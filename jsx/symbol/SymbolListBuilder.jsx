#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

ドキュメントに登録されたシンボルを一覧表示する専用アートボード「シンボル一覧」を自動生成します。
ダイアログでパラメーターを操作しながらライブプレビューで確認でき、［OK］で確定します。

詳細は README を参照してください。

### Overview

Generates a dedicated "Symbol List" artboard that lays out every symbol registered in the document.
Parameters are adjusted with a live preview in the dialog, and OK commits the result.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SymbolListBuilder";            /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.2.3";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-05-09";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-16";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SymbolListBuilder.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SymbolListBuilder.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/ncac687d0a3a0"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /**
     * @typedef {object} ListSettings
     * @property {string} position - 作成方向（"right" / "below"）
     * @property {string} baseMode - 基準（"last" = 最終アートボード / "specified" = 番号指定）
     * @property {number} baseArtboardNumber - 基準にするアートボード番号（1 始まり）
     * @property {number} artboardGap - 基準アートボードとの間隔（pt）
     * @property {number} margin - アートボード内側の余白（pt）
     * @property {boolean} update - 既存のシンボル一覧を削除して作り直すか
     * @property {boolean} showCaption - シンボル名を表示するか
     * @property {string} filter - 収集対象（"all" / "used"）
     * @property {number} symbolGap - シンボル同士の間隔（pt）
     * @property {number} maxRowWidth - 1 行の最大幅（pt）
     * @property {string} bgColor - 背景（BACKGROUND_CHOICES のいずれか）
     * @property {string} captionPosition - キャプションの位置（"above" / "below"）
     * @property {number} fontSize - キャプションのフォントサイズ（pt）
     * @property {?number} [widthOverridePt] - 固定する幅（pt、null なら自動）
     * @property {?number} [heightOverridePt] - 固定する高さ（pt、null なら自動）
     */

    /* 設定の既定値（寸法は pt）。［更新］と［収集対象］は保存値を使わず、起動時は常に ON／すべて
     * Default settings (sizes in pt). Update and Collect ignore saved values and always start as on / all */
    var DEFAULT_SETTINGS = {
        position: "below",
        baseMode: "last",
        baseArtboardNumber: 1,
        artboardGap: 100,
        margin: 50,
        update: true,
        showCaption: false,
        filter: "all",
        symbolGap: 20,
        maxRowWidth: 800,
        bgColor: "none",
        captionPosition: "below",
        fontSize: 9
    };

    /* シンボルとキャプションの間隔（pt）/ Gap between a symbol and its caption (pt) */
    var CAPTION_GAP_PT = 6;

    /* キャプションのフォント（UI 言語別。見つからなければ既定のまま）/ Caption font per UI language (left as is if missing) */
    var CAPTION_FONT_NAMES = { ja: "HiraginoSans-W3", en: "MyriadPro-Regular" };

    /* 起動直後のプレビューのズーム倍率 / Zoom factor for the initial preview */
    var INITIAL_PREVIEW_ZOOM = 0.6;

    /* ウィンドウに合わせた後に掛けるズーム倍率 / Zoom ratio applied after fitting to the window */
    var FIT_ZOOM_RATIO = 0.9;

    // =========================================
    // レイアウト / Layout
    // =========================================

    var WINDOW_MARGINS          = 16;                 /* ウィンドウ外周の余白 / window margin */
    var WINDOW_SPACING          = 12;                 /* ウィンドウ内の要素間隔 / window spacing */
    var PANEL_MARGINS           = [16, 20, 16, 12];   /* パネル余白 [左,上,右,下] / panel margins */
    var PANEL_SPACING           = 8;                  /* パネル内の要素間隔 / panel spacing */
    var COLUMN_SPACING          = 12;                 /* 2カラムの間隔 / column spacing */
    var ROW_SPACING             = 6;                  /* 行内の要素間隔 / spacing within a row */
    var RADIO_SPACING           = 10;                 /* 横並びラジオの間隔 / spacing between radios */
    var BUTTON_SPACING          = 8;                  /* ボタンの間隔 / spacing between buttons */
    var FIELD_LABEL_WIDTH       = 70;                 /* 項目名の幅 / field label width */
    var NUMBER_INPUT_CHARACTERS = 4;                  /* 数値入力欄の文字数 / numeric field width */
    var BACKGROUND_RADIO_WIDTH  = 60;                 /* 背景ラジオの幅（2×2 の列揃え）/ background radio width */

    /**
     * ウィンドウの共通レイアウトを設定する
     * @param {Window} targetWindow - 対象のウィンドウ
     * @returns {void}
     */
    function setupWindow(targetWindow) {
        targetWindow.orientation = "column";
        targetWindow.alignChildren = "fill";
        targetWindow.margins = WINDOW_MARGINS;
        targetWindow.spacing = WINDOW_SPACING;
    }

    /**
     * タイトル付きのパネルを追加し、共通レイアウトを設定する
     * @param {object} parent - 追加先のパネルまたはグループ
     * @param {object} titleSet - パネル名のラベル（ja/en）
     * @returns {Panel} 追加したパネル
     */
    function addPanel(parent, titleSet) {
        var newPanel = parent.add("panel", undefined, getLabel(titleSet));
        newPanel.orientation = "column";
        newPanel.alignChildren = ["fill", "top"];
        newPanel.alignment = "fill";
        newPanel.margins = PANEL_MARGINS;
        newPanel.spacing = PANEL_SPACING;
        return newPanel;
    }

    /**
     * 2カラムの列にする縦並びのグループを追加する
     * @param {Group} parent - 追加先のグループ
     * @returns {Group} 追加した列グループ
     */
    function addColumnGroup(parent) {
        var columnGroup = parent.add("group");
        columnGroup.orientation = "column";
        columnGroup.alignChildren = "fill";
        return columnGroup;
    }

    /**
     * 横並びの行グループを追加する（子は左詰め・天地中央）
     * @param {object} parent - 追加先のパネルまたはグループ
     * @param {number} [spacing] - 要素間隔（省略時は ROW_SPACING）
     * @returns {Group} 追加した行グループ
     */
    function addRowGroup(parent, spacing) {
        var rowGroup = parent.add("group");
        rowGroup.orientation = "row";
        rowGroup.alignChildren = ["left", "center"];
        rowGroup.spacing = (typeof spacing === "number") ? spacing : ROW_SPACING;
        return rowGroup;
    }

    /**
     * 行の先頭に右揃えの項目名を追加する
     * @param {Group} rowGroup - 追加先の行グループ
     * @param {object} labelSet - 項目名のラベル（ja/en）
     * @returns {StaticText} 追加した項目名
     */
    function addFieldLabel(rowGroup, labelSet) {
        var fieldLabel = rowGroup.add("statictext", undefined, labelText(labelSet));
        fieldLabel.preferredSize.width = FIELD_LABEL_WIDTH;
        fieldLabel.justify = "right";
        return fieldLabel;
    }

    /**
     * ラベルとツールチップ付きのラジオボタン・チェックボックス・ボタンを追加する
     * @param {object} parent - 追加先のパネルまたはグループ
     * @param {string} controlType - "radiobutton" / "checkbox" / "button"
     * @param {object} labelSet - 表示するラベル（ja/en）
     * @param {object} tooltipSet - ツールチップ（ja/en）
     * @returns {object} 追加したコントロール
     */
    function addLabeledControl(parent, controlType, labelSet, tooltipSet) {
        var control = parent.add(controlType, undefined, getLabel(labelSet));
        control.helpTip = getLabel(tooltipSet);
        return control;
    }

    /**
     * ツールチップ付きの数値入力欄を追加する
     * @param {Group} parent - 追加先の行グループ
     * @param {string} initialText - 初期表示する値
     * @param {object} tooltipSet - ツールチップ（ja/en）
     * @returns {EditText} 追加した入力欄
     */
    function addNumberInput(parent, initialText, tooltipSet) {
        var numberInput = parent.add("edittext", undefined, initialText);
        numberInput.characters = NUMBER_INPUT_CHARACTERS;
        numberInput.helpTip = getLabel(tooltipSet);
        return numberInput;
    }

    /**
     * 親の異なるラジオボタンを手動で排他選択にする
     * @param {RadioButton[]} radios - 排他にするラジオボタン
     * @param {RadioButton} selectedRadio - 選択するラジオボタン
     * @returns {void}
     */
    function selectExclusiveRadio(radios, selectedRadio) {
        for (var i = 0; i < radios.length; i++) {
            radios[i].value = (radios[i] === selectedRadio);
        }
    }

    /**
     * ↑↓キーで数値入力欄の値を増減する（↑↓: ±1、Shift+↑↓: ±10 で 10 の倍数にスナップ）
     * @param {EditText} editText - 対象の入力欄
     * @param {boolean} allowNegative - 負の値を許すか（false なら 0 が下限）
     * @param {function(): void} [onChange] - 値を変えた後に呼ぶ処理
     * @param {number} [minValue] - 下限値
     * @returns {void}
     */
    function changeValueByArrowKey(editText, allowNegative, onChange, minValue) {
        editText.addEventListener("keydown", function (event) {
            if (event.keyName !== "Up" && event.keyName !== "Down") return;
            var current = Number(editText.text);
            if (isNaN(current)) return;

            var sign = (event.keyName === "Up") ? 1 : -1;
            var step = ScriptUI.environment.keyboardState.shiftKey ? 10 : 1;
            var next = (step === 10)
                ? (sign > 0 ? Math.ceil((current + 1) / step) * step
                    : Math.floor((current - 1) / step) * step)
                : current + sign * step;

            next = Math.round(next);
            if (!allowNegative && next < 0) next = 0;
            if (typeof minValue === "number" && next < minValue) next = minValue;

            editText.text = next;
            event.preventDefault();
            if (typeof onChange === "function") onChange();
        });
    }

    // =========================================
    // 定数 / Constants
    // =========================================

    /* 設定保存用キー / Preference key for saved settings */
    var PREF_KEY = "swwwitch.listupallsymbol.settings";

    /* 保存する設定のキー（保存文字列の並び順）/ Keys of saved settings (serialized order) */
    var SETTINGS_KEYS = [
        "position", "baseMode", "baseArtboardNumber", "artboardGap", "margin",
        "update", "showCaption", "filter", "symbolGap", "maxRowWidth",
        "bgColor", "captionPosition", "fontSize"
    ];

    /* 背景の選択肢（ラジオの並び順）と、塗りに使う墨の濃度（%）/ Background choices (radio order) and black ink percentage */
    var BACKGROUND_CHOICES = ["none", "black", "white", "gray"];
    var BACKGROUND_BLACK_PERCENT = { black: 100, white: 0, gray: 50 };

    /* 選択肢で持つ設定と、取りうる値 / Settings stored as one of fixed choices */
    var SETTING_CHOICES = {
        position: ["right", "below"],
        baseMode: ["last", "specified"],
        filter: ["all", "used"],
        bgColor: BACKGROUND_CHOICES,
        captionPosition: ["above", "below"]
    };

    /* 旧アートボード上のアイテムを拾う許容幅（pt）/ Tolerance for picking items on an old artboard (pt) */
    var ARTBOARD_HIT_TOLERANCE_PT = 1.0;

    /* プレビュー用アートボードの矩形を照合する許容誤差（pt）/ Tolerance for matching the preview artboard rect (pt) */
    var RECT_MATCH_TOLERANCE_PT = 0.01;

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /* 現在の UI 言語 / Current UI language */
    var uiLang = ($.locale.indexOf("ja") === 0) ? "ja" : "en";

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "シンボル一覧を作成", en: "Create Symbol List" }
        },
        panel: {
            artboard: { ja: "作成するアートボード", en: "New artboard" },
            location: { ja: "作成位置", en: "Location" },
            sizeAndPadding: { ja: "サイズと余白", en: "Size & padding" },
            background: { ja: "背景", en: "Background" },
            symbols: { ja: "収集するシンボル", en: "Symbols" },
            collectTarget: { ja: "収集対象", en: "Collect" },
            arrangement: { ja: "並べ方", en: "Layout" },
            caption: { ja: "キャプション", en: "Caption" }
        },
        radio: {
            baseLast: { ja: "最終アートボード", en: "Last artboard" },
            baseSpecified: { ja: "指定", en: "Specified" },
            directionRight: { ja: "右側", en: "Right" },
            directionBelow: { ja: "下側", en: "Below" },
            background: {
                none: { ja: "なし", en: "None" },
                black: { ja: "黒", en: "Black" },
                white: { ja: "白", en: "White" },
                gray: { ja: "グレー", en: "Gray" }
            },
            captionAbove: { ja: "上", en: "Top" },
            captionBelow: { ja: "下", en: "Bottom" },
            filterAll: { ja: "すべて", en: "All" },
            filterUsed: { ja: "使用中のみ", en: "Used only" }
        },
        checkbox: {
            update: { ja: "更新", en: "Update" },
            showCaption: { ja: "シンボル名を表示", en: "Show symbol name" }
        },
        fieldLabel: {
            direction: { ja: "方向", en: "Direction" },
            gap: { ja: "間隔", en: "Gap" },
            width: { ja: "幅", en: "Width" },
            height: { ja: "高さ", en: "Height" },
            margin: { ja: "余白", en: "Padding" },
            maxRowWidth: { ja: "最大幅", en: "Max row width" },
            captionPosition: { ja: "位置", en: "Position" },
            fontSize: { ja: "フォントサイズ", en: "Font size" }
        },
        tooltip: {
            baseLast: {
                ja: "アートボード一覧の末尾を基準にする。更新 ON の場合は既存のシンボル一覧を除外",
                en: "Use the last artboard in the list as the base; with Update on, the existing symbol list is excluded"
            },
            baseSpecified: { ja: "指定した番号のアートボードを基準にする（1 始まり）", en: "Use the artboard with the given number as the base (1-based)" },
            baseArtboardNumber: { ja: "基準にするアートボード番号（1 始まり）", en: "Base artboard number (1-based)" },
            directionRight: { ja: "基準アートボードの右に新規アートボードを作成", en: "Place the new artboard to the right of the base" },
            directionBelow: { ja: "基準アートボードの下に新規アートボードを作成", en: "Place the new artboard below the base" },
            artboardGap: { ja: "基準アートボードと新規アートボードの間隔", en: "Distance between the base and new artboards" },
            margin: { ja: "新規アートボードの内側に設ける余白", en: "Padding inside the new artboard around the symbols" },
            symbolGap: { ja: "シンボル同士の間隔", en: "Spacing between adjacent symbols" },
            maxRowWidth: { ja: "1行に並べる最大幅。超えると次の行に折り返し", en: "Max width per row; symbols wrap to the next row when exceeded" },
            width: {
                ja: "アートボードの幅。空欄または 0 で自動計算、入力すると指定値で固定",
                en: "Artboard width. Empty/0 = auto; a number forces that exact size"
            },
            height: {
                ja: "アートボードの高さ。空欄または 0 で自動計算、入力すると指定値で固定",
                en: "Artboard height. Empty/0 = auto; a number forces that exact size"
            },
            update: {
                ja: "既存の「シンボル一覧」アートボードと対象オブジェクトを削除して作り直す",
                en: "Delete the existing Symbol List artboard and its objects, then rebuild"
            },
            showCaption: { ja: "各シンボルの近くにシンボル名をキャプションとして表示", en: "Show the symbol name as a caption near each symbol" },
            captionAbove: { ja: "シンボルの上にシンボル名を表示", en: "Place the name above the symbol" },
            captionBelow: { ja: "シンボルの下にシンボル名を表示", en: "Place the name below the symbol" },
            fontSize: {
                ja: "キャプションのフォントサイズ。単位は Illustrator の文字設定に従う",
                en: "Caption font size; unit follows Illustrator's type preferences"
            },
            filterAll: { ja: "ドキュメントに登録されているすべてのシンボルを並べる", en: "List every symbol registered in the document" },
            filterUsed: { ja: "ドキュメント内に配置されているシンボルだけを並べる", en: "List only symbols placed in the document" },
            background: {
                none: { ja: "背景の塗りを作成しない", en: "Do not create a background fill" },
                black: { ja: "アートボード背面に黒（K100）の塗りを敷く", en: "Place a solid black (K100) fill behind the artboard" },
                white: { ja: "アートボード背面に白の塗りを敷く", en: "Place a solid white fill behind the artboard" },
                gray: { ja: "アートボード背面にグレー（K50）の塗りを敷く", en: "Place a 50% gray (K50) fill behind the artboard" }
            },
            fitSymbolList: { ja: "作成したアートボードのみをウィンドウに合わせて表示", en: "Fit the created artboard to the window" },
            fitAll: { ja: "すべてのアートボードをウィンドウに合わせて表示", en: "Fit all artboards to the window" }
        },
        button: {
            fitSymbolList: { ja: "シンボル一覧", en: "Symbol list" },
            fitAll: { ja: "全体表示", en: "Fit all" },
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" }
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noSymbols: { ja: "登録されているシンボルがありません。", en: "No symbols are registered." },
            noUsedSymbols: { ja: "ドキュメント内で使用中のシンボルがありません。", en: "No symbols are currently used in the document." }
        },
        itemName: {
            symbolList: { ja: "シンボル一覧", en: "Symbol List" }
        }
    };

    /**
     * ラベル（ja/en）を現在の UI 言語の文字列にする
     * @param {object} labelSet - ja/en を持つラベル
     * @returns {string} 現在の言語の文字列
     */
    function getLabel(labelSet) {
        return (labelSet && labelSet[uiLang]) || "";
    }

    /**
     * 項目名にコロンを付ける（日本語は全角、英語は半角）
     * @param {object} labelSet - ja/en を持つラベル
     * @returns {string} コロン付きの項目名
     */
    function labelText(labelSet) {
        return getLabel(labelSet) + (uiLang === "ja" ? "：" : ":");
    }

    /* 作成するレイヤーとアートボードの名前 / Name of the created layer and artboard */
    var SYMBOL_LIST_NAME = getLabel(LABELS.itemName.symbolList);

    /**
     * シンボル一覧のレイヤー名・アートボード名か（別の言語で作成したものも含む）
     * @param {string} name - 調べる名前
     * @returns {boolean} シンボル一覧の名前なら true
     */
    function isSymbolListName(name) {
        return name === LABELS.itemName.symbolList.ja || name === LABELS.itemName.symbolList.en;
    }

    // =========================================
    // 単位 / Units
    // =========================================

    /**
     * @typedef {object} UnitInfo
     * @property {string} label - 表示する単位名
     * @property {number} factor - 1 単位あたりの pt
     */

    /**
     * 単位のプリファレンスから表示単位と pt への換算係数を取得する
     * （0=inch / 1=mm / 3=pica / 4=cm / 5=Q / 6=px / その他=pt）
     * @param {string} preferenceKey - 整数プリファレンスのキー（"rulerType" / "text/units"）
     * @returns {UnitInfo} 単位情報
     */
    function getUnitInfo(preferenceKey) {
        switch (app.preferences.getIntegerPreference(preferenceKey)) {
            case 0: return { label: "inch", factor: 72.0 };
            case 1: return { label: "mm", factor: 72.0 / 25.4 };
            case 3: return { label: "pica", factor: 12.0 };
            case 4: return { label: "cm", factor: 72.0 / 2.54 };
            case 5: return { label: "Q", factor: 72.0 / 25.4 * 0.25 };
            case 6: return { label: "px", factor: 1.0 };
            default: return { label: "pt", factor: 1.0 };
        }
    }

    /* ルーラー単位（寸法・余白・間隔）/ Ruler unit for sizes, padding and gaps */
    var RULER_UNIT = getUnitInfo("rulerType");

    /* 文字の単位（フォントサイズ）/ Type unit for font size */
    var TYPE_UNIT = getUnitInfo("text/units");

    /**
     * pt の値を表示単位の整数の文字列にする
     * @param {number} valuePt - 値（pt）
     * @param {UnitInfo} unitInfo - 表示単位
     * @returns {string} 表示用の文字列
     */
    function formatUnitValue(valuePt, unitInfo) {
        return String(Math.round(valuePt / unitInfo.factor));
    }

    /**
     * ルーラー単位の入力欄を pt で読む
     * @param {EditText} valueInput - 入力欄
     * @param {number} defaultPt - 数値でないときの値（pt）
     * @returns {number} 値（pt）
     */
    function readUnitValuePt(valueInput, defaultPt) {
        var parsedValue = parseFloat(valueInput.text);
        return isNaN(parsedValue) ? defaultPt : parsedValue * RULER_UNIT.factor;
    }

    /**
     * 文字の単位のフォントサイズ欄を pt で読む（0 以下や数値以外は既定値）
     * @param {EditText} fontSizeInput - フォントサイズの入力欄
     * @returns {number} フォントサイズ（pt）
     */
    function readFontSizePt(fontSizeInput) {
        var fontSize = parseFloat(fontSizeInput.text);
        return (isNaN(fontSize) || fontSize <= 0) ? DEFAULT_SETTINGS.fontSize : fontSize * TYPE_UNIT.factor;
    }

    /**
     * アートボード番号の文字列を 1 以上の整数にする
     * @param {string} text - 入力された文字列
     * @returns {number} アートボード番号（1 始まり）
     */
    function parseArtboardNumber(text) {
        var artboardNumber = parseInt(text, 10);
        return (isNaN(artboardNumber) || artboardNumber < 1) ? 1 : artboardNumber;
    }

    // =========================================
    // 設定の保存・復元 / Save & restore settings
    // =========================================

    /**
     * 設定を JSON 形式の文字列にする（ES3 互換の手書き。文字列は引用し、数値・真偽値はそのまま）
     * @param {ListSettings} settings - 保存する設定
     * @returns {string} 保存用の文字列
     */
    function serializeSettings(settings) {
        var jsonParts = [];
        for (var i = 0; i < SETTINGS_KEYS.length; i++) {
            var key = SETTINGS_KEYS[i];
            var value = settings[key];
            jsonParts.push('"' + key + '":' + ((typeof value === "string") ? '"' + value + '"' : String(value)));
        }
        return "{" + jsonParts.join(",") + "}";
    }

    /**
     * 保存文字列からキーと値の組を読み出す（eval は使わない）
     * @param {string} settingsText - 保存文字列
     * @returns {object} キーごとの値（文字列・数値・真偽値）
     */
    function readSavedValues(settingsText) {
        var savedValues = {};
        var pairPattern = /"(\w+)"\s*:\s*("[^"\\]*"|-?\d+(?:\.\d+)?|true|false)/g;
        var pairMatch;
        while ((pairMatch = pairPattern.exec(settingsText)) !== null) {
            var rawValue = pairMatch[2];
            if (rawValue.charAt(0) === '"') {
                savedValues[pairMatch[1]] = rawValue.slice(1, -1);
            } else if (rawValue === "true" || rawValue === "false") {
                savedValues[pairMatch[1]] = (rawValue === "true");
            } else {
                savedValues[pairMatch[1]] = Number(rawValue);
            }
        }
        return savedValues;
    }

    /**
     * 保存値として使えるか（型が既定値と同じで、選択肢や下限を満たす）
     * @param {string} key - 設定のキー
     * @param {*} value - 保存値
     * @returns {boolean} 使えるなら true
     */
    function isValidSettingValue(key, value) {
        if (typeof value !== typeof DEFAULT_SETTINGS[key]) return false;
        if (SETTING_CHOICES.hasOwnProperty(key)) {
            for (var i = 0; i < SETTING_CHOICES[key].length; i++) {
                if (value === SETTING_CHOICES[key][i]) return true;
            }
            return false;
        }
        if (key === "fontSize") return value > 0;
        if (key === "baseArtboardNumber") return value >= 1;
        return true;
    }

    /**
     * 保存文字列から設定を復元する（使えない値は既定値に戻す）
     * @param {string} settingsText - 保存文字列（空なら全項目が既定値）
     * @returns {ListSettings} 復元した設定
     */
    function parseSettings(settingsText) {
        var savedValues = readSavedValues(settingsText || "");
        var settings = {};
        for (var i = 0; i < SETTINGS_KEYS.length; i++) {
            var key = SETTINGS_KEYS[i];
            settings[key] = isValidSettingValue(key, savedValues[key]) ? savedValues[key] : DEFAULT_SETTINGS[key];
        }
        return settings;
    }

    /**
     * 設定を app.preferences に保存する
     * @param {ListSettings} settings - 保存する設定
     * @returns {void}
     */
    function saveSettings(settings) {
        try { app.preferences.setStringPreference(PREF_KEY, serializeSettings(settings)); } catch (e) { }
    }

    /**
     * app.preferences から設定を読み出す（保存値が無ければ既定値）
     * @returns {ListSettings} 設定
     */
    function loadSettings() {
        var settingsText = "";
        try { settingsText = app.preferences.getStringPreference(PREF_KEY); } catch (e) { }
        return parseSettings(settingsText);
    }

    // =========================================
    // シンボルの収集 / Collect symbols
    // =========================================

    /**
     * 配置されているシンボル名の集合を返す（シンボル一覧レイヤー上のインスタンスは数えない）
     * @param {Document} doc - 対象のドキュメント
     * @returns {object} シンボル名をキー、true を値にしたオブジェクト
     */
    function getUsedSymbolNames(doc) {
        var usedNames = {};
        var symbolItems = doc.symbolItems;
        for (var i = 0; i < symbolItems.length; i++) {
            if (isSymbolListName(symbolItems[i].layer.name)) continue;
            usedNames[symbolItems[i].symbol.name] = true;
        }
        return usedNames;
    }

    /**
     * 収集対象に応じて並べるシンボルを返す
     * @param {Document} doc - 対象のドキュメント
     * @param {string} filter - 収集対象（"all" / "used"）
     * @returns {Symbol[]} 並べるシンボル
     */
    function getTargetSymbols(doc, filter) {
        var usedNames = (filter === "used") ? getUsedSymbolNames(doc) : null;
        var targetSymbols = [];
        for (var i = 0; i < doc.symbols.length; i++) {
            if (!usedNames || usedNames[doc.symbols[i].name]) targetSymbols.push(doc.symbols[i]);
        }
        return targetSymbols;
    }

    // =========================================
    // シンボル一覧の作成 / Build the symbol list
    // =========================================

    /**
     * @typedef {object} SymbolEntry
     * @property {SymbolItem} symbolItem - 配置したシンボルインスタンス
     * @property {number} symbolWidth - シンボルの幅（pt）
     * @property {number} symbolHeight - シンボルの高さ（pt）
     * @property {TextFrame} [caption] - シンボル名のキャプション
     * @property {number} [captionWidth] - キャプションの幅（pt）
     * @property {number} [captionHeight] - キャプションの高さ（pt）
     * @property {number} width - シンボルとキャプションを合わせた枠の幅（pt）
     * @property {number} height - シンボルとキャプションを合わせた枠の高さ（pt）
     * @property {number} x - 枠の左上の相対位置（pt、パッキングで決まる）
     * @property {number} y - 枠の左上の相対位置（pt、下方向が負）
     */

    /**
     * @typedef {object} ArtboardInfo
     * @property {number} index - 作成時のアートボードの index
     * @property {number[]} rect - 作成時の矩形 [left, top, right, bottom]
     */

    /**
     * @typedef {object} SymbolListLayout
     * @property {Layer} layer - シンボル一覧のレイヤー
     * @property {ArtboardInfo} artboardInfo - シンボル一覧のアートボード
     */

    /**
     * シンボル一覧用のレイヤーを作成する
     * @param {Document} doc - 対象のドキュメント
     * @returns {?Layer} 作成したレイヤー（作成できなければ null）
     */
    function createSymbolListLayer(doc) {
        try {
            var listLayer = doc.layers.add();
            listLayer.name = SYMBOL_LIST_NAME;
            return listLayer;
        } catch (e) {
            return null;
        }
    }

    /**
     * ページアイテムの表示上の幅と高さを返す
     * @param {PageItem} pageItem - 対象のアイテム
     * @returns {{width: number, height: number}} 幅と高さ（pt）
     */
    function getVisibleSize(pageItem) {
        var bounds = pageItem.visibleBounds; // [left, top, right, bottom]
        return { width: bounds[2] - bounds[0], height: bounds[1] - bounds[3] };
    }

    /**
     * 墨の濃度から、ドキュメントのカラーモードに合ったグレーの色を作る
     * @param {Document} doc - 対象のドキュメント
     * @param {number} blackPercent - 墨の濃度（0〜100）
     * @returns {Color} CMYKColor または RGBColor
     */
    function createGrayColor(doc, blackPercent) {
        if (doc.documentColorSpace === DocumentColorSpace.CMYK) {
            var cmykColor = new CMYKColor();
            cmykColor.cyan = 0; cmykColor.magenta = 0; cmykColor.yellow = 0; cmykColor.black = blackPercent;
            return cmykColor;
        }
        var channelValue = Math.round(255 * (1 - blackPercent / 100));
        var rgbColor = new RGBColor();
        rgbColor.red = channelValue; rgbColor.green = channelValue; rgbColor.blue = channelValue;
        return rgbColor;
    }

    /**
     * シンボル名のキャプションを作成する
     * @param {Layer} listLayer - 作成先のレイヤー
     * @param {string} symbolName - シンボル名
     * @param {number} fontSizePt - フォントサイズ（pt）
     * @param {?Color} fillColor - 文字の色（null なら既定のまま）
     * @returns {TextFrame} 作成したキャプション
     */
    function createCaption(listLayer, symbolName, fontSizePt, fillColor) {
        var caption = listLayer.textFrames.add();
        caption.contents = symbolName;
        var textRange = caption.textRange;
        textRange.paragraphAttributes.justification = Justification.CENTER;
        if (fillColor) textRange.characterAttributes.fillColor = fillColor;
        /* 範囲外のサイズは既定のまま / Keep the default size when out of range */
        try { textRange.characterAttributes.size = fontSizePt; } catch (e) { }
        /* フォントが無ければ既定のまま / Keep the default font when not installed */
        try { textRange.characterAttributes.textFont = app.textFonts.getByName(CAPTION_FONT_NAMES[uiLang]); } catch (e) { }
        return caption;
    }

    /**
     * シンボルを配置し（必要ならキャプションも作成して）、大きさを測ったエントリを返す
     * @param {Document} doc - 対象のドキュメント
     * @param {Layer} listLayer - 配置先のレイヤー
     * @param {ListSettings} settings - 設定
     * @param {Symbol[]} targetSymbols - 並べるシンボル
     * @returns {SymbolEntry[]} エントリ
     */
    function createSymbolEntries(doc, listLayer, settings, targetSymbols) {
        /* 背景が黒のときだけキャプションを白に / White captions only on a black background */
        var captionFillColor = (settings.bgColor === "black") ? createGrayColor(doc, 0) : null;
        var entries = [];
        for (var i = 0; i < targetSymbols.length; i++) {
            var symbolItem = doc.symbolItems.add(targetSymbols[i]);
            symbolItem.moveToBeginning(listLayer);
            var symbolSize = getVisibleSize(symbolItem);
            var entry = {
                symbolItem: symbolItem,
                symbolWidth: symbolSize.width,
                symbolHeight: symbolSize.height,
                width: symbolSize.width,
                height: symbolSize.height
            };

            if (settings.showCaption) {
                entry.caption = createCaption(listLayer, targetSymbols[i].name, settings.fontSize, captionFillColor);
                var captionSize = getVisibleSize(entry.caption);
                entry.captionWidth = captionSize.width;
                entry.captionHeight = captionSize.height;
                entry.width = Math.max(symbolSize.width, captionSize.width);
                entry.height = symbolSize.height + CAPTION_GAP_PT + captionSize.height;
            }
            entries.push(entry);
        }
        return entries;
    }

    /**
     * エントリを左から詰め、最大幅を超えたら次の行に折り返して位置を決める（シェルフパッキング）
     * @param {SymbolEntry[]} entries - エントリ（x, y を書き込む）
     * @param {number} maxRowWidth - 1 行の最大幅（pt）
     * @param {number} gap - エントリ同士の間隔（pt）
     * @returns {{width: number, height: number}} 並べた全体の幅と高さ（pt）
     */
    function packEntriesIntoRows(entries, maxRowWidth, gap) {
        var rowX = 0, rowY = 0, rowHeight = 0, totalWidth = 0;
        for (var i = 0; i < entries.length; i++) {
            var entry = entries[i];
            if (rowX > 0 && rowX + entry.width > maxRowWidth) {
                rowY -= rowHeight + gap;
                rowX = 0;
                rowHeight = 0;
            }
            entry.x = rowX;
            entry.y = rowY;
            rowX += entry.width + gap;
            if (rowX - gap > totalWidth) totalWidth = rowX - gap;
            if (entry.height > rowHeight) rowHeight = entry.height;
        }
        return { width: totalWidth, height: -rowY + rowHeight };
    }

    /**
     * キャンバス上で最も右下にあるアートボードの番号を返す（シンボル一覧は除く）
     * 右端が大きく下端が小さいほど右下とみなし、right − bottom で比べる
     * @param {Document} doc - 対象のドキュメント
     * @returns {number} アートボード番号（1 始まり）
     */
    function findBottomRightArtboardNumber(doc) {
        var bestNumber = 1;
        var bestScore = null;
        for (var i = 0; i < doc.artboards.length; i++) {
            if (isSymbolListName(doc.artboards[i].name)) continue;
            var artboardRect = doc.artboards[i].artboardRect; // [left, top, right, bottom]
            var score = artboardRect[2] - artboardRect[3];
            if (bestScore === null || score > bestScore) {
                bestScore = score;
                bestNumber = i + 1;
            }
        }
        return bestNumber;
    }

    /**
     * 基準アートボードの矩形を返す
     * - 指定：その番号のアートボード（範囲外は最後のアートボード）
     * - 最終・更新 ON：OK 時に消える既存のシンボル一覧を除いた、最後のアートボード
     * - 最終・更新 OFF：最後のアートボード
     * @param {Document} doc - 対象のドキュメント
     * @param {ListSettings} settings - 設定
     * @returns {number[]} 矩形 [left, top, right, bottom]
     */
    function resolveBaseArtboardRect(doc, settings) {
        var artboards = doc.artboards;
        var lastIndex = artboards.length - 1;
        if (settings.baseMode === "specified") {
            return artboards[Math.min(settings.baseArtboardNumber - 1, lastIndex)].artboardRect;
        }
        if (settings.update) {
            for (var i = lastIndex; i >= 0; i--) {
                if (!isSymbolListName(artboards[i].name)) return artboards[i].artboardRect;
            }
        }
        return artboards[lastIndex].artboardRect;
    }

    /**
     * 基準アートボードの右または下に、新しいアートボードを置く左上の座標を返す
     * @param {Document} doc - 対象のドキュメント
     * @param {ListSettings} settings - 設定
     * @returns {{left: number, top: number}} 左上の座標
     */
    function computeArtboardOrigin(doc, settings) {
        var baseRect = resolveBaseArtboardRect(doc, settings);
        if (settings.position === "right") {
            return { left: baseRect[2] + settings.artboardGap, top: baseRect[1] };
        }
        return { left: baseRect[0], top: baseRect[3] - settings.artboardGap };
    }

    /**
     * シンボル一覧のアートボードを追加する
     * @param {Document} doc - 対象のドキュメント
     * @param {{left: number, top: number}} origin - 左上の座標
     * @param {number} width - 幅（pt）
     * @param {number} height - 高さ（pt）
     * @returns {ArtboardInfo} 追加したアートボードの情報
     */
    function addSymbolListArtboard(doc, origin, width, height) {
        var artboardIndex = doc.artboards.length;
        var rect = [origin.left, origin.top, origin.left + width, origin.top - height];
        doc.artboards.add(rect).name = SYMBOL_LIST_NAME;
        return { index: artboardIndex, rect: rect };
    }

    /**
     * パッキングの結果に従ってシンボルとキャプションを配置する
     * @param {SymbolEntry[]} entries - エントリ
     * @param {{left: number, top: number}} origin - アートボードの左上
     * @param {number} margin - アートボード内側の余白（pt）
     * @param {string} captionPosition - キャプションの位置（"above" / "below"）
     * @returns {void}
     */
    function placeEntries(entries, origin, margin, captionPosition) {
        var isCaptionAbove = (captionPosition === "above");
        for (var i = 0; i < entries.length; i++) {
            var entry = entries[i];
            var slotLeft = origin.left + margin + entry.x;
            var slotTop = origin.top - margin + entry.y;

            /* 枠の中で水平中央に揃える / Center horizontally within the slot */
            var symbolTop = (entry.caption && isCaptionAbove) ? slotTop - entry.captionHeight - CAPTION_GAP_PT : slotTop;
            entry.symbolItem.position = [slotLeft + (entry.width - entry.symbolWidth) / 2, symbolTop];

            if (entry.caption) {
                var captionTop = isCaptionAbove ? slotTop : slotTop - entry.symbolHeight - CAPTION_GAP_PT;
                entry.caption.position = [slotLeft + (entry.width - entry.captionWidth) / 2, captionTop];
            }
        }
    }

    /**
     * アートボードいっぱいの背景の塗りを作り、レイヤーの最背面に送る
     * @param {Layer} listLayer - 作成先のレイヤー
     * @param {number[]} rect - アートボードの矩形 [left, top, right, bottom]
     * @param {Color} fillColor - 塗りの色
     * @returns {void}
     */
    function createBackgroundFill(listLayer, rect, fillColor) {
        var backgroundRect = listLayer.pathItems.rectangle(rect[1], rect[0], rect[2] - rect[0], rect[1] - rect[3]);
        backgroundRect.filled = true;
        backgroundRect.stroked = false;
        backgroundRect.fillColor = fillColor;
        backgroundRect.zOrder(ZOrderMethod.SENDTOBACK);
    }

    /**
     * シンボル一覧（レイヤー・シンボル・キャプション・アートボード・背景）を作成する
     * 既存のシンボル一覧はここでは消さない（削除は OK 時の removeOtherSymbolLists）
     * @param {Document} doc - 対象のドキュメント
     * @param {ListSettings} settings - 設定
     * @returns {?SymbolListLayout} 作成した一覧（並べるシンボルが無ければ null）
     */
    function buildLayout(doc, settings) {
        var targetSymbols = getTargetSymbols(doc, settings.filter);
        if (targetSymbols.length === 0) return null;
        var listLayer = createSymbolListLayer(doc);
        if (!listLayer) return null;

        var entries = createSymbolEntries(doc, listLayer, settings, targetSymbols);
        var contentSize = packEntriesIntoRows(entries, settings.maxRowWidth, settings.symbolGap);
        var origin = computeArtboardOrigin(doc, settings);
        var artboardWidth = (settings.widthOverridePt > 0) ? settings.widthOverridePt : contentSize.width + settings.margin * 2;
        var artboardHeight = (settings.heightOverridePt > 0) ? settings.heightOverridePt : contentSize.height + settings.margin * 2;
        var artboardInfo = addSymbolListArtboard(doc, origin, artboardWidth, artboardHeight);
        placeEntries(entries, origin, settings.margin, settings.captionPosition);

        if (settings.bgColor !== "none") {
            createBackgroundFill(listLayer, artboardInfo.rect, createGrayColor(doc, BACKGROUND_BLACK_PERCENT[settings.bgColor]));
        }
        return { layer: listLayer, artboardInfo: artboardInfo };
    }

    // =========================================
    // シンボル一覧の削除 / Remove symbol lists
    // =========================================

    /**
     * 2 つの矩形が許容誤差の範囲で一致するか
     * @param {number[]} rectA - 矩形 [left, top, right, bottom]
     * @param {number[]} rectB - 矩形 [left, top, right, bottom]
     * @returns {boolean} 一致すれば true
     */
    function isSameRect(rectA, rectB) {
        for (var i = 0; i < 4; i++) {
            if (Math.abs(rectA[i] - rectB[i]) >= RECT_MATCH_TOLERANCE_PT) return false;
        }
        return true;
    }

    /**
     * 作成したシンボル一覧のアートボードの現在の index を探す（後ろから探す）
     * @param {Document} doc - 対象のドキュメント
     * @param {ArtboardInfo} artboardInfo - 作成時のアートボード情報
     * @returns {number} index（見つからなければ -1）
     */
    function findArtboardIndex(doc, artboardInfo) {
        for (var i = doc.artboards.length - 1; i >= 0; i--) {
            var artboard = doc.artboards[i];
            if (artboard.name === SYMBOL_LIST_NAME && isSameRect(artboard.artboardRect, artboardInfo.rect)) return i;
        }
        return -1;
    }

    /**
     * プレビューで作成したシンボル一覧（アートボードとレイヤー）を削除する
     * @param {Document} doc - 対象のドキュメント
     * @param {SymbolListLayout} layout - 削除する一覧
     * @returns {void}
     */
    function removeLayout(doc, layout) {
        var artboardIndex = findArtboardIndex(doc, layout.artboardInfo);
        if (artboardIndex >= 0 && doc.artboards.length > 1) doc.artboards.remove(artboardIndex);
        layout.layer.remove();
    }

    /**
     * 中心が矩形の内側にあるか（許容幅つき。y は上が正）
     * @param {number[]} itemBounds - アイテムの境界 [left, top, right, bottom]
     * @param {number[]} rect - 矩形 [left, top, right, bottom]
     * @param {number} tolerance - 許容幅（pt）
     * @returns {boolean} 内側なら true
     */
    function isCenterInRect(itemBounds, rect, tolerance) {
        var centerX = (itemBounds[0] + itemBounds[2]) / 2;
        var centerY = (itemBounds[1] + itemBounds[3]) / 2;
        return centerX >= rect[0] - tolerance &&
            centerX <= rect[2] + tolerance &&
            centerY <= rect[1] + tolerance &&
            centerY >= rect[3] - tolerance;
    }

    /**
     * 旧アートボードの矩形に中心が入るページアイテムを全レイヤーから削除する
     * 保護レイヤー上のアイテムは残す（同じ位置に作った新しい一覧を巻き込まないため）
     * @param {Document} doc - 対象のドキュメント
     * @param {number[]} rect - 旧アートボードの矩形
     * @param {Layer} protectedLayer - 削除しないレイヤー
     * @returns {void}
     */
    function removeItemsInRect(doc, rect, protectedLayer) {
        var pageItems = doc.pageItems;
        var itemsToRemove = [];
        for (var i = 0; i < pageItems.length; i++) {
            if (pageItems[i].layer === protectedLayer) continue;
            if (isCenterInRect(pageItems[i].geometricBounds, rect, ARTBOARD_HIT_TOLERANCE_PT)) itemsToRemove.push(pageItems[i]);
        }
        for (var j = 0; j < itemsToRemove.length; j++) {
            /* 親グループごと消えた子やロックされたアイテムは例外になるので飛ばす / Skip children of removed groups and locked items */
            try { itemsToRemove[j].remove(); } catch (e) { }
        }
    }

    /**
     * 確定した一覧を残して、ほかのシンボル一覧（アートボード・その上のアイテム・レイヤー）を削除する
     * @param {Document} doc - 対象のドキュメント
     * @param {SymbolListLayout} keptLayout - 残す一覧
     * @returns {void}
     */
    function removeOtherSymbolLists(doc, keptLayout) {
        var keptIndex = findArtboardIndex(doc, keptLayout.artboardInfo);
        if (keptIndex < 0) keptIndex = keptLayout.artboardInfo.index;

        /* 後ろから消すので、keptIndex より前を消しても比較はずれない / Removing backwards keeps keptIndex valid for comparison */
        for (var i = doc.artboards.length - 1; i >= 0; i--) {
            if (i === keptIndex || doc.artboards.length <= 1 || !isSymbolListName(doc.artboards[i].name)) continue;
            removeItemsInRect(doc, doc.artboards[i].artboardRect, keptLayout.layer);
            doc.artboards.remove(i);
        }

        for (var j = doc.layers.length - 1; j >= 0; j--) {
            var existingLayer = doc.layers[j];
            if (existingLayer === keptLayout.layer || doc.layers.length <= 1 || !isSymbolListName(existingLayer.name)) continue;
            /* ロックされたレイヤーは削除できないので残す / Locked layers cannot be removed */
            try { existingLayer.remove(); } catch (e) { }
        }
    }

    // =========================================
    // 表示 / View
    // =========================================

    /**
     * @typedef {object} ViewState
     * @property {View} view - 対象のビュー
     * @property {number} zoom - ズーム倍率
     * @property {number[]} centerPoint - 表示の中心
     */

    /**
     * 現在のズームと表示の中心を控える
     * @param {Document} doc - 対象のドキュメント
     * @returns {ViewState} 控えた表示状態
     */
    function captureViewState(doc) {
        var activeView = doc.activeView;
        return { view: activeView, zoom: activeView.zoom, centerPoint: activeView.centerPoint };
    }

    /**
     * 控えたズームと表示の中心に戻す
     * @param {ViewState} viewState - 控えた表示状態
     * @returns {void}
     */
    function restoreViewState(viewState) {
        viewState.view.zoom = viewState.zoom;
        viewState.view.centerPoint = viewState.centerPoint;
    }

    /**
     * アートボードをアクティブにし、その中心を指定倍率で表示する
     * @param {Document} doc - 対象のドキュメント
     * @param {number} artboardIndex - アートボードの index
     * @param {number} zoomFactor - ズーム倍率
     * @returns {void}
     */
    function zoomToArtboard(doc, artboardIndex, zoomFactor) {
        doc.artboards.setActiveArtboardIndex(artboardIndex);
        var rect = doc.artboards[artboardIndex].artboardRect; // [left, top, right, bottom]
        var activeView = doc.activeView;
        activeView.zoom = zoomFactor;
        activeView.centerPoint = [(rect[0] + rect[2]) / 2, (rect[1] + rect[3]) / 2];
    }

    /**
     * メニューコマンドでウィンドウに合わせた後、少し縮小して表示する
     * @param {Document} doc - 対象のドキュメント
     * @param {string} menuCommand - "fitin"（アクティブなアートボード）/ "fitall"（すべて）
     * @returns {void}
     */
    function fitWindowThenZoomOut(doc, menuCommand) {
        app.executeMenuCommand(menuCommand);
        /* 最小倍率を下回ると例外になるので、そのときはフィットのまま / Stay fitted if below the minimum zoom */
        try { doc.activeView.zoom = doc.activeView.zoom * FIT_ZOOM_RATIO; } catch (e) { }
        app.redraw();
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 項目名・数値入力欄・単位の行を追加する
     * @param {Panel} parent - 追加先のパネル
     * @param {object} labelSet - 項目名のラベル（ja/en）
     * @param {number} valuePt - 初期値（pt）
     * @param {UnitInfo} unitInfo - 表示単位
     * @param {object} tooltipSet - 項目名と入力欄のツールチップ（ja/en）
     * @returns {EditText} 追加した入力欄（行グループは parent で参照できる）
     */
    function addUnitValueRow(parent, labelSet, valuePt, unitInfo, tooltipSet) {
        var valueRow = addRowGroup(parent);
        addFieldLabel(valueRow, labelSet).helpTip = getLabel(tooltipSet);
        var valueInput = addNumberInput(valueRow, formatUnitValue(valuePt, unitInfo), tooltipSet);
        valueRow.add("statictext", undefined, unitInfo.label);
        return valueInput;
    }

    /**
     * 「作成位置」パネル（基準・間隔・方向）を構築する
     * @param {Panel} parent - 追加先のパネル
     * @param {object} controls - コントロールの参照を書き込むオブジェクト
     * @param {ListSettings} initialSettings - 初期値
     * @param {Document} doc - 対象のドキュメント
     * @returns {void}
     */
    function buildLocationPanel(parent, controls, initialSettings, doc) {
        var locationPanel = addPanel(parent, LABELS.panel.location);

        /* 基準：最終アートボード／指定［番号］（親が異なるので排他はイベント側で処理）
         * Base: last artboard or specified [number] (different parents; exclusivity is handled in events) */
        var baseModeGroup = locationPanel.add("group");
        baseModeGroup.orientation = "column";
        baseModeGroup.alignChildren = "left";
        baseModeGroup.spacing = ROW_SPACING;
        controls.baseLastRadio = addLabeledControl(baseModeGroup, "radiobutton", LABELS.radio.baseLast, LABELS.tooltip.baseLast);
        var specifiedBaseRow = addRowGroup(baseModeGroup);
        controls.baseSpecifiedRadio = addLabeledControl(specifiedBaseRow, "radiobutton", LABELS.radio.baseSpecified, LABELS.tooltip.baseSpecified);
        /* 番号の初期値はキャンバスの最も右下にあるアートボード / Default number: bottom-right-most artboard */
        controls.baseArtboardNumberInput = addNumberInput(specifiedBaseRow, String(findBottomRightArtboardNumber(doc)), LABELS.tooltip.baseArtboardNumber);
        controls.baseLastRadio.value = (initialSettings.baseMode === "last");
        controls.baseSpecifiedRadio.value = (initialSettings.baseMode === "specified");

        controls.artboardGapInput = addUnitValueRow(locationPanel, LABELS.fieldLabel.gap, initialSettings.artboardGap, RULER_UNIT, LABELS.tooltip.artboardGap);

        /* 方向：右側／下側 / Direction: right or below */
        var directionRow = addRowGroup(locationPanel);
        addFieldLabel(directionRow, LABELS.fieldLabel.direction);
        var directionRadioGroup = addRowGroup(directionRow, RADIO_SPACING);
        controls.directionRightRadio = addLabeledControl(directionRadioGroup, "radiobutton", LABELS.radio.directionRight, LABELS.tooltip.directionRight);
        controls.directionBelowRadio = addLabeledControl(directionRadioGroup, "radiobutton", LABELS.radio.directionBelow, LABELS.tooltip.directionBelow);
        controls.directionRightRadio.value = (initialSettings.position === "right");
        controls.directionBelowRadio.value = (initialSettings.position === "below");
    }

    /**
     * 「サイズと余白」パネル（幅・高さ・余白）を構築する
     * @param {Panel} parent - 追加先のパネル
     * @param {object} controls - コントロールの参照を書き込むオブジェクト
     * @param {ListSettings} initialSettings - 初期値
     * @returns {void}
     */
    function buildSizePanel(parent, controls, initialSettings) {
        var sizePanel = addPanel(parent, LABELS.panel.sizeAndPadding);
        controls.widthInput = addUnitValueRow(sizePanel, LABELS.fieldLabel.width, 0, RULER_UNIT, LABELS.tooltip.width);
        controls.heightInput = addUnitValueRow(sizePanel, LABELS.fieldLabel.height, 0, RULER_UNIT, LABELS.tooltip.height);
        controls.marginInput = addUnitValueRow(sizePanel, LABELS.fieldLabel.margin, initialSettings.margin, RULER_UNIT, LABELS.tooltip.margin);
    }

    /**
     * 「背景」パネル（なし／黒／白／グレーを 2 行 2 列）を構築する
     * @param {Panel} parent - 追加先のパネル
     * @param {object} controls - コントロールの参照を書き込むオブジェクト
     * @param {ListSettings} initialSettings - 初期値
     * @returns {void}
     */
    function buildBackgroundPanel(parent, controls, initialSettings) {
        var backgroundPanel = addPanel(parent, LABELS.panel.background);
        var radioRow;
        controls.backgroundRadios = [];
        for (var i = 0; i < BACKGROUND_CHOICES.length; i++) {
            if (i % 2 === 0) radioRow = addRowGroup(backgroundPanel, RADIO_SPACING);
            var choice = BACKGROUND_CHOICES[i];
            var backgroundRadio = addLabeledControl(radioRow, "radiobutton", LABELS.radio.background[choice], LABELS.tooltip.background[choice]);
            backgroundRadio.preferredSize.width = BACKGROUND_RADIO_WIDTH;
            backgroundRadio.value = (choice === initialSettings.bgColor);
            controls.backgroundRadios.push(backgroundRadio);
        }
    }

    /**
     * 「収集対象」パネル（すべて／使用中のみ）を構築する
     * @param {Panel} parent - 追加先のパネル
     * @param {object} controls - コントロールの参照を書き込むオブジェクト
     * @returns {void}
     */
    function buildCollectTargetPanel(parent, controls) {
        var collectTargetPanel = addPanel(parent, LABELS.panel.collectTarget);
        var filterRow = addRowGroup(collectTargetPanel, RADIO_SPACING);
        controls.filterAllRadio = addLabeledControl(filterRow, "radiobutton", LABELS.radio.filterAll, LABELS.tooltip.filterAll);
        controls.filterUsedRadio = addLabeledControl(filterRow, "radiobutton", LABELS.radio.filterUsed, LABELS.tooltip.filterUsed);
        /* 保存値は使わず、起動時は常に「すべて」/ Always start with "all" regardless of saved settings */
        controls.filterAllRadio.value = true;
    }

    /**
     * 「並べ方」パネル（間隔・最大幅）を構築する
     * @param {Panel} parent - 追加先のパネル
     * @param {object} controls - コントロールの参照を書き込むオブジェクト
     * @param {ListSettings} initialSettings - 初期値
     * @returns {void}
     */
    function buildArrangementPanel(parent, controls, initialSettings) {
        var arrangementPanel = addPanel(parent, LABELS.panel.arrangement);
        controls.symbolGapInput = addUnitValueRow(arrangementPanel, LABELS.fieldLabel.gap, initialSettings.symbolGap, RULER_UNIT, LABELS.tooltip.symbolGap);
        controls.maxRowWidthInput = addUnitValueRow(arrangementPanel, LABELS.fieldLabel.maxRowWidth, initialSettings.maxRowWidth, RULER_UNIT, LABELS.tooltip.maxRowWidth);
    }

    /**
     * 「キャプション」パネル（表示・位置・フォントサイズ）を構築する
     * @param {Panel} parent - 追加先のパネル
     * @param {object} controls - コントロールの参照を書き込むオブジェクト
     * @param {ListSettings} initialSettings - 初期値
     * @returns {void}
     */
    function buildCaptionPanel(parent, controls, initialSettings) {
        var captionPanel = addPanel(parent, LABELS.panel.caption);
        controls.showCaptionCheckbox = addLabeledControl(captionPanel, "checkbox", LABELS.checkbox.showCaption, LABELS.tooltip.showCaption);
        controls.showCaptionCheckbox.value = initialSettings.showCaption;

        /* 位置：上／下 / Position: above or below */
        controls.captionPositionRow = addRowGroup(captionPanel);
        controls.captionPositionRow.add("statictext", undefined, labelText(LABELS.fieldLabel.captionPosition));
        var captionPositionRadioGroup = addRowGroup(controls.captionPositionRow, RADIO_SPACING);
        controls.captionAboveRadio = addLabeledControl(captionPositionRadioGroup, "radiobutton", LABELS.radio.captionAbove, LABELS.tooltip.captionAbove);
        controls.captionBelowRadio = addLabeledControl(captionPositionRadioGroup, "radiobutton", LABELS.radio.captionBelow, LABELS.tooltip.captionBelow);
        controls.captionAboveRadio.value = (initialSettings.captionPosition === "above");
        controls.captionBelowRadio.value = (initialSettings.captionPosition === "below");

        /* フォントサイズ：項目名が項目名の幅に収まらないので、入力欄の上に左揃えで置く（単位は環境設定の文字の単位）
         * Font size: the label does not fit the label width, so it sits above the field, left-aligned (type unit preference) */
        controls.fontSizeGroup = captionPanel.add("group");
        controls.fontSizeGroup.orientation = "column";
        controls.fontSizeGroup.alignChildren = "left";
        controls.fontSizeGroup.spacing = ROW_SPACING;
        controls.fontSizeGroup.add("statictext", undefined, labelText(LABELS.fieldLabel.fontSize)).helpTip = getLabel(LABELS.tooltip.fontSize);
        var fontSizeInputRow = addRowGroup(controls.fontSizeGroup);
        controls.fontSizeInput = addNumberInput(fontSizeInputRow, formatUnitValue(initialSettings.fontSize, TYPE_UNIT), LABELS.tooltip.fontSize);
        fontSizeInputRow.add("statictext", undefined, TYPE_UNIT.label);
    }

    /**
     * 最下段のボタン行（左：表示合わせ／右：キャンセル・OK）を構築する
     * @param {Window} dialog - 追加先のダイアログ
     * @param {object} controls - コントロールの参照を書き込むオブジェクト
     * @returns {void}
     */
    function buildButtonRow(dialog, controls) {
        // メイングループ（横並び） / Main group (horizontal layout)
        var btnRowGroup = dialog.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.alignment = ["fill", "bottom"];
        btnRowGroup.spacing = BUTTON_SPACING;

        // 左側グループ / Left-side button group
        var btnLeftGroup = btnRowGroup.add("group");
        btnLeftGroup.alignChildren = ["left", "center"];
        btnLeftGroup.spacing = BUTTON_SPACING;
        controls.btnFitSymbolList = addLabeledControl(btnLeftGroup, "button", LABELS.button.fitSymbolList, LABELS.tooltip.fitSymbolList);
        controls.btnFitAll = addLabeledControl(btnLeftGroup, "button", LABELS.button.fitAll, LABELS.tooltip.fitAll);

        // スペーサー（伸縮）/ Spacer (stretchable)
        var spacer = btnRowGroup.add("group");
        spacer.alignment = ["fill", "fill"];
        spacer.minimumSize.width = 0;

        // 右側グループ / Right-side button group
        var btnRightGroup = btnRowGroup.add("group");
        btnRightGroup.alignChildren = ["right", "center"];
        btnRightGroup.spacing = BUTTON_SPACING;
        btnRightGroup.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });
        btnRightGroup.add("button", undefined, getLabel(LABELS.button.ok), { name: "ok" });
    }

    /**
     * ダイアログを構築し、コントロールの参照を返す
     * @param {Document} doc - 対象のドキュメント
     * @param {ListSettings} initialSettings - 初期値
     * @returns {object} dialog と各コントロールの参照
     */
    function buildDialog(doc, initialSettings) {
        var dialog = new Window("dialog", getLabel(LABELS.dialog.title) + " " + SCRIPT_VERSION);
        setupWindow(dialog);
        var controls = { dialog: dialog };

        /* 2 カラム構成 / Two-column layout */
        var columnsGroup = dialog.add("group");
        columnsGroup.orientation = "row";
        columnsGroup.alignChildren = ["fill", "top"];
        columnsGroup.spacing = COLUMN_SPACING;

        /* 左列：作成するアートボード / Left column: new artboard */
        var artboardPanel = addPanel(addColumnGroup(columnsGroup), LABELS.panel.artboard);
        buildLocationPanel(artboardPanel, controls, initialSettings, doc);
        buildSizePanel(artboardPanel, controls, initialSettings);
        buildBackgroundPanel(artboardPanel, controls, initialSettings);
        /* 保存値は使わず、起動時は常に ON / Always start checked regardless of saved settings */
        controls.updateCheckbox = addLabeledControl(artboardPanel, "checkbox", LABELS.checkbox.update, LABELS.tooltip.update);
        controls.updateCheckbox.value = true;

        /* 右列：収集するシンボル / Right column: symbols */
        var symbolsPanel = addPanel(addColumnGroup(columnsGroup), LABELS.panel.symbols);
        buildCollectTargetPanel(symbolsPanel, controls);
        buildArrangementPanel(symbolsPanel, controls, initialSettings);
        buildCaptionPanel(symbolsPanel, controls, initialSettings);

        buildButtonRow(dialog, controls);
        return controls;
    }

    /**
     * 選択中の背景ラジオの値を返す
     * @param {RadioButton[]} backgroundRadios - 背景ラジオ（BACKGROUND_CHOICES と同じ並び）
     * @returns {string} 背景の選択肢
     */
    function readBackgroundChoice(backgroundRadios) {
        for (var i = 0; i < backgroundRadios.length; i++) {
            if (backgroundRadios[i].value) return BACKGROUND_CHOICES[i];
        }
        return "none";
    }

    /**
     * ダイアログの入力から設定を読む（幅・高さの固定は含まない）
     * @param {object} controls - コントロールの参照
     * @returns {ListSettings} 設定
     */
    function readDialogSettings(controls) {
        return {
            position: controls.directionRightRadio.value ? "right" : "below",
            baseMode: controls.baseSpecifiedRadio.value ? "specified" : "last",
            baseArtboardNumber: parseArtboardNumber(controls.baseArtboardNumberInput.text),
            artboardGap: readUnitValuePt(controls.artboardGapInput, DEFAULT_SETTINGS.artboardGap),
            margin: readUnitValuePt(controls.marginInput, DEFAULT_SETTINGS.margin),
            update: controls.updateCheckbox.value,
            showCaption: controls.showCaptionCheckbox.value,
            filter: controls.filterUsedRadio.value ? "used" : "all",
            symbolGap: readUnitValuePt(controls.symbolGapInput, DEFAULT_SETTINGS.symbolGap),
            maxRowWidth: readUnitValuePt(controls.maxRowWidthInput, DEFAULT_SETTINGS.maxRowWidth),
            bgColor: readBackgroundChoice(controls.backgroundRadios),
            captionPosition: controls.captionBelowRadio.value ? "below" : "above",
            fontSize: readFontSizePt(controls.fontSizeInput)
        };
    }

    // =========================================
    // プレビュー / Preview
    // =========================================

    /**
     * @typedef {object} PreviewSession
     * @property {?SymbolListLayout} currentLayout - 表示中のプレビュー
     * @property {?number} widthOverridePt - 入力で固定した幅（pt、null なら自動）
     * @property {?number} heightOverridePt - 入力で固定した高さ（pt、null なら自動）
     * @property {function(): ListSettings} readSettings - 入力と固定サイズから設定を読む
     * @property {function(): void} clear - プレビューを削除する
     * @property {function(): void} refresh - プレビューを作り直す
     * @property {function(): void} refreshWithAutoSize - 幅・高さの固定を解除して作り直す
     */

    /**
     * 自動サイズのとき、作成したアートボードの幅・高さを入力欄に書き戻す（onChange は発火しない）
     * @param {object} controls - コントロールの参照
     * @param {PreviewSession} previewSession - プレビューの状態
     * @returns {void}
     */
    function writeBackAutoSize(controls, previewSession) {
        var rect = previewSession.currentLayout.artboardInfo.rect;
        if (previewSession.widthOverridePt === null) controls.widthInput.text = formatUnitValue(rect[2] - rect[0], RULER_UNIT);
        if (previewSession.heightOverridePt === null) controls.heightInput.text = formatUnitValue(rect[1] - rect[3], RULER_UNIT);
    }

    /**
     * ダイアログの入力に追従するプレビューを作る（前回のプレビューを消してから作り直す）
     * @param {Document} doc - 対象のドキュメント
     * @param {object} controls - コントロールの参照
     * @returns {PreviewSession} プレビューの状態と操作
     */
    function createPreviewSession(doc, controls) {
        var previewSession = { currentLayout: null, widthOverridePt: null, heightOverridePt: null };

        previewSession.readSettings = function () {
            var settings = readDialogSettings(controls);
            settings.widthOverridePt = previewSession.widthOverridePt;
            settings.heightOverridePt = previewSession.heightOverridePt;
            return settings;
        };
        previewSession.clear = function () {
            if (!previewSession.currentLayout) return;
            removeLayout(doc, previewSession.currentLayout);
            previewSession.currentLayout = null;
        };
        previewSession.refresh = function () {
            previewSession.clear();
            previewSession.currentLayout = buildLayout(doc, previewSession.readSettings());
            if (previewSession.currentLayout) writeBackAutoSize(controls, previewSession);
            app.redraw();
        };
        previewSession.refreshWithAutoSize = function () {
            previewSession.widthOverridePt = null;
            previewSession.heightOverridePt = null;
            previewSession.refresh();
        };
        return previewSession;
    }

    // =========================================
    // イベント / Events
    // =========================================

    /**
     * 複数のコントロールに同じ onClick を設定する
     * @param {object[]} controlList - 対象のコントロール
     * @param {function(): void} handler - クリック時の処理
     * @returns {void}
     */
    function setClickHandler(controlList, handler) {
        for (var i = 0; i < controlList.length; i++) controlList[i].onClick = handler;
    }

    /**
     * 親の異なるラジオボタンを排他にし、選択後の処理を設定する
     * @param {RadioButton[]} radios - 排他にするラジオボタン
     * @param {function(): void} onSelect - 選択後の処理
     * @returns {void}
     */
    function bindExclusiveRadios(radios, onSelect) {
        for (var i = 0; i < radios.length; i++) {
            radios[i].onClick = function () {
                selectExclusiveRadio(radios, this);
                onSelect();
            };
        }
    }

    /**
     * 数値入力欄の確定と↑↓キーに同じ処理を設定する
     * @param {EditText} numberInput - 対象の入力欄
     * @param {function(): void} onValueChange - 値が変わったときの処理
     * @param {number} [minValue] - ↑↓キーでの下限値
     * @returns {void}
     */
    function bindNumberInput(numberInput, onValueChange, minValue) {
        numberInput.onChange = onValueChange;
        changeValueByArrowKey(numberInput, false, onValueChange, minValue);
    }

    /**
     * 「作成位置」パネルのイベントを設定する
     * @param {object} controls - コントロールの参照
     * @param {PreviewSession} previewSession - プレビュー
     * @returns {void}
     */
    function bindLocationEvents(controls, previewSession) {
        var baseNumberInput = controls.baseArtboardNumberInput;

        /* 基準が「指定」のときだけ番号を入力できる / Enable the number only when the base is "specified" */
        function updateBaseNumberEnabled() {
            baseNumberInput.enabled = controls.baseSpecifiedRadio.value;
        }

        bindExclusiveRadios([controls.baseLastRadio, controls.baseSpecifiedRadio], function () {
            updateBaseNumberEnabled();
            previewSession.refreshWithAutoSize();
        });
        setClickHandler([controls.directionRightRadio, controls.directionBelowRadio], previewSession.refreshWithAutoSize);

        /* 番号は 1 未満を 1 に補正してから作り直す / Clamp the number to 1 or more, then rebuild */
        bindNumberInput(baseNumberInput, function () {
            baseNumberInput.text = String(parseArtboardNumber(baseNumberInput.text));
            previewSession.refreshWithAutoSize();
        }, 1);
        updateBaseNumberEnabled();
    }

    /**
     * 幅・高さの入力を設定する（正の値で固定、0 や空欄で自動）
     * 幅を固定したときは［最大幅］も「幅 − 余白 × 2」に合わせる
     * @param {object} controls - コントロールの参照
     * @param {PreviewSession} previewSession - プレビュー
     * @returns {void}
     */
    function bindSizeEvents(controls, previewSession) {
        bindNumberInput(controls.widthInput, function () {
            var widthPt = readUnitValuePt(controls.widthInput, 0);
            previewSession.widthOverridePt = (widthPt > 0) ? widthPt : null;
            var maxRowWidthPt = widthPt - 2 * readUnitValuePt(controls.marginInput, DEFAULT_SETTINGS.margin);
            if (widthPt > 0 && maxRowWidthPt > 0) controls.maxRowWidthInput.text = formatUnitValue(maxRowWidthPt, RULER_UNIT);
            previewSession.refresh();
        });
        bindNumberInput(controls.heightInput, function () {
            var heightPt = readUnitValuePt(controls.heightInput, 0);
            previewSession.heightOverridePt = (heightPt > 0) ? heightPt : null;
            previewSession.refresh();
        });
    }

    /**
     * 「キャプション」パネルのイベントを設定する
     * @param {object} controls - コントロールの参照
     * @param {PreviewSession} previewSession - プレビュー
     * @returns {void}
     */
    function bindCaptionEvents(controls, previewSession) {
        /* 「シンボル名を表示」OFF のときは位置とフォントサイズをディム / Dim position and font size when captions are off */
        function updateCaptionRowsEnabled() {
            var isCaptionShown = controls.showCaptionCheckbox.value;
            controls.captionPositionRow.enabled = isCaptionShown;
            controls.fontSizeGroup.enabled = isCaptionShown;
        }

        controls.showCaptionCheckbox.onClick = function () {
            updateCaptionRowsEnabled();
            previewSession.refreshWithAutoSize();
        };
        setClickHandler([controls.captionAboveRadio, controls.captionBelowRadio], previewSession.refreshWithAutoSize);
        updateCaptionRowsEnabled();
    }

    /**
     * 表示合わせボタンのイベントを設定する
     * @param {Document} doc - 対象のドキュメント
     * @param {object} controls - コントロールの参照
     * @param {PreviewSession} previewSession - プレビュー
     * @returns {void}
     */
    function bindViewButtons(doc, controls, previewSession) {
        controls.btnFitSymbolList.onClick = function () {
            if (!previewSession.currentLayout) return;
            doc.artboards.setActiveArtboardIndex(previewSession.currentLayout.artboardInfo.index);
            fitWindowThenZoomOut(doc, "fitin");
        };
        controls.btnFitAll.onClick = function () {
            fitWindowThenZoomOut(doc, "fitall");
        };
    }

    /**
     * ダイアログのすべてのイベントを設定する
     * @param {Document} doc - 対象のドキュメント
     * @param {object} controls - コントロールの参照
     * @param {PreviewSession} previewSession - プレビュー
     * @returns {void}
     */
    function bindDialogEvents(doc, controls, previewSession) {
        bindLocationEvents(controls, previewSession);
        bindSizeEvents(controls, previewSession);
        bindCaptionEvents(controls, previewSession);
        bindViewButtons(doc, controls, previewSession);

        /* 寸法と関係のない操作は幅・高さの固定を解除して作り直す / Other changes release the fixed size and rebuild */
        bindExclusiveRadios(controls.backgroundRadios, previewSession.refreshWithAutoSize);
        setClickHandler([controls.updateCheckbox, controls.filterAllRadio, controls.filterUsedRadio], previewSession.refreshWithAutoSize);
        var numberInputs = [
            controls.artboardGapInput,
            controls.marginInput,
            controls.symbolGapInput,
            controls.maxRowWidthInput,
            controls.fontSizeInput
        ];
        for (var i = 0; i < numberInputs.length; i++) {
            bindNumberInput(numberInputs[i], previewSession.refreshWithAutoSize);
        }
    }

    // =========================================
    // メイン / Main
    // =========================================

    /**
     * OK 時にプレビューを消して一覧を作り直し、更新 ON なら既存の一覧を削除して設定を保存する
     * @param {Document} doc - 対象のドキュメント
     * @param {PreviewSession} previewSession - プレビュー
     * @returns {void}
     */
    function commitLayout(doc, previewSession) {
        var finalSettings = previewSession.readSettings();
        previewSession.clear();
        var finalLayout = buildLayout(doc, finalSettings);
        if (!finalLayout) {
            alert(getLabel(LABELS.alert.noUsedSymbols));
        } else {
            if (finalSettings.update) removeOtherSymbolLists(doc, finalLayout);
            saveSettings(finalSettings);
        }
        app.redraw();
    }

    /**
     * ダイアログを表示し、OK なら確定、キャンセルならプレビューを消して表示を戻す
     * @param {Document} doc - 対象のドキュメント
     * @returns {void}
     */
    function showDialog(doc) {
        var viewState = captureViewState(doc);
        var controls = buildDialog(doc, loadSettings());
        var previewSession = createPreviewSession(doc, controls);
        bindDialogEvents(doc, controls, previewSession);

        /* 起動直後にプレビューを作り、新しいアートボードの中心を表示 / Initial preview centered on the new artboard */
        previewSession.refresh();
        if (previewSession.currentLayout) {
            zoomToArtboard(doc, previewSession.currentLayout.artboardInfo.index, INITIAL_PREVIEW_ZOOM);
            app.redraw();
        }

        if (controls.dialog.show() === 1) {
            commitLayout(doc, previewSession);
            return;
        }
        previewSession.clear();
        restoreViewState(viewState);
        app.redraw();
    }

    /**
     * ドキュメントとシンボルの有無を確認してダイアログを開く
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel(LABELS.alert.noDocument));
            return;
        }
        var doc = app.activeDocument;
        if (doc.symbols.length === 0) {
            alert(getLabel(LABELS.alert.noSymbols));
            return;
        }
        showDialog(doc);
    }

    main();

})();
