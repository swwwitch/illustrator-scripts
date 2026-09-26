#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

アクティブなドキュメントで使われている書式（フォント・サイズ・行送り・字間・行揃え・段落前後のアキ）を段落ごとに調べて一覧にし、
一覧で選んだ書式を編集・置換すると、その書式の段落をまとめて更新します（段落スタイルを使わない再定義）。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/FontPresetCollector.md

### Overview

Lists every format (font, size, leading, letter spacing, justification and paragraph spacing) used in the active document,
paragraph by paragraph. Edit or replace a format and every paragraph in that format is updated at once,
much like redefining a paragraph style without using styles.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/FontPresetCollector.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "FontPresetCollector";          /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-09-26";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-26";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/FontPresetCollector.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/FontPresetCollector.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // レイアウト / Layout
    // =========================================
    var WINDOW_MARGINS = 16;                  /* ウィンドウ外周の余白 / Window margins */
    var WINDOW_SPACING = 12;                  /* ウィンドウ内の要素間隔 / Window spacing */
    var ROW_SPACING = 8;                      /* 行グループ内の間隔 / Spacing inside a row group */
    var PANEL_MARGINS = [12, 16, 12, 12];     /* パネル内の余白 / Margins inside a panel */
    var PANEL_ROW_SPACING = 6;                /* パネル内の行間隔 / Spacing between rows inside a panel */
    var BUTTON_ROW_TOP_MARGIN = 5;            /* ボタン行の上余白 / Top margin above the button row */
    var LIST_WIDTH = 340;                     /* 書式一覧の幅 / Width of the format list */
    var LIST_ROW_HEIGHT = 22;                 /* 一覧1行の高さの目安 / Estimated height of one list row */
    var LIST_ROW_COUNT = 14;                  /* 一覧に出す行数 / Rows shown in the list */
    var LABEL_WIDTH = 120;                    /* 項目名の幅 / Width of the row labels */
    var NUMBER_FIELD_CHARS = 6;               /* 数値欄の文字数 / Characters in a number field */
    var DROPDOWN_WIDTH = 220;                 /* ドロップダウンの幅 / Width of the dropdowns */

    /**
     * 行グループの向き・揃え・間隔をまとめて設定する
     * @param {Group} targetGroup - 対象の行グループ
     * @param {string} alignment - 横方向の揃え（"left" / "right" / "fill"）
     * @param {number} spacing - 要素間の間隔（省略時は ROW_SPACING）
     * @returns {void}
     */
    function setupRow(targetGroup, alignment, spacing) {
        targetGroup.orientation = "row";
        targetGroup.alignment = [alignment || "left", "center"];
        targetGroup.alignChildren = ["left", "center"];
        targetGroup.spacing = (typeof spacing === "number") ? spacing : ROW_SPACING;
    }

    /**
     * パネルの向き・揃え・余白・間隔をまとめて設定する
     * @param {Panel} targetPanel - 対象のパネル
     * @returns {void}
     */
    function setupPanel(targetPanel) {
        targetPanel.orientation = "column";
        targetPanel.alignChildren = ["left", "top"];
        targetPanel.margins = PANEL_MARGINS;
        targetPanel.spacing = PANEL_ROW_SPACING;
    }

    /**
     * 数値欄に ↑↓ キーでの増減を付ける（shift で ±10、option で ±0.1）
     * @param {EditText} editText - 対象の数値欄
     * @returns {void}
     */
    function changeValueByArrowKey(editText) {
        editText.addEventListener("keydown", function(event) {
            if (event.keyName != 'Up' && event.keyName != 'Down') return;
            var value = Number(editText.text);
            if (isNaN(value)) return;

            var keyboard = ScriptUI.environment.keyboardState;
            var delta = 1;

            if (keyboard.shiftKey) {
                delta = 10;
                /* Shiftキー押下時は10の倍数にスナップ / Snap to multiples of 10 when Shift is held */
                if (event.keyName === "Up") {
                    value = Math.ceil((value + 1) / delta) * delta;
                    event.preventDefault();
                } else if (event.keyName === "Down") {
                    value = Math.floor((value - 1) / delta) * delta;
                    if (value < 0) value = 0;
                    event.preventDefault();
                }
            } else if (keyboard.altKey) {
                delta = 0.1;
                /* Optionキー押下時は0.1単位で増減 / Step by 0.1 when Option is held */
                if (event.keyName === "Up") {
                    value += delta;
                    event.preventDefault();
                } else if (event.keyName === "Down") {
                    value -= delta;
                    event.preventDefault();
                }
            } else {
                delta = 1;
                if (event.keyName === "Up") {
                    value += delta;
                    event.preventDefault();
                } else if (event.keyName === "Down") {
                    value -= delta;
                    if (value < 0) value = 0;
                    event.preventDefault();
                }
            }

            if (keyboard.altKey) {
                /* 小数第1位までに丸め / Round to one decimal place */
                value = Math.round(value * 10) / 10;
            } else {
                /* 整数に丸め / Round to an integer */
                value = Math.round(value);
            }

            editText.text = value;
        });
    }

    // =========================================
    // 単位 / Units
    // =========================================

    /* 単位コードに対応する表示ラベルと、1単位あたりのポイント数
       Unit code -> display label and points per unit */
    var UNITS = [
        { label: "in",    pointsPerUnit: 72 },                /* 0 */
        { label: "mm",    pointsPerUnit: 72 / 25.4 },         /* 1 */
        { label: "pt",    pointsPerUnit: 1 },                 /* 2 */
        { label: "pica",  pointsPerUnit: 12 },                /* 3 */
        { label: "cm",    pointsPerUnit: 72 / 2.54 },         /* 4 */
        { label: "Q",     pointsPerUnit: 72 / 25.4 * 0.25 },  /* 5 */
        { label: "px",    pointsPerUnit: 1 },                 /* 6 */
        { label: "ft/in", pointsPerUnit: 72 * 12 },           /* 7 */
        { label: "m",     pointsPerUnit: 72 / 25.4 * 1000 },  /* 8 */
        { label: "yd",    pointsPerUnit: 72 * 36 },           /* 9 */
        { label: "ft",    pointsPerUnit: 72 * 12 }            /* 10 */
    ];

    /* 単位コード5を「歯（H）」と表示する環境設定キー。文字サイズ（text/units）だけ「級（Q）」
       Preference keys that show unit code 5 as H; only the type size (text/units) shows Q */
    var HA_UNIT_PREF_KEYS = { "rulerType": true, "strokeUnits": true, "text/asianunits": true };

    /**
     * 環境設定キーの単位を返す
     * @param {string} [prefKey] - "rulerType"（既定）/ "strokeUnits" / "text/units" / "text/asianunits"
     * @returns {{code: number, label: string, pointsPerUnit: number}} 単位の情報
     */
    function getUnitInfo(prefKey) {
        var unitKey = prefKey || "rulerType";
        var unitCode = app.preferences.getIntegerPreference(unitKey);
        /* 未知のコードは pt に寄せる / unknown codes fall back to points */
        var unit = UNITS[unitCode] || UNITS[2];
        /* 級（Q）と歯（H）は同じ長さだが、文字サイズは「Q」、距離は「H」と呼び分ける */
        var label = (unitCode === 5 && HA_UNIT_PREF_KEYS[unitKey]) ? "H" : unit.label;
        return { code: unitCode, label: label, pointsPerUnit: unit.pointsPerUnit };
    }

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * UI 言語を判定する
     * @returns {string} "ja" または "en"
     */
    function getUiLanguage() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = getUiLanguage();

    /* ラベル定義 / Label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "段落書式の一括変更", en: "Batch Paragraph Formatting" }
        },
        panel: {
            edit: { ja: "書式の編集", en: "Edit Format" },
            editPending: { ja: "書式の編集（未反映の変更あり）", en: "Edit Format (unapplied changes)" }
        },
        fieldLabel: {
            font: { ja: "フォント", en: "Font" },
            size: { ja: "フォントサイズ", en: "Font Size" },
            leading: { ja: "行送り", en: "Leading" },
            tracking: { ja: "トラッキング", en: "Tracking" },
            tsume: { ja: "文字ツメ", en: "Tsume" },
            kern: { ja: "カーニング", en: "Kerning" },
            justify: { ja: "行揃え", en: "Justification" },
            spaceBefore: { ja: "段落前のアキ", en: "Space Before" },
            spaceAfter: { ja: "段落後のアキ", en: "Space After" }
        },
        checkbox: {
            autoLeading: { ja: "自動", en: "Auto" }
        },
        dropdown: {
            noReplaceSource: { ja: "—", en: "—" },
            kernMetrics: { ja: "メトリクス", en: "Metrics" },
            kernOptical: { ja: "オプティカル", en: "Optical" },
            kernRomanOnly: { ja: "和文等幅", en: "Metrics - Roman Only" },
            kernNone: { ja: "0", en: "0" },
            justifyLeft: { ja: "左揃え", en: "Align Left" },
            justifyCenter: { ja: "中央揃え", en: "Align Center" },
            justifyRight: { ja: "右揃え", en: "Align Right" },
            justifyLastLeft: { ja: "均等配置（最終行左揃え）", en: "Justify with Last Line Aligned Left" },
            justifyLastCenter: { ja: "均等配置（最終行中央揃え）", en: "Justify with Last Line Aligned Center" },
            justifyLastRight: { ja: "均等配置（最終行右揃え）", en: "Justify with Last Line Aligned Right" },
            justifyAllLines: { ja: "両端揃え", en: "Justify All Lines" }
        },
        button: {
            replace: { ja: "置換", en: "Replace" },
            update: { ja: "更新", en: "Update" },
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" }
        },
        tooltip: {
            formatList: {
                ja: "アクティブなドキュメントで使われている書式（かっこ内は段落数）。1つ選ぶと右側で値を編集できます。" +
                    "shift／command（Ctrl）＋クリックで複数選ぶと、まとめて置換できます。",
                en: "Formats used in the active document (paragraph count in brackets). Choose one to edit its values on the right; " +
                    "Shift/Command (Ctrl)-click to pick several and replace them together."
            },
            replace: {
                ja: "一覧で選んでいる書式（複数可）の段落を、ポップアップで選んだ書式にすぐに置き換えます。",
                en: "Replace the paragraphs in the formats chosen in the list (one or more) with the one picked in the popup, right away."
            },
            font: { ja: "ドキュメントで使われているフォントから選びます。", en: "Choose from the fonts used in the document." },
            tracking: { ja: "単位は 1/1000 em です。", en: "In 1/1000 em." },
            autoLeading: {
                ja: "オンにすると、行送りはフォントサイズに合わせて自動で決まります（自動行送りの割合はそのまま）。",
                en: "When on, leading follows the font size automatically (the auto-leading percentage is kept)."
            },
            update: { ja: "欄で編集した値を、その書式の段落へすぐに反映します。", en: "Apply the edited values to the paragraphs in this format right away." },
            ok: {
                ja: "変更を確定して閉じます（［更新］していない編集もここで反映します）。",
                en: "Keep the changes and close (edits not yet applied with Update are applied now)."
            },
            cancel: { ja: "反映した変更をすべて元に戻して閉じます。", en: "Undo every change that was applied and close." }
        },
        alert: {
            noDocument: { ja: "ドキュメントを開いてください。", en: "Open a document first." },
            noText: { ja: "テキストが見つかりませんでした。", en: "No text was found." }
        }
    };

    /* カーニングの選択肢（並び順＝表示順）/ Kerning choices, in display order */
    var KERN_OPTIONS = [
        { id: "metrics", label: LABELS.dropdown.kernMetrics },
        { id: "optical", label: LABELS.dropdown.kernOptical },
        { id: "mono", label: LABELS.dropdown.kernRomanOnly },
        { id: "zero", label: LABELS.dropdown.kernNone }
    ];

    /* 行揃えの選択肢（並び順＝表示順）/ Justification choices, in display order */
    var JUSTIFY_OPTIONS = [
        { id: "left", label: LABELS.dropdown.justifyLeft },
        { id: "center", label: LABELS.dropdown.justifyCenter },
        { id: "right", label: LABELS.dropdown.justifyRight },
        { id: "justifyLeft", label: LABELS.dropdown.justifyLastLeft },
        { id: "justifyCenter", label: LABELS.dropdown.justifyLastCenter },
        { id: "justifyRight", label: LABELS.dropdown.justifyLastRight },
        { id: "justifyAll", label: LABELS.dropdown.justifyAllLines }
    ];

    /**
     * 言語に応じたラベル文字列を取得する
     * @param {Object} labelSet - { ja, en } 形式のラベル定義
     * @returns {string} 現在の UI 言語のラベル
     */
    function getLabel(labelSet) {
        if (!labelSet) return "";
        return labelSet[uiLang] || labelSet.ja || labelSet.en || "";
    }

    /**
     * コロン付きの項目名を返す（日本語は全角、英語は半角）
     * @param {Object} labelSet - { ja, en } 形式のラベル定義
     * @returns {string} コロンを付けた項目名
     */
    function labelText(labelSet) {
        return getLabel(labelSet) + (uiLang === "ja" ? "：" : ":");
    }

    // =========================================
    // 書式の読み書き / Reading and writing formats
    // =========================================

    /* 書式が持つ項目。一覧のまとまりと、更新時の比較はこの項目で行う
       The settings a format holds; grouping and the update diff both use these */
    var FORMAT_KEYS = [
        "psName", "size", "autoLeading", "leading", "tracking", "tsume",
        "kern", "justify", "spaceBefore", "spaceAfter"
    ];

    /**
     * 小数第2位までに丸める（10.5pt のような端数は残す）
     * @param {number} value - 丸めたい値
     * @returns {number} 丸めた値
     */
    function roundToHundredths(value) {
        return Math.round(value * 100) / 100;
    }

    /**
     * AutoKernType を方式 ID へ変換する
     * @param {AutoKernType} kernValue - 読み取ったカーニング
     * @returns {string} "metrics" / "optical" / "mono" / "zero"。判定できなければ ""
     */
    function kernMethodToId(kernValue) {
        var kernText = String(kernValue);
        if (kernText === String(AutoKernType.AUTO)) return "metrics";
        if (kernText === String(AutoKernType.OPTICAL)) return "optical";
        if (kernText === String(AutoKernType.METRICSROMANONLY)) return "mono";
        if (kernText === String(AutoKernType.NOAUTOKERN)) return "zero";
        return "";
    }

    /**
     * 方式 ID を AutoKernType へ変換する
     * @param {string} methodId - "metrics" / "optical" / "mono" / "zero"
     * @returns {AutoKernType} カーニング
     */
    function resolveAutoKernType(methodId) {
        if (methodId === "mono") return AutoKernType.METRICSROMANONLY;
        if (methodId === "metrics") return AutoKernType.AUTO;
        if (methodId === "optical") return AutoKernType.OPTICAL;
        return AutoKernType.NOAUTOKERN;
    }

    /**
     * Justification を ID 文字列へ変換する
     * @param {Justification} justifyValue - 読み取った行揃え
     * @returns {string} 行揃えの ID。判定できなければ ""
     */
    function justificationToId(justifyValue) {
        var justifyText = String(justifyValue);
        if (justifyText === String(Justification.LEFT)) return "left";
        if (justifyText === String(Justification.CENTER)) return "center";
        if (justifyText === String(Justification.RIGHT)) return "right";
        if (justifyText === String(Justification.FULLJUSTIFYLASTLINELEFT)) return "justifyLeft";
        if (justifyText === String(Justification.FULLJUSTIFYLASTLINECENTER)) return "justifyCenter";
        if (justifyText === String(Justification.FULLJUSTIFYLASTLINERIGHT)) return "justifyRight";
        if (justifyText === String(Justification.FULLJUSTIFY)) return "justifyAll";
        return "";
    }

    /**
     * 行揃え ID を Justification へ変換する
     * @param {string} justifyId - 行揃えの ID
     * @returns {Justification} 行揃え
     */
    function resolveJustification(justifyId) {
        if (justifyId === "center") return Justification.CENTER;
        if (justifyId === "right") return Justification.RIGHT;
        if (justifyId === "justifyLeft") return Justification.FULLJUSTIFYLASTLINELEFT;
        if (justifyId === "justifyCenter") return Justification.FULLJUSTIFYLASTLINECENTER;
        if (justifyId === "justifyRight") return Justification.FULLJUSTIFYLASTLINERIGHT;
        if (justifyId === "justifyAll") return Justification.FULLJUSTIFY;
        return Justification.LEFT;
    }

    /**
     * PostScript 名からフォントを探す
     * @param {string} psName - フォントの PostScript 名
     * @returns {TextFont|null} フォント。見つからなければ null
     */
    function findFontByName(psName) {
        /* getByName は見つからないと例外を投げる / getByName throws when there is no match */
        try { return app.textFonts.getByName(psName); } catch (e) { return null; }
    }

    /**
     * 段落の書式を読む。文字属性は先頭の文字、段落属性は段落から読む
     * 行送りは、自動なら自動行送り量から算出した値、固定ならその値（pt）
     * @param {TextRange} paragraph - 対象の段落（空でないこと）
     * @returns {Object|null} 書式。フォントが読めなければ null
     */
    function readParagraphFormat(paragraph) {
        var charAttrs = paragraph.characters[0].characterAttributes;
        var paragraphAttrs = paragraph.paragraphAttributes;
        var currentFont = charAttrs.textFont;
        if (!currentFont || !currentFont.name) return null;
        var fontSize = roundToHundredths(charAttrs.size);
        var isAutoLeading = (charAttrs.autoLeading === true);
        var autoLeadingAmount = roundToHundredths(paragraphAttrs.autoLeadingAmount);
        return {
            psName: currentFont.name,
            size: fontSize,
            autoLeading: isAutoLeading,
            autoLeadingAmount: autoLeadingAmount,
            leading: isAutoLeading ? roundToHundredths(fontSize * autoLeadingAmount / 100) : roundToHundredths(charAttrs.leading),
            tracking: Math.round(charAttrs.tracking),
            tsume: Math.round(charAttrs.Tsume),
            kern: kernMethodToId(charAttrs.kerningMethod),
            justify: justificationToId(paragraphAttrs.justification),
            spaceBefore: roundToHundredths(paragraphAttrs.spaceBefore),
            spaceAfter: roundToHundredths(paragraphAttrs.spaceAfter)
        };
    }

    /**
     * 書式を、一覧のまとまりを決めるキー文字列にする
     * 自動行送りは算出値ではなく自動行送り量で比べる（サイズが同じなら結果も同じ）
     * @param {Object} format - 書式
     * @returns {string} キー文字列
     */
    function buildFormatKey(format) {
        var keyParts = [];
        for (var i = 0; i < FORMAT_KEYS.length; i++) {
            var keyName = FORMAT_KEYS[i];
            keyParts.push((keyName === "leading" && format.autoLeading) ? "auto" + format.autoLeadingAmount : String(format[keyName]));
        }
        return keyParts.join("\t");
    }

    /**
     * 書式の複製を返す
     * @param {Object} format - 複製したい書式
     * @returns {Object} 複製
     */
    function copyFormat(format) {
        var copied = {};
        for (var keyName in format) {
            if (format.hasOwnProperty(keyName)) copied[keyName] = format[keyName];
        }
        return copied;
    }

    /**
     * 編集の出発点になる書式を返す。行送りは既定で「自動」にする
     * @param {Object} original - 編集前の書式
     * @returns {Object} 編集用の書式
     */
    function createEditedFormat(original) {
        var edited = copyFormat(original);
        edited.autoLeading = true;
        edited.leading = roundToHundredths(edited.size * edited.autoLeadingAmount / 100);
        return edited;
    }

    /**
     * 2つの書式で値が違う項目の名前を返す
     * 行送り（自動・値）は、ほかの項目が違うときか、includeLeading が true のときだけ数える
     * @param {Object} fromFormat - 比べる元の書式
     * @param {Object} toFormat - 比べる先の書式
     * @param {boolean} includeLeading - 行送りだけの違いも数えるなら true
     * @returns {string[]} 違う項目の名前
     */
    function listDifferingKeys(fromFormat, toFormat, includeLeading) {
        var differingKeys = [], leadingKeys = [];
        for (var i = 0; i < FORMAT_KEYS.length; i++) {
            var keyName = FORMAT_KEYS[i];
            if (fromFormat[keyName] === toFormat[keyName]) continue;
            if (keyName === "autoLeading" || keyName === "leading") leadingKeys.push(keyName);
            else differingKeys.push(keyName);
        }
        /* 自動行送り量はキーの一部なので、自動どうしで量だけ違うときも行送りの違いとして数える
           The auto leading amount is part of the grouping key, so an amount-only difference counts as a leading change */
        if (toFormat.autoLeading && fromFormat.autoLeadingAmount !== toFormat.autoLeadingAmount) leadingKeys.push("autoLeadingAmount");
        if (differingKeys.length || includeLeading) differingKeys = differingKeys.concat(leadingKeys);
        return differingKeys;
    }

    /**
     * 段落へ、指定の項目だけを書き込む
     * 行送りは「自動」ならサイズに追従させ、固定なら値を入れる
     * @param {TextRange} paragraph - 書き込む段落
     * @param {Object} toFormat - 書き込む書式
     * @param {string[]} keyNames - 書き込む項目の名前
     * @param {TextFont} targetFont - フォントを変えるときに当てるフォント（見つからなければ null）
     * @returns {void}
     */
    function writeParagraphFormat(paragraph, toFormat, keyNames, targetFont) {
        var changedFlags = {};
        for (var i = 0; i < keyNames.length; i++) changedFlags[keyNames[i]] = true;
        var charAttrs = paragraph.characterAttributes;
        var paragraphAttrs = paragraph.paragraphAttributes;
        if (changedFlags.psName && targetFont) charAttrs.textFont = targetFont;
        if (changedFlags.size) charAttrs.size = toFormat.size;
        if (changedFlags.autoLeading || changedFlags.leading || changedFlags.autoLeadingAmount) {
            charAttrs.autoLeading = toFormat.autoLeading;
            if (!toFormat.autoLeading) charAttrs.leading = toFormat.leading;
        }
        if (changedFlags.autoLeadingAmount) paragraphAttrs.autoLeadingAmount = toFormat.autoLeadingAmount;
        if (changedFlags.tracking) charAttrs.tracking = toFormat.tracking;
        if (changedFlags.tsume) charAttrs.Tsume = toFormat.tsume;
        if (changedFlags.kern) {
            charAttrs.kerningMethod = resolveAutoKernType(toFormat.kern);
            /* メトリクスのときだけプロポーショナルメトリクスを ON / Proportional metrics ON only for Metrics */
            charAttrs.proportionalMetrics = (toFormat.kern === "metrics");
        }
        if (changedFlags.justify) paragraphAttrs.justification = resolveJustification(toFormat.justify);
        if (changedFlags.spaceBefore) paragraphAttrs.spaceBefore = toFormat.spaceBefore;
        if (changedFlags.spaceAfter) paragraphAttrs.spaceAfter = toFormat.spaceAfter;
    }

    /**
     * 段落へ書式を書き込む。書き込めない段落（ロック中のレイヤーなど）は飛ばす
     * @param {TextRange} paragraph - 書き込む段落
     * @param {Object} toFormat - 書き込む書式
     * @param {string[]} keyNames - 書き込む項目の名前
     * @param {TextFont} targetFont - フォントを変えるときに当てるフォント
     * @returns {void}
     */
    function tryWriteParagraphFormat(paragraph, toFormat, keyNames, targetFont) {
        try { writeParagraphFormat(paragraph, toFormat, keyNames, targetFont); } catch (e) { }
    }

    /**
     * 編集後の書式を、カンバス上の段落へすぐに反映する（反映済みの状態から変わった項目だけ）
     * 再描画は呼び出し側でまとめて行う
     * @param {Object} formatGroup - 書式のまとまり（applied / edited / leadingTouched / paragraphs）
     * @returns {void}
     */
    function applyEditedToCanvas(formatGroup) {
        var keyNames = listDifferingKeys(formatGroup.applied, formatGroup.edited, formatGroup.leadingTouched);
        if (!keyNames.length) return;
        var targetFont = findFontByName(formatGroup.edited.psName);
        for (var i = 0; i < formatGroup.paragraphs.length; i++) {
            tryWriteParagraphFormat(formatGroup.paragraphs[i], formatGroup.edited, keyNames, targetFont);
        }
        formatGroup.applied = copyFormat(formatGroup.edited);
    }

    /**
     * ダイアログを開いたときの書式へ、段落を戻す（今の書式を読み、違う項目だけ書き戻す）
     * 置換のたびに一覧を作り直すので、戻すときは開いたときの控えを使う
     * @param {Object[]} initialGroups - 開いたときに調べた書式のまとまり
     * @returns {void}
     */
    function restoreInitialFormats(initialGroups) {
        for (var i = 0; i < initialGroups.length; i++) {
            var original = initialGroups[i].original;
            var targetFont = findFontByName(original.psName);
            for (var j = 0; j < initialGroups[i].paragraphs.length; j++) {
                var paragraph = initialGroups[i].paragraphs[j];
                var keyNames = listDifferingKeys(readParagraphFormat(paragraph), original, true);
                if (keyNames.length) tryWriteParagraphFormat(paragraph, original, keyNames, targetFont);
            }
        }
    }

    // =========================================
    // ドキュメントの調査 / Surveying the document
    // =========================================

    /**
     * 全テキストフレームを段落ごとに調べ、書式ごとのまとまりにする（使われた文字数の多い順）
     * 空の段落は飛ばす
     * @param {Document} doc - 対象のドキュメント
     * @returns {Object[]} 書式のまとまり（original / applied（カンバス上の状態）/ edited / leadingTouched / paragraphs / chars）
     */
    function collectFormatGroups(doc) {
        var groupsByKey = {}, formatGroups = [];
        var textFrames = doc.textFrames;
        for (var i = 0; i < textFrames.length; i++) {
            var paragraphs = textFrames[i].paragraphs;
            for (var j = 0; j < paragraphs.length; j++) {
                var paragraph = paragraphs[j];
                var charCount = paragraph.characters.length;
                if (!charCount) continue;
                var paragraphFormat = readParagraphFormat(paragraph);
                if (!paragraphFormat) continue;
                var groupKey = "k" + buildFormatKey(paragraphFormat);
                if (!groupsByKey.hasOwnProperty(groupKey)) {
                    groupsByKey[groupKey] = {
                        original: paragraphFormat,
                        applied: copyFormat(paragraphFormat),
                        edited: createEditedFormat(paragraphFormat),
                        leadingTouched: false,
                        paragraphs: [],
                        chars: 0
                    };
                    formatGroups.push(groupsByKey[groupKey]);
                }
                groupsByKey[groupKey].paragraphs.push(paragraph);
                groupsByKey[groupKey].chars += charCount;
            }
        }
        formatGroups.sort(function (a, b) { return b.chars - a.chars; });
        return formatGroups;
    }

    /**
     * カーソルだけの選択（長さ0）から、そのカーソルがある段落をストーリーの段落から探す
     * @param {TextRange} insertionPoint - カーソル位置の選択
     * @returns {TextRange|null} 段落。見つからなければ null
     */
    function paragraphAtInsertionPoint(insertionPoint) {
        var insertionOffset = insertionPoint.characterOffset;
        var storyParagraphs = insertionPoint.story.paragraphs;
        for (var i = 0; i < storyParagraphs.length; i++) {
            var paragraphStart = storyParagraphs[i].characterOffset;
            if (insertionOffset >= paragraphStart && insertionOffset <= paragraphStart + storyParagraphs[i].length) return storyParagraphs[i];
        }
        return null;
    }

    /**
     * 起動時に選択していた段落を返す
     * 文字の選択ならその最初の段落、オブジェクトの選択は段落が1つのテキストフレーム1つに限る
     * @param {Document} doc - 対象のドキュメント
     * @returns {TextRange|null} 段落。決められなければ null
     */
    function findSelectedParagraph(doc) {
        var currentSelection = doc.selection;
        if (!currentSelection) return null;
        if (currentSelection.typename === "TextRange") {
            var touchedParagraphs = currentSelection.paragraphs;
            return touchedParagraphs.length ? touchedParagraphs[0] : paragraphAtInsertionPoint(currentSelection);
        }
        if (currentSelection.length !== 1 || currentSelection[0].typename !== "TextFrame") return null;
        var frameParagraphs = currentSelection[0].paragraphs;
        return (frameParagraphs.length === 1) ? frameParagraphs[0] : null;
    }

    /**
     * 起動時に選択していた段落の書式が、一覧の何番目かを返す
     * @param {Document} doc - 対象のドキュメント
     * @param {Object[]} formatGroups - 書式のまとまり
     * @returns {number} 一覧の位置。見つからなければ 0（先頭）
     */
    function findInitialIndex(doc, formatGroups) {
        var selectedParagraph = findSelectedParagraph(doc);
        if (!selectedParagraph || !selectedParagraph.characters.length) return 0;
        var selectedFormat = readParagraphFormat(selectedParagraph);
        if (!selectedFormat) return 0;
        var selectedKey = buildFormatKey(selectedFormat);
        for (var i = 0; i < formatGroups.length; i++) {
            if (buildFormatKey(formatGroups[i].original) === selectedKey) return i;
        }
        return 0;
    }

    // =========================================
    // 一覧の表示 / Presenting formats
    // =========================================

    /* 文字サイズ・行送り・段落のアキに使う単位 / Unit for type size, leading and paragraph spacing */
    var textUnit = null;

    /* psName → 表示名（ファミリー＋スタイル）/ psName → display name (family + style) */
    var fontDisplayNames = {};

    /**
     * PostScript 名の表示名を返す（見つからないフォントは PostScript 名のまま）
     * @param {string} psName - フォントの PostScript 名
     * @returns {string} 表示名
     */
    function fontDisplayName(psName) {
        if (!fontDisplayNames.hasOwnProperty(psName)) {
            var font = findFontByName(psName);
            fontDisplayNames[psName] = font ? font.family + (font.style ? " " + font.style : "") : psName;
        }
        return fontDisplayNames[psName];
    }

    /**
     * pt の値を文字の単位で表示する文字列にする
     * @param {number} points - 値（pt）
     * @returns {string} 単位の数値（小数第2位まで）
     */
    function pointsToUnitText(points) {
        return String(roundToHundredths(points / textUnit.pointsPerUnit));
    }

    /**
     * 書式のまとまりの要約を返す（一覧とポップアップの行）
     * @param {Object} formatGroup - 書式のまとまり
     * @returns {string} 「フォント名  サイズ / 行送り  (段落数)」の文字列
     */
    function buildFormatSummary(formatGroup) {
        var format = formatGroup.original;
        return fontDisplayName(format.psName) + "  " +
            pointsToUnitText(format.size) + " / " + pointsToUnitText(format.leading) + " " + textUnit.label +
            "  (" + formatGroup.paragraphs.length + ")";
    }

    /**
     * フォントの選択肢を返す（ドキュメントで使われているフォント、重複なし）
     * @param {Object[]} formatGroups - 書式のまとまり
     * @returns {Object[]} 選択肢（id / label）
     */
    function buildFontChoices(formatGroups) {
        var fontChoices = [], seenFontNames = {};
        for (var i = 0; i < formatGroups.length; i++) {
            var psName = formatGroups[i].original.psName;
            if (seenFontNames.hasOwnProperty(psName)) continue;
            seenFontNames[psName] = true;
            var displayName = fontDisplayName(psName);
            fontChoices.push({ id: psName, label: { ja: displayName, en: displayName } });
        }
        return fontChoices;
    }

    // =========================================
    // ダイアログの組み立て / Building the dialog
    // =========================================

    /**
     * 右揃えの項目名を付けた行を追加する
     * @param {Panel} parentPanel - 追加先のパネル
     * @param {Object} labelSet - 項目名のラベル定義
     * @returns {Group} 追加した行
     */
    function addLabeledRow(parentPanel, labelSet) {
        var fieldRowGroup = parentPanel.add("group");
        setupRow(fieldRowGroup, "left");
        var rowLabel = fieldRowGroup.add("statictext", undefined, labelText(labelSet));
        rowLabel.preferredSize.width = LABEL_WIDTH;
        rowLabel.justify = "right";
        return fieldRowGroup;
    }

    /**
     * 数値欄（と単位）の行を追加する
     * @param {Panel} parentPanel - 追加先のパネル
     * @param {Object} labelSet - 項目名のラベル定義
     * @param {string} unitLabel - 欄の後ろに出す単位（空なら出さない）
     * @param {boolean} useArrowKeys - ↑↓キーでの増減を付けるなら true
     * @returns {EditText} 追加した数値欄
     */
    function addNumberRow(parentPanel, labelSet, unitLabel, useArrowKeys) {
        var fieldRowGroup = addLabeledRow(parentPanel, labelSet);
        var numberField = fieldRowGroup.add("edittext", undefined, "");
        numberField.characters = NUMBER_FIELD_CHARS;
        if (useArrowKeys) changeValueByArrowKey(numberField);
        if (unitLabel) fieldRowGroup.add("statictext", undefined, unitLabel);
        return numberField;
    }

    /**
     * 選択肢のドロップダウンの行を追加する
     * @param {Panel} parentPanel - 追加先のパネル
     * @param {Object} labelSet - 項目名のラベル定義
     * @param {Object[]} choiceOptions - 選択肢（id / label）
     * @returns {DropDownList} 追加したドロップダウン
     */
    function addDropdownRow(parentPanel, labelSet, choiceOptions) {
        var fieldRowGroup = addLabeledRow(parentPanel, labelSet);
        var choiceDropdown = fieldRowGroup.add("dropdownlist", undefined, []);
        choiceDropdown.preferredSize.width = DROPDOWN_WIDTH;
        for (var i = 0; i < choiceOptions.length; i++) choiceDropdown.add("item", getLabel(choiceOptions[i].label));
        return choiceDropdown;
    }

    /**
     * 選択肢の ID から位置を返す
     * @param {Object[]} choiceOptions - 選択肢（id / label）
     * @param {string} optionId - 探す ID
     * @returns {number|null} 位置。無ければ null（ドロップダウンは未選択になる）
     */
    function findOptionIndex(choiceOptions, optionId) {
        for (var i = 0; i < choiceOptions.length; i++) {
            if (choiceOptions[i].id === optionId) return i;
        }
        return null;
    }

    /**
     * 左側（書式の一覧と、置換のポップアップ・ボタン）を組み立てる
     * @param {Group} parentGroup - 追加先のグループ
     * @param {Object} controls - 作ったコントロールの追加先
     * @returns {void}
     */
    function buildListColumn(parentGroup, controls) {
        var listColumnGroup = parentGroup.add("group");
        listColumnGroup.orientation = "column";
        listColumnGroup.alignChildren = ["fill", "top"];
        listColumnGroup.spacing = PANEL_ROW_SPACING;

        controls.formatList = listColumnGroup.add("listbox", undefined, [], { multiselect: true });
        controls.formatList.preferredSize = [LIST_WIDTH, LIST_ROW_HEIGHT * LIST_ROW_COUNT];
        controls.formatList.helpTip = getLabel(LABELS.tooltip.formatList);

        controls.replaceRowGroup = listColumnGroup.add("group");
        setupRow(controls.replaceRowGroup, "fill");
        controls.replaceSourceDropdown = controls.replaceRowGroup.add("dropdownlist", undefined, []);
        controls.replaceSourceDropdown.alignment = ["fill", "center"];
        controls.replaceSourceDropdown.helpTip = getLabel(LABELS.tooltip.replace);
        controls.btnReplace = controls.replaceRowGroup.add("button", undefined, getLabel(LABELS.button.replace));
        controls.btnReplace.helpTip = getLabel(LABELS.tooltip.replace);
    }

    /**
     * 右側（書式の編集欄と［更新］）を組み立てる
     * @param {Group} parentGroup - 追加先のグループ
     * @param {Object[]} fontChoices - フォントの選択肢
     * @param {Object} controls - 作ったコントロールの追加先
     * @returns {void}
     */
    function buildEditPanel(parentGroup, fontChoices, controls) {
        var formatEditPanel = parentGroup.add("panel", undefined, getLabel(LABELS.panel.edit));
        setupPanel(formatEditPanel);
        controls.formatEditPanel = formatEditPanel;

        controls.fontDropdown = addDropdownRow(formatEditPanel, LABELS.fieldLabel.font, fontChoices);
        controls.fontDropdown.helpTip = getLabel(LABELS.tooltip.font);
        controls.sizeField = addNumberRow(formatEditPanel, LABELS.fieldLabel.size, textUnit.label, true);
        controls.leadingField = addNumberRow(formatEditPanel, LABELS.fieldLabel.leading, textUnit.label, true);
        controls.autoLeadingCheck = controls.leadingField.parent.add("checkbox", undefined, getLabel(LABELS.checkbox.autoLeading));
        controls.autoLeadingCheck.helpTip = getLabel(LABELS.tooltip.autoLeading);
        controls.trackingField = addNumberRow(formatEditPanel, LABELS.fieldLabel.tracking, "", false);
        controls.trackingField.helpTip = getLabel(LABELS.tooltip.tracking);
        controls.tsumeField = addNumberRow(formatEditPanel, LABELS.fieldLabel.tsume, "%", true);
        controls.kernDropdown = addDropdownRow(formatEditPanel, LABELS.fieldLabel.kern, KERN_OPTIONS);
        controls.justifyDropdown = addDropdownRow(formatEditPanel, LABELS.fieldLabel.justify, JUSTIFY_OPTIONS);
        controls.spaceBeforeField = addNumberRow(formatEditPanel, LABELS.fieldLabel.spaceBefore, textUnit.label, true);
        controls.spaceAfterField = addNumberRow(formatEditPanel, LABELS.fieldLabel.spaceAfter, textUnit.label, true);

        var editButtonRowGroup = formatEditPanel.add("group");
        setupRow(editButtonRowGroup, "right");
        editButtonRowGroup.margins = [0, BUTTON_ROW_TOP_MARGIN, 0, 0];
        controls.btnUpdate = editButtonRowGroup.add("button", undefined, getLabel(LABELS.button.update));
        controls.btnUpdate.helpTip = getLabel(LABELS.tooltip.update);
    }

    /**
     * ボタンエリア（キャンセル／OK）を組み立てる
     * @param {Window} formatDialog - 追加先のダイアログ
     * @param {Object} controls - 作ったコントロールの追加先
     * @returns {void}
     */
    function buildButtonRow(formatDialog, controls) {
        var btnRowGroup = formatDialog.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.margins = [0, BUTTON_ROW_TOP_MARGIN, 0, 0];
        btnRowGroup.alignment = ["fill", "bottom"];
        var spacer = btnRowGroup.add("group");
        spacer.alignment = ["fill", "fill"];
        spacer.minimumSize.width = 0;
        var btnRightGroup = btnRowGroup.add("group");
        btnRightGroup.alignChildren = ["right", "center"];
        var btnCancel = btnRightGroup.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });
        btnCancel.helpTip = getLabel(LABELS.tooltip.cancel);
        controls.btnOK = btnRightGroup.add("button", undefined, getLabel(LABELS.button.ok), { name: "ok" });
        controls.btnOK.helpTip = getLabel(LABELS.tooltip.ok);
    }

    /**
     * ダイアログを組み立て、操作に使うコントロールをまとめて返す
     * @param {Object[]} fontChoices - フォントの選択肢
     * @returns {Object} コントロール（window / formatList / replaceSourceDropdown / 各欄 / 各ボタン など）
     */
    function buildFormatDialog(fontChoices) {
        var controls = {};
        var formatDialog = new Window("dialog", getLabel(LABELS.dialog.title) + " " + SCRIPT_VERSION);
        formatDialog.orientation = "column";
        formatDialog.alignChildren = ["fill", "top"];
        formatDialog.margins = WINDOW_MARGINS;
        formatDialog.spacing = WINDOW_SPACING;
        controls.window = formatDialog;

        var mainRowGroup = formatDialog.add("group");
        mainRowGroup.orientation = "row";
        mainRowGroup.alignChildren = ["fill", "top"];
        mainRowGroup.spacing = WINDOW_SPACING;

        buildListColumn(mainRowGroup, controls);
        buildEditPanel(mainRowGroup, fontChoices, controls);
        buildButtonRow(formatDialog, controls);
        return controls;
    }

    // =========================================
    // ダイアログの操作 / Dialog behavior
    // =========================================

    /**
     * 書式の一覧と編集欄のダイアログを表示する。［置換］［更新］はその場でカンバスへ反映する
     * @param {Object} dialogState - doc（対象のドキュメント）と groups（書式のまとまり。置換・更新で差し替える）
     * @param {number} initialIndex - 最初に選んでおく一覧の位置
     * @returns {boolean} OK で閉じたら true
     */
    function showFormatDialog(dialogState, initialIndex) {
        /* 書式のまとまり。置換・更新のたびに調べ直して差し替える / Format groups; replaced by a fresh survey after each Replace/Update */
        var formatGroups = dialogState.groups;
        var fontChoices = buildFontChoices(formatGroups);
        var controls = buildFormatDialog(fontChoices);
        var formatList = controls.formatList;
        var replaceSourceDropdown = controls.replaceSourceDropdown;

        /* 一覧を作り直す間は onChange を止める / Suppress onChange while the list is rebuilt */
        var suppressListEvents = false;

        /* 編集欄で編集中のまとまりの位置。1つだけ選んでいるときだけ 0 以上
           Index of the group being edited in the fields; 0 or more only while exactly one is selected */
        var editingIndex = -1;

        /* 一覧で選んでいるまとまりの位置（昇順）/ Indices of the groups selected in the list, ascending */
        var selectedIndices = [];

        /**
         * 一覧を作り直し、選択位置を戻す（行の文字を1行ずつ書き換えない。Mac では別の行に反映されることがある）
         * @returns {void}
         */
        function fillFormatList() {
            suppressListEvents = true;
            formatList.removeAll();
            for (var i = 0; i < formatGroups.length; i++) formatList.add("item", buildFormatSummary(formatGroups[i]));
            formatList.selection = selectedIndices.length ? selectedIndices : null;
            suppressListEvents = false;
        }

        /**
         * 置換のポップアップへ書式を並べ直す（先頭は「選んでいない」状態の目印）
         * @returns {void}
         */
        function fillReplaceSourceDropdown() {
            replaceSourceDropdown.removeAll();
            replaceSourceDropdown.add("item", getLabel(LABELS.dropdown.noReplaceSource));
            for (var i = 0; i < formatGroups.length; i++) replaceSourceDropdown.add("item", buildFormatSummary(formatGroups[i]));
            replaceSourceDropdown.selection = 0;
        }

        /**
         * 一覧で選んでいる行の位置を昇順で返す
         * @returns {number[]} 選んでいる行の位置
         */
        function readSelectedIndices() {
            var pickedIndices = [];
            var listSelection = formatList.selection;
            if (!listSelection) return pickedIndices;
            /* 複数選択の一覧でも、1行だけのときは配列でなく ListItem が返ることがある / A single pick may come back as a lone ListItem */
            if (listSelection.length === undefined) listSelection = [listSelection];
            for (var i = 0; i < listSelection.length; i++) pickedIndices.push(listSelection[i].index);
            pickedIndices.sort(function (a, b) { return a - b; });
            return pickedIndices;
        }

        /**
         * 位置が選んでいる行に含まれるかを返す
         * @param {number} groupIndex - 調べる位置
         * @returns {boolean} 含まれれば true
         */
        function isSelectedIndex(groupIndex) {
            for (var i = 0; i < selectedIndices.length; i++) {
                if (selectedIndices[i] === groupIndex) return true;
            }
            return false;
        }

        /**
         * まだカンバスに出していない編集を反映してから、ドキュメントを調べ直して一覧を作り直す
         * 同じ書式になったまとまりは1行に統合され、段落数も数え直される
         * @param {string} preferredKey - 作り直したあとに選んでおく書式のキー
         * @returns {void}
         */
        function rescanFormats(preferredKey) {
            for (var i = 0; i < formatGroups.length; i++) applyEditedToCanvas(formatGroups[i]);
            app.redraw();
            formatGroups = collectFormatGroups(dialogState.doc);
            dialogState.groups = formatGroups;
            selectedIndices = [];
            for (var j = 0; j < formatGroups.length; j++) {
                if (buildFormatKey(formatGroups[j].original) === preferredKey) { selectedIndices.push(j); break; }
            }
            fillReplaceSourceDropdown();
            fillFormatList();
            showSelectedFormat();
        }

        /**
         * 編集欄の見出しに、まだカンバスに出していない編集があるかを出す
         * @returns {void}
         */
        function refreshEditPanelTitle() {
            var editingGroup = (editingIndex >= 0) ? formatGroups[editingIndex] : null;
            var hasPending = editingGroup && listDifferingKeys(editingGroup.applied, editingGroup.edited, editingGroup.leadingTouched).length > 0;
            controls.formatEditPanel.text = getLabel(hasPending ? LABELS.panel.editPending : LABELS.panel.edit);
        }

        /**
         * 数値欄の値を pt で読む（読めなければ元の値）
         * @param {EditText} numberField - 対象の数値欄
         * @param {number} fallbackPoints - 読めないときの値（pt）
         * @returns {number} 値（pt、小数第2位まで）
         */
        function readPointsField(numberField, fallbackPoints) {
            var parsedValue = parseFloat(numberField.text);
            if (isNaN(parsedValue)) return fallbackPoints;
            /* 表示を丸めただけで値を変えたことにしない / Unchanged text must not count as an edit */
            if (numberField.text === pointsToUnitText(fallbackPoints)) return fallbackPoints;
            return roundToHundredths(parsedValue * textUnit.pointsPerUnit);
        }

        /**
         * 数値欄の値を整数で読む（読めなければ元の値）
         * @param {EditText} numberField - 対象の数値欄
         * @param {number} fallbackValue - 読めないときの値
         * @returns {number} 整数
         */
        function readIntegerField(numberField, fallbackValue) {
            var parsedValue = parseFloat(numberField.text);
            return isNaN(parsedValue) ? fallbackValue : Math.round(parsedValue);
        }

        /**
         * 行送りの欄と［自動］を、編集後の書式へ書き戻す
         * 自分で触ったら、既定の「自動」ではなく指定として扱う
         * @param {Object} editingGroup - 編集中の書式のまとまり
         * @returns {void}
         */
        function storeLeadingFields(editingGroup) {
            var edited = editingGroup.edited;
            var previousAuto = edited.autoLeading, previousLeading = edited.leading;
            edited.autoLeading = controls.autoLeadingCheck.value;
            if (edited.autoLeading) {
                /* 自動はサイズに追従するので、表示用の値だけ計算し直す / Auto follows the size; recompute the shown value */
                edited.leading = roundToHundredths(edited.size * edited.autoLeadingAmount / 100);
            } else if (previousAuto) {
                /* 自動を外したら、元の固定値（元も自動なら算出値）から始める / Unticking Auto starts from the original fixed value */
                edited.leading = editingGroup.original.autoLeading ? previousLeading : editingGroup.original.leading;
            } else {
                edited.leading = readPointsField(controls.leadingField, edited.leading);
            }
            if (edited.autoLeading !== previousAuto || (!edited.autoLeading && edited.leading !== previousLeading)) {
                editingGroup.leadingTouched = true;
            }
        }

        /**
         * 編集欄の値を、編集中のまとまりの編集後の書式へ書き戻す
         * 表示前には呼ばない（表示前のチェックボックスは値が読み戻せない）
         * @returns {void}
         */
        function storeFieldsToEdited() {
            if (editingIndex < 0) return;
            var editingGroup = formatGroups[editingIndex];
            var edited = editingGroup.edited;
            if (controls.fontDropdown.selection) edited.psName = fontChoices[controls.fontDropdown.selection.index].id;
            edited.size = readPointsField(controls.sizeField, edited.size);
            storeLeadingFields(editingGroup);
            edited.tracking = readIntegerField(controls.trackingField, edited.tracking);
            edited.tsume = readIntegerField(controls.tsumeField, edited.tsume);
            if (controls.kernDropdown.selection) edited.kern = KERN_OPTIONS[controls.kernDropdown.selection.index].id;
            if (controls.justifyDropdown.selection) edited.justify = JUSTIFY_OPTIONS[controls.justifyDropdown.selection.index].id;
            edited.spaceBefore = readPointsField(controls.spaceBeforeField, edited.spaceBefore);
            edited.spaceAfter = readPointsField(controls.spaceAfterField, edited.spaceAfter);
            refreshEditPanelTitle();
        }

        /**
         * 書式を編集欄へ表示する
         * @param {Object} format - 表示する書式
         * @returns {void}
         */
        function showFormatInFields(format) {
            controls.fontDropdown.selection = findOptionIndex(fontChoices, format.psName);
            controls.sizeField.text = pointsToUnitText(format.size);
            controls.leadingField.text = pointsToUnitText(format.leading);
            controls.autoLeadingCheck.value = format.autoLeading;
            controls.leadingField.enabled = !format.autoLeading;
            controls.trackingField.text = String(format.tracking);
            controls.tsumeField.text = String(format.tsume);
            controls.kernDropdown.selection = findOptionIndex(KERN_OPTIONS, format.kern);
            controls.justifyDropdown.selection = findOptionIndex(JUSTIFY_OPTIONS, format.justify);
            controls.spaceBeforeField.text = pointsToUnitText(format.spaceBefore);
            controls.spaceAfterField.text = pointsToUnitText(format.spaceAfter);
        }

        /**
         * 一覧で選んでいる書式を編集欄に出し、置換のポップアップを合わせる（欄の書き戻しはしない）
         * @returns {void}
         */
        function showSelectedFormat() {
            selectedIndices = readSelectedIndices();
            /* 編集欄は1つだけ選んでいるときに使う。複数のときは先頭の値をディム表示 / Fields work with a single pick; with several, the first is shown dimmed */
            editingIndex = (selectedIndices.length === 1) ? selectedIndices[0] : -1;
            controls.formatEditPanel.enabled = (editingIndex >= 0);
            controls.replaceRowGroup.enabled = (selectedIndices.length > 0);
            /* 一覧で選んでいる書式は、置換元としてディム表示 / Dim the list's own formats among the replace-with choices */
            for (var i = 0; i < formatGroups.length; i++) replaceSourceDropdown.items[i + 1].enabled = !isSelectedIndex(i);
            if (replaceSourceDropdown.selection && isSelectedIndex(replaceSourceDropdown.selection.index - 1)) replaceSourceDropdown.selection = 0;
            if (selectedIndices.length) showFormatInFields(formatGroups[selectedIndices[0]].edited);
            refreshEditPanelTitle();
        }

        /**
         * 一覧の選択に合わせて編集欄を切り替える（前のまとまりの編集は書き戻してから）
         * @returns {void}
         */
        function onFormatSelected() {
            if (suppressListEvents) return;
            storeFieldsToEdited();
            showSelectedFormat();
        }

        /**
         * 編集欄を変えたら、値を書き戻して表示を整える
         * @returns {void}
         */
        function onFieldChanged() {
            storeFieldsToEdited();
            if (editingIndex >= 0) showFormatInFields(formatGroups[editingIndex].edited);
        }

        /**
         * 更新：欄の値をその書式の段落へ反映し、調べ直す
         * @returns {void}
         */
        function onUpdateClick() {
            if (editingIndex < 0) return;
            storeFieldsToEdited();
            var updatedGroup = formatGroups[editingIndex];
            applyEditedToCanvas(updatedGroup);
            rescanFormats(buildFormatKey(updatedGroup.applied));
        }

        /**
         * 置換：ポップアップで選んだ書式の値を、一覧で選んでいる書式へそのまま写して反映する
         * 置換したものは置換元の書式に統合されるので、調べ直したあとはそれを選んでおく
         * @returns {void}
         */
        function onReplaceClick() {
            var sourceIndex = replaceSourceDropdown.selection ? replaceSourceDropdown.selection.index - 1 : -1;
            if (sourceIndex < 0 || !selectedIndices.length) return;
            var sourceFormat = formatGroups[sourceIndex].original;
            var replacedCount = 0;
            for (var i = 0; i < selectedIndices.length; i++) {
                /* 置換元そのものはディム表示だが、選べてしまう環境に備えて飛ばす / Skip the source itself in case its dimmed entry got picked */
                if (selectedIndices[i] === sourceIndex) continue;
                var targetGroup = formatGroups[selectedIndices[i]];
                targetGroup.edited = copyFormat(sourceFormat);
                /* 置換元の行送り（自動か固定か）をそのまま使う / Keep the source's leading as it is */
                targetGroup.leadingTouched = true;
                replacedCount++;
            }
            if (replacedCount) rescanFormats(buildFormatKey(sourceFormat));
            else replaceSourceDropdown.selection = 0;
        }

        formatList.onChange = onFormatSelected;
        var editControls = [controls.fontDropdown, controls.sizeField, controls.leadingField, controls.trackingField,
            controls.tsumeField, controls.kernDropdown, controls.justifyDropdown, controls.spaceBeforeField, controls.spaceAfterField];
        for (var i = 0; i < editControls.length; i++) editControls[i].onChange = onFieldChanged;
        controls.autoLeadingCheck.onClick = onFieldChanged;
        controls.btnUpdate.onClick = onUpdateClick;
        controls.btnReplace.onClick = onReplaceClick;
        /* OK の前に、表示中の欄を書き戻す / Store the shown fields before closing with OK */
        controls.btnOK.onClick = function () {
            storeFieldsToEdited();
            controls.window.close(1);
        };

        /* 最初の選択では欄を書き戻さない。表示前のチェックボックスは値が読み戻せず、「自動」が外れたことになるため
           The first selection must not store the fields: before show, the checkbox reads back as off and Auto would look unticked */
        selectedIndices = [initialIndex];
        fillReplaceSourceDropdown();
        fillFormatList();
        showSelectedFormat();

        return controls.window.show() === 1;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * ドキュメントを調べてダイアログを表示する。OK なら変更を確定し、キャンセルなら元に戻す
     * 起動時に選択していた段落の書式を最初に選んでおく
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) { alert(getLabel(LABELS.alert.noDocument)); return; }
        var doc = app.activeDocument;
        textUnit = getUnitInfo("text/units");

        var formatGroups = collectFormatGroups(doc);
        if (!formatGroups.length) { alert(getLabel(LABELS.alert.noText)); return; }

        /* キャンセルで戻すための控え（ダイアログ側の一覧は置換のたびに作り直される）
           Snapshot for Cancel; the dialog rebuilds its own list after every Replace */
        var initialGroups = formatGroups.slice(0);
        var dialogState = { doc: doc, groups: formatGroups };
        if (showFormatDialog(dialogState, findInitialIndex(doc, formatGroups))) {
            /* OK なら、まだカンバスに出していない編集も反映 / OK applies any edits not yet on the canvas */
            for (var i = 0; i < dialogState.groups.length; i++) applyEditedToCanvas(dialogState.groups[i]);
        } else {
            /* キャンセルなら、開いたときの書式へ戻す / Cancel puts everything back as it was when the dialog opened */
            restoreInitialFormats(initialGroups);
        }
        app.redraw();
    }

    main();

})();
