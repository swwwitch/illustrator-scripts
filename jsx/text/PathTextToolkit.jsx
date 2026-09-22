#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

パス上文字の作成と調整をまとめて行うツールです。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/PathTextToolkit.md

### Overview

A toolkit for creating and adjusting text on a path.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/PathTextToolkit.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "PathTextToolkit";              /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.3.4";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-23";                             /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/PathTextToolkit.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/PathTextToolkit.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // レイアウト / Layout
    // =========================================
    var DIALOG_MARGINS        = [15, 20, 15, 15];  /* ダイアログの余白 / dialog margins */
    var PANEL_MARGINS         = [15, 20, 15, 10];  /* パネルの余白 / panel margins */
    var PANEL_SPACING         = 8;                 /* パネル内の間隔 / panel spacing */
    var ARC_LABEL_WIDTH       = 70;                /* ［アーチ方向］の項目名の幅 / width of the arc direction label */
    var ARC_SLIDER_WIDTH      = 130;               /* ［まるみ］スライダーの幅 / width of the roundness slider */
    var FIT_PANEL_GAP         = 5;                 /* ［パス幅に合わせる］パネルの上の余白 / gap above the fit-to-path-width panel */
    var POSITION_LABEL_WIDTH  = 60;                /* 開始／終了位置の項目名の幅 / width of the start/end labels */
    var POSITION_FIELD_CHARS  = 5;                 /* 開始／終了位置の入力欄の幅 / width of the start/end fields */
    var ADJUST_LABEL_WIDTH    = 95;                /* テキスト調整の項目名の幅 / width of the text-adjust labels */
    var ADJUST_FIELD_CHARS    = 6;                 /* テキスト調整の入力欄の幅 / width of the text-adjust fields */
    var ROW_SLIDER_WIDTH      = 180;               /* 各行のスライダーの幅 / width of the row sliders */
    var EDIT_HINT_WIDTH       = 360;               /* テキスト編集ダイアログの説明の幅 / width of the text-edit hint */
    var EDIT_FIELD_SIZE       = [320, 80];         /* テキスト編集ダイアログの入力欄 / size of the text-edit field */

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * Illustrator の UI 言語から表示言語を判定する
     * @returns {string} "ja" または "en"
     */
    function detectUILanguage() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }

    var uiLang = detectUILanguage();

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "パス上文字（変換と調整）", en: "Text on a Path (Convert & Adjust)" },
            textEditTitle: { ja: "テキスト編集", en: "Edit Text" },
            textEditHint: {
                ja: "内容を編集してOKで反映します。複数選択の場合は全て置換します。",
                en: "Edit the content and press OK to apply. If multiple items are selected, it replaces all."
            }
        },
        panel: {
            process: { ja: "パス上文字にする", en: "Create Path Text" },
            split: { ja: "テキストを分離", en: "Split Text" },
            option: { ja: "オプション", en: "Options" },
            effect: { ja: "効果", en: "Effect" },
            position: { ja: "位置", en: "Position" },
            textAdjust: { ja: "テキスト調整", en: "Text Adjust" },
            fitWidth: { ja: "パス幅に合わせる", en: "Fit to Path Width" }
        },
        radio: {
            toPathText: { ja: "パス上文字にする", en: "Convert to Path Text" },
            genArcPath: { ja: "アーチ状のパスを生成", en: "Generate Arc Path" },
            genCircle: { ja: "正円を生成", en: "Generate Circle" },
            splitKeepFormat: { ja: "書式を保持", en: "Keep Formatting" },
            splitNoFormat: { ja: "書式を保持しない", en: "Do Not Keep Formatting" },
            alignLeft: { ja: "左揃え", en: "Left" },
            alignCenter: { ja: "中央", en: "Center" },
            alignRight: { ja: "右揃え", en: "Right" },
            alignFullJustify: { ja: "両端揃え", en: "Justify" },
            arcUp: { ja: "上", en: "Up" },
            arcDown: { ja: "下", en: "Down" },
            fitWidthNone: { ja: "しない", en: "None" },
            fitWidthFontSize: { ja: "文字サイズを調整", en: "Adjust Font Size" },
            fitWidthTracking: { ja: "トラッキングを調整", en: "Adjust Tracking" },
            effectRainbow: { ja: "虹", en: "Rainbow" },
            effectDistort: { ja: "歪み", en: "Skew" },
            effectRibbon: { ja: "3D リボン", en: "3D Ribbon" },
            effectStep: { ja: "階段", en: "Stair Step" },
            effectGravity: { ja: "引力", en: "Gravity" }
        },
        checkbox: {
            splitDeletePath: { ja: "パスを削除", en: "Remove Path" },
            reverse: { ja: "内側配置", en: "Inside" },
            preview: { ja: "プレビュー", en: "Preview" }
        },
        fieldLabel: {
            align: { ja: "行揃え", en: "Alignment" },
            arcDirection: { ja: "アーチ方向", en: "Arc Direction" },
            arcRoundness: { ja: "まるみ", en: "Roundness" },
            startPos: { ja: "開始位置", en: "Start" },
            endPos: { ja: "終了位置", en: "End" },
            baseShift: { ja: "ベースライン", en: "Baseline Shift" },
            tracking: { ja: "トラッキング", en: "Tracking" },
            fontSize: { ja: "文字サイズ", en: "Font Size" }
        },
        button: {
            textEdit: { ja: "テキスト編集", en: "Edit Text" },
            ok: { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        tooltip: {
            textEdit: { ja: "選択したテキストの内容を、別ダイアログで書き換えます。", en: "Opens a separate dialog for rewriting the selected text." },
            toPathText: { ja: "選択したテキストを、一緒に選んだパスの上に流し込みます。", en: "Flows the selected text onto the path selected with it." },
            genArcPath: { ja: "テキストの幅に合わせたアーチ状のパスを作り、その上に流し込みます。", en: "Builds an arc sized to the text and flows the text onto it." },
            genCircle: {
                ja: "テキストの幅を円周とする正円を作り、その上に流し込みます。",
                en: "Builds a circle whose circumference matches the text width and flows the text onto it."
            },
            splitKeepFormat: { ja: "パス上文字を、書式を保ったままテキストとパスに分けます。", en: "Splits path text into text and path, keeping the formatting." },
            splitNoFormat: { ja: "パス上文字を、書式を落としてテキストとパスに分けます。", en: "Splits path text into text and path, dropping the formatting." },
            splitDeletePath: { ja: "分離したあと、元のパスを削除します。", en: "Deletes the original path after the split." },
            reverse: { ja: "文字をパスの反対側（内側）に配置します。", en: "Places the characters on the other side of the path." },
            arcUp: { ja: "上に膨らんだアーチにします。", en: "Bulges the arc upward." },
            arcDown: { ja: "下に膨らんだアーチにします。", en: "Bulges the arc downward." },
            arcRoundness: { ja: "アーチの曲がり具合です。0で直線に近づきます。", en: "How strongly the arc curves. 0 is nearly straight." },
            fitWidthNone: { ja: "文字の大きさも字間もそのままにします。", en: "Leaves both the size and the spacing as they are." },
            fitWidthFontSize: { ja: "文字サイズを変えて、パスの長さいっぱいに収めます。", en: "Changes the font size so the text fills the path." },
            fitWidthTracking: { ja: "字間を変えて、パスの長さいっぱいに収めます。", en: "Changes the tracking so the text fills the path." },
            effectRainbow: { ja: "1文字ずつ色を変えて虹色にします。", en: "Colours each character to make a rainbow." },
            effectDistort: { ja: "1文字ずつ斜めに傾けます。", en: "Skews each character." },
            effectRibbon: { ja: "奥行きのあるリボンのように、1文字ずつ変形します。", en: "Transforms each character like a ribbon with depth." },
            effectStep: { ja: "1文字ずつ高さをずらして階段状にします。", en: "Offsets each character vertically, like stairs." },
            effectGravity: { ja: "中央へ引き寄せられたように、1文字ずつ大きさと位置を変えます。", en: "Varies each character as if pulled toward the centre." },
            startPosEnabled: {
                ja: "開始位置を指定します。オフのときはパスの先頭から始めます。",
                en: "Sets the start position. When off, the text starts at the beginning of the path."
            },
            startPos: { ja: "文字の流し込みを始める位置です。", en: "Where the text begins along the path." },
            endPosEnabled: {
                ja: "終了位置を指定します。オフのときはパスの終端まで使います。",
                en: "Sets the end position. When off, the text runs to the end of the path."
            },
            endPos: { ja: "文字の流し込みを終える位置です。", en: "Where the text ends along the path." },
            align: { ja: "開始位置と終了位置の間での文字の揃え方です。", en: "How the characters are aligned between the start and end positions." },
            baseShift: { ja: "パスから文字を浮かせる量です。負の値で沈みます。", en: "How far the characters sit above the path. Negative values sink below it." },
            tracking: { ja: "文字と文字の間隔です。単位は1/1000em。", en: "Spacing between characters, in 1/1000 em." },
            fontSize: { ja: "元の文字サイズからの増減です。", en: "Change applied to the original font size." },
            preview: { ja: "結果を画面で確認します。キャンセルすると元に戻ります。", en: "Shows the result on the canvas. Cancel restores the original state." }
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません", en: "No document is open." },
            noText: { ja: "対象のテキストが見つかりません", en: "No target text found." },
            needPath: { ja: "パスを一緒に選択してください", en: "Select a path together." },
            dupPathFail: { ja: "パスの複製に失敗しました", en: "Failed to duplicate the path." },
            arcFail: { ja: "アーチ状のパス生成に失敗しました", en: "Failed to generate an arc path." },
            needPathText: { ja: "パス上文字を選択してください", en: "Select path text." }
        }
    };

    /**
     * ドット区切りのパスで表示言語のラベルを取得する
     * @param {string} labelPath - "panel.effect" のようなドット区切りキー
     * @returns {string} 表示言語のラベル（無ければ ja、それも無ければパスをそのまま返す）
     */
    function getLabel(labelPath) {
        var pathKeys = labelPath.split(".");
        var labelNode = LABELS;
        for (var i = 0; i < pathKeys.length; i++) {
            labelNode = labelNode[pathKeys[i]];
            if (!labelNode) return labelPath;
        }
        if (labelNode[uiLang]) return labelNode[uiLang];
        if (labelNode.ja) return labelNode.ja;
        return labelPath;
    }

    /**
     * コロン付きの項目名を返す（日本語は全角、英語は半角）
     * @param {string} labelPath - ラベルのパス
     * @returns {string} コロン付きの項目名
     */
    function labelText(labelPath) {
        return getLabel(labelPath) + (uiLang === "ja" ? "：" : ":");
    }

    // =========================================
    // 共通ヘルパー / Shared helpers
    // =========================================

    /**
     * 入力欄の値を ↑↓ キーで増減する（↑↓=±1、Shift=±10、Option=±0.1）
     * @param {EditText} editText - 対象の入力欄
     * @param {boolean} allowNegative - false なら 0 未満にしない
     * @param {Function} [onChanged] - 値が変わったあとに呼ぶ関数
     * @returns {void}
     */
    function changeValueByArrowKey(editText, allowNegative, onChanged) {
        if (!editText) return;
        editText.addEventListener("keydown", function (event) {
            if (!event || (event.keyName !== "Up" && event.keyName !== "Down")) return;

            var value = Number(editText.text);
            if (isNaN(value)) return;

            var keyboard = ScriptUI.environment.keyboardState;
            var delta = 1;

            if (keyboard.shiftKey) {
                delta = 10;
                /* Shift は 10 の倍数へ寄せる / Snap to multiples of 10 for Shift */
                if (event.keyName === "Up") {
                    value = Math.ceil((value + 1) / delta) * delta;
                } else {
                    value = Math.floor((value - 1) / delta) * delta;
                }
            } else if (keyboard.altKey) {
                delta = 0.1;
                if (event.keyName === "Up") value += delta;
                else value -= delta;
            } else {
                delta = 1;
                if (event.keyName === "Up") value += delta;
                else value -= delta;
            }

            if (keyboard.altKey) value = Math.round(value * 10) / 10;
            else value = Math.round(value);

            if (!allowNegative && value < 0) value = 0;

            editText.text = String(value);

            /* 矢印キー本来のカーソル移動を止める / Prevent the default cursor move */
            event.preventDefault();

            if (onChanged && typeof onChanged === "function") {
                /* プレビュー更新で DOM の例外が出てもキー操作は続ける / keep the key handler alive if the preview throws */
                try { onChanged(); } catch (e) { }
            }
        });
    }

    /**
     * 値を min〜max の範囲に収める
     * @param {number} value - 対象の値
     * @param {number} min - 下限
     * @param {number} max - 上限
     * @returns {number} 範囲に収めた値
     */
    function clampNumber(value, min, max) {
        if (value < min) return min;
        if (value > max) return max;
        return value;
    }

    /**
     * 文字列を数値へ変換する（失敗したら既定値）
     * @param {string} text - 対象の文字列
     * @param {number} fallback - 既定値
     * @returns {number} 数値
     */
    function parseNumberOr(text, fallback) {
        var parsed = Number(text);
        if (isNaN(parsed)) return fallback;
        return parsed;
    }

    /**
     * 文字列を t 値（0〜5）として解釈する
     * @param {string} text - 対象の文字列
     * @param {number} fallback - 既定値
     * @returns {number} t 値
     */
    function parseTValue(text, fallback) {
        var parsed = Number(text);
        if (isNaN(parsed)) return fallback;
        return clampNumber(parsed, 0, 5);
    }

    /**
     * 数値を小数1桁の文字列にする（0.1pt 単位）
     * @param {number} value - 対象の値
     * @returns {string} 小数1桁の文字列
     */
    function formatOneDecimal(value) {
        return (Math.round(value * 10) / 10).toFixed(1);
    }

    /**
     * スミ100%の CMYKColor を作る
     * @returns {CMYKColor} 黒
     */
    function makeBlackCMYK() {
        var black = new CMYKColor();
        black.cyan = 0;
        black.magenta = 0;
        black.yellow = 0;
        black.black = 100;
        return black;
    }

    // =========================================
    // 選択の解析 / Selection analysis
    // =========================================

    /**
     * 選択内のポイント文字・パス上文字を再帰的に集める
     * @param {PageItem[]} pageItems - 対象のオブジェクト
     * @returns {TextFrame[]} 見つかったテキスト
     */
    function getTargetTextItems(pageItems) {
        var textItems = [];
        for (var i = 0; i < pageItems.length; i++) {
            var pageItem = pageItems[i];
            if (pageItem.typename === "TextFrame") {
                try {
                    if (pageItem.kind === TextType.POINTTEXT || pageItem.kind === TextType.PATHTEXT) {
                        textItems.push(pageItem);
                    }
                } catch (e) { }
            } else if (pageItem.typename === "GroupItem") {
                textItems = textItems.concat(getTargetTextItems(pageItem.pageItems));
            }
        }
        return textItems;
    }

    /**
     * 選択内のパス（PathItem / CompoundPathItem）を再帰的に集める（グループ内も）
     * @param {PageItem[]} pageItems - 対象のオブジェクト
     * @returns {PageItem[]} 見つかったパス
     */
    function getSelectedPathItems(pageItems) {
        var pathItems = [];
        for (var i = 0; i < pageItems.length; i++) {
            var pageItem = pageItems[i];
            if (pageItem.typename === "PathItem") {
                pathItems.push(pageItem);
            } else if (pageItem.typename === "CompoundPathItem") {
                pathItems.push(pageItem);
            } else if (pageItem.typename === "GroupItem") {
                pathItems = pathItems.concat(getSelectedPathItems(pageItem.pageItems));
            }
        }
        return pathItems;
    }

    /**
     * 選択から、使える処理（パス上文字にする／分離）を判定する
     * @param {TextFrame[]} textItems - 対象のテキスト
     * @param {PageItem[]} pathItems - 一緒に選択されたパス
     * @returns {{hasPathTextTarget: boolean, canToPathText: boolean, canSplit: boolean}} 判定結果
     */
    function detectModeAvailability(textItems, pathItems) {
        var hasSelectedPath = (pathItems && pathItems.length > 0);
        var hasPathTextTarget = false;
        for (var i = 0; i < textItems.length; i++) {
            try {
                var textItem = textItems[i];
                if (textItem && textItem.typename === "TextFrame" && textItem.kind === TextType.PATHTEXT && textItem.textPath) {
                    hasPathTextTarget = true;
                    break;
                }
            } catch (e) { }
        }
        /* 「パス上文字にする」はパスを選んでいるか、パス上文字が自分のパスを持っていれば使える
           「分離」はパス上文字を選んでいるときだけ */
        return {
            hasPathTextTarget: hasPathTextTarget,
            canToPathText: hasSelectedPath || hasPathTextTarget,
            canSplit: hasPathTextTarget
        };
    }

    // =========================================
    // 前提チェック / Preconditions
    // =========================================
    if (app.documents.length === 0) {
        alert(getLabel("alert.noDocument"));
        return false;
    }
    var doc = app.activeDocument;
    var initialSelection = doc.selection;

    /* ダイアログ表示中のプレビューを安定させるための、最初の選択の控え / Base selection snapshot for a stable preview */
    var baseSelection = [];
    /* テキスト編集中の selection は配列ではない（slice が無い） / the selection is not an array while editing text */
    try { baseSelection = initialSelection.slice(0); } catch (e) { baseSelection = []; }

    var targetItems = getTargetTextItems(initialSelection);
    var selectedPaths = getSelectedPathItems(initialSelection);

    if (targetItems.length === 0) {
        alert(getLabel("alert.noText"));
        return false;
    }

    /**
     * 現在の選択を返す（プレビュー中に選択が空になったら最初の選択の控えを使う）
     * @returns {Object[]} 選択
     */
    function readCurrentSelection() {
        var currentSelection = [];
        try { currentSelection = doc.selection; } catch (e) { currentSelection = []; }
        if (!currentSelection || currentSelection.length === 0) currentSelection = baseSelection;
        return currentSelection;
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * パネルを縦並び・共通の余白と間隔で初期化する
     * @param {Panel} targetPanel - 対象のパネル
     * @param {number} [spacing] - 間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function setupPanel(targetPanel, spacing) {
        targetPanel.orientation = "column";
        targetPanel.alignChildren = ["fill", "top"];
        targetPanel.alignment = "fill";
        targetPanel.margins = PANEL_MARGINS;
        targetPanel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /**
     * 横並び（左寄せ・上下中央）の行を追加する
     * @param {Object} parentContainer - 追加先
     * @returns {Group} 作成した行
     */
    function addRow(parentContainer) {
        var newRow = parentContainer.add("group");
        newRow.orientation = "row";
        newRow.alignChildren = ["left", "center"];
        return newRow;
    }

    /**
     * 上寄せ・幅いっぱいのグループを追加する
     * @param {Object} parentContainer - 追加先
     * @param {string} orientation - "row" または "column"
     * @returns {Group} 作成したグループ
     */
    function addFillGroup(parentContainer, orientation) {
        var newGroup = parentContainer.add("group");
        newGroup.orientation = orientation;
        newGroup.alignChildren = ["fill", "top"];
        newGroup.alignment = ["fill", "top"];
        return newGroup;
    }

    /**
     * tooltip 付きのラジオボタンを追加する
     * @param {Object} parentContainer - 追加先
     * @param {string} labelPath - ラベルのパス
     * @param {string} tooltipPath - tooltip のパス
     * @returns {RadioButton} 作成したラジオボタン
     */
    function addRadio(parentContainer, labelPath, tooltipPath) {
        var radio = parentContainer.add("radiobutton", undefined, getLabel(labelPath));
        radio.helpTip = getLabel(tooltipPath);
        return radio;
    }

    /**
     * 「項目名＋入力欄＋スライダー」の行を追加する（テキスト調整用）
     * @param {Panel} parentPanel - 追加先のパネル
     * @param {string} labelPath - 項目名のパス
     * @param {string} tooltipPath - 入力欄とスライダーの tooltip のパス
     * @param {string} initialText - 入力欄の初期値
     * @param {number} sliderMin - スライダーの最小値
     * @param {number} sliderMax - スライダーの最大値
     * @returns {{valueInput: EditText, slider: Slider}} 作成した入力欄とスライダー
     */
    function addAdjustRow(parentPanel, labelPath, tooltipPath, initialText, sliderMin, sliderMax) {
        var adjustRow = addRow(parentPanel);
        var adjustLabel = adjustRow.add("statictext", undefined, labelText(labelPath));
        adjustLabel.preferredSize.width = ADJUST_LABEL_WIDTH;

        var valueInput = adjustRow.add("edittext", undefined, initialText);
        valueInput.helpTip = getLabel(tooltipPath);
        valueInput.characters = ADJUST_FIELD_CHARS;
        var slider = adjustRow.add("slider", undefined, 0, sliderMin, sliderMax);
        slider.helpTip = getLabel(tooltipPath);
        slider.preferredSize.width = ROW_SLIDER_WIDTH;
        return { valueInput: valueInput, slider: slider };
    }

    /**
     * 「チェック＋項目名＋入力欄＋スライダー」の行を追加する（開始／終了位置用）
     * @param {Panel} parentPanel - 追加先のパネル
     * @param {string} labelPath - 項目名のパス
     * @param {string} enabledTooltipPath - チェックボックスの tooltip のパス
     * @param {string} tooltipPath - 入力欄とスライダーの tooltip のパス
     * @param {string} initialText - 入力欄の初期値
     * @param {number} sliderValue - スライダーの初期値
     * @param {number} sliderMax - スライダーの最大値
     * @returns {{enabledCheckbox: Checkbox, valueInput: EditText, slider: Slider}} 作成したコントロール
     */
    function addPositionRow(parentPanel, labelPath, enabledTooltipPath, tooltipPath, initialText, sliderValue, sliderMax) {
        var positionRow = addRow(parentPanel);
        var enabledCheckbox = positionRow.add("checkbox", undefined, "");
        enabledCheckbox.helpTip = getLabel(enabledTooltipPath);
        enabledCheckbox.value = false;
        var positionLabel = positionRow.add("statictext", undefined, getLabel(labelPath));
        positionLabel.preferredSize.width = POSITION_LABEL_WIDTH;
        var valueInput = positionRow.add("edittext", undefined, initialText);
        valueInput.characters = POSITION_FIELD_CHARS;
        valueInput.helpTip = getLabel(tooltipPath);
        var slider = positionRow.add("slider", undefined, sliderValue, 0, sliderMax);
        slider.preferredSize.width = ROW_SLIDER_WIDTH;
        slider.helpTip = getLabel(tooltipPath);
        return { enabledCheckbox: enabledCheckbox, valueInput: valueInput, slider: slider };
    }

    var pathTextDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
    pathTextDialog.orientation = "column";
    pathTextDialog.alignChildren = ["fill", "top"];
    pathTextDialog.margins = DIALOG_MARGINS;

    /* 2カラムレイアウト / Two-column layout */
    var columnsRow = pathTextDialog.add("group");
    columnsRow.orientation = "row";
    columnsRow.alignChildren = ["fill", "top"];
    columnsRow.alignment = ["fill", "top"];

    var leftCol = addFillGroup(columnsRow, "column");
    var rightCol = addFillGroup(columnsRow, "column");

    /* 処理パネル（パス上文字にする / アーチ生成 / 正円生成） / Process panel */
    var pnlProcess = leftCol.add("panel", undefined, getLabel("panel.process"));
    setupPanel(pnlProcess);

    var rbToPathText = addRadio(pnlProcess, "radio.toPathText", "tooltip.toPathText");
    var rbGenArcPath = addRadio(pnlProcess, "radio.genArcPath", "tooltip.genArcPath");
    var rbGenCircle = addRadio(pnlProcess, "radio.genCircle", "tooltip.genCircle");
    rbToPathText.value = true;

    var btnTextEdit = pnlProcess.add("button", undefined, getLabel("button.textEdit"));
    btnTextEdit.helpTip = getLabel("tooltip.textEdit");
    btnTextEdit.alignment = ["left", "center"];

    /* 分離パネル（処理パネルの下） / Split panel (under Process) */
    var pnlSplit = leftCol.add("panel", undefined, getLabel("panel.split"));
    setupPanel(pnlSplit);

    var rbSplitTextAndPath = addRadio(pnlSplit, "radio.splitKeepFormat", "tooltip.splitKeepFormat");
    var rbSplitTextAndPathNoFormat = addRadio(pnlSplit, "radio.splitNoFormat", "tooltip.splitNoFormat");
    rbSplitTextAndPath.value = false;
    rbSplitTextAndPathNoFormat.value = false;

    var cbSplitDeletePath = pnlSplit.add("checkbox", undefined, getLabel("checkbox.splitDeletePath"));
    cbSplitDeletePath.helpTip = getLabel("tooltip.splitDeletePath");
    cbSplitDeletePath.value = false; /* 既定はオフ / off by default */

    /* 選択に応じて使える処理を切り替える（アーチの既定値は UI をすべて作ってから入れる）
       Enable modes per selection (arc presets are applied once the whole UI exists) */
    var initialAvailability = detectModeAvailability(targetItems, selectedPaths);
    applyModeAvailability(initialAvailability);
    switchToArcIfUnavailable(initialAvailability);

    /* オプションパネル（右カラム） / Option panel (right column) */
    var pnlOption = rightCol.add("panel", undefined, getLabel("panel.option"));
    setupPanel(pnlOption);

    var cbReverse = pnlOption.add("checkbox", undefined, getLabel("checkbox.reverse"));
    cbReverse.helpTip = getLabel("tooltip.reverse");
    cbReverse.value = false;

    /* アーチ方向 / Arc direction */
    var arcDirectionRow = addRow(pnlOption);
    var arcDirectionLabel = arcDirectionRow.add("statictext", undefined, labelText("fieldLabel.arcDirection"));
    arcDirectionLabel.preferredSize.width = ARC_LABEL_WIDTH;

    var rbArcUp = addRadio(arcDirectionRow, "radio.arcUp", "tooltip.arcUp");
    var rbArcDown = addRadio(arcDirectionRow, "radio.arcDown", "tooltip.arcDown");
    rbArcUp.value = true;

    /* まるみ（0=平ら、50=既定、100=最も丸い） / Arc roundness (0 = flat, 50 = default, 100 = roundest) */
    var arcRoundnessRow = addRow(pnlOption);
    arcRoundnessRow.add("statictext", undefined, getLabel("fieldLabel.arcRoundness"));
    var slArcRoundness = arcRoundnessRow.add("slider", undefined, 50, 0, 100);
    slArcRoundness.helpTip = getLabel("tooltip.arcRoundness");
    slArcRoundness.preferredSize.width = ARC_SLIDER_WIDTH;

    /* パス幅パネルの上の余白 / Spacer above the Fit-to-path-width panel */
    var fitWidthSpacer = pnlOption.add("group");
    fitWidthSpacer.preferredSize.height = FIT_PANEL_GAP;

    /* パス幅に合わせるパネル（オプションパネル内の最下部） / Fit-to-path-width panel (bottom of the Options panel) */
    var pnlFitWidth = pnlOption.add("panel", undefined, getLabel("panel.fitWidth"));
    setupPanel(pnlFitWidth);

    var rbFitWidthNone = addRadio(pnlFitWidth, "radio.fitWidthNone", "tooltip.fitWidthNone");
    var rbFitWidthFontSize = addRadio(pnlFitWidth, "radio.fitWidthFontSize", "tooltip.fitWidthFontSize");
    var rbFitWidthTracking = addRadio(pnlFitWidth, "radio.fitWidthTracking", "tooltip.fitWidthTracking");
    rbFitWidthNone.value = true;

    /* 全幅（2カラムの下）：効果 → 位置 → テキスト調整 / Full width: effect -> position -> text adjust */
    var fullWidthColumn = addFillGroup(pathTextDialog, "column");

    /* 効果パネル：5項目を横並びにするため setupPanel（縦並び）は使わない / Effect panel (5 radios in a row) */
    var pnlEffect = fullWidthColumn.add("panel", undefined, getLabel("panel.effect"));
    pnlEffect.orientation = "row";
    pnlEffect.alignChildren = ["center", "center"];
    pnlEffect.alignment = "fill";
    pnlEffect.margins = PANEL_MARGINS;
    pnlEffect.spacing = PANEL_SPACING;

    var rbEffectRainbow = addRadio(pnlEffect, "radio.effectRainbow", "tooltip.effectRainbow");
    var rbEffectDistort = addRadio(pnlEffect, "radio.effectDistort", "tooltip.effectDistort");
    var rbEffectRibbon = addRadio(pnlEffect, "radio.effectRibbon", "tooltip.effectRibbon");
    var rbEffectStep = addRadio(pnlEffect, "radio.effectStep", "tooltip.effectStep");
    var rbEffectGravity = addRadio(pnlEffect, "radio.effectGravity", "tooltip.effectGravity");
    var effectRadios = [rbEffectRainbow, rbEffectDistort, rbEffectRibbon, rbEffectStep, rbEffectGravity];

    /* 既定は効果なし。パス上文字を選んでいるときだけ「虹」 / No effect by default; Rainbow when path text is selected */
    for (var effectIndex = 0; effectIndex < effectRadios.length; effectIndex++) effectRadios[effectIndex].value = false;
    if (initialAvailability.hasPathTextTarget) rbEffectRainbow.value = true;

    /* 位置（開始／終了）パネル / Position (start/end) panel */
    var pnlPosition = fullWidthColumn.add("panel", undefined, getLabel("panel.position"));
    setupPanel(pnlPosition);

    /* スライダーは t 値×100（開始 0.0〜4.0、終了 0.0〜5.0） / sliders hold t × 100 */
    var startPosRow = addPositionRow(pnlPosition, "fieldLabel.startPos", "tooltip.startPosEnabled", "tooltip.startPos", "0.0", 0, 400);
    var cbStartT = startPosRow.enabledCheckbox;
    var etStartT = startPosRow.valueInput;
    var slStartT = startPosRow.slider;
    changeValueByArrowKey(etStartT, false, function () { syncPositionFromEdits(); refreshPreviewIfNeeded(); });

    var endPosRow = addPositionRow(pnlPosition, "fieldLabel.endPos", "tooltip.endPosEnabled", "tooltip.endPos", "1.0", 100, 500);
    var cbEndT = endPosRow.enabledCheckbox;
    var etEndT = endPosRow.valueInput;
    var slEndT = endPosRow.slider;
    changeValueByArrowKey(etEndT, false, function () { syncPositionFromEdits(); refreshPreviewIfNeeded(); });

    /* テキスト調整パネル / Text adjust panel */
    var pnlTextAdjust = fullWidthColumn.add("panel", undefined, getLabel("panel.textAdjust"));
    setupPanel(pnlTextAdjust);

    /* 行揃え（パネル先頭の行） / Alignment (top row) */
    var alignRow = addRow(pnlTextAdjust);
    var alignLabel = alignRow.add("statictext", undefined, labelText("fieldLabel.align"));
    alignLabel.preferredSize.width = ADJUST_LABEL_WIDTH;

    var rbAlignLeft = addRadio(alignRow, "radio.alignLeft", "tooltip.align");
    var rbAlignCenter = addRadio(alignRow, "radio.alignCenter", "tooltip.align");
    var rbAlignRight = addRadio(alignRow, "radio.alignRight", "tooltip.align");
    var rbAlignFullJustify = addRadio(alignRow, "radio.alignFullJustify", "tooltip.align");
    rbAlignCenter.value = true;

    /* ベースライン・文字サイズのスライダーは 0.1pt 単位で、範囲は ±文字サイズに合わせて更新する
       Baseline and font-size sliders use 0.1pt steps; the range follows ±font size */
    var baseShiftRow = addAdjustRow(pnlTextAdjust, "fieldLabel.baseShift", "tooltip.baseShift", "0.0", -1000, 1000);
    var etBaseShift = baseShiftRow.valueInput;
    var slBaseShift = baseShiftRow.slider;

    var trackingRow = addAdjustRow(pnlTextAdjust, "fieldLabel.tracking", "tooltip.tracking", "0", -100, 500);
    var etTracking = trackingRow.valueInput;
    var slTracking = trackingRow.slider;

    var fontSizeRow = addAdjustRow(pnlTextAdjust, "fieldLabel.fontSize", "tooltip.fontSize", "0.0", -1000, 1000);
    var etFontSize = fontSizeRow.valueInput;
    var slFontSize = fontSizeRow.slider;

    changeValueByArrowKey(etBaseShift, true, function () { baseShiftSync.fromEdit(); refreshPreviewIfNeeded(); });
    changeValueByArrowKey(etTracking, true, function () { syncTrackingFromEdit(); refreshPreviewIfNeeded(); });
    changeValueByArrowKey(etFontSize, true, function () { fontSizeSync.fromEdit(); refreshPreviewIfNeeded(); });

    /* ボタンエリア（左＝プレビュー、右＝キャンセル／OK） / Buttons (preview on the left, Cancel/OK on the right) */
    var btnRowGroup = pathTextDialog.add("group");
    btnRowGroup.orientation = "row";
    btnRowGroup.alignChildren = ["fill", "center"];
    btnRowGroup.alignment = ["fill", "top"];

    var btnLeftGroup = btnRowGroup.add("group");
    btnLeftGroup.orientation = "row";
    btnLeftGroup.alignChildren = ["left", "center"];
    btnLeftGroup.alignment = ["left", "center"];

    var cbPreview = btnLeftGroup.add("checkbox", undefined, getLabel("checkbox.preview"));
    cbPreview.helpTip = getLabel("tooltip.preview");
    cbPreview.value = true;

    var btnRightGroup = btnRowGroup.add("group");
    btnRightGroup.orientation = "row";
    btnRightGroup.alignChildren = ["right", "center"];
    btnRightGroup.alignment = ["right", "center"];

    var btnCancel = btnRightGroup.add("button", undefined, getLabel("button.cancel"));
    var btnOK = btnRightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });

    // =========================================
    // ダイアログの状態 / Dialog state
    // =========================================

    /**
     * 「処理」ラジオをすべてオフにする（別パネルの「分離」と手動で排他にする）
     * @returns {void}
     */
    function clearProcessRadios() {
        rbToPathText.value = false;
        rbGenArcPath.value = false;
        rbGenCircle.value = false;
    }

    /**
     * 「分離」ラジオをすべてオフにする（別パネルの「処理」と手動で排他にする）
     * @returns {void}
     */
    function clearSplitRadios() {
        rbSplitTextAndPath.value = false;
        rbSplitTextAndPathNoFormat.value = false;
    }

    /**
     * 使える処理に合わせてラジオの有効／無効を切り替える
     * @param {{canToPathText: boolean, canSplit: boolean}} availability - detectModeAvailability() の結果
     * @returns {void}
     */
    function applyModeAvailability(availability) {
        rbToPathText.enabled = availability.canToPathText;
        rbSplitTextAndPath.enabled = availability.canSplit;
        rbSplitTextAndPathNoFormat.enabled = availability.canSplit;
        cbSplitDeletePath.enabled = availability.canSplit;
        rbGenCircle.enabled = true;
    }

    /**
     * 選んでいる処理が使えないときだけ「アーチ状のパスを生成」へ切り替える
     * @param {{canToPathText: boolean, canSplit: boolean}} availability - detectModeAvailability() の結果
     * @returns {boolean} 切り替えたら true
     */
    function switchToArcIfUnavailable(availability) {
        var switched = false;
        if (rbToPathText.value && !availability.canToPathText) {
            rbToPathText.value = false;
            rbGenArcPath.value = true;
            switched = true;
        }
        if ((rbSplitTextAndPath.value || rbSplitTextAndPathNoFormat.value) && !availability.canSplit) {
            clearSplitRadios();
            rbGenArcPath.value = true;
            switched = true;
        }
        return switched;
    }

    /**
     * アーチモードの既定値（中央揃え・虹・文字サイズでフィット）を入れる
     * @returns {void}
     */
    function applyArcModePresets() {
        rbAlignCenter.value = true;
        rbEffectRainbow.value = true;
        rbFitWidthFontSize.value = true;
    }

    /**
     * 処理に応じてパネルの有効／無効を切り替える
     * （分離では処理・分離以外を無効に。正円ではオープンパス用のパス幅・アーチ方向・まるみも無効に）
     * @returns {void}
     */
    function updatePanelsByMode() {
        var isSplitMode = rbSplitTextAndPath.value || rbSplitTextAndPathNoFormat.value;
        var isCircleMode = rbGenCircle.value;
        var isAdjustable = !isSplitMode;

        pnlProcess.enabled = true;
        pnlSplit.enabled = true;
        pnlOption.enabled = isAdjustable;
        pnlEffect.enabled = isAdjustable;
        pnlTextAdjust.enabled = isAdjustable;
        pnlPosition.enabled = isAdjustable;
        pnlFitWidth.enabled = isAdjustable && !isCircleMode;
        arcDirectionRow.enabled = isAdjustable && !isCircleMode;
        arcRoundnessRow.enabled = isAdjustable && !isCircleMode;
    }

    /**
     * 現在の選択から対象テキスト・パスと、処理ラジオの有効状態を取り直す
     * @returns {void}
     */
    function refreshTargetsFromSelection() {
        var currentSelection = readCurrentSelection();
        targetItems = getTargetTextItems(currentSelection);
        selectedPaths = getSelectedPathItems(currentSelection);

        var availability = detectModeAvailability(targetItems, selectedPaths);
        applyModeAvailability(availability);
        if (switchToArcIfUnavailable(availability)) applyArcModePresets();
        updatePanelsByMode();
    }

    // =========================================
    // 数値欄とスライダーの同期 / Field-slider sync
    // =========================================

    /**
     * 基準の文字サイズ（先頭のテキスト）を pt で返す
     * @returns {number} 文字サイズ（取れなければ 100）
     */
    function getReferenceFontSizePt() {
        try {
            var firstTarget = (targetItems && targetItems.length > 0) ? targetItems[0] : null;
            if (firstTarget && firstTarget.typename === "TextFrame") {
                /* 先頭の textRange を優先し、無ければ全体 / Prefer the first textRange, fall back to the whole range */
                if (firstTarget.textRanges && firstTarget.textRanges.length > 0) {
                    var firstRangeSize = firstTarget.textRanges[0].characterAttributes.size;
                    if (firstRangeSize && !isNaN(firstRangeSize)) return firstRangeSize;
                }
                var fallbackSize = firstTarget.textRange.characterAttributes.size;
                if (fallbackSize && !isNaN(fallbackSize)) return fallbackSize;
            }
        } catch (e) { }
        return 100;
    }

    /**
     * 増減スライダーの範囲を ±文字サイズ（0.1pt 単位）にする
     * @param {Slider} deltaSlider - 対象のスライダー
     * @returns {void}
     */
    function updateDeltaSliderRange(deltaSlider) {
        var sliderLimit = Math.max(1, Math.round(getReferenceFontSizePt() * 10));
        deltaSlider.minvalue = -sliderLimit;
        deltaSlider.maxvalue = sliderLimit;
    }

    /**
     * 0.1pt 単位の増減欄とスライダーを相互に同期する関数の組を作る（ベースライン・文字サイズ用）
     * @param {EditText} deltaInput - 入力欄
     * @param {Slider} deltaSlider - スライダー（値＝増減×10）
     * @returns {{fromEdit: Function, fromSlider: Function}} 入力欄→スライダー、スライダー→入力欄の同期
     */
    function createDeltaSliderSync(deltaInput, deltaSlider) {
        var isSyncing = false;
        return {
            fromEdit: function () {
                if (isSyncing) return;
                isSyncing = true;
                updateDeltaSliderRange(deltaSlider);
                /* スライダーの範囲に収める / clamp to the slider range */
                var deltaValue = clampNumber(parseNumberOr(deltaInput.text, 0), deltaSlider.minvalue / 10.0, deltaSlider.maxvalue / 10.0);
                deltaInput.text = formatOneDecimal(deltaValue);
                deltaSlider.value = Math.round(deltaValue * 10);
                isSyncing = false;
            },
            fromSlider: function () {
                if (isSyncing) return;
                isSyncing = true;
                updateDeltaSliderRange(deltaSlider);
                deltaInput.text = formatOneDecimal(deltaSlider.value / 10.0);
                isSyncing = false;
            }
        };
    }

    var baseShiftSync = createDeltaSliderSync(etBaseShift, slBaseShift);
    var fontSizeSync = createDeltaSliderSync(etFontSize, slFontSize);

    var isTrackingSyncing = false;

    /**
     * 入力欄からトラッキングのスライダーへ同期する（-100〜500 の整数）
     * @returns {void}
     */
    function syncTrackingFromEdit() {
        if (isTrackingSyncing) return;
        isTrackingSyncing = true;
        var trackingValue = Math.round(parseNumberOr(etTracking.text, 0));
        if (trackingValue > 500) trackingValue = 500;
        if (trackingValue < -100) trackingValue = -100;
        etTracking.text = String(trackingValue);
        slTracking.value = trackingValue;
        isTrackingSyncing = false;
    }

    /**
     * スライダーからトラッキングの入力欄へ同期する
     * @returns {void}
     */
    function syncTrackingFromSlider() {
        if (isTrackingSyncing) return;
        isTrackingSyncing = true;
        etTracking.text = String(Math.round(slTracking.value));
        isTrackingSyncing = false;
    }

    var isPositionSyncing = false;

    /**
     * チェックボックスに応じて開始／終了位置の入力を有効にする
     * @returns {void}
     */
    function updatePositionInputsEnabled() {
        etStartT.enabled = cbStartT.value;
        slStartT.enabled = cbStartT.value;
        etEndT.enabled = cbEndT.value;
        slEndT.enabled = cbEndT.value;
    }

    /**
     * 開始／終了位置の値を表示欄の範囲に収めて書き戻す（オフの側は既定値 0.0 / 1.0）
     * @param {number} startValue - 開始位置
     * @param {number} endValue - 終了位置
     * @returns {{startValue: number, endValue: number}} 収めた値
     */
    function writePositionFields(startValue, endValue) {
        startValue = clampNumber(startValue, 0.0, 4.0);
        endValue = clampNumber(endValue, 0.0, 5.0);
        if (!cbStartT.value) startValue = 0.0;
        if (!cbEndT.value) endValue = 1.0;
        etStartT.text = formatOneDecimal(startValue);
        etEndT.text = formatOneDecimal(endValue);
        return { startValue: startValue, endValue: endValue };
    }

    /**
     * 入力欄から開始／終了位置のスライダーへ同期する
     * @returns {void}
     */
    function syncPositionFromEdits() {
        if (isPositionSyncing) return;
        isPositionSyncing = true;
        updatePositionInputsEnabled();
        var startValue = cbStartT.value ? parseTValue(etStartT.text, 0.0) : 0.0;
        var endValue = cbEndT.value ? parseTValue(etEndT.text, 1.0) : 1.0;
        var clampedValues = writePositionFields(startValue, endValue);
        slStartT.value = Math.round(clampedValues.startValue * 100);
        slEndT.value = Math.round(clampedValues.endValue * 100);
        isPositionSyncing = false;
    }

    /**
     * スライダーから開始／終了位置の入力欄へ同期する
     * @returns {void}
     */
    function syncPositionFromSliders() {
        if (isPositionSyncing) return;
        isPositionSyncing = true;
        updatePositionInputsEnabled();
        var startValue = cbStartT.value ? (slStartT.value / 100.0) : 0.0;
        var endValue = cbEndT.value ? (slEndT.value / 100.0) : 1.0;
        writePositionFields(startValue, endValue);
        isPositionSyncing = false;
    }

    // =========================================
    // プレビュー（取り消しを使わない） / Preview (no undo)
    // =========================================
    var previewTempItems = [];        /* プレビューで作ったもの / items created during preview */
    var previewHiddenOriginals = [];  /* プレビュー中に隠した元のオブジェクト / originals hidden during preview */
    var previewPathStates = [];       /* { item, stroked, filled, strokeWidth, strokeColor, opacity } */

    /**
     * プレビューで作った一時オブジェクトを削除し、元の状態へ戻す
     * @returns {void}
     */
    function clearPreview() {
        for (var i = previewTempItems.length - 1; i >= 0; i--) {
            try { previewTempItems[i].remove(); } catch (e) { }
        }
        previewTempItems = [];

        for (var j = previewHiddenOriginals.length - 1; j >= 0; j--) {
            try { previewHiddenOriginals[j].hidden = false; } catch (e) { }
        }
        /* プレビュー中に変えたパスの見た目を戻す / Restore path appearance changed during preview */
        for (var k = previewPathStates.length - 1; k >= 0; k--) {
            try {
                var pathState = previewPathStates[k];
                pathState.item.stroked = pathState.stroked;
                pathState.item.filled = pathState.filled;
                pathState.item.strokeWidth = pathState.strokeWidth;
                pathState.item.strokeColor = pathState.strokeColor;
                pathState.item.opacity = pathState.opacity;
            } catch (e) { }
        }
        previewPathStates = [];
        previewHiddenOriginals = [];
    }

    /**
     * プレビューの復元用に、パスの元の見た目を控える
     * @param {PathItem} pathItem - 対象のパス
     * @returns {void}
     */
    function recordPreviewPathState(pathItem) {
        try {
            if (!pathItem) return;
            for (var i = 0; i < previewPathStates.length; i++) {
                if (previewPathStates[i].item === pathItem) return;
            }
            previewPathStates.push({
                item: pathItem,
                stroked: pathItem.stroked,
                filled: pathItem.filled,
                strokeWidth: pathItem.strokeWidth,
                strokeColor: pathItem.strokeColor,
                opacity: pathItem.opacity
            });
        } catch (e) { }
    }

    /**
     * 複合パスも展開して、各 PathItem に処理を適用する
     * @param {PageItem} pathItem - PathItem または CompoundPathItem
     * @param {Function} styleFunction - 各 PathItem を受け取る関数
     * @returns {void}
     */
    function forEachPathItem(pathItem, styleFunction) {
        try {
            if (!pathItem || !styleFunction) return;
            if (pathItem.typename === "CompoundPathItem") {
                for (var i = 0; i < pathItem.pathItems.length; i++) {
                    try { styleFunction(pathItem.pathItems[i]); } catch (e) { }
                }
            } else {
                try { styleFunction(pathItem); } catch (e) { }
            }
        } catch (e) { }
    }

    /**
     * パスを見えるガイド（塗りなし・黒1pt・不透明度50%）にする
     * @param {PathItem} pathItem - 対象のパス
     * @returns {void}
     */
    function styleVisiblePath(pathItem) {
        pathItem.filled = false;
        pathItem.stroked = true;
        pathItem.strokeWidth = 1;
        pathItem.strokeColor = makeBlackCMYK();
        pathItem.opacity = 50;
    }

    /**
     * パスを見えない状態（塗り・線なし）にする
     * @param {PathItem} pathItem - 対象のパス
     * @returns {void}
     */
    function styleInvisiblePath(pathItem) {
        pathItem.filled = false;
        pathItem.stroked = false;
        pathItem.strokeWidth = 0;
    }

    /**
     * プレビュー：元の見た目を控えてから、見えるガイドにする
     * @param {PageItem} basePath - 対象のパス
     * @returns {void}
     */
    function applyPreviewPathStyle(basePath) {
        forEachPathItem(basePath, function (pathItem) {
            recordPreviewPathState(pathItem);
            styleVisiblePath(pathItem);
        });
    }

    /**
     * 実行：パスを見えるガイドのまま残す（戻さない）
     * @param {PageItem} pathItem - 対象のパス
     * @returns {void}
     */
    function applyExecutePathStyle(pathItem) {
        forEachPathItem(pathItem, styleVisiblePath);
    }

    /**
     * 生成したパスを見えないガイド（線幅0）にする
     * @param {PageItem} pathItem - 対象のパス
     * @returns {void}
     */
    function applyInvisiblePathStyle(pathItem) {
        forEachPathItem(pathItem, styleInvisiblePath);
    }

    /**
     * 元のオブジェクトをプレビュー中だけ隠す
     * @param {PageItem} originalItem - 対象のオブジェクト
     * @returns {void}
     */
    function hideOriginalForPreview(originalItem) {
        /* ロック中のオブジェクトは hidden の代入が例外になりうる / hiding a locked item can throw */
        try {
            if (!originalItem) return;
            for (var i = 0; i < previewHiddenOriginals.length; i++) {
                if (previewHiddenOriginals[i] === originalItem) return;
            }
            originalItem.hidden = true;
            previewHiddenOriginals.push(originalItem);
        } catch (e) { }
    }

    /**
     * 選んでいる処理を実行する
     * @param {boolean} showAlerts - 失敗を alert で知らせるか
     * @param {boolean} previewMode - プレビューとして実行するか
     * @returns {void}
     */
    function runSelectedMode(showAlerts, previewMode) {
        if (rbSplitTextAndPath.value) {
            splitPathTextAndPath(showAlerts, previewMode, true);
        } else if (rbSplitTextAndPathNoFormat.value) {
            splitPathTextAndPath(showAlerts, previewMode, false);
        } else if (rbToPathText.value) {
            createPathTextOnSelectedPath(showAlerts, previewMode);
        } else if (rbGenArcPath.value) {
            createPathTextOnArc(showAlerts, previewMode);
        } else if (rbGenCircle.value) {
            createPathTextOnCircle(showAlerts, previewMode);
        }
    }

    /**
     * 現在の設定でプレビューを作り直す（取り消しは使わない）
     * @returns {void}
     */
    function rebuildPreview() {
        clearPreview();
        /* 選択が変わってもプレビューが安定するよう、最初の選択に戻す / Restore the base selection for a stable preview */
        try { doc.selection = baseSelection; } catch (e) { }
        refreshTargetsFromSelection();

        if (!targetItems || targetItems.length === 0) {
            cbPreview.value = false;
            return;
        }

        runSelectedMode(false, true);
        app.redraw();
    }

    /**
     * プレビューがオンのときだけ作り直す
     * @returns {void}
     */
    function refreshPreviewIfNeeded() {
        if (cbPreview.value) {
            rebuildPreview();
        }
    }

    // =========================================
    // テキスト編集ダイアログ / Text edit dialog
    // =========================================

    /**
     * 選択中のテキストの内容を書き換える小さなダイアログを開く
     * @returns {void}
     */
    function showTextEditDialog() {
        var textEditDialog = new Window("dialog", getLabel("dialog.textEditTitle"));
        textEditDialog.orientation = "column";
        textEditDialog.alignChildren = ["fill", "top"];
        textEditDialog.margins = DIALOG_MARGINS;

        var hintLabel = textEditDialog.add("statictext", undefined, getLabel("dialog.textEditHint"), { multiline: true });
        hintLabel.preferredSize.width = EDIT_HINT_WIDTH;

        /* 対象はクリック時の選択から取り直す（古い控えに頼らない） / Resolve targets at click time */
        var editTargets = [];
        /* テキスト編集中の選択は配列でなく、たどると例外になる / a text-editing selection throws when walked */
        try { editTargets = getTargetTextItems(readCurrentSelection()); } catch (e) { editTargets = []; }

        var initialText = "";
        try {
            if (editTargets && editTargets.length > 0 && editTargets[0] && editTargets[0].typename === "TextFrame") {
                initialText = String(editTargets[0].contents);
            }
        } catch (e) { initialText = ""; }

        var textEditInput = textEditDialog.add("edittext", undefined, initialText, { multiline: true });
        textEditInput.preferredSize = EDIT_FIELD_SIZE;

        var editBtnRowGroup = textEditDialog.add("group");
        editBtnRowGroup.orientation = "row";
        editBtnRowGroup.alignChildren = ["center", "center"];
        var btnEditCancel = editBtnRowGroup.add("button", undefined, getLabel("button.cancel"));
        var btnEditOK = editBtnRowGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });

        btnEditCancel.onClick = function () {
            textEditDialog.close(0);
        };
        btnEditOK.onClick = function () {
            var newText = textEditInput.text;
            for (var i = 0; i < editTargets.length; i++) {
                var textFrame = editTargets[i];
                if (!textFrame || textFrame.typename !== "TextFrame") continue;
                try { textFrame.contents = newText; } catch (e) { }
            }
            refreshTargetsFromSelection();
            refreshPreviewIfNeeded();
            app.redraw();
            textEditDialog.close(1);
        };
        textEditDialog.show();
    }

    // =========================================
    // イベント / Event handlers
    // =========================================

    btnTextEdit.onClick = showTextEditDialog;

    cbPreview.onClick = function () {
        if (cbPreview.value) {
            rebuildPreview();
        } else {
            clearPreview();
            app.redraw();
        }
    };

    /**
     * 処理・分離ラジオを押したときの共通処理
     * @param {Function} clearOtherGroup - もう一方のグループのラジオをオフにする関数
     * @param {boolean} isArcMode - アーチモードなら true（既定値を入れる。ほかはパス幅フィットを「しない」に戻す）
     * @returns {void}
     */
    function onModeRadioClick(clearOtherGroup, isArcMode) {
        clearOtherGroup();
        if (!isArcMode) rbFitWidthNone.value = true;
        updatePanelsByMode();
        if (isArcMode) applyArcModePresets();
        refreshPreviewIfNeeded();
    }

    rbToPathText.onClick = function () { onModeRadioClick(clearSplitRadios, false); };
    rbGenCircle.onClick = function () { onModeRadioClick(clearSplitRadios, false); };
    rbGenArcPath.onClick = function () { onModeRadioClick(clearSplitRadios, true); };
    rbSplitTextAndPath.onClick = function () { onModeRadioClick(clearProcessRadios, false); };
    rbSplitTextAndPathNoFormat.onClick = function () { onModeRadioClick(clearProcessRadios, false); };

    /* 押したらプレビューを作り直すだけのコントロール / Controls that only refresh the preview */
    var previewRefreshControls = [
        cbReverse,
        rbEffectRainbow, rbEffectDistort, rbEffectRibbon, rbEffectStep, rbEffectGravity,
        rbAlignLeft, rbAlignCenter, rbAlignRight, rbAlignFullJustify,
        rbArcUp, rbArcDown,
        rbFitWidthNone, rbFitWidthFontSize, rbFitWidthTracking
    ];
    for (var controlIndex = 0; controlIndex < previewRefreshControls.length; controlIndex++) {
        previewRefreshControls[controlIndex].onClick = refreshPreviewIfNeeded;
    }

    /* まるみ：スライダーを離したら更新 / Roundness: refresh on release */
    slArcRoundness.onChange = refreshPreviewIfNeeded;

    etBaseShift.onChanging = function () { baseShiftSync.fromEdit(); refreshPreviewIfNeeded(); };
    slBaseShift.onChanging = function () { baseShiftSync.fromSlider(); };
    slBaseShift.onChange = function () { baseShiftSync.fromSlider(); refreshPreviewIfNeeded(); };
    baseShiftSync.fromEdit();

    etTracking.onChanging = function () { syncTrackingFromEdit(); refreshPreviewIfNeeded(); };
    slTracking.onChanging = function () { syncTrackingFromSlider(); };
    slTracking.onChange = function () { syncTrackingFromSlider(); refreshPreviewIfNeeded(); };
    syncTrackingFromEdit();

    etFontSize.onChanging = function () { fontSizeSync.fromEdit(); refreshPreviewIfNeeded(); };
    slFontSize.onChanging = function () { fontSizeSync.fromSlider(); };
    slFontSize.onChange = function () { fontSizeSync.fromSlider(); refreshPreviewIfNeeded(); };
    fontSizeSync.fromEdit();

    etStartT.onChanging = function () { syncPositionFromEdits(); refreshPreviewIfNeeded(); };
    etEndT.onChanging = function () { syncPositionFromEdits(); refreshPreviewIfNeeded(); };
    slStartT.onChanging = function () { syncPositionFromSliders(); };
    slStartT.onChange = function () { syncPositionFromSliders(); refreshPreviewIfNeeded(); };
    slEndT.onChanging = function () { syncPositionFromSliders(); };
    slEndT.onChange = function () { syncPositionFromSliders(); refreshPreviewIfNeeded(); };

    cbStartT.onClick = function () { syncPositionFromEdits(); refreshPreviewIfNeeded(); };
    cbEndT.onClick = function () { syncPositionFromEdits(); refreshPreviewIfNeeded(); };

    syncPositionFromEdits();

    btnCancel.onClick = function () {
        clearPreview();
        pathTextDialog.close(0);
    };

    btnOK.onClick = function () {
        /* プレビューが出ていれば先に消し、一時オブジェクトを重ねない / Clear the preview first so temp items do not stack */
        if (cbPreview.value) {
            clearPreview();
        }
        runSelectedMode(true, false);
        pathTextDialog.close(1);
    };

    // =========================================
    // メイン処理 / Main
    // =========================================

    /* 起動時にアーチモードへ切り替わっていたら、UI がそろったここで既定値を入れる
       If arc mode was chosen at startup, apply its presets now that the whole UI exists */
    if (rbGenArcPath.value) applyArcModePresets();

    updatePanelsByMode();

    /* 開いた直後に一度プレビュー（取り消しは使わない） / Preview once on open (no undo) */
    /* 変換処理の DOM 例外でダイアログが開かなくならないように / keep the dialog opening even if the preview throws */
    try {
        if (cbPreview.value) {
            rebuildPreview();
            app.redraw();
        }
    } catch (e) { }

    var dialogResult = pathTextDialog.show();
    if (dialogResult !== 1) return false;
    return true;

    // =========================================
    // 変換処理 / Conversion
    // =========================================

    /**
     * 選んでいる効果をメニューコマンドでパス上文字に適用する
     * @param {TextFrame} textFrame - 対象のパス上文字
     * @returns {void}
     */
    function applyPathTextEffect(textFrame) {
        var effectCommand = getSelectedEffectCommand();
        if (!effectCommand) return;

        var previousSelection = null;
        try { previousSelection = doc.selection; } catch (e) { previousSelection = null; }

        try {
            try { doc.selection = []; } catch (e) { }
            try { textFrame.selected = true; } catch (e) { }
            app.executeMenuCommand(effectCommand);
        } catch (e) { }

        /* 選択は必ず戻す / Always restore the selection */
        try { doc.selection = previousSelection; } catch (e) { }
    }

    /**
     * 選んでいる効果に対応するメニューコマンド名を返す
     * @returns {string|null} メニューコマンド名（効果なしなら null）
     */
    function getSelectedEffectCommand() {
        if (rbEffectRainbow.value) return "Rainbow";
        if (rbEffectDistort.value) return "Skew";
        if (rbEffectRibbon.value) return "3D ribbon";
        if (rbEffectStep.value) return "Stair Step";
        if (rbEffectGravity.value) return "Gravity";
        return null;
    }

    /**
     * ［行揃え］の選択を Justification に変換する（左揃えは「最終行左揃え」）
     * @returns {Justification} 行揃え
     */
    function readSelectedJustification() {
        if (rbAlignFullJustify.value) return Justification.FULLJUSTIFY;
        if (rbAlignRight.value) return Justification.RIGHT;
        if (rbAlignCenter.value) return Justification.CENTER;
        return Justification.FULLJUSTIFYLASTLINELEFT;
    }

    /**
     * テキストフレームの全段落と全体に行揃えを設定する（内容を複製したあとに呼ぶ）
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @param {Justification} justification - 行揃え
     * @returns {void}
     */
    function setFrameJustification(textFrame, justification) {
        if (!textFrame) return;
        try {
            if (textFrame.paragraphs && textFrame.paragraphs.length > 0) {
                for (var i = 0; i < textFrame.paragraphs.length; i++) {
                    try { textFrame.paragraphs[i].paragraphAttributes.justification = justification; } catch (e) { }
                }
            }
        } catch (e) { }
        /* 念のため全体にも / Also apply to the whole range */
        try { textFrame.textRange.paragraphAttributes.justification = justification; } catch (e) { }
    }

    /**
     * UI の開始／終了位置をパス上文字へ適用する
     * @param {TextFrame} textFrame - 対象のパス上文字
     * @returns {void}
     */
    function applyStartEndTValue(textFrame) {
        try {
            if (!textFrame) return;
            if (cbStartT.value) textFrame.startTValue = parseTValue(etStartT.text, 0.0);
            if (cbEndT.value) textFrame.endTValue = parseTValue(etEndT.text, 1.0);
        } catch (e) { }
    }

    /**
     * テキストフレームの textRange を配列で返す（無ければ textRange 単体）
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @returns {TextRange[]} textRange の配列
     */
    function collectTextRanges(textFrame) {
        var textRanges = [];
        try {
            if (textFrame.textRanges && textFrame.textRanges.length > 0) {
                for (var i = 0; i < textFrame.textRanges.length; i++) textRanges.push(textFrame.textRanges[i]);
            }
        } catch (e) { }
        if (textRanges.length === 0) {
            try { if (textFrame.textRange) textRanges = [textFrame.textRange]; } catch (e) { textRanges = []; }
        }
        return textRanges;
    }

    /**
     * 各 textRange の文字属性に増減を加える（clampMin があれば下限を適用）
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @param {string} attributeName - 文字属性の名前
     * @param {number} delta - 増減
     * @param {number} [clampMin] - 下限
     * @returns {void}
     */
    function addDeltaToTextRanges(textFrame, attributeName, delta, clampMin) {
        if (!textFrame || !delta) return;
        var textRanges = collectTextRanges(textFrame);
        for (var i = 0; i < textRanges.length; i++) {
            try {
                var attributes = textRanges[i].characterAttributes;
                var nextValue = attributes[attributeName] + delta;
                if (typeof clampMin === "number" && nextValue < clampMin) nextValue = clampMin;
                attributes[attributeName] = nextValue;
            } catch (e) { }
        }
    }

    /**
     * 各 textRange の文字サイズに［文字サイズ］の増減を加える（最小 0.1pt）
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @returns {void}
     */
    function applyFontSizeDelta(textFrame) {
        addDeltaToTextRanges(textFrame, "size", parseNumberOr(etFontSize.text, 0), 0.1);
    }

    /**
     * 各 textRange のベースラインとトラッキングに［テキスト調整］の増減を加える（行揃えのあとに呼ぶ）
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @returns {void}
     */
    function applyShiftAndTrackingDelta(textFrame) {
        addDeltaToTextRanges(textFrame, "baselineShift", parseNumberOr(etBaseShift.text, 0));
        addDeltaToTextRanges(textFrame, "tracking", Math.round(parseNumberOr(etTracking.text, 0)));
    }

    /**
     * 指定のパス上にパス上文字を作り、元のテキストの直前へ置く
     * @param {PathItem} textPath - 使うパス
     * @param {TextFrame} originalText - 元のテキスト
     * @param {Layer} currentLayer - 作成先のレイヤー
     * @param {boolean} previewMode - プレビューなら true
     * @returns {TextFrame} 作成したパス上文字
     */
    function createPathTextFrame(textPath, originalText, currentLayer, previewMode) {
        var textOnAPath = currentLayer.textFrames.pathText(textPath);
        /* 重ね順を保つ（ほかのオブジェクトの背面に回り込まないように） / Keep stacking order */
        try { textOnAPath.move(originalText, ElementPlacement.PLACEBEFORE); } catch (e) { }
        if (previewMode) previewTempItems.push(textOnAPath);
        return textOnAPath;
    }

    /**
     * パス上文字の共通の仕上げ：内側配置・効果・内容の複製・各種調整を適用し、元のテキストを隠す／削除する
     * @param {TextFrame} textOnAPath - 作成したパス上文字
     * @param {TextFrame} originalText - 元のテキスト
     * @param {boolean} previewMode - プレビューなら true
     * @param {{forceCenter: boolean, circleInsideShift: boolean}} [decorateOptions] - 正円用の指定
     * @returns {void}
     */
    function decoratePathText(textOnAPath, originalText, previewMode, decorateOptions) {
        decorateOptions = decorateOptions || {};

        /* 内側配置 / Inside placement */
        if (cbReverse.value) {
            try { textOnAPath.textPath.polarity = PolarityValues.NEGATIVE; } catch (e) { }
        }

        applyPathTextEffect(textOnAPath);
        applyStartEndTValue(textOnAPath);

        /* 正円＋内側：開始／終了位置を +0.5 ずらして見かけの開始位置を合わせる（上限は 5、開始は 1 を超えてよい）
           Circle + inside: shift start/end by +0.5 (clamped to the global cap of 5) */
        if (decorateOptions.circleInsideShift) {
            try {
                var shiftAmount = 0.5;
                textOnAPath.startTValue = clampNumber(textOnAPath.startTValue + shiftAmount, 0.0, 5.0);
                textOnAPath.endTValue = clampNumber(textOnAPath.endTValue + shiftAmount, 0.0, 5.0);
            } catch (e) { }
        }

        /* テキストの内容を元から複製 / Copy the text content from the original */
        for (var i = 0; i < originalText.textRanges.length; i++) {
            originalText.textRanges[i].duplicate(textOnAPath);
        }

        applyFontSizeDelta(textOnAPath);

        /* 行揃え（複製のあとでないと上書きされる） / Justification (apply after the content is copied) */
        setFrameJustification(textOnAPath, decorateOptions.forceCenter ? Justification.CENTER : readSelectedJustification());

        applyShiftAndTrackingDelta(textOnAPath);

        if (previewMode) {
            hideOriginalForPreview(originalText);
        } else {
            try { originalText.remove(); } catch (e) { }
            /* 実行時は作ったパス上文字を選択 / Select the created path text on execute */
            try { textOnAPath.selected = true; } catch (e) { }
        }
    }

    /**
     * 各テキストに使う元のパスを決める（選択したパスを優先し、無ければ自分の textPath）
     * @param {PageItem[]} pathItems - 選択したパス
     * @param {number} textCount - テキストの数
     * @param {number} textIndex - このテキストの位置
     * @param {TextFrame} originalText - 元のテキスト
     * @returns {PageItem|null} 使うパス
     */
    function resolveBasePathForText(pathItems, textCount, textIndex, originalText) {
        if (pathItems && pathItems.length > 0) {
            if (pathItems.length === textCount) return pathItems[textIndex];
            return pathItems[0];
        }
        try {
            if (originalText && originalText.typename === "TextFrame") {
                if (originalText.kind === TextType.PATHTEXT && originalText.textPath) {
                    return originalText.textPath;
                }
            }
        } catch (e) { }
        return null;
    }

    /**
     * 選択した元のパスを見えなくする（塗り・線なし）
     * @param {PageItem} basePath - 対象のパス
     * @returns {void}
     */
    function makeOriginalPathInvisible(basePath) {
        try {
            if (basePath.typename === "CompoundPathItem") {
                for (var i = 0; i < basePath.pathItems.length; i++) {
                    basePath.pathItems[i].stroked = false;
                    basePath.pathItems[i].filled = false;
                }
            } else {
                basePath.stroked = false;
                basePath.filled = false;
            }
        } catch (e) { }
    }

    /**
     * 複製に使う PathItem を返す（複合パスは先頭）
     * @param {PageItem} basePath - 対象のパス
     * @returns {PathItem|null} 複製に使うパス
     */
    function resolveSourcePath(basePath) {
        try {
            return (basePath.typename === "CompoundPathItem") ? basePath.pathItems[0] : basePath;
        } catch (e) {
            return null;
        }
    }

    /**
     * ハンドルとポイントの種類を保ったままパスを複製する
     * @param {PathItem} originalPath - 元のパス
     * @param {Layer} targetLayer - 作成先のレイヤー
     * @returns {PathItem|null} 複製したパス
     */
    function duplicatePathWithHandles(originalPath, targetLayer) {
        try {
            if (!originalPath) return null;
            var pathContainer = (targetLayer && targetLayer.pathItems) ? targetLayer.pathItems : doc.pathItems;
            var newPath = pathContainer.add();

            for (var i = 0; i < originalPath.pathPoints.length; i++) {
                var originalPoint = originalPath.pathPoints[i];
                var newPoint = newPath.pathPoints.add();
                newPoint.anchor = originalPoint.anchor;
                newPoint.leftDirection = originalPoint.leftDirection;
                newPoint.rightDirection = originalPoint.rightDirection;
                newPoint.pointType = originalPoint.pointType;
            }
            newPath.closed = originalPath.closed;
            return newPath;
        } catch (e) {
            return null;
        }
    }

    /**
     * テキスト用にパスを複製し、同じレイヤーへ置く
     * @param {PageItem} basePath - 元のパス
     * @param {Layer} currentLayer - 置き先のレイヤー
     * @returns {PathItem|null} 複製したパス
     */
    function duplicatePathForText(basePath, currentLayer) {
        var sourcePath = resolveSourcePath(basePath);
        if (!sourcePath) return null;

        /* 1) まず標準の複製 / Try the native duplicate first */
        try {
            var duplicatedPath = sourcePath.duplicate();
            try { duplicatedPath.move(currentLayer, ElementPlacement.PLACEATBEGINNING); } catch (e) { }
            return duplicatedPath;
        } catch (e) {
            /* 2) だめならアンカーとハンドルを写して作る（失敗すると null） / Fall back to copying anchors and handles */
            return duplicatePathWithHandles(sourcePath, currentLayer);
        }
    }

    /**
     * 処理：選択したパスに沿ってパス上文字を作る
     * @param {boolean} showAlerts - 失敗を alert で知らせるか
     * @param {boolean} previewMode - プレビューなら true
     * @returns {void}
     */
    function createPathTextOnSelectedPath(showAlerts, previewMode) {
        var createdTexts = [];

        for (var i = 0; i < targetItems.length; i++) {
            var originalText = targetItems[i];
            var currentLayer = originalText.layer;

            var basePath = resolveBasePathForText(selectedPaths, targetItems.length, i, originalText);
            if (!basePath) {
                /* 事前に確かめているので通常は来ない / Normally unreachable after the availability check */
                if (showAlerts) alert(getLabel("alert.needPath"));
                return;
            }

            if (!previewMode) {
                makeOriginalPathInvisible(basePath);
            } else {
                /* プレビュー：元のパスを見えるガイドにする（プレビューを消すと戻る） / Preview: show the base path as a guide */
                applyPreviewPathStyle(basePath);
            }

            var textPath = duplicatePathForText(basePath, currentLayer);
            if (!textPath) {
                if (showAlerts) alert(getLabel("alert.dupPathFail"));
                continue;
            }
            if (previewMode) {
                previewTempItems.push(textPath);
            } else {
                applyExecutePathStyle(textPath);
            }

            var textOnAPath = createPathTextFrame(textPath, originalText, currentLayer, previewMode);
            decoratePathText(textOnAPath, originalText, previewMode);
            createdTexts.push(textOnAPath);
        }
        applyFitToPathWidth(createdTexts);
    }

    /**
     * アーチ／正円モードで一緒に選択したパスを、プレビューでは隠し、実行では削除する
     * @param {boolean} previewMode - プレビューなら true
     * @returns {void}
     */
    function handleArcModeSelectedPaths(previewMode) {
        if (!selectedPaths || selectedPaths.length === 0) return;
        for (var i = selectedPaths.length - 1; i >= 0; i--) {
            var selectedPath = selectedPaths[i];
            if (!selectedPath) continue;
            if (previewMode) {
                hideOriginalForPreview(selectedPath);
            } else {
                try { selectedPath.remove(); } catch (e) { }
            }
        }
    }

    /**
     * 処理：アーチ状のパスを作ってパス上文字にする
     * @param {boolean} showAlerts - 失敗を alert で知らせるか
     * @param {boolean} previewMode - プレビューなら true
     * @returns {void}
     */
    function createPathTextOnArc(showAlerts, previewMode) {
        var createdTexts = [];
        handleArcModeSelectedPaths(previewMode);

        for (var i = 0; i < targetItems.length; i++) {
            var originalText = targetItems[i];
            var currentLayer = originalText.layer;

            var textPath = createArcPathFromText(originalText, currentLayer);
            if (!textPath) {
                if (showAlerts) alert(getLabel("alert.arcFail"));
                continue;
            }

            /* 生成したアーチは見えないガイド / The generated arc is an invisible guide */
            applyInvisiblePathStyle(textPath);
            if (previewMode) previewTempItems.push(textPath);

            var textOnAPath = createPathTextFrame(textPath, originalText, currentLayer, previewMode);

            /* 変換時に Illustrator がスタイルを上書きするので、パス上文字のパスも見えなくする
               Illustrator restyles the path on conversion, so hide the path-text path too */
            try {
                if (textOnAPath.textPath) applyInvisiblePathStyle(textOnAPath.textPath);
            } catch (e) { }

            decoratePathText(textOnAPath, originalText, previewMode);
            createdTexts.push(textOnAPath);
        }
        applyFitToPathWidth(createdTexts);
    }

    /**
     * 処理：テキストの幅を直径とする正円を作ってパス上文字にする
     * @param {boolean} showAlerts - 失敗を alert で知らせるか
     * @param {boolean} previewMode - プレビューなら true
     * @returns {void}
     */
    function createPathTextOnCircle(showAlerts, previewMode) {
        var createdTexts = [];
        handleArcModeSelectedPaths(previewMode);

        for (var i = 0; i < targetItems.length; i++) {
            var originalText = targetItems[i];
            var currentLayer = originalText.layer;

            var textPath = createCirclePathFromText(originalText, currentLayer);
            if (!textPath) {
                if (showAlerts) alert(getLabel("alert.arcFail"));
                continue;
            }

            if (previewMode) {
                applyPreviewPathStyle(textPath);
                previewTempItems.push(textPath);
            } else {
                applyExecutePathStyle(textPath);
            }

            /* 正円を中心まわりに 90° 回す / Rotate the circle 90° around its center */
            try { textPath.rotate(90, true, true, true, true, Transformation.CENTER); } catch (e) { }

            var textOnAPath = createPathTextFrame(textPath, originalText, currentLayer, previewMode);

            /* 変換時に Illustrator がスタイルを上書きするので、パス上文字のパスを整え直す
               Illustrator restyles the path on conversion, so restyle the path-text path */
            try {
                if (textOnAPath.textPath) {
                    if (previewMode) {
                        applyPreviewPathStyle(textOnAPath.textPath);
                    } else {
                        applyExecutePathStyle(textOnAPath.textPath);
                    }
                }
            } catch (e) { }

            /* 正円は常に中央揃え。内側配置のときは開始／終了位置を +0.5 補正 / Circles are always centered */
            decoratePathText(textOnAPath, originalText, previewMode, {
                forceCenter: true,
                circleInsideShift: cbReverse.value
            });
            createdTexts.push(textOnAPath);
        }
        applyFitToPathWidth(createdTexts);
    }

    /**
     * 1件のパス上文字を、パスとポイント文字に分ける
     * @param {TextFrame} pathText - 対象のパス上文字
     * @param {boolean} previewMode - プレビューなら true
     * @param {boolean} keepFormatting - 書式を保つなら true
     * @returns {boolean} 処理したら true
     */
    function splitOnePathText(pathText, previewMode, keepFormatting) {
        if (!pathText || pathText.typename !== "TextFrame") return false;

        var isPathText = false;
        try { isPathText = (pathText.kind === TextType.PATHTEXT); } catch (e) { isPathText = false; }
        if (!isPathText) return false;

        var originalPath = null;
        try { originalPath = pathText.textPath; } catch (e) { originalPath = null; }
        if (!originalPath) return false;

        var pathTextLayer = null;
        try { pathTextLayer = pathText.layer; } catch (e) { pathTextLayer = null; }

        /* 1) 「パスを削除」がオフなら、ハンドルを保ってパスを複製（塗りなし・スミ1pt） / Keep a copy of the path */
        if (!cbSplitDeletePath.value) {
            var newPath = duplicatePathWithHandles(originalPath, pathTextLayer);
            if (newPath) {
                try {
                    newPath.filled = false;
                    newPath.stroked = true;
                    newPath.strokeColor = makeBlackCMYK();
                    newPath.strokeWidth = 1;
                } catch (e) { }
                if (previewMode) previewTempItems.push(newPath);
            }
        }

        /* 2) 元の開始アンカーの位置にポイント文字を作る / Create point text at the first anchor */
        var newText = null;
        try {
            newText = (pathTextLayer ? pathTextLayer.textFrames.add() : doc.textFrames.add());
        } catch (e) {
            newText = null;
        }

        if (newText) {
            try {
                var anchorPoint = originalPath.pathPoints[0].anchor;
                newText.position = [anchorPoint[0], anchorPoint[1]];
            } catch (e) { }

            /* 内容のコピー（書式を保つ／保たない） / Copy the content (with or without formatting) */
            if (keepFormatting) {
                try { pathText.textRange.duplicate(newText); } catch (e) { }
            } else {
                try { newText.contents = pathText.contents; } catch (e) { }
            }

            applyFontSizeDelta(newText);
            applyShiftAndTrackingDelta(newText);

            /* 書式を保つときだけ塗り／線の色を引き継ぐ / Carry over fill/stroke colors only when keeping formatting */
            if (keepFormatting) {
                for (var i = 0; i < pathText.textRanges.length; i++) {
                    try {
                        var sourceAttributes = pathText.textRanges[i].characterAttributes;
                        var newAttributes = newText.textRanges[i].characterAttributes;
                        newAttributes.fillColor = sourceAttributes.fillColor;
                        newAttributes.strokeColor = sourceAttributes.strokeColor;
                    } catch (e) { }
                }
            }

            /* 行揃え（複製のあと）と重ね順の保持 / Justification after the copy, and keep the stacking order */
            setFrameJustification(newText, readSelectedJustification());
            try { newText.move(pathText, ElementPlacement.PLACEBEFORE); } catch (e) { }

            if (previewMode) previewTempItems.push(newText);
        }

        /* 3) ポイント文字を作れたときだけ元のパス上文字を隠す／削除（失敗したら元を残す） / Keep the original on failure */
        if (newText) {
            if (previewMode) {
                hideOriginalForPreview(pathText);
            } else {
                try { pathText.remove(); } catch (e) { }
            }
        }
        return true;
    }

    /**
     * パス上文字を、見えるパスとポイント文字に分ける
     * @param {boolean} showAlerts - 何も処理しなかったとき alert で知らせるか
     * @param {boolean} previewMode - プレビューなら true
     * @param {boolean} keepFormatting - 書式を保つなら true
     * @returns {void}
     */
    function splitPathTextAndPath(showAlerts, previewMode, keepFormatting) {
        var didAny = false;
        for (var i = targetItems.length - 1; i >= 0; i--) {
            if (splitOnePathText(targetItems[i], previewMode, keepFormatting)) didAny = true;
        }

        if (!didAny && showAlerts) {
            alert(getLabel("alert.needPathText"));
        }
    }

    // =========================================
    // パスの生成 / Path generation
    // =========================================

    /**
     * テキストの範囲をアウトライン化して測る（ライブテキストの geometricBounds より安定する）
     * @param {TextFrame} originalText - 対象のテキスト
     * @returns {number[]} [左, 上, 右, 下]
     */
    function measureArcTextBounds(originalText) {
        var measureTexts = [originalText.duplicate(), originalText.duplicate()];
        measureTexts[0].contents = "";

        /* 先頭行の textRange を measureTexts[0] に複製（ポイント文字＋1行以上が前提の旧来の手法）
           Copy the first line's textRanges into measureTexts[0] (legacy approach) */
        for (var i = 0; i < originalText.lines[0].length; i++) {
            originalText.textRanges[i].duplicate(measureTexts[0]);
        }
        for (var j = 0; j < measureTexts.length; j++) {
            measureTexts[j] = measureTexts[j].createOutline();
        }

        var textBounds = measureTexts[1].geometricBounds;
        textBounds[3] = measureTexts[0].geometricBounds[3];

        for (var k = 0; k < measureTexts.length; k++) {
            try { measureTexts[k].remove(); } catch (e) { }
        }
        return textBounds;
    }

    /**
     * ［まるみ］スライダーを読む（0〜100、既定 50）
     * @returns {number} まるみ（%）
     */
    function readArcRoundnessPercent() {
        var percent = Number(slArcRoundness.value);
        if (isNaN(percent)) percent = 50;
        if (percent < 0) percent = 0;
        if (percent > 100) percent = 100;
        return percent;
    }

    /**
     * アーチ方向の符号を返す（上=+1、下=-1）
     * @returns {number} 符号
     */
    function readArcDirectionSign() {
        return rbArcDown.value ? -1 : 1;
    }

    /**
     * 直線のパスのハンドルを動かしてアーチ状に曲げる
     * @param {PathItem} arcPath - 2点の直線パス
     * @param {number} roundnessPercent - まるみ（%）
     * @param {number} directionSign - 上=+1、下=-1
     * @returns {void}
     */
    function applyArcHandles(arcPath, roundnessPercent, directionSign) {
        var handleOffsetDivisors = [3.5, 4];
        var pathLength = arcPath.length;

        /* handleOffsets[0]: 水平ハンドル（固定）/ handleOffsets[1]: 垂直ハンドル＝カーブの深さ
           50% で垂直ハンドルは pathLength / handleOffsetDivisors[1]（旧来の既定） */
        var handleOffsets = [
            pathLength / handleOffsetDivisors[0],
            pathLength * (roundnessPercent / 100) * (2 / handleOffsetDivisors[1])
        ];
        var pathPoints = arcPath.pathPoints;

        pathPoints[0].rightDirection = [
            pathPoints[0].rightDirection[0] + handleOffsets[0],
            pathPoints[0].rightDirection[1] + (directionSign * handleOffsets[1])
        ];
        pathPoints[1].leftDirection = [
            pathPoints[1].leftDirection[0] - handleOffsets[0],
            pathPoints[1].leftDirection[1] + (directionSign * handleOffsets[1])
        ];
    }

    /**
     * テキストのアウトラインの範囲からアーチ状のパスを作る
     * @param {TextFrame} originalText - 対象のテキスト
     * @param {Layer} targetLayer - 作成先のレイヤー
     * @returns {PathItem|null} 作ったパス
     */
    function createArcPathFromText(originalText, targetLayer) {
        try {
            if (!originalText || originalText.typename !== "TextFrame") return null;
            if (!originalText.lines || originalText.lines.length === 0) return null;
            if (!originalText.textRanges || originalText.textRanges.length === 0) return null;

            var textBounds = measureArcTextBounds(originalText);

            /* ベースラインに沿った直線を作る（1.02 は旧来の既定） / Straight line along the baseline */
            var baselineY = textBounds[3] * 1.02;
            var arcPath = targetLayer.pathItems.add();
            arcPath.setEntirePath([
                [textBounds[0], baselineY],
                [textBounds[2], baselineY]
            ]);

            try {
                arcPath.stroked = false;
                arcPath.filled = false;
            } catch (e) { }

            applyArcHandles(arcPath, readArcRoundnessPercent(), readArcDirectionSign());
            return arcPath;
        } catch (e) {
            return null;
        }
    }

    /**
     * テキストの幅を直径とする正円のパスを作る
     * @param {TextFrame} originalText - 対象のテキスト
     * @param {Layer} targetLayer - 作成先のレイヤー
     * @returns {PathItem|null} 作ったパス
     */
    function createCirclePathFromText(originalText, targetLayer) {
        try {
            if (!originalText || originalText.typename !== "TextFrame") return null;
            if (!originalText.textRanges || originalText.textRanges.length === 0) return null;

            /* 安定させるためアウトラインの範囲で測る / Measure the outline for stability */
            var measureText = originalText.duplicate();
            var measureOutline = null;
            try { measureOutline = measureText.createOutline(); } catch (e) { measureOutline = null; }
            /* createOutline() は複製を消費するので通常ここは例外になる / createOutline() consumes the duplicate */
            try { measureText.remove(); } catch (e) { }
            if (!measureOutline) return null;

            var outlineBounds = measureOutline.geometricBounds;
            try { measureOutline.remove(); } catch (e) { }

            var textWidth = outlineBounds[2] - outlineBounds[0];
            if (!textWidth || isNaN(textWidth) || textWidth <= 0) return null;

            var centerX = (outlineBounds[0] + outlineBounds[2]) / 2;
            var centerY = (outlineBounds[1] + outlineBounds[3]) / 2;

            var diameter = textWidth;
            var circlePath = targetLayer.pathItems.ellipse(centerY + (diameter / 2), centerX - (diameter / 2), diameter, diameter);

            /* 既定は見えない状態（プレビューの見た目は呼び出し側で付ける） / Invisible by default */
            try {
                circlePath.stroked = false;
                circlePath.filled = false;
            } catch (e) { }

            return circlePath;
        } catch (e) {
            return null;
        }
    }

    // =========================================
    // パス幅に合わせる / Fit to path width
    // =========================================

    /**
     * ［パス幅に合わせる］の選択に応じてフィットする
     * @param {TextFrame[]} textFrames - 作成したパス上文字
     * @returns {void}
     */
    function applyFitToPathWidth(textFrames) {
        if (rbFitWidthFontSize.value) {
            fitTextToOpenPath(textFrames);
        } else if (rbFitWidthTracking.value) {
            fitTextToOpenPathByTracking(textFrames);
        }
    }

    /**
     * パスが閉じているか判定する
     * @param {PageItem} pathItem - 対象のパス
     * @returns {boolean} 閉じていれば true
     */
    function isClosedPathItem(pathItem) {
        try {
            if (!pathItem) return false;
            if (pathItem.typename === "CompoundPathItem") {
                if (pathItem.pathItems && pathItem.pathItems.length > 0) return !!pathItem.pathItems[0].closed;
                return false;
            }
            if (pathItem.typename === "PathItem") return !!pathItem.closed;
        } catch (e) { }
        return false;
    }

    /**
     * フィットの対象（編集できるオープンパス上のパス上文字）か判定する
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @returns {boolean} 対象なら true
     */
    function isTargetPathText(textFrame) {
        try {
            if (!textFrame || textFrame.typename !== "TextFrame") return false;
            if (textFrame.kind !== TextType.PATHTEXT) return false;
            if (!textFrame.editable || textFrame.locked || textFrame.hidden) return false;
            var textPath = textFrame.textPath;
            if (!textPath) return false;
            if (isClosedPathItem(textPath)) return false;
            return true;
        } catch (e) { }
        return false;
    }

    /**
     * 表示されている行数を返す（最低1）
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @returns {number} 行数
     */
    function getLineAmount(textFrame) {
        try {
            if (textFrame.lines && textFrame.lines.length > 0) return textFrame.lines.length;
        } catch (e) { }
        return 1;
    }

    /**
     * テキストがあふれているか判定する
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @param {number} lineAmount - 表示する行数
     * @returns {boolean} あふれていれば true
     */
    function isOverset(textFrame, lineAmount) {
        try {
            if (!textFrame) return false;
            if (textFrame.lines.length > 0) {
                var charactersOnVisibleLines = 0;
                if (typeof (lineAmount) === "undefined" || lineAmount === null) {
                    lineAmount = 1;
                } else {
                    lineAmount = Math.floor(lineAmount);
                    if (lineAmount < 1) lineAmount = 1;
                    if (lineAmount > textFrame.lines.length) lineAmount = textFrame.lines.length;
                }
                for (var i = 0; i < lineAmount; i++) {
                    charactersOnVisibleLines += textFrame.lines[i].characters.length;
                }
                return (charactersOnVisibleLines < textFrame.characters.length);
            } else if (textFrame.characters.length > 0) {
                return true;
            }
        } catch (e) { }
        return false;
    }

    /**
     * オープンパスの端まで、文字サイズだけでパス上文字を合わせる
     * あふれていれば少しずつ縮め、収まっていれば一度あふれるまで広げてから縮める（ほぼ最大）
     * @param {TextFrame[]} textFrames - 対象のパス上文字
     * @returns {boolean} 対象があれば true
     */
    function fitTextToOpenPath(textFrames) {
        if (!textFrames || textFrames.length === 0) return false;

        var fitOptions = {
            increment: 0.1,
            minFontSize: 0.1,
            maxShrinkIter: 2000,
            maxGrowIter: 10,
            maxFontSize: 2000
        };

        /**
         * あふれるまで拡大してから、収まるまで縮小する
         * @param {TextFrame} textFrame - 対象のパス上文字
         * @returns {void}
         */
        function shrinkFont(textFrame) {
            try {
                if (!textFrame || textFrame.characters.length <= 0) return;

                var lineAmount = getLineAmount(textFrame);

                /* 収まっていれば、あふれるまで倍にしていく（上限あり） / Grow by doubling until it overflows */
                if (!isOverset(textFrame, lineAmount)) {
                    var growIterations = 0;
                    while (!isOverset(textFrame, lineAmount) && growIterations < fitOptions.maxGrowIter) {
                        var growSize = textFrame.textRange.characterAttributes.size;
                        if (growSize >= fitOptions.maxFontSize) break;
                        textFrame.textRange.characterAttributes.size = Math.min(fitOptions.maxFontSize, growSize * 2);
                        growIterations++;
                    }
                }

                /* 収まるまで少しずつ縮める / Then shrink in small steps until it fits */
                var shrinkIterations = 0;
                while (isOverset(textFrame, lineAmount)) {
                    var currentSize = textFrame.textRange.characterAttributes.size;
                    if (currentSize <= fitOptions.minFontSize) break;

                    textFrame.textRange.characterAttributes.size = Math.max(fitOptions.minFontSize, currentSize - fitOptions.increment);

                    shrinkIterations++;
                    if (shrinkIterations >= fitOptions.maxShrinkIter) break;
                }
            } catch (e) { }
        }

        for (var i = 0; i < textFrames.length; i++) {
            var textFrame = textFrames[i];
            if (!isTargetPathText(textFrame)) continue;
            shrinkFont(textFrame);
        }

        return true;
    }

    /**
     * オープンパスの端まで、トラッキングだけで合わせる（文字サイズは保つ）
     * @param {TextFrame[]} textFrames - 対象のパス上文字
     * @returns {boolean} 対象があれば true
     */
    function fitTextToOpenPathByTracking(textFrames) {
        if (!textFrames || textFrames.length === 0) return false;

        var fitOptions = {
            coarseStep: 50,     /* 粗調整1ステップのトラッキング量 / tracking units per coarse step */
            fineStep: 1,        /* 微調整1ステップのトラッキング量 / tracking units per fine step */
            minTracking: -1000, /* 累積トラッキングの下限 / tightest allowed cumulative delta */
            maxTracking: 20000, /* 累積トラッキングの上限 / loosest allowed cumulative delta */
            maxIter: 4000
        };

        /**
         * 1つのパス上文字をトラッキングでパス幅に合わせる
         * @param {TextFrame} textFrame - 対象のパス上文字
         * @returns {void}
         */
        function fitByTracking(textFrame) {
            try {
                if (!textFrame || textFrame.characters.length <= 0) return;

                var lineAmount = getLineAmount(textFrame);
                var appliedTracking = 0; /* これまでに加えた累積トラッキング / cumulative tracking delta applied so far */
                var iterations;

                if (isOverset(textFrame, lineAmount)) {
                    /* あふれている：収まるまで粗く詰める / too wide: tighten (coarse) until it fits */
                    iterations = 0;
                    while (isOverset(textFrame, lineAmount) && iterations < fitOptions.maxIter) {
                        if (appliedTracking - fitOptions.coarseStep < fitOptions.minTracking) break;
                        addDeltaToTextRanges(textFrame, "tracking", -fitOptions.coarseStep);
                        appliedTracking -= fitOptions.coarseStep;
                        iterations++;
                    }
                    /* 再びあふれるまで細かく広げる / loosen back (fine) until it overflows again */
                    iterations = 0;
                    while (!isOverset(textFrame, lineAmount) && iterations < fitOptions.maxIter) {
                        if (appliedTracking + fitOptions.fineStep > fitOptions.maxTracking) break;
                        addDeltaToTextRanges(textFrame, "tracking", fitOptions.fineStep);
                        appliedTracking += fitOptions.fineStep;
                        iterations++;
                    }
                    /* 1ステップ行き過ぎたら戻して収める / stepped one step too far: pull back once so it fits */
                    if (isOverset(textFrame, lineAmount)) {
                        addDeltaToTextRanges(textFrame, "tracking", -fitOptions.fineStep);
                        appliedTracking -= fitOptions.fineStep;
                    }
                } else {
                    /* 余裕あり：あふれるまで粗く広げる / fits with room: loosen (coarse) until it overflows */
                    iterations = 0;
                    while (!isOverset(textFrame, lineAmount) && iterations < fitOptions.maxIter) {
                        if (appliedTracking + fitOptions.coarseStep > fitOptions.maxTracking) break;
                        addDeltaToTextRanges(textFrame, "tracking", fitOptions.coarseStep);
                        appliedTracking += fitOptions.coarseStep;
                        iterations++;
                    }
                    /* 収まるまで細かく詰める / tighten back (fine) until it fits */
                    iterations = 0;
                    while (isOverset(textFrame, lineAmount) && iterations < fitOptions.maxIter) {
                        if (appliedTracking - fitOptions.fineStep < fitOptions.minTracking) break;
                        addDeltaToTextRanges(textFrame, "tracking", -fitOptions.fineStep);
                        appliedTracking -= fitOptions.fineStep;
                        iterations++;
                    }
                }
            } catch (e) { }
        }

        for (var i = 0; i < textFrames.length; i++) {
            var textFrame = textFrames[i];
            if (!isTargetPathText(textFrame)) continue;
            fitByTracking(textFrame);
        }

        return true;
    }
}());
