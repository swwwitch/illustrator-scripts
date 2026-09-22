#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

アートボードを現在のサイズを基準にスケール変更します。ダイアログを開いたままライブプレビューできます。
対象は現在のアートボード／すべて／指定から選べ、9つの基準点のいずれかを固定して再計算します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AiArtboardScaler.md

### Overview

Scales artboards relative to their current size, with a live preview while the dialog stays open.
The target can be the current artboard, all of them, or a specified set, recalculated around one of nine reference points.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AiArtboardScaler.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AiArtboardScaler";             /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-07-15";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-23";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AiArtboardScaler.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AiArtboardScaler.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // レイアウト / Layout
    // =========================================

    /* ウィンドウ・パネルの余白と間隔 / Window & panel margins and spacing */
    var WINDOW_MARGINS = 16;                 /* ウィンドウ外周の余白 / window margin */
    var WINDOW_SPACING = 12;                 /* ウィンドウ内の要素間隔 / window spacing */
    var PANEL_MARGINS  = [16, 20, 16, 12];   /* パネル余白 [左,上,右,下] / panel margins */
    var PANEL_SPACING  = 8;                  /* パネル内の要素間隔 / panel spacing */
    var OPTION_PANEL_SPACING = 6;            /* 対象・サイズパネル内の要素間隔 / spacing inside the target and size panels */
    var FIELD_LABEL_WIDTH = 64;              /* ラベル幅を揃えるための固定幅（「スケール:」が収まる幅）/ fixed label width (fits "Scale:") */
    var NUMBER_FIELD_CHARS = 4;              /* スケール・幅・高さ欄の文字数 / width of the scale, width and height fields */
    var SELECTION_FIELD_CHARS = 8;           /* 「指定」欄の文字数（通常の2倍）/ width of the Specify field (twice the usual) */
    var SCALE_COLUMN_GAP = 16;               /* 入力欄の列と基準点の列の間隔 / gap between the field column and the anchor column */
    var CHECKBOX_GROUP_MARGINS = [0, 10, 0, 0]; /* チェックボックス群の余白（上に10px）/ checkbox group margins (10px on top) */
    var ANCHOR_WIDGET_SIZE = 66;             /* 基準点ウィジェットの一辺 / side of the anchor widget */

    /**
     * ウィンドウに共通のレイアウト設定を適用する
     * @param {Window} targetWindow - 対象のウィンドウ
     * @param {number} [spacing] - 要素間隔（省略時は WINDOW_SPACING）
     * @returns {void}
     */
    function setupWindow(targetWindow, spacing) {
        targetWindow.orientation = "column";
        targetWindow.alignChildren = "fill";
        targetWindow.margins = WINDOW_MARGINS;
        targetWindow.spacing = (typeof spacing === "number") ? spacing : WINDOW_SPACING;
    }

    /**
     * パネルに共通のレイアウト設定を適用する
     * @param {Panel} targetPanel - 対象のパネル
     * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function setupPanel(targetPanel, spacing) {
        targetPanel.orientation = "column";
        targetPanel.alignChildren = ["fill", "top"];
        targetPanel.alignment = "fill";
        targetPanel.margins = PANEL_MARGINS;
        targetPanel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
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

    /**
     * 数値を小数2桁に丸めて文字列で返す / Round a number to 2 decimals and return as string
     * @param {number} value - 丸める値
     * @returns {string} 表示用の文字列
     */
    function formatNumber(value) {
        return "" + (Math.round(value * 100) / 100);
    }

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * 現在のUI言語を返す / Return current UI language
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = getCurrentLang();

    var LABELS = {
        dialog: {
            title: { ja: "アートボードサイズ変更", en: "Resize Artboards" }
        },
        panel: {
            target: { ja: "対象", en: "Target" },
            scale:  { ja: "サイズ・スケール", en: "Size & Scale" }
        },
        radio: {
            current: { ja: "現在のアートボード", en: "Current artboard" },
            all:     { ja: "すべてのアートボード", en: "All artboards" },
            specify: { ja: "指定", en: "Specify" }
        },
        checkbox: {
            scaleObjects: { ja: "オブジェクトと一緒に拡大・縮小", en: "Scale objects together" },
            pixelGrid: { ja: "ピクセルグリッドに最適化", en: "Optimize for pixel grid" }
        },
        fieldLabel: {
            scale:  { ja: "スケール", en: "Scale" },
            width:  { ja: "幅", en: "Width" },
            height: { ja: "高さ", en: "Height" },
            anchor: { ja: "基準点", en: "Anchor" }
        },
        tooltip: {
            current: { ja: "いま選ばれているアートボードだけを変更します。", en: "Changes only the artboard that is currently active." },
            all:     { ja: "ドキュメント内のすべてのアートボードを変更します。", en: "Changes every artboard in the document." },
            specify: { ja: "番号で対象を指定します（例: 3, 4 または 3-5）。", en: "Picks the artboards by number (for example 3, 4 or 3-5)." },
            scale:   { ja: "現在のサイズに対する倍率（％）です。幅・高さと連動します。", en: "Percentage of the current size. It is linked to the width and height." },
            size:    { ja: "変更後のサイズです。入力するとスケールが連動して変わります。", en: "The size after resizing. Typing here updates the scale." },
            anchor:  { ja: "サイズを変えるときに動かさない位置です。3×3のマスで選びます。", en: "The point that stays put while the artboard is resized. Pick it on the 3x3 grid." },
            scaleObjects: {
                ja: "アートボードの拡大・縮小に合わせて、載っているオブジェクトも一緒に変形します。",
                en: "Scales the objects on the artboard along with the artboard itself."
            },
            pixelGrid: {
                ja: "アートボードの位置とサイズを整数ピクセルにそろえます。",
                en: "Snaps the artboard position and size to whole pixels."
            }
        },
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            apply:  { ja: "適用", en: "Apply" }
        },
        alert: {
            noDocument:      { ja: "開いているドキュメントがありません。", en: "No documents are open." },
            invalidNumber:   { ja: "正の数値を入力してください。", en: "Please enter positive numbers." },
            invalidSelection: {
                ja: "対象アートボードの指定が正しくありません（例: 3, 4 または 3-5）。",
                en: "Invalid artboard selection (e.g. 3, 4 or 3-5)."
            },
            transformError: {
                ja: "変形の適用に失敗したため、処理を中止して元に戻しました。",
                en: "Failed to apply the transform; the operation was cancelled and reverted."
            },
            restoreError: {
                ja: "アートボードの復元に失敗しました。手動で元に戻してください（取り消し等）。",
                en: "Failed to restore the artboards. Please revert manually (e.g. Undo)."
            }
        }
    };

    /**
     * ドット区切りパスでラベルを取得しローカライズする（キー欠落時は labelPath を返す）
     * Resolve a localized label by dot path (returns the path itself if a key is missing)
     * @param {string} labelPath - 例 "panel.target"
     * @returns {string} ローカライズ済み文字列（見つからなければ labelPath）
     */
    function getLabel(labelPath) {
        var labelPathKeys = labelPath.split(".");
        var labelNode = LABELS;
        /* 各階層を安全に辿る（途中で欠落したら labelPath を返す） / Walk each level safely; return the path if anything is missing */
        for (var i = 0; i < labelPathKeys.length; i++) {
            if (labelNode === null || typeof labelNode !== "object" || !labelNode.hasOwnProperty(labelPathKeys[i])) {
                return labelPath;
            }
            labelNode = labelNode[labelPathKeys[i]];
        }
        /* 現在言語→英語→日本語の順にフォールバック / Fall back current language → English → Japanese */
        var localizedText = labelNode[uiLang];
        if (localizedText === undefined || localizedText === null) { localizedText = labelNode.en; }
        if (localizedText === undefined || localizedText === null) { localizedText = labelNode.ja; }
        if (typeof localizedText !== "string") { return labelPath; } /* どの言語値も無ければ labelPath / no language value → path */
        return localizedText.replace(/\{slash\}/g, "/");
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
    // 対象指定の解析 / Target spec parsing
    // =========================================

    /**
     * 前後の空白を除去する（ES3にString.trimが無いため） / Trim whitespace (ES3 has no String.trim)
     * @param {string} sourceText - 対象の文字列
     * @returns {string} 前後の空白を除いた文字列
     */
    function trimWhitespace(sourceText) {
        return ("" + sourceText).replace(/^\s+/, "").replace(/\s+$/, "");
    }

    /**
     * 「3, 4」「3-5」形式の指定を0始まりのアートボード索引配列に厳密変換する
     * Strictly parse a "3, 4" / "3-5" style spec into 0-based artboard indices
     * 各トークンを正規表現で完全一致検証し、"3abc"・"1-2-3"・空トークン(",")等は不正扱い。
     * @param {string} selectionText - 入力文字列（1始まり）
     * @param {number} artboardCount - アートボード総数
     * @returns {number[]|null} 索引配列。不正な場合は null
     */
    function parseArtboardSelection(selectionText, artboardCount) {
        var singlePattern = /^\d+$/;                 /* 単一番号 / single number */
        var rangePattern = /^\d+\s*-\s*\d+$/;        /* 範囲（前後の空白可） / range (spaces allowed) */
        var artboardIndices = [];
        var seenIndices = {};
        var tokens = ("" + selectionText).split(",");
        for (var i = 0; i < tokens.length; i++) {
            var token = trimWhitespace(tokens[i]);
            /* 空トークン（"1," ",2" "1,,2" 等）は不正扱い / empty token (from "1," ",2" "1,,2") is invalid */
            if (token === "") { return null; }

            var rangeStart, rangeEnd;
            if (rangePattern.test(token)) {
                var dashPosition = token.indexOf("-");
                rangeStart = parseInt(trimWhitespace(token.substring(0, dashPosition)), 10);
                rangeEnd = parseInt(trimWhitespace(token.substring(dashPosition + 1)), 10);
            } else if (singlePattern.test(token)) {
                rangeStart = rangeEnd = parseInt(token, 10);
            } else {
                return null; /* 形式不一致（"3abc" "1-2-3" 等） / format mismatch */
            }
            if (rangeStart > rangeEnd) { var swappedStart = rangeStart; rangeStart = rangeEnd; rangeEnd = swappedStart; } /* 逆順は入れ替え / swap reversed range */

            for (var artboardNumber = rangeStart; artboardNumber <= rangeEnd; artboardNumber++) {
                if (artboardNumber < 1 || artboardNumber > artboardCount) { return null; } /* 範囲外は不正 / out of range is invalid */
                var artboardIndex = artboardNumber - 1;
                if (!seenIndices[artboardIndex]) { seenIndices[artboardIndex] = true; artboardIndices.push(artboardIndex); } /* 重複は1回だけ / dedupe */
            }
        }
        return artboardIndices.length ? artboardIndices : null;
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 「ラベル: [入力欄] 単位」の1行を生成する / Build one "label: [input] unit" row
     * @param {Group} parentGroup - 追加先
     * @param {string} captionText - コロン付きの項目名
     * @param {string} defaultValue - 入力欄の初期値
     * @param {string} unitLabel - 入力欄の後ろに表示する単位
     * @param {string} [tooltipText] - 入力欄に付けるツールチップ
     * @returns {EditText} 生成した入力欄
     */
    function addSizeField(parentGroup, captionText, defaultValue, unitLabel, tooltipText) {
        var fieldRow = parentGroup.add("group");
        fieldRow.orientation = "row";
        var captionLabel = fieldRow.add("statictext", undefined, captionText);
        captionLabel.preferredSize.width = FIELD_LABEL_WIDTH; /* ラベル幅を固定して揃える / Fix width to align labels */
        captionLabel.justify = "right";                       /* 右揃え / Right-align the label */
        var valueInput = fieldRow.add("edittext", undefined, defaultValue);
        if (tooltipText) valueInput.helpTip = tooltipText;
        valueInput.characters = NUMBER_FIELD_CHARS;
        fieldRow.add("statictext", undefined, unitLabel); /* 入力欄の後ろに単位 / Unit after the input */
        return valueInput;
    }

    /**
     * テキストフィールドで↑↓キーによる値の増減を有効にする / Enable arrow-key value change on a text field
     * ↑↓で±1、Shift併用で±10（10の倍数にスナップ）。optionキーは使わない。
     * @param {EditText} editText - 対象の入力欄（数値を保持していること）
     * @returns {void}
     */
    function changeValueByArrowKey(editText) {
        editText.addEventListener("keydown", function (event) {
            var value = Number(editText.text);
            if (isNaN(value)) { return; }

            /* 修飾キーは event 優先、取得できなければ keyboardState にフォールバック / Read modifiers from the event first, fall back to keyboardState */
            var shiftPressed = event.shiftKey;
            if (shiftPressed === undefined) { shiftPressed = ScriptUI.environment.keyboardState.shiftKey; }
            var delta = shiftPressed ? 10 : 1;

            if (event.keyName === "Up") {
                /* Shift時は10の倍数にスナップ / Snap to multiples of 10 with Shift */
                value = shiftPressed ? Math.ceil((value + 1) / delta) * delta : value + delta;
                event.preventDefault();
            } else if (event.keyName === "Down") {
                value = shiftPressed ? Math.floor((value - 1) / delta) * delta : value - delta;
                if (value < 0) { value = 0; }
                event.preventDefault();
            } else {
                return;
            }

            editText.text = Math.round(value);
            /* 連動している依存フィールドを更新する / Fire onChanging so linked fields update */
            if (typeof editText.onChanging === "function") { editText.onChanging(); }
        });
    }

    /**
     * プレビュー更新を要求する（コントローラ未接続なら何もしない） / Request a preview refresh (no-op until wired)
     * @param {Window} resizeDialog - onPreview を持つダイアログ
     * @returns {void}
     */
    function requestPreview(resizeDialog) {
        if (resizeDialog.onPreview) { resizeDialog.onPreview(); }
    }

    /**
     * 対象パネル（現在のアートボード／すべてのアートボード／指定）を構築する
     * Build the target panel (Current artboard / All artboards / Specify)
     * @param {Window} resizeDialog - 追加先ダイアログ（コントロールをプロパティとして公開する）
     * @param {string} defaultSelection - 指定入力欄の初期値（例: "1-6"）
     * @returns {void}
     */
    function addTargetPanel(resizeDialog, defaultSelection) {
        var targetPanel = resizeDialog.add("panel", undefined, getLabel("panel.target"));
        setupPanel(targetPanel, OPTION_PANEL_SPACING);

        /**
         * パネルに左寄せの行を追加する
         * @returns {Group} 追加した行
         */
        function addLeftRow() {
            var radioRow = targetPanel.add("group");
            radioRow.orientation = "row";
            radioRow.alignment = "left";
            return radioRow;
        }

        /* 「現在のアートボード」ラジオ / "Current artboard" radio */
        var currentRadio = addLeftRow().add("radiobutton", undefined, getLabel("radio.current"));
        currentRadio.helpTip = getLabel("tooltip.current");

        /* 「すべてのアートボード」ラジオ / "All artboards" radio */
        var allRadio = addLeftRow().add("radiobutton", undefined, getLabel("radio.all"));
        allRadio.helpTip = getLabel("tooltip.all");

        /* 「指定」ラジオ＋範囲入力 / "Specify" radio with range input */
        var specifyRow = addLeftRow();
        var specifyRadio = specifyRow.add("radiobutton", undefined, getLabel("radio.specify"));
        specifyRadio.helpTip = getLabel("tooltip.specify");
        var selectInput = specifyRow.add("edittext", undefined, defaultSelection);
        selectInput.helpTip = getLabel("tooltip.specify");
        selectInput.characters = SELECTION_FIELD_CHARS;

        var targetRadios = [currentRadio, allRadio, specifyRadio];

        /**
         * ラジオは親が異なると排他にならないため手動で同期する / Sync manually since radios in different parents are not exclusive
         * @param {RadioButton} activeRadio - 選ばれたラジオ
         * @returns {void}
         */
        function selectTarget(activeRadio) {
            for (var i = 0; i < targetRadios.length; i++) { targetRadios[i].value = (targetRadios[i] === activeRadio); }
            selectInput.enabled = (activeRadio === specifyRadio); /* 入力欄は「指定」時のみ有効 / input enabled only for "Specify" */
            requestPreview(resizeDialog);
        }
        currentRadio.onClick = function () { selectTarget(currentRadio); };
        allRadio.onClick = function () { selectTarget(allRadio); };
        specifyRadio.onClick = function () { selectTarget(specifyRadio); };
        selectInput.onChanging = function () { requestPreview(resizeDialog); };

        /* 初期状態は「すべてのアートボード」 / Default to "All artboards" */
        allRadio.value = true;
        selectInput.enabled = false;

        resizeDialog.currentRadio = currentRadio;
        resizeDialog.allRadio = allRadio;
        resizeDialog.specifyRadio = specifyRadio;
        resizeDialog.selectInput = selectInput;
    }

    /**
     * サイズ・スケール統合パネルを構築する / Build the merged size & scale panel
     * スケール%を一意の倍率とし、幅・高さ（現在の定規単位）はアクティブアートボードの現在サイズ基準で相互連動する。
     * The scale % is a single uniform ratio; width/height (in the current ruler unit) are linked to the active artboard's current size.
     * @param {Window} resizeDialog - 追加先ダイアログ（コントロールをプロパティとして公開する）
     * @param {string} unitLabel - 幅・高さの単位ラベル（例: "mm"）
     * @param {number} baseWidth - 幅の基準値：アクティブアートボードの現在幅を現在の定規単位へ変換済み（ptではない）
     * @param {number} baseHeight - 高さの基準値：アクティブアートボードの現在高さを現在の定規単位へ変換済み（ptではない）
     * @returns {void}
     */
    function addScalePanel(resizeDialog, unitLabel, baseWidth, baseHeight) {
        var scalePanel = resizeDialog.add("panel", undefined, getLabel("panel.scale"));
        setupPanel(scalePanel, OPTION_PANEL_SPACING);

        /* 2列レイアウト：左=スケール/幅/高さの3行、右=基準点グリッド / Two columns: left has scale/width/height rows, right holds the anchor grid */
        var columnsRow = scalePanel.add("group");
        columnsRow.orientation = "row";
        columnsRow.alignChildren = ["left", "center"];
        columnsRow.spacing = SCALE_COLUMN_GAP;

        var fieldColumn = columnsRow.add("group");
        fieldColumn.orientation = "column";
        fieldColumn.alignChildren = ["left", "top"];
        fieldColumn.spacing = OPTION_PANEL_SPACING;
        resizeDialog.scaleInput = addSizeField(fieldColumn, labelText("fieldLabel.scale"), "100", "%", getLabel("tooltip.scale"));
        resizeDialog.widthInput = addSizeField(fieldColumn, labelText("fieldLabel.width"), formatNumber(baseWidth), unitLabel, getLabel("tooltip.size"));
        resizeDialog.heightInput = addSizeField(fieldColumn, labelText("fieldLabel.height"), formatNumber(baseHeight), unitLabel, getLabel("tooltip.size"));

        /* 基準点グリッド（2列目） / Anchor reference-point grid (second column) */
        addAnchorWidget(resizeDialog, columnsRow);

        /* チェックボックス群（上に10pxの余白） / Checkbox group (10px top margin) */
        var checkboxGroup = scalePanel.add("group");
        checkboxGroup.orientation = "column";
        checkboxGroup.alignChildren = ["left", "top"];
        checkboxGroup.margins = CHECKBOX_GROUP_MARGINS;

        /* オブジェクトも一緒に拡大・縮小するか / Whether to scale the objects along with the artboard */
        var scaleObjectsCheckbox = checkboxGroup.add("checkbox", undefined, getLabel("checkbox.scaleObjects"));
        scaleObjectsCheckbox.helpTip = getLabel("tooltip.scaleObjects");
        scaleObjectsCheckbox.value = true;

        /* アートボードのX/Y/W/Hを整数化してピクセルグリッドに合わせる / Round artboard X/Y/W/H to integers for the pixel grid */
        var pixelGridCheckbox = checkboxGroup.add("checkbox", undefined, getLabel("checkbox.pixelGrid"));
        pixelGridCheckbox.helpTip = getLabel("tooltip.pixelGrid");
        pixelGridCheckbox.value = false;

        linkScaleFields(resizeDialog, baseWidth, baseHeight);
        scaleObjectsCheckbox.onClick = function () {
            requestPreview(resizeDialog);
        };
        pixelGridCheckbox.onClick = function () {
            requestPreview(resizeDialog);
        };

        resizeDialog.scaleObjectsCheckbox = scaleObjectsCheckbox;
        resizeDialog.pixelGridCheckbox = pixelGridCheckbox;
    }

    /**
     * スケール・幅・高さの3欄を連動させる（入力中は相互に更新し、確定時に正規化する）
     * Link the scale, width and height fields (update each other while typing, normalize on commit)
     * @param {Window} resizeDialog - scaleInput / widthInput / heightInput を持つダイアログ
     * @param {number} baseWidth - 幅の基準値（定規単位）
     * @param {number} baseHeight - 高さの基準値（定規単位）
     * @returns {void}
     */
    function linkScaleFields(resizeDialog, baseWidth, baseHeight) {
        var scaleInput = resizeDialog.scaleInput;
        var widthInput = resizeDialog.widthInput;
        var heightInput = resizeDialog.heightInput;

        /**
         * スケール%から幅・高さ(現在サイズ×%)を再計算する / Recalc width/height (current size × %) from the scale
         * @returns {void}
         */
        function applyScaleToSize() {
            var percent = parseFloat(scaleInput.text);
            if (isNaN(percent)) { return; }
            widthInput.text = formatNumber(baseWidth * percent / 100);
            heightInput.text = formatNumber(baseHeight * percent / 100);
        }

        /**
         * 編集中の寸法欄からスケール%を逆算し、スケール欄ともう一方の寸法欄だけ更新する
         * Derive the scale from the edited size field, updating only the scale field and the OTHER size field
         * （編集中の欄自身は書き換えない＝小数点入力が消える不具合を防ぐ / never rewrite the field being edited, so decimals can be typed）
         * @param {number} editedValue - 編集中の欄の値
         * @param {number} editedBase - 編集中の欄の基準サイズ
         * @param {EditText} otherField - もう一方（連動更新する）寸法欄
         * @param {number} otherBase - もう一方の基準サイズ
         * @returns {void}
         */
        function applySizeToScale(editedValue, editedBase, otherField, otherBase) {
            if (isNaN(editedValue) || editedBase === 0) { return; }
            var percent = editedValue / editedBase * 100;
            scaleInput.text = formatNumber(percent);
            otherField.text = formatNumber(otherBase * percent / 100);
        }

        /* 直近の有効なスケール%（入力強化の復帰先） / Last valid scale % (fallback for input hardening) */
        var lastValidScale = 100;

        /**
         * 確定時に3欄をスケール基準へ正規化する（不正値は直近の有効値へ復帰） / On commit, canonicalize all three fields to the scale (invalid → last valid)
         * @returns {void}
         */
        function normalizeFields() {
            var percent = parseFloat(scaleInput.text);
            if (isNaN(percent) || percent <= 0) {
                percent = lastValidScale; /* 空・0・負・非数値は直近の有効値へ / empty/0/negative/NaN falls back */
            } else {
                lastValidScale = percent;
            }
            scaleInput.text = formatNumber(percent);
            widthInput.text = formatNumber(baseWidth * percent / 100);
            heightInput.text = formatNumber(baseHeight * percent / 100);
            requestPreview(resizeDialog);
        }

        scaleInput.onChanging = function () {
            applyScaleToSize();
            requestPreview(resizeDialog);
        };
        widthInput.onChanging = function () {
            applySizeToScale(parseFloat(widthInput.text), baseWidth, heightInput, baseHeight);
            requestPreview(resizeDialog);
        };
        heightInput.onChanging = function () {
            applySizeToScale(parseFloat(heightInput.text), baseHeight, widthInput, baseWidth);
            requestPreview(resizeDialog);
        };
        /* 確定（Enter/フォーカスアウト）で正規化 / Normalize on commit (Enter / focus-out) */
        scaleInput.onChange = normalizeFields;
        widthInput.onChange = normalizeFields;
        heightInput.onChange = normalizeFields;
    }

    /**
     * サイズ入力ダイアログを構築する / Build the size-input dialog
     * @param {string} unitLabel - 表示する単位ラベル（例: "mm"）
     * @param {string} defaultWidth - 幅入力欄の初期値
     * @param {string} defaultHeight - 高さ入力欄の初期値
     * @param {number} artboardCount - アートボード総数（選択の初期値に使用）
     * @returns {Window} ダイアログウィンドウ（各入力コントロールを公開）
     */
    function createResizeDialog(unitLabel, defaultWidth, defaultHeight, artboardCount) {
        /* 基準点ウィジェットの配色をUI明暗に合わせる / Match the anchor widget colors to the light/dark UI */
        initAnchorColors();

        /* タイトルバーにバージョンを表示 / Show version in the title bar */
        var resizeDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        setupWindow(resizeDialog);

        /* 対象パネル / Target panel */
        addTargetPanel(resizeDialog, "1-" + artboardCount);

        /* サイズ・スケール統合パネル（アクティブアートボードの現在サイズを基準に） / Merged size & scale panel (based on the active artboard's current size) */
        addScalePanel(resizeDialog, unitLabel, parseFloat(defaultWidth), parseFloat(defaultHeight));

        /* 数値入力欄に↑↓キーでの増減を付与 / Enable arrow-key value change on the numeric fields */
        changeValueByArrowKey(resizeDialog.scaleInput);
        changeValueByArrowKey(resizeDialog.widthInput);
        changeValueByArrowKey(resizeDialog.heightInput);

        /* ボタン類はパネル幅いっぱいには広げず右寄せ / Keep buttons right-aligned, not full width */
        var btnRowGroup = resizeDialog.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.alignment = "right";
        btnRowGroup.add("button", undefined, getLabel("button.cancel"), {name: "cancel"});
        btnRowGroup.add("button", undefined, getLabel("button.apply"), {name: "ok"});

        return resizeDialog;
    }

    // =========================================
    // 基準点ウィジェット / Anchor widget
    // =========================================

    /**
     * UI明度(0..1)を取得する（失敗時は1=明るい） / Get UI brightness (0..1); 1 on failure
     * @returns {number} 0..1 の明度
     */
    function getUIBrightness() {
        try {
            var brightness = app.preferences.getRealPreference("uiBrightness");
            if (brightness < 0) { brightness = 0; }
            if (brightness > 1) { brightness = 1; }
            return brightness;
        } catch (e) {
            return 1;
        }
    }

    /**
     * 明るいUIかどうか / Whether the UI is light
     * @returns {boolean} 明るいUIなら true
     */
    function isLightUI() {
        return getUIBrightness() > 0.5;
    }

    /* 基準点セルの配色（initAnchorColorsでUI明暗に合わせて上書き） / Anchor-cell colors (overwritten by initAnchorColors) */
    var ANCHOR_LINE_COLOR = [0.6, 0.6, 0.6, 1];      /* 枠・ケイ線：薄いグレー / border & rules: light gray */
    var ANCHOR_SELECTED_FILL = [0.4, 0.4, 0.4, 1];   /* 選択セルの塗り / fill of the selected cell */

    /**
     * UI明暗に合わせて選択セルの塗り色を決める（表示前に呼ぶ） / Decide the selected-cell fill from the light/dark UI (call before showing)
     * @returns {void}
     */
    function initAnchorColors() {
        ANCHOR_SELECTED_FILL = isLightUI() ? [0.4, 0.4, 0.4, 1] : [0.8, 0.8, 0.8, 1];
    }

    /**
     * コントロールを再描画する（notifyは環境により例外を投げ得るので保護） / Redraw a control (notify can throw in some environments)
     * @param {Object} control - 再描画するコントロール
     * @returns {void}
     */
    function redrawControl(control) {
        try { control.notify("onDraw"); } catch (e) {}
    }

    /**
     * 正方形のサブパスを1つ作る / Build one square subpath
     * @param {ScriptUIGraphics} graphics - 描画先
     * @param {number} x - 左上のX
     * @param {number} y - 左上のY
     * @param {number} size - 一辺
     * @returns {void}
     */
    function squarePath(graphics, x, y, size) {
        graphics.newPath();
        graphics.moveTo(x, y);
        graphics.lineTo(x + size, y);
        graphics.lineTo(x + size, y + size);
        graphics.lineTo(x, y + size);
        graphics.closePath();
    }

    /**
     * 基準点セルの□を1つ描画する（選択時のみ塗り、枠は常時） / Draw one anchor-cell square (fill only when selected, always bordered)
     * @param {ScriptUIGraphics} graphics - 描画先
     * @param {number} x - 左上のX
     * @param {number} y - 左上のY
     * @param {number} size - 一辺
     * @param {boolean} selected - 選択中のセルなら true
     * @returns {void}
     */
    function drawAnchorCell(graphics, x, y, size, selected) {
        if (selected) {
            squarePath(graphics, x, y, size);
            graphics.fillPath(graphics.newBrush(graphics.BrushType.SOLID_COLOR, ANCHOR_SELECTED_FILL));
        }
        /* 枠線は常に薄いグレー（塗りの上に描く） / Border always light gray (drawn over the fill) */
        squarePath(graphics, x, y, size);
        graphics.strokePath(graphics.newPen(graphics.PenType.SOLID_COLOR, ANCHOR_LINE_COLOR, 1));
    }

    /**
     * 9軸ウィジェットを描画する（外周の□をケイ線でつなぐ・中央は独立） / Draw the 9-axis widget (outer squares joined by rules; center stands alone)
     * @param {Button} widget - 描画するウィジェット（anchorIndex を持つ）
     * @returns {void}
     */
    function drawAnchorWidget(widget) {
        var graphics = widget.graphics;
        var width = widget.size[0];
        var height = widget.size[1];

        try {
            /* コントロール地色で塗って透過に見せる / Paint the control's own background so it looks transparent */
            graphics.rectPath(0, 0, width, height);
            graphics.fillPath(graphics.backgroundColor);
        } catch (e0) {}

        var cellSize = 9;
        var cellGap = 7.5;
        var cellStep = cellSize + cellGap;
        var gridSize = cellSize * 3 + cellGap * 2;
        var originX = Math.round((width - gridSize) / 2);
        var originY = Math.round((height - gridSize) / 2);

        /**
         * セルの左上X / Left X of a cell
         * @param {number} index - セル番号（0..8）
         * @returns {number} X座標
         */
        function anchorCellX(index) { return originX + (index % 3) * cellStep; }

        /**
         * セルの左上Y / Top Y of a cell
         * @param {number} index - セル番号（0..8）
         * @returns {number} Y座標
         */
        function anchorCellY(index) { return originY + Math.floor(index / 3) * cellStep; }

        /* 中央(4)を除く外周の□どうしをケイ線でつなぐ / Join the outer squares (except center 4) with rules */
        var connections = [[0, 1], [1, 2], [6, 7], [7, 8], [0, 3], [3, 6], [2, 5], [5, 8]];
        var linePen = graphics.newPen(graphics.PenType.SOLID_COLOR, ANCHOR_LINE_COLOR, 1);
        for (var i = 0; i < connections.length; i++) {
            var cellA = connections[i][0];
            var cellB = connections[i][1];
            var cellAX = anchorCellX(cellA);
            var cellAY = anchorCellY(cellA);
            var cellBX = anchorCellX(cellB);
            var cellBY = anchorCellY(cellB);
            graphics.newPath();
            if (cellB - cellA === 1) {
                graphics.moveTo(cellAX + cellSize, cellAY + cellSize / 2);
                graphics.lineTo(cellBX, cellBY + cellSize / 2);
            } else {
                graphics.moveTo(cellAX + cellSize / 2, cellAY + cellSize);
                graphics.lineTo(cellBX + cellSize / 2, cellBY);
            }
            graphics.strokePath(linePen);
        }

        for (var index = 0; index < 9; index++) {
            drawAnchorCell(graphics, anchorCellX(index), anchorCellY(index), cellSize, index === widget.anchorIndex);
        }
    }

    /**
     * 3×3の基準点ウィジェット（onDrawで描画するproxy）を構築する / Build a 3×3 anchor proxy widget drawn via onDraw
     * 選択に応じて resizeDialog.anchorX / resizeDialog.anchorY に 0/0.5/1 の割合を設定する（既定=左上）。
     * @param {Window} resizeDialog - 割合を書き込むダイアログ
     * @param {Group} parentGroup - 追加先グループ
     * @returns {void}
     */
    function addAnchorWidget(resizeDialog, parentGroup) {
        var anchorColumn = parentGroup.add("group");
        anchorColumn.orientation = "column";
        anchorColumn.alignChildren = ["center", "top"];
        anchorColumn.spacing = 4;
        anchorColumn.add("statictext", undefined, getLabel("fieldLabel.anchor"));

        var anchorWidget = anchorColumn.add("button", undefined, "");
        anchorWidget.helpTip = getLabel("tooltip.anchor");
        anchorWidget.preferredSize = [ANCHOR_WIDGET_SIZE, ANCHOR_WIDGET_SIZE];
        anchorWidget.minimumSize = [ANCHOR_WIDGET_SIZE, ANCHOR_WIDGET_SIZE];
        anchorWidget.maximumSize = [ANCHOR_WIDGET_SIZE, ANCHOR_WIDGET_SIZE];
        anchorWidget.anchorIndex = 0; /* 0..8 行優先（0=左上, 4=中央, 8=右下）/ 0..8 row-major (0=top-left, 4=center, 8=bottom-right) */
        anchorWidget.onDraw = function () { drawAnchorWidget(this); };
        anchorWidget.onClick = function () {}; /* セル判定は mousedown 側 / hit-testing happens in mousedown */

        /**
         * 選択索引を割合(0/0.5/1)に変換して resizeDialog に反映する / Map the index to fractions and store on the dialog
         * @returns {void}
         */
        function commitAnchor() {
            resizeDialog.anchorX = (anchorWidget.anchorIndex % 3) * 0.5;
            resizeDialog.anchorY = Math.floor(anchorWidget.anchorIndex / 3) * 0.5;
        }
        commitAnchor(); /* 既定=左上 / default: top-left */

        /* クリックした3×3のセルを基準点に設定（座標はコントロール基準）/ Set the anchor from the clicked 3x3 cell (control-relative coords) */
        try {
            anchorWidget.addEventListener("mousedown", function (event) {
                var cellColumn = Math.floor(event.clientX / (anchorWidget.size[0] / 3));
                var cellRow = Math.floor(event.clientY / (anchorWidget.size[1] / 3));
                if (cellColumn < 0) { cellColumn = 0; }
                if (cellColumn > 2) { cellColumn = 2; }
                if (cellRow < 0) { cellRow = 0; }
                if (cellRow > 2) { cellRow = 2; }
                anchorWidget.anchorIndex = cellRow * 3 + cellColumn;
                commitAnchor();
                redrawControl(anchorWidget);
                requestPreview(resizeDialog);
            });
        } catch (e) {}
    }

    // =========================================
    // オブジェクトの拡縮 / Object scaling
    // =========================================

    /**
     * 1アイテムを逆拡縮して元の位置へ戻す（rollback用） / Inverse-scale one item and put it back (for rollback)
     * @param {PageItem} pageItem - 戻すアイテム
     * @param {number} ratioX - 掛けた横の倍率
     * @param {number} ratioY - 掛けた縦の倍率
     * @param {number[]} originalPosition - 元の position [左, 上]
     * @returns {void}
     */
    function inverseResizeItem(pageItem, ratioX, ratioY, originalPosition) {
        try {
            var inverseX = (1 / ratioX) * 100;
            var inverseY = (1 / ratioY) * 100;
            /* 第7引数(changeLineWidths)は線幅拡縮のパーセント値を渡す（参考実装準拠） / 7th arg is the line-width scale percentage (per the reference) */
            pageItem.resize(inverseX, inverseY, true, true, true, true, inverseX, Transformation.TOPLEFT);
            pageItem.position = [originalPosition[0], originalPosition[1]];
        } catch (e) {}
    }

    /**
     * 【確定用】指定比率でオブジェクト群を基準点(anchor)基準に一発で拡縮する（線幅も拡縮／途中失敗は自前で完全復元）
     * [For commit] Scale items about the anchor once (also scales line width); on mid-way failure, fully reverts its own items
     * artboardsResizeWithObjects.jsx(Alexander Ladygin) を参考にした resize()+position 方式。
     * resize() でサイズ・線幅を拡縮（アイテム左上基準）し、position で基準点からのオフセットを拡縮して再配置する。
     * オブジェクトごとに { pageItem, originalPosition, resized } を記録し、失敗時はこの呼び出しで変形済みの分を確実に戻す。
     * @param {PageItem[]} targetItems - 拡縮するアイテム
     * @param {number} ratioX - 横の倍率
     * @param {number} ratioY - 縦の倍率
     * @param {number} anchorX - 基準点のX
     * @param {number} anchorY - 基準点のY
     * @returns {object} { success:boolean, transformedItems:PageItem[][, error:Error] }
     */
    function resizeItemsAbout(targetItems, ratioX, ratioY, anchorX, anchorY) {
        var transformedItems = [];
        if (!targetItems || !targetItems.length) { return { success: true, transformedItems: transformedItems }; }
        var resizeStates = []; /* {pageItem, originalPosition, resized} 途中失敗時の復元用 / for rollback on failure */
        for (var i = 0; i < targetItems.length; i++) {
            var pageItem = targetItems[i];
            var originalPosition = pageItem.position; // [左, 上]（ドキュメント座標） / [left, top] in document coordinates
            var resizeState = { pageItem: pageItem, originalPosition: [originalPosition[0], originalPosition[1]], resized: false };
            resizeStates.push(resizeState);
            try {
                /* サイズ・線幅を拡縮（アイテム左上基準）。第7引数=線幅拡縮のパーセント値（参考実装準拠） / Scale size & line width about the item's top-left; 7th arg = line-width scale percentage (per the reference) */
                pageItem.resize(ratioX * 100, ratioY * 100, true, true, true, true, ratioX * 100, Transformation.TOPLEFT);
                resizeState.resized = true; /* resize成功。以降のpositionで失敗しても逆resizeで戻せる / resize done; still recoverable if position fails */
                /* 基準点からのオフセットを拡縮して再配置 / Reposition by scaling the offset from the anchor point */
                pageItem.position = [
                    anchorX + (originalPosition[0] - anchorX) * ratioX,
                    anchorY + (originalPosition[1] - anchorY) * ratioY
                ];
                transformedItems.push(pageItem);
            } catch (e) {
                /* 途中失敗：この呼び出しで resize 済み（position前後どちらも）を確実に復元 / mid-way failure: revert every item resized in this call */
                for (var j = resizeStates.length - 1; j >= 0; j--) {
                    if (resizeStates[j].resized) {
                        inverseResizeItem(resizeStates[j].pageItem, ratioX, ratioY, resizeStates[j].originalPosition);
                    }
                }
                return { success: false, transformedItems: [], error: e };
            }
        }
        return { success: true, transformedItems: transformedItems };
    }

    // =========================================
    // オブジェクトの所属 / Object ownership
    // =========================================

    /**
     * 指定アートボード上の編集可能オブジェクトを安全に取得する（selection/active が途中例外でも壊れない）
     * Safely collect editable objects on the given artboard (selection/active stay consistent even if an exception occurs)
     * @param {Document} doc - 対象ドキュメント
     * @param {number} artboardIndex - アートボードの索引
     * @returns {PageItem[]} アートボード上のオブジェクト
     */
    function getObjectsOnArtboard(doc, artboardIndex) {
        var artboardItems = [];
        try {
            doc.selection = null;
            doc.artboards.setActiveArtboardIndex(artboardIndex);
            doc.selectObjectsOnActiveArtboard();
            var currentSelection = doc.selection;
            /* selection が null/未定義相当や配列でない場合も安全に扱う / Handle null / non-array selection safely */
            if (currentSelection && typeof currentSelection.length === "number") {
                for (var i = 0; i < currentSelection.length; i++) { artboardItems.push(currentSelection[i]); }
            }
        } finally {
            /* 例外の有無に関わらず選択を解除する / clear the selection whether or not an exception occurred */
            try { doc.selection = null; } catch (e) {}
        }
        return artboardItems;
    }

    /**
     * オブジェクトの一意識別子を返す（uuid優先、無ければ null） / Return an object's unique id (prefer uuid; null if unavailable)
     * @param {PageItem} pageItem - 対象のオブジェクト
     * @returns {string|null} uuid
     */
    function getItemUuid(pageItem) {
        try {
            if (pageItem.uuid) { return pageItem.uuid; }
        } catch (e) {}
        return null;
    }

    /**
     * 点(x,y)がアートボード矩形[左,上,右,下]の内側か / Whether point (x,y) is inside the artboard rect [L,T,R,B]
     * @param {number[]} artboardRect - アートボードの矩形
     * @param {number} x - 点のX
     * @param {number} y - 点のY
     * @returns {boolean} 内側なら true
     */
    function rectContainsPoint(artboardRect, x, y) {
        return x >= artboardRect[0] && x <= artboardRect[2] && y <= artboardRect[1] && y >= artboardRect[3];
    }

    /**
     * オブジェクト境界[左,上,右,下]とアートボード矩形の重なり面積 / Overlap area between object bounds [L,T,R,B] and an artboard rect
     * @param {number[]} itemBounds - オブジェクトの境界
     * @param {number[]} artboardRect - アートボードの矩形
     * @returns {number} 重なり面積（重ならなければ 0）
     */
    function overlapArea(itemBounds, artboardRect) {
        var overlapWidth = Math.min(itemBounds[2], artboardRect[2]) - Math.max(itemBounds[0], artboardRect[0]);
        var overlapHeight = Math.min(itemBounds[1], artboardRect[1]) - Math.max(itemBounds[3], artboardRect[3]);
        return (overlapWidth > 0 && overlapHeight > 0) ? overlapWidth * overlapHeight : 0;
    }

    /**
     * オブジェクトの所属アートボードを決める（中心点包含→重なり最大→番号が小さい方）
     * Decide which artboard owns an object (center containment → max overlap → smallest index)
     * @param {number[]} itemBounds - オブジェクト境界 [左,上,右,下]
     * @param {number[]} sortedTargets - 昇順の対象アートボード索引
     * @param {object} rectByIndex - index→矩形
     * @returns {number} 所属アートボード索引（どこにも重ならなければ -1）
     */
    function pickOwnerArtboard(itemBounds, sortedTargets, rectByIndex) {
        var centerX = (itemBounds[0] + itemBounds[2]) / 2;
        var centerY = (itemBounds[1] + itemBounds[3]) / 2;
        /* 1. 中心点が含まれるアートボード（昇順で最初） / center-point containment (first in ascending order) */
        for (var i = 0; i < sortedTargets.length; i++) {
            if (rectContainsPoint(rectByIndex[sortedTargets[i]], centerX, centerY)) { return sortedTargets[i]; }
        }
        /* 2/3. 重なり面積が最大のもの（同じなら昇順の先勝ち＝番号が小さい方） / largest overlap (ties break to the smallest index via ascending scan) */
        var bestIndex = -1;
        var bestArea = 0;
        for (var j = 0; j < sortedTargets.length; j++) {
            var area = overlapArea(itemBounds, rectByIndex[sortedTargets[j]]);
            if (area > bestArea) { bestArea = area; bestIndex = sortedTargets[j]; }
        }
        return bestArea > 0 ? bestIndex : -1;
    }

    /**
     * 対象アートボード群から、重複を排除して所属先ごとにオブジェクトをまとめる
     * Collect objects across target artboards, deduped and grouped by owning artboard
     * @param {Document} doc - 対象ドキュメント
     * @param {number[]} sortedTargets - 昇順の対象アートボード索引
     * @param {object} rectByIndex - index→元の矩形
     * @returns {object} index→[所属オブジェクト] / index → owned items
     */
    function collectOwnedObjectsByArtboard(doc, sortedTargets, rectByIndex) {
        var ownedItemsByIndex = {};
        for (var i = 0; i < sortedTargets.length; i++) { ownedItemsByIndex[sortedTargets[i]] = []; }

        /* 重複判定はオブジェクト参照で行う（uuid優先、無ければ === 比較） / Dedupe by object identity (uuid first; === fallback) */
        var seenUuids = {};
        var seenItems = [];

        /**
         * 既に処理したオブジェクトか判定し、未処理なら記録する
         * @param {PageItem} pageItem - 対象のオブジェクト
         * @returns {boolean} 既に処理済みなら true
         */
        function alreadySeen(pageItem) {
            var uuid = getItemUuid(pageItem);
            if (uuid !== null) {
                if (seenUuids[uuid]) { return true; }
                seenUuids[uuid] = true;
                return false;
            }
            for (var k = 0; k < seenItems.length; k++) {
                if (seenItems[k] === pageItem) { return true; } /* 同一参照なら重複 / same reference = duplicate */
            }
            seenItems.push(pageItem);
            return false;
        }

        for (var j = 0; j < sortedTargets.length; j++) {
            var artboardItems = getObjectsOnArtboard(doc, sortedTargets[j]);
            for (var k = 0; k < artboardItems.length; k++) {
                if (alreadySeen(artboardItems[k])) { continue; }
                var ownerIndex = pickOwnerArtboard(artboardItems[k].geometricBounds, sortedTargets, rectByIndex);
                if (ownerIndex !== -1) { ownedItemsByIndex[ownerIndex].push(artboardItems[k]); }
            }
        }
        return ownedItemsByIndex;
    }

    // =========================================
    // 設定の読み取りと寸法計算 / Settings & geometry
    // =========================================

    /**
     * 基準点を固定したまま新しいアートボード矩形を算出する（ピクセルグリッド時は幅高さと基準点を整数化）
     * Compute the new artboard rect keeping the anchor fixed (integerize size & pivot when pixel-grid is on)
     *
     * ピクセルグリッド仕様 / Pixel-grid behavior:
     *   基準点固定と「幅・高さの整数化」を優先する。基準点も整数化するため、
     *   基準点でない側の端（右端／下端など）は 0.5px 単位になる場合がある。
     *   Anchor-fixed and integer width/height take priority; since the pivot is also integerized,
     *   the non-anchored edges (e.g. right/bottom) may land on 0.5px.
     *
     * effectiveRatioX/Y は整数化後の実効倍率（オブジェクト変形はこれを使い、アートボードと一致させる）。
     * effectiveRatioX/Y are the post-rounding effective ratios; object scaling uses them so objects match the artboard.
     * @param {number[]} originalRect - 元のアートボード矩形 [左,上,右,下]
     * @param {object} scaleSettings - readScaleSettings() の戻り値
     * @returns {object} { rect:[左,上,右,下], pivotX, pivotY, effectiveRatioX, effectiveRatioY }
     */
    function computeNewGeometry(originalRect, scaleSettings) {
        var width = originalRect[2] - originalRect[0];
        var height = originalRect[1] - originalRect[3];
        /* 1. 元の矩形から基準点（固定される点）を計算 / pivot (fixed point) from the original rect */
        var pivotX = originalRect[0] + scaleSettings.anchorFx * width;
        var pivotY = originalRect[1] - scaleSettings.anchorFy * height;

        /* 2. 新しい幅・高さ / new width & height */
        var newWidth = width * scaleSettings.ratioX;
        var newHeight = height * scaleSettings.ratioY;

        if (scaleSettings.pixelGrid) {
            /* 3. 幅・高さを整数化（0以下防止） / integerize size (prevent <= 0) */
            newWidth = Math.round(newWidth);
            newHeight = Math.round(newHeight);
            if (newWidth < 1) { newWidth = 1; }
            if (newHeight < 1) { newHeight = 1; }
            /* 4. 基準点も整数化 / integerize the pivot too */
            pivotX = Math.round(pivotX);
            pivotY = Math.round(pivotY);
        }

        /* 5. 基準点を固定して矩形を再計算（Y座標は下方向がマイナス） / rebuild the rect keeping the pivot fixed (Y grows downward negative) */
        var newLeft = pivotX - scaleSettings.anchorFx * newWidth;
        var newTop = pivotY + scaleSettings.anchorFy * newHeight;
        return {
            rect: [newLeft, newTop, newLeft + newWidth, newTop - newHeight],
            pivotX: pivotX,
            pivotY: pivotY,
            /* 整数化後の実効倍率（幅・高さが0のケースは1倍扱い） / post-rounding effective ratio (treat 0 size as 1x) */
            effectiveRatioX: width !== 0 ? newWidth / width : 1,
            effectiveRatioY: height !== 0 ? newHeight / height : 1
        };
    }

    /**
     * ダイアログの対象指定から処理するアートボード索引配列を得る / Resolve target artboard indices
     * @param {Window} resizeDialog - 設定を読むダイアログ
     * @param {number} artboardCount - アートボード総数
     * @returns {number[]|null} 索引配列（不正は null）
     */
    function resolveTargetIndices(resizeDialog, artboardCount) {
        if (resizeDialog.currentRadio.value) {
            /* 現在（起動時）のアクティブアートボードのみ / only the active artboard at launch */
            return [resizeDialog.currentArtboardIndex];
        }
        if (resizeDialog.allRadio.value) {
            var allIndices = [];
            for (var i = 0; i < artboardCount; i++) { allIndices.push(i); }
            return allIndices;
        }
        return parseArtboardSelection(resizeDialog.selectInput.text, artboardCount);
    }

    /**
     * ダイアログの現在値を読み取る / Read the current dialog settings
     * @param {Window} resizeDialog - 設定を読むダイアログ
     * @param {number} artboardCount - アートボード総数
     * @returns {object|null} indices / ratioX / ratioY / anchorFx / anchorFy / scaleObjects / pixelGrid（不正なら null）
     */
    function readScaleSettings(resizeDialog, artboardCount) {
        var targetIndices = resolveTargetIndices(resizeDialog, artboardCount);
        if (!targetIndices) { return null; }
        var scalePercent = parseFloat(resizeDialog.scaleInput.text);
        if (isNaN(scalePercent) || scalePercent <= 0) { return null; }
        var scaleRatio = scalePercent / 100; /* スケールは一意（縦横同率） / Uniform scale (same ratio for both axes) */
        return {
            indices: targetIndices,
            ratioX: scaleRatio,
            ratioY: scaleRatio,
            anchorFx: resizeDialog.anchorX, /* 0=左,0.5=中央,1=右 / 0=left,0.5=center,1=right */
            anchorFy: resizeDialog.anchorY, /* 0=上,0.5=中央,1=下 / 0=top,0.5=center,1=bottom */
            scaleObjects: resizeDialog.scaleObjectsCheckbox.value,
            pixelGrid: resizeDialog.pixelGridCheckbox.value
        };
    }

    // =========================================
    // プレビューと確定 / Preview & commit
    // =========================================

    /**
     * プレビューと確定を行うコントローラを生成する / Create the preview & commit controller
     * プレビューはアートボード矩形のみを更新（完全可逆・線幅に無関係）。オブジェクトはOK確定時に一度だけ resize() する。
     * The live preview updates artboard rectangles only (fully reversible, unrelated to line width);
     * objects are scaled once with resize() at commit. Initial state is the baseline for restore/re-apply to limit drift.
     * オブジェクトは複数アートボードにまたがっても所属ルールで一意化し1回だけ変形する。
     * @param {Document} doc - 対象ドキュメント
     * @param {Window} resizeDialog - 設定を読むダイアログ
     * @param {number} artboardCount - アートボード総数
     * @returns {object} { update, commit, restore, restoreSelectionAndActive, hasError }
     */
    function createPreviewController(doc, resizeDialog, artboardCount) {
        /* 初期状態を保存（キャンセル・確定時に復元） / Save initial state (restored on cancel/commit) */
        var originalActiveIndex = doc.artboards.getActiveArtboardIndex();
        var originalSelection = [];
        var initialSelection = doc.selection;
        if (initialSelection && typeof initialSelection.length === "number") {
            for (var i = 0; i < initialSelection.length; i++) { originalSelection.push(initialSelection[i]); }
        }

        var capturedRects = {};  /* index -> [元rect] / index -> original rect */
        var lastError = null;    /* 直近のエラー / last error */

        /**
         * アートボードの元rectを一度だけ保存する / Capture an artboard's original rect once
         * @param {number} artboardIndex - アートボードの索引
         * @returns {number[]} 元の矩形
         */
        function captureRect(artboardIndex) {
            if (!capturedRects[artboardIndex]) {
                var artboardRect = doc.artboards[artboardIndex].artboardRect;
                capturedRects[artboardIndex] = [artboardRect[0], artboardRect[1], artboardRect[2], artboardRect[3]];
            }
            return capturedRects[artboardIndex];
        }

        /**
         * アートボード矩形を初期スナップショットへ戻す（成功可否を返す） / Reset artboard rects to the initial snapshot (returns success)
         * @returns {object} { success:boolean[, error:Error] }
         */
        function restore() {
            try {
                for (var key in capturedRects) {
                    if (!capturedRects.hasOwnProperty(key)) { continue; }
                    doc.artboards[Number(key)].artboardRect = capturedRects[key];
                }
                return { success: true };
            } catch (e) {
                /* 復元失敗時は履歴（capturedRects）を消さない / keep the history (capturedRects) on failure */
                return { success: false, error: e };
            }
        }

        /**
         * 元の選択状態とアクティブアートボードを可能な範囲で復元する / Restore the original selection and active artboard as far as possible
         * @returns {void}
         */
        function restoreSelectionAndActive() {
            try { doc.selection = null; } catch (e0) {}
            for (var i = 0; i < originalSelection.length; i++) {
                /* 削除・無効化された項目は個別にスキップ / skip deleted/invalid items individually */
                try { originalSelection[i].selected = true; } catch (e1) {}
            }
            try { doc.artboards.setActiveArtboardIndex(originalActiveIndex); } catch (e2) {}
        }

        /**
         * 昇順の対象索引配列を返す（所属ルールのタイブレーク用） / Return target indices sorted ascending (for ownership tie-breaks)
         * @param {object} scaleSettings - readScaleSettings() の戻り値
         * @returns {number[]} 昇順の索引
         */
        function sortedTargetsOf(scaleSettings) {
            var sortedIndices = scaleSettings.indices.concat();
            sortedIndices.sort(function (indexA, indexB) { return indexA - indexB; });
            return sortedIndices;
        }

        /**
         * プレビュー：対象アートボードの矩形だけを更新する（オブジェクトは触らない） / Preview: update only the target artboard rectangles (objects untouched)
         * @returns {void}
         */
        function update() {
            lastError = null;
            try {
                restore(); /* まず矩形を初期状態へ / reset rects to the initial state first */

                var scaleSettings = readScaleSettings(resizeDialog, artboardCount);
                if (!scaleSettings) { app.redraw(); return; } /* 不正入力中は何も適用しない / apply nothing while input is invalid */

                var sortedTargets = sortedTargetsOf(scaleSettings);
                for (var i = 0; i < sortedTargets.length; i++) {
                    var artboardIndex = sortedTargets[i];
                    captureRect(artboardIndex);
                    doc.artboards[artboardIndex].artboardRect = computeNewGeometry(capturedRects[artboardIndex], scaleSettings).rect;
                }
            } catch (e) {
                lastError = e;
                restore();
            }
            app.redraw();
        }

        /**
         * 確定：矩形を初期状態へ戻し、初期状態から一度だけ「矩形＋オブジェクト」を適用する（線幅も正しく拡縮・累積誤差なし）
         * Commit: reset rects, then apply "rects + objects" once from the initial state (correct line width, no drift)
         * @returns {void}
         */
        function commit() {
            lastError = null;
            /* position/geometricBounds を artboardRect と同じドキュメント座標で扱うため明示設定（終了時に復元） / Force document coordinates so position/geometricBounds match artboardRect (restored at the end) */
            var previousCoordinateSystem = app.coordinateSystem;
            app.coordinateSystem = CoordinateSystem.DOCUMENTCOORDINATESYSTEM;
            try {
                restore(); /* プレビューの矩形変更を初期状態へ戻す / reset the preview rects to the initial state */

                var scaleSettings = readScaleSettings(resizeDialog, artboardCount);
                if (!scaleSettings) { app.redraw(); return; }

                var sortedTargets = sortedTargetsOf(scaleSettings);
                var i;
                for (i = 0; i < sortedTargets.length; i++) { captureRect(sortedTargets[i]); }
                var ownedObjects = scaleSettings.scaleObjects
                    ? collectOwnedObjectsByArtboard(doc, sortedTargets, capturedRects)
                    : {};

                var committedResizes = []; /* 失敗時のベストエフォート巻き戻し用 / for best-effort rollback on failure */
                for (i = 0; i < sortedTargets.length; i++) {
                    var artboardIndex = sortedTargets[i];
                    var geometry = computeNewGeometry(capturedRects[artboardIndex], scaleSettings);

                    /* オブジェクトはアートボードと同じ基準点・実効倍率（整数化後）で拡縮 / Objects use the artboard's pivot and post-rounding effective ratio */
                    if (scaleSettings.scaleObjects) {
                        var ownedItems = ownedObjects[artboardIndex] || [];
                        var resizeResult = resizeItemsAbout(ownedItems, geometry.effectiveRatioX, geometry.effectiveRatioY, geometry.pivotX, geometry.pivotY);
                        committedResizes.push({ items: resizeResult.transformedItems, ratioX: geometry.effectiveRatioX, ratioY: geometry.effectiveRatioY, pivotX: geometry.pivotX, pivotY: geometry.pivotY });
                        if (!resizeResult.success) {
                            /* 途中失敗：確定済みを逆拡縮し、矩形を戻して中止 / mid-way failure: inverse-resize committed items, reset rects, and stop */
                            lastError = resizeResult.error;
                            for (var j = committedResizes.length - 1; j >= 0; j--) {
                                resizeItemsAbout(committedResizes[j].items, 1 / committedResizes[j].ratioX, 1 / committedResizes[j].ratioY, committedResizes[j].pivotX, committedResizes[j].pivotY);
                            }
                            restore();
                            app.redraw();
                            return;
                        }
                    }

                    doc.artboards[artboardIndex].artboardRect = geometry.rect;
                }
            } catch (e) {
                lastError = e;
                restore();
            } finally {
                /* 座標系を元へ戻す / restore the coordinate system */
                try { app.coordinateSystem = previousCoordinateSystem; } catch (e2) {}
            }
            app.redraw();
        }

        return {
            update: update,
            commit: commit,
            restore: restore,
            restoreSelectionAndActive: restoreSelectionAndActive,
            hasError: function () { return lastError !== null; }
        };
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * キャンセル・失敗時の後始末：矩形と選択・アクティブアートボードを戻して再描画する
     * @param {object} previewController - createPreviewController() の戻り値
     * @returns {object} 矩形の復元結果 { success:boolean[, error:Error] }
     */
    function revertAll(previewController) {
        var restoreResult = previewController.restore();
        previewController.restoreSelectionAndActive();
        app.redraw();
        return restoreResult;
    }

    /**
     * エントリポイント：ダイアログを出し、アートボードサイズをプレビューしつつ確定でオブジェクトも拡縮する
     * Entry point: show dialog, preview artboard size, scale objects on commit
     * @returns {void}
     */
    function resizeArtboards() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }

        var doc = app.activeDocument;
        var rulerUnit = getUnitInfo();
        var artboardCount = doc.artboards.length;

        /* アクティブアートボードの現在サイズを初期値にする / Use the active artboard's current size as defaults */
        var activeRect = doc.artboards[doc.artboards.getActiveArtboardIndex()].artboardRect;
        var defaultWidth = formatNumber((activeRect[2] - activeRect[0]) / rulerUnit.pointsPerUnit);
        var defaultHeight = formatNumber((activeRect[1] - activeRect[3]) / rulerUnit.pointsPerUnit);

        var resizeDialog = createResizeDialog(rulerUnit.label, defaultWidth, defaultHeight, artboardCount);
        /* 「現在のアートボード」対象用に起動時のアクティブ索引を保持 / Remember the launch-time active index for the "Current artboard" target */
        resizeDialog.currentArtboardIndex = doc.artboards.getActiveArtboardIndex();

        /* プレビュー配線（矩形のみ）＋表示時にスケール欄へフォーカス / Wire up the preview (rects only) and focus the scale field on show */
        var previewController = createPreviewController(doc, resizeDialog, artboardCount);
        resizeDialog.onPreview = previewController.update;
        resizeDialog.onShow = function () {
            previewController.update();
            resizeDialog.scaleInput.active = true;
        };

        if (resizeDialog.show() !== 1) {
            /* キャンセル：矩形を復元（失敗は通知）し、選択・アクティブアートボードを戻す / Cancel: restore rects (notify on failure), restore selection & active artboard */
            if (!revertAll(previewController).success) { alert(getLabel("alert.restoreError")); }
            return;
        }

        /* OK：不正値は矩形を戻してエラー表示 / OK: revert rects and report on invalid input */
        if (!readScaleSettings(resizeDialog, artboardCount)) {
            revertAll(previewController);
            alert(resolveTargetIndices(resizeDialog, artboardCount) ? getLabel("alert.invalidNumber") : getLabel("alert.invalidSelection"));
            return;
        }

        /* 確定：矩形を戻し、resize()で「矩形＋オブジェクト」を一度だけ適用 / Commit: reset rects, then apply rects + objects once via resize() */
        previewController.commit();
        if (previewController.hasError()) {
            /* 変形失敗時は確定せず、矩形と選択・アクティブを復元 / on failure, do not commit; restore rects, selection & active */
            revertAll(previewController);
            alert(getLabel("alert.transformError"));
            return;
        }

        /* 適用済みの状態を確定。選択・アクティブアートボードのみ復元（完了メッセージは表示しない） / Keep the applied result; restore only selection & active artboard (no completion message) */
        previewController.restoreSelectionAndActive();
        app.redraw();
    }

    resizeArtboards();

})();
