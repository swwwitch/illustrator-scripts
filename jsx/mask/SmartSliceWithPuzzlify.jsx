#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択した画像やオブジェクトを、グリッド（格子）またはジグソーパズル形状のピースに分割し、各ピースでマスクします。
行数・列数・ピース数のほか、オフセット、オーバーラップ、バラけ、ケイ線、角丸を指定できます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartSliceWithPuzzlify.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n89f63325c0bc

### Overview

Slices the selected image or artwork into grid cells or jigsaw pieces and masks each piece.
Besides rows, columns and piece count, you can set offset, overlap, scatter, stroke and rounded corners.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartSliceWithPuzzlify.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SmartSliceWithPuzzlify";       /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.5.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-06-07";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-27";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartSliceWithPuzzlify.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartSliceWithPuzzlify.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n89f63325c0bc"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* ダイアログの初期値 / Initial dialog values */
    var DEFAULT_PIECES_PUZZLE = "25";  /* パズル時のピース数 / piece count in puzzle mode */
    var DEFAULT_PIECES_GRID   = "2";   /* グリッド時のピース数 / piece count in grid mode */
    var DEFAULT_COLUMNS       = "6";   /* 選択の寸法が取れないときの列数 / columns when the selection size is unknown */
    var DEFAULT_ROWS          = "4";   /* 選択の寸法が取れないときの行数 / rows when the selection size is unknown */
    var DEFAULT_OFFSET        = "-2";  /* オフセット / offset */
    var DEFAULT_OVERLAP       = "10";  /* オーバーラップ / overlap */
    var DEFAULT_SCATTER       = "30";  /* バラけの最大移動量 / maximum scatter distance */
    var DEFAULT_ROUND_RADIUS  = "3";   /* 角丸の半径 / round corner radius */

    // =========================================
    // レイアウト / Layout
    // =========================================

    var MODE_ROW_MARGINS     = [10, 5, 10, 5];   /* 分割方法の行の余白 / margins of the mode row */
    var PANEL_MARGINS        = [15, 20, 15, 10]; /* パネル余白 [左,上,右,下] / panel margins */
    var SHAPE_ROW_MARGINS    = [0, 10, 0, 10];   /* 形状の行の余白 / margins of the shape row */
    var PROGRESS_ROW_MARGINS = [10, 0, 10, 0];   /* プログレスバーの行の余白 / margins of the progress row */
    var PROGRESS_BAR_SIZE    = [200, 7];         /* プログレスバーの寸法 / progress bar size */
    var BUTTON_ROW_MARGINS   = [0, 0, 0, 0];     /* ボタンエリアの余白 / margins of the button row */

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
     * 現在の表示言語を取得する
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = getCurrentLang();

    /* カテゴリ分けした日英ラベル定義 / Categorized Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "グリッド／パズルに分割", en: "Slice into Grid or Puzzle" }
        },
        panel: {
            slice:   { ja: "分割", en: "Slice" },
            options: { ja: "オプション", en: "Options" }
        },
        fieldLabel: {
            mode:        { ja: "分割方法", en: "Method" },
            totalPieces: { ja: "ピース数", en: "Pieces" },
            columns:     { ja: "列数", en: "Columns" },
            rows:        { ja: "行数", en: "Rows" },
            shape:       { ja: "形状", en: "Shape" }
        },
        radio: {
            grid:        { ja: "グリッド", en: "Grid" },
            puzzle:      { ja: "パズル", en: "Puzzle" },
            traditional: { ja: "トラディショナル", en: "Traditional" },
            random:      { ja: "ランダム", en: "Random" }
        },
        checkbox: {
            offset:       { ja: "オフセット", en: "Offset" },
            overlap:      { ja: "オーバーラップ", en: "Overlap" },
            scatter:      { ja: "バラけさせる", en: "Scatter" },
            stroke:       { ja: "ケイ線を追加", en: "Add stroke" },
            roundCorners: { ja: "角丸", en: "Round corners" }
        },
        tooltip: {
            modeGrid:         { ja: "画像を格子状に切り分けます。", en: "Cuts the image into a plain grid." },
            modePuzzle:       { ja: "画像をジグソーパズルのピース状に切り分けます。", en: "Cuts the image into jigsaw puzzle pieces." },
            totalPieces: {
                ja: "作るピースのおおよその数です。オブジェクトを1つ選択しているときは、その縦横比から行数・列数を決めます。",
                en: "Approximate number of pieces. With one object selected, the rows and columns follow from its aspect ratio."
            },
            columns:          { ja: "横に並べるピースの数です。0 にすると行数と縦横比から決めます。", en: "Pieces across. Enter 0 to derive it from the rows and aspect ratio." },
            rows:             { ja: "縦に並べるピースの数です。0 にすると列数と縦横比から決めます。", en: "Pieces down. Enter 0 to derive it from the columns and aspect ratio." },
            shapeTraditional: { ja: "はめ込みの突起を規則的に並べた、よくあるパズル形状にします。", en: "Uses the familiar puzzle shape with regularly placed tabs." },
            shapeRandom:      { ja: "突起の向きをランダムにします。", en: "Randomizes the direction of the tabs." },
            offset:           { ja: "ピースの輪郭をずらします。マイナスで内側に縮みます。", en: "Offsets the outline of each piece. A negative value shrinks it inward." },
            overlap:          { ja: "隣り合うピースを重ねる幅です。継ぎ目のすき間を防ぎます。", en: "How far neighbouring pieces overlap. Use it to hide the seams." },
            scatter:          { ja: "切り分けたピースを少しずつずらして散らします。", en: "Nudges the finished pieces apart so they scatter." },
            scatterDistance:  { ja: "ピースをずらす最大距離です。", en: "Maximum distance a piece moves." },
            stroke:           { ja: "各ピースに線を追加します。", en: "Adds a stroke to each piece." },
            roundCorners:     { ja: "各ピースに効果［角を丸くする］を適用します。", en: "Applies the Round Corners effect to each piece." },
            roundRadius:      { ja: "角丸の半径です。", en: "Corner radius." }
        },
        button: {
            ok:     { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        alert: {
            noSelection:       { ja: "分割するオブジェクトを選択してください。", en: "Select the artwork to slice." },
            symbolizeMultiple: { ja: "複数オブジェクトのシンボル化に失敗しました：", en: "Failed to symbolize multiple objects: " },
            symbolizeRaster:   { ja: "埋め込み画像のシンボル化に失敗しました：", en: "Failed to symbolize embedded artwork: " },
            symbolizeVector:   { ja: "ベクターオブジェクトのシンボル化に失敗しました：", en: "Failed to symbolize vector artwork: " },
            maskNotPath: {
                ja: "マスク用のパスを作れなかったため、このピースはマスクせずに残します。",
                en: "No mask path was available, so this piece is left unmasked."
            },
            offsetNoPath: {
                ja: "オフセット後にパスが見つからないため、このピースのマスクをスキップします。",
                en: "No path was found after the offset, so the mask for this piece is skipped."
            },
            offsetFailed:      { ja: "オフセットの適用中にエラーが発生しました：", en: "An error occurred while applying the offset: " },
            scriptError:       { ja: "スクリプトの実行中にエラーが発生しました：", en: "An error occurred while running the script: " }
        }
    };

    /**
     * LABELS からカテゴリを辿って現在の言語のラベルを取得する（例: getLabel("radio", "grid")）
     * @param {...string} keys - LABELS を辿るキー列
     * @returns {string} 該当するラベル（見つからない場合は空文字）
     */
    function getLabel() {
        var labelNode = LABELS;
        for (var i = 0; i < arguments.length; i++) {
            if (labelNode == null) break;
            labelNode = labelNode[arguments[i]];
        }
        return (labelNode && labelNode[uiLang] != null) ? labelNode[uiLang] : "";
    }

    /**
     * 項目名にコロンを付けて返す（日本語は全角、英語は半角）
     * @param {string} labelKey - LABELS.fieldLabel のキー
     * @returns {string} コロン付きの項目名
     */
    function labelText(labelKey) {
        return getLabel("fieldLabel", labelKey) + (uiLang === "ja" ? "：" : ":");
    }

    // =========================================
    // 入力補助 / Input helpers
    // =========================================

    /**
     * 数値欄を↑↓キーで増減できるようにする（shift で10刻み、option で0.1刻み）
     * @param {EditText} editText - 対象の入力欄
     * @param {boolean} allowNegative - マイナスを許すなら true
     * @param {Function} [onValueChanged] - 値を変えたあとに呼ぶ関数
     * @returns {void}
     */
    function changeValueByArrowKey(editText, allowNegative, onValueChanged) {
        editText.addEventListener("keydown", function (event) {
            if (event.keyName != "Up" && event.keyName != "Down") return;
            var value = Number(editText.text);
            if (isNaN(value)) return;

            var keyboard = ScriptUI.environment.keyboardState;
            var isUp = (event.keyName == "Up");
            if (keyboard.shiftKey) {
                /* 10の倍数にスナップ / Snap to multiples of 10 */
                value = isUp ? Math.ceil((value + 1) / 10) * 10 : Math.floor((value - 1) / 10) * 10;
            } else if (keyboard.altKey) {
                value = Math.round((value + (isUp ? 0.1 : -0.1)) * 10) / 10;
            } else {
                value = Math.round(value + (isUp ? 1 : -1));
            }

            if (!allowNegative && value < 0) value = 0;

            event.preventDefault();
            editText.text = value;
            if (typeof onValueChanged === "function") onValueChanged(editText, value);
        });
    }

    /**
     * 入力欄の数値を pt に換算する（数値でなければ 0）
     * @param {EditText} editText - 定規の単位で入力された欄
     * @returns {number} pt 値
     */
    function readLengthInPoints(editText) {
        var inputValue = parseFloat(editText.text);
        return (isNaN(inputValue) ? 0 : inputValue) * getUnitInfo().pointsPerUnit;
    }

    /**
     * ピース数と縦横比から行数・列数を計算する
     * @param {number} artworkWidth - 対象の幅
     * @param {number} artworkHeight - 対象の高さ
     * @param {number} pieceCount - 作りたいピース数
     * @returns {{rows: number, columns: number}} 行数と列数
     */
    function calcGridSizeFromPieceCount(artworkWidth, artworkHeight, pieceCount) {
        var aspectRatio = artworkWidth / artworkHeight;
        var columns = Math.round(Math.sqrt(pieceCount * aspectRatio));
        if (columns < 1) columns = 1;
        var rows = Math.round(pieceCount / columns);
        if (rows < 1) rows = 1;
        return { rows: rows, columns: columns };
    }

    /**
     * 1つだけ選択している対象の寸法を返す（対象外なら null）
     * @param {Document} doc - 対象ドキュメント
     * @returns {{width: number, height: number}|null} 幅と高さ
     */
    function getSelectedArtworkSize(doc) {
        if (doc.selection.length != 1) return null;
        var selectedItem = doc.selection[0];
        var sizableTypes = { RasterItem: 1, PlacedItem: 1, SymbolItem: 1, PathItem: 1, GroupItem: 1, CompoundPathItem: 1 };
        if (!sizableTypes[selectedItem.typename]) return null;
        var bounds = selectedItem.geometricBounds;
        return { width: bounds[2] - bounds[0], height: Math.abs(bounds[1] - bounds[3]) };
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 項目名＋数値欄の組を追加する
     * @param {Group|Panel} parentContainer - 追加先
     * @param {string} labelKey - LABELS.fieldLabel のキー
     * @param {string} defaultText - 初期値
     * @param {number} characters - 欄の文字数
     * @param {string} tooltipKey - LABELS.tooltip のキー
     * @returns {{label: StaticText, input: EditText}} 追加したコントロール
     */
    function addNumberField(parentContainer, labelKey, defaultText, characters, tooltipKey) {
        var fieldGroup = parentContainer.add("group");
        fieldGroup.orientation = "row";
        var fieldLabel = fieldGroup.add("statictext", undefined, labelText(labelKey));
        var fieldInput = fieldGroup.add("edittext", undefined, defaultText);
        fieldInput.characters = characters;
        fieldInput.helpTip = getLabel("tooltip", tooltipKey);
        return { label: fieldLabel, input: fieldInput };
    }

    /**
     * チェックボックス＋数値欄＋単位の行を追加する
     * @param {Panel} parentPanel - 追加先
     * @param {string} checkboxKey - LABELS.checkbox のキー
     * @param {string} defaultText - 初期値
     * @param {number} characters - 欄の文字数
     * @param {string} inputTooltipKey - 数値欄の LABELS.tooltip のキー
     * @returns {{checkbox: Checkbox, input: EditText, unitLabel: StaticText}} 追加したコントロール
     */
    function addCheckboxValueRow(parentPanel, checkboxKey, defaultText, characters, inputTooltipKey) {
        var valueRow = parentPanel.add("group");
        valueRow.orientation = "row";
        valueRow.alignChildren = "left";
        var rowCheckbox = valueRow.add("checkbox", undefined, getLabel("checkbox", checkboxKey));
        rowCheckbox.helpTip = getLabel("tooltip", checkboxKey);
        var rowInput = valueRow.add("edittext", undefined, defaultText);
        rowInput.helpTip = getLabel("tooltip", inputTooltipKey);
        rowInput.characters = characters;
        var rowUnitLabel = valueRow.add("statictext", undefined, getUnitInfo().label);
        return { checkbox: rowCheckbox, input: rowInput, unitLabel: rowUnitLabel };
    }

    /**
     * 分割方法（グリッド／パズル）の行を作る
     * @param {Window} dlg - ダイアログ
     * @param {Object} controls - コントロールの格納先
     * @returns {void}
     */
    function buildModeRow(dlg, controls) {
        var modeRow = dlg.add("group");
        modeRow.orientation = "row";
        modeRow.alignChildren = "left";
        modeRow.margins = MODE_ROW_MARGINS;
        modeRow.add("statictext", undefined, labelText("mode"));
        controls.modeGridRadio = modeRow.add("radiobutton", undefined, getLabel("radio", "grid"));
        controls.modeGridRadio.helpTip = getLabel("tooltip", "modeGrid");
        controls.modePuzzleRadio = modeRow.add("radiobutton", undefined, getLabel("radio", "puzzle"));
        controls.modePuzzleRadio.helpTip = getLabel("tooltip", "modePuzzle");
        controls.modeGridRadio.value = true;
        controls.modeRow = modeRow;
    }

    /**
     * 分割パネル（ピース数／列数・行数／形状／オフセット／オーバーラップ）を作る
     * @param {Group} parentGroup - 追加先
     * @param {Object} controls - コントロールの格納先
     * @returns {void}
     */
    function buildSlicePanel(parentGroup, controls) {
        var slicePanel = parentGroup.add("panel", undefined, getLabel("panel", "slice"));
        slicePanel.orientation = "column";
        slicePanel.alignChildren = "left";
        slicePanel.margins = PANEL_MARGINS;

        var totalPiecesField = addNumberField(slicePanel, "totalPieces", DEFAULT_PIECES_PUZZLE, 4, "totalPieces");
        controls.totalPiecesLabel = totalPiecesField.label;
        controls.totalPiecesInput = totalPiecesField.input;

        var gridSizeRow = slicePanel.add("group");
        gridSizeRow.orientation = "row";
        gridSizeRow.alignChildren = "left";
        controls.columnsInput = addNumberField(gridSizeRow, "columns", DEFAULT_COLUMNS, 3, "columns").input;
        controls.rowsInput = addNumberField(gridSizeRow, "rows", DEFAULT_ROWS, 3, "rows").input;

        var shapeRow = slicePanel.add("group");
        shapeRow.orientation = "row";
        shapeRow.alignChildren = ["left", "top"];
        shapeRow.margins = SHAPE_ROW_MARGINS;
        shapeRow.add("statictext", undefined, labelText("shape"));
        var shapeRadioColumn = shapeRow.add("group");
        shapeRadioColumn.orientation = "column";
        shapeRadioColumn.alignChildren = "left";
        controls.shapeTraditionalRadio = shapeRadioColumn.add("radiobutton", undefined, getLabel("radio", "traditional"));
        controls.shapeTraditionalRadio.helpTip = getLabel("tooltip", "shapeTraditional");
        controls.shapeRandomRadio = shapeRadioColumn.add("radiobutton", undefined, getLabel("radio", "random"));
        controls.shapeRandomRadio.helpTip = getLabel("tooltip", "shapeRandom");
        controls.shapeTraditionalRadio.value = true;
        controls.shapeRow = shapeRow;

        controls.offsetRow = addCheckboxValueRow(slicePanel, "offset", DEFAULT_OFFSET, 4, "offset");
        controls.overlapRow = addCheckboxValueRow(slicePanel, "overlap", DEFAULT_OVERLAP, 4, "overlap");
        controls.slicePanel = slicePanel;
    }

    /**
     * オプションパネル（バラけ／ケイ線／角丸）を作る
     * @param {Group} parentGroup - 追加先
     * @param {Object} controls - コントロールの格納先
     * @returns {void}
     */
    function buildOptionsPanel(parentGroup, controls) {
        var optionsPanel = parentGroup.add("panel", undefined, getLabel("panel", "options"));
        optionsPanel.orientation = "column";
        optionsPanel.alignChildren = "left";
        optionsPanel.margins = PANEL_MARGINS;

        controls.scatterRow = addCheckboxValueRow(optionsPanel, "scatter", DEFAULT_SCATTER, 4, "scatterDistance");

        var strokeRow = optionsPanel.add("group");
        strokeRow.orientation = "row";
        strokeRow.alignChildren = "left";
        controls.strokeCheckbox = strokeRow.add("checkbox", undefined, getLabel("checkbox", "stroke"));
        controls.strokeCheckbox.helpTip = getLabel("tooltip", "stroke");

        controls.roundCornerRow = addCheckboxValueRow(optionsPanel, "roundCorners", DEFAULT_ROUND_RADIUS, 5, "roundRadius");
        controls.optionsPanel = optionsPanel;
    }

    /**
     * プログレスバー（処理中のみ表示）とボタンエリアを作る
     * @param {Window} dlg - ダイアログ
     * @param {Object} controls - コントロールの格納先
     * @returns {void}
     */
    function buildProgressAndButtons(dlg, controls) {
        /* プログレスバーとボタンを同じ位置に重ね、非表示の行でボタンの上に余白ができないようにする
           Stack the progress bar and buttons so the hidden row adds no gap above the buttons */
        var bottomStack = dlg.add("group");
        bottomStack.orientation = "stack";
        bottomStack.alignment = ["fill", "bottom"];

        var progressRow = bottomStack.add("group");
        progressRow.orientation = "column";
        progressRow.alignment = ["fill", "center"];
        progressRow.alignChildren = "fill";
        progressRow.margins = PROGRESS_ROW_MARGINS;
        controls.progressBar = progressRow.add("progressbar", undefined, 0, 100);
        controls.progressBar.preferredSize = PROGRESS_BAR_SIZE;
        progressRow.visible = false;
        controls.progressRow = progressRow;

        var btnRowGroup = bottomStack.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.margins = BUTTON_ROW_MARGINS;
        btnRowGroup.alignment = "center";
        controls.btnCancel = btnRowGroup.add("button", undefined, getLabel("button", "cancel"), { name: "cancel" });
        controls.btnOK = btnRowGroup.add("button", undefined, getLabel("button", "ok"), { name: "ok" });
        controls.btnOK.active = true;
        controls.btnRowGroup = btnRowGroup;
    }

    /**
     * チェックボックス付きの行を有効／無効にする（数値欄はチェック時のみ有効）
     * @param {Object} valueRow - addCheckboxValueRow() の戻り値
     * @param {boolean} rowEnabled - 行を有効にするなら true
     * @returns {void}
     */
    function setValueRowEnabled(valueRow, rowEnabled) {
        valueRow.checkbox.enabled = rowEnabled;
        valueRow.input.enabled = rowEnabled && valueRow.checkbox.value;
        valueRow.unitLabel.enabled = rowEnabled && valueRow.checkbox.value;
    }

    /**
     * 分割方法とチェック状態に合わせて各コントロールを有効／無効にする
     * @param {Object} controls - ダイアログのコントロール
     * @returns {void}
     */
    function syncEnabledStates(controls) {
        var isPuzzle = controls.modePuzzleRadio.value;
        /* パズル時のみ有効 / Puzzle only */
        controls.totalPiecesLabel.enabled = isPuzzle;
        controls.totalPiecesInput.enabled = isPuzzle;
        controls.shapeRow.enabled = isPuzzle;
        setValueRowEnabled(controls.offsetRow, isPuzzle);
        setValueRowEnabled(controls.scatterRow, isPuzzle);
        /* グリッド時のみ有効 / Grid only */
        setValueRowEnabled(controls.overlapRow, !isPuzzle);
        setValueRowEnabled(controls.roundCornerRow, !isPuzzle);
    }

    /**
     * 分割方法ごとの初期値に戻す
     * @param {Object} controls - ダイアログのコントロール
     * @returns {void}
     */
    function applyModeDefaults(controls) {
        controls.totalPiecesInput.text = controls.modeGridRadio.value ? DEFAULT_PIECES_GRID : DEFAULT_PIECES_PUZZLE;
        controls.offsetRow.checkbox.value = false;
        controls.offsetRow.input.text = DEFAULT_OFFSET;
        controls.overlapRow.checkbox.value = false;
        controls.overlapRow.input.text = DEFAULT_OVERLAP;
        controls.scatterRow.checkbox.value = false;
        controls.scatterRow.input.text = DEFAULT_SCATTER;
        controls.strokeCheckbox.value = false;
        controls.roundCornerRow.checkbox.value = false;
        controls.roundCornerRow.input.text = DEFAULT_ROUND_RADIUS;
    }

    /**
     * ダイアログのイベントを結び付け、初期状態を整える
     * @param {Object} controls - ダイアログのコントロール
     * @param {{width: number, height: number}|null} artworkSize - 選択対象の寸法
     * @returns {void}
     */
    function bindDialogEvents(controls, artworkSize) {
        /* ピース数から行数・列数を決める / Derive rows and columns from the piece count */
        function updateGridSizeFromPieces() {
            var pieceCount = parseInt(controls.totalPiecesInput.text, 10);
            if (isNaN(pieceCount) || pieceCount < 1 || !artworkSize) return;
            var gridSize = calcGridSizeFromPieceCount(artworkSize.width, artworkSize.height, pieceCount);
            controls.rowsInput.text = String(gridSize.rows);
            controls.columnsInput.text = String(gridSize.columns);
        }

        function onSyncEnabled() {
            syncEnabledStates(controls);
        }

        function onModeChange() {
            applyModeDefaults(controls);
            updateGridSizeFromPieces();
            syncEnabledStates(controls);
        }

        controls.totalPiecesInput.onChanging = updateGridSizeFromPieces;
        controls.modeGridRadio.onClick = onModeChange;
        controls.modePuzzleRadio.onClick = onModeChange;
        controls.offsetRow.checkbox.onClick = onSyncEnabled;
        controls.overlapRow.checkbox.onClick = onSyncEnabled;
        controls.scatterRow.checkbox.onClick = onSyncEnabled;
        controls.roundCornerRow.checkbox.onClick = onSyncEnabled;

        changeValueByArrowKey(controls.columnsInput, false);
        changeValueByArrowKey(controls.rowsInput, false);
        changeValueByArrowKey(controls.totalPiecesInput, false, updateGridSizeFromPieces);
        changeValueByArrowKey(controls.overlapRow.input, false);
        changeValueByArrowKey(controls.roundCornerRow.input, false);
        changeValueByArrowKey(controls.offsetRow.input, true);
        changeValueByArrowKey(controls.scatterRow.input, false);

        onModeChange();
    }

    /**
     * ダイアログを作る
     * @param {{width: number, height: number}|null} artworkSize - 選択対象の寸法
     * @returns {Object} ダイアログ（dialog）と各コントロール
     */
    function buildDialog(artworkSize) {
        var dlg = new Window("dialog", getLabel("dialog", "title") + " " + SCRIPT_VERSION);
        dlg.orientation = "column";
        var controls = { dialog: dlg };

        buildModeRow(dlg, controls);

        var panelStack = dlg.add("group");
        panelStack.orientation = "column";
        panelStack.alignChildren = "fill";
        buildSlicePanel(panelStack, controls);
        buildOptionsPanel(panelStack, controls);

        buildProgressAndButtons(dlg, controls);
        bindDialogEvents(controls, artworkSize);
        return controls;
    }

    /**
     * ダイアログの値を分割の設定として読み取る（分割方法で使わない項目は無効扱い）
     * @param {Object} controls - ダイアログのコントロール
     * @returns {Object} 分割の設定
     */
    function readSliceSettings(controls) {
        var isGridMode = controls.modeGridRadio.value;
        return {
            isGridMode: isGridMode,
            isRandomShape: controls.shapeRandomRadio.value,
            columnCount: Math.round(Number(controls.columnsInput.text)),
            rowCount: Math.round(Number(controls.rowsInput.text)),
            shouldApplyOffset: !isGridMode && controls.offsetRow.checkbox.value,
            offsetInPoints: readLengthInPoints(controls.offsetRow.input),
            overlapInPoints: (isGridMode && controls.overlapRow.checkbox.value) ? readLengthInPoints(controls.overlapRow.input) : 0,
            shouldScatter: !isGridMode && controls.scatterRow.checkbox.value,
            scatterDistance: readLengthInPoints(controls.scatterRow.input),
            shouldAddStroke: controls.strokeCheckbox.value,
            shouldApplyRoundCorners: isGridMode && controls.roundCornerRow.checkbox.value,
            roundRadiusInPoints: readLengthInPoints(controls.roundCornerRow.input)
        };
    }

    /**
     * 処理中の表示（入力を無効化してプログレスバーを出す）に切り替える
     * @param {Object} controls - ダイアログのコントロール
     * @returns {void}
     */
    function showProgressState(controls) {
        controls.modeRow.enabled = false;
        controls.slicePanel.enabled = false;
        controls.optionsPanel.enabled = false;
        controls.btnRowGroup.visible = false;
        controls.progressRow.visible = true;
        controls.progressBar.value = 0;
        controls.dialog.layout.layout(true);
        controls.dialog.update();
    }

    // =========================================
    // 元オブジェクトの準備 / Source preparation
    // =========================================

    /**
     * オブジェクトをシンボル化し、同じ位置にインスタンスを置いて元を削除する
     * @param {Document} doc - 対象ドキュメント
     * @param {PageItem} sourceItem - シンボル化する対象
     * @returns {SymbolItem} 置き換えたインスタンス
     */
    function convertToSymbolItem(doc, sourceItem) {
        var sourceLeft = sourceItem.left;
        var sourceTop = sourceItem.top;
        var sourceParent = sourceItem.parent;
        var createdSymbol = doc.symbols.add(sourceItem);
        var symbolInstance = sourceParent.symbolItems.add(createdSymbol);
        symbolInstance.left = sourceLeft;
        symbolInstance.top = sourceTop;
        sourceItem.remove();
        return symbolInstance;
    }

    /**
     * 選択中のオブジェクトを重ね順を保ったまま1つのグループにまとめる
     * @param {Document} doc - 対象ドキュメント
     * @returns {GroupItem} まとめたグループ
     */
    function groupSelectedItems(doc) {
        var selectedItems = [];
        for (var i = 0; i < doc.selection.length; i++) {
            selectedItems.push(doc.selection[i]);
        }
        var selectionGroup = doc.groupItems.add();
        for (var j = selectedItems.length - 1; j >= 0; j--) {
            selectedItems[j].move(selectionGroup, ElementPlacement.PLACEATBEGINNING);
        }
        return selectionGroup;
    }

    /**
     * 選択を分割できる形に整える（複数選択・埋め込み画像・ベクターはシンボル化、
     * 画像やシンボルはマスク用の矩形を用意）
     * @param {Document} doc - 対象ドキュメント
     * @returns {{contentSourceItem: PageItem, maskSourceItem: PageItem, isTemporaryBoundsRect: boolean}|null} 準備した対象（失敗時は null）
     */
    function prepareSourceItems(doc) {
        var vectorTypes = { PathItem: 1, GroupItem: 1, CompoundPathItem: 1 };
        var failureAlertKey = "symbolizeMultiple";
        var workingItem;

        try {
            if (doc.selection.length > 1) {
                workingItem = convertToSymbolItem(doc, groupSelectedItems(doc));
            } else {
                workingItem = doc.selection[0];
                if (workingItem.typename === "RasterItem") {
                    failureAlertKey = "symbolizeRaster";
                    workingItem = convertToSymbolItem(doc, workingItem);
                } else if (vectorTypes[workingItem.typename]) {
                    failureAlertKey = "symbolizeVector";
                    workingItem = convertToSymbolItem(doc, workingItem);
                }
            }
        } catch (e) {
            alert(getLabel("alert", failureAlertKey) + e);
            return null;
        }

        /* 配置画像とシンボルは外接矩形をマスクの元にする / Placed images and symbols get a bounding rectangle as the mask source */
        if (workingItem.typename === "PlacedItem" || workingItem.typename === "SymbolItem") {
            var sourceBounds = workingItem.geometricBounds;
            var rectWidth = sourceBounds[2] - sourceBounds[0];
            var rectHeight = Math.abs(sourceBounds[1] - sourceBounds[3]);
            return {
                contentSourceItem: workingItem,
                maskSourceItem: doc.pathItems.rectangle(sourceBounds[1], sourceBounds[0], rectWidth, rectHeight),
                isTemporaryBoundsRect: true
            };
        }
        return { contentSourceItem: null, maskSourceItem: workingItem, isTemporaryBoundsRect: false };
    }

    /**
     * 分割の元にしたオブジェクトと一時矩形を削除する
     * @param {Object} preparedItems - prepareSourceItems() の戻り値
     * @returns {void}
     */
    function cleanupSourceItems(preparedItems) {
        if (preparedItems.contentSourceItem) preparedItems.contentSourceItem.remove();
        if (preparedItems.isTemporaryBoundsRect) preparedItems.maskSourceItem.remove();
    }

    /**
     * 0 の列数／行数を、もう一方と縦横比から補う
     * @param {number} columnCount - 列数（0 なら自動）
     * @param {number} rowCount - 行数（0 なら自動）
     * @param {number[]} bounds - マスク元の geometricBounds
     * @returns {{columnCount: number, rowCount: number}} 補った列数と行数
     */
    function resolveGridCounts(columnCount, rowCount, bounds) {
        var boundsWidth = bounds[2] - bounds[0];
        var boundsHeight = bounds[3] - bounds[1];
        if (columnCount == 0) {
            columnCount = Math.max(1, Math.round(Math.abs(boundsWidth / boundsHeight) * rowCount));
        }
        if (rowCount == 0) {
            rowCount = Math.max(1, Math.round(Math.abs(boundsHeight / boundsWidth) * columnCount));
        }
        return { columnCount: columnCount, rowCount: rowCount };
    }

    /**
     * 列数・行数の入力が分割できる組み合わせか（どちらかが1以上、もう一方は0以上）
     * @param {number} columnCount - 列数
     * @param {number} rowCount - 行数
     * @returns {boolean} 分割できるなら true
     */
    function isValidGridCount(columnCount, rowCount) {
        return columnCount >= 0 && rowCount >= 0 && (columnCount >= 1 || rowCount >= 1);
    }

    // =========================================
    // マスク形状 / Mask shapes
    // =========================================

    /**
     * 分割の基準となる格子を作る（パズル時は各ピースの突起の向きとずれも決める）
     * @param {number[]} bounds - マスク元の geometricBounds
     * @param {number} columnCount - 列数
     * @param {number} rowCount - 行数
     * @param {boolean} isPuzzle - パズル形状なら true
     * @param {boolean} isRandomShape - 突起の向きをランダムにするなら true
     * @returns {Object} 格子の情報
     */
    function buildSliceGrid(bounds, columnCount, rowCount, isPuzzle, isRandomShape) {
        var sliceGrid = {
            originX: bounds[0],
            originY: bounds[1],
            right: bounds[2],
            bottom: bounds[3],
            columnCount: columnCount,
            rowCount: rowCount,
            pieceWidth: (bounds[2] - bounds[0]) / columnCount,
            pieceHeight: (bounds[1] - bounds[3]) / rowCount,
            edgeData: null
        };
        if (!isPuzzle) return sliceGrid;

        var edgeData = new Array(rowCount);
        for (var y = 0; y < rowCount; y++) {
            edgeData[y] = new Array(columnCount);
            for (var x = 0; x < columnCount; x++) {
                var isTopOut = isRandomShape ? Math.random() < 0.5 : (x & 1) ^ (y & 1);
                var isRightOut = isRandomShape ? Math.random() < 0.5 : !((x & 1) ^ (y & 1));
                edgeData[y][x] = {
                    topOut: isTopOut,
                    rightOut: isRightOut,
                    verticalOffset: sliceGrid.pieceHeight * (Math.random() - 0.5) / 10,
                    horizontalOffset: sliceGrid.pieceWidth * (Math.random() - 0.5) / 10
                };
            }
        }
        sliceGrid.edgeData = edgeData;
        return sliceGrid;
    }

    /**
     * グリッド1マス分の矩形マスクを作る（オーバーラップ分広げ、元の範囲内に収める）
     * @param {Document} doc - 対象ドキュメント
     * @param {Object} sliceGrid - buildSliceGrid() の戻り値
     * @param {number} columnIndex - 列番号
     * @param {number} rowIndex - 行番号
     * @param {number} overlapInPoints - オーバーラップ（pt）
     * @returns {PathItem} 矩形マスク
     */
    function createGridMask(doc, sliceGrid, columnIndex, rowIndex, overlapInPoints) {
        var rectLeft = sliceGrid.originX + columnIndex * sliceGrid.pieceWidth - overlapInPoints / 2;
        var rectWidth = sliceGrid.pieceWidth + overlapInPoints;
        if (rectLeft < sliceGrid.originX) {
            rectWidth -= (sliceGrid.originX - rectLeft);
            rectLeft = sliceGrid.originX;
        }
        if (rectLeft + rectWidth > sliceGrid.right) {
            rectWidth = sliceGrid.right - rectLeft;
        }

        var rectTop = sliceGrid.originY - rowIndex * sliceGrid.pieceHeight + overlapInPoints / 2;
        var rectHeight = sliceGrid.pieceHeight + overlapInPoints;
        if (rectTop > sliceGrid.originY) {
            rectHeight -= (rectTop - sliceGrid.originY);
            rectTop = sliceGrid.originY;
        }
        if (rectTop - rectHeight < sliceGrid.bottom) {
            rectHeight = rectTop - sliceGrid.bottom;
        }

        var gridMask = doc.pathItems.rectangle(rectTop, rectLeft, rectWidth, rectHeight);
        gridMask.closed = true;
        gridMask.filled = false;
        gridMask.stroked = false;
        return gridMask;
    }

    /**
     * コーナーポイントを追加する
     * @param {PathItem} pathItem - 対象パス
     * @param {number} anchorX - X座標
     * @param {number} anchorY - Y座標
     * @returns {void}
     */
    function addCornerPoint(pathItem, anchorX, anchorY) {
        var cornerPoint = pathItem.pathPoints.add();
        cornerPoint.anchor = [anchorX, anchorY];
        cornerPoint.leftDirection = [anchorX, anchorY];
        cornerPoint.rightDirection = [anchorX, anchorY];
        cornerPoint.pointType = PointType.CORNER;
    }

    /**
     * スムーズポイントを追加する
     * @param {PathItem} pathItem - 対象パス
     * @param {number} anchorX - アンカーのX座標
     * @param {number} anchorY - アンカーのY座標
     * @param {number} leftX - 前側ハンドルのX座標
     * @param {number} leftY - 前側ハンドルのY座標
     * @param {number} rightX - 後側ハンドルのX座標
     * @param {number} rightY - 後側ハンドルのY座標
     * @returns {void}
     */
    function addCurvePoint(pathItem, anchorX, anchorY, leftX, leftY, rightX, rightY) {
        var curvePoint = pathItem.pathPoints.add();
        curvePoint.anchor = [anchorX, anchorY];
        curvePoint.leftDirection = [leftX, leftY];
        curvePoint.rightDirection = [rightX, rightY];
        curvePoint.pointType = PointType.SMOOTH;
    }

    /**
     * パズルピース1つ分のマスクパスを作る（下辺→右辺→上辺→左辺の順に突起を描く）
     * @param {Document} doc - 対象ドキュメント
     * @param {Object} sliceGrid - buildSliceGrid() の戻り値
     * @param {number} x - 列番号
     * @param {number} y - 行番号
     * @returns {PathItem} マスクパス
     */
    function createPuzzleMaskPath(doc, sliceGrid, x, y) {
        var leftX = sliceGrid.originX + x * sliceGrid.pieceWidth;
        var rightX = sliceGrid.originX + (x + 1) * sliceGrid.pieceWidth;
        var topY = sliceGrid.originY - y * sliceGrid.pieceHeight;
        var bottomY = sliceGrid.originY - (y + 1) * sliceGrid.pieceHeight;

        var maskPath = doc.pathItems.add();
        addCornerPoint(maskPath, leftX, bottomY);
        appendLowerTab(maskPath, sliceGrid, x, y);
        addCornerPoint(maskPath, rightX, bottomY);
        appendRightTab(maskPath, sliceGrid, x, y);
        addCornerPoint(maskPath, rightX, topY);
        appendUpperTab(maskPath, sliceGrid, x, y);
        addCornerPoint(maskPath, leftX, topY);
        appendLeftTab(maskPath, sliceGrid, x, y);
        maskPath.closed = true;
        return maskPath;
    }

    /**
     * 下隣のピースとの境界（y+1 行目との辺）の突起を追加する
     * @param {PathItem} maskPath - 描画中のパス
     * @param {Object} sliceGrid - buildSliceGrid() の戻り値
     * @param {number} x - 列番号
     * @param {number} y - 行番号
     * @returns {void}
     */
    function appendLowerTab(maskPath, sliceGrid, x, y) {
        if (y >= sliceGrid.rowCount - 1) return;
        var edge = sliceGrid.edgeData[y + 1][x];
        var left = sliceGrid.originX + x * sliceGrid.pieceWidth;
        var rowTop = sliceGrid.originY - y * sliceGrid.pieceHeight;
        var rowBottom = sliceGrid.originY - (y + 1) * sliceGrid.pieceHeight;
        var thirdWidth = sliceGrid.pieceWidth / 3;
        var quarterHeight = sliceGrid.pieceHeight / 4;
        var shift = edge.verticalOffset;

        if (edge.topOut) {
            addCurvePoint(maskPath,
                left + thirdWidth, rowBottom - shift,
                left + 0.67 * thirdWidth, rowBottom + 0.33 * quarterHeight - shift,
                left + 1.33 * thirdWidth, rowBottom - 0.33 * quarterHeight - shift);
            addCurvePoint(maskPath,
                left + thirdWidth, rowBottom - quarterHeight - shift,
                left + 0.67 * thirdWidth, rowBottom - quarterHeight + 0.33 * quarterHeight - shift,
                left + 1.33 * thirdWidth, rowBottom - quarterHeight - 0.33 * quarterHeight - shift);
            addCurvePoint(maskPath,
                left + 2 * thirdWidth, rowBottom - quarterHeight - shift,
                left + 1.67 * thirdWidth, rowBottom - quarterHeight - 0.33 * quarterHeight - shift,
                left + 2.33 * thirdWidth, rowBottom - quarterHeight + 0.33 * quarterHeight - shift);
            addCurvePoint(maskPath,
                left + 2 * thirdWidth, rowBottom - shift,
                left + 1.67 * thirdWidth, rowBottom - 0.33 * quarterHeight - shift,
                left + 2.33 * thirdWidth, rowBottom + 0.33 * quarterHeight - shift);
        } else {
            addCurvePoint(maskPath,
                left + thirdWidth, rowBottom - shift,
                left + 0.67 * thirdWidth, rowBottom - 0.33 * quarterHeight - shift,
                left + 1.33 * thirdWidth, rowBottom + 0.33 * quarterHeight - shift);
            addCurvePoint(maskPath,
                left + thirdWidth, rowTop - 3 * quarterHeight - shift,
                left + 0.67 * thirdWidth, rowTop - 3.5 * quarterHeight - shift,
                left + 1.33 * thirdWidth, rowTop - 2.5 * quarterHeight - shift);
            addCurvePoint(maskPath,
                left + 2 * thirdWidth, rowTop - 3 * quarterHeight - shift,
                left + 1.67 * thirdWidth, rowTop - 2.5 * quarterHeight - shift,
                left + 2.33 * thirdWidth, rowTop - 3.5 * quarterHeight - shift);
            addCurvePoint(maskPath,
                left + 2 * thirdWidth, rowBottom - shift,
                left + 1.67 * thirdWidth, rowBottom + 0.33 * quarterHeight - shift,
                left + 2.33 * thirdWidth, rowBottom - 0.33 * quarterHeight - shift);
        }
    }

    /**
     * 右隣のピースとの境界の突起を追加する
     * @param {PathItem} maskPath - 描画中のパス
     * @param {Object} sliceGrid - buildSliceGrid() の戻り値
     * @param {number} x - 列番号
     * @param {number} y - 行番号
     * @returns {void}
     */
    function appendRightTab(maskPath, sliceGrid, x, y) {
        if (x >= sliceGrid.columnCount - 1) return;
        var edge = sliceGrid.edgeData[y][x + 1];
        var left = sliceGrid.originX + x * sliceGrid.pieceWidth;
        var rowTop = sliceGrid.originY - y * sliceGrid.pieceHeight;
        var thirdHeight = sliceGrid.pieceHeight / 3;
        var quarterWidth = sliceGrid.pieceWidth / 4;
        var shift = edge.horizontalOffset;

        if (edge.rightOut) {
            addCurvePoint(maskPath,
                left + 4 * quarterWidth - shift, rowTop - 2 * thirdHeight,
                left + 3.5 * quarterWidth - shift, rowTop - 2.33 * thirdHeight,
                left + 4.5 * quarterWidth - shift, rowTop - 1.67 * thirdHeight);
            addCurvePoint(maskPath,
                left + 5 * quarterWidth - shift, rowTop - 2 * thirdHeight,
                left + 4.5 * quarterWidth - shift, rowTop - 2.33 * thirdHeight,
                left + 5.5 * quarterWidth - shift, rowTop - 1.67 * thirdHeight);
            addCurvePoint(maskPath,
                left + 5 * quarterWidth - shift, rowTop - thirdHeight,
                left + 5.5 * quarterWidth - shift, rowTop - 1.33 * thirdHeight,
                left + 4.5 * quarterWidth - shift, rowTop - 0.67 * thirdHeight);
            addCurvePoint(maskPath,
                left + 4 * quarterWidth - shift, rowTop - thirdHeight,
                left + 4.5 * quarterWidth - shift, rowTop - 1.33 * thirdHeight,
                left + 3.5 * quarterWidth - shift, rowTop - 0.67 * thirdHeight);
        } else {
            addCurvePoint(maskPath,
                left + 4 * quarterWidth - shift, rowTop - 2 * thirdHeight,
                left + 4.5 * quarterWidth - shift, rowTop - 2.33 * thirdHeight,
                left + 3.5 * quarterWidth - shift, rowTop - 1.67 * thirdHeight);
            addCurvePoint(maskPath,
                left + 3 * quarterWidth - shift, rowTop - 2 * thirdHeight,
                left + 3.5 * quarterWidth - shift, rowTop - 2.33 * thirdHeight,
                left + 2.5 * quarterWidth - shift, rowTop - 1.67 * thirdHeight);
            addCurvePoint(maskPath,
                left + 3 * quarterWidth - shift, rowTop - thirdHeight,
                left + 2.5 * quarterWidth - shift, rowTop - 1.33 * thirdHeight,
                left + 3.5 * quarterWidth - shift, rowTop - 0.67 * thirdHeight);
            addCurvePoint(maskPath,
                left + 4 * quarterWidth - shift, rowTop - thirdHeight,
                left + 3.5 * quarterWidth - shift, rowTop - 1.33 * thirdHeight,
                left + 4.5 * quarterWidth - shift, rowTop - 0.67 * thirdHeight);
        }
    }

    /**
     * 上隣のピースとの境界（y 行目の上辺）の突起を追加する
     * @param {PathItem} maskPath - 描画中のパス
     * @param {Object} sliceGrid - buildSliceGrid() の戻り値
     * @param {number} x - 列番号
     * @param {number} y - 行番号
     * @returns {void}
     */
    function appendUpperTab(maskPath, sliceGrid, x, y) {
        if (y <= 0) return;
        var edge = sliceGrid.edgeData[y][x];
        var left = sliceGrid.originX + x * sliceGrid.pieceWidth;
        var rowTop = sliceGrid.originY - y * sliceGrid.pieceHeight;
        var thirdWidth = sliceGrid.pieceWidth / 3;
        var quarterHeight = sliceGrid.pieceHeight / 4;
        var shift = edge.verticalOffset;

        if (edge.topOut) {
            addCurvePoint(maskPath,
                left + 2 * thirdWidth, rowTop - shift,
                left + 2.33 * thirdWidth, rowTop + 0.33 * quarterHeight - shift,
                left + 1.67 * thirdWidth, rowTop - 0.33 * quarterHeight - shift);
            addCurvePoint(maskPath,
                left + 2 * thirdWidth, rowTop - quarterHeight - shift,
                left + 2.33 * thirdWidth, rowTop - quarterHeight + 0.33 * quarterHeight - shift,
                left + 1.67 * thirdWidth, rowTop - quarterHeight - 0.33 * quarterHeight - shift);
            addCurvePoint(maskPath,
                left + thirdWidth, rowTop - quarterHeight - shift,
                left + 1.33 * thirdWidth, rowTop - quarterHeight - 0.33 * quarterHeight - shift,
                left + 0.67 * thirdWidth, rowTop - quarterHeight + 0.33 * quarterHeight - shift);
            addCurvePoint(maskPath,
                left + thirdWidth, rowTop - shift,
                left + 1.33 * thirdWidth, rowTop - 0.33 * quarterHeight - shift,
                left + 0.67 * thirdWidth, rowTop + 0.33 * quarterHeight - shift);
        } else {
            addCurvePoint(maskPath,
                left + 2 * thirdWidth, rowTop - shift,
                left + 2.33 * thirdWidth, rowTop - 0.33 * quarterHeight - shift,
                left + 1.67 * thirdWidth, rowTop + 0.33 * quarterHeight - shift);
            addCurvePoint(maskPath,
                left + 2 * thirdWidth, rowTop + quarterHeight - shift,
                left + 2.33 * thirdWidth, rowTop + 0.5 * quarterHeight - shift,
                left + 1.67 * thirdWidth, rowTop + 1.5 * quarterHeight - shift);
            addCurvePoint(maskPath,
                left + thirdWidth, rowTop + quarterHeight - shift,
                left + 1.33 * thirdWidth, rowTop + 1.5 * quarterHeight - shift,
                left + 0.67 * thirdWidth, rowTop + 0.5 * quarterHeight - shift);
            addCurvePoint(maskPath,
                left + thirdWidth, rowTop - shift,
                left + 1.33 * thirdWidth, rowTop + 0.33 * quarterHeight - shift,
                left + 0.67 * thirdWidth, rowTop - 0.33 * quarterHeight - shift);
        }
    }

    /**
     * 左隣のピースとの境界の突起を追加する
     * @param {PathItem} maskPath - 描画中のパス
     * @param {Object} sliceGrid - buildSliceGrid() の戻り値
     * @param {number} x - 列番号
     * @param {number} y - 行番号
     * @returns {void}
     */
    function appendLeftTab(maskPath, sliceGrid, x, y) {
        if (x <= 0) return;
        var edge = sliceGrid.edgeData[y][x];
        var left = sliceGrid.originX + x * sliceGrid.pieceWidth;
        var rowTop = sliceGrid.originY - y * sliceGrid.pieceHeight;
        var thirdHeight = sliceGrid.pieceHeight / 3;
        var quarterWidth = sliceGrid.pieceWidth / 4;
        var shift = edge.horizontalOffset;

        if (edge.rightOut) {
            addCurvePoint(maskPath,
                left - shift, rowTop - thirdHeight,
                left - 0.5 * quarterWidth - shift, rowTop - 0.67 * thirdHeight,
                left + 0.5 * quarterWidth - shift, rowTop - 1.33 * thirdHeight);
            addCurvePoint(maskPath,
                left + quarterWidth - shift, rowTop - thirdHeight,
                left + 0.5 * quarterWidth - shift, rowTop - 0.67 * thirdHeight,
                left + 1.5 * quarterWidth - shift, rowTop - 1.33 * thirdHeight);
            addCurvePoint(maskPath,
                left + quarterWidth - shift, rowTop - 2 * thirdHeight,
                left + 1.5 * quarterWidth - shift, rowTop - 1.67 * thirdHeight,
                left + 0.5 * quarterWidth - shift, rowTop - 2.33 * thirdHeight);
            addCurvePoint(maskPath,
                left - shift, rowTop - 2 * thirdHeight,
                left + 0.5 * quarterWidth - shift, rowTop - 1.67 * thirdHeight,
                left - 0.5 * quarterWidth - shift, rowTop - 2.33 * thirdHeight);
        } else {
            addCurvePoint(maskPath,
                left - shift, rowTop - thirdHeight,
                left + 0.5 * quarterWidth - shift, rowTop - 0.67 * thirdHeight,
                left - 0.5 * quarterWidth - shift, rowTop - 1.33 * thirdHeight);
            addCurvePoint(maskPath,
                left - quarterWidth - shift, rowTop - thirdHeight,
                left - 0.5 * quarterWidth - shift, rowTop - 0.67 * thirdHeight,
                left - 1.5 * quarterWidth - shift, rowTop - 1.33 * thirdHeight);
            addCurvePoint(maskPath,
                left - quarterWidth - shift, rowTop - 2 * thirdHeight,
                left - 1.5 * quarterWidth - shift, rowTop - 1.67 * thirdHeight,
                left - 0.5 * quarterWidth - shift, rowTop - 2.33 * thirdHeight);
            addCurvePoint(maskPath,
                left - shift, rowTop - 2 * thirdHeight,
                left - 0.5 * quarterWidth - shift, rowTop - 1.67 * thirdHeight,
                left + 0.5 * quarterWidth - shift, rowTop - 2.33 * thirdHeight);
        }
    }

    // =========================================
    // オフセット / Offset
    // =========================================

    /**
     * 効果［パスのオフセット］を適用して分割し、分割後のオブジェクトを返す
     * @param {Document} doc - 対象ドキュメント
     * @param {PathItem} maskPath - 対象のマスクパス（処理後は削除される）
     * @param {number} offsetInPoints - オフセット量（pt）
     * @returns {PageItem|null} 分割後のオブジェクト
     */
    function applyOffsetEffect(doc, maskPath, offsetInPoints) {
        var previousInteractionLevel = app.userInteractionLevel;
        app.userInteractionLevel = UserInteractionLevel.DONTDISPLAYALERTS;
        try {
            doc.selection = null;
            var offsetTarget = maskPath.duplicate(maskPath, ElementPlacement.PLACEAFTER);
            maskPath.remove();
            offsetTarget.selected = true;
            offsetTarget.applyEffect('<LiveEffect name="Adobe Offset Path"><Dict data="R mlim 4 R ofst ' + offsetInPoints + ' I jntp 2 "/></LiveEffect>');
            app.redraw();
            app.executeMenuCommand("expandStyle");
            var expandedItem = (doc.selection.length > 0) ? doc.selection[0] : null;
            doc.selection = null;
            return expandedItem;
        } finally {
            /* 例外時も警告の表示設定を戻す / Restore the alert level even on failure */
            app.userInteractionLevel = previousInteractionLevel;
        }
    }

    /**
     * オブジェクトからマスクに使えるパスを取り出す（グループなら最初のパス）
     * @param {PageItem|null} targetItem - 対象
     * @returns {PathItem|null} 見つかったパス
     */
    function findMaskPath(targetItem) {
        if (!targetItem) return null;
        if (targetItem.typename === "PathItem") return targetItem;
        if (targetItem.typename === "GroupItem") {
            for (var i = 0; i < targetItem.pageItems.length; i++) {
                if (targetItem.pageItems[i].typename === "PathItem") return targetItem.pageItems[i];
            }
        }
        return null;
    }

    /**
     * マスクパスにオフセットを掛けたパスを返す（パスが得られなければ警告して元の結果を返す）
     * @param {Document} doc - 対象ドキュメント
     * @param {PathItem} maskPath - 対象のマスクパス
     * @param {number} offsetInPoints - オフセット量（pt）
     * @returns {PageItem|null} マスクに使うオブジェクト
     */
    function offsetMaskPath(doc, maskPath, offsetInPoints) {
        try {
            var expandedItem = applyOffsetEffect(doc, maskPath, offsetInPoints);
            var offsetPath = findMaskPath(expandedItem);
            if (offsetPath) return offsetPath;
            alert(getLabel("alert", "offsetNoPath"));
            return expandedItem;
        } catch (e) {
            alert(getLabel("alert", "offsetFailed") + e.message);
            return maskPath;
        }
    }

    // =========================================
    // ピースの仕上げ / Piece finishing
    // =========================================

    /**
     * 中央寄りの正規乱数を返す（Box-Muller 法）
     * @returns {number} 平均0・標準偏差1の乱数
     */
    function randomGaussian() {
        var u = 0, v = 0;
        while (u === 0) u = Math.random();
        while (v === 0) v = Math.random();
        return Math.sqrt(-2.0 * Math.log(u)) * Math.cos(2.0 * Math.PI * v);
    }

    /**
     * 値を範囲内に収める
     * @param {number} value - 対象の値
     * @param {number} minValue - 下限
     * @param {number} maxValue - 上限
     * @returns {number} 範囲内に収めた値
     */
    function clampValue(value, minValue, maxValue) {
        return Math.max(minValue, Math.min(maxValue, value));
    }

    /**
     * 元オブジェクトの複製をマスクパスでクリップしたグループを作る
     * @param {Document} doc - 対象ドキュメント
     * @param {PageItem} contentSourceItem - 中身にする配置画像／シンボル
     * @param {PageItem|null} maskPath - マスクにするパス
     * @returns {GroupItem} クリップグループ
     */
    function createClippedPiece(doc, contentSourceItem, maskPath) {
        var contentCopy = contentSourceItem.duplicate();
        var clippingGroup = doc.groupItems.add();
        contentCopy.moveToBeginning(clippingGroup);
        if (maskPath && maskPath.typename === "PathItem") {
            maskPath.moveToBeginning(clippingGroup);
            maskPath.clipping = true;
            clippingGroup.clipped = true;
        } else {
            alert(getLabel("alert", "maskNotPath"));
        }
        return clippingGroup;
    }

    /**
     * ピースにバラけ・ケイ線・角丸を適用する
     * @param {Document} doc - 対象ドキュメント
     * @param {PageItem|null} pieceItem - 対象のピース
     * @param {Object} sliceSettings - readSliceSettings() の戻り値
     * @returns {void}
     */
    function finishPiece(doc, pieceItem, sliceSettings) {
        if (!pieceItem) return;
        if (sliceSettings.shouldScatter && sliceSettings.scatterDistance > 0) {
            /* ガウス分布（中央寄り）で動かし、最大移動量で抑える / Gaussian nudge clamped to the maximum distance */
            var maxDistance = sliceSettings.scatterDistance;
            var dx = clampValue(randomGaussian() * maxDistance * 0.5, -maxDistance, maxDistance);
            var dy = clampValue(randomGaussian() * maxDistance * 0.5, -maxDistance, maxDistance);
            pieceItem.transform(app.getTranslationMatrix(dx, dy));
        }
        pieceItem.selected = true;

        if (sliceSettings.shouldAddStroke) {
            doc.selection = null;
            pieceItem.selected = true;
            app.executeMenuCommand("Adobe New Stroke Shortcut");
            app.executeMenuCommand("Live Pathfinder Add");
        }
        if (sliceSettings.shouldApplyRoundCorners && sliceSettings.roundRadiusInPoints > 0) {
            pieceItem.applyEffect('<LiveEffect name="Adobe Round Corners"><Dict data="R radius ' + sliceSettings.roundRadiusInPoints + ' "/></LiveEffect>');
        }
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択オブジェクトを分割してピースを作る
     * @param {Document} doc - 対象ドキュメント
     * @param {Object} sliceSettings - readSliceSettings() の戻り値
     * @param {Function} onProgress - 進み具合を受け取る関数 (完了数, 総数)
     * @returns {void}
     */
    function executeSlice(doc, sliceSettings, onProgress) {
        var preparedItems = prepareSourceItems(doc);
        if (!preparedItems) return;

        var maskSourceItem = preparedItems.maskSourceItem;
        var contentSourceItem = preparedItems.contentSourceItem;
        maskSourceItem.selected = false;

        var bounds = maskSourceItem.geometricBounds;
        var gridCounts = resolveGridCounts(sliceSettings.columnCount, sliceSettings.rowCount, bounds);
        var sliceGrid = buildSliceGrid(bounds, gridCounts.columnCount, gridCounts.rowCount,
            !sliceSettings.isGridMode, sliceSettings.isRandomShape);

        var totalPieces = sliceGrid.rowCount * sliceGrid.columnCount;
        var finishedPieces = 0;
        onProgress(0, totalPieces);

        for (var y = 0; y < sliceGrid.rowCount; y++) {
            for (var x = 0; x < sliceGrid.columnCount; x++) {
                var maskPath;
                if (sliceSettings.isGridMode) {
                    maskPath = createGridMask(doc, sliceGrid, x, y, sliceSettings.overlapInPoints);
                } else {
                    maskPath = createPuzzleMaskPath(doc, sliceGrid, x, y);
                    if (sliceSettings.shouldApplyOffset) {
                        maskPath = offsetMaskPath(doc, maskPath, sliceSettings.offsetInPoints);
                    }
                }

                var pieceItem = contentSourceItem ? createClippedPiece(doc, contentSourceItem, maskPath) : maskPath;
                finishPiece(doc, pieceItem, sliceSettings);

                finishedPieces++;
                onProgress(finishedPieces, totalPieces);
            }
        }

        cleanupSourceItems(preparedItems);
    }

    /**
     * エントリーポイント：ダイアログを表示し、OK で分割を実行する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) return;
        var doc = app.activeDocument;
        if (doc.selection.length === 0) {
            alert(getLabel("alert", "noSelection"));
            return;
        }

        var controls = buildDialog(getSelectedArtworkSize(doc));
        var sliceDialog = controls.dialog;

        controls.btnOK.onClick = function () {
            var sliceSettings = readSliceSettings(controls);
            /* 分割できない列数・行数なら閉じずに待つ / Stay open when the counts cannot be sliced */
            if (!isValidGridCount(sliceSettings.columnCount, sliceSettings.rowCount)) return;

            showProgressState(controls);
            try {
                executeSlice(doc, sliceSettings, function (current, total) {
                    controls.progressBar.value = total > 0 ? (current / total) * 100 : 0;
                    sliceDialog.update();
                });
            } catch (e) {
                alert(getLabel("alert", "scriptError") + e);
            }
            sliceDialog.close(1);
        };

        sliceDialog.show();
    }

    main();

})();
