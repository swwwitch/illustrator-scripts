#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択オブジェクトのすべてのアンカーポイントに、マーカー（自動生成の正方形／最前面オブジェクトの複製／シンボルのインスタンス）を配置します。
ダイアログを閉じずに、ライブプレビューで仕上がりを確かめながら設定を調整できます。

詳しい仕様と注意事項は README を参照してください。

*/

/*

### Overview

Places a marker — an auto-generated square, a duplicate of the frontmost object, or a symbol instance — at every anchor point of the selection.
Settings can be adjusted with a live preview, without closing the dialog.

See the README for the full specification and notes.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AiAnchorPointMarker";          /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-07-05";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-08-12";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AiAnchorPointMarker.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AiAnchorPointMarker.md
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n757f8802dc4b"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================
    /**
     * 追加するオブジェクトの種別 / Kind of object to place
     * @type {Object}
     */
    var OBJECT_SOURCE = {
        autoGenerate: "autoGenerate", /* 正方形を自動生成 / Auto-generated square */
        frontObject: "frontObject",   /* 最前面のオブジェクト / Frontmost object */
        symbol: "symbol"              /* シンボル / Symbol */
    };

    var PREVIEW_LAYER_NAME = "__ANCHOR_MARKER_PREVIEW__"; /* プレビュー専用レイヤー名 / Preview-only layer name */
    var ANCHOR_LAYER_NAME = "_anchorpoint";               /* マーカー移動先レイヤー名 / Destination layer for markers */
    var SQUARE_SYMBOL_NAME = "アンカーポイント";           /* 自動生成シンボルの名前 / Name for the generated symbol */

    /**
     * ダイアログの初期値 / Initial dialog values
     * @type {Object}
     */
    var DEFAULTS = {
        objectSource: OBJECT_SOURCE.autoGenerate, /* 追加するオブジェクトの種類 / Kind of object to add */
        squareSize: 6,                            /* 正方形の一辺（pt） / Square edge size (pt) */
        squareColor: { r: 79, g: 128, b: 255 },   /* 塗り色（RGB） / Fill color (RGB) */
        symbolize: true,                          /* 自動生成の正方形をシンボル化して配置 / Place squares as symbol instances */
        moveToLayer: false,                       /* 配置後のマーカーを専用レイヤーへ移動 / Move markers to a dedicated layer */
        groupItems: true,                         /* 配置後のマーカーを1つのグループにまとめる / Group the placed markers */
        scalePercent: 100,                        /* 最前面オブジェクト／シンボルの拡大縮小率（%） / Scale for frontmost object / symbol (%) */
        registrationIndex: 4                      /* 基準点 0..8（行優先, 4=中央）/ Registration point 0..8 row-major (4=center) */
    };

    // =========================================
    // レイアウト / Layout
    // =========================================
    var WINDOW_MARGINS = 16;                 /* ウィンドウ外周の余白 / window margin */
    var WINDOW_SPACING = 12;                 /* ウィンドウ内の要素間隔 / window spacing */
    var PANEL_MARGINS  = [16, 20, 16, 12];   /* パネル余白 [左,上,右,下] / panel margins */
    var PANEL_SPACING  = 10;                 /* パネル内の要素間隔 / panel spacing */
    var COLUMN_SPACING = 16;                 /* 2カラムの間隔 / gap between columns */
    var WIDGET_SIZE    = 56;                 /* 9軸ウィジェットの一辺（px）/ edge size of the 9-axis widget */
    var SWATCH_SIZE    = 20;                 /* カラースウォッチの一辺（px）/ edge size of the color swatch */

    /**
     * ダイアログウィンドウの共通レイアウトを設定します。
     * @param {Window} dialogWindow - 対象のダイアログ
     * @param {number} [spacing] - 要素間隔（省略時は WINDOW_SPACING）
     * @returns {Window} 設定後のダイアログ
     */
    function setupWindow(dialogWindow, spacing) {
        dialogWindow.orientation = "column";
        dialogWindow.alignChildren = ["fill", "top"];
        dialogWindow.margins = WINDOW_MARGINS;
        dialogWindow.spacing = (typeof spacing === "number") ? spacing : WINDOW_SPACING;
        return dialogWindow;
    }

    /**
     * パネルの共通レイアウトを設定します。
     * @param {Panel} panel - 対象のパネル
     * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
     * @returns {Panel} 設定後のパネル
     */
    function setupPanel(panel, spacing) {
        panel.orientation = "column";
        panel.alignChildren = ["fill", "top"];
        panel.margins = PANEL_MARGINS;
        panel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
        return panel;
    }

    /**
     * グループ／パネルを横並び（row）に設定します。
     * @param {Group|Panel} container - 対象のグループまたはパネル
     * @param {Array<string>} alignChildren - 子要素の整列（例 ["left", "top"]）
     * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
     * @returns {Group|Panel} 設定後のコンテナ
     */
    function setupRow(container, alignChildren, spacing) {
        container.orientation = "row";
        container.alignChildren = alignChildren;
        container.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
        return container;
    }

    /**
     * 共通レイアウト済みのパネルを追加します。
     * @param {Window|Group} parent - 追加先
     * @param {string} labelText - パネルのタイトル
     * @returns {Panel} 追加したパネル
     */
    function addPanel(parent, labelText) {
        return setupPanel(parent.add("panel", undefined, labelText));
    }

    /**
     * 縦並びのカラムグループを追加します。
     * @param {Window|Group|Panel} parent - 追加先
     * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
     * @returns {Group} 追加したグループ
     */
    function addColumnGroup(parent, spacing) {
        var columnGroup = parent.add("group");
        columnGroup.orientation = "column";
        columnGroup.alignChildren = ["left", "center"];
        columnGroup.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
        return columnGroup;
    }

    // =========================================
    // ローカライズ / Localization
    // =========================================
    /**
     * 現在の UI 言語（"ja" / "en"）を判定して返します。
     * @returns {string} 言語コード
     */
    function getCurrentLang() {
        return ($.locale && $.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var currentLanguage = getCurrentLang();

    /**
     * 日英ラベル定義（カテゴリ別） / Japanese-English label definitions (by category)
     * getLabel("dialog.title") のようにドット区切りで参照する。短い文言は1行で記述。
     * @type {Object}
     */
    var LABELS = {
        dialog: {
            title: { ja: "アンカーポイントに複製", en: "Duplicate to Anchor Points" }
        },
        colorPicker: {
            title: { ja: "カラーを選択", en: "Choose Color" }
        },
        panel: {
            objectSource: { ja: "追加するオブジェクト", en: "Object to Add" },
            anchorPoint: { ja: "アンカーポイント", en: "Anchor Point" },
            options: { ja: "オプション", en: "Options" }
        },
        radio: {
            autoGenerate: { ja: "アンカーポイントを自動生成", en: "Auto-generate square" },
            frontObject: { ja: "最前面のオブジェクト", en: "Frontmost object" },
            symbol: { ja: "シンボル：", en: "Symbol:" }
        },
        label: {
            squareSize: { ja: "大きさ：", en: "Size:" },
            squareColor: { ja: "カラー：", en: "Color:" },
            scale: { ja: "スケール：", en: "Scale:" },
            unit: { ja: "pt", en: "pt" },
            percent: { ja: "%", en: "%" }
        },
        tooltip: {
            autoGenerate: { ja: "指定した大きさ・カラーの正方形を生成して各アンカーポイントに配置します。", en: "Generate a square of the given size/color and place one at each anchor point." },
            frontObject: { ja: "最前面のオブジェクトを複製して各アンカーポイントに配置します。", en: "Duplicate the frontmost object and place it at each anchor point." },
            symbol: { ja: "選択したシンボルのインスタンスを各アンカーポイントに配置します。", en: "Place an instance of the selected symbol at each anchor point." },
            symbolDropdown: { ja: "配置するシンボルを選びます。", en: "Choose the symbol to place." },
            squareSize: { ja: "正方形の一辺（pt）。↑↓：±1　Shift：±10　⌘：±0.1", en: "Square size (pt). ↑↓: ±1, Shift: ±10, ⌘: ±0.1" },
            squareColor: { ja: "正方形の塗り色。［選択...］でカラーを変更します。", en: "Fill color of the square. Click Choose... to change it." },
            symbolize: { ja: "生成した正方形をシンボル（アンカーポイント）として登録し、インスタンスで配置します。", en: "Register the generated square as a symbol and place instances." },
            moveToLayer: { ja: "配置後のマーカーを「_anchorpoint」レイヤーへ移動します。", en: "Move the placed markers to the \"_anchorpoint\" layer." },
            group: { ja: "配置後のマーカーを1つのグループにまとめます。", en: "Group the placed markers into a single group." },
            scale: { ja: "最前面オブジェクト・シンボルの拡大縮小率（%）。↑↓：±1　Shift：±10　⌘：±0.1", en: "Scale for the frontmost object / symbol (%). ↑↓: ±1, Shift: ±10, ⌘: ±0.1" },
            registration: { ja: "基準点。マーカーのどの位置をアンカーポイントに合わせるかを選びます（中央が既定）。", en: "Registration point: which part of the marker aligns to the anchor (center by default)." }
        },
        checkbox: {
            symbolize: { ja: "シンボル化", en: "Symbolize" },
            moveToLayer: { ja: "レイヤーに移動", en: "Move to layer" },
            group: { ja: "グループ化", en: "Group" }
        },
        button: {
            chooseColor: { ja: "選択...", en: "Choose..." },
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" }
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noSelection: { ja: "オブジェクトを選択してください。", en: "Please select an object." },
            noAnchorPoints: { ja: "選択オブジェクトにアンカーポイントがありません。\nパスを含むオブジェクトを選択してください。", en: "The selection has no anchor points.\nSelect an object that contains paths." },
            invalidSize: { ja: "大きさには 0 より大きい数値を入力してください。", en: "Enter a size greater than 0." },
            invalidScale: { ja: "スケールには 0 より大きい数値を入力してください。", en: "Enter a scale greater than 0." },
            noSymbolChosen: { ja: "配置するシンボルを選択してください。", en: "Please choose a symbol to place." }
        }
    };

    /**
     * ドット区切りキーで LABELS を辿り、現在の言語のラベルを返します。
     * キー漏れ・言語漏れのときは、キー自身／英語にフォールバックします。
     * @param {string} labelPath - ラベルのパス（例 "dialog.title"）
     * @returns {string} ラベル文字列
     */
    function getLabel(labelPath) {
        var pathParts = labelPath.split(".");
        var labelNode = LABELS;
        for (var i = 0; i < pathParts.length; i++) {
            if (labelNode == null) return labelPath;
            labelNode = labelNode[pathParts[i]];
        }
        if (labelNode == null) return labelPath;
        return labelNode[currentLanguage] || labelNode.en || labelPath;
    }

    // =========================================
    // 状態・UI配色 / State & UI colors
    // =========================================
    var registrationIndex = DEFAULTS.registrationIndex; /* 基準点 0..8（行優先, 4=中央）/ Registration point 0..8 (row-major, 4=center) */
    var activePreviewLayer = null; /* 生成したプレビュー用レイヤーの参照（名前でなく参照で管理）/ Reference to the preview layer we created */
    var cachedFrontmostItem;       /* 複製元の走査結果。undefined＝未計算, null＝該当なし / Cached duplication source (undefined = not resolved yet) */

    var lightUI = isLightUI();
    var widgetForeColor = lightUI ? [0.25, 0.25, 0.25, 1] : [0.85, 0.85, 0.85, 1]; /* セル・ケイ線 / Cells and rules */
    var widgetDimColor = lightUI ? [0.70, 0.70, 0.70, 1] : [0.42, 0.42, 0.42, 1];  /* 無効時 / Disabled */

    // =========================================
    // メイン / Main
    // =========================================
    if (app.documents.length === 0) {
        alert(getLabel("alert.noDocument"));
        return;
    }

    var doc = app.activeDocument;
    var selectedItems = doc.selection;

    if (!selectedItems || selectedItems.length === 0) {
        alert(getLabel("alert.noSelection"));
        return;
    }

    if (collectAllAnchorPoints(selectedItems).length === 0) {
        alert(getLabel("alert.noAnchorPoints"));
        return; /* パスが無い（テキスト・画像のみ等）/ No path anchors (e.g. text/image only) */
    }

    toggleCanvasHelpers(); /* 実行時：エッジ等を隠す / On run: hide edges & annotator */
    try {
        placeMarkers();
    } finally {
        // 途中で例外が出ても、プレビューとエッジ表示は必ず元に戻す
        // Always clean up the preview and restore edges, even if something throws
        removePreviewLayer();
        toggleCanvasHelpers();
    }

    /**
     * 設定ダイアログを表示し、確定した設定でマーカーを配置します。
     * キャンセル時・配置元を用意できなかったときは、何もせずに戻ります。
     * @returns {void}
     */
    function placeMarkers() {
        var userSettings = showSettingsDialog();
        if (!userSettings) return; /* キャンセル / Cancelled */

        var placeMarkerAtPoint = buildMarkerPlacer(userSettings, false);
        if (!placeMarkerAtPoint) return; /* 配置元が用意できなかった / No placement source available */

        var placementPoints = resolveAnchorPoints(userSettings.objectSource);
        var placedItems = [];
        for (var i = 0; i < placementPoints.length; i++) {
            placedItems.push(placeMarkerAtPoint(placementPoints[i][0], placementPoints[i][1]));
        }

        // グループ化：まとめてから（必要なら）レイヤーへ移す / Group first, then move to the layer if requested
        var resultItems = (userSettings.groupItems && placedItems.length > 0) ? [groupPlacedItems(placedItems)] : placedItems;

        if (userSettings.moveToLayer) {
            var targetLayer = getOrCreateLayer(ANCHOR_LAYER_NAME);
            for (var m = 0; m < resultItems.length; m++) {
                moveItemToLayer(resultItems[m], targetLayer);
            }
        }

        // 配置物を選択状態にして直後の移動・整列をしやすく / Select the placed items for easy move / align
        if (resultItems.length > 0) {
            doc.selection = resultItems;
        }
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================
    /**
     * マーカー配置の設定 / Marker placement settings
     * @typedef {Object} MarkerSettings
     * @property {string} objectSource - 追加するオブジェクトの種別（OBJECT_SOURCE の値）
     * @property {number} squareSize - 正方形の一辺（pt）
     * @property {Object} squareColor - 塗り色 {r, g, b}
     * @property {boolean} symbolize - 自動生成の正方形をシンボル化するか
     * @property {number} symbolIndex - 配置するシンボルのインデックス（未選択は -1）
     * @property {boolean} moveToLayer - 専用レイヤーへ移動するか
     * @property {boolean} groupItems - 1つのグループにまとめるか
     * @property {number} scalePercent - 拡大縮小率（%）
     * @property {number} registrationIndex - 基準点 0..8（行優先, 4=中央）
     */

    /**
     * 設定ダイアログを表示し、確定した設定を返します。
     * @returns {MarkerSettings|null} 確定した設定（キャンセル時は null）
     */
    function showSettingsDialog() {
        var pickedColor = {
            r: DEFAULTS.squareColor.r,
            g: DEFAULTS.squareColor.g,
            b: DEFAULTS.squareColor.b
        };

        var symbolNames = getSymbolNames();
        var hasSymbols = symbolNames.length > 0;

        var dialogWindow = setupWindow(new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION));

        // --- 追加するオブジェクト パネル / Object to Add panel ---
        var objectSourcePanel = addPanel(dialogWindow, getLabel("panel.objectSource"));

        var autoGenerateRadio = objectSourcePanel.add("radiobutton", undefined, getLabel("radio.autoGenerate"));
        autoGenerateRadio.helpTip = getLabel("tooltip.autoGenerate");
        var frontObjectRadio = objectSourcePanel.add("radiobutton", undefined, getLabel("radio.frontObject"));
        frontObjectRadio.helpTip = getLabel("tooltip.frontObject");

        var symbolSourceGroup = setupRow(objectSourcePanel.add("group"), ["left", "center"], 8);
        var symbolRadio = symbolSourceGroup.add("radiobutton", undefined, getLabel("radio.symbol"));
        symbolRadio.helpTip = getLabel("tooltip.symbol");
        var symbolDropdown = symbolSourceGroup.add("dropdownlist", undefined, symbolNames);
        symbolDropdown.helpTip = getLabel("tooltip.symbolDropdown");
        symbolDropdown.preferredSize.width = 130;
        if (hasSymbols) {
            symbolDropdown.selection = 0;
        }

        // 選択が1つ（単一オブジェクト／グループ1つ）だと複製元を除くと配置先が無いのでディム
        var canUseFrontObject = (selectedItems.length > 1);

        autoGenerateRadio.value = (DEFAULTS.objectSource === OBJECT_SOURCE.autoGenerate);
        frontObjectRadio.value = (DEFAULTS.objectSource === OBJECT_SOURCE.frontObject) && canUseFrontObject;
        symbolRadio.value = (DEFAULTS.objectSource === OBJECT_SOURCE.symbol) && hasSymbols;
        frontObjectRadio.enabled = canUseFrontObject;
        symbolRadio.enabled = hasSymbols;

        // --- アンカーポイント パネル（2カラム：左＝大きさ・シンボル化 / 右＝カラー）---
        var anchorPointPanel = addPanel(dialogWindow, getLabel("panel.anchorPoint"));
        setupRow(anchorPointPanel, ["left", "top"], COLUMN_SPACING);

        // 左カラム：大きさ・シンボル化 / Left column: size, symbolize
        var anchorLeft = addColumnGroup(anchorPointPanel);

        var sizeInput = addLabeledField(anchorLeft, getLabel("label.squareSize"), DEFAULTS.squareSize, 2, getLabel("label.unit"), renderPreview, undefined, true, 0.1);
        sizeInput.helpTip = getLabel("tooltip.squareSize"); /* ⌘併用で ±0.1 などのキー説明 / Key modifiers incl. ⌘ = ±0.1 */

        var symbolizeCheckbox = anchorLeft.add("checkbox", undefined, getLabel("checkbox.symbolize"));
        symbolizeCheckbox.value = DEFAULTS.symbolize;
        symbolizeCheckbox.helpTip = getLabel("tooltip.symbolize");

        // 右カラム：カラー（2行：カラー■ / 選択...）/ Right column: color (2 rows)
        // OSのカラーパレットはモーダルを壊すため、自前のRGBダイアログを使う
        var anchorRight = addColumnGroup(anchorPointPanel, 0);

        var colorRow = setupRow(anchorRight.add("group"), ["left", "center"], 8);
        colorRow.add("statictext", undefined, getLabel("label.squareColor"));
        var colorSwatch = colorRow.add("panel");
        colorSwatch.preferredSize = [SWATCH_SIZE, SWATCH_SIZE]; /* 正方形・高さは短いまま / Square, keep the short height */
        colorSwatch.helpTip = getLabel("tooltip.squareColor");
        colorSwatch.onDraw = makeSwatchDrawer(colorSwatch, pickedColor);

        // 選択ボタンは上マージンを持つグループで包む（コントロール直接の margins は効かない環境があるため）
        var chooseColorWrap = anchorRight.add("group");
        chooseColorWrap.margins = [0, 10, 0, 0]; /* 上にマージン10 / 10px top margin */
        var chooseColorButton = chooseColorWrap.add("button", undefined, getLabel("button.chooseColor"));
        chooseColorButton.helpTip = getLabel("tooltip.squareColor");
        chooseColorButton.onClick = function () {
            var chosenColor = chooseRgbColor(pickedColor);
            if (chosenColor) {
                pickedColor.r = chosenColor.r;
                pickedColor.g = chosenColor.g;
                pickedColor.b = chosenColor.b;
                redrawControl(colorSwatch);
                renderPreview();
            }
        };

        // --- オプション パネル（2カラム：左＝3設定 / 右＝9軸）/ Options panel (2 columns) ---
        var optionsPanel = addPanel(dialogWindow, getLabel("panel.options"));
        setupRow(optionsPanel, ["left", "top"], COLUMN_SPACING);

        // 左カラム：スケール・レイヤーに移動・グループ化 / Left column: scale, move-to-layer, group
        var optionsLeft = addColumnGroup(optionsPanel);

        var scaleInput = addLabeledField(optionsLeft, getLabel("label.scale"), DEFAULTS.scalePercent, 3, getLabel("label.percent"), renderPreview);
        scaleInput.helpTip = getLabel("tooltip.scale");

        var moveToLayerCheckbox = optionsLeft.add("checkbox", undefined, getLabel("checkbox.moveToLayer"));
        moveToLayerCheckbox.value = DEFAULTS.moveToLayer;
        moveToLayerCheckbox.helpTip = getLabel("tooltip.moveToLayer");

        var groupCheckbox = optionsLeft.add("checkbox", undefined, getLabel("checkbox.group"));
        groupCheckbox.value = DEFAULTS.groupItems;
        groupCheckbox.helpTip = getLabel("tooltip.group");

        // 右カラム：9軸（基準点）を天地左右中央に / Right column: 9-axis widget, centered both ways
        var optionsRight = addColumnGroup(optionsPanel);
        optionsRight.alignChildren = ["center", "center"];
        optionsRight.alignment = ["center", "center"]; /* 左カラムの高さに対して天地中央 / Vertically center against the left column */
        var registrationWidget = addRegistrationWidget(optionsRight, renderPreview);

        /**
         * 現在選択されている「追加するオブジェクト」の種別を返します。
         * @returns {string} OBJECT_SOURCE の値
         */
        function getChosenSource() {
            if (frontObjectRadio.value) return OBJECT_SOURCE.frontObject;
            if (symbolRadio.value) return OBJECT_SOURCE.symbol;
            return OBJECT_SOURCE.autoGenerate;
        }

        /**
         * 追加するオブジェクトの選択に応じて、各コントロールの有効／無効を切り替えます。
         * @returns {void}
         */
        function syncControlState() {
            var isAutoGenerate = autoGenerateRadio.value;
            sizeInput.enabled = isAutoGenerate;
            chooseColorButton.enabled = isAutoGenerate;
            symbolizeCheckbox.enabled = isAutoGenerate;
            // 自動生成では基準点を中央へ戻す / Reset registration to center in auto-generate
            if (isAutoGenerate) {
                registrationIndex = DEFAULTS.registrationIndex;
            }
            // 自動生成時はスケールと9軸のみディム（レイヤー移動・グループ化は常時有効）
            scaleInput.parent.enabled = !isAutoGenerate;
            registrationWidget.enabled = !isAutoGenerate;
            redrawControl(registrationWidget); /* 自作描画なので色を更新 / Redraw the custom widget */
            symbolDropdown.enabled = symbolRadio.value;
        }

        /**
         * ラジオが別コンテナに分かれているため、排他選択を手動で担保します。
         * @param {RadioButton} selectedRadio - 選択されたラジオボタン
         * @returns {void}
         */
        function selectObjectSource(selectedRadio) {
            autoGenerateRadio.value = (selectedRadio === autoGenerateRadio);
            frontObjectRadio.value = (selectedRadio === frontObjectRadio);
            symbolRadio.value = (selectedRadio === symbolRadio);
            syncControlState();
            renderPreview();
        }
        autoGenerateRadio.onClick = function () { selectObjectSource(autoGenerateRadio); };
        frontObjectRadio.onClick = function () { selectObjectSource(frontObjectRadio); };
        symbolRadio.onClick = function () { selectObjectSource(symbolRadio); };
        symbolizeCheckbox.onClick = renderPreview;
        symbolDropdown.onChange = renderPreview;
        syncControlState();

        /**
         * 現在のコントロール値から設定オブジェクトを読み取ります（プレビュー・本適用の共通ソース）。
         * @returns {MarkerSettings} 現在の設定
         */
        function readCurrentSettings() {
            var sizeValue = parseFloat(sizeInput.text);
            var scaleValue = parseFloat(scaleInput.text);
            return {
                objectSource: getChosenSource(),
                squareSize: (isNaN(sizeValue) || sizeValue <= 0) ? DEFAULTS.squareSize : sizeValue,
                squareColor: { r: pickedColor.r, g: pickedColor.g, b: pickedColor.b },
                symbolize: symbolizeCheckbox.value,
                symbolIndex: symbolDropdown.selection ? symbolDropdown.selection.index : -1,
                moveToLayer: moveToLayerCheckbox.value,
                groupItems: groupCheckbox.value,
                scalePercent: (isNaN(scaleValue) || scaleValue <= 0) ? DEFAULTS.scalePercent : scaleValue,
                registrationIndex: registrationIndex
            };
        }

        /**
         * プレビュー専用レイヤーに、現在の設定でマーカーを描画します。
         * @returns {void}
         */
        function renderPreview() {
            removePreviewLayer();

            var currentSettings = readCurrentSettings();
            var previewPlacer = buildMarkerPlacer(currentSettings, true);
            if (previewPlacer) {
                var previewPoints = resolveAnchorPoints(currentSettings.objectSource);
                activePreviewLayer = doc.layers.add();
                activePreviewLayer.name = PREVIEW_LAYER_NAME;
                for (var p = 0; p < previewPoints.length; p++) {
                    moveItemToLayer(previewPlacer(previewPoints[p][0], previewPoints[p][1]), activePreviewLayer);
                }
            }
            app.redraw();
        }

        // --- ボタン / Buttons（Mac 規約: Cancel → OK）---
        var dialogButtonGroup = setupRow(dialogWindow.add("group"), ["right", "center"], 8);
        dialogButtonGroup.alignment = ["right", "center"];
        var cancelButton = dialogButtonGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var okButton = dialogButtonGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });

        var dialogResult = null;
        okButton.onClick = function () {
            var chosenSource = getChosenSource();
            if (chosenSource === OBJECT_SOURCE.autoGenerate) {
                var sizeValue = parseFloat(sizeInput.text);
                if (isNaN(sizeValue) || sizeValue <= 0) {
                    alert(getLabel("alert.invalidSize"));
                    return;
                }
            } else {
                var scaleValue = parseFloat(scaleInput.text);
                if (isNaN(scaleValue) || scaleValue <= 0) {
                    alert(getLabel("alert.invalidScale"));
                    return;
                }
                if (chosenSource === OBJECT_SOURCE.symbol && !symbolDropdown.selection) {
                    alert(getLabel("alert.noSymbolChosen"));
                    return;
                }
            }

            removePreviewLayer(); /* プレビューを片付けてから本適用 / Clear preview before the real placement */
            dialogResult = readCurrentSettings();
            dialogWindow.close();
        };
        cancelButton.onClick = function () {
            removePreviewLayer();
            app.redraw();
            dialogResult = null;
            dialogWindow.close();
        };

        // 表示直後に初回プレビュー / Render the first preview once shown
        dialogWindow.onShow = function () {
            renderPreview();
        };

        // レイアウト確定後、選択ボタンを通常ボタンより -2 に詰める / After layout, trim the choose button by 2px
        dialogWindow.layout.layout(true);
        trimButtonHeight(chooseColorButton, 2);

        dialogWindow.center();
        dialogWindow.show();
        return dialogResult;
    }

    // =========================================
    // ダイアログ部品 / Dialog Helpers
    // =========================================
    /**
     * 「ラベル [入力] 単位」の1行を作り、入力欄を返します。
     * @param {Group|Panel} parentPanel - 追加先
     * @param {string} labelText - ラベル文字列
     * @param {number} initialValue - 初期値
     * @param {number} charCount - 入力欄の文字数（最小幅の指定）
     * @param {string} unitText - 単位表記
     * @param {function} onChange - 値が変わったときに呼ぶ関数
     * @param {number} [labelWidth] - ラベルの固定幅（px）
     * @param {boolean} [allowDecimal] - 小数第1位まで保持するか
     * @param {number} [minValue] - 下限値（省略時は 0）
     * @returns {EditText} 追加した入力欄
     */
    function addLabeledField(parentPanel, labelText, initialValue, charCount, unitText, onChange, labelWidth, allowDecimal, minValue) {
        var fieldRow = setupRow(parentPanel.add("group"), ["left", "center"], 8);
        var fieldLabel = fieldRow.add("statictext", undefined, labelText);
        if (labelWidth) {
            fieldLabel.preferredSize.width = labelWidth; /* 指定時のみ固定幅 / Fixed width only when given */
        }
        var fieldInput = fieldRow.add("edittext", undefined, String(initialValue));
        fieldInput.characters = charCount;
        fieldInput.onChanging = onChange;
        changeValueByArrowKey(fieldInput, onChange, allowDecimal, minValue);
        fieldRow.add("statictext", undefined, unitText);
        return fieldInput;
    }

    /**
     * スウォッチ（panel）を現在の色で塗る onDraw ハンドラを生成します。
     * @param {Panel} swatch - 描画対象のパネル
     * @param {Object} color - 塗り色 {r, g, b}
     * @returns {function} onDraw ハンドラ
     */
    function makeSwatchDrawer(swatch, color) {
        return function () {
            var swatchGraphics = swatch.graphics;
            var fillBrush = swatchGraphics.newBrush(
                swatchGraphics.BrushType.SOLID_COLOR,
                [color.r / 255, color.g / 255, color.b / 255, 1]
            );
            swatchGraphics.newPath();
            swatchGraphics.rectPath(0, 0, swatch.size[0], swatch.size[1]);
            swatchGraphics.fillPath(fillBrush);
        };
    }

    /**
     * 自前の RGB カラーダイアログを表示します（OS のカラーパレットはモーダルを壊すため）。
     * @param {Object} startColor - 初期色 {r, g, b}
     * @returns {Object|null} 選択した色 {r, g, b}（キャンセル時は null）
     */
    function chooseRgbColor(startColor) {
        var workingColor = { r: startColor.r, g: startColor.g, b: startColor.b };
        var confirmed = false;

        var pickerWindow = new Window("dialog", getLabel("colorPicker.title"));
        setupRow(pickerWindow, ["fill", "fill"], WINDOW_SPACING);
        pickerWindow.margins = WINDOW_MARGINS;

        var previewSwatch = pickerWindow.add("panel");
        previewSwatch.preferredSize = [64, 64];
        previewSwatch.onDraw = makeSwatchDrawer(previewSwatch, workingColor);

        var fieldsColumn = addColumnGroup(pickerWindow, 6);

        /**
         * 3つの数値欄から作業色を読み直し、プレビューを再描画します。
         * @returns {void}
         */
        function refreshFromFields() {
            workingColor.r = clampColorChannel(Number(redInput.text));
            workingColor.g = clampColorChannel(Number(greenInput.text));
            workingColor.b = clampColorChannel(Number(blueInput.text));
            redrawControl(previewSwatch);
        }
        var redInput = addColorChannelField(fieldsColumn, "R", workingColor.r, refreshFromFields);
        var greenInput = addColorChannelField(fieldsColumn, "G", workingColor.g, refreshFromFields);
        var blueInput = addColorChannelField(fieldsColumn, "B", workingColor.b, refreshFromFields);

        var pickerButtonGroup = setupRow(fieldsColumn.add("group"), ["right", "center"], 8);
        pickerButtonGroup.alignment = ["right", "center"];
        var pickerCancelButton = pickerButtonGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var pickerOkButton = pickerButtonGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });

        // show() の戻り値に依存せず、明示的な onClick で確定する（環境差で OK が 1 を返さない対策）
        pickerOkButton.onClick = function () {
            refreshFromFields();
            confirmed = true;
            pickerWindow.close();
        };
        pickerCancelButton.onClick = function () {
            confirmed = false;
            pickerWindow.close();
        };

        pickerWindow.center();
        pickerWindow.show();
        return confirmed ? { r: workingColor.r, g: workingColor.g, b: workingColor.b } : null;
    }

    /**
     * RGB チャンネル1つ（ラベル＋スライダー＋数値入力を相互同期）を作り、入力欄を返します。
     * @param {Group} parentGroup - 追加先
     * @param {string} channelLabel - チャンネル名（"R" / "G" / "B"）
     * @param {number} initialValue - 初期値（0〜255）
     * @param {function} onChange - 値が変わったときに呼ぶ関数
     * @returns {EditText} 追加した入力欄
     */
    function addColorChannelField(parentGroup, channelLabel, initialValue, onChange) {
        var channelRow = setupRow(parentGroup.add("group"), ["left", "center"], 6);
        var channelLabelText = channelRow.add("statictext", undefined, channelLabel);
        channelLabelText.preferredSize.width = 14;
        var channelSlider = channelRow.add("slider", undefined, initialValue, 0, 255);
        channelSlider.preferredSize = [140, 18];
        var channelInput = channelRow.add("edittext", undefined, String(initialValue));
        channelInput.characters = 4;

        /**
         * スライダーの値を数値欄へ反映します。
         * @returns {void}
         */
        function syncFromSlider() {
            channelInput.text = Math.round(channelSlider.value);
            onChange();
        }

        /**
         * 数値欄の値をスライダーへ反映します。
         * @returns {void}
         */
        function syncFromInput() {
            channelSlider.value = clampColorChannel(Number(channelInput.text));
            onChange();
        }
        channelSlider.onChanging = syncFromSlider; /* ドラッグ中（発火する環境）/ During drag where supported */
        channelSlider.onChange = syncFromSlider;   /* ドラッグ後（リリース）で確実に反映 / On release, reliably */
        channelInput.onChanging = syncFromInput;
        changeValueByArrowKey(channelInput, syncFromInput);
        return channelInput;
    }

    /**
     * 数値を 0〜255 に丸めてクランプします。
     * @param {number} value - 入力値
     * @returns {number} 0〜255 の整数
     */
    function clampColorChannel(value) {
        if (isNaN(value)) return 0;
        value = Math.round(value);
        if (value < 0) return 0;
        if (value > 255) return 255;
        return value;
    }

    /**
     * 生成したプレビューレイヤーだけを参照で削除します（同名の既存レイヤーは触らない）。
     * @returns {void}
     */
    function removePreviewLayer() {
        if (activePreviewLayer) {
            try {
                activePreviewLayer.remove();
            } catch (e) {
                /* 既に無ければ無視 / Ignore if already gone */
            }
            activePreviewLayer = null;
        }
    }

    /**
     * 指定名のレイヤーを取得し、無ければ作成します。
     * @param {string} layerName - レイヤー名
     * @returns {Layer} 取得または作成したレイヤー
     */
    function getOrCreateLayer(layerName) {
        try {
            return doc.layers.getByName(layerName);
        } catch (e) {
            var layer = doc.layers.add();
            layer.name = layerName;
            return layer;
        }
    }

    /**
     * エッジ表示とライブコーナー注釈をトグルします（実行中は隠す）。
     * @returns {void}
     */
    function toggleCanvasHelpers() {
        app.executeMenuCommand('edge');
        app.executeMenuCommand('Live Corner Annotator');
    }

    /**
     * ボタンの高さを指定 px 詰めます（レイアウト確定後に呼ぶ）。
     * @param {Button} button - 対象のボタン
     * @param {number} px - 詰める量（px）
     * @returns {void}
     */
    function trimButtonHeight(button, px) {
        try {
            button.size = [button.size.width, button.size.height - px];
        } catch (e) {}
    }

    /**
     * 生成物を指定レイヤーの最前面へ移動します。
     * @param {PageItem} item - 移動するオブジェクト
     * @param {Layer|GroupItem} layer - 移動先
     * @returns {void}
     */
    function moveItemToLayer(item, layer) {
        item.move(layer, ElementPlacement.PLACEATBEGINNING);
    }

    /**
     * コントロールを再描画します（notify は環境により例外を投げ得るので保護）。
     * @param {Object} control - 再描画するコントロール
     * @returns {void}
     */
    function redrawControl(control) {
        try {
            control.notify("onDraw");
        } catch (e) {}
    }

    /**
     * 配置済みアイテムを1つのグループにまとめます。
     * @param {Array<PageItem>} items - まとめる対象
     * @returns {GroupItem} 作成したグループ
     */
    function groupPlacedItems(items) {
        var markerGroup = doc.groupItems.add();
        for (var i = 0; i < items.length; i++) {
            items[i].move(markerGroup, ElementPlacement.PLACEATEND);
        }
        return markerGroup;
    }

    /**
     * ↑↓キーで数値を増減できるようにします（Shift＝±10・10スナップ／⌘＝±0.1／通常＝±1）。
     * @param {EditText} editText - 対象の入力欄
     * @param {function} onValueChange - 値が変わったときに呼ぶ関数
     * @param {boolean} [allowDecimal] - true で小数第1位まで保持、false で整数に丸め
     * @param {number} [minValue] - 下限値（省略時は 0）
     * @returns {void}
     */
    function changeValueByArrowKey(editText, onValueChange, allowDecimal, minValue) {
        var lowerBound = (typeof minValue === "number") ? minValue : 0;
        editText.addEventListener("keydown", function (event) {
            if (event.keyName !== "Up" && event.keyName !== "Down") return;
            var value = Number(editText.text);
            if (isNaN(value)) return;

            // 修飾キーはイベントから読む（macOS では keyboardState が false のことがある）
            // Read modifiers from the event (keyboardState can be false on macOS)
            var withShift = readModifier(event, "shiftKey");
            var withCommand = readModifier(event, "metaKey"); /* ⌘ = metaKey */
            var direction = (event.keyName === "Up") ? 1 : -1;

            if (withShift) {
                value = (direction > 0) ? Math.ceil((value + 1) / 10) * 10 : Math.floor((value - 1) / 10) * 10;
            } else if (withCommand) {
                value += direction * 0.1;
            } else {
                value += direction;
            }

            if (allowDecimal || withCommand) {
                value = Math.round(value * 10) / 10; /* 小数第1位まで / Round to 1 decimal */
            } else {
                value = Math.round(value); /* 整数に丸め / Round to integer */
            }
            if (value < lowerBound) value = lowerBound; /* 下限でクランプ / Clamp to the lower bound */

            editText.text = value;
            event.preventDefault();
            if (typeof onValueChange === "function") {
                onValueChange();
            }
        });
    }

    /**
     * キーイベントの修飾キー状態を返します（event を優先し、keyboardState はフォールバック）。
     * @param {Object} event - キーイベント
     * @param {string} name - 修飾キー名（"shiftKey" / "metaKey" など）
     * @returns {boolean} 押されていれば true
     */
    function readModifier(event, name) {
        if (event[name] === true) return true;
        try {
            return ScriptUI.environment.keyboardState[name] === true;
        } catch (e) {
            return false;
        }
    }

    // =========================================
    // アンカーポイント収集 / Anchor Point Collection
    // =========================================
    /**
     * 選択オブジェクト群からアンカー座標を収集します。
     * @param {Array<PageItem>} items - 走査対象
     * @param {PageItem} [excludeItem] - 除外するオブジェクト
     * @returns {Array<Array<number>>} アンカー座標 [x, y] の配列
     */
    function collectAllAnchorPoints(items, excludeItem) {
        var collectedPoints = [];
        for (var i = 0; i < items.length; i++) {
            collectAnchorPoints(items[i], collectedPoints, excludeItem);
        }
        return collectedPoints;
    }

    /**
     * オブジェクトの種別に応じてアンカー座標を集めます（excludeItem は対象外）。
     * @param {PageItem} item - 対象オブジェクト
     * @param {Array<Array<number>>} collectedPoints - 収集先の配列
     * @param {PageItem} [excludeItem] - 除外するオブジェクト
     * @returns {void}
     */
    function collectAnchorPoints(item, collectedPoints, excludeItem) {
        if (excludeItem && item === excludeItem) {
            return; /* このオブジェクト自身のアンカーは対象外 / Skip this object's own anchors */
        }
        if (item.typename === "PathItem") {
            for (var i = 0; i < item.pathPoints.length; i++) {
                collectedPoints.push(item.pathPoints[i].anchor);
            }
        } else if (item.typename === "CompoundPathItem") {
            for (var j = 0; j < item.pathItems.length; j++) {
                collectAnchorPoints(item.pathItems[j], collectedPoints, excludeItem);
            }
        } else if (item.typename === "GroupItem") {
            for (var k = 0; k < item.pageItems.length; k++) {
                collectAnchorPoints(item.pageItems[k], collectedPoints, excludeItem);
            }
        }
    }

    /**
     * 配置に使うアンカー座標を返します。
     * 「最前面のオブジェクト」モードでは、複製元（最前面オブジェクト）自身のアンカーを除外します。
     * @param {string} objectSource - 追加するオブジェクトの種別（OBJECT_SOURCE の値）
     * @returns {Array<Array<number>>} アンカー座標 [x, y] の配列
     */
    function resolveAnchorPoints(objectSource) {
        var excludeItem = (objectSource === OBJECT_SOURCE.frontObject) ? getFrontmostItem() : null;
        return collectAllAnchorPoints(selectedItems, excludeItem);
    }

    // =========================================
    // 配置処理 / Placement
    // =========================================
    /**
     * ユーザー設定から、アンカー座標に配置する処理（placer 関数）を組み立てます。
     * placer は生成した pageItem を返します（プレビュー時にレイヤーへ移すため）。
     * forPreview=true のときは自動生成の「シンボル化」を無視し、見た目が同じ正方形パスを描きます
     * （プレビューのたびに新規シンボルを登録してシンボルパネルを汚さないため）。
     * @param {MarkerSettings} settings - 現在の設定
     * @param {boolean} forPreview - プレビュー用なら true
     * @returns {function|null} (anchorX, anchorY) を受け取り PageItem を返す関数（用意できなければ null）
     */
    function buildMarkerPlacer(settings, forPreview) {
        var fractionX = (settings.registrationIndex % 3) / 2;         /* 0=左, 0.5=中央, 1=右 / 0=left, .5=center, 1=right */
        var fractionY = Math.floor(settings.registrationIndex / 3) / 2; /* 0=上, 0.5=中央, 1=下 / 0=top, .5=middle, 1=bottom */

        if (settings.objectSource === OBJECT_SOURCE.frontObject) {
            var frontmostItem = getFrontmostItem();
            if (!frontmostItem) return null;
            return function (anchorX, anchorY) {
                return duplicateItemAtPoint(frontmostItem, anchorX, anchorY, settings.scalePercent, fractionX, fractionY);
            };
        }

        if (settings.objectSource === OBJECT_SOURCE.symbol) {
            if (settings.symbolIndex < 0) return null;
            var chosenSymbol = doc.symbols[settings.symbolIndex];
            return function (anchorX, anchorY) {
                return placeSymbolInstance(chosenSymbol, anchorX, anchorY, settings.scalePercent, fractionX, fractionY);
            };
        }

        // OBJECT_SOURCE.autoGenerate（大きさは pt 指定なのでスケールは 100 固定）
        var squareColor = toRgbColor(settings.squareColor);
        if (settings.symbolize && !forPreview) {
            var squareSymbol = createSquareSymbol(settings.squareSize, squareColor);
            return function (anchorX, anchorY) {
                return placeSymbolInstance(squareSymbol, anchorX, anchorY, 100, fractionX, fractionY);
            };
        }
        return function (anchorX, anchorY) {
            return placeSquareRect(settings.squareSize, squareColor, anchorX, anchorY, fractionX, fractionY);
        };
    }

    /**
     * 基準点の割合（0〜1）に合わせて生成物を配置します。
     * @param {PageItem} item - 配置するオブジェクト
     * @param {number} anchorX - アンカーのX座標
     * @param {number} anchorY - アンカーのY座標
     * @param {number} fractionX - 横方向の基準位置（0=左, 0.5=中央, 1=右）
     * @param {number} fractionY - 縦方向の基準位置（0=上, 0.5=中央, 1=下）
     * @returns {void}
     */
    function positionItemAtAnchor(item, anchorX, anchorY, fractionX, fractionY) {
        item.left = anchorX - fractionX * item.width;
        item.top = anchorY + fractionY * item.height;
    }

    /**
     * {r, g, b} を RGBColor へ変換します。
     * @param {Object} rgb - 色 {r, g, b}
     * @returns {RGBColor} 変換した色
     */
    function toRgbColor(rgb) {
        var color = new RGBColor();
        color.red = rgb.r;
        color.green = rgb.g;
        color.blue = rgb.b;
        return color;
    }

    /**
     * 現在の大きさ・カラーで正方形シンボルを毎回新規登録して返します（既存は再利用しない）。
     * プレビュー（正方形パス）と本適用（同設定のシンボルインスタンス）の見た目を一致させるためです。
     * @param {number} squareSize - 正方形の一辺（pt）
     * @param {RGBColor} squareColor - 塗り色
     * @returns {Symbol} 登録したシンボル
     */
    function createSquareSymbol(squareSize, squareColor) {
        var masterSquare = doc.pathItems.rectangle(squareSize / 2, -squareSize / 2, squareSize, squareSize);
        masterSquare.filled = true;
        masterSquare.fillColor = squareColor;
        masterSquare.stroked = false;

        var symbolDefinition = doc.symbols.add(masterSquare);
        try {
            symbolDefinition.name = SQUARE_SYMBOL_NAME;
        } catch (e) {
            /* 同名シンボルが既にある場合は既定名のまま（重複を許容）/ Keep the default name on collision */
        }
        masterSquare.remove();
        return symbolDefinition;
    }

    /**
     * 基準点に合わせてシンボルインスタンスを配置します。
     * @param {Symbol} symbolDefinition - 配置するシンボル
     * @param {number} anchorX - アンカーのX座標
     * @param {number} anchorY - アンカーのY座標
     * @param {number} scalePercent - 拡大縮小率（%）
     * @param {number} fractionX - 横方向の基準位置
     * @param {number} fractionY - 縦方向の基準位置
     * @returns {SymbolItem} 配置したインスタンス
     */
    function placeSymbolInstance(symbolDefinition, anchorX, anchorY, scalePercent, fractionX, fractionY) {
        var symbolInstance = doc.symbolItems.add(symbolDefinition);
        if (scalePercent !== 100) {
            symbolInstance.resize(scalePercent, scalePercent);
        }
        positionItemAtAnchor(symbolInstance, anchorX, anchorY, fractionX, fractionY);
        return symbolInstance;
    }

    /**
     * 基準点に合わせて最前面オブジェクトを複製します。
     * @param {PageItem} sourceItem - 複製元
     * @param {number} anchorX - アンカーのX座標
     * @param {number} anchorY - アンカーのY座標
     * @param {number} scalePercent - 拡大縮小率（%）
     * @param {number} fractionX - 横方向の基準位置
     * @param {number} fractionY - 縦方向の基準位置
     * @returns {PageItem} 複製したオブジェクト
     */
    function duplicateItemAtPoint(sourceItem, anchorX, anchorY, scalePercent, fractionX, fractionY) {
        var duplicatedItem = sourceItem.duplicate();
        if (scalePercent !== 100) {
            duplicatedItem.resize(scalePercent, scalePercent);
        }
        positionItemAtAnchor(duplicatedItem, anchorX, anchorY, fractionX, fractionY);
        return duplicatedItem;
    }

    /**
     * 基準点に合わせて正方形パスを配置します。
     * @param {number} squareSize - 正方形の一辺（pt）
     * @param {RGBColor} squareColor - 塗り色
     * @param {number} anchorX - アンカーのX座標
     * @param {number} anchorY - アンカーのY座標
     * @param {number} fractionX - 横方向の基準位置
     * @param {number} fractionY - 縦方向の基準位置
     * @returns {PathItem} 配置した正方形
     */
    function placeSquareRect(squareSize, squareColor, anchorX, anchorY, fractionX, fractionY) {
        var squareRect = doc.pathItems.rectangle(0, 0, squareSize, squareSize);
        squareRect.filled = true;
        squareRect.fillColor = squareColor;
        squareRect.stroked = false;
        positionItemAtAnchor(squareRect, anchorX, anchorY, fractionX, fractionY);
        return squareRect;
    }

    /**
     * 「選択範囲内で最前面のオブジェクト」を返します（未選択の最前面は対象外）。
     * 複製元はスクリプト実行中ずっと同じもの（選択は起動時に確定し、モーダル表示中は変更できず、
     * プレビューの複製物は除外レイヤーへ逃がしている）なので、初回の走査結果を使い回します。
     * @returns {PageItem|null} 選択範囲内の最前面オブジェクト（無ければ null）
     */
    function getFrontmostItem() {
        if (cachedFrontmostItem === undefined) {
            cachedFrontmostItem = findFrontmostSelectedItem();
        }
        return cachedFrontmostItem;
    }

    /**
     * ドキュメントを前面から走査して、最初に selected なアイテムを返します。
     * 選択されているものだけを対象にすることで、未選択オブジェクトを複製元に選んでしまう事故を防ぎます。
     * @returns {PageItem|null} 最前面の選択アイテム（無ければ null）
     */
    function findFrontmostSelectedItem() {
        for (var i = 0; i < doc.layers.length; i++) {
            var layer = doc.layers[i];
            if (!layer.visible || layer.locked || layer.name === PREVIEW_LAYER_NAME) {
                continue;
            }
            var foundItem = frontmostSelectedInContainer(layer);
            if (foundItem) return foundItem;
        }
        return null;
    }

    /**
     * コンテナ（レイヤー／グループ）を前面から走査し、最初に選択されているアイテムを返します。
     * @param {Layer|GroupItem} container - 走査するコンテナ
     * @returns {PageItem|null} 最初に見つかった選択アイテム（無ければ null）
     */
    function frontmostSelectedInContainer(container) {
        var childItems = container.pageItems;
        for (var i = 0; i < childItems.length; i++) {
            var item = childItems[i];
            if (item.selected) {
                return item; /* この選択アイテムが最前面 / This selected item is frontmost */
            }
            if (item.typename === "GroupItem") {
                var nestedItem = frontmostSelectedInContainer(item);
                if (nestedItem) return nestedItem;
            }
        }
        return null;
    }

    /**
     * ドキュメント内のシンボル名の一覧を返します。
     * @returns {Array<string>} シンボル名の配列
     */
    function getSymbolNames() {
        var names = [];
        for (var i = 0; i < doc.symbols.length; i++) {
            names.push(doc.symbols[i].name);
        }
        return names;
    }

    // =========================================
    // 基準点ウィジェット（9軸）/ Registration widget (9-axis)
    // =========================================
    /**
     * 3×3 の基準点ウィジェットを生成します。クリックで registrationIndex を更新し onChange を呼びます。
     * @param {Group} parentGroup - 追加先
     * @param {function} onChange - 基準点が変わったときに呼ぶ関数
     * @returns {Button} 追加したウィジェット
     */
    function addRegistrationWidget(parentGroup, onChange) {
        var widget = parentGroup.add("button", undefined, "");
        widget.helpTip = getLabel("tooltip.registration");
        widget.preferredSize = [WIDGET_SIZE, WIDGET_SIZE];
        widget.minimumSize = [WIDGET_SIZE, WIDGET_SIZE];
        widget.maximumSize = [WIDGET_SIZE, WIDGET_SIZE];
        widget.onDraw = function () {
            drawRegistrationWidget(this);
        };
        widget.addEventListener("mousedown", function (event) {
            var columnIndex = clampCell(Math.floor(event.clientX / (widget.size[0] / 3)));
            var rowIndex = clampCell(Math.floor(event.clientY / (widget.size[1] / 3)));
            registrationIndex = rowIndex * 3 + columnIndex;
            redrawControl(widget);
            if (typeof onChange === "function") {
                onChange();
            }
        });
        return widget;
    }

    /**
     * セル位置を 0〜2 にクランプします。
     * @param {number} value - 入力値
     * @returns {number} 0〜2 の値
     */
    function clampCell(value) {
        if (value < 0) return 0;
        if (value > 2) return 2;
        return value;
    }

    /**
     * 9軸ウィジェットを描画します（外周の□をケイ線でつなぎ、中央は独立）。
     * @param {Button} widget - 描画対象のウィジェット
     * @returns {void}
     */
    function drawRegistrationWidget(widget) {
        var graphics = widget.graphics;
        var width = widget.size[0];
        var height = widget.size[1];
        var foreColor = widget.enabled ? widgetForeColor : widgetDimColor; /* 無効時はディム / Dim when disabled */

        // 背景は塗らない（透過・親ウィンドウ色）/ No background fill (transparent, parent window color)
        var cellSize = 8;   /* 四角のサイズ / Square size */
        var cellGap = 6;    /* 四角どうしの間隔 / Gap between squares */
        var cellStep = cellSize + cellGap;
        var gridSize = cellSize * 3 + cellGap * 2;
        var originX = Math.round((width - gridSize) / 2);
        var originY = Math.round((height - gridSize) / 2);

        /**
         * セルの左上X座標を返します。
         * @param {number} index - セル番号 0..8
         * @returns {number} X座標
         */
        function cellX(index) { return originX + (index % 3) * cellStep; }

        /**
         * セルの左上Y座標を返します。
         * @param {number} index - セル番号 0..8
         * @returns {number} Y座標
         */
        function cellY(index) { return originY + Math.floor(index / 3) * cellStep; }

        // 中央(4)を除く外周の□どうしをケイ線でつなぐ / Join the outer squares (skipping center) with rules
        var connections = [[0, 1], [1, 2], [6, 7], [7, 8], [0, 3], [3, 6], [2, 5], [5, 8]];
        var linePen = graphics.newPen(graphics.PenType.SOLID_COLOR, foreColor, 1);
        for (var i = 0; i < connections.length; i++) {
            var startCell = connections[i][0];
            var endCell = connections[i][1];
            graphics.newPath();
            if (endCell - startCell === 1) {
                graphics.moveTo(cellX(startCell) + cellSize, cellY(startCell) + cellSize / 2);
                graphics.lineTo(cellX(endCell), cellY(endCell) + cellSize / 2);
            } else {
                graphics.moveTo(cellX(startCell) + cellSize / 2, cellY(startCell) + cellSize);
                graphics.lineTo(cellX(endCell) + cellSize / 2, cellY(endCell));
            }
            graphics.strokePath(linePen);
        }

        for (var index = 0; index < 9; index++) {
            drawRegistrationCell(graphics, cellX(index), cellY(index), cellSize, index === registrationIndex, foreColor);
        }
    }

    /**
     * 基準点セルの□を1つ描画します。
     * 選択中は「塗り＋ケイ線」で描き、非選択の枠線と外周サイズをそろえます。
     * @param {Object} graphics - ScriptUI の graphics オブジェクト
     * @param {number} x - セルの左上X座標
     * @param {number} y - セルの左上Y座標
     * @param {number} size - セルの一辺（px）
     * @param {boolean} selected - 選択中なら true
     * @param {Array<number>} foreColor - 描画色 [r, g, b, a]
     * @returns {void}
     */
    function drawRegistrationCell(graphics, x, y, size, selected, foreColor) {
        if (selected) {
            buildCellPath(graphics, x, y, size);
            graphics.fillPath(graphics.newBrush(graphics.BrushType.SOLID_COLOR, foreColor));
        }
        buildCellPath(graphics, x, y, size);
        graphics.strokePath(graphics.newPen(graphics.PenType.SOLID_COLOR, foreColor, 1));
    }

    /**
     * セルの正方形パスを組み立てます。
     * @param {Object} graphics - ScriptUI の graphics オブジェクト
     * @param {number} x - 左上X座標
     * @param {number} y - 左上Y座標
     * @param {number} size - 一辺（px）
     * @returns {void}
     */
    function buildCellPath(graphics, x, y, size) {
        graphics.newPath();
        graphics.moveTo(x, y);
        graphics.lineTo(x + size, y);
        graphics.lineTo(x + size, y + size);
        graphics.lineTo(x, y + size);
        graphics.closePath();
    }

    // =========================================
    // UI の明暗判定 / Light vs. dark UI
    // =========================================
    /**
     * UI 明度からライトテーマかどうかを判定します（取得に失敗したらダーク扱い）。
     * @returns {boolean} ライトテーマなら true
     */
    function isLightUI() {
        try {
            return app.preferences.getRealPreference("uiBrightness") > 0.5;
        } catch (e) {
            return false;
        }
    }
})();
