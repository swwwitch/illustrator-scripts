#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択オブジェクトの全アンカーポイントに、マーカー（正方形／最前面オブジェクト／シンボル）を配置するユーティリティです。ダイアログを閉じずにライブプレビューしながら設定できます。

- 追加するオブジェクトを選択：アンカーポイントを自動生成（正方形）／最前面のオブジェクトを複製／ドキュメント内シンボルのインスタンス
- 正方形は「大きさ（pt・小数可、⌘＋↑↓で±0.1）」「カラー（自前のRGBダイアログ）」「シンボル化（既定ON・シンボル名『アンカーポイント』）」を指定
- オプション：スケール（%、最前面オブジェクト・シンボルに適用）／レイヤーに移動（_anchorpoint）／グループ化（既定ON）／9軸の基準点（マーカーをアンカーに合わせる位置）
- 自動生成時はスケール・9軸をディム表示し、基準点は中央に固定
- ライブプレビューは専用レイヤーに描画し、OK／キャンセルで確実に片付け
- 実行中はエッジ表示とライブコーナー注釈を一時的に隠す（開始・終了でトグル）
- 日英ローカライズ、ライト／ダークUIに追従

### Overview

Places a marker (auto-generated square / frontmost object / symbol) at every anchor point of the selection, with a live preview you can tweak without closing the dialog.

- Choose what to add: an auto-generated square, a duplicate of the frontmost object, or an instance of a document symbol
- The square takes size (pt, decimals allowed, ⌘+↑↓ = ±0.1), color (a self-contained RGB dialog), and Symbolize (on by default; symbol named "アンカーポイント")
- Options: Scale (%, applied to the frontmost object / symbol), Move to layer (_anchorpoint), Group (on by default), and a 9-axis registration point
- In auto-generate mode, Scale and the 9-axis widget are dimmed and the registration point is fixed to center
- The live preview draws into a dedicated layer and is always cleaned up on OK / Cancel
- Edges and the Live Corner Annotator are hidden during the run (toggled on start and finish)
- Japanese / English localization; adapts to the light / dark UI

### 更新履歴 / Change Log

- v1.0.0: 初期バージョン。マーカー配置（正方形／最前面／シンボル）、基準点（9軸）、スケール、レイヤー移動、グループ化、ライブプレビュー。

*/

// =========================================
// バージョン / Version
// =========================================
var SCRIPT_VERSION = "v1.0.0";

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================
    var OBJECT_SOURCE = {
        autoGenerate: "autoGenerate", /* 正方形を自動生成 / Auto-generated square */
        frontObject: "frontObject",   /* 最前面のオブジェクト / Frontmost object */
        symbol: "symbol"              /* シンボル / Symbol */
    };

    var PREVIEW_LAYER_NAME = "__ANCHOR_MARKER_PREVIEW__"; /* プレビュー専用レイヤー名 / Preview-only layer name */
    var ANCHOR_LAYER_NAME = "_anchorpoint";               /* マーカー移動先レイヤー名 / Destination layer for markers */
    var SQUARE_SYMBOL_NAME = "アンカーポイント";           /* 自動生成シンボルの名前 / Name for the generated symbol */

    var DEFAULTS = {
        objectSource: OBJECT_SOURCE.autoGenerate, /* 追加するオブジェクトの種類 / Kind of object to add */
        squareSize: 6,                            /* 正方形の一辺（pt） / Square edge size (pt) */
        squareColor: { r: 0, g: 149, b: 212 },    /* 塗り色（RGB） / Fill color (RGB) */
        symbolize: true,                          /* 自動生成の正方形をシンボル化して配置 / Place squares as symbol instances */
        moveToLayer: false,                       /* 配置後のマーカーを専用レイヤーへ移動 / Move markers to a dedicated layer */
        groupItems: true,                         /* 配置後のマーカーを1つのグループにまとめる / Group the placed markers */
        scalePercent: 100,                        /* 最前面オブジェクト／シンボルの拡大縮小率（%） / Scale for frontmost object / symbol (%) */
        registrationIndex: 4                      /* 基準点 0..8（行優先, 4=中央）/ Registration point 0..8 row-major (4=center) */
    };

    // =========================================
    // ローカライズ / Localization
    // =========================================
    function getCurrentLang() {
        return ($.locale && $.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var currentLanguage = getCurrentLang();

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

    /* ドット区切りキーで LABELS を辿る。キー漏れ・言語漏れ時はキー自身/英語にフォールバック */
    function L(labelPath) {
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

    var lightUI = isLightUI();
    var widgetForeColor = lightUI ? [0.25, 0.25, 0.25, 1] : [0.85, 0.85, 0.85, 1]; /* セル・ケイ線 / Cells and rules */
    var widgetDimColor = lightUI ? [0.70, 0.70, 0.70, 1] : [0.42, 0.42, 0.42, 1];  /* 無効時 / Disabled */

    // =========================================
    // メイン / Main
    // =========================================
    if (app.documents.length === 0) {
        alert(L("alert.noDocument"));
        return;
    }

    var doc = app.activeDocument;
    var selectedItems = doc.selection;

    if (!selectedItems || selectedItems.length === 0) {
        alert(L("alert.noSelection"));
        return;
    }

    var anchorPoints = collectAllAnchorPoints(selectedItems);
    if (anchorPoints.length === 0) {
        alert(L("alert.noAnchorPoints"));
        return; /* パスが無い（テキスト・画像のみ等）/ No path anchors (e.g. text/image only) */
    }

    toggleCanvasHelpers(); /* 実行時：エッジ等を隠す / On run: hide edges & annotator */

    var userSettings = showSettingsDialog();
    if (!userSettings) {
        toggleCanvasHelpers(); /* 終了時：エッジを元に戻す / On exit: restore edges */
        return; /* キャンセル / Cancelled */
    }

    var placeMarkerAtPoint = buildMarkerPlacer(userSettings, false);
    if (!placeMarkerAtPoint) {
        toggleCanvasHelpers();
        return; /* 配置元が用意できなかった / No placement source available */
    }

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

    toggleCanvasHelpers(); /* 終了時：エッジ等を元に戻す / On finish: restore edges & annotator */

    // =========================================
    // ダイアログ / Dialog
    // =========================================
    /* 設定ダイアログを表示し、ユーザー設定オブジェクトを返す（キャンセル時は null） */
    function showSettingsDialog() {
        var pickedColor = {
            r: DEFAULTS.squareColor.r,
            g: DEFAULTS.squareColor.g,
            b: DEFAULTS.squareColor.b
        };

        var symbolNames = getSymbolNames();
        var hasSymbols = symbolNames.length > 0;

        var dialogWindow = new Window("dialog", L("dialog.title") + " " + SCRIPT_VERSION);
        dialogWindow.orientation = "column";
        dialogWindow.alignChildren = ["fill", "top"];
        dialogWindow.margins = 16;
        dialogWindow.spacing = 12;

        // --- 追加するオブジェクト パネル / Object to Add panel ---
        var objectSourcePanel = dialogWindow.add("panel", undefined, L("panel.objectSource"));
        objectSourcePanel.orientation = "column";
        objectSourcePanel.alignChildren = ["left", "center"];
        objectSourcePanel.margins = [16, 20, 16, 16];
        objectSourcePanel.spacing = 10;

        var autoGenerateRadio = objectSourcePanel.add("radiobutton", undefined, L("radio.autoGenerate"));
        autoGenerateRadio.helpTip = L("tooltip.autoGenerate");
        var frontObjectRadio = objectSourcePanel.add("radiobutton", undefined, L("radio.frontObject"));
        frontObjectRadio.helpTip = L("tooltip.frontObject");

        var symbolSourceGroup = objectSourcePanel.add("group");
        symbolSourceGroup.orientation = "row";
        symbolSourceGroup.spacing = 8;
        var symbolRadio = symbolSourceGroup.add("radiobutton", undefined, L("radio.symbol"));
        symbolRadio.helpTip = L("tooltip.symbol");
        var symbolDropdown = symbolSourceGroup.add("dropdownlist", undefined, symbolNames);
        symbolDropdown.helpTip = L("tooltip.symbolDropdown");
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
        var anchorPointPanel = dialogWindow.add("panel", undefined, L("panel.anchorPoint"));
        anchorPointPanel.orientation = "row";
        anchorPointPanel.alignChildren = ["left", "top"];
        anchorPointPanel.margins = [16, 20, 16, 16];
        anchorPointPanel.spacing = 24;

        // 左カラム：大きさ・シンボル化 / Left column: size, symbolize
        var anchorLeft = anchorPointPanel.add("group");
        anchorLeft.orientation = "column";
        anchorLeft.alignChildren = ["left", "center"];
        anchorLeft.spacing = 12;

        var sizeInput = addLabeledField(anchorLeft, L("label.squareSize"), DEFAULTS.squareSize, 2, L("label.unit"), renderPreview, undefined, true, 0.1);
        sizeInput.helpTip = L("tooltip.squareSize"); /* ⌘併用で ±0.1 などのキー説明 / Key modifiers incl. ⌘ = ±0.1 */

        var symbolizeCheckbox = anchorLeft.add("checkbox", undefined, L("checkbox.symbolize"));
        symbolizeCheckbox.value = DEFAULTS.symbolize;
        symbolizeCheckbox.helpTip = L("tooltip.symbolize");

        // 右カラム：カラー（2行：カラー■ / 選択...）/ Right column: color (2 rows)
        // OSのカラーパレットはモーダルを壊すため、自前のRGBダイアログを使う
        var anchorRight = anchorPointPanel.add("group");
        anchorRight.orientation = "column";
        anchorRight.alignChildren = ["left", "center"];
        anchorRight.spacing = 0;

        var colorRow = anchorRight.add("group");
        colorRow.orientation = "row";
        colorRow.spacing = 8;
        colorRow.add("statictext", undefined, L("label.squareColor"));
        var colorSwatch = colorRow.add("panel");
        colorSwatch.preferredSize = [20, 20]; /* 正方形・高さは短いまま / Square, keep the short height */
        colorSwatch.helpTip = L("tooltip.squareColor");
        colorSwatch.onDraw = makeSwatchDrawer(colorSwatch, pickedColor);

        // 選択ボタンは上マージン5のグループで包む（コントロール直接の margins は効かない環境があるため）
        var chooseColorWrap = anchorRight.add("group");
        chooseColorWrap.margins = [0, 10, 0, 0]; /* 上にマージン5 / 5px top margin */
        var chooseColorButton = chooseColorWrap.add("button", undefined, L("button.chooseColor"));
        chooseColorButton.helpTip = L("tooltip.squareColor");
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

        // --- オプション パネル（2カラム：左＝3設定 / 右＝9axis）/ Options panel (2 columns) ---
        var optionsPanel = dialogWindow.add("panel", undefined, L("panel.options"));
        optionsPanel.orientation = "row";
        optionsPanel.alignChildren = ["left", "top"];
        optionsPanel.margins = [16, 20, 16, 16];
        optionsPanel.spacing = 16;

        // 左カラム：スケール・レイヤーに移動・グループ化 / Left column: scale, move-to-layer, group
        var optionsLeft = optionsPanel.add("group");
        optionsLeft.orientation = "column";
        optionsLeft.alignChildren = ["left", "center"];
        optionsLeft.spacing = 10;

        var scaleInput = addLabeledField(optionsLeft, L("label.scale"), DEFAULTS.scalePercent, 3, L("label.percent"), renderPreview);
        scaleInput.helpTip = L("tooltip.scale");

        var moveToLayerCheckbox = optionsLeft.add("checkbox", undefined, L("checkbox.moveToLayer"));
        moveToLayerCheckbox.value = DEFAULTS.moveToLayer;
        moveToLayerCheckbox.helpTip = L("tooltip.moveToLayer");

        var groupCheckbox = optionsLeft.add("checkbox", undefined, L("checkbox.group"));
        groupCheckbox.value = DEFAULTS.groupItems;
        groupCheckbox.helpTip = L("tooltip.group");

        // 右カラム：9axis（基準点）を天地左右中央に / Right column: 9-axis widget, centered both ways
        var optionsRight = optionsPanel.add("group");
        optionsRight.orientation = "column";
        optionsRight.alignChildren = ["center", "center"];
        optionsRight.alignment = ["center", "center"]; /* 左カラムの高さに対して天地中央 / Vertically center against the left column */
        var registrationWidget = addRegistrationWidget(optionsRight, renderPreview);

        // 現在選択されている「追加するオブジェクト」種別を返す
        function getChosenSource() {
            if (frontObjectRadio.value) return OBJECT_SOURCE.frontObject;
            if (symbolRadio.value) return OBJECT_SOURCE.symbol;
            return OBJECT_SOURCE.autoGenerate;
        }

        // 追加オブジェクトの選択に応じて各コントロールの有効/無効を切り替え
        function syncControlState() {
            var isAutoGenerate = autoGenerateRadio.value;
            sizeInput.enabled = isAutoGenerate;
            chooseColorButton.enabled = isAutoGenerate;
            symbolizeCheckbox.enabled = isAutoGenerate;
            // 自動生成では基準点を中央へ戻す / Reset registration to center in auto-generate
            if (isAutoGenerate) {
                registrationIndex = DEFAULTS.registrationIndex;
            }
            // 自動生成時はスケールと9axisのみディム（レイヤー移動・グループ化は常時有効）
            scaleInput.parent.enabled = !isAutoGenerate;
            registrationWidget.enabled = !isAutoGenerate;
            redrawControl(registrationWidget); /* 自作描画なので色を更新 / Redraw the custom widget */
            symbolDropdown.enabled = symbolRadio.value;
        }
        // ラジオが別コンテナに分かれているため排他選択を手動で担保
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

        // 現在のコントロール値から設定オブジェクトを読み取る（プレビュー・本適用の共通ソース）
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

        // プレビュー専用レイヤーに現在設定でマーカーを描画 / Draw markers into the preview layer
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
        var dialogButtonGroup = dialogWindow.add("group");
        dialogButtonGroup.orientation = "row";
        dialogButtonGroup.alignment = ["right", "center"];
        dialogButtonGroup.spacing = 8;
        var cancelButton = dialogButtonGroup.add("button", undefined, L("button.cancel"), { name: "cancel" });
        var okButton = dialogButtonGroup.add("button", undefined, L("button.ok"), { name: "ok" });

        var dialogResult = null;
        okButton.onClick = function () {
            var chosenSource = getChosenSource();
            if (chosenSource === OBJECT_SOURCE.autoGenerate) {
                var sizeValue = parseFloat(sizeInput.text);
                if (isNaN(sizeValue) || sizeValue <= 0) {
                    alert(L("alert.invalidSize"));
                    return;
                }
            } else {
                var scaleValue = parseFloat(scaleInput.text);
                if (isNaN(scaleValue) || scaleValue <= 0) {
                    alert(L("alert.invalidScale"));
                    return;
                }
                if (chosenSource === OBJECT_SOURCE.symbol && !symbolDropdown.selection) {
                    alert(L("alert.noSymbolChosen"));
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
    /* 「ラベル [入力] 単位」の 1 行を作り、入力欄を返す / Build a "label [input] unit" row and return the input */
    function addLabeledField(parentPanel, labelText, initialValue, charCount, unitText, onChange, labelWidth, allowDecimal, minValue) {
        var fieldRow = parentPanel.add("group");
        fieldRow.orientation = "row";
        fieldRow.spacing = 8;
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

    /* スウォッチ（panel）を現在色で塗る onDraw ハンドラを生成 / Make an onDraw handler that fills a swatch */
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

    /* 自前の RGB カラーダイアログ。{r,g,b} を返す（キャンセル時は null） / In-dialog RGB color chooser */
    function chooseRgbColor(startColor) {
        var workingColor = { r: startColor.r, g: startColor.g, b: startColor.b };
        var confirmed = false;

        var pickerWindow = new Window("dialog", L("colorPicker.title"));
        pickerWindow.orientation = "row";
        pickerWindow.alignChildren = ["fill", "fill"];
        pickerWindow.margins = 16;
        pickerWindow.spacing = 12;

        var previewSwatch = pickerWindow.add("panel");
        previewSwatch.preferredSize = [64, 64];
        previewSwatch.onDraw = makeSwatchDrawer(previewSwatch, workingColor);

        var fieldsColumn = pickerWindow.add("group");
        fieldsColumn.orientation = "column";
        fieldsColumn.alignChildren = ["left", "center"];
        fieldsColumn.spacing = 6;

        function refreshFromFields() {
            workingColor.r = clampColorChannel(Number(redInput.text));
            workingColor.g = clampColorChannel(Number(greenInput.text));
            workingColor.b = clampColorChannel(Number(blueInput.text));
            redrawControl(previewSwatch);
        }
        var redInput = addColorChannelField(fieldsColumn, "R", workingColor.r, refreshFromFields);
        var greenInput = addColorChannelField(fieldsColumn, "G", workingColor.g, refreshFromFields);
        var blueInput = addColorChannelField(fieldsColumn, "B", workingColor.b, refreshFromFields);

        var pickerButtonGroup = fieldsColumn.add("group");
        pickerButtonGroup.orientation = "row";
        pickerButtonGroup.alignment = ["right", "center"];
        pickerButtonGroup.spacing = 8;
        var pickerCancelButton = pickerButtonGroup.add("button", undefined, L("button.cancel"), { name: "cancel" });
        var pickerOkButton = pickerButtonGroup.add("button", undefined, L("button.ok"), { name: "ok" });

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

    /* RGB チャンネル 1 つ（ラベル＋スライダー＋数値入力を相互同期）を作り、入力欄を返す */
    /* Build one RGB channel (label + slider + numeric field, kept in sync); returns the field */
    function addColorChannelField(parentGroup, channelLabel, initialValue, onChange) {
        var channelRow = parentGroup.add("group");
        channelRow.orientation = "row";
        channelRow.spacing = 6;
        var channelLabelText = channelRow.add("statictext", undefined, channelLabel);
        channelLabelText.preferredSize.width = 14;
        var channelSlider = channelRow.add("slider", undefined, initialValue, 0, 255);
        channelSlider.preferredSize = [140, 18];
        var channelInput = channelRow.add("edittext", undefined, String(initialValue));
        channelInput.characters = 4;

        /* スライダー → 数値欄 / Slider drives the field */
        function syncFromSlider() {
            channelInput.text = Math.round(channelSlider.value);
            onChange();
        }
        /* 数値欄 → スライダー / Field drives the slider */
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

    /* 0〜255 に丸めてクランプ / Round and clamp to 0..255 */
    function clampColorChannel(value) {
        if (isNaN(value)) return 0;
        value = Math.round(value);
        if (value < 0) return 0;
        if (value > 255) return 255;
        return value;
    }

    /* 生成したプレビューレイヤーだけを参照で削除（同名の既存レイヤーは触らない）/ Remove only the preview layer we created, by reference */
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

    /* 指定名のレイヤーを取得、無ければ作成 / Get a layer by name, creating it if absent */
    function getOrCreateLayer(layerName) {
        try {
            return doc.layers.getByName(layerName);
        } catch (e) {
            var layer = doc.layers.add();
            layer.name = layerName;
            return layer;
        }
    }

    /* エッジ表示＋ライブコーナー注釈のトグル（実行中は隠す）/ Toggle edges and the Live Corner Annotator */
    function toggleCanvasHelpers() {
        app.executeMenuCommand('edge');
        app.executeMenuCommand('Live Corner Annotator');
    }

    /* ボタンの高さを指定 px 詰める（レイアウト確定後に呼ぶ）/ Trim a button's height by the given px (call after layout) */
    function trimButtonHeight(button, px) {
        try {
            button.size = [button.size.width, button.size.height - px];
        } catch (e) {}
    }

    /* 生成物を指定レイヤーの最前面へ移動 / Move a created item to the front of a layer */
    function moveItemToLayer(item, layer) {
        item.move(layer, ElementPlacement.PLACEATBEGINNING);
    }

    /* コントロールを再描画。notify は環境により例外を投げ得るので保護 / Redraw a control (notify can throw) */
    function redrawControl(control) {
        try {
            control.notify("onDraw");
        } catch (e) {}
    }

    /* 配置済みアイテムを1つのグループにまとめ、そのグループを返す / Group placed items and return the group */
    function groupPlacedItems(items) {
        var markerGroup = doc.groupItems.add();
        for (var i = 0; i < items.length; i++) {
            items[i].move(markerGroup, ElementPlacement.PLACEATEND);
        }
        return markerGroup;
    }

    /*
     * ↑↓キーで値を増減（Shift=±10・10スナップ / ⌘=±0.1 / 通常=±1）
     * allowDecimal=true で小数第1位まで保持、false で整数に丸め
     * minValue を渡すとその値未満に下げられない（未指定は 0 を下限）
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

    /* キーイベントの修飾キー状態を返す（event 優先, keyboardState はフォールバック）/ Read a modifier flag, event first */
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
    /* 選択オブジェクト群からアンカー座標 [x, y] の配列を収集 / Collect all anchor coordinates */
    function collectAllAnchorPoints(items, excludeItem) {
        var collectedPoints = [];
        for (var i = 0; i < items.length; i++) {
            collectAnchorPoints(items[i], collectedPoints, excludeItem);
        }
        return collectedPoints;
    }

    /* オブジェクトの種別に応じてアンカー座標を集める（excludeItem は対象外）/ Gather anchor coordinates, skipping excludeItem */
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

    /*
     * 配置に使うアンカー座標を返す
     * 「最前面のオブジェクト」モードでは複製元（最前面オブジェクト）自身のアンカーを除外する
     * Anchors to place onto; in frontmost-object mode, exclude the source object's own anchors
     */
    function resolveAnchorPoints(objectSource) {
        var excludeItem = (objectSource === OBJECT_SOURCE.frontObject) ? getFrontmostItem() : null;
        return collectAllAnchorPoints(selectedItems, excludeItem);
    }

    // =========================================
    // 配置処理 / Placement
    // =========================================
    /*
     * ユーザー設定から、アンカー座標に配置する処理（placer 関数）を組み立てる
     * placer は生成した pageItem を返す（プレビュー時にレイヤー移動するため）
     * forPreview=true のときは自動生成の「シンボル化」を無視し、見た目が同じ正方形パスを描く
     * （プレビューのたびに新規シンボルを登録してシンボルパネルを汚さないため）
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

    /* 基準点の割合(0..1)に合わせて生成物を配置 / Position an item so its registration point sits on the anchor */
    function positionItemAtAnchor(item, anchorX, anchorY, fractionX, fractionY) {
        item.left = anchorX - fractionX * item.width;
        item.top = anchorY + fractionY * item.height;
    }

    /* {r,g,b} を RGBColor へ変換 / Convert {r,g,b} to an RGBColor */
    function toRgbColor(rgb) {
        var color = new RGBColor();
        color.red = rgb.r;
        color.green = rgb.g;
        color.blue = rgb.b;
        return color;
    }

    /*
     * 現在の大きさ・カラーで正方形シンボルを毎回新規登録して返す（既存は再利用しない）
     * → プレビュー（正方形パス）と本適用（同設定のシンボルインスタンス）の見た目が一致する
     * Always register a fresh square symbol from the current size/color (no reuse),
     * so the preview and the actual placement look identical.
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

    /* 基準点に合わせてシンボルインスタンスを配置し、生成物を返す / Place a symbol instance at the anchor */
    function placeSymbolInstance(symbolDefinition, anchorX, anchorY, scalePercent, fractionX, fractionY) {
        var symbolInstance = doc.symbolItems.add(symbolDefinition);
        if (scalePercent !== 100) {
            symbolInstance.resize(scalePercent, scalePercent);
        }
        positionItemAtAnchor(symbolInstance, anchorX, anchorY, fractionX, fractionY);
        return symbolInstance;
    }

    /* 基準点に合わせて最前面オブジェクトを複製し、生成物を返す / Duplicate the frontmost object at the anchor */
    function duplicateItemAtPoint(sourceItem, anchorX, anchorY, scalePercent, fractionX, fractionY) {
        var duplicatedItem = sourceItem.duplicate();
        if (scalePercent !== 100) {
            duplicatedItem.resize(scalePercent, scalePercent);
        }
        positionItemAtAnchor(duplicatedItem, anchorX, anchorY, fractionX, fractionY);
        return duplicatedItem;
    }

    /* 基準点に合わせて正方形パスを配置し、生成物を返す / Place a square path at the anchor */
    function placeSquareRect(squareSize, squareColor, anchorX, anchorY, fractionX, fractionY) {
        var squareRect = doc.pathItems.rectangle(0, 0, squareSize, squareSize);
        squareRect.filled = true;
        squareRect.fillColor = squareColor;
        squareRect.stroked = false;
        positionItemAtAnchor(squareRect, anchorX, anchorY, fractionX, fractionY);
        return squareRect;
    }

    /*
     * 「選択範囲内で最前面のオブジェクト」を返す（未選択の最前面は対象外）
     * ドキュメントを前面から走査し、最初に selected なアイテムを返す。未選択オブジェクトを
     * 複製元に選んでしまう事故を防ぐ。
     * Return the frontmost object *within the selection* (never an unselected one):
     * scan the document front-to-back and return the first selected item.
     */
    function getFrontmostItem() {
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

    /* コンテナ（レイヤー／グループ）を前面から走査し、最初に選択されているアイテムを返す */
    /* Scan a container (layer / group) front-to-back for the first selected item */
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

    /* ドキュメント内シンボルの名前配列を返す / Return document symbol names */
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
    /* 3×3 の基準点ウィジェットを生成。クリックで registrationIndex を更新し onChange を呼ぶ */
    function addRegistrationWidget(parentGroup, onChange) {
        var widget = parentGroup.add("button", undefined, "");
        widget.helpTip = L("tooltip.registration");
        widget.preferredSize = [56, 56];
        widget.minimumSize = [56, 56];
        widget.maximumSize = [56, 56];
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

    /* 0..2 にクランプ / Clamp to 0..2 */
    function clampCell(value) {
        if (value < 0) return 0;
        if (value > 2) return 2;
        return value;
    }

    /* 9軸ウィジェットを描画（外周の□をケイ線でつなぐ・中央は独立）/ Draw the 9-axis widget */
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

        function cellX(index) { return originX + (index % 3) * cellStep; }
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

    /* 基準点セルの□を1つ描画。選択中は「塗り＋ケイ線」で、非選択の枠線と外周サイズを合わせる */
    /* Draw one cell square; selected uses fill + rule so its outer size matches the outlined cells */
    function drawRegistrationCell(graphics, x, y, size, selected, foreColor) {
        if (selected) {
            buildCellPath(graphics, x, y, size);
            graphics.fillPath(graphics.newBrush(graphics.BrushType.SOLID_COLOR, foreColor));
        }
        buildCellPath(graphics, x, y, size);
        graphics.strokePath(graphics.newPen(graphics.PenType.SOLID_COLOR, foreColor, 1));
    }

    /* セルの正方形パスを組む / Build the square path for a cell */
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
    /* UI 明度 > 0.5 ならライト、取得失敗時はダーク扱い / Light if uiBrightness > 0.5; dark on failure */
    function isLightUI() {
        try {
            return app.preferences.getRealPreference("uiBrightness") > 0.5;
        } catch (e) {
            return false;
        }
    }
})();
