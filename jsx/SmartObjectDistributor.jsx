#target illustrator

/*
 * スクリプト名：SmartObjectDistributor.jsx
 * バージョン: 0.5.8
 * 概要 / Overview：
 * アートボードまたは「_target」レイヤー内の図形を基準にグリッド分割し、
 * 選択オブジェクトを各セル中央に配置するスクリプト。
 * 背景長方形の描画、ガイド化、アートボード化、リアルタイムプレビューに対応。
 *
 * 主な機能 / Features：
 * - グリッド分割（行・列・マージン・ガター）
 * - 背景長方形の描画（色／不透明度／透過）とガイド・アートボード変換
 * - オブジェクトの中央配置とランダム配置
 * - プレビュー対応、UIリアルタイム反映、アンドゥ可能
 * - 「_target」レイヤー図形の一時アートボード化／非表示処理
 *
 * 対象 / Target：アクティブなアートボード、または _target レイヤー内の長方形
 *
 * 作成日：2025-05-20
 * 更新日：2025-06-05
 * - 0.5.3 行間と列間を統一し、マージンも共通値のみに簡素化
 * - 0.5.4 UI構造とラベル整理、cell-background レイヤーの自動削除を追加
 * - 0.5.5 行列、行間・段間、マージンを簡易化、直前に実行したセルを削除
 * - 0.5.6 UI構成と有効制御を調整
 * - 0.5.7 マージンの変更が即時プレビューに反映されるよう調整
 * - 0.5.8 _targetレイヤーの長方形をアートボードとして一時使用／非表示化処理を追加
 */

// 言語判定関数とラベル定義（グローバル）
function getCurrentLang() {
    return ($.locale && $.locale.indexOf('ja') === 0) ? 'ja' : 'en';
}
var lang = getCurrentLang();

var labels = {
    dialogTitle: {
        ja: "グリッド状に配置",
        en: "Distribute in Grid"
    },
    infoText: {
        ja: "「_target」レイヤーの長方形を優先",
        en: "Use rectangle in _target layer first"
    },
    gridSettingsLabel: {
        ja: "分割とマージン",
        en: "Grid Settings"
    },
    rowsLabel: {
        ja: "行数：",
        en: "Rows:"
    },
    columnsLabel: {
        ja: "列数：",
        en: "Columns:"
    },
    rowGutterLabel: {
        ja: "間隔：",
        en: "Gutter:"
    },
    commonMarginLabel: {
        ja: "マージン：",
        en: "Margin:"
    },
    cellDrawingLabel: {
        ja: "セル描画",
        en: "Cell Drawing"
    },
    cellRectLabel: {
        ja: "残す",
        en: "Keep Cell"
    },
    convertToGuideLabel: {
        ja: "ガイド化",
        en: "Convert to Guides"
    },
    convertToArtboardLabel: {
        ja: "アートボード化",
        en: "Convert to Artboards"
    },
    colorLabel: {
        ja: "カラー：",
        en: "Color:"
    },
    black: {
        ja: "黒",
        en: "Black"
    },
    white: {
        ja: "白",
        en: "White"
    },
    none: {
        ja: "透過",
        en: "Transparent"
    },
    opacityLabel: {
        ja: "不透明度：",
        en: "Opacity:"
    },
    randomLabel: {
        ja: "ランダム",
        en: "Random"
    },
    cancelLabel: {
        ja: "キャンセル",
        en: "Cancel"
    },
    okLabel: {
        ja: "OK",
        en: "OK"
    },
    noSelection: {
        ja: "オブジェクトが選択されていません。",
        en: "No objects selected."
    },
    artboardError: {
        ja: "アートボードの作成中にエラーが発生しました。",
        en: "Error occurred while creating artboards."
    },
    artboardCreated: {
        ja: " 個のアートボードを作成しました。",
        en: " artboards created."
    }
};

// 選択オブジェクトがアートボードに重なっている場合、他と重ならない位置へ一時的に退避
function moveObjectsOutsideArtboard(doc, selection) {
    var activeArtboardIndex = doc.artboards.getActiveArtboardIndex();
    var activeAB = doc.artboards[activeArtboardIndex].artboardRect;
    var left = activeAB[0];
    var top = activeAB[1];
    var right = activeAB[2];
    var bottom = activeAB[3];

    var movedObjects = [];
    var BOUNDARY_MARGIN = 1;
    var ARTBOARD_BUFFER = 10;

    function isOverlappingWithMovedObjects(newBounds) {
        for (var i = 0; i < movedObjects.length; i++) {
            var b = movedObjects[i];
            if (!(newBounds[2] < b[0] - BOUNDARY_MARGIN || newBounds[0] > b[2] + BOUNDARY_MARGIN ||
                    newBounds[1] < b[3] - BOUNDARY_MARGIN || newBounds[3] > b[1] + BOUNDARY_MARGIN)) {
                return true;
            }
        }
        return false;
    }

    function isOverlappingWithArtboards(newBounds) {
        for (var i = 0; i < doc.artboards.length; i++) {
            if (i === activeArtboardIndex) continue;
            var ab = doc.artboards[i].artboardRect;
            if (!(newBounds[2] < ab[0] - BOUNDARY_MARGIN - ARTBOARD_BUFFER ||
                    newBounds[0] > ab[2] + BOUNDARY_MARGIN + ARTBOARD_BUFFER ||
                    newBounds[1] < ab[3] - BOUNDARY_MARGIN - ARTBOARD_BUFFER ||
                    newBounds[3] > ab[1] + BOUNDARY_MARGIN + ARTBOARD_BUFFER)) {
                return true;
            }
        }
        return false;
    }

    for (var i = 0; i < selection.length; i++) {
        var item = selection[i];
        var bounds = item.visibleBounds;
        var isOverlapping = !(bounds[2] < left || bounds[0] > right || bounds[1] < bottom || bounds[3] > top);

        if (isOverlapping) {
            var dx = (right - bounds[0]) + 200;
            do {
                var newBounds = [bounds[0] + dx, bounds[1], bounds[2] + dx, bounds[3]];
                dx += 200;
            } while (isOverlappingWithMovedObjects(newBounds) || isOverlappingWithArtboards(newBounds));

            item.position = [item.position[0] + dx - 200, item.position[1]];
            movedObjects.push(newBounds);
        }
    }

}
var originalPositions = [];

function saveOriginalPositions(items) {
    originalPositions = [];
    for (var i = 0; i < items.length; i++) {
        var bounds = items[i].visibleBounds;
        var centerX = (bounds[0] + bounds[2]) / 2;
        var centerY = (bounds[1] + bounds[3]) / 2;
        originalPositions.push([centerX, centerY]);
    }
}

// 記録した中心座標へ戻す（drawGuides 内で呼び出し）
function restoreOriginalPositions(items) {
    if (!originalPositions || originalPositions.length === 0) return;
    for (var i = 0; i < items.length && i < originalPositions.length; i++) {
        var bounds = items[i].visibleBounds;
        var itemCenterX = (bounds[0] + bounds[2]) / 2;
        var itemCenterY = (bounds[1] + bounds[3]) / 2;
        var dx = originalPositions[i][0] - itemCenterX;
        var dy = originalPositions[i][1] - itemCenterY;
        items[i].translate(dx, dy);
    }
}

function main() {
    addTemporaryArtboardFromTarget();
    if (app.documents.length === 0) {
        alert("ドキュメントを開いてください。\nPlease open a document.");
        return;
    }
    // "cell-background" レイヤーが存在する場合は削除
    try {
        var existingLayer = app.activeDocument.layers.getByName("cell-background");
        if (existingLayer) existingLayer.remove();
    } catch (e) {
        // 存在しない場合はエラーを無視
    }

    var doc = app.activeDocument;
    var rulerUnit = app.preferences.getIntegerPreference("rulerType");
    var unitLabel = "pt";
    var unitFactor = 1.0;
    // 単位設定 (switch文)
    switch (rulerUnit) {
        case 0: // inch
            unitLabel = "inch";
            unitFactor = 72.0;
            break;
        case 1: // mm
            unitLabel = "mm";
            unitFactor = 72.0 / 25.4;
            break;
        case 2: // pt
            unitLabel = "pt";
            unitFactor = 1.0;
            break;
        case 3: // pica
            unitLabel = "pica";
            unitFactor = 12.0;
            break;
        case 4: // cm
            unitLabel = "cm";
            unitFactor = 72.0 / 2.54;
            break;
        case 5: // Q
            unitLabel = "Q";
            unitFactor = 72.0 / 25.4 * 0.25;
            break;
        case 6: // px
            unitLabel = "px";
            unitFactor = 1.0;
            break;
        default:
            unitLabel = "pt";
            unitFactor = 1.0;
    }
    // ダイアログ作成
    var dlg = new Window("dialog", labels.dialogTitle[lang]);
    dlg.orientation = "column";
    dlg.alignChildren = "fill";

    var infoGroup = dlg.add("group");
    infoGroup.orientation = "column";
    infoGroup.alignChildren = "center";
    infoGroup.alignment = "fill";
    infoGroup.margins = [0, 0, 0, 10];

    var infoText = infoGroup.add("statictext", undefined, labels.infoText[lang]);
    infoText.preferredSize.height = 24;

    // グリッド設定パネル（2列構成、タイトル付き）
    var gridPanel = dlg.add("panel", undefined, labels.gridSettingsLabel[lang]);
    gridPanel.orientation = "row";
    gridPanel.alignChildren = "top";
    gridPanel.margins = [15, 20, 15, 15];
    gridPanel.spacing = 20;

    // ラベル幅を日本語環境のときに 40/55、それ以外は 55/70 に個別指定
    var leftLabelWidth = (lang === 'ja') ? 40 : 55;
    var rightLabelWidth = (lang === 'ja') ? 65 : 70;

    // 2列グループ（gridPanel直下へ追加）
    var leftColGroup = gridPanel.add("group");
    leftColGroup.orientation = "column";
    leftColGroup.alignChildren = "left";

    var rightColGroup = gridPanel.add("group");
    rightColGroup.orientation = "column";
    rightColGroup.alignChildren = "left";

    // 左側：行数・列数（ペアで整える）
    var inputY = leftColGroup.add("group");
    inputY.add("statictext", undefined, labels.rowsLabel[lang]).preferredSize.width = leftLabelWidth;
    var inputYText = inputY.add("edittext", undefined, "5");
    inputYText.characters = 3;

    var inputX = leftColGroup.add("group");
    inputX.add("statictext", undefined, labels.columnsLabel[lang]).preferredSize.width = leftLabelWidth;
    var inputXText = inputX.add("edittext", undefined, "5");
    inputXText.characters = 3;

    // 右側：行間・マージン（ペアで整える）
    var rowGutterGroup = rightColGroup.add("group");
    rowGutterGroup.alignChildren = "center";
    rowGutterGroup.add("statictext", undefined, labels.rowGutterLabel[lang]).preferredSize.width = rightLabelWidth;
    var inputRowGutter = rowGutterGroup.add("edittext", undefined, "10");
    inputRowGutter.characters = 4;
    rowGutterGroup.add("statictext", undefined, unitLabel);

    var marginGroup = rightColGroup.add("group");
    marginGroup.orientation = "row";
    marginGroup.alignChildren = "center";
    marginGroup.add("statictext", undefined, labels.commonMarginLabel[lang]).preferredSize.width = rightLabelWidth;
    var marginInput = marginGroup.add("edittext", undefined, "10");
    marginInput.characters = 4;
    marginGroup.add("statictext", undefined, unitLabel);

    // オプション設定グループ
    var optGroup = dlg.add("group");
    optGroup.orientation = "column";
    optGroup.alignChildren = "left";

    // === 長方形オプションパネル ===
    var rectPanel = optGroup.add("panel", undefined, labels.cellDrawingLabel[lang]);
    rectPanel.orientation = "column";
    rectPanel.alignChildren = "left";
    rectPanel.margins = [15, 20, 15, 15];

    // 横並びのグループにチェックボックスをまとめる
    var rectOptionsGroup = rectPanel.add("group");
    rectOptionsGroup.orientation = "row";
    rectOptionsGroup.alignChildren = "left";

    var cellRectCheckbox = rectOptionsGroup.add("checkbox", undefined, labels.cellRectLabel[lang]);
    cellRectCheckbox.value = true;

    var convertToGuideCheckbox = rectOptionsGroup.add("checkbox", undefined, labels.convertToGuideLabel[lang]);
    convertToGuideCheckbox.value = false;

    // === 「アートボード化」チェックボックス追加 ===
    var convertToArtboardCheckbox = rectOptionsGroup.add("checkbox", undefined, labels.convertToArtboardLabel[lang]);
    convertToArtboardCheckbox.value = false;
    // オプション切り替え時のUI状態反映（冗長な分岐を整理）
    function handleToggleOptions(isArtboardToggled, isGuideToggled) {
        var isArtboard = isArtboardToggled !== undefined ? isArtboardToggled : convertToArtboardCheckbox.value;
        var isGuide = isGuideToggled !== undefined ? isGuideToggled : convertToGuideCheckbox.value;
        if (isArtboard) {
            setArtboardOptionUI();
            renderGrid(true, true, true);
        } else if (isGuide) {
            setGuideOptionUI();
            renderGrid(true, true, isGuide);
        } else {
            setNormalOptionUI();
            renderGrid(true);
        }
    }
    // サブ関数：アートボード化時のUI
    function setArtboardOptionUI() {
        cellRectCheckbox.enabled = false;
        cellRectCheckbox.value = true;
        convertToGuideCheckbox.enabled = false;
        convertToGuideCheckbox.value = false;
        rbBlack.enabled = false;
        rbWhite.enabled = false;
        rbNone.enabled = false;
        rbNone.value = true;
        inputOpacity.enabled = false;
    }
    // サブ関数：ガイド化時のUI
    function setGuideOptionUI() {
        cellRectCheckbox.enabled = true;
        convertToGuideCheckbox.enabled = true;
        rbBlack.enabled = false;
        rbWhite.enabled = false;
        rbNone.enabled = false;
        rbNone.value = true;
        inputOpacity.enabled = false;
        updateColorOpacity();
    }
    // サブ関数：通常時のUI
    function setNormalOptionUI() {
        cellRectCheckbox.enabled = true;
        cellRectCheckbox.value = true;
        convertToGuideCheckbox.enabled = cellRectCheckbox.value;
        rbBlack.enabled = true;
        rbWhite.enabled = true;
        rbNone.enabled = true;
        rbBlack.value = true;
        inputOpacity.enabled = rbBlack.value || rbWhite.value;
    }
    cellRectCheckbox.onClick = function() {
        convertToGuideCheckbox.enabled = cellRectCheckbox.value;
        if (!cellRectCheckbox.value) {
            convertToGuideCheckbox.value = false;
        }
        renderGrid(true);
    };
    convertToArtboardCheckbox.onClick = function() {
        handleToggleOptions(convertToArtboardCheckbox.value, false);
    };
    convertToGuideCheckbox.onClick = function() {
        handleToggleOptions(false, convertToGuideCheckbox.value);
    };

    // カラー選択ラジオ（黒・白・透過）
    var colorGroup = rectPanel.add("group");
    colorGroup.orientation = "row";
    colorGroup.alignChildren = "left";

    colorGroup.add("statictext", undefined, labels.colorLabel[lang]);

    var rbBlack = colorGroup.add("radiobutton", undefined, labels.black[lang]);
    var rbWhite = colorGroup.add("radiobutton", undefined, labels.white[lang]);
    var rbNone = colorGroup.add("radiobutton", undefined, labels.none[lang]);

    // 不透明度設定行を追加
    var opacityGroup = rectPanel.add("group");
    opacityGroup.orientation = "row";
    opacityGroup.alignChildren = "center";

    opacityGroup.add("statictext", undefined, labels.opacityLabel[lang]);

    var inputOpacity = opacityGroup.add("edittext", undefined, "15");
    inputOpacity.characters = 4;
    opacityGroup.add("statictext", undefined, "%");

    rbBlack.value = true; // 初期選択は黒

    // 不透明度欄：黒・白のとき有効、透過時は無効
    function updateColorOpacity() {
        var previousValue = inputOpacity.text;
        var previousEnabled = inputOpacity.enabled;
        inputOpacity.enabled = rbBlack.value || rbWhite.value;
        // カラー選択時の不透明度値をUIに反映
        if (rbBlack.value) {
            inputOpacity.text = "15";
        } else if (rbWhite.value) {
            inputOpacity.text = "100";
        }
        if (previousValue !== inputOpacity.text || previousEnabled !== (rbBlack.value || rbWhite.value)) {
            renderGrid(true, true, convertToGuideCheckbox.value);
        }
    }
    rbBlack.onClick = rbWhite.onClick = rbNone.onClick = updateColorOpacity;
    // 初期状態反映
    updateColorOpacity();

    // 不透明度変更時のみプレビュー更新
    var lastOpacityValue = inputOpacity.text;
    inputOpacity.onChanging = function() {
        if (inputOpacity.text !== lastOpacityValue) {
            lastOpacityValue = inputOpacity.text;
            renderGrid(true, true, convertToGuideCheckbox.value);
        }
    };


    // === ボタンエリア（レイアウト変更版）===
    var outerGroup = dlg.add("group");
    outerGroup.orientation = "row";
    outerGroup.alignChildren = ["fill", "center"];
    outerGroup.margins = [0, 10, 0, 0];
    outerGroup.spacing = 0;



    // --- 左グループ（キャンセルボタンのみ） ---
    var leftGroup = outerGroup.add("group");
    leftGroup.orientation = "row";
    leftGroup.alignChildren = "left";

    var btnRandom = leftGroup.add("button", undefined, labels.randomLabel[lang]);

    // ランダムボタン押下時
    btnRandom.onClick = function() {
        var selectedItems = doc.selection;
        if (!selectedItems || selectedItems.length === 0) {
            alert(labels.noSelection[lang]);
            return;
        }
        randomizedSelection = [];
        for (var i = 0; i < selectedItems.length; i++) {
            randomizedSelection.push(selectedItems[i]);
        }
        // Fisher-Yates シャッフルで順序をランダム化
        for (var i = randomizedSelection.length - 1; i > 0; i--) {
            var j = Math.floor(Math.random() * (i + 1));
            var temp = randomizedSelection[i];
            randomizedSelection[i] = randomizedSelection[j];
            randomizedSelection[j] = temp;
        }
        // プレビューを更新
        renderGrid(true, true, convertToGuideCheckbox.value);
        app.redraw();
    };

    // スペーサー（横に伸びる空白）
    var spacer = outerGroup.add("group");
    spacer.alignment = ["fill", "fill"];
    spacer.minimumSize.width = (lang === 'ja') ? 80 : 120;
    spacer.maximumSize.height = 0;

    // 右グループ（OKボタンのみ）
    var rightGroup = outerGroup.add("group");
    rightGroup.orientation = "row";
    rightGroup.alignChildren = "right";
    rightGroup.spacing = 10;
    var btnCancel = rightGroup.add("button", undefined, labels.cancelLabel[lang], {
        name: "cancel"
    });

    var btnOK = rightGroup.add("button", undefined, labels.okLabel[lang], {
        name: "ok"
    });

    // 行数・列数に応じてガター入力欄の有効/無効を切り替える
    function updateGutterEnable() {
        var yVal = parseInt(inputYText.text, 10);
        inputRowGutter.enabled = (yVal > 1);
    }
    inputXText.onChanging = inputYText.onChanging = function() {
        updateGutterEnable();
        renderGrid(true);
    };


    inputRowGutter.onChanging = function() {
        renderGrid(true);
    };
    marginInput.onChanging = function() {
        renderGrid(true);
    };


    // グローバル変数：ランダム順序保持用
    var randomizedSelection = null;

    // OKボタン押下時
    btnOK.onClick = function() {
        updateGutterEnable();
        randomizedSelection = null;
        originalPositions = [];

        app.executeMenuCommand("deselectall");

        // === convertToArtboardCheckboxがONの場合は_Preview_Backgroundレイヤーの長方形をアートボードに変換 ===
        if (convertToArtboardCheckbox.value) {
            try {
                var bgLayer = doc.layers.getByName("_Preview_Background");
                var bgRects = [];
                for (var i = 0; i < bgLayer.pathItems.length; i++) {
                    var item = bgLayer.pathItems[i];
                    if (item.typename === "PathItem") {
                        bgRects.push(item);
                    }
                }
                var createdCount = 0;
                for (var j = 0; j < bgRects.length; j++) {
                    var bounds = bgRects[j].geometricBounds;
                    var left = bounds[0],
                        top = bounds[1],
                        right = bounds[2],
                        bottom = bounds[3];
                    if (isNaN(left) || isNaN(top) || isNaN(right) || isNaN(bottom) || right <= left || top <= bottom) {
                        continue;
                    }
                    try {
                        doc.artboards.add([left, top, right, bottom]);
                        createdCount++;
                    } catch (e) {
                        alert("アートボード作成エラー：" + e.message + "\n座標：" + bounds.join(", "));
                    }
                }
                alert(createdCount + labels.artboardCreated[lang]);
            } catch (e) {
                alert(labels.artboardError[lang]);
            }
        }

        // === OK時にcell-backgroundレイヤー内の図形を選択状態にする ===
        try {
            var bgLayerForSelect = doc.layers.getByName("cell-background");
            var bgItems = [];
            for (var i = 0; i < bgLayerForSelect.pathItems.length; i++) {
                var it = bgLayerForSelect.pathItems[i];
                if (it.typename === "PathItem") {
                    bgItems.push(it);
                }
            }
            doc.selection = bgItems;
        } catch (e) {
            // レイヤーが見つからない場合は無視
        }

        // 移動: OK時のグリッド描画・ガイド消去
        removePreviewGuides();
        var shouldDraw = !convertToArtboardCheckbox.value;
        renderGrid(false, shouldDraw, convertToGuideCheckbox.value);

        dlg.close(1);
        removeTemporaryArtboard();
    };

    // キャンセルボタン押下時
    btnCancel.onClick = function() {
        removePreviewGuides(); // プレビュー削除
        randomizedSelection = null;
        originalPositions = [];
        dlg.close(0); // ダイアログを閉じる
        removeTemporaryArtboard();
    };
    // 長方形描画処理（背景長方形は専用レイヤーに描画する）
    function drawCellRectangles(cellLayer, baseLeft, baseTop, cellWidth, cellHeight, xDiv, yDiv, colGutter, rowGutter, convertToGuide, isPreview) {
        var opacityValue = parseFloat(inputOpacity.text);
        for (var row = 0; row < yDiv; row++) {
            var cellY = baseTop - (cellHeight + rowGutter) * row;
            for (var col = 0; col < xDiv; col++) {
                var cellX = baseLeft + (cellWidth + colGutter) * col;
                var rect = cellLayer.pathItems.rectangle(cellY, cellX, cellWidth, cellHeight);
                rect.stroked = false;
                // カラーラジオボタンの値による分岐
                if (rbBlack.value) {
                    rect.filled = true;
                    rect.fillColor = createBlackColor();
                } else if (rbWhite.value) {
                    rect.filled = true;
                    rect.fillColor = createWhiteColor();
                } else {
                    rect.filled = false;
                }
                if (!isNaN(opacityValue)) {
                    rect.opacity = opacityValue;
                }
                if (convertToGuide) {
                    rect.guides = true;
                }
            }
        }
        // 背景長方形レイヤーを最背面に移動（プレビュー時"_Preview_Guides"または確定描画時"cell-rectangle"）
        if ((isPreview && cellLayer.name === "_Preview_Guides") || (!isPreview && cellLayer.name === "cell-rectangle")) {
            cellLayer.zOrder(ZOrderMethod.SENDTOBACK);
        }
    }

    // 各長方形に1つずつオブジェクトを配置
    function distributeSelectionToCells(selectedItems, baseLeft, baseTop, cellWidth, cellHeight, xDiv, yDiv, colGutter, rowGutter) {
        if (!selectedItems || selectedItems.length === 0) return;
        var items = randomizedSelection || selectedItems;
        var index = 0;
        for (var row = 0; row < yDiv; row++) {
            var cellY = baseTop - (cellHeight + rowGutter) * row;
            for (var col = 0; col < xDiv; col++) {
                if (index >= items.length) return;
                var cellX = baseLeft + (cellWidth + colGutter) * col;
                var cellCenterX = cellX + cellWidth / 2;
                var cellCenterY = cellY - cellHeight / 2;
                var item = items[index];
                var bounds = item.visibleBounds;
                var itemCenterX = (bounds[0] + bounds[2]) / 2;
                var itemCenterY = (bounds[1] + bounds[3]) / 2;
                var dx = cellCenterX - itemCenterX;
                var dy = cellCenterY - itemCenterY;
                item.translate(dx, dy);
                index++;
            }
        }
    }

    // ガイド描画／長方形描画／オブジェクト配置を制御
    // 背景長方形は専用レイヤーに描画する
    function renderGrid(isPreview, keepRects, convertToGuide) {
        // プレビュー時、選択オブジェクトの初期位置を復元（cellRectCheckbox.value が true の場合のみ）
        if (isPreview && doc.selection.length > 0 && cellRectCheckbox.value) {
            if (originalPositions.length === 0) saveOriginalPositions(doc.selection);
            restoreOriginalPositions(doc.selection);
            moveObjectsOutsideArtboard(doc, doc.selection);
        }

        if (isPreview) {
            removePreviewGuides();
        }

        var xDiv = parseInt(inputXText.text, 10);
        var yDiv = parseInt(inputYText.text, 10);
        // 個別マージン削除・共通マージンのみ使用
        var marginVal = parseFloat(marginInput.text) * unitFactor;
        var top = marginVal;
        var bottom = marginVal;
        var left = marginVal;
        var right = marginVal;
        var rowGutter = parseFloat(inputRowGutter.text) * unitFactor;
        var colGutter = rowGutter;
        var drawCells = cellRectCheckbox.value;

        if (isNaN(xDiv) || xDiv <= 0 || isNaN(yDiv) || yDiv <= 0) return;

        // レイヤー名整理：ガイドレイヤーは現在は長方形と配置用にのみ利用
        var gridLayerName = isPreview ? "_Preview_Guides" : "placement_layer";
        var gridLayer;
        try {
            gridLayer = doc.layers.getByName(gridLayerName);
        } catch (e) {
            gridLayer = doc.layers.add();
            gridLayer.name = gridLayerName;
        }
        gridLayer.locked = false;

        // 背景長方形は専用レイヤーに描画する
        var bgLayerName = isPreview ? "_Preview_Background" : "cell-background";
        var cellLayer;
        try {
            cellLayer = doc.layers.getByName(bgLayerName);
        } catch (e) {
            cellLayer = doc.layers.add();
            cellLayer.name = bgLayerName;
        }
        cellLayer.locked = false;

        var b = doc.artboards.getActiveArtboardIndex();
        var ab = doc.artboards[b];
        var rect = ab.artboardRect;
        var abLeft = rect[0],
            abTop = rect[1],
            abRight = rect[2],
            abBottom = rect[3];
        var baseLeft = abLeft + left;
        var baseRight = abRight - right;
        var baseTop = abTop - top;
        var baseBottom = abBottom + bottom;

        var usableWidth = baseRight - baseLeft;
        var usableHeight = baseTop - baseBottom;
        var totalColGutter = (xDiv - 1) * colGutter;
        var totalRowGutter = (yDiv - 1) * rowGutter;
        var cellWidth = (usableWidth - totalColGutter) / xDiv;
        var cellHeight = (usableHeight - totalRowGutter) / yDiv;

        if ((drawCells || (convertToArtboardCheckbox.value && isPreview)) && cellLayer) {
            drawCellRectangles(cellLayer, baseLeft, baseTop, cellWidth, cellHeight, xDiv, yDiv, colGutter, rowGutter, convertToGuide, isPreview);

            if (isPreview) {
                try {
                    var fg = doc.layers.getByName("_Preview_Guides");
                    var bg = doc.layers.getByName("_Preview_Background");
                    fg.zOrder(ZOrderMethod.BRINGTOFRONT);
                    bg.zOrder(ZOrderMethod.SENDTOBACK);
                } catch (e) {}
            }

            if (drawCells) {
                distributeSelectionToCells(doc.selection, baseLeft, baseTop, cellWidth, cellHeight, xDiv, yDiv, colGutter, rowGutter);
            }
        }

        // OK実行時：描画オフなら長方形レイヤーを削除
        if (!isPreview && drawCells && !keepRects) {
            try {
                var cellLayerToRemove = doc.layers.getByName("cell-background");
                if (cellLayerToRemove) cellLayerToRemove.remove();
            } catch (e) {}
        }

        // OK実行時の後処理
        if (!isPreview) {
            // placement_layer を削除
            try {
                var placement = doc.layers.getByName("placement_layer");
                if (placement) placement.remove();
            } catch (e) {}

            // cell-background を最背面に
            try {
                var bgFinal = doc.layers.getByName("cell-background");
                if (bgFinal) bgFinal.zOrder(ZOrderMethod.SENDTOBACK);
            } catch (e) {}

            // gridLayerをロック（削除されていなければ）
            try {
                if (gridLayer) gridLayer.locked = true;
            } catch (e) {}
        }

        if (isPreview) {
            app.redraw();
        }
    }
    // 黒色（CMYK or RGB）を返す
    function createBlackColor() {
        if (doc.documentColorSpace === DocumentColorSpace.CMYK) {
            var cmyk = new CMYKColor();
            cmyk.cyan = 0;
            cmyk.magenta = 0;
            cmyk.yellow = 0;
            cmyk.black = 100;
            return cmyk;
        } else {
            var rgb = new RGBColor();
            rgb.red = 0;
            rgb.green = 0;
            rgb.blue = 0;
            return rgb;
        }
    }
    // 白色（CMYK or RGB）を返す
    function createWhiteColor() {
        if (doc.documentColorSpace === DocumentColorSpace.CMYK) {
            var cmyk = new CMYKColor();
            cmyk.cyan = 0;
            cmyk.magenta = 0;
            cmyk.yellow = 0;
            cmyk.black = 0;
            return cmyk;
        } else {
            var rgb = new RGBColor();
            rgb.red = 255;
            rgb.green = 255;
            rgb.blue = 255;
            return rgb;
        }
    }

    // プレビュー用レイヤー削除（背景レイヤーも削除）
    function removePreviewGuides() {
        try {
            var previewLayer = doc.layers.getByName("_Preview_Guides");
            if (previewLayer) previewLayer.remove();
        } catch (e) {}
        try {
            var bgLayer = doc.layers.getByName("_Preview_Background");
            if (bgLayer) bgLayer.remove();
        } catch (e) {}
    }


    // convertToGuideCheckboxの有効/無効をcellRectCheckbox.valueに応じて初期設定
    convertToGuideCheckbox.enabled = cellRectCheckbox.value;

    // ダイアログ初期プレビュー＆終了時処理
    updateGutterEnable();
    renderGrid(true, true, convertToGuideCheckbox.value); // プレビュー時は仮でtrue

    dlg.show();
}
main();

var tempArtboardIndex = -1;

function addTemporaryArtboardFromTarget() {
    var doc = app.activeDocument;
    try {
        var layer = doc.layers.getByName("_target");
        for (var i = 0; i < layer.pathItems.length; i++) {
            var item = layer.pathItems[i];
            if (item.typename === "PathItem") {
                var bounds = item.geometricBounds;
                tempArtboardIndex = doc.artboards.length;
                doc.artboards.add(bounds);
                item.hidden = true;
                return;
            }
        }
    } catch (e) {
        // _targetレイヤーが見つからない or 無視
    }
}

function removeTemporaryArtboard() {
    try {
        if (tempArtboardIndex >= 0 && tempArtboardIndex < app.activeDocument.artboards.length) {
            app.activeDocument.artboards.remove(tempArtboardIndex);
            tempArtboardIndex = -1;
        }
    } catch (e) {
        // アートボード削除時エラーは無視
    }
}