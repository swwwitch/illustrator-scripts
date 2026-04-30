#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
### スクリプト名：

SmartObjectDistributor.jsx

### 概要

- 選択オブジェクトを、指定した行数・列数のグリッドに沿って整列配置するIllustrator用スクリプトです。
- 配置先は現在のアートボード／最背面オブジェクト／「_target」レイヤー内の長方形から選択できます。
- セル描画モード（残す／ガイド化／アートボード化）をラジオボタンで切り替え、排他的に処理します。
- セル数を超えるオブジェクトはアートボード外へ退避し、プレビュー状態のまま確定します。
- PreviewManager により、プレビュー更新時のUndo履歴を汚さず、OK時は1回のUndoで戻せます。
- 透明グリッドの表示はスクリプト内で状態を保持し、終了時に元へ戻します。

### 主な機能

- グリッド分割（行数・列数・マージン・ガター）設定
- 現在のアートボード／最背面オブジェクト／「_target」レイヤー内長方形を配置先に指定
- セル背景の描画、ガイド化、アートボード化
- セル中央配置およびランダム配置
- 即時プレビューとUndo履歴を抑えた確定処理
- 日本語／英語インターフェース対応

### 更新履歴

- v1.0 (20250605) : 初版。選択オブジェクトをグリッド状に配置する基本機能を追加
- v1.8.0 (20260430) : セル描画モードのラジオボタン化、配置処理の分離、透明グリッド復元、プレビュー確定処理を整理
---

### Script Name:

SmartObjectDistributor.jsx

### Overview

- Distributes selected objects into a grid based on specified rows and columns in Adobe Illustrator.
- The target area can be selected from the current artboard, the backmost object, or a rectangle in the "_target" layer.
- Cell drawing mode (Keep / Guides / Artboards) is handled via radio buttons with mutually exclusive behavior.
- Objects exceeding the number of cells are moved outside the artboard and preserved as-is from preview on confirmation.
- PreviewManager keeps the undo history clean during preview and commits changes as a single undoable step.
- Transparency grid toggling is tracked and restored to its original state when the script finishes.

### Main Features

- Grid division settings for rows, columns, margins, and gutters
- Target selection from the current artboard, backmost object, or rectangle in the "_target" layer
- Cell background drawing with guide or artboard conversion
- Center placement and random placement
- Instant preview with controlled Undo history
- Japanese and English UI support

### Update History

- v1.0 (20250605) : Initial release with basic grid distribution for selected objects
- v1.8.0 (20260430) : Refined cell drawing modes as radio buttons, separated placement from cell drawing, restored transparency grid state, and improved preview confirmation behavior

*/

(function () {

    var SCRIPT_VERSION = "v1.8.0";

    // 言語判定関数とラベル定義（グローバル）
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var currentLanguage = getCurrentLang();

    var labels = {
        dialogTitle: {
            ja: "グリッド状に配置",
            en: "Distribute in Grid"
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
        transparencyGridLabel: {
            ja: "透明グリッド",
            en: "Transparency Grid"
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
        },
        targetPanelLabel: {
            ja: "対象",
            en: "Target"
        },
        targetCurrentArtboard: {
            ja: "現在のアートボード",
            en: "Current Artboard"
        },
        targetBackmostObject: {
            ja: "最背面のオブジェクト",
            en: "Backmost Object"
        },
        targetRectangleLayer: {
            ja: "「_target」レイヤーの長方形",
            en: "Rectangle in '_target' Layer"
        }
    };

    // 汎用 Undo/Preview 管理クラス (PreviewManager)
    // - プレビュー更新のたびに rollback → addStep で「履歴を汚さない」
    // - OK時は confirm(finalAction) で「1回のUndoで戻せる」状態にする
    function PreviewManager(doc, historyName) {
        this.doc = doc;
        this.historyName = historyName || "Preview";
        this.undoDepth = 0; // プレビュー中に実行されたアクションの回数

        /**
         * 変更操作を実行し、履歴としてカウントする
         * @param {Function} func - 実行したい処理（無名関数で渡す）
         */
        this.addStep = function (func) {
            try {
                // 可能なら1ステップにまとめる（環境によっては使えない場合がある）
                if (this.doc && this.doc.suspendHistory) {
                    var previewAction = func;
                    PreviewManager._tempFunc = function () { previewAction(); };
                    this.doc.suspendHistory(this.historyName, "PreviewManager._tempFunc()");
                    PreviewManager._tempFunc = null;
                } else {
                    func();
                }
                this.undoDepth++;
                app.redraw();
            } catch (previewError) {
                alert("Preview Error: " + previewError);
            }
        };

        /**
         * プレビューのために行った変更を全て取り消す（キャンセル時など）
         */
        this.rollback = function () {
            while (this.undoDepth > 0) {
                try { app.undo(); } catch (undoError) { break; }
                this.undoDepth--;
            }
            app.redraw();
        };

        /**
         * 現在の状態を確定する（OK時）
         * @param {Function} [finalAction] - (任意) 全てUndoした後に実行する「本番」の処理
         */
        this.confirm = function (finalAction) {
            if (finalAction) {
                this.rollback();

                // 最終処理も可能なら1ステップにまとめる
                try {
                    if (this.doc && this.doc.suspendHistory) {
                        PreviewManager._tempFunc = function () { finalAction(); };
                        this.doc.suspendHistory("SmartObjectDistributor", "PreviewManager._tempFunc()");
                        PreviewManager._tempFunc = null;
                    } else {
                        finalAction();
                    }
                } catch (historyError) {
                    // suspendHistory が失敗した場合は通常実行にフォールバック
                    try { finalAction(); } catch (finalActionError) { throw finalActionError; }
                }
            } else {
                this.undoDepth = 0;
            }
        };
    }
    PreviewManager._tempFunc = null;

    // 選択オブジェクトがアートボードに重なっている場合、他と重ならない位置へ一時的に退避
    function moveObjectsOutsideArtboard(doc, selection) {
        var activeArtboardIndex = doc.artboards.getActiveArtboardIndex();
        var activeArtboardRect = doc.artboards[activeArtboardIndex].artboardRect;
        var artboardLeft = activeArtboardRect[0];
        var artboardTop = activeArtboardRect[1];
        var artboardRight = activeArtboardRect[2];
        var artboardBottom = activeArtboardRect[3];

        var movedObjects = [];
        var BOUNDARY_MARGIN = 1;
        var ARTBOARD_BUFFER = 10;

        function isOverlappingWithMovedObjects(newBounds) {
            for (var i = 0; i < movedObjects.length; i++) {
                var movedBounds = movedObjects[i];
                if (!(newBounds[2] < movedBounds[0] - BOUNDARY_MARGIN || newBounds[0] > movedBounds[2] + BOUNDARY_MARGIN ||
                    newBounds[1] < movedBounds[3] - BOUNDARY_MARGIN || newBounds[3] > movedBounds[1] + BOUNDARY_MARGIN)) {
                    return true;
                }
            }
            return false;
        }

        function isOverlappingWithArtboards(newBounds) {
            for (var i = 0; i < doc.artboards.length; i++) {
                if (i === activeArtboardIndex) continue;
                var otherArtboardRect = doc.artboards[i].artboardRect;
                if (!(newBounds[2] < otherArtboardRect[0] - BOUNDARY_MARGIN - ARTBOARD_BUFFER ||
                    newBounds[0] > otherArtboardRect[2] + BOUNDARY_MARGIN + ARTBOARD_BUFFER ||
                    newBounds[1] < otherArtboardRect[3] - BOUNDARY_MARGIN - ARTBOARD_BUFFER ||
                    newBounds[3] > otherArtboardRect[1] + BOUNDARY_MARGIN + ARTBOARD_BUFFER)) {
                    return true;
                }
            }
            return false;
        }

        for (var i = 0; i < selection.length; i++) {
            var item = selection[i];
            var bounds = item.visibleBounds;
            var overlapsActiveArtboard = !(bounds[2] < artboardLeft || bounds[0] > artboardRight || bounds[1] < artboardBottom || bounds[3] > artboardTop);

            if (overlapsActiveArtboard) {
                var horizontalOffset = (artboardRight - bounds[0]) + 200;
                do {
                    var newBounds = [bounds[0] + horizontalOffset, bounds[1], bounds[2] + horizontalOffset, bounds[3]];
                    horizontalOffset += 200;
                } while (isOverlappingWithMovedObjects(newBounds) || isOverlappingWithArtboards(newBounds));

                item.position = [item.position[0] + horizontalOffset - 200, item.position[1]];
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
            var horizontalOffset = originalPositions[i][0] - itemCenterX;
            var verticalOffset = originalPositions[i][1] - itemCenterY;
            items[i].translate(horizontalOffset, verticalOffset);
        }
    }

    // 「_target」レイヤー内の最初の PathItem を返す（無ければ null）
    function findTargetLayerRectangle() {
        if (app.documents.length === 0) return null;
        var doc = app.activeDocument;

        for (var layerIndex = 0; layerIndex < doc.layers.length; layerIndex++) {
            var layer = doc.layers[layerIndex];
            if (layer.name !== "_target") continue;

            for (var itemIndex = 0; itemIndex < layer.pathItems.length; itemIndex++) {
                var item = layer.pathItems[itemIndex];
                if (item.typename === "PathItem") return item;
            }
            break;
        }
        return null;
    }

    // システムレイヤーを除外したドキュメント最背面の pageItem を返す（無ければ null）
    // 選択中のアイテムも候補に含める：枠として使う場合の利便性を優先
    function findBackmostPageItem() {
        if (app.documents.length === 0) return null;
        var doc = app.activeDocument;
        var SYSTEM_LAYERS = {
            "_target": 1,
            "_Preview_Guides": 1,
            "_Preview_Background": 1,
            "cell-background": 1,
            "placement_layer": 1
        };
        for (var layerIndex = doc.layers.length - 1; layerIndex >= 0; layerIndex--) {
            var layer = doc.layers[layerIndex];
            try {
                if (layer.visible === false || layer.locked) continue;
            } catch (layerAccessError) { continue; }
            if (SYSTEM_LAYERS[layer.name]) continue;
            for (var i = layer.pageItems.length - 1; i >= 0; i--) {
                var item = layer.pageItems[i];
                try {
                    if (item.hidden || item.locked) continue;
                } catch (pageItemAccessError) { continue; }
                return item;
            }
        }
        return null;
    }

    // 選択数とアートボードのアスペクト比から、セルが正方形に近くなる行数・列数を計算
    // 余白セルが行/列1本分以上になる構成は除外し、その範囲内でセルアスペクト比1に最も近いものを選ぶ
    function computeDefaultGridDivision(count, artboardRect) {
        if (!count || count <= 0) return { rows: 5, cols: 5 };
        if (count === 1) return { rows: 1, cols: 1 };

        var width = artboardRect[2] - artboardRect[0];
        var height = artboardRect[1] - artboardRect[3];
        if (width <= 0 || height <= 0) return { rows: 5, cols: 5 };

        var bestRowCount = 1;
        var bestColumnCount = count;
        var bestAspectRatio = Infinity;
        for (var columnCount = 1; columnCount <= count; columnCount++) {
            var rowCount = Math.ceil(count / columnCount);
            var emptyCellCount = rowCount * columnCount - count;
            if (emptyCellCount > 0 && emptyCellCount >= Math.min(rowCount, columnCount)) continue;
            var cellWidth = width / columnCount;
            var cellHeight = height / rowCount;
            var aspectRatio = (cellWidth > cellHeight) ? (cellWidth / cellHeight) : (cellHeight / cellWidth);
            if (aspectRatio < bestAspectRatio) {
                bestAspectRatio = aspectRatio;
                bestRowCount = rowCount;
                bestColumnCount = columnCount;
            }
        }
        return { rows: bestRowCount, cols: bestColumnCount };
    }

    function run() {
        if (app.documents.length === 0) {
            alert("ドキュメントを開いてください。\nPlease open a document.");
            return;
        }

        // "cell-background" レイヤーが存在する場合は削除
        try {
            var existingLayer = app.activeDocument.layers.getByName("cell-background");
            if (existingLayer) existingLayer.remove();
        } catch (existingLayerError) {
            // 存在しない場合はエラーを無視
        }

        var doc = app.activeDocument;

        // 以後の計算は「ダイアログ起動時点のアクティブアートボード」を基準に固定
        var baseArtboardIndex = doc.artboards.getActiveArtboardIndex();
        var baseArtboardRect = doc.artboards[baseArtboardIndex].artboardRect;

        // 対象候補の検出（変更は加えない）
        var targetRectangleItem = findTargetLayerRectangle();
        var backmostPageItem = findBackmostPageItem();

        // 初期ターゲット矩形（既存挙動を踏襲：_target 矩形があればそれ、なければ現在のアートボード）
        var initialTargetRect;
        if (targetRectangleItem) {
            try { initialTargetRect = targetRectangleItem.geometricBounds; } catch (targetRectangleBoundsError) { }
        }
        if (!initialTargetRect) initialTargetRect = baseArtboardRect;

        // _target レイヤーの矩形は対象として選ばれている間は非表示にする
        if (targetRectangleItem) {
            try { targetRectangleItem.hidden = true; } catch (targetRectangleHideError) { }
        }

        // PreviewManager
        var previewMgr = new PreviewManager(doc, "SmartObjectDistributor Preview");

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
        var dlg = new Window("dialog", labels.dialogTitle[currentLanguage] + " " + SCRIPT_VERSION);
        dlg.orientation = "column";
        dlg.alignChildren = "fill";

        // === 対象パネル ===
        var targetPanel = dlg.add("panel", undefined, labels.targetPanelLabel[currentLanguage]);
        targetPanel.orientation = "column";
        targetPanel.alignChildren = "left";
        targetPanel.alignment = "fill";
        targetPanel.margins = [15, 20, 15, 15];

        var rbCurrentArtboard = targetPanel.add("radiobutton", undefined, labels.targetCurrentArtboard[currentLanguage]);
        var rbBackmostObject = targetPanel.add("radiobutton", undefined, labels.targetBackmostObject[currentLanguage]);
        var rbTargetRectLayer = targetPanel.add("radiobutton", undefined, labels.targetRectangleLayer[currentLanguage]);

        rbBackmostObject.enabled = (backmostPageItem !== null);
        rbTargetRectLayer.enabled = (targetRectangleItem !== null);

        if (rbTargetRectLayer.enabled) {
            rbTargetRectLayer.value = true;
        } else {
            rbCurrentArtboard.value = true;
        }

        // 選ばれている対象の矩形（[left, top, right, bottom]）を返す
        function computeTargetRect() {
            if (rbBackmostObject.value && backmostPageItem) {
                try { return backmostPageItem.geometricBounds; } catch (backmostBoundsError) { }
            }
            if (rbTargetRectLayer.value && targetRectangleItem) {
                try { return targetRectangleItem.geometricBounds; } catch (targetRectangleBoundsError) { }
            }
            if (doc.artboards.length > baseArtboardIndex) {
                return doc.artboards[baseArtboardIndex].artboardRect;
            }
            return baseArtboardRect;
        }

        // グリッド設定パネル（2列構成、タイトル付き）
        var gridPanel = dlg.add("panel", undefined, labels.gridSettingsLabel[currentLanguage]);
        gridPanel.orientation = "row";
        gridPanel.alignChildren = "top";
        gridPanel.margins = [15, 20, 15, 15];
        gridPanel.spacing = 20;

        // ラベル幅を日本語環境のときに 40/55、それ以外は 55/70 に個別指定
        var leftLabelWidth = (currentLanguage === 'ja') ? 40 : 55;
        var rightLabelWidth = (currentLanguage === 'ja') ? 65 : 70;

        // 2列グループ（gridPanel直下へ追加）
        var leftColGroup = gridPanel.add("group");
        leftColGroup.orientation = "column";
        leftColGroup.alignChildren = "left";
        leftColGroup.margins.right = 20; // 右列とのスペース確保

        var rightColGroup = gridPanel.add("group");
        rightColGroup.orientation = "column";
        rightColGroup.alignChildren = "left";

        // 選択オブジェクト数に応じて行・列の初期値を計算（セルが正方形に近くなるように）
        var initialSelectionCount = 0;
        try {
            if (doc.selection && doc.selection.length) initialSelectionCount = doc.selection.length;
        } catch (eSel) { }
        var defaultGrid = computeDefaultGridDivision(initialSelectionCount, initialTargetRect);

        // 左側：行数・列数（ペアで整える）
        var inputY = leftColGroup.add("group");
        inputY.add("statictext", undefined, labels.rowsLabel[currentLanguage]).preferredSize.width = leftLabelWidth;
        var inputYText = inputY.add("edittext", undefined, String(defaultGrid.rows));
        inputYText.characters = 3;

        var inputX = leftColGroup.add("group");
        inputX.add("statictext", undefined, labels.columnsLabel[currentLanguage]).preferredSize.width = leftLabelWidth;
        var inputXText = inputX.add("edittext", undefined, String(defaultGrid.cols));
        inputXText.characters = 3;

        // 右側：行間・マージン（ペアで整える）
        var rowGutterGroup = rightColGroup.add("group");
        rowGutterGroup.alignChildren = "center";
        var rowGutterLabel = rowGutterGroup.add("statictext", undefined, labels.rowGutterLabel[currentLanguage]);
        rowGutterLabel.preferredSize.width = rightLabelWidth;
        rowGutterLabel.justify = "right";
        var inputRowGutter = rowGutterGroup.add("edittext", undefined, "10");
        inputRowGutter.characters = 4;
        rowGutterGroup.add("statictext", undefined, unitLabel);

        var marginGroup = rightColGroup.add("group");
        marginGroup.orientation = "row";
        marginGroup.alignChildren = "center";
        var marginLabel = marginGroup.add("statictext", undefined, labels.commonMarginLabel[currentLanguage]);
        marginLabel.preferredSize.width = rightLabelWidth;
        marginLabel.justify = "right";
        var marginInput = marginGroup.add("edittext", undefined, "10");
        marginInput.characters = 4;
        marginGroup.add("statictext", undefined, unitLabel);

        // オプション設定グループ
        var optGroup = dlg.add("group");
        optGroup.orientation = "column";
        optGroup.alignChildren = "fill";
        optGroup.alignment = "fill";

        // === 長方形オプションパネル ===
        var rectPanel = optGroup.add("panel", undefined, labels.cellDrawingLabel[currentLanguage]);
        rectPanel.orientation = "column";
        rectPanel.alignChildren = "left";
        rectPanel.alignment = "fill";
        rectPanel.margins = [15, 20, 15, 15];

        // 横並びのグループにラジオボタンをまとめる
        var rectOptionsGroup = rectPanel.add("group");
        rectOptionsGroup.orientation = "row";
        rectOptionsGroup.alignChildren = "left";

        var cellRectRadio = rectOptionsGroup.add("radiobutton", undefined, labels.cellRectLabel[currentLanguage]);
        var convertToGuideRadio = rectOptionsGroup.add("radiobutton", undefined, labels.convertToGuideLabel[currentLanguage]);
        var convertToArtboardRadio = rectOptionsGroup.add("radiobutton", undefined, labels.convertToArtboardLabel[currentLanguage]);

        cellRectRadio.value = true;

        // カラー選択ラジオ（黒・白・透過）
        var colorGroup = rectPanel.add("group");
        colorGroup.orientation = "row";
        colorGroup.alignChildren = "left";

        colorGroup.add("statictext", undefined, labels.colorLabel[currentLanguage]);

        var rbBlack = colorGroup.add("radiobutton", undefined, labels.black[currentLanguage]);
        var rbWhite = colorGroup.add("radiobutton", undefined, labels.white[currentLanguage]);
        var rbNone = colorGroup.add("radiobutton", undefined, labels.none[currentLanguage]);

        // 不透明度設定行を追加
        var opacityGroup = rectPanel.add("group");
        opacityGroup.orientation = "row";
        opacityGroup.alignChildren = "center";

        opacityGroup.add("statictext", undefined, labels.opacityLabel[currentLanguage]);

        var inputOpacity = opacityGroup.add("edittext", undefined, "15");
        inputOpacity.characters = 4;
        opacityGroup.add("statictext", undefined, "%");

        var transparencyGridToggleCount = 0;
        var transparencyGridButton = opacityGroup.add("button", undefined, labels.transparencyGridLabel[currentLanguage]);
        transparencyGridButton.onClick = function () {
            app.executeMenuCommand('TransparencyGrid Menu Item');
            transparencyGridToggleCount++;
        };

        rbBlack.value = true; // 初期選択は黒

        // プレビュー更新（Undo汚染防止）
        function updatePreviewFromCurrentSettings(convertToGuideOverride) {
            var convertToGuide = (convertToGuideOverride !== undefined) ? convertToGuideOverride : convertToGuideRadio.value;
            previewMgr.rollback();
            previewMgr.addStep(function () {
                renderDistributionGrid(true, true, convertToGuide);
            });
        }

        // オプション切り替え時のUI状態反映（ラジオボタン用に再実装）
        function syncCellDrawingOptions() {
            if (convertToArtboardRadio.value) {
                applyArtboardCellDrawingUIState();
                updatePreviewFromCurrentSettings(true);
            } else if (convertToGuideRadio.value) {
                applyGuideCellDrawingUIState();
                updatePreviewFromCurrentSettings(true);
            } else {
                applyNormalCellDrawingUIState();
                updatePreviewFromCurrentSettings(false);
            }
        }

        // サブ関数：アートボード化時のUI
        function applyArtboardCellDrawingUIState() {
            cellRectRadio.enabled = true;
            convertToGuideRadio.enabled = true;
            convertToArtboardRadio.enabled = true;
            cellRectRadio.value = false;
            convertToGuideRadio.value = false;
            convertToArtboardRadio.value = true;

            rbBlack.enabled = false;
            rbWhite.enabled = false;
            rbNone.enabled = false;
            rbNone.value = true;
            inputOpacity.enabled = false;
        }

        // サブ関数：ガイド化時のUI
        function applyGuideCellDrawingUIState() {
            cellRectRadio.enabled = true;
            convertToGuideRadio.enabled = true;
            convertToArtboardRadio.enabled = true;
            cellRectRadio.value = false;
            convertToGuideRadio.value = true;
            convertToArtboardRadio.value = false;

            rbBlack.enabled = false;
            rbWhite.enabled = false;
            rbNone.enabled = false;
            rbNone.value = true;
            inputOpacity.enabled = false;
            syncColorOpacityUI(); // 内部でプレビュー更新
        }

        // サブ関数：通常時のUI
        function applyNormalCellDrawingUIState() {
            cellRectRadio.enabled = true;
            convertToGuideRadio.enabled = true;
            convertToArtboardRadio.enabled = true;
            cellRectRadio.value = true;
            convertToGuideRadio.value = false;
            convertToArtboardRadio.value = false;

            rbBlack.enabled = true;
            rbWhite.enabled = true;
            rbNone.enabled = true;
            rbBlack.value = true;
            inputOpacity.enabled = rbBlack.value || rbWhite.value;
        }

        // 対象ラジオの切り替え：_target 矩形は選択中だけ非表示にし、プレビュー再描画
        function handleTargetSelectionChange() {
            if (targetRectangleItem) {
                try { targetRectangleItem.hidden = (rbTargetRectLayer.value === true); } catch (targetRectangleVisibilityError) { }
            }
            updatePreviewFromCurrentSettings();
        }
        rbCurrentArtboard.onClick = handleTargetSelectionChange;
        rbBackmostObject.onClick = handleTargetSelectionChange;
        rbTargetRectLayer.onClick = handleTargetSelectionChange;

        cellRectRadio.onClick = syncCellDrawingOptions;
        convertToGuideRadio.onClick = syncCellDrawingOptions;
        convertToArtboardRadio.onClick = syncCellDrawingOptions;

        // 不透明度欄：黒・白のとき有効、透過時は無効
        function syncColorOpacityUI() {
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
                updatePreviewFromCurrentSettings(convertToGuideRadio.value);
            }
        }
        rbBlack.onClick = rbWhite.onClick = rbNone.onClick = syncColorOpacityUI;

        // 初期状態反映
        syncColorOpacityUI();

        // 不透明度変更時のみプレビュー更新
        var lastOpacityValue = inputOpacity.text;
        inputOpacity.onChanging = function () {
            if (inputOpacity.text !== lastOpacityValue) {
                lastOpacityValue = inputOpacity.text;
                updatePreviewFromCurrentSettings(convertToGuideRadio.value);
            }
        };

        // === ボタンエリア（レイアウト変更版）===
        var outerGroup = dlg.add("group");
        outerGroup.orientation = "row";
        outerGroup.alignChildren = ["fill", "center"];
        outerGroup.margins = [0, 10, 0, 0];
        outerGroup.spacing = 0;

        // --- 左グループ（ランダム） ---
        var leftGroup = outerGroup.add("group");
        leftGroup.orientation = "row";
        leftGroup.alignChildren = "left";

        var btnRandom = leftGroup.add("button", undefined, labels.randomLabel[currentLanguage]);

        // スペーサー（横に伸びる空白）
        var spacer = outerGroup.add("group");
        spacer.alignment = ["fill", "fill"];
        spacer.minimumSize.width = (currentLanguage === 'ja') ? 80 : 120;
        spacer.maximumSize.height = 0;

        // 右グループ（Cancel/OK）
        var rightGroup = outerGroup.add("group");
        rightGroup.orientation = "row";
        rightGroup.alignChildren = "right";
        rightGroup.spacing = 10;

        var btnCancel = rightGroup.add("button", undefined, labels.cancelLabel[currentLanguage], { name: "cancel" });
        var btnOK = rightGroup.add("button", undefined, labels.okLabel[currentLanguage], { name: "ok" });

        // 行数・列数に応じてガター入力欄の有効/無効を切り替える
        function syncGutterInputEnabled() {
            var yVal = parseInt(inputYText.text, 10);
            inputRowGutter.enabled = (yVal > 1);
        }

        inputXText.onChanging = inputYText.onChanging = function () {
            syncGutterInputEnabled();
            updatePreviewFromCurrentSettings();
        };

        inputRowGutter.onChanging = function () {
            updatePreviewFromCurrentSettings();
        };
        marginInput.onChanging = function () {
            updatePreviewFromCurrentSettings();
        };


        // 選択状態を退避／復元（Undo rollback 後に selection が空になるケース対策）
        function getSelectionSnapshot() {
            var selectedItems = [];
            var currentSelection = doc.selection;
            if (currentSelection && currentSelection.length) {
                for (var i = 0; i < currentSelection.length; i++) selectedItems.push(currentSelection[i]);
            }
            return selectedItems;
        }

        function restoreSelectionSnapshot(snapshot) {
            if (!snapshot || snapshot.length === 0) return;
            try {
                doc.selection = snapshot;
            } catch (selectionRestoreError) { }
        }
        // グローバル変数：ランダム順序保持用
        var randomizedSelection = null;

        // 配置対象アイテムを返す（「最背面のオブジェクト」が対象のときは枠アイテムを除外）
        function getDistributionItems() {
            var distributionItems = [];
            if (!doc.selection || !doc.selection.length) return distributionItems;

            for (var i = 0; i < doc.selection.length; i++) {
                var item = doc.selection[i];
                if (rbBackmostObject.value && backmostPageItem && item === backmostPageItem) continue;
                distributionItems.push(item);
            }
            return distributionItems;
        }

        // ランダムボタン押下時
        btnRandom.onClick = function () {
            var selectedItems = getDistributionItems();
            if (!selectedItems || selectedItems.length === 0) {
                alert(labels.noSelection[currentLanguage]);
                return;
            }
            randomizedSelection = [];
            for (var i = 0; i < selectedItems.length; i++) {
                randomizedSelection.push(selectedItems[i]);
            }
            // Fisher-Yates シャッフルで順序をランダム化
            for (var k = randomizedSelection.length - 1; k > 0; k--) {
                var randomIndex = Math.floor(Math.random() * (k + 1));
                var temporaryItem = randomizedSelection[k];
                randomizedSelection[k] = randomizedSelection[randomIndex];
                randomizedSelection[randomIndex] = temporaryItem;
            }
            updatePreviewFromCurrentSettings(convertToGuideRadio.value);
            app.redraw();
        };

        // OKボタン押下時（確定を1アクション化）
        btnOK.onClick = function () {
            syncGutterInputEnabled();
            var selectionSnapshot = getSelectionSnapshot();

            previewMgr.confirm(function () {

                // app.executeMenuCommand("deselectall");

                // プレビュー用レイヤーを削除（Undoで消えないケースの保険）
                removePreviewLayers();

                // rollback後に選択が外れてしまう環境向けに復元
                restoreSelectionSnapshot(selectionSnapshot);

                // アートボード化
                if (convertToArtboardRadio.value) {
                    try {
                        var geomForArtboard = computeDistributionGridGeometry();
                        var createdCount = createArtboardsFromGridGeometry(geomForArtboard);
                        alert(createdCount + labels.artboardCreated[currentLanguage]);
                    } catch (artboardCreationError) {
                        alert(labels.artboardError[currentLanguage]);
                    }
                }
                // 念のため：本番処理前に基準アートボードへ戻す
                try {
                    if (doc.artboards.setActiveArtboardIndex) {
                        doc.artboards.setActiveArtboardIndex(baseArtboardIndex);
                    }
                } catch (activeArtboardRestoreError) { }

                // 本番描画・配置（プレビュー時と同じ退避・配置状態で確定する）
                var keepRects = !convertToArtboardRadio.value;
                renderDistributionGrid(false, keepRects, convertToGuideRadio.value);

                randomizedSelection = null;
                originalPositions = [];

                // === OK時にcell-backgroundレイヤー内の図形を選択状態にする ===
                try {
                    var bgLayerForSelect = doc.layers.getByName("cell-background");
                    var bgItems = [];
                    for (var ii = 0; ii < bgLayerForSelect.pathItems.length; ii++) {
                        var backgroundPathItem = bgLayerForSelect.pathItems[ii];
                        if (backgroundPathItem.typename === "PathItem") {
                            bgItems.push(backgroundPathItem);
                        }
                    }
                    doc.selection = bgItems;
                } catch (backgroundSelectionError) {
                    // レイヤーが見つからない場合は無視
                }

                // 「_target」レイヤーの矩形を再表示（プレビュー時に隠していた場合）
                if (targetRectangleItem) {
                    try { targetRectangleItem.hidden = false; } catch (targetRectangleUnhideError) { }
                }
            });

            dlg.close(1);
        };

        // キャンセルボタン押下時
        btnCancel.onClick = function () {
            previewMgr.rollback();
            removePreviewLayers(); // 念のため（Undoで消えないケース対策）
            randomizedSelection = null;
            originalPositions = [];
            if (targetRectangleItem) {
                try { targetRectangleItem.hidden = false; } catch (targetRectangleUnhideError) { }
            }
            dlg.close(0);
        };

        // 長方形描画処理（背景長方形は専用レイヤーに描画する）
        function drawCellRectangles(cellLayer, baseLeft, baseTop, cellWidth, cellHeight, columnCount, rowCount, colGutter, rowGutter, convertToGuide, isPreview) {
            var opacityValue = parseFloat(inputOpacity.text);
            for (var row = 0; row < rowCount; row++) {
                var cellY = baseTop - (cellHeight + rowGutter) * row;
                for (var col = 0; col < columnCount; col++) {
                    var cellX = baseLeft + (cellWidth + colGutter) * col;
                    var cellRectangle = cellLayer.pathItems.rectangle(cellY, cellX, cellWidth, cellHeight);
                    cellRectangle.stroked = false;

                    // カラーラジオボタンの値による分岐
                    if (rbBlack.value) {
                        cellRectangle.filled = true;
                        cellRectangle.fillColor = createBlackColor();
                    } else if (rbWhite.value) {
                        cellRectangle.filled = true;
                        cellRectangle.fillColor = createWhiteColor();
                    } else {
                        cellRectangle.filled = false;
                    }

                    if (!isNaN(opacityValue)) {
                        cellRectangle.opacity = opacityValue;
                    }

                    if (convertToGuide) {
                        cellRectangle.guides = true;
                    }
                }
            }

            // 背景長方形レイヤーを最背面に移動
            if ((isPreview && cellLayer.name === "_Preview_Guides") || (!isPreview && cellLayer.name === "cell-rectangle")) {
                cellLayer.zOrder(ZOrderMethod.SENDTOBACK);
            }
        }

        // 各長方形に1つずつオブジェクトを配置
        function distributeSelectionToCells(selectedItems, baseLeft, baseTop, cellWidth, cellHeight, columnCount, rowCount, colGutter, rowGutter) {
            if (!selectedItems || selectedItems.length === 0) return;
            var items = randomizedSelection || selectedItems;
            var index = 0;
            for (var row = 0; row < rowCount; row++) {
                var cellY = baseTop - (cellHeight + rowGutter) * row;
                for (var col = 0; col < columnCount; col++) {
                    if (index >= items.length) return;

                    var cellX = baseLeft + (cellWidth + colGutter) * col;
                    var cellCenterX = cellX + cellWidth / 2;
                    var cellCenterY = cellY - cellHeight / 2;

                    var item = items[index];
                    var bounds = item.visibleBounds;
                    var itemCenterX = (bounds[0] + bounds[2]) / 2;
                    var itemCenterY = (bounds[1] + bounds[3]) / 2;
                    var horizontalOffset = cellCenterX - itemCenterX;
                    var verticalOffset = cellCenterY - itemCenterY;
                    item.translate(horizontalOffset, verticalOffset);

                    index++;
                }
            }
        }

        // グリッド計算結果をまとめて返す（プレビュー/本番で共通利用）
        function computeDistributionGridGeometry() {
            var columnCount = parseInt(inputXText.text, 10);
            var rowCount = parseInt(inputYText.text, 10);
            if (isNaN(columnCount) || columnCount <= 0 || isNaN(rowCount) || rowCount <= 0) return null;

            var marginVal = parseFloat(marginInput.text) * unitFactor;
            var top = marginVal;
            var bottom = marginVal;
            var left = marginVal;
            var right = marginVal;

            var rowGutter = parseFloat(inputRowGutter.text) * unitFactor;
            var colGutter = rowGutter;

            // 対象パネルの選択に応じた矩形を基準にする
            var targetBounds = computeTargetRect();
            var targetLeft = targetBounds[0];
            var targetTop = targetBounds[1];
            var targetRight = targetBounds[2];
            var targetBottom = targetBounds[3];

            var baseLeft = targetLeft + left;
            var baseRight = targetRight - right;
            var baseTop = targetTop - top;
            var baseBottom = targetBottom + bottom;

            var usableWidth = baseRight - baseLeft;
            var usableHeight = baseTop - baseBottom;
            var totalColumnGutter = (columnCount - 1) * colGutter;
            var totalRowGutter = (rowCount - 1) * rowGutter;
            var cellWidth = (usableWidth - totalColumnGutter) / columnCount;
            var cellHeight = (usableHeight - totalRowGutter) / rowCount;

            return {
                columnCount: columnCount,
                rowCount: rowCount,
                baseLeft: baseLeft,
                baseTop: baseTop,
                cellWidth: cellWidth,
                cellHeight: cellHeight,
                colGutter: colGutter,
                rowGutter: rowGutter
            };
        }

        // セルの幾何情報からアートボードを作成
        function createArtboardsFromGridGeometry(geom) {

            var previousArtboardIndex = -1;
            try { previousArtboardIndex = doc.artboards.getActiveArtboardIndex(); } catch (activeArtboardReadError) { }

            if (!geom) return 0;
            var createdCount = 0;
            for (var row = 0; row < geom.rowCount; row++) {
                var cellY = geom.baseTop - (geom.cellHeight + geom.rowGutter) * row;
                for (var col = 0; col < geom.columnCount; col++) {
                    var cellX = geom.baseLeft + (geom.cellWidth + geom.colGutter) * col;
                    var left = cellX;
                    var top = cellY;
                    var right = cellX + geom.cellWidth;
                    var bottom = cellY - geom.cellHeight;
                    if (isNaN(left) || isNaN(top) || isNaN(right) || isNaN(bottom) || right <= left || top <= bottom) continue;
                    try {
                        doc.artboards.add([left, top, right, bottom]);
                        createdCount++;
                    } catch (artboardAddError) {
                        // 個別エラーは握りつぶして続行
                    }
                }
            }
            // アクティブアートボードを元に戻す（最後の追加ABがアクティブになるのを防ぐ）
            try {
                if (previousArtboardIndex >= 0 && doc.artboards.setActiveArtboardIndex) {
                    doc.artboards.setActiveArtboardIndex(previousArtboardIndex);
                }
            } catch (activeArtboardRestoreError) { }
            return createdCount;
        }

        // ガイド描画／長方形描画／オブジェクト配置を制御
        // 背景長方形は専用レイヤーに描画する
        function renderDistributionGrid(isPreview, keepRects, convertToGuide) {
            var distributionItems = getDistributionItems();

            // プレビュー／確定ともに、いったん元位置へ戻してから対象セルに配置する
            // セル数を超えるオブジェクトは、プレビューと同じくアートボード外へ退避した状態で確定する
            if (doc.selection.length > 0) {
                if (originalPositions.length === 0) saveOriginalPositions(doc.selection);
                restoreOriginalPositions(doc.selection);
                moveObjectsOutsideArtboard(doc, distributionItems);
            }

            if (isPreview) {
                removePreviewLayers();
            }

            var geom = computeDistributionGridGeometry();
            if (!geom) return;

            var columnCount = geom.columnCount;
            var rowCount = geom.rowCount;
            var baseLeft = geom.baseLeft;
            var baseTop = geom.baseTop;
            var cellWidth = geom.cellWidth;
            var cellHeight = geom.cellHeight;
            var colGutter = geom.colGutter;
            var rowGutter = geom.rowGutter;

            var shouldDrawCellRectangles = cellRectRadio.value || convertToGuideRadio.value || (convertToArtboardRadio.value && isPreview);
            var shouldKeepCellRectangles = cellRectRadio.value;

            var gridLayerName = isPreview ? "_Preview_Guides" : "placement_layer";
            var gridLayer;
            try {
                gridLayer = doc.layers.getByName(gridLayerName);
            } catch (gridLayerError) {
                gridLayer = doc.layers.add();
                gridLayer.name = gridLayerName;
            }
            gridLayer.locked = false;

            var bgLayerName = isPreview ? "_Preview_Background" : "cell-background";
            var cellLayer;
            try {
                cellLayer = doc.layers.getByName(bgLayerName);
            } catch (cellLayerError) {
                cellLayer = doc.layers.add();
                cellLayer.name = bgLayerName;
            }
            cellLayer.locked = false;

            if (shouldDrawCellRectangles && cellLayer) {
                drawCellRectangles(cellLayer, baseLeft, baseTop, cellWidth, cellHeight, columnCount, rowCount, colGutter, rowGutter, convertToGuide, isPreview);

                if (isPreview) {
                    try {
                        var previewGuideLayer = doc.layers.getByName("_Preview_Guides");
                        var previewBackgroundLayer = doc.layers.getByName("_Preview_Background");
                        previewGuideLayer.zOrder(ZOrderMethod.BRINGTOFRONT);
                        previewBackgroundLayer.zOrder(ZOrderMethod.SENDTOBACK);
                    } catch (previewLayerOrderError) { }
                }
            }

            distributeSelectionToCells(distributionItems, baseLeft, baseTop, cellWidth, cellHeight, columnCount, rowCount, colGutter, rowGutter);

            // OK実行時：アートボード化では背景長方形レイヤーを削除
            if (!isPreview && !shouldKeepCellRectangles && !keepRects) {
                try {
                    var cellLayerToRemove = doc.layers.getByName("cell-background");
                    if (cellLayerToRemove) cellLayerToRemove.remove();
                } catch (cellLayerRemovalError) { }
            }

            // OK実行時の後処理
            if (!isPreview) {
                try {
                    var placementLayer = doc.layers.getByName("placement_layer");
                    if (placementLayer) placementLayer.remove();
                } catch (placementLayerRemovalError) { }

                try {
                    var backgroundLayer = doc.layers.getByName("cell-background");
                    if (backgroundLayer) backgroundLayer.zOrder(ZOrderMethod.SENDTOBACK);
                } catch (backgroundLayerOrderError) { }

                try {
                    if (gridLayer) gridLayer.locked = true;
                } catch (gridLayerLockError) { }
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
        function removePreviewLayers() {
            try {
                var previewLayer = doc.layers.getByName("_Preview_Guides");
                if (previewLayer) previewLayer.remove();
            } catch (previewGuideLayerError) { }
            try {
                var bgLayer = doc.layers.getByName("_Preview_Background");
                if (bgLayer) bgLayer.remove();
            } catch (previewBackgroundLayerError) { }
        }

        // セル描画モードと入力欄の初期状態を反映
        syncGutterInputEnabled();
        applyNormalCellDrawingUIState();
        updatePreviewFromCurrentSettings(false);

        // 値を矢印キーで増減する関数
        function changeValueByArrowKey(editText) {
            editText.addEventListener("keydown", function (event) {
                var value = Number(editText.text);
                if (isNaN(value)) return;

                var keyboard = ScriptUI.environment.keyboardState;
                var increment = 1;

                if (keyboard.shiftKey) {
                    increment = 10;
                    if (event.keyName === "Up") {
                        value = Math.ceil((value + 1) / increment) * increment;
                        event.preventDefault();
                    } else if (event.keyName === "Down") {
                        value = Math.floor((value - 1) / increment) * increment;
                        if (value < 0) value = 0;
                        event.preventDefault();
                    }
                } else if (keyboard.altKey) {
                    increment = 0.1;
                    if (event.keyName === "Up") {
                        value += increment;
                        event.preventDefault();
                    } else if (event.keyName === "Down") {
                        value -= increment;
                        event.preventDefault();
                    }
                } else {
                    increment = 1;
                    if (event.keyName === "Up") {
                        value += increment;
                        event.preventDefault();
                    } else if (event.keyName === "Down") {
                        value -= increment;
                        if (value < 0) value = 0;
                        event.preventDefault();
                    }
                }

                if (keyboard.altKey) {
                    value = Math.round(value * 10) / 10;
                } else {
                    value = Math.round(value);
                }

                editText.text = value;
                updatePreviewFromCurrentSettings(convertToGuideRadio.value);
            });
        }

        // 各入力欄に適用
        changeValueByArrowKey(inputYText);
        changeValueByArrowKey(inputXText);
        changeValueByArrowKey(inputRowGutter);
        changeValueByArrowKey(marginInput);
        changeValueByArrowKey(inputOpacity);

        dlg.show();

        if (transparencyGridToggleCount % 2 !== 0) {
            app.executeMenuCommand('TransparencyGrid Menu Item');
        }
    }

    run();

}());