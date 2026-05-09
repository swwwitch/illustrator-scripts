#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
概要（日本語）

選択オブジェクトの端または中心を、アクティブアートボードの端・中央、または条件に合うガイドへ整列します。

主な機能
- 3×3ラジオボタンで整列方向を指定
- ファイル名による整列方向の自動判定に対応
  例：GroupEdgeAlignLEFT.jsx / GroupEdgeAlignRIGHT.jsx / GroupEdgeAlignCENTER.jsx などのファイル名で既定の整列方向を切り替えます。
- ダイアログでラジオボタンを選択した場合は、ラジオボタンの指定を優先
- ラジオボタン使用時は「ガイドを使用」をディム表示にし、整列処理でもガイドを無視
- 線幅込み／線幅なしの境界を切り替え可能
- プレビューと矢印キーによるステップ整列に対応
- ラジオボタンのキーボードショートカット：w/e/r, s/d/f, x/c/v

対象
- 選択されているすべてのオブジェクト

対象外
- 未選択、ロック、非表示のオブジェクト
- 指定方向に対応しないガイド、アートボード外のガイド

補足
- 座標系はIllustrator準拠（Yは上方向が大きい）
- クリッピンググループは内部要素の境界を参照

オリジナルアイデア
Gorolib Designさん
https://gorolib.blog.jp/archives/63149753.html

キーによるステップ移動のアイデア
ken @ken_rainy

作成日：2025-04-06

--------------------------------------------------

Summary (English)

Aligns the edge or center of selected objects to the active artboard edge/center, or to a matching guide.

Main features
- Choose alignment direction with a 3×3 radio button grid
- Supports automatic alignment direction detection from the script file name
  Example: GroupEdgeAlignLEFT.jsx / RIGHT / CENTER
- Radio button selection takes priority over file-name defaults
- When using radio buttons, "Use guides" is disabled and guides are ignored
- Switch between visible bounds and geometric bounds
- Supports preview and step alignment with arrow keys
- Radio button shortcuts: w/e/r, s/d/f, x/c/v

Target
- All selected objects

Excluded
- Unselected, locked, or hidden objects
- Guides that do not match the selected direction or are outside the artboard

Notes
- Coordinates follow Illustrator conventions (Y increases upward)
- Clipping groups use internal item bounds
*/

// =========================================
// バージョンとローカライズ / Version and localization
// =========================================

var SCRIPT_VERSION = "v1.0";

function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}

var lang = getCurrentLang();

/* 日英ラベル定義 / Japanese-English label definitions */
var LABELS = {
    dialogTitle: {
        ja: "アートボードに整列",
        en: "Align to Artboard"
    },
    alignmentPanelTitle: {
        ja: "整列",
        en: "Alignment"
    },
    optionPanelTitle: {
        ja: "オプション",
        en: "Options"
    },
    previewBoundsCheckbox: {
        ja: "プレビュー境界を使用",
        en: "Use preview bounds"
    },
    useGuidesCheckbox: {
        ja: "ガイドを使用",
        en: "Use guides"
    },
    previewCheckbox: {
        ja: "プレビュー",
        en: "Preview"
    },
    cancelButton: {
        ja: "キャンセル",
        en: "Cancel"
    },
    okButton: {
        ja: "OK",
        en: "OK"
    },
    noSelectionError: {
        ja: "オブジェクトが選択されていません。",
        en: "No objects are selected."
    },
    generalError: {
        ja: "エラーが発生しました",
        en: "An error occurred"
    },
    invalidGuideSearchModeError: {
        ja: "GUIDE_SEARCH_MODE の指定が不正です",
        en: "Invalid GUIDE_SEARCH_MODE"
    }
};

function getLabel(labelKey) {
    return LABELS[labelKey][lang] || LABELS[labelKey].ja;
}

/* コロン付きラベル（日本語は全角、英語は半角）/ Label with colon (full-width JA, half-width EN) */
function labelText(labelKey) {
    return getLabel(labelKey) + (lang === 'ja' ? '：' : ':');
}

function main() {
    try {
        var documentRef = app.activeDocument;
        var selectedItems = documentRef.selection;
        var USE_GUIDES = true; // ガイドを使用するかどうかのスイッチ
        // ファイル名から揃え方向を自動判定（例: GroupEdgeAlignRIGHT.jsx → "right"）
        var scriptFileName = File($.fileName).name;
        var fileNameUpper = scriptFileName.toUpperCase();

        var ALIGNMENT_SIDE = "right"; // デフォルト

        if (fileNameUpper.indexOf("CENTERX") !== -1) {
            ALIGNMENT_SIDE = "CENTER_X";
        } else if (fileNameUpper.indexOf("CENTERY") !== -1) {
            ALIGNMENT_SIDE = "CENTER_Y";
        } else if (fileNameUpper.indexOf("CENTER") !== -1) {
            ALIGNMENT_SIDE = "CENTER";
        } else if (fileNameUpper.indexOf("LEFT") !== -1) {
            ALIGNMENT_SIDE = "left";
        } else if (fileNameUpper.indexOf("RIGHT") !== -1) {
            ALIGNMENT_SIDE = "right";
        } else if (fileNameUpper.indexOf("TOP") !== -1) {
            ALIGNMENT_SIDE = "top";
        } else if (fileNameUpper.indexOf("BOTTOM") !== -1) {
            ALIGNMENT_SIDE = "bottom";
        }
        var GUIDE_SEARCH_MODE = "inside"; // "inside" | "nearest"
        var GUIDE_ORIENTATION_TOLERANCE = 0.01; // ガイドの水平・垂直判定に使う許容値
        var SHOW_DIALOG = true; // ダイアログを表示するかどうかのスイッチ

        // 選択されていない場合は処理中止
        if (selectedItems.length === 0) {
            alert(getLabel("noSelectionError"));
            return;
        }

        // プレビュー境界使用の初期値は環境設定から取得
        var includeStrokeInBounds = app.preferences.getBooleanPreference("includeStrokeInBounds");

        var artboards = documentRef.artboards;
        var activeArtboardRect = artboards[artboards.getActiveArtboardIndex()].artboardRect;

        // 選択オブジェクトの元位置を保存
        // - trueOriginalPositions: キャンセル時の完全復元用（不変）
        // - originalPositions: プレビュー復元用ベースライン（矢印キーのステップ移動で更新）
        var trueOriginalPositions = [];
        var originalPositions = [];
        for (var saveIndex = 0; saveIndex < selectedItems.length; saveIndex++) {
            var savedPosition = selectedItems[saveIndex].position;
            trueOriginalPositions.push([savedPosition[0], savedPosition[1]]);
            originalPositions.push([savedPosition[0], savedPosition[1]]);
        }

        // プレビュー適用関数（settings === null の場合は元位置への復元のみ）
        var previewAlignment = function (settings) {
            for (var restoreIndex = 0; restoreIndex < selectedItems.length; restoreIndex++) {
                selectedItems[restoreIndex].position = originalPositions[restoreIndex];
            }
            if (settings !== null) {
                var previewBounds = computeSelectionBoundsFromItems(selectedItems, settings.usePreviewBounds);
                var previewOffsets = computeAlignmentOffsets(previewBounds, settings.selectedHorizontal, settings.selectedVertical, ALIGNMENT_SIDE, settings.useGuides, activeArtboardRect, documentRef, GUIDE_SEARCH_MODE, GUIDE_ORIENTATION_TOLERANCE);
                for (var translateIndex = 0; translateIndex < selectedItems.length; translateIndex++) {
                    selectedItems[translateIndex].translate(previewOffsets.x, previewOffsets.y);
                }
            }
            app.redraw();
        };

        // 矢印キー押下時のステップ移動（スクリプトを1回実行したのと同等の挙動）
        // 現在の baseline を起点に1段階整列し、新しい位置を baseline として保存する
        var stepAlignment = function (alignmentSide, settings) {
            // プレビュー状態をリセットしてから、独立した1回の整列として適用
            for (var restoreIndex = 0; restoreIndex < selectedItems.length; restoreIndex++) {
                selectedItems[restoreIndex].position = originalPositions[restoreIndex];
            }
            var stepHorizontal = (alignmentSide === "left" || alignmentSide === "right") ? alignmentSide : null;
            var stepVertical = (alignmentSide === "top" || alignmentSide === "bottom") ? alignmentSide : null;
            var stepBounds = computeSelectionBoundsFromItems(selectedItems, settings.usePreviewBounds);
            var stepOffsets = computeAlignmentOffsets(stepBounds, stepHorizontal, stepVertical, ALIGNMENT_SIDE, settings.useGuides, activeArtboardRect, documentRef, GUIDE_SEARCH_MODE, GUIDE_ORIENTATION_TOLERANCE);
            for (var moveStepIndex = 0; moveStepIndex < selectedItems.length; moveStepIndex++) {
                selectedItems[moveStepIndex].translate(stepOffsets.x, stepOffsets.y);
            }
            // 新しい位置を baseline に更新（次回の → でさらに先へ進めるように）
            for (var updateIndex = 0; updateIndex < selectedItems.length; updateIndex++) {
                var newPosition = selectedItems[updateIndex].position;
                originalPositions[updateIndex] = [newPosition[0], newPosition[1]];
            }
            app.redraw();
        };

        var dialogResult = null;
        if (SHOW_DIALOG) {
            dialogResult = showAlignmentDialog(includeStrokeInBounds, USE_GUIDES, previewAlignment, stepAlignment);
            if (dialogResult === null) {
                // キャンセル: 矢印キーでのステップ移動も含めて完全復元
                for (var cancelIndex = 0; cancelIndex < selectedItems.length; cancelIndex++) {
                    selectedItems[cancelIndex].position = trueOriginalPositions[cancelIndex];
                }
                app.redraw();
                return;
            }
            // ラジオ未選択なら、矢印キーでの最終位置をそのまま確定（finalAlignment は適用しない）
            if (dialogResult.selectedHorizontal === undefined && dialogResult.selectedVertical === undefined) return;
            includeStrokeInBounds = dialogResult.usePreviewBounds;
            USE_GUIDES = dialogResult.useGuides;
        }

        // ダイアログ閉じ直後は元位置の状態。最終整列を fresh に適用する
        var finalHorizontal = (dialogResult !== null) ? dialogResult.selectedHorizontal : undefined;
        var finalVertical = (dialogResult !== null) ? dialogResult.selectedVertical : undefined;
        var finalBounds = computeSelectionBoundsFromItems(selectedItems, includeStrokeInBounds);
        var finalOffsets = computeAlignmentOffsets(finalBounds, finalHorizontal, finalVertical, ALIGNMENT_SIDE, USE_GUIDES, activeArtboardRect, documentRef, GUIDE_SEARCH_MODE, GUIDE_ORIENTATION_TOLERANCE);

        for (var moveItemIndex = 0; moveItemIndex < selectedItems.length; moveItemIndex++) {
            selectedItems[moveItemIndex].translate(finalOffsets.x, finalOffsets.y);
        }
    } catch (error) {
        alert(labelText("generalError") + error.message);
    }
}

// 整列オプションのダイアログを表示し、結果を返す（キャンセル時は null）
// previewCallback(settings) を渡すと「プレビュー」ON 時に都度呼ばれる。settings === null は復元要求
function showAlignmentDialog(initialUsePreviewBounds, initialUseGuides, previewCallback, stepCallback) {
    var dialog = new Window("dialog", getLabel("dialogTitle") + " " + SCRIPT_VERSION);
    dialog.alignChildren = "left";
    dialog.margins = 16;
    dialog.spacing = 12;

    // 3行3列のラジオボタン（ラベルなし、手動排他）
    // horizontal: "left" | "CENTER_X" | "right"
    // vertical: "top" | "CENTER_Y" | "bottom"
    var radioButtonCellDefinitions = [
        { horizontal: "left", vertical: "top" }, { horizontal: "CENTER_X", vertical: "top" }, { horizontal: "right", vertical: "top" },
        { horizontal: "left", vertical: "CENTER_Y" }, { horizontal: "CENTER_X", vertical: "CENTER_Y" }, { horizontal: "right", vertical: "CENTER_Y" },
        { horizontal: "left", vertical: "bottom" }, { horizontal: "CENTER_X", vertical: "bottom" }, { horizontal: "right", vertical: "bottom" }
    ];

    var alignmentPanel = dialog.add("panel", undefined, getLabel("alignmentPanelTitle"));
    alignmentPanel.orientation = "column";
    alignmentPanel.alignChildren = "center";
    alignmentPanel.alignment = "center";
    alignmentPanel.margins = [15, 20, 15, 10];

    var radioPanel = alignmentPanel.add("group");
    radioPanel.orientation = "column";
    radioPanel.alignChildren = "left";
    radioPanel.spacing = 4;
    radioPanel.margins = 8;

    var RADIO_CELL_SIZE = 24;

    var alignmentRadioButtons = [];
    var alignmentRadioMetadata = [];

    for (var rowIndex = 0; rowIndex < 3; rowIndex++) {
        var rowGroup = radioPanel.add("group");
        rowGroup.orientation = "row";
        rowGroup.spacing = 6;
        for (var columnIndex = 0; columnIndex < 3; columnIndex++) {
            var cellDefinition = radioButtonCellDefinitions[rowIndex * 3 + columnIndex];
            var cellGroup = rowGroup.add("group");
            cellGroup.orientation = "row";
            cellGroup.alignChildren = ["center", "center"];
            cellGroup.margins = 0;
            cellGroup.spacing = 0;
            cellGroup.preferredSize.width = RADIO_CELL_SIZE;
            cellGroup.preferredSize.height = RADIO_CELL_SIZE;
            var alignmentRadioButton = cellGroup.add("radiobutton", undefined, "");
            alignmentRadioButtons.push(alignmentRadioButton);
            alignmentRadioMetadata.push({ horizontal: cellDefinition.horizontal, vertical: cellDefinition.vertical });
        }
    }

    var optionPanel = dialog.add("panel", undefined, getLabel("optionPanelTitle"));
    optionPanel.orientation = "column";
    optionPanel.alignChildren = "left";
    optionPanel.alignment = "fill";
    optionPanel.margins = [15, 20, 15, 10];
    optionPanel.spacing = 6;

    var previewBoundsCheckbox = optionPanel.add("checkbox", undefined, getLabel("previewBoundsCheckbox"));
    previewBoundsCheckbox.value = initialUsePreviewBounds;

    var useGuidesCheckbox = optionPanel.add("checkbox", undefined, getLabel("useGuidesCheckbox"));
    useGuidesCheckbox.value = initialUseGuides;

    var previewCheckboxGroup = dialog.add("group");
    previewCheckboxGroup.orientation = "row";
    previewCheckboxGroup.alignment = "center";
    previewCheckboxGroup.alignChildren = ["center", "center"];

    var previewCheckbox = previewCheckboxGroup.add("checkbox", undefined, getLabel("previewCheckbox"));
    previewCheckbox.value = false;

    var buttonGroup = dialog.add("group");
    buttonGroup.alignment = "center";
    buttonGroup.add("button", undefined, getLabel("cancelButton"), { name: "cancel" });
    buttonGroup.add("button", undefined, getLabel("okButton"), { name: "ok" });

    function getSelectedAlignment() {
        for (var checkIndex = 0; checkIndex < alignmentRadioButtons.length; checkIndex++) {
            if (alignmentRadioButtons[checkIndex].value) return alignmentRadioMetadata[checkIndex];
        }
        return { horizontal: undefined, vertical: undefined };
    }

    function isRadioAlignmentSelected() {
        var alignment = getSelectedAlignment();
        return alignment.horizontal !== undefined || alignment.vertical !== undefined;
    }

    function syncUseGuidesCheckboxEnabled() {
        useGuidesCheckbox.enabled = !isRadioAlignmentSelected();
    }

    function getEffectiveUseGuides() {
        if (isRadioAlignmentSelected()) return false;
        return useGuidesCheckbox.value;
    }

    function triggerPreview() {
        if (!previewCallback) return;
        if (previewCheckbox.value) {
            var alignment = getSelectedAlignment();
            // ラジオ未選択なら元位置に戻すだけ（ALIGNMENT_SIDE フォールバックを抑止）
            if (alignment.horizontal === undefined && alignment.vertical === undefined) {
                previewCallback(null);
                return;
            }
            previewCallback({
                usePreviewBounds: previewBoundsCheckbox.value,
                useGuides: getEffectiveUseGuides(),
                selectedHorizontal: alignment.horizontal,
                selectedVertical: alignment.vertical
            });
        } else {
            previewCallback(null);
        }
    }

    // 別グループ間でも排他になるよう、クリック時に他をオフにし、プレビューを更新
    for (var radioIndex = 0; radioIndex < alignmentRadioButtons.length; radioIndex++) {
        alignmentRadioButtons[radioIndex].onClick = (function (clickedIndex) {
            return function () {
                for (var otherIndex = 0; otherIndex < alignmentRadioButtons.length; otherIndex++) {
                    if (otherIndex !== clickedIndex) alignmentRadioButtons[otherIndex].value = false;
                }
                syncUseGuidesCheckboxEnabled();
                triggerPreview();
            };
        })(radioIndex);
    }

    previewBoundsCheckbox.onClick = triggerPreview;
    useGuidesCheckbox.onClick = triggerPreview;
    previewCheckbox.onClick = triggerPreview;

    // 3x3 ラジオボタンのキーボードショートカット（w/e/r/s/d/f/x/c/v）
    var radioShortcutMap = {
        "W": 0, "E": 1, "R": 2,
        "S": 3, "D": 4, "F": 5,
        "X": 6, "C": 7, "V": 8
    };

    dialog.addEventListener("keyup", function (keyEvent) {
        // 矢印キー：押すたびに「スクリプトを1回実行」相当のステップ移動
        // ラジオ選択は触らない（OK 時に finalAlignment が二重適用されないようにするため）
        var stepSide = null;
        if (keyEvent.keyName === "Up") stepSide = "top";
        else if (keyEvent.keyName === "Down") stepSide = "bottom";
        else if (keyEvent.keyName === "Left") stepSide = "left";
        else if (keyEvent.keyName === "Right") stepSide = "right";
        if (stepSide !== null) {
            if (!stepCallback) return;
            for (var clearIndex = 0; clearIndex < alignmentRadioButtons.length; clearIndex++) {
                alignmentRadioButtons[clearIndex].value = false;
            }
            syncUseGuidesCheckboxEnabled();
            stepCallback(stepSide, {
                usePreviewBounds: previewBoundsCheckbox.value,
                useGuides: getEffectiveUseGuides()
            });
            return;
        }

        // ラジオボタンのショートカット：対応するラジオを選択してプレビューを更新
        if (radioShortcutMap.hasOwnProperty(keyEvent.keyName)) {
            var targetRadioIndex = radioShortcutMap[keyEvent.keyName];
            for (var radioToggleIndex = 0; radioToggleIndex < alignmentRadioButtons.length; radioToggleIndex++) {
                alignmentRadioButtons[radioToggleIndex].value = (radioToggleIndex === targetRadioIndex);
            }
            syncUseGuidesCheckboxEnabled();
            triggerPreview();
            return;
        }

        // チェックボックスのトグルショートカット
        if (keyEvent.keyName === "G") {
            if (!useGuidesCheckbox.enabled) return;
            useGuidesCheckbox.value = !useGuidesCheckbox.value;
            triggerPreview();
            return;
        }
        if (keyEvent.keyName === "B") {
            previewBoundsCheckbox.value = !previewBoundsCheckbox.value;
            triggerPreview();
            return;
        }
    });

    syncUseGuidesCheckboxEnabled();
    var dialogShowResult = dialog.show();

    // OK / キャンセルどちらでも、ダイアログを閉じる際は元位置に復元（main 側で最終整列を fresh に適用）
    if (previewCallback) previewCallback(null);

    if (dialogShowResult !== 1) return null;

    var finalAlignment = getSelectedAlignment();
    return {
        usePreviewBounds: previewBoundsCheckbox.value,
        useGuides: getEffectiveUseGuides(),
        selectedHorizontal: finalAlignment.horizontal,
        selectedVertical: finalAlignment.vertical
    };
}

// 選択オブジェクト群の合成バウンディングボックスを返す（[L, T, R, B]、Illustrator座標）
function computeSelectionBoundsFromItems(pageItems, includeStrokeInBounds) {
    var leftEdge, topEdge, rightEdge, bottomEdge;
    for (var itemIndex = 0; itemIndex < pageItems.length; itemIndex++) {
        var itemBounds = getItemBounds(pageItems[itemIndex], includeStrokeInBounds);
        if (itemIndex === 0) {
            leftEdge = itemBounds[0];
            topEdge = itemBounds[1];
            rightEdge = itemBounds[2];
            bottomEdge = itemBounds[3];
        } else {
            if (itemBounds[0] < leftEdge) leftEdge = itemBounds[0];
            if (itemBounds[1] > topEdge) topEdge = itemBounds[1];
            if (itemBounds[2] > rightEdge) rightEdge = itemBounds[2];
            if (itemBounds[3] < bottomEdge) bottomEdge = itemBounds[3];
        }
    }
    return [leftEdge, topEdge, rightEdge, bottomEdge];
}

// ラジオ選択（horizontal/vertical）または ALIGNMENT_SIDE に基づき X/Y のオフセットを返す
// horizontal/vertical の各値: "left"|"right"|"CENTER_X" / "top"|"bottom"|"CENTER_Y" / null（その軸は動かさない） / undefined（ラジオ未選択 → ALIGNMENT_SIDE フォールバック）
function computeAlignmentOffsets(selectionBounds, selectedHorizontal, selectedVertical, alignmentSide, useGuides, artboardRect, documentRef, guideSearchMode, guideTolerance) {
    var offsetX = 0;
    var offsetY = 0;

    var hasExplicitRadioSelection = (selectedHorizontal !== undefined) || (selectedVertical !== undefined);

    if (hasExplicitRadioSelection) {
        if (selectedHorizontal) {
            offsetX = computeAxisOffset(selectedHorizontal, selectionBounds, artboardRect, useGuides, documentRef, guideSearchMode, guideTolerance);
        }
        if (selectedVertical) {
            offsetY = computeAxisOffset(selectedVertical, selectionBounds, artboardRect, useGuides, documentRef, guideSearchMode, guideTolerance);
        }
    } else if (alignmentSide === "CENTER_X") {
        offsetX = computeAxisOffset("CENTER_X", selectionBounds, artboardRect, useGuides, documentRef, guideSearchMode, guideTolerance);
    } else if (alignmentSide === "CENTER_Y") {
        offsetY = computeAxisOffset("CENTER_Y", selectionBounds, artboardRect, useGuides, documentRef, guideSearchMode, guideTolerance);
    } else if (alignmentSide === "CENTER") {
        offsetX = computeAxisOffset("CENTER_X", selectionBounds, artboardRect, useGuides, documentRef, guideSearchMode, guideTolerance);
        offsetY = computeAxisOffset("CENTER_Y", selectionBounds, artboardRect, useGuides, documentRef, guideSearchMode, guideTolerance);
    } else if (alignmentSide === "left" || alignmentSide === "right") {
        offsetX = computeAxisOffset(alignmentSide, selectionBounds, artboardRect, useGuides, documentRef, guideSearchMode, guideTolerance);
    } else if (alignmentSide === "top" || alignmentSide === "bottom") {
        offsetY = computeAxisOffset(alignmentSide, selectionBounds, artboardRect, useGuides, documentRef, guideSearchMode, guideTolerance);
    }

    return { x: offsetX, y: offsetY };
}

// 1軸分の整列オフセットを計算する（left/right/top/bottom/CENTER_X/CENTER_Y に対応）
function computeAxisOffset(alignmentSide, selectionBounds, artboardRect, useGuides, documentRef, guideSearchMode, guideTolerance) {
    if (alignmentSide === null) return 0;
    if (alignmentSide === "CENTER_X") return getHorizontalCenterValue(artboardRect) - getHorizontalCenterValue(selectionBounds);
    if (alignmentSide === "CENTER_Y") return getVerticalCenterValue(artboardRect) - getVerticalCenterValue(selectionBounds);

    var selectionEdge = getEdgeValueForAlignmentSide(selectionBounds, alignmentSide);
    var artboardEdge = getEdgeValueForAlignmentSide(artboardRect, alignmentSide);
    if (selectionEdge === null || artboardEdge === null) return 0;

    var targetEdge = artboardEdge;
    if (useGuides) {
        var guideValue = findGuideSnapValue(documentRef, artboardRect, selectionEdge, alignmentSide, guideSearchMode, guideTolerance);
        if (guideValue !== null) targetEdge = guideValue;
    }

    return targetEdge - selectionEdge;
}

// 指定した方向に対応する境界値を返す（不正な方向の場合は null）
function getEdgeValueForAlignmentSide(bounds, alignmentSide) {
    if (alignmentSide === "left") return bounds[0];
    if (alignmentSide === "top") return bounds[1];
    if (alignmentSide === "right") return bounds[2];
    if (alignmentSide === "bottom") return bounds[3];
    return null;
}

// 左右中央の座標を返す
function getHorizontalCenterValue(bounds) {
    return (bounds[0] + bounds[2]) / 2;
}

// 上下中央の座標を返す
function getVerticalCenterValue(bounds) {
    return (bounds[1] + bounds[3]) / 2;
}

// アートボード内側にあるガイドのうち、指定した方向と探索範囲に合う吸着先座標を返す（なければ null）
function findGuideSnapValue(documentRef, artboardRect, selectionEdge, alignmentSide, guideSearchMode, guideOrientationTolerance) {
    var nearestGuideValue = null;
    var nearestGuideDistance = null;
    var guidePathItems = documentRef.pathItems;

    for (var guidePathIndex = 0; guidePathIndex < guidePathItems.length; guidePathIndex++) {
        var guidePathItem = guidePathItems[guidePathIndex];
        if (guidePathItem.guides !== true) continue;

        var guideBounds = guidePathItem.geometricBounds; // [L, T, R, B]
        var guideValue = getGuideValueForAlignmentSide(guideBounds, alignmentSide, guideOrientationTolerance);
        if (guideValue === null) continue;
        if (!isGuideValueInsideArtboard(guideValue, artboardRect, alignmentSide)) continue;

        if (guideSearchMode === "inside") {
            if (!isGuideOnAlignmentSide(guideValue, selectionEdge, alignmentSide)) continue;

            if (nearestGuideValue === null || isGuideCloserFromInside(guideValue, nearestGuideValue, alignmentSide)) {
                nearestGuideValue = guideValue;
            }

        } else if (guideSearchMode === "nearest") {
            var guideDistance = Math.abs(guideValue - selectionEdge);

            if (nearestGuideDistance === null || guideDistance < nearestGuideDistance) {
                nearestGuideValue = guideValue;
                nearestGuideDistance = guideDistance;
            }

        } else {
            alert(labelText("invalidGuideSearchModeError") + guideSearchMode);
            return null;
        }
    }

    return nearestGuideValue;
}

// ALIGNMENT_SIDE に対応するガイド座標を返す（対応しない向きのガイドは null）
function getGuideValueForAlignmentSide(guideBounds, alignmentSide, guideOrientationTolerance) {
    var isVerticalGuide = Math.abs(guideBounds[2] - guideBounds[0]) <= guideOrientationTolerance;
    var isHorizontalGuide = Math.abs(guideBounds[1] - guideBounds[3]) <= guideOrientationTolerance;

    if ((alignmentSide === "left" || alignmentSide === "right") && isVerticalGuide) {
        return guideBounds[0];
    }

    if ((alignmentSide === "top" || alignmentSide === "bottom") && isHorizontalGuide) {
        return guideBounds[1];
    }

    return null;
}

// ガイド座標がアクティブアートボード内にあるかを返す
function isGuideValueInsideArtboard(guideValue, artboardRect, alignmentSide) {
    if (alignmentSide === "left" || alignmentSide === "right") {
        return guideValue >= artboardRect[0] && guideValue <= artboardRect[2];
    }

    if (alignmentSide === "top" || alignmentSide === "bottom") {
        return guideValue <= artboardRect[1] && guideValue >= artboardRect[3];
    }

    return false;
}

// GUIDE_SEARCH_MODE が inside のとき、揃える方向側にあるガイドかを返す
function isGuideOnAlignmentSide(guideValue, selectionEdge, alignmentSide) {
    if (alignmentSide === "left") return guideValue < selectionEdge;
    if (alignmentSide === "right") return guideValue > selectionEdge;
    if (alignmentSide === "top") return guideValue > selectionEdge;
    if (alignmentSide === "bottom") return guideValue < selectionEdge;
    return false;
}

// inside 側にある複数のガイドのうち、選択範囲に最も近いかを判定する
function isGuideCloserFromInside(guideValue, currentBestGuideValue, alignmentSide) {
    if (alignmentSide === "left") return guideValue > currentBestGuideValue;
    if (alignmentSide === "right") return guideValue < currentBestGuideValue;
    if (alignmentSide === "top") return guideValue < currentBestGuideValue;
    if (alignmentSide === "bottom") return guideValue > currentBestGuideValue;
    return false;
}

// 指定オブジェクトの境界を取得する関数
function getItemBounds(pageItem, includeStrokeInBounds) {
    var itemBounds;

    // クリッピンググループの場合、最初のアイテムのgeometricBoundsを使用
    if (pageItem.typename === "GroupItem" && pageItem.clipped === true) {
        itemBounds = pageItem.pageItems[0].geometricBounds;

    } else if (includeStrokeInBounds === true) {
        // 線を含める設定の場合、visibleBoundsを使用
        itemBounds = pageItem.visibleBounds;

    } else {
        // 線を含めない場合、geometricBoundsを使用
        itemBounds = pageItem.geometricBounds;
    }

    return itemBounds;
}

main(); 