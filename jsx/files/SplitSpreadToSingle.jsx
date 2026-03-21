#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
概要 / Overview
- 見開きページ相当のオブジェクトを検出し、左右2つの片ページに分割します。 / Detects spread-like objects and splits them into left and right single pages.
- 対象は選択オブジェクトのみ、またはドキュメント内のすべてから選べます。 / You can process either only the selected objects or all matching objects in the document.
- 偶数ページが左／右の綴じ方向に対応します。 / Supports both left-binding and right-binding page orders.
- 必要に応じて、アートボード名の連番リネームとアートボード再配置を実行できます。 / Optionally renames artboards sequentially and rearranges artboards after processing.
更新日 / Updated: 2026-03-21
*/

// =========================================
// バージョンとローカライズ
// =========================================
var SCRIPT_VERSION = "v1.1";

function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

var LABELS = {
    dialogTitle: {
        ja: "見開きページを片ページに",
        en: "Split Spread Pages into Single Pages"
    },
    panelTarget: {
        ja: "対象",
        en: "Target"
    },
    modeSelectionOnly: {
        ja: "選択したオブジェクトのみ",
        en: "Selected Objects Only"
    },
    modeAll: {
        ja: "すべて",
        en: "All"
    },
    panelEvenPage: {
        ja: "偶数ページ",
        en: "Even Pages"
    },
    sideRight: {
        ja: "右",
        en: "Right"
    },
    sideLeft: {
        ja: "左",
        en: "Left"
    },
    panelPostProcess: {
        ja: "後処理",
        en: "Post-Process"
    },
    renameArtboards: {
        ja: "アートボード名を連番でリネーム",
        en: "Rename Artboards Sequentially"
    },
    rearrangeArtboards: {
        ja: "アートボードの再配置",
        en: "Rearrange Artboards"
    },
    spacingHorizontal: {
        ja: "左右",
        en: "Horizontal"
    },
    spacingVertical: {
        ja: "上下",
        en: "Vertical"
    },
    cancel: {
        ja: "キャンセル",
        en: "Cancel"
    },
    ok: {
        ja: "OK",
        en: "OK"
    },
    groupNameLeftHalf: {
        ja: "左半分",
        en: "Left_Half"
    },
    groupNameRightHalf: {
        ja: "右半分",
        en: "Right_Half"
    },
    alertInvalidSpacing: {
        ja: "アートボードの再配置の間隔には数値を入力してください。",
        en: "Enter numeric values for artboard rearrangement spacing."
    },
    alertNoSpreadFoundAll: {
        ja: "見開きページ相当の PlacedItem / GroupItem / RasterItem が見つかりませんでした。",
        en: "No spread-like PlacedItem / GroupItem / RasterItem was found."
    },
    alertSelectSpreadObject: {
        ja: "分割したい見開きオブジェクトを選択してください。",
        en: "Select a spread object to split."
    },
    alertSelectionIsSingle: {
        ja: "選択オブジェクトは片ページ相当です。見開きページ相当のオブジェクトを選択してください。",
        en: "The selected object looks like a single page. Select a spread-like object."
    },
    alertSelectValidSpread: {
        ja: "見開きページ相当の PlacedItem / GroupItem / RasterItem を選択してください。",
        en: "Select a spread-like PlacedItem / GroupItem / RasterItem."
    },
    alertNoTargetArtboard: {
        ja: "処理対象のアートボードが見つかりませんでした。",
        en: "No target artboard was found for processing."
    },
    artboardNamePrefix: {
        ja: "アートボード ",
        en: "Artboard "
    }
};

function L(key) {
    return LABELS[key][lang];
}

var SPLIT_GROUP_NOTE_PREFIX = "__SplitSpreadToSingle__";

// =========================================
// メイン処理
// =========================================

(function () {
    if (app.documents.length === 0) {
        return;
    }

    var doc = app.activeDocument;

    /* ダイアログ / Dialog */
    var dlg = new Window("dialog", L('dialogTitle') + ' ' + SCRIPT_VERSION);
    dlg.alignChildren = "fill";

    var modePanel = dlg.add("panel", undefined, L('panelTarget'));
    modePanel.orientation = "column";
    modePanel.alignChildren = "left";
    modePanel.margins = [15, 20, 15, 10];
    var rbSelection = modePanel.add("radiobutton", undefined, L('modeSelectionOnly'));
    var rbAll = modePanel.add("radiobutton", undefined, L('modeAll'));
    rbAll.value = true;

    var evenPanel = dlg.add("panel", undefined, L('panelEvenPage'));
    evenPanel.orientation = "row";
    evenPanel.alignChildren = "left";
    evenPanel.margins = [15, 20, 15, 10];
    var rbEvenRight = evenPanel.add("radiobutton", undefined, L('sideRight'));
    var rbEvenLeft = evenPanel.add("radiobutton", undefined, L('sideLeft'));
    rbEvenLeft.value = true;

    var optionPanel = dlg.add("panel", undefined, L('panelPostProcess'));
    optionPanel.orientation = "column";
    optionPanel.alignChildren = "left";
    optionPanel.margins = [15, 20, 15, 10];

    var cbRenameArtboards = optionPanel.add("checkbox", undefined, L('renameArtboards'));
    cbRenameArtboards.value = true;

    var cbRearrangeArtboards = optionPanel.add("checkbox", undefined, L('rearrangeArtboards'));
    cbRearrangeArtboards.value = true;

    var spacingGroup = optionPanel.add("group");
    spacingGroup.orientation = "row";
    spacingGroup.alignChildren = ["left", "center"];

    var stSpacingHorizontal = spacingGroup.add("statictext", undefined, L('spacingHorizontal'));
    var etSpacingHorizontal = spacingGroup.add("edittext", undefined, "20");
    etSpacingHorizontal.characters = 5;

    var stSpacingVertical = spacingGroup.add("statictext", undefined, L('spacingVertical'));
    var etSpacingVertical = spacingGroup.add("edittext", undefined, "20");
    etSpacingVertical.characters = 5;

    function updateRearrangeUiEnabled() {
        var enabled = cbRearrangeArtboards.value;
        stSpacingHorizontal.enabled = enabled;
        etSpacingHorizontal.enabled = enabled;
        stSpacingVertical.enabled = enabled;
        etSpacingVertical.enabled = enabled;
    }

    cbRearrangeArtboards.onClick = updateRearrangeUiEnabled;
    updateRearrangeUiEnabled();

    var btnGroup = dlg.add("group");
    btnGroup.alignment = "center";
    btnGroup.add("button", undefined, L('cancel'), { name: "cancel" });
    btnGroup.add("button", undefined, L('ok'), { name: "ok" });

    if (dlg.show() !== 1) return;

    var processAll = rbAll.value;
    var evenOnRight = rbEvenRight.value;
    var renameArtboards = cbRenameArtboards.value;
    var rearrangeArtboards = cbRearrangeArtboards.value;
    var spacingHorizontal = parseFloat(etSpacingHorizontal.text);
    var spacingVertical = parseFloat(etSpacingVertical.text);

    if (rearrangeArtboards && (isNaN(spacingHorizontal) || isNaN(spacingVertical))) {
        alert(L('alertInvalidSpacing'));
        return;
    }

    var pageOffset;
    try {
        var pref = app.preferences;
        var spacing = pref.getRealPreference('artnewdialog/artboardSpacing');
        pageOffset = spacing / 2;
    } catch (e) {
        pageOffset = 5; // fallback
    }

    /* 対象オブジェクトの収集 / Collect target objects */
    var items = [];

    if (processAll) {
        var allCandidates = collectTargetItems(doc);
        for (var ai = 0; ai < allCandidates.length; ai++) {
            var allItem = allCandidates[ai];
            var allPageType = getPageTypeInfo(doc, allItem);
            if (allPageType.artboardIndex < 0) continue;
            if (allPageType.artboardType !== "spread") continue;
            if (allPageType.pageType === "spread") {
                items.push(allItem);
            }
        }

        if (items.length === 0) {
            alert(L('alertNoSpreadFoundAll'));
            return;
        }
    } else {
        if (doc.selection.length < 1) {
            alert(L('alertSelectSpreadObject'));
            return;
        }

        var spreadCount = 0;
        var singleCount = 0;
        var otherCount = 0;

        for (var s = 0; s < doc.selection.length; s++) {
            var sel = doc.selection[s];
            if (isTargetItem(sel)) {
                var selectedPageType = getPageTypeInfo(doc, sel);
                if (selectedPageType.artboardIndex < 0) {
                    otherCount++;
                    continue;
                }
                if (selectedPageType.artboardType === "spread" && selectedPageType.pageType === "spread") {
                    items.push(sel);
                    spreadCount++;
                } else if (selectedPageType.pageType === "single") {
                    singleCount++;
                } else {
                    otherCount++;
                }
            } else {
                otherCount++;
            }
        }

        if (items.length === 0) {
            if (singleCount > 0 && otherCount === 0) {
                alert(L('alertSelectionIsSingle'));
            } else {
                alert(L('alertSelectValidSpread'));
            }
            return;
        }
    }

    /* クリッピンググループを作る関数 / Create clipping group */
    function makeClip(targetItem, clipLeft, clipTop, clipRight, clipBottom, groupName) {
        var parent = targetItem.parent;

        var rect = parent.pathItems.rectangle(
            clipTop,
            clipLeft,
            clipRight - clipLeft,
            clipTop - clipBottom
        );

        rect.stroked = false;
        rect.filled = false;

        var group = parent.groupItems.add();
        group.name = groupName;

        targetItem.move(group, ElementPlacement.PLACEATEND);
        rect.move(group, ElementPlacement.PLACEATBEGINNING);

        rect.clipping = true;
        group.clipped = true;

        return group;
    }

    /* アートボード挿入でインデックスがずれるため、大きい番号順に処理する / Process in descending artboard order because insertions shift indices */
    /* 同一アートボードは最初の1件だけ採用し、重複処理を防ぐ / Use only the first item per artboard to avoid duplicate processing */
    var workList = [];
    var usedArtboardMap = {};
    for (var i = 0; i < items.length; i++) {
        var abIndex = getPageTypeInfo(doc, items[i]).artboardIndex;
        if (abIndex < 0) continue;
        if (usedArtboardMap[abIndex]) continue;

        usedArtboardMap[abIndex] = true;
        workList.push({
            item: items[i],
            abIndex: abIndex
        });
    }
    workList.sort(function (a, b) { return b.abIndex - a.abIndex; });

    if (workList.length === 0) {
        alert(L('alertNoTargetArtboard'));
        return;
    }

    /* 各オブジェクトを処理 / Process each object */
    doc.selection = null;

    for (var j = 0; j < workList.length; j++) {
        var item = workList[j].item;
        var abIndex = workList[j].abIndex;

        /* 元アートボードの位置とサイズを基準にする / Use the source artboard bounds as the base */
        var abRect = doc.artboards[abIndex].artboardRect;
        var left = abRect[0];
        var top = abRect[1];
        var right = abRect[2];
        var bottom = abRect[3];

        var width = right - left;
        var centerX = left + width / 2;
        var pageWidth = centerX - left;

        /* 複製（B） / Duplicate for right half */
        var dup = item.duplicate();

        /* アートボード基準の左右ページ矩形 / Left and right page rects based on the artboard */
        var leftClipLeft = left;
        var leftClipTop = top;
        var leftClipRight = centerX;
        var leftClipBottom = bottom;
        var rightClipLeft = centerX;
        var rightClipTop = top;
        var rightClipRight = right;
        var rightClipBottom = bottom;

        /* クリップ後の各グループは左右半分の矩形を基準位置とする / Use the clipped half rects as base positions for each group */

        /* 元オブジェクト(A) → 左半分 / Original object to left half */
        var groupA = makeClip(item, leftClipLeft, leftClipTop, leftClipRight, leftClipBottom, L('groupNameLeftHalf'));

        /* 複製(B) → 右半分 / Duplicate object to right half */
        var groupB = makeClip(dup, rightClipLeft, rightClipTop, rightClipRight, rightClipBottom, L('groupNameRightHalf'));

        /* アートボード処理 / Artboard processing */
        if (abIndex >= 0) {
            var leftPageRect;
            var rightPageRect;
            var moveAX = 0;
            var moveBX = 0;
            var targetALeft;
            var targetBLeft;

            if (evenOnRight) {
                rightPageRect = [centerX, top, right, bottom];
                leftPageRect = [left - pageOffset, top, centerX - pageOffset, bottom];

                /* 見た目の左→右に合わせて、元のアートボードを左ページ、追加アートボードを右ページにする / Keep artboard order consistent with visual left-to-right order */
                doc.artboards[abIndex].artboardRect = leftPageRect;
                doc.artboards.setActiveArtboardIndex(abIndex);
                doc.artboards.insert(rightPageRect, abIndex + 1);

                /* 左半分（groupA）は右ページへ、右半分（groupB）は左ページへ移す / Move left half to right page and right half to left page */
                targetALeft = rightPageRect[0];
                targetBLeft = leftPageRect[0];
            } else {
                /* 偶数ページが左（デフォルト） / Even pages on the left (default) */
                leftPageRect = [left, top, centerX, bottom];
                rightPageRect = [centerX + pageOffset, top, right + pageOffset, bottom];

                /* 元のアートボードを左ページ、追加アートボードを右ページにする / Keep original artboard as left page and inserted artboard as right page */
                doc.artboards[abIndex].artboardRect = leftPageRect;
                doc.artboards.setActiveArtboardIndex(abIndex);
                doc.artboards.insert(rightPageRect, abIndex + 1);

                /* 左半分（groupA）は左ページのまま、右半分（groupB）は右ページへ移す / Keep left half on left page and move right half to right page */
                targetALeft = leftPageRect[0];
                targetBLeft = rightPageRect[0];
            }

            /* クリップ前提の基準位置から配置先までの移動量を算出 / Calculate translation from clipped base positions to destination artboards */
            moveAX = targetALeft - leftClipLeft;
            moveBX = targetBLeft - rightClipLeft;

            groupA.translate(moveAX, 0);
            groupB.translate(moveBX, 0);

            groupA.note = SPLIT_GROUP_NOTE_PREFIX + ":role=A";
            groupB.note = SPLIT_GROUP_NOTE_PREFIX + ":role=B";
        }

        groupA.selected = true;
        groupB.selected = true;
    }

    /* アートボード名をリネーム（オプション） / Rename artboards (optional) */
    if (renameArtboards) {
        for (var k = 0; k < doc.artboards.length; k++) {
            doc.artboards[k].name = L('artboardNamePrefix') + (k + 1);
        }
    }

    /* アートボード再配置（オプション） / Rearrange artboards (optional) */
    if (rearrangeArtboards) {
        rearrangeArtboardsByPageOrder(doc, spacingHorizontal, spacingVertical, evenOnRight);
    }

    // =========================================
    // アートボード再配置
    // =========================================
    function unlockAndUnhideAll(container) {
        if (!container) return;

        if (container.typename === "Document" || container.typename === "Layer") {
            var layers = container.layers;
            for (var i = 0; i < layers.length; i++) {
                var lr = layers[i];
                if (lr.locked) lr.locked = false;
                if (!lr.visible) lr.visible = true;
                unlockAndUnhideAll(lr);
            }
        }

        if (container.pageItems) {
            var items = container.pageItems;
            for (var j = 0; j < items.length; j++) {
                var item = items[j];
                if (item.locked) item.locked = false;
                if (item.hidden) item.hidden = false;
                if (item.typename === "GroupItem") {
                    unlockAndUnhideAll(item);
                }
            }
        }
    }

    function rearrangeArtboardsByPageOrder(doc, spacingX, spacingY, evenOnRight) {
        if (!doc || !doc.artboards || doc.artboards.length === 0) return;

        unlockAndUnhideAll(doc);

        var artboards = doc.artboards;
        var numArtboards = artboards.length;
        var firstRect = artboards[0].artboardRect;
        var baseX = firstRect[0];
        var baseY = firstRect[1];
        var firstPageRight = !evenOnRight;
        var artboardItemMap = [];
        var tempMarkedItems = [];
        var tempOriginalNotes = [];
        var tempMarker = "__SplitSpreadToSingle_Rearrange__";

        for (var ai = 0; ai < numArtboards; ai++) {
            artboards.setActiveArtboardIndex(ai);
            doc.selection = null;
            doc.selectObjectsOnActiveArtboard();

            var pageItems = [];
            for (var aj = 0; aj < doc.selection.length; aj++) {
                var selectedItem = doc.selection[aj];
                if (!selectedItem) continue;

                var noteText = String(selectedItem.note || "");
                if (noteText.indexOf(tempMarker) >= 0) continue;

                tempMarkedItems.push(selectedItem);
                tempOriginalNotes.push(noteText);
                selectedItem.note = noteText ? (noteText + "\n" + tempMarker) : tempMarker;
                pageItems.push(selectedItem);
            }

            artboardItemMap[ai] = pageItems;
            doc.selection = null;
        }

        for (var am = 0; am < tempMarkedItems.length; am++) {
            tempMarkedItems[am].note = tempOriginalNotes[am];
        }

        for (var i = 0; i < numArtboards; i++) {
            var ab = artboards[i];
            var sel = artboardItemMap[i] || [];
            var oldRect = ab.artboardRect;
            var w = oldRect[2] - oldRect[0];
            var h = oldRect[1] - oldRect[3];
            var pageNumber = i + 1;
            var isThisPageRight;
            var row;
            var newLeft;

            if (pageNumber === 1) {
                row = 0;
                newLeft = baseX;
            } else {
                if (firstPageRight) {
                    isThisPageRight = (pageNumber % 2 !== 0);
                    newLeft = isThisPageRight ? baseX : (baseX - w - spacingX);
                } else {
                    isThisPageRight = (pageNumber % 2 === 0);
                    newLeft = isThisPageRight ? (baseX + w + spacingX) : baseX;
                }
                row = Math.floor(pageNumber / 2);
            }

            var newTop = baseY - row * (h + spacingY);
            var newRight = newLeft + w;
            var newBottom = newTop - h;
            var deltaX = newLeft - oldRect[0];
            var deltaY = newTop - oldRect[1];

            ab.artboardRect = [newLeft, newTop, newRight, newBottom];

            if (deltaX !== 0 || deltaY !== 0) {
                for (var k = 0; k < sel.length; k++) {
                    sel[k].translate(deltaX, deltaY, true, true, true, true);
                }
            }
        }

        app.executeMenuCommand("fitall");
    }

    // =========================================
    // ヘルパー関数
    // =========================================
    /* ヘルパー関数 / Helper functions */



    function isTargetItem(obj) {
        return obj && (
            obj.typename === "PlacedItem" ||
            obj.typename === "GroupItem" ||
            obj.typename === "RasterItem"
        );
    }

    function collectTargetItems(container) {
        var results = [];
        if (!container || !container.pageItems) return results;

        for (var i = 0; i < container.pageItems.length; i++) {
            var item = container.pageItems[i];
            if (isTargetItem(item)) {
                results.push(item);
            }
        }
        return results;
    }

    /* 最も重なり面積の大きいアートボードを返す / Return the artboard with the largest overlap area */
    function getPrimaryArtboardIndex(doc, obj) {
        var bounds = obj.geometricBounds;
        var left = bounds[0];
        var top = bounds[1];
        var right = bounds[2];
        var bottom = bounds[3];
        var bestIndex = -1;
        var bestArea = 0;

        for (var i = 0; i < doc.artboards.length; i++) {
            var abRect = doc.artboards[i].artboardRect;
            var overlapLeft = Math.max(left, abRect[0]);
            var overlapTop = Math.min(top, abRect[1]);
            var overlapRight = Math.min(right, abRect[2]);
            var overlapBottom = Math.max(bottom, abRect[3]);
            var overlapWidth = overlapRight - overlapLeft;
            var overlapHeight = overlapTop - overlapBottom;

            if (overlapWidth > 0 && overlapHeight > 0) {
                var overlapArea = overlapWidth * overlapHeight;
                if (overlapArea > bestArea) {
                    bestArea = overlapArea;
                    bestIndex = i;
                }
            }
        }
        return bestIndex;
    }

    /* ドキュメント内の基準ページ幅（最小アートボード幅）を返す / Return the reference page width as the minimum artboard width */
    function getReferencePageWidth(doc) {
        var minWidth = null;
        for (var i = 0; i < doc.artboards.length; i++) {
            var abRect = doc.artboards[i].artboardRect;
            var abWidth = abRect[2] - abRect[0];
            if (abWidth <= 0) continue;
            if (minWidth === null || abWidth < minWidth) {
                minWidth = abWidth;
            }
        }
        return (minWidth === null) ? 0 : minWidth;
    }

    /* 指定アートボードが片ページ相当か見開き相当かを返す / Determine whether an artboard is single-page-like or spread-like */
    function getArtboardType(doc, artboardIndex, referencePageWidth) {
        if (artboardIndex < 0 || referencePageWidth <= 0) return "other";

        var abRect = doc.artboards[artboardIndex].artboardRect;
        var abWidth = abRect[2] - abRect[0];
        var ratio = abWidth / referencePageWidth;
        var singleTolerance = 0.15;
        var spreadTolerance = 0.15;

        if (Math.abs(ratio - 1) <= singleTolerance) return "single";
        if (Math.abs(ratio - 2) <= spreadTolerance) return "spread";
        return "other";
    }

    /* オブジェクト幅と基準ページ幅から片ページ/見開き/その他を判定 / Classify object as single, spread, or other by width */
    function getPageTypeInfo(doc, obj) {
        var artboardIndex = getPrimaryArtboardIndex(doc, obj);
        if (artboardIndex < 0) {
            return {
                artboardIndex: -1,
                pageType: "other",
                ratio: 0,
                artboardType: "other"
            };
        }

        var referencePageWidth = getReferencePageWidth(doc);
        if (referencePageWidth <= 0) {
            return {
                artboardIndex: artboardIndex,
                pageType: "other",
                ratio: 0,
                artboardType: "other"
            };
        }

        var bounds = obj.geometricBounds;
        var objWidth = bounds[2] - bounds[0];
        var abRect = doc.artboards[artboardIndex].artboardRect;
        var abWidth = abRect[2] - abRect[0];
        var ratio = objWidth / referencePageWidth;
        var fitToArtboardRatio = objWidth / abWidth;
        var singleTolerance = 0.15;
        var spreadTolerance = 0.15;
        var fitTolerance = 0.15;
        var artboardType = getArtboardType(doc, artboardIndex, referencePageWidth);
        var pageType = "other";

        /* 見開きアートボード上でアートボード幅にほぼ一致するオブジェクトは見開き扱い / Treat objects that match artboard width on spread artboards as spread */
        if (artboardType === "spread" && Math.abs(fitToArtboardRatio - 1) <= fitTolerance) {
            pageType = "spread";
        } else if (artboardType === "single" && Math.abs(ratio - 1) <= singleTolerance) {
            pageType = "single";
        } else if (artboardType === "spread" && Math.abs(ratio - 2) <= spreadTolerance) {
            pageType = "spread";
        }

        return {
            artboardIndex: artboardIndex,
            pageType: pageType,
            ratio: ratio,
            artboardType: artboardType
        };
    }
})();
