#target illustrator
#targetengine "session"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

var SCRIPT_VERSION = "v1.3";

/*
 * ひとつの長方形に変換.jsx / Convert to One Rectangle.jsx
 *
 * 概要 / Overview:
 * 選択した複数オブジェクト全体の外接矩形をもとに、ひとつの長方形へ統合します。
 * 属性の引き継ぎ元を選択でき、プレビュー境界 / オブジェクト境界の切り替え、
 * 元の図形を残す、プレビュー表示、属性パネルでの中心表示に対応しています。
 * 引き継ぎ元が長方形の場合は、そのオブジェクト自体を再利用して座標と大きさを更新し、
 * visibleBounds 使用時は stroke alignment / stroke width を考慮して補正します。
 * ダイアログボックスの位置はセッション中に記憶され、次回表示時に復元されます。
 * Creates one merged rectangle from the overall bounds of multiple selected objects.
 * You can choose the source object for appearance, switch between preview bounds and geometric bounds,
 * keep the original objects, use preview display, and enable center display in the Attributes panel.
 * If the source object is a rectangle, the object itself is reused and its position and size are updated.
 * When visibleBounds is used, the rectangle geometry is compensated for stroke alignment and stroke width.
 * The dialog position is remembered during the session and restored the next time it is shown.
 *
 * 更新日 / Updated: 2026-03-08
 */

function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

/* 日英ラベル定義 / Japanese-English label definitions */
var LABELS = {
    dialogTitle: {
        ja: "ひとつの長方形に変換",
        en: "Convert to One Rectangle"
    },
    noDocument: {
        ja: "ドキュメントが開かれていません。",
        en: "No document is open."
    },
    noSelection: {
        ja: "オブジェクトを選択してから実行してください。",
        en: "Please select objects before running this script."
    },
    pathItem: {
        ja: "パス",
        en: "Path"
    },
    compoundPathItem: {
        ja: "複合パス",
        en: "Compound Path"
    },
    groupItem: {
        ja: "グループ",
        en: "Group"
    },
    textFrame: {
        ja: "テキスト",
        en: "Text"
    },
    placedItem: {
        ja: "配置画像",
        en: "Placed Image"
    },
    rasterItem: {
        ja: "ラスター画像",
        en: "Raster Image"
    },
    symbolItem: {
        ja: "シンボル",
        en: "Symbol"
    },
    meshItem: {
        ja: "メッシュ",
        en: "Mesh"
    },
    pluginItem: {
        ja: "プラグイン",
        en: "Plugin"
    },
    objectLabelPrefix: {
        ja: "",
        en: ""
    },
    inheritSourcePanel: {
        ja: "属性の引き継ぎ元",
        en: "Source for Attributes"
    },
    optionsPanel: {
        ja: "オプション",
        en: "Options"
    },
    useVisibleBounds: {
        ja: "プレビュー境界を使用",
        en: "Use Preview Bounds"
    },
    useVisibleBoundsTip: {
        ja: "線幅を含む見た目の外接矩形で計算",
        en: "Use visible bounds including stroke width"
    },
    keepOriginals: {
        ja: "元の図形を残す",
        en: "Keep Original Objects"
    },
    keepOriginalsTip: {
        ja: "元のオブジェクトを残して新しい長方形を作成",
        en: "Keep the originals and create a new rectangle"
    },
    showCenter: {
        ja: "属性パネルで中心を表示",
        en: "Show Center"
    },
    showCenterTip: {
        ja: "確定後に反映",
        en: "Applied after confirmation"
    },
    preview: {
        ja: "プレビュー",
        en: "Preview"
    },
    previewTip: {
        ja: "合体後の長方形を一時表示",
        en: "Temporarily show the merged rectangle"
    },
    cancel: {
        ja: "キャンセル",
        en: "Cancel"
    },
    ok: {
        ja: "OK",
        en: "OK"
    },
    tempLayerName: {
        ja: "_選択プレビュー（一時）",
        en: "_Selection Preview (Temp)"
    },
    resultLayerName: {
        ja: "",
        en: ""
    }
};

function L(key) {
    var item = LABELS[key];
    if (!item) return key;
    if (lang === "ja") {
        return (item.ja !== undefined) ? item.ja : key;
    }
    return (item.en !== undefined) ? item.en : ((item.ja !== undefined) ? item.ja : key);
}

var __OneRectDialogState__ = $.global.__OneRectDialogState__ || ($.global.__OneRectDialogState__ = {});

function main() {
    function saveDialogLocation(dlg) {
        try {
            if (!dlg || !dlg.location) return;
            var x = dlg.location[0];
            var y = dlg.location[1];
            if (isNaN(x) || isNaN(y)) return;
            __OneRectDialogState__.location = [x, y];
        } catch (e) {
            $.writeln("[OneRect] saveDialogLocation error: " + e);
        }
    }

    function restoreDialogLocation(dlg) {
        try {
            if (!dlg || !__OneRectDialogState__ || !__OneRectDialogState__.location) return;
            var loc = __OneRectDialogState__.location;
            if (loc.length !== 2 || isNaN(loc[0]) || isNaN(loc[1])) return;
            dlg.location = [loc[0], loc[1]];
        } catch (e) {
            $.writeln("[OneRect] restoreDialogLocation error: " + e);
        }
    }

    /* 選択情報を収集 / Collect selection data */
    function collectSelectionData(selection) {
        var objects = [];
        var visibleBoundsData = [];
        var geometricBoundsData = [];
        var originalHiddenStates = [];

        for (var i = 0; i < selection.length; i++) {
            objects.push(selection[i]);
            originalHiddenStates.push(selection[i].hidden);

            var vb = selection[i].visibleBounds;
            visibleBoundsData.push({ left: vb[0], top: vb[1], right: vb[2], bottom: vb[3] });

            var gb = selection[i].geometricBounds;
            geometricBoundsData.push({ left: gb[0], top: gb[1], right: gb[2], bottom: gb[3] });
        }

        return {
            objects: objects,
            visibleBoundsData: visibleBoundsData,
            geometricBoundsData: geometricBoundsData,
            originalHiddenStates: originalHiddenStates
        };
    }

    /* ダイアログを表示して設定を取得 / Show dialog and get options */
    function showDialogAndGetOptions() {
        /* ダイアログボックスを作成 / Build dialog */
        var dlg = new Window("dialog", L("dialogTitle") + " " + SCRIPT_VERSION);
        dlg.alignChildren = ["left", "top"];

        var radioPanel = dlg.add("panel", undefined, L("inheritSourcePanel"));
        radioPanel.alignment = ["fill", "fill"];
        radioPanel.alignChildren = ["left", "top"];
        radioPanel.margins = [15, 20, 15, 10];

        var radioButtons = [];
        for (var i = 0; i < objects.length; i++) {
            var rb = radioPanel.add("radiobutton", undefined, getObjectLabel(objects[i], i));
            if (i === 0) rb.value = true;
            rb.onClick = function () { updatePreview(radioButtons, previewCheck); };
            radioButtons.push(rb);
        }

        var optionPanel = dlg.add("panel", undefined, L("optionsPanel"));
        optionPanel.alignment = ["fill", "fill"];
        optionPanel.alignChildren = ["left", "top"];
        optionPanel.margins = [15, 20, 15, 10];

        var useVisibleBounds = optionPanel.add("checkbox", undefined, L("useVisibleBounds"));
        useVisibleBounds.value = true;
        useVisibleBounds.helpTip = L("useVisibleBoundsTip");
        useVisibleBounds.onClick = function () {
            activeBoundsData = useVisibleBounds.value ? visibleBoundsData : geometricBoundsData;
            updatePreview(radioButtons, previewCheck);
        };

        var keepOriginals = optionPanel.add("checkbox", undefined, L("keepOriginals"));
        keepOriginals.value = false;
        keepOriginals.helpTip = L("keepOriginalsTip");
        var showCenter = optionPanel.add("checkbox", undefined, L("showCenter"));
        showCenter.value = true;
        showCenter.helpTip = L("showCenterTip");

        var bottomGroup = dlg.add("group");
        bottomGroup.alignment = ["fill", "top"];
        bottomGroup.alignChildren = ["center", "center"];

        var previewCheck = bottomGroup.add("checkbox", undefined, L("preview"));
        previewCheck.value = false;
        previewCheck.helpTip = L("previewTip");
        previewCheck.onClick = function () { updatePreview(radioButtons, previewCheck); };

        var btnGroup = dlg.add("group");
        btnGroup.add("button", undefined, L("cancel"), { name: "cancel" });
        btnGroup.add("button", undefined, L("ok"), { name: "ok" });

        dlg.layout.layout(true);
        restoreDialogLocation(dlg);
        dlg.onMove = function () {
            saveDialogLocation(dlg);
        };

        /* キーでラジオボタンを選択（j=1番目, k=2番目, l=3番目, ...） / Select radio buttons by keyboard */
        var keyMap = ["J", "K", "L", "Semicolon", "A", "S", "D", "F", "G"];
        dlg.addEventListener("keydown", function (e) {
            for (var i = 0; i < keyMap.length && i < radioButtons.length; i++) {
                if (e.keyName == keyMap[i]) {
                    for (var k = 0; k < radioButtons.length; k++) {
                        radioButtons[k].value = (k === i);
                    }
                    updatePreview(radioButtons, previewCheck);
                    e.preventDefault();
                    break;
                }
            }
        });

        /* 初期プレビューを描画 / Draw initial preview */
        updatePreview(radioButtons, previewCheck);

        var result = dlg.show();
        saveDialogLocation(dlg);

        var sourceIndex = 0;
        if (result === 1) {
            for (var i = 0; i < radioButtons.length; i++) {
                if (radioButtons[i].value) {
                    sourceIndex = i;
                    break;
                }
            }
        }

        return {
            result: result,
            sourceIndex: sourceIndex,
            keepOriginals: keepOriginals.value,
            showCenter: showCenter.value
        };
    }

    function act_ShowCenter() {

        var str = '/version 3' + '/name [ 9' + ' 417474726962757465' + ']' + '/isOpen 1' + '/actionCount 1' + '/action-1 {' + ' /name [ 10' + ' 53686f7743656e746572' + ' ]' + ' /keyIndex 0' + ' /colorIndex 0' + ' /isOpen 1' + ' /eventCount 1' + ' /event-1 {' + ' /useRulersIn1stQuadrant 0' + ' /internalName (adobe_attributePalette)' + ' /localizedName [ 12' + ' e5b19ee680a7e8a8ade5ae9a' + ' ]' + ' /isOpen 1' + ' /isOn 1' + ' /hasDialog 0' + ' /parameterCount 1' + ' /parameter-1 {' + ' /key 1668183154' + ' /showInPalette 4294967295' + ' /type (boolean)' + ' /value 1' + ' }' + ' }' + '}';

        var f = new File('~/ScriptAction.aia');
        f.open('w');
        f.write(str);
        f.close();

        app.loadAction(f);
        f.remove();
        app.doScript("ShowCenter", "Attribute", false); // action name, set name
        app.unloadAction("Attribute", ""); // set name
    }


    /* 結果を適用 / Apply final result */
    function applyMergeResult(bounds, sourceIndex, keepOriginals) {
        var baseObj = objects[sourceIndex];
        var reuseSourceRectangle = (!keepOriginals && isAxisAlignedRectanglePathItem(baseObj));

        if (reuseSourceRectangle) {
            /* 引き継ぎ元が長方形なら、そのオブジェクト自体を変形して使う */
            /* 同じオブジェクトを再利用するため copyBasicAppearance() は不要 / Reusing the same object, so copyBasicAppearance() is unnecessary */
            reshapeRectanglePathItem(baseObj, bounds, baseObj);

            for (var j = objects.length - 1; j >= 0; j--) {
                if (j === sourceIndex) continue;
                objects[j].remove();
            }

            baseObj.selected = true;
            return;
        }

        if (keepOriginals) {
            /* 元の図形を残す: 新しい長方形を作成 */
            var newRect = createMergedRectangleFromSource(bounds, baseObj);
            newRect.selected = true;
        } else {
            /* 元の図形を削除して新しい長方形を作成 */
            var mergedRect = createMergedRectangleFromSource(bounds, baseObj);

            for (var k = objects.length - 1; k >= 0; k--) {
                objects[k].remove();
            }

            mergedRect.selected = true;
        }
    }

    if (app.documents.length === 0) {
        alert(L("noDocument"));
        return;
    }

    var doc = app.activeDocument;
    var sel = doc.selection;

    if (sel.length === 0) {
        alert(L("noSelection"));
        return;
    }

    /* 選択オブジェクトの情報を事前に保存 / Cache selected object data in advance */
    var selectionData = collectSelectionData(sel);
    var objects = selectionData.objects;
    var visibleBoundsData = selectionData.visibleBoundsData;
    var geometricBoundsData = selectionData.geometricBoundsData;
    var originalHiddenStates = selectionData.originalHiddenStates;


    /* 使用するboundsデータ（初期値はgeometricBounds） / Active bounds data (default: geometricBounds) */
    var activeBoundsData = geometricBoundsData;

    /* 選択を解除（赤枠が見やすいように） / Clear selection to make preview easier to see */
    doc.selection = null;

    /* typenameの表示名マッピング / Localized typename label mapping */
    var typenameMap = {
        "PathItem": L("pathItem"),
        "CompoundPathItem": L("compoundPathItem"),
        "GroupItem": L("groupItem"),
        "TextFrame": L("textFrame"),
        "PlacedItem": L("placedItem"),
        "RasterItem": L("rasterItem"),
        "SymbolItem": L("symbolItem"),
        "MeshItem": L("meshItem"),
        "PluginItem": L("pluginItem")
    };

    /* オブジェクトの表示名を取得 / Get display label for an object */
    function getObjectLabel(obj, index) {
        var label = "";
        if (obj.name && obj.name !== "") {
            label = obj.name;
        } else {
            label = typenameMap[obj.typename] || obj.typename;
        }
        return (index + 1) + ": " + L("objectLabelPrefix") + label;
    }

    /* 一時レイヤー（遅延生成） / Temporary preview layer (lazy creation) */
    var tempLayer = null;

    function ensureTempLayer() {
        if (!tempLayer) {
            tempLayer = doc.layers.add();
            tempLayer.name = L("tempLayerName");
        }
    }

    var redColor = new RGBColor();
    redColor.red = 255;
    redColor.green = 0;
    redColor.blue = 0;

    /* 赤枠を描画する関数 / Draw red highlight rectangle */
    function drawHighlight(index) {
        ensureTempLayer();
        var b = activeBoundsData[index];
        var rect = tempLayer.pathItems.rectangle(b.top, b.left, b.right - b.left, b.top - b.bottom);
        rect.filled = false;
        rect.stroked = true;
        rect.strokeColor = redColor;
        rect.strokeWidth = 2;
    }

    /* 合体後のバウンディングボックスを計算する関数 / Calculate merged bounds */
    function calcMergedBounds() {
        var minLeft = Infinity, maxTop = -Infinity, maxRight = -Infinity, minBottom = Infinity;
        for (var i = 0; i < activeBoundsData.length; i++) {
            var b = activeBoundsData[i];
            if (b.left < minLeft) minLeft = b.left;
            if (b.top > maxTop) maxTop = b.top;
            if (b.right > maxRight) maxRight = b.right;
            if (b.bottom < minBottom) minBottom = b.bottom;
        }
        return { left: minLeft, top: maxTop, right: maxRight, bottom: minBottom };
    }

    function isAxisAlignedRectanglePathItem(obj) {
        if (!obj || obj.typename !== "PathItem") return false;
        if (!obj.closed) return false;
        if (!obj.pathPoints || obj.pathPoints.length !== 4) return false;

        try {
            var pts = [];
            for (var i = 0; i < obj.pathPoints.length; i++) {
                var a = obj.pathPoints[i].anchor;
                pts.push([a[0], a[1]]);
            }

            var xs = {};
            var ys = {};
            for (var j = 0; j < pts.length; j++) {
                xs[String(pts[j][0])] = true;
                ys[String(pts[j][1])] = true;
            }

            var xCount = 0, yCount = 0;
            for (var k in xs) if (xs.hasOwnProperty(k)) xCount++;
            for (var m in ys) if (ys.hasOwnProperty(m)) yCount++;

            return (xCount === 2 && yCount === 2);
        } catch (e) {
            $.writeln("[OneRect] isAxisAlignedRectanglePathItem error: " + e);
            return false;
        }
    }

    function reshapeRectanglePathItem(rectObj, bounds, sourceObj) {
        var geomBounds = getGeometryBoundsForRectangle(bounds, sourceObj);

        rectObj.setEntirePath([
            [geomBounds.left, geomBounds.top],
            [geomBounds.right, geomBounds.top],
            [geomBounds.right, geomBounds.bottom],
            [geomBounds.left, geomBounds.bottom]
        ]);
        rectObj.closed = true;

        for (var i = 0; i < rectObj.pathPoints.length; i++) {
            var pt = rectObj.pathPoints[i];
            pt.leftDirection = pt.anchor;
            pt.rightDirection = pt.anchor;
            pt.pointType = PointType.CORNER;
        }
    }

    function getStrokeAlignmentName(obj) {
        try {
            if (!obj || !obj.stroked || !obj.strokeWidth) return "center";
            var a = obj.strokeAlignment;
            if (a === undefined || a === null) return "center";

            var s = String(a).toLowerCase();
            if (s.indexOf("inside") >= 0) return "inside";
            if (s.indexOf("outside") >= 0) return "outside";
            if (s.indexOf("center") >= 0) return "center";
        } catch (e) {
            $.writeln("[OneRect] getStrokeAlignmentName error: " + e);
        }
        return "center";
    }

    function getGeometryBoundsForRectangle(bounds, sourceObj) {
        var result = {
            left: bounds.left,
            top: bounds.top,
            right: bounds.right,
            bottom: bounds.bottom
        };

        if (activeBoundsData !== visibleBoundsData) {
            return result;
        }
        if (!sourceObj || !sourceObj.stroked || !sourceObj.strokeWidth) {
            return result;
        }

        var sw = Number(sourceObj.strokeWidth);
        if (!(sw > 0)) return result;

        var align = getStrokeAlignmentName(sourceObj);
        var inset = 0;

        if (align === "center") {
            inset = sw / 2;
        } else if (align === "outside") {
            inset = sw;
        } else {
            inset = 0; // inside
        }

        result.left += inset;
        result.top -= inset;
        result.right -= inset;
        result.bottom += inset;

        if (result.right < result.left) {
            var cx = (result.left + result.right) / 2;
            result.left = cx;
            result.right = cx;
        }
        if (result.top < result.bottom) {
            var cy = (result.top + result.bottom) / 2;
            result.top = cy;
            result.bottom = cy;
        }

        return result;
    }

    function copyBasicAppearance(sourceObj, targetObj) {
        if (sourceObj.filled) {
            targetObj.filled = true;
            targetObj.fillColor = sourceObj.fillColor;
        } else {
            targetObj.filled = false;
        }

        if (sourceObj.stroked) {
            targetObj.stroked = true;
            targetObj.strokeColor = sourceObj.strokeColor;
            targetObj.strokeWidth = sourceObj.strokeWidth;

            try {
                targetObj.strokeDashes = sourceObj.strokeDashes;
            } catch (eDashes) {
                $.writeln("[OneRect] copy strokeDashes error: " + eDashes);
            }
            try {
                targetObj.strokeDashOffset = sourceObj.strokeDashOffset;
            } catch (eDashOffset) {
                $.writeln("[OneRect] copy strokeDashOffset error: " + eDashOffset);
            }
            try {
                targetObj.strokeCap = sourceObj.strokeCap;
            } catch (eCap) {
                $.writeln("[OneRect] copy strokeCap error: " + eCap);
            }
            try {
                targetObj.strokeJoin = sourceObj.strokeJoin;
            } catch (eJoin) {
                $.writeln("[OneRect] copy strokeJoin error: " + eJoin);
            }
            try {
                targetObj.strokeMiterLimit = sourceObj.strokeMiterLimit;
            } catch (eMiter) {
                $.writeln("[OneRect] copy strokeMiterLimit error: " + eMiter);
            }
            try {
                targetObj.strokeOverprint = sourceObj.strokeOverprint;
            } catch (eStrokeOverprint) {
                $.writeln("[OneRect] copy strokeOverprint error: " + eStrokeOverprint);
            }
            try {
                targetObj.strokeAlignment = sourceObj.strokeAlignment;
            } catch (eAlign) {
                $.writeln("[OneRect] copy strokeAlignment error: " + eAlign);
            }
        } else {
            targetObj.stroked = false;
        }

        try {
            targetObj.fillOverprint = sourceObj.fillOverprint;
        } catch (eFillOverprint) {
            $.writeln("[OneRect] copy fillOverprint error: " + eFillOverprint);
        }
        try {
            targetObj.opacity = sourceObj.opacity;
        } catch (eOpacity) {
            $.writeln("[OneRect] copy opacity error: " + eOpacity);
        }
        try {
            targetObj.blendingMode = sourceObj.blendingMode;
        } catch (eBlend) {
            $.writeln("[OneRect] copy blendingMode error: " + eBlend);
        }
    }

    /* 属性をコピーした長方形を作成 / Create rectangle and inherit appearance */
    function createMergedRectangleFromSource(bounds, sourceObj) {
        var geomBounds = getGeometryBoundsForRectangle(bounds, sourceObj);
        var rect = doc.activeLayer.pathItems.rectangle(
            geomBounds.top,
            geomBounds.left,
            geomBounds.right - geomBounds.left,
            geomBounds.top - geomBounds.bottom
        );

        copyBasicAppearance(sourceObj, rect);
        return rect;
    }

    /* プレビューを更新する関数 / Update preview */
    function updatePreview(radioButtons, previewCheck) {
        ensureTempLayer();
        while (tempLayer.pageItems.length > 0) {
            tempLayer.pageItems[0].remove();
        }

        var selectedIndex = 0;
        for (var i = 0; i < radioButtons.length; i++) {
            if (radioButtons[i].value) {
                selectedIndex = i;
                break;
            }
        }

        if (typeof previewCheck !== "undefined" && previewCheck.value) {
            /* プレビュー: 合体後の長方形を描画 / Preview merged rectangle */
            var mb = calcMergedBounds();
            var srcObj = objects[selectedIndex];
            var previewGeomBounds = getGeometryBoundsForRectangle(mb, srcObj);
            var previewRect = tempLayer.pathItems.rectangle(
                previewGeomBounds.top,
                previewGeomBounds.left,
                previewGeomBounds.right - previewGeomBounds.left,
                previewGeomBounds.top - previewGeomBounds.bottom
            );

            copyBasicAppearance(srcObj, previewRect);

            /* 一時的に非表示にする（元のhidden状態は originalHiddenStates に保存済み） / Temporarily hide objects */
            for (var j = 0; j < objects.length; j++) {
                objects[j].hidden = true;
            }
        } else {
            /* 通常: 赤枠で選択中のオブジェクトをハイライト / Highlight selected source object */
            drawHighlight(selectedIndex);

            /* 元のオブジェクトの表示状態を復元 / Restore original hidden states */
            for (var j = 0; j < objects.length; j++) {
                objects[j].hidden = originalHiddenStates[j];
            }
        }

        app.redraw();
    }

    /* 一時レイヤーを削除し、オブジェクトの表示を復元する関数 / Remove temp layer and restore visibility */
    function removeTempLayer() {
        try {
            for (var i = 0; i < objects.length; i++) {
                objects[i].hidden = originalHiddenStates[i];
            }
            if (tempLayer) {
                tempLayer.remove();
                tempLayer = null;
            }
            app.redraw();
        } catch (e) {
            $.writeln("[OneRect] removeTempLayer error: " + e);
        }
    }


    var dialogResult = showDialogAndGetOptions();

    /* 一時レイヤーを削除 / Remove temporary layer */
    removeTempLayer();

    /* キャンセル時 / On cancel */
    if (dialogResult.result !== 1) {
        for (var i = 0; i < objects.length; i++) {
            try {
                objects[i].selected = true;
            } catch (eSel) {
                $.writeln("[OneRect] restore selection error: " + eSel);
            }
        }
        return;
    }

    /* バウンディングボックスを計算（全オブジェクト対象） / Calculate bounds for all objects */
    var mb = calcMergedBounds();
    applyMergeResult(mb, dialogResult.sourceIndex, dialogResult.keepOriginals);

    if (dialogResult.showCenter) {
        act_ShowCenter();
    }
}

main();