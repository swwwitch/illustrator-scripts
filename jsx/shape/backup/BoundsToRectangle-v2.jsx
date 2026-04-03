#target illustrator
#targetengine "session"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

var SCRIPT_VERSION = "v1.2";

/*
 * ひとつの長方形に変換.jsx / BoundsToRectangle.jsx
 *
 * 概要 / Overview:
 * 選択した複数オブジェクトの外接矩形をもとに、新しい長方形を作成してひとつに統合します。
 * 属性の引き継ぎ元を選択でき、プレビュー境界 / オブジェクト境界の切り替え、
 * 元の図形を残す、プレビュー表示に対応しています。
 * また、ダイアログボックスの位置をセッション中に記憶し、次回表示時に復元します。
 * Creates a new rectangle from the overall bounds of multiple selected objects and merges them into one result.
 * You can choose which object to inherit attributes from, switch between preview bounds
 * and geometric bounds, keep the original objects, and use preview display.
 * The dialog position is also remembered during the session and restored next time.
 *
 * 更新日 / Updated: 2026-03-07
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
    keepOriginals: {
        ja: "元の図形を残す",
        en: "Keep Original Objects"
    },
    preview: {
        ja: "プレビュー",
        en: "Preview"
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
    },
    rowPanel: {
        ja: "行",
        en: "Rows"
    },
    colPanel: {
        ja: "列",
        en: "Columns"
    },
    stepCount: {
        ja: "段数:",
        en: "Count:"
    },
    gapLabel: {
        ja: "間隔:",
        en: "Gap:"
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
            rb.onClick = function () { refreshPreview(); };
            radioButtons.push(rb);
        }

        /* 行・列パネル（2カラム） / Row & Column panels (2-column layout) */
        var gridGroup = dlg.add("group");
        gridGroup.alignment = ["fill", "top"];
        gridGroup.alignChildren = ["fill", "top"];

        var rowPanel = gridGroup.add("panel", undefined, L("rowPanel"));
        rowPanel.alignChildren = ["left", "center"];
        rowPanel.margins = [15, 20, 15, 10];
        var rowCountGroup = rowPanel.add("group");
        var rowCountCheck = rowCountGroup.add("checkbox", undefined, "");
        rowCountCheck.value = false;
        var rowCountLabel = rowCountGroup.add("statictext", undefined, L("stepCount"));
        var rowCountInput = rowCountGroup.add("edittext", undefined, "1");
        rowCountInput.characters = 4;
        rowCountInput.enabled = false;
        var rowGapGroup = rowPanel.add("group");
        var rowGapCheck = rowGapGroup.add("checkbox", undefined, "");
        rowGapCheck.value = false;
        var rowGapLabel = rowGapGroup.add("statictext", undefined, L("gapLabel"));
        var rowGapInput = rowGapGroup.add("edittext", undefined, "0");
        rowGapInput.characters = 4;
        rowGapInput.enabled = false;

        /* ラベルクリックでチェック切替 / Toggle checkbox by clicking label */
        function bindLabelToCheck(label, check, input, defaultVal) {
            label.addEventListener("click", function () {
                check.value = !check.value;
                input.enabled = check.value;
                if (check.value && defaultVal !== undefined && parseFloat(input.text) < defaultVal) {
                    input.text = defaultVal;
                }
                refreshPreview();
            });
        }

        /* 段数チェックON時に値が1以下なら2にする / Set count to 2 when enabling */
        function onCountCheckClick(check, input) {
            input.enabled = check.value;
            if (check.value && parseFloat(input.text) < 2) {
                input.text = "2";
            }
            refreshPreview();
        }

        /* 間隔チェックON時に値が0なら10にする / Set gap to 10 when enabling */
        function onGapCheckClick(check, input) {
            input.enabled = check.value;
            if (check.value && parseFloat(input.text) === 0) {
                input.text = "10";
            }
            refreshPreview();
        }

        bindLabelToCheck(rowCountLabel, rowCountCheck, rowCountInput, 2);
        bindLabelToCheck(rowGapLabel, rowGapCheck, rowGapInput, 10);
        rowCountCheck.onClick = function () {
            onCountCheckClick(rowCountCheck, rowCountInput);
        };
        rowGapCheck.onClick = function () {
            onGapCheckClick(rowGapCheck, rowGapInput);
        };

        var colPanel = gridGroup.add("panel", undefined, L("colPanel"));
        colPanel.alignChildren = ["left", "center"];
        colPanel.margins = [15, 20, 15, 10];
        var colCountGroup = colPanel.add("group");
        var colCountCheck = colCountGroup.add("checkbox", undefined, "");
        colCountCheck.value = false;
        var colCountLabel = colCountGroup.add("statictext", undefined, L("stepCount"));
        var colCountInput = colCountGroup.add("edittext", undefined, "1");
        colCountInput.characters = 4;
        colCountInput.enabled = false;
        var colGapGroup = colPanel.add("group");
        var colGapCheck = colGapGroup.add("checkbox", undefined, "");
        colGapCheck.value = false;
        var colGapLabel = colGapGroup.add("statictext", undefined, L("gapLabel"));
        var colGapInput = colGapGroup.add("edittext", undefined, "0");
        colGapInput.characters = 4;
        colGapInput.enabled = false;
        bindLabelToCheck(colCountLabel, colCountCheck, colCountInput, 2);
        bindLabelToCheck(colGapLabel, colGapCheck, colGapInput, 10);
        colCountCheck.onClick = function () {
            onCountCheckClick(colCountCheck, colCountInput);
        };
        colGapCheck.onClick = function () {
            onGapCheckClick(colGapCheck, colGapInput);
        };

        var optionPanel = dlg.add("panel", undefined, L("optionsPanel"));
        optionPanel.alignment = ["fill", "fill"];
        optionPanel.alignChildren = ["left", "top"];
        optionPanel.margins = [15, 20, 15, 10];

        var useVisibleBounds = optionPanel.add("checkbox", undefined, L("useVisibleBounds"));
        useVisibleBounds.value = true;
        useVisibleBounds.onClick = function () {
            activeBoundsData = useVisibleBounds.value ? visibleBoundsData : geometricBoundsData;
            refreshPreview();
        };

        var keepOriginals = optionPanel.add("checkbox", undefined, L("keepOriginals"));
        keepOriginals.value = false;

        /* ↑↓キーで値を増減 / Change value by arrow keys */
        function changeValueByArrowKey(editText, allowNegative) {
            editText.addEventListener("keydown", function (event) {
                var value = Number(editText.text);
                if (isNaN(value)) return;

                var keyboard = ScriptUI.environment.keyboardState;
                var delta = 1;

                if (keyboard.shiftKey) {
                    delta = 10;
                    if (event.keyName == "Up") {
                        value = Math.ceil((value + 1) / delta) * delta;
                        event.preventDefault();
                    } else if (event.keyName == "Down") {
                        value = Math.floor((value - 1) / delta) * delta;
                        event.preventDefault();
                    }
                } else if (keyboard.altKey) {
                    delta = 0.1;
                    if (event.keyName == "Up") {
                        value += delta;
                        event.preventDefault();
                    } else if (event.keyName == "Down") {
                        value -= delta;
                        event.preventDefault();
                    }
                } else {
                    if (event.keyName == "Up") {
                        value += delta;
                        event.preventDefault();
                    } else if (event.keyName == "Down") {
                        value -= delta;
                        event.preventDefault();
                    }
                }

                if (keyboard.altKey) {
                    value = Math.round(value * 10) / 10;
                } else {
                    value = Math.round(value);
                }

                if (!allowNegative && value < 0) value = 0;

                editText.text = value;
                refreshPreview();
            });
        }

        changeValueByArrowKey(rowCountInput, false);
        changeValueByArrowKey(rowGapInput, true);
        changeValueByArrowKey(colCountInput, false);
        changeValueByArrowKey(colGapInput, true);

        /* グリッド入力値を読み取ってプレビュー更新するラッパー / Wrapper that reads grid inputs */
        function refreshPreview() {
            var gc = {
                rowCount: rowCountCheck.value ? Math.max(1, Math.round(parseFloat(rowCountInput.text) || 1)) : 1,
                colCount: colCountCheck.value ? Math.max(1, Math.round(parseFloat(colCountInput.text) || 1)) : 1,
                rowGap: rowGapCheck.value ? (parseFloat(rowGapInput.text) || 0) : 0,
                colGap: colGapCheck.value ? (parseFloat(colGapInput.text) || 0) : 0
            };
            updatePreview(radioButtons, previewCheck, gc);
        }

        /* 各入力のonChangeをラッパーに接続 / Connect onChange handlers */
        rowCountInput.onChange = refreshPreview;
        rowGapInput.onChange = refreshPreview;
        colCountInput.onChange = refreshPreview;
        colGapInput.onChange = refreshPreview;

        var bottomGroup = dlg.add("group");
        bottomGroup.alignment = ["fill", "top"];
        bottomGroup.alignChildren = ["center", "center"];

        var previewCheck = bottomGroup.add("checkbox", undefined, L("preview"));
        previewCheck.value = false;
        previewCheck.onClick = function () { refreshPreview(); };

        var btnGroup = dlg.add("group");
        btnGroup.alignment = ["center", "top"];
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
                    refreshPreview();
                    e.preventDefault();
                    break;
                }
            }
        });

        /* 初期プレビューを描画 / Draw initial preview */
        refreshPreview();

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

        var rowCount = rowCountCheck.value ? Math.max(1, Math.round(parseFloat(rowCountInput.text) || 1)) : 1;
        var colCount = colCountCheck.value ? Math.max(1, Math.round(parseFloat(colCountInput.text) || 1)) : 1;
        var rowGap = rowGapCheck.value ? (parseFloat(rowGapInput.text) || 0) : 0;
        var colGap = colGapCheck.value ? (parseFloat(colGapInput.text) || 0) : 0;

        return {
            result: result,
            sourceIndex: sourceIndex,
            keepOriginals: keepOriginals.value,
            rowCount: rowCount,
            colCount: colCount,
            rowGap: rowGap,
            colGap: colGap
        };
    }

    /* 結果を適用 / Apply final result */
    function applyMergeResult(bounds, sourceIndex, keepOriginals, rowCount, colCount, rowGap, colGap) {
        /* 削除前に属性をキャッシュ / Cache attributes before removing originals */
        var baseObj = objects[sourceIndex];
        var srcAttrs = {
            filled: baseObj.filled,
            fillColor: baseObj.filled ? baseObj.fillColor : null,
            stroked: baseObj.stroked,
            strokeColor: baseObj.stroked ? baseObj.strokeColor : null,
            strokeWidth: baseObj.stroked ? baseObj.strokeWidth : 0,
            opacity: baseObj.opacity
        };

        if (!keepOriginals) {
            for (var j = objects.length - 1; j >= 0; j--) {
                objects[j].remove();
            }
        }

        /* グリッド分割で長方形を作成 / Create grid of rectangles */
        var totalWidth = bounds.right - bounds.left;
        var totalHeight = bounds.top - bounds.bottom;
        var cellWidth = (totalWidth - colGap * (colCount - 1)) / colCount;
        var cellHeight = (totalHeight - rowGap * (rowCount - 1)) / rowCount;

        for (var r = 0; r < rowCount; r++) {
            for (var c = 0; c < colCount; c++) {
                var cellLeft = bounds.left + c * (cellWidth + colGap);
                var cellTop = bounds.top - r * (cellHeight + rowGap);
                var rect = doc.activeLayer.pathItems.rectangle(
                    cellTop, cellLeft, cellWidth, cellHeight
                );
                if (srcAttrs.filled) {
                    rect.filled = true;
                    rect.fillColor = srcAttrs.fillColor;
                } else {
                    rect.filled = false;
                }
                if (srcAttrs.stroked) {
                    rect.stroked = true;
                    rect.strokeColor = srcAttrs.strokeColor;
                    rect.strokeWidth = srcAttrs.strokeWidth;
                } else {
                    rect.stroked = false;
                }
                rect.opacity = srcAttrs.opacity;
                rect.selected = true;
            }
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


    /* 使用するboundsデータ（初期値はvisibleBounds） / Active bounds data (default: visibleBounds) */
    var activeBoundsData = visibleBoundsData;

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

    /* プレビューを更新する関数 / Update preview
     * gridConfig: { rowCount, colCount, rowGap, colGap } */
    function updatePreview(radioButtons, previewCheck, gridConfig) {
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

        var gc = gridConfig || {};
        var pRowCount = gc.rowCount || 1;
        var pColCount = gc.colCount || 1;
        var pRowGap = gc.rowGap || 0;
        var pColGap = gc.colGap || 0;

        if (typeof previewCheck !== "undefined" && previewCheck.value) {
            /* プレビュー: グリッド分割で長方形を描画 / Preview grid of rectangles */
            var mb = calcMergedBounds();
            var srcObj = objects[selectedIndex];

            var totalW = mb.right - mb.left;
            var totalH = mb.top - mb.bottom;
            var cellW = (totalW - pColGap * (pColCount - 1)) / pColCount;
            var cellH = (totalH - pRowGap * (pRowCount - 1)) / pRowCount;

            for (var r = 0; r < pRowCount; r++) {
                for (var c = 0; c < pColCount; c++) {
                    var cx = mb.left + c * (cellW + pColGap);
                    var cy = mb.top - r * (cellH + pRowGap);
                    var previewRect = tempLayer.pathItems.rectangle(cy, cx, cellW, cellH);

                    if (srcObj.filled) {
                        previewRect.filled = true;
                        previewRect.fillColor = srcObj.fillColor;
                    } else {
                        previewRect.filled = false;
                    }
                    if (srcObj.stroked) {
                        previewRect.stroked = true;
                        previewRect.strokeColor = srcObj.strokeColor;
                        previewRect.strokeWidth = srcObj.strokeWidth;
                    } else {
                        previewRect.stroked = false;
                    }
                    previewRect.opacity = srcObj.opacity;
                }
            }

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
            objects[i].selected = true;
        }
        return;
    }

    /* バウンディングボックスを計算（全オブジェクト対象） / Calculate bounds for all objects */
    var mb = calcMergedBounds();
    applyMergeResult(mb, dialogResult.sourceIndex, dialogResult.keepOriginals, dialogResult.rowCount, dialogResult.colCount, dialogResult.rowGap, dialogResult.colGap);
}

main();