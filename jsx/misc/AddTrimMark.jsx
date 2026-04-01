#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
### スクリプト名：

AddTrimMark.jsx

### 概要

- 選択オブジェクト（単純な長方形1点）、現在のアートボード、またはすべてのアートボードを対象に、トンボを作成するIllustrator用スクリプトです。
- 実行時にScriptUIダイアログを表示し、トンボの対象および「ガイドを残す」「日本式トンボ」のON/OFFを選択できます。
- 「選択オブジェクト」を選択した場合は、単純な長方形1点のみを対象とし、対象オブジェクトをトンボレイヤーに複製して〈塗り〉〈線〉をなしにしてからトンボを作成します（元オブジェクトは変更しません）。
- 「現在のアートボード」を選択した場合は、アートボードサイズの矩形を生成してトンボを作成します。
- 「すべてのアートボード」を選択した場合は、アートボードごとに「トンボ_アートボード名」レイヤーを作成し、それぞれにトンボを作成します（同名の場合は「（2）」などでユニーク化されます）。
- 「ガイドを残す」がONのときは対象オブジェクトをガイド化し、OFFのときは対象オブジェクトを削除します。
- トンボは各対象レイヤー上に作成され、処理後はレイヤーのロック状態を元に戻します。
- ダイアログのUIおよびラベルは日本語／英語に自動対応します。

### 主な機能

- 選択オブジェクト（単純な長方形1点）、現在のアートボード、またはすべてのアートボードを対象として選択可能
- 実行時ダイアログで「ガイドを残す」「日本式トンボ」のON/OFFを切り替え
- 「トンボ」レイヤーを自動取得／未存在時は新規作成
- 単純な長方形1点、現在のアートボード矩形1つ、または全アートボード矩形を対象に処理
- 同一オブジェクトでトンボ生成とガイド化、または削除を実行
- 処理後は「トンボ」レイヤーをロック

### 処理の流れ

1. ダイアログボックスでトンボの対象と「ガイドを残す」「日本式トンボ」のON/OFFを選択
2. 「トンボ」レイヤーを取得し、なければ新規作成
3. 環境設定で日本式トンボをONに設定
4. 選択オブジェクトなら、単純な長方形1点のみを「トンボ」レイヤーへ複製し、塗りと線をなしに設定
5. 現在のアートボードなら「トンボ」レイヤー上にアートボード矩形を1つ作成
6. すべてのアートボードならアートボードごとに「トンボ_アートボード名」レイヤーを作成し、矩形を1つずつ作成
7. その対象オブジェクトを元にトリムマーク作成メニューを実行
8. 「ガイドを残す」がONなら同じ対象オブジェクトをガイド化
9. 「ガイドを残す」がOFFなら同じ対象オブジェクトを削除

### 更新日

- 2026-04-01

### 更新履歴

- v1.2.0 (20260401) : すべてのアートボード対応、選択オブジェクト対応、ScriptUIダイアログ追加、ローカライズ対応
- v1.1.0 (20260401) : 選択オブジェクト対応、ScriptUIダイアログ追加、ローカライズ対応
- v1.0 (20250205) : 初期バージョン
*/

// =========================================
// バージョンとローカライズ
// =========================================

var SCRIPT_VERSION = "v1.2.0";

function sanitizeLayerName(name) {
    return name.replace(new RegExp('[\\\\/:*?"<>|]', 'g'), '_');
}

function getOrCreateLayerByName(doc, layerName) {
    var i;
    for (i = 0; i < doc.layers.length; i++) {
        if (doc.layers[i].name === layerName) {
            return doc.layers[i];
        }
    }
    var layer = doc.layers.add();
    layer.name = layerName;
    return layer;
}

function getUniqueLayerName(doc, baseName) {
    var name = baseName;
    var count = 2;
    var exists, i;

    while (true) {
        exists = false;
        for (i = 0; i < doc.layers.length; i++) {
            if (doc.layers[i].name === name) {
                exists = true;
                break;
            }
        }
        if (!exists) {
            return name;
        }
        name = baseName + "（" + count + "）";
        count++;
    }
}

function buildTrimLayerNameForArtboard(doc, artboard) {
    var baseName = "トンボ_" + sanitizeLayerName(artboard.name || "アートボード");
    return getUniqueLayerName(doc, baseName);
}

function createArtboardRectangle(doc, trimLayer, artboard) {
    var rect = artboard.artboardRect;
    doc.activeLayer = trimLayer;
    var targetObj = doc.activeLayer.pathItems.rectangle(rect[1], rect[0], rect[2] - rect[0], rect[1] - rect[3]);
    targetObj.filled = false;
    targetObj.stroked = false;
    return targetObj;
}

function applyTrimMarksToTarget(doc, trimLayer, targetObj, createGuide) {
    doc.activeLayer = trimLayer;
    doc.selection = null;
    doc.selection = [targetObj];
    app.executeMenuCommand('TrimMark v25');

    if (createGuide) {
        targetObj.guides = true;
    } else {
        targetObj.remove();
    }
}

function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

var LABELS = {
    dialogTitle: {
        ja: "トンボ作成",
        en: "Create Trim Marks"
    },
    panelTarget: {
        ja: "トンボの対象",
        en: "Trim Mark Target"
    },
    radioSelection: {
        ja: "選択オブジェクト",
        en: "Selected Object"
    },
    radioCurrentArtboard: {
        ja: "現在のアートボード",
        en: "Current Artboard"
    },
    radioAllArtboards: {
        ja: "すべてのアートボード",
        en: "All Artboards"
    },
    panelOptions: {
        ja: "オプション",
        en: "Options"
    },
    chkGuide: {
        ja: "ガイドを残す",
        en: "Keep Guides"
    },
    chkJapaneseTrim: {
        ja: "日本式トンボ",
        en: "Japanese-style Trim Marks"
    },
    btnCancel: {
        ja: "キャンセル",
        en: "Cancel"
    },
    btnOk: {
        ja: "OK",
        en: "OK"
    }
};

function L(key) {
    return LABELS[key][lang];
}

// =========================================
// Helper functions for rectangle validation
// =========================================
function isNearlyEqual(a, b) {
    return Math.abs(a - b) < 0.01;
}

function isAxisAlignedRectangle(pathItem) {
    var pts, xs, ys, i, left, right, top, bottom, anchors;

    if (pathItem.pathPoints.length !== 4 || !pathItem.closed) {
        return false;
    }

    pts = pathItem.pathPoints;
    xs = [];
    ys = [];

    for (i = 0; i < 4; i++) {
        xs.push(pts[i].anchor[0]);
        ys.push(pts[i].anchor[1]);

        if (!isNearlyEqual(pts[i].leftDirection[0], pts[i].anchor[0]) ||
            !isNearlyEqual(pts[i].leftDirection[1], pts[i].anchor[1]) ||
            !isNearlyEqual(pts[i].rightDirection[0], pts[i].anchor[0]) ||
            !isNearlyEqual(pts[i].rightDirection[1], pts[i].anchor[1])) {
            return false;
        }
    }

    left = Math.min.apply(null, xs);
    right = Math.max.apply(null, xs);
    top = Math.max.apply(null, ys);
    bottom = Math.min.apply(null, ys);

    anchors = {};
    for (i = 0; i < 4; i++) {
        if (!isNearlyEqual(xs[i], left) && !isNearlyEqual(xs[i], right)) {
            return false;
        }
        if (!isNearlyEqual(ys[i], top) && !isNearlyEqual(ys[i], bottom)) {
            return false;
        }
        anchors[Math.round(xs[i] * 1000) + ',' + Math.round(ys[i] * 1000)] = true;
    }

    return anchors[Math.round(left * 1000) + ',' + Math.round(top * 1000)] &&
        anchors[Math.round(right * 1000) + ',' + Math.round(top * 1000)] &&
        anchors[Math.round(right * 1000) + ',' + Math.round(bottom * 1000)] &&
        anchors[Math.round(left * 1000) + ',' + Math.round(bottom * 1000)];
}

// =========================================
// ScriptUIダイアログ
// =========================================
function showOptionsDialog() {
    var dlg = new Window('dialog', L('dialogTitle') + ' ' + SCRIPT_VERSION);
    var pnlTarget = dlg.add('panel', undefined, L('panelTarget'));
    var rdoSelection = pnlTarget.add('radiobutton', undefined, L('radioSelection'));
    var rdoCurrentArtboard = pnlTarget.add('radiobutton', undefined, L('radioCurrentArtboard'));
    var rdoAllArtboards = pnlTarget.add('radiobutton', undefined, L('radioAllArtboards'));
    var pnlOptions = dlg.add('panel', undefined, L('panelOptions'));
    var chkGuide = pnlOptions.add('checkbox', undefined, L('chkGuide'));
    var chkJapanese = pnlOptions.add('checkbox', undefined, L('chkJapaneseTrim'));
    var btnGroup = dlg.add('group');
    var btnCancel = btnGroup.add('button', undefined, L('btnCancel'), { name: 'cancel' });
    var btnOk = btnGroup.add('button', undefined, L('btnOk'), { name: 'ok' });

    if (app.activeDocument.selection.length === 0) {
        rdoCurrentArtboard.value = true;
    } else if (app.activeDocument.selection.length === 1 &&
        app.activeDocument.selection[0].typename === 'PathItem' &&
        !app.activeDocument.selection[0].guides &&
        !app.activeDocument.selection[0].clipping &&
        isAxisAlignedRectangle(app.activeDocument.selection[0])) {
        rdoSelection.value = true;
    } else {
        rdoCurrentArtboard.value = true;
    }
    chkGuide.value = true;
    chkJapanese.value = true;

    pnlTarget.orientation = 'column';
    pnlTarget.alignChildren = 'left';
    pnlTarget.alignment = 'fill';
    pnlTarget.margins = [15, 20, 15, 10];
    pnlOptions.orientation = 'column';
    pnlOptions.alignChildren = 'left';
    pnlOptions.alignment = 'fill';
    pnlOptions.margins = [15, 20, 15, 10];
    btnGroup.alignment = 'right';
    btnCancel.preferredSize.width = 80;
    btnOk.preferredSize.width = 80;

    if (dlg.show() !== 1) {
        return null;
    }

    return {
        targetType: rdoSelection.value ? 'selection' : (rdoAllArtboards.value ? 'allArtboards' : 'artboard'),
        createGuide: chkGuide.value,
        japaneseStyle: chkJapanese.value
    };
}

function getSingleSelectedObject(doc) {
    var targetObj;

    if (doc.selection.length !== 1) {
        alert(lang === 'ja'
            ? '「選択オブジェクト」を使うには、単純な長方形を1つだけ選択してください。'
            : 'To use "Selected Object", select exactly one simple rectangle.');
        return null;
    }

    targetObj = doc.selection[0];

    if (targetObj.typename !== 'PathItem' || targetObj.guides || targetObj.clipping) {
        alert(lang === 'ja'
            ? '「選択オブジェクト」で使えるのは、単純な長方形だけです。'
            : 'Only a simple rectangle can be used for "Selected Object".');
        return null;
    }

    if (!isAxisAlignedRectangle(targetObj)) {
        alert(lang === 'ja'
            ? '「選択オブジェクト」で使えるのは、各辺が水平・垂直で、4点が直交している単純な長方形だけです。'
            : 'Only a simple rectangle with horizontal/vertical edges and right-angle corners can be used for "Selected Object".');
        return null;
    }

    return targetObj;
}

function prepareSelectedObjectForTrim(targetObj, trimLayer) {
    var duplicatedObj = targetObj.duplicate(trimLayer, ElementPlacement.PLACEATBEGINNING);
    if ('filled' in duplicatedObj) {
        duplicatedObj.filled = false;
    }
    if ('stroked' in duplicatedObj) {
        duplicatedObj.stroked = false;
    }
    return duplicatedObj;
}

function main() {
    var doc = app.activeDocument;
    var options = showOptionsDialog();
    var targetObj = null;
    var trimLayer = null;
    var trimLayers = [];
    var trimLayerLockStates = [];
    var wasLocked = false;
    var i;

    if (!options) {
        return;
    }

    if (options.targetType !== 'allArtboards') {
        /* 「トンボ」レイヤーを取得（なければ作成） / Get the "Trim" layer, or create it if missing */
        trimLayer = getOrCreateLayerByName(doc, "トンボ");
        wasLocked = trimLayer.locked;
        if (wasLocked) {
            trimLayer.locked = false;
        }
    }

    try {

        /* 日本式トンボの設定を切り替える / Set Japanese-style trim marks */
        app.preferences.setBooleanPreference('cropMarkStyle', options.japaneseStyle ? 1 : 0);

        if (options.targetType === 'selection') {
            /* 選択オブジェクトを取得 / Get the selected object */
            targetObj = getSingleSelectedObject(doc);
            if (!targetObj) {
                return;
            }

            /* 選択オブジェクトを複製してトンボ用に準備 / Duplicate the selected object for trim marks */
            targetObj = prepareSelectedObjectForTrim(targetObj, trimLayer);
            applyTrimMarksToTarget(doc, trimLayer, targetObj, options.createGuide);
        } else if (options.targetType === 'allArtboards') {
            for (i = 0; i < doc.artboards.length; i++) {
                trimLayer = getOrCreateLayerByName(doc, buildTrimLayerNameForArtboard(doc, doc.artboards[i]));
                trimLayers.push(trimLayer);
                trimLayerLockStates.push(trimLayer.locked);
                if (trimLayer.locked) {
                    trimLayer.locked = false;
                }
                targetObj = createArtboardRectangle(doc, trimLayer, doc.artboards[i]);
                applyTrimMarksToTarget(doc, trimLayer, targetObj, options.createGuide);
            }
        } else {
            /* アクティブなアートボードを取得 / Get the active artboard */
            var artboard = doc.artboards[doc.artboards.getActiveArtboardIndex()];

            /* アートボード矩形を作成 / Create the artboard rectangle */
            targetObj = createArtboardRectangle(doc, trimLayer, artboard);
            applyTrimMarksToTarget(doc, trimLayer, targetObj, options.createGuide);
        }

    } finally {
        doc.selection = null;

        if (options.targetType === 'allArtboards') {
            for (i = 0; i < trimLayers.length; i++) {
                trimLayers[i].locked = trimLayerLockStates[i];
            }
        } else if (trimLayer) {
            trimLayer.locked = wasLocked;
        }
    }
}

main();