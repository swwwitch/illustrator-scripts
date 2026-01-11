#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
UI messages support Japanese/English. (Note format remains Japanese for compatibility.)

### スクリプト名：

OutlineTextRestore.jsx

### Readme （GitHub）：

https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/outline/OutlineTextRestore.jsx

### 概要：

- アウトライン化されたテキストを、メモ情報（note）をもとに元のテキストとして復元するスクリプトです。
- 選択しているパス／グループのみを対象とし、復元されたテキストは毎回「新規レイヤー」に作成されます。
- 元のアウトライン（選択オブジェクトのみ）は、新規作成した退避レイヤーに移動され、そのレイヤーを「outlined_text」として最背面に配置し、テンプレートレイヤーに設定します。

### 更新日：

20260111

### 主な機能：

- PathItem または GroupItem の note 情報からテキスト内容、フォント情報、座標などを抽出
- 新しいテキストフレームを元の位置に再生成し、元のアウトラインを非表示または移動
- メモが存在しない・情報が不正な場合には警告を表示
- 復元テキストは毎回「新規レイヤー」に作成／元アウトラインは毎回「新規退避レイヤー」に移動して outlined_text（最背面・テンプレート化）

### 処理の流れ：

1. 対象オブジェクトを選択（アウトライン化された Path または Group）
2. メモ情報を解析し、属性データを抽出
3. 復元テキストを「新規レイヤー」に再構築
4. 元オブジェクトを「outlined_text」レイヤーへ移動（このレイヤーは常に最背面・テンプレート化）

### 更新履歴：

- v1.0 (20240811) : 初期バージョン
- v1.1 (20250721) : 微調整
- v1.2 (20260111) : ローカライズ（アラート文言の英語対応）
- v1.3 (20260111) : 退避レイヤー名を outlined_text に変更（旧名 outlined-text から自動移行）

*/

// スクリプトバージョン
var SCRIPT_VERSION = "v1.3";
function normalizeOutlinedTextLayerNames(doc, keepLayer) {
    if (!doc || !keepLayer) { return; }

    var dupIndex = 1;
    for (var k = 0; k < doc.layers.length; k++) {
        var lyr = doc.layers[k];
        if (lyr !== keepLayer && lyr.name === TEMPLATE_LAYER_NAME) {
            // Unlock/visible to allow rename
            try { lyr.locked = false; } catch (e0) { }
            try { lyr.visible = true; } catch (e1) { }

            var newName = TEMPLATE_LAYER_NAME + "_dup" + dupIndex;
            dupIndex++;
            try {
                lyr.name = newName;
            } catch (e2) {
                // ignore
            }
        }
    }
}

// outlined_text という名前の既存レイヤーがあれば退避（元レイヤーの改名誤爆を防ぐ）
function archiveExistingOutlinedTextLayers(doc) {
    if (!doc) { return; }

    function exists(name) {
        for (var i = 0; i < doc.layers.length; i++) {
            if (doc.layers[i].name === name) { return true; }
        }
        return false;
    }

    var idx = 1;
    for (var i = 0; i < doc.layers.length; i++) {
        var lyr = doc.layers[i];
        if (lyr.name === TEMPLATE_LAYER_NAME) {
            try { lyr.locked = false; } catch (e0) { }
            try { lyr.visible = true; } catch (e1) { }

            var newName;
            do {
                newName = TEMPLATE_LAYER_NAME + "_archive" + idx;
                idx++;
            } while (exists(newName));

            try { lyr.name = newName; } catch (e2) { }
        }
    }
}

// 新規退避レイヤーを作成（このレイヤーに選択パスのみを移動する）
function createNewOutlineStashLayer(doc) {
    if (!doc) { return null; }

    var lyr = doc.layers.add();
    try { lyr.locked = false; } catch (e0) { }
    try { lyr.visible = true; } catch (e1) { }

    // Temporary unique name (will be renamed to outlined_text at finalize)
    lyr.name = "__outlined_text_stash__";

    // Marker group to identify the intended outlined_text layer after action
    try {
        var marker = lyr.groupItems.add();
        marker.name = "__outlined_text_marker__";
    } catch (e2) { }

    return lyr;
}

// act_setLayTmplAttr 実行後に同名 outlined_text ができた場合、意図したレイヤー以外は改名して退避（中身は移動しない）
function cleanupOutlinedTextDuplicateNames(doc) {
    if (!doc) { return null; }

    var target = null;
    // Find the outlined_text layer that contains our marker group
    for (var i = 0; i < doc.layers.length; i++) {
        var lyr = doc.layers[i];
        if (lyr.name !== TEMPLATE_LAYER_NAME) { continue; }
        try {
            for (var g = 0; g < lyr.groupItems.length; g++) {
                if (lyr.groupItems[g].name === "__outlined_text_marker__") {
                    target = lyr;
                    break;
                }
            }
        } catch (e0) { }
        if (target) { break; }
    }

    // If we couldn't find by marker, prefer the bottom-most outlined_text
    if (!target) {
        for (var j = doc.layers.length - 1; j >= 0; j--) {
            if (doc.layers[j].name === TEMPLATE_LAYER_NAME) {
                target = doc.layers[j];
                break;
            }
        }
    }

    // Rename other outlined_text layers to avoid duplicates
    var dupIndex = 1;
    for (var k = 0; k < doc.layers.length; k++) {
        var lyr2 = doc.layers[k];
        if (lyr2 !== target && lyr2.name === TEMPLATE_LAYER_NAME) {
            try { lyr2.locked = false; } catch (e1) { }
            try { lyr2.visible = true; } catch (e2) { }
            try { lyr2.name = TEMPLATE_LAYER_NAME + "_dup" + dupIndex; } catch (e3) { }
            dupIndex++;
        }
    }

    // Remove marker group if present
    if (target) {
        // Template action may lock the layer; unlock temporarily to remove marker
        try { target.locked = false; } catch (eLock) { }
        try { target.visible = true; } catch (eVis) { }

        try {
            for (var gg = target.groupItems.length - 1; gg >= 0; gg--) {
                if (target.groupItems[gg].name === "__outlined_text_marker__") {
                    target.groupItems[gg].remove();
                }
            }
        } catch (e4) { }
    }

    return target;
}


// 退避レイヤー名（新/旧）
var TEMPLATE_LAYER_NAME = "outlined_text";
// Backward compatibility: previous names
var TEMPLATE_LAYER_NAME_OLD_1 = "テンプレート";
var TEMPLATE_LAYER_NAME_OLD_2 = "outlined-text";

// ==============================
// Localization (JP / EN)
// ==============================
var LOCALE = (app.locale && app.locale.indexOf('ja') === 0) ? 'ja' : 'en';

var I18N = {
    ja: {
        ERR_NO_DOC: 'ドキュメントが開かれていません。',
        ERR_NO_SELECTION: 'オブジェクトが選択されていません。',
        ERR_NO_NOTE_PATH: '選択されたパスにはメモが付いていません。',
        ERR_NO_NOTE_GROUP: '選択されたグループにはメモが付いていません。',
        ERR_INVALID_NOTE_PATH: 'メモに有効な属性情報が含まれていません。',
        ERR_INVALID_NOTE_GROUP: 'グループのメモに有効な属性情報が含まれていません。',
        ERR_INVALID_SELECTION: 'Path または Group を選択してください。',
        WARN_FONT_FALLBACK: '指定されたフォントやその他の属性が見つかりませんでした。デフォルトの設定が使用されます。'
    },
    en: {
        ERR_NO_DOC: 'No document is open.',
        ERR_NO_SELECTION: 'No objects are selected.',
        ERR_NO_NOTE_PATH: 'The selected path has no note.',
        ERR_NO_NOTE_GROUP: 'The selected group has no note.',
        ERR_INVALID_NOTE_PATH: 'The note does not contain valid attribute data.',
        ERR_INVALID_NOTE_GROUP: 'The group note does not contain valid attribute data.',
        ERR_INVALID_SELECTION: 'Please select a Path or Group.',
        WARN_FONT_FALLBACK: 'The specified font or attributes were not found. Default settings were used.'
    }
};

function _(key) {
    return (I18N[LOCALE] && I18N[LOCALE][key]) ? I18N[LOCALE][key] : key;
}


// ==============================
// Template Layer (Action-based)
// ==============================

// Apply Template Layer attribute to a named layer (must exist)
function act_setLayTmplAttr() {
    // スクリプトのバージョンとアクションの設定を行う文字列
    var actionString =
        '/version 3' +
        '/name [ 5' +
        ' 6c61796572' +
        ' ]' +
        '/isOpen 1' +
        '/actionCount 1' +
        '/action-1 {' +
        ' /name [ 24' +
        ' 6368616e67652d746f2d74656d706c6174652d6c61796572' +
        ' ]' +
        ' /keyIndex 0' +
        ' /colorIndex 0' +
        ' /isOpen 1' +
        ' /eventCount 1' +
        ' /event-1 {' +
        ' /useRulersIn1stQuadrant 0' +
        ' /internalName (ai_plugin_Layer)' +
        ' /localizedName [ 9' +
        ' e8a1a8e7a4ba203a20' +
        ' ]' +
        ' /isOpen 1' +
        ' /isOn 1' +
        ' /hasDialog 1' +
        ' /showDialog 0' +
        ' /parameterCount 10' +
        ' /parameter-1 {' +
        ' /key 1836411236' +
        ' /showInPalette 4294967295' +
        ' /type (integer)' +
        ' /value 4' +
        ' }' +
        ' /parameter-2 {' +
        ' /key 1851878757' +
        ' /showInPalette 4294967295' +
        ' /type (ustring)' +
        ' /value [ 36' +
        ' e383ace382a4e383a4e383bce38391e3838de383abe382aae38397e382b7e383' +
        ' a7e383b3' +
        ' ]' +
        ' }' +
        ' /parameter-3 {' +
        ' /key 1953068140' +
        ' /showInPalette 4294967295' +
        ' /type (ustring)' +
        ' /value [ 13' +
        ' 6f75746c696e65645f74657874' +
        ' ]' +
        ' }' +
        ' /parameter-4 {' +
        ' /key 1953329260' +
        ' /showInPalette 4294967295' +
        ' /type (boolean)' +
        ' /value 1' +
        ' }' +
        ' /parameter-5 {' +
        ' /key 1936224119' +
        ' /showInPalette 4294967295' +
        ' /type (boolean)' +
        ' /value 1' +
        ' }' +
        ' /parameter-6 {' +
        ' /key 1819239275' +
        ' /showInPalette 4294967295' +
        ' /type (boolean)' +
        ' /value 1' +
        ' }' +
        ' /parameter-7 {' +
        ' /key 1886549623' +
        ' /showInPalette 4294967295' +
        ' /type (boolean)' +
        ' /value 1' +
        ' }' +
        ' /parameter-8 {' +
        ' /key 1886547572' +
        ' /showInPalette 4294967295' +
        ' /type (boolean)' +
        ' /value 0' +
        ' }' +
        ' /parameter-9 {' +
        ' /key 1684630830' +
        ' /showInPalette 4294967295' +
        ' /type (boolean)' +
        ' /value 1' +
        ' }' +
        ' /parameter-10 {' +
        ' /key 1885564532' +
        ' /showInPalette 4294967295' +
        ' /type (unit real)' +
        ' /value 50.0' +
        ' /unit 592474723' +
        ' }' +
        ' }' +
        '}';

    // スクリプトアクションファイルを作成し、アクションを設定
    var actionFile = new File('~/ScriptAction.aia');
    actionFile.encoding = 'UTF-8';
    actionFile.lineFeed = 'Unix';
    actionFile.open('w');
    actionFile.write(actionString);
    actionFile.close();

    // アクションを読み込み、実行し、削除する
    app.loadAction(actionFile);
    actionFile.remove();
    app.doScript("change-to-template-layer", "layer", false); // action name, set name
    app.unloadAction("layer", ""); // set name
}


// メモからテキスト属性（文字列、フォント、フォントサイズなど）を抽出する関数
function extractTextAttributes(noteText) {
    var lines = noteText.split("\n");
    var attributes = {
        text: null,
        font: null,
        fontSize: null,
        leading: null,
        orientation: null,
        tracking: null,
        proportionalMetrics: null,
        x: null, // X座標
        y: null, // Y座標
        savedBounds: null // [L, T, R, B]（geometricBounds）
    };

    for (var i = 0; i < lines.length; i++) {
        if (lines[i].indexOf("文字列：") === 0 && i + 1 < lines.length) {
            attributes.text = lines[i + 1];
        }
        if (lines[i].indexOf("フォント：") === 0 && i + 1 < lines.length) {
            attributes.font = lines[i + 1];
        }
        if (lines[i].indexOf("フォントサイズ：") === 0 && i + 1 < lines.length) {
            attributes.fontSize = parseFloat(lines[i + 1]);
        }
        if (lines[i].indexOf("行送り：") === 0 && i + 1 < lines.length) {
            attributes.leading = parseFloat(lines[i + 1]);
        }
        if (lines[i].indexOf("組み方向：") === 0 && i + 1 < lines.length) {
            attributes.orientation = lines[i + 1];
        }
        if (lines[i].indexOf("トラッキング：") === 0 && i + 1 < lines.length) {
            attributes.tracking = parseFloat(lines[i + 1]);
        }
        if (lines[i].indexOf("プロポーショナルメトリクス：") === 0 && i + 1 < lines.length) {
            attributes.proportionalMetrics = lines[i + 1] === "true";
        }
        if (lines[i].match(/^座標：\s*X\s*=\s*([-]?\d+(\.\d+)?),\s*Y\s*=\s*([-]?\d+(\.\d+)?)/)) {
            attributes.x = parseFloat(RegExp.$1); // X座標
            attributes.y = parseFloat(RegExp.$3); // Y座標
        }
        // 座標（geometricBounds）：L/T/R/B を note から復元（複製時のズレ対策：最優先で使用）
        // 例："L = 12.34, T = 56.78, R = 90.12, B = 34.56"
        if (lines[i].match(/L\s*=\s*([-]?\d+(?:\.\d+)?),\s*T\s*=\s*([-]?\d+(?:\.\d+)?),\s*R\s*=\s*([-]?\d+(?:\.\d+)?),\s*B\s*=\s*([-]?\d+(?:\.\d+)?)/)) {
            attributes.savedBounds = [
                parseFloat(RegExp.$1),
                parseFloat(RegExp.$2),
                parseFloat(RegExp.$3),
                parseFloat(RegExp.$4)
            ];
        }
    }

    return attributes;
}

// 位置合わせ（geometricBounds ベース）
// 参考スクリプトに準拠：
// 1) いったん position で置く
// 2) フォント/サイズ等を適用
// 3) geometricBounds の left/top 差分で translate して見た目位置を揃える
function alignTextFrameToTargetByGeometricBounds(target, textFrame) {
    try {
        // bounds が更新されない環境対策（不要な場合もあります）
        app.redraw();

        var b1;
        // target が [L,T,R,B] 配列のとき
        if (target && target.length === 4 && typeof target[0] === 'number') {
            b1 = target;
        } else if (target && target.geometricBounds) {
            // target が PageItem のとき
            b1 = target.geometricBounds;
        } else {
            return;
        }

        var b2 = textFrame.geometricBounds;  // [left, top, right, bottom]

        var dx = b1[0] - b2[0];
        var dy = b1[1] - b2[1];

        textFrame.translate(dx, dy);
    } catch (e) {
        // 位置合わせに失敗しても処理を止めない
    }
}

// 退避（アウトライン）レイヤー：毎回新規作成して返す
function getOrCreateOutlinedTextLayer(doc) {
    return createNewOutlineStashLayer(doc);
}

// 復元テキスト用：毎回新規レイヤーを作成して返す
function createNewRestoredTextLayer(doc) {
    var baseName = "restored_text";
    var name = baseName;

    function layerNameExists(nm) {
        for (var i = 0; i < doc.layers.length; i++) {
            if (doc.layers[i].name === nm) { return true; }
        }
        return false;
    }

    if (layerNameExists(name)) {
        var idx = 2;
        while (layerNameExists(baseName + "_" + idx)) {
            idx++;
        }
        name = baseName + "_" + idx;
    }

    var lyr = doc.layers.add();
    lyr.name = name;

    // Make sure it's usable
    try { lyr.locked = false; } catch (e) { }
    try { lyr.visible = true; } catch (e2) { }

    return lyr;
}

// Ensure the given layer is actually the active layer (some environments may ignore the assignment)
function setActiveLayerStrict(doc, layer) {
    if (!doc || !layer) { return false; }

    // Try direct assignment
    try { doc.activeLayer = layer; } catch (e0) { }

    // Verify
    try {
        if (doc.activeLayer === layer) { return true; }
    } catch (e1) { }

    // Fallback: try to re-find by name (outlined_text) and activate it
    try {
        for (var i = 0; i < doc.layers.length; i++) {
            if (doc.layers[i] === layer) {
                doc.activeLayer = doc.layers[i];
                break;
            }
        }
    } catch (e2) { }

    try {
        return doc.activeLayer === layer;
    } catch (e3) {
        return false;
    }
}

// 処理完了後にテンプレート属性を付与し、最後にロックする
function finalizeTemplateLayer(outlinedTextLayer) {
    if (!outlinedTextLayer) { return; }

    // Ensure usable state before applying action
    try {
        outlinedTextLayer.locked = false;
    } catch (e) { }
    try {
        outlinedTextLayer.visible = true;
    } catch (e) { }

    // Archive existing outlined_text layers to avoid renaming the original layer by accident
    archiveExistingOutlinedTextLayers(app.activeDocument);

    // Ensure THIS stash layer becomes the only outlined_text
    if (outlinedTextLayer.name !== "outlined_text") {
        outlinedTextLayer.name = "outlined_text";
    }

    // Keep at back
    try { outlinedTextLayer.zOrder(ZOrderMethod.SENDTOBACK); } catch (eBack) { }

    // outlined_text の重複名を再度解消（処理中に増減した場合に備える）
    normalizeOutlinedTextLayerNames(app.activeDocument, outlinedTextLayer);

    // Apply template layer attribute (action-based)
    try {
        // Make it the active layer for safety (action will rename/template the ACTIVE layer)
        var __doc = app.activeDocument;
        var __okActive = setActiveLayerStrict(__doc, outlinedTextLayer);
        if (!__okActive) {
            // If we cannot ensure the target layer is active, do NOT run the action.
            // Otherwise, Illustrator may template/rename another layer (moving unintended artwork).
            return;
        }

        act_setLayTmplAttr();

        // Some environments may create another outlined_text layer when applying template attribute.
        // Keep the intended one (marker) as outlined_text, and rename others (do NOT move artwork).
        outlinedTextLayer = cleanupOutlinedTextDuplicateNames(app.activeDocument) || outlinedTextLayer;
    } catch (e) {
        // ignore
    }

    // Final cleanup: ensure marker is removed from the finalized outlined_text layer
    try { outlinedTextLayer.locked = false; } catch (eUnlockFinal) { }
    try {
        for (var mm = outlinedTextLayer.groupItems.length - 1; mm >= 0; mm--) {
            if (outlinedTextLayer.groupItems[mm].name === "__outlined_text_marker__") {
                outlinedTextLayer.groupItems[mm].remove();
            }
        }
    } catch (eMarkerFinal) { }

    // Final state: keep at back and locked
    try { outlinedTextLayer.zOrder(ZOrderMethod.SENDTOBACK); } catch (e9) { }
    try { outlinedTextLayer.locked = true; } catch (e10) { }
}

// パスオブジェクトを処理する関数
function processPathItem(pathItem, outlinedTextLayer, restoredTextLayer) {
    var noteText = pathItem.note;

    if (noteText) {
        var attributes = extractTextAttributes(noteText);
        if (attributes) {
            var textFrame = createTextFrame(pathItem, attributes, restoredTextLayer);
            alignTextFrameToTargetByGeometricBounds(attributes.savedBounds || pathItem, textFrame);
            moveToOutlinedTextLayer(pathItem, outlinedTextLayer);
        } else {
            alert(_("ERR_INVALID_NOTE_PATH"));
        }
    } else {
        alert(_("ERR_NO_NOTE_PATH"));
    }
}

// グループオブジェクトを処理する関数
function processGroupItem(groupItem, outlinedTextLayer, restoredTextLayer) {
    var noteText = groupItem.note;

    if (noteText) {
        var attributes = extractTextAttributes(noteText);
        if (attributes) {
            var textFrame = createTextFrame(groupItem, attributes, restoredTextLayer);
            alignTextFrameToTargetByGeometricBounds(attributes.savedBounds || groupItem, textFrame);
            groupItem.opacity = 30;
            groupItem.locked = true;
            moveToOutlinedTextLayer(groupItem, outlinedTextLayer);
        } else {
            alert(_("ERR_INVALID_NOTE_GROUP"));
        }
    } else {
        alert(_("ERR_NO_NOTE_GROUP"));
    }
}

// テキストフレームを作成する関数
function createTextFrame(item, attributes, restoredLayer) {
    var targetLayer = restoredLayer || item.layer;
    var textFrame = targetLayer.textFrames.add();
    textFrame.contents = attributes.text;

    // X, Y座標がある場合は使用し、なければ元のオブジェクトの座標を使用
    var posX = (attributes.x !== null) ? attributes.x : item.left;
    var posY = (attributes.y !== null) ? attributes.y : item.top;

    textFrame.position = [posX, posY];

    try {
        var font = app.textFonts.getByName(attributes.font);
        textFrame.textRange.characterAttributes.textFont = font;
        textFrame.textRange.characterAttributes.size = attributes.fontSize;
        textFrame.textRange.characterAttributes.autoLeading = false;
        textFrame.textRange.characterAttributes.leading = attributes.leading;
        textFrame.textRange.characterAttributes.tracking = attributes.tracking;
        textFrame.textRange.characterAttributes.proportionalMetrics = attributes.proportionalMetrics;

        textFrame.orientation = (attributes.orientation === "縦組み") ? TextOrientation.VERTICAL : TextOrientation.HORIZONTAL;
    } catch (e) {
        alert(_("WARN_FONT_FALLBACK"));
    }

    return textFrame;
}

// パスまたはグループを テンプレート（退避）レイヤーに移動する関数
function moveToOutlinedTextLayer(item, outlinedTextLayer) {
    item.moveToBeginning(outlinedTextLayer.groupItems.add());
}

// オブジェクトを再帰的に処理する関数
function processObject(obj, outlinedTextLayer, restoredTextLayer) {
    if (obj.typename == "GroupItem") {
        processGroupItem(obj, outlinedTextLayer, restoredTextLayer);
    } else if (obj.typename == "PathItem") {
        processPathItem(obj, outlinedTextLayer, restoredTextLayer);
    } else {
        alert(_("ERR_INVALID_SELECTION"));
    }
}

function main() {
    if (app.documents.length === 0) {
        alert(_("ERR_NO_DOC"));
        return;
    }
    var doc = app.activeDocument;
    if (app.selection.length === 0) {
        alert(_("ERR_NO_SELECTION"));
        return;
    }
    var selection = app.selection;

    var hasProcessableObject = false;
    for (var i = 0; i < selection.length; i++) {
        var item = selection[i];
        if (item.typename === "GroupItem" || item.typename === "PathItem") {
            hasProcessableObject = true;
            break;
        }
    }

    if (!hasProcessableObject) {
        alert(_("ERR_INVALID_SELECTION"));
        return;
    }

    var outlinedTextLayer = getOrCreateOutlinedTextLayer(doc);
    var restoredTextLayer = createNewRestoredTextLayer(doc);

    for (var i = 0; i < selection.length; i++) {
        var item = selection[i];
        if (item.typename === "GroupItem" || item.typename === "PathItem") {
            processObject(item, outlinedTextLayer, restoredTextLayer);
        }
    }

    finalizeTemplateLayer(outlinedTextLayer);

    // Make restored text layer active (avoid leaving outlined_text as active layer)
    try {
        doc.activeLayer = restoredTextLayer;
    } catch (e) { }
}

main();