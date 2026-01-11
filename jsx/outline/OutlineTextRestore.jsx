#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
UI messages support Japanese/English. (Note format remains Japanese for compatibility.)

### スクリプト名：

OutlineTextRestore.jsx

### Readme （GitHub）：

https://github.com/swwwitch/illustrator-scripts

### 概要：

- アウトライン化されたテキストの復元を支援するスクリプト
- オブジェクトに埋め込まれたメモ情報から、元のテキストとスタイルを再構築

### 主な機能：

- PathItem または GroupItem の note 情報からテキスト内容、フォント情報、座標などを抽出
- 新しいテキストフレームを元の位置に再生成し、元のアウトラインを非表示または移動
- メモが存在しない・情報が不正な場合には警告を表示

### 処理の流れ：

1. 対象オブジェクトを選択（アウトライン化された Path または Group）
2. メモ情報を解析し、属性データを抽出
3. 指定位置にテキストフレームを再構築
4. 元オブジェクトを「outlined-text」レイヤーへ移動または透明化

### 更新履歴：

- v1.0 (20240811) : 初期バージョン
- v1.1 (20250721) : 微調整
- v1.2 (20260111) : ローカライズ（アラート文言の英語対応）

*/

// スクリプトバージョン
var SCRIPT_VERSION = "v1.2";

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

// outlined-textレイヤーを取得または作成し、ロック解除および表示する関数
function getOrCreateOutlinedTextLayer(doc) {
    var outlinedTextLayer = null;

    for (var i = 0; i < doc.layers.length; i++) {
        if (doc.layers[i].name === "outlined-text") {
            outlinedTextLayer = doc.layers[i];
            break;
        }
    }

    if (!outlinedTextLayer) {
        outlinedTextLayer = doc.layers.add();
        outlinedTextLayer.name = "outlined-text";
    }

    if (outlinedTextLayer.locked) {
        outlinedTextLayer.locked = false;
    }

    if (!outlinedTextLayer.visible) {
        outlinedTextLayer.visible = true;
    }

    return outlinedTextLayer;
}

// パスオブジェクトを処理する関数
function processPathItem(pathItem, outlinedTextLayer) {
    var noteText = pathItem.note;

    if (noteText) {
        var attributes = extractTextAttributes(noteText);
        if (attributes) {
            var textFrame = createTextFrame(pathItem, attributes);
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
function processGroupItem(groupItem, outlinedTextLayer) {
    var noteText = groupItem.note;

    if (noteText) {
        var attributes = extractTextAttributes(noteText);
        if (attributes) {
            var textFrame = createTextFrame(groupItem, attributes);
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
function createTextFrame(item, attributes) {
    var originalLayer = item.layer;
    var textFrame = originalLayer.textFrames.add();
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

// パスまたはグループを outlined-text レイヤーに移動する関数
function moveToOutlinedTextLayer(item, outlinedTextLayer) {
    item.moveToBeginning(outlinedTextLayer.groupItems.add());
}

// オブジェクトを再帰的に処理する関数
function processObject(obj, outlinedTextLayer) {
    if (obj.typename == "GroupItem") {
        processGroupItem(obj, outlinedTextLayer);
    } else if (obj.typename == "PathItem") {
        processPathItem(obj, outlinedTextLayer);
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

    for (var i = 0; i < selection.length; i++) {
        var item = selection[i];
        if (item.typename === "GroupItem" || item.typename === "PathItem") {
            processObject(item, outlinedTextLayer);
        }
    }
}

main();