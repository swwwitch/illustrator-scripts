#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
This script supports only Japanese language.

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

*/

// スクリプトバージョン
var SCRIPT_VERSION = "v1.1";


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
        y: null // Y座標
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
    }

    return attributes;
}

// テキストフレームの位置を調整する関数
function adjustTextFramePosition(textFrame, attributes) {
    textFrame.top += attributes.fontSize / 50; // 文字サイズの2%下に移動
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
            adjustTextFramePosition(textFrame, attributes);
            moveToOutlinedTextLayer(pathItem, outlinedTextLayer);
        } else {
            alert("メモに有効な属性情報が含まれていません。");
        }
    } else {
        alert("選択されたパスにはメモが付いていません。");
    }
}

// グループオブジェクトを処理する関数
function processGroupItem(groupItem, outlinedTextLayer) {
    var noteText = groupItem.note;

    if (noteText) {
        var attributes = extractTextAttributes(noteText);
        if (attributes) {
            var textFrame = createTextFrame(groupItem, attributes);
            adjustTextFramePosition(textFrame, attributes);
            groupItem.opacity = 30;
            groupItem.locked = true;
            moveToOutlinedTextLayer(groupItem, outlinedTextLayer);
        } else {
            alert("グループのメモに有効な属性情報が含まれていません。");
        }
    } else {
        alert("選択されたグループにはメモが付いていません。");
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
        alert("指定されたフォントやその他の属性が見つかりませんでした。デフォルトの設定が使用されます。");
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
        alert("選択されたオブジェクトはパスまたはグループではありません。");
    }
}

function main() {
    var doc = app.activeDocument;
    if (app.selection.length === 0) {
        alert("オブジェクトが選択されていません。");
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
        alert("Path または Group を選択してください。");
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