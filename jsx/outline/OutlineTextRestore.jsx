#target illustrator
app.preferences.setBooleanPreference("ShowExternalJSXWarning", false);

/*

### スクリプト名：

OutlineTextRestore.jsx

### Readme （GitHub）：

https://github.com/swwwitch/illustrator-scripts

### 概要：

- アウトライン化されたパスまたはグループの note 情報をもとに、元のテキストを復元するスクリプト
- フォント、サイズ、行送り、トラッキング、組み方向、座標などを再現可能

### 主な機能：

- note に含まれるテキスト情報を抽出
- テキストフレームを作成し属性を適用
- 元のアウトラインオブジェクトは非表示処理

### 処理の流れ：

1. 選択オブジェクトを判定（PathItem または GroupItem）
2. note から属性情報を抽出
3. テキストフレームを復元・位置調整
4. 元オブジェクトを outlined-text レイヤーへ移動

### 更新履歴：

- v1.0 (20240721) : 初期バージョン
- v1.1 (20250721) : 日本語/英語のローカライズ対応

---

### Script Name:

OutlineTextRestore.jsx

### Readme (GitHub):

https://github.com/swwwitch/illustrator-scripts

### Description:

- Restores original text from outlined objects using metadata in note
- Supports restoration of font, size, leading, tracking, orientation, and position

### Features:

- Extracts text and style info from note field
- Recreates a text frame and applies attributes
- Moves original outline to a dedicated layer

### Workflow:

1. Identify selected object (PathItem or GroupItem)
2. Extract attributes from note
3. Create and position text frame
4. Move original object to "outlined-text" layer

### Update History:

- v1.0 (20240721): Initial release
- v1.1 (20250721): Added Japanese/English localization support

*/
var SCRIPT_VERSION = "v1.1";

// ロケール取得関数 / Get current locale
function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

/* 日英ラベル定義 / Japanese-English label definitions */
var LABELS = {
    noSelection: {
        ja: "オブジェクトが選択されていません。",
        en: "No object is selected."
    },
    notPathOrGroup: {
        ja: "選択されたオブジェクトはパスまたはグループではありません。",
        en: "The selected object is not a path or group."
    },
    noNoteOnPath: {
        ja: "選択されたパスにはメモが付いていません。",
        en: "The selected path has no note."
    },
    noNoteOnGroup: {
        ja: "選択されたグループにはメモが付いていません。",
        en: "The selected group has no note."
    },
    invalidAttributesInNote: {
        ja: "メモに有効な属性情報が含まれていません。",
        en: "The note does not contain valid attribute information."
    },
    invalidAttributesInGroupNote: {
        ja: "グループのメモに有効な属性情報が含まれていません。",
        en: "The group note does not contain valid attribute information."
    },
    fallbackFont: {
        ja: "指定されたフォントやその他の属性が見つかりませんでした。デフォルトの設定が使用されます。",
        en: "Specified font or other attributes were not found. Default settings will be used."
    }
};

// メモからテキスト属性（文字列、フォント、フォントサイズなど）を抽出する関数 / Extract text attributes (string, font, font size, etc.) from note
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
        x: null, // X座標 / X coordinate
        y: null  // Y座標 / Y coordinate
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
            attributes.x = parseFloat(RegExp.$1); // X座標 / X coordinate
            attributes.y = parseFloat(RegExp.$3); // Y座標 / Y coordinate
        }
    }

    return attributes;
}

// テキストフレームの位置を調整する関数 / Adjust text frame position
function adjustTextFramePosition(textFrame, attributes) {
    textFrame.top += attributes.fontSize / 50; // 文字サイズの2%下に移動 / Move down by 2% of font size
}

// outlined-textレイヤーを取得または作成し、ロック解除および表示する関数 / Get or create outlined-text layer, unlock and make visible
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

// パスオブジェクトを処理する関数 / Process path object
function processPathItem(pathItem, outlinedTextLayer) {
    var noteText = pathItem.note;

    if (noteText) {
        var attributes = extractTextAttributes(noteText);
        if (attributes) {
            var textFrame = createTextFrame(pathItem, attributes);
            adjustTextFramePosition(textFrame, attributes);
            moveToOutlinedTextLayer(pathItem, outlinedTextLayer);
        } else {
            alert(LABELS.invalidAttributesInNote[lang]);
        }
    } else {
        alert(LABELS.noNoteOnPath[lang]);
    }
}

// グループオブジェクトを処理する関数 / Process group object
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
            alert(LABELS.invalidAttributesInGroupNote[lang]);
        }
    } else {
        alert(LABELS.noNoteOnGroup[lang]);
    }
}

// テキストフレームを作成する関数 / Create text frame
function createTextFrame(item, attributes) {
    var originalLayer = item.layer;
    var textFrame = originalLayer.textFrames.add();
    textFrame.contents = attributes.text;

    // X, Y座標がある場合は使用し、なければ元のオブジェクトの座標を使用 / Use X, Y coordinates if available, otherwise use original object's coordinates
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
        alert(LABELS.fallbackFont[lang]);
    }

    return textFrame;
}

// パスまたはグループを outlined-text レイヤーに移動する関数 / Move path or group to outlined-text layer
function moveToOutlinedTextLayer(item, outlinedTextLayer) {
    item.moveToBeginning(outlinedTextLayer.groupItems.add());
}

// オブジェクトを再帰的に処理する関数 / Recursively process object
function processObject(obj, outlinedTextLayer) {
    if (obj.typename == "GroupItem") {
        processGroupItem(obj, outlinedTextLayer);
    } else if (obj.typename == "PathItem") {
        processPathItem(obj, outlinedTextLayer);
    } else {
        alert(LABELS.notPathOrGroup[lang]);
    }
}

// メイン処理 / Main process
function main() {
    var doc = app.activeDocument;
    var selection = doc.selection;

    if (selection.length > 0) {
        var outlinedTextLayer = getOrCreateOutlinedTextLayer(doc);
        for (var i = 0; i < selection.length; i++) {
            processObject(selection[i], outlinedTextLayer);
        }
    } else {
        alert(LABELS.noSelection[lang]);
    }
}

main();