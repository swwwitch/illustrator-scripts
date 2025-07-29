#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
スクリプト名：現在のアートボードサイズの長方形作成 / Create Rectangle Same as Artboard

概要 / Overview:
- 現在のアートボードと同じ大きさの長方形を作成
- CMYKドキュメントの場合は塗りをK20に設定
- RGBドキュメントの場合は塗りを#999999に設定
- 「bg-template」レイヤーを作成し、テンプレート化＆最背面に移動

更新履歴 / Update History:
- v1.0 (20250729) : 初期バージョン
*/

var SCRIPT_VERSION = "v1.0";

function getCurrentLang() {
  return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

// UIラベル定義 / UI Label Definitions
var LABELS = {
    dialogTitle: {
        ja: "背景色レイヤーを作成 " + SCRIPT_VERSION,
        en: "Create Background Color Layer " + SCRIPT_VERSION
    },
    colorPanel: {
        ja: "カラー設定",
        en: "Color Settings"
    },
    marginLabel: {
        ja: "マージン:",
        en: "Margin:"
    },
    templateCheckbox: {
        ja: "「テンプレート」レイヤーに",
        en: "Set as Template Layer"
    },
    okButton: {
        ja: "OK",
        en: "OK"
    },
    cancelButton: {
        ja: "キャンセル",
        en: "Cancel"
    }
};

// RGB → HEX 更新 / Update HEX from RGB
function updateHexFromRGB(rInput, gInput, bInput, hexInput) {
    var r = Math.min(255, Math.max(0, parseInt(rInput.text) || 0));
    var g = Math.min(255, Math.max(0, parseInt(gInput.text) || 0));
    var b = Math.min(255, Math.max(0, parseInt(bInput.text) || 0));
    var hex = "#" +
        ("0" + r.toString(16)).slice(-2) +
        ("0" + g.toString(16)).slice(-2) +
        ("0" + b.toString(16)).slice(-2);
    hexInput.text = hex.toUpperCase();
}

// HEX → RGB 更新 / Update RGB from HEX
function updateRGBFromHex(hexInput, rInput, gInput, bInput) {
    var hexValue = hexInput.text;
    var cleanHex = hexValue.replace(/^#/, ""); // 先頭の#を削除
    if (/^[0-9A-Fa-f]{6}$/.test(cleanHex)) {
        rInput.text = parseInt(cleanHex.substr(0, 2), 16);
        gInput.text = parseInt(cleanHex.substr(2, 2), 16);
        bInput.text = parseInt(cleanHex.substr(4, 2), 16);
    }
}

// 値を矢印キーで増減する関数 / Function to change value by arrow keys
function changeValueByArrowKey(editText) {
    editText.addEventListener("keydown", function(event) {
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
            delta = 1;
            if (event.keyName == "Up") {
                value += delta;
                event.preventDefault();
            } else if (event.keyName == "Down") {
                value -= delta;
                event.preventDefault();
            }
        }

        if (keyboard.altKey) {
            value = Math.round(value * 10) / 10; // 小数第1位まで
        } else {
            value = Math.round(value); // 整数
        }

        editText.text = value;
    });
}

function act_setLayTmplAttr() {
    var actionSetName = "layer";
    var actionName = "template";

    // アクション定義テキスト / Action definition text
    var actionCode ='''
 /version 3
/name [ 5
	6c61796572
]
/isOpen 1
/actionCount 1
/action-1 {
	/name [ 8
		74656d706c617465
	]
	/keyIndex 0
	/colorIndex 0
	/isOpen 1
	/eventCount 1
	/event-1 {
		/useRulersIn1stQuadrant 0
		/internalName (ai_plugin_Layer)
		/localizedName [ 9
			e8a1a8e7a4ba203a20
		]
		/isOpen 1
		/isOn 1
		/hasDialog 1
		/showDialog 0
		/parameterCount 9
		/parameter-1 {
			/key 1836411236
			/showInPalette 4294967295
			/type (integer)
			/value 4
		}
		/parameter-2 {
			/key 1851878757
			/showInPalette 4294967295
			/type (ustring)
			/value [ 36
				e383ace382a4e383a4e383bce38391e3838de383abe382aae38397e382b7e383
				a7e383b3
			]
		}
		/parameter-3 {
			/key 1953329260
			/showInPalette 4294967295
			/type (boolean)
			/value 1
		}
		/parameter-4 {
			/key 1936224119
			/showInPalette 4294967295
			/type (boolean)
			/value 1
		}
		/parameter-5 {
			/key 1819239275
			/showInPalette 4294967295
			/type (boolean)
			/value 1
		}
		/parameter-6 {
			/key 1886549623
			/showInPalette 4294967295
			/type (boolean)
			/value 1
		}
		/parameter-7 {
			/key 1886547572
			/showInPalette 4294967295
			/type (boolean)
			/value 0
		}
		/parameter-8 {
			/key 1684630830
			/showInPalette 4294967295
			/type (boolean)
			/value 1
		}
		/parameter-9 {
			/key 1885564532
			/showInPalette 4294967295
			/type (unit real)
			/value 50.0
			/unit 592474723
		}
	}
}
''';
    try {
        var tempFile = new File(Folder.temp + "/temp_action.aia");
        tempFile.open("w");
        tempFile.write(actionCode);
        tempFile.close();

        // アクション読み込み / Load action
        app.loadAction(tempFile);
        // アクション実行 / Execute action
        app.doScript(actionName, actionSetName);
        // アクション削除 / Unload action
        app.unloadAction(actionSetName, "");
        // 一時ファイル削除 / Remove temporary file
        tempFile.remove();
    } catch (e) {
        alert("エラーが発生しました: " + e);
    }
}

// 塗り色を設定する関数 / Function to set fill color
function setFillColor(doc, rectItem, useCMYK) {
    if (useCMYK) {
        // CMYKドキュメント：K45 / CMYK document: K45
        var fillColor = new CMYKColor();
        fillColor.cyan = 0;
        fillColor.magenta = 0;
        fillColor.yellow = 0;
        fillColor.black = 45;
        rectItem.fillColor = fillColor;
    } else {
        // RGBドキュメント：#999999 / RGB document: #999999
        var rgbColor = new RGBColor();
        rgbColor.red = 153;
        rgbColor.green = 153;
        rgbColor.blue = 153;
        rectItem.fillColor = rgbColor;
    }
    rectItem.filled = true;
    rectItem.stroked = false; // 線なし / No stroke
}

// 単位コードとラベルのマップ / Map of unit codes to labels
var unitLabelMap = {
  0: "in",
  1: "mm",
  2: "pt",
  3: "pica",
  4: "cm",
  5: "Q/H",
  6: "px",
  7: "ft/in",
  8: "m",
  9: "yd",
  10: "ft"
};

// 現在の単位ラベルを取得 / Get current unit label
function getCurrentUnitLabel() {
  var unitCode = app.preferences.getIntegerPreference("rulerType");
  return unitLabelMap[unitCode] || "pt";
}

// CMYK入力欄を作成する共通関数 / Common function to create CMYK input
function createCMYKInput(parent, label, defaultValue) {
    var group = parent.add("group");
    var lbl = group.add("statictext", undefined, label);
    lbl.preferredSize.width = 20;
    var input = group.add("edittext", undefined, defaultValue);
    input.characters = 4;
    input.preferredSize.width = 40;
    changeValueByArrowKey(input);
    return input;
}

// RGB入力欄を作成する共通関数 / Common function to create RGB input
function createRGBInput(parent, label, defaultValue) {
    var group = parent.add("group");
    var lbl = group.add("statictext", undefined, label);
    lbl.preferredSize.width = 20;
    var input = group.add("edittext", undefined, defaultValue);
    input.characters = 4;
    input.preferredSize.width = 40;
    changeValueByArrowKey(input);
    return input;
}

function main() {
    try {
        if (app.documents.length === 0) {
            alert("ドキュメントが開かれていません。");
            return;
        }

        var doc = app.activeDocument;

        /* ダイアログボックスを表示 / Show dialog box */
        var dlg = new Window("dialog", LABELS.dialogTitle[lang]);
        dlg.orientation = "column";
        dlg.alignChildren = "left";

        /* --- Two-column color panel (CMYK left, RGB right) --- */
        var colorPanel = dlg.add("panel", undefined, LABELS.colorPanel[lang]);
        colorPanel.orientation = "row";
        colorPanel.alignChildren = ["fill", "top"];
        colorPanel.margins = [15, 20, 15, 10];
        colorPanel.spacing = 20;

        var leftCol = colorPanel.add("group");
        leftCol.orientation = "column";
        leftCol.alignChildren = ["left", "top"];

        var rightCol = colorPanel.add("group");
        rightCol.orientation = "column";
        rightCol.alignChildren = ["left", "top"];

        // CMYK inputs
        var cInput = createCMYKInput(leftCol, "C:", "0");
        var mInput = createCMYKInput(leftCol, "M:", "0");
        var yInput = createCMYKInput(leftCol, "Y:", "0");
        var kInput = createCMYKInput(leftCol, "K:", "45");

        // RGB inputs
        var rInput = createRGBInput(rightCol, "R:", "153");
        var gInput = createRGBInput(rightCol, "G:", "153");
        var bInput = createRGBInput(rightCol, "B:", "153");

        // HEX input under RGB
        var hexGroup = rightCol.add("group");
        hexGroup.orientation = "row";
        var hexLabel = hexGroup.add("statictext", undefined, "#");
        hexLabel.preferredSize.width = 20;
        var hexInput = hexGroup.add("edittext", undefined, "999999");
        hexInput.characters = 7;

        // RGBとHEXの連動 / Link RGB and HEX
        rInput.onChanging = function() { updateHexFromRGB(rInput, gInput, bInput, hexInput); };
        gInput.onChanging = function() { updateHexFromRGB(rInput, gInput, bInput, hexInput); };
        bInput.onChanging = function() { updateHexFromRGB(rInput, gInput, bInput, hexInput); };
        hexInput.onChanging = function() { updateRGBFromHex(hexInput, rInput, gInput, bInput); };

        /* カラーモードに応じて項目をディム表示 / Dim inputs depending on document color mode */
        if (doc.documentColorSpace === DocumentColorSpace.CMYK) {
            // Disable RGB inputs
            rInput.enabled = false;
            gInput.enabled = false;
            bInput.enabled = false;
            hexInput.enabled = false;
            // Keep CMYK inputs active
            if (cInput) cInput.enabled = true;
            if (mInput) mInput.enabled = true;
            if (yInput) yInput.enabled = true;
            if (kInput) kInput.enabled = true;
        } else {
            // Disable CMYK inputs (dim but still visible)
            if (cInput) cInput.enabled = false;
            if (mInput) mInput.enabled = false;
            if (yInput) yInput.enabled = false;
            if (kInput) kInput.enabled = false;
            // Keep RGB inputs active
            rInput.enabled = true;
            gInput.enabled = true;
            bInput.enabled = true;
            hexInput.enabled = true;
        }

        var marginGroup = dlg.add("group");
        marginGroup.orientation = "row";

        marginGroup.add("statictext", undefined, LABELS.marginLabel[lang]);
        var marginInput = marginGroup.add("edittext", undefined, "0");
        marginInput.characters = 4;
        var unitLabel = marginGroup.add("statictext", undefined, getCurrentUnitLabel());

        changeValueByArrowKey(marginInput);

        var templateCheckbox = dlg.add("checkbox", undefined, LABELS.templateCheckbox[lang]);
        templateCheckbox.value = true; // デフォルトでON

        var btnGroup = dlg.add("group");
        btnGroup.orientation = "row";
        btnGroup.alignment = "center"; // 中央揃え / Center alignment

        var cancelBtn = btnGroup.add("button", undefined, LABELS.cancelButton[lang], {name:"cancel"});
        var okBtn = btnGroup.add("button", undefined, LABELS.okButton[lang], {name:"ok"});

        if (dlg.show() != 1) {
            return; // キャンセル時は終了
        }

        var margin = parseFloat(marginInput.text) || 0;
        var setAsTemplate = templateCheckbox.value;

        var abIndex = doc.artboards.getActiveArtboardIndex();
        var ab = doc.artboards[abIndex];
        var rect = ab.artboardRect; // [left, top, right, bottom]

        /* 「bg-template」レイヤーを再作成 / Recreate "bg-template" layer */
        var bgLayer;
        var existingLayer = null;
        try {
            existingLayer = doc.layers.getByName("bg-template");
        } catch (e) {}

        if (existingLayer) {
            existingLayer.locked = false; // 既存レイヤーをロック解除
            existingLayer.remove();       // 既存レイヤーを削除
        }

        bgLayer = doc.layers.add();
        bgLayer.name = "bg-template";

        /* 長方形を作成 / Create rectangle */
        var newRect = bgLayer.pathItems.rectangle(
            rect[1] + margin, // top expands with positive margin
            rect[0] - margin, // left expands with positive margin
            (rect[2] - rect[0]) + (margin * 2), // width expands with positive margin
            (rect[1] - rect[3]) + (margin * 2)  // height expands with positive margin
        );

        /* 塗り色を設定 / Set fill color */
        if (doc.documentColorSpace === DocumentColorSpace.CMYK) {
            var c = parseFloat(cInput.text) || 0;
            var m = parseFloat(mInput.text) || 0;
            var y = parseFloat(yInput.text) || 0;
            var k = parseFloat(kInput.text) || 0;
            var fillColor = new CMYKColor();
            fillColor.cyan = c;
            fillColor.magenta = m;
            fillColor.yellow = y;
            fillColor.black = k;
            newRect.fillColor = fillColor;
        } else {
            var rgbColor = new RGBColor();
            var hexValue = hexInput.text;
            var cleanHex = hexValue.replace(/^#/, ""); // 先頭の#を削除
            if (/^[0-9A-Fa-f]{6}$/.test(cleanHex)) {
                rgbColor.red = parseInt(cleanHex.substr(0, 2), 16);
                rgbColor.green = parseInt(cleanHex.substr(2, 2), 16);
                rgbColor.blue = parseInt(cleanHex.substr(4, 2), 16);
            } else {
                var r = parseFloat(rInput.text) || 0;
                var g = parseFloat(gInput.text) || 0;
                var b = parseFloat(bInput.text) || 0;
                rgbColor.red = r;
                rgbColor.green = g;
                rgbColor.blue = b;
            }
            newRect.fillColor = rgbColor;
        }
        newRect.filled = true;
        newRect.stroked = false;

        if (setAsTemplate) {
            /* テンプレートレイヤーに設定（アクションを使用）/ Set as template layer (using action) */
            act_setLayTmplAttr();
        }

        /* 一時的にロック解除 / Temporarily unlock */
        bgLayer.locked = false;
        /* 最背面に移動 / Send to back */
        bgLayer.zOrder(ZOrderMethod.SENDTOBACK);

bgLayer.name = "bg-template";

        bgLayer.locked = true;

    } catch (e) {
        alert("エラーが発生しました: " + e);
    }
}

main();