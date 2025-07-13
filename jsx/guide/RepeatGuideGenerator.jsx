#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
### スクリプト名
RepeatGuideGenerator.jsx

### 概要
- ダイアログで1本目の位置、間隔、本数を指定し、アートボード左上基準で横・縦方向にガイドを作成します。
- プレビューは常にONで、値を変更すると即時反映されます。
- シフトキー押下時は矢印キーで±10、通常は±1の値変更が可能です。

### 更新履歴
- v1.0 (20250713) : 初版
- v1.1 (20250713) : 全体ロジック見直し、安定性改善、プレビュー常時ON、矢印キー対応
*/

function getCurrentLang() {
    return ($.locale && $.locale.indexOf('ja') === 0) ? 'ja' : 'en';
}

var lang = getCurrentLang();

// -------------------------------
// 日英ラベル定義 / Define labels
// -------------------------------
var LABELS = {
    dialogTitle: { ja: "ガイド作成", en: "Guide Generator" },
    horizontal: { ja: "横", en: "Horizontal" },
    vertical: { ja: "縦", en: "Vertical" },
    first: { ja: "1本目", en: "First" },
    count: { ja: "本数", en: "Count" },
    spacing: { ja: "間隔", en: "Spacing" },
    ok: { ja: "OK", en: "OK" },
    cancel: { ja: "キャンセル", en: "Cancel" }
};

var unitLabelMap = {
  0: "in", 1: "mm", 2: "pt", 3: "pica", 4: "cm",
  5: "Q/H", 6: "px", 7: "ft/in", 8: "m", 9: "yd", 10: "ft"
};

// 現在の単位ラベルを取得 / Get current unit label
function getCurrentUnitLabel() {
  var unitCode = app.preferences.getIntegerPreference("rulerType");
  return unitLabelMap[unitCode] || "pt";
}

// 単位コードからポイント換算係数を取得 / Get point factor from unit code
function getPtFactorFromUnitCode(code) {
    switch (code) {
        case 0: return 72.0;
        case 1: return 72.0 / 25.4;
        case 2: return 1.0;
        case 3: return 12.0;
        case 4: return 72.0 / 2.54;
        case 5: return 72.0 / 25.4 * 0.25;
        case 6: return 1.0;
        case 7: return 72.0 * 12.0;
        case 8: return 72.0 / 25.4 * 1000.0;
        case 9: return 72.0 * 36.0;
        case 10: return 72.0 * 12.0;
        default: return 1.0;
    }
}

function changeValueByArrowKey(editText, onUpdate) {
    editText.addEventListener("keydown", function(event) {
        var value = Number(editText.text);
        if (isNaN(value)) return;
        var keyboard = ScriptUI.environment.keyboardState;
        var delta = keyboard.shiftKey ? 10 : 1;
        if (event.keyName == "Up") {
            value += delta;
            event.preventDefault();
        } else if (event.keyName == "Down") {
            value -= delta;
            event.preventDefault();
        }
        editText.text = value;
        if (typeof onUpdate === "function") {
            onUpdate(editText.text);
        }
    });
}

function main() {
    if (app.documents.length == 0) {
        alert("ドキュメントを開いてください。\nPlease open a document.");
        return;
    }

    var doc = app.activeDocument;
    var ab = doc.artboards[doc.artboards.getActiveArtboardIndex()];
    var abRect = ab.artboardRect;

    var dlg = new Window("dialog", LABELS.dialogTitle[lang]);
    dlg.orientation = "column";
    dlg.alignChildren = ["left", "top"];

    // 方向選択ラジオボタン / Direction selection radio buttons
    var rbGroup = dlg.add("group");
    rbGroup.orientation = "row";
    var rbHorizontal = rbGroup.add("radiobutton", undefined, LABELS.horizontal[lang]);
    var rbVertical = rbGroup.add("radiobutton", undefined, LABELS.vertical[lang]);
    rbHorizontal.value = true;

    // 1本目入力フィールド / First line position
    var xInputGroup = dlg.add("group");
    xInputGroup.add("statictext", undefined, LABELS.first[lang]);
    var xInput = xInputGroup.add("edittext", undefined, "100");
    xInput.characters = 5;
    xInputGroup.add("statictext", undefined, getCurrentUnitLabel());

    // 本数入力フィールド / Number of lines
    var repeatGroup = dlg.add("group");
    repeatGroup.add("statictext", undefined, LABELS.count[lang]);
    var repeatInput = repeatGroup.add("edittext", undefined, "1");
    repeatInput.characters = 5;

    // 間隔入力フィールド / Spacing
    var spacingGroup = dlg.add("group");
    spacingGroup.add("statictext", undefined, LABELS.spacing[lang]);
    var spacingInput = spacingGroup.add("edittext", undefined, "10");
    spacingInput.characters = 5;
    spacingGroup.add("statictext", undefined, getCurrentUnitLabel());

    var previewLines = [];

    // プレビュークリア / Clear preview
    function clearPreview() {
        for (var i = 0; i < previewLines.length; i++) {
            try { previewLines[i].remove(); } catch (e) {}
        }
        previewLines = [];
    }

    // ガイド描画 / Draw guides
    function drawGuides() {
        clearPreview();

        var xSize = parseFloat(xInput.text);
        var spacing = parseFloat(spacingInput.text);
        var repeatCount = parseInt(repeatInput.text, 10);

        if (isNaN(xSize) || isNaN(repeatCount) || repeatCount < 1) return;

        var unitCode = app.preferences.getIntegerPreference("rulerType");
        var factor = getPtFactorFromUnitCode(unitCode);
        var xSizePt = xSize * factor;
        var spacingPt = isNaN(spacing) ? 0 : spacing * factor;

        var guideLayer = null;
        for (var i = 0; i < doc.layers.length; i++) {
            if (doc.layers[i].name == "_guide") {
                guideLayer = doc.layers[i];
                break;
            }
        }
        if (!guideLayer) {
            guideLayer = doc.layers.add();
            guideLayer.name = "_guide";
        }
        guideLayer.locked = false;
        guideLayer.visible = true;

        for (var i = 0; i < repeatCount; i++) {
            var offset = spacingPt * i;
            var line;
            if (rbHorizontal.value) {
                var lineTop = abRect[1] - xSizePt - offset;
                line = doc.pathItems.add();
                line.stroked = true;
                line.strokeColor = new GrayColor();
                line.strokeWidth = 1;
                line.filled = false;
                line.guides = false; // プレビューではガイドにしない
                line.setEntirePath([[abRect[0], lineTop], [abRect[2], lineTop]]);
            } else if (rbVertical.value) {
                var lineLeft = abRect[0] + xSizePt + offset;
                line = doc.pathItems.add();
                line.stroked = true;
                line.strokeColor = new GrayColor();
                line.strokeWidth = 1;
                line.filled = false;
                line.guides = false; // プレビューではガイドにしない
                line.setEntirePath([[lineLeft, abRect[1]], [lineLeft, abRect[3]]]);
            }
            line.move(guideLayer, ElementPlacement.PLACEATBEGINNING);
            previewLines.push(line);
        }
    }

    repeatInput.onChanging = function() {
        var count = parseInt(repeatInput.text, 10);
        if (isNaN(count) || count < 1) {
            repeatInput.text = "1";
            count = 1;
        }
        spacingInput.enabled = true;
        drawGuides();
    };
    repeatInput.onChanging();

    xInput.onChanging = spacingInput.onChanging = function() {
        drawGuides();
    };

    var btnGroup = dlg.add("group");
    btnGroup.alignment = ["right", "top"];
    var okBtn = btnGroup.add("button", undefined, LABELS.ok[lang]);
    var cancelBtn = btnGroup.add("button", undefined, LABELS.cancel[lang]);

    okBtn.onClick = function() {
        for (var i = 0; i < previewLines.length; i++) {
            previewLines[i].guides = true;
        }
        dlg.close();
    };

    cancelBtn.onClick = function() {
        clearPreview();
        dlg.close();
    };

    changeValueByArrowKey(xInput, function() { drawGuides(); });
    changeValueByArrowKey(spacingInput, function() { drawGuides(); });
    changeValueByArrowKey(repeatInput, function() { drawGuides(); });

    dlg.show();
}

main();