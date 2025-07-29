#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);


var persistentBaseSize = null;

/*
### スクリプト名：

TypeScaler.jsx

### GitHub：

https://github.com/swwwitch/illustrator-scripts

### 概要：

- 基準フォントサイズと倍率からタイプスケールを自動生成
- 選択テキストにフォントサイズを適用、または見本を生成
- プレビュー機能でサイズを即時確認可能

### 主な機能：

- 倍率に基づいたサイズリスト生成
- 選択テキストへのフォントサイズ適用
- 複数サイズの見本テキスト生成
- 単位に応じたラベル表示

### 処理の流れ：

1. ダイアログで基準サイズと倍率を入力
2. 自動生成されたサイズリストを表示
3. 選択したサイズを適用または見本を作成

### 参考

https://note.com/hiro_design_n/n/nc95a1d2d86a4

### 更新履歴：

- v1.0 (20250728) : 初期バージョン
- v1.1 (20250729) : UI改善とローカライズ対応
*/

/*
### Script Name：

TypeScaler.jsx

### GitHub：

https://github.com/swwwitch/illustrator-scripts

### Overview：

- Automatically generate a type scale from a base font size and ratio
- Apply font sizes to selected text or generate sample text in the document
- Preview function for immediate size confirmation

### Key Features：

- Generate size list based on ratio
- Apply font sizes to selected text
- Create multiple sample texts with sizes
- Display labels according to units

### Process Flow：

1. Enter base size and ratio in the dialog
2. Display the automatically generated size list
3. Apply the selected size or create samples

### Update History：

- v1.0 (20250728) : Initial version
- v1.1 (20250729) : UI improvements and localization support

*/

// スクリプトバージョン
var SCRIPT_VERSION = "v1.1";

function getCurrentLang() {
  return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

/* 日英ラベル定義 / Japanese-English label definitions */

/* UIラベル定義 / UI Label Definitions */
var LABELS = {
    dialogTitle: {
        ja: "タイプスケール " + SCRIPT_VERSION,
        en: "Type Scale " + SCRIPT_VERSION
    },
    baseLabel: {
        ja: "基準：",
        en: "Base:"
    },
    unitLabel: {
        ja: "（単位）",
        en: "(Unit)"
    },
    ratioDropdown: {
        ja: "倍率",
        en: "Ratio"
    },
    samplePanel: {
        ja: "見本作成",
        en: "Create Sample"
    },
    sampleText: {
        ja: "山路を登りながら",
        en: "Sample Text"
    },
    showSizeCheckbox: {
        ja: "サイズ表示",
        en: "Show Size"
    },
    sampleBtn: {
        ja: "見本作成",
        en: "Create"
    },
    cancelBtn: {
        ja: "キャンセル",
        en: "Cancel"
    },
    okBtn: {
        ja: "OK",
        en: "OK"
    },
    alertSelectSize: {
        ja: "リストからサイズを選択してください。",
        en: "Please select a size from the list."
    },
    alertInvalidSize: {
        ja: "正しいサイズを選択してください。",
        en: "Please select a valid size."
    },
    alertInvalidBase: {
        ja: "基準フォントサイズが不正です。",
        en: "Invalid base font size."
    },
    alertApplyError: {
        ja: "フォントサイズの適用に失敗しました：",
        en: "Failed to apply font size: "
    },
    alertFontError: {
        ja: "フォントの適用に失敗しました：",
        en: "Failed to apply font: "
    }
};

/* 単位ラベル取得関数とマップ / Get unit label and mapping */
function getUnitLabel(code, prefKey) {
  if (code === 5 && prefKey === "text/asianunits") {
    return "H";
  }
  return unitLabelMap[code] || "不明";
}

var unitLabelMap = {
 0: "in",
 1: "mm",
 2: "pt",
 3: "pica",
 4: "cm",
 5: "Q/H",
 6: "px"
};

/* 番号付きラベル生成関数 / Generate numbered labels */
function createNumberedLabels(items, offset, unitLabel) {
  var labels = [];
  for (var i = 0; i < items.length; i++) {
    var label = (offset + i) + " - " + items[i] + (unitLabel ? " " + unitLabel : "");
    labels.push(label);
  }
  return labels;
}


function getSelectedTextFrames() {
  var sel = app.activeDocument.selection;
  var frames = [];
  for (var i = 0; i < sel.length; i++) {
    if (sel[i].typename === "TextFrame") {
      frames.push(sel[i]);
    }
  }
  return frames;
}

// =========================
// フォントサイズ配列を倍率に基づいて生成 / Generate array of font sizes based on ratio
// =========================
function generateTypeScaleSizes(baseSize, ratio) {
  var sizes = [];
  var down = baseSize;
  for (var i = 0; i < 2; i++) {
    down /= ratio;
  }
  for (var i = 0; i < 7; i++) {
    sizes.push(Math.round(down * 10) / 10);
    down *= ratio;
  }
  return sizes;
}

// =========================
// リストボックスを更新 / Update listbox with calculated font sizes
// =========================
function updateSizeList(sizeList, ratioPopup, baseSize, textUnitLabel) {
  sizeList.removeAll();
  if (isNaN(baseSize) || baseSize <= 0) return;
  var ratio = ratioPopup.selection ? ratioValues[ratioPopup.selection.index] : 1.25;
  var sizes = generateTypeScaleSizes(baseSize, ratio);
  for (var j = 0; j < sizes.length; j++) {
    sizeList.add("item", sizes[j] + " " + textUnitLabel);
  }
}

// =========================
// 倍率ラベル・値配列
// =========================
var ratioLabels = [
  "Minor Second 1.067",
  "Major Second 1.125",
  "Minor Third 1.2",
  "Major Third 1.25",
  "Golden Ratio: ½ 1.309",
  "Perfect Fourth 1.333",
  "Augmented Fourth 1.414",
  "Golden Ratio 1.618"
];
var ratioValues = [1.067, 1.125, 1.2, 1.25, 1.309, 1.333, 1.414, 1.618];

(function () {
  if (app.documents.length > 0) main();
})();

/* メイン処理 / Main process */
function main() {
  var previewFrame = null;


  // =========================
  // UI構築
  // =========================
  var dialog = new Window("dialog", LABELS.dialogTitle[lang]);
  dialog.orientation = "column";
  dialog.alignChildren = "left";

  var textUnitCode = app.preferences.getIntegerPreference("text/units");
  var textUnitLabel = getUnitLabel(textUnitCode, "text/units");
  var sizeGroup = dialog.add("group");
  sizeGroup.orientation = "row";
  sizeGroup.margins = [0, 0, 0, 15];
  sizeGroup.spacing = 5;
  sizeGroup.add("statictext", undefined, LABELS.baseLabel[lang]);
  var sizeInput = sizeGroup.add("edittext", undefined, "12");
  sizeInput.characters = 4;
  sizeGroup.add("statictext", undefined, LABELS.unitLabel[lang].replace("単位", textUnitLabel));
  changeValueByArrowKey(sizeInput);

  var ratioPopup = sizeGroup.add("dropdownlist", undefined, ratioLabels);
  ratioPopup.selection = 3; // "Major Third 1.25"


  sizeGroup.alignment = "center";

  var mainGroup = dialog.add("group");
  mainGroup.orientation = "row";

  // 左カラム
  var leftPanel = mainGroup.add("group");
  leftPanel.orientation = "column";
  leftPanel.alignChildren = "left";



  var sizeList = leftPanel.add("listbox", undefined, [], { multiselect: false });
  sizeList.preferredSize = [85, 136];

  // 右カラム
  var rightPanel = mainGroup.add("group");
  rightPanel.orientation = "column";
  rightPanel.alignment = "top";
  rightPanel.alignChildren = "left";

  var samplePanel = rightPanel.add("panel", undefined, LABELS.samplePanel[lang]);
  samplePanel.orientation = "column";
  samplePanel.alignChildren = "left";
  samplePanel.margins = [15, 20, 15, 10];

  var sampleInput = samplePanel.add("edittext", undefined, LABELS.sampleText[lang]);
  sampleInput.characters = 20;
  var showSizeCheckbox = samplePanel.add("checkbox", undefined, LABELS.showSizeCheckbox[lang]);
  showSizeCheckbox.value = true;
  var sampleBtn = samplePanel.add("button", undefined, LABELS.sampleBtn[lang]);
  sampleBtn.alignment = "right";

  // ボタングループをダイアログ下部に追加
  var buttonGroup = dialog.add("group");
  buttonGroup.orientation = "row";
  buttonGroup.alignment = "center";
  var cancelBtn = buttonGroup.add("button", undefined, LABELS.cancelBtn[lang]);
  var okBtn = buttonGroup.add("button", undefined, LABELS.okBtn[lang]);

  // =========================
  // イベント定義
  // =========================


  // 基準サイズ入力変更時
  sizeInput.onChanging = function () {
    var inputValue = parseFloat(sizeInput.text);
    if (!isNaN(inputValue) && inputValue > 0) {
      persistentBaseSize = inputValue;
    }
    updateSizeList(sizeList, ratioPopup, inputValue, textUnitLabel);
  };

  // 倍率変更時
  ratioPopup.onChange = function () {
    var inputValue = parseFloat(sizeInput.text);
    updateSizeList(sizeList, ratioPopup, inputValue, textUnitLabel);
  };

  // OKボタン押下時（選択テキストにサイズ適用）
  okBtn.onClick = function () {
    if (!sizeList.selection) {
      alert(LABELS.alertSelectSize[lang]);
      return;
    }
    var selectedText = sizeList.selection.text;
    // Split by spaces and take the part that contains "pt" or numeric size
    var parts = selectedText.split(" ");
    var sizeValue = NaN;
    for (var i = 0; i < parts.length; i++) {
        if (parts[i].indexOf("pt") !== -1 || !isNaN(parseFloat(parts[i]))) {
            sizeValue = parseFloat(parts[i]);
            break;
        }
    }
    if (isNaN(sizeValue) || sizeValue <= 0) {
      alert(LABELS.alertInvalidSize[lang]);
      return;
    }

    var sel = app.activeDocument.selection;
    for (var i = 0; i < sel.length; i++) {
      var item = sel[i];
      try {
        if (item.typename === "TextRange") {
          item.characterAttributes.size = sizeValue;
        } else if (item.typename === "TextFrame") {
          if (item.textRange && item.textRange.characters.length > 0) {
            item.textRange.characterAttributes.size = sizeValue;
          }
        }
      } catch (e) {
        alert(LABELS.alertApplyError[lang] + e.message);
      }
    }
    app.redraw(); // ← この行を追加
    // Remove previewFrame if present before closing dialog
    if (previewFrame && previewFrame.isValid) {
      try { previewFrame.remove(); } catch (e) {}
    }
    dialog.close();
  };

  // キャンセルボタン押下時
  cancelBtn.onClick = function () {
    dialog.close();
  };

  // 見本作成ボタン押下時（サイズリストの見本テキストを作成）
  sampleBtn.onClick = function () {
    var baseSize;
    if (sizeList.selection) {
      baseSize = parseFloat(sizeList.selection.text.split(" ")[0]);
    } else if (persistentBaseSize !== null) {
      baseSize = persistentBaseSize;
    } else {
      baseSize = parseFloat(sizeInput.text);
    }
    persistentBaseSize = baseSize;
    if (isNaN(baseSize) || baseSize <= 0) {
      alert(LABELS.alertInvalidBase[lang]);
      return;
    }

    // 現在の選択からフォントを取得（最初に見つかったテキストフレームから）
    var sel = app.activeDocument.selection;
    var selectedFont = null;
    for (var i = 0; i < sel.length; i++) {
      if (sel[i].typename === "TextFrame") {
        try {
          selectedFont = sel[i].textRange.characterAttributes.textFont;
          break;
        } catch (e) {
          // 無視して次のオブジェクトを見る
        }
      }
    }

    var doc = app.activeDocument;
    var x = 20;
    var y = -20;
    var yOffset = 20;

    var ratio = ratioPopup.selection ? ratioValues[ratioPopup.selection.index] : 1.25;
    var sizes = generateTypeScaleSizes(baseSize, ratio);
    for (var i = 0; i < sizes.length; i++) {
      var fontSize = sizes[i];
      var tf = doc.textFrames.add();
      var contentText = sampleInput.text;
      if (showSizeCheckbox.value) {
        contentText += "（" + fontSize + textUnitLabel + "）";
      }
      tf.contents = contentText;
      tf.left = x;
      tf.top = y;

      if (selectedFont) {
        try {
          tf.textRange.characterAttributes.textFont = selectedFont;
        } catch (e) {
          alert(LABELS.alertFontError[lang] + e.message);
        }
      }

      try {
        tf.textRange.characterAttributes.size = fontSize;
      } catch (e) {
        alert(LABELS.alertApplyError[lang] + e.message);
      }

      y -= fontSize + yOffset;
    }

    dialog.close();
  };


  // =========================
  // 初期表示更新とダイアログ表示
  // =========================
  // --- Ensure text object selection when in text edit mode ---
  if (app.documents.length && app.selection.constructor.name === "TextRange") {
    var textFramesInStory = app.selection.story.textFrames;
    if (textFramesInStory.length === 1) {
      app.executeMenuCommand("deselectall");
      app.selection = [textFramesInStory[0]];
      try {
        app.selectTool("Adobe Select Tool");
      } catch (e) {}
    }
  }
  // -----------------------------------------------------------
  updateSizeList(sizeList, ratioPopup, parseFloat(sizeInput.text), textUnitLabel);
  // 選択テキストがない場合は listbox を無効化
  if (getSelectedTextFrames().length === 0) {
    sizeList.enabled = false;
  }

  // --- Opacity and position adjustment ---
  var offsetX = 300;
  var dialogOpacity = 0.97;

  function shiftDialogPosition(dlg, offsetX, offsetY) {
      dlg.onShow = function () {
          var currentX = dlg.location[0];
          var currentY = dlg.location[1];
          dlg.location = [currentX + offsetX, currentY + offsetY];
      };
  }

  function setDialogOpacity(dlg, opacityValue) {
      dlg.opacity = opacityValue;
  }

  setDialogOpacity(dialog, dialogOpacity);
  shiftDialogPosition(dialog, offsetX, 0);
  // --- End Opacity and position adjustment ---

  dialog.center();
  dialog.onShow = function() {
      sizeInput.active = true;
  };
  dialog.show();

  /* 上下キーで値を変更する関数 / Change value with up/down arrow keys */
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
                  if (value < 0) value = 0;
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
                  if (value < 0) value = 0;
                  event.preventDefault();
              }
          }

          if (keyboard.altKey) {
              value = Math.round(value * 10) / 10;
          } else {
              value = Math.round(value);
          }

          editText.text = value;
      });
  }
}