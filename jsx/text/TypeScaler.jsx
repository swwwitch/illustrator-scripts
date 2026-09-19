#targetengine "TypeScalerEngine"
#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したテキストのフォントサイズを、基準サイズと比率から算出して適用します。

詳細は README を参照してください。

### Overview

Sets the font size of the selected text from a base size and a ratio.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "TypeScaler";                   /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.2.1";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-19";                             /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/TypeScaler.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/TypeScaler.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    var persistentBaseSize = null;
    $.global.__sizeValue = $.global.__sizeValue || "12";
    $.global.__ratioIndex = ($.global.__ratioIndex !== undefined) ? $.global.__ratioIndex : 3;

    // スクリプトバージョン

    function getCurrentLang() {
      return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = getCurrentLang();

    /* 日英ラベル定義 / Japanese-English label definitions */

    /* UIラベル定義 / UI Label Definitions */
    var LABELS = {
        dialogTitle: {
            ja: "タイプスケール " + SCRIPT_VERSION,
            en: "Type Scale " + SCRIPT_VERSION
        },
        baseLabel: {
            ja: "基準",
            en: "Base"
        },
        baseLabelSuffix: {
            ja: "（{unit}）",
            en: " ({unit})"
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
        tipBaseSize: {
            ja: "タイプスケールの基準になるフォントサイズです。",
            en: "Font size the type scale is built from."
        },
        tipRatio: {
            ja: "1段ごとに掛ける倍率です。大きいほどサイズの差が開きます。",
            en: "Multiplier applied at each step. A larger ratio spreads the sizes further apart."
        },
        tipSizeList: {
            ja: "計算したサイズの一覧です。選んで OK すると、選択中のテキストに適用します。",
            en: "The calculated sizes. Pick one and press OK to apply it to the selected text."
        },
        tipSampleText: {
            ja: "見本に使う文字列です。",
            en: "Text used for the sample."
        },
        tipShowSize: {
            ja: "見本の各行にサイズの数値を添えます。",
            en: "Adds the size value to each line of the sample."
        },
        tipSampleBtn: {
            ja: "一覧のすべてのサイズで見本を作り、ドキュメントに配置します。",
            en: "Creates a sample in every size on the list and places it in the document."
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
    /* 単位テーブル（配列の添字が rulerType コードと一致：0=in, 1=mm, 2=pt …）/ Unit table; the array index equals the rulerType code */
    var UNITS = [
      { label: "in",    pointsPerUnit: 72 },                /* 0 */
      { label: "mm",    pointsPerUnit: 72 / 25.4 },         /* 1 */
      { label: "pt",    pointsPerUnit: 1 },                 /* 2 */
      { label: "pica",  pointsPerUnit: 12 },                /* 3 */
      { label: "cm",    pointsPerUnit: 72 / 2.54 },         /* 4 */
      { label: "Q",     pointsPerUnit: 72 / 25.4 * 0.25 },  /* 5 */
      { label: "px",    pointsPerUnit: 1 },                 /* 6 */
      { label: "ft/in", pointsPerUnit: 72 * 12 },           /* 7 */
      { label: "m",     pointsPerUnit: 72 / 25.4 * 1000 },  /* 8 */
      { label: "yd",    pointsPerUnit: 72 * 36 },           /* 9 */
      { label: "ft",    pointsPerUnit: 72 * 12 }            /* 10 */
    ];

    /* 単位コード5を「歯（H）」と表示する環境設定キー。文字サイズ（text/units）だけ「級（Q）」
       Preference keys that show unit code 5 as H; only the type size (text/units) shows Q */
    var HA_UNIT_PREF_KEYS = { "rulerType": true, "strokeUnits": true, "text/asianunits": true };

    /**
     * 設定キーごとの単位情報を取得する
     * @param {string} prefKey - 環境設定キー（省略時は "rulerType"）
     * @returns {{code: number, label: string, pointsPerUnit: number}} 単位情報
     */
    function getUnitInfo(prefKey) {
      var unitKey = prefKey || "rulerType";
      var unitCode = app.preferences.getIntegerPreference(unitKey);
      var unit = UNITS[unitCode] || UNITS[2];
      var label = (unitCode === 5 && HA_UNIT_PREF_KEYS[unitKey]) ? "H" : unit.label;
      return { code: unitCode, label: label, pointsPerUnit: unit.pointsPerUnit };
    }

    function getSelectedTextFrames() {
      var currentSelection = app.activeDocument.selection;
      var frames = [];
      for (var i = 0; i < currentSelection.length; i++) {
        if (currentSelection[i].typename === "TextFrame") {
          frames.push(currentSelection[i]);
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
      var dialog = new Window("dialog", LABELS.dialogTitle[uiLang]);
      dialog.orientation = "column";
      dialog.alignChildren = "left";

      var textUnitLabel = getUnitInfo("text/units").label;
      var sizeGroup = dialog.add("group");
      sizeGroup.orientation = "row";
      sizeGroup.margins = [0, 0, 0, 15];
      sizeGroup.spacing = 5;
      sizeGroup.add("statictext", undefined, LABELS.baseLabel[uiLang] + (uiLang === "ja" ? "：" : ": "));
      var sizeInput = sizeGroup.add("edittext", undefined, $.global.__sizeValue);
      sizeInput.characters = 4;
      sizeInput.helpTip = LABELS.tipBaseSize[uiLang];
      sizeGroup.add("statictext", undefined, LABELS.baseLabelSuffix[uiLang].replace("{unit}", textUnitLabel));
      changeValueByArrowKey(sizeInput);

      var ratioPopup = sizeGroup.add("dropdownlist", undefined, ratioLabels);
      ratioPopup.selection = $.global.__ratioIndex;
      ratioPopup.helpTip = LABELS.tipRatio[uiLang];

      sizeGroup.alignment = "center";

      var mainGroup = dialog.add("group");
      mainGroup.orientation = "row";

      // 左カラム
      var leftPanel = mainGroup.add("group");
      leftPanel.orientation = "column";
      leftPanel.alignChildren = "left";

      var sizeList = leftPanel.add("listbox", undefined, [], { multiselect: false });
      sizeList.preferredSize = [85, 136];
      sizeList.helpTip = LABELS.tipSizeList[uiLang];

      // 右カラム
      var rightPanel = mainGroup.add("group");
      rightPanel.orientation = "column";
      rightPanel.alignment = "top";
      rightPanel.alignChildren = "left";

      var samplePanel = rightPanel.add("panel", undefined, LABELS.samplePanel[uiLang]);
      samplePanel.orientation = "column";
      samplePanel.alignChildren = "left";
      samplePanel.margins = [15, 20, 15, 10];

      var sampleInput = samplePanel.add("edittext", undefined, LABELS.sampleText[uiLang]);
      sampleInput.characters = 20;
      sampleInput.helpTip = LABELS.tipSampleText[uiLang];
      var showSizeCheckbox = samplePanel.add("checkbox", undefined, LABELS.showSizeCheckbox[uiLang]);
      showSizeCheckbox.value = true;
      showSizeCheckbox.helpTip = LABELS.tipShowSize[uiLang];
      var sampleBtn = samplePanel.add("button", undefined, LABELS.sampleBtn[uiLang]);
      sampleBtn.alignment = "right";
      sampleBtn.helpTip = LABELS.tipSampleBtn[uiLang];

      // ボタングループをダイアログ下部に追加
      var buttonGroup = dialog.add("group");
      buttonGroup.orientation = "row";
      buttonGroup.alignment = "center";
      var cancelBtn = buttonGroup.add("button", undefined, LABELS.cancelBtn[uiLang]);
      var okBtn = buttonGroup.add("button", undefined, LABELS.okBtn[uiLang]);

      // =========================
      // イベント定義
      // =========================

      // 基準サイズ入力変更時
      sizeInput.onChanging = function () {
        var inputValue = parseFloat(sizeInput.text);
        if (!isNaN(inputValue) && inputValue > 0) {
          persistentBaseSize = inputValue;
          $.global.__sizeInput = sizeInput;
        }
        $.global.__sizeValue = sizeInput.text;
        updateSizeList(sizeList, ratioPopup, inputValue, textUnitLabel);
      };

      // 倍率変更時
      ratioPopup.onChange = function () {
        var inputValue = parseFloat(sizeInput.text);
        updateSizeList(sizeList, ratioPopup, inputValue, textUnitLabel);
        $.global.__ratioPopup = ratioPopup;
        $.global.__ratioIndex = ratioPopup.selection ? ratioPopup.selection.index : 0;
      };

      // OKボタン押下時（選択テキストにサイズ適用）
      okBtn.onClick = function () {
        if (!sizeList.selection) {
          alert(LABELS.alertSelectSize[uiLang]);
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
          alert(LABELS.alertInvalidSize[uiLang]);
          return;
        }

        var currentSelection = app.activeDocument.selection;
        for (var i = 0; i < currentSelection.length; i++) {
          var item = currentSelection[i];
          try {
            if (item.typename === "TextRange") {
              item.characterAttributes.size = sizeValue;
            } else if (item.typename === "TextFrame") {
              if (item.textRange && item.textRange.characters.length > 0) {
                item.textRange.characterAttributes.size = sizeValue;
              }
            }
          } catch (e) {
            alert(LABELS.alertApplyError[uiLang] + e.message);
          }
        }
        app.redraw(); // ← この行を追加
        $.global.__sizeValue = sizeInput.text;
        $.global.__ratioIndex = ratioPopup.selection ? ratioPopup.selection.index : 0;
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
        $.global.__sizeValue = sizeInput.text;
        $.global.__ratioIndex = ratioPopup.selection ? ratioPopup.selection.index : 0;
        if (isNaN(baseSize) || baseSize <= 0) {
          alert(LABELS.alertInvalidBase[uiLang]);
          return;
        }

        // 現在の選択からフォントを取得（最初に見つかったテキストフレームから）
        var currentSelection = app.activeDocument.selection;
        var selectedFont = null;
        for (var i = 0; i < currentSelection.length; i++) {
          if (currentSelection[i].typename === "TextFrame") {
            try {
              selectedFont = currentSelection[i].textRange.characterAttributes.textFont;
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
              alert(LABELS.alertFontError[uiLang] + e.message);
            }
          }

          try {
            tf.textRange.characterAttributes.size = fontSize;
          } catch (e) {
            alert(LABELS.alertApplyError[uiLang] + e.message);
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

})();
