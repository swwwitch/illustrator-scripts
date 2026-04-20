#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
### スクリプト名：

SmartSliceWithPuzzlify.jsx

### 概要

- 選択した画像や図形を、グリッドまたはジグソーパズル形状に分割し、それぞれにマスクを適用するIllustrator用スクリプトです。
- 行数・列数・ピース数の指定に対応し、オフセット、オーバーラップ、バラけ（Scatter）、ケイ、角丸などを組み合わせて調整できます。
- 日本語／英語UIに対応し、入力欄では ↑↓ / Shift+↑↓ / Option+↑↓ による値の増減も行えます。

### 主な機能

- グリッド分割、トラディショナル、ランダムなジグソー形状の生成
- 行数・列数・ピース数の指定と、選択オブジェクト比率に応じた自動行列算出
- オフセット、オーバーラップ、バラけ（Scatter）処理の適用
- ケイの追加、角丸の適用
- 画像、シンボル、ベクターアートワーク、複数選択オブジェクトへの対応
- 複数選択時の一時グループ化とシンボル化処理
- 日本語／英語インターフェース対応
- 入力欄でのキーボード増減操作
  - ↑↓キー：±1
  - Shift+↑↓：±10
  - Option+↑↓：±0.1

### 処理の流れ

1. 対象オブジェクト（画像、シンボル、ベクターなど）を選択
2. ダイアログで分割モード、ピース数、行数、列数、形状、各種効果を指定
3. モードに応じてグリッドまたはジグソー形状の各ピースを生成
4. 必要に応じてオフセットやバラけ処理を適用
5. 元オブジェクトを削除し、新しいピースを配置

### オリジナル、謝辞

Originally created by Jongware on 7-Oct-2010
https://community.adobe.com/t5/illustrator-discussions/cut-multiple-jigsaw-shapes-out-of-image-simultaneously/td-p/8709621#9185144

### 更新履歴

- v1.0 (20250607) : 初期バージョン
- v1.5.0 (20260421) : scatterを一様分布→ガウス分布（中央寄り・クランプ付き）に変更、scatterを含む単位のpt正規化、オーバーラップ初期OFF＋値保持、オフセット初期OFF＋デフォルト値-2、モードごとに既定値で初期化（UI仕様整理）、offset処理のtry/finally化による安全化

---

### Script Name:

SmartSliceWithPuzzlify.jsx

### Overview

- An Illustrator script that splits a selected image or artwork into grid or jigsaw puzzle pieces and applies a mask to each piece.
- Supports rows, columns, and total piece count, along with offset, overlap, scatter, stroke, and round-corner adjustments.
- Supports Japanese and English UI, and numeric fields can be adjusted using ↑↓ / Shift+↑↓ / Option+↑↓ keyboard shortcuts.

### Main Features

- Generate grid, traditional, or random jigsaw-shaped pieces
- Specify rows, columns, and total pieces, with automatic row/column calculation based on artwork ratio
- Apply offset, overlap, and scatter effects
- Add stroke and apply round corners
- Supports images, symbols, vector artwork, and multiple selected objects
- Temporary grouping and symbolization for multi-selection workflows
- Japanese and English UI support
- Keyboard increment/decrement support in numeric fields
  - Up/Down: ±1
  - Shift+Up/Down: ±10
  - Option+Up/Down: ±0.1

### Process Flow

1. Select a target object (image, symbol, vector artwork, etc.)
2. Configure split mode, total pieces, rows, columns, shape type, and optional effects in the dialog
3. Generate grid or jigsaw-shaped pieces according to the selected mode
4. Optionally apply offset and scatter effects
5. Remove the original object and place the generated pieces

### Original / Acknowledgements

Originally created by Jongware on 7-Oct-2010
https://community.adobe.com/t5/illustrator-discussions/cut-multiple-jigsaw-shapes-out-of-image-simultaneously/td-p/8709621#9185144

### Update History

- v1.0 (20250607): Initial version
- v1.5.0 (20260421): Scatter changed from uniform to Gaussian distribution (center-biased with clamp), unit normalization to pt including scatter, overlap default OFF with value retention, offset default OFF with default value -2, mode-based initialization (UI behavior clarified), and safer offset processing using try/finally
*/

(function () {
  var SCRIPT_VERSION = "v1.5.0";

  function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
  }
  var lang = getCurrentLang();

  /* 日英ラベル定義 / Japanese-English label definitions */
  var LABELS = {
    modePuzzle: { ja: "パズル", en: "Puzzle" },
    modeGridSplit: { ja: "グリッド", en: "Grid" },
    shapeTraditional: { ja: "トラディショナル", en: "Traditional" },
    shapeRandom: { ja: "ランダム", en: "Random" },
    totalPieces: { ja: "ピース数", en: "Total pieces" },
    columns: { ja: "列数", en: "Columns" },
    rows: { ja: "行数", en: "Rows" },
    explode: { ja: "バラけ処理", en: "Scatter" },
    offsetLabel: { ja: "オフセット", en: "Offset" },
    overlap: { ja: "オーバーラップ", en: "Overlap" },
    ruleCheck: { ja: "ケイ（1ptの罫線を追加）", en: "Add Stroke (1pt)" },
    roundCheck: { ja: "角丸", en: "Apply Round Corners" },
    shapeLabel: { ja: "形状：", en: "Shape:" },
    alertMaskNotPath: {
      ja: "マスク用オブジェクトが PathItem ではないため、マスクをスキップします。",
      en: "The mask object is not a PathItem, so the mask step will be skipped."
    },
    alertMultiSymbolizeFailed: {
      ja: "複数オブジェクトのシンボル化に失敗しました: ",
      en: "Failed to symbolize multiple objects: "
    },
    alertRasterSymbolizeFailed: {
      ja: "埋め込み画像のシンボル化に失敗しました: ",
      en: "Failed to symbolize embedded artwork: "
    },
    alertVectorSymbolizeFailed: {
      ja: "ベクターオブジェクトのシンボル化に失敗しました: ",
      en: "Failed to symbolize vector artwork: "
    },
    alertOffsetGroupNoPath: {
      ja: "オフセット後の GroupItem に PathItem が含まれていません。マスク処理をスキップします。",
      en: "The offset GroupItem does not contain a PathItem, so the mask step will be skipped."
    },
    alertOffsetUnexpectedType: {
      ja: "オフセット後のオブジェクトが予期しない型です。マスク処理をスキップします。",
      en: "The object after offsetting has an unexpected type, so the mask step will be skipped."
    },
    alertOffsetError: {
      ja: "オフセット適用中にエラーが発生しました: ",
      en: "An error occurred while applying the offset: "
    },
    alertScriptError: {
      ja: "スクリプト実行中にエラーが発生しました: ",
      en: "An error occurred while running the script: "
    },
    alertGeneralError: {
      ja: "エラーが発生しました：\n",
      en: "An error occurred:\n"
    },
    dialogTitle: {
      ja: "オブジェクトの分割",
      en: "Split Artwork"
    },
    panelSplit: { ja: "分割", en: "Split" },
    panelOptions: { ja: "オプション", en: "Options" },
    okBtn: { ja: "OK", en: "OK" },
    cancel: { ja: "キャンセル", en: "Cancel" }
  };

  function L(key) {
    return LABELS[key][lang];
  }

  // 単位コードとラベルのマップ
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

  // 現在の単位ラベルを取得
  function getCurrentUnitLabel() {
    var unitCode = app.preferences.getIntegerPreference("rulerType");
    return unitLabelMap[unitCode] || "pt";
  }

  // 現在の単位→pt 変換係数 / Unit-to-pt conversion factor
  function getUnitToPtFactor() {
    var unitCode = app.preferences.getIntegerPreference("rulerType");
    switch (unitCode) {
      case 0: return 72;               /* in */
      case 1: return 72 / 25.4;        /* mm */
      case 2: return 1;                /* pt */
      case 3: return 12;               /* pica */
      case 4: return 72 / 2.54;        /* cm */
      case 5: return 72 / 25.4 * 0.25; /* Q */
      case 6: return 1;                /* px */
      case 7: return 72;               /* ft/in */
      case 8: return 72 / 0.0254;      /* m */
      case 9: return 72 * 36;          /* yd */
      case 10: return 72 * 12;         /* ft */
      default: return 1;
    }
  }

  // 画像サイズに基づく初期グリッドサイズを計算
  function getInitialGridSize(imageWidth, imageHeight, totalPieces) {
    var aspectRatio = imageWidth / imageHeight;
    var cols = Math.round(Math.sqrt(totalPieces * aspectRatio));
    if (cols < 1) cols = 1;
    var rows = Math.round(totalPieces / cols);
    if (rows < 1) rows = 1;
    return [rows, cols];
  }

  // EditText arrow key increment/decrement support
  function changeValueByArrowKey(editText, allowNegative, onValueChanged) {
    editText.addEventListener("keydown", function (event) {
      var value = Number(editText.text);
      if (isNaN(value)) return;

      var keyboard = ScriptUI.environment.keyboardState;
      var delta = 1;
      var handled = false;

      if (keyboard.shiftKey) {
        delta = 10;
        if (event.keyName == "Up") {
          value = Math.ceil((value + 1) / delta) * delta;
          handled = true;
        } else if (event.keyName == "Down") {
          value = Math.floor((value - 1) / delta) * delta;
          handled = true;
        }
      } else if (keyboard.altKey) {
        delta = 0.1;
        if (event.keyName == "Up") {
          value += delta;
          handled = true;
        } else if (event.keyName == "Down") {
          value -= delta;
          handled = true;
        }
      } else {
        delta = 1;
        if (event.keyName == "Up") {
          value += delta;
          handled = true;
        } else if (event.keyName == "Down") {
          value -= delta;
          handled = true;
        }
      }

      if (!handled) return;

      if (keyboard.altKey) {
        value = Math.round(value * 10) / 10;
      } else {
        value = Math.round(value);
      }

      if (!allowNegative && value < 0) value = 0;

      event.preventDefault();
      editText.text = value;
      if (typeof onValueChanged === "function") onValueChanged(editText, value);
    });
  }

  /* ダイアログボックス作成 / Build the dialog window */
  function createDialog() {
    var dlg = new Window('dialog', L('dialogTitle') + ' ' + SCRIPT_VERSION);
    dlg.orientation = 'column';
    dlg.alignment = 'right';

    /* 共通セクション（上部） / Common section (top) */
    var commonGroup = dlg.add('group');
    commonGroup.orientation = 'column';
    commonGroup.alignChildren = 'left';
    commonGroup.margins = [10, 5, 10, 5];

    /* モード選択（パズル / グリッド分割） */
    var modeGroup = commonGroup.add("group");
    modeGroup.orientation = "row";
    modeGroup.alignChildren = "left";
    var modeRadioGrid = modeGroup.add("radiobutton", undefined, LABELS.modeGridSplit[lang]);
    var modeRadioPuzzle = modeGroup.add("radiobutton", undefined, LABELS.modePuzzle[lang]);
    modeRadioGrid.value = true;

    /* 縦積みのパネル群 / Vertically stacked panels */
    var panelsGroup = dlg.add('group');
    panelsGroup.orientation = 'column';
    panelsGroup.alignChildren = 'fill';

    /* 分割パネル: ピース数 / 列数・行数 / 形状 / オフセット / オーバーラップ */
    var splitPanel = panelsGroup.add('panel', undefined, LABELS.panelSplit[lang]);
    splitPanel.orientation = 'column';
    splitPanel.alignChildren = 'left';
    splitPanel.margins = [15, 20, 15, 10];

    /* ピース数 */
    var totalPiecesGroup = splitPanel.add("group");
    totalPiecesGroup.orientation = "row";
    totalPiecesGroup.alignment = "left";
    var totalPiecesLabel = totalPiecesGroup.add("statictext", undefined, LABELS.totalPieces[lang]);
    var totalPiecesInput = totalPiecesGroup.add("edittext", undefined, "25");
    totalPiecesInput.characters = 4;

    /* 列数・行数 */
    var rowColGroup = splitPanel.add('group');
    rowColGroup.orientation = 'row';
    rowColGroup.alignChildren = 'left';
    var colGroup = rowColGroup.add('group');
    colGroup.orientation = 'row';
    colGroup.add('statictext', undefined, LABELS.columns[lang]);
    var columnsInput = colGroup.add('edittext', undefined, "6");
    columnsInput.characters = 3;
    var rowGroup = rowColGroup.add('group');
    rowGroup.orientation = 'row';
    rowGroup.add('statictext', undefined, LABELS.rows[lang]);
    var rowsInput = rowGroup.add('edittext', undefined, "4");
    rowsInput.characters = 3;

    function getSelectedArtworkSize() {
      if (app.documents.length > 0 && app.selection.length == 1) {
        var selectedArtworkItem = app.selection[0];
        if (
          selectedArtworkItem.typename === "RasterItem" ||
          selectedArtworkItem.typename === "PlacedItem" ||
          selectedArtworkItem.typename === "SymbolItem" ||
          selectedArtworkItem.typename === "PathItem" ||
          selectedArtworkItem.typename === "GroupItem" ||
          selectedArtworkItem.typename === "CompoundPathItem"
        ) {
          var artworkBounds = selectedArtworkItem.geometricBounds;
          var artworkWidth = artworkBounds[2] - artworkBounds[0];
          var artworkHeight = artworkBounds[1] - artworkBounds[3];
          if (artworkHeight < 0) artworkHeight = -artworkHeight;
          return { width: artworkWidth, height: artworkHeight };
        }
      }
      return null;
    }
    var selectedArtwork = getSelectedArtworkSize();
    if (selectedArtwork) {
      var grid = getInitialGridSize(selectedArtwork.width, selectedArtwork.height, 25);
      rowsInput.text = String(grid[0]);
      columnsInput.text = String(grid[1]);
    }



    function updateRowsColsFromTotalPieces() {
      var val = parseInt(totalPiecesInput.text, 10);
      if (isNaN(val) || val < 1) return;
      var selectedArtwork = getSelectedArtworkSize();
      if (selectedArtwork) {
        var grid = getInitialGridSize(selectedArtwork.width, selectedArtwork.height, val);
        rowsInput.text = String(grid[0]);
        columnsInput.text = String(grid[1]);
      }
    }

    totalPiecesInput.onChanging = function () {
      updateRowsColsFromTotalPieces();
    };


    /* 形状 */
    var shapeGroup = splitPanel.add("group");
    shapeGroup.orientation = "row";
    shapeGroup.alignChildren = ["left", "top"];
    shapeGroup.margins = [0, 10, 0, 10];

    var shapeLabel = shapeGroup.add('statictext', undefined, L('shapeLabel'));
    shapeLabel.preferredSize.width = 28;

    var shapeOptions = shapeGroup.add("group");
    shapeOptions.orientation = "column";
    shapeOptions.alignChildren = "left";

    var shapeRadioTraditional = shapeOptions.add("radiobutton", undefined, LABELS.shapeTraditional[lang]);
    var shapeRadioRandom = shapeOptions.add("radiobutton", undefined, LABELS.shapeRandom[lang]);
    shapeRadioTraditional.value = true;

    /* オフセット */
    var offsetGroup = splitPanel.add("group");
    offsetGroup.orientation = "row";
    offsetGroup.alignChildren = "left";
    var offsetCheckbox = offsetGroup.add('checkbox', undefined, LABELS.offsetLabel[lang]);
    offsetCheckbox.value = false;
    var offsetValueInput = offsetGroup.add("edittext", undefined, "-2");
    offsetValueInput.characters = 4;
    var offsetUnitLabel = offsetGroup.add("statictext", undefined, getCurrentUnitLabel());
    offsetValueInput.enabled = offsetCheckbox.value;
    offsetUnitLabel.enabled = offsetCheckbox.value;
    offsetCheckbox.onClick = function () {
      var puzzle = modeRadioPuzzle.value;
      offsetValueInput.enabled = puzzle && offsetCheckbox.value;
      offsetUnitLabel.enabled = puzzle && offsetCheckbox.value;
    };

    /* オーバーラップ */
    var overlapGroup = splitPanel.add("group");
    overlapGroup.orientation = "row";
    overlapGroup.alignChildren = "left";
    var overlapCheckbox = overlapGroup.add('checkbox', undefined, LABELS.overlap[lang]);
    overlapCheckbox.value = false;
    var overlapInput = overlapGroup.add("edittext", undefined, "10");
    overlapInput.characters = 4;
    var overlapUnitLabel = overlapGroup.add("statictext", undefined, getCurrentUnitLabel());
    overlapInput.enabled = overlapCheckbox.value;
    overlapUnitLabel.enabled = overlapCheckbox.value;
    overlapCheckbox.onClick = function () {
      var gridMode = modeRadioGrid.value;
      overlapInput.enabled = gridMode && overlapCheckbox.value;
      overlapUnitLabel.enabled = gridMode && overlapCheckbox.value;
    };

    /* オプションパネル: バラけ処理 / ケイ / 角丸 */
    var optionsPanel = panelsGroup.add('panel', undefined, LABELS.panelOptions[lang]);
    optionsPanel.orientation = 'column';
    optionsPanel.alignChildren = 'left';
    optionsPanel.margins = [15, 20, 15, 10];

    /* バラけ処理 */
    var scatterGroup = optionsPanel.add("group");
    scatterGroup.orientation = "row";
    scatterGroup.alignChildren = "left";
    var scatterCheckbox = scatterGroup.add('checkbox', undefined, LABELS.explode[lang]);
    scatterCheckbox.value = false;
    var scatterStrengthInput = scatterGroup.add("edittext", undefined, "30");
    scatterStrengthInput.characters = 4;
    var scatterUnitLabel = scatterGroup.add("statictext", undefined, getCurrentUnitLabel());
    scatterStrengthInput.enabled = scatterCheckbox.value;
    scatterUnitLabel.enabled = scatterCheckbox.value;
    scatterCheckbox.onClick = function () {
      scatterStrengthInput.enabled = scatterCheckbox.value;
      scatterUnitLabel.enabled = scatterCheckbox.value;
    };

    /* ケイ */
    var ruleGroup = optionsPanel.add("group");
    ruleGroup.orientation = "row";
    ruleGroup.alignChildren = "left";
    var ruleCheckbox = ruleGroup.add('checkbox', undefined, LABELS.ruleCheck[lang]);
    ruleCheckbox.value = false;

    /* 角丸 */
    var roundCornerGroup = optionsPanel.add("group");
    roundCornerGroup.orientation = "row";
    roundCornerGroup.alignChildren = "left";
    var roundCornerCheckbox = roundCornerGroup.add('checkbox', undefined, LABELS.roundCheck[lang]);
    roundCornerCheckbox.value = false;
    var roundRadiusInput = roundCornerGroup.add("edittext", undefined, "3");
    roundRadiusInput.characters = 5;
    var roundCornerUnitLabel = roundCornerGroup.add("statictext", undefined, getCurrentUnitLabel());
    roundRadiusInput.enabled = roundCornerCheckbox.value;
    roundCornerUnitLabel.enabled = roundCornerCheckbox.value;
    roundCornerCheckbox.onClick = function () {
      var gridMode = modeRadioGrid.value;
      roundRadiusInput.enabled = gridMode && roundCornerCheckbox.value;
      roundCornerUnitLabel.enabled = gridMode && roundCornerCheckbox.value;
    };

    /* モードに応じてディム化 / Toggle enabled state by mode */
    function updateModeDependentEnabled() {
      var puzzle = modeRadioPuzzle.value;
      /* パズル時のみ有効 / Puzzle only */
      shapeGroup.enabled = puzzle;
      totalPiecesInput.enabled = puzzle;
      totalPiecesLabel.enabled = puzzle;
      offsetCheckbox.enabled = puzzle;
      offsetValueInput.enabled = puzzle && offsetCheckbox.value;
      offsetUnitLabel.enabled = puzzle && offsetCheckbox.value;
      scatterCheckbox.enabled = puzzle;
      scatterStrengthInput.enabled = puzzle && scatterCheckbox.value;
      /* グリッド分割時のみ有効 / Grid only */
      overlapCheckbox.enabled = !puzzle;
      overlapInput.enabled = !puzzle && overlapCheckbox.value;
      overlapUnitLabel.enabled = !puzzle && overlapCheckbox.value;
      ruleCheckbox.enabled = !puzzle;
      roundCornerCheckbox.enabled = !puzzle;
      roundRadiusInput.enabled = !puzzle && roundCornerCheckbox.value;
      roundCornerUnitLabel.enabled = !puzzle && roundCornerCheckbox.value;
    }

    /* モードごとの既定値を適用 / Apply mode-specific defaults */
    function applyModeDefaults() {
      if (modeRadioGrid.value) {
        totalPiecesInput.text = "2";
        offsetValueInput.text = "-2";
        offsetCheckbox.value = false;
        overlapCheckbox.value = false;
        overlapInput.text = "10";
        ruleCheckbox.value = false;
        roundCornerCheckbox.value = false;
        roundRadiusInput.text = "3";
      } else {
        totalPiecesInput.text = "25";
        offsetValueInput.text = "-2";
        offsetCheckbox.value = false;
        scatterCheckbox.value = false;
        scatterStrengthInput.text = "30";
        overlapCheckbox.value = false;
        ruleCheckbox.value = false;
        roundCornerCheckbox.value = false;
        roundRadiusInput.text = "3";
      }
    }

    function onModeChange() {
      applyModeDefaults();
      updateRowsColsFromTotalPieces();
      updateModeDependentEnabled();
    }
    modeRadioPuzzle.onClick = onModeChange;
    modeRadioGrid.onClick = onModeChange;
    /* デフォルトはグリッド分割 / Default to grid split */
    applyModeDefaults();
    updateRowsColsFromTotalPieces();
    updateModeDependentEnabled();

    /* プログレスバー（処理中のみ表示） / Progress bar (shown during processing) */
    var progressGroup = dlg.add('group');
    progressGroup.orientation = 'column';
    progressGroup.alignChildren = 'fill';
    progressGroup.margins = [10, 0, 10, 0];
    var progressBar = progressGroup.add('progressbar', undefined, 0, 100);
    progressBar.preferredSize = [200, 7];
    progressGroup.visible = false;

    /* OK・キャンセルボタン / OK and Cancel buttons */
    var buttonGroup = dlg.add('group');
    buttonGroup.orientation = 'row';
    buttonGroup.alignment = "center";
    var cancelBtn = buttonGroup.add('button', undefined, LABELS.cancel[lang], { name: "cancel" });
    var okBtn = buttonGroup.add('button', undefined, LABELS.okBtn[lang], { name: "ok" });
    okBtn.active = true;

    // Add arrow-key increment/decrement support for edittext fields
    changeValueByArrowKey(columnsInput, false);
    changeValueByArrowKey(rowsInput, false);
    changeValueByArrowKey(totalPiecesInput, false, updateRowsColsFromTotalPieces);
    changeValueByArrowKey(overlapInput, false);
    changeValueByArrowKey(roundRadiusInput, false);
    changeValueByArrowKey(offsetValueInput, true);
    changeValueByArrowKey(scatterStrengthInput, false);

    return {
      dialog: dlg,
      modeRadioPuzzle: modeRadioPuzzle,
      modeRadioGrid: modeRadioGrid,
      totalPiecesInput: totalPiecesInput,
      columnsInput: columnsInput,
      rowsInput: rowsInput,
      shapeRadioTraditional: shapeRadioTraditional,
      shapeRadioRandom: shapeRadioRandom,
      scatterCheckbox: scatterCheckbox,
      scatterStrengthInput: scatterStrengthInput,
      offsetCheckbox: offsetCheckbox,
      offsetValueInput: offsetValueInput,
      overlapCheckbox: overlapCheckbox,
      overlapInput: overlapInput,
      ruleCheckbox: ruleCheckbox,
      roundCornerCheckbox: roundCornerCheckbox,
      roundRadiusInput: roundRadiusInput,
      okBtn: okBtn,
      cancelBtn: cancelBtn,
      commonGroup: commonGroup,
      splitPanel: splitPanel,
      optionsPanel: optionsPanel,
      buttonGroup: buttonGroup,
      progressGroup: progressGroup,
      progressBar: progressBar
    };
  }

  /* オフセットパスエフェクトユーティリティ / Offset Path Effect Utility */
  function createOffsetEffectXML(offsetVal) {
    var xml = '<LiveEffect name="Adobe Offset Path"><Dict data="R mlim 4 R ofst value I jntp 2 "/></LiveEffect>';
    return xml.replace("value", offsetVal);
  }

  function applyOffsetPathToSelection(offsetVal) {
    var doc = app.activeDocument;
    var prevUIL = app.userInteractionLevel;
    if (!doc.selection || doc.selection.length === 0) {
      return null;
    }
    var prevSelection = doc.selection ? doc.selection.slice(0) : null;
    try {
      var sourceItem = doc.selection[0];

      app.userInteractionLevel = UserInteractionLevel.DONTDISPLAYALERTS;
      doc.selection = null;

      var duplicatedItem = sourceItem.duplicate(sourceItem, ElementPlacement.PLACEAFTER);
      sourceItem.remove();

      duplicatedItem.selected = true;
      duplicatedItem.applyEffect(createOffsetEffectXML(offsetVal));
      app.redraw();
      app.executeMenuCommand('expandStyle');

      if (app.selection && app.selection.length > 0) {
        return app.selection[0];
      }
      return null;
    } catch (err) {
      alert(L("alertGeneralError") + err.message);
      return null;
    } finally {
      // 必ず UIレベルと選択状態を戻す
      try { app.userInteractionLevel = prevUIL; } catch (e) {}
      try {
        if (prevSelection && prevSelection.length) {
          doc.selection = null;
          for (var i = 0; i < prevSelection.length; i++) {
            try { prevSelection[i].selected = true; } catch (e) {}
          }
        }
      } catch (e) {}
    }
  }

  /* ケイ・角丸を適用 / Apply stroke (rule) and round corners to target */
  function applyRuleAndRoundCorners(target, shouldAddStroke, shouldApplyRoundCorners, roundRadiusInPoints) {
    if (!target) return;
    if (shouldAddStroke) {
      try {
        app.activeDocument.selection = null;
        target.selected = true;
        app.executeMenuCommand('Adobe New Stroke Shortcut');
        app.executeMenuCommand('Live Pathfinder Add');
      } catch (e) { }
    }
    if (shouldApplyRoundCorners && roundRadiusInPoints > 0) {
      try {
        var roundXML = '<LiveEffect name="Adobe Round Corners"><Dict data="R radius ' + roundRadiusInPoints + ' "/></LiveEffect>';
        target.applyEffect(roundXML);
      } catch (e) { }
    }
  }


  function readDialogValues(ui) {
    var ruleCheckbox = ui.ruleCheckbox;
    var roundCornerCheckbox = ui.roundCornerCheckbox;
    var roundRadiusInput = ui.roundRadiusInput;
    var overlapCheckbox = ui.overlapCheckbox;
    var overlapInput = ui.overlapInput;
    var columnsInput = ui.columnsInput;
    var rowsInput = ui.rowsInput;
    var scatterCheckbox = ui.scatterCheckbox;
    var scatterStrengthInput = ui.scatterStrengthInput;
    var offsetCheckbox = ui.offsetCheckbox;
    var offsetValueInput = ui.offsetValueInput;

    var shouldAddStroke = ruleCheckbox.value;
    var shouldApplyRoundCorners = roundCornerCheckbox.value;
    var roundRadiusInputValue = parseFloat(roundRadiusInput.text);
    var roundRadiusInPoints = (isNaN(roundRadiusInputValue) ? 0 : roundRadiusInputValue) * getUnitToPtFactor();

    var overlapInPoints = 0;
    if (overlapCheckbox.value) {
      var overlapInputValue = parseFloat(overlapInput.text);
      overlapInPoints = (isNaN(overlapInputValue) ? 0 : overlapInputValue) * getUnitToPtFactor();
    }

    var shouldScatter = scatterCheckbox.value;
    var scatterStrengthInputValue = parseFloat(scatterStrengthInput.text);
    var scatterStrength = (isNaN(scatterStrengthInputValue) ? 0 : scatterStrengthInputValue) * getUnitToPtFactor();

    var shouldApplyOffset = offsetCheckbox.value;
    var offsetInputValue = parseFloat(offsetValueInput.text);
    var offsetInPoints = (isNaN(offsetInputValue) ? 0 : offsetInputValue) * getUnitToPtFactor();

    var columnCount = Math.round(Number(columnsInput.text));
    var rowCount = Math.round(Number(rowsInput.text));

    return {
      shouldAddStroke: shouldAddStroke,
      shouldApplyRoundCorners: shouldApplyRoundCorners,
      roundRadiusInPoints: roundRadiusInPoints,
      overlapInPoints: overlapInPoints,
      shouldScatter: shouldScatter,
      scatterStrength: scatterStrength,
      shouldApplyOffset: shouldApplyOffset,
      offsetInPoints: offsetInPoints,
      columnCount: columnCount,
      rowCount: rowCount
    };
  }

  function prepareSourceItems() {
    var workingSourceItem;
    var contentSourceItem = null;
    var maskSourceItem;
    var isTemporaryBoundsRect = false;

    /* 複数選択時はグループ化してシンボル化（重ね順を保持） */
    if (app.selection.length > 1) {
      try {
        var tempGroup = app.activeDocument.groupItems.add();
        var selectedItems = [];
        for (var selectedIndex = 0; selectedIndex < app.selection.length; selectedIndex++) {
          selectedItems.push(app.selection[selectedIndex]);
        }
        for (var selectedItemIndex = selectedItems.length - 1; selectedItemIndex >= 0; selectedItemIndex--) {
          selectedItems[selectedItemIndex].move(tempGroup, ElementPlacement.PLACEATBEGINNING);
        }
        var groupLeft = tempGroup.left;
        var groupTop = tempGroup.top;
        var groupParent = tempGroup.parent;
        var groupedSymbol = app.activeDocument.symbols.add(tempGroup);
        var symbolItem = groupParent.symbolItems.add(groupedSymbol);
        symbolItem.left = groupLeft;
        symbolItem.top = groupTop;
        tempGroup.remove();
        workingSourceItem = symbolItem;
      } catch (e) {
        alert(L("alertMultiSymbolizeFailed") + e);
        return null;
      }
    } else {
      workingSourceItem = app.selection[0];
    }
    maskSourceItem = workingSourceItem;

    /* 埋め込み画像(RasterItem)をSymbolItemに変換 */
    if (workingSourceItem.typename === "RasterItem") {
      try {
        var rasterLeft = workingSourceItem.left;
        var rasterTop = workingSourceItem.top;
        var rasterParent = workingSourceItem.parent;
        var tempSymbol = app.activeDocument.symbols.add(workingSourceItem);
        var rasterSymbolItem = rasterParent.symbolItems.add(tempSymbol);
        rasterSymbolItem.left = rasterLeft;
        rasterSymbolItem.top = rasterTop;
        workingSourceItem.remove();
        workingSourceItem = rasterSymbolItem;
      } catch (e) {
        alert(L("alertRasterSymbolizeFailed") + e);
        return null;
      }
    }
    /* ベクターオブジェクトをSymbolItemに変換 */
    else if (
      workingSourceItem.typename === "PathItem" ||
      workingSourceItem.typename === "GroupItem" ||
      workingSourceItem.typename === "CompoundPathItem"
    ) {
      try {
        var vectorLeft = workingSourceItem.left;
        var vectorTop = workingSourceItem.top;
        var vectorParent = workingSourceItem.parent;
        var vectorSymbol = app.activeDocument.symbols.add(workingSourceItem);
        var vectorSymbolItem = vectorParent.symbolItems.add(vectorSymbol);
        vectorSymbolItem.left = vectorLeft;
        vectorSymbolItem.top = vectorTop;
        workingSourceItem.remove();
        workingSourceItem = vectorSymbolItem;
      } catch (e) {
        alert(L("alertVectorSymbolizeFailed") + e);
        return null;
      }
    }

    /* PlacedItem/RasterItem/SymbolItemは矩形化 */
    var workingSourceType = workingSourceItem.typename;
    if (workingSourceType == "PlacedItem" || workingSourceType == "RasterItem" || workingSourceType == "SymbolItem") {
      var sourceBounds = workingSourceItem.geometricBounds;
      var rectWidth = sourceBounds[2] - sourceBounds[0];
      var rectHeight = sourceBounds[1] - sourceBounds[3];
      if (rectHeight < 0) rectHeight = -rectHeight;
      var tempBoundsRect = app.activeDocument.pathItems.rectangle(sourceBounds[1], sourceBounds[0], rectWidth, rectHeight);
      maskSourceItem = tempBoundsRect;
      contentSourceItem = workingSourceItem;
      isTemporaryBoundsRect = true;
    }

    return {
      contentSourceItem: contentSourceItem,
      maskSourceItem: maskSourceItem,
      isTemporaryBoundsRect: isTemporaryBoundsRect
    };
  }

  function createClippedPiece(contentSourceItem, maskPath) {
    var placedCopy = contentSourceItem.duplicate();
    var clippingGroup = app.activeDocument.groupItems.add();
    placedCopy.moveToBeginning(clippingGroup);
    if (maskPath && maskPath.typename === "PathItem") {
      maskPath.moveToBeginning(clippingGroup);
      maskPath.clipping = true;
      clippingGroup.clipped = true;
    } else {
      alert(L("alertMaskNotPath"));
    }
    return clippingGroup;
  }

  function isClippableContent(item) {
    return item && (item.typename === "PlacedItem" || item.typename === "SymbolItem");
  }

  function finalizePieceAppearance(pieceItem, shouldScatter, scatterStrength, shouldAddStroke, shouldApplyRoundCorners, roundRadiusInPoints) {
    if (shouldScatter && scatterStrength > 0) {
      // ガウス分布（中央寄り）+ 最大移動量でクランプ / Gaussian distribution (center-biased) with max-offset clamp
      function randomGaussian() {
        var u = 0, v = 0;
        while (u === 0) u = Math.random();
        while (v === 0) v = Math.random();
        return Math.sqrt(-2.0 * Math.log(u)) * Math.cos(2.0 * Math.PI * v);
      }

      function clamp(value, minValue, maxValue) {
        return Math.max(minValue, Math.min(maxValue, value));
      }

      var maxOffset = scatterStrength;
      var offsetX = clamp(randomGaussian() * scatterStrength * 0.5, -maxOffset, maxOffset);
      var offsetY = clamp(randomGaussian() * scatterStrength * 0.5, -maxOffset, maxOffset);

      var moveMatrix = app.getTranslationMatrix(offsetX, offsetY);
      pieceItem.transform(moveMatrix);
    }
    pieceItem.selected = true;
    applyRuleAndRoundCorners(pieceItem, shouldAddStroke, shouldApplyRoundCorners, roundRadiusInPoints);
  }

  function cleanupSourceItems(contentSourceItem, maskSourceItem, isTemporaryBoundsRect) {
    if (contentSourceItem && typeof contentSourceItem.remove === "function") {
      contentSourceItem.remove();
    }
    if (isTemporaryBoundsRect && maskSourceItem && typeof maskSourceItem.remove === "function") {
      maskSourceItem.remove();
    }
  }

  function createGridMask(originX, originY, columnIndex, rowIndex, pieceWidth, pieceHeight, overlapInPoints, imgLeft, imgRight, imgTop, imgBottom) {
    /* オーバーラップ考慮の矩形計算 / Rect with overlap, clamped to image bounds */
    var rectLeft = originX + columnIndex * pieceWidth - overlapInPoints / 2;
    var rectWidthActual = pieceWidth + overlapInPoints;
    if (rectLeft < imgLeft) {
      rectWidthActual -= (imgLeft - rectLeft);
      rectLeft = imgLeft;
    }
    if (rectLeft + rectWidthActual > imgRight) {
      rectWidthActual = imgRight - rectLeft;
    }

    var rectTop = originY - rowIndex * pieceHeight + overlapInPoints / 2;
    var rectHeightActual = pieceHeight + overlapInPoints;
    if (rectTop > imgTop) {
      rectHeightActual -= (rectTop - imgTop);
      rectTop = imgTop;
    }
    if (rectTop - rectHeightActual < imgBottom) {
      rectHeightActual = rectTop - imgBottom;
    }

    var gridMask = app.activeDocument.pathItems.rectangle(
      rectTop,
      rectLeft,
      rectWidthActual,
      rectHeightActual
    );
    gridMask.closed = true;
    gridMask.filled = false;
    gridMask.stroked = false;
    return gridMask;
  }

  function addPoint(obj, x, y) {
    var np = obj.pathPoints.add();
    np.anchor = [x, y];
    np.leftDirection = [x, y];
    np.rightDirection = [x, y];
    np.pointType = PointType.CORNER;
  }

  function addCurvePoint(pathItem, anchorX, anchorY, leftX, leftY, rightX, rightY) {
    var curvePoint = pathItem.pathPoints.add();
    curvePoint.anchor = [anchorX, anchorY];
    curvePoint.leftDirection = [leftX, leftY];
    curvePoint.rightDirection = [rightX, rightY];
    curvePoint.pointType = PointType.SMOOTH;
    return curvePoint;
  }

  function createPuzzleMaskPath(originX, originY, x, y, pieceWidth, pieceHeight, columnCount, rowCount, oneThirdWide, oneQuarterHigh, oneThirdHigh, oneQuarterWide, pieceEdgeData) {
    var maskPath = app.activeDocument.pathItems.add();
    addPoint(maskPath, originX + x * pieceWidth, originY - (y + 1) * pieceHeight);
    appendTopEdge(maskPath, originX, originY, x, y, pieceWidth, pieceHeight, rowCount, oneThirdWide, oneQuarterHigh, pieceEdgeData);
    addPoint(maskPath, originX + (x + 1) * pieceWidth, originY - (y + 1) * pieceHeight);
    appendRightEdge(maskPath, originX, originY, x, y, pieceWidth, pieceHeight, columnCount, oneThirdHigh, oneQuarterWide, pieceEdgeData);
    addPoint(maskPath, originX + (x + 1) * pieceWidth, originY - y * pieceHeight);
    appendBottomEdge(maskPath, originX, originY, x, y, pieceWidth, pieceHeight, oneThirdWide, oneQuarterHigh, pieceEdgeData);
    addPoint(maskPath, originX + x * pieceWidth, originY - y * pieceHeight);
    appendLeftEdge(maskPath, originX, originY, x, y, pieceWidth, pieceHeight, oneThirdHigh, oneQuarterWide, pieceEdgeData);

    maskPath.closed = true;

    return maskPath;
  }

  function appendTopEdge(maskPath, originX, originY, x, y, pieceWidth, pieceHeight, rowCount, oneThirdWide, oneQuarterHigh, pieceEdgeData) {
    /* 上辺突起 */
    if (y < rowCount - 1) {
      if (pieceEdgeData[y + 1][x].topOut) {
        addCurvePoint(maskPath,
          originX + x * pieceWidth + oneThirdWide, originY - (y + 1) * pieceHeight - pieceEdgeData[y + 1][x].verticalOffset,
          originX + x * pieceWidth + 0.67 * oneThirdWide, originY - (y + 1) * pieceHeight + 0.33 * oneQuarterHigh - pieceEdgeData[y + 1][x].verticalOffset,
          originX + x * pieceWidth + 1.33 * oneThirdWide, originY - (y + 1) * pieceHeight - 0.33 * oneQuarterHigh - pieceEdgeData[y + 1][x].verticalOffset
        );

        addCurvePoint(maskPath,
          originX + x * pieceWidth + oneThirdWide, originY - (y + 1) * pieceHeight - oneQuarterHigh - pieceEdgeData[y + 1][x].verticalOffset,
          originX + x * pieceWidth + 0.67 * oneThirdWide, originY - (y + 1) * pieceHeight - oneQuarterHigh + 0.33 * oneQuarterHigh - pieceEdgeData[y + 1][x].verticalOffset,
          originX + x * pieceWidth + 1.33 * oneThirdWide, originY - (y + 1) * pieceHeight - oneQuarterHigh - 0.33 * oneQuarterHigh - pieceEdgeData[y + 1][x].verticalOffset
        );

        addCurvePoint(maskPath,
          originX + x * pieceWidth + 2 * oneThirdWide, originY - (y + 1) * pieceHeight - oneQuarterHigh - pieceEdgeData[y + 1][x].verticalOffset,
          originX + x * pieceWidth + 1.67 * oneThirdWide, originY - (y + 1) * pieceHeight - oneQuarterHigh - 0.33 * oneQuarterHigh - pieceEdgeData[y + 1][x].verticalOffset,
          originX + x * pieceWidth + 2.33 * oneThirdWide, originY - (y + 1) * pieceHeight - oneQuarterHigh + 0.33 * oneQuarterHigh - pieceEdgeData[y + 1][x].verticalOffset
        );

        addCurvePoint(maskPath,
          originX + x * pieceWidth + 2 * oneThirdWide, originY - (y + 1) * pieceHeight - pieceEdgeData[y + 1][x].verticalOffset,
          originX + x * pieceWidth + 1.67 * oneThirdWide, originY - (y + 1) * pieceHeight - 0.33 * oneQuarterHigh - pieceEdgeData[y + 1][x].verticalOffset,
          originX + x * pieceWidth + 2.33 * oneThirdWide, originY - (y + 1) * pieceHeight + 0.33 * oneQuarterHigh - pieceEdgeData[y + 1][x].verticalOffset
        );
      } else {
        addCurvePoint(maskPath,
          originX + x * pieceWidth + oneThirdWide, originY - (y + 1) * pieceHeight - pieceEdgeData[y + 1][x].verticalOffset,
          originX + x * pieceWidth + 0.67 * oneThirdWide, originY - (y + 1) * pieceHeight - 0.33 * oneQuarterHigh - pieceEdgeData[y + 1][x].verticalOffset,
          originX + x * pieceWidth + 1.33 * oneThirdWide, originY - (y + 1) * pieceHeight + 0.33 * oneQuarterHigh - pieceEdgeData[y + 1][x].verticalOffset
        );

        addCurvePoint(maskPath,
          originX + x * pieceWidth + oneThirdWide, originY - y * pieceHeight - 3 * oneQuarterHigh - pieceEdgeData[y + 1][x].verticalOffset,
          originX + x * pieceWidth + 0.67 * oneThirdWide, originY - y * pieceHeight - 3.5 * oneQuarterHigh - pieceEdgeData[y + 1][x].verticalOffset,
          originX + x * pieceWidth + 1.33 * oneThirdWide, originY - y * pieceHeight - 2.5 * oneQuarterHigh - pieceEdgeData[y + 1][x].verticalOffset
        );

        addCurvePoint(maskPath,
          originX + x * pieceWidth + 2 * oneThirdWide, originY - y * pieceHeight - 3 * oneQuarterHigh - pieceEdgeData[y + 1][x].verticalOffset,
          originX + x * pieceWidth + 1.67 * oneThirdWide, originY - y * pieceHeight - 2.5 * oneQuarterHigh - pieceEdgeData[y + 1][x].verticalOffset,
          originX + x * pieceWidth + 2.33 * oneThirdWide, originY - y * pieceHeight - 3.5 * oneQuarterHigh - pieceEdgeData[y + 1][x].verticalOffset
        );

        addCurvePoint(maskPath,
          originX + x * pieceWidth + 2 * oneThirdWide, originY - (y + 1) * pieceHeight - pieceEdgeData[y + 1][x].verticalOffset,
          originX + x * pieceWidth + 1.67 * oneThirdWide, originY - (y + 1) * pieceHeight + 0.33 * oneQuarterHigh - pieceEdgeData[y + 1][x].verticalOffset,
          originX + x * pieceWidth + 2.33 * oneThirdWide, originY - (y + 1) * pieceHeight - 0.33 * oneQuarterHigh - pieceEdgeData[y + 1][x].verticalOffset
        );
      }
    }
  }

  function appendRightEdge(maskPath, originX, originY, x, y, pieceWidth, pieceHeight, columnCount, oneThirdHigh, oneQuarterWide, pieceEdgeData) {
    /* 右辺突起 */
    if (x < columnCount - 1) {
      if (pieceEdgeData[y][x + 1].rightOut) {
        addCurvePoint(maskPath,
          originX + x * pieceWidth + 4 * oneQuarterWide - pieceEdgeData[y][x + 1].horizontalOffset, originY - y * pieceHeight - 2 * oneThirdHigh,
          originX + x * pieceWidth + 3.5 * oneQuarterWide - pieceEdgeData[y][x + 1].horizontalOffset, originY - y * pieceHeight - 2.33 * oneThirdHigh,
          originX + x * pieceWidth + 4.5 * oneQuarterWide - pieceEdgeData[y][x + 1].horizontalOffset, originY - y * pieceHeight - 1.67 * oneThirdHigh
        );

        addCurvePoint(maskPath,
          originX + x * pieceWidth + 5 * oneQuarterWide - pieceEdgeData[y][x + 1].horizontalOffset, originY - y * pieceHeight - 2 * oneThirdHigh,
          originX + x * pieceWidth + 4.5 * oneQuarterWide - pieceEdgeData[y][x + 1].horizontalOffset, originY - y * pieceHeight - 2.33 * oneThirdHigh,
          originX + x * pieceWidth + 5.5 * oneQuarterWide - pieceEdgeData[y][x + 1].horizontalOffset, originY - y * pieceHeight - 1.67 * oneThirdHigh
        );

        addCurvePoint(maskPath,
          originX + x * pieceWidth + 5 * oneQuarterWide - pieceEdgeData[y][x + 1].horizontalOffset, originY - y * pieceHeight - 1 * oneThirdHigh,
          originX + x * pieceWidth + 5.5 * oneQuarterWide - pieceEdgeData[y][x + 1].horizontalOffset, originY - y * pieceHeight - 1.33 * oneThirdHigh,
          originX + x * pieceWidth + 4.5 * oneQuarterWide - pieceEdgeData[y][x + 1].horizontalOffset, originY - y * pieceHeight - 0.67 * oneThirdHigh
        );

        addCurvePoint(maskPath,
          originX + x * pieceWidth + 4 * oneQuarterWide - pieceEdgeData[y][x + 1].horizontalOffset, originY - y * pieceHeight - oneThirdHigh,
          originX + x * pieceWidth + 4.5 * oneQuarterWide - pieceEdgeData[y][x + 1].horizontalOffset, originY - y * pieceHeight - 1.33 * oneThirdHigh,
          originX + x * pieceWidth + 3.5 * oneQuarterWide - pieceEdgeData[y][x + 1].horizontalOffset, originY - y * pieceHeight - 0.67 * oneThirdHigh
        );
      } else {
        addCurvePoint(maskPath,
          originX + x * pieceWidth + 4 * oneQuarterWide - pieceEdgeData[y][x + 1].horizontalOffset, originY - y * pieceHeight - 2 * oneThirdHigh,
          originX + x * pieceWidth + 4.5 * oneQuarterWide - pieceEdgeData[y][x + 1].horizontalOffset, originY - y * pieceHeight - 2.33 * oneThirdHigh,
          originX + x * pieceWidth + 3.5 * oneQuarterWide - pieceEdgeData[y][x + 1].horizontalOffset, originY - y * pieceHeight - 1.67 * oneThirdHigh
        );

        addCurvePoint(maskPath,
          originX + x * pieceWidth + 3 * oneQuarterWide - pieceEdgeData[y][x + 1].horizontalOffset, originY - y * pieceHeight - 2 * oneThirdHigh,
          originX + x * pieceWidth + 3.5 * oneQuarterWide - pieceEdgeData[y][x + 1].horizontalOffset, originY - y * pieceHeight - 2.33 * oneThirdHigh,
          originX + x * pieceWidth + 2.5 * oneQuarterWide - pieceEdgeData[y][x + 1].horizontalOffset, originY - y * pieceHeight - 1.67 * oneThirdHigh
        );

        addCurvePoint(maskPath,
          originX + x * pieceWidth + 3 * oneQuarterWide - pieceEdgeData[y][x + 1].horizontalOffset, originY - y * pieceHeight - 1 * oneThirdHigh,
          originX + x * pieceWidth + 2.5 * oneQuarterWide - pieceEdgeData[y][x + 1].horizontalOffset, originY - y * pieceHeight - 1.33 * oneThirdHigh,
          originX + x * pieceWidth + 3.5 * oneQuarterWide - pieceEdgeData[y][x + 1].horizontalOffset, originY - y * pieceHeight - 0.67 * oneThirdHigh
        );

        addCurvePoint(maskPath,
          originX + x * pieceWidth + 4 * oneQuarterWide - pieceEdgeData[y][x + 1].horizontalOffset, originY - y * pieceHeight - oneThirdHigh,
          originX + x * pieceWidth + 3.5 * oneQuarterWide - pieceEdgeData[y][x + 1].horizontalOffset, originY - y * pieceHeight - 1.33 * oneThirdHigh,
          originX + x * pieceWidth + 4.5 * oneQuarterWide - pieceEdgeData[y][x + 1].horizontalOffset, originY - y * pieceHeight - 0.67 * oneThirdHigh
        );
      }
    }
  }

  function appendBottomEdge(maskPath, originX, originY, x, y, pieceWidth, pieceHeight, oneThirdWide, oneQuarterHigh, pieceEdgeData) {
    /* 下辺突起 */
    if (y > 0) {
      if (pieceEdgeData[y][x].topOut) {
        addCurvePoint(maskPath,
          originX + x * pieceWidth + 2 * oneThirdWide, originY - y * pieceHeight - pieceEdgeData[y][x].verticalOffset,
          originX + x * pieceWidth + 2.33 * oneThirdWide, originY - y * pieceHeight + 0.33 * oneQuarterHigh - pieceEdgeData[y][x].verticalOffset,
          originX + x * pieceWidth + 1.67 * oneThirdWide, originY - y * pieceHeight - 0.33 * oneQuarterHigh - pieceEdgeData[y][x].verticalOffset
        );

        addCurvePoint(maskPath,
          originX + x * pieceWidth + 2 * oneThirdWide, originY - y * pieceHeight - oneQuarterHigh - pieceEdgeData[y][x].verticalOffset,
          originX + x * pieceWidth + 2.33 * oneThirdWide, originY - y * pieceHeight - oneQuarterHigh + 0.33 * oneQuarterHigh - pieceEdgeData[y][x].verticalOffset,
          originX + x * pieceWidth + 1.67 * oneThirdWide, originY - y * pieceHeight - oneQuarterHigh - 0.33 * oneQuarterHigh - pieceEdgeData[y][x].verticalOffset
        );

        addCurvePoint(maskPath,
          originX + x * pieceWidth + oneThirdWide, originY - y * pieceHeight - oneQuarterHigh - pieceEdgeData[y][x].verticalOffset,
          originX + x * pieceWidth + 1.33 * oneThirdWide, originY - y * pieceHeight - oneQuarterHigh - 0.33 * oneQuarterHigh - pieceEdgeData[y][x].verticalOffset,
          originX + x * pieceWidth + 0.67 * oneThirdWide, originY - y * pieceHeight - oneQuarterHigh + 0.33 * oneQuarterHigh - pieceEdgeData[y][x].verticalOffset
        );

        addCurvePoint(maskPath,
          originX + x * pieceWidth + oneThirdWide, originY - y * pieceHeight - pieceEdgeData[y][x].verticalOffset,
          originX + x * pieceWidth + 1.33 * oneThirdWide, originY - y * pieceHeight - 0.33 * oneQuarterHigh - pieceEdgeData[y][x].verticalOffset,
          originX + x * pieceWidth + 0.67 * oneThirdWide, originY - y * pieceHeight + 0.33 * oneQuarterHigh - pieceEdgeData[y][x].verticalOffset
        );
      } else {
        addCurvePoint(maskPath,
          originX + x * pieceWidth + 2 * oneThirdWide, originY - y * pieceHeight - pieceEdgeData[y][x].verticalOffset,
          originX + x * pieceWidth + 2.33 * oneThirdWide, originY - y * pieceHeight - 0.33 * oneQuarterHigh - pieceEdgeData[y][x].verticalOffset,
          originX + x * pieceWidth + 1.67 * oneThirdWide, originY - y * pieceHeight + 0.33 * oneQuarterHigh - pieceEdgeData[y][x].verticalOffset
        );

        addCurvePoint(maskPath,
          originX + x * pieceWidth + 2 * oneThirdWide, originY - y * pieceHeight + oneQuarterHigh - pieceEdgeData[y][x].verticalOffset,
          originX + x * pieceWidth + 2.33 * oneThirdWide, originY - y * pieceHeight + 0.5 * oneQuarterHigh - pieceEdgeData[y][x].verticalOffset,
          originX + x * pieceWidth + 1.67 * oneThirdWide, originY - y * pieceHeight + 1.5 * oneQuarterHigh - pieceEdgeData[y][x].verticalOffset
        );

        addCurvePoint(maskPath,
          originX + x * pieceWidth + oneThirdWide, originY - y * pieceHeight + oneQuarterHigh - pieceEdgeData[y][x].verticalOffset,
          originX + x * pieceWidth + 1.33 * oneThirdWide, originY - y * pieceHeight + 1.5 * oneQuarterHigh - pieceEdgeData[y][x].verticalOffset,
          originX + x * pieceWidth + 0.67 * oneThirdWide, originY - y * pieceHeight + 0.5 * oneQuarterHigh - pieceEdgeData[y][x].verticalOffset
        );

        addCurvePoint(maskPath,
          originX + x * pieceWidth + oneThirdWide, originY - y * pieceHeight - pieceEdgeData[y][x].verticalOffset,
          originX + x * pieceWidth + 1.33 * oneThirdWide, originY - y * pieceHeight + 0.33 * oneQuarterHigh - pieceEdgeData[y][x].verticalOffset,
          originX + x * pieceWidth + 0.67 * oneThirdWide, originY - y * pieceHeight - 0.33 * oneQuarterHigh - pieceEdgeData[y][x].verticalOffset
        );
      }
    }
  }

  function appendLeftEdge(maskPath, originX, originY, x, y, pieceWidth, pieceHeight, oneThirdHigh, oneQuarterWide, pieceEdgeData) {
    /* 左辺突起 */
    if (x > 0) {
      if (pieceEdgeData[y][x].rightOut) {
        addCurvePoint(maskPath,
          originX + x * pieceWidth - pieceEdgeData[y][x].horizontalOffset, originY - y * pieceHeight - oneThirdHigh,
          originX + x * pieceWidth - 0.5 * oneQuarterWide - pieceEdgeData[y][x].horizontalOffset, originY - y * pieceHeight - 0.67 * oneThirdHigh,
          originX + x * pieceWidth + 0.5 * oneQuarterWide - pieceEdgeData[y][x].horizontalOffset, originY - y * pieceHeight - 1.33 * oneThirdHigh
        );

        addCurvePoint(maskPath,
          originX + x * pieceWidth + oneQuarterWide - pieceEdgeData[y][x].horizontalOffset, originY - y * pieceHeight - 1 * oneThirdHigh,
          originX + x * pieceWidth + 0.5 * oneQuarterWide - pieceEdgeData[y][x].horizontalOffset, originY - y * pieceHeight - 0.67 * oneThirdHigh,
          originX + x * pieceWidth + 1.5 * oneQuarterWide - pieceEdgeData[y][x].horizontalOffset, originY - y * pieceHeight - 1.33 * oneThirdHigh
        );

        addCurvePoint(maskPath,
          originX + x * pieceWidth + oneQuarterWide - pieceEdgeData[y][x].horizontalOffset, originY - y * pieceHeight - 2 * oneThirdHigh,
          originX + x * pieceWidth + 1.5 * oneQuarterWide - pieceEdgeData[y][x].horizontalOffset, originY - y * pieceHeight - 1.67 * oneThirdHigh,
          originX + x * pieceWidth + 0.5 * oneQuarterWide - pieceEdgeData[y][x].horizontalOffset, originY - y * pieceHeight - 2.33 * oneThirdHigh
        );

        addCurvePoint(maskPath,
          originX + x * pieceWidth - pieceEdgeData[y][x].horizontalOffset, originY - y * pieceHeight - 2 * oneThirdHigh,
          originX + x * pieceWidth + 0.5 * oneQuarterWide - pieceEdgeData[y][x].horizontalOffset, originY - y * pieceHeight - 1.67 * oneThirdHigh,
          originX + x * pieceWidth - 0.5 * oneQuarterWide - pieceEdgeData[y][x].horizontalOffset, originY - y * pieceHeight - 2.33 * oneThirdHigh
        );
      } else {
        addCurvePoint(maskPath,
          originX + x * pieceWidth - pieceEdgeData[y][x].horizontalOffset, originY - y * pieceHeight - oneThirdHigh,
          originX + x * pieceWidth + 0.5 * oneQuarterWide - pieceEdgeData[y][x].horizontalOffset, originY - y * pieceHeight - 0.67 * oneThirdHigh,
          originX + x * pieceWidth - 0.5 * oneQuarterWide - pieceEdgeData[y][x].horizontalOffset, originY - y * pieceHeight - 1.33 * oneThirdHigh
        );

        addCurvePoint(maskPath,
          originX + x * pieceWidth - oneQuarterWide - pieceEdgeData[y][x].horizontalOffset, originY - y * pieceHeight - 1 * oneThirdHigh,
          originX + x * pieceWidth - 0.5 * oneQuarterWide - pieceEdgeData[y][x].horizontalOffset, originY - y * pieceHeight - 0.67 * oneThirdHigh,
          originX + x * pieceWidth - 1.5 * oneQuarterWide - pieceEdgeData[y][x].horizontalOffset, originY - y * pieceHeight - 1.33 * oneThirdHigh
        );

        addCurvePoint(maskPath,
          originX + x * pieceWidth - oneQuarterWide - pieceEdgeData[y][x].horizontalOffset, originY - y * pieceHeight - 2 * oneThirdHigh,
          originX + x * pieceWidth - 1.5 * oneQuarterWide - pieceEdgeData[y][x].horizontalOffset, originY - y * pieceHeight - 1.67 * oneThirdHigh,
          originX + x * pieceWidth - 0.5 * oneQuarterWide - pieceEdgeData[y][x].horizontalOffset, originY - y * pieceHeight - 2.33 * oneThirdHigh
        );

        addCurvePoint(maskPath,
          originX + x * pieceWidth - pieceEdgeData[y][x].horizontalOffset, originY - y * pieceHeight - 2 * oneThirdHigh,
          originX + x * pieceWidth - 0.5 * oneQuarterWide - pieceEdgeData[y][x].horizontalOffset, originY - y * pieceHeight - 1.67 * oneThirdHigh,
          originX + x * pieceWidth + 0.5 * oneQuarterWide - pieceEdgeData[y][x].horizontalOffset, originY - y * pieceHeight - 2.33 * oneThirdHigh
        );
      }
    }
  }

  function executeSlice(ui, onProgress) {
    var modeRadioGrid = ui.modeRadioGrid;
    var shapeRadioRandom = ui.shapeRadioRandom;

    /* ダイアログ値取得 / Read dialog values */
    var dialogValues = readDialogValues(ui);
    var shouldAddStroke = dialogValues.shouldAddStroke;
    var shouldApplyRoundCorners = dialogValues.shouldApplyRoundCorners;
    var roundRadiusInPoints = dialogValues.roundRadiusInPoints;
    var overlapInPoints = dialogValues.overlapInPoints;
    var shouldScatter = dialogValues.shouldScatter;
    var scatterStrength = dialogValues.scatterStrength;
    var shouldApplyOffset = dialogValues.shouldApplyOffset;
    var offsetInPoints = dialogValues.offsetInPoints;

    // 選択オブジェクト取得 / Get selected object
    var preparedItems = prepareSourceItems();
    if (!preparedItems) {
      return;
    }

    var contentSourceItem = preparedItems.contentSourceItem;
    var maskSourceItem = preparedItems.maskSourceItem;
    var isTemporaryBoundsRect = preparedItems.isTemporaryBoundsRect;

    /* 列・行数取得 */
    var columnCount = dialogValues.columnCount;
    var rowCount = dialogValues.rowCount;

    if (
      (columnCount >= 1 && rowCount >= 1) ||
      (columnCount == 0 && rowCount >= 1) ||
      (columnCount >= 1 && rowCount == 0)
    ) {
      maskSourceItem.selected = false;

      var aspectRatio;
      if (columnCount == 0) {
        aspectRatio = Math.abs(
          (maskSourceItem.geometricBounds[2] - maskSourceItem.geometricBounds[0]) /
          (maskSourceItem.geometricBounds[3] - maskSourceItem.geometricBounds[1])
        );
        columnCount = Math.round(aspectRatio * rowCount);
      }
      if (rowCount == 0) {
        aspectRatio = Math.abs(
          (maskSourceItem.geometricBounds[3] - maskSourceItem.geometricBounds[1]) /
          (maskSourceItem.geometricBounds[2] - maskSourceItem.geometricBounds[0])
        );
        rowCount = Math.round(aspectRatio * columnCount);
      }

      var maskBounds = maskSourceItem.geometricBounds;
      var maskLeft = maskBounds[0];
      var maskTop = maskBounds[1];
      var maskRight = maskBounds[2];
      var maskBottom = maskBounds[3];

      var originX = maskLeft;
      var originY = maskTop;

      var pieceWidth = (maskRight - maskLeft) / columnCount;
      var pieceHeight = (maskTop - maskBottom) / rowCount;

      var oneThirdWide = pieceWidth / 3;
      var oneQuarterHigh = pieceHeight / 4;
      var oneThirdHigh = pieceHeight / 3;
      var oneQuarterWide = pieceWidth / 4;

      var pieceEdgeData = new Array(rowCount);
      if (shapeRadioRandom.value) {
        for (var y = 0; y < rowCount; y++) {
          pieceEdgeData[y] = new Array(columnCount);
          for (var x = 0; x < columnCount; x++) {
            pieceEdgeData[y][x] = {
              topOut: Math.random() < 0.5,
              rightOut: Math.random() < 0.5,
              verticalOffset: pieceHeight * (Math.random() - 0.5) / 10,
              horizontalOffset: pieceWidth * (Math.random() - 0.5) / 10
            };
          }
        }
      } else {
        for (var y = 0; y < rowCount; y++) {
          pieceEdgeData[y] = new Array(columnCount);
          for (var x = 0; x < columnCount; x++) {
            pieceEdgeData[y][x] = {
              topOut: (x & 1) ^ (y & 1),
              rightOut: !((x & 1) ^ (y & 1)),
              verticalOffset: pieceHeight * (Math.random() - 0.5) / 10,
              horizontalOffset: pieceWidth * (Math.random() - 0.5) / 10
            };
          }
        }
      }

      var totalPieces = rowCount * columnCount;
      var piecesDone = 0;
      if (onProgress) onProgress(0, totalPieces);

      /* グリッド形状（矩形マスク）モード */
      if (modeRadioGrid.value) {
        var imgLeft = maskLeft;
        var imgRight = maskRight;
        var imgTop = maskTop;
        var imgBottom = maskBottom;
        for (var y = 0; y < rowCount; y++) {
          for (var x = 0; x < columnCount; x++) {
            var gridMask = createGridMask(
              originX,
              originY,
              x,
              y,
              pieceWidth,
              pieceHeight,
              overlapInPoints,
              imgLeft,
              imgRight,
              imgTop,
              imgBottom
            );
            if (isClippableContent(contentSourceItem)) {
              var clippedPiece = createClippedPiece(contentSourceItem, gridMask);
              finalizePieceAppearance(clippedPiece, shouldScatter, scatterStrength, shouldAddStroke, shouldApplyRoundCorners, roundRadiusInPoints);
            } else {
              finalizePieceAppearance(gridMask, shouldScatter, scatterStrength, shouldAddStroke, shouldApplyRoundCorners, roundRadiusInPoints);
            }
            piecesDone++;
            if (onProgress) onProgress(piecesDone, totalPieces);
          }
        }

        cleanupSourceItems(contentSourceItem, maskSourceItem, isTemporaryBoundsRect);
        return;
      }

      /* ジグソー形状モード */
      for (var y = 0; y < rowCount; y++) {
        for (var x = 0; x < columnCount; x++) {
          var maskPath = createPuzzleMaskPath(
            originX,
            originY,
            x,
            y,
            pieceWidth,
            pieceHeight,
            columnCount,
            rowCount,
            oneThirdWide,
            oneQuarterHigh,
            oneThirdHigh,
            oneQuarterWide,
            pieceEdgeData
          );

          /* オフセット効果を適用 */
          if (shouldApplyOffset && maskPath) {
            var offsetVal = offsetInPoints;
            app.selection = null;
            try {
              maskPath.selected = true;
              var offsetResultItem = applyOffsetPathToSelection(offsetVal);
              if (offsetResultItem) {
                if (offsetResultItem.typename === "PathItem") {
                  maskPath = offsetResultItem;
                } else if (offsetResultItem.typename === "GroupItem") {
                  for (var pageItemIndex = 0; pageItemIndex < offsetResultItem.pageItems.length; pageItemIndex++) {
                    if (offsetResultItem.pageItems[pageItemIndex].typename === "PathItem") {
                      maskPath = offsetResultItem.pageItems[pageItemIndex];
                      break;
                    }
                  }
                  if (maskPath.typename !== "PathItem") {
                    alert(L("alertOffsetGroupNoPath"));
                  }
                } else {
                  alert(L("alertOffsetUnexpectedType"));
                }
              }
            } catch (e) {
              alert(L("alertOffsetError") + e.message);
            }
          }

          /* 画像クリッピンググループ処理 */
          if (isClippableContent(contentSourceItem)) {
            var clippedPiece = createClippedPiece(contentSourceItem, maskPath);
            finalizePieceAppearance(clippedPiece, shouldScatter, scatterStrength, shouldAddStroke, shouldApplyRoundCorners, roundRadiusInPoints);
          } else {
            finalizePieceAppearance(maskPath, shouldScatter, scatterStrength, shouldAddStroke, shouldApplyRoundCorners, roundRadiusInPoints);
          }

          piecesDone++;
          if (onProgress) onProgress(piecesDone, totalPieces);
        }
      }

      cleanupSourceItems(contentSourceItem, maskSourceItem, isTemporaryBoundsRect);
    }
  }

  function main() {
    try {
      /* Illustrator状態チェック / Illustrator state check */
      if (app.documents.length === 0 || app.selection.length === 0) {
        return;
      }

      var ui = createDialog();
      var puzzlifyDialog = ui.dialog;

      ui.okBtn.onClick = function () {
        /* 入力チェック / Input validation */
        if (!(ui.columnsInput.text || ui.rowsInput.text)) {
          return;
        }

        /* プログレス表示モードへ切替 / Switch to progress display */
        ui.commonGroup.enabled = false;
        ui.splitPanel.enabled = false;
        ui.optionsPanel.enabled = false;
        ui.buttonGroup.visible = false;
        ui.progressGroup.visible = true;
        ui.progressBar.value = 0;
        puzzlifyDialog.layout.layout(true);
        puzzlifyDialog.update();

        try {
          executeSlice(ui, function (current, total) {
            ui.progressBar.value = total > 0 ? (current / total) * 100 : 0;
            puzzlifyDialog.update();
          });
        } catch (e) {
          alert(L("alertScriptError") + e);
        }

        /* 処理完了後にダイアログを閉じる / Close dialog after processing completes */
        puzzlifyDialog.close(1);
      };

      puzzlifyDialog.show();
    } catch (e) {
      alert(L("alertScriptError") + e);
    }
  }

  main();

})();