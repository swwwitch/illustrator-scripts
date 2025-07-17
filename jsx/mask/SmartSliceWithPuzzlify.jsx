#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false); 

/*
### スクリプト名：

SmartSliceWithPuzzlify.jsx

### 概要

- 選択した画像や図形をグリッドまたはジグソーパズル形状に分割し、それぞれにマスクを適用するIllustrator用スクリプトです。
- 任意の行数・列数設定、オフセットやバラけ効果など柔軟なカスタマイズが可能です。

### 主な機能

- グリッド分割、トラディショナル、ランダムなジグソー形状の生成
- 各ピースに対するオフセットや散乱（バラけ）効果の適用
- 画像、シンボル、ベクターオブジェクトなど多様なオブジェクトに対応
- 自動行列算出、複数選択オブジェクトの自動シンボル化処理
- 日本語／英語インターフェース対応

### 処理の流れ

1. 対象オブジェクト（画像、シンボル、ベクターなど）を選択
2. ダイアログで分割設定（形状、ピース数、行数、列数、オフセットなど）を指定
3. 形状に応じた各ピースを生成
4. 必要に応じてオフセットやバラけを適用
5. 元オブジェクトを削除し、新しいピースを配置

### オリジナル、謝辞

Originally created by Jongware on 7-Oct-2010  
https://community.adobe.com/t5/illustrator-discussions/cut-multiple-jigsaw-shapes-out-of-image-simultaneously/td-p/8709621#9185144

### 更新履歴

- v1.0 (20250607) : 初期バージョン
- v1.1 (20250608) : シンボル対応、ベクターアートワーク対応
- v1.2 (20250609) : グリッド形状対応
- v1.3 (20250610) : オフセット機能追加、単位コード対応

---

### Script Name:

SmartSliceWithPuzzlify.jsx

### Overview

- An Illustrator script to split a selected image or shape into a grid or jigsaw puzzle pieces, each with individual masks.
- Supports custom rows/columns, offsets, and scatter (explode) effects.

### Main Features

- Generate grid, traditional, or random jigsaw-shaped pieces
- Apply offset and scatter (explode) effects to each piece
- Supports images, symbols, vector artwork, and more
- Auto calculate rows/columns, auto symbolization for multiple selections
- Japanese and English UI support

### Process Flow

1. Select a target object (image, symbol, vector, etc.)
2. Configure split settings (shape type, number of pieces, rows, columns, offset, etc.) via dialog
3. Generate each piece based on chosen shape
4. Optionally apply offset and scatter effects
5. Remove original object and place new pieces

### Original / Acknowledgements

Originally created by Jongware on 7-Oct-2010  
https://community.adobe.com/t5/illustrator-discussions/cut-multiple-jigsaw-shapes-out-of-image-simultaneously/td-p/8709621#9185144

### Update History

- v1.0 (20250607): Initial version
- v1.1 (20250608): Added symbol and vector artwork support
- v1.2 (20250609): Added grid shape support
- v1.3 (20250610): Added offset feature and unit code support
*/

var SCRIPT_VERSION = "v1.3";

function getCurrentLang() {
  return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

/* 日英ラベル定義 / Japanese-English label definitions */

var LABELS = {
  shapePanel: { ja: "形状", en: "Shape" },
  shapeTraditional: { ja: "トラディショナル", en: "Traditional" },
  shapeRandom: { ja: "ランダム", en: "Random" },
  shapeGrid: { ja: "グリッド", en: "Grid" },
  totalPieces: { ja: "ピース数", en: "Total pieces" },
  columns: { ja: "列数", en: "Columns" },
  rows: { ja: "行数", en: "Rows" },
  explode: { ja: "バラけ処理", en: "Explode" },
  offsetLabel: { ja: "オフセット", en: "Offset" },
  dialogTitle: {
    ja: "ジグソー作成 " + SCRIPT_VERSION,
    en: "Create Jigsaw " + SCRIPT_VERSION
  },
  okBtn: { ja: "実行", en: "Run" },
  cancel: { ja: "キャンセル", en: "Cancel" }
};

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

// 画像サイズに基づく初期グリッドサイズを計算
function getInitialGridSize(imageWidth, imageHeight, totalPieces) {
  var aspectRatio = imageWidth / imageHeight;
  var cols = Math.round(Math.sqrt(totalPieces * aspectRatio));
  if (cols < 1) cols = 1;
  var rows = Math.round(totalPieces / cols);
  if (rows < 1) rows = 1;
  return [rows, cols];
}

function main() {
  /* Undoグループで全体をまとめる / Group the entire operation with undo */
  app.undoGroup = "ジグソー作成";
  try {

  /* ダイアログ作成 / Create dialog */
  var puzzlifyDialog = new Window('dialog', LABELS.dialogTitle[lang]);
  puzzlifyDialog.orientation = 'column';
  puzzlifyDialog.alignment = 'right';

  /* 入力パネル（順序調整用のグループ） / Input panel (group for layout) */
  var inputPanel = puzzlifyDialog.add('group');
  inputPanel.orientation = 'column';
  inputPanel.alignChildren = 'left';

  /* ピース設定グループ（ピース数＋行列入力欄） / Piece settings group (piece count + row/col inputs) */
  var pieceSettingGroup = inputPanel.add("group");
  pieceSettingGroup.orientation = "column";
  pieceSettingGroup.alignChildren = "left";
  pieceSettingGroup.margins = [10, 5, 10, 15];

  /* ピース数（目安）入力欄 / Total pieces input */
  var totalPiecesGroup = pieceSettingGroup.add("group");
  totalPiecesGroup.alignment = "left";
  totalPiecesGroup.add("statictext", undefined, LABELS.totalPieces[lang]);
  var totalPiecesInput = totalPiecesGroup.add("edittext", undefined, "25");
  totalPiecesInput.characters = 4;

  /* 「行数」「列数」入力欄を group にまとめて1行に並べる / Row and column input fields in a group */
  var rowColGroup = pieceSettingGroup.add('group');
  rowColGroup.orientation = 'row';
  rowColGroup.alignChildren = 'left';

  /* 列数入力欄 / Column count input */
  var colGroup = rowColGroup.add('group');
  colGroup.orientation = 'row';
  colGroup.add('statictext', undefined, LABELS.columns[lang]);
  var columnText = colGroup.add('edittext', undefined, "5");
  columnText.characters = 3;

  /* 行数入力欄 / Row count input */
  var rowGroup = rowColGroup.add('group');
  rowGroup.orientation = 'row';
  rowGroup.add('statictext', undefined, LABELS.rows[lang]);
  var rowText = rowGroup.add('edittext', undefined, "5");
  rowText.characters = 3;

  /* 選択画像やベクターアートワークがあれば初期値を自動設定 / Set initial values based on selected image/artwork */
  function getSelectedArtworkItem() {
    if (app.documents.length > 0 && app.selection.length == 1) {
      var sel = app.selection[0];
      if (
        sel.typename === "RasterItem" ||
        sel.typename === "PlacedItem" ||
        sel.typename === "SymbolItem" ||
        sel.typename === "PathItem" ||
        sel.typename === "GroupItem" ||
        sel.typename === "CompoundPathItem"
      ) {
        var gb = sel.geometricBounds;
        // [左, 上, 右, 下] / [left, top, right, bottom]
        var width = gb[2] - gb[0];
        var height = gb[1] - gb[3];
        if (height < 0) height = -height;
        return { width: width, height: height };
      }
    }
    return null;
  }
  var selectedImage = getSelectedArtworkItem();
  if (selectedImage) {
    var grid = getInitialGridSize(selectedImage.width, selectedImage.height, 25);
    rowText.text = String(grid[0]);
    columnText.text = String(grid[1]);
  }

  /* 画像・シンボル・ベクターアートワーク選択時の自動計算ロジック / Auto-calculate rows/cols for image, symbol, vector selection */
  /* 指定の条件でピース数から行・列を自動算出 / Auto-calculate rows/cols from total pieces */
  function autoCalcRowsCols() {
    if (
      app.documents.length > 0 &&
      app.selection.length === 1
    ) {
      var selItem = app.activeDocument.selection[0];
      if (
        selItem.typename === "PlacedItem" ||
        selItem.typename === "SymbolItem" ||
        selItem.typename === "PathItem" ||
        selItem.typename === "GroupItem" ||
        selItem.typename === "CompoundPathItem"
      ) {
        // ピース数が明示的に入力されたときのみ行列数を再計算 / Only recalc if total pieces entered
        if (!isNaN(parseInt(totalPiecesInput.text, 10)) && parseInt(totalPiecesInput.text, 10) > 0) {
          var pieceCount = parseInt(totalPiecesInput.text, 10);
        } else {
          return; // 行数・列数が直接指定されている場合は何もしない / Do nothing if row/col manually entered
        }
        var bounds = selItem.geometricBounds;
        var w = bounds[2] - bounds[0];
        var h = bounds[1] - bounds[3];
        if (h < 0) h = -h;
        var aspect = w / h;
        var cols = Math.round(Math.sqrt(pieceCount * aspect));
        if (cols < 1) cols = 1;
        var rows = Math.round(pieceCount / cols);
        if (rows < 1) rows = 1;
        rowText.text = rows;
        columnText.text = cols;
      }
    }
  }

  /* ピース数入力変更時、行・列数を自動更新（ベクターアートワーク含む） / Update rows/cols on total pieces change (includes vector artwork) */
  totalPiecesInput.onChanging = function () {
    var val = parseInt(totalPiecesInput.text, 10);
    if (isNaN(val) || val < 1) return;
    var selectedImage = getSelectedArtworkItem();
    if (selectedImage) {
      var grid = getInitialGridSize(selectedImage.width, selectedImage.height, val);
      rowText.text = String(grid[0]);
      columnText.text = String(grid[1]);
    }
    // 追加: ベクターアートワーク/Placed/Symbol時の自動計算 / Also auto-calc for vector/placed/symbol
    autoCalcRowsCols();
  };

  /* 行・列数入力変更時には自動計算を行わない / Don't auto-calc for manual row/col input */
  columnText.onChanging = function () {}; // 手動入力時には自動計算を行わない / No auto-calc for manual input
  rowText.onChanging = function () {}; // 手動入力時には自動計算を行わない / No auto-calc for manual input

  /* 分割方式（形状パネル内ラジオボタンを上下に並べる） / Split type (shape panel radio buttons vertical) */
  var shapePanel = inputPanel.add("panel", undefined, LABELS.shapePanel[lang]);
  shapePanel.orientation = "column";
  shapePanel.alignChildren = "left";
  shapePanel.margins = [15, 20, 15, 10];
  var shapeRadioTraditional = shapePanel.add("radiobutton", undefined, LABELS.shapeTraditional[lang]);
  var shapeRadioRandom = shapePanel.add("radiobutton", undefined, LABELS.shapeRandom[lang]);
  var shapeRadioGrid = shapePanel.add("radiobutton", undefined, LABELS.shapeGrid[lang]);
  shapeRadioTraditional.value = true;

  /* ばらけ調整グループ / Scatter adjustment group */
  var scatterGroup = inputPanel.add("group");
  scatterGroup.orientation = "row";
  scatterGroup.alignChildren = "left";
  scatterGroup.margins = [10, 5, 3, 10];
  /* オプションチェックボックス / Option checkbox */
  var scatterCheckbox = scatterGroup.add('checkbox', undefined, LABELS.explode[lang]);
  scatterCheckbox.value = false;
  /* バラけさせる強さ入力欄 / Scatter strength input */
  var scatterStrengthInput = scatterGroup.add("edittext", undefined, "30");
  scatterStrengthInput.characters = 4;
  /* 初期状態でチェックボックスがオフなら入力欄を無効化 / Disable input if checkbox off by default */
  scatterStrengthInput.enabled = scatterCheckbox.value;
  /* チェックボックスの onClick で有効/無効を切り替え / Enable/disable input on checkbox click */
  scatterCheckbox.onClick = function () {
    scatterStrengthInput.enabled = scatterCheckbox.value;
  };
  /* オフセット調整グループ（非パネル） / Offset adjustment group (non-panel) */
  var offsetGroup = inputPanel.add("group");
  offsetGroup.orientation = "row";
  offsetGroup.alignChildren = "left";
  offsetGroup.margins = [10, 0, 10, 10];
  /* オプションチェックボックス / Option checkbox */
  var offsetCheckbox = offsetGroup.add('checkbox', undefined, LABELS.offsetLabel[lang]);
  offsetCheckbox.value = true;
  /* オフセット値入力欄 / Offset value input */
  var offsetValueInput = offsetGroup.add("edittext", undefined, "-2");
  offsetValueInput.characters = 4;
  /* 単位ラベルを変数で保持 / Unit label variable */
  var offsetUnitLabel = offsetGroup.add("statictext", undefined, getCurrentUnitLabel());
  /* 初期状態でチェックボックスがオフなら入力欄と単位ラベルを無効化 / Disable input and label if checkbox off */
  offsetValueInput.enabled = offsetCheckbox.value;
  offsetUnitLabel.enabled = offsetCheckbox.value;
  /* チェックボックスの onClick で有効/無効を切り替え / Enable/disable input and label on checkbox click */
  offsetCheckbox.onClick = function () {
    offsetValueInput.enabled = offsetCheckbox.value;
    offsetUnitLabel.enabled = offsetCheckbox.value;
  };

  /* OK・キャンセルボタン / OK and Cancel buttons */
  var groupButtons = puzzlifyDialog.add('group');
  groupButtons.orientation = 'row';
  groupButtons.alignment = "right";
  groupButtons.add('button', undefined, LABELS.cancel[lang]);
  /* OKボタンの定義を show() の戻り値が "ok" になるよう name: "ok" を追加 / OK button with name: "ok" for dialog result */
  var okBtn = groupButtons.add('button', undefined, LABELS.okBtn[lang], { name: "ok" });
  okBtn.active = true;

  /* Illustrator状態とダイアログ入力チェック / Illustrator state and dialog input check */
  /* 画像選択時の変数を事前宣言 / Predeclare selected image variable */
  var selectedImage = null;
  if (
    app.documents.length > 0 &&
    app.selection.length > 0 &&
    puzzlifyDialog.show() == 1 &&
    (columnText.text || rowText.text)
  ) {
    // 選択オブジェクト取得 / Get selected object
    var selectedObj;
    var origImageObj = null;
    var selectedObjectForPuzzle;
    var isTempRect = false;

    /* 複数選択時はグループ化してシンボル化（重ね順を保持） / Group and symbolize when multiple objects selected (preserve stacking order) */
    if (app.selection.length > 1) {
      try {
        var tempGroup = app.activeDocument.groupItems.add();
        var selectedItems = [];
        for (var i = 0; i < app.selection.length; i++) {
          selectedItems.push(app.selection[i]);
        }
        // 後ろのものから先に移動することで重ね順を保持
        for (var j = selectedItems.length - 1; j >= 0; j--) {
          selectedItems[j].move(tempGroup, ElementPlacement.PLACEATBEGINNING);
        }
        var groupLeft = tempGroup.left;
        var groupTop = tempGroup.top;
        var groupParent = tempGroup.parent;
        var groupedSymbol = app.activeDocument.symbols.add(tempGroup);
        var symbolItem = groupParent.symbolItems.add(groupedSymbol);
        symbolItem.left = groupLeft;
        symbolItem.top = groupTop;
        tempGroup.remove();
        selectedObj = symbolItem;
      } catch (e) {
        alert("複数オブジェクトのシンボル化に失敗しました: " + e);
        return;
      }
    } else {
      selectedObj = app.selection[0];
    }
    selectedObjectForPuzzle = selectedObj;

    /* 埋め込み画像(RasterItem)をSymbolItemに変換 / Convert embedded image (RasterItem) to SymbolItem */
    if (selectedObj.typename === "RasterItem") {
      try {
        var rasterGB = selectedObj.geometricBounds;
        var rasterLeft = selectedObj.left;
        var rasterTop = selectedObj.top;
        var rasterParent = selectedObj.parent;
        var tempSymbol = app.activeDocument.symbols.add(selectedObj);
        var symbolItem = rasterParent.symbolItems.add(tempSymbol);
        symbolItem.left = rasterLeft;
        symbolItem.top = rasterTop;
        selectedObj.remove();
        selectedObj = symbolItem;
      } catch (e) {
        alert("埋め込み画像のシンボル化に失敗しました: " + e);
        return;
      }
    }
    /* ベクターオブジェクト(PathItem, GroupItem, CompoundPathItem)をSymbolItemに変換 / Convert vector object (PathItem, GroupItem, CompoundPathItem) to SymbolItem */
    else if (
      selectedObj.typename === "PathItem" ||
      selectedObj.typename === "GroupItem" ||
      selectedObj.typename === "CompoundPathItem"
    ) {
      try {
        var vectorLeft = selectedObj.left;
        var vectorTop = selectedObj.top;
        var vectorParent = selectedObj.parent;
        var vectorSymbol = app.activeDocument.symbols.add(selectedObj);
        var vectorSymbolItem = vectorParent.symbolItems.add(vectorSymbol);
        vectorSymbolItem.left = vectorLeft;
        vectorSymbolItem.top = vectorTop;
        selectedObj.remove();
        selectedObj = vectorSymbolItem;
      } catch (e) {
        alert("ベクターオブジェクトのシンボル化に失敗しました: " + e);
        return;
      }
    }

    /* PlacedItemまたはRasterItemまたはSymbolItemの場合は矩形化 / Convert PlacedItem, RasterItem, or SymbolItem to rectangle */
    var selType = selectedObj.typename;
    if (selType == "PlacedItem" || selType == "RasterItem" || selType == "SymbolItem") {
      var gb = selectedObj.geometricBounds;
      var rectWidth = gb[2] - gb[0];
      var rectHeight = gb[1] - gb[3];
      if (rectHeight < 0) rectHeight = -rectHeight;
      var rect = app.activeDocument.pathItems.rectangle(gb[1], gb[0], rectWidth, rectHeight);
      selectedObjectForPuzzle = rect;
      origImageObj = selectedObj; // 元画像保持（SymbolItem も含む）
      isTempRect = true;
      selectedImage = app.selection[0];
    } else {
      isTempRect = false;
      selectedImage = app.selection[0];
    }

    /* 列・行数取得 / Get column and row count */
    var columnCount = Math.round(Number(columnText.text));
    var rowCount = Math.round(Number(rowText.text));

    /* 入力値検証 / Input value validation */
    if (
      (columnCount >= 1 && rowCount >= 1) ||
      (columnCount == 0 && rowCount >= 1) ||
      (columnCount >= 1 && rowCount == 0)
    ) {
      /* 選択解除 / Deselect object */
      selectedObjectForPuzzle.selected = false;

      /* 列または行が0の場合は比率で自動算出 / Auto-calculate rows/cols from aspect ratio if either is 0 */
      var aspectRatio;
      if (columnCount == 0) {
        aspectRatio = Math.abs(
          (selectedObjectForPuzzle.geometricBounds[2] - selectedObjectForPuzzle.geometricBounds[0]) /
          (selectedObjectForPuzzle.geometricBounds[3] - selectedObjectForPuzzle.geometricBounds[1])
        );
        columnCount = Math.round(aspectRatio * rowCount);
      }
      if (rowCount == 0) {
        aspectRatio = Math.abs(
          (selectedObjectForPuzzle.geometricBounds[3] - selectedObjectForPuzzle.geometricBounds[1]) /
          (selectedObjectForPuzzle.geometricBounds[2] - selectedObjectForPuzzle.geometricBounds[0])
        );
        rowCount = Math.round(aspectRatio * columnCount);
      }

      /* オブジェクト左上座標 / Top-left coordinates of object */
      var originX = selectedObjectForPuzzle.geometricBounds[0];
      var originY = selectedObjectForPuzzle.geometricBounds[3];

      /* ピース1つあたりの幅・高さ / Width and height per piece */
      var pieceWidth = (selectedObjectForPuzzle.geometricBounds[2] - selectedObjectForPuzzle.geometricBounds[0]) / columnCount;
      var pieceHeight = (selectedObjectForPuzzle.geometricBounds[3] - selectedObjectForPuzzle.geometricBounds[1]) / rowCount;

      /* ジグソー突起用定数 / Constants for jigsaw nubs */
      var oneThirdWide = pieceWidth / 3;
      var oneQuarterHigh = pieceHeight / 4;
      var oneThirdHigh = pieceHeight / 3;
      var oneQuarterWide = pieceWidth / 4;

      /* 突起方向・ジッター配列 / Array for nub direction and jitter */
      var nubDirArray = new Array(rowCount);
      if (shapeRadioRandom.value) {
        // ランダム方式
        for (var y = 0; y < rowCount; y++) {
          nubDirArray[y] = new Array(columnCount);
          for (var x = 0; x < columnCount; x++) {
            nubDirArray[y][x] = new Array(4);
            // 上辺突起方向 true=上, false=下
            nubDirArray[y][x][0] = Math.random() < 0.5;
            // 右辺突起方向 true=右, false=左
            nubDirArray[y][x][1] = Math.random() < 0.5;
            // 上下ジッター
            nubDirArray[y][x][2] = pieceHeight * (Math.random() - 0.5) / 10;
            // 左右ジッター
            nubDirArray[y][x][3] = pieceWidth * (Math.random() - 0.5) / 10;
          }
        }
      } else {
        // トラディショナル方式
        for (var y = 0; y < rowCount; y++) {
          nubDirArray[y] = new Array(columnCount);
          for (var x = 0; x < columnCount; x++) {
            nubDirArray[y][x] = new Array(4);
            nubDirArray[y][x][0] = (x & 1) ^ (y & 1);
            nubDirArray[y][x][1] = !((x & 1) ^ (y & 1));
            nubDirArray[y][x][2] = pieceHeight * (Math.random() - 0.5) / 10;
            nubDirArray[y][x][3] = pieceWidth * (Math.random() - 0.5) / 10;
          }
        }
      }

      /* グリッド形状（矩形マスク）モード / Grid shape (rectangle mask) mode */
      if (shapeRadioGrid.value) {
        var generatedPieces = [];
        for (var y = 0; y < rowCount; y++) {
          for (var x = 0; x < columnCount; x++) {
            var gridMask = app.activeDocument.pathItems.rectangle(
              originY - y * pieceHeight,
              originX + x * pieceWidth,
              pieceWidth,
              pieceHeight
            );
            /* オフセット効果を適用（有効な場合） / Apply offset effect if enabled */
            if (offsetCheckbox.value) {
              var offsetVal = parseFloat(offsetValueInput.text);
              gridMask = offsetRectangleBounds(gridMask, offsetVal);
            }
            gridMask.closed = true;
            gridMask.filled = false;
            gridMask.stroked = false;
            if (origImageObj && (origImageObj.typename === "PlacedItem" || origImageObj.typename === "SymbolItem")) {
              var placedCopy = origImageObj.duplicate();
              var group = app.activeDocument.groupItems.add();
              placedCopy.moveToBeginning(group);
              gridMask.moveToBeginning(group);
              gridMask.clipping = true;
              group.clipped = true;
              generatedPieces.push(group);
              if (scatterCheckbox.value) {
                var strength = parseFloat(scatterStrengthInput.text);
                var offsetX = Math.random() * strength * 2 - strength;
                var offsetY = Math.random() * strength * 2 - strength;
                var moveMatrix = app.getTranslationMatrix(offsetX, offsetY);
                group.transform(moveMatrix);
              }
              group.selected = true;
            } else {
              generatedPieces.push(gridMask);
              if (scatterCheckbox.value) {
                var strength = parseFloat(scatterStrengthInput.text);
                var offsetX = Math.random() * strength * 2 - strength;
                var offsetY = Math.random() * strength * 2 - strength;
                var moveMatrix = app.getTranslationMatrix(offsetX, offsetY);
                gridMask.transform(moveMatrix);
              }
              gridMask.selected = true;
            }
          }
        }

        if (origImageObj && typeof origImageObj.remove === "function") {
          origImageObj.remove();
        }
        if (isTempRect && selectedObjectForPuzzle && typeof selectedObjectForPuzzle.remove === "function") {
          selectedObjectForPuzzle.remove();
        }
        return;
      }
      /* 各ピース生成・配置処理 / Generate and place each piece */
      var generatedPieces = [];
      for (var y = 0; y < rowCount; y++) {
        for (var x = 0; x < columnCount; x++) {
          /* マスク用パス作成 / Create mask path */
          var maskPath = app.activeDocument.pathItems.add();
          /* 左上点 / Top-left point */
          addPoint(maskPath, originX + x * pieceWidth, originY - (y + 1) * pieceHeight);

          /* 上辺突起 / Top edge nub */
          if (y < rowCount - 1) {
            if (nubDirArray[y + 1][x][0]) {
              // 上突起（上向き）
              var curvept = maskPath.pathPoints.add();
              curvept.anchor = [originX + x * pieceWidth + oneThirdWide, originY - (y + 1) * pieceHeight - nubDirArray[y + 1][x][2]];
              curvept.leftDirection = [originX + x * pieceWidth + 0.67 * oneThirdWide, originY - (y + 1) * pieceHeight + 0.33 * oneQuarterHigh - nubDirArray[y + 1][x][2]];
              curvept.rightDirection = [originX + x * pieceWidth + 1.33 * oneThirdWide, originY - (y + 1) * pieceHeight - 0.33 * oneQuarterHigh - nubDirArray[y + 1][x][2]];
              curvept.pointType = PointType.SMOOTH;

              curvept = maskPath.pathPoints.add();
              curvept.anchor = [originX + x * pieceWidth + oneThirdWide, originY - (y + 1) * pieceHeight - oneQuarterHigh - nubDirArray[y + 1][x][2]];
              curvept.rightDirection = [originX + x * pieceWidth + 1.33 * oneThirdWide, originY - (y + 1) * pieceHeight - oneQuarterHigh - 0.33 * oneQuarterHigh - nubDirArray[y + 1][x][2]];
              curvept.leftDirection = [originX + x * pieceWidth + 0.67 * oneThirdWide, originY - (y + 1) * pieceHeight - oneQuarterHigh + 0.33 * oneQuarterHigh - nubDirArray[y + 1][x][2]];
              curvept.pointType = PointType.SMOOTH;

              curvept = maskPath.pathPoints.add();
              curvept.anchor = [originX + x * pieceWidth + 2 * oneThirdWide, originY - (y + 1) * pieceHeight - oneQuarterHigh - nubDirArray[y + 1][x][2]];
              curvept.rightDirection = [originX + x * pieceWidth + 2.33 * oneThirdWide, originY - (y + 1) * pieceHeight - oneQuarterHigh + 0.33 * oneQuarterHigh - nubDirArray[y + 1][x][2]];
              curvept.leftDirection = [originX + x * pieceWidth + 1.67 * oneThirdWide, originY - (y + 1) * pieceHeight - oneQuarterHigh - 0.33 * oneQuarterHigh - nubDirArray[y + 1][x][2]];
              curvept.pointType = PointType.SMOOTH;

              curvept = maskPath.pathPoints.add();
              curvept.anchor = [originX + x * pieceWidth + 2 * oneThirdWide, originY - (y + 1) * pieceHeight - nubDirArray[y + 1][x][2]];
              curvept.rightDirection = [originX + x * pieceWidth + 2.33 * oneThirdWide, originY - (y + 1) * pieceHeight + 0.33 * oneQuarterHigh - nubDirArray[y + 1][x][2]];
              curvept.leftDirection = [originX + x * pieceWidth + 1.67 * oneThirdWide, originY - (y + 1) * pieceHeight - 0.33 * oneQuarterHigh - nubDirArray[y + 1][x][2]];
              curvept.pointType = PointType.SMOOTH;
            } else {
              // 上突起（下向き）
              var curvept = maskPath.pathPoints.add();
              curvept.anchor = [originX + x * pieceWidth + oneThirdWide, originY - (y + 1) * pieceHeight - nubDirArray[y + 1][x][2]];
              curvept.leftDirection = [originX + x * pieceWidth + 0.67 * oneThirdWide, originY - (y + 1) * pieceHeight - 0.33 * oneQuarterHigh - nubDirArray[y + 1][x][2]];
              curvept.rightDirection = [originX + x * pieceWidth + 1.33 * oneThirdWide, originY - (y + 1) * pieceHeight + 0.33 * oneQuarterHigh - nubDirArray[y + 1][x][2]];
              curvept.pointType = PointType.SMOOTH;

              curvept = maskPath.pathPoints.add();
              curvept.anchor = [originX + x * pieceWidth + oneThirdWide, originY - y * pieceHeight - 3 * oneQuarterHigh - nubDirArray[y + 1][x][2]];
              curvept.rightDirection = [originX + x * pieceWidth + 1.33 * oneThirdWide, originY - y * pieceHeight - 2.5 * oneQuarterHigh - nubDirArray[y + 1][x][2]];
              curvept.leftDirection = [originX + x * pieceWidth + 0.67 * oneThirdWide, originY - y * pieceHeight - 3.5 * oneQuarterHigh - nubDirArray[y + 1][x][2]];
              curvept.pointType = PointType.SMOOTH;

              curvept = maskPath.pathPoints.add();
              curvept.anchor = [originX + x * pieceWidth + 2 * oneThirdWide, originY - y * pieceHeight - 3 * oneQuarterHigh - nubDirArray[y + 1][x][2]];
              curvept.rightDirection = [originX + x * pieceWidth + 2.33 * oneThirdWide, originY - y * pieceHeight - 3.5 * oneQuarterHigh - nubDirArray[y + 1][x][2]];
              curvept.leftDirection = [originX + x * pieceWidth + 1.67 * oneThirdWide, originY - y * pieceHeight - 2.5 * oneQuarterHigh - nubDirArray[y + 1][x][2]];
              curvept.pointType = PointType.SMOOTH;

              curvept = maskPath.pathPoints.add();
              curvept.anchor = [originX + x * pieceWidth + 2 * oneThirdWide, originY - (y + 1) * pieceHeight - nubDirArray[y + 1][x][2]];
              curvept.rightDirection = [originX + x * pieceWidth + 2.33 * oneThirdWide, originY - (y + 1) * pieceHeight - 0.33 * oneQuarterHigh - nubDirArray[y + 1][x][2]];
              curvept.leftDirection = [originX + x * pieceWidth + 1.67 * oneThirdWide, originY - (y + 1) * pieceHeight + 0.33 * oneQuarterHigh - nubDirArray[y + 1][x][2]];
              curvept.pointType = PointType.SMOOTH;
            }
          }

          /* 右上点 / Top-right point */
          addPoint(maskPath, originX + (x + 1) * pieceWidth, originY - (y + 1) * pieceHeight);

          /* 右辺突起 / Right edge nub */
          if (x < columnCount - 1) {
            if (nubDirArray[y][x + 1][1]) {
              // 右突起（右向き）
              var curvept = maskPath.pathPoints.add();
              curvept.anchor = [originX + x * pieceWidth + 4 * oneQuarterWide - nubDirArray[y][x + 1][3], originY - y * pieceHeight - 2 * oneThirdHigh];
              curvept.leftDirection = [originX + x * pieceWidth + 3.5 * oneQuarterWide - nubDirArray[y][x + 1][3], originY - y * pieceHeight - 2.33 * oneThirdHigh];
              curvept.rightDirection = [originX + x * pieceWidth + 4.5 * oneQuarterWide - nubDirArray[y][x + 1][3], originY - y * pieceHeight - 1.67 * oneThirdHigh];
              curvept.pointType = PointType.SMOOTH;

              curvept = maskPath.pathPoints.add();
              curvept.anchor = [originX + x * pieceWidth + 5 * oneQuarterWide - nubDirArray[y][x + 1][3], originY - y * pieceHeight - 2 * oneThirdHigh];
              curvept.leftDirection = [originX + x * pieceWidth + 4.5 * oneQuarterWide - nubDirArray[y][x + 1][3], originY - y * pieceHeight - 2.33 * oneThirdHigh];
              curvept.rightDirection = [originX + x * pieceWidth + 5.5 * oneQuarterWide - nubDirArray[y][x + 1][3], originY - y * pieceHeight - 1.67 * oneThirdHigh];
              curvept.pointType = PointType.SMOOTH;

              curvept = maskPath.pathPoints.add();
              curvept.anchor = [originX + x * pieceWidth + 5 * oneQuarterWide - nubDirArray[y][x + 1][3], originY - y * pieceHeight - 1 * oneThirdHigh];
              curvept.leftDirection = [originX + x * pieceWidth + 5.5 * oneQuarterWide - nubDirArray[y][x + 1][3], originY - y * pieceHeight - 1.33 * oneThirdHigh];
              curvept.rightDirection = [originX + x * pieceWidth + 4.5 * oneQuarterWide - nubDirArray[y][x + 1][3], originY - y * pieceHeight - 0.67 * oneThirdHigh];
              curvept.pointType = PointType.SMOOTH;

              curvept = maskPath.pathPoints.add();
              curvept.anchor = [originX + x * pieceWidth + 4 * oneQuarterWide - nubDirArray[y][x + 1][3], originY - y * pieceHeight - oneThirdHigh];
              curvept.leftDirection = [originX + x * pieceWidth + 4.5 * oneQuarterWide - nubDirArray[y][x + 1][3], originY - y * pieceHeight - 1.33 * oneThirdHigh];
              curvept.rightDirection = [originX + x * pieceWidth + 3.5 * oneQuarterWide - nubDirArray[y][x + 1][3], originY - y * pieceHeight - 0.67 * oneThirdHigh];
              curvept.pointType = PointType.SMOOTH;
            } else {
              // 右突起（左向き）
              var curvept = maskPath.pathPoints.add();
              curvept.anchor = [originX + x * pieceWidth + 4 * oneQuarterWide - nubDirArray[y][x + 1][3], originY - y * pieceHeight - 2 * oneThirdHigh];
              curvept.rightDirection = [originX + x * pieceWidth + 3.5 * oneQuarterWide - nubDirArray[y][x + 1][3], originY - y * pieceHeight - 1.67 * oneThirdHigh];
              curvept.leftDirection = [originX + x * pieceWidth + 4.5 * oneQuarterWide - nubDirArray[y][x + 1][3], originY - y * pieceHeight - 2.33 * oneThirdHigh];
              curvept.pointType = PointType.SMOOTH;

              curvept = maskPath.pathPoints.add();
              curvept.anchor = [originX + x * pieceWidth + 3 * oneQuarterWide - nubDirArray[y][x + 1][3], originY - y * pieceHeight - 2 * oneThirdHigh];
              curvept.rightDirection = [originX + x * pieceWidth + 2.5 * oneQuarterWide - nubDirArray[y][x + 1][3], originY - y * pieceHeight - 1.67 * oneThirdHigh];
              curvept.leftDirection = [originX + x * pieceWidth + 3.5 * oneQuarterWide - nubDirArray[y][x + 1][3], originY - y * pieceHeight - 2.33 * oneThirdHigh];
              curvept.pointType = PointType.SMOOTH;

              curvept = maskPath.pathPoints.add();
              curvept.anchor = [originX + x * pieceWidth + 3 * oneQuarterWide - nubDirArray[y][x + 1][3], originY - y * pieceHeight - 1 * oneThirdHigh];
              curvept.leftDirection = [originX + x * pieceWidth + 2.5 * oneQuarterWide - nubDirArray[y][x + 1][3], originY - y * pieceHeight - 1.33 * oneThirdHigh];
              curvept.rightDirection = [originX + x * pieceWidth + 3.5 * oneQuarterWide - nubDirArray[y][x + 1][3], originY - y * pieceHeight - 0.67 * oneThirdHigh];
              curvept.pointType = PointType.SMOOTH;

              curvept = maskPath.pathPoints.add();
              curvept.anchor = [originX + x * pieceWidth + 4 * oneQuarterWide - nubDirArray[y][x + 1][3], originY - y * pieceHeight - oneThirdHigh];
              curvept.rightDirection = [originX + x * pieceWidth + 4.5 * oneQuarterWide - nubDirArray[y][x + 1][3], originY - y * pieceHeight - 0.67 * oneThirdHigh];
              curvept.leftDirection = [originX + x * pieceWidth + 3.5 * oneQuarterWide - nubDirArray[y][x + 1][3], originY - y * pieceHeight - 1.33 * oneThirdHigh];
              curvept.pointType = PointType.SMOOTH;
            }
          }

          /* 右下点 / Bottom-right point */
          addPoint(maskPath, originX + (x + 1) * pieceWidth, originY - y * pieceHeight);

          /* 下辺突起 / Bottom edge nub */
          if (y > 0) {
            if (nubDirArray[y][x][0]) {
              // 下突起（上向き）
              var curvept = maskPath.pathPoints.add();
              curvept.anchor = [originX + x * pieceWidth + 2 * oneThirdWide, originY - y * pieceHeight - nubDirArray[y][x][2]];
              curvept.leftDirection = [originX + x * pieceWidth + 2.33 * oneThirdWide, originY - y * pieceHeight + 0.33 * oneQuarterHigh - nubDirArray[y][x][2]];
              curvept.rightDirection = [originX + x * pieceWidth + 1.67 * oneThirdWide, originY - y * pieceHeight - 0.33 * oneQuarterHigh - nubDirArray[y][x][2]];
              curvept.pointType = PointType.SMOOTH;

              curvept = maskPath.pathPoints.add();
              curvept.anchor = [originX + x * pieceWidth + 2 * oneThirdWide, originY - y * pieceHeight - oneQuarterHigh - nubDirArray[y][x][2]];
              curvept.leftDirection = [originX + x * pieceWidth + 2.33 * oneThirdWide, originY - y * pieceHeight - oneQuarterHigh + 0.33 * oneQuarterHigh - nubDirArray[y][x][2]];
              curvept.rightDirection = [originX + x * pieceWidth + 1.67 * oneThirdWide, originY - y * pieceHeight - oneQuarterHigh - 0.33 * oneQuarterHigh - nubDirArray[y][x][2]];
              curvept.pointType = PointType.SMOOTH;

              curvept = maskPath.pathPoints.add();
              curvept.anchor = [originX + x * pieceWidth + oneThirdWide, originY - y * pieceHeight - oneQuarterHigh - nubDirArray[y][x][2]];
              curvept.leftDirection = [originX + x * pieceWidth + 1.33 * oneThirdWide, originY - y * pieceHeight - oneQuarterHigh - 0.33 * oneQuarterHigh - nubDirArray[y][x][2]];
              curvept.rightDirection = [originX + x * pieceWidth + 0.67 * oneThirdWide, originY - y * pieceHeight - oneQuarterHigh + 0.33 * oneQuarterHigh - nubDirArray[y][x][2]];
              curvept.pointType = PointType.SMOOTH;

              curvept = maskPath.pathPoints.add();
              curvept.anchor = [originX + x * pieceWidth + oneThirdWide, originY - y * pieceHeight - nubDirArray[y][x][2]];
              curvept.rightDirection = [originX + x * pieceWidth + 0.67 * oneThirdWide, originY - y * pieceHeight + 0.33 * oneQuarterHigh - nubDirArray[y][x][2]];
              curvept.leftDirection = [originX + x * pieceWidth + 1.33 * oneThirdWide, originY - y * pieceHeight - 0.33 * oneQuarterHigh - nubDirArray[y][x][2]];
              curvept.pointType = PointType.SMOOTH;
            } else {
              // 下突起（下向き）
              var curvept = maskPath.pathPoints.add();
              curvept.anchor = [originX + x * pieceWidth + 2 * oneThirdWide, originY - y * pieceHeight - nubDirArray[y][x][2]];
              curvept.leftDirection = [originX + x * pieceWidth + 2.33 * oneThirdWide, originY - y * pieceHeight - 0.33 * oneQuarterHigh - nubDirArray[y][x][2]];
              curvept.rightDirection = [originX + x * pieceWidth + 1.67 * oneThirdWide, originY - y * pieceHeight + 0.33 * oneQuarterHigh - nubDirArray[y][x][2]];
              curvept.pointType = PointType.SMOOTH;

              curvept = maskPath.pathPoints.add();
              curvept.anchor = [originX + x * pieceWidth + 2 * oneThirdWide, originY - y * pieceHeight + oneQuarterHigh - nubDirArray[y][x][2]];
              curvept.rightDirection = [originX + x * pieceWidth + 1.67 * oneThirdWide, originY - y * pieceHeight + 1.5 * oneQuarterHigh - nubDirArray[y][x][2]];
              curvept.leftDirection = [originX + x * pieceWidth + 2.33 * oneThirdWide, originY - y * pieceHeight + 0.5 * oneQuarterHigh - nubDirArray[y][x][2]];
              curvept.pointType = PointType.SMOOTH;

              curvept = maskPath.pathPoints.add();
              curvept.anchor = [originX + x * pieceWidth + oneThirdWide, originY - y * pieceHeight + oneQuarterHigh - nubDirArray[y][x][2]];
              curvept.rightDirection = [originX + x * pieceWidth + 0.67 * oneThirdWide, originY - y * pieceHeight + 0.5 * oneQuarterHigh - nubDirArray[y][x][2]];
              curvept.leftDirection = [originX + x * pieceWidth + 1.33 * oneThirdWide, originY - y * pieceHeight + 1.5 * oneQuarterHigh - nubDirArray[y][x][2]];
              curvept.pointType = PointType.SMOOTH;

              curvept = maskPath.pathPoints.add();
              curvept.anchor = [originX + x * pieceWidth + oneThirdWide, originY - y * pieceHeight - nubDirArray[y][x][2]];
              curvept.rightDirection = [originX + x * pieceWidth + 0.67 * oneThirdWide, originY - y * pieceHeight - 0.33 * oneQuarterHigh - nubDirArray[y][x][2]];
              curvept.leftDirection = [originX + x * pieceWidth + 1.33 * oneThirdWide, originY - y * pieceHeight + 0.33 * oneQuarterHigh - nubDirArray[y][x][2]];
              curvept.pointType = PointType.SMOOTH;
            }
          }

          /* 左下点 / Bottom-left point */
          addPoint(maskPath, originX + x * pieceWidth, originY - y * pieceHeight);

          /* 左辺突起 / Left edge nub */
          if (x > 0) {
            if (nubDirArray[y][x][1]) {
              // 左突起（右向き）
              var curvept = maskPath.pathPoints.add();
              curvept.anchor = [originX + x * pieceWidth - nubDirArray[y][x][3], originY - y * pieceHeight - oneThirdHigh];
              curvept.rightDirection = [originX + x * pieceWidth + 0.5 * oneQuarterWide - nubDirArray[y][x][3], originY - y * pieceHeight - 1.33 * oneThirdHigh];
              curvept.leftDirection = [originX + x * pieceWidth - 0.5 * oneQuarterWide - nubDirArray[y][x][3], originY - y * pieceHeight - 0.67 * oneThirdHigh];
              curvept.pointType = PointType.SMOOTH;

              curvept = maskPath.pathPoints.add();
              curvept.anchor = [originX + x * pieceWidth + oneQuarterWide - nubDirArray[y][x][3], originY - y * pieceHeight - 1 * oneThirdHigh];
              curvept.rightDirection = [originX + x * pieceWidth + 1.5 * oneQuarterWide - nubDirArray[y][x][3], originY - y * pieceHeight - 1.33 * oneThirdHigh];
              curvept.leftDirection = [originX + x * pieceWidth + 0.5 * oneQuarterWide - nubDirArray[y][x][3], originY - y * pieceHeight - 0.67 * oneThirdHigh];
              curvept.pointType = PointType.SMOOTH;

              curvept = maskPath.pathPoints.add();
              curvept.anchor = [originX + x * pieceWidth + oneQuarterWide - nubDirArray[y][x][3], originY - y * pieceHeight - 2 * oneThirdHigh];
              curvept.rightDirection = [originX + x * pieceWidth + 0.5 * oneQuarterWide - nubDirArray[y][x][3], originY - y * pieceHeight - 2.33 * oneThirdHigh];
              curvept.leftDirection = [originX + x * pieceWidth + 1.5 * oneQuarterWide - nubDirArray[y][x][3], originY - y * pieceHeight - 1.67 * oneThirdHigh];
              curvept.pointType = PointType.SMOOTH;

              curvept = maskPath.pathPoints.add();
              curvept.anchor = [originX + x * pieceWidth - nubDirArray[y][x][3], originY - y * pieceHeight - 2 * oneThirdHigh];
              curvept.rightDirection = [originX + x * pieceWidth - 0.5 * oneQuarterWide - nubDirArray[y][x][3], originY - y * pieceHeight - 2.33 * oneThirdHigh];
              curvept.leftDirection = [originX + x * pieceWidth + 0.5 * oneQuarterWide - nubDirArray[y][x][3], originY - y * pieceHeight - 1.67 * oneThirdHigh];
              curvept.pointType = PointType.SMOOTH;
            } else {
              // 左突起（左向き）
              var curvept = maskPath.pathPoints.add();
              curvept.anchor = [originX + x * pieceWidth - nubDirArray[y][x][3], originY - y * pieceHeight - oneThirdHigh];
              curvept.leftDirection = [originX + x * pieceWidth + 0.5 * oneQuarterWide - nubDirArray[y][x][3], originY - y * pieceHeight - 0.67 * oneThirdHigh];
              curvept.rightDirection = [originX + x * pieceWidth - 0.5 * oneQuarterWide - nubDirArray[y][x][3], originY - y * pieceHeight - 1.33 * oneThirdHigh];
              curvept.pointType = PointType.SMOOTH;

              curvept = maskPath.pathPoints.add();
              curvept.anchor = [originX + x * pieceWidth - oneQuarterWide - nubDirArray[y][x][3], originY - y * pieceHeight - 1 * oneThirdHigh];
              curvept.leftDirection = [originX + x * pieceWidth - 0.5 * oneQuarterWide - nubDirArray[y][x][3], originY - y * pieceHeight - 0.67 * oneThirdHigh];
              curvept.rightDirection = [originX + x * pieceWidth - 1.5 * oneQuarterWide - nubDirArray[y][x][3], originY - y * pieceHeight - 1.33 * oneThirdHigh];
              curvept.pointType = PointType.SMOOTH;

              curvept = maskPath.pathPoints.add();
              curvept.anchor = [originX + x * pieceWidth - oneQuarterWide - nubDirArray[y][x][3], originY - y * pieceHeight - 2 * oneThirdHigh];
              curvept.rightDirection = [originX + x * pieceWidth - 0.5 * oneQuarterWide - nubDirArray[y][x][3], originY - y * pieceHeight - 2.33 * oneThirdHigh];
              curvept.leftDirection = [originX + x * pieceWidth - 1.5 * oneQuarterWide - nubDirArray[y][x][3], originY - y * pieceHeight - 1.67 * oneThirdHigh];
              curvept.pointType = PointType.SMOOTH;

              curvept = maskPath.pathPoints.add();
              curvept.anchor = [originX + x * pieceWidth - nubDirArray[y][x][3], originY - y * pieceHeight - 2 * oneThirdHigh];
              curvept.leftDirection = [originX + x * pieceWidth - 0.5 * oneQuarterWide - nubDirArray[y][x][3], originY - y * pieceHeight - 1.67 * oneThirdHigh];
              curvept.rightDirection = [originX + x * pieceWidth + 0.5 * oneQuarterWide - nubDirArray[y][x][3], originY - y * pieceHeight - 2.33 * oneThirdHigh];
              curvept.pointType = PointType.SMOOTH;
            }
          }

          /* パスを閉じる / Close path */
          maskPath.closed = true;

          /* 元の長方形（baseRectやpieceRectなど）は remove する（この場合、isTempRect/selectedObjectForPuzzle のみが該当）
             → ピースごとには不要（ピース生成用一時矩形は main の最後で削除）
             もし baseRect や pieceBase を作成していた場合は remove() を呼び出す（このスクリプトではピース生成用一時矩形は未使用）
             / Remove base rectangle if created (not used here) */

          /* オフセット効果を適用（グループ化/マスク前） / Apply offset effect if enabled (before grouping/masking) */
          if (offsetCheckbox.value && maskPath) {
            var offsetVal = parseFloat(offsetValueInput.text);
            app.selection = null;
            try {
              maskPath.selected = true;
              var result = applyOffsetPathToSelection(offsetVal);
              if (result) {
                if (result.typename === "PathItem") {
                  maskPath = result;
                } else if (result.typename === "GroupItem") {
                  for (var i = 0; i < result.pageItems.length; i++) {
                    if (result.pageItems[i].typename === "PathItem") {
                      maskPath = result.pageItems[i];
                      break;
                    }
                  }
                  if (maskPath.typename !== "PathItem") {
                    alert("オフセット後の GroupItem に PathItem が含まれていません。マスク処理をスキップします。");
                  }
                } else {
                  alert("オフセット後のオブジェクトが予期しない型です。マスク処理をスキップします。");
                }
              }
            } catch (e) {
              alert("オフセット適用中にエラーが発生しました：" + e.message);
            }
          }
          /* 画像クリッピンググループ処理 / Image clipping group process */
          /* 複製画像の位置は元のまま、マスクパスを画像と同じ位置に重ねてマスク / Overlay mask path at original image position */
          if (origImageObj && (origImageObj.typename === "PlacedItem" || origImageObj.typename === "SymbolItem")) {
            var placedCopy = origImageObj.duplicate();
            // グループ作成、画像とパスを追加（位置は変更しない）
            var group = app.activeDocument.groupItems.add();
            placedCopy.moveToBeginning(group);
            // マスクパスが有効なPathItemかどうか確認
            if (maskPath && maskPath.typename === "PathItem") {
              maskPath.moveToBeginning(group);
              maskPath.clipping = true;
              group.clipped = true; // グループのclippedは画像の位置に影響しない
            } else {
              alert("マスク用オブジェクトが PathItem ではないため、マスクをスキップします。");
            }
            generatedPieces.push(group);
            // バラけさせる場合、グループごと移動
            if (scatterCheckbox.value) {
              var strength = parseFloat(scatterStrengthInput.text);
              var offsetX = Math.random() * strength * 2 - strength;
              var offsetY = Math.random() * strength * 2 - strength;
              var moveMatrix = app.getTranslationMatrix(offsetX, offsetY);
              group.transform(moveMatrix);
            }
            group.selected = true;
          } else {
            /* 通常パスのみ / Just the mask path */
            generatedPieces.push(maskPath);
            if (scatterCheckbox.value) {
              var strength = parseFloat(scatterStrengthInput.text);
              var offsetX = Math.random() * strength * 2 - strength;
              var offsetY = Math.random() * strength * 2 - strength;
              var moveMatrix = app.getTranslationMatrix(offsetX, offsetY);
              maskPath.transform(moveMatrix);
            }
            maskPath.selected = true;
          }
        }
      }

      /* 元画像削除（PlacedItem/RasterItemの場合） / Delete original image (PlacedItem/RasterItem) */
      if (origImageObj && typeof origImageObj.remove === "function") {
        origImageObj.remove();
      }
      /* 一時的に作成した矩形があれば削除 / Remove temp rectangle if created */
      if (isTempRect && selectedObjectForPuzzle && typeof selectedObjectForPuzzle.remove === "function") {
        selectedObjectForPuzzle.remove();
      }
    }
  } else {
    puzzlifyDialog.hide();
  }

  /* パス上に点を追加する関数 / Function to add a point to the path */
  function addPoint(obj, x, y) {
    var np = obj.pathPoints.add();
    np.anchor = [x, y];
    np.leftDirection = [x, y];
    np.rightDirection = [x, y];
    np.pointType = PointType.CORNER;
  }

  /* --- end main --- */
  } catch (e) {
    alert("スクリプト実行中にエラーが発生しました: " + e);
  }
}

/* オフセットパスエフェクトユーティリティ / Offset Path Effect Utility */
function createOffsetEffectXML(offsetVal) {
  var xml = '<LiveEffect name="Adobe Offset Path"><Dict data="R mlim 4 R ofst value I jntp 2 "/></LiveEffect>';
  return xml.replace("value", offsetVal);
}

function applyOffsetPathToSelection(offsetVal) {
  try {
    var doc = app.activeDocument;
    var original = doc.selection[0];

    app.userInteractionLevel = UserInteractionLevel.DONTDISPLAYALERTS;
    doc.selection = null;

    var copy = original.duplicate(original, ElementPlacement.PLACEAFTER);
    original.remove();

    copy.selected = true;
    copy.applyEffect(createOffsetEffectXML(offsetVal));
    app.redraw();
    app.executeMenuCommand('expandStyle');

    app.userInteractionLevel = UserInteractionLevel.DISPLAYALERTS;

    // Return the updated selection (should be the new item)
    if (app.selection && app.selection.length > 0) {
      return app.selection[0];
    } else {
      return null;
    }
  } catch (err) {
    alert("エラーが発生しました：\n" + err.message);
    return null;
  }
}

main();

/* 矩形の座標をオフセットする関数 / Function to offset rectangle bounds */
function offsetRectangleBounds(pathItem, offsetVal) {
  var gb = pathItem.geometricBounds;
  var newWidth = (gb[2] - gb[0]) + offsetVal * 2;
  var newHeight = (gb[1] - gb[3]) + offsetVal * 2;
  pathItem.left = pathItem.left - offsetVal;
  pathItem.top = pathItem.top + offsetVal;
  pathItem.width = newWidth;
  pathItem.height = newHeight;
  return pathItem;
}