#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false); 

/*

Slice2Artboards2.jsx


### 更新履歴

- v1.0 (20250607) : 初期バージョン


【日本語説明】
選択した画像やベクターオブジェクトを指定した行数・列数で分割し、それぞれを個別の矩形マスクでクリッピングします。分割ピースはバラけさせたり、オフセットをつけてサイズを調整することも可能です。複数オブジェクト選択時は自動的にグループ化・シンボル化して扱います。主にIllustratorでのアートボード分割やパズル風レイアウト作成に利用できます。

*/

function getCurrentLang() {
  // 現在のロケールから日本語か英語かを判定 / Determine language by locale
  return ($.locale.indexOf('ja') === 0) ? 'ja' : 'en';
}

/* 単位コードとラベルのマップ / Map from unit code to label */
var unitLabelMap = {
  0: "in",  1: "mm",  2: "pt",  3: "pica",  4: "cm",
  5: "Q/H",  6: "px",  7: "ft/in",  8: "m",  9: "yd",  10: "ft"
};

function getCurrentUnitLabel() {
  // 現在の単位ラベルを取得 / Get current document unit label
  var unitCode = app.preferences.getIntegerPreference("rulerType");
  return unitLabelMap[unitCode] || "pt";
}

/* UIで使用するラベル定義 / Labels for UI */
var LABELS = {
  columns:     { ja: "列数", en: "Columns" },
  rows:        { ja: "行数", en: "Rows" },
  offsetLabel: { ja: "オフセット", en: "Offset" },
  okBtn:       { ja: "実行", en: "Run" },
  cancel:      { ja: "キャンセル", en: "Cancel" }
};

/* 画像サイズに基づく初期グリッドサイズ計算 / Calculate initial grid size based on image size */
function getInitialGridSize(imageWidth, imageHeight) {
  // デフォルトの分割数 (5x5)
  var defaultRows = 5;
  var defaultCols = 5;
  var aspectRatio = imageWidth / imageHeight;
  var cols = defaultCols;
  var rows = defaultRows;
  if (aspectRatio >= 1) {
    cols = Math.max(1, Math.round(defaultCols * aspectRatio));
    rows = defaultRows;
  } else {
    rows = Math.max(1, Math.round(defaultRows / aspectRatio));
    cols = defaultCols;
  }
  return [rows, cols];
}

function main() {
  app.undoGroup = "ジグソー作成";
  try {
    var lang = getCurrentLang();

    /* --- ダイアログ作成 / Create dialog --- */
    var Dialog = new Window('dialog', 'Puzzlify');
    Dialog.orientation = 'column';
    Dialog.alignment = 'right';

    // 入力パネル / Input panel
    var inputPanel = Dialog.add('group');
    inputPanel.orientation = 'column';
    inputPanel.alignChildren = 'left';

    /* --- ピース設定グループ / Piece settings (total, rows, columns) --- */
    var pieceSettingGroup = inputPanel.add("group");
    pieceSettingGroup.orientation = "column";
    pieceSettingGroup.alignChildren = "left";
    pieceSettingGroup.margins = [10, 5, 10, 15];


    // 行数・列数入力欄 / Row and column inputs
    var rowColGroup = pieceSettingGroup.add('group');
    rowColGroup.orientation = 'row';
    rowColGroup.alignChildren = 'left';

    // 列数 / Columns
    var colGroup = rowColGroup.add('group');
    colGroup.orientation = 'row';
    colGroup.add('statictext', undefined, LABELS.columns[lang]);
    var columnText = colGroup.add('edittext', undefined, "5");
    columnText.characters = 3;

    // 行数 / Rows
    var rowGroup = rowColGroup.add('group');
    rowGroup.orientation = 'row';
    rowGroup.add('statictext', undefined, LABELS.rows[lang]);
    var rowText = rowGroup.add('edittext', undefined, "5");
    rowText.characters = 3;

    /* 選択画像やベクターアートワークがあれば初期値を自動設定 / Auto-set grid if artwork selected */
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
      var grid = getInitialGridSize(selectedImage.width, selectedImage.height);
      rowText.text = String(grid[0]);
      columnText.text = String(grid[1]);
    }

    // 行・列数入力時は自動計算しない / Do not auto-calc on row/col edit
    columnText.onChanging = function () {};
    rowText.onChanging = function () {};

    // 形状パネルは削除（グリッド固定）

    // --- バラけ調整グループ（削除済み） ---

    /* --- オフセットグループ / Offset group --- */
    var offsetGroup = inputPanel.add("group");
    offsetGroup.orientation = "row";
    offsetGroup.alignChildren = "left";
    offsetGroup.margins = [10, 0, 10, 10];
    var offsetCheckbox = offsetGroup.add('checkbox', undefined, LABELS.offsetLabel[lang]);
    offsetCheckbox.value = true;
    var offsetValueInput = offsetGroup.add("edittext", undefined, "-2");
    offsetValueInput.characters = 4;
    var offsetUnitLabel = offsetGroup.add("statictext", undefined, getCurrentUnitLabel());
    offsetValueInput.enabled = offsetCheckbox.value;
    offsetUnitLabel.enabled = offsetCheckbox.value;
    offsetCheckbox.onClick = function () {
      offsetValueInput.enabled = offsetCheckbox.value;
      offsetUnitLabel.enabled = offsetCheckbox.value;
    };

    /* --- OK・キャンセルボタン / OK & Cancel buttons --- */
    var groupButtons = Dialog.add('group');
    groupButtons.orientation = 'row';
    groupButtons.alignment = "right";
    groupButtons.add('button', undefined, LABELS.cancel[lang]);
    var okBtn = groupButtons.add('button', undefined, LABELS.okBtn[lang], { name: "ok" });
    okBtn.active = true;

    // 画像選択時の変数を事前宣言
    var selectedImage = null;
    if (
      app.documents.length > 0 &&
      app.selection.length > 0 &&
      Dialog.show() == 1 &&
      (columnText.text || rowText.text)
    ) {
      // --- オブジェクト準備 / Prepare selected object ---
      var selectedObj;
      var origImageObj = null;
      var selectedObjectForPuzzle;
      var isTempRect = false;

      /* 複数選択時はグループ化してシンボル化 / Group and symbolize if multiple selected */
      if (app.selection.length > 1) {
        try {
          var tempGroup = app.activeDocument.groupItems.add();
          var selectedItems = [];
          for (var i = 0; i < app.selection.length; i++) {
            selectedItems.push(app.selection[i]);
          }
          // 重ね順を保持 / Preserve stacking order
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

      /* 埋め込み画像(RasterItem)やベクターをシンボル化 / Symbolize raster or vector */
      if (selectedObj.typename === "RasterItem") {
        try {
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

      /* シンボル/画像の場合は矩形化 / For symbol/image, create rectangle */
      var selType = selectedObj.typename;
      if (selType == "PlacedItem" || selType == "RasterItem" || selType == "SymbolItem") {
        var gb = selectedObj.geometricBounds;
        var rectWidth = gb[2] - gb[0];
        var rectHeight = gb[1] - gb[3];
        if (rectHeight < 0) rectHeight = -rectHeight;
        var rect = app.activeDocument.pathItems.rectangle(gb[1], gb[0], rectWidth, rectHeight);
        selectedObjectForPuzzle = rect;
        origImageObj = selectedObj;
        isTempRect = true;
        selectedImage = app.selection[0];
      } else {
        isTempRect = false;
        selectedImage = app.selection[0];
      }

      // 列・行数取得 / Get row and column count
      var columnCount = Math.round(Number(columnText.text));
      var rowCount = Math.round(Number(rowText.text));

      // 入力値検証 / Validate input
      if (
        (columnCount >= 1 && rowCount >= 1) ||
        (columnCount == 0 && rowCount >= 1) ||
        (columnCount >= 1 && rowCount == 0)
      ) {
        selectedObjectForPuzzle.selected = false;

        // 列または行が0の場合は比率で自動算出 / Auto-calc if row/col is zero
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

        // オブジェクト左上座標 / Top-left coordinate
        var originX = selectedObjectForPuzzle.geometricBounds[0];
        var originY = selectedObjectForPuzzle.geometricBounds[3];
        // ピース1つあたりの幅・高さ / Piece width & height
        var pieceWidth = (selectedObjectForPuzzle.geometricBounds[2] - selectedObjectForPuzzle.geometricBounds[0]) / columnCount;
        var pieceHeight = (selectedObjectForPuzzle.geometricBounds[3] - selectedObjectForPuzzle.geometricBounds[1]) / rowCount;

        /* --- グリッド形状（矩形マスク）モード / Grid mask mode --- */
        if (true) {
          var generatedPieces = [];
          for (var y = 0; y < rowCount; y++) {
            for (var x = 0; x < columnCount; x++) {
              var gridMask = app.activeDocument.pathItems.rectangle(
                originY - y * pieceHeight,
                originX + x * pieceWidth,
                pieceWidth,
                pieceHeight
              );
              // オフセット適用 / Apply offset if enabled
              if (offsetCheckbox.value) {
                var offsetVal = parseFloat(offsetValueInput.text);
                gridMask = offsetRectangleBounds(gridMask, offsetVal);
              }
              gridMask.closed = true;
              gridMask.filled = false;
              gridMask.stroked = false;
              // 画像やシンボルならクリッピング / Clip if image or symbol
              if (origImageObj && (origImageObj.typename === "PlacedItem" || origImageObj.typename === "SymbolItem")) {
                var placedCopy = origImageObj.duplicate();
                var group = app.activeDocument.groupItems.add();
                placedCopy.moveToBeginning(group);
                gridMask.moveToBeginning(group);
                gridMask.clipping = true;
                group.clipped = true;
                generatedPieces.push(group);
                group.selected = true;
              } else {
                generatedPieces.push(gridMask);
                gridMask.selected = true;
              }
            }
          }
          // 元画像・一時矩形削除 / Remove original image and temp rect
          if (origImageObj && typeof origImageObj.remove === "function") {
            origImageObj.remove();
          }
          if (isTempRect && selectedObjectForPuzzle && typeof selectedObjectForPuzzle.remove === "function") {
            selectedObjectForPuzzle.remove();
          }
          return;
        }
        // ジグソー形状の生成ロジックは削除されました / Jigsaw logic removed
        if (origImageObj && typeof origImageObj.remove === "function") {
          origImageObj.remove();
        }
        if (isTempRect && selectedObjectForPuzzle && typeof selectedObjectForPuzzle.remove === "function") {
          selectedObjectForPuzzle.remove();
        }
      }
    } else {
      Dialog.hide();
    }
  } catch (e) {
    alert("スクリプト実行中にエラーが発生しました: " + e);
  }
}

/* --- Offset Path Effect Utility --- */
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

/* --- 矩形の座標をオフセット調整 / Offset rectangle bounds --- */
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

/*
【最適化提案 / Optimization Suggestions】
- シンボル化処理やベクター・ラスターのtry-catchを関数化し共通化することで冗長な重複を削減可能。
- offsetRectangleBoundsの座標調整をより汎用的なユーティリティ関数として抽出できる。
- getInitialGridSize, autoCalcRowsColsなど補助関数を共通モジュール化することで再利用性向上が見込める。
*/