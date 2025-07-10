#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
Slice2Artboards2.jsx

【日本語説明】
選択した画像やベクターオブジェクトを、指定した行数・列数のグリッドに分割し、それぞれを個別の矩形マスクでクリッピングします。分割ピースはグリッド上に整列し、オフセット値で各ピースのサイズを調整できます。複数オブジェクト選択時は自動的にグループ化・シンボル化して扱います。主にIllustratorでのアートボード分割やパズル風レイアウト作成、印刷物の面付けなどに利用できます。

【English Description】
Splits selected image or vector objects into a grid of specified rows and columns, masking each piece with a separate rectangle. Each piece is arranged in a grid, and offset values allow you to adjust the size of each piece. When multiple objects are selected, they are automatically grouped and symbolized. Useful for dividing artboards, creating puzzle-like layouts, or imposition for printing in Adobe Illustrator.

### 更新履歴 / Change Log
- v1.0 (20250607): 初期バージョン / Initial version
- v1.1 (20250710): コメント粒度統一・UIラベル整理・説明文拡充・未使用項目削除・構造整理 / Unified comment granularity, cleaned up UI labels, expanded documentation, removed unused items, improved code structure
*/

function getCurrentLang() {
  /* 現在のロケールから日本語か英語かを判定 / Determine language by locale */
  return ($.locale.indexOf('ja') === 0) ? 'ja' : 'en';
}

/* 単位コードとラベルのマップ / Map from unit code to label */
var unitLabelMap = {
  0: "in",  1: "mm",  2: "pt",  3: "pica",  4: "cm",
  5: "Q/H",  6: "px",  7: "ft/in",  8: "m",  9: "yd",  10: "ft"
};

function getCurrentUnitLabel() {
  /* 現在の単位ラベルを取得 / Get current document unit label */
  var unitCode = app.preferences.getIntegerPreference("rulerType");
  return unitLabelMap[unitCode] || "pt";
}

/* UIで使用するラベル定義 / Labels for UI */
var LABELS = {
  shapePanel:  { ja: "形状", en: "Shape" },
  shapeA4:     { ja: "A4 (210 x 297)", en: "A4 (210 x 297)" },
  shapeSquare: { ja: "スクエア", en: "Square" },
  columns:     { ja: "列数", en: "Columns" },
  rows:        { ja: "行数", en: "Rows" },
  offsetLabel: { ja: "オフセット", en: "Offset" },
  okBtn:       { ja: "実行", en: "Run" },
  cancel:      { ja: "キャンセル", en: "Cancel" }
};

/* 画像サイズに基づく初期グリッドサイズ計算 / Calculate initial grid size based on image size */
function getInitialGridSize(imageWidth, imageHeight) {
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

    /* --- ダイアログ初期設定を関数化することを推奨（最適化提案） --- */
    // TODO: ダイアログ初期設定を別関数に分離して可読性・再利用性向上

    /* --- ダイアログ作成 / Create dialog --- */
    var dialogTitle = (lang === "ja" ? "分割してアートボード化" : "Slice and Create Artboards") + " v1.1";
    var Dialog = new Window('dialog', dialogTitle);
    Dialog.orientation = 'column';
    Dialog.alignment = 'right';

    /* --- 入力パネル / Input Panel --- */
    var inputPanel = Dialog.add('group');
    inputPanel.orientation = 'column';
    inputPanel.alignChildren = 'left';

    /* --- 形状パネル / Shape panel (A4/Square) --- */
    var shapePanel = inputPanel.add("panel", undefined, LABELS.shapePanel[lang]);
    shapePanel.orientation = "column";
    shapePanel.alignChildren = "left";
    shapePanel.margins = [15, 20, 15, 10];
    var shapeRadioA4 = shapePanel.add("radiobutton", undefined, LABELS.shapeA4[lang]);
    var shapeRadioSquare = shapePanel.add("radiobutton", undefined, LABELS.shapeSquare[lang]);
    shapeRadioA4.value = true; // デフォルトA4 / Default is A4

    /* --- ピース設定グループ / Piece setting group --- */
    var pieceSettingGroup = inputPanel.add("group");
    pieceSettingGroup.orientation = "column";
    pieceSettingGroup.alignChildren = "left";
    pieceSettingGroup.margins = [10, 5, 10, 15];

    /* --- 行数・列数入力欄 / Row and Column Inputs --- */
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

    /* --- オフセットグループ / Offset Group --- */
    var offsetGroup = inputPanel.add("group");
    offsetGroup.orientation = "row";
    offsetGroup.alignChildren = "left";
    offsetGroup.margins = [10, 0, 10, 10];
    var offsetCheckbox = offsetGroup.add('checkbox', undefined, LABELS.offsetLabel[lang]);
    offsetCheckbox.value = true;
    var offsetValueInput = offsetGroup.add("edittext", undefined, "-20");
    offsetValueInput.characters = 4;
    var offsetUnitLabel = offsetGroup.add("statictext", undefined, getCurrentUnitLabel());
    offsetValueInput.enabled = offsetCheckbox.value;
    offsetUnitLabel.enabled = offsetCheckbox.value;
    offsetCheckbox.onClick = function () {
      offsetValueInput.enabled = offsetCheckbox.value;
      offsetUnitLabel.enabled = offsetCheckbox.value;
    };

    /* --- OK・キャンセルボタン / OK & Cancel Buttons --- */
    var groupButtons = Dialog.add('group');
    groupButtons.orientation = 'row';
    groupButtons.alignment = "right";
    groupButtons.add('button', undefined, LABELS.cancel[lang]);
    var okBtn = groupButtons.add('button', undefined, LABELS.okBtn[lang], { name: "ok" });
    okBtn.active = true;

    /* --- 選択画像やベクターアートワークがあれば初期値を自動設定 / Auto-set grid if artwork selected --- */
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
    /* --- 行列数の自動計算（A4比率 or スクエア） / Auto-calc grid by shape --- */
    var selectedImage = getSelectedArtworkItem();
    function updateGridByShape() {
      if (!selectedImage) return;
      var aspect = selectedImage.width / selectedImage.height;
      var targetAspect = shapeRadioA4.value ? (210 / 297) : 1.0;
      if (aspect > targetAspect) {
        rowText.text = "1";
        columnText.text = String(Math.max(1, Math.round(aspect / targetAspect)));
      } else {
        columnText.text = "1";
        rowText.text = String(Math.max(1, Math.round(targetAspect / aspect)));
      }
    }
    if (selectedImage) {
      updateGridByShape();
    }
    shapeRadioA4.onClick = updateGridByShape;
    shapeRadioSquare.onClick = updateGridByShape;

    // --- メイン処理（OK押下時） / Main logic on OK ---
    if (
      app.documents.length > 0 &&
      app.selection.length > 0 &&
      Dialog.show() == 1 &&
      (columnText.text || rowText.text)
    ) {
      var selectedObj;
      var origImageObj = null;
      var selectedObjectForPuzzle;
      var isTempRect = false;

      // シンボル化処理を共通関数で実行
      var symbolResult = convertToSymbol(app.selection);
      selectedObj = symbolResult.symbolItem;
      selectedObjectForPuzzle = symbolResult.maskRect ? symbolResult.maskRect : selectedObj;
      origImageObj = symbolResult.origImageObj ? symbolResult.origImageObj : null;
      isTempRect = symbolResult.isTempRect ? true : false;
      selectedImage = app.selection[0];

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

        /* マスク矩形のアスペクト比決定 / Mask aspect ratio by shape radio */
        var maskAspectRatio = 1.0;
        if (shapeRadioA4.value) {
          maskAspectRatio = 210 / 297;
        } else if (shapeRadioSquare.value) {
          maskAspectRatio = 1.0;
        }

        // オブジェクト左上座標 / Top-left coordinate
        var originX = selectedObjectForPuzzle.geometricBounds[0];
        var originY = selectedObjectForPuzzle.geometricBounds[3];
        var objWidth = selectedObjectForPuzzle.geometricBounds[2] - selectedObjectForPuzzle.geometricBounds[0];
        var objHeight = selectedObjectForPuzzle.geometricBounds[1] - selectedObjectForPuzzle.geometricBounds[3];
        if (objHeight < 0) objHeight = -objHeight;

        // ピースごとの幅・高さを形状比率に固定し中央揃え / Fix piece size by aspect and center
        var pieceWidth, pieceHeight;
        var gridTotalWidth, gridTotalHeight;
        var w0 = objWidth / columnCount;
        var h0 = objHeight / rowCount;
        if (w0 / h0 > maskAspectRatio) {
          pieceHeight = h0;
          pieceWidth = pieceHeight * maskAspectRatio;
        } else {
          pieceWidth = w0;
          pieceHeight = pieceWidth / maskAspectRatio;
        }
        gridTotalWidth = pieceWidth * columnCount;
        gridTotalHeight = pieceHeight * rowCount;
        originX = selectedObjectForPuzzle.geometricBounds[0] + (objWidth - gridTotalWidth) / 2;
        originY = selectedObjectForPuzzle.geometricBounds[3] + (objHeight - gridTotalHeight) / 2 + gridTotalHeight;

        /* --- マスクピース生成ループの共通処理化を推奨（最適化提案） --- */
        // TODO: マスクピース生成ループを関数化して共通処理にすることで、拡張や再利用性を高められます

        /* --- グリッド形状（矩形マスク）モード / Grid mask mode --- */
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
              gridMask = offsetRectangle(gridMask, offsetVal, offsetVal);
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
    } else {
      Dialog.hide();
    }
  } catch (e) {
    alert("スクリプト実行中にエラーが発生しました: " + e);
  }
}

/* --- Offset Path Effect Utility --- */
function createOffsetEffectXML(offsetVal) {
  /* Offset PathエフェクトXMLを生成 / Generate Offset Path effect XML */
  var xml = '<LiveEffect name="Adobe Offset Path"><Dict data="R mlim 4 R ofst value I jntp 2 "/></LiveEffect>';
  return xml.replace("value", offsetVal);
}

function applyOffsetPathToSelection(offsetVal) {
  /* 選択アイテムにオフセットパス効果を適用 / Apply Offset Path effect to selection */
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

/* --- 矩形の座標をオフセット調整（左右・上下独立） / Offset rectangle bounds (independent X/Y) --- */
function offsetRectangle(pathItem, offsetX, offsetY) {
  var gb = pathItem.geometricBounds;
  var newWidth = (gb[2] - gb[0]) + offsetX * 2;
  var newHeight = (gb[1] - gb[3]) + offsetY * 2;
  pathItem.left = pathItem.left - offsetX;
  pathItem.top = pathItem.top + offsetY;
  pathItem.width = newWidth;
  pathItem.height = newHeight;
  return pathItem;
}

/**
 * 選択オブジェクトをシンボル化し、必要なら矩形を生成して返す
 * Symbolize selected object(s), create rectangle if needed
 * @param {Array|Object} selection - app.selection
 * @return {Object} {symbolItem, maskRect, origImageObj, isTempRect}
 */
function convertToSymbol(selection) {
  var result = {
    symbolItem: null,
    maskRect: null,
    origImageObj: null,
    isTempRect: false
  };
  var doc = app.activeDocument;
  var sel;
  if (selection instanceof Array || selection.length > 1) {
    /* 複数選択時: グループ化してシンボル化 / When multiple selected: group and symbolize */
    try {
      var tempGroup = doc.groupItems.add();
      var selectedItems = [];
      for (var i = 0; i < selection.length; i++) {
        selectedItems.push(selection[i]);
      }
      // 重ね順保持 / Keep stacking order
      for (var j = selectedItems.length - 1; j >= 0; j--) {
        selectedItems[j].move(tempGroup, ElementPlacement.PLACEATBEGINNING);
      }
      var groupLeft = tempGroup.left;
      var groupTop = tempGroup.top;
      var groupParent = tempGroup.parent;
      var groupedSymbol = doc.symbols.add(tempGroup);
      var symbolItem = groupParent.symbolItems.add(groupedSymbol);
      symbolItem.left = groupLeft;
      symbolItem.top = groupTop;
      tempGroup.remove();
      sel = symbolItem;
    } catch (e) {
      alert("複数オブジェクトのシンボル化に失敗しました: " + e);
      throw e;
    }
  } else {
    sel = selection[0];
  }

  // ラスターやベクターの場合はシンボル化 / Symbolize if raster or vector
  if (sel.typename === "RasterItem") {
    try {
      var rasterLeft = sel.left;
      var rasterTop = sel.top;
      var rasterParent = sel.parent;
      var tempSymbol = doc.symbols.add(sel);
      var symbolItem = rasterParent.symbolItems.add(tempSymbol);
      symbolItem.left = rasterLeft;
      symbolItem.top = rasterTop;
      sel.remove();
      sel = symbolItem;
    } catch (e) {
      alert("埋め込み画像のシンボル化に失敗しました: " + e);
      throw e;
    }
  } else if (
    sel.typename === "PathItem" ||
    sel.typename === "GroupItem" ||
    sel.typename === "CompoundPathItem"
  ) {
    try {
      var vectorLeft = sel.left;
      var vectorTop = sel.top;
      var vectorParent = sel.parent;
      var vectorSymbol = doc.symbols.add(sel);
      var vectorSymbolItem = vectorParent.symbolItems.add(vectorSymbol);
      vectorSymbolItem.left = vectorLeft;
      vectorSymbolItem.top = vectorTop;
      sel.remove();
      sel = vectorSymbolItem;
    } catch (e) {
      alert("ベクターオブジェクトのシンボル化に失敗しました: " + e);
      throw e;
    }
  }

  // シンボル/画像の場合は矩形を作成 / Create rectangle if symbol/image
  var selType = sel.typename;
  if (selType == "PlacedItem" || selType == "RasterItem" || selType == "SymbolItem") {
    var gb = sel.geometricBounds;
    var rectWidth = gb[2] - gb[0];
    var rectHeight = gb[1] - gb[3];
    if (rectHeight < 0) rectHeight = -rectHeight;
    var rect = doc.pathItems.rectangle(gb[1], gb[0], rectWidth, rectHeight);
    result.maskRect = rect;
    result.origImageObj = sel;
    result.isTempRect = true;
  }
  result.symbolItem = sel;
  return result;
}