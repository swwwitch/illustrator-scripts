#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
### スクリプト名：

SwapNearestItem.jsx

### 概要

- 選択したオブジェクトを基準に、指定方向（右／左／上／下）にある最も近いオブジェクトと自然な見た目で位置を入れ替えるIllustrator用スクリプトです。
- 複数選択時の特別処理や幅・高さ、隙間を考慮した入れ替えに対応します。

### 主な機能

- 上下左右方向での最短距離判定によるスワップ
- 幅・高さ、隙間を考慮した自然な位置調整
- 複数選択時の手動入れ替え処理対応
- 日本語／英語インターフェース対応

### 処理の流れ

1. ドキュメントと選択オブジェクトを確認
2. 指定方向に最も近いオブジェクトを検索
3. 高さ・幅・隙間を考慮して位置を入れ替え
4. 複数選択時は中心座標または端基準で入れ替え

### 更新履歴

- v1.0.0 (20250610) : 初版リリース
- v1.0.1 (20250612) : グループ・複合パスの除外、ロック対応追加
- v1.0.2 (20250613) : getCenter() と getSize() の導入による整理
- v1.0.3 (20250614) : 複数選択時の一時グループ処理追加
- v1.0.4 (20250615) : 一時グループ化処理の削除、整理

---

### Script Name:

SwapNearestItem.jsx

### Overview

- An Illustrator script that swaps the position of a selected object with the nearest object in a specified direction (right, left, up, or down), maintaining a natural visual layout.
- Supports special handling for multi-selection and considers widths, heights, and gaps.

### Main Features

- Swap based on shortest distance in four directions
- Adjust positions naturally considering widths, heights, and gaps
- Special handling for manual swaps with two selections
- Japanese and English UI support

### Process Flow

1. Check document and selected objects
2. Find the nearest object in the specified direction
3. Swap positions considering sizes and gaps
4. Handle multi-selection with center or edge-based swaps

### Update History

- v1.0.0 (20250610): Initial release
- v1.0.1 (20250612): Added exclusion of group/compound path children and lock support
- v1.0.2 (20250613): Refactored with getCenter() and getSize()
- v1.0.3 (20250614): Added temporary group handling for multiple selection
- v1.0.4 (20250615): Removed temporary grouping, cleaned up logic
*/

var directionMap = {
  1: "right",
  2: "left",
  3: "up",
  4: "down"
};

var direction = directionMap[2];

// 有効なオブジェクト型かどうかを判定する
function isValidType(name) {
  if (name === "LegacyTextItem" || name === "NonNativeItem" ||
      name === "Guide" || name === "Annotation" || name === "PathPoint") {
    return false;
  }

  var types = [
    "PathItem", "CompoundPathItem", "GroupItem", "TextFrame",
    "PlacedItem", "RasterItem", "SymbolItem", "MeshItem",
    "PluginItem", "GraphItem"
  ];

  for (var i = 0; i < types.length; i++) {
    if (types[i] === name) return true;
  }

  return false;
}

// オブジェクトの中心座標 [x, y] を返す
function getCenter(item) {
  return [item.left + item.width / 2, item.top - item.height / 2];
}
// オブジェクトの境界 [left, top, right, bottom] を返す
function getBounds(item) {
  return item.visibleBounds;
}

// オブジェクトの幅と高さ [width, height] を返す
function getSize(item) {
  var b = item.visibleBounds;
  return [b[2] - b[0], b[1] - b[3]];
}

/**
 * 指定方向にある最も近いオブジェクトを検索する
 * 基準オブジェクトと比較し、方向と重なり条件を満たすものの中で最短距離のものを返す
 * @param {PageItem} referenceItem - 基準オブジェクト
 * @param {string} direction - 検索方向 ("right", "left", "up", "down")
 * @returns {PageItem|null} - 最も近いオブジェクト、なければnull
 */
function findNearestObjectInDirection(referenceItem, direction) {
  var doc = app.activeDocument;
  var nearest = null;
  var minDist = Number.MAX_VALUE;

  for (var i = 0; i < doc.pageItems.length; i++) {
    var item = doc.pageItems[i];
    if (item === referenceItem) continue;
    if (item.locked || item.hidden) continue;
    if (item.layer.locked || item.layer.visible === false) continue;
    if (!isValidType(item.typename)) continue;
    if (item.parent && (item.parent.typename === "GroupItem" || item.parent.typename === "CompoundPathItem")) {
      if (item !== item.parent) {
        continue; // グループや複合パス内の子オブジェクトは除外（親のみ処理対象）
      }
    }

    var centerRef = getCenter(referenceItem);
    var centerItem = getCenter(item);
    var dx = centerItem[0] - centerRef[0];
    var dy = centerItem[1] - centerRef[1];

    var distance = Math.sqrt(dx * dx + dy * dy);

    var aBounds = referenceItem.visibleBounds;
    var bBounds = item.visibleBounds;

    var ay1 = aBounds[1]; // top
    var ay2 = aBounds[3]; // bottom
    var by1 = bBounds[1];
    var by2 = bBounds[3];
    var verticalOverlap = !(ay2 > by1 || ay1 < by2);

    if (direction === "right" && dx > 0 && verticalOverlap) {
      if (distance < minDist) {
        minDist = distance;
        nearest = item;
      }
    } else if (direction === "left" && dx < 0 && verticalOverlap) {
      if (distance < minDist) {
        minDist = distance;
        nearest = item;
      }
    } else if (direction === "down" && dy < 0 && Math.abs(dx) < referenceItem.width / 2) {
      if (distance < minDist) {
        minDist = distance;
        nearest = item;
      }
    } else if (direction === "up" && dy > 0 && Math.abs(dx) < referenceItem.width / 2) {
      if (distance < minDist) {
        minDist = distance;
        nearest = item;
      }
    }
  }

  return nearest;
}

function main() {
  if (app.documents.length === 0) {
    alert("ドキュメントが開かれていません。");
    return;
  }

  var doc = app.activeDocument;
  var sel = doc.selection;

  if (!sel || sel.length < 1) {
    alert("1つ以上のオブジェクトを選択してください。");
    return;
  }

  // 2つ選択時かつ左右方向なら手動入れ替え
  if (sel.length === 2 && (direction === "right" || direction === "left")) {
    var itemA = sel[0];
    var itemB = sel[1];

    var boundsA = getBounds(itemA);
    var boundsB = getBounds(itemB);

    var xA = boundsA[0];
    var xB = boundsB[0];

    var wA = boundsA[2] - boundsA[0];
    var wB = boundsB[2] - boundsB[0];

    // itemAが左、itemBが右になるように並び替え
    if (xA > xB) {
      var tmp = itemA;
      itemA = itemB;
      itemB = tmp;

      var tmpBounds = boundsA;
      boundsA = boundsB;
      boundsB = tmpBounds;

      var tmpX = xA;
      xA = xB;
      xB = tmpX;

      var tmpW = wA;
      wA = wB;
      wB = tmpW;
    }

    // itemB（右）のX座標を itemA（左）と同じに
    itemB.left = xA;

    // itemA（左）のX座標を itemBの右端から自身の幅だけ左に
    itemA.left = xA + wB + (xB - (xA + wA));

    return;
  }

  // 2つ選択時かつ上下方向なら手動入れ替え（中心座標を入れ替える）
  if (sel.length === 2 && (direction === "up" || direction === "down")) {
    var itemA = sel[0];
    var itemB = sel[1];

    var centerA = getCenter(itemA);
    var centerB = getCenter(itemB);

    if (direction === "up") {
      var deltaY = centerA[1] - centerB[1];
      itemA.top -= deltaY;
      itemB.top += deltaY;
    } else {
      var deltaY = centerB[1] - centerA[1];
      itemA.top += deltaY;
      itemB.top -= deltaY;
    }

    return;
  }

  var target = sel[0];

  if (!target.hasOwnProperty("left") || !target.hasOwnProperty("top")) {
    alert("位置情報を持つオブジェクトを選択してください。");
    return;
  }

  var nearest = findNearestObjectInDirection(target, direction);
  if (!nearest || nearest === target) {
    return;
  }

  if (direction === "up" || direction === "down") {
    var boundsA = getBounds(target);
    var boundsB = getBounds(nearest);

    var yA = boundsA[1];
    var yB = boundsB[1];

    var sizeA = getSize(target);
    var sizeB = getSize(nearest);
    var hA = sizeA[1];
    var hB = sizeB[1];

    if (direction === "up") {
      target.top = yB;

      var gap = Math.abs(yA - boundsB[3]);
      var totalHeight = hA + hB + gap;
      nearest.top = yB - totalHeight + hB;
    } else {
      nearest.top = yA;

      var gap = Math.abs(yB - boundsA[3]);
      var totalHeight = hA + hB + gap;
      target.top = yA - totalHeight + hA;
    }
  } else {
    var boundsA = getBounds(target);
    var boundsB = getBounds(nearest);

    var xA = boundsA[0];
    var xB = boundsB[0];

    var sizeA = getSize(target);
    var sizeB = getSize(nearest);
    var wA = sizeA[0];
    var wB = sizeB[0];

    if (direction === "left") {
      var gap = Math.abs(boundsA[0] - boundsB[2]);
      target.left = xB;
      nearest.left = xB + wA + gap;
    } else {
      var gap = Math.abs(xB - boundsA[2]);

      nearest.left = xA;
      var totalWidth = wA + wB + gap;
      target.left = xA + totalWidth - wA;
    }
  }
}

main();