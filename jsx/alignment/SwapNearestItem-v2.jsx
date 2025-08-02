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
- positionMode が "center" の場合、2つのオブジェクトを選択しているときに限り、互いの中心を入れ替え（隙間や全体幅は無視）

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
- When positionMode is "center" and exactly two objects are selected, swap their centers (ignoring gaps and total width/height)

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

// 基準位置の設定: "topLeft" または "center"
var positionMode = "center";

var directionMap = {
  1: "right",
  2: "left",
  3: "up",
  4: "down"
};

var direction = directionMap[1];

// 左右の端を固定してスワップ
function swapObjectsKeepEdges(objA, objB) {
    var boundsA = objA.visibleBounds; // [x1, y1, x2, y2]
    var boundsB = objB.visibleBounds;

    var leftA  = boundsA[0];
    var rightA = boundsA[2];
    var leftB  = boundsB[0];
    var rightB = boundsB[2];

    var widthA = rightA - leftA;
    var widthB = rightB - leftB;

    // Aの左端、Bの右端を基準に位置決定
    objA.left = rightB - widthA; // AをBの右端に合わせる
    objB.left = leftA;           // BをAの左端に合わせる
}

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
// 中心基準でスワップする処理（幅や高さは無視）
function swapByCenter(itemA, itemB, direction) {
  var centerA = getCenter(itemA);
  var centerB = getCenter(itemB);

  // AをBの中心へ移動
  itemA.left += centerB[0] - centerA[0];
  itemA.top  += centerB[1] - centerA[1];

  // BをAの中心へ移動
  itemB.left += centerA[0] - centerB[0];
  itemB.top  += centerA[1] - centerB[1];
}
// 基準点を positionMode に応じて返す
function getReferencePoint(item) {
  if (positionMode === "center") {
    return getCenter(item);
  } else {
    return [item.left, item.top];
  }
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

    var refPoint = getReferencePoint(referenceItem);
    var itemPoint = getReferencePoint(item);
    var dx = itemPoint[0] - refPoint[0];
    var dy = itemPoint[1] - refPoint[1];

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

// ユーザー提供の swapObjectsByCenter 関数
function swapObjectsByCenter(objA, objB) {
    // 中心点を取得する関数
    function getCenter(obj) {
        var bounds;
        if (obj.typename === "GroupItem" && obj.clipped) {
            // クリップグループの場合はマスクパスを基準に
            var mask = null;
            for (var i = 0; i < obj.pageItems.length; i++) {
                if (obj.pageItems[i].clipping) {
                    mask = obj.pageItems[i];
                    break;
                }
            }
            if (mask) {
                bounds = mask.geometricBounds; // マスクパスの範囲
            } else {
                bounds = obj.visibleBounds; // 保険：マスクが見つからなければ全体
            }
        } else {
            bounds = obj.visibleBounds;
        }
        return [(bounds[0] + bounds[2]) / 2, (bounds[1] + bounds[3]) / 2];
    }

    // 中心座標を取得
    var centerA = getCenter(objA);
    var centerB = getCenter(objB);

    // 移動量を計算
    var deltaA = [centerB[0] - centerA[0], centerB[1] - centerA[1]];
    var deltaB = [centerA[0] - centerB[0], centerA[1] - centerB[1]];

    // 位置を入れ替え
    objA.translate(deltaA[0], deltaA[1]);
    objB.translate(deltaB[0], deltaB[1]);
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

  if (sel.length === 2) {
    swapObjectsByCenter(sel[0], sel[1]);
    return;
  }

  var target = sel[0];

  if (!target.hasOwnProperty("left") || !target.hasOwnProperty("top")) {
    alert("位置情報を持つオブジェクトを選択してください。");
    return;
  }

  var nearest = findNearestObjectInDirection(target, "right");
  if (!nearest) return;

  // 1つ選択 → 左右端固定でスワップ
  swapObjectsKeepEdges(target, nearest);
}

main();