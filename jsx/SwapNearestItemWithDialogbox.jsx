#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
  スクリプト名：SwapNearestItem.jsx
  
  概要：
    選択中のオブジェクトと、指定方向（上下左右）にある最も近いオブジェクトの位置を入れ替える。
    ダイアログボックスを表示中、↑↓キーで上下、←→キーで左右のオブジェクトを入れ替える
  処理の流れ：
    1. ドキュメントと選択状態を確認
    2. 指定方向にある最も近いPageItemを検索
    3. 該当オブジェクトがあれば、元のオブジェクトと座標を交換する
  備考：
  ・Escapeキーでダイアログボックスを閉じる
  ・⌘ + returnキー、option + returnキーでもOS が特別に解釈してkeyイベントとして伝えるため閉じる
  謝辞：
  @ken https://x.com/ken_rainy
*/

var directionMap = {
  1: "right",
  2: "left",
  3: "up",
  4: "down"
};

var direction = directionMap[4];

var LABELS = {
  ja: {
    title: "オブジェクトを入れ替え",
    message: "矢印キーで入れ替え",
    close: "終了"
  },
  en: {
    title: "Swap Position",
    message: "Press arrow key to swap",
    close: "Close"
  }
};

/**
 * 現在のロケールに基づき言語コードを返す
 * @returns {string} "ja" または "en"
 */
function getCurrentLang() {
  return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var L = LABELS[getCurrentLang()];

(function(){
  if(app.activeDocument.selection.length != "1"){
    alert("1つのオブジェクトを選択してください。");
    return;
  }

/**
 * 指定したタイプ名が処理対象か判定する
 * @param {string} name - オブジェクトの typename
 * @returns {boolean} 処理対象ならtrue
 */
function isValidType(name) {
  // 対象外の型（ガイドや注釈など）
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

/**
 * 指定方向にある最も近いオブジェクトを検索する
 * @param {PageItem} referenceItem - 基準オブジェクト
 * @param {string} direction - 検索方向 ("right", "left", "up", "down")
 * @returns {PageItem|null} - 見つかった最も近いオブジェクト、なければnull
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

    var cxRef = referenceItem.left + referenceItem.width / 2;
    var cyRef = referenceItem.top - referenceItem.height / 2;
    var cxItem = item.left + item.width / 2;
    var cyItem = item.top - item.height / 2;

    var dx = cxItem - cxRef;
    var dy = cyItem - cyRef;

    var distance = Math.sqrt(dx * dx + dy * dy);

    // 縦方向の重なり判定（左右方向の場合）
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

/**
 * 選択中オブジェクトと指定方向の最も近いオブジェクトの位置を入れ替える
 */
function swapNearestObject() {
  if (app.documents.length === 0) {
    alert("ドキュメントが開かれていません。");
    return;
  }

  var doc = app.activeDocument;
  var sel = doc.selection;

  if (!sel || sel.length !== 1) {
    alert("1つのオブジェクトを選択してください。");
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

  var boundsA = target.visibleBounds;
  var boundsB = nearest.visibleBounds;

  if (direction === "up" || direction === "down") {
    var yA = boundsA[1];
    var hA = boundsA[1] - boundsA[3];
    var yB = boundsB[1];
    var hB = boundsB[1] - boundsB[3];

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
    var xA = boundsA[0];
    var wA = boundsA[2] - boundsA[0];
    var xB = boundsB[0];
    var wB = boundsB[2] - boundsB[0];

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

var winObj = new Window("dialog", L.title);
winObj.txt = winObj.add("statictext", undefined, L.message);
winObj.txt.alignment = "center";

var bt = winObj.add("button", undefined, L.close);
bt.onClick = function(){
  winObj.close();
};

var isProcessing = false;

winObj.addEventListener("keydown", function (e) {
  if (isProcessing) return;
  isProcessing = true;

  var k = e.keyName;
  switch (k) {
    case "Up": direction = directionMap[3]; break;
    case "Down": direction = directionMap[4]; break;
    case "Left": direction = directionMap[2]; break;
    case "Right": direction = directionMap[1]; break;
    case "Escape":
    case "Enter":
    case "Return":
      bt.notify();
      return;
    default:
      isProcessing = false;
      return;
  }

  swapNearestObject();
  redraw();
}, false);

winObj.addEventListener("keyup", function (e) {
  isProcessing = false;
}, false);

bt.active = true;

winObj.show();
})();
