#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したオブジェクトを基準に、指定した方向（右／左／上／下）にある最も近いオブジェクトと位置を入れ替えます。

詳細は README を参照してください。

### Overview

Swaps the selected object with the nearest object in a chosen direction — right, left, up or down.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SwapNearestItem";              /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.4";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-06-10";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-06";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SwapNearestItem.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SwapNearestItem.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n21a03e135423"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // 基本設定 / Basic settings
    // =========================================

    var directionMap = {
      1: "right",
      2: "left",
      3: "up",
      4: "down"
    };

    /* 入れ替え先を探す方向。1=右／2=左／3=上／4=下 / Search direction: 1=right, 2=left, 3=up, 4=down */
    var direction = directionMap[2];

    // =========================================
    // ラベル定義 / Label definitions
    // =========================================

    /**
     * 現在のロケールに基づき言語コードを返す
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
      return (String($.locale || "").indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = getCurrentLang();

    /* カテゴリ分けした日英ラベル定義 / Categorized Japanese-English label definitions */
    var LABELS = {
      alert: {
        noDocument:  { ja: "ドキュメントが開かれていません。", en: "No document is open." },
        noSelection: { ja: "1つ以上のオブジェクトを選択してください。", en: "Select at least one object." },
        noPosition:  { ja: "位置情報を持つオブジェクトを選択してください。", en: "Select an object that has a position." },
        noTarget:    { ja: "その方向に入れ替えられるオブジェクトが見つかりません。", en: "No object to swap with in that direction." }
      }
    };

    /**
     * ラベル定義からUI言語に応じた文字列を取得する
     * @param {string} labelPath - "alert.noDocument" のようなドット区切りのパス
     * @returns {string} 対応する文字列。見つからない場合はパスをそのまま返す
     */
    function getLabel(labelPath) {
      var pathKeys = labelPath.split(".");
      var labelNode = LABELS;
      for (var keyIndex = 0; keyIndex < pathKeys.length; keyIndex++) {
        labelNode = labelNode[pathKeys[keyIndex]];
        if (!labelNode) return labelPath;
      }
      return labelNode[uiLang] || labelNode.en || labelPath;
    }

    // =========================================
    // 判定・計測 / Checks and measurement
    // =========================================

    /**
     * 入れ替えの対象にできるオブジェクトか判定する
     * @param {PageItem} item - 判定するオブジェクト
     * @returns {boolean} 対象にできる場合はtrue
     */
    function isValidType(item) {
      /* ガイドは typename が "PathItem" のため guides で判定する / Guides are PathItems, so test .guides */
      if (item.typename === "PathItem" && item.guides === true) return false;

      var types = [
        "PathItem", "CompoundPathItem", "GroupItem", "TextFrame",
        "PlacedItem", "RasterItem", "SymbolItem", "MeshItem",
        "PluginItem", "GraphItem"
      ];

      for (var i = 0; i < types.length; i++) {
        if (types[i] === item.typename) return true;
      }

      return false;
    }

    /**
     * ロック・非表示のレイヤーに属していないか判定する
     * @param {Layer} layer - 判定するレイヤー
     * @returns {boolean} 触れる場合はtrue
     */
    function isUsableLayer(layer) {
      /* サブレイヤーは自身がロックされていなくても親レイヤー側でロックされることがある
         A sublayer can be locked or hidden by an ancestor layer even when its own flags are clear */
      for (var node = layer; node && node.typename === "Layer"; node = node.parent) {
        if (node.locked || node.visible === false) return false;
      }
      return true;
    }

    /**
     * 配列に指定したオブジェクトが含まれるか判定する
     * @param {PageItem[]} items - 探す対象の配列
     * @param {PageItem} item - 探すオブジェクト
     * @returns {boolean} 含まれる場合はtrue
     */
    function containsItem(items, item) {
      if (!items) return false;
      for (var i = 0; i < items.length; i++) {
        if (items[i] === item) return true;
      }
      return false;
    }

    /**
     * オブジェクトの境界を返す
     * @param {PageItem} item - 対象オブジェクト
     * @returns {number[]} [left, top, right, bottom]
     */
    function getBounds(item) {
      return item.visibleBounds;
    }

    /**
     * オブジェクトの中心座標を返す
     * @param {PageItem} item - 対象オブジェクト
     * @returns {number[]} [x, y]
     */
    function getCenter(item) {
      var b = getBounds(item);
      return [(b[0] + b[2]) / 2, (b[1] + b[3]) / 2];
    }

    // =========================================
    // 探索と入れ替え / Search and swap
    // =========================================

    /**
     * 指定方向にある最も近いオブジェクトを検索する
     * @param {PageItem} referenceItem - 基準オブジェクト
     * @param {string} direction - 検索方向（"right" / "left" / "up" / "down"）
     * @param {PageItem[]} excludeItems - 候補から除外するオブジェクト（省略可）
     * @returns {PageItem|null} 見つかったオブジェクト。なければnull
     */
    function findNearestObjectInDirection(referenceItem, direction, excludeItems) {
      var doc = app.activeDocument;
      var items = doc.pageItems;
      var itemCount = items.length;

      var refBounds = getBounds(referenceItem);
      var refCenter = getCenter(referenceItem);

      var nearest = null;
      var minGap = Number.MAX_VALUE;
      var minDist = Number.MAX_VALUE;

      for (var i = 0; i < itemCount; i++) {
        var item = items[i];
        if (item === referenceItem) continue;
        if (containsItem(excludeItems, item)) continue;
        if (item.locked || item.hidden) continue;
        if (!isUsableLayer(item.layer)) continue;
        if (!isValidType(item)) continue;
        /* グループや複合パス内の子オブジェクトは除外（親のみ処理対象）/ Skip children of groups and compound paths */
        if (item.parent && (item.parent.typename === "GroupItem" || item.parent.typename === "CompoundPathItem")) continue;

        var bounds = getBounds(item);
        var center = getCenter(item);
        var dx = center[0] - refCenter[0];
        var dy = center[1] - refCenter[1];

        /* 探索軸と直交する方向で範囲が重なっているか / Overlap on the axis perpendicular to the search */
        var verticalOverlap = !(refBounds[3] > bounds[1] || refBounds[1] < bounds[3]);
        var horizontalOverlap = !(refBounds[0] > bounds[2] || refBounds[2] < bounds[0]);

        /* 近さは探索方向の隙間で測る（中心間の直線距離では斜めのオブジェクトが勝つ）/ Score by the gap along the search axis */
        var gap;
        if (direction === "right" && dx > 0 && verticalOverlap) {
          gap = bounds[0] - refBounds[2];
        } else if (direction === "left" && dx < 0 && verticalOverlap) {
          gap = refBounds[0] - bounds[2];
        } else if (direction === "up" && dy > 0 && horizontalOverlap) {
          gap = bounds[3] - refBounds[1];
        } else if (direction === "down" && dy < 0 && horizontalOverlap) {
          gap = refBounds[3] - bounds[1];
        } else {
          continue;
        }

        if (gap < 0) gap = 0; /* 重なっている場合は隙間0として扱う / Treat overlapping items as a zero gap */

        var distance = Math.sqrt(dx * dx + dy * dy);
        if (gap < minGap || (gap === minGap && distance < minDist)) {
          minGap = gap;
          minDist = distance;
          nearest = item;
        }
      }

      return nearest;
    }

    /**
     * 2つのオブジェクトを左右に入れ替える（全体の占有範囲と隙間を保つ）
     * @param {PageItem} itemA - 入れ替える一方
     * @param {PageItem} itemB - 入れ替えるもう一方
     * @returns {void}
     */
    function swapHorizontally(itemA, itemB) {
      var boundsA = getBounds(itemA);
      var boundsB = getBounds(itemB);

      /* itemAが左、itemBが右になるように並び替え / Order so that itemA is the left one */
      if (boundsA[0] > boundsB[0]) {
        var tmpItem = itemA;
        itemA = itemB;
        itemB = tmpItem;

        var tmpBounds = boundsA;
        boundsA = boundsB;
        boundsB = tmpBounds;
      }

      var widthB = boundsB[2] - boundsB[0];
      var gap = boundsB[0] - boundsA[2];

      itemB.left = boundsA[0];
      itemA.left = boundsA[0] + widthB + gap;
    }

    /**
     * 2つのオブジェクトを上下に入れ替える（全体の占有範囲と隙間を保つ）
     * @param {PageItem} itemA - 入れ替える一方
     * @param {PageItem} itemB - 入れ替えるもう一方
     * @returns {void}
     */
    function swapVertically(itemA, itemB) {
      var boundsA = getBounds(itemA);
      var boundsB = getBounds(itemB);

      /* itemAが上、itemBが下になるように並び替え / Order so that itemA is the upper one */
      if (boundsA[1] < boundsB[1]) {
        var tmpItem = itemA;
        itemA = itemB;
        itemB = tmpItem;

        var tmpBounds = boundsA;
        boundsA = boundsB;
        boundsB = tmpBounds;
      }

      var heightB = boundsB[1] - boundsB[3];
      var gap = boundsA[3] - boundsB[1];

      itemB.top = boundsA[1];
      itemA.top = boundsA[1] - heightB - gap;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択内容に応じて入れ替えを実行する
     * @returns {void}
     */
    function main() {
      if (app.documents.length === 0) {
        alert(getLabel("alert.noDocument"));
        return;
      }

      var doc = app.activeDocument;
      var sel = doc.selection;

      /* 文字カーソルが立っているとselはTextRangeで、lengthは文字数になる / A text caret gives a TextRange whose length counts characters */
      if (!sel || sel.typename === "TextRange" || sel.length < 1) {
        alert(getLabel("alert.noSelection"));
        return;
      }

      /* 2つ選択時は探索せずにその2つを入れ替える / With two objects selected, swap them directly */
      if (sel.length === 2) {
        if (direction === "up" || direction === "down") {
          swapVertically(sel[0], sel[1]);
        } else {
          swapHorizontally(sel[0], sel[1]);
        }
        return;
      }

      var target = sel[0];

      if (!isValidType(target)) {
        alert(getLabel("alert.noPosition"));
        return;
      }

      /* 選択中のオブジェクト同士で入れ替わらないよう、選択全体を候補から外す / Keep the selection out of the candidates */
      var nearest = findNearestObjectInDirection(target, direction, sel);
      if (!nearest) {
        alert(getLabel("alert.noTarget"));
        return;
      }

      if (direction === "up" || direction === "down") {
        swapVertically(target, nearest);
      } else {
        swapHorizontally(target, nearest);
      }
    }

    main();

})();
