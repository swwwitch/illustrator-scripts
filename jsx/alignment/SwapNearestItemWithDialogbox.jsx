#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択中のオブジェクトと、指定した方向（上下左右）にある最も近いオブジェクトの位置を入れ替えます。方向はダイアログで指定します。

詳細は README を参照してください。

### Overview

Swaps the selected object with the nearest object in a chosen direction, picked from a dialog.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SwapNearestItemWithDialogbox"; /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-07-06";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-06";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SwapNearestItemWithDialogbox.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SwapNearestItemWithDialogbox.md

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

    /* 起動直後の方向。ダイアログの矢印キーで切り替わる / Initial direction; the arrow keys change it */
    var direction = directionMap[4];

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
      dialog: {
        title:    { ja: "オブジェクトを入れ替え", en: "Swap Position" },
        message:  { ja: "矢印キーで入れ替え", en: "Press arrow key to swap" },
        notFound: { ja: "見つかりません", en: "Nothing found" }
      },
      button: {
        close: { ja: "終了", en: "Close" }
      },
      alert: {
        noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
        selectOne:  { ja: "1つのオブジェクトを選択してください。", en: "Select one object." },
        noPosition: { ja: "位置情報を持つオブジェクトを選択してください。", en: "Select an object that has a position." }
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

    /**
     * 選択中のオブジェクトと、指定方向で最も近いオブジェクトの位置を入れ替える
     * @returns {boolean} 入れ替えた場合はtrue
     */
    function swapNearestObject() {
      var sel = app.activeDocument.selection;
      /* 実行中に選択が変わることがあるため、毎回確認する（案内はmain側で済ませている）/ The selection can change while the dialog is open */
      if (!sel || sel.typename === "TextRange" || sel.length !== 1) return false;

      var target = sel[0];
      if (!isValidType(target)) return false;

      var nearest = findNearestObjectInDirection(target, direction, sel);
      if (!nearest) return false;

      if (direction === "up" || direction === "down") {
        swapVertically(target, nearest);
      } else {
        swapHorizontally(target, nearest);
      }

      return true;
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 矢印キーで入れ替えを実行するダイアログを表示する
     * @returns {void}
     */
    function showDialog() {
      var dlg = new Window("dialog", getLabel("dialog.title"));

      var msgText = dlg.add("statictext", undefined, getLabel("dialog.message"));
      msgText.alignment = "center";

      var btnClose = dlg.add("button", undefined, getLabel("button.close"));
      btnClose.onClick = function () {
        dlg.close();
      };

      var isProcessing = false;

      dlg.addEventListener("keydown", function (e) {
        if (isProcessing) return;
        isProcessing = true;

        switch (e.keyName) {
          case "Up": direction = directionMap[3]; break;
          case "Down": direction = directionMap[4]; break;
          case "Left": direction = directionMap[2]; break;
          case "Right": direction = directionMap[1]; break;
          case "Escape":
          case "Enter":
          case "Return":
            btnClose.notify();
            return;
          default:
            isProcessing = false;
            return;
        }

        /* 見つからないときはメッセージを差し替えて知らせる。連打を邪魔しないようアラートは出さない
           Report a miss by swapping the message; an alert would interrupt repeated key presses */
        msgText.text = swapNearestObject() ? getLabel("dialog.message") : getLabel("dialog.notFound");
        app.redraw();
      }, false);

      dlg.addEventListener("keyup", function () {
        isProcessing = false;
      }, false);

      btnClose.active = true;
      dlg.show();
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択内容を確認してダイアログを表示する
     * @returns {void}
     */
    function main() {
      if (app.documents.length === 0) {
        alert(getLabel("alert.noDocument"));
        return;
      }

      var sel = app.activeDocument.selection;

      /* 文字カーソルが立っているとselはTextRangeで、lengthは文字数になる / A text caret gives a TextRange whose length counts characters */
      if (!sel || sel.typename === "TextRange" || sel.length !== 1) {
        alert(getLabel("alert.selectOne"));
        return;
      }

      if (!isValidType(sel[0])) {
        alert(getLabel("alert.noPosition"));
        return;
      }

      showDialog();
    }

    main();

})();
