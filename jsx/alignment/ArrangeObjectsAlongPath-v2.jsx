// ArrangeObjectsAlongPath.jsx (restored full script)
#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
  概要 / Overview
  ----------------------------------------------------------------------
  複数のオブジェクト（A）を、選択範囲内で最背面にある1本のパス（B）に沿って
  等間隔に自動配置するIllustrator用スクリプトです。

  This script arranges multiple objects (A) at even intervals
  along a single base path (B), which is determined as the
  backmost path within the current selection.

  使い方 / How to use
  ----------------------------------------------------------------------
  1. 配置したいオブジェクトを複数選択します（A）
  2. 基準にしたいパスを含めて選択します（Bは最背面のパス）
  3. スクリプトを実行し、ダイアログでオプションを指定します
  4. OKを押すと、オブジェクトがパスに沿って配置されます

  Select multiple objects to place (A), including one path to be used
  as the base path (B). The script uses the backmost path in the selection
  as B, then arranges the objects evenly along it.

  主な仕様 / Key features
  ----------------------------------------------------------------------
  ・開パス／クローズパスの両方に対応
    - クローズパスでは始点と終点の重複を避けて均等配置されます
  ・配置基準は各オブジェクトの geometricBounds の中心
  ・基準パス（B）の扱いを選択可能
      - 何もしない
      - 塗り／線をなしに
      - 削除
  ・配置後のオブジェクトをグループ化するオプションあり
  ・ダイアログは半透明表示で、画面右側にオフセット表示されます

  Supports both open and closed paths, optional grouping of placed objects,
  and configurable handling of the base path (keep, hide appearance, or delete).

  注意事項 / Notes
  ----------------------------------------------------------------------
  ・選択順は配置順として保証されません
  ・クローズパスの開始位置はパスの始点に依存します
  ・「塗り／線をなしに」「削除」は基準パス（B）に対して直接適用されます
    （元の外観は自動では復元されません）

  更新日 / Last updated: 2026-02-10
*/

// Script version / スクリプトバージョン
var SCRIPT_VERSION = "v1.0";

function getCurrentLang() {
  return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

/* 日英ラベル定義 / Japanese-English label definitions */
var LABELS = {
  dialogTitle: {
    ja: "パスに沿って配置",
    en: "Arrange Objects Along Path"
  },
  panelBasePath: {
    ja: "基準パス",
    en: "Base Path"
  },
  groupPlaced: {
    ja: "配置したオブジェクトをグループ化",
    en: "Group placed objects"
  },
  basePathModeNone: {
    ja: "何もしない",
    en: "Do nothing"
  },
  basePathModeHide: {
    ja: "塗り/線をなしに",
    en: "No fill / no stroke"
  },
  basePathModeDelete: {
    ja: "削除",
    en: "Delete"
  },
  btnCancel: {
    ja: "キャンセル",
    en: "Cancel"
  },
  btnOK: {
    ja: "OK",
    en: "OK"
  },
  alertNoDocument: {
    ja: "ドキュメントがありません。",
    en: "No document is open."
  },
  alertNeedSelection: {
    ja: "A（複数オブジェクト）とB（パス）を選択してください。基準パス（B）は最背面のパスです。",
    en: "Select A (objects) and B (a path). The base path (B) is the backmost path in the selection."
  },
  alertNoBasePath: {
    ja: "選択範囲に基準となるパス（B）が見つかりません。最背面にパスを含めて選択してください。",
    en: "No base path (B) was found. Include a path as the backmost item in the selection."
  },
  alertNoItems: {
    ja: "配置対象（A）が見つかりません。",
    en: "No placeable objects (A) were found."
  },
  alertPathTooShort: {
    ja: "パスが短すぎます。",
    en: "The path is too short."
  },
  alertPathAnalyzeFailed: {
    ja: "パスの解析に失敗しました。",
    en: "Failed to analyze the path."
  },
  alertPathLengthZero: {
    ja: "パス長が0です。",
    en: "The path length is zero."
  }
};

function L(key) {
  var entry = LABELS[key];
  if (!entry) return key;
  return entry[lang] || entry.en || key;
}

/* ダイアログ外観設定 / Dialog appearance settings */
var DIALOG_OFFSET_X = 300;
var DIALOG_OFFSET_Y = 0;
var DIALOG_OPACITY = 0.98;

function shiftDialogPosition(dlg, offsetX, offsetY) {
  dlg.onShow = function () {
    var currentX = dlg.location[0];
    var currentY = dlg.location[1];
    dlg.location = [currentX + offsetX, currentY + offsetY];
  };
}

function setDialogOpacity(dlg, opacityValue) {
  try {
    dlg.opacity = opacityValue;
  } catch (e) {
    // opacity not supported in some environments
  }
}

(function () {
  if (app.documents.length === 0) {
    alert(L('alertNoDocument'));
    return;
  }

  var doc = app.activeDocument;
  var sel = doc.selection;

  /* ダイアログ / Dialog */
  var dlg = new Window("dialog", L('dialogTitle') + ' ' + SCRIPT_VERSION);
  dlg.orientation = "column";
  dlg.alignChildren = "fill";

  // Apply dialog appearance / 外観適用
  setDialogOpacity(dlg, DIALOG_OPACITY);
  shiftDialogPosition(dlg, DIALOG_OFFSET_X, DIALOG_OFFSET_Y);

  // Output option / 出力オプション
  var cbGroupPlaced = dlg.add("checkbox", undefined, L('groupPlaced'));
  cbGroupPlaced.value = true; // default ON

  // Base path options / 基準パスオプション
  var optPanel = dlg.add("panel", undefined, L('panelBasePath'));
  optPanel.orientation = "column";
  optPanel.alignChildren = "left";
  optPanel.margins = [15, 20, 15, 10];

  /* 基準パスの扱い（排他） / Base path handling (exclusive) */
  var rbNone = optPanel.add("radiobutton", undefined, L('basePathModeNone'));
  var rbHide = optPanel.add("radiobutton", undefined, L('basePathModeHide'));
  var rbDelete = optPanel.add("radiobutton", undefined, L('basePathModeDelete'));
  rbHide.value = true; // default（旧「塗り/線をなしに」ON相当）

  // Buttons / ボタン
  var btns = dlg.add("group");
  btns.orientation = "row";
  btns.alignment = "right";
  btns.alignChildren = ["right", "center"];

  var btnCancel = btns.add("button", undefined, L('btnCancel'));
  var btnOK = btns.add("button", undefined, L('btnOK'), { name: "ok" });

  // OK / Cancel
  var dialogResult = dlg.show();
  if (dialogResult !== 1) {
    return; // Cancel
  }

  if (!sel || sel.length < 2) {
    alert(L('alertNeedSelection'));
    return;
  }

  /* B = 選択範囲の中で「最背面」のパス / B = backmost path in selection */
  var pathItem = getBackmostPathItem(sel);
  if (!pathItem) {
    alert(L('alertNoBasePath'));
    return;
  }

  /* 基準パスの扱い / Base path handling */
  if (rbHide.value) {
    try {
      pathItem.filled = false;
      pathItem.stroked = false;
    } catch (e) {
      // ignore
    }
  }

  /* A = 選択範囲のうち、基準パス（B）以外で配置可能なもの / A = placeable items excluding base path (B) */
  var items = [];
  for (var i = 0; i < sel.length; i++) {
    if (sel[i] === pathItem) continue;
    if (isPlaceableItem(sel[i])) items.push(sel[i]);
  }

  if (items.length === 0) {
    alert(L('alertNoItems'));
    return;
  }

  // Options / オプション
  var rotateAlongTangent = false; // trueにすると接線方向へ回転（簡易）
  var useEndpoints = true;        // true: 始点〜終点を含めて等分 / false: 端を避ける

  // PathPoint 配列（開パス/閉パス対応） / Path points (open/closed)
  var pts = pathItem.pathPoints;
  if (!pts || pts.length < 2) {
    alert(L('alertPathTooShort'));
    return;
  }

  /* ベジェ曲線を分割近似して折れ線化 / Convert Bezier to polyline by sampling */
  var samplesPerSegment = 30; // 精度（増やすほど重い）
  var poly = buildPolylineFromBezierPath(pathItem, samplesPerSegment); // [{x,y},...]
  if (poly.length < 2) {
    alert(L('alertPathAnalyzeFailed'));
    return;
  }

  var cum = [0];
  for (var p = 1; p < poly.length; p++) {
    cum[p] = cum[p - 1] + dist(poly[p - 1], poly[p]);
  }
  var totalLen = cum[cum.length - 1];
  if (totalLen <= 0) {
    alert(L('alertPathLengthZero'));
    return;
  }

  // 等間隔配置する距離リストを作る / Build distances for even placement
  // items.length 個の位置を作る / Create n positions
  var n = items.length;
  var ds = [];
  var isClosedPath = !!pathItem.closed;

  if (n === 1) {
    ds.push(totalLen / 2);
  } else {
    if (isClosedPath) {
      /* クローズパスは「始点=終点」なので、0〜totalLen を n 等分（終点は含めない） / Closed path excludes endpoint */
      for (var kc = 0; kc < n; kc++) {
        ds.push((totalLen * kc) / n);
      }
    } else {
      if (useEndpoints) {
        /* 0 〜 totalLen を (n-1) 分割 / Include endpoints */
        for (var k = 0; k < n; k++) {
          ds.push((totalLen * k) / (n - 1));
        }
      } else {
        /* 端を避けて 0〜totalLen の内側を n 分割（中心配置） / Avoid endpoints (centered) */
        for (var k2 = 0; k2 < n; k2++) {
          ds.push(totalLen * (k2 + 0.5) / n);
        }
      }
    }
  }

  /* 配置処理 / Place items */
  for (var j = 0; j < n; j++) {
    var d = ds[j];
    var pos = pointAtDistance(poly, cum, d);

    // 対象アイテムの中心 / Center of item
    var c = getItemCenter(items[j]);

    // 移動量 / Delta
    var dx = pos.x - c.x;
    var dy = pos.y - c.y;

    items[j].translate(dx, dy);

    if (rotateAlongTangent) {
      // 簡易：近傍点で接線角を推定 / Approx tangent angle
      var pos2 = pointAtDistance(poly, cum, Math.min(totalLen, d + totalLen * 0.001));
      var ang = Math.atan2(pos2.y - pos.y, pos2.x - pos.x) * 180 / Math.PI;

      // Illustratorの rotate は中心基準にする / Rotate around center
      items[j].rotate(ang, true, true, true, true, Transformation.CENTER);
    }
  }

  /* グループ化 / Group placed objects (exclude base path) */
  if (cbGroupPlaced.value) {
    try {
      var g = doc.groupItems.add();
      g.name = "ArrangedAlongPath";
      for (var gi = 0; gi < items.length; gi++) {
        items[gi].move(g, ElementPlacement.PLACEATEND);
      }
    } catch (eG) {
      // ignore
    }
  }

  /* 基準パス削除 / Remove base path */
  if (rbDelete.value) {
    try {
      pathItem.remove();
    } catch (eR) {
      // ignore
    }
  }

  /* ヘルパー / Helpers */

  function getBackmostPathItem(selectionArray) {
    var backmost = null;
    var minZ = null;

    for (var i = 0; i < selectionArray.length; i++) {
      var o = selectionArray[i];
      if (!isPathItem(o)) continue;

      var z = null;
      try {
        // 小さいほど背面（ドキュメント内のスタッキング順）
        z = o.zOrderPosition;
      } catch (_) {
        z = null;
      }

      if (z === null || z === undefined) {
        continue;
      }

      if (minZ === null || z < minZ) {
        minZ = z;
        backmost = o;
      }
    }

    // Fallback / フォールバック
    if (!backmost) {
      for (var j = 0; j < selectionArray.length; j++) {
        if (isPathItem(selectionArray[j])) return selectionArray[j];
      }
    }
    return backmost;
  }

  function isPathItem(o) {
    return o && o.typename === "PathItem" && o.pathPoints && o.pathPoints.length >= 2;
  }

  function isPlaceableItem(o) {
    if (!o) return false;
    var t = o.typename;
    return (
      t === "PathItem" ||
      t === "CompoundPathItem" ||
      t === "GroupItem" ||
      t === "TextFrame" ||
      t === "PlacedItem" ||
      t === "RasterItem" ||
      t === "SymbolItem" ||
      t === "MeshItem"
    );
  }

  function dist(a, b) {
    var dx = b.x - a.x;
    var dy = b.y - a.y;
    return Math.sqrt(dx * dx + dy * dy);
  }

  /* ベジェパスを折れ線近似に変換 / Convert Bezier path to polyline */
  function buildPolylineFromBezierPath(pItem, samplesPerSeg) {
    var pts = pItem.pathPoints;
    var closed = pItem.closed;

    var poly = [];
    var segCount = closed ? pts.length : (pts.length - 1);

    for (var i = 0; i < segCount; i++) {
      var p0 = pts[i];
      var p1 = pts[(i + 1) % pts.length];

      var a = { x: p0.anchor[0], y: p0.anchor[1] };
      var c1 = { x: p0.rightDirection[0], y: p0.rightDirection[1] };
      var c2 = { x: p1.leftDirection[0], y: p1.leftDirection[1] };
      var b = { x: p1.anchor[0], y: p1.anchor[1] };

      for (var s = 0; s <= samplesPerSeg; s++) {
        if (i > 0 && s === 0) continue; // 重複回避 / avoid duplicate
        var t = s / samplesPerSeg;
        var pt = cubicBezier(a, c1, c2, b, t);
        poly.push(pt);
      }
    }
    return poly;
  }

  function cubicBezier(a, c1, c2, b, t) {
    var mt = 1 - t;
    var mt2 = mt * mt;
    var t2 = t * t;

    var x =
      a.x * mt2 * mt +
      3 * c1.x * mt2 * t +
      3 * c2.x * mt * t2 +
      b.x * t2 * t;
    var y =
      a.y * mt2 * mt +
      3 * c1.y * mt2 * t +
      3 * c2.y * mt * t2 +
      b.y * t2 * t;

    return { x: x, y: y };
  }

  /* cumLength配列を使って距離dの位置を線形補間で求める / Find point at distance d */
  function pointAtDistance(poly, cum, d) {
    if (d <= 0) return { x: poly[0].x, y: poly[0].y };
    var total = cum[cum.length - 1];
    if (d >= total) return { x: poly[poly.length - 1].x, y: poly[poly.length - 1].y };

    // Linear search / 線形探索
    var idx = 1;
    while (idx < cum.length && cum[idx] < d) idx++;

    var d0 = cum[idx - 1];
    var d1 = cum[idx];
    var tt = (d - d0) / (d1 - d0);

    var p0 = poly[idx - 1];
    var p1 = poly[idx];

    return {
      x: p0.x + (p1.x - p0.x) * tt,
      y: p0.y + (p1.y - p0.y) * tt
    };
  }

  function getItemCenter(item) {
    // geometricBounds: [left, top, right, bottom]
    var b = item.geometricBounds;
    var left = b[0], top = b[1], right = b[2], bottom = b[3];
    return { x: (left + right) / 2, y: (top + bottom) / 2 };
  }

})();
