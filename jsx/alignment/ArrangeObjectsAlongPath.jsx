// ArrangeObjectsAlongPath.jsx (restored full script)
#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
  概要 / Overview
  ----------------------------------------------------------------------
  複数のオブジェクト（A）を、選択範囲内で「自動（面積最大）/ 最前面 / 最背面」で指定した1本のパス（B）に沿って
  等間隔に自動配置するIllustrator用スクリプトです。

  This script arranges multiple objects (A) at even intervals
  along a single base path (B), which is chosen by
  Auto (largest-area), Frontmost, or Backmost within the current selection.

  使い方 / How to use
  ----------------------------------------------------------------------
  1. 配置したいオブジェクトを複数選択します（A）
  2. 基準にしたいパスを含めて選択します（Bは「自動（面積最大）/ 最前面 / 最背面」で指定）
  3. スクリプトを実行し、ダイアログでオプションを指定します
  4. OKを押すと、オブジェクトがパスに沿って配置されます

  Select multiple objects to place (A), including one path to be used
  as the base path (B). The script uses the base path selected by Auto/Frontmost/Backmost
  as B, then arranges the objects evenly along it.

  主な仕様 / Key features
  ----------------------------------------------------------------------
  ・開パス／クローズパスの両方に対応
    - クローズパスでは始点と終点の重複を避けて均等配置されます
  ・配置基準は各オブジェクトの geometricBounds の中心
  ・基準パス（B）は「自動（面積最大）/ 最前面 / 最背面」で選択可能（PathItem）
  ・基準パス（B）の扱いを選択可能
      - 何もしない
      - 塗り／線をなしに
      - 削除
  ・配置後のオブジェクトをグループ化するオプションあり
  ・プレビュー（非破壊）：複製をプレビュー用レイヤーに表示（OK/キャンセル/OFFで自動削除）
  ・ダイアログは2カラム表示（回転は右カラム）。半透明表示で、画面右側にオフセット表示されます

  Supports both open and closed paths, optional grouping of placed objects,
  and configurable handling of the base path (keep, hide appearance, or delete).

  Also includes a non-destructive Preview that shows duplicated items on a
  temporary preview layer (auto-removed on OK/Cancel/Preview OFF).

  注意事項 / Notes
  ----------------------------------------------------------------------
  ・選択順は配置順として保証されません
  ・クローズパスの開始位置はパスの始点に依存します
  ・「塗り／線をなしに」「削除」は基準パス（B）に対して直接適用されます
    （元の外観は自動では復元されません）
  ・プレビューは複製で表示します（元オブジェクトは移動しません）
  ・プレビュー中は元のオブジェクトを一時的に非表示にします（OFF/OK/キャンセルで復元）

  更新日 / Last updated: 2026-02-24
*/

// Script version / スクリプトバージョン
var SCRIPT_VERSION = "v1.2";

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
    ja: "パス処理",
    en: "Base Path Handling"
  },
  preview: {
    ja: "プレビュー",
    en: "Preview"
  },
  panelPathOrder: {
    ja: "基準",
    en: "Base Path"
  },
  panelTargetPath: {
    ja: "対象パス",
    en: "Target Path"
  },
  panelPlaceObjects: {
    ja: "配置するオブジェクト",
    en: "Objects to Arrange"
  },
  panelOrder: {
    ja: "順番",
    en: "Order"
  },
  panelSpacing: {
    ja: "間隔",
    en: "Spacing"
  },
  spacingEven: {
    ja: "均等（現状）",
    en: "Even (Current)"
  },
  spacingRandom: {
    ja: "ランダム",
    en: "Random"
  },
  orderCurrent: {
    ja: "重ね順",
    en: "Current"
  },
  orderReverse: {
    ja: "逆順",
    en: "Reverse"
  },
  orderRandom: {
    ja: "ランダム",
    en: "Random"
  },
  panelRotation: {
    ja: "回転",
    en: "Rotation"
  },
  rotationNone: {
    ja: "0°",
    en: "Do nothing"
  },
  rotationAngle: {
    ja: "角度指定",
    en: "Angle"
  },
  rotationRandom: {
    ja: "ランダム",
    en: "Random"
  },
  autoLargest: {
    ja: "自動（面積最大）",
    en: "Auto (Largest)"
  },
  frontmost: {
    ja: "最前面",
    en: "Frontmost"
  },
  backmost: {
    ja: "最背面",
    en: "Backmost"
  },
  groupPlaced: {
    ja: "グループ化",
    en: "Group placed objects"
  },
  allRandom: {
    ja: "一括ランダム",
    en: "Random"
  },
  basePathModeNone: {
    ja: "何もしない",
    en: "Do nothing"
  },
  basePathModeHide: {
    ja: "「塗り／線」なし",
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
    ja: "A（複数オブジェクト）とB（基準パス）を選択してください。基準パス（B）は「自動（面積最大）/ 最前面 / 最背面」で指定できます。",
    en: "Select A (objects) and B (a path). Choose the base path by Auto (largest), Frontmost, or Backmost."
  },
  alertNoBasePath: {
    ja: "選択範囲に基準となるパス（B）が見つかりません。基準パスにしたいパス（PathItem）を含めて選択してください。",
    en: "No base path (B) was found. Include a PathItem to be used as the base path."
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

  // Options (kept as current defaults) / オプション（現状の既定値を維持）
  var rotateAlongTangent = false; // trueにすると接線方向へ回転（簡易）
  var useEndpoints = true;        // true: 始点〜終点を含めて等分 / false: 端を避ける
  var samplesPerSegment = 30;     // 精度（増やすほど重い）

  // Random spacing strength (ratio of step) / ランダム間隔の強さ（ステップに対する比率）
  // 0.1〜1.0 の範囲で調整
  var spacingRandomJitterRatio = 0.4;
  var _spacingJitterDragging = false;


  // Preview internals / プレビュー用内部
  var PREVIEW_LAYER_NAME = "__PREVIEW_ArrangeAlongPath";
  var previewLayer = null;

  // Selection snapshot for preview rebuild / プレビュー再構築用の選択スナップショット
  var previewSelSnapshot = null;

  // Remember last random order so OK matches the latest Preview / ランダム順を保持（プレビューとOKを一致）
  var lastRandomOrder = null;

  // Remember last random spacing so OK matches the latest Preview / ランダム間隔を保持（プレビューとOKを一致）
  var lastRandomSpacing = null; // { totalLen: number, ds: number[] }

  // Original visibility backup for preview / プレビュー中の元オブジェクト非表示（退避）
  var previewHiddenItems = [];
  var previewHiddenStates = [];

  function restoreHiddenItems() {
    if (!previewHiddenItems || previewHiddenItems.length === 0) return;
    for (var i = 0; i < previewHiddenItems.length; i++) {
      var it = previewHiddenItems[i];
      var st = previewHiddenStates[i];
      try { it.hidden = st; } catch (_) { }
    }
    previewHiddenItems = [];
    previewHiddenStates = [];
  }

  function hideOriginalItemsForPreview(baseItem, itemsArray) {
    // reset
    previewHiddenItems = [];
    previewHiddenStates = [];

    function pushItem(it) {
      if (!it) return;
      // avoid duplicates
      for (var k = 0; k < previewHiddenItems.length; k++) {
        if (previewHiddenItems[k] === it) return;
      }
      var prevHidden = false;
      try { prevHidden = !!it.hidden; } catch (_) { prevHidden = false; }
      previewHiddenItems.push(it);
      previewHiddenStates.push(prevHidden);
      try { it.hidden = true; } catch (_) { }
    }

    pushItem(baseItem);
    if (itemsArray && itemsArray.length) {
      for (var i = 0; i < itemsArray.length; i++) {
        pushItem(itemsArray[i]);
      }
    }
  }

  function clearPreview() {
    try {
      if (previewLayer) {
        try { previewLayer.locked = false; } catch (_) { }
        previewLayer.remove();
      }
    } catch (e) {
      // ignore
    }
    previewLayer = null;
    restoreHiddenItems();

    // Restore selection if it was dropped while preview hid originals / プレビューで選択が落ちた場合に復元
    try {
      var cur = doc.selection;
      if ((!cur || cur.length === 0) && previewSelSnapshot && previewSelSnapshot.length) {
        doc.selection = previewSelSnapshot;
      }
    } catch (_) { }

    try { app.redraw(); } catch (_) { }
  }

  function buildPreview(rbNone, rbHide, rbDelete, rbAutoLargest, rbFrontmost, rbBackmost) {
    // Always rebuild / 常に作り直す
    clearPreview();

    var curSel = doc.selection;
    if (!curSel || curSel.length < 2) {
      // When originals are hidden, Illustrator may drop the selection.
      // Use the last snapshot if available.
      if (previewSelSnapshot && previewSelSnapshot.length >= 2) {
        curSel = previewSelSnapshot;
      }
    }
    if (!curSel || curSel.length < 2) {
      alert(L('alertNeedSelection'));
      return false;
    }

    // Save snapshot for subsequent rebuilds while preview is ON
    previewSelSnapshot = [];
    for (var si = 0; si < curSel.length; si++) {
      previewSelSnapshot.push(curSel[si]);
    }

    var base = rbAutoLargest.value ? getLargestPathItem(curSel) : (rbFrontmost.value ? getFrontmostPathItem(curSel) : getBackmostPathItem(curSel));
    if (!base) {
      alert(L('alertNoBasePath'));
      return false;
    }

    var srcItems = [];
    for (var i = 0; i < curSel.length; i++) {
      if (curSel[i] === base) continue;
      if (isPlaceableItem(curSel[i])) srcItems.push(curSel[i]);
    }

    if (srcItems.length === 0) {
      alert(L('alertNoItems'));
      return false;
    }

    var orderMode = getOrderMode();
    var orderedSrcItems = applyOrderToArray(srcItems, orderMode, true);

    // Keep current selection (duplication may change selection) / 選択状態を保持
    var keepSelArr = [];
    try {
      for (var ks = 0; ks < curSel.length; ks++) keepSelArr.push(curSel[ks]);
    } catch (_) { }

    // Create preview layer on top / プレビュー用レイヤーを最前面に作成
    try {
      previewLayer = doc.layers.add();
      previewLayer.name = PREVIEW_LAYER_NAME;
    } catch (eL) {
      previewLayer = null;
      return false;
    }

    var gPrev = null;
    try {
      gPrev = previewLayer.groupItems.add();
      gPrev.name = "Preview_ArrangedAlongPath";
    } catch (eG) {
      try { previewLayer.remove(); } catch (_) { }
      previewLayer = null;
      return false;
    }

    // Duplicate base path for geometry / 形状計算用に基準パスを複製
    var prevPath = null;
    try {
      prevPath = base.duplicate(gPrev, ElementPlacement.PLACEATBEGINNING);
    } catch (eP) {
      try { previewLayer.remove(); } catch (_) { }
      previewLayer = null;
      return false;
    }

    // Apply base-path appearance mode to preview only / プレビュー側だけ外観モード反映
    if (rbHide.value) {
      try {
        prevPath.filled = false;
        prevPath.stroked = false;
      } catch (_) { }
    }

    // Duplicate items / 配置対象を複製
    var prevItems = [];
    for (var di = 0; di < orderedSrcItems.length; di++) {
      try {
        var dup = orderedSrcItems[di].duplicate(gPrev, ElementPlacement.PLACEATEND);
        prevItems.push(dup);
      } catch (_) {
        // skip
      }
    }

    if (prevItems.length === 0) {
      clearPreview();
      alert(L('alertNoItems'));
      return false;
    }

    // Arrange duplicates / 複製を配置
    var rot = getRotationSettings();
    var spacingMode = getSpacingMode();
    var ok = arrangeAlongPath(prevPath, prevItems, true, rot, spacingMode, true);
    if (!ok) {
      clearPreview();
      return false;
    }

    // If delete is selected, remove preview path after arranging / 削除選択時は配置後にパスを消す
    if (rbDelete.value) {
      try { prevPath.remove(); } catch (_) { }
    }

    // Hide originals while preview is shown / プレビュー表示中は元オブジェクトを隠す
    hideOriginalItemsForPreview(base, srcItems);

    // Restore selection / 選択を復元
    try { doc.selection = keepSelArr; } catch (_) { }
    try { app.redraw(); } catch (_) { }

    return true;
  }

  /* ダイアログ / Dialog */
  var dlg = new Window("dialog", L('dialogTitle') + ' ' + SCRIPT_VERSION);
  dlg.orientation = "column";
  dlg.alignChildren = "fill";

  // Apply dialog appearance / 外観適用
  setDialogOpacity(dlg, DIALOG_OPACITY);
  shiftDialogPosition(dlg, DIALOG_OFFSET_X, DIALOG_OFFSET_Y);

  // Two columns / 2カラム
  var cols = dlg.add("group");
  cols.orientation = "row";
  cols.alignChildren = ["fill", "top"];
  cols.spacing = 15;

  var colL = cols.add("group");
  colL.orientation = "column";
  colL.alignChildren = "fill";
  colL.alignment = "top";

  var colR = cols.add("group");
  colR.orientation = "column";
  colR.alignChildren = "fill";
  colR.alignment = "top";

  // Target path / 対象パス
  var targetPathPanel = colL.add("panel", undefined, L('panelTargetPath'));
  targetPathPanel.orientation = "column";
  targetPathPanel.alignChildren = "left";
  targetPathPanel.margins = [15, 20, 15, 15];

  // 対象 / Target
  var orderPanel = targetPathPanel.add("panel", undefined, L('panelPathOrder'));
  orderPanel.orientation = "column";
  orderPanel.alignChildren = "left";
  orderPanel.margins = [15, 20, 15, 10];

  var rbAutoLargest = orderPanel.add("radiobutton", undefined, L('autoLargest'));
  var rbFrontmost = orderPanel.add("radiobutton", undefined, L('frontmost'));
  var rbBackmost = orderPanel.add("radiobutton", undefined, L('backmost'));
  rbAutoLargest.value = true; // default: Auto (largest-area)

  // 処理 / Handling
  var optPanel = targetPathPanel.add("panel", undefined, L('panelBasePath'));
  optPanel.orientation = "column";
  optPanel.alignChildren = "left";
  optPanel.margins = [15, 20, 15, 10];

  /* 基準パスの扱い（排他） / Base path handling (exclusive) */
  var rbNone = optPanel.add("radiobutton", undefined, L('basePathModeNone'));
  var rbHide = optPanel.add("radiobutton", undefined, L('basePathModeHide'));
  var rbDelete = optPanel.add("radiobutton", undefined, L('basePathModeDelete'));
  rbHide.value = true; // default（旧「塗り/線をなしに」ON相当）




  // Objects to arrange / 配置するオブジェクト
  var placeObjPanel = colR.add("panel", undefined, L('panelPlaceObjects'));
  placeObjPanel.orientation = "column";
  placeObjPanel.alignChildren = "fill";
  placeObjPanel.margins = [15, 20, 15, 10];

  // Put all rotation radios under ONE parent group (for grouping & order) / 回転ラジオを同一groupに
  var rotPanel = placeObjPanel.add("panel", undefined, L('panelRotation'));
  rotPanel.orientation = "column";
  rotPanel.alignChildren = "left";
  rotPanel.margins = [15, 20, 15, 10];

  // Put all rotation radios under ONE parent group (for grouping & order) / 回転ラジオを同一groupに
  var gRotRadios = rotPanel.add("group");
  gRotRadios.orientation = "column";
  gRotRadios.alignChildren = "left";

  // Rotation modes / 回転モード（順番：何もしない → ランダム → 角度指定）
  var rbRotNone = gRotRadios.add("radiobutton", undefined, L('rotationNone'));
  var rbRotRandom = gRotRadios.add("radiobutton", undefined, L('rotationRandom'));

  // Angle (single line) / 角度指定（1行）
  var gRotAngleRow = gRotRadios.add("group");
  gRotAngleRow.orientation = "row";
  gRotAngleRow.alignChildren = ["left", "center"];
  gRotAngleRow.spacing = 6;

  var rbRotAngle = gRotAngleRow.add("radiobutton", undefined, L('rotationAngle'));
  var etRotAngle = gRotAngleRow.add("edittext", undefined, "0");
  etRotAngle.characters = 3;
  var stRotDeg = gRotAngleRow.add("statictext", undefined, "°");

  // Order / 順番
  var orderModePanel = placeObjPanel.add("panel", undefined, L('panelOrder'));
  orderModePanel.orientation = "row";
  orderModePanel.alignChildren = ["left", "center"];
  orderModePanel.margins = [15, 20, 15, 10];
  orderModePanel.spacing = 12;

  var rbOrderCurrent = orderModePanel.add("radiobutton", undefined, L('orderCurrent'));
  var rbOrderReverse = orderModePanel.add("radiobutton", undefined, L('orderReverse'));
  var rbOrderRandom = orderModePanel.add("radiobutton", undefined, L('orderRandom'));
  rbOrderCurrent.value = true; // default

  // Spacing / 間隔
  var spacingPanel = placeObjPanel.add("panel", undefined, L('panelSpacing'));
  spacingPanel.orientation = "column";
  spacingPanel.alignChildren = "fill";
  spacingPanel.margins = [15, 20, 15, 10];

  // Radios (horizontal) / ラジオ（横並び）
  var gSpacingRadios = spacingPanel.add("group");
  gSpacingRadios.orientation = "row";
  gSpacingRadios.alignChildren = ["left", "center"];
  gSpacingRadios.spacing = 12;

  var rbSpacingEven = gSpacingRadios.add("radiobutton", undefined, L('spacingEven'));
  var rbSpacingRandom = gSpacingRadios.add("radiobutton", undefined, L('spacingRandom'));
  rbSpacingEven.value = true; // default

  // Slider (enabled only when Random) / スライダー（ランダム選択時のみ有効）
  var gSpacingSld = spacingPanel.add("group");
  gSpacingSld.orientation = "row";
  gSpacingSld.alignChildren = ["left", "center"];

  // small indent under radios
  // var stSpIndent = gSpacingSld.add("statictext", undefined, "");
  // stSpIndent.preferredSize.width = 18;

  var sldSpacingJitter = gSpacingSld.add("slider", undefined, spacingRandomJitterRatio, 0.1, 1.0);
  sldSpacingJitter.preferredSize.width = 180;

  // Hidden label (kept for logic; no layout space) / 非表示ラベル（ロジック用・余白なし）
  var stSpacingJitterVal = gSpacingSld.add("statictext", undefined, "");
  stSpacingJitterVal.visible = false;
  stSpacingJitterVal.minimumSize.width = 0;
  stSpacingJitterVal.maximumSize.width = 0;
  stSpacingJitterVal.preferredSize.width = 0;

  sldSpacingJitter.enabled = false;
  stSpacingJitterVal.enabled = false;

  // Grouping / グループ化（右カラム・中央寄せ）
  var gGroupPlaced = placeObjPanel.add("group");
  gGroupPlaced.orientation = "row";
  gGroupPlaced.alignment = "fill";
  gGroupPlaced.alignChildren = ["center", "center"];

  // left spacer
  var stGrpSpL = gGroupPlaced.add("statictext", undefined, "");
  stGrpSpL.alignment = "fill";
  stGrpSpL.minimumSize.width = 10;
  stGrpSpL.maximumSize.width = 10000;

  // checkboxes (centered)
  var gGrpChecks = gGroupPlaced.add("group");
  gGrpChecks.orientation = "row";
  gGrpChecks.alignChildren = ["center", "center"];
  gGrpChecks.spacing = 12;

  var cbGroupPlaced = gGrpChecks.add("checkbox", undefined, L('groupPlaced'));
  cbGroupPlaced.value = true; // default ON

  var cbAllRandom = gGrpChecks.add("checkbox", undefined, L('allRandom'));
  cbAllRandom.value = false;

  // right spacer
  var stGrpSpR = gGroupPlaced.add("statictext", undefined, "");
  stGrpSpR.alignment = "fill";
  stGrpSpR.minimumSize.width = 10;
  stGrpSpR.maximumSize.width = 10000;

  // Bottom bar (3 columns) / 下部バー（3カラム）
  var bottomBar = dlg.add("group");
  bottomBar.orientation = "row";
  bottomBar.alignment = "fill";
  bottomBar.alignChildren = ["left", "center"];

  // Left: Preview / 左：プレビュー
  var gBottomL = bottomBar.add("group");
  gBottomL.orientation = "row";
  gBottomL.alignChildren = ["left", "center"];
  var cbPreview = gBottomL.add("checkbox", undefined, L('preview'));
  cbPreview.value = false;
  gBottomL.margins = [0, 0, 0, 0];

  // Middle: spacer / 中央：スペーサー
  var stSpacer = bottomBar.add("statictext", undefined, "");
  stSpacer.alignment = "fill";
  stSpacer.minimumSize.width = 10;
  stSpacer.maximumSize.width = 10000;
  stSpacer.preferredSize.width = 170;

  // Right: Buttons / 右：ボタン
  var btns = bottomBar.add("group");
  btns.orientation = "row";
  btns.alignment = "right";
  btns.alignChildren = ["right", "center"];

  var btnCancel = btns.add("button", undefined, L('btnCancel'), { name: "cancel" });
  var btnOK = btns.add("button", undefined, L('btnOK'), { name: "ok" });

  // Ensure Cancel always closes / キャンセルで必ず閉じる
  btnCancel.onClick = function () {
    try { dlg.close(0); } catch (_) { }
  };

  // Wire preview handlers / プレビュー連動
  function rebuildPreviewIfNeeded() {
    if (!cbPreview.value) return;
    var ok = buildPreview(rbNone, rbHide, rbDelete, rbAutoLargest, rbFrontmost, rbBackmost);
    if (!ok) cbPreview.value = false;
  }

  function setRotationMode(mode) {
    rbRotNone.value = (mode === "none");
    rbRotAngle.value = (mode === "angle");
    rbRotRandom.value = (mode === "random");
    updateRotationUI();
  }

  function updateRotationUI() {
    etRotAngle.enabled = !!rbRotAngle.value;
  }

  function getRotationSettings() {
    var mode = "none";
    try {
      if (rbRotAngle.value) mode = "angle";
      else if (rbRotRandom.value) mode = "random";
    } catch (_) { mode = "none"; }

    var ang = 0;
    try {
      ang = Number(etRotAngle.text);
      if (isNaN(ang)) ang = 0;
    } catch (_) { ang = 0; }

    return { mode: mode, angle: ang };
  }


  function getOrderMode() {
    try {
      if (rbOrderReverse.value) return "reverse";
      if (rbOrderRandom.value) return "random";
    } catch (_) { }
    return "current";
  }

  function getSpacingMode() {
    try {
      if (rbSpacingRandom.value) return "random";
    } catch (_) { }
    return "even";
  }

  function updateSpacingUI() {
    var en = !!rbSpacingRandom.value;
    sldSpacingJitter.enabled = en;
    stSpacingJitterVal.enabled = en;
  }

  function getZOrderPosSafe(it) {
    var z = null;
    try { z = it.zOrderPosition; } catch (_) { z = null; }
    if (z === null || z === undefined) return null;
    return z;
  }

  // “現状” = stacking order (frontmost → backmost) / 「現状」＝重ね順（最前面→最背面）
  function getCurrentOrderedArray(arr) {
    var deco = [];
    for (var i = 0; i < arr.length; i++) {
      var it = arr[i];
      var z = getZOrderPosSafe(it);
      deco.push({ it: it, z: z, hasZ: (z !== null), idx: i });
    }

    deco.sort(function (a, b) {
      if (a.hasZ && b.hasZ) {
        if (a.z === b.z) return a.idx - b.idx;
        return b.z - a.z; // larger = frontmost
      }
      if (a.hasZ) return -1;
      if (b.hasZ) return 1;
      return a.idx - b.idx;
    });

    var out = [];
    for (var j = 0; j < deco.length; j++) out.push(deco[j].it);
    return out;
  }

  function applyStoredOrder(baseArr, storedArr) {
    if (!storedArr || !storedArr.length) return null;

    var used = [];
    for (var i = 0; i < baseArr.length; i++) used[i] = false;

    var out = [];
    var matched = 0;

    for (var s = 0; s < storedArr.length; s++) {
      var target = storedArr[s];
      for (var i2 = 0; i2 < baseArr.length; i2++) {
        if (!used[i2] && baseArr[i2] === target) {
          used[i2] = true;
          out.push(baseArr[i2]);
          matched++;
          break;
        }
      }
    }

    // append leftovers
    for (var k = 0; k < baseArr.length; k++) {
      if (!used[k]) out.push(baseArr[k]);
    }

    // if nothing matched, treat as invalid (different selection)
    if (matched === 0) return null;

    return out;
  }

  function applyOrderToArray(arr, mode, isPreview) {
    var base = getCurrentOrderedArray(arr);

    if (mode === "reverse") {
      lastRandomOrder = null;
      base.reverse();
      return base;
    }

    if (mode !== "random") {
      lastRandomOrder = null;
      return base; // current
    }

    // Random:
    // If we already have a random order (from Preview), reuse it on OK so it matches.
    if (!isPreview) {
      var reused = applyStoredOrder(base, lastRandomOrder);
      if (reused) return reused;
    }

    // New shuffle (Preview builds the latest random order)
    var out = [];
    for (var i = 0; i < base.length; i++) out.push(base[i]);

    for (var j = out.length - 1; j > 0; j--) {
      var r = Math.floor(Math.random() * (j + 1));
      var tmp = out[j];
      out[j] = out[r];
      out[r] = tmp;
    }

    lastRandomOrder = [];
    for (var t = 0; t < out.length; t++) lastRandomOrder.push(out[t]);

    return out;
  }

  cbPreview.onClick = function () {
    if (cbPreview.value) {
      var ok = buildPreview(rbNone, rbHide, rbDelete, rbAutoLargest, rbFrontmost, rbBackmost);
      if (!ok) cbPreview.value = false;
    } else {
      clearPreview();
    }
  };

  rbNone.onClick = function () { rebuildPreviewIfNeeded(); };
  rbHide.onClick = function () { rebuildPreviewIfNeeded(); };
  rbDelete.onClick = function () { rebuildPreviewIfNeeded(); };
  rbAutoLargest.onClick = function () { rebuildPreviewIfNeeded(); };
  rbFrontmost.onClick = function () { rebuildPreviewIfNeeded(); };
  rbBackmost.onClick = function () { rebuildPreviewIfNeeded(); };

  function onRotModeChanged(mode) {
    setRotationMode(mode);

    // When switching to Angle mode, set default to 10°
    if (mode === "angle") {
      etRotAngle.text = "10";
    }

    // Focus angle field when Angle is selected so ↑↓ works immediately
    if (mode === "angle") {
      try { etRotAngle.active = true; } catch (_) { }
      try { etRotAngle.selection = [0, etRotAngle.text.length]; } catch (_) { }
    }

    rebuildPreviewIfNeeded();
  }

  rbRotNone.onClick = function () { onRotModeChanged("none"); };
  rbRotAngle.onClick = function () { onRotModeChanged("angle"); };
  rbRotRandom.onClick = function () { onRotModeChanged("random"); };

  rbOrderCurrent.onClick = function () { rebuildPreviewIfNeeded(); };
  rbOrderReverse.onClick = function () { rebuildPreviewIfNeeded(); };
  rbOrderRandom.onClick = function () { rebuildPreviewIfNeeded(); };

  rbSpacingEven.onClick = function () {
    lastRandomSpacing = null;
    updateSpacingUI();
    rebuildPreviewIfNeeded();
  };
  rbSpacingRandom.onClick = function () {
    updateSpacingUI();
    rebuildPreviewIfNeeded();
  };

  sldSpacingJitter.onChanging = function () {
    _spacingJitterDragging = true;

    // Update label/value, but DO NOT rebuild preview while dragging
    spacingRandomJitterRatio = Number(sldSpacingJitter.value);
    if (isNaN(spacingRandomJitterRatio)) spacingRandomJitterRatio = 0.4;
    if (spacingRandomJitterRatio < 0.1) spacingRandomJitterRatio = 0.1;
    if (spacingRandomJitterRatio > 1.0) spacingRandomJitterRatio = 1.0;
    if (stSpacingJitterVal) stSpacingJitterVal.text = spacingRandomJitterRatio.toFixed(2);
  };

  sldSpacingJitter.onChange = function () {
    _spacingJitterDragging = false;

    // Apply final value and rebuild preview on mouse release
    spacingRandomJitterRatio = Number(sldSpacingJitter.value);
    if (isNaN(spacingRandomJitterRatio)) spacingRandomJitterRatio = 0.4;
    if (spacingRandomJitterRatio < 0.1) spacingRandomJitterRatio = 0.1;
    if (spacingRandomJitterRatio > 1.0) spacingRandomJitterRatio = 1.0;
    if (stSpacingJitterVal) stSpacingJitterVal.text = spacingRandomJitterRatio.toFixed(2);

    // Force regenerate random spacing with new strength
    lastRandomSpacing = null;
    rebuildPreviewIfNeeded();
  };

  etRotAngle.onChange = function () {
    // Editing angle implies Angle mode
    setRotationMode("angle");
    rebuildPreviewIfNeeded();
  };

  // One-shot: set Rotation/Order/Spacing to Random when checked / チェックで一括ランダム
  cbAllRandom.onClick = function () {
    if (cbAllRandom.value) {
      // ON: set all to Random
      setRotationMode("random");
      rbOrderRandom.value = true;
      rbSpacingRandom.value = true;

      // Refresh UI enable states
      updateRotationUI();
      updateSpacingUI();

      // Force regenerate random results
      lastRandomOrder = null;
      lastRandomSpacing = null;

      rebuildPreviewIfNeeded();
      return;
    }

    // OFF: reset to defaults
    setRotationMode("none");
    rbOrderCurrent.value = true;
    rbSpacingEven.value = true;

    // Refresh UI enable states
    updateRotationUI();
    updateSpacingUI();

    // Clear random caches
    lastRandomOrder = null;
    lastRandomSpacing = null;

    rebuildPreviewIfNeeded();
  };

  updateRotationUI();
  updateSpacingUI();

  // OK / Cancel
  var dialogResult = dlg.show();

  // Cleanup preview on close (OK/Cancel) / ダイアログ終了時にプレビュー掃除
  clearPreview();

  if (dialogResult !== 1) {
    previewSelSnapshot = null;
    lastRandomOrder = null;
    lastRandomSpacing = null;
    return; // Cancel
  }

  // Re-fetch selection at OK time / OK時に選択を取り直す
  sel = doc.selection;

  if (!sel || sel.length < 2) {
    // Fallback to snapshot if selection was dropped during preview
    if (previewSelSnapshot && previewSelSnapshot.length >= 2) {
      try { doc.selection = previewSelSnapshot; } catch (_) { }
      try { sel = doc.selection; } catch (_) { }
    }
  }

  if (!sel || sel.length < 2) {
    alert(L('alertNeedSelection'));
    previewSelSnapshot = null;
    return;
  }

  /* B = 選択範囲の中で選択されたルールに従うパス / B = base path by selected rule */
  var pathItem = rbAutoLargest.value ? getLargestPathItem(sel) : (rbFrontmost.value ? getFrontmostPathItem(sel) : getBackmostPathItem(sel));
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
  for (var i2 = 0; i2 < sel.length; i2++) {
    if (sel[i2] === pathItem) continue;
    if (isPlaceableItem(sel[i2])) items.push(sel[i2]);
  }

  if (items.length === 0) {
    alert(L('alertNoItems'));
    return;
  }

  items = applyOrderToArray(items, getOrderMode(), false);

  // Arrange actual items / 実体を配置
  var rotFinal = getRotationSettings();
  var spacingMode = getSpacingMode();
  var okActual = arrangeAlongPath(pathItem, items, true, rotFinal, spacingMode, false);
  if (!okActual) {
    return;
  }

  /* グループ化 / Group placed objects (exclude base path) */
  if (cbGroupPlaced.value) {
    try {
      var g = doc.groupItems.add();
      g.name = "ArrangedAlongPath";
      for (var gi = 0; gi < items.length; gi++) {
        items[gi].move(g, ElementPlacement.PLACEATEND);
      }
    } catch (eG2) {
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

  previewSelSnapshot = null;
  lastRandomOrder = null;
  lastRandomSpacing = null;

  /* ヘルパー / Helpers */

  function arrangeAlongPath(pathItem, itemsArray, showAlerts, rot, spacingMode, isPreview) {
    var pts = pathItem.pathPoints;
    if (!pts || pts.length < 2) {
      if (showAlerts) alert(L('alertPathTooShort'));
      return false;
    }

    var poly = buildPolylineFromBezierPath(pathItem, samplesPerSegment); // [{x,y},...]
    if (poly.length < 2) {
      if (showAlerts) alert(L('alertPathAnalyzeFailed'));
      return false;
    }

    var cum = [0];
    for (var p = 1; p < poly.length; p++) {
      cum[p] = cum[p - 1] + dist(poly[p - 1], poly[p]);
    }
    var totalLen = cum[cum.length - 1];
    if (totalLen <= 0) {
      if (showAlerts) alert(L('alertPathLengthZero'));
      return false;
    }

    var n = itemsArray.length;
    var ds = [];
    var isClosedPath = !!pathItem.closed;

    // Random spacing / ランダム間隔
    if (spacingMode === "random") {
      // Reuse the last previewed spacing so OK matches Preview
      if (!isPreview && lastRandomSpacing && lastRandomSpacing.ds && lastRandomSpacing.ds.length === n) {
        var okReuse = false;
        try {
          var diff = Math.abs(Number(lastRandomSpacing.totalLen) - Number(totalLen));
          okReuse = (diff < 0.01);
        } catch (_) { okReuse = false; }

        if (okReuse) {
          ds = [];
          for (var rr = 0; rr < lastRandomSpacing.ds.length; rr++) ds.push(lastRandomSpacing.ds[rr]);
        }
      }

      // Generate new if not reusable
      if (ds.length !== n) {
        ds = [];

        // Range limits
        var minD = 0;
        var maxD = totalLen;
        if (!isClosedPath && !useEndpoints) {
          var pad = totalLen * 0.02;
          if (pad > 0) {
            minD = pad;
            maxD = Math.max(pad, totalLen - pad);
          }
        }

        // Start from even spacing, then add small jitter (weaker random)
        var baseDs = [];
        if (n === 1) {
          baseDs.push(totalLen / 2);
        } else {
          if (isClosedPath) {
            for (var kc2 = 0; kc2 < n; kc2++) {
              baseDs.push((totalLen * kc2) / n);
            }
          } else {
            if (useEndpoints) {
              for (var k3 = 0; k3 < n; k3++) {
                baseDs.push((totalLen * k3) / (n - 1));
              }
            } else {
              for (var k4 = 0; k4 < n; k4++) {
                baseDs.push(totalLen * (k4 + 0.5) / n);
              }
            }
          }
        }

        // Estimate step for jitter amplitude
        var step = (n <= 1) ? totalLen : (isClosedPath ? (totalLen / n) : (useEndpoints ? (totalLen / (n - 1)) : (totalLen / n)));
        var ratio = spacingRandomJitterRatio;
        if (isNaN(ratio)) ratio = 0.4;
        if (ratio < 0.1) ratio = 0.1;
        if (ratio > 1.0) ratio = 1.0;
        var jitter = step * ratio;

        for (var iR = 0; iR < n; iR++) {
          var v = baseDs[iR];

          // Keep endpoints stable when open-path & endpoints mode
          if (!isClosedPath && useEndpoints && n > 1 && (iR === 0 || iR === n - 1)) {
            // keep v
          } else {
            v += (Math.random() * 2 - 1) * jitter;
          }

          // clamp into range
          if (v < minD) v = minD;
          if (v > maxD) v = maxD;
          ds.push(v);
        }

        // sort => random gaps but monotonic along the path
        ds.sort(function (a, b) { return a - b; });

        // store latest random spacing
        lastRandomSpacing = { totalLen: totalLen, ds: [] };
        for (var ss = 0; ss < ds.length; ss++) lastRandomSpacing.ds.push(ds[ss]);
      }

    } else {
      // Even spacing (current behavior) / 均等（現状）
      lastRandomSpacing = null;

      if (n === 1) {
        ds.push(totalLen / 2);
      } else {
        if (isClosedPath) {
          for (var kc = 0; kc < n; kc++) {
            ds.push((totalLen * kc) / n);
          }
        } else {
          if (useEndpoints) {
            for (var k = 0; k < n; k++) {
              ds.push((totalLen * k) / (n - 1));
            }
          } else {
            for (var k2 = 0; k2 < n; k2++) {
              ds.push(totalLen * (k2 + 0.5) / n);
            }
          }
        }
      }
    }

    for (var j = 0; j < n; j++) {
      var d = ds[j];
      var pos = pointAtDistance(poly, cum, d);

      var c = getItemCenter(itemsArray[j]);
      var dx = pos.x - c.x;
      var dy = pos.y - c.y;

      itemsArray[j].translate(dx, dy);

      // Rotation / 回転
      if (rot && rot.mode && rot.mode !== "none") {
        var angDeg = 0;
        if (rot.mode === "angle") {
          angDeg = rot.angle || 0;
        } else if (rot.mode === "random") {
          angDeg = (Math.random() * 360) - 180;
        }
        try {
          itemsArray[j].rotate(angDeg, true, true, true, true, Transformation.CENTER);
        } catch (_) { }
      }

      if (rotateAlongTangent) {
        var pos2 = pointAtDistance(poly, cum, Math.min(totalLen, d + totalLen * 0.001));
        var ang = Math.atan2(pos2.y - pos.y, pos2.x - pos.x) * 180 / Math.PI;
        itemsArray[j].rotate(ang, true, true, true, true, Transformation.CENTER);
      }
    }
    return true;
  }

  function getLargestPathItem(selectionArray) {
    var best = null;
    var bestArea = null;

    for (var i = 0; i < selectionArray.length; i++) {
      var o = selectionArray[i];
      if (!isPathItem(o)) continue;

      var a = null;
      try {
        // PathItem.area may be 0 for open paths, and may be negative depending on direction
        a = Math.abs(o.area);
      } catch (_) {
        a = null;
      }

      if (a === null || a === undefined || isNaN(a) || a <= 0) {
        a = getBoundsArea(o);
      }

      if (bestArea === null || a > bestArea) {
        bestArea = a;
        best = o;
      }
    }

    if (!best) {
      best = getBackmostPathItem(selectionArray);
    }
    return best;
  }

  function getBoundsArea(item) {
    var b = null;
    try { b = item.geometricBounds; } catch (_) { b = null; }
    if (!b || b.length < 4) return 0;
    var w = Math.abs(b[2] - b[0]);
    var h = Math.abs(b[1] - b[3]);
    return w * h;
  }

  function getBackmostPathItem(selectionArray) {
    var backmost = null;
    var minZ = null;

    for (var i = 0; i < selectionArray.length; i++) {
      var o = selectionArray[i];
      if (!isPathItem(o)) continue;

      var z = null;
      try { z = o.zOrderPosition; } catch (_) { z = null; }
      if (z === null || z === undefined) continue;

      if (minZ === null || z < minZ) {
        minZ = z;
        backmost = o;
      }
    }

    if (!backmost) {
      for (var j = 0; j < selectionArray.length; j++) {
        if (isPathItem(selectionArray[j])) return selectionArray[j];
      }
    }
    return backmost;
  }

  function getFrontmostPathItem(selectionArray) {
    var frontmost = null;
    var maxZ = null;

    for (var i = 0; i < selectionArray.length; i++) {
      var o = selectionArray[i];
      if (!isPathItem(o)) continue;

      var z = null;
      try { z = o.zOrderPosition; } catch (_) { z = null; }
      if (z === null || z === undefined) continue;

      if (maxZ === null || z > maxZ) {
        maxZ = z;
        frontmost = o;
      }
    }

    // Fallback
    if (!frontmost) {
      for (var j = selectionArray.length - 1; j >= 0; j--) {
        if (isPathItem(selectionArray[j])) return selectionArray[j];
      }
    }
    return frontmost;
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
        if (i > 0 && s === 0) continue;
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

  function pointAtDistance(poly, cum, d) {
    if (d <= 0) return { x: poly[0].x, y: poly[0].y };
    var total = cum[cum.length - 1];
    if (d >= total) return { x: poly[poly.length - 1].x, y: poly[poly.length - 1].y };

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
    var b = item.geometricBounds; // [left, top, right, bottom]
    var left = b[0], top = b[1], right = b[2], bottom = b[3];
    return { x: (left + right) / 2, y: (top + bottom) / 2 };
  }

})();