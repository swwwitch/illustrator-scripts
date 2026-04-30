#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
  互い違いに選択 / Alternate Select
  - 選択中のオブジェクトを「奇数 / 偶数」で互い違いに選択
  - 方向（垂直 / 水平）でカウント順を切り替え
*/

/* バージョン / Version */
var SCRIPT_VERSION = "v1.1.0";

function getCurrentLang() {
  return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var currentLanguage = getCurrentLang();

/* UIラベル定義 / UI label definitions */
var UI_LABELS = {
  dialogTitle: {
    ja: "互い違いに選択",
    en: "Alternate Select"
  },
  panelSelect: {
    ja: "選択",
    en: "Selection"
  },
  odd: {
    ja: "奇数",
    en: "Odd"
  },
  even: {
    ja: "偶数",
    en: "Even"
  },
  panelDirection: {
    ja: "方向",
    en: "Direction"
  },
  vertical: {
    ja: "垂直",
    en: "Vertical"
  },
  horizontal: {
    ja: "水平",
    en: "Horizontal"
  },
  zOrder: {
    ja: "重ね順",
    en: "Z-order"
  },
  ok: {
    ja: "OK",
    en: "OK"
  },
  cancel: {
    ja: "キャンセル",
    en: "Cancel"
  }
};

/* エラー／ログ文言定義 / Error and log message definitions */
var ERROR_LABELS = {
  noDocument: {
    ja: "ドキュメントを開いてください。",
    en: "Please open a document."
  },
  noSelection: {
    ja: "オブジェクトを選択してください。",
    en: "Please select objects."
  },
  noValidItems: {
    ja: "有効なオブジェクトが選択されていません。",
    en: "No valid objects are selected."
  },
  preview: {
    ja: "プレビューエラー：",
    en: "Preview Error: "
  },
  prefix: {
    ja: "エラー：",
    en: "Error: "
  },
  debugPrefix: {
    ja: "デバッグ：",
    en: "Debug: "
  }
};

function getLocalizedText(table, key) {
  if (!table[key]) return key;
  return table[key][currentLanguage] || table[key].en || key;
}

function getUILabel(key) {
  return getLocalizedText(UI_LABELS, key);
}

function getErrorLabel(key) {
  return getLocalizedText(ERROR_LABELS, key);
}

function getErrorMessage(key, detail) {
  var message = getErrorLabel(key);
  if (detail !== undefined && detail !== null && String(detail) !== "") {
    message += String(detail);
  }
  return message;
}

function showError(key, detail) {
  alert(getErrorMessage(key, detail));
}

function throwLocalizedError(key, detail) {
  throw new Error(getErrorMessage(key, detail));
}

function debugLog(key, detail) {
  if (typeof $.writeln !== "function") return;
  $.writeln(getErrorLabel("debugPrefix") + getErrorMessage(key, detail));
}

/* プレビュー管理 / Preview manager
   - 選択のプレビューは通常 Undo 履歴に乗らないため、選択はスナップショット復元
   - ドキュメント変更を伴う場合のみ app.undo() を使って巻き戻せるようにする
*/
function PreviewManager() {
  this.undoDepth = 0; // Undoable step count during preview
  this.selectionSnapshot = null; // Original selection snapshot

  /* 選択状態を退避 / Capture selection */
  this.captureSelection = function (doc) {
    var snap = [];
    try {
      var currentSelection = doc.selection;
      if (currentSelection && currentSelection.length) {
        for (var selectionIndex = 0; selectionIndex < currentSelection.length; selectionIndex++) {
          snap.push(currentSelection[selectionIndex]);
        }
      }
    } catch (e) { }
    this.selectionSnapshot = snap;
  };

  /* 選択状態を復元 / Restore selection */
  this.restoreSelection = function (doc) {
    try {
      doc.selection = null;
      if (this.selectionSnapshot && this.selectionSnapshot.length) {
        for (var snapshotIndex = 0; snapshotIndex < this.selectionSnapshot.length; snapshotIndex++) {
          try { this.selectionSnapshot[snapshotIndex].selected = true; } catch (e) { }
        }
      }
      app.redraw();
    } catch (e) { }
  };

  /**
   * 変更操作を実行し、必要なら履歴としてカウント / Run step and optionally count as undoable
   * @param {Function} func - 実行処理 / action
   * @param {Boolean} [undoable=false] - Undo 対象なら true / true if it creates undo history
   */
  this.addStep = function (func, undoable) {
    try {
      func();
      if (undoable) this.undoDepth++;
      app.redraw();
    } catch (e) {
      showError("preview", e);
    }
  };

  /* プレビュー分の変更を巻き戻し / Rollback preview changes */
  this.rollback = function (doc) {
    // Undoable steps rollback
    while (this.undoDepth > 0) {
      try { app.undo(); } catch (e) { break; }
      this.undoDepth--;
    }
    // Selection rollback (always)
    this.restoreSelection(doc);
  };

  /**
   * 確定 / Confirm
   * @param {Document} doc
   * @param {Function} [finalAction] - 一度戻してから本番処理 / optional final action
   */
  this.confirm = function (doc, finalAction) {
    if (finalAction) {
      this.rollback(doc);
      finalAction();
      this.undoDepth = 0;
      this.captureSelection(doc); // keep current as baseline
    } else {
      this.undoDepth = 0;
      this.captureSelection(doc); // keep current as baseline
    }
  };
}

/* 方向推定の設定 / Direction detection settings */
var DIRECTION_THRESHOLD = 1.20; // しきい値（1.20 = 20%差） / Threshold ratio
var PREF_KEY_LAST_DIR = "AlternateSelect.LastDirectionMode";

/* 前回の方向モードを取得 / Load last direction mode (custom options)
   - 取得できない場合は null を返す
   - mode: "vertical" | "horizontal" | "zorder"
*/
function loadLastDirectionMode() {
  try {
    var desc = app.getCustomOptions(PREF_KEY_LAST_DIR);
    if (desc && desc.hasKey(stringIDToTypeID("mode"))) {
      return desc.getString(stringIDToTypeID("mode"));
    }
  } catch (e) { }
  return null;
}

/* 前回の方向モードを保存 / Save last direction mode (custom options) */
function saveLastDirectionMode(mode) {
  try {
    var m = String(mode);
    if (m !== "vertical" && m !== "horizontal" && m !== "zorder") return;
    var desc = new ActionDescriptor();
    desc.putString(stringIDToTypeID("mode"), m);
    app.putCustomOptions(PREF_KEY_LAST_DIR, desc, true);
  } catch (e) { }
}

/* 方向の自動推定 / Auto detect direction (initial)
   - 中心点のX/Yレンジで判定（外れ値耐性あり）
   - rangeX が rangeY の DIRECTION_THRESHOLD 倍より大きい → "horizontal"
   - rangeY が rangeX の DIRECTION_THRESHOLD 倍より大きい → "vertical"
   - それ以外（僅差/グリッド等）は「前回mode優先（fallback）」を返す
*/
function guessInitialDirectionMode(items, fallbackMode) {
  var fallbackDirectionMode = (fallbackMode !== undefined && fallbackMode !== null) ? String(fallbackMode) : "vertical";
  if (fallbackDirectionMode !== "vertical" && fallbackDirectionMode !== "horizontal" && fallbackDirectionMode !== "zorder") fallbackDirectionMode = "vertical";

  if (!items || items.length < 2) return fallbackDirectionMode;

  var xs = [];
  var ys = [];

  for (var i = 0; i < items.length; i++) {
    var candidateItem = items[i];
    if (!candidateItem) continue;

    // geometricBounds: [left, top, right, bottom]
    var geometricBounds;
    try { geometricBounds = candidateItem.geometricBounds; } catch (e) { continue; }
    if (!geometricBounds || geometricBounds.length < 4) continue;

    var centerX = (geometricBounds[0] + geometricBounds[2]) / 2;
    var centerY = (geometricBounds[1] + geometricBounds[3]) / 2;
    xs.push(centerX);
    ys.push(centerY);
  }

  if (xs.length < 2 || ys.length < 2) return fallbackDirectionMode;

  xs.sort(function (firstItem, secondItem) { return firstItem - secondItem; });
  ys.sort(function (firstItem, secondItem) { return firstItem - secondItem; });

  // 外れ値耐性：両端を少し落としてレンジを取る（nが大きいほど効果）
  var n = xs.length;
  var trim = 0;
  if (n >= 10) trim = Math.floor(n * 0.10); // 10% trimming
  else if (n >= 6) trim = 1;

  var minX = xs[trim];
  var maxX = xs[n - 1 - trim];
  var minY = ys[trim];
  var maxY = ys[n - 1 - trim];

  var rangeX = maxX - minX;
  var rangeY = maxY - minY;

  // しきい値判定（明確な差があるときだけ自動切替）
  if (rangeX > rangeY * DIRECTION_THRESHOLD) return "horizontal";
  if (rangeY > rangeX * DIRECTION_THRESHOLD) return "vertical";

  // 僅差（グリッド等）は前回mode優先
  return fallbackDirectionMode;
}

(function () {
  /* ドキュメント確認 / Check document */
  if (app.documents.length === 0) {
    showError("noDocument");
    return;
  }

  var doc = app.activeDocument;

  /* プレビューマネージャ / Preview manager */
  var previewMgr = new PreviewManager();
  previewMgr.captureSelection(doc);

  var originalSelection = doc.selection;

  /* 選択確認 / Check selection */
  if (!originalSelection || originalSelection.length === 0) {
    showError("noSelection");
    return;
  }

  /* 有効なオブジェクト抽出 / Collect valid items */
  var items = [];
  for (var selectionIndex = 0; selectionIndex < originalSelection.length; selectionIndex++) {
    var selectedItem = originalSelection[selectionIndex];
    if (!selectedItem.locked && !selectedItem.hidden) {
      items.push(selectedItem);
    }
  }

  if (items.length === 0) {
    showError("noValidItems");
    return;
  }

  function applySelectionPreview() {
    // まず前回プレビューを巻き戻す（選択はスナップショットで復元）
    previewMgr.rollback(doc);

    // 今回のプレビューを適用（選択変更は通常 Undoable ではないので undoable=false）
    previewMgr.addStep(function () {
      var selectOdd = oddRadio.value;
      var mode = verticalRadio.value ? "vertical" : (horizontalRadio.value ? "horizontal" : "zorder");

      /* 並び順ソート / Sort order */
      if (mode === "vertical") {
        /* 垂直：Y降順（上→下） / Vertical: Y desc (top to bottom) */
        items.sort(function (firstItem, secondItem) { return secondItem.position[1] - firstItem.position[1]; });
      } else if (mode === "horizontal") {
        /* 水平：X昇順（左→右） / Horizontal: X asc (left to right) */
        items.sort(function (firstItem, secondItem) { return firstItem.position[0] - secondItem.position[0]; });
      } else {
        /* 重ね順：zOrderPosition 昇順 / Z-order: zOrderPosition asc */
        items.sort(function (firstItem, secondItem) {
          var firstZOrderPosition = 0, secondZOrderPosition = 0;
          try { firstZOrderPosition = firstItem.zOrderPosition; } catch (e) { firstZOrderPosition = 0; }
          try { secondZOrderPosition = secondItem.zOrderPosition; } catch (e) { secondZOrderPosition = 0; }
          return firstZOrderPosition - secondZOrderPosition;
        });
      }

      /* 互い違い選択 / Alternate selection */
      doc.selection = null;
      for (var itemIndex = 0; itemIndex < items.length; itemIndex++) {
        var isOddIndex = (itemIndex % 2 === 0); // 0,2,4... => 1,3,5...
        if ((selectOdd && isOddIndex) || (!selectOdd && !isOddIndex)) {
          items[itemIndex].selected = true;
        }
      }
    }, false);
  }


  /* キー入力でラジオ切替 / Key handler for radio buttons
     Odd: O / Even: E / Vertical: V / Horizontal: H / Z-Order: A
  */
  function addRadioKeyHandler(dialog) {
    dialog.addEventListener("keydown", function (event) {
      var pressedKeyName = event.keyName;

      // 奇数 / 偶数
      if (pressedKeyName === "O") {
        setAlternateSelectionMode("odd");
        event.preventDefault();
        return;
      } else if (pressedKeyName === "E") {
        setAlternateSelectionMode("even");
        event.preventDefault();
        return;
      }

      // 方向（垂直 / 水平 / 重ね順）
      if (pressedKeyName === "V") {
        setDirectionMode("vertical");
        event.preventDefault();
        return;
      } else if (pressedKeyName === "H") {
        setDirectionMode("horizontal");
        event.preventDefault();
        return;
      } else if (pressedKeyName === "A") {
        setDirectionMode("zorder");
        event.preventDefault();
        return;
      }
    });
  }

  /* ダイアログボックス / Dialog box */
  var dialog = new Window("dialog", getUILabel("dialogTitle") + " " + SCRIPT_VERSION);
  dialog.orientation = "column";
  dialog.alignChildren = "left";

  var selectionPanel = dialog.add("panel", undefined, getUILabel("panelSelect"));
  selectionPanel.margins = [15, 20, 15, 10];
  selectionPanel.orientation = "column";
  selectionPanel.alignChildren = "left";

  var selectGroup = selectionPanel.add("group");
  selectGroup.orientation = "row";
  selectGroup.alignChildren = "left";

  var oddRadio = selectGroup.add("radiobutton", undefined, getUILabel("odd"));
  var evenRadio = selectGroup.add("radiobutton", undefined, getUILabel("even"));
  oddRadio.value = true; // デフォルトは奇数 / Default is odd

  function setAlternateSelectionMode(selectionMode) {
    oddRadio.value = (selectionMode === "odd");
    evenRadio.value = (selectionMode === "even");
    applySelectionPreview();
  }

  /* 方向パネル / Direction panel */
  var dirPanel = dialog.add("panel", undefined, getUILabel("panelDirection"));
  dirPanel.margins = [15, 20, 15, 10];
  dirPanel.orientation = "column";
  dirPanel.alignChildren = "left";

  var dirGroup = dirPanel.add("group");
  dirGroup.orientation = "row";
  dirGroup.alignChildren = "left";

  var verticalRadio = dirGroup.add("radiobutton", undefined, getUILabel("vertical"));
  var horizontalRadio = dirGroup.add("radiobutton", undefined, getUILabel("horizontal"));
  var zOrderRadio = dirGroup.add("radiobutton", undefined, getUILabel("zOrder"));

  function setDirectionMode(directionMode) {
    verticalRadio.value = (directionMode === "vertical");
    horizontalRadio.value = (directionMode === "horizontal");
    zOrderRadio.value = (directionMode === "zorder");
    saveLastDirectionMode(directionMode);
    applySelectionPreview();
  }
  // 前回値を基本にし、差が明確なときだけ自動切替 / Prefer last value; auto-switch only if clear
  var lastMode = loadLastDirectionMode();
  if (lastMode === null) lastMode = "vertical";

  var initialMode = guessInitialDirectionMode(items, lastMode);
  verticalRadio.value = (initialMode === "vertical");
  horizontalRadio.value = (initialMode === "horizontal");
  zOrderRadio.value = (initialMode === "zorder");

  addRadioKeyHandler(dialog);

  var buttonGroup = dialog.add("group");
  buttonGroup.alignment = "right";
  var cancelBtn = buttonGroup.add("button", undefined, getUILabel("cancel"));
  var okBtn = buttonGroup.add("button", undefined, getUILabel("ok"));

  oddRadio.onClick = function () {
    setAlternateSelectionMode("odd");
  };
  evenRadio.onClick = function () {
    setAlternateSelectionMode("even");
  };
  verticalRadio.onClick = function () {
    setDirectionMode("vertical");
  };
  horizontalRadio.onClick = function () {
    setDirectionMode("horizontal");
  };
  zOrderRadio.onClick = function () {
    setDirectionMode("zorder");
  };

  /* ボタン動作 / Button handlers */
  cancelBtn.onClick = function () {
    // キャンセル：プレビューを巻き戻して閉じる / Cancel: rollback and close
    previewMgr.rollback(doc);
    dialog.close(0);
  };

  okBtn.onClick = function () {
    // OK：Undo を綺麗にするため「一度戻して再実行」 / OK: rollback then re-apply once
    previewMgr.confirm(doc, function () {
      // 最終適用（1回） / Final apply (single step)
      var selectOdd = oddRadio.value;
      var mode = verticalRadio.value ? "vertical" : (horizontalRadio.value ? "horizontal" : "zorder");

      if (mode === "vertical") {
        items.sort(function (firstItem, secondItem) { return secondItem.position[1] - firstItem.position[1]; });
      } else if (mode === "horizontal") {
        items.sort(function (firstItem, secondItem) { return firstItem.position[0] - secondItem.position[0]; });
      } else {
        items.sort(function (firstItem, secondItem) {
          var firstZOrderPosition = 0, secondZOrderPosition = 0;
          try { firstZOrderPosition = firstItem.zOrderPosition; } catch (e) { firstZOrderPosition = 0; }
          try { secondZOrderPosition = secondItem.zOrderPosition; } catch (e) { secondZOrderPosition = 0; }
          return firstZOrderPosition - secondZOrderPosition;
        });
      }

      doc.selection = null;
      for (var itemIndex = 0; itemIndex < items.length; itemIndex++) {
        var isOddIndex = (itemIndex % 2 === 0);
        if ((selectOdd && isOddIndex) || (!selectOdd && !isOddIndex)) {
          items[itemIndex].selected = true;
        }
      }
      app.redraw();
    });
    var modeToSave = verticalRadio.value ? "vertical" : (horizontalRadio.value ? "horizontal" : "zorder");
    saveLastDirectionMode(modeToSave);
    dialog.close(1);
  };

  // 初期状態をプレビュー反映（ダイアログ表示直後に選択を更新） / Initial preview
  applySelectionPreview();


  // ダイアログ表示 / Show dialog
  dialog.show();
})();
