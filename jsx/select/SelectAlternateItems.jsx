#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
  互い違いに選択 / Alternate Select
  - 選択中のオブジェクトを「奇数 / 偶数」で互い違いに選択
  - 方向（垂直 / 水平）でカウント順を切り替え
*/

/* バージョン / Version */
var SCRIPT_VERSION = "v1.1";

function getCurrentLang() {
  return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

/* 日英ラベル定義 / Japanese-English label definitions */
var LABELS = {
  dialogTitle: {
    ja: "互い違いに選択",
    en: "Alternate Select"
  },
  alertNoDoc: {
    ja: "ドキュメントを開いてください。",
    en: "Please open a document."
  },
  alertNoSelection: {
    ja: "オブジェクトを選択してください。",
    en: "Please select objects."
  },
  alertNoValidItems: {
    ja: "有効なオブジェクトが選択されていません。",
    en: "No valid objects are selected."
  },
  panelSelect: {
    ja: "選択",
    en: "Select"
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
    en: "Z-Order"
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

function L(key) {
  if (!LABELS[key]) return key;
  return LABELS[key][lang] || LABELS[key].en || key;
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
      var s = doc.selection;
      if (s && s.length) {
        for (var i = 0; i < s.length; i++) snap.push(s[i]);
      }
    } catch (e) {}
    this.selectionSnapshot = snap;
  };

  /* 選択状態を復元 / Restore selection */
  this.restoreSelection = function (doc) {
    try {
      doc.selection = null;
      if (this.selectionSnapshot && this.selectionSnapshot.length) {
        for (var i = 0; i < this.selectionSnapshot.length; i++) {
          try { this.selectionSnapshot[i].selected = true; } catch (e) {}
        }
      }
      app.redraw();
    } catch (e) {}
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
      alert("Preview Error: " + e);
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
    } catch (e) {}
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
    } catch (e) {}
}

/* 方向の自動推定 / Auto detect direction (initial)
   - 中心点のX/Yレンジで判定（外れ値耐性あり）
   - rangeX が rangeY の DIRECTION_THRESHOLD 倍より大きい → "horizontal"
   - rangeY が rangeX の DIRECTION_THRESHOLD 倍より大きい → "vertical"
   - それ以外（僅差/グリッド等）は「前回mode優先（fallback）」を返す
*/
function guessInitialDirectionMode(items, fallbackMode) {
    var fb = (fallbackMode !== undefined && fallbackMode !== null) ? String(fallbackMode) : "vertical";
    if (fb !== "vertical" && fb !== "horizontal" && fb !== "zorder") fb = "vertical";

    if (!items || items.length < 2) return fb;

    var xs = [];
    var ys = [];

    for (var i = 0; i < items.length; i++) {
        var it = items[i];
        if (!it) continue;

        // geometricBounds: [left, top, right, bottom]
        var b;
        try { b = it.geometricBounds; } catch (e) { continue; }
        if (!b || b.length < 4) continue;

        var cx = (b[0] + b[2]) / 2;
        var cy = (b[1] + b[3]) / 2;
        xs.push(cx);
        ys.push(cy);
    }

    if (xs.length < 2 || ys.length < 2) return fb;

    xs.sort(function (a, b) { return a - b; });
    ys.sort(function (a, b) { return a - b; });

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
    return fb;
}

(function () {
    /* ドキュメント確認 / Check document */
    if (app.documents.length === 0) {
        alert(L("alertNoDoc"));
        return;
    }

    var doc = app.activeDocument;

    /* プレビューマネージャ / Preview manager */
    var previewMgr = new PreviewManager();
    previewMgr.captureSelection(doc);

    var sel = doc.selection;

    /* 選択確認 / Check selection */
    if (!sel || sel.length === 0) {
        alert(L("alertNoSelection"));
        return;
    }

    /* 有効なオブジェクト抽出 / Collect valid items */
    var items = [];
    for (var i = 0; i < sel.length; i++) {
        if (!sel[i].locked && !sel[i].hidden) {
            items.push(sel[i]);
        }
    }

    if (items.length === 0) {
        alert(L("alertNoValidItems"));
        return;
    }

    function applySelectionPreview() {
        // まず前回プレビューを巻き戻す（選択はスナップショットで復元）
        previewMgr.rollback(doc);

        // 今回のプレビューを適用（選択変更は通常 Undoable ではないので undoable=false）
        previewMgr.addStep(function () {
            var selectOdd = rbOdd.value;
            var mode = rbVertical.value ? "vertical" : (rbHorizontal.value ? "horizontal" : "zorder");

            /* 並び順ソート / Sort order */
            if (mode === "vertical") {
                /* 垂直：Y降順（上→下） / Vertical: Y desc (top to bottom) */
                items.sort(function (a, b) { return b.position[1] - a.position[1]; });
            } else if (mode === "horizontal") {
                /* 水平：X昇順（左→右） / Horizontal: X asc (left to right) */
                items.sort(function (a, b) { return a.position[0] - b.position[0]; });
            } else {
                /* 重ね順：zOrderPosition 昇順 / Z-order: zOrderPosition asc */
                items.sort(function (a, b) {
                    var za = 0, zb = 0;
                    try { za = a.zOrderPosition; } catch (e) { za = 0; }
                    try { zb = b.zOrderPosition; } catch (e) { zb = 0; }
                    return za - zb;
                });
            }

            /* 互い違い選択 / Alternate selection */
            doc.selection = null;
            for (var i = 0; i < items.length; i++) {
                var isOddIndex = (i % 2 === 0); // 0,2,4... => 1,3,5...
                if ((selectOdd && isOddIndex) || (!selectOdd && !isOddIndex)) {
                    items[i].selected = true;
                }
            }
        }, false);
    }

    /* ダイアログ表示調整用パラメータ / Dialog display adjustment parameters */
    var offsetX = 300;
    var offsetY = 0;
    var dialogOpacity = 0.975;

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
            // opacity 非対応環境対策（無視） / Ignore if opacity not supported
        }
    }

    /* キー入力でラジオ切替 / Key handler for radio buttons
       Odd: O / Even: E / Vertical: V / Horizontal: H / Z-Order: A
    */
    function addRadioKeyHandler(dialog) {
        dialog.addEventListener("keydown", function (event) {
            var k = event.keyName;

            // 奇数 / 偶数
            if (k === "O") {
                rbOdd.value = true;
                rbEven.value = false;
                applySelectionPreview();
                event.preventDefault();
                return;
            } else if (k === "E") {
                rbOdd.value = false;
                rbEven.value = true;
                applySelectionPreview();
                event.preventDefault();
                return;
            }

            // 方向（垂直 / 水平 / 重ね順）
            if (k === "V") {
                rbVertical.value = true;
                rbHorizontal.value = false;
                rbZOrder.value = false;
                saveLastDirectionMode("vertical");
                applySelectionPreview();
                event.preventDefault();
                return;
            } else if (k === "H") {
                rbVertical.value = false;
                rbHorizontal.value = true;
                rbZOrder.value = false;
                saveLastDirectionMode("horizontal");
                applySelectionPreview();
                event.preventDefault();
                return;
            } else if (k === "A") {
                rbVertical.value = false;
                rbHorizontal.value = false;
                rbZOrder.value = true;
                saveLastDirectionMode("zorder");
                applySelectionPreview();
                event.preventDefault();
                return;
            }
        });
    }

    /* ダイアログボックス / Dialog box */
    var dlg = new Window("dialog", L("dialogTitle") + " " + SCRIPT_VERSION);
    setDialogOpacity(dlg, dialogOpacity);
    shiftDialogPosition(dlg, offsetX, offsetY);
    dlg.orientation = "column";
    dlg.alignChildren = "left";

    var panel = dlg.add("panel", undefined, L("panelSelect"));
    panel.margins = [15, 20, 15, 10];
    panel.orientation = "column";
    panel.alignChildren = "left";

    var selectGroup = panel.add("group");
    selectGroup.orientation = "row";
    selectGroup.alignChildren = "left";

    var rbOdd = selectGroup.add("radiobutton", undefined, L("odd"));
    var rbEven = selectGroup.add("radiobutton", undefined, L("even"));
    rbOdd.value = true; // デフォルトは奇数 / Default is odd

    /* 方向パネル / Direction panel */
    var dirPanel = dlg.add("panel", undefined, L("panelDirection"));
    dirPanel.margins = [15, 20, 15, 10];
    dirPanel.orientation = "column";
    dirPanel.alignChildren = "left";

    var dirGroup = dirPanel.add("group");
    dirGroup.orientation = "row";
    dirGroup.alignChildren = "left";

    var rbVertical = dirGroup.add("radiobutton", undefined, L("vertical"));
    var rbHorizontal = dirGroup.add("radiobutton", undefined, L("horizontal"));
    var rbZOrder = dirGroup.add("radiobutton", undefined, L("zOrder"));
    // 前回値を基本にし、差が明確なときだけ自動切替 / Prefer last value; auto-switch only if clear
    var lastMode = loadLastDirectionMode();
    if (lastMode === null) lastMode = "vertical";

    var initialMode = guessInitialDirectionMode(items, lastMode);
    rbVertical.value = (initialMode === "vertical");
    rbHorizontal.value = (initialMode === "horizontal");
    rbZOrder.value = (initialMode === "zorder");

    addRadioKeyHandler(dlg);

    var btnGroup = dlg.add("group");
    btnGroup.alignment = "right";
    var cancelBtn = btnGroup.add("button", undefined, L("cancel"));
    var okBtn = btnGroup.add("button", undefined, L("ok"));

    rbOdd.onClick = applySelectionPreview;
    rbEven.onClick = applySelectionPreview;
    rbVertical.onClick = function () {
        saveLastDirectionMode("vertical");
        applySelectionPreview();
    };
    rbHorizontal.onClick = function () {
        saveLastDirectionMode("horizontal");
        applySelectionPreview();
    };
    rbZOrder.onClick = function () {
        saveLastDirectionMode("zorder");
        applySelectionPreview();
    };

    /* ボタン動作 / Button handlers */
    cancelBtn.onClick = function () {
        // キャンセル：プレビューを巻き戻して閉じる / Cancel: rollback and close
        previewMgr.rollback(doc);
        dlg.close(0);
    };

    okBtn.onClick = function () {
        // OK：Undo を綺麗にするため「一度戻して再実行」 / OK: rollback then re-apply once
        previewMgr.confirm(doc, function () {
            // 最終適用（1回） / Final apply (single step)
            var selectOdd = rbOdd.value;
            var mode = rbVertical.value ? "vertical" : (rbHorizontal.value ? "horizontal" : "zorder");

            if (mode === "vertical") {
                items.sort(function (a, b) { return b.position[1] - a.position[1]; });
            } else if (mode === "horizontal") {
                items.sort(function (a, b) { return a.position[0] - b.position[0]; });
            } else {
                items.sort(function (a, b) {
                    var za = 0, zb = 0;
                    try { za = a.zOrderPosition; } catch (e) { za = 0; }
                    try { zb = b.zOrderPosition; } catch (e) { zb = 0; }
                    return za - zb;
                });
            }

            doc.selection = null;
            for (var i = 0; i < items.length; i++) {
                var isOddIndex = (i % 2 === 0);
                if ((selectOdd && isOddIndex) || (!selectOdd && !isOddIndex)) {
                    items[i].selected = true;
                }
            }
            app.redraw();
        });
        var modeToSave = rbVertical.value ? "vertical" : (rbHorizontal.value ? "horizontal" : "zorder");
        saveLastDirectionMode(modeToSave);
        dlg.close(1);
    };

    // 初期状態をプレビュー反映（ダイアログ表示直後に選択を更新） / Initial preview
    applySelectionPreview();

    // ダイアログ表示 / Show dialog
    dlg.show();
})();
