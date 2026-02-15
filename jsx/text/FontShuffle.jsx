#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
  FontShuffle.jsx
  選択したテキストの1文字ごとにランダムなフォントを適用するスクリプト（プレビュー対応）

  - ダイアログボックスを表示
  - ［再実行］ボタンでランダム適用（プレビュー更新）

  更新日: 2026-02-16
  変更: 犯行声明文風の適用で再ランダム化しない（プレビュー維持）
  変更: 英数字以外を含む場合は和文フォントに限るを自動ON
*/


var SCRIPT_VERSION = "v1.0.2";

function getCurrentLang() {
  return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

/* 日英ラベル定義 / Japanese-English label definitions */
var LABELS = {
  dialogTitle: { ja: "フォントシャッフル", en: "Font Shuffle" },
  info: {
    ja: "選択したテキストに対して、1文字ごとにランダムなフォントを適用します。",
    en: "Applies a random font to each character in the selected text."
  },
  limitToJPFonts: { ja: "和文フォントに限る", en: "Limit to JP fonts" },
  manifestoStyle: { ja: "犯行声明文風", en: "Manifesto style" },
  rerun: { ja: "再実行", en: "Rerun" },
  ok: { ja: "OK", en: "OK" },
  cancel: { ja: "キャンセル", en: "Cancel" },
  alertNoDoc: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
  alertSelectText: { ja: "テキストオブジェクトを選択してください。", en: "Please select at least one text object." },
  alertNoJPFonts: { ja: "対象の和文フォント（Pr6 / Pr6N）が見つかりません。", en: "No target JP fonts (Pr6 / Pr6N) were found." },
  historyPreview: { ja: "FontShuffle プレビュー", en: "FontShuffle Preview" }
};

function L(key) {
  try {
    var o = LABELS[key];
    if (!o) return String(key);
    return (o[lang] != null) ? o[lang] : (o.ja != null ? o.ja : String(key));
  } catch (e) {
    return String(key);
  }
}

/* DialogPersist util (extractable)
 * ダイアログの不透明度・初期位置・位置記憶を共通化
 * 使い方:
 *   DialogPersist.setOpacity(dlg, 0.95);
 *   DialogPersist.restorePosition(dlg, "__YourDialogKey", offsetX, offsetY);
 *   DialogPersist.rememberOnMove(dlg, "__YourDialogKey");
 *   DialogPersist.savePosition(dlg, "__YourDialogKey");
 */
(function(g){
  if (!g.DialogPersist) {
    g.DialogPersist = {
      setOpacity: function(dlg, v){ try{ dlg.opacity=v; }catch(e){} },
      _getSaved: function(key){ return g[key] && g[key].length===2 ? g[key] : null; },
      _setSaved: function(key, loc){ g[key] = [loc[0], loc[1]]; },
      _clampToScreen: function(loc){
        try{
          var vb = ($.screens && $.screens.length) ? $.screens[0].visibleBounds : [0,0,1920,1080];
          var x = Math.max(vb[0]+10, Math.min(loc[0], vb[2]-10));
          var y = Math.max(vb[1]+10, Math.min(loc[1], vb[3]-10));
          return [x,y];
        }catch(e){ return loc; }
      },
      restorePosition: function(dlg, key, offsetX, offsetY){
        var loc = this._getSaved(key);
        try{
          if (loc) dlg.location = this._clampToScreen(loc);
          else { var l = dlg.location; dlg.location = [l[0]+(offsetX|0), l[1]+(offsetY|0)]; }
        }catch(e){}
      },
      rememberOnMove: function(dlg, key){
        var self = this;
        dlg.onMove = function(){
          try{ self._setSaved(key, [dlg.location[0], dlg.location[1]]); }catch(e){}
        };
      },
      savePosition: function(dlg, key){
        try{ this._setSaved(key, [dlg.location[0], dlg.location[1]]); }catch(e){}
      }
    };
  }
})($.global);
(function () {
    // ドキュメントが開かれているか確認
    if (app.documents.length === 0) {
        alert(L('alertNoDoc'));
        return;
    }

    var doc = app.activeDocument;

    // インストールされている全フォントを取得（数が多いと処理に時間がかかる場合があります）
    var allFonts = app.textFonts;

var fontCount = allFonts.length;

// 背景矩形（1文字ごと・アウトライン経由） / Background rectangles (per-character via outline)
var BG_LAYER_NAME = '__FontShuffle_BG__';
var BG_GROUP_NAME = '__BGRects__';

function ensureBgLayer() {
    var lyr = null;
    for (var i = 0; i < doc.layers.length; i++) {
        try {
            if (doc.layers[i].name === BG_LAYER_NAME) { lyr = doc.layers[i]; break; }
        } catch (_) { }
    }
    if (!lyr) {
        lyr = doc.layers.add();
        lyr.name = BG_LAYER_NAME;
    }
    return lyr;
}

function removeBgGroupIfExists(bgLayer) {
    try {
        for (var i = bgLayer.groupItems.length - 1; i >= 0; i--) {
            if (bgLayer.groupItems[i].name === BG_GROUP_NAME) {
                bgLayer.groupItems[i].remove();
            }
        }
    } catch (_) { }
}

function randInt(min, max) {
    return Math.floor(Math.random() * (max - min + 1)) + min;
}

function makeRandomFill() {
    var g = new GrayColor();
    // GrayColor.gray: 0 = white, 100 = black
    // K10〜K50 の間でランダム
    g.gray = randInt(10, 50);
    return g;
}

function setRectStyle(rect) {
    try {
        rect.stroked = false;
        rect.filled = true;
        rect.fillColor = makeRandomFill();
    } catch (_) { }
}

function jitterRectOutward(rect, maxPt) {
    if (!rect || rect.typename !== 'PathItem') return;
    if (!rect.pathPoints || rect.pathPoints.length < 4) return;

    var b;
    try { b = rect.geometricBounds; } catch (_) { return; }
    // b: [L, T, R, B]
    var cx = (b[0] + b[2]) / 2;
    var cy = (b[1] + b[3]) / 2;

    for (var i = 0; i < rect.pathPoints.length; i++) {
        try {
            var pp = rect.pathPoints[i];
            var ax = pp.anchor[0], ay = pp.anchor[1];
            var dx = ax - cx;
            var dy = ay - cy;
            var len = Math.sqrt(dx * dx + dy * dy);
            if (!len || len === 0) continue;

            // random outward distance
            var d = Math.random() * maxPt;
            var ux = dx / len;
            var uy = dy / len;
            var ox = ux * d;
            var oy = uy * d;

            // move anchor + handles together to keep corner shape
            pp.anchor = [ax + ox, ay + oy];
            pp.leftDirection = [pp.leftDirection[0] + ox, pp.leftDirection[1] + oy];
            pp.rightDirection = [pp.rightDirection[0] + ox, pp.rightDirection[1] + oy];
        } catch (_) { }
    }
}

function addRectFromBounds(bgGroup, b) {
    // b: [L, T, R, B]
    var PAD = 1; // 1pt expansion on each side

    var L = b[0] - PAD;
    var T = b[1] + PAD;
    var R = b[2] + PAD;
    var B = b[3] - PAD;

    var w = R - L;
    var h = T - B;
    if (w <= 0 || h <= 0) return null;

    var rect = bgGroup.pathItems.rectangle(T, L, w, h);
    setRectStyle(rect);
    // Randomly expand anchors outward (rough, organic feel)
    jitterRectOutward(rect, 1);
    return rect;
}

function collectCharGroupsFromOutlined(outlined, out) {
    // C: アウトライン結果の「サブレベル groupItems」を 1文字単位として扱う。
    // 典型例: outlined(GroupItem)
    //   ├─ groupItems[0] (行/ブロック)
    //   │    ├─ groupItems[0] (1文字)
    //   │    ├─ groupItems[1] (1文字)
    //   │    ...
    //   ├─ groupItems[1] (行/ブロック)
    //   ...
    // 例外ケースに備えてフォールバックも用意。
    if (!outlined) return;

    function pushLeafUnits(container) {
        try {
            if (!container || !container.pageItems) return;
            for (var k = 0; k < container.pageItems.length; k++) {
                var it = container.pageItems[k];
                if (!it) continue;
                if (it.typename === 'GroupItem') {
                    pushLeafUnits(it);
                } else if (it.typename === 'PathItem' || it.typename === 'CompoundPathItem') {
                    out.push(it);
                }
            }
        } catch (_) { }
    }

    try {
        if (outlined.typename === 'GroupItem') {
            var hasPushed = false;

            // Prefer: sub-level groupItems (child group -> its groupItems)
            if (outlined.groupItems && outlined.groupItems.length > 0) {
                for (var i = 0; i < outlined.groupItems.length; i++) {
                    var g1 = outlined.groupItems[i];
                    if (!g1 || g1.typename !== 'GroupItem') continue;

                    if (g1.groupItems && g1.groupItems.length > 0) {
                        for (var j = 0; j < g1.groupItems.length; j++) {
                            out.push(g1.groupItems[j]);
                            hasPushed = true;
                        }
                    } else {
                        // If the child has no further groups, use it as a unit
                        out.push(g1);
                        hasPushed = true;
                    }
                }
            }

            if (hasPushed) return;

            // Fallback: leaf units so it won't collapse to a single unit
            pushLeafUnits(outlined);
            if (out.length > 0) return;

            // Last resort: whole outlined
            out.push(outlined);
            return;
        }

        out.push(outlined);
    } catch (_) { }
}

function createBackgroundRectsByOutlining(frames) {
    if (!frames || frames.length === 0) return;

    var bgLayer = ensureBgLayer();
    removeBgGroupIfExists(bgLayer);
    var bgGroup = bgLayer.groupItems.add();
    bgGroup.name = BG_GROUP_NAME;

    for (var i = 0; i < frames.length; i++) {
        var tf = frames[i];
        if (!tf || tf.typename !== 'TextFrame') continue;

        var dup = null;
        var outlined = null;

        // A) duplicate behind original
        try {
            dup = tf.duplicate();
            try { dup.move(tf, ElementPlacement.PLACEAFTER); } catch (_) { }
        } catch (_) {
            continue;
        }

        // B) outline the duplicate
        try {
            outlined = dup.createOutline();
        } catch (_) {
            try { if (dup) dup.remove(); } catch (_) { }
            continue;
        }

        // remove duplicate text
        try { if (dup) dup.remove(); } catch (_) { }

        // C) treat each top-level group as a character group
        var charGroups = [];
        collectCharGroupsFromOutlined(outlined, charGroups);

        // D) rectangle per character group, random fill
        for (var j = 0; j < charGroups.length; j++) {
            try {
                var b = charGroups[j].geometricBounds;
                addRectFromBounds(bgGroup, b);
            } catch (_) { }
        }

        // E) delete outlined artwork (C)
        try { if (outlined) outlined.remove(); } catch (_) { }
    }

    // Keep backgrounds behind content
    try { bgGroup.zOrder(ZOrderMethod.SENDTOBACK); } catch (_) { }
    try { bgLayer.zOrder(ZOrderMethod.SENDTOBACK); } catch (_) { }
}

function setTrackingForFrames(frames, trackingVal) {
    if (!frames || frames.length === 0) return;
    for (var i = 0; i < frames.length; i++) {
        var item = frames[i];
        if (!item || item.typename !== 'TextFrame') continue;
        try {
            // tracking is in 1/1000 em
            item.textRange.characterAttributes.tracking = trackingVal;
        } catch (_) { }
    }
}

function getBgLayerIfExists() {
    try {
        for (var i = 0; i < doc.layers.length; i++) {
            if (doc.layers[i] && doc.layers[i].name === BG_LAYER_NAME) return doc.layers[i];
        }
    } catch (_) { }
    return null;
}

function clearBackgroundRectsIfAny() {
    var lyr = getBgLayerIfExists();
    if (!lyr) return;
    removeBgGroupIfExists(lyr);
}

    function getMorisawaFonts() {
        var list = [];
        var i, f;

        function matchJAKeyword(s) {
            if (!s) return false;
            s = String(s);
            // 和文フォント判定キーワード
            // Pr6 / Pr6N / AB- / FOT を含むものを対象
            return (
                s.indexOf('Pr6N') !== -1 ||
                s.indexOf('Pr6') !== -1 ||
                s.indexOf('AB-') !== -1 ||
                s.indexOf('FOT') !== -1 ||
                s.indexOf('-OTF') !== -1
            );
        }

        for (i = 0; i < allFonts.length; i++) {
            try {
                f = allFonts[i];
                if (!f) continue;

                var n1 = f.fullName ? String(f.fullName) : '';
                var n2 = f.name ? String(f.name) : '';
                var n3 = (f.postScriptName) ? String(f.postScriptName) : '';

                if (matchJAKeyword(n1) || matchJAKeyword(n2) || matchJAKeyword(n3)) {
                    list.push(f);
                }
            } catch (_) { }
        }

        return list;
    }

    function getSelectionTextFrames() {
        var sel = doc.selection;
        if (!sel || sel.length === 0) return [];
        var frames = [];
        for (var i = 0; i < sel.length; i++) {
            try {
                if (sel[i] && sel[i].typename === "TextFrame") frames.push(sel[i]);
            } catch (_) { }
        }
        return frames;
    }
    function selectionContainsNonAlnum(frames) {
        if (!frames || frames.length === 0) return false;

        for (var i = 0; i < frames.length; i++) {
            var tf = frames[i];
            if (!tf || tf.typename !== 'TextFrame') continue;

            try {
                var txt = tf.contents;
                for (var j = 0; j < txt.length; j++) {
                    var ch = txt.charAt(j);
                    // ASCII英数字のみ許可
                    if (!(/[A-Za-z0-9]/.test(ch))) {
                        // 改行は除外
                        if (ch === '\r' || ch === '\n') continue;
                        return true;
                    }
                }
            } catch (_) { }
        }
        return false;
    }

    function applyRandomFontsToFrames(frames, fontList) {
        if (!frames || frames.length === 0) return;
        if (!fontList || fontList.length === 0) return;

        for (var i = 0; i < frames.length; i++) {
            var item = frames[i];
            if (!item || item.typename !== "TextFrame") continue;

            var chars = item.textRange.characters;
            for (var j = 0; j < chars.length; j++) {
                // 改行や空白文字はスキップ（エラー回避と見た目のため）
                if (chars[j].contents === "\r" || chars[j].contents === " ") continue;

                try {
                    var randomFontIndex = Math.floor(Math.random() * fontList.length);
                    chars[j].characterAttributes.textFont = fontList[randomFontIndex];
                } catch (_) {
                    // 特定のフォントが適用できない場合は無視
                }
            }
        }

        app.redraw();
    }

    // --- Dialog ---
// ダイアログ / Dialog
    var DIALOG_KEY = "__FontShuffle_Dialog__";
    var dlg = new Window('dialog', L('dialogTitle') + ' ' + SCRIPT_VERSION);
    DialogPersist.setOpacity(dlg, 0.98);
    dlg.onShow = function () {
        DialogPersist.restorePosition(dlg, DIALOG_KEY, 300, 0);
    };
    DialogPersist.rememberOnMove(dlg, DIALOG_KEY);
    dlg.orientation = 'column';
    dlg.alignChildren = ['fill', 'top'];

    var info = dlg.add('statictext', undefined, L('info'));
    info.characters = 52;

    var chkMorisawa = dlg.add('checkbox', undefined, L('limitToJPFonts'));
    chkMorisawa.value = false;
    var chkManifesto = dlg.add('checkbox', undefined, L('manifestoStyle'));
    chkManifesto.value = false;


    // 犯行声明文風のON/OFFは、プレビュー済みなら「ランダムを回さず」見た目だけ更新
    chkManifesto.onClick = function () {
        if (!_previewApplied) return;
        // 背景の有無だけ切り替え（ONなら作成、OFFなら削除）。フォントは維持。
        runPreview(!!(chkManifesto && chkManifesto.value), false);
    };

    // 下部ボタン行（3カラム） / Bottom bar (3 columns)
    var bottom = dlg.add('group');
    bottom.orientation = 'row';
    bottom.alignChildren = ['fill', 'center'];

    // Left
    var colL = bottom.add('group');
    colL.orientation = 'row';
    colL.alignChildren = ['left', 'center'];
    var btnRerun = colL.add('button', undefined, L('rerun'));

    // Center spacer
    var colC = bottom.add('group');
    colC.orientation = 'row';
    colC.alignment = ['fill', 'center'];
    colC.add('statictext', undefined, '');

    // Right
    var colR = bottom.add('group');
    colR.orientation = 'row';
    colR.alignChildren = ['right', 'center'];
    var btnCancel = colR.add('button', undefined, L('cancel'), { name: 'cancel' });
    var btnOK = colR.add('button', undefined, L('ok'), { name: 'ok' });

    try { colC.preferredSize.width = 20; } catch (_) { }
    try { bottom.alignment = ['fill', 'bottom']; } catch (_) { }
    try { colR.alignment = ['right', 'center']; } catch (_) { }

    // プレビュー状態：プレビュー描画を閉じるときに一括Undo / Preview state: undo all preview steps on close
    var _previewApplied = false;

    // Preview items (reserved) / プレビュー項目（予約）
    var __previewItems = [];

    /* =========================================
     * PreviewHistory util (extractable)
     * ヒストリーを残さないプレビューのための小さなユーティリティ。
     * 他スクリプトでもこのブロックをコピペすれば再利用できます。
     * 使い方:
     *   PreviewHistory.start();     // ダイアログ表示時などにカウンタ初期化
     *   PreviewHistory.bump();      // プレビュー描画ごとにカウント(+1)
     *   PreviewHistory.undo();      // 閉じる/キャンセル時に一括Undo
     *   PreviewHistory.cancelTask(t);// app.scheduleTaskのキャンセル補助
     * ========================================= */
    (function(g){
        if (!g.PreviewHistory) {
            g.PreviewHistory = {
                start: function(){ g.__previewUndoCount = 0; },
                bump:  function(){ g.__previewUndoCount = (g.__previewUndoCount | 0) + 1; },
                undo:  function(){
                    var n = g.__previewUndoCount | 0;
                    try { for (var i = 0; i < n; i++) app.executeMenuCommand('undo'); } catch (e) {}
                    g.__previewUndoCount = 0;
                },
                cancelTask: function(taskId){
                    try { if (taskId) app.cancelTask(taskId); } catch (e) {}
                }
            };
        }
    })($.global);

    function runPreview(makeRects, randomizeFonts, countHistory) {
        if (typeof randomizeFonts === 'undefined') randomizeFonts = true;
        if (typeof countHistory === 'undefined') countHistory = true;

        function _do() {
            var frames = getSelectionTextFrames();
            if (frames.length === 0) {
                alert(L('alertSelectText'));
                return;
            }

            // ランダムを走らせる場合のみ、直前プレビューを undo で戻す
            if (randomizeFonts && _previewApplied) {
                try { app.undo(); } catch (_) { }
            }

            // 1) random fonts (optional)
            if (randomizeFonts) {
                var fontList = null;
                if (chkMorisawa && chkMorisawa.value) {
                    fontList = getMorisawaFonts();
                    if (!fontList || fontList.length === 0) {
                        alert(L('alertNoJPFonts'));
                        return;
                    }
                } else {
                    fontList = allFonts;
                }

                applyRandomFontsToFrames(frames, fontList);
            }

            // 2) 犯行声明文風: tracking=200（ランダムを走らせない場合でも適用可）
            if (chkManifesto && chkManifesto.value) {
                setTrackingForFrames(frames, 200);
            }

            // 3) Background rectangles: only when requested AND 犯行声明文風 ON
            if (makeRects && chkManifesto && chkManifesto.value) {
                createBackgroundRectsByOutlining(frames);
            } else {
                // OFF の場合は背景を付けない（既存があれば消す）
                clearBackgroundRectsIfAny();
            }

            _previewApplied = true;
        }

        // Try to keep the operation to a single history step (so undo once works)
        try {
            if (doc && doc.suspendHistory) {
                doc.suspendHistory(L('historyPreview'), '_do()');
                if (countHistory) PreviewHistory.bump();
                return;
            }
        } catch (_) { }

        _do();
        if (countHistory) PreviewHistory.bump();
    }


    btnRerun.onClick = function () {
        // プレビュー更新（ランダム再実行）
        runPreview(false, true);
    };

    btnOK.onClick = function () {
        // プレビューの履歴を消してから最終適用 / Clear preview history before final apply
        PreviewHistory.undo();
        // プレビュー未実行ならここで1回適用（犯行声明文風 ON の場合のみ長方形も作る）
        if (!_previewApplied) {
            runPreview(true, true, false);
        } else {
            // 既にプレビュー済み：フォントは維持したまま、OK時に必要な見た目だけ確定
            if (chkManifesto && chkManifesto.value) {
                // 背景矩形を付ける（再ランダム化しない）
                runPreview(true, false, false);
            } else {
                // OFF の場合は長方形なし（既存があれば消す）
                clearBackgroundRectsIfAny();
            }
        }
        DialogPersist.savePosition(dlg, DIALOG_KEY);
        dlg.close(1);
    };

    btnCancel.onClick = function () {
        // プレビューを戻す（履歴を残さない） / Undo all previews
        PreviewHistory.undo();
        DialogPersist.savePosition(dlg, DIALOG_KEY);
        dlg.close(0);
    };

    // ダイアログ表示前に選択テキストをチェック
    var _initialFrames = getSelectionTextFrames();
    if (selectionContainsNonAlnum(_initialFrames)) {
        chkMorisawa.value = true;
    }

    PreviewHistory.start();
    dlg.center();
    dlg.show();
})();
