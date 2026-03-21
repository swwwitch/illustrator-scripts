#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*****
### スクリプト名 / Script name:
AutoFitTextFrame

### 更新日 / Updated:
20260303

### 概要 / Description:
パス上文字（TextType.PATHTEXT）とエリア内文字（TextType.AREATEXT）に対応したテキストサイズ調整スクリプト。
Adjusts font size for PathText and AreaText to eliminate overset.
選択中の TextFrame のみに対して実行します。
Operates only on selected TextFrames.
ダイアログで処理を選択（チェックボックスは排他的にしない）：
Select processing mode via dialog (checkboxes are not mutually exclusive):
  ・あふれ処理 / Shrink to fit：文字あふれ（オーバーセット）がなくなるまで文字サイズを縮小
  ・ぴったり / Maximize fit：いったん文字サイズを倍程度にしてオーバーセットを発生させ、その後「あふれ処理」を実行
    -> 余白があっても「最大であふれないサイズ」まで詰める
  ・両方ON / Both：ぴったり → あふれ処理 の順で実行
データセット使用時、最初のデータセットで元の文字サイズをタグに記録し、
各データセット適用時に一度リセットしてから処理します。
最終データセット処理後にタグを削除します。

※ パス上文字は従来通り「表示行に収まっている文字数」で判定。
※ エリア内文字は可能な場合 `overflows`（Illustratorのオーバーフロー判定）を優先して判定します。
*****/

(function () {

  /* バージョン / Version */
  var SCRIPT_VERSION = "v2.0";

  /* 言語設定 / Language setting */
  function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
  }
  var lang = getCurrentLang();

  /* 日英ラベル定義 / Japanese-English label definitions */
  var LABELS = {
    dialogTitle: {
      ja: "文字サイズ自動調整",
      en: "Auto Fit Text Frame"
    },
    panelLabel: {
      ja: "処理",
      en: "Mode"
    },
    cbOverflow: {
      ja: "あふれ処理",
      en: "Shrink to fit"
    },
    cbFit: {
      ja: "ぴったり",
      en: "Maximize fit"
    },
    btnCancel: {
      ja: "閉じる",
      en: "Close"
    },
    btnOk: {
      ja: "OK",
      en: "OK"
    },
    alertSelectObject: {
      ja: "対象オブジェクトを選択してください。",
      en: "Please select target objects."
    },
    alertNoValidText: {
      ja: "選択中に処理可能なテキストがありません。\n（グループ内のテキストはグループごと選択でもOK）",
      en: "No processable text in selection.\n(Groups containing text can be selected as a group.)"
    },
    alertSelectMode: {
      ja: "処理を選択してください。\n（あふれ処理 / ぴったり）",
      en: "Please select a mode.\n(Shrink to fit / Maximize fit)"
    },
    alertHardReturn: {
      ja: "改行コードが含まれているテキストには対応していません。\n対象テキスト: ",
      en: "Text containing line breaks is not supported.\nTarget: "
    },
    alertHardReturnError: {
      ja: "改行コード判定中にエラーが発生しました。\n",
      en: "An error occurred while checking for line breaks.\n"
    },
    alertMaxIter: {
      ja: "縮小処理が上限回数に達しました:\n",
      en: "Shrink iteration limit reached:\n"
    }
  };

  /* ローカライズ文字列取得 / Get localized string */
  function L(key) {
    var entry = LABELS[key];
    if (!entry) return key;
    return entry[lang] || entry["en"] || key;
  }

  var DealWithOversetText = (function () {

    /* デフォルト値 / Defaults */
    var DEFAULTS = {
      tagName: "overset_text_default_size",
      increment: 0.1,
      minFontSize: 0.1,       // pt
      maxShrinkIter: 2000,    /* 安全上限 / Safety limit */
      alertOnMaxIter: true
    };

    function mergeOptions(userOpt) {
      var opt = {};
      var k;

      for (k in DEFAULTS) {
        if (DEFAULTS.hasOwnProperty(k)) opt[k] = DEFAULTS[k];
      }
      if (userOpt) {
        for (k in userOpt) {
          if (userOpt.hasOwnProperty(k)) opt[k] = userOpt[k];
        }
      }

      opt.increment = (opt.increment * 1) || DEFAULTS.increment;
      opt.minFontSize = (opt.minFontSize * 1) || DEFAULTS.minFontSize;
      opt.maxShrinkIter = Math.floor((opt.maxShrinkIter * 1) || DEFAULTS.maxShrinkIter);
      if (opt.maxShrinkIter < 1) opt.maxShrinkIter = DEFAULTS.maxShrinkIter;
      if (opt.increment <= 0) opt.increment = DEFAULTS.increment;
      if (opt.minFontSize <= 0) opt.minFontSize = DEFAULTS.minFontSize;

      return opt;
    }

    function defaultFilter(tf) {
      return ((tf.kind == TextType.PATHTEXT || tf.kind == TextType.AREATEXT) && tf.editable && !tf.locked && !tf.hidden);
    }

    function collectTextFramesFromItem(item, out, filterFn) {
      if (!item) return;

      /* TextRange 選択の場合 / TextRange selection case */
      try {
        if (item.typename === "TextRange") {
          if (item.parent && item.parent.typename === "TextFrame") {
            collectTextFramesFromItem(item.parent, out, filterFn);
          }
          return;
        }
      } catch (eTR) { }

      /* 直接 TextFrame の場合 / Direct TextFrame */
      try {
        if (item.typename === "TextFrame") {
          if (filterFn(item)) out.push(item);
          return;
        }
      } catch (eTF) { }

      /* GroupItem の場合、pageItems を再帰 / GroupItem: recurse into pageItems */
      try {
        if (item.typename === "GroupItem" && item.pageItems && item.pageItems.length) {
          for (var i = 0; i < item.pageItems.length; i++) {
            collectTextFramesFromItem(item.pageItems[i], out, filterFn);
          }
          return;
        }
      } catch (eG) { }

      /* CompoundPathItem: pathItems を再帰 / CompoundPathItem: recurse into pathItems */
      try {
        if (item.typename === "CompoundPathItem" && item.pathItems && item.pathItems.length) {
          for (var j = 0; j < item.pathItems.length; j++) {
            collectTextFramesFromItem(item.pathItems[j], out, filterFn);
          }
          return;
        }
      } catch (eC) { }

      /* その他のコンテナ（レイヤーなど）/ Other containers (Layer, etc.) */
      try {
        if (item.pageItems && item.pageItems.length) {
          for (var k = 0; k < item.pageItems.length; k++) {
            collectTextFramesFromItem(item.pageItems[k], out, filterFn);
          }
        }
      } catch (eAny) { }
    }

    function getSelectedTextFrames(doc, filterFn) {
      var res = [];
      if (!doc || !doc.selection || doc.selection.length === 0) return res;

      /* 収集（重複あり）/ Collect (may include duplicates) */
      for (var s = 0; s < doc.selection.length; s++) {
        collectTextFramesFromItem(doc.selection[s], res, filterFn);
      }

      /* 重複除去 / De-duplicate by object reference */
      var uniq = [];
      for (var i = 0; i < res.length; i++) {
        var tf = res[i];
        var exists = false;
        for (var j = 0; j < uniq.length; j++) {
          if (uniq[j] === tf) {
            exists = true;
            break;
          }
        }
        if (!exists) uniq.push(tf);
      }

      return uniq;
    }

    /* 互換性のため保持（未使用）/ Kept for compatibility (not used: selection-only) */
    function getTargets(doc, filterFn) {
      var res = [];
      var i, tf;
      for (i = 0; i < doc.textFrames.length; i++) {
        tf = doc.textFrames[i];
        try {
          if (filterFn(tf)) res.push(tf);
        } catch (e) { }
      }
      return res;
    }

    function recordFontSizeInTag(tf, tagName) {
      var tag;
      var tags = tf.tags;
      var size = tf.textRange.characterAttributes.size;
      try {
        tag = tags.getByName(tagName);
        tag.value = size;
      } catch (e) {
        tag = tags.add();
        tag.name = tagName;
        tag.value = size;
      }
    }

    function readFontSizeFromTag(tf, tagName) {
      try {
        return tf.tags.getByName(tagName).value * 1;
      } catch (e) {
        return null;
      }
    }

    function resetSize(tf, tagName) {
      if (tf.contents === "") return;
      var size = readFontSizeFromTag(tf, tagName);
      if (size != null) {
        tf.textRange.characterAttributes.size = size;
      }
    }

    function isOverset(tf, lineAmt) {
      if (tf.lines.length > 0) {
        var charactersOnVisibleLines = 0;

        if (typeof (lineAmt) === "undefined" || lineAmt === null) {
          lineAmt = 1;
        } else {
          lineAmt = Math.floor(lineAmt);
          if (lineAmt < 1) lineAmt = 1;
          if (lineAmt > tf.lines.length) lineAmt = tf.lines.length;
        }

        for (var i = 0; i < lineAmt; i++) {
          charactersOnVisibleLines += tf.lines[i].characters.length;
        }
        return (charactersOnVisibleLines < tf.characters.length);
      } else if (tf.characters.length > 0) {
        return true;
      }
      return false;
    }

    function safeOverflows(tf) {
      try {
        if (tf && typeof tf.overflows !== "undefined") return !!tf.overflows;
      } catch (e) { }
      return null;
    }

    function isOversetFrame(tf) {
      /* エリア内文字：overflows を優先 / AREATEXT: prefer built-in overflows when available */
      if (tf && tf.kind == TextType.AREATEXT) {
        var ov = safeOverflows(tf);
        if (ov !== null) return ov;
        try {
          return isOverset(tf, (tf.lines && tf.lines.length) ? tf.lines.length : 1);
        } catch (eA) {
          return false;
        }
      }

      /* パス上文字：表示行の文字数で判定 / PATHTEXT: use visible-line character count */
      /* NOTE: overflows can exist but may not be reliable for PathText. */
      try {
        return isOverset(tf, (tf.lines && tf.lines.length) ? tf.lines.length : 1);
      } catch (eP) {
        return false;
      }
    }

    /* 手動行送り時のみ比率を返す / Returns leading ratio only for manual leading (null if auto) */
    function getLeadingInfo(tf) {
      try {
        var attrs = tf.textRange.characterAttributes;
        if (attrs.autoLeading) return null;
        var size = attrs.size;
        var leading = attrs.leading;
        if (size > 0 && leading > 0) return { ratio: leading / size };
      } catch (e) { }
      return null;
    }

    function applyLeading(tf, newSize, leadingInfo) {
      if (!leadingInfo) return;
      try {
        tf.textRange.characterAttributes.leading = newSize * leadingInfo.ratio;
      } catch (e) { }
    }

    function hasHardReturn(tf) {
      try {
        return /[\r\n]/.test(tf.contents);
      } catch (e) {
        return false;
      }
    }

    function stopIfHardReturn(tf) {
      if (hasHardReturn(tf)) {
        alert(L("alertHardReturn") + (tf.name ? tf.name : "[unnamed]"));
        return true;
      }
      return false;
    }

    function shrinkFont(tf, opt) {
      var inc = opt.increment;

      try {
        if (stopIfHardReturn(tf)) return false;
      } catch (eCheckReturn) {
        alert(L("alertHardReturnError") + eCheckReturn);
        return false;
      }

      if (tf.characters.length <= 0) return true;
      if (!isOversetFrame(tf)) return true;

      var leadingInfo = getLeadingInfo(tf);
      var iter = 0;
      while (true) {
        var oversetNow;
        try {
          oversetNow = isOversetFrame(tf);
        } catch (eCheck) {
          break;
        }
        if (!oversetNow) break;

        var cur = tf.textRange.characterAttributes.size;
        if (cur <= opt.minFontSize) break;

        var newSize = Math.max(opt.minFontSize, cur - inc);
        tf.textRange.characterAttributes.size = newSize;
        applyLeading(tf, newSize, leadingInfo);

        iter++;
        if (iter >= opt.maxShrinkIter) {
          if (opt.alertOnMaxIter) {
            try {
              alert(L("alertMaxIter") + (tf.name ? tf.name : "[unnamed]"));
            } catch (eA) { }
          }
          break;
        }
      }

      return true;
    }

    function fitFont(tf, opt) {
      /* ぴったり / Maximize fit:
         1) 文字サイズを倍にしてオーバーセットを発生 / Grow font to force overset
         2) あふれ処理で縮小 / Then shrink to fit */

      try {
        if (stopIfHardReturn(tf)) return false;
      } catch (eCheckReturn) {
        alert(L("alertHardReturnError") + eCheckReturn);
        return false;
      }

      if (tf.characters.length <= 0) return true;

      var leadingInfo = getLeadingInfo(tf);
      var original = tf.textRange.characterAttributes.size;

      /* Step 1: オーバーセットが出るまで拡大 / Grow until overset (x2) */
      if (!isOversetFrame(tf)) {
        var high = original;
        var guardUp = 0;
        while (!isOversetFrame(tf) && guardUp < 25) {
          guardUp++;
          high = high * 2;
          if (high > 100000) break;
          try {
            tf.textRange.characterAttributes.size = high;
            applyLeading(tf, high, leadingInfo);
          } catch (eSet) {
            break;
          }
        }
      }

      /* オーバーセットが出ない場合は元に戻す / If still not overset, keep original */
      if (!isOversetFrame(tf)) {
        try {
          tf.textRange.characterAttributes.size = original;
          applyLeading(tf, original, leadingInfo);
        } catch (eBack) { }
        return true;
      }

      /* Step 2: あふれ処理で縮小 / Shrink to fit */
      return shrinkFont(tf, opt);
    }

    function removeTag(tf, tagName) {
      try { tf.tags.getByName(tagName).remove(); } catch (e) { }
    }

    function isFirstDataSet(doc) {
      return (doc.dataSets.length > 0 && doc.activeDataSet == doc.dataSets[0]);
    }

    function isLastDataSet(doc) {
      return (doc.dataSets.length > 0 && doc.activeDataSet == doc.dataSets[doc.dataSets.length - 1]);
    }

    function run(doc, options) {
      if (!doc) return false;

      var opt = mergeOptions(options);
      var filterFn = opt.filter || defaultFilter;

      /* 選択確認 / Check selection */
      if (!doc.selection || doc.selection.length === 0) {
        alert(L("alertSelectObject"));
        return false;
      }

      var targets = getSelectedTextFrames(doc, filterFn);

      if (targets.length === 0) {
        alert(L("alertNoValidText"));
        return false;
      }

      /* データセット：最初 → デフォルトサイズを記録 / Dataset: first => record defaults */
      if (isFirstDataSet(doc)) {
        for (var i = 0; i < targets.length; i++) {
          recordFontSizeInTag(targets[i], opt.tagName);
        }
      }

      /* サイズをリセット / Reset size */
      for (var j = 0; j < targets.length; j++) {
        resetSize(targets[j], opt.tagName);
      }

      /* 処理実行 / Process selection */
      /* 外部から options.mode が渡された場合は旧来の挙動を維持 / If options.mode is provided externally, keep legacy behavior */
      if (options && options.mode) {
        var mode = options.mode;
        if (mode === "fit") {
          for (var k0 = 0; k0 < targets.length; k0++) {
            var okFit0 = fitFont(targets[k0], opt);
            if (okFit0 === false) return false;
          }
        } else {
          for (var k1 = 0; k1 < targets.length; k1++) {
            var ok1 = shrinkFont(targets[k1], opt);
            if (ok1 === false) return false;
          }
        }
      } else {
        var doFit = !!(options && options.doFit);
        var doOverflow = !!(options && options.doOverflow);

        if (!doFit && !doOverflow) {
          alert(L("alertSelectMode"));
          return false;
        }

        /* 両方ON：ぴったり → あふれ処理 / Both ON: maximize fit then shrink */
        if (doFit) {
          for (var kF = 0; kF < targets.length; kF++) {
            var okFit = fitFont(targets[kF], opt);
            if (okFit === false) return false;
          }
        }
        if (doOverflow) {
          for (var kO = 0; kO < targets.length; kO++) {
            var okO = shrinkFont(targets[kO], opt);
            if (okO === false) return false;
          }
        }
      }

      /* データセット：最後 → タグ削除 / Dataset: last => cleanup */
      if (isLastDataSet(doc)) {
        for (var m = 0; m < targets.length; m++) {
          removeTag(targets[m], opt.tagName);
        }
      }

      return true;
    }

    return { run: run };
  })();


  /* セッション中のダイアログ位置を保持 / Retain dialog position within session */
  var _dialogLocation = null;

  function showDialogAndRun() {
    if (app.documents.length === 0) return;

    var win = new Window("dialog", L("dialogTitle") + " " + SCRIPT_VERSION);
    win.orientation = "column";
    win.alignChildren = ["fill", "top"];

    /* 前回位置を復元 / Restore previous position */
    if (_dialogLocation) {
      try { win.location = _dialogLocation; } catch (e) { }
    }

    /* 閉じるとき位置を記録 / Save position on close */
    win.onClose = function () {
      try { _dialogLocation = [win.location[0], win.location[1]]; } catch (e) { }
    };

    var p = win.add("panel", undefined, L("panelLabel"));
    p.orientation = "column";
    p.alignChildren = ["left", "top"];
    p.margins = [15, 20, 15, 10];

    var cbOverflow = p.add("checkbox", undefined, L("cbOverflow"));
    var cbFit = p.add("checkbox", undefined, L("cbFit"));

    cbOverflow.value = true;
    cbFit.value = true;

    var gBtn = win.add("group");
    gBtn.alignment = ["right", "center"];
    gBtn.alignChildren = ["right", "center"];
    var btnCancel = gBtn.add("button", undefined, L("btnCancel"), { name: "cancel" });
    var btnOk = gBtn.add("button", undefined, L("btnOk"), { name: "ok" });

    btnOk.onClick = function () {
      var doFit = !!cbFit.value;
      var doOverflow = !!cbOverflow.value;
      if (!doFit && !doOverflow) {
        alert(L("alertSelectMode"));
        return;
      }
      win.close(1);
      DealWithOversetText.run(app.activeDocument, {
        doFit: doFit,
        doOverflow: doOverflow
      });
    };

    btnCancel.onClick = function () {
      win.close(0);
    };

    win.show();
  }

  showDialogAndRun();
})();
