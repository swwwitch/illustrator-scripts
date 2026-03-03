#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*****
### スクリプト名 / Script name:
AutoFitTextFrame

### 更新日：
20260304

### 概要 / Description:
処理（文字サイズ）：
Font size modes:
  ・文字サイズ：あふれ処理 / Shrink to fit：文字あふれ（オーバーセット）がなくなるまで文字サイズを縮小
  ・文字サイズ：ぴったり / Maximize fit：いったん文字サイズを増やしてオーバーセットを発生させ、その後「あふれ処理」を実行
    -> 余白があっても「最大であふれないサイズ」まで詰める
  ・両方ON / Both：ぴったり → あふれ処理 の順で実行

処理（エリア内文字）：
AreaText mode:
  ・エリア内文字の高さ調整 / Adjust AreaText height：ONのとき、文字サイズの処理は無効になり、オプションを選択して実行

オプション（※「エリア内文字の高さ調整」ONのときのみ有効）：
Options (enabled only when “Adjust AreaText height” is ON):
  ・高さを調整 / Adjust height：自動サイズ調整を一時的にON→OFFにして、必要な分だけ高さを拡張して固定
  ・自動サイズ調整 / Auto size：自動サイズ調整をONにします（拡張のみ・OFFには戻しません）

データセット使用時、最初のデータセットで元の値（文字サイズ/高さ）をタグに記録し、
各データセット適用時に一度リセットしてから処理します。
最終データセット処理後にタグを削除します。

※ パス上文字は従来通り「表示行に収まっている文字数」で判定。
※ エリア内文字は可能な場合 `overflows`（Illustratorのオーバーフロー判定）を優先して判定します。
*****/

/* バージョン / Version */
var SCRIPT_VERSION = "v2.3.1";

(function () {

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
    panelAdjust: {
      ja: "オプション",
      en: "Options"
    },
    rbHeight: {
      ja: "高さを調整",
      en: "Adjust height"
    },
    rbAutoSizeArea: {
      ja: "自動サイズ調整",
      en: "Auto size"
    },
    tipAutoSizeArea: {
      ja: "エリア内文字に自動サイズ調整を適用します（拡張のみ）。",
      en: "Applies Auto Size to AreaText (expand only)."
    },
    alertHeightOnlyNoArea: {
      ja: "高さ調整はエリア内テキストのみ対応です。\n選択中にエリア内テキストがありません。",
      en: "Height adjustment is supported for AreaText only.\nNo AreaText found in selection."
    },
    cbOverflow: {
      ja: "文字サイズ：あふれ処理",
      en: "Shrink to fit"
    },
    cbFit: {
      ja: "文字サイズ：ぴったり",
      en: "Maximize fit"
    },
    cbHeightMode: {
      ja: "エリア内文字の高さ調整",
      en: "Adjust AreaText height"
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
      ja: "処理を選択してください。\n（文字サイズ：あふれ処理 / 文字サイズ：ぴったり）",
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
    alertSelectHeightOption: {
      ja: "オプションを選択してください。\n（高さを調整 / 自動サイズ調整）",
      en: "Please select an option.\n(Adjust height / Auto size)"
    },
    alertMaxIter: {
      ja: "縮小処理が上限回数に達しました:\n",
      en: "Shrink iteration limit reached:\n"
    },
    unnamedText: {
      ja: "[無名のテキスト]",
      en: "[Unnamed Text]"
    }
  };

  /* ローカライズ文字列取得 / Get localized string */
  function L(key) {
    var entry = LABELS[key];
    if (!entry) return key;
    return entry[lang] || entry["en"] || key;
  }


  var DealWithOversetText = (function () {

    // Defaults
    var DEFAULTS = {
      tagName: "overset_text_default_size",
      heightTagName: "overset_text_default_height",
      increment: 0.1,
      heightIncrement: 0.5,   // pt (height adjustment step)
      minFontSize: 0.1,       // pt
      maxShrinkIter: 2000,    // safety limit
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

      // TextRange selection case
      try {
        if (item.typename === "TextRange") {
          if (item.parent && item.parent.typename === "TextFrame") {
            collectTextFramesFromItem(item.parent, out, filterFn);
          }
          return;
        }
      } catch (eTR) { }

      // Direct TextFrame
      try {
        if (item.typename === "TextFrame") {
          if (filterFn(item)) out.push(item);
          return;
        }
      } catch (eTF) { }

      // GroupItem: recurse into pageItems
      try {
        if (item.typename === "GroupItem" && item.pageItems && item.pageItems.length) {
          for (var i = 0; i < item.pageItems.length; i++) {
            collectTextFramesFromItem(item.pageItems[i], out, filterFn);
          }
          return;
        }
      } catch (eG) { }

      // CompoundPathItem: recurse into pathItems (text shouldn't be inside, but safe)
      try {
        if (item.typename === "CompoundPathItem" && item.pathItems && item.pathItems.length) {
          for (var j = 0; j < item.pathItems.length; j++) {
            collectTextFramesFromItem(item.pathItems[j], out, filterFn);
          }
          return;
        }
      } catch (eC) { }

      // Other containers (Layer, etc.) - try pageItems if present
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

      // collect (may include duplicates)
      for (var s = 0; s < doc.selection.length; s++) {
        collectTextFramesFromItem(doc.selection[s], res, filterFn);
      }

      // de-dup by real object reference (ExtendScript stringification is not unique)
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

    // kept for compatibility (not used: selection-only)
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
      // AREATEXT: prefer built-in overflows when available
      if (tf && tf.kind == TextType.AREATEXT) {
        var ov = safeOverflows(tf);
        if (ov !== null) return ov;
        try {
          return isOverset(tf, (tf.lines && tf.lines.length) ? tf.lines.length : 1);
        } catch (eA) {
          return false;
        }
      }

      // PATHTEXT (and others): use legacy visible-line character count.
      // NOTE: overflows can exist but may not be reliable for PathText.
      try {
        return isOverset(tf, (tf.lines && tf.lines.length) ? tf.lines.length : 1);
      } catch (eP) {
        return false;
      }
    }

    // 手動行送り時のみ比率を返す（autoLeadingの場合はnull）
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

    // --- Height adjustment functions (AreaText only) ---

    function recordHeightInTag(tf, tagName) {
      var tag;
      var tags = tf.tags;
      var h = tf.height;
      try {
        tag = tags.getByName(tagName);
        tag.value = h;
      } catch (e) {
        tag = tags.add();
        tag.name = tagName;
        tag.value = h;
      }
    }

    function readHeightFromTag(tf, tagName) {
      try {
        return tf.tags.getByName(tagName).value * 1;
      } catch (e) {
        return null;
      }
    }

    function resetHeight(tf, tagName) {
      if (tf.contents === "") return;
      var h = readHeightFromTag(tf, tagName);
      if (h != null) {
        tf.height = h;
      }
    }

    // Illustrator アクションで自動サイズ調整を ON/OFF にする
    function act_setAutoSizeAdjust(valueInt) {
      // valueInt: 1 = ON, 2 = OFF
      if (valueInt !== 1 && valueInt !== 2) return;

      var str = '/version 3'
        + '/name [ 8 4172656154797065]'
        + '/isOpen 1'
        + '/actionCount 1'
        + '/action-1 {'
        + ' /name [ 8 4175746f53697a65 ]'
        + ' /keyIndex 0'
        + ' /colorIndex 0'
        + ' /isOpen 1'
        + ' /eventCount 1'
        + ' /event-1 {'
        + ' /useRulersIn1stQuadrant 0'
        + ' /internalName (adobe_SLOAreaTextDialog)'
        + ' /localizedName [ 33'
        + ' e382a8e383aae382a2e58685e69687e5ad97e382aae38397e382b7e383a7e383b3'
        + ' ]'
        + ' /isOpen 1'
        + ' /isOn 1'
        + ' /hasDialog 0'
        + ' /parameterCount 1'
        + ' /parameter-1 {'
        + ' /key 1952539754'
        + ' /showInPalette 4294967295'
        + ' /type (integer)'
        + ' /value ' + valueInt
        + ' }'
        + ' }'
        + '}';

      var f = new File('~/ScriptAction.aia');
      f.open('w');
      f.write(str);
      f.close();
      app.loadAction(f);
      f.remove();

      app.doScript("AutoSize", "AreaType", false);
      app.unloadAction("AreaType", "");
    }

    function expandFrameToFit(tf) {
      app.activeDocument.selection = [tf];
      act_setAutoSizeAdjust(1);
    }

    function collapseFrameAuto(tf) {
      app.activeDocument.selection = [tf];
      act_setAutoSizeAdjust(2);
    }

    function growHeight(tf, opt) {
      if (tf.characters.length <= 0) return true;
      if (!isOversetFrame(tf)) return true;
      expandFrameToFit(tf);
      collapseFrameAuto(tf);
      return true;
    }

    function fitHeight(tf, opt) {
      if (tf.characters.length <= 0) return true;
      expandFrameToFit(tf);
      collapseFrameAuto(tf);
      return true;
    }

    // --- End height adjustment functions ---

    function hasHardReturn(tf) {
      try {
        return /[\r\n]/.test(tf.contents);
      } catch (e) {
        return false;
      }
    }

    function stopIfHardReturn(tf) {
      if (hasHardReturn(tf)) {
        alert(L("alertHardReturn") + (tf.name ? tf.name : L("unnamedText")));
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
              alert(L("alertMaxIter") + (tf.name ? tf.name : L("unnamedText")));
            } catch (eA) { }
          }
          break;
        }
      }

      return true;
    }

    function fitFont(tf, opt) {
      // ぴったり:
      // 1) いったん文字サイズを倍程度にしてオーバーセットを発生させる
      // 2) その後、あふれ処理（shrinkFont）で詰める

      try {
        if (stopIfHardReturn(tf)) return false;
      } catch (eCheckReturn) {
        alert(L("alertHardReturnError") + eCheckReturn);
        return false;
      }

      if (tf.characters.length <= 0) return true;

      var leadingInfo = getLeadingInfo(tf);
      var original = tf.textRange.characterAttributes.size;

      // Step 1: grow until overset (x2)
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

      // If still not overset, keep original
      if (!isOversetFrame(tf)) {
        try {
          tf.textRange.characterAttributes.size = original;
          applyLeading(tf, original, leadingInfo);
        } catch (eBack) { }
        return true;
      }

      // Step 2: shrink to fit using existing logic
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
      var adjustMode = (options && options.adjustMode) || "fontSize";
      var isHeightMode = (adjustMode === "height");

      // selection-only (TextFrame / TextRange / GroupItem etc.)
      if (!doc.selection || doc.selection.length === 0) {
        alert(L("alertSelectObject"));
        return false;
      }

      var targets = getSelectedTextFrames(doc, filterFn);

      var autoSizeArea = !!(options && options.autoSizeArea);
      if (autoSizeArea) {
        // Auto size (AreaText): run ONLY expandFrameToFit for AreaText frames
        var areaOnly = [];
        for (var aa = 0; aa < targets.length; aa++) {
          if (targets[aa].kind == TextType.AREATEXT) areaOnly.push(targets[aa]);
        }

        if (areaOnly.length === 0) {
          alert(L("alertHeightOnlyNoArea"));
          return false;
        }

        for (var ab = 0; ab < areaOnly.length; ab++) {
          try { expandFrameToFit(areaOnly[ab]); } catch (eAuto) { }
        }
        return true;
      }

      // Height mode: filter to AreaText only
      if (isHeightMode) {
        var areaTargets = [];
        for (var a = 0; a < targets.length; a++) {
          if (targets[a].kind == TextType.AREATEXT) areaTargets.push(targets[a]);
        }
        if (areaTargets.length === 0) {
          alert(L("alertHeightOnlyNoArea"));
          return false;
        }
        targets = areaTargets;
      }

      if (targets.length === 0) {
        alert(L("alertNoValidText"));
        return false;
      }

      var tagKey = isHeightMode ? opt.heightTagName : opt.tagName;

      // dataset: first => record defaults
      if (isFirstDataSet(doc)) {
        for (var i = 0; i < targets.length; i++) {
          if (isHeightMode) {
            recordHeightInTag(targets[i], tagKey);
          } else {
            recordFontSizeInTag(targets[i], tagKey);
          }
        }
      }

      // reset
      for (var j = 0; j < targets.length; j++) {
        if (isHeightMode) {
          resetHeight(targets[j], tagKey);
        } else {
          resetSize(targets[j], tagKey);
        }
      }

      // Helper references for current mode
      var fnShrink = isHeightMode ? growHeight : shrinkFont;
      var fnFit = isHeightMode ? fitHeight : fitFont;

      // process selection
      // If `options.mode` is provided externally, keep legacy behavior.
      if (options && options.mode) {
        var mode = options.mode;
        if (mode === "fit") {
          for (var k0 = 0; k0 < targets.length; k0++) {
            var okFit0 = fnFit(targets[k0], opt);
            if (okFit0 === false) return false;
          }
        } else {
          for (var k1 = 0; k1 < targets.length; k1++) {
            var ok1 = fnShrink(targets[k1], opt);
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

        // 両方ONなら → ぴったり → あふれ処理
        if (doFit) {
          for (var kF = 0; kF < targets.length; kF++) {
            var okFit = fnFit(targets[kF], opt);
            if (okFit === false) return false;
          }
        }
        if (doOverflow) {
          for (var kO = 0; kO < targets.length; kO++) {
            var okO = fnShrink(targets[kO], opt);
            if (okO === false) return false;
          }
        }
      }

      // dataset: last => cleanup
      if (isLastDataSet(doc)) {
        for (var m = 0; m < targets.length; m++) {
          removeTag(targets[m], tagKey);
        }
      }

      return true;
    }

    return { run: run };
  })();


  var _dialogLocation = null;

  function hasAreaTextInSelection(doc) {
    if (!doc || !doc.selection || doc.selection.length === 0) return false;
    var sel = doc.selection;
    for (var i = 0; i < sel.length; i++) {
      try {
        if (sel[i].typename === "TextFrame" && sel[i].kind == TextType.AREATEXT) return true;
      } catch (e) { }
      try {
        if (sel[i].typename === "GroupItem" && sel[i].pageItems) {
          for (var j = 0; j < sel[i].pageItems.length; j++) {
            var pi = sel[i].pageItems[j];
            if (pi.typename === "TextFrame" && pi.kind == TextType.AREATEXT) return true;
          }
        }
      } catch (e2) { }
    }
    return false;
  }

  function showDialogAndRun() {
    if (app.documents.length === 0) return;
    // If nothing is selected, alert and do not open the dialog
    try {
      if (!app.activeDocument.selection || app.activeDocument.selection.length === 0) {
        alert(L("alertSelectObject"));
        return;
      }
    } catch (eSel) {
      alert(L("alertSelectObject"));
      return;
    }

    var hasArea = hasAreaTextInSelection(app.activeDocument);

    var dlg = new Window('dialog', L('dialogTitle') + ' ' + SCRIPT_VERSION);
    dlg.orientation = "column";
    dlg.alignChildren = ["fill", "top"];

    if (_dialogLocation) {
      try { dlg.location = _dialogLocation; } catch (e) { }
    }

    dlg.onClose = function () {
      try { _dialogLocation = [dlg.location[0], dlg.location[1]]; } catch (e) { }
    };

    var p = dlg.add("panel", undefined, L("panelLabel"));
    p.orientation = "column";
    p.alignChildren = ["left", "top"];
    p.margins = [15, 20, 15, 10];

    var cbOverflow = p.add("checkbox", undefined, L("cbOverflow"));
    var cbFit = p.add("checkbox", undefined, L("cbFit"));
    var cbHeightMode = p.add("checkbox", undefined, L("cbHeightMode"));

    cbOverflow.value = true;
    cbFit.value = true;
    cbHeightMode.value = false;
    cbHeightMode.enabled = hasArea;

    var pAdjust = dlg.add("panel", undefined, L("panelAdjust"));
    pAdjust.orientation = "column";
    pAdjust.alignChildren = ["left", "top"];
    pAdjust.margins = [15, 20, 15, 10];

    var rbHeight = pAdjust.add("radiobutton", undefined, L("rbHeight"));
    rbHeight.value = false;

    var rbAutoSizeArea = pAdjust.add("radiobutton", undefined, L("rbAutoSizeArea"));
    rbAutoSizeArea.value = false;
    try { rbAutoSizeArea.helpTip = L("tipAutoSizeArea"); } catch (eTip) { }

    // Remember previous font-size checkbox states
    var _prevFontOverflow = cbOverflow.value;
    var _prevFontFit = cbFit.value;

    function updateOptionsEnabled() {
      var enabled = !!cbHeightMode.value && !!hasArea;

      // Options panel is available only in AreaText height mode
      try { pAdjust.enabled = enabled; } catch (eP) { }
      try { rbHeight.enabled = enabled; } catch (eH) { }
      try { rbAutoSizeArea.enabled = enabled; } catch (eA) { }

      // When height mode is ON, disable font-size processing checkboxes and ignore their values
      if (enabled) {
        _prevFontOverflow = cbOverflow.value;
        _prevFontFit = cbFit.value;
        cbOverflow.value = false;
        cbFit.value = false;
        try { cbOverflow.enabled = false; } catch (eO1) { }
        try { cbFit.enabled = false; } catch (eF1) { }
        // When height mode is enabled, select "高さを調整" by default
        rbHeight.value = true;
        rbAutoSizeArea.value = false;
      } else {
        try { cbOverflow.enabled = true; } catch (eO2) { }
        try { cbFit.enabled = true; } catch (eF2) { }
        cbOverflow.value = _prevFontOverflow;
        cbFit.value = _prevFontFit;
      }

      if (!enabled) {
        rbHeight.value = false;
        rbAutoSizeArea.value = false;
      }
    }

    cbHeightMode.onClick = updateOptionsEnabled;
    updateOptionsEnabled();

    var gBtn = dlg.add("group");
    gBtn.alignment = ["right", "center"];
    gBtn.alignChildren = ["right", "center"];
    var btnCancel = gBtn.add("button", undefined, L("btnCancel"), { name: "cancel" });
    var btnOk = gBtn.add("button", undefined, L("btnOk"), { name: "ok" });

    btnOk.onClick = function () {
      var autoSizeArea = !!rbAutoSizeArea.value;

      // Font-size modes (checkboxes)
      var doFit = !!cbFit.value;
      var doOverflow = !!cbOverflow.value;

      // AreaText height mode: map option radios to internal doFit/doOverflow
      // - 「高さを調整」 => run fitHeight (mapped to doFit)
      // - 「自動サイズ調整」 => handled by autoSizeArea early branch
      if (!!cbHeightMode.value) {
        doFit = !!rbHeight.value;
        doOverflow = false;
      }

      if (!!cbHeightMode.value) {
        if (!rbHeight.value && !autoSizeArea) {
          alert(L("alertSelectHeightOption"));
          return;
        }
      } else {
        if (!autoSizeArea && !doFit && !doOverflow) {
          alert(L("alertSelectMode"));
          return;
        }
      }

      var adjustMode = cbHeightMode.value ? "height" : "fontSize";
      dlg.close(1);
      DealWithOversetText.run(app.activeDocument, {
        doFit: doFit,
        doOverflow: doOverflow,
        adjustMode: adjustMode,
        autoSizeArea: autoSizeArea
      });
    };

    btnCancel.onClick = function () {
      dlg.close(0);
    };

    dlg.show();
  }

  showDialogAndRun();
})();