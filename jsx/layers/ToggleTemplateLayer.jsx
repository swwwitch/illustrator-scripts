#target illustrator
#targetengine "ToggleTemplateLayerEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

アクティブレイヤーの「テンプレート」属性（ロック・印刷不可・画像を薄く表示）をON/OFFします。
小さなダイアログで切り替えを選び、ダイナミックアクションで実行します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ToggleTemplateLayer.md

### Overview

Toggles the template attribute — locked, non-printing and dimmed images — on the active layer.
A small dialog picks the direction, and the change is applied through a dynamic action.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ToggleTemplateLayer.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ToggleTemplateLayer";          /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.3.2";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2024-07-21";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-28";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ToggleTemplateLayer.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ToggleTemplateLayer.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

  // 一時アクション設定 / Temporary action settings
  // =========================================

  var ACTION_SET_NAME = "DynamicActionLayerTemplate";
  var ACTION_NAME_ON = "TemplateON";
  var ACTION_NAME_OFF = "TemplateOFF";
  var ACTION_FILE_NAME = "~/ToggleTemplateLayerAction.aia";

  /*
    レイヤーオプションのプリセット（チェックボックスの ON/OFF）
    Layer-option presets (dialog checkbox states)
  */
  var TEMPLATE_ON_OPTIONS = {
    template: true,   /* tmpl テンプレート */
    show: true,       /* show 表示 */
    lock: true,       /* lock ロック */
    preview: true,    /* prvw プレビュー */
    print: false,     /* prnt プリント */
    dim: true,        /* dim. 画像を薄く表示 */
    dimPercent: 50    /* 薄く表示の％（dim が true のときだけ書き出す） */
  };

  var TEMPLATE_OFF_OPTIONS = {
    template: false,
    show: true,
    lock: false,
    preview: true,
    print: true,
    dim: false,
    dimPercent: 50
  };

  // =========================================
  // 一時アクション生成 / Temporary action generation
  // =========================================

  /*
    記録済みの .aia から採取した internalName / key を使い、オプションから組み立てる
    Build the action from recorded internalName / keys, driven by the options preset
  */
  // =========================================
  // ローカライズ / Localization
  // =========================================

  /**
   * 現在のUI言語を判定する
   * @returns {string} "ja" または "en"
   */
  function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
  }
  var uiLang = getCurrentLang();

  /* 日英ラベル定義 / Japanese-English label definitions */
  var LABELS = {
    dialogTitle: { ja: "レイヤーテンプレート", en: "Layer Template" },
    description: {
      ja: "アクティブレイヤーのテンプレート属性を切り替えます。",
      en: "Toggles the template attribute of the active layer."
    },
    buttonOn:  { ja: "ON（テンプレート化）", en: "ON (make template)" },
    buttonOff: { ja: "OFF（解除）", en: "OFF (release)" },
    cancel:    { ja: "キャンセル", en: "Cancel" },
    tipOn: {
      ja: "アクティブレイヤーをテンプレートレイヤーにします。印刷・書き出しの対象から外れ、ロックされます。",
      en: "Makes the active layer a template layer: it is locked and left out of printing and export."
    },
    tipOff: {
      ja: "テンプレート属性を外して、通常のレイヤーに戻します。",
      en: "Clears the template attribute and returns the layer to a normal one."
    }
  };

  /**
   * ラベルを取得する
   * @param {string} key - LABELS のキー
   * @returns {string} 現在のUI言語のラベル
   */
  function getLabel(key) {
    return LABELS[key] ? LABELS[key][uiLang] : key;
  }

  function buildActionSource(setName, actionName, layerName, options) {
    var parameterLines = buildLayerParameterLines(layerName, options);

    var parameterBlock = '';
    for (var i = 0; i < parameterLines.length; i++) {
      parameterBlock += ' /parameter-' + (i + 1) + ' { ' + parameterLines[i] + ' }\n';
    }

    return ''
      + '/version 3\n'
      + buildActionNameLine(setName)
      + '/isOpen 1\n'
      + '/actionCount 1\n'
      + '/action-1 {\n'
      + ' ' + buildActionNameLine(actionName)
      + ' /keyIndex 0\n'
      + ' /colorIndex 0\n'
      + ' /isOpen 1\n'
      + ' /eventCount 1\n'
      + ' /event-1 {\n'
      + ' /useRulersIn1stQuadrant 0\n'
      + ' /internalName (ai_plugin_Layer)\n'
      + ' /localizedName [ 9 e8a1a8e7a4ba203a20 ]\n'
      + ' /isOpen 1\n'
      + ' /isOn 1\n'
      + ' /hasDialog 1\n'
      + ' /showDialog 0\n'
      + ' /parameterCount ' + parameterLines.length + '\n'
      + parameterBlock
      + ' }\n'
      + '}\n';
  }

  /*
    parameter-* を1行ずつ生成。key は記録済み .aia の FourCC（tmpl/show/lock/prvw/prnt/dim.）
    Build each parameter line; keys are FourCC from the recorded .aia
    dim が true のときだけ末尾に「薄く表示の％」行を追加する（OFF では出さない）
  */
  function buildLayerParameterLines(layerName, options) {
    var lines = [];
    lines.push('/key 1836411236 /showInPalette 4294967295 /type (integer) /value 4');                                   /* カラー（ラベル色） */
    lines.push('/key 1851878757 /showInPalette 4294967295 /type (ustring) /value [ 36 e383ace382a4e383a4e383bce38391e3838de383abe382aae38397e382b7e383a7e383b3 ]'); /* 名前ラベル */
    lines.push('/key 1953068140 /showInPalette 4294967295 /type (ustring) /value ' + buildUstringValue(layerName));     /* レイヤー名（動的注入） */
    lines.push('/key 1953329260 /showInPalette 4294967295 /type (boolean) /value ' + boolBit(options.template));        /* tmpl テンプレート */
    lines.push('/key 1936224119 /showInPalette 4294967295 /type (boolean) /value ' + boolBit(options.show));            /* show 表示 */
    lines.push('/key 1819239275 /showInPalette 4294967295 /type (boolean) /value ' + boolBit(options.lock));            /* lock ロック */
    lines.push('/key 1886549623 /showInPalette 4294967295 /type (boolean) /value ' + boolBit(options.preview));         /* prvw プレビュー */
    lines.push('/key 1886547572 /showInPalette 4294967295 /type (boolean) /value ' + boolBit(options.print));           /* prnt プリント */
    lines.push('/key 1684630830 /showInPalette 4294967295 /type (boolean) /value ' + boolBit(options.dim));             /* dim. 画像を薄く表示 */
    if (options.dim) {
      lines.push('/key 1885564532 /showInPalette 4294967295 /type (unit real) /value ' + formatDimPercent(options.dimPercent) + ' /unit 592474723'); /* 薄く表示の％ */
    }
    return lines;
  }

  function boolBit(flag) {
    return flag ? 1 : 0;
  }

  function formatDimPercent(percent) {
    return percent.toFixed(1);
  }

  function buildActionNameLine(actionName) {
    return '/name [ ' + actionName.length + ' ' + stringToHex(actionName) + ' ]\n';
  }

  function stringToHex(sourceText) {
    var hexText = "";
    for (var i = 0; i < sourceText.length; i++) {
      var hexValue = sourceText.charCodeAt(i).toString(16);
      if (hexValue.length < 2) hexValue = "0" + hexValue;
      hexText += hexValue;
    }
    return hexText;
  }

  /*
    ustring 値（[ バイト数 UTF-8のhex ]）を生成する
    Build a ustring value ([ byteCount UTF-8 hex ]); handles multi-byte names
  */
  function buildUstringValue(sourceText) {
    var byteString = unescape(encodeURIComponent(sourceText));
    var hexText = "";
    for (var i = 0; i < byteString.length; i++) {
      var hexValue = byteString.charCodeAt(i).toString(16);
      if (hexValue.length < 2) hexValue = "0" + hexValue;
      hexText += hexValue;
    }
    return '[ ' + byteString.length + ' ' + hexText + ' ]';
  }

  // =========================================
  // 一時アクション実行 / Temporary action playback
  // =========================================

  function playTemporaryAction(actionSource, setName, actionName, actionFilePath) {
    var actionFile = new File(actionFilePath);
    var isActionLoaded = false;
    var isActionFileOpen = false;

    try { app.unloadAction(setName, ""); } catch (e) { }

    try {
      if (!actionFile.open("w")) {
        throw new Error("Failed to open temporary action file.");
      }
      isActionFileOpen = true;

      actionFile.write(actionSource);
      actionFile.close();
      isActionFileOpen = false;

      app.loadAction(actionFile);
      isActionLoaded = true;

      app.doScript(actionName, setName, false);

    } catch (e) {
      alert("テンプレート属性の適用に失敗しました。\nFailed to apply the template attribute.\n\n" + e);

    } finally {
      if (isActionFileOpen) {
        try { actionFile.close(); } catch (e) { }
      }

      if (actionFile.exists) {
        try { actionFile.remove(); } catch (e) { }
      }

      if (isActionLoaded) {
        try { app.unloadAction(setName, ""); } catch (e) { }
      }
    }
  }

  // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
  // ダイアログの位置と不透明度（再利用パーツ） / Dialog position and opacity (reusable)
  //
  // 【移植手順 / How to port】
  // 1. ▼〜▲ をまるごと、コピー先の IIFE 内に貼る。
  //    識別子は DIALOG_* / prepareDialogWindow / *DialogLeft* / getSelectionViewSpan の名前
  // 2. スクリプトの先頭（#target の次の行）に #targetengine "<SCRIPT_NAME>Engine" を置く。
  //    #targetengine が無いと $.global が実行ごとに消え、位置を覚えられない。すでにあればそのまま使う
  // 3. ダイアログの show() の直前で prepareDialogWindow(dialog, SCRIPT_NAME) を呼ぶ。
  //    それまでに入れた onShow / onMove / onClose はそのまま生かし、あとに位置の復元・記録をつなぐ
  //      prepareDialogWindow(mainDialog, SCRIPT_NAME);
  //      var dialogResult = mainDialog.show();
  //    同じスクリプトで複数のダイアログを開くときは、2つ目以降のキーを変える（SCRIPT_NAME + "_colorPicker" など）
  //    同じダイアログを何度も開くときも、毎回 show() の直前で呼んでよい（2回目からは選択範囲を測り直すだけ）
  // 4. 初めて開くとき（記録が無いとき）は、スクリプト側の配置（中央・オフセットなど）がそのまま効く
  // 5. 開く位置が選択中のオブジェクトに重なりそうなら左右の反対側へずらす（Illustrator のみ）。
  //    ずらした位置は記録せず、ユーザーが動かしたときだけ記録する
  // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

  var DIALOG_OPACITY = 0.97;       /* ダイアログの不透明度 / dialog opacity */
  var DIALOG_AVOID_MARGIN = 60;    /* 選択範囲の推定位置の両側に取る余裕（px）/ margin on each side of the estimated selection (px) */
  var DIALOG_AVOID_MAX_ITEMS = 100; /* 選択範囲を測るオブジェクトの上限 / max items measured for the selection bounds */

  /**
   * ダイアログの不透明度を設定し、前回閉じた位置で開いて、動かした位置を記録するようにする。
   * 開く位置が選択中のオブジェクトに重なりそうなときは、左右の反対側へずらす（Illustrator のみ）。
   * 既存の onShow / onMove / onClose は先に呼んでから、位置の復元・記録を行う。
   * @param {Window} dialog - 対象のダイアログ
   * @param {string} storageKey - 位置を覚えるキー（ふつうは SCRIPT_NAME）
   * @returns {void}
   */
  function prepareDialogWindow(dialog, storageKey) {
      /* 同じダイアログを開き直すときは、選択範囲を測り直すだけにする（ハンドラーを重ねない）
         When the same dialog is shown again, only re-measure the selection (don't stack handlers) */
      if (dialog.dialogWindowState) {
          dialog.dialogWindowState.selectionSpan = getSelectionViewSpan();
          dialog.dialogWindowState.avoidedLocation = null;
          return;
      }
      var locationKey = "__" + storageKey + "_DialogLocation";
      var previousOnShow = dialog.onShow;
      var previousOnMove = dialog.onMove;
      var previousOnClose = dialog.onClose;
      var windowState = {
          selectionSpan: getSelectionViewSpan(), /* 選択範囲は show() の前に測る / measured before show() */
          screenWidth: null,                     /* 最初に開いたときに推定する / estimated on the first show */
          avoidedLocation: null                  /* 避けるためにずらした位置（記録しない）/ location set to avoid the selection (not remembered) */
      };
      dialog.dialogWindowState = windowState;

      dialog.opacity = DIALOG_OPACITY;

      /* 今の位置を記録する / Remember the current location */
      function rememberDialogLocation() {
          var currentLocation = [dialog.location[0], dialog.location[1]];
          var avoidedLocation = windowState.avoidedLocation;
          if (avoidedLocation && currentLocation[0] === avoidedLocation[0] && currentLocation[1] === avoidedLocation[1]) return;
          $.global[locationKey] = currentLocation;
      }

      dialog.onShow = function () {
          /* 最初に開くときの既定の位置は画面の横中央なので、画面の幅を逆算できる。2回目からは前回の位置なので使い回す
             On the first show the default location is centered horizontally, which gives the screen width; reuse it afterwards */
          if (windowState.screenWidth === null) windowState.screenWidth = dialog.location[0] * 2 + dialog.bounds.width;
          if (previousOnShow) previousOnShow.apply(this, arguments);
          /* $.screens は実際の画面の大きさと合わない（Mac で 1280×524 など）ので、画面内かは判定しない
             $.screens does not match the real display (e.g. 1280x524 on a Mac), so no on-screen check */
          var savedLocation = $.global[locationKey];
          if (savedLocation) dialog.location = [savedLocation[0], savedLocation[1]];
          if (windowState.selectionSpan) {
              var avoidLeft = findDialogLeftAvoidingSelection(dialog.location[0], dialog.bounds.width, windowState.screenWidth, windowState.selectionSpan);
              if (avoidLeft !== null) {
                  dialog.location = [avoidLeft, dialog.location[1]];
                  /* 代入後の値で比べる（丸められることがある）/ Compare with the value after assignment, which may be rounded */
                  windowState.avoidedLocation = [dialog.location[0], dialog.location[1]];
              }
          }
      };
      dialog.onMove = function () {
          if (previousOnMove) previousOnMove.apply(this, arguments);
          rememberDialogLocation();
      };
      dialog.onClose = function () {
          rememberDialogLocation();
          /* false を返すと閉じるのを取りやめるので、戻り値は元の onClose のものを返す
             Returning false cancels the close, so pass the original onClose result through */
          if (previousOnClose) return previousOnClose.apply(this, arguments);
      };
  }

  /**
   * 選択中のオブジェクトが、ドキュメントの表示域の左端から画面上で何 px の範囲にあるかを返す。
   * @returns {{left: number, right: number, viewWidth: number}|null} 選択が無い・測れないときは null
   */
  function getSelectionViewSpan() {
      try {
          if (app.name !== "Adobe Illustrator" || !app.documents.length) return null;
          var targetDoc = app.activeDocument;
          var selectedItems = targetDoc.selection;
          if (!selectedItems || !selectedItems.length || !selectedItems[0].visibleBounds) return null;
          var itemCount = Math.min(selectedItems.length, DIALOG_AVOID_MAX_ITEMS);
          var spanLeft = Infinity;
          var spanRight = -Infinity;
          for (var i = 0; i < itemCount; i++) {
              var itemBounds = selectedItems[i].visibleBounds;
              if (itemBounds[0] < spanLeft) spanLeft = itemBounds[0];
              if (itemBounds[2] > spanRight) spanRight = itemBounds[2];
          }
          var activeView = targetDoc.activeView; /* 複数ウィンドウで開いていても今のウィンドウ / the current window even with multiple windows */
          var viewBounds = activeView.bounds;
          var zoom = activeView.zoom;
          var viewWidth = (viewBounds[2] - viewBounds[0]) * zoom;
          /* 表示域の外にはみ出した部分は数えない / Ignore the part outside the view */
          var left = Math.max(0, (spanLeft - viewBounds[0]) * zoom);
          var right = Math.min(viewWidth, (spanRight - viewBounds[0]) * zoom);
          if (right <= left) return null;
          return { left: left, right: right, viewWidth: viewWidth };
      } catch (e) {
          /* テキスト編集中など測れないときは避けない / Do not avoid when it cannot be measured, e.g. while editing text */
          return null;
      }
  }

  /**
   * ダイアログが選択範囲に重なるなら、重ならない左端の位置を返す。
   * 表示域は画面の横中央にあるとみなし、ずれは DIALOG_AVOID_MARGIN で吸収する。
   * @param {number} dialogLeft - 今のダイアログの左端
   * @param {number} dialogWidth - ダイアログの幅
   * @param {number} screenWidth - 画面の幅
   * @param {{left: number, right: number, viewWidth: number}} selectionSpan - getSelectionViewSpan() の結果
   * @returns {number|null} ずらした左端。重ならない・どちらにも収まらないときは null
   */
  function findDialogLeftAvoidingSelection(dialogLeft, dialogWidth, screenWidth, selectionSpan) {
      var viewLeft = (screenWidth - selectionSpan.viewWidth) / 2;
      var avoidLeft = viewLeft + selectionSpan.left - DIALOG_AVOID_MARGIN;
      var avoidRight = viewLeft + selectionSpan.right + DIALOG_AVOID_MARGIN;
      if (dialogLeft + dialogWidth <= avoidLeft || dialogLeft >= avoidRight) return null;

      var leftSideLeft = avoidLeft - dialogWidth;   /* 選択範囲の左に置くとき / placed left of the selection */
      var rightSideLeft = avoidRight;               /* 選択範囲の右に置くとき / placed right of the selection */
      var fitsLeft = leftSideLeft >= 0;
      var fitsRight = rightSideLeft + dialogWidth <= screenWidth;
      /* 選択範囲が画面の右寄りなら左へ、左寄りなら右へ逃がす / Move away from the side the selection leans to */
      var preferLeft = (avoidLeft + avoidRight) / 2 > screenWidth / 2;
      if (preferLeft && fitsLeft) return leftSideLeft;
      if (fitsRight) return rightSideLeft;
      if (fitsLeft) return leftSideLeft;
      return null;
  }

  // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
  // ダイアログの位置と不透明度（再利用パーツ）ここまで / End of the reusable dialog position and opacity
  // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

  // =========================================
  // ダイアログ / Dialog
  // =========================================

  /*
    ON / OFF を選ぶ小ダイアログ。戻り値は "on" / "off" / null（キャンセル）
    Small dialog to choose ON / OFF; returns "on" / "off" / null (cancel)
  */
  function chooseTemplateMode() {
    var dialog = new Window("dialog", getLabel("dialogTitle") + " " + SCRIPT_VERSION);
    dialog.orientation = "column";
    dialog.alignChildren = "fill";
    dialog.margins = 16;
    dialog.spacing = 12;

    dialog.add("statictext", undefined, getLabel("description"));

    var buttonGroup = dialog.add("group");
    buttonGroup.alignment = "right";

    var cancelButton = buttonGroup.add("button", undefined, getLabel("cancel"), { name: "cancel" });
    var offButton = buttonGroup.add("button", undefined, getLabel("buttonOff"));
    offButton.helpTip = getLabel("tipOff");
    var onButton = buttonGroup.add("button", undefined, getLabel("buttonOn"), { name: "ok" });
    onButton.helpTip = getLabel("tipOn");

    var chosenMode = null;
    onButton.onClick = function () { chosenMode = "on"; dialog.close(); };
    offButton.onClick = function () { chosenMode = "off"; dialog.close(); };
    cancelButton.onClick = function () { chosenMode = null; dialog.close(); };

    prepareDialogWindow(dialog, SCRIPT_NAME);
    dialog.show();
    return chosenMode;
  }

  // =========================================
  // メイン処理 / Main
  // =========================================

  function applyLayerTemplate() {
    if (app.documents.length === 0) {
      alert("ドキュメントが開かれていません。\nNo document is open.");
      return;
    }

    var mode = chooseTemplateMode();
    if (!mode) return; /* キャンセル / Cancelled */

    var activeLayer = app.activeDocument.activeLayer;

    /*
      ロック確認はテンプレート化（ON）のときのみ。
      OFF はテンプレート（＝ロック済み）を解除する用途なので、ロックは許容する。
      Lock check applies only to ON; OFF intentionally allows locked layers (templates are locked).
    */
    if (mode === "on" && activeLayer.locked) {
      alert("アクティブレイヤーがロックされているため、実行できません。\nThe active layer is locked.");
      return;
    }

    /* 非表示レイヤーは ON/OFF とも対象外 / Hidden layers are skipped in both modes */
    if (!activeLayer.visible) {
      alert("アクティブレイヤーが非表示のため、実行できません。\nThe active layer is hidden.");
      return;
    }

    /* 実行前にアクティブレイヤー名を取得して parameter-3 に注入（リネーム防止） */
    /* Capture the active layer name and inject it into parameter-3 (avoid renaming) */
    var layerName = activeLayer.name;

    var options = (mode === "on") ? TEMPLATE_ON_OPTIONS : TEMPLATE_OFF_OPTIONS;
    var actionName = (mode === "on") ? ACTION_NAME_ON : ACTION_NAME_OFF;

    var actionSource = buildActionSource(ACTION_SET_NAME, actionName, layerName, options);
    playTemporaryAction(actionSource, ACTION_SET_NAME, actionName, ACTION_FILE_NAME);
  }

  applyLayerTemplate();

})();
