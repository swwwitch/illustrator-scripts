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
var SCRIPT_VERSION  = "v1.3.4";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2024-07-21";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-29";                   /* 更新日 / last updated */

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

  // ローカライズ（再利用パーツ） / Localization (reusable)

  /**
   * UI の言語を返す（"ja" で始まるロケールは日本語、それ以外は英語）
   * @returns {string} "ja" または "en"
   */
  function getCurrentLang() {
      return (String($.locale || "").indexOf("ja") === 0) ? "ja" : "en";
  }

  var uiLang = getCurrentLang();

  /**
   * LABELS から今の UI 言語の文言を取り出す。
   * @param {string|Object} labelRef - "dialog.title" のようなパス、または { ja, en }
   * @param {Object|Array} [placeholderValues] - { name: 値 } なら {name} を、[値, …] なら %1, %2 … を差し込む
   * @returns {string} 文言。パスが見つからなければパスの文字列、{ ja, en } が無ければ空文字
   */
  function getLabel(labelRef, placeholderValues) {
      var labelEntry = labelRef;
      if (typeof labelRef === "string") {
          var labelPathKeys = labelRef.split(".");
          labelEntry = LABELS;
          for (var i = 0; i < labelPathKeys.length && labelEntry != null; i++) {
              labelEntry = labelEntry[labelPathKeys[i]];
          }
      }
      var labelString;
      if (typeof labelEntry === "string") labelString = labelEntry;
      else if (labelEntry != null && labelEntry[uiLang] != null) labelString = labelEntry[uiLang];
      else if (labelEntry != null && labelEntry.en != null) labelString = labelEntry.en;
      else return (typeof labelRef === "string") ? labelRef : "";
      return fillLabelPlaceholders(String(labelString), placeholderValues);
  }

  /**
   * 項目名の文言の末尾にコロンを付ける（日本語は全角「：」、英語は半角「:」）
   * @param {string|Object} labelRef - getLabel と同じ
   * @param {Object|Array} [placeholderValues] - getLabel と同じ
   * @returns {string} コロン付きの文言
   */
  function labelText(labelRef, placeholderValues) {
      return getLabel(labelRef, placeholderValues) + (uiLang === "ja" ? "：" : ":");
  }

  /**
   * 「項目名：値」の1行を返す（日本語は「件数：5」、英語は「Count: 5」とコロンのあとに空白を入れる）
   * @param {string|Object} labelRef - getLabel と同じ
   * @param {string|number} value - コロンのあとに続ける値
   * @returns {string} 項目名と値をつないだ文字列
   */
  function labelValueText(labelRef, value) {
      return labelText(labelRef) + (uiLang === "ja" ? "" : " ") + value;
  }

  /**
   * 文言の {name} や %1 に値を差し込む
   * @param {string} labelString - 文言
   * @param {Object|Array} [placeholderValues] - { name: 値 } または [値, …]
   * @returns {string} 差し込んだ文言
   */
  function fillLabelPlaceholders(labelString, placeholderValues) {
      if (placeholderValues == null) return labelString;
      if (placeholderValues instanceof Array) {
          /* 大きい番号から置き換え、%1 が %10 の一部を置き換えないようにする / Replace from the highest index so %1 does not eat into %10 */
          for (var i = placeholderValues.length; i >= 1; i--) {
              labelString = labelString.split("%" + i).join(String(placeholderValues[i - 1]));
          }
          return labelString;
      }
      for (var placeholderKey in placeholderValues) {
          if (!placeholderValues.hasOwnProperty(placeholderKey)) continue;
          labelString = labelString.split("{" + placeholderKey + "}").join(String(placeholderValues[placeholderKey]));
      }
      return labelString;
  }

  // ローカライズ（再利用パーツ）ここまで / End of the reusable localization

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
    var nameHex = toActionHex(actionName);
    return '/name [ ' + (nameHex.length / 2) + ' ' + nameHex + ' ]\n';
  }

  /*
    ustring 値（[ バイト数 UTF-8のhex ]）を生成する
    Build a ustring value ([ byteCount UTF-8 hex ]); handles multi-byte names
  */
  function buildUstringValue(sourceText) {
    var valueHex = toActionHex(sourceText);
    return '[ ' + (valueHex.length / 2) + ' ' + valueHex + ' ]';
  }

  // =========================================
  // 一時アクション実行 / Temporary action playback
  // =========================================

  // 一時アクション（再利用パーツ） / Temporary action (reusable)

  /**
   * 文字列を UTF-8 のバイト列の16進にする（アクション定義の /name・/localizedName 用）
   * @param {string} sourceText - 変換する文字列
   * @returns {string} 16進の文字列（2文字で1バイト）
   */
  function toActionHex(sourceText) {
      var utf8Text = unescape(encodeURIComponent(String(sourceText)));
      var hexText = "";
      for (var i = 0; i < utf8Text.length; i++) {
          var hexByte = utf8Text.charCodeAt(i).toString(16);
          hexText += (hexByte.length < 2 ? "0" : "") + hexByte;
      }
      return hexText;
  }

  /**
   * アクション定義の「/name [ バイト数 16進 ]」の3行を返す
   * @param {string} indent - 行頭の字下げ（"\t" など）
   * @param {string} nameText - 名前
   * @param {string} [fieldName] - 項目名（既定は "name"。"localizedName" など）
   * @returns {string[]} 3行ぶんの配列
   */
  function buildActionNameLines(indent, nameText, fieldName) {
      var nameHex = toActionHex(nameText);
      return [
          indent + "/" + (fieldName || "name") + " [ " + (nameHex.length / 2),
          indent + "\t" + nameHex,
          indent + "]"
      ];
  }

  /**
   * アクション定義を一時ファイルに書き出してセットを読み込む。読み込んだら一時ファイルは消す
   * （読み込んだ時点で解釈済みなので、以降の失敗でファイルが残らない）
   * @param {string} actionSource - アクション定義のテキスト
   * @param {string} setName - アクションセット名
   * @returns {boolean} 読み込めたら true
   */
  function loadTemporaryActionSet(actionSource, setName) {
      var actionFile = new File(Folder.temp + "/" + setName + "_" + new Date().getTime() + ".aia");
      try {
          actionFile.encoding = "UTF-8";
          if (!actionFile.open("w")) throw new Error("cannot open " + actionFile.fsName);
          actionFile.write(actionSource);
          actionFile.close();
          /* 前回の失敗で同じ名前のセットが残っていれば外す / Remove a same-name set left by an earlier failure */
          unloadTemporaryActionSet(setName);
          app.loadAction(actionFile);
          return true;
      } catch (e) {
          $.writeln("loadTemporaryActionSet: " + e);
          return false;
      } finally {
          try { actionFile.close(); } catch (closeError) { /* 閉じ済み / already closed */ }
          try { actionFile.remove(); } catch (removeError) { /* 消せなくても続ける / keep going */ }
      }
  }

  /**
   * 一時アクションのセットを解除する（読み込まれていなくてもエラーにしない）
   * @param {string} setName - アクションセット名
   * @returns {void}
   */
  function unloadTemporaryActionSet(setName) {
      try {
          app.unloadAction(setName, "");
      } catch (e) {
          /* 読み込まれていない / not loaded */
      }
  }

  /**
   * アクション定義を読み込んで1回実行し、解除する。途中で失敗しても解除は必ず試みる
   * @param {string} actionSource - アクション定義のテキスト
   * @param {string} setName - アクションセット名
   * @param {string} actionName - 実行するアクション名
   * @returns {boolean} 実行できたら true
   */
  function runTemporaryAction(actionSource, setName, actionName) {
      if (!loadTemporaryActionSet(actionSource, setName)) return false;
      try {
          app.doScript(actionName, setName);
          return true;
      } catch (e) {
          $.writeln("runTemporaryAction: " + e);
          return false;
      } finally {
          unloadTemporaryActionSet(setName);
      }
  }

  // 一時アクション（再利用パーツ）ここまで / End of the reusable temporary action

  // ダイアログの位置と不透明度（再利用パーツ） / Dialog position and opacity (reusable)

  var DIALOG_OPACITY = 0.98;       /* ダイアログの不透明度 / dialog opacity */
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

  // ダイアログの位置と不透明度（再利用パーツ）ここまで / End of the reusable dialog position and opacity

  // ボタン行（再利用パーツ） / Button row (reusable)

  var BUTTON_ROW_TOP_MARGIN = 5; /* ボタン行の上の余白 / top margin of the button row */
  var BUTTON_ROW_SPACING = 10;   /* ボタンどうしの間隔 / spacing between buttons */

  /**
   * ダイアログ下部のボタン行を作る。
   * 通常は「左のグループ・伸びるスペーサー・右のグループ」、centered なら行そのものを左右中央に置く
   * @param {Window|Group|Panel} parent - 行を足す先（ふつうはダイアログ）
   * @param {Object} [rowOptions] - { centered: true } で左右中央に並べる
   * @returns {{rowGroup: Group, leftGroup: Group|null, rightGroup: Group|null}} 行と左右のグループ（centered のときは左右が null）
   */
  function addButtonRow(parent, rowOptions) {
      var isCentered = !!(rowOptions && rowOptions.centered);
      var btnRowGroup = parent.add("group");
      btnRowGroup.orientation = "row";
      btnRowGroup.margins = [0, BUTTON_ROW_TOP_MARGIN, 0, 0];
      btnRowGroup.spacing = BUTTON_ROW_SPACING;

      if (isCentered) {
          btnRowGroup.alignment = ["center", "bottom"];
          btnRowGroup.alignChildren = ["center", "center"];
          return { rowGroup: btnRowGroup, leftGroup: null, rightGroup: null };
      }

      btnRowGroup.alignment = ["fill", "bottom"];

      var btnLeftGroup = btnRowGroup.add("group");
      btnLeftGroup.alignChildren = ["left", "center"];
      btnLeftGroup.spacing = BUTTON_ROW_SPACING;

      /* 余りの幅を吸って、右のグループを右端に寄せる / Absorbs the extra width so the right group sits at the right edge */
      var spacer = btnRowGroup.add("group");
      spacer.alignment = ["fill", "fill"];
      spacer.minimumSize.width = 0;

      var btnRightGroup = btnRowGroup.add("group");
      btnRightGroup.alignChildren = ["right", "center"];
      btnRightGroup.spacing = BUTTON_ROW_SPACING;

      return { rowGroup: btnRowGroup, leftGroup: btnLeftGroup, rightGroup: btnRightGroup };
  }

  // ボタン行（再利用パーツ）ここまで / End of the reusable button row

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

    var buttonRow = addButtonRow(dialog);

    var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("cancel"), { name: "cancel" });
    var btnOff = buttonRow.rightGroup.add("button", undefined, getLabel("buttonOff"));
    btnOff.helpTip = getLabel("tipOff");
    var btnOn = buttonRow.rightGroup.add("button", undefined, getLabel("buttonOn"), { name: "ok" });
    btnOn.helpTip = getLabel("tipOn");

    var chosenMode = null;
    btnOn.onClick = function () { chosenMode = "on"; dialog.close(); };
    btnOff.onClick = function () { chosenMode = "off"; dialog.close(); };
    btnCancel.onClick = function () { chosenMode = null; dialog.close(); };

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
    if (!runTemporaryAction(actionSource, ACTION_SET_NAME, actionName)) {
      alert("テンプレート属性の適用に失敗しました。\nFailed to apply the template attribute.");
    }
  }

  applyLayerTemplate();

})();
