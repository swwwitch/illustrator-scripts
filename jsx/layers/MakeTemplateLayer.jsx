#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要：

- アクティブレイヤーを「テンプレート」属性（ロック・印刷不可・画像を薄く表示）にする（ON 専用）
- ダイアログなしで即実行
- ダイナミックアクションで実行
- 実行前にアクティブレイヤー名を取得し、リネームせず属性のみ適用
- ロック／非表示のレイヤーには実行しない

*/

(function () {

  // =========================================
  // バージョン / Version
  // =========================================

  // スクリプトバージョン
  var SCRIPT_VERSION = "v1.0";

  // =========================================
  // 一時アクション設定 / Temporary action settings
  // =========================================

  var ACTION_SET_NAME = "DynamicActionMakeTemplate";
  var ACTION_NAME = "TemplateON";
  var ACTION_FILE_NAME = "~/MakeTemplateLayerAction.aia";

  /*
    レイヤーオプションのプリセット（テンプレート ON 固定）
    Layer-option preset (template ON)
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

  // =========================================
  // 一時アクション生成 / Temporary action generation
  // =========================================

  /*
    記録済みの .aia から採取した internalName / key を使い、オプションから組み立てる
    Build the action from recorded internalName / keys, driven by the options preset
  */
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
    dim が true のときだけ末尾に「薄く表示の％」行を追加する
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

  // =========================================
  // メイン処理 / Main
  // =========================================

  function makeTemplateLayer() {
    if (app.documents.length === 0) {
      alert("ドキュメントが開かれていません。\nNo document is open.");
      return;
    }

    var activeLayer = app.activeDocument.activeLayer;

    /* ロック・非表示のレイヤーには適用しない / Skip locked or hidden layers */
    if (activeLayer.locked) {
      alert("アクティブレイヤーがロックされているため、実行できません。\nThe active layer is locked.");
      return;
    }
    if (!activeLayer.visible) {
      alert("アクティブレイヤーが非表示のため、実行できません。\nThe active layer is hidden.");
      return;
    }

    /* 実行前にアクティブレイヤー名を取得して parameter-3 に注入（リネーム防止） */
    /* Capture the active layer name and inject it into parameter-3 (avoid renaming) */
    var layerName = activeLayer.name;

    var actionSource = buildActionSource(ACTION_SET_NAME, ACTION_NAME, layerName, TEMPLATE_ON_OPTIONS);
    playTemporaryAction(actionSource, ACTION_SET_NAME, ACTION_NAME, ACTION_FILE_NAME);
  }

  makeTemplateLayer();

})();
