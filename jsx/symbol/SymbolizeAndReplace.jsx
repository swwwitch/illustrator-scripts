#target illustrator
#targetengine "SymbolizeAndReplaceEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択オブジェクトをシンボル化し、書類内の一致するアイテムをそのシンボルインスタンスへまとめて置き換えます。
シンボル名と基準点はダイアログで指定でき、テキスト選択時は同じフォント・スタイル・文字列のフレームを対象にします。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SymbolizeAndReplace.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n650a4b91329d

### Overview

Registers the selection as a symbol and replaces every matching item in the document with an instance of it.
The symbol name and registration point are set in a dialog; with text selected, frames sharing the same font, style and string are targeted.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SymbolizeAndReplace.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SymbolizeAndReplace";          /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.5";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-30";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SymbolizeAndReplace.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SymbolizeAndReplace.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n650a4b91329d"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    /* 3×3 基準点の単一テーブル（行優先：上行 → 中行 → 下行）。AiReferencePoint と SymbolRegistrationPoint への対応をここから派生させる / Single source of truth for the 3×3 reference points (row-major: top → middle → bottom). Both AiReferencePoint and SymbolRegistrationPoint mappings are derived from this table */
    var REFERENCE_POINTS = [
      { key: 'TOP_LEFT', symbolPoint: SymbolRegistrationPoint.SYMBOLTOPLEFTPOINT },
      { key: 'TOP_MIDDLE', symbolPoint: SymbolRegistrationPoint.SYMBOLTOPMIDDLEPOINT },
      { key: 'TOP_RIGHT', symbolPoint: SymbolRegistrationPoint.SYMBOLTOPRIGHTPOINT },
      { key: 'MIDDLE_LEFT', symbolPoint: SymbolRegistrationPoint.SYMBOLMIDDLELEFTPOINT },
      { key: 'CENTER', symbolPoint: SymbolRegistrationPoint.SYMBOLCENTERPOINT },
      { key: 'MIDDLE_RIGHT', symbolPoint: SymbolRegistrationPoint.SYMBOLMIDDLERIGHTPOINT },
      { key: 'BOTTOM_LEFT', symbolPoint: SymbolRegistrationPoint.SYMBOLBOTTOMLEFTPOINT },
      { key: 'BOTTOM_MIDDLE', symbolPoint: SymbolRegistrationPoint.SYMBOLBOTTOMMIDDLEPOINT },
      { key: 'BOTTOM_RIGHT', symbolPoint: SymbolRegistrationPoint.SYMBOLBOTTOMRIGHTPOINT }
    ];

    /* キー名 → インデックス（0-8）の別名を REFERENCE_POINTS から派生 / Derive key-name → index (0-8) aliases from REFERENCE_POINTS */
    var AiReferencePoint = {};
    for (var rpIndex = 0; rpIndex < REFERENCE_POINTS.length; rpIndex++) {
      AiReferencePoint[REFERENCE_POINTS[rpIndex].key] = rpIndex;
    }

    /* 既定のシンボル名（ダイアログ初期値）。空欄なら OK が無効化される / Default symbol name shown in the dialog (empty = OK disabled until typed) */
    var DEFAULT_SYMBOL_NAME = '';

    /* 一括選択した対象オブジェクトをマーキングする note 値。buildActionSource がこの文字列を 16進化して .aia に埋め込むため、手動での一致合わせは不要 / Note value used to tag bulk-selected items; buildActionSource hex-encodes this string into the .aia, so no manual matching is needed */
    var TEMP_NOTE_VALUE = 'temp_memo';

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

    /* 日英ラベル定義（カテゴリ別。参照はドット区切りキーで getLabel('dialog.title') のように取得）/ Japanese-English label definitions grouped by category; access with dotted keys such as getLabel('dialog.title') */
    var LABELS = {
      dialog: {
        title: { ja: 'シンボル化して置換', en: 'Symbolize and Replace' }
      },
      panel: {
        symbolName: { ja: 'シンボル名', en: 'Symbol Name' },
        referencePoint: { ja: 'シンボルの基準点', en: 'Symbol Reference Point' },
        text: { ja: 'テキスト', en: 'Text' }
      },
      checkbox: {
        allowSizeMismatch: { ja: 'フォントサイズ違いも対象にする', en: 'Include different font sizes' }
      },
      button: {
        cancel: { ja: 'キャンセル', en: 'Cancel' }
      },
      help: {
        symbolName: {
          ja: '新しく作成するシンボルの名前です。空欄、および既存シンボルと同じ名前は使用できません。',
          en: 'Name of the new symbol. Empty names and names already used by existing symbols are not allowed.'
        },
        referencePoint: {
          ja: 'シンボル登録時の基準点です。置換時もこの点を使って元オブジェクトの位置に揃えます。',
          en: 'Sets the registration point for the new symbol. The same point is used to align each replacement instance to the original object.'
        },
        allowSizeMismatch: {
          ja: 'ON のときは、フォント・スタイル・文字列が一致していれば、フォントサイズが異なるテキストフレームも置換対象に含めます。',
          en: 'When enabled, TextFrames with matching font, style, and contents are included even if their font size differs.'
        }
      },
      alert: {
        skipped: {
          ja: '{count} 件はロック／非表示、または親レイヤーの状態により置換できませんでした。',
          en: '{count} item(s) could not be replaced because they or their parent layers were locked or hidden.'
        },
        duplicateSymbol: {
          ja: 'シンボル「{name}」は既に存在します。別の名前を指定してください。',
          en: 'A symbol named "{name}" already exists. Please choose a different name.'
        },
        noTargets: {
          ja: '置換対象が見つからなかったため、作成したシンボルを削除しました。',
          en: 'No replacement targets were found, so the created symbol was removed.'
        },
        multiSelectionNotGroups: {
          ja: '複数選択時は、すべてのアイテムがグループである必要があります。グループ以外を含む選択では実行できません。',
          en: 'When multiple items are selected, every item must be a group. The script cannot run if the selection includes non-group items.'
        }
      }
    };

    /* ドット区切りキー（例 "dialog.title"）で LABELS の { ja, en } エントリを辿る。見つからなければ null / Resolve a dotted key (e.g. "dialog.title") to its { ja, en } entry in LABELS; null when not found */
    function resolveLabelEntry(key) {
      var parts = String(key).split('.');
      var node = LABELS;
      for (var i = 0; i < parts.length; i++) {
        if (!node || typeof node !== 'object') { return null; }
        node = node[parts[i]];
      }
      return node || null;
    }

    /* 同じキーの ja / en メッセージを改行でまとめて返すユーティリティ。{key:value} のプレースホルダ置換に対応 / Build a JA+EN bilingual message from a LABELS key, supporting {key:value} placeholders */
    function buildBilingualMessage(key, replacements) {
      var entry = resolveLabelEntry(key);
      if (!entry) { return key; }
      var jaText = entry.ja;
      var enText = entry.en;
      if (replacements) {
        for (var placeholderKey in replacements) {
          if (!replacements.hasOwnProperty(placeholderKey)) { continue; }
          var token = '{' + placeholderKey + '}';
          jaText = jaText.split(token).join(replacements[placeholderKey]);
          enText = enText.split(token).join(replacements[placeholderKey]);
        }
      }
      return jaText + '\n' + enText;
    }

    // =========================================
    // ヘルパー関数 / Helper functions
    // =========================================

    /* AiReferencePoint インデックスを REFERENCE_POINTS 経由で SymbolRegistrationPoint に変換。範囲外は CENTER にフォールバック / Map an AiReferencePoint index to SymbolRegistrationPoint via REFERENCE_POINTS; out-of-range falls back to CENTER */
    function toSymbolRegistrationPoint(referencePointIndex) {
      if (referencePointIndex < 0 || referencePointIndex >= REFERENCE_POINTS.length) {
        return SymbolRegistrationPoint.SYMBOLCENTERPOINT;
      }
      return REFERENCE_POINTS[referencePointIndex].symbolPoint;
    }

    /* 変形パレットの要領で基準点に対応する座標を返す / Return the coordinate matching a reference point (Transform-palette style) */
    function getReferencePointPosition(targetItem, referencePoint) {
      // AiReferencePoint は 0-8 の 3x3 グリッド（列：左/中/右、行：上/中/下）/ AiReferencePoint is a 0-8 3x3 grid (col: L/C/R, row: T/M/B)
      var bounds = targetItem.visibleBounds;
      var col = referencePoint % 3;
      var row = Math.floor(referencePoint / 3);
      var x = col === 0 ? bounds[0] : col === 2 ? bounds[2] : (bounds[0] + bounds[2]) / 2;
      var y = row === 0 ? bounds[1] : row === 2 ? bounds[3] : (bounds[1] + bounds[3]) / 2;
      return [x, y];
    }

    // =========================================
    // 一時アクション設定 / Temporary action settings
    // =========================================

    /* 一括選択用ダイナミックアクションのアクションセット名・アクション名。セット名／アクション名は任意の内部ラベルで、記録済みの内容と一致する必要はない（doScript と /name 生成で同じ定数を使うため整合する）/ Action-set name and action name for the bulk-select dynamic action. The set/action names are arbitrary internal labels and need not match any recording (doScript and the /name lines use the same constants, so they stay consistent) */
    var ACTION_SET_NAME = 'SymbolizeAndReplaceNote';
    var ACTION_NAME = 'AttachTempNote';

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

    // =========================================
    // 一時アクション生成 / Temporary action generation
    // =========================================

    /* 選択アイテムの note 属性に値を書き込む .aia アクションソースを生成する。internalName(adobe_attributePalette) と parameter-1(/key 1852798053, /type ustring) は実際に記録した .aia から採取した値なので勘で書き換えないこと。note 値だけを引数から差し込む / Build the .aia source for an action that writes a value into the note attribute. internalName (adobe_attributePalette) and parameter-1 (/key 1852798053, /type ustring) are taken from a real recording — do not guess them; only the note value is injected from the argument */
    function buildActionSource(setName, actionName, noteValue) {
      return ''
        + '/version 3\n'
        + buildActionNameLine(setName)
        + '/isOpen 1\n'
        + '/actionCount 1\n'
        + '/action-1 {\n'
        + buildActionNameLine(actionName)
        + ' /keyIndex 0\n'
        + ' /colorIndex 0\n'
        + ' /isOpen 1\n'
        + ' /eventCount 1\n'
        + ' /event-1 {\n'
        + ' /useRulersIn1stQuadrant 0\n'
        + ' /internalName (adobe_attributePalette)\n'
        + ' /localizedName [ 0  ]\n'
        + ' /isOpen 1\n'
        + ' /isOn 1\n'
        + ' /hasDialog 0\n'
        + ' /parameterCount 1\n'
        + ' /parameter-1 { /key 1852798053 /showInPalette 4294967295 /type (ustring) /value ' + buildHexByteArray(noteValue) + ' }\n'
        + ' }\n'
        + '}\n';
    }

    /* アクション名／セット名を .aia の /name 行に変換 / Convert a name to the .aia /name line */
    function buildActionNameLine(actionName) {
      return '/name ' + buildHexByteArray(actionName) + '\n';
    }

    /* 文字列を .aia の "[ バイト長 16進 ]" 形式に変換 / Convert a string to the .aia "[ byteLength hex ]" form */
    function buildHexByteArray(sourceText) {
      var hexText = toActionHex(sourceText);
      return '[ ' + (hexText.length / 2) + ' ' + hexText + ' ]';
    }

    /* 現在の選択アイテムの note 属性に TEMP_NOTE_VALUE を書き込む一時アクションを生成・再生 / Build and play a temporary action that writes TEMP_NOTE_VALUE into the note attribute of the current selection */
    function attachNoteToSelection() {
      var actionSource = buildActionSource(ACTION_SET_NAME, ACTION_NAME, TEMP_NOTE_VALUE);
      /* 失敗しても note が付かないだけで、後段の「対象なし」の警告で終わる / On failure no note is attached and the run ends with the no-targets alert */
      runTemporaryAction(actionSource, ACTION_SET_NAME, ACTION_NAME);
    }

    /* 指定 note 値を持つ pageItem を全件収集 / Collect every pageItem whose note property equals the given value */
    function collectItemsByNote(targetDocument, noteValue) {
      var matches = [];
      var allItems = targetDocument.pageItems;
      for (var i = 0, len = allItems.length; i < len; i++) {
        try {
          if (allItems[i].note === noteValue) {
            matches.push(allItems[i]);
          }
        } catch (collectNoteError) { }
      }
      return matches;
    }

    /* 指定アイテム群の note 属性を空文字でクリア / Clear the note property on the given items */
    function clearNoteOnItems(items) {
      for (var i = 0; i < items.length; i++) {
        try { items[i].note = ''; } catch (clearNoteError) { }
      }
    }

    /* 対象が TextFrame か判定 / Whether the item is a TextFrame */
    function isTextFrame(item) {
      return !!item && item.typename === 'TextFrame';
    }

    /* 配列内の全要素が GroupItem か（空配列は false）/ Whether every item in the array is a GroupItem (false for empty arrays) */
    function areAllGroups(items) {
      if (!items || items.length === 0) { return false; }
      for (var i = 0; i < items.length; i++) {
        if (!items[i] || items[i].typename !== 'GroupItem') { return false; }
      }
      return true;
    }

    /* テキストフレームの内容と一致するか判定（typename と文字列を一括チェック）/ Whether the item is a TextFrame whose contents equal the given text */
    function isMatchingTextFrame(item, sourceTextContent) {
      return isTextFrame(item) && item.contents === sourceTextContent;
    }

    /* シンボル名向けにテキストを整形（改行を半角スペース化）/ Normalize text for use as a symbol name (collapse newlines to spaces) */
    function sanitizeTextForSymbolName(text) {
      return (text || '').replace(/[\r\n]+/g, ' ');
    }

    /* 前後の空白をトリム（ExtendScript の String.prototype.trim 非対応に備えた手動実装）/ Trim leading/trailing whitespace (manual implementation for ExtendScript compatibility) */
    function trimWhitespace(text) {
      return (text || '').replace(/^\s+|\s+$/g, '');
    }

    /* 同じフォント・スタイル（オプションでサイズ）かつ同一文字列のテキストフレームを取得。allowSizeMismatch=true のときはサイズを無視するメニューコマンドを使用する。副作用：内部のメニューコマンドにより document.selection が書き換わる / Collect TextFrames matching font/style (and optionally size) plus exact text content; when allowSizeMismatch is true, uses the size-agnostic menu command. Side effect: the inner menu command mutates document.selection */
    function findMatchingTextFrames(targetDocument, sourceTextContent, allowSizeMismatch) {
      var menuCommand = allowSizeMismatch
        ? 'Find Text Font Family Style menu item'
        : 'Find Text Font Family Style Size menu item';
      app.executeMenuCommand(menuCommand);
      var matches = [];
      var currentSelection = targetDocument.selection;
      for (var i = 0; i < currentSelection.length; i++) {
        if (isMatchingTextFrame(currentSelection[i], sourceTextContent)) {
          matches.push(currentSelection[i]);
        }
      }
      return matches;
    }

    /* 指定名のシンボルが既に存在するか / Whether a symbol with the given name already exists */
    function symbolNameExists(targetDocument, name) {
      try {
        targetDocument.symbols.getByName(name);
        return true;
      } catch (e) {
        return false;
      }
    }

    /* シンボル名の重複を検証し、重複時は警告を表示 / Validate duplicate symbol names and alert when duplicated */
    function validateDuplicateSymbolName(targetDocument, candidateName, nameInput) {
      if (!symbolNameExists(targetDocument, candidateName)) {
        return true;
      }

      alert(buildBilingualMessage('alert.duplicateSymbol', { name: candidateName }));

      nameInput.active = true;
      return false;
    }

    /* 選択 1 個を複製してシンボル化、複製インスタンスは削除 / Duplicate the selection, convert it to a symbol, then remove the duplicate instance */
    function createSymbolFromSelection(targetDocument, symbolName, referencePoint) {
      var sourceItem = targetDocument.selection[0];
      var duplicatedItem = sourceItem.duplicate();
      var createdSymbol = targetDocument.symbols.add(duplicatedItem, toSymbolRegistrationPoint(referencePoint));
      createdSymbol.name = symbolName;
      duplicatedItem.remove();

      /* 後続の検索・SmartEdit が元選択を参照できるよう、元オブジェクトを選択し直す / Reselect the original item so the following search or SmartEdit can use it */
      targetDocument.selection = null;
      sourceItem.selected = true;
      return createdSymbol;
    }

    /* 置換先として使えるレイヤーか判定（親レイヤーも確認）/ Whether the layer can be used for replacement, including parent layers */
    function isReplacementLayerAvailable(targetLayer) {
      try {
        var currentLayer = targetLayer;
        while (currentLayer && currentLayer.typename === 'Layer') {
          if (currentLayer.locked || !currentLayer.visible) {
            return false;
          }
          currentLayer = currentLayer.parent;
        }
        return true;
      } catch (layerStateError) {
        return false;
      }
    }

    /* 置換対象として処理できるアイテムか判定 / Whether the item can be processed as a replacement target */
    function isReplacementItemAvailable(targetItem) {
      try {
        if (!targetItem || targetItem.locked || targetItem.hidden) {
          return false;
        }
        return isReplacementLayerAvailable(targetItem.layer);
      } catch (itemStateError) {
        return false;
      }
    }

    /* シンボルインスタンスを作成し、指定基準点同士が揃うよう移動 / Create a symbol instance and align the specified reference points */
    function createAlignedSymbolItem(destinationLayer, destinationSymbol, targetItem, referencePoint) {
      var destinationPosition = getReferencePointPosition(targetItem, referencePoint);
      var newSymbolItem = destinationLayer.symbolItems.add(destinationSymbol);
      var sourcePosition = getReferencePointPosition(newSymbolItem, referencePoint);
      newSymbolItem.translate(destinationPosition[0] - sourcePosition[0], destinationPosition[1] - sourcePosition[1]);
      return newSymbolItem;
    }

    /* 対象を順次シンボルインスタンスに置換 / Replace each target with an instance of the given symbol */
    function replaceItemsWithSymbol(targetDocument, targetItems, destinationSymbol, referencePoint) {
      var createdSymbolItems = [];
      var skippedCount = 0;

      for (var i = 0; i < targetItems.length; i++) {
        var currentItem = targetItems[i];
        var newSymbolItem = null;

        try {
          if (!isReplacementItemAvailable(currentItem)) {
            skippedCount++;
            continue;
          }

          newSymbolItem = createAlignedSymbolItem(
            currentItem.layer,
            destinationSymbol,
            currentItem,
            referencePoint
          );
          currentItem.remove();
          createdSymbolItems.push(newSymbolItem);
        } catch (replaceItemError) {
          if (newSymbolItem) {
            try {
              newSymbolItem.remove();
            } catch (removeNewSymbolItemError) { }
          }
          skippedCount++;
        }
      }

      // 選択を 1 個ずつ .selected=true で立てると多数選択時に固まるため、配列で一括代入する / Bulk assignment avoids the per-item freeze that .selected = true triggers on large counts
      targetDocument.selection = createdSymbolItems;

      if (skippedCount > 0) {
        alert(buildBilingualMessage('alert.skipped', { count: skippedCount }));
      }
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /* パネル共通の余白 [左, 上, 右, 下] / Common panel margins [left, top, right, bottom] */
    var PANEL_MARGINS = [16, 20, 16, 12];

    /* パネルの共通設定（縦並び・左寄せ・横一杯）/ Common panel setup (column orientation, left-aligned, fills width) */
    function setupPanel(panel, spacing) {
      panel.orientation = 'column';
      panel.alignChildren = 'left';
      panel.alignment = 'fill';
      panel.margins = PANEL_MARGINS;
      if (typeof spacing === 'number') {
        panel.spacing = spacing;
      }
    }

    /* 基準点ラジオボタンを排他的に選択 / Select one reference-point radio button exclusively */
    function selectReferencePointButton(radioButtons, selectedIndex) {
      for (var i = 0; i < radioButtons.length; i++) {
        radioButtons[i].value = (i === selectedIndex);
      }
      radioButtons.selectedIndex = selectedIndex;
    }

    /* 基準点 3x3 ラジオボタンを生成（手動排他、選択値は radioButtons.selectedIndex に保持）/ Build a 3x3 radio grid with manual mutual exclusion; the selected index is exposed as radioButtons.selectedIndex */
    function createReferencePointGrid(parentPanel, defaultIndex) {
      var radioButtons = [];
      for (var row = 0; row < 3; row++) {
        var rowGroup = parentPanel.add('group');
        rowGroup.orientation = 'row';
        rowGroup.spacing = 4;
        for (var col = 0; col < 3; col++) {
          var referencePointButton = rowGroup.add('radiobutton', undefined, '');
          referencePointButton.helpTip = getLabel('help.referencePoint');
          radioButtons.push(referencePointButton);
        }
      }
      selectReferencePointButton(radioButtons, defaultIndex);
      for (var i = 0; i < radioButtons.length; i++) {
        (function (buttonIndex) {
          radioButtons[buttonIndex].onClick = function () {
            selectReferencePointButton(radioButtons, buttonIndex);
          };
        })(i);
      }
      return radioButtons;
    }

    /* シンボル名入力に応じて OK ボタンの有効状態を更新 / Update OK button availability from the symbol-name input */
    function updateOkButtonState(nameInput, okButton) {
      okButton.enabled = trimWhitespace(nameInput.text).length > 0;
    }

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
            /* 文字ツールで文字を選択しているときは TextRange が返り、[0] が無い / Selecting characters with the Type tool returns a TextRange, which has no [0] */
            if (!selectedItems || selectedItems.typename === "TextRange" || !selectedItems.length || !selectedItems[0].visibleBounds) return null;
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

    /* シンボル化の設定ダイアログを表示 / Show the symbolize settings dialog */
    function showSymbolizeDialog(targetDocument, defaultName, defaultReferencePoint, isTextSelection) {
      var dialog = new Window('dialog', getLabel('dialog.title') + ' ' + SCRIPT_VERSION);
      dialog.alignChildren = ['fill', 'top'];
      dialog.margins = 16;
      dialog.spacing = 12;

      /* シンボル名 / Symbol name */
      var symbolNamePanel = dialog.add('panel', undefined, getLabel('panel.symbolName'));
      setupPanel(symbolNamePanel, 4);
      var nameInput = symbolNamePanel.add('edittext', undefined, defaultName);
      nameInput.characters = 18;
      nameInput.preferredSize.width = 210;
      nameInput.helpTip = getLabel('help.symbolName');
      nameInput.active = true;

      /* 基準点 / Registration point */
      var referencePointPanel = dialog.add('panel', undefined, getLabel('panel.referencePoint'));
      setupPanel(referencePointPanel, 4);
      referencePointPanel.alignChildren = 'center';
      referencePointPanel.helpTip = getLabel('help.referencePoint');
      var referencePointButtons = createReferencePointGrid(referencePointPanel, defaultReferencePoint);

      /* テキスト（テキスト選択時のみ。一番下に配置）/ Text options (only when a TextFrame is selected; placed at the bottom) */
      var allowSizeMismatchCheckbox = null;
      if (isTextSelection) {
        var textPanel = dialog.add('panel', undefined, getLabel('panel.text'));
        setupPanel(textPanel, 4);
        allowSizeMismatchCheckbox = textPanel.add('checkbox', undefined, getLabel('checkbox.allowSizeMismatch'));
        allowSizeMismatchCheckbox.helpTip = getLabel('help.allowSizeMismatch');
        allowSizeMismatchCheckbox.value = false;
      }

      /* OK / Cancel ボタン。OK は重複チェックでダイアログを保持できるよう name:'ok' を付けず、defaultElement で Enter に紐付け / OK / Cancel buttons. OK omits name:'ok' so duplicate-name validation can keep the dialog open; defaultElement wires Enter to it. */
      var buttonRow = addButtonRow(dialog);
      var btnCancel = buttonRow.rightGroup.add('button', undefined, getLabel('button.cancel'), { name: 'cancel' });
      var btnOK = buttonRow.rightGroup.add('button', undefined, 'OK');
      btnOK.onClick = function () {
        var candidateName = trimWhitespace(nameInput.text);
        if (!validateDuplicateSymbolName(targetDocument, candidateName, nameInput)) {
          return;
        }
        dialog.close(1);
      };
      dialog.defaultElement = btnOK;

      /* 入力中はトリム後の長さで OK の有効状態を更新（空欄時は押せない）/ Toggle OK availability live; the trimmed name length must be > 0 */
      nameInput.onChanging = function () {
        updateOkButtonState(nameInput, btnOK);
      };
      updateOkButtonState(nameInput, btnOK);

      prepareDialogWindow(dialog, SCRIPT_NAME);
      if (dialog.show() !== 1) { return null; }
      return {
        symbolName: trimWhitespace(nameInput.text),
        referencePoint: referencePointButtons.selectedIndex,
        allowSizeMismatch: allowSizeMismatchCheckbox ? !!allowSizeMismatchCheckbox.value : false
      };
    }

    // =========================================
    // メイン処理 / Main flow
    // =========================================

    (function main() {
      if (app.documents.length <= 0) { return; }
      var activeDoc = app.activeDocument;
      var rawSelection = activeDoc.selection;
      if (rawSelection.length < 1) { return; }

      /* 複数選択時は全要素が GroupItem の場合のみ対応（"選択した全グループを 1 シンボルにまとめる" モード）。それ以外の複数選択は通知して終了 / Multi-selection is only supported when every selected item is a GroupItem (the "merge all selected groups into a single symbol" mode); other multi-selection cases notify and bail out */
      var isMultiGroupSelection = false;
      if (rawSelection.length > 1) {
        if (!areAllGroups(rawSelection)) {
          alert(buildBilingualMessage('alert.multiSelectionNotGroups'));
          return;
        }
        isMultiGroupSelection = true;
      }

      /* 選択をローカル配列にコピーしておく（複数グループ選択時は createSymbolFromSelection が selection を 1 個に絞るため、参照を別途保持する必要がある）/ Snapshot the selection so references survive createSymbolFromSelection narrowing the selection down to one item */
      var sourceSelection = [];
      for (var snapshotIndex = 0; snapshotIndex < rawSelection.length; snapshotIndex++) {
        sourceSelection.push(rawSelection[snapshotIndex]);
      }

      /* 単独選択時のみ TextFrame 分岐の判定を行う（複数グループ選択時は常に false）/ TextFrame branching only applies for single selection (always false in multi-group mode) */
      var sourceItem = sourceSelection[0];
      var isTextSelection = !isMultiGroupSelection && isTextFrame(sourceItem);
      var sourceTextContent = isTextSelection ? sourceItem.contents : null;

      /* 既定値（ダイアログの結果で上書きされる）。テキストならその文字列をシンボル名の初期値に / Defaults overridden by the dialog result; TextFrames seed the symbol name from their content */
      var symbolName = (isTextSelection && sourceTextContent) ? sanitizeTextForSymbolName(sourceTextContent) : DEFAULT_SYMBOL_NAME;
      var referencePoint = AiReferencePoint.CENTER;

      /* ダイアログで設定し、複製をシンボル化（重複チェックはダイアログ内で実施済み）/ Configure via dialog and symbolize a duplicate (duplicate-name check is done inside the dialog) */
      var dialogResult = showSymbolizeDialog(activeDoc, symbolName, referencePoint, isTextSelection);
      if (!dialogResult) { return; }
      symbolName = dialogResult.symbolName;
      referencePoint = dialogResult.referencePoint;
      var allowSizeMismatch = dialogResult.allowSizeMismatch;
      var createdSymbol = createSymbolFromSelection(activeDoc, symbolName, referencePoint);

      /* 置換対象の収集：複数グループ選択時は選択そのものを使用、テキスト単独はフォント・スタイル（＋オプションでサイズ）＋文字列で検索、それ以外は SmartEdit + note 収集 / Collect replacement targets: multi-group uses the selection itself; a single TextFrame uses Find Text (with optional size match); otherwise SmartEdit + note collection */
      var replaceTargets;
      if (isMultiGroupSelection) {
        replaceTargets = sourceSelection;
      } else if (isTextSelection) {
        replaceTargets = findMatchingTextFrames(activeDoc, sourceTextContent, allowSizeMismatch);
      } else {
        /* 過去の実行が clearNoteOnItems 前にクラッシュしていた場合に備え、SmartEdit ブランチに入る前に同じ note 値を持つアイテムを一掃する / Guard against pre-existing temp_memo notes from a previous run that crashed before clearNoteOnItems */
        clearNoteOnItems(collectItemsByNote(activeDoc, TEMP_NOTE_VALUE));

        /* SmartEdit ON → ダイナミックアクションで note を付与 → note を持つアイテムを収集 → note クリア / Toggle SmartEdit on, tag matching items with a note via a dynamic action, collect every pageItem carrying that note, then clear */
        app.executeMenuCommand('SmartEdit Menu Item');
        attachNoteToSelection();
        /* attachNoteToSelection() の doScript 実行で SmartEdit は自動的に OFF に戻るため、ここで再度 'SmartEdit Menu Item' を呼ぶと逆に ON になってしまう。よって明示的なトグル OFF は入れない / Illustrator turns SmartEdit back off automatically once attachNoteToSelection()'s doScript runs, so calling 'SmartEdit Menu Item' again here would toggle it back ON — do not add an explicit off-toggle */
        replaceTargets = collectItemsByNote(activeDoc, TEMP_NOTE_VALUE);
        clearNoteOnItems(replaceTargets);
      }
      /* 元オブジェクト以外の置換対象がない場合は、作成したシンボルを削除して終了。複数グループ選択は 2 件以上が前提なので、通常この分岐には入らない / Bail out (removing the created symbol) when nothing but the original item was found. Multi-group selection always carries 2+ items, so this branch is normally not taken there */
      if (replaceTargets.length <= 1) {
        try {
          createdSymbol.remove();
        } catch (removeSymbolError) { }

        alert(buildBilingualMessage('alert.noTargets'));
        return;
      }

      /* 置換実行（createSymbolFromSelection が返したシンボルをそのまま使う）/ Replace items using the symbol returned by createSymbolFromSelection */
      replaceItemsWithSymbol(activeDoc, replaceTargets, createdSymbol, referencePoint);
    })();

})();
