#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

概要

選択中の 1 オブジェクトをシンボル化し、条件に一致するアイテムを一括で
そのシンボルインスタンスに置き換える Illustrator 用 JSX スクリプト。

- 選択1個を複製し、指定したシンボル名と基準点でシンボル登録する
- ダイアログでシンボル名を入力し、空欄時は OK を無効化する
- 既存シンボルと同名の場合は警告し、ダイアログを閉じずに再入力できる
- 基準点は 3×3 のラジオボタンで指定し、置換時の位置合わせにも使用する
- 選択がテキストの場合は「同じフォント・スタイル・サイズ」かつ「同一文字列」で対象を検索する
- それ以外は SmartEdit で類似オブジェクトを一括選択し、ダイナミックアクションで一時 note を付与してから SmartEdit を抜け、note を持つアイテムを収集する
- 収集後は note をクリアして書類状態を元に戻す
- 検出した各アイテムを、作成したシンボルインスタンスに置換する
- 置換後は新規シンボルインスタンス群を選択状態にする
- 置換対象が見つからない場合は、シンボル作成済みであることを通知して終了する

Overview

Illustrator JSX script that converts a single selected object into a symbol
and replaces matching items with instances of that symbol.

- Duplicates the selected item and registers it as a symbol with the specified
name and registration point
- Lets the user enter the symbol name in a dialog, disabling OK while the name is empty
- Warns when the name already exists and keeps the dialog open for correction
- Uses a 3×3 radio-button grid for the registration point and uses it again for alignment
- For TextFrames, searches for items with the same font/style/size and identical contents
- For other objects, toggles SmartEdit on to bulk-select similar items, tags them with a temporary note via a dynamic action, toggles SmartEdit off, and then collects every pageItem carrying that note
- Clears the temporary note after collection so the document state is restored
- Replaces each detected item with an instance of the created symbol
- Leaves the newly created symbol instances selected after replacement
- If no replacement targets are found, notifies the user that the symbol has still been created

*/

// =========================================
// バージョンとローカライズ / Version and localization
// =========================================

var SCRIPT_VERSION = "v1.0.0";

/* 現在の UI 言語 / Current UI language */
var currentLang = ($.locale.indexOf('ja') === 0) ? 'ja' : 'en';

/* 日英ラベル定義 / Japanese-English label definitions */
var LABELS = {
  dialogTitle: {
    ja: 'シンボル化して置換',
    en: 'Symbolize and Replace'
  },
  panelSymbolName: {
    ja: 'シンボル名',
    en: 'Symbol Name'
  },
  panelReferencePoint: {
    ja: '位置合わせの基準点',
    en: 'Alignment Reference Point'
  },
  buttonCancel: {
    ja: 'キャンセル',
    en: 'Cancel'
  },
  helpSymbolName: {
    ja: '作成するシンボル名を入力します。既存シンボルと同じ名前は使用できません。',
    en: 'Enter the name of the symbol to create. Existing symbol names cannot be used.'
  },
  helpReferencePoint: {
    ja: '元オブジェクトと置換後のシンボルインスタンスを揃える基準点です。',
    en: 'Sets the point used to align the original object and the replacement symbol instance.'
  },
  alertSkipped: {
    ja: '{count} 件はロック／非表示、または親レイヤーの状態により置換できませんでした。',
    en: '{count} item(s) could not be replaced because they or their parent layers were locked or hidden.'
  },
  alertResolveSymbolFailed: {
    ja: '作成したシンボル「{name}」を取得できなかったため、置換を中止しました。',
    en: 'Could not resolve the created symbol "{name}", so replacement was aborted.'
  }
};

/* キーから現在言語のラベルを取得（無ければ英語、それも無ければキー文字列）/ Look up a label for the current language; falls back to English then to the key */
function getLabel(key) {
  return (LABELS[key] && (LABELS[key][currentLang] || LABELS[key].en)) || key;
}

// =========================================
// 設定 / Settings
// =========================================

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

/* 一括選択した対象オブジェクトをマーキングする note 値。attachNoteToSelection 内の AIA 文字列に埋め込まれた値（74656d705f6d656d6f = "temp_memo"）と一致させる / Note value used to tag bulk-selected items. Must match the hex-encoded value (74656d705f6d656d6f = "temp_memo") embedded in the AIA string inside attachNoteToSelection */
var TEMP_NOTE_VALUE = 'temp_memo';

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

/* 選択中アイテムの note 属性に TEMP_NOTE_VALUE を書き込むダイナミックアクションを実行 / Run a dynamic action that writes TEMP_NOTE_VALUE into the note attribute of the current selection */
function attachNoteToSelection() {
  /* AIA 文字列内に "note"（アクションセット名）, "temp"（アクション名）, "temp_memo"（値）を UTF-8 16進で埋め込み済み / The AIA string already embeds "note" (set name), "temp" (action name), and "temp_memo" (value) as UTF-8 hex */
  var aiaContent = '/version 3\n' +
    '/name [ 4\n' +
    '\t6e6f7465\n' +
    ']\n' +
    '/isOpen 1\n' +
    '/actionCount 1\n' +
    '/action-1 {\n' +
    '\t/name [ 4\n' +
    '\t\t74656d70\n' +
    '\t]\n' +
    '\t/keyIndex 0\n' +
    '\t/colorIndex 0\n' +
    '\t/isOpen 1\n' +
    '\t/eventCount 1\n' +
    '\t/event-1 {\n' +
    '\t\t/useRulersIn1stQuadrant 0\n' +
    '\t\t/internalName (adobe_attributePalette)\n' +
    '\t\t/localizedName [ 12\n' +
    '\t\t\te5b19ee680a7e8a8ade5ae9a\n' +
    '\t\t]\n' +
    '\t\t/isOpen 1\n' +
    '\t\t/isOn 1\n' +
    '\t\t/hasDialog 0\n' +
    '\t\t/parameterCount 1\n' +
    '\t\t/parameter-1 {\n' +
    '\t\t\t/key 1852798053\n' +
    '\t\t\t/showInPalette 4294967295\n' +
    '\t\t\t/type (ustring)\n' +
    '\t\t\t/value [ 9\n' +
    '\t\t\t\t74656d705f6d656d6f\n' +
    '\t\t\t]\n' +
    '\t\t}\n' +
    '\t}\n' +
    '}';
  var actionFile = new File('~/ScriptAction.aia');
  actionFile.open('w');
  actionFile.write(aiaContent);
  actionFile.close();
  app.loadAction(actionFile);
  actionFile.remove();
  app.doScript('temp', 'note', false);
  app.unloadAction('note', '');
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

/* 同じフォント・スタイル・サイズかつ同一文字列のテキストフレームを取得（副作用：内部のメニューコマンドにより document.selection が書き換わる）/ Collect TextFrames matching font/style/size and exact text content (side effect: the inner menu command mutates document.selection) */
function findMatchingTextFrames(targetDocument, sourceTextContent) {
  app.executeMenuCommand('Find Text Font Family Style Size menu item');
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

  alert('シンボル「' + candidateName + '」は既に存在します。別の名前を指定してください。\n' +
    'A symbol named "' + candidateName + '" already exists. Please choose a different name.');

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

      var destinationLayer = currentItem.layer;
      if (!isReplacementLayerAvailable(destinationLayer)) {
        skippedCount++;
        continue;
      }

      newSymbolItem = createAlignedSymbolItem(
        destinationLayer,
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
    alert(LABELS.alertSkipped.ja.replace('{count}', skippedCount) + '\n' +
      LABELS.alertSkipped.en.replace('{count}', skippedCount));
  }
}

// =========================================
// ダイアログ / Dialog
// =========================================

/* パネル共通の余白 [左, 上, 右, 下] / Common panel margins [left, top, right, bottom] */
var PANEL_MARGINS = [15, 20, 15, 10];

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
      referencePointButton.helpTip = getLabel('helpReferencePoint');
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

/* シンボル化の設定ダイアログを表示 / Show the symbolize settings dialog */
function showSymbolizeDialog(targetDocument, defaultName, defaultReferencePoint) {
  var dialog = new Window('dialog', getLabel('dialogTitle') + ' ' + SCRIPT_VERSION);
  dialog.alignChildren = ['fill', 'top'];
  dialog.margins = 16;
  dialog.spacing = 12;

  /* シンボル名 / Symbol name */
  var symbolNamePanel = dialog.add('panel', undefined, getLabel('panelSymbolName'));
  setupPanel(symbolNamePanel, 4);
  var nameInput = symbolNamePanel.add('edittext', undefined, defaultName);
  nameInput.characters = 18;
  nameInput.helpTip = getLabel('helpSymbolName');
  nameInput.active = true;

  /* 基準点 / Registration point */
  var referencePointPanel = dialog.add('panel', undefined, getLabel('panelReferencePoint'));
  setupPanel(referencePointPanel, 4);
  referencePointPanel.alignChildren = 'center';
  referencePointPanel.helpTip = getLabel('helpReferencePoint');
  var referencePointButtons = createReferencePointGrid(referencePointPanel, defaultReferencePoint);

  /* OK / Cancel ボタン。OK は重複チェックでダイアログを保持できるよう name:'ok' を付けず、defaultElement で Enter に紐付け / OK / Cancel buttons. OK omits name:'ok' so duplicate-name validation can keep the dialog open; defaultElement wires Enter to it. */
  var buttonGroup = dialog.add('group');
  buttonGroup.alignment = 'right';
  buttonGroup.add('button', undefined, getLabel('buttonCancel'), { name: 'cancel' });
  var okButton = buttonGroup.add('button', undefined, 'OK');
  okButton.onClick = function () {
    var candidateName = trimWhitespace(nameInput.text);
    if (!validateDuplicateSymbolName(targetDocument, candidateName, nameInput)) {
      return;
    }
    dialog.close(1);
  };
  dialog.defaultElement = okButton;

  /* 入力中はトリム後の長さで OK の有効状態を更新（空欄時は押せない）/ Toggle OK availability live; the trimmed name length must be > 0 */
  nameInput.onChanging = function () {
    updateOkButtonState(nameInput, okButton);
  };
  updateOkButtonState(nameInput, okButton);

  if (dialog.show() !== 1) { return null; }
  return {
    symbolName: trimWhitespace(nameInput.text),
    referencePoint: referencePointButtons.selectedIndex
  };
}

// =========================================
// メイン処理 / Main flow
// =========================================

(function main() {
  if (app.documents.length <= 0) { return; }
  var activeDoc = app.activeDocument;
  if (activeDoc.selection.length !== 1) { return; }

  /* 選択 1 個の typename を保存（テキストかどうかで後段の検索方法を分岐）/ Capture the selected typename to branch the search strategy below */
  var sourceItem = activeDoc.selection[0];
  var isTextSelection = isTextFrame(sourceItem);
  var sourceTextContent = isTextSelection ? sourceItem.contents : null;

  /* 既定値（ダイアログの結果で上書きされる）。テキストならその文字列をシンボル名の初期値に / Defaults overridden by the dialog result; TextFrames seed the symbol name from their content */
  var symbolName = (isTextSelection && sourceTextContent) ? sanitizeTextForSymbolName(sourceTextContent) : DEFAULT_SYMBOL_NAME;
  var referencePoint = AiReferencePoint.CENTER;

  /* ダイアログで設定し、複製をシンボル化（重複チェックはダイアログ内で実施済み）/ Configure via dialog and symbolize a duplicate (duplicate-name check is done inside the dialog) */
  var dialogResult = showSymbolizeDialog(activeDoc, symbolName, referencePoint);
  if (!dialogResult) { return; }
  symbolName = dialogResult.symbolName;
  referencePoint = dialogResult.referencePoint;
  var createdSymbol = createSymbolFromSelection(activeDoc, symbolName, referencePoint);

  /* テキストは「同じフォント・サイズ・スタイル」かつ「同じ文字列」で検索、それ以外は SmartEdit + 属性アクションで selection を拡張 / TextFrames: same font/style/size + same content. Others: SmartEdit + attribute action expands document.selection */
  var replaceTargets;
  if (isTextSelection) {
    replaceTargets = findMatchingTextFrames(activeDoc, sourceTextContent);
  } else {
    /* SmartEdit を ON にしてから note を書き込む属性アクションを流すと、SmartEdit にぶら下がる類似アイテムが doc.selection に集まる。先に SmartEdit を OFF にすると選択が縮むので、ここで選択をスナップショットしてから OFF にする / Turning SmartEdit ON and then running the note-writing attribute action makes Illustrator collect every linked similar item into doc.selection. Toggling SmartEdit OFF first collapses that selection, so snapshot it before turning SmartEdit OFF */
    app.executeMenuCommand('SmartEdit Menu Item');
    attachNoteToSelection();
    var smartEditSelection = activeDoc.selection;
    replaceTargets = [];
    for (var rtIndex = 0; rtIndex < smartEditSelection.length; rtIndex++) {
      replaceTargets.push(smartEditSelection[rtIndex]);
    }
    clearNoteOnItems(replaceTargets);
    app.executeMenuCommand('SmartEdit Menu Item');
  }
  /* 元オブジェクト以外の置換対象がない場合は、シンボル作成のみで終了 / If only the original item is found, keep the created symbol and stop without replacement */
  if (replaceTargets.length <= 1) {
    try {
      createdSymbol.remove();
    } catch (removeSymbolError) { }

    alert('置換対象が見つからなかったため、作成したシンボルを削除しました。\n' +
      'No replacement targets were found, so the created symbol was removed.');
    return;
  }

  /* 置換するシンボルを取得 / Resolve the destination symbol */
  var destinationSymbol;
  try {
    destinationSymbol = activeDoc.symbols.getByName(symbolName);
  } catch (resolveSymbolError) {
    alert(LABELS.alertResolveSymbolFailed.ja.replace('{name}', symbolName) + '\n' +
      LABELS.alertResolveSymbolFailed.en.replace('{name}', symbolName));
    try {
      createdSymbol.remove();
    } catch (removeSymbolError) { }
    return;
  }

  /* 置換実行 / Replace items */
  replaceItemsWithSymbol(activeDoc, replaceTargets, destinationSymbol, referencePoint);
})();