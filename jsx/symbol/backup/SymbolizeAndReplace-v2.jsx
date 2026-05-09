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
- それ以外はグローバル編集（Global Edit / SmartEdit）で類似オブジェクトを検出する
- 検出した各アイテムを、作成したシンボルインスタンスに置換する
- 置換後は新規シンボルインスタンス群を選択状態にする
- 置換対象が見つからない場合は、シンボル作成済みであることを通知して終了する

参考：
sttk3.com
https://note.com/sttk3com/n/n134404369442

追加点：

- 選択オブジェクトから新規シンボルを作成する処理
- シンボル名入力ダイアログ
- 重複チェック
- 3×3 基準点 UI
- テキストフレーム専用検索
- ロック／非表示／親レイヤーのスキップ処理
- 作成済みシンボルの削除処理
- tooltip / ローカライズ / 概要コメント

Overview

Illustrator JSX script that converts a single selected object into a symbol
and replaces matching items with instances of that symbol.

- Duplicates the selected item and registers it as a symbol with the specified
name and registration point
- Lets the user enter the symbol name in a dialog, disabling OK while the name is empty
- Warns when the name already exists and keeps the dialog open for correction
- Uses a 3×3 radio-button grid for the registration point and uses it again for alignment
- For TextFrames, searches for items with the same font/style/size and identical contents
- For other objects, uses Global Edit (SmartEdit) to detect similar items
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

/* 変形基準点の別名（SymbolRegistrationPoint を参考）/ Reference point aliases (modeled after SymbolRegistrationPoint) */
var AiReferencePoint = {
  TOP_LEFT: 0,
  TOP_MIDDLE: 1,
  TOP_RIGHT: 2,
  MIDDLE_LEFT: 3,
  CENTER: 4,
  MIDDLE_RIGHT: 5,
  BOTTOM_LEFT: 6,
  BOTTOM_MIDDLE: 7,
  BOTTOM_RIGHT: 8
};

/* 既定のシンボル名（ダイアログ初期値）。空欄なら OK が無効化される / Default symbol name shown in the dialog (empty = OK disabled until typed) */
var DEFAULT_SYMBOL_NAME = '';

// =========================================
// ヘルパー関数 / Helper functions
// =========================================

/* AiReferencePoint インデックスから SymbolRegistrationPoint へのルックアップ / Lookup table from AiReferencePoint index to SymbolRegistrationPoint */
var SYMBOL_REGISTRATION_POINT_BY_INDEX = [
  SymbolRegistrationPoint.SYMBOLTOPLEFTPOINT,
  SymbolRegistrationPoint.SYMBOLTOPMIDDLEPOINT,
  SymbolRegistrationPoint.SYMBOLTOPRIGHTPOINT,
  SymbolRegistrationPoint.SYMBOLMIDDLELEFTPOINT,
  SymbolRegistrationPoint.SYMBOLCENTERPOINT,
  SymbolRegistrationPoint.SYMBOLMIDDLERIGHTPOINT,
  SymbolRegistrationPoint.SYMBOLBOTTOMLEFTPOINT,
  SymbolRegistrationPoint.SYMBOLBOTTOMMIDDLEPOINT,
  SymbolRegistrationPoint.SYMBOLBOTTOMRIGHTPOINT
];

/* AiReferencePoint を SymbolRegistrationPoint に変換 / Map AiReferencePoint to SymbolRegistrationPoint */
function toSymbolRegistrationPoint(referencePointIndex) {
  // 0-8 範囲外（不正値）は CENTER にフォールバック / Out-of-range (invalid) indices fall back to CENTER
  if (referencePointIndex < 0 || referencePointIndex > 8) {
    return SymbolRegistrationPoint.SYMBOLCENTERPOINT;
  }
  return SYMBOL_REGISTRATION_POINT_BY_INDEX[referencePointIndex];
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

/* グローバル編集で選択されている SymbolItem 群を取得 / Collect SymbolItems flagged by Global Edit */
function getGlobalEditSymbolItems(targetDocument) {
  var globalEditItems = [];
  var currentSelection = targetDocument.selection;
  if (currentSelection.length !== 1) { return globalEditItems; }

  var allSymbolItems = targetDocument.symbolItems;
  for (var i = 0, len = allSymbolItems.length; i < len; i++) {
    var currentItem = allSymbolItems[i];
    try {
      // グローバル編集側の幽霊シンボルは getByName でエラーになる / Phantom Global Edit symbols throw on getByName
      targetDocument.symbols.getByName(currentItem.symbol.name);
    } catch (e) {
      globalEditItems.push(currentItem);
    }
  }
  // 元オブジェクト自身もシンボルインスタンスに置換するため、対象列の先頭に追加 / Prepend the original item so it is also replaced with a symbol instance
  if (globalEditItems.length > 0) { globalEditItems.unshift(currentSelection[0]); }
  return globalEditItems;
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

  /* 後続の検索・Global Edit が元選択を参照できるよう、元オブジェクトを選択し直す / Reselect the original item so the following search or Global Edit can use it */
  targetDocument.selection = null;
  sourceItem.selected = true;
  return createdSymbol;
}

/* 置換先レイヤーを取得 / Resolve the layer used for replacement */
function getReplacementLayer(currentItem, itemIndex, isGlobalEdit) {
  // グローバル編集時、元の選択アイテムは分離コンテキスト内なので親レイヤーへ / Under Global Edit isolation, place the original selected item on the parent layer
  return (itemIndex === 0 && isGlobalEdit) ? currentItem.layer.parent : currentItem.layer;
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
function replaceItemsWithSymbol(targetDocument, targetItems, destinationSymbol, referencePoint, isGlobalEdit) {
  var createdSymbolItems = [];
  var skippedCount = 0;

  // 逆順：元アイテム (index 0) を最後に処理し、Global Edit 時の親レイヤー切替（getReplacementLayer）を有効化 / Reverse order: process the original item (index 0) last so the parent-layer switch in getReplacementLayer applies under Global Edit
  for (var i = targetItems.length - 1; i >= 0; i--) {
    var currentItem = targetItems[i];
    var newSymbolItem = null;

    try {
      if (!isReplacementItemAvailable(currentItem)) {
        skippedCount++;
        continue;
      }

      var destinationLayer = getReplacementLayer(currentItem, i, isGlobalEdit);
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

  /* テキストは「同じフォント・サイズ・スタイル」かつ「同じ文字列」で検索、それ以外はグローバル編集 / TextFrames: same font/style/size + same content. Others: Global Edit */
  var replaceTargets;
  if (isTextSelection) {
    replaceTargets = findMatchingTextFrames(activeDoc, sourceTextContent);
  } else {
    app.executeMenuCommand('SmartEdit Menu Item');
    replaceTargets = getGlobalEditSymbolItems(activeDoc);
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
  replaceItemsWithSymbol(activeDoc, replaceTargets, destinationSymbol, referencePoint, !isTextSelection);
})();