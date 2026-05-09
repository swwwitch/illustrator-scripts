#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

概要

選択中の 1 オブジェクトをシンボル化し、条件に一致するアイテムを一括で
そのシンボルインスタンスに置き換える Illustrator 用 JSX スクリプト。

- 選択1個を複製し、指定したシンボル名と基準点でシンボル登録する
- ダイアログでシンボル名を入力し、空欄（前後の空白を除いた文字数 0）時は OK を無効化する
- 選択がテキストの場合は、その文字列（改行は半角スペース化）をシンボル名の初期値として流用する
- 既存シンボルと同名の場合は警告し、ダイアログを閉じずに再入力できる
- 基準点は 3×3 のラジオボタンで指定し、置換時の位置合わせにも使用する
- 選択がテキストの場合は「同じフォント・スタイル・サイズ」かつ「同一文字列」で対象を検索する（ダイアログで「サイズ違いを許容」を ON にすると、サイズは無視してフォント・スタイル・文字列のみで検索する。チェックボックスはテキスト選択時のみ表示）
- それ以外は［オブジェクトを一括選択］で類似オブジェクトを一括選択し、ダイナミックアクションで一時 note を付与してから SmartEdit を抜け、note を持つアイテムを収集する
- 過去の実行がクリーンアップ前に中断していた場合に備え、note 付与の前に書類内の同一 note 値を一掃する
- 収集後は note をクリアして書類状態を元に戻す
- 複数のグループが選択されている場合は、それらをすべて同一とみなし、先頭のグループを雛形にして選択全体をその新シンボルのインスタンスに置き換える（自動検出は行わずユーザーの選択がそのまま対象）
- 複数選択かつ全要素がグループでない場合は、警告を表示して何も行わずに終了する
- 検出した各アイテムを、作成したシンボルインスタンスに置換する
- 置換後は新規シンボルインスタンス群を選択状態にする
- ロックや非表示（親レイヤー含む）で置換できなかったアイテムがあれば、件数をまとめて通知する
- 置換対象が元オブジェクトしか見つからない場合は、作成したシンボルを削除して通知し終了する

Overview

Illustrator JSX script that converts a single selected object into a symbol
and replaces matching items with instances of that symbol.

- Duplicates the selected item and registers it as a symbol with the specified
name and registration point
- Lets the user enter the symbol name in a dialog, disabling OK while the trimmed name is empty
- Seeds the symbol-name field with the TextFrame's contents (newlines collapsed to spaces) when a TextFrame is selected
- Warns when the name already exists and keeps the dialog open for correction
- Uses a 3×3 radio-button grid for the registration point and uses it again for alignment
- For TextFrames, searches for items with the same font/style/size and identical contents (when "Allow different sizes" is enabled in the dialog, ignores size and matches by font/style/contents only; the checkbox is shown only for TextFrame selections)
- For other objects, toggles SmartEdit on to bulk-select similar items, tags them with a temporary note via a dynamic action, toggles SmartEdit off, and then collects every pageItem carrying that note
- Sweeps any pre-existing items carrying the same note value before tagging, in case a previous run aborted before its cleanup
- Clears the temporary note after collection so the document state is restored
- When multiple GroupItems are selected, treats them as identical: uses the first group as the blueprint and replaces every selected group with an instance of the new symbol (no auto-detection; the user's selection is the target set)
- When the multi-selection includes any non-GroupItem, alerts the user and exits without making changes
- Replaces each detected item with an instance of the created symbol
- Leaves the newly created symbol instances selected after replacement
- Reports the total number of items that could not be replaced because they (or a parent layer) were locked or hidden
- If only the original item is found, removes the created symbol and notifies the user before exiting

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
    ja: 'シンボルの基準点',
    en: 'Symbol Reference Point'
  },
  panelText: {
    ja: 'テキスト',
    en: 'Text'
  },
  checkboxAllowSizeMismatch: {
    ja: 'フォントサイズ違いも対象にする',
    en: 'Include different font sizes'
  },
  helpAllowSizeMismatch: {
    ja: 'ON のときは、フォント・スタイル・文字列が一致していれば、フォントサイズが異なるテキストフレームも置換対象に含めます。',
    en: 'When enabled, TextFrames with matching font, style, and contents are included even if their font size differs.'
  },
  buttonCancel: {
    ja: 'キャンセル',
    en: 'Cancel'
  },
  helpSymbolName: {
    ja: '新しく作成するシンボルの名前です。空欄、および既存シンボルと同じ名前は使用できません。',
    en: 'Name of the new symbol. Empty names and names already used by existing symbols are not allowed.'
  },
  helpReferencePoint: {
    ja: 'シンボル登録時の基準点です。置換時もこの点を使って元オブジェクトの位置に揃えます。',
    en: 'Sets the registration point for the new symbol. The same point is used to align each replacement instance to the original object.'
  },
  alertSkipped: {
    ja: '{count} 件はロック／非表示、または親レイヤーの状態により置換できませんでした。',
    en: '{count} item(s) could not be replaced because they or their parent layers were locked or hidden.'
  },
  alertDuplicateSymbol: {
    ja: 'シンボル「{name}」は既に存在します。別の名前を指定してください。',
    en: 'A symbol named "{name}" already exists. Please choose a different name.'
  },
  alertNoTargets: {
    ja: '置換対象が見つからなかったため、作成したシンボルを削除しました。',
    en: 'No replacement targets were found, so the created symbol was removed.'
  },
  alertMultiSelectionNotGroups: {
    ja: '複数選択時は、すべてのアイテムがグループである必要があります。グループ以外を含む選択では実行できません。',
    en: 'When multiple items are selected, every item must be a group. The script cannot run if the selection includes non-group items.'
  }
};

/* 同じキーの ja / en メッセージを改行でまとめて返すユーティリティ。{key:value} のプレースホルダ置換に対応 / Build a JA+EN bilingual message from a LABELS key, supporting {key:value} placeholders */
function buildBilingualMessage(key, replacements) {
  var entry = LABELS[key];
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
  var aiaContent = '/version 3/name [ 4 6e6f7465 ]/isOpen 1/actionCount 1/action-1 {/name [ 4 74656d70 ]/keyIndex 0/colorIndex 0/isOpen 1/eventCount 1/event-1 {/useRulersIn1stQuadrant 0/internalName (adobe_attributePalette)/localizedName [ 12 e5b19ee680a7e8a8ade5ae9a ]/isOpen 1/isOn 1/hasDialog 0/parameterCount 1/parameter-1 {/key 1852798053/showInPalette 4294967295/type (ustring)/value [ 9 74656d705f6d656d6f ]}}}';
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

  alert(buildBilingualMessage('alertDuplicateSymbol', { name: candidateName }));

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
    alert(buildBilingualMessage('alertSkipped', { count: skippedCount }));
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
function showSymbolizeDialog(targetDocument, defaultName, defaultReferencePoint, isTextSelection) {
  var dialog = new Window('dialog', getLabel('dialogTitle') + ' ' + SCRIPT_VERSION);
  dialog.alignChildren = ['fill', 'top'];
  dialog.margins = 16;
  dialog.spacing = 12;

  /* シンボル名 / Symbol name */
  var symbolNamePanel = dialog.add('panel', undefined, getLabel('panelSymbolName'));
  setupPanel(symbolNamePanel, 4);
  var nameInput = symbolNamePanel.add('edittext', undefined, defaultName);
  nameInput.characters = 18;
  nameInput.preferredSize.width = 210;
  nameInput.helpTip = getLabel('helpSymbolName');
  nameInput.active = true;

  /* 基準点 / Registration point */
  var referencePointPanel = dialog.add('panel', undefined, getLabel('panelReferencePoint'));
  setupPanel(referencePointPanel, 4);
  referencePointPanel.alignChildren = 'center';
  referencePointPanel.helpTip = getLabel('helpReferencePoint');
  var referencePointButtons = createReferencePointGrid(referencePointPanel, defaultReferencePoint);

  /* テキスト（テキスト選択時のみ。一番下に配置）/ Text options (only when a TextFrame is selected; placed at the bottom) */
  var allowSizeMismatchCheckbox = null;
  if (isTextSelection) {
    var textPanel = dialog.add('panel', undefined, getLabel('panelText'));
    setupPanel(textPanel, 4);
    allowSizeMismatchCheckbox = textPanel.add('checkbox', undefined, getLabel('checkboxAllowSizeMismatch'));
    allowSizeMismatchCheckbox.helpTip = getLabel('helpAllowSizeMismatch');
    allowSizeMismatchCheckbox.value = false;
  }

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
      alert(buildBilingualMessage('alertMultiSelectionNotGroups'));
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

    /* SmartEdit ON → ダイナミックアクションで note を付与 → note を持つアイテムを収集 → note クリア / Toggle SmartEdit on, tag with a note via a dynamic action, toggle SmartEdit off, collect every pageItem carrying that note, then clear */
    app.executeMenuCommand('SmartEdit Menu Item');
    attachNoteToSelection();
    replaceTargets = collectItemsByNote(activeDoc, TEMP_NOTE_VALUE);
    clearNoteOnItems(replaceTargets);
  }
  /* 元オブジェクト以外の置換対象がない場合は、シンボル作成のみで終了（複数グループ選択時は到達しない）/ If only the original item is found, keep the created symbol and stop without replacement (unreachable in multi-group mode) */
  if (replaceTargets.length <= 1) {
    try {
      createdSymbol.remove();
    } catch (removeSymbolError) { }

    alert(buildBilingualMessage('alertNoTargets'));
    return;
  }

  /* 置換実行（createSymbolFromSelection が返したシンボルをそのまま使う）/ Replace items using the symbol returned by createSymbolFromSelection */
  replaceItemsWithSymbol(activeDoc, replaceTargets, createdSymbol, referencePoint);
})();