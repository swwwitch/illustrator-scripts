#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

概要

選択中の 1 オブジェクトをドキュメントにシンボル登録し、ユーザーが指定した
「一致条件」で他の同種オブジェクトを検索して、その全てを生成シンボルの
インスタンスに一括置換する。

- 元の選択は複製してシンボルとして登録し、複製は削除（原本は維持）
- 一致条件はダイアログで選択：
  - 通常オブジェクト：塗りカラー／線カラー／塗り＋線／アピアランス
  - テキスト：同フォント・スタイル・サイズ ＋ 同一文字列
- Illustrator 標準の「選択 > 共通 ...」系メニューコマンドを使う
  （グローバル編集には依存しない）
- 各置換アイテムは元の visibleBounds 上の指定基準点に整列して配置

Overview

Register the single selected object as a symbol, then find every other
object that shares the user-chosen attribute and replace each with an
instance of that symbol.

- The selection is duplicated, the duplicate is registered as a symbol,
  and the duplicate is removed so the original geometry remains.
- Match criterion is chosen via dialog:
  - Generic items: fill / stroke / fill & stroke / appearance
  - TextFrames: same font/style/size + identical contents
- Uses Illustrator's built-in "Select > Same ..." menu commands
  (no dependence on Global Edit).
- Each replacement instance is positioned so the chosen anchor lies on
  the original's visibleBounds anchor.

*/

// =========================================
// バージョンとローカライズ / Version and localization
// =========================================

var SCRIPT_VERSION = "v1.0.0";

function getCurrentLang() {
  return ($.locale.indexOf('ja') === 0) ? 'ja' : 'en';
}
var lang = getCurrentLang();

var LABELS = {
  dialogTitle:           { ja: 'シンボル化して一括置換', en: 'Symbolize and Replace' },
  fieldName:             { ja: 'シンボル名：', en: 'Name:' },
  panelAnchor:           { ja: '基準点', en: 'Anchor' },
  panelMatch:            { ja: '一致条件', en: 'Match by' },
  matchFill:             { ja: '塗りカラー', en: 'Fill color' },
  matchStroke:           { ja: '線カラー', en: 'Stroke color' },
  matchFillStroke:       { ja: '塗り＋線カラー', en: 'Fill & stroke color' },
  matchAppearance:       { ja: 'アピアランス', en: 'Appearance' },
  matchTextFontContents: { ja: 'フォント・スタイル・サイズ＋文字列', en: 'Font/style/size + contents' },
  buttonCancel:          { ja: 'キャンセル', en: 'Cancel' }
};

function L(key) {
  return (LABELS[key] && (LABELS[key][lang] || LABELS[key].en)) || key;
}

// =========================================
// 設定 / Settings
// =========================================

var DEFAULT_SYMBOL_NAME = 'symbol_source';

/* 一致条件 ID と対応する Illustrator メニューコマンド / Match strategy IDs and corresponding Illustrator menu commands */
var MATCH = {
  FILL:          'fill',
  STROKE:        'stroke',
  FILL_STROKE:   'fillStroke',
  APPEARANCE:    'appearance',
  TEXT_FONT_BODY: 'textFontBody'
};

var MATCH_MENU_COMMAND = {
  fill:          'Find Fill Color menu item',
  stroke:        'Find Stroke Color menu item',
  fillStroke:    'Find Fill & Stroke Color menu item',
  appearance:    'Find Style menu item',
  textFontBody:  'Find Text Font Family Style Size menu item'
};

/* 0-8 の基準点インデックスから SymbolRegistrationPoint への対応表（行優先 3x3）/ Anchor-index (0-8, row-major 3x3) to SymbolRegistrationPoint */
var ANCHOR_REGISTRATION_POINT = [
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

var DEFAULT_ANCHOR_INDEX = 4; // center

// =========================================
// テキストヘルパー / Text helpers
// =========================================

function isTextFrame(item) {
  return !!item && item.typename === 'TextFrame';
}

function hasSameTextContents(item, contents) {
  return isTextFrame(item) && item.contents === contents;
}

/* シンボル名向けにテキストを整形（改行を半角スペース化）/ Normalize text for use as a symbol name (collapse newlines into spaces) */
function sanitizeForSymbolName(text) {
  return (text || '').replace(/[\r\n]+/g, ' ');
}

// =========================================
// 座標計算 / Geometry helpers
// =========================================

/* 0-8 の基準点インデックスから item.visibleBounds 上の [x, y] を返す
   / Return the [x, y] on item.visibleBounds for an anchor index 0-8 */
function anchorPoint(item, anchorIndex) {
  var b = item.visibleBounds;
  var col = anchorIndex % 3;
  var row = Math.floor(anchorIndex / 3);
  var x = (col === 0) ? b[0] : (col === 2) ? b[2] : (b[0] + b[2]) / 2;
  var y = (row === 0) ? b[1] : (row === 2) ? b[3] : (b[1] + b[3]) / 2;
  return [x, y];
}

// =========================================
// シンボル登録 / Symbol registration
// =========================================

/* 選択 1 個を複製してシンボル登録、複製は削除して原本を残す
   / Register the selection as a symbol via a temporary duplicate; the original geometry stays */
function registerSymbolFromSelection(targetDocument, symbolName, anchorIndex) {
  var original = targetDocument.selection[0];
  var temporary = original.duplicate();
  var symbol = targetDocument.symbols.add(temporary, ANCHOR_REGISTRATION_POINT[anchorIndex]);
  symbol.name = symbolName;
  temporary.remove();
  targetDocument.selection = null;
  original.selected = true;
  return symbol;
}

// =========================================
// 候補抽出 / Candidate collection
// =========================================

/* 一致条件のメニューを実行し、候補配列を返す。テキスト経路は最後に文字列でフィルタ
   / Run the menu command for the chosen strategy and gather candidates;
     for text the result is post-filtered by exact contents */
function collectCandidates(targetDocument, strategy, sourceTextContents) {
  var menuCommand = MATCH_MENU_COMMAND[strategy];
  if (!menuCommand) { return []; }
  app.executeMenuCommand(menuCommand);

  var current = targetDocument.selection;
  var collected = [];
  var i;
  if (strategy === MATCH.TEXT_FONT_BODY) {
    for (i = 0; i < current.length; i++) {
      if (hasSameTextContents(current[i], sourceTextContents)) { collected.push(current[i]); }
    }
  } else {
    for (i = 0; i < current.length; i++) { collected.push(current[i]); }
  }
  return collected;
}

// =========================================
// 置換 / Replacement
// =========================================

/* 候補を生成シンボルのインスタンスに置換 / Replace each candidate with an instance of the symbol */
function replaceWithSymbolInstances(candidates, symbol, anchorIndex) {
  for (var i = candidates.length - 1; i >= 0; i--) {
    var current = candidates[i];
    var targetXY = anchorPoint(current, anchorIndex);
    var instance = current.layer.symbolItems.add(symbol);
    var instanceXY = anchorPoint(instance, anchorIndex);
    instance.translate(targetXY[0] - instanceXY[0], targetXY[1] - instanceXY[1]);
    current.remove();
  }
}

// =========================================
// ダイアログ / Dialog
// =========================================

/* 3x3 ラジオグリッド（手動排他、buttons.selectedIndex に保持）/ 3x3 mutually-exclusive radio grid; selection exposed as buttons.selectedIndex */
function buildAnchorGrid(parent, defaultIndex) {
  var buttons = [];
  for (var row = 0; row < 3; row++) {
    var rowGroup = parent.add('group');
    rowGroup.orientation = 'row';
    rowGroup.spacing = 4;
    for (var col = 0; col < 3; col++) {
      buttons.push(rowGroup.add('radiobutton', undefined, ''));
    }
  }
  buttons[defaultIndex].value = true;
  buttons.selectedIndex = defaultIndex;
  for (var i = 0; i < buttons.length; i++) {
    (function (idx) {
      buttons[idx].onClick = function () {
        for (var j = 0; j < buttons.length; j++) { buttons[j].value = (j === idx); }
        buttons.selectedIndex = idx;
      };
    })(i);
  }
  return buttons;
}

/* 一致条件のラジオ群（同一親なので ScriptUI が自動排他）/ Match-strategy radios; same parent gives automatic mutual exclusion */
function buildStrategyRadios(parent, isText, defaultStrategy) {
  var entries = isText
    ? [[L('matchTextFontContents'), MATCH.TEXT_FONT_BODY]]
    : [
        [L('matchFill'),       MATCH.FILL],
        [L('matchStroke'),     MATCH.STROKE],
        [L('matchFillStroke'), MATCH.FILL_STROKE],
        [L('matchAppearance'), MATCH.APPEARANCE]
      ];
  var radios = [];
  var strategies = [];
  for (var i = 0; i < entries.length; i++) {
    var rb = parent.add('radiobutton', undefined, entries[i][0]);
    if (entries[i][1] === defaultStrategy) { rb.value = true; }
    radios.push(rb);
    strategies.push(entries[i][1]);
  }
  return { radios: radios, strategies: strategies };
}

function pickedStrategy(group, fallback) {
  for (var i = 0; i < group.radios.length; i++) {
    if (group.radios[i].value) { return group.strategies[i]; }
  }
  return fallback;
}

function showSettingsDialog(defaultName, defaultAnchorIndex, isText, defaultStrategy) {
  var dialog = new Window('dialog', L('dialogTitle') + ' ' + SCRIPT_VERSION);
  dialog.alignChildren = ['fill', 'top'];
  dialog.margins = 16;
  dialog.spacing = 12;

  /* シンボル名 / Symbol name */
  var nameRow = dialog.add('group');
  nameRow.orientation = 'row';
  nameRow.alignChildren = ['left', 'center'];
  nameRow.add('statictext', undefined, L('fieldName'));
  var nameInput = nameRow.add('edittext', undefined, defaultName);
  nameInput.characters = 24;
  nameInput.active = true;

  /* 基準点 / Anchor */
  var anchorPanel = dialog.add('panel', undefined, L('panelAnchor'));
  anchorPanel.orientation = 'column';
  anchorPanel.alignChildren = 'center';
  anchorPanel.margins = 12;
  anchorPanel.spacing = 4;
  var anchorButtons = buildAnchorGrid(anchorPanel, defaultAnchorIndex);

  /* 一致条件 / Match criterion */
  var matchPanel = dialog.add('panel', undefined, L('panelMatch'));
  matchPanel.orientation = 'column';
  matchPanel.alignChildren = 'left';
  matchPanel.margins = 12;
  matchPanel.spacing = 4;
  var strategyGroup = buildStrategyRadios(matchPanel, isText, defaultStrategy);

  /* OK / Cancel（OK はローカライズ不要）/ OK / Cancel (OK left as-is) */
  var actionRow = dialog.add('group');
  actionRow.alignment = 'right';
  actionRow.add('button', undefined, L('buttonCancel'), { name: 'cancel' });
  actionRow.add('button', undefined, 'OK', { name: 'ok' });

  if (dialog.show() !== 1) { return null; }
  return {
    name:        nameInput.text || defaultName,
    anchorIndex: anchorButtons.selectedIndex,
    strategy:    pickedStrategy(strategyGroup, defaultStrategy)
  };
}

// =========================================
// メイン / Main
// =========================================

(function main() {
  if (app.documents.length <= 0) { return; }
  var activeDoc = app.activeDocument;
  if (activeDoc.selection.length !== 1) { return; }

  var sourceItem = activeDoc.selection[0];
  var isText = isTextFrame(sourceItem);
  var sourceTextContents = isText ? sourceItem.contents : null;

  var defaultName = (isText && sourceTextContents)
    ? sanitizeForSymbolName(sourceTextContents)
    : DEFAULT_SYMBOL_NAME;
  var defaultStrategy = isText ? MATCH.TEXT_FONT_BODY : MATCH.FILL;

  var settings = showSettingsDialog(defaultName, DEFAULT_ANCHOR_INDEX, isText, defaultStrategy);
  if (!settings) { return; }

  var symbol = registerSymbolFromSelection(activeDoc, settings.name, settings.anchorIndex);

  var candidates = collectCandidates(activeDoc, settings.strategy, sourceTextContents);
  if (candidates.length === 0) { return; }

  replaceWithSymbolInstances(candidates, symbol, settings.anchorIndex);
})();
