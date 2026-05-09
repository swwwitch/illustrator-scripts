#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

(function () {

    // =========================================
    // バージョンとローカライズ / Version and localization
    // =========================================

    var SCRIPT_VERSION = "v1.0.0";

    /* 現在の UI 言語 / Current UI language */
    var currentLang = ($.locale.indexOf('ja') === 0) ? 'ja' : 'en';

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialogTitle: { ja: 'シンボル化して置換', en: 'Symbolize and Replace' },
        symbolName: { ja: 'シンボル名', en: 'Symbol Name' },
        referencePoint: { ja: '位置合わせの基準点', en: 'Alignment Reference Point' },
        cancel: { ja: 'キャンセル', en: 'Cancel' },
        alertDuplicateName: {
            ja: 'シンボル「{name}」は既に存在します。別の名前を指定してください。',
            en: 'A symbol named "{name}" already exists. Please choose a different name.'
        }
    };

    /* キーから現在言語のラベルを取得 / Look up a label for the current language */
    function getLabel(key) {
        return (LABELS[key] && (LABELS[key][currentLang] || LABELS[key].en)) || key;
    }

    // =========================================
    // 設定 / Settings
    // =========================================

    /* 3×3 基準点テーブル（行優先：上 → 中 → 下、列：左 → 中 → 右） / 3×3 reference-point table (row-major: top → middle → bottom, col: L → C → R) */
    var REFERENCE_POINTS = [
        { key: 'TOP_LEFT',      symbolPoint: SymbolRegistrationPoint.SYMBOLTOPLEFTPOINT },
        { key: 'TOP_MIDDLE',    symbolPoint: SymbolRegistrationPoint.SYMBOLTOPMIDDLEPOINT },
        { key: 'TOP_RIGHT',     symbolPoint: SymbolRegistrationPoint.SYMBOLTOPRIGHTPOINT },
        { key: 'MIDDLE_LEFT',   symbolPoint: SymbolRegistrationPoint.SYMBOLMIDDLELEFTPOINT },
        { key: 'CENTER',        symbolPoint: SymbolRegistrationPoint.SYMBOLCENTERPOINT },
        { key: 'MIDDLE_RIGHT',  symbolPoint: SymbolRegistrationPoint.SYMBOLMIDDLERIGHTPOINT },
        { key: 'BOTTOM_LEFT',   symbolPoint: SymbolRegistrationPoint.SYMBOLBOTTOMLEFTPOINT },
        { key: 'BOTTOM_MIDDLE', symbolPoint: SymbolRegistrationPoint.SYMBOLBOTTOMMIDDLEPOINT },
        { key: 'BOTTOM_RIGHT',  symbolPoint: SymbolRegistrationPoint.SYMBOLBOTTOMRIGHTPOINT }
    ];

    /* キー → インデックス（0-8） / key → index (0-8) */
    var AiReferencePoint = {};
    for (var rpIndex = 0; rpIndex < REFERENCE_POINTS.length; rpIndex++) {
        AiReferencePoint[REFERENCE_POINTS[rpIndex].key] = rpIndex;
    }

    /* 既定の基準点 / Default reference point */
    var DEFAULT_REFERENCE_POINT = AiReferencePoint.CENTER;

    /* インデックスを SymbolRegistrationPoint に変換、範囲外は CENTER / Map index to SymbolRegistrationPoint, fallback to CENTER */
    function toSymbolRegistrationPoint(referencePointIndex) {
        if (referencePointIndex < 0 || referencePointIndex >= REFERENCE_POINTS.length) {
            return SymbolRegistrationPoint.SYMBOLCENTERPOINT;
        }
        return REFERENCE_POINTS[referencePointIndex].symbolPoint;
    }

    /* 変形パレット要領で基準点に対応する座標を返す / Return the coordinate matching a reference point (Transform-palette style) */
    function getReferencePointPosition(targetItem, referencePointIndex) {
        var bounds = targetItem.visibleBounds;
        var col = referencePointIndex % 3;
        var row = Math.floor(referencePointIndex / 3);
        var x = col === 0 ? bounds[0] : col === 2 ? bounds[2] : (bounds[0] + bounds[2]) / 2;
        var y = row === 0 ? bounds[1] : row === 2 ? bounds[3] : (bounds[1] + bounds[3]) / 2;
        return [x, y];
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /* 指定名のシンボルが既に存在するか / Whether a symbol with the given name already exists */
    function symbolNameExists(symbolName) {
        try {
            app.activeDocument.symbols.getByName(symbolName);
            return true;
        } catch (e) {
            return false;
        }
    }

    /* 基準点ラジオボタンを排他的に選択 / Select one reference-point radio button exclusively */
    function selectReferencePointButton(radioButtons, selectedIndex) {
        for (var i = 0; i < radioButtons.length; i++) {
            radioButtons[i].value = (i === selectedIndex);
        }
        radioButtons.selectedIndex = selectedIndex;
    }

    /* 3x3 基準点ラジオを生成（手動排他、選択値は radioButtons.selectedIndex で参照） / Build a 3x3 radio grid with manual mutual exclusion; selected index is exposed via radioButtons.selectedIndex */
    function createReferencePointGrid(parentPanel, defaultIndex) {
        var radioButtons = [];
        for (var row = 0; row < 3; row++) {
            var rowGroup = parentPanel.add('group');
            rowGroup.orientation = 'row';
            rowGroup.spacing = 4;
            for (var col = 0; col < 3; col++) {
                var radioButton = rowGroup.add('radiobutton', undefined, '');
                radioButtons.push(radioButton);
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

    function showSymbolizeDialog(defaultName, defaultReferencePoint) {
        var dialog = new Window('dialog', getLabel('dialogTitle') + ' ' + SCRIPT_VERSION);
        dialog.alignChildren = ['fill', 'top'];
        dialog.margins = 16;
        dialog.spacing = 12;

        /* シンボル名 / Symbol name */
        var symbolNamePanel = dialog.add('panel', undefined, getLabel('symbolName'));
        symbolNamePanel.orientation = 'column';
        symbolNamePanel.alignChildren = 'fill';
        symbolNamePanel.margins = [15, 20, 15, 10];
        var nameInput = symbolNamePanel.add('edittext', undefined, defaultName || '');
        nameInput.characters = 18;
        nameInput.active = true;

        /* 基準点 / Reference point */
        var referencePointPanel = dialog.add('panel', undefined, getLabel('referencePoint'));
        referencePointPanel.orientation = 'column';
        referencePointPanel.alignChildren = 'center';
        referencePointPanel.margins = [15, 20, 15, 10];
        referencePointPanel.spacing = 4;
        var referencePointButtons = createReferencePointGrid(referencePointPanel, defaultReferencePoint);

        var buttonGroup = dialog.add('group');
        buttonGroup.alignment = 'right';
        buttonGroup.add('button', undefined, getLabel('cancel'), { name: 'cancel' });
        /* OK は重複チェックでダイアログを保持できるよう name:'ok' を付けず、defaultElement で Enter に紐付け / OK omits name:'ok' so duplicate-name validation can keep the dialog open; defaultElement wires Enter to it */
        var okButton = buttonGroup.add('button', undefined, 'OK');
        okButton.onClick = function () {
            var trimmedName = nameInput.text.replace(/^\s+|\s+$/g, '');
            if (!trimmedName) { nameInput.active = true; return; }
            if (symbolNameExists(trimmedName)) {
                alert(LABELS.alertDuplicateName.ja.replace('{name}', trimmedName) + '\n' +
                    LABELS.alertDuplicateName.en.replace('{name}', trimmedName));
                nameInput.active = true;
                return;
            }
            dialog.close(1);
        };
        dialog.defaultElement = okButton;

        if (dialog.show() !== 1) return null;
        return {
            symbolName: nameInput.text.replace(/^\s+|\s+$/g, ''),
            referencePoint: referencePointButtons.selectedIndex
        };
    }

    // =========================================
    // メイン処理 / Main flow
    // =========================================

    if (app.documents.length <= 0) return;
    var doc = app.activeDocument;
    if (!doc.selection.length) return;

    /* 操作前に選択をスナップショット（後段の操作で書き換わるため） / Snapshot the selection before any operation that mutates it */
    var initialSelection = [];
    for (var s = 0; s < doc.selection.length; s++) initialSelection.push(doc.selection[s]);
    var isTextOnly = isTextOnlySelection(initialSelection);
    var sourceTextContent = isTextOnly ? initialSelection[0].contents : null;

    /* テキスト選択時はその文字列（改行は半角スペース化）を初期値に / For text selections, seed the dialog with the contents (newlines collapsed to spaces) */
    var defaultSymbolName = isTextOnly ? sanitizeTextForSymbolName(sourceTextContent) : '';
    var dialogResult = showSymbolizeDialog(defaultSymbolName, DEFAULT_REFERENCE_POINT);
    if (!dialogResult) return;
    var symbolName = dialogResult.symbolName;
    var referencePoint = dialogResult.referencePoint;

    if (isTextOnly) {
        /* テキスト分岐: 1個目のテキストのみシンボル化 → フォント/スタイル/サイズで類似検索 → 文字列一致で絞り込み / Text branch: symbolize the first text only, then expand by font/style/size and filter by exact contents */
        var sourceText = initialSelection[0];
        doc.selection = null;
        sourceText.selected = true;
        createSymbolFromSelection(symbolName, referencePoint);
        doc.selection = null;
        sourceText.selected = true;
        var textTargets = findMatchingTextFrames(doc, sourceTextContent);
        replaceItemsWithSymbol(textTargets, symbolName, referencePoint);
    } else {
        createSymbolFromSelection(symbolName, referencePoint);
        app.executeMenuCommand('SmartEdit Menu Item');
        tagSelectionWithTempNote();
        var smartEditTargets = [];
        for (var k = 0; k < doc.selection.length; k++) smartEditTargets.push(doc.selection[k]);
        replaceItemsWithSymbol(smartEditTargets, symbolName, referencePoint);
    }

    // =========================================
    // 関数 / Functions
    // =========================================

    /* 選択を複製してシンボル化（指定基準点で登録） / Duplicate the selection and register it as a symbol with the given registration point */
    function createSymbolFromSelection(symbolName, referencePoint) {
        var doc = app.activeDocument;
        var selection = doc.selection;
        if (!selection.length) return;
        var symbolSource;
        if (selection.length === 1) {
            symbolSource = selection[0].duplicate();
        } else {
            symbolSource = doc.groupItems.add();
            for (var i = selection.length - 1; i >= 0; i--) {
                selection[i].duplicate(symbolSource, ElementPlacement.PLACEATBEGINNING);
            }
        }
        doc.symbols.add(symbolSource, toSymbolRegistrationPoint(referencePoint)).name = symbolName;
        symbolSource.remove();
    }

    /* 一括選択した対象に note="temp_memo" を書き込むアクションを実行 / Run an action that writes note="temp_memo" onto the current selection */
    function tagSelectionWithTempNote() {
        var actionDefinition = '/version 3/name [ 4 6e6f7465 ]/isOpen 1/actionCount 1/action-1 {/name [ 4 74656d70 ]/keyIndex 0/colorIndex 0/isOpen 1/eventCount 1/event-1 {/useRulersIn1stQuadrant 0/internalName (adobe_attributePalette)/localizedName [ 12 e5b19ee680a7e8a8ade5ae9a ]/isOpen 1/isOn 1/hasDialog 0/parameterCount 1/parameter-1 {/key 1852798053/showInPalette 4294967295/type (ustring)/value [ 9 74656d705f6d656d6f ]}}}';
        var actionFile = new File('~/ScriptAction.aia');
        actionFile.open('w');
        actionFile.write(actionDefinition);
        actionFile.close();
        app.loadAction(actionFile);
        actionFile.remove();
        app.doScript('temp', 'note', false);
        app.unloadAction('note', '');
    }

    /* 指定アイテム群を、各アイテムの基準点を揃えながらシンボルインスタンスに置換し、置換後のインスタンスを選択状態にする / Replace the given items with symbol instances aligned by the reference point, leaving the new instances selected */
    function replaceItemsWithSymbol(targetItems, symbolName, referencePoint) {
        var doc = app.activeDocument;
        var targetSymbol = findSymbolByName(symbolName);
        if (!targetSymbol) return;
        var createdSymbolItems = [];
        for (var j = 0; j < targetItems.length; j++) {
            var targetItem = targetItems[j];
            var targetPosition = getReferencePointPosition(targetItem, referencePoint);
            var symbolInstance = doc.symbolItems.add(targetSymbol);
            if (symbolInstance.parent !== targetItem.parent) {
                symbolInstance.move(targetItem.parent, ElementPlacement.PLACEATBEGINNING);
            }
            var instancePosition = getReferencePointPosition(symbolInstance, referencePoint);
            symbolInstance.translate(targetPosition[0] - instancePosition[0], targetPosition[1] - instancePosition[1]);
            targetItem.remove();
            createdSymbolItems.push(symbolInstance);
        }
        /* 1 個ずつ .selected=true は多数で固まるので配列で一括代入 / Bulk assignment avoids the per-item freeze that .selected = true triggers on large counts */
        doc.selection = createdSymbolItems;
    }

    function findSymbolByName(symbolName) {
        var symbols = app.activeDocument.symbols;
        for (var i = 0; i < symbols.length; i++) {
            if (symbols[i].name === symbolName) return symbols[i];
        }
        return null;
    }

    /* TextFrame か判定 / Whether the item is a TextFrame */
    function isTextFrame(item) {
        return !!item && item.typename === 'TextFrame';
    }

    /* 選択がすべて TextFrame か / Whether every selected item is a TextFrame */
    function isTextOnlySelection(selection) {
        if (!selection || !selection.length) return false;
        for (var i = 0; i < selection.length; i++) {
            if (!isTextFrame(selection[i])) return false;
        }
        return true;
    }

    /* シンボル名向けにテキストを整形（改行を半角スペース化）/ Normalize text for use as a symbol name (collapse newlines to spaces) */
    function sanitizeTextForSymbolName(text) {
        return (text || '').replace(/[\r\n]+/g, ' ');
    }

    /* 同フォント・スタイル・サイズかつ同一文字列の TextFrame を取得（メニュー実行で doc.selection が書き換わる）/ Collect TextFrames matching font/style/size with identical contents (menu command mutates doc.selection) */
    function findMatchingTextFrames(targetDocument, sourceTextContent) {
        app.executeMenuCommand('Find Text Font Family Style Size menu item');
        var matches = [];
        var currentSelection = targetDocument.selection;
        for (var i = 0; i < currentSelection.length; i++) {
            if (isTextFrame(currentSelection[i]) && currentSelection[i].contents === sourceTextContent) {
                matches.push(currentSelection[i]);
            }
        }
        return matches;
    }
})();
