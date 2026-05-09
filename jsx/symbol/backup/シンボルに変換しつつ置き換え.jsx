#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
概要：選択中のオブジェクトのうち、選択内で zOrderPosition が最大のものをシンボル化し、種類とサイズが一致する選択オブジェクトをそのシンボルのインスタンスに置き換える。
置換後は、生成されたシンボルインスタンスのみを選択状態にする。
Overview: Convert the selected object with the largest zOrderPosition into a new symbol, then replace selected objects whose type and size match it with instances of that symbol.
After replacement, only the generated symbol instances remain selected.
*/

(function () {

	// =========================================
	// バージョンとローカライズ / Version and localization
	// =========================================

	var SCRIPT_VERSION = "v1.0.0";

	function getCurrentLanguage() {
		return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
	}
	var currentLanguage = getCurrentLanguage();

	/* 日英ラベル定義 / Japanese-English label definitions */
	var LABELS = {
		dialogTitle: { ja: 'シンボルに変換しつつ置き換え', en: 'Convert to Symbol & Replace' },
		nameLabel: { ja: '名前', en: 'Name' },
		namePanel: { ja: 'シンボル名', en: 'Symbol Name' },
		anchorPanel: { ja: '基準点', en: 'Reference Point' },
		cancel: { ja: 'キャンセル', en: 'Cancel' },
		defaultSymbolName: { ja: '新規シンボル', en: 'New Symbol' },
		nameMissing: { ja: 'シンボル名が指定されていません', en: 'Symbol name is not specified.' },
		needTwoOrMore: { ja: '2 つ以上のオブジェクトを選択してください', en: 'Please select two or more objects.' },
		activeLayerLockedOrHidden: { ja: 'アクティブレイヤーがロックされているか非表示になっています', en: 'The active layer is locked or hidden.' },
		slowConfirmSuffix: { ja: '個のオブジェクトが選択されており、処理にとても時間がかかる可能性があります。継続しますか？', en: ' objects are selected. Processing may take a long time. Continue?' },
		errorPrefix: { ja: 'エラーが発生して処理を実行できませんでした\nエラー内容：', en: 'An error occurred and the process could not be completed.\nDetails: ' }
	};

	/* ラベル取得 / Get a localized label */
	function L(key) {
		return LABELS[key][currentLanguage];
	}

	/* コロン付きラベル（日本語は全角、英語は半角）/ Label with colon (full-width JA, half-width EN) */
	function labelText(key) {
		return L(key) + (currentLanguage === 'ja' ? '：' : ':');
	}

	// =========================================
	// 定数 / Constants
	// =========================================

	var SLOW_PROCESS_WARNING_THRESHOLD = 20;
	var PANEL_MARGINS = [15, 20, 15, 10];
	var SIZE_TOLERANCE = 0.1; // 幅・高さの一致判定許容差 (pt) / Width/height tolerance (pt)

	/* 3x3 基準点グリッド（行優先：上→下、左→右）/ 3x3 anchor grid (row-major: top→bottom, left→right) */
	var ANCHOR_GRID = [
		[SymbolRegistrationPoint.SYMBOLTOPLEFTPOINT, SymbolRegistrationPoint.SYMBOLTOPMIDDLEPOINT, SymbolRegistrationPoint.SYMBOLTOPRIGHTPOINT],
		[SymbolRegistrationPoint.SYMBOLMIDDLELEFTPOINT, SymbolRegistrationPoint.SYMBOLCENTERPOINT, SymbolRegistrationPoint.SYMBOLMIDDLERIGHTPOINT],
		[SymbolRegistrationPoint.SYMBOLBOTTOMLEFTPOINT, SymbolRegistrationPoint.SYMBOLBOTTOMMIDDLEPOINT, SymbolRegistrationPoint.SYMBOLBOTTOMRIGHTPOINT]
	];
	var DEFAULT_ANCHOR_INDEX = 4; // center

	// =========================================
	// エントリポイント / Entry point
	// =========================================

	var activeDoc = app.activeDocument;
	var activeLayer = activeDoc.activeLayer;
	var selectedItems = activeDoc.selection;

	if (canRun()) showDialog();

	// =========================================
	// 事前チェックとダイアログ / Pre-flight check and dialog
	// =========================================

	/* 実行前チェック（選択数・アクティブレイヤー状態・処理時間の警告）/ Pre-flight check before showing the dialog */
	function canRun() {
		if (selectedItems.length < 2) {
			alert(L('needTwoOrMore'));
			return false;
		}
		if (!activeLayer.visible || activeLayer.locked) {
			alert(L('activeLayerLockedOrHidden'));
			return false;
		}
		if (selectedItems.length > SLOW_PROCESS_WARNING_THRESHOLD) {
			return confirm(selectedItems.length + L('slowConfirmSuffix'));
		}
		return true;
	}

	/* ダイアログを構築して表示し、OK 時に変換を実行 / Build and show the dialog; convert on OK */
	function showDialog() {
		var dialog = new Window('dialog', L('dialogTitle') + ' ' + SCRIPT_VERSION);
		dialog.orientation = 'column';
		dialog.alignChildren = 'fill';
		dialog.margins = 16;

		var nameInput = buildNameRow(dialog);
		var anchorRadios = buildAnchorPanel(dialog);
		buildButtons(dialog, nameInput, anchorRadios);

		nameInput.active = true;
		dialog.show();
	}

	/* シンボル名パネル / Symbol name panel */
	function buildNameRow(parent) {
		var panel = parent.add('panel', undefined, labelText('namePanel'));
		setupPanel(panel, 0);

		var nameGroup = panel.add('group');
		nameGroup.orientation = 'row';
		nameGroup.alignChildren = 'center';


		var nameInput = nameGroup.add('edittext', undefined, L('defaultSymbolName'));
		nameInput.preferredSize.width = 160;

		return nameInput;
	}

	/* 3×3 基準点パネル（手動排他制御）/ 3x3 anchor panel with manual exclusivity */
	function buildAnchorPanel(parent) {
		var panel = parent.add('panel', undefined, labelText('anchorPanel'));
		setupPanel(panel, 4);
		panel.alignChildren = 'center';

		var radios = [];
		function onAnchorClick() {
			for (var i = 0; i < radios.length; i++) {
				radios[i].value = (radios[i] === this);
			}
		}

		for (var rowIndex = 0; rowIndex < 3; rowIndex++) {
			var row = panel.add('group');
			row.orientation = 'row';
			row.spacing = 14;
			for (var columnIndex = 0; columnIndex < 3; columnIndex++) {
				var radioButton = row.add('radiobutton', undefined, '');
				radioButton.regPoint = ANCHOR_GRID[rowIndex][columnIndex];
				radioButton.onClick = onAnchorClick;
				radios.push(radioButton);
			}
		}
		radios[DEFAULT_ANCHOR_INDEX].value = true;
		return radios;
	}

	/* OK / キャンセルボタン / OK and Cancel buttons */
	function buildButtons(dialog, nameInput, anchorRadios) {
		var buttonGroup = dialog.add('group');
		buttonGroup.alignment = 'center';
		buttonGroup.orientation = 'row';
		var cancelButton = buttonGroup.add('button', undefined, L('cancel'), { name: 'cancel' });
		var okButton = buttonGroup.add('button', undefined, 'OK', { name: 'ok' });
		okButton.onClick = function () { onOk(nameInput, anchorRadios, dialog); };
		cancelButton.onClick = function () { dialog.close(); };
	}

	/* OK 押下時：名前検証 → 変換実行 / On OK: validate name, run conversion */
	function onOk(nameInput, anchorRadios, dialog) {
		var name = trim(nameInput.text);
		if (name === '') {
			alert(L('nameMissing'));
			return;
		}
		try {
			convertAndReplace(name, getSelectedAnchor(anchorRadios));
			dialog.close();
		} catch (e) {
			alert(L('errorPrefix') + e);
		}
	}

	// =========================================
	// UI ヘルパ / UI helpers
	// =========================================

	/* パネル共通スタイルを適用 / Apply common panel style */
	function setupPanel(panel, spacing) {
		panel.orientation = 'column';
		panel.alignChildren = 'left';
		panel.alignment = 'fill';
		panel.margins = PANEL_MARGINS;
		if (typeof spacing === 'number') {
			panel.spacing = spacing;
		}
	}

	/* 選択中の基準点を返す / Return the registration point of the currently selected anchor radio */
	function getSelectedAnchor(radios) {
		for (var i = 0; i < radios.length; i++) {
			if (radios[i].value) return radios[i].regPoint;
		}
		return SymbolRegistrationPoint.SYMBOLCENTERPOINT;
	}

	// =========================================
	// メイン処理 / Main process
	// =========================================

	/* 最前面の選択をシンボル化し、種類とサイズが一致する選択のみインスタンスに置換（不一致は除外）/ Convert the topmost selection to a new symbol; replace only selections whose type and size match (others are excluded) */
	function convertAndReplace(symbolName, regPoint) {
		var allItems = toItemArray(selectedItems);
		var sourceItem = findTopmost(allItems);
		var targetItems = filterSimilarItems(allItems, sourceItem);

		activeDoc.selection = null;

		var uniqueName = getUniqueSymbolName(symbolName);
		var newSymbol = activeDoc.symbols.add(sourceItem, regPoint);
		try {
			newSymbol.name = uniqueName;
		} catch (e) {
			newSymbol.remove();
			throw e;
		}

		for (var i = 0; i < targetItems.length; i++) {
			var targetItem = targetItems[i];
			var symbolInstance = targetItem.parent.symbolItems.add(newSymbol);
			centerOnTarget(symbolInstance, targetItem);
			symbolInstance.selected = true;
			targetItem.remove();
		}
	}

	/* reference と typename・幅・高さ（許容差付き）が一致する項目のみ抽出 / Keep only items whose typename and size match the reference (with tolerance) */
	function filterSimilarItems(items, reference) {
		var referenceType = reference.typename;
		var referenceSize = getItemSize(reference);
		var result = [];
		for (var i = 0; i < items.length; i++) {
			if (items[i].typename !== referenceType) continue;
			var itemSize = getItemSize(items[i]);
			if (Math.abs(itemSize.width - referenceSize.width) > SIZE_TOLERANCE) continue;
			if (Math.abs(itemSize.height - referenceSize.height) > SIZE_TOLERANCE) continue;
			result.push(items[i]);
		}
		return result;
	}

	/* geometricBounds から幅・高さを取得 / Get width and height from geometricBounds */
	function getItemSize(item) {
		var bounds = item.geometricBounds; // [left, top, right, bottom]
		return {
			width: bounds[2] - bounds[0],
			height: bounds[1] - bounds[3]
		};
	}

	/* 既存シンボル名と衝突しない名前を返す（重複時は _2, _3, ... を付与）/ Return a unique symbol name, suffixing _2, _3, ... on collision */
	function getUniqueSymbolName(baseName) {
		var name = baseName;
		var suffixNumber = 2;
		while (symbolNameExists(name)) name = baseName + '_' + (suffixNumber++);
		return name;
	}

	/* シンボル名が既存シンボルと重複しているか確認 / Check whether the symbol name already exists */
	function symbolNameExists(name) {
		var symbols = activeDoc.symbols;
		for (var i = 0; i < symbols.length; i++) {
			if (symbols[i].name === name) return true;
		}
		return false;
	}

	/* 選択内の zOrderPosition 最大の項目を返す / Pick the item with the largest zOrderPosition within the selection */
	function findTopmost(items) {
		var topmostItem = items[0];
		for (var i = 1; i < items.length; i++) {
			if (items[i].zOrderPosition > topmostItem.zOrderPosition) {
				topmostItem = items[i];
			}
		}
		return topmostItem;
	}

	/* シンボルインスタンスをターゲットの中心に配置 / Center the symbol item on the target item's bounding box */
	function centerOnTarget(symbolItem, targetItem) {
		var targetBounds = targetItem.geometricBounds;
		var symbolBounds = symbolItem.geometricBounds;

		var targetCenterX = (targetBounds[0] + targetBounds[2]) / 2;
		var targetCenterY = (targetBounds[1] + targetBounds[3]) / 2;

		var symbolWidth = symbolBounds[2] - symbolBounds[0];
		var symbolHeight = symbolBounds[1] - symbolBounds[3];

		symbolItem.left = targetCenterX - (symbolWidth / 2);
		symbolItem.top = targetCenterY + (symbolHeight / 2);
	}

	// =========================================
	// ユーティリティ / Utilities
	// =========================================

	/* コレクション（selection／PageItems）を Array に変換 / Convert a collection to a plain Array */
	function toItemArray(collection) {
		var items = [];
		for (var i = 0; i < collection.length; i++) {
			items.push(collection[i]);
		}
		return items;
	}

	/* 文字列の前後の空白を削除 / Remove leading and trailing whitespace */
	function trim(s) {
		return s.replace(/^\s+|\s+$/g, '');
	}
}());
