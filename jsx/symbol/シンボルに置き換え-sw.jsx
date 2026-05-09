/*
シンボルに置き換え.jsx
Copyright (c) 2015 Toshiyuki Takahashi
Released under the MIT license
http://opensource.org/licenses/mit-license.php
http://www.graphicartsunit.com/
ver. 0.5.0
*/

(function () {

	var SCRIPT_TITLE = 'シンボルに置き換え';
	var SCRIPT_VERSION = '0.5.0';

	var MAX_VISIBLE_SYMBOLS = 10;
	var RADIO_ROW_HEIGHT = 20;
	var SLOW_SELECTION_THRESHOLD = 20;
	var ERROR_PREFIX = 'エラーが発生して処理を実行できませんでした\nエラー内容：';

	// Settings
	var settings = {
		'symbolIndex': 0
	};

	var activeDoc = app.activeDocument;
	var activeLayer = activeDoc.activeLayer;
	var selectedItems = activeDoc.selection;
	var documentSymbols = activeDoc.symbols;
	var symbolNames = getSymbolNames(documentSymbols);

	if (canRun()) {
		createDialog().show();
	}

	// UI dialog
	function createDialog() {
		var window = new Window('dialog', SCRIPT_TITLE + ' - ver.' + SCRIPT_VERSION);

		var symbolPanel = window.add('panel', undefined, 'シンボル：');
		symbolPanel.alignment = 'left';
		symbolPanel.margins = [15, 20, 15, 10];
		symbolPanel.orientation = 'column';
		symbolPanel.alignChildren = 'left';

		var buttonGroup = window.add('group');
		buttonGroup.alignment = 'center';
		buttonGroup.orientation = 'row';

		var symbolRadios = buildSymbolList(symbolPanel);
		buildButtons(buttonGroup);

		return {
			show: function () { window.show(); }
		};

		function buildSymbolList(parent) {
			var listContainer = parent.add('group');
			listContainer.orientation = 'row';
			listContainer.alignChildren = 'top';
			listContainer.spacing = 4;

			var radioGroup = listContainer.add('group');
			radioGroup.orientation = 'column';
			radioGroup.alignChildren = 'left';
			radioGroup.spacing = 2;

			var visibleCount = Math.min(symbolNames.length, MAX_VISIBLE_SYMBOLS);
			var radios = [];
			for (var i = 0; i < visibleCount; i++) {
				var radio = radioGroup.add('radiobutton', undefined, symbolNames[i]);
				radio.symbolIndex = i;
				radio.value = (i === settings.symbolIndex);
				radio.onClick = onRadioClick;
				radios.push(radio);
			}
			if (radios.length > 0) {
				radios[0].active = true;
			}

			if (symbolNames.length > MAX_VISIBLE_SYMBOLS) {
				addScrollbar(listContainer, radios);
			}
			return radios;
		}

		function addScrollbar(parent, radios) {
			var maxOffset = symbolNames.length - MAX_VISIBLE_SYMBOLS;
			var scrollbar = parent.add('scrollbar', undefined, 0, 0, maxOffset);
			scrollbar.preferredSize.width = 16;
			scrollbar.preferredSize.height = RADIO_ROW_HEIGHT * MAX_VISIBLE_SYMBOLS;
			scrollbar.onChanging = function () {
				var offset = Math.round(scrollbar.value);
				for (var idx = 0; idx < radios.length; idx++) {
					var symbolIdx = offset + idx;
					radios[idx].text = symbolNames[symbolIdx];
					radios[idx].symbolIndex = symbolIdx;
					radios[idx].value = (symbolIdx === settings.symbolIndex);
				}
			};
		}

		function buildButtons(parent) {
			var cancelButton = parent.add('button', undefined, 'キャンセル', { name: 'cancel' });
			var okButton = parent.add('button', undefined, '実行', { name: 'ok' });

			okButton.onClick = function () {
				try {
					replaceSelectionWithSymbol(false);
					window.close();
				} catch (e) {
					alert(ERROR_PREFIX + e);
				}
			};
			cancelButton.onClick = function () {
				window.close();
			};
		}

		function onRadioClick() {
			try {
				settings.symbolIndex = this.symbolIndex;
				previewReplace();
			} catch (e) {
				alert(ERROR_PREFIX + e);
			}
		}

		function previewReplace() {
			replaceSelectionWithSymbol(true);
			app.redraw();
			app.undo();
		}
	}

	// Pre-flight check before showing the dialog
	function canRun() {
		if (!activeDoc || selectedItems.length < 1) {
			alert('オブジェクトが選択されていません');
			return false;
		}
		if (!activeLayer.visible || activeLayer.locked) {
			alert('選択レイヤーがロックされているか非表示になっています');
			return false;
		}
		if (selectedItems.length > SLOW_SELECTION_THRESHOLD) {
			return confirm(selectedItems.length + '個のオブジェクトが選択されており、処理にとても時間がかかる可能性があります。継続しますか？');
		}
		return true;
	}

	// Main process
	function replaceSelectionWithSymbol(isPreview) {
		var targetItems = toItemArray(selectedItems);
		for (var i = 0; i < targetItems.length; i++) {
			var newSymbolItem = activeLayer.symbolItems.add(documentSymbols[settings.symbolIndex]);
			centerOnTarget(newSymbolItem, targetItems[i]);
			if (isPreview) {
				selectedItems[i].hidden = true;
			} else {
				newSymbolItem.selected = true;
				selectedItems[i].remove();
			}
		}
	}

	// Center the symbol item on the target item's bounding box
	function centerOnTarget(symbolItem, targetItem) {
		var t = targetItem.geometricBounds;
		var s = symbolItem.geometricBounds;
		symbolItem.top = (t[1] + t[3]) / 2 - (s[3] - s[1]) / 2;
		symbolItem.left = (t[0] + t[2]) / 2 - (s[2] - s[0]) / 2;
	}

	// Convert a collection (selection / PageItems) to a plain Array
	function toItemArray(collection) {
		var items = [];
		for (var i = 0; i < collection.length; i++) {
			items.push(collection[i]);
		}
		return items;
	}

	// Get symbol names
	function getSymbolNames(symbolCollection) {
		var names = [];
		for (var i = 0; i < symbolCollection.length; i++) {
			names.push(symbolCollection[i].name);
		}
		return names;
	}
}());
