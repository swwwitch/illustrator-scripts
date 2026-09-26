#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

アクティブなアートボードの外にあるオブジェクト、またはアートボード内の選択していないオブジェクトを、削除するか保管用レイヤーへ移します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/DeleteOutsideArtboard.md

### Overview

Deletes the objects outside the active artboard, or the unselected objects inside it, or moves them to a backup layer.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/DeleteOutsideArtboard.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "DeleteOutsideArtboard";        /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.4.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-07-08";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-27";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/DeleteOutsideArtboard.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/DeleteOutsideArtboard.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* 保管用レイヤーの名前（このレイヤーのオブジェクトは対象外）/ Backup layer name (its objects are never touched) */
    var BACKUP_LAYER_NAME = "// backup";

    // =========================================
    // レイアウト / Layout
    // =========================================

    var PANEL_MARGINS  = [15, 20, 15, 10];  /* パネル余白 [左,上,右,下] */
    var OPTION_MARGINS = [15, 0, 15, 10];   /* オプション欄の余白 [左,上,右,下] */

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * UI言語を判定する
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = getCurrentLang();

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "オブジェクトを削除", en: "Delete Objects" }
        },
        panel: {
            inside:  { ja: "アートボード内", en: "Inside Artboard" },
            outside: { ja: "アートボード外", en: "Outside Artboard" }
        },
        radio: {
            keepSelected: { ja: "選択オブジェクトを残す", en: "Keep Selected Objects" },
            keepAll:      { ja: "すべて残す", en: "Keep All Objects" },
            remove:       { ja: "削除", en: "Delete" },
            ignore:       { ja: "無視（残す）", en: "Ignore (Keep)" }
        },
        checkbox: {
            includeLocked: { ja: "ロックされたオブジェクトを含む", en: "Include Locked Objects" },
            moveToBackup:  { ja: "保管用レイヤーに移す", en: "Move to Backup Layer" }
        },
        tooltip: {
            keepSelected: {
                ja: "アクティブなアートボード内のオブジェクトのうち、選択しているものだけを残して他を削除します。このときアートボード外の設定は使いません。",
                en: "Inside the active artboard, keeps only the selected objects and deletes the rest. The Outside Artboard setting is not used."
            },
            keepAll: {
                ja: "アートボード内のオブジェクトはすべて残します。",
                en: "Keeps every object inside the artboard."
            },
            remove: {
                ja: "アクティブなアートボードの外にはみ出したオブジェクトを削除します。",
                en: "Deletes the objects that sit outside the active artboard."
            },
            ignore: { ja: "アートボードの外のオブジェクトはそのまま残します。", en: "Leaves the objects outside the artboard untouched." },
            includeLocked: {
                ja: "ロックされたオブジェクトも処理の対象にします。オフのときは触りません。",
                en: "Includes locked objects. They are left alone when this is off."
            },
            moveToBackup: {
                ja: "削除せずに保管用のレイヤーへ移します。あとから戻せます。",
                en: "Moves the objects to a backup layer instead of deleting them, so they can be brought back."
            }
        },
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok:     { ja: "削除", en: "Delete" }
        },
        alert: {
            noDocument:  { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noSelection: { ja: "選択オブジェクトがありません。", en: "No objects selected." },
            noTargets:   { ja: "削除対象のオブジェクトはありません。", en: "No objects to delete." }
        }
    };

    /**
     * ラベルを取得する
     * @param {string} labelPath - "dialog.title" のようなドット区切りのキー
     * @returns {string} 現在のUI言語のラベル
     */
    function getLabel(labelPath) {
        var pathKeys = String(labelPath).split('.');
        var labelNode = LABELS;
        for (var i = 0; i < pathKeys.length; i++) {
            labelNode = labelNode[pathKeys[i]];
            if (!labelNode) return labelPath;
        }
        return (labelNode[uiLang] != null) ? labelNode[uiLang] : labelPath;
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * ラジオボタンを1つ追加し、ツールチップを設定する
     * @param {Group|Panel} parentContainer - 追加先のコンテナ
     * @param {string} labelKey - radio / tooltip 共通のキー名
     * @returns {RadioButton} 追加したラジオボタン
     */
    function addRadio(parentContainer, labelKey) {
        var radioButton = parentContainer.add("radiobutton", undefined, getLabel("radio." + labelKey));
        radioButton.helpTip = getLabel("tooltip." + labelKey);
        return radioButton;
    }

    /**
     * チェックボックスを1つ追加し、ツールチップと初期値を設定する
     * @param {Group|Panel} parentContainer - 追加先のコンテナ
     * @param {string} labelKey - checkbox / tooltip 共通のキー名
     * @param {boolean} initialValue - 初期値
     * @returns {Checkbox} 追加したチェックボックス
     */
    function addCheckbox(parentContainer, labelKey, initialValue) {
        var checkbox = parentContainer.add("checkbox", undefined, getLabel("checkbox." + labelKey));
        checkbox.helpTip = getLabel("tooltip." + labelKey);
        checkbox.value = initialValue;
        return checkbox;
    }

    /**
     * 縦並びのパネルを追加する
     * @param {Window} dialog - 追加先のダイアログ
     * @param {string} labelKey - panel のキー名
     * @returns {Panel} 追加したパネル
     */
    function addPanel(dialog, labelKey) {
        var panel = dialog.add("panel", undefined, getLabel("panel." + labelKey));
        panel.orientation = "column";
        panel.alignChildren = "left";
        panel.margins = PANEL_MARGINS;
        return panel;
    }

    /**
     * ダイアログを表示し、選ばれた設定を返す
     * @param {boolean} hasSelection - 選択オブジェクトがあるか（「選択オブジェクトを残す」の初期値）
     * @returns {{keepSelected: boolean, deleteOutside: boolean, moveToBackup: boolean, includeLocked: boolean}|null} 設定。キャンセル時は null
     */
    function showDialog(hasSelection) {
        var dialog = new Window("dialog", getLabel("dialog.title"));
        dialog.orientation = "column";
        dialog.alignChildren = "fill";

        /* アートボード内パネル / Inside-artboard panel */
        var insidePanel = addPanel(dialog, "inside");
        var keepSelectedRadio = addRadio(insidePanel, "keepSelected");
        var keepAllRadio = addRadio(insidePanel, "keepAll");
        /* 選択があれば「選択オブジェクトを残す」を初期値に / Default to "keep selected" when something is selected */
        keepSelectedRadio.value = hasSelection;
        keepAllRadio.value = !hasSelection;

        /* アートボード外パネル / Outside-artboard panel */
        var outsidePanel = addPanel(dialog, "outside");
        var removeRadio = addRadio(outsidePanel, "remove");
        var ignoreRadio = addRadio(outsidePanel, "ignore");
        ignoreRadio.value = true;

        /* オプション（ロック含む、保管用レイヤー）/ Options (include locked, backup layer) */
        var optionGroup = dialog.add("group");
        optionGroup.orientation = "column";
        optionGroup.alignChildren = "left";
        optionGroup.margins = OPTION_MARGINS;
        var includeLockedCheckbox = addCheckbox(optionGroup, "includeLocked", true);
        var moveToBackupCheckbox = addCheckbox(optionGroup, "moveToBackup", false);

        /* ボタン / Buttons */
        var btnRowGroup = dialog.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.alignment = "center";
        btnRowGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        btnRowGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });

        if (dialog.show() !== 1) return null;
        return {
            keepSelected: keepSelectedRadio.value,
            deleteOutside: removeRadio.value,
            moveToBackup: moveToBackupCheckbox.value,
            includeLocked: includeLockedCheckbox.value
        };
    }

    // =========================================
    // 判定と収集 / Hit testing and collection
    // =========================================

    /**
     * オブジェクトがアートボードと重なっているか判定する（visibleBounds の矩形で判定）
     * @param {PageItem} item - 判定するオブジェクト
     * @param {Artboard} artboard - アートボード
     * @returns {boolean} 重なっていれば true
     */
    function isOverlappingArtboard(item, artboard) {
        var itemRect = item.visibleBounds;
        var abRect = artboard.artboardRect;
        var overlapWidth = Math.min(itemRect[2], abRect[2]) - Math.max(itemRect[0], abRect[0]);
        var overlapHeight = Math.min(itemRect[1], abRect[1]) - Math.max(itemRect[3], abRect[3]);
        return overlapWidth > 0 && overlapHeight > 0;
    }

    /**
     * 重なり判定に使うオブジェクトを返す（クリップグループはクリッピングパス）
     * @param {PageItem} item - 対象オブジェクト
     * @returns {PageItem} 判定に使うオブジェクト
     */
    function getHitTestTarget(item) {
        if (item.typename === "GroupItem" && item.clipped) {
            for (var k = 0; k < item.pageItems.length; k++) {
                if (item.pageItems[k].clipping) return item.pageItems[k];
            }
        }
        return item;
    }

    /**
     * 収集の対象にするか判定し、対象ならロックと非表示を解除する
     * @param {PageItem} item - 対象オブジェクト
     * @param {boolean} includeLocked - ロックされたオブジェクトも対象にするか
     * @returns {boolean} 対象なら true
     */
    function prepareCandidate(item, includeLocked) {
        if (item.layer && item.layer.name === BACKUP_LAYER_NAME) return false;
        if (!includeLocked && item.locked) return false;
        if (item.locked) item.locked = false;
        if (!item.visible) item.visible = true;
        return true;
    }

    /**
     * アートボードと重ならないオブジェクトを集める（重なるグループは中身も調べる）
     * @param {PageItems} items - 調べるオブジェクト
     * @param {Artboard} artboard - 基準のアートボード
     * @param {boolean} includeLocked - ロックされたオブジェクトも対象にするか
     * @param {PageItem[]} result - 結果を追加する配列
     * @returns {void}
     */
    function collectOutsideItems(items, artboard, includeLocked, result) {
        for (var i = items.length - 1; i >= 0; i--) {
            var item = items[i];
            if (!prepareCandidate(item, includeLocked)) continue;
            var overlaps = isOverlappingArtboard(getHitTestTarget(item), artboard);
            if (!overlaps) {
                result.push(item);
            } else if (item.typename === "GroupItem") {
                collectOutsideItems(item.pageItems, artboard, includeLocked, result);
            }
        }
    }

    /**
     * アートボードと重なるオブジェクトを集める（重なるグループは中身も調べる）
     * @param {PageItems} items - 調べるオブジェクト
     * @param {Artboard} artboard - 基準のアートボード
     * @param {boolean} includeLocked - ロックされたオブジェクトも対象にするか
     * @param {PageItem[]} result - 結果を追加する配列
     * @returns {void}
     */
    function collectInsideItems(items, artboard, includeLocked, result) {
        for (var i = items.length - 1; i >= 0; i--) {
            var item = items[i];
            if (!prepareCandidate(item, includeLocked)) continue;
            if (!isOverlappingArtboard(getHitTestTarget(item), artboard)) continue;
            result.push(item);
            if (item.typename === "GroupItem") {
                collectInsideItems(item.pageItems, artboard, includeLocked, result);
            }
        }
    }

    /**
     * 選択に含まれないものだけを返す
     * @param {PageItem[]} items - 対象オブジェクト
     * @param {Array} selection - 選択オブジェクト
     * @returns {PageItem[]} 選択されていないオブジェクト
     */
    function excludeSelected(items, selection) {
        var filteredItems = [];
        for (var i = 0; i < items.length; i++) {
            var isSelected = false;
            for (var j = 0; j < selection.length; j++) {
                if (items[i] === selection[j]) {
                    isSelected = true;
                    break;
                }
            }
            if (!isSelected) filteredItems.push(items[i]);
        }
        return filteredItems;
    }

    // =========================================
    // 削除と移動 / Delete and move
    // =========================================

    /**
     * 保管用レイヤーを取得する（無ければ作る）
     * @param {Document} doc - 対象ドキュメント
     * @returns {Layer} 保管用レイヤー
     */
    function getBackupLayer(doc) {
        var backupLayer;
        try {
            backupLayer = doc.layers.getByName(BACKUP_LAYER_NAME);
        } catch (e) {
            /* 見つからないと例外 / getByName throws when missing */
            backupLayer = doc.layers.add();
            backupLayer.name = BACKUP_LAYER_NAME;
        }
        return backupLayer;
    }

    /**
     * オブジェクトと、その中身すべてのロックと非表示を解除する
     * @param {PageItem} target - 対象オブジェクト
     * @returns {void}
     */
    function unlockAndShowAll(target) {
        if (target.locked) target.locked = false;
        if (!target.visible) target.visible = true;
        if (typeof target.pageItems !== "undefined") {
            for (var i = 0; i < target.pageItems.length; i++) {
                unlockAndShowAll(target.pageItems[i]);
            }
        }
    }

    /**
     * オブジェクトを保管用レイヤーへ移し、レイヤーを非表示にする
     * @param {PageItem} item - 移すオブジェクト
     * @param {Document} doc - 対象ドキュメント
     * @returns {void}
     */
    function moveItemToBackupLayer(item, doc) {
        var backupLayer = getBackupLayer(doc);
        if (backupLayer.locked) backupLayer.locked = false;
        if (!backupLayer.visible) backupLayer.visible = true;
        unlockAndShowAll(item);
        item.move(backupLayer, ElementPlacement.PLACEATBEGINNING);
        backupLayer.visible = false;
    }

    /**
     * 対象オブジェクトを削除するか保管用レイヤーへ移す
     * @param {PageItem[]} targetItems - 対象オブジェクト
     * @param {Document} doc - 対象ドキュメント
     * @param {boolean} moveToBackup - 削除せずに保管用レイヤーへ移すか
     * @returns {void}
     */
    function disposeItems(targetItems, doc, moveToBackup) {
        for (var i = 0; i < targetItems.length; i++) {
            var item = targetItems[i];
            /* doc.pageItems はグループの中身も含むため、同じオブジェクトが重複して入り、
               先に処理した親ごと消えたものは例外になる / Duplicates via nested pageItems may already be gone */
            try {
                if (item.locked) item.locked = false;
                if (!item.visible) item.visible = true;
                if (item.layer && item.layer.locked) item.layer.locked = false;
                if (moveToBackup) {
                    moveItemToBackupLayer(item, doc);
                } else {
                    item.remove();
                }
            } catch (e) {
                $.writeln("Skipped invalid object: " + e);
            }
        }
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * メイン処理
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }
        var doc = app.activeDocument;
        var currentSelection = doc.selection;
        var hasSelection = !!(currentSelection && currentSelection.length > 0);

        var deleteOptions = showDialog(hasSelection);
        if (!deleteOptions) return;
        /* 「すべて残す」＋「無視」なら何もしない / Nothing to do when both panels keep everything */
        if (!deleteOptions.keepSelected && !deleteOptions.deleteOutside) return;

        var activeArtboard = doc.artboards[doc.artboards.getActiveArtboardIndex()];
        var targetItems = [];

        if (deleteOptions.keepSelected) {
            /* アートボード内の選択していないオブジェクト / Unselected objects inside the artboard */
            if (!hasSelection) {
                alert(getLabel("alert.noSelection"));
                return;
            }
            var insideItems = [];
            collectInsideItems(doc.pageItems, activeArtboard, deleteOptions.includeLocked, insideItems);
            targetItems = excludeSelected(insideItems, currentSelection);
        } else {
            /* アクティブなアートボードの外のオブジェクト / Objects outside the active artboard */
            collectOutsideItems(doc.pageItems, activeArtboard, deleteOptions.includeLocked, targetItems);
        }

        if (targetItems.length === 0) {
            alert(getLabel("alert.noTargets"));
            return;
        }
        disposeItems(targetItems, doc, deleteOptions.moveToBackup);
    }

    main();

})();
