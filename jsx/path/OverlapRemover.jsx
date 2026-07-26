#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

- 選択オブジェクトに対して、重なりをならすためのメニューコマンドを順に実行します。
- オフセットパス → グループ化 → パスファインダー：合流 → アピアランスを分割、の順で処理します。

*/

/*

### Overview

- Runs a sequence of menu commands against the current selection to flatten overlaps.
- The order is Offset Path, Group, Pathfinder: Merge, then Expand Appearance.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "OverlapRemover";               /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-03-01";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-07-27";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/OverlapRemover.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/OverlapRemover.md

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User configuration
    // =========================================
    var CONFIG = {
        offsetPathCommand: "OffsetPath v22",        /* オフセットパス / Offset Path */
        groupCommand: "group",                      /* グループ化 / Group */
        mergeCommand: "Live Pathfinder Merge",      /* パスファインダー：合流 / Pathfinder: Merge */
        expandCommand: "expandStyle"                /* アピアランスを分割 / Expand Appearance */
    };

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /* 実行環境のロケールから表示言語を決める / Pick the UI language from the locale */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var currentLanguage = getCurrentLang();

    var LABELS = {
        message: {
            noDocument: {
                ja: "ドキュメントを開いてください。",
                en: "Please open a document."
            },
            noSelection: {
                ja: "オブジェクトを選択してから実行してください。",
                en: "Please select an object before running this script."
            },
            commandFailed: {
                ja: "メニューコマンドの実行に失敗しました：\n",
                en: "Failed to run the menu command:\n"
            }
        }
    };

    /* ラベルをドット区切りのキーで引く / Look a label up by a dot-separated key */
    function L(key) {
        var parts = key.split(".");
        var entry = LABELS;
        for (var i = 0; i < parts.length; i++) {
            if (!entry) return key;
            entry = entry[parts[i]];
        }
        return (entry && entry[currentLanguage]) ? entry[currentLanguage] : key;
    }

    // =========================================
    // オブジェクトの走査 / Item inspection
    // =========================================

    /*
     * 選択できるオブジェクトかを判定する。
     * ロック中・非表示のほか、参照が失われたオブジェクトも除外する。
     * Test whether an item can be selected: not locked, not hidden,
     * and still a live reference.
     */
    function isSelectable(item) {
        if (!item) return false;
        try {
            return !item.locked && !item.hidden;
        } catch (e) {
            return false;
        }
    }

    /* 配列に同一参照が含まれるかを判定する / Test whether the array already holds this exact reference */
    function containsRef(items, target) {
        for (var i = 0; i < items.length; i++) {
            if (items[i] === target) return true;
        }
        return false;
    }

    /* ドキュメント内の全オブジェクトを記録する（新規生成物の検出用）/ Snapshot every page item, used to spot newly created ones */
    function snapshotPageItems(doc) {
        var items = [];
        for (var i = 0; i < doc.pageItems.length; i++) {
            items.push(doc.pageItems[i]);
        }
        return items;
    }

    /* 記録に無く、かつ選択できるオブジェクトだけを取り出す / Collect the selectable items that were not in the snapshot */
    function findNewItems(before, after) {
        var items = [];
        for (var i = 0; i < after.length; i++) {
            if (!containsRef(before, after[i]) && isSelectable(after[i])) {
                items.push(after[i]);
            }
        }
        return items;
    }

    /* 複数の配列を、重複と選択不可を除いて1つにまとめる / Merge lists, dropping duplicates and unselectable items */
    function unionSelectable(first, second) {
        var items = [];
        var lists = [first, second];
        for (var i = 0; i < lists.length; i++) {
            var list = lists[i];
            if (!list) continue;
            for (var j = 0; j < list.length; j++) {
                if (isSelectable(list[j]) && !containsRef(items, list[j])) {
                    items.push(list[j]);
                }
            }
        }
        return items;
    }

    // =========================================
    // 選択とコマンド / Selection and commands
    // =========================================

    /* 選択オブジェクトを配列に写し取る（selection は操作で変化するため）/ Copy the selection into a plain array, since selection changes as we work */
    function captureSelection(doc) {
        var items = [];
        var selection = doc.selection;
        if (!selection || !selection.length) return items;
        for (var i = 0; i < selection.length; i++) {
            items.push(selection[i]);
        }
        return items;
    }

    /* 選択できるものだけを選択状態にする / Select only the items that can be selected */
    function setSelection(doc, items) {
        var selectable = unionSelectable(items, null);
        try {
            doc.selection = selectable.length > 0 ? selectable : null;
        } catch (e) {
            // Illustrator が選択を拒む場合は現在の選択のままにする / Keep whatever stays selected
        }
    }

    /* メニューコマンドを実行し、失敗したらアラートを出す / Run a menu command, alerting on failure */
    function runMenu(command) {
        try {
            app.executeMenuCommand(command);
            return true;
        } catch (e) {
            alert(L("message.commandFailed") + command + "\n\n" + e);
            return false;
        }
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /*
     * オフセットパスの結果を元の選択と合わせて選び直す。
     * 新規オブジェクトが検出できなかった場合は、選択が空のときだけ元へ戻す。
     * Reselect the Offset Path results together with the original selection.
     * When nothing new is detected, restore the original selection only if the selection went empty.
     */
    function reselectAfterOffset(doc, originalSelection, newItems) {
        if (newItems.length > 0) {
            setSelection(doc, unionSelectable(originalSelection, newItems));
            return;
        }
        if (!doc.selection || doc.selection.length === 0) {
            setSelection(doc, originalSelection);
        }
    }

    /* コマンドを順に実行して重なりをならす / Run the command sequence that flattens the overlaps */
    function main() {
        var doc = app.activeDocument;
        var originalSelection = captureSelection(doc);

        /* オフセットパスの前後を比較して、新しく作られたオブジェクトを見つける / Compare before and after to spot the new objects */
        var before = snapshotPageItems(doc);
        if (!runMenu(CONFIG.offsetPathCommand)) return;
        reselectAfterOffset(doc, originalSelection, findNewItems(before, snapshotPageItems(doc)));

        if (!runMenu(CONFIG.groupCommand)) return;
        if (!runMenu(CONFIG.mergeCommand)) return;
        runMenu(CONFIG.expandCommand);
    }

    if (app.documents.length === 0) {
        alert(L("message.noDocument"));
    } else if (!app.activeDocument.selection || app.activeDocument.selection.length === 0) {
        alert(L("message.noSelection"));
    } else {
        main();
    }

})();
