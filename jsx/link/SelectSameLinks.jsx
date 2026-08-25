#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択中のリンク画像と同じファイルを参照する配置画像をドキュメント全体から探し、まとめて選択または削除します。
同一かどうかの判定は、リンクの絶対パスとファイル名のどちらでも行えます。

詳細は README を参照してください。

### Overview

Searches the whole document for placed images that reference the same file as the selection, then selects or deletes them together.
The match can be made either on the absolute path of the link or on the file name.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SelectSameLinks";              /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-05-20";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-08-17";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SelectSameLinks.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SelectSameLinks.md

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    /* リンクの同一判定に使うキー / Key used to compare links */
    var MATCH_BY_PATH = "path";  /* placedItem.file.fsName（絶対パス）で判定 / compare by absolute path */
    var MATCH_BY_NAME = "name";  /* placedItem.file.name（ファイル名のみ）で判定 / compare by file name only */

    /* 見つかった配置画像に対する動作 / What to do with the matched items */
    var ACTION_SELECT            = "select";           /* 選択する / select them */
    var ACTION_DELETE_IMAGE      = "deleteImage";      /* 配置画像だけ削除する / delete the placed image only */
    var ACTION_DELETE_CLIP_GROUP = "deleteClipGroup";  /* クリップグループごと削除する / delete the enclosing clip group */

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* ダイアログを開いたときの初期選択 / Initial dialog state */
    var DEFAULT_MATCH_MODE = MATCH_BY_PATH;
    var DEFAULT_ACTION     = ACTION_SELECT;

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * UIの表示言語を判定する
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        return ($.locale && $.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var currentLanguage = getCurrentLang();

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        /* ダイアログ / Dialog */
        dialog: {
            title: { ja: "同一リンクを選択／削除", en: "Select / Delete Same Links" }
        },
        /* パネル見出し / Panel titles */
        panel: {
            matchMode: { ja: "判定方法", en: "Match by" },
            action: { ja: "動作", en: "Action" }
        },
        /* ラジオボタン / Radio buttons */
        radio: {
            matchByPath: { ja: "同じパス", en: "Same path" },
            matchByName: { ja: "ファイル名一致", en: "File name only" },
            actionSelect: { ja: "同一リンクを選択", en: "Select same links" },
            actionDeleteImage: { ja: "リンク画像のみを削除", en: "Delete linked image only" },
            actionDeleteClipGroup: { ja: "クリップグループごと削除", en: "Delete with clip group" }
        },
        /* ボタン / Buttons */
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" }
        },
        /* メッセージ / Messages */
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noSelection: {
                ja: "リンク画像（配置画像）を選択してから実行してください。",
                en: "Please select at least one linked (placed) image first."
            },
            noLinkPath: {
                ja: "選択中の配置画像からリンクパスを取得できませんでした。",
                en: "Could not read the linked file path of the selection."
            },
            selected: { ja: "#count#件選択しました。", en: "#count# item(s) selected." },
            deleted: { ja: "#count#件削除しました。", en: "#count# item(s) deleted." }
        }
    };

    /**
     * "category.key" 形式のラベルを現在の言語で取得する
     * @param {string} key - ラベルキー（例: "alert.noDocument"）
     * @returns {string} 現在の言語のラベル文字列（未定義のときはキーをそのまま返す）
     */
    function getLabel(key) {
        var keyParts = key.split(".");
        var labelNode = LABELS;
        for (var i = 0; i < keyParts.length; i++) {
            if (!labelNode) break;
            labelNode = labelNode[keyParts[i]];
        }
        if (labelNode) {
            if (typeof labelNode[currentLanguage] === "string") return labelNode[currentLanguage];
            if (typeof labelNode.en === "string") return labelNode.en;
        }
        return key;
    }

    /**
     * 件数を埋め込んだメッセージを組み立てる
     * @param {string} key - ラベルキー（`#count#` を含むもの）
     * @param {number} count - 埋め込む件数
     * @returns {string} 件数を差し替えた文字列
     */
    function formatCount(key, count) {
        return getLabel(key).replace("#count#", String(count));
    }

    // =========================================
    // UIレイアウトの共通設定 / Shared UI layout
    // =========================================

    /* ウィンドウ・パネルの余白と間隔 / Window & panel margins and spacing */
    var WINDOW_MARGINS = 16;                 /* ウィンドウ外周の余白 / window margin */
    var WINDOW_SPACING = 12;                 /* ウィンドウ内の要素間隔 / window spacing */
    var PANEL_MARGINS  = [16, 20, 16, 12];   /* パネル余白 [左,上,右,下] / panel margins */
    var PANEL_SPACING  = 6;                  /* パネル内の要素間隔 / panel spacing */

    /**
     * ウィンドウの共通設定
     * @param {Window} targetWindow - 設定するウィンドウ
     * @param {number} [spacing] - 要素間隔（省略時は WINDOW_SPACING）
     * @returns {void}
     */
    function setupWindow(targetWindow, spacing) {
        targetWindow.orientation = "column";
        targetWindow.alignChildren = "fill";
        targetWindow.margins = WINDOW_MARGINS;
        targetWindow.spacing = (typeof spacing === "number") ? spacing : WINDOW_SPACING;
    }

    /**
     * パネルの共通設定
     * @param {Panel} targetPanel - 設定するパネル
     * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function setupPanel(targetPanel, spacing) {
        targetPanel.orientation = "column";
        targetPanel.alignChildren = ["fill", "top"];
        targetPanel.alignment = "fill";
        targetPanel.margins = PANEL_MARGINS;
        targetPanel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /**
     * 行グループの共通設定（ボタン列など）
     * @param {Group} targetGroup - 設定するグループ
     * @param {string} [alignment] - 揃え位置（省略時は "left"）
     * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function setupRow(targetGroup, alignment, spacing) {
        targetGroup.orientation = "row";
        /*
           揃えは横と天地を必ず対で指定する。文字列だけを渡すと天地の指定が外れる。
           Always pass both axes: a bare string drops the vertical one.
        */
        targetGroup.alignment = [alignment || "left", "center"];
        /*
           親の alignChildren（fill）を引き継ぐと、行の中のボタンまで横いっぱいに伸びる。
           Without this, buttons inherit the parent's fill and stretch across the row.
        */
        targetGroup.alignChildren = ["left", "center"];
        targetGroup.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /**
     * 見出し付きのパネルを追加する
     * @param {Window|Group} parent - パネルを追加する親
     * @param {string} labelString - パネルの見出し
     * @returns {Panel} 追加したパネル
     */
    function addPanel(parent, labelString) {
        var newPanel = parent.add("panel", undefined, labelString);
        setupPanel(newPanel);
        return newPanel;
    }

    /**
     * ラジオボタンを追加する（パネルのfillを打ち消して左揃えにする）
     * @param {Panel} parentPanel - ラジオボタンを追加するパネル
     * @param {string} labelString - 表示する文言
     * @param {boolean} isSelected - 初期状態で選択するかどうか
     * @returns {RadioButton} 追加したラジオボタン
     */
    function addRadio(parentPanel, labelString, isSelected) {
        var radio = parentPanel.add("radiobutton", undefined, labelString);
        radio.alignment = "left";
        radio.value = isSelected;
        return radio;
    }

    // =========================================
    // リンクの判定 / Link matching
    // =========================================

    /**
     * 選択範囲（およびその子孫）から配置画像を集める
     * @param {Array} selectedItems - ドキュメントの選択範囲
     * @returns {Array<PlacedItem>} 見つかった配置画像
     */
    function collectPlacedItemsFromSelection(selectedItems) {
        var collected = [];
        if (!selectedItems) return collected;

        /* グループの中の配置画像も拾う / Descend into groups and clip groups */
        function visit(node) {
            if (node.typename === "PlacedItem") {
                collected.push(node);
            } else if (node.typename === "GroupItem") {
                for (var i = 0; i < node.pageItems.length; i++) visit(node.pageItems[i]);
            }
        }

        for (var i = 0; i < selectedItems.length; i++) visit(selectedItems[i]);
        return collected;
    }

    /**
     * 配置画像のリンクを表すキーを返す
     * @param {PlacedItem} placedItem - 対象の配置画像
     * @param {string} matchMode - 判定方法（MATCH_BY_PATH / MATCH_BY_NAME）
     * @returns {string|null} 絶対パスまたはファイル名。リンクが取得できないときは null
     */
    function getLinkKey(placedItem, matchMode) {
        try {
            /* リンク切れや埋め込み画像では file の参照で例外になる / Throws for missing links and embedded images */
            var linkedFile = placedItem.file;
            return (matchMode === MATCH_BY_NAME) ? linkedFile.name : linkedFile.fsName;
        } catch (e) {
            return null;
        }
    }

    /**
     * 配置画像を内包するクリップグループを取得する
     * @param {PlacedItem} placedItem - 対象の配置画像
     * @returns {GroupItem|null} 最も内側のクリップグループ。無ければ null
     */
    function getEnclosingClipGroup(placedItem) {
        var node = placedItem.parent;
        while (node && node.typename === "GroupItem") {
            if (node.clipped) return node;
            node = node.parent;
        }
        return null;
    }

    /**
     * 指定したリンクキーを持つ配置画像をドキュメント全体から集める
     * @param {Document} doc - 対象ドキュメント
     * @param {Object} keySet - リンクキーをプロパティに持つ集合
     * @param {string} matchMode - 判定方法（MATCH_BY_PATH / MATCH_BY_NAME）
     * @returns {Array<PlacedItem>} 一致した配置画像
     */
    function findMatchingPlacedItems(doc, keySet, matchMode) {
        var matched = [];
        var placedItems = doc.placedItems;
        for (var i = 0; i < placedItems.length; i++) {
            var key = getLinkKey(placedItems[i], matchMode);
            if (key && keySet[key]) matched.push(placedItems[i]);
        }
        return matched;
    }

    /**
     * 削除するオブジェクトを決める（クリップグループごと削除する場合は重複を除く）
     * @param {Array<PlacedItem>} matchedItems - 一致した配置画像
     * @param {string} action - 動作（ACTION_DELETE_IMAGE / ACTION_DELETE_CLIP_GROUP）
     * @returns {Array<PageItem>} 実際に削除するオブジェクト
     */
    function collectRemovalTargets(matchedItems, action) {
        if (action !== ACTION_DELETE_CLIP_GROUP) return matchedItems;

        /* 同じクリップグループに複数の配置画像がある場合、削除対象は1つにまとめる / Collapse siblings sharing a clip group */
        var targets = [];
        for (var i = 0; i < matchedItems.length; i++) {
            var target = getEnclosingClipGroup(matchedItems[i]) || matchedItems[i];
            var isDuplicate = false;
            for (var j = 0; j < targets.length; j++) {
                if (targets[j] === target) { isDuplicate = true; break; }
            }
            if (!isDuplicate) targets.push(target);
        }
        return targets;
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 判定方法と動作を選ぶダイアログを表示する
     * @param {string} initialMatchMode - 初期選択の判定方法
     * @param {string} initialAction - 初期選択の動作
     * @returns {{matchMode: string, action: string}|null} 選択内容。キャンセル時は null
     */
    function showOptionsDialog(initialMatchMode, initialAction) {
        var dialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        setupWindow(dialog);

        /* 判定方法 / Match mode */
        var matchModePanel = addPanel(dialog, getLabel("panel.matchMode"));
        addRadio(matchModePanel, getLabel("radio.matchByPath"), initialMatchMode !== MATCH_BY_NAME);
        var matchByNameRadio = addRadio(matchModePanel, getLabel("radio.matchByName"), initialMatchMode === MATCH_BY_NAME);

        /* 動作 / Action */
        var actionPanel = addPanel(dialog, getLabel("panel.action"));
        addRadio(actionPanel, getLabel("radio.actionSelect"),
            initialAction !== ACTION_DELETE_IMAGE && initialAction !== ACTION_DELETE_CLIP_GROUP);
        var deleteImageRadio = addRadio(actionPanel, getLabel("radio.actionDeleteImage"),
            initialAction === ACTION_DELETE_IMAGE);
        var deleteClipGroupRadio = addRadio(actionPanel, getLabel("radio.actionDeleteClipGroup"),
            initialAction === ACTION_DELETE_CLIP_GROUP);

        /* ボタン / Buttons（Mac 規約：Cancel → OK） */
        var buttonRow = dialog.add("group");
        setupRow(buttonRow, "right", PANEL_SPACING);
        buttonRow.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var okButton = buttonRow.add("button", undefined, getLabel("button.ok"), { name: "ok" });

        var result = null;
        okButton.onClick = function () {
            var action = ACTION_SELECT;
            if (deleteImageRadio.value) action = ACTION_DELETE_IMAGE;
            else if (deleteClipGroupRadio.value) action = ACTION_DELETE_CLIP_GROUP;
            result = {
                matchMode: matchByNameRadio.value ? MATCH_BY_NAME : MATCH_BY_PATH,
                action: action
            };
            dialog.close(1);
        };

        dialog.show();
        return result;
    }

    // =========================================
    // メイン / Main
    // =========================================

    /**
     * 選択中のリンク画像と同じリンクを参照する配置画像を選択／削除する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }

        var doc = app.activeDocument;
        var selectedPlacedItems = collectPlacedItemsFromSelection(doc.selection);
        if (selectedPlacedItems.length === 0) {
            alert(getLabel("alert.noSelection"));
            return;
        }

        var options = showOptionsDialog(DEFAULT_MATCH_MODE, DEFAULT_ACTION);
        if (!options) return;

        /* 選択中の配置画像からリンクキーの集合を作る / Build the set of link keys to look for */
        var keySet = {};
        var hasKey = false;
        for (var i = 0; i < selectedPlacedItems.length; i++) {
            var key = getLinkKey(selectedPlacedItems[i], options.matchMode);
            if (key) {
                keySet[key] = true;
                hasKey = true;
            }
        }
        if (!hasKey) {
            alert(getLabel("alert.noLinkPath"));
            return;
        }

        /* 削除中に参照が無効にならないよう、対象を先に集めておく / Buffer the matches before touching the document */
        var matchedItems = findMatchingPlacedItems(doc, keySet, options.matchMode);
        doc.selection = null;

        if (options.action === ACTION_SELECT) {
            var selectedCount = 0;
            for (var j = 0; j < matchedItems.length; j++) {
                /* ロックレイヤーや非表示レイヤーの項目は選択できない / Locked or hidden items cannot be selected */
                try {
                    matchedItems[j].selected = true;
                    selectedCount++;
                } catch (e) { }
            }
            app.redraw();
            alert(formatCount("alert.selected", selectedCount));
            return;
        }

        var removalTargets = collectRemovalTargets(matchedItems, options.action);
        var deletedCount = 0;
        for (var k = removalTargets.length - 1; k >= 0; k--) {
            /* ロックレイヤーや非表示レイヤーの項目は削除できない / Locked or hidden items cannot be removed */
            try {
                removalTargets[k].remove();
                deletedCount++;
            } catch (e) { }
        }
        app.redraw();
        alert(formatCount("alert.deleted", deletedCount));
    }

    main();

})();
