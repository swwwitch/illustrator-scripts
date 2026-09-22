#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したオブジェクトを「既存のグループに入れる」「グループから出す」「1つずつグループにする」のどれで処理するかを、ダイアログで選んで実行します。
初期値は選択の状態から自動で決まります。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/GroupMembership.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n36fbd4162721

### Overview

Adds the selection to an existing group, releases it from its groups, or wraps each object in its own group, as chosen in a dialog.
The default choice is picked from the current selection.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/GroupMembership.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "GroupMembership";              /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-09-22";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-22";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/GroupMembership.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/GroupMembership.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n36fbd4162721"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* ［グループから出す］の出す位置の初期値（"layerTop" / "beforeGroup"） / Default placement for "Release from group" */
    var DEFAULT_RELEASE_PLACEMENT = "layerTop";

    /* ［グループ化済みのものは飛ばす］の初期値 / Default of "Skip objects that are already groups" */
    var DEFAULT_SKIP_GROUPS = false;

    // =========================================
    // レイアウト / Layout
    // =========================================

    var WINDOW_MARGINS     = 16;               /* ウィンドウ外周の余白 / window margin */
    var WINDOW_SPACING     = 12;               /* ウィンドウ内の要素間隔 / window spacing */
    var PANEL_MARGINS      = [16, 20, 16, 12]; /* パネル余白 [左,上,右,下] / panel margins */
    var PANEL_SPACING      = 6;                /* パネル内の要素間隔 / panel spacing */
    var BUTTON_BAR_MARGINS = [0, 10, 0, 0];    /* ボタンバーの余白 / margins of the bottom button bar */

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * 表示言語を判定する
     * @returns {string} 日本語環境なら "ja"、それ以外は "en"
     */
    function getCurrentLang() {
        return ($.locale && $.locale.indexOf('ja') === 0) ? 'ja' : 'en';
    }

    var uiLang = getCurrentLang();

    /* 日英ラベル定義（UIパーツ別） / Bilingual labels grouped by UI part */
    var LABELS = {
        dialog: {
            title: { ja: "グループの出し入れ", en: "Group Membership" }
        },
        panel: {
            action: { ja: "処理", en: "Action" },
            options: { ja: "オプション", en: "Options" }
        },
        radio: {
            addToGroup: { ja: "既存のグループに入れる", en: "Add to existing group" },
            releaseFromGroup: { ja: "グループから出す", en: "Release from group" },
            groupEach: { ja: "1つずつグループにする", en: "Group each object" },
            layerTop: { ja: "レイヤーの最前面", en: "Top of the layer" },
            beforeGroup: { ja: "元のグループのすぐ前面", en: "Just in front of the group" }
        },
        fieldLabel: {
            releasePlacement: { ja: "出す位置", en: "Place at" }
        },
        checkbox: {
            skipGroups: { ja: "グループ化済みのものは飛ばす", en: "Skip objects that are already groups" }
        },
        tooltip: {
            addToGroup: {
                ja: "選択したオブジェクトを、一緒に選んだ既存のグループへ入れます。グループは解除しないので、効果やクリッピングマスクは残ります。グループが無いか複数あるときは、1つのグループにまとめ直します（クリップグループは解除しません）。",
                en: "Moves the selected objects into the existing group selected with them. The group is not released, so its effects and clipping mask are kept. With no group or several groups, everything is regrouped as one (clipping groups are kept intact)."
            },
            releaseFromGroup: {
                ja: "グループの中で選択したオブジェクトを、グループの外へ出します。ダイレクト選択ツールなどでグループ内のオブジェクトを選んでから実行します。",
                en: "Moves the objects selected inside groups out of their groups. Select them with the Direct Selection tool or similar first."
            },
            groupEach: {
                ja: "選択したオブジェクトを、1つずつ別々のグループにします。重ね順は変わりません。",
                en: "Wraps each selected object in its own group. The stacking order does not change."
            },
            layerTop: {
                ja: "所属しているレイヤーの最前面へ出します。",
                en: "Places the objects at the top of the layer they belong to."
            },
            beforeGroup: {
                ja: "元のグループ（レイヤー直下のグループ）のすぐ前面へ出します。ほかのオブジェクトとの重なりはほとんど変わりません。",
                en: "Places the objects just in front of their outermost group, so they stay in place relative to other objects."
            },
            skipGroups: {
                ja: "ONにすると、選択の中のグループはグループ化せずにそのまま残します。",
                en: "When on, selected groups are left as they are instead of being wrapped in another group."
            }
        },
        button: {
            ok: { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noSelection: { ja: "オブジェクトを選択して実行してください。", en: "Please select objects and run the script." }
        }
    };

    /**
     * 現在の言語のラベルを返す
     * @param {Object} labelSet - { ja: string, en: string }
     * @returns {string} ラベル文字列
     */
    function getLabel(labelSet) {
        return (labelSet && labelSet[uiLang]) || "";
    }

    /**
     * 項目名にコロンを付けて返す（日本語は全角、英語は半角）
     * @param {Object} labelSet - { ja: string, en: string }
     * @returns {string} コロン付きのラベル
     */
    function labelText(labelSet) {
        return getLabel(labelSet) + (uiLang === "ja" ? "：" : ":");
    }

    // =========================================
    // 選択 / Selection
    // =========================================

    /**
     * 処理対象の選択オブジェクトを返す
     * @returns {PageItem[]|null} 選択オブジェクト（前面→背面の順）。ドキュメントがない・未選択・文字の選択中は null
     */
    function getValidSelection() {
        if (!app.documents.length) return null;
        var selectedItems = app.activeDocument.selection;
        /* 文字ツールで文字を選択中は、添字で要素を取れない TextRange が返る
           A text selection returns a TextRange that cannot be indexed */
        if (!selectedItems || !selectedItems.length || selectedItems.typename === "TextRange") return null;
        return selectedItems;
    }

    /**
     * 選択の中でグループが並んでいる位置を集める
     * @param {PageItem[]} selectedItems - 選択オブジェクト（前面→背面の順）
     * @returns {number[]} GroupItem の添字
     */
    function findGroupIndexes(selectedItems) {
        var groupIndexes = [];
        for (var i = 0; i < selectedItems.length; i++) {
            if (selectedItems[i].typename === "GroupItem") groupIndexes.push(i);
        }
        return groupIndexes;
    }

    /**
     * グループの中にあるオブジェクトを選んでいるか調べる
     * @param {PageItem[]} selectedItems - 選択オブジェクト
     * @returns {boolean} 1つでもグループの中にあれば true
     */
    function hasItemInsideGroup(selectedItems) {
        for (var i = 0; i < selectedItems.length; i++) {
            if (selectedItems[i].parent.typename === "GroupItem") return true;
        }
        return false;
    }

    /**
     * @typedef {Object} SelectionState
     * @property {boolean} canAdd - ［既存のグループに入れる］を選べるか（2つ以上選んでいる）
     * @property {boolean} canRelease - ［グループから出す］を選べるか（グループの中のオブジェクトを選んでいる）
     * @property {string} defaultAction - 初期値の処理（"addToGroup" / "releaseFromGroup" / "groupEach"）
     */

    /**
     * 選択の状態から、選べる処理と初期値を決める
     * グループの中を選んでいれば［グループから出す］、グループ1つとほかのオブジェクトなら［既存のグループに入れる］、それ以外は［1つずつグループにする］
     * @param {PageItem[]} selectedItems - 選択オブジェクト
     * @param {number[]} groupIndexes - selectedItems の中でのグループの添字
     * @returns {SelectionState} 選択の状態
     */
    function analyzeSelection(selectedItems, groupIndexes) {
        var selectionState = {
            canAdd: selectedItems.length >= 2,
            canRelease: hasItemInsideGroup(selectedItems),
            defaultAction: "groupEach"
        };
        if (selectionState.canRelease) {
            selectionState.defaultAction = "releaseFromGroup";
        } else if (groupIndexes.length === 1 && selectionState.canAdd) {
            selectionState.defaultAction = "addToGroup";
        }
        return selectionState;
    }

    // =========================================
    // 既存のグループに入れる / Add to existing group
    // =========================================

    /**
     * クリップグループのマスクを探す
     * パスは clipping、複合パスは先頭のサブパスの clipping で見分ける。テキストのマスクにはフラグが無いので、最前面の項目をマスクとみなす
     * @param {GroupItem} clipGroup - クリップグループ（オブジェクトを入れる前の状態）
     * @returns {PageItem|null} マスク（中身が空なら null）
     */
    function findMaskItem(clipGroup) {
        for (var i = 0; i < clipGroup.pageItems.length; i++) {
            var childItem = clipGroup.pageItems[i];
            if (childItem.clipping) return childItem;
            if (childItem.typename === "CompoundPathItem" && childItem.pathItems.length > 0 && childItem.pathItems[0].clipping) return childItem;
        }
        /* マスクはクリップグループの最前面にある / The mask sits at the top of the clipping group */
        return (clipGroup.pageItems.length > 0) ? clipGroup.pageItems[0] : null;
    }

    /**
     * グループ以外の選択オブジェクトを、重ね順を保ったままグループへ移す
     * グループより前面のものはグループ内の最前面へ、背面のものは最背面へ入れる
     * @param {PageItem[]} selectedItems - 選択オブジェクト（前面→背面の順）
     * @param {number} groupIndex - selectedItems の中でのグループの添字
     * @returns {void}
     */
    function moveItemsIntoGroup(selectedItems, groupIndex) {
        var targetGroup = selectedItems[groupIndex];
        /* 前面側はグループに近いものから最前面へ入れる / Front side: nearest first, each to the top */
        for (var i = groupIndex - 1; i >= 0; i--) {
            selectedItems[i].move(targetGroup, ElementPlacement.PLACEATBEGINNING);
        }
        /* 背面側はグループに近いものから最背面へ入れる / Back side: nearest first, each to the bottom */
        for (var j = groupIndex + 1; j < selectedItems.length; j++) {
            selectedItems[j].move(targetGroup, ElementPlacement.PLACEATEND);
        }
    }

    /**
     * 既存のグループへ、ほかの選択オブジェクトを加える
     * グループは解除しないので、効果・名前・クリッピングマスク・ロックや非表示の中身はそのまま残る
     * @param {Document} doc - 対象ドキュメント
     * @param {PageItem[]} selectedItems - 選択オブジェクト（前面→背面の順）
     * @param {number} groupIndex - selectedItems の中でのグループの添字
     * @returns {void}
     */
    function addItemsToGroup(doc, selectedItems, groupIndex) {
        var targetGroup = selectedItems[groupIndex];
        /* 前面側のオブジェクトはマスクより上に入るので、入れる前にマスクを控えて最前面へ戻す
           Front-side objects land above the mask; remember the mask first and bring it back to the top */
        var maskItem = targetGroup.clipped ? findMaskItem(targetGroup) : null;
        moveItemsIntoGroup(selectedItems, groupIndex);
        if (maskItem) maskItem.zOrder(ZOrderMethod.BRINGTOFRONT);
        doc.selection = [targetGroup];
    }

    /**
     * グループを1段解除してから、選択全体を1つのグループにまとめ直す
     * クリップグループは解除するとマスクが外れるので、解除せずにそのまま入れる
     * @param {Document} doc - 対象ドキュメント
     * @param {PageItem[]} selectedItems - 選択オブジェクト（前面→背面の順）
     * @returns {void}
     */
    function regroupSelection(doc, selectedItems) {
        var releasedItems = [];
        var keptClipGroups = [];
        for (var i = 0; i < selectedItems.length; i++) {
            if (selectedItems[i].typename === "GroupItem" && selectedItems[i].clipped) {
                keptClipGroups.push(selectedItems[i]);
            } else {
                releasedItems.push(selectedItems[i]);
            }
        }

        if (releasedItems.length > 0) {
            doc.selection = releasedItems;
            app.executeMenuCommand("ungroup");
            /* 解除後の選択は、グループの中身とグループ以外のオブジェクト / Afterwards the selection holds the released contents and the other objects */
            releasedItems = doc.selection;
        }

        var itemsToGroup = [];
        for (var j = 0; j < releasedItems.length; j++) itemsToGroup.push(releasedItems[j]);
        for (var k = 0; k < keptClipGroups.length; k++) itemsToGroup.push(keptClipGroups[k]);
        doc.selection = itemsToGroup;
        app.executeMenuCommand("group");
    }

    /**
     * ［既存のグループに入れる］を実行する（グループが1つならそこへ加え、それ以外は1つのグループにまとめ直す）
     * @param {Document} doc - 対象ドキュメント
     * @param {PageItem[]} selectedItems - 選択オブジェクト（前面→背面の順）
     * @param {number[]} groupIndexes - selectedItems の中でのグループの添字
     * @returns {void}
     */
    function addToGroup(doc, selectedItems, groupIndexes) {
        if (groupIndexes.length === 1) {
            addItemsToGroup(doc, selectedItems, groupIndexes[0]);
            return;
        }
        regroupSelection(doc, selectedItems);
    }

    // =========================================
    // グループから出す / Release from group
    // =========================================

    /**
     * いちばん外側の親グループ（レイヤー直下のグループ）を返す
     * @param {PageItem} targetItem - 対象オブジェクト
     * @returns {GroupItem|null} 親グループ（グループの中に無ければ null）
     */
    function findOutermostGroup(targetItem) {
        var outermostGroup = null;
        var currentContainer = targetItem.parent;
        while (currentContainer.typename === "GroupItem") {
            outermostGroup = currentContainer;
            currentContainer = currentContainer.parent;
        }
        return outermostGroup;
    }

    /**
     * グループの中で選んだオブジェクトを、グループの外へ出す
     * @param {Document} doc - 対象ドキュメント
     * @param {PageItem[]} selectedItems - 選択オブジェクト（前面→背面の順）
     * @param {string} releasePlacement - "layerTop"（レイヤーの最前面）/ "beforeGroup"（元のグループのすぐ前面）
     * @returns {void}
     */
    function releaseFromGroup(doc, selectedItems, releasePlacement) {
        var outermostGroup;
        if (releasePlacement === "beforeGroup") {
            /* グループのすぐ前面へは前面側から出す（後から出したものほど、グループ寄りの背面に入る）
               Front to back: each one lands between the group and the previous one */
            for (var i = 0; i < selectedItems.length; i++) {
                outermostGroup = findOutermostGroup(selectedItems[i]);
                if (!outermostGroup) continue;
                selectedItems[i].move(outermostGroup, ElementPlacement.PLACEBEFORE);
            }
        } else {
            /* レイヤーの最前面へは背面側から出す（最後に出した最前面のものが一番上になる）
               Back to front: the frontmost one is moved last and ends up on top */
            for (var j = selectedItems.length - 1; j >= 0; j--) {
                outermostGroup = findOutermostGroup(selectedItems[j]);
                if (!outermostGroup) continue;
                selectedItems[j].move(outermostGroup.parent, ElementPlacement.PLACEATBEGINNING);
            }
        }
        /* 移動で外れた選択を元に戻す / Restore the selection the moves dropped */
        doc.selection = selectedItems;
        /* グループの中を選ぶのに使ったダイレクト選択ツールから戻す / Switch back from the Direct Selection tool */
        app.selectTool("Adobe Select Tool");
    }

    // =========================================
    // 1つずつグループにする / Group each object
    // =========================================

    /**
     * 選択したオブジェクトを、1つずつ別々のグループにする
     * @param {Document} doc - 対象ドキュメント
     * @param {PageItem[]} selectedItems - 選択オブジェクト
     * @param {boolean} skipGroups - グループはグループ化せずに残す
     * @returns {void}
     */
    function groupEachItem(doc, selectedItems, skipGroups) {
        var itemsToSelect = [];
        for (var i = 0; i < selectedItems.length; i++) {
            var targetItem = selectedItems[i];
            if (skipGroups && targetItem.typename === "GroupItem") {
                itemsToSelect.push(targetItem);
                continue;
            }
            var wrapperGroup = targetItem.parent.groupItems.add();
            /* 元の位置のすぐ前面にグループを置いてから中へ入れ、重ね順を保つ / Put the group right in front of the item, then move the item in */
            wrapperGroup.move(targetItem, ElementPlacement.PLACEBEFORE);
            targetItem.move(wrapperGroup, ElementPlacement.PLACEATEND);
            itemsToSelect.push(wrapperGroup);
        }
        doc.selection = itemsToSelect;
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * @typedef {Object} MembershipOptions
     * @property {string} action - "addToGroup" / "releaseFromGroup" / "groupEach"
     * @property {string} releasePlacement - "layerTop" / "beforeGroup"
     * @property {boolean} skipGroups - グループ化済みのものは飛ばす
     */

    /**
     * パネルを追加する
     * @param {Window} parentWindow - 追加先のダイアログ
     * @param {string} panelTitle - パネルの見出し
     * @returns {Panel} 追加したパネル
     */
    function addPanel(parentWindow, panelTitle) {
        var addedPanel = parentWindow.add("panel", undefined, panelTitle);
        addedPanel.orientation = "column";
        addedPanel.alignChildren = ["left", "top"];
        addedPanel.alignment = "fill";
        addedPanel.margins = PANEL_MARGINS;
        addedPanel.spacing = PANEL_SPACING;
        return addedPanel;
    }

    /**
     * ラジオボタンかチェックボックスを、LABELS のキーで追加する
     * @param {Panel|Group} parentContainer - 追加先のパネルかグループ
     * @param {string} controlType - "radiobutton" / "checkbox"
     * @param {string} labelKey - LABELS.radio（または LABELS.checkbox）と LABELS.tooltip のキー
     * @returns {RadioButton|Checkbox} 追加したコントロール
     */
    function addLabeledControl(parentContainer, controlType, labelKey) {
        var labelCategory = (controlType === "radiobutton") ? LABELS.radio : LABELS.checkbox;
        var addedControl = parentContainer.add(controlType, undefined, getLabel(labelCategory[labelKey]));
        addedControl.helpTip = getLabel(LABELS.tooltip[labelKey]);
        return addedControl;
    }

    /**
     * ダイアログを表示して、選ばれた設定を返す
     * @param {SelectionState} selectionState - 選択の状態
     * @returns {MembershipOptions|null} キャンセル時は null
     */
    function showDialog(selectionState) {
        var membershipDialog = new Window("dialog", getLabel(LABELS.dialog.title) + " " + SCRIPT_VERSION);
        membershipDialog.orientation = "column";
        membershipDialog.alignChildren = ["fill", "top"];
        membershipDialog.spacing = WINDOW_SPACING;
        membershipDialog.margins = WINDOW_MARGINS;

        /* 処理 / Action */
        var actionPanel = addPanel(membershipDialog, getLabel(LABELS.panel.action));
        var addToGroupRadio = addLabeledControl(actionPanel, "radiobutton", "addToGroup");
        var releaseFromGroupRadio = addLabeledControl(actionPanel, "radiobutton", "releaseFromGroup");
        var groupEachRadio = addLabeledControl(actionPanel, "radiobutton", "groupEach");
        addToGroupRadio.enabled = selectionState.canAdd;
        releaseFromGroupRadio.enabled = selectionState.canRelease;
        addToGroupRadio.value = (selectionState.defaultAction === "addToGroup");
        releaseFromGroupRadio.value = (selectionState.defaultAction === "releaseFromGroup");
        groupEachRadio.value = (selectionState.defaultAction === "groupEach");

        /* オプション / Options */
        var optionsPanel = addPanel(membershipDialog, getLabel(LABELS.panel.options));

        /* 出す位置のラジオは別グループなので、処理のラジオとは排他にならない / These radios sit in their own group, apart from the action radios */
        var placementRow = optionsPanel.add("group");
        placementRow.orientation = "row";
        placementRow.alignChildren = ["left", "center"];
        placementRow.spacing = PANEL_SPACING;
        placementRow.add("statictext", undefined, labelText(LABELS.fieldLabel.releasePlacement));
        var layerTopRadio = addLabeledControl(placementRow, "radiobutton", "layerTop");
        var beforeGroupRadio = addLabeledControl(placementRow, "radiobutton", "beforeGroup");
        layerTopRadio.value = (DEFAULT_RELEASE_PLACEMENT !== "beforeGroup");
        beforeGroupRadio.value = !layerTopRadio.value;

        var skipGroupsCheckbox = addLabeledControl(optionsPanel, "checkbox", "skipGroups");
        skipGroupsCheckbox.value = DEFAULT_SKIP_GROUPS;

        /**
         * 選んだ処理に関係するオプションだけを有効にする
         * @returns {void}
         */
        function updateOptionState() {
            placementRow.enabled = releaseFromGroupRadio.value;
            skipGroupsCheckbox.enabled = groupEachRadio.value;
        }
        addToGroupRadio.onClick = updateOptionState;
        releaseFromGroupRadio.onClick = updateOptionState;
        groupEachRadio.onClick = updateOptionState;
        updateOptionState();

        /* ボタンエリア（右寄せ） / Button row (right-aligned) */
        var btnRowGroup = membershipDialog.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.margins = BUTTON_BAR_MARGINS;
        btnRowGroup.alignment = ["fill", "bottom"];

        var spacer = btnRowGroup.add("group");
        spacer.alignment = ["fill", "fill"];
        spacer.minimumSize.width = 0;

        var btnRightGroup = btnRowGroup.add("group");
        btnRightGroup.alignChildren = ["right", "center"];
        var btnCancel = btnRightGroup.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });
        var btnOK = btnRightGroup.add("button", undefined, getLabel(LABELS.button.ok), { name: "ok" });

        if (membershipDialog.show() !== 1) return null;

        var selectedAction = "groupEach";
        if (addToGroupRadio.value) selectedAction = "addToGroup";
        if (releaseFromGroupRadio.value) selectedAction = "releaseFromGroup";
        return {
            action: selectedAction,
            releasePlacement: beforeGroupRadio.value ? "beforeGroup" : "layerTop",
            skipGroups: skipGroupsCheckbox.value
        };
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * ダイアログで選んだ処理を実行する
     * @returns {void}
     */
    function main() {
        if (!app.documents.length) {
            alert(getLabel(LABELS.alert.noDocument));
            return;
        }
        var selectedItems = getValidSelection();
        if (!selectedItems) {
            alert(getLabel(LABELS.alert.noSelection));
            return;
        }

        var groupIndexes = findGroupIndexes(selectedItems);
        var membershipOptions = showDialog(analyzeSelection(selectedItems, groupIndexes));
        if (!membershipOptions) return;

        var doc = app.activeDocument;
        if (membershipOptions.action === "addToGroup") {
            addToGroup(doc, selectedItems, groupIndexes);
        } else if (membershipOptions.action === "releaseFromGroup") {
            releaseFromGroup(doc, selectedItems, membershipOptions.releasePlacement);
        } else {
            groupEachItem(doc, selectedItems, membershipOptions.skipGroups);
        }
    }

    main();

})();
