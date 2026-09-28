#target illustrator
#targetengine "SimplifyGroupsDialogEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したグループの入れ子をダイアログで選んだ方法で解消します。外側のグループを残して中だけ解除するか、すべて解除して1つのグループに作り直すかを選べます。
クリップグループや、ロック・非表示のサブグループを解除せずに残すこともできます。

### Overview

Removes nested groups from the selection using the method chosen in a dialog: keep the outer group and ungroup only what is inside, or ungroup everything and rebuild it as one group.
Clipping groups and locked or hidden subgroups can be left intact.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SimplifyGroupsDialog";         /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-09-22";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-28";                   /* 更新日 / last updated */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* 処理方法の初期値（"keepOuter" / "rebuild"） / Default method ("keepOuter" / "rebuild") */
    var DEFAULT_METHOD = "keepOuter";

    /* 「クリップグループは解除しない」の初期値 / Default of "Keep clipping groups" */
    var DEFAULT_KEEP_CLIP_GROUPS = true;

    /* 「ロック・非表示のサブグループは解除しない」の初期値 / Default of "Keep locked or hidden subgroups" */
    var DEFAULT_SKIP_LOCKED_HIDDEN = true;

    // =========================================
    // レイアウト / Layout
    // =========================================

    var DIALOG_MARGINS = [25, 20, 25, 20];  /* ダイアログ余白 [左,上,右,下] / Dialog margins */
    var PANEL_MARGINS  = [15, 20, 15, 10];  /* パネル余白 [左,上,右,下] / Panel margins */

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
            title: { ja: "グループの簡素化", en: "Simplify Groups" }
        },
        panel: {
            method: { ja: "処理方法", en: "Method" },
            options: { ja: "オプション", en: "Options" }
        },
        radio: {
            keepOuter: { ja: "外側のグループを残す", en: "Keep outer group" },
            rebuild: { ja: "1つのグループに作り直す", en: "Rebuild as one group" }
        },
        checkbox: {
            keepClipGroups: { ja: "クリップグループは解除しない", en: "Keep clipping groups" },
            skipLockedHidden: { ja: "ロック・非表示のサブグループは解除しない", en: "Keep locked or hidden subgroups" }
        },
        tooltip: {
            keepOuter: {
                ja: "選択したグループを残したまま、中のサブグループだけを解除します。グループの名前やアピアランスは保たれます。グループ以外のオブジェクトも選んでいるときは、そのグループに入れます（グループが複数なら、全体を新しいグループにまとめます）。",
                en: "Keeps each selected group and ungroups only the subgroups inside it, so its name and appearance are preserved. Other selected objects are moved into the group (with several groups, everything is combined into a new group)."
            },
            rebuild: {
                ja: "選択したグループも含めてすべて解除し、選択全体を新しい1つのグループにまとめます。選択したグループの名前やアピアランスは失われます。",
                en: "Ungroups everything, including the selected groups, and combines the whole selection into one new group. The selected groups' names and appearance are lost."
            },
            keepClipGroups: {
                ja: "ONにすると、クリップグループは解除せずに残し、マスクが外れないようにします。クリップグループの中のサブグループは解除します。",
                en: "When on, clipping groups are left intact so their masks stay in place. Subgroups inside them are still ungrouped."
            },
            skipLockedHidden: {
                ja: "ONにすると、ロックまたは非表示のサブグループには手を付けません。OFFにすると解除し、ロック・非表示の状態は中のオブジェクトに引き継ぎます。",
                en: "When on, locked or hidden subgroups are left untouched. When off, they are ungrouped and their contents stay locked or hidden."
            }
        },
        button: {
            ok: { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
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

    // =========================================
    // クリップグループ / Clipping groups
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

    // =========================================
    // グループの解除 / Ungrouping
    // =========================================

    /**
     * @typedef {Object} SimplifyOptions
     * @property {string} method - "keepOuter"（外側のグループを残す）/ "rebuild"（1つのグループに作り直す）
     * @property {boolean} keepClipGroups - クリップグループを解除しない
     * @property {boolean} skipLockedHidden - ロック・非表示のサブグループを解除しない
     */

    /**
     * ロック・非表示を一時的に外す
     * @param {GroupItem} groupItem - 対象グループ
     * @param {{locked: boolean, hidden: boolean}} lockState - 元の状態
     * @returns {void}
     */
    function unlockAndShow(groupItem, lockState) {
        if (lockState.locked) groupItem.locked = false;
        if (lockState.hidden) groupItem.hidden = false;
    }

    /**
     * ロック・非表示を付け直す（非表示にしてからロックする）
     * @param {PageItem[]} targetItems - 対象オブジェクト
     * @param {{locked: boolean, hidden: boolean}} lockState - 付け直す状態
     * @returns {void}
     */
    function applyLockState(targetItems, lockState) {
        for (var i = 0; i < targetItems.length; i++) {
            if (lockState.hidden) targetItems[i].hidden = true;
            if (lockState.locked) targetItems[i].locked = true;
        }
    }

    /**
     * グループを1段だけ解除し、グループのロック・非表示を中のオブジェクトへ引き継ぐ
     * @param {Document} doc - 対象ドキュメント
     * @param {GroupItem} groupItem - 解除するグループ（ロック・非表示は外してある）
     * @param {{locked: boolean, hidden: boolean}} lockState - グループの元の状態
     * @returns {void}
     */
    function ungroupKeepingState(doc, groupItem, lockState) {
        doc.selection = [groupItem];
        app.executeMenuCommand("ungroup");
        /* 解除直後は中身が選択されている / The released contents are selected right after ungrouping */
        if (lockState.locked || lockState.hidden) applyLockState(doc.selection, lockState);
    }

    /**
     * グループ内のサブグループを再帰的に解除する（parentGroup 自身は残す）
     * @param {Document} doc - 対象ドキュメント
     * @param {GroupItem} parentGroup - 解除せずに残すグループ
     * @param {SimplifyOptions} simplifyOptions - ダイアログの設定
     * @returns {void}
     */
    function ungroupSubGroups(doc, parentGroup, simplifyOptions) {
        /* 解除すると後ろの添字がずれるので末尾から処理 / Walk backwards; ungrouping shifts later indexes */
        for (var i = parentGroup.pageItems.length - 1; i >= 0; i--) {
            var childItem = parentGroup.pageItems[i];
            if (childItem.typename !== "GroupItem") continue;

            var lockState = { locked: childItem.locked, hidden: childItem.hidden };
            if ((lockState.locked || lockState.hidden) && simplifyOptions.skipLockedHidden) continue;

            /* 中を選択できるよう先にロック・非表示を外す / Unlock and show first so the contents can be selected */
            unlockAndShow(childItem, lockState);
            ungroupSubGroups(doc, childItem, simplifyOptions);

            if (simplifyOptions.keepClipGroups && childItem.clipped) {
                applyLockState([childItem], lockState);
            } else {
                ungroupKeepingState(doc, childItem, lockState);
            }
        }
    }

    // =========================================
    // グループへの取り込み / Moving into a group
    // =========================================

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

    // =========================================
    // 処理方法 / Methods
    // =========================================

    /**
     * 選択全体を新しいグループにまとめ、中のサブグループを解除する
     * @param {Document} doc - 対象ドキュメント
     * @param {SimplifyOptions} simplifyOptions - ダイアログの設定
     * @returns {void}
     */
    function rebuildAsOneGroup(doc, simplifyOptions) {
        app.executeMenuCommand("group");
        var groupedItems = doc.selection;
        if (groupedItems.length !== 1 || groupedItems[0].typename !== "GroupItem") return;

        var newGroup = groupedItems[0];
        ungroupSubGroups(doc, newGroup, simplifyOptions);
        doc.selection = [newGroup];

        /* 残したグループ1つだけを包んでいるなら、外側は不要なので解除（解除後は中のグループが選択される）
           If the new group only wraps one kept group, release the wrapper; the inner group stays selected */
        if (newGroup.pageItems.length === 1 && newGroup.pageItems[0].typename === "GroupItem") {
            app.executeMenuCommand("ungroup");
        }
    }

    /**
     * 選択したグループを残したまま、中のサブグループを解除する
     * グループ1つと非グループなら非グループをそのグループへ入れ、グループが複数なら全体を新しいグループにまとめる
     * @param {Document} doc - 対象ドキュメント
     * @param {PageItem[]} selectedItems - 選択オブジェクト（前面→背面の順）
     * @param {number[]} groupIndexes - selectedItems の中でのグループの添字（1つ以上）
     * @param {SimplifyOptions} simplifyOptions - ダイアログの設定
     * @returns {void}
     */
    function simplifyKeepingOuter(doc, selectedItems, groupIndexes, simplifyOptions) {
        /* グループ1つと非グループ：既存のグループへ取り込む / One group plus other objects: move them into it */
        if (groupIndexes.length === 1 && selectedItems.length > 1) {
            var targetGroup = selectedItems[groupIndexes[0]];
            /* 前面側のオブジェクトはマスクより上に入るので、入れる前にマスクを控えて最前面へ戻す
               Front-side objects land above the mask; remember the mask first and bring it back to the top */
            var maskItem = targetGroup.clipped ? findMaskItem(targetGroup) : null;
            moveItemsIntoGroup(selectedItems, groupIndexes[0]);
            if (maskItem) maskItem.zOrder(ZOrderMethod.BRINGTOFRONT);
            ungroupSubGroups(doc, targetGroup, simplifyOptions);
            doc.selection = [targetGroup];
            return;
        }

        /* 選択したグループはそれぞれ残して中だけ解除 / Keep each selected group and flatten only its inside */
        for (var i = 0; i < groupIndexes.length; i++) {
            ungroupSubGroups(doc, selectedItems[groupIndexes[i]], simplifyOptions);
        }
        doc.selection = selectedItems;

        /* グループ以外も選んでいれば、全体を新しいグループにまとめる / Wrap everything in a new group when other objects are selected */
        if (groupIndexes.length < selectedItems.length) app.executeMenuCommand("group");
    }

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // ダイアログの位置と不透明度（再利用パーツ） / Dialog position and opacity (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内に貼る。
    //    識別子は DIALOG_* / prepareDialogWindow / *DialogLeft* / getSelectionViewSpan の名前
    // 2. スクリプトの先頭（#target の次の行）に #targetengine "<SCRIPT_NAME>Engine" を置く。
    //    #targetengine が無いと $.global が実行ごとに消え、位置を覚えられない。すでにあればそのまま使う
    // 3. ダイアログの show() の直前で prepareDialogWindow(dialog, SCRIPT_NAME) を呼ぶ。
    //    それまでに入れた onShow / onMove / onClose はそのまま生かし、あとに位置の復元・記録をつなぐ
    //      prepareDialogWindow(mainDialog, SCRIPT_NAME);
    //      var dialogResult = mainDialog.show();
    //    同じスクリプトで複数のダイアログを開くときは、2つ目以降のキーを変える（SCRIPT_NAME + "_colorPicker" など）
    //    同じダイアログを何度も開くときも、毎回 show() の直前で呼んでよい（2回目からは選択範囲を測り直すだけ）
    // 4. 初めて開くとき（記録が無いとき）は、スクリプト側の配置（中央・オフセットなど）がそのまま効く
    // 5. 開く位置が選択中のオブジェクトに重なりそうなら左右の反対側へずらす（Illustrator のみ）。
    //    ずらした位置は記録せず、ユーザーが動かしたときだけ記録する
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

    var DIALOG_OPACITY = 0.97;       /* ダイアログの不透明度 / dialog opacity */
    var DIALOG_AVOID_MARGIN = 60;    /* 選択範囲の推定位置の両側に取る余裕（px）/ margin on each side of the estimated selection (px) */
    var DIALOG_AVOID_MAX_ITEMS = 100; /* 選択範囲を測るオブジェクトの上限 / max items measured for the selection bounds */

    /**
     * ダイアログの不透明度を設定し、前回閉じた位置で開いて、動かした位置を記録するようにする。
     * 開く位置が選択中のオブジェクトに重なりそうなときは、左右の反対側へずらす（Illustrator のみ）。
     * 既存の onShow / onMove / onClose は先に呼んでから、位置の復元・記録を行う。
     * @param {Window} dialog - 対象のダイアログ
     * @param {string} storageKey - 位置を覚えるキー（ふつうは SCRIPT_NAME）
     * @returns {void}
     */
    function prepareDialogWindow(dialog, storageKey) {
        /* 同じダイアログを開き直すときは、選択範囲を測り直すだけにする（ハンドラーを重ねない）
           When the same dialog is shown again, only re-measure the selection (don't stack handlers) */
        if (dialog.dialogWindowState) {
            dialog.dialogWindowState.selectionSpan = getSelectionViewSpan();
            dialog.dialogWindowState.avoidedLocation = null;
            return;
        }
        var locationKey = "__" + storageKey + "_DialogLocation";
        var previousOnShow = dialog.onShow;
        var previousOnMove = dialog.onMove;
        var previousOnClose = dialog.onClose;
        var windowState = {
            selectionSpan: getSelectionViewSpan(), /* 選択範囲は show() の前に測る / measured before show() */
            screenWidth: null,                     /* 最初に開いたときに推定する / estimated on the first show */
            avoidedLocation: null                  /* 避けるためにずらした位置（記録しない）/ location set to avoid the selection (not remembered) */
        };
        dialog.dialogWindowState = windowState;

        dialog.opacity = DIALOG_OPACITY;

        /* 今の位置を記録する / Remember the current location */
        function rememberDialogLocation() {
            var currentLocation = [dialog.location[0], dialog.location[1]];
            var avoidedLocation = windowState.avoidedLocation;
            if (avoidedLocation && currentLocation[0] === avoidedLocation[0] && currentLocation[1] === avoidedLocation[1]) return;
            $.global[locationKey] = currentLocation;
        }

        dialog.onShow = function () {
            /* 最初に開くときの既定の位置は画面の横中央なので、画面の幅を逆算できる。2回目からは前回の位置なので使い回す
               On the first show the default location is centered horizontally, which gives the screen width; reuse it afterwards */
            if (windowState.screenWidth === null) windowState.screenWidth = dialog.location[0] * 2 + dialog.bounds.width;
            if (previousOnShow) previousOnShow.apply(this, arguments);
            /* $.screens は実際の画面の大きさと合わない（Mac で 1280×524 など）ので、画面内かは判定しない
               $.screens does not match the real display (e.g. 1280x524 on a Mac), so no on-screen check */
            var savedLocation = $.global[locationKey];
            if (savedLocation) dialog.location = [savedLocation[0], savedLocation[1]];
            if (windowState.selectionSpan) {
                var avoidLeft = findDialogLeftAvoidingSelection(dialog.location[0], dialog.bounds.width, windowState.screenWidth, windowState.selectionSpan);
                if (avoidLeft !== null) {
                    dialog.location = [avoidLeft, dialog.location[1]];
                    /* 代入後の値で比べる（丸められることがある）/ Compare with the value after assignment, which may be rounded */
                    windowState.avoidedLocation = [dialog.location[0], dialog.location[1]];
                }
            }
        };
        dialog.onMove = function () {
            if (previousOnMove) previousOnMove.apply(this, arguments);
            rememberDialogLocation();
        };
        dialog.onClose = function () {
            rememberDialogLocation();
            /* false を返すと閉じるのを取りやめるので、戻り値は元の onClose のものを返す
               Returning false cancels the close, so pass the original onClose result through */
            if (previousOnClose) return previousOnClose.apply(this, arguments);
        };
    }

    /**
     * 選択中のオブジェクトが、ドキュメントの表示域の左端から画面上で何 px の範囲にあるかを返す。
     * @returns {{left: number, right: number, viewWidth: number}|null} 選択が無い・測れないときは null
     */
    function getSelectionViewSpan() {
        try {
            if (app.name !== "Adobe Illustrator" || !app.documents.length) return null;
            var targetDoc = app.activeDocument;
            var selectedItems = targetDoc.selection;
            if (!selectedItems || !selectedItems.length || !selectedItems[0].visibleBounds) return null;
            var itemCount = Math.min(selectedItems.length, DIALOG_AVOID_MAX_ITEMS);
            var spanLeft = Infinity;
            var spanRight = -Infinity;
            for (var i = 0; i < itemCount; i++) {
                var itemBounds = selectedItems[i].visibleBounds;
                if (itemBounds[0] < spanLeft) spanLeft = itemBounds[0];
                if (itemBounds[2] > spanRight) spanRight = itemBounds[2];
            }
            var activeView = targetDoc.activeView; /* 複数ウィンドウで開いていても今のウィンドウ / the current window even with multiple windows */
            var viewBounds = activeView.bounds;
            var zoom = activeView.zoom;
            var viewWidth = (viewBounds[2] - viewBounds[0]) * zoom;
            /* 表示域の外にはみ出した部分は数えない / Ignore the part outside the view */
            var left = Math.max(0, (spanLeft - viewBounds[0]) * zoom);
            var right = Math.min(viewWidth, (spanRight - viewBounds[0]) * zoom);
            if (right <= left) return null;
            return { left: left, right: right, viewWidth: viewWidth };
        } catch (e) {
            /* テキスト編集中など測れないときは避けない / Do not avoid when it cannot be measured, e.g. while editing text */
            return null;
        }
    }

    /**
     * ダイアログが選択範囲に重なるなら、重ならない左端の位置を返す。
     * 表示域は画面の横中央にあるとみなし、ずれは DIALOG_AVOID_MARGIN で吸収する。
     * @param {number} dialogLeft - 今のダイアログの左端
     * @param {number} dialogWidth - ダイアログの幅
     * @param {number} screenWidth - 画面の幅
     * @param {{left: number, right: number, viewWidth: number}} selectionSpan - getSelectionViewSpan() の結果
     * @returns {number|null} ずらした左端。重ならない・どちらにも収まらないときは null
     */
    function findDialogLeftAvoidingSelection(dialogLeft, dialogWidth, screenWidth, selectionSpan) {
        var viewLeft = (screenWidth - selectionSpan.viewWidth) / 2;
        var avoidLeft = viewLeft + selectionSpan.left - DIALOG_AVOID_MARGIN;
        var avoidRight = viewLeft + selectionSpan.right + DIALOG_AVOID_MARGIN;
        if (dialogLeft + dialogWidth <= avoidLeft || dialogLeft >= avoidRight) return null;

        var leftSideLeft = avoidLeft - dialogWidth;   /* 選択範囲の左に置くとき / placed left of the selection */
        var rightSideLeft = avoidRight;               /* 選択範囲の右に置くとき / placed right of the selection */
        var fitsLeft = leftSideLeft >= 0;
        var fitsRight = rightSideLeft + dialogWidth <= screenWidth;
        /* 選択範囲が画面の右寄りなら左へ、左寄りなら右へ逃がす / Move away from the side the selection leans to */
        var preferLeft = (avoidLeft + avoidRight) / 2 > screenWidth / 2;
        if (preferLeft && fitsLeft) return leftSideLeft;
        if (fitsRight) return rightSideLeft;
        if (fitsLeft) return leftSideLeft;
        return null;
    }

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // ダイアログの位置と不透明度（再利用パーツ）ここまで / End of the reusable dialog position and opacity
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * パネルを追加する
     * @param {Window} parentWindow - 追加先のダイアログ
     * @param {string} panelTitle - パネルの見出し
     * @returns {Panel} 追加したパネル
     */
    function addPanel(parentWindow, panelTitle) {
        var addedPanel = parentWindow.add("panel", undefined, panelTitle);
        addedPanel.orientation = "column";
        addedPanel.alignChildren = "left";
        addedPanel.margins = PANEL_MARGINS;
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
     * @param {boolean} hasGroup - 選択にグループが含まれるか（無ければ「外側のグループを残す」は選べない）
     * @returns {SimplifyOptions|null} キャンセル時は null
     */
    function showDialog(hasGroup) {
        var simplifyDialog = new Window("dialog", getLabel(LABELS.dialog.title) + " " + SCRIPT_VERSION);
        simplifyDialog.orientation = "column";
        simplifyDialog.alignChildren = "fill";
        simplifyDialog.margins = DIALOG_MARGINS;

        var methodPanel = addPanel(simplifyDialog, getLabel(LABELS.panel.method));
        var keepOuterRadio = addLabeledControl(methodPanel, "radiobutton", "keepOuter");
        var rebuildRadio = addLabeledControl(methodPanel, "radiobutton", "rebuild");
        keepOuterRadio.enabled = hasGroup;
        keepOuterRadio.value = hasGroup && DEFAULT_METHOD === "keepOuter";
        rebuildRadio.value = !keepOuterRadio.value;

        var optionsPanel = addPanel(simplifyDialog, getLabel(LABELS.panel.options));
        var keepClipGroupsCheckbox = addLabeledControl(optionsPanel, "checkbox", "keepClipGroups");
        keepClipGroupsCheckbox.value = DEFAULT_KEEP_CLIP_GROUPS;
        var skipLockedHiddenCheckbox = addLabeledControl(optionsPanel, "checkbox", "skipLockedHidden");
        skipLockedHiddenCheckbox.value = DEFAULT_SKIP_LOCKED_HIDDEN;

        /* ボタンエリア（右寄せ） / Button row (right-aligned) */
        var btnRowGroup = simplifyDialog.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.alignment = ["right", "bottom"];
        btnRowGroup.alignChildren = ["right", "center"];
        var btnCancel = btnRowGroup.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });
        var btnOK = btnRowGroup.add("button", undefined, getLabel(LABELS.button.ok), { name: "ok" });

        prepareDialogWindow(simplifyDialog, SCRIPT_NAME);
        if (simplifyDialog.show() !== 1) return null;
        return {
            method: keepOuterRadio.value ? "keepOuter" : "rebuild",
            keepClipGroups: keepClipGroupsCheckbox.value,
            skipLockedHidden: skipLockedHiddenCheckbox.value
        };
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * ダイアログを表示して、選ばれた方法でグループを整理する
     * @returns {void}
     */
    function main() {
        var selectedItems = getValidSelection();
        if (!selectedItems) return;

        var groupIndexes = findGroupIndexes(selectedItems);
        /* グループを含まない1つだけの選択は、まとめ直すものがない / A lone non-group object has nothing to simplify */
        if (groupIndexes.length === 0 && selectedItems.length < 2) return;

        var simplifyOptions = showDialog(groupIndexes.length > 0);
        if (!simplifyOptions) return;

        var doc = app.activeDocument;
        if (simplifyOptions.method === "rebuild") {
            rebuildAsOneGroup(doc, simplifyOptions);
        } else {
            simplifyKeepingOuter(doc, selectedItems, groupIndexes, simplifyOptions);
        }
    }

    main();

})();
