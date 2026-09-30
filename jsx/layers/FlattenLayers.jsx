#target illustrator
#targetengine "FlattenLayersEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

除外名を持つレイヤーを残しつつ、その他のレイヤー配下のオブジェクトを指定したレイヤーへ集約してフラット化します。
ロック／非表示の扱いやガイドの行き先は、ダイアログで個別に指定できます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/FlattenLayers.md

### Overview

Flattens the document by moving the objects under every layer except the excluded ones into a single target layer.
How locked and hidden items are treated, and where guides end up, are chosen in a dialog.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/FlattenLayers.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "FlattenLayers";                /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.7.9";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-04-14";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-30";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/FlattenLayers.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/FlattenLayers.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // ローカライズ（再利用パーツ） / Localization (reusable)

    /**
     * UI の言語を返す（"ja" で始まるロケールは日本語、それ以外は英語）
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        return (String($.locale || "").indexOf("ja") === 0) ? "ja" : "en";
    }

    var uiLang = getCurrentLang();

    /**
     * LABELS から今の UI 言語の文言を取り出す。
     * @param {string|Object} labelRef - "dialog.title" のようなパス、または { ja, en }
     * @param {Object|Array} [placeholderValues] - { name: 値 } なら {name} を、[値, …] なら %1, %2 … を差し込む
     * @returns {string} 文言。パスが見つからなければパスの文字列、{ ja, en } が無ければ空文字
     */
    function getLabel(labelRef, placeholderValues) {
        var labelEntry = labelRef;
        if (typeof labelRef === "string") {
            var labelPathKeys = labelRef.split(".");
            labelEntry = LABELS;
            for (var i = 0; i < labelPathKeys.length && labelEntry != null; i++) {
                labelEntry = labelEntry[labelPathKeys[i]];
            }
        }
        var labelString;
        if (typeof labelEntry === "string") labelString = labelEntry;
        else if (labelEntry != null && labelEntry[uiLang] != null) labelString = labelEntry[uiLang];
        else if (labelEntry != null && labelEntry.en != null) labelString = labelEntry.en;
        else return (typeof labelRef === "string") ? labelRef : "";
        return fillLabelPlaceholders(String(labelString), placeholderValues);
    }

    /**
     * 項目名の文言の末尾にコロンを付ける（日本語は全角「：」、英語は半角「:」）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {Object|Array} [placeholderValues] - getLabel と同じ
     * @returns {string} コロン付きの文言
     */
    function labelText(labelRef, placeholderValues) {
        return getLabel(labelRef, placeholderValues) + (uiLang === "ja" ? "：" : ":");
    }

    /**
     * 「項目名：値」の1行を返す（日本語は「件数：5」、英語は「Count: 5」とコロンのあとに空白を入れる）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {string|number} value - コロンのあとに続ける値
     * @returns {string} 項目名と値をつないだ文字列
     */
    function labelValueText(labelRef, value) {
        return labelText(labelRef) + (uiLang === "ja" ? "" : " ") + value;
    }

    /**
     * 文言の {name} や %1 に値を差し込む
     * @param {string} labelString - 文言
     * @param {Object|Array} [placeholderValues] - { name: 値 } または [値, …]
     * @returns {string} 差し込んだ文言
     */
    function fillLabelPlaceholders(labelString, placeholderValues) {
        if (placeholderValues == null) return labelString;
        if (placeholderValues instanceof Array) {
            /* 大きい番号から置き換え、%1 が %10 の一部を置き換えないようにする / Replace from the highest index so %1 does not eat into %10 */
            for (var i = placeholderValues.length; i >= 1; i--) {
                labelString = labelString.split("%" + i).join(String(placeholderValues[i - 1]));
            }
            return labelString;
        }
        for (var placeholderKey in placeholderValues) {
            if (!placeholderValues.hasOwnProperty(placeholderKey)) continue;
            labelString = labelString.split("{" + placeholderKey + "}").join(String(placeholderValues[placeholderKey]));
        }
        return labelString;
    }

    // ローカライズ（再利用パーツ）ここまで / End of the reusable localization

    var LABELS = {
        dialogTitle: {
            ja: 'レイヤー統合（フラット化）',
            en: 'Flatten Layers'
        },
        options: {
            ja: '後処理',
            en: 'Cleanup'
        },
        process: {
            ja: '処理',
            en: 'Process'
        },
        promoteSublayers: {
            ja: 'サブレイヤーを上位レベルのレイヤーに変更',
            en: 'Move sublayers to top-level layers'
        },
        guides: {
            ja: 'ガイド',
            en: 'Guides'
        },
        separateGuides: {
            ja: '別レイヤーに移動',
            en: 'Move to another layer'
        },
        keepGuidesInCurrentLayer: {
            ja: '現在のレイヤーに保持',
            en: 'Keep in the current layer'
        },
        integrateGuides: {
            ja: '統合',
            en: 'Integrate'
        },
        exclude: {
            ja: '対象外にする',
            en: 'Exclude'
        },
        destination: {
            ja: 'まとめ先',
            en: 'Destination'
        },
        layerName: {
            ja: 'レイヤー名',
            en: 'Layer name'
        },
        layerColor: {
            ja: 'レイヤーカラー',
            en: 'Layer color'
        },
        reuseExistingMergedLayer: {
            ja: '既存のまとめ先を再利用',
            en: 'Reuse existing destination layer'
        },
        lockedPanelTitle: {
            ja: 'ロック',
            en: 'Locked'
        },
        layersPanelTitle: {
            ja: 'レイヤー',
            en: 'Layers'
        },
        hiddenPanelTitle: {
            ja: '非表示',
            en: 'Hidden'
        },
        slashSlashLayer: {
            ja: '「//」レイヤー',
            en: '“//” layers'
        },
        objectsPanelTitle: {
            ja: 'オブジェクト',
            en: 'Objects'
        },
        skipLockedLayers: {
            ja: 'レイヤー',
            en: 'Layers'
        },
        skipHiddenLayers: {
            ja: 'レイヤー',
            en: 'Layers'
        },
        skipLockedObjects: {
            ja: 'オブジェクト',
            en: 'Objects'
        },
        skipHiddenObjects: {
            ja: 'オブジェクト',
            en: 'Objects'
        },
        toggleAllExclusions: {
            ja: '対象外設定を一括切替',
            en: 'Toggle all exclusions'
        },
        includeGuidesFromExcludedLayers: {
            ja: '除外レイヤーのガイドも対象にする',
            en: 'Include guides from excluded layers'
        },
        tipPromoteSublayers: {
            ja: 'サブレイヤーを親から出して、トップレベルのレイヤーに並べ直します。',
            en: 'Pulls sublayers out of their parents and lists them as top-level layers.'
        },
        tipIntegrateGuides: {
            ja: 'ガイドもまとめ先のレイヤーへ移します。',
            en: 'Moves the guides into the destination layer too.'
        },
        tipKeepGuidesInCurrentLayer: {
            ja: 'ガイドは元のレイヤーに残します。',
            en: 'Leaves the guides on their original layers.'
        },
        tipSeparateGuides: {
            ja: 'ガイドだけを別のレイヤーにまとめます。名前は右の欄で決めます。',
            en: 'Collects the guides onto a layer of their own. The field on the right names it.'
        },
        tipIncludeGuidesFromExcludedLayers: {
            ja: 'ロックや非表示などで除外したレイヤーにあるガイドも、まとめる対象にします。',
            en: 'Also collects guides from layers excluded as locked or hidden.'
        },
        tipMergedLayerName: {
            ja: 'まとめ先のレイヤー名です。',
            en: 'Name of the destination layer.'
        },
        tipReuseExistingMergedLayer: {
            ja: '同じ名前のレイヤーがあれば作り直さず、そこへまとめます。',
            en: 'Reuses a layer of that name instead of creating a new one.'
        },
        tipLayerColor: {
            ja: 'まとめ先レイヤーの色をRGBで指定します（例: 79,127,255）。',
            en: 'Color of the destination layer, as RGB (for example 79,127,255).'
        },
        tipSkipLocked: {
            ja: 'ロックされたレイヤーはまとめません。',
            en: 'Leaves locked layers out of the merge.'
        },
        tipSkipHidden: {
            ja: '非表示のレイヤーはまとめません。',
            en: 'Leaves hidden layers out of the merge.'
        },
        tipSkipSlashSlash: {
            ja: '名前が「//」で始まるレイヤーはまとめません。作業用レイヤーを残すのに使います。',
            en: 'Leaves layers whose name starts with "//" out of the merge, so scratch layers survive.'
        },
        deleteEmptyLayers: {
            ja: '空のレイヤー／サブレイヤーを削除',
            en: 'Delete empty layers / sublayers'
        },
        cancel: {
            ja: 'キャンセル',
            en: 'Cancel'
        },
        ok: {
            ja: 'OK',
            en: 'OK'
        }
    };

    function createProcessStats() {
        return {
            moveFailureCount: 0,
            deleteFailureCount: 0,
            visibilityRestoreFailureCount: 0,
            layerLockRestoreFailureCount: 0,
            itemLockRestoreFailureCount: 0,
            topLevelPromotionFailureCount: 0,
            guideSeparationFailureCount: 0,
            parentAccessFailureCount: 0
        };
    }

    // ダイアログの位置と不透明度（再利用パーツ） / Dialog position and opacity (reusable)

    var DIALOG_OPACITY = 0.98;       /* ダイアログの不透明度 / dialog opacity */
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
            /* 文字ツールで文字を選択しているときは TextRange が返り、[0] が無い / Selecting characters with the Type tool returns a TextRange, which has no [0] */
            if (!selectedItems || selectedItems.typename === "TextRange" || !selectedItems.length || !selectedItems[0].visibleBounds) return null;
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

    // ダイアログの位置と不透明度（再利用パーツ）ここまで / End of the reusable dialog position and opacity

    // ボタン行（再利用パーツ） / Button row (reusable)

    var BUTTON_ROW_TOP_MARGIN = 5; /* ボタン行の上の余白 / top margin of the button row */
    var BUTTON_ROW_SPACING = 10;   /* ボタンどうしの間隔 / spacing between buttons */
    var BUTTON_ROW_CENTER_MAX_WIDTH = 200; /* 右のボタンだけの行を中央に置く、ダイアログの内側の最大幅（px、左右の余白を除く）。広いダイアログは右揃え / max inner dialog width (px, margins excluded) that centers a right-only row; wider dialogs keep it right-aligned */

    /**
     * ダイアログ下部のボタン行を作る。
     * 通常は「左のグループ・伸びるスペーサー・右のグループ」、centered なら行そのものを左右中央に置く
     * @param {Window|Group|Panel} parent - 行を足す先（ふつうはダイアログ）
     * @param {Object} [rowOptions] - { centered: true } で左右中央に並べる
     * @returns {{rowGroup: Group, leftGroup: Group|null, rightGroup: Group|null}} 行と左右のグループ（centered のときは左右が null）
     */
    function addButtonRow(parent, rowOptions) {
        var isCentered = !!(rowOptions && rowOptions.centered);
        var btnRowGroup = parent.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.margins = [0, BUTTON_ROW_TOP_MARGIN, 0, 0];
        btnRowGroup.spacing = BUTTON_ROW_SPACING;

        if (isCentered) {
            btnRowGroup.alignment = ["center", "bottom"];
            btnRowGroup.alignChildren = ["center", "center"];
            return { rowGroup: btnRowGroup, leftGroup: null, rightGroup: null };
        }

        btnRowGroup.alignment = ["fill", "bottom"];

        var btnLeftGroup = btnRowGroup.add("group");
        btnLeftGroup.alignChildren = ["left", "center"];
        btnLeftGroup.spacing = BUTTON_ROW_SPACING;

        /* 余りの幅を吸って、右のグループを右端に寄せる / Absorbs the extra width so the right group sits at the right edge */
        var spacer = btnRowGroup.add("group");
        spacer.alignment = ["fill", "fill"];
        spacer.minimumSize.width = 0;

        var btnRightGroup = btnRowGroup.add("group");
        btnRightGroup.alignChildren = ["right", "center"];
        btnRightGroup.spacing = BUTTON_ROW_SPACING;

        return { rowGroup: btnRowGroup, leftGroup: btnLeftGroup, rightGroup: btnRightGroup };
    }

    /**
     * 左のグループにボタンが無い（右のボタンだけの）行を、ダイアログの幅に合わせて揃える。
     * 内側の幅（左右の余白を除く）が BUTTON_ROW_CENTER_MAX_WIDTH 以下なら左右中央、それより広ければ右揃えのまま。
     * 幅はレイアウトが決まるまで分からないので、ダイアログを表示した時点（show イベント）で判定する。
     * ボタンをすべて足したあと、show() の前に呼ぶ。centered で作った行や、左にボタンがある行はそのまま
     * @param {{rowGroup: Group, leftGroup: Group|null, rightGroup: Group|null}} buttonRow - addButtonRow() の戻り値
     * @returns {void}
     */
    function alignRightOnlyButtonRow(buttonRow) {
        if (!buttonRow.leftGroup || buttonRow.leftGroup.children.length > 0) return;
        var dialogWindow = buttonRow.rowGroup.window;
        dialogWindow.addEventListener("show", function () {
            if (!buttonRow.leftGroup) return;
            var btnRowGroup = buttonRow.rowGroup;
            /* 行の幅＝ダイアログの内側の幅（左右の余白を除く）/ The row spans the dialog's inner width (margins excluded) */
            if (!btnRowGroup.size || btnRowGroup.size.width > BUTTON_ROW_CENTER_MAX_WIDTH) return;
            /* 左のグループとスペーサーを外し、右のグループだけを中央に置く / Drop the left group and the spacer so only the right group remains, centered */
            btnRowGroup.remove(buttonRow.leftGroup);
            btnRowGroup.remove(btnRowGroup.children[0]); /* 左のグループを外すと先頭はスペーサー / the spacer is first once the left group is gone */
            btnRowGroup.alignment = ["center", "bottom"];
            btnRowGroup.alignChildren = ["center", "center"];
            buttonRow.leftGroup = null;
            dialogWindow.layout.layout(true);
        });
    }

    // ボタン行（再利用パーツ）ここまで / End of the reusable button row

    function showOptionsDialog(documentRef, hasExistingMergedLayer) {
        var dlg = new Window('dialog', getLabel('dialogTitle') + ' ' + SCRIPT_VERSION);
        dlg.orientation = 'column';
        dlg.alignChildren = 'fill';

        var processPanel = dlg.add('panel', undefined, getLabel('process'));
        processPanel.orientation = 'column';
        processPanel.alignChildren = 'fill';
        processPanel.margins = [15, 20, 15, 10];

        var cbPromoteSublayers = processPanel.add('checkbox', undefined, getLabel('promoteSublayers'));
        cbPromoteSublayers.helpTip = getLabel('tipPromoteSublayers');
        cbPromoteSublayers.value = true;

        function documentHasAnyLockedLayers(container) {
            var layers = container.layers;
            for (var i = 0; i < layers.length; i++) {
                if (layerHasAnyLockedLayers(layers[i])) {
                    return true;
                }
            }
            return false;
        }

        function layerHasAnyLockedLayers(layerRef) {
            if (layerRef.locked) {
                return true;
            }

            var childLayers = layerRef.layers;
            for (var i = 0; i < childLayers.length; i++) {
                if (layerHasAnyLockedLayers(childLayers[i])) {
                    return true;
                }
            }

            return false;
        }

        function documentHasAnyHiddenLayers(container) {
            var layers = container.layers;
            for (var i = 0; i < layers.length; i++) {
                if (layerHasAnyHiddenLayers(layers[i])) {
                    return true;
                }
            }
            return false;
        }

        function layerHasAnyHiddenLayers(layerRef) {
            if (!layerRef.visible) {
                return true;
            }

            var childLayers = layerRef.layers;
            for (var i = 0; i < childLayers.length; i++) {
                if (layerHasAnyHiddenLayers(childLayers[i])) {
                    return true;
                }
            }

            return false;
        }

        function documentHasAnyHiddenObjects(container) {
            var layers = container.layers;
            for (var i = 0; i < layers.length; i++) {
                if (layerHasAnyHiddenObjects(layers[i])) {
                    return true;
                }
            }
            return false;
        }

        function documentHasAnySlashSlashLayers(container) {
            var layers = container.layers;
            for (var i = 0; i < layers.length; i++) {
                if (layerHasAnySlashSlashLayers(layers[i])) {
                    return true;
                }
            }
            return false;
        }

        function layerHasAnySlashSlashLayers(layerRef) {
            if (isSlashSlashLayer(layerRef)) {
                return true;
            }

            var childLayers = layerRef.layers;
            for (var i = 0; i < childLayers.length; i++) {
                if (layerHasAnySlashSlashLayers(childLayers[i])) {
                    return true;
                }
            }

            return false;
        }

        function layerHasAnyHiddenObjects(layerRef) {
            var pageItems = layerRef.pageItems;
            for (var i = 0; i < pageItems.length; i++) {
                if (pageItems[i].parent === layerRef && isActuallyHiddenItem(pageItems[i])) {
                    return true;
                }
            }

            var childLayers = layerRef.layers;
            for (var j = 0; j < childLayers.length; j++) {
                if (layerHasAnyHiddenObjects(childLayers[j])) {
                    return true;
                }
            }

            return false;
        }

        function isActuallyHiddenItem(item) {
            try {
                return item.hidden === true;
            } catch (e) {
                return false;
            }
        }

        function documentHasAnySublayers(container) {
            var layers = container.layers;
            for (var i = 0; i < layers.length; i++) {
                if (layers[i].parent && layers[i].parent.typename === 'Layer') {
                    return true;
                }
                if (layers[i].layers && layers[i].layers.length > 0) {
                    return true;
                }
                if (documentHasAnySublayers(layers[i])) {
                    return true;
                }
            }
            return false;
        }

        function documentHasAnyGuides(container) {
            var layers = container.layers;
            for (var i = 0; i < layers.length; i++) {
                if (layerHasAnyGuides(layers[i])) {
                    return true;
                }
            }
            return false;
        }

        function layerHasAnyGuides(layerRef) {
            var pageItems = layerRef.pageItems;
            for (var i = 0; i < pageItems.length; i++) {
                if (pageItems[i].parent === layerRef && isGuideItem(pageItems[i])) {
                    return true;
                }
            }

            var childLayers = layerRef.layers;
            for (var j = 0; j < childLayers.length; j++) {
                if (layerHasAnyGuides(childLayers[j])) {
                    return true;
                }
            }

            return false;
        }

        var hasAnySublayers = documentHasAnySublayers(documentRef);
        var hasAnyLockedLayers = documentHasAnyLockedLayers(documentRef);
        var hasAnyHiddenLayers = documentHasAnyHiddenLayers(documentRef);
        var hasAnyHiddenObjects = documentHasAnyHiddenObjects(documentRef);
        var hasAnySlashSlashLayers = documentHasAnySlashSlashLayers(documentRef);
        var hasAnyGuides = documentHasAnyGuides(documentRef);
        cbPromoteSublayers.enabled = hasAnySublayers;
        if (!hasAnySublayers) cbPromoteSublayers.value = false;

        // ガイド設定パネル / Guides panel
        var guidesPanel = processPanel.add('panel', undefined, getLabel('guides'));
        guidesPanel.orientation = 'column';
        guidesPanel.alignChildren = 'left';
        guidesPanel.margins = [15, 20, 15, 10];
        guidesPanel.enabled = hasAnyGuides;

        var rbIntegrateGuides = guidesPanel.add('radiobutton', undefined, getLabel('integrateGuides'));
        rbIntegrateGuides.helpTip = getLabel('tipIntegrateGuides');
        rbIntegrateGuides.value = false;

        var rbKeepGuidesInCurrentLayer = guidesPanel.add('radiobutton', undefined, getLabel('keepGuidesInCurrentLayer'));
        rbKeepGuidesInCurrentLayer.helpTip = getLabel('tipKeepGuidesInCurrentLayer');
        rbKeepGuidesInCurrentLayer.value = false;

        var separateGuidesGroup = guidesPanel.add('group');
        separateGuidesGroup.orientation = 'row';
        separateGuidesGroup.alignChildren = ['left', 'center'];

        var rbSeparateGuides = separateGuidesGroup.add('radiobutton', undefined, getLabel('separateGuides'));
        rbSeparateGuides.helpTip = getLabel('tipSeparateGuides');
        rbSeparateGuides.value = true;

        var etGuideLayerName = separateGuidesGroup.add('edittext', undefined, '_guide');
        etGuideLayerName.helpTip = getLabel('tipSeparateGuides');
        etGuideLayerName.characters = 12;

        function documentHasAnyGuidesInExcludedLayers(container) {
            var layers = container.layers;
            for (var i = 0; i < layers.length; i++) {
                if (isExcludedLayer(layers[i].name) && layerHasAnyGuides(layers[i])) {
                    return true;
                }
            }
            return false;
        }

        var hasGuidesInExcludedLayers = documentHasAnyGuidesInExcludedLayers(documentRef);

        var cbIncludeGuidesFromExcludedLayers = guidesPanel.add('checkbox', undefined, getLabel('includeGuidesFromExcludedLayers'));
        cbIncludeGuidesFromExcludedLayers.helpTip = getLabel('tipIncludeGuidesFromExcludedLayers');
        cbIncludeGuidesFromExcludedLayers.value = hasGuidesInExcludedLayers;
        cbIncludeGuidesFromExcludedLayers.enabled = hasGuidesInExcludedLayers && rbSeparateGuides.value;

        function updateGuideOptionsState(selected) {
            if (selected === 'integrate' && rbIntegrateGuides.value) {
                rbKeepGuidesInCurrentLayer.value = false;
                rbSeparateGuides.value = false;
            } else if (selected === 'keep' && rbKeepGuidesInCurrentLayer.value) {
                rbIntegrateGuides.value = false;
                rbSeparateGuides.value = false;
            } else if (selected === 'separate' && rbSeparateGuides.value) {
                rbIntegrateGuides.value = false;
                rbKeepGuidesInCurrentLayer.value = false;
            }

            etGuideLayerName.enabled = rbSeparateGuides.value;
            cbIncludeGuidesFromExcludedLayers.enabled = hasGuidesInExcludedLayers && rbSeparateGuides.value;
            if (!rbSeparateGuides.value) {
                cbIncludeGuidesFromExcludedLayers.value = false;
            }
        }

        rbIntegrateGuides.onClick = function () {
            updateGuideOptionsState('integrate');
        };
        rbKeepGuidesInCurrentLayer.onClick = function () {
            updateGuideOptionsState('keep');
        };
        rbSeparateGuides.onClick = function () {
            updateGuideOptionsState('separate');
        };

        if (hasAnyGuides) {
            rbIntegrateGuides.value = false;
            rbKeepGuidesInCurrentLayer.value = false;
            rbSeparateGuides.value = true;
            updateGuideOptionsState('separate');
        } else {
            rbIntegrateGuides.value = true;
            rbKeepGuidesInCurrentLayer.value = false;
            rbSeparateGuides.value = false;
            updateGuideOptionsState('integrate');
        }

        var destPanel = dlg.add('panel', undefined, getLabel('destination'));
        destPanel.orientation = 'column';
        destPanel.alignChildren = 'left';
        destPanel.margins = [15, 20, 15, 10];

        var nameGroup = destPanel.add('group');
        nameGroup.orientation = 'row';
        nameGroup.alignChildren = ['left', 'center'];
        var nameLabel = nameGroup.add('statictext', undefined, getLabel('layerName'));
        var etLayerName = nameGroup.add('edittext', undefined, '_mergedLayer');
        etLayerName.helpTip = getLabel('tipMergedLayerName');
        etLayerName.characters = 19;

        var cbReuseExistingMergedLayer = destPanel.add('checkbox', undefined, getLabel('reuseExistingMergedLayer'));
        cbReuseExistingMergedLayer.helpTip = getLabel('tipReuseExistingMergedLayer');
        cbReuseExistingMergedLayer.value = hasExistingMergedLayer;
        cbReuseExistingMergedLayer.enabled = hasExistingMergedLayer;

        function updateReuseExistingMergedLayerState() {
            var layerName = normalizeLayerName(etLayerName.text, '');
            var hasMatchingLayer = false;

            if (layerName !== '') {
                try {
                    documentRef.layers.getByName(layerName);
                    hasMatchingLayer = true;
                } catch (e) {
                    hasMatchingLayer = false;
                }
            }

            cbReuseExistingMergedLayer.enabled = hasMatchingLayer;
            if (!hasMatchingLayer) {
                cbReuseExistingMergedLayer.value = false;
            }
        }

        var colorGroup = destPanel.add('group');
        colorGroup.orientation = 'row';
        colorGroup.alignChildren = ['left', 'center'];
        var colorLabel = colorGroup.add('statictext', undefined, getLabel('layerColor'));
        var colorSwatch = colorGroup.add('panel');
        colorSwatch.preferredSize = [14, 14];
        colorSwatch.minimumSize = [14, 14];
        var etLayerColor = colorGroup.add('edittext', undefined, '79,127,255');
        etLayerColor.helpTip = getLabel('tipLayerColor');
        etLayerColor.characters = 12;

        function updateColorSwatch() {
            var colorValues = parseLayerColorValue(etLayerColor.text);
            colorSwatch.onDraw = function () {
                var g = this.graphics;
                var brush = g.newBrush(g.BrushType.SOLID_COLOR, [colorValues[0] / 255, colorValues[1] / 255, colorValues[2] / 255, 1]);
                var pen = g.newPen(g.PenType.SOLID_COLOR, [0.4, 0.4, 0.4, 1], 1);
                g.rectPath(0, 0, this.size[0], this.size[1]);
                g.fillPath(brush);
                g.strokePath(pen);
            };
            if (colorSwatch.parent && colorSwatch.parent.layout) {
                colorSwatch.parent.layout.layout(true);
            }
        }

        etLayerColor.onChanging = updateColorSwatch;
        etLayerColor.onChange = updateColorSwatch;
        updateColorSwatch();

        etLayerName.onChanging = updateReuseExistingMergedLayerState;
        etLayerName.onChange = updateReuseExistingMergedLayerState;
        updateReuseExistingMergedLayerState();

        var excludePanel = processPanel.add('panel', undefined, getLabel('exclude'));
        excludePanel.orientation = 'column';
        excludePanel.alignChildren = 'fill';
        excludePanel.margins = [15, 20, 15, 10];

        var excludeGroup = excludePanel.add('group');
        excludeGroup.orientation = 'row';
        excludeGroup.alignChildren = ['left', 'top'];
        excludeGroup.spacing = 15;

        // 左カラム：レイヤー / Left column: layers
        var layerExcludePanel = excludeGroup.add('panel', undefined, getLabel('layersPanelTitle'));
        layerExcludePanel.orientation = 'column';
        layerExcludePanel.alignChildren = 'left';
        layerExcludePanel.margins = [15, 20, 15, 10];

        var cbSkipLocked = layerExcludePanel.add('checkbox', undefined, getLabel('lockedPanelTitle'));
        cbSkipLocked.helpTip = getLabel('tipSkipLocked');
        cbSkipLocked.value = false;
        cbSkipLocked.enabled = hasAnyLockedLayers;
        if (!hasAnyLockedLayers) cbSkipLocked.value = false;

        var cbSkipHidden = layerExcludePanel.add('checkbox', undefined, getLabel('hiddenPanelTitle'));
        cbSkipHidden.helpTip = getLabel('tipSkipHidden');
        cbSkipHidden.value = false;
        cbSkipHidden.enabled = hasAnyHiddenLayers;
        if (!hasAnyHiddenLayers) cbSkipHidden.value = false;

        var cbSkipSlashSlashLayers = layerExcludePanel.add('checkbox', undefined, getLabel('slashSlashLayer'));
        cbSkipSlashSlashLayers.helpTip = getLabel('tipSkipSlashSlash');
        cbSkipSlashSlashLayers.value = false;
        cbSkipSlashSlashLayers.enabled = hasAnySlashSlashLayers;
        if (!hasAnySlashSlashLayers) cbSkipSlashSlashLayers.value = false;

        // 右カラム：オブジェクト / Right column: objects
        var objectExcludePanel = excludeGroup.add('panel', undefined, getLabel('objectsPanelTitle'));
        objectExcludePanel.orientation = 'column';
        objectExcludePanel.alignChildren = 'left';
        objectExcludePanel.margins = [15, 20, 15, 10];

        var cbSkipLockedObjects = objectExcludePanel.add('checkbox', undefined, getLabel('lockedPanelTitle'));
        cbSkipLockedObjects.value = false;

        var cbSkipHiddenObjects = objectExcludePanel.add('checkbox', undefined, getLabel('hiddenPanelTitle'));
        cbSkipHiddenObjects.value = false;
        cbSkipHiddenObjects.enabled = hasAnyHiddenObjects;
        if (!hasAnyHiddenObjects) cbSkipHiddenObjects.value = false;

        var toggleAllGroup = excludePanel.add('group');
        toggleAllGroup.orientation = 'row';
        toggleAllGroup.alignment = ['left', 'top'];
        var cbToggleAllExclusions = toggleAllGroup.add('checkbox', undefined, getLabel('toggleAllExclusions'));
        cbToggleAllExclusions.value = false;

        function updateToggleAllExclusionsState() {
            var enabledValues = [];

            if (cbSkipLocked.enabled) {
                enabledValues.push(cbSkipLocked.value);
            }
            if (cbSkipHidden.enabled) {
                enabledValues.push(cbSkipHidden.value);
            }
            if (cbSkipSlashSlashLayers.enabled) {
                enabledValues.push(cbSkipSlashSlashLayers.value);
            }
            if (cbSkipLockedObjects.enabled) {
                enabledValues.push(cbSkipLockedObjects.value);
            }
            if (cbSkipHiddenObjects.enabled) {
                enabledValues.push(cbSkipHiddenObjects.value);
            }

            cbToggleAllExclusions.value = enabledValues.length > 0 && allValuesAreTrue(enabledValues);
        }

        function allValuesAreTrue(values) {
            for (var i = 0; i < values.length; i++) {
                if (!values[i]) {
                    return false;
                }
            }
            return true;
        }

        cbToggleAllExclusions.onClick = function () {
            var newValue = cbToggleAllExclusions.value;
            if (cbSkipLocked.enabled) {
                cbSkipLocked.value = newValue;
            }
            if (cbSkipHidden.enabled) {
                cbSkipHidden.value = newValue;
            }
            if (cbSkipSlashSlashLayers.enabled) {
                cbSkipSlashSlashLayers.value = newValue;
            }
            if (cbSkipLockedObjects.enabled) {
                cbSkipLockedObjects.value = newValue;
            }
            if (cbSkipHiddenObjects.enabled) {
                cbSkipHiddenObjects.value = newValue;
            }
            updateToggleAllExclusionsState();
        };

        cbSkipLocked.onClick = updateToggleAllExclusionsState;
        cbSkipHidden.onClick = updateToggleAllExclusionsState;
        cbSkipSlashSlashLayers.onClick = updateToggleAllExclusionsState;
        cbSkipLockedObjects.onClick = updateToggleAllExclusionsState;
        cbSkipHiddenObjects.onClick = updateToggleAllExclusionsState;
        updateToggleAllExclusionsState();

        var optionsPanel = dlg.add('panel', undefined, getLabel('options'));
        optionsPanel.orientation = 'column';
        optionsPanel.alignChildren = 'left';
        optionsPanel.margins = [15, 20, 15, 10];

        var cbDeleteEmpty = optionsPanel.add('checkbox', undefined, getLabel('deleteEmptyLayers'));
        cbDeleteEmpty.value = true;

        var buttonRow = addButtonRow(dlg, { centered: true });
        var btnCancel = buttonRow.rowGroup.add('button', undefined, getLabel('cancel'), { name: 'cancel' });
        var btnOK = buttonRow.rowGroup.add('button', undefined, getLabel('ok'), { name: 'ok' });

        prepareDialogWindow(dlg, SCRIPT_NAME);
        if (dlg.show() !== 1) {
            return null;
        }

        return {
            skipLockedLayers: cbSkipLocked.value,
            skipHiddenLayers: cbSkipHidden.value,
            skipSlashSlashLayers: cbSkipSlashSlashLayers.value,
            skipLockedObjects: cbSkipLockedObjects.value,
            skipHiddenObjects: cbSkipHiddenObjects.value,
            promoteSublayersToTopLevel: cbPromoteSublayers.value,
            guideMode: rbSeparateGuides.value ? 'separate' : (rbKeepGuidesInCurrentLayer.value ? 'keep' : 'integrate'),
            guideLayerName: normalizeLayerName(etGuideLayerName.text, '_guide'),
            mergedLayerName: etLayerName.text,
            deleteEmptyLayers: cbDeleteEmpty.value,
            reuseExistingMergedLayer: cbReuseExistingMergedLayer.value,
            layerColorValue: etLayerColor.text,
            includeGuidesFromExcludedLayers: cbIncludeGuidesFromExcludedLayers.value
        };
    }

    function getOrCreateMergedLayer(documentRef, mergedLayerName, reuseExistingMergedLayer) {
        var normalizedName = normalizeLayerName(mergedLayerName, '_mergedLayer');
        var mergedLayer = null;

        if (reuseExistingMergedLayer) {
            try {
                mergedLayer = documentRef.layers.getByName(normalizedName);
            } catch (e) {
                mergedLayer = null;
            }
        }

        if (mergedLayer === null) {
            mergedLayer = documentRef.layers.add();
            mergedLayer.name = reuseExistingMergedLayer ? normalizedName : generateUniqueLayerName(documentRef, normalizedName);
        }

        return mergedLayer;
    }

    function normalizeLayerName(layerName, fallbackName) {
        var normalized = layerName == null ? '' : String(layerName).replace(/^\s+|\s+$/g, '');
        return normalized !== '' ? normalized : fallbackName;
    }

    function generateUniqueLayerName(documentRef, baseName) {
        var candidate = baseName;
        var suffix = 2;

        while (layerExistsByName(documentRef, candidate)) {
            candidate = baseName + ' ' + suffix;
            suffix++;
        }

        return candidate;
    }

    function layerExistsByName(documentRef, layerName) {
        try {
            documentRef.layers.getByName(layerName);
            return true;
        } catch (e) {
            return false;
        }
    }

    function clampColorValue(value) {
        var num = parseInt(value, 10);
        if (isNaN(num)) return 0;
        if (num < 0) return 0;
        if (num > 255) return 255;
        return num;
    }

    function parseLayerColorValue(layerColorValue) {
        var defaultColor = [79, 127, 255];
        if (!layerColorValue || layerColorValue === '') {
            return defaultColor;
        }

        var parts = String(layerColorValue).split(',');
        if (parts.length !== 3) {
            return defaultColor;
        }

        return [
            clampColorValue(parts[0]),
            clampColorValue(parts[1]),
            clampColorValue(parts[2])
        ];
    }

    function applyMergedLayerSettings(mergedLayer, layerColorValue) {
        mergedLayer.visible = true;
        mergedLayer.locked = false;
        mergedLayer.printable = true;
        var colorValues = parseLayerColorValue(layerColorValue);
        var customRGBColor = new RGBColor();
        customRGBColor.red = colorValues[0];
        customRGBColor.green = colorValues[1];
        customRGBColor.blue = colorValues[2];
        mergedLayer.color = customRGBColor;
    }

    function main() {
        var documentRef;
        try {
            documentRef = app.activeDocument;
        } catch (error) {
            return;
        }

        var mergedLayerName = "_mergedLayer";
        var hasExistingMergedLayer = false;

        try {
            documentRef.layers.getByName(mergedLayerName);
            hasExistingMergedLayer = true;
        } catch (e) {
            hasExistingMergedLayer = false;
        }

        var options = showOptionsDialog(documentRef, hasExistingMergedLayer);
        if (!options) {
            return;
        }

        mergedLayerName = normalizeLayerName(options.mergedLayerName, '_mergedLayer');

        var stats = createProcessStats();
        var mergedLayer = getOrCreateMergedLayer(documentRef, mergedLayerName, options.reuseExistingMergedLayer);
        applyMergedLayerSettings(mergedLayer, options.layerColorValue);

        var allTopLayers = documentRef.layers;
        for (var i = allTopLayers.length - 1; i >= 0; i--) {
            var currentLayer = allTopLayers[i];
            if (currentLayer === mergedLayer) {
                continue;
            }
            if (isExcludedLayer(currentLayer.name)) {
                continue;
            }
            moveItemsToTargetLayer(currentLayer, mergedLayer, options, stats);
        }

        // 「現在のレイヤーに保持」時は、サブレイヤー直下のガイドを上位レイヤーへ繰り上げる。
        if (options.guideMode === 'keep') {
            hoistGuideItemsFromSublayers(documentRef, stats);
        }

        if (options.promoteSublayersToTopLevel) {
            promoteSublayersToTop(documentRef, mergedLayer, options, stats);
        }

        if (options.guideMode === 'separate') {
            separateGuidesToLayer(documentRef, mergedLayer, options, stats);
        }

        if (options.guideMode === 'separate' && options.includeGuidesFromExcludedLayers) {
            moveGuidesFromExcludedLayers(documentRef, options, stats);
        }

        if (options.deleteEmptyLayers) {
            deleteEmptyLayersRecursively(documentRef, mergedLayer, options, stats);
        }

        var messages = [];
        if (stats.moveFailureCount > 0) {
            messages.push((uiLang === 'ja' ? '移動失敗' : 'Move failures') + ': ' + stats.moveFailureCount);
        }
        if (stats.deleteFailureCount > 0) {
            messages.push((uiLang === 'ja' ? '削除失敗' : 'Delete failures') + ': ' + stats.deleteFailureCount);
        }
        if (stats.visibilityRestoreFailureCount > 0) {
            messages.push((uiLang === 'ja' ? '表示状態の復元失敗' : 'Visibility restore failures') + ': ' + stats.visibilityRestoreFailureCount);
        }
        if (stats.layerLockRestoreFailureCount > 0) {
            messages.push((uiLang === 'ja' ? 'レイヤーロック復元失敗' : 'Layer lock restore failures') + ': ' + stats.layerLockRestoreFailureCount);
        }
        if (stats.itemLockRestoreFailureCount > 0) {
            messages.push((uiLang === 'ja' ? 'オブジェクトロック復元失敗' : 'Item lock restore failures') + ': ' + stats.itemLockRestoreFailureCount);
        }
        if (stats.topLevelPromotionFailureCount > 0) {
            messages.push((uiLang === 'ja' ? '最上位化失敗' : 'Top-level promotion failures') + ': ' + stats.topLevelPromotionFailureCount);
        }
        if (stats.guideSeparationFailureCount > 0) {
            messages.push((uiLang === 'ja' ? 'ガイド分離失敗' : 'Guide separation failures') + ': ' + stats.guideSeparationFailureCount);
        }
        if (stats.parentAccessFailureCount > 0) {
            messages.push((uiLang === 'ja' ? '親レイヤーアクセス失敗' : 'Parent access failures') + ': ' + stats.parentAccessFailureCount);
        }
        if (messages.length > 0) {
            alert(messages.join('\n'));
        }
    }

    function moveGuidesFromExcludedLayers(documentRef, options, stats) {
        var guideLayerName = normalizeLayerName(options.guideLayerName, '_guide');
        var guideLayer = getOrCreateGuideLayer(documentRef, guideLayerName);

        var allTopLayers = documentRef.layers;
        for (var i = allTopLayers.length - 1; i >= 0; i--) {
            var currentLayer = allTopLayers[i];
            if (isExcludedLayer(currentLayer.name)) {
                moveGuideItemsFromLayerRecursive(currentLayer, guideLayer, stats);
            }
        }
    }

    function moveGuideItemsFromLayerRecursive(sourceLayer, destinationLayer, stats) {
        for (var i = sourceLayer.layers.length - 1; i >= 0; i--) {
            moveGuideItemsFromLayerRecursive(sourceLayer.layers[i], destinationLayer, stats);
        }

        var restoreLayerLock = false;
        var restoreVisibility = false;
        try {
            if (sourceLayer.locked) {
                sourceLayer.locked = false;
                restoreLayerLock = true;
            }
            if (!sourceLayer.visible) {
                sourceLayer.visible = true;
                restoreVisibility = true;
            }
        } catch (e) { }

        moveGuideItemsFromLayer(sourceLayer, destinationLayer, stats);

        if (restoreVisibility) {
            try { sourceLayer.visible = false; } catch (e1) {
                if (stats) stats.visibilityRestoreFailureCount++;
            }
        }
        if (restoreLayerLock) {
            try { sourceLayer.locked = true; } catch (e2) {
                if (stats) stats.layerLockRestoreFailureCount++;
            }
        }
    }

    function separateGuidesToLayer(documentRef, mergedLayer, options, stats) {
        var guideLayerName = normalizeLayerName(options.guideLayerName, '_guide');
        var guideLayer = getOrCreateGuideLayer(documentRef, guideLayerName);
        applyMergedLayerSettings(guideLayer, options.layerColorValue);

        moveGuideItemsFromLayer(mergedLayer, guideLayer, stats);
    }

    function getOrCreateGuideLayer(documentRef, guideLayerName) {
        guideLayerName = normalizeLayerName(guideLayerName, '_guide');
        try {
            return documentRef.layers.getByName(guideLayerName);
        } catch (e) {
            var guideLayer = documentRef.layers.add();
            guideLayer.name = guideLayerName;
            return guideLayer;
        }
    }

    function moveGuideItemsFromLayer(sourceLayer, destinationLayer, stats) {
        var pageItems = sourceLayer.pageItems;
        for (var i = pageItems.length - 1; i >= 0; i--) {
            var item = pageItems[i];
            if (item.parent !== sourceLayer) {
                continue;
            }
            if (!isGuideItem(item)) {
                continue;
            }

            var wasLocked = false;
            try {
                wasLocked = item.locked;
                if (wasLocked) {
                    item.locked = false;
                }
                item.move(destinationLayer, ElementPlacement.PLACEATBEGINNING);
            } catch (e) {
                if (stats) {
                    stats.guideSeparationFailureCount++;
                }
            } finally {
                try {
                    if (wasLocked) {
                        item.locked = true;
                    }
                } catch (restoreError) {
                    if (stats) {
                        stats.itemLockRestoreFailureCount++;
                    }
                }
            }
        }
    }

    function isGuideItem(item) {
        if (!item) {
            return false;
        }

        try {
            if (item.guides === true) {
                return true;
            }
        } catch (e) { }

        return false;
    }

    function hoistGuideItemsFromSublayers(documentRef, stats) {
        var topLayers = documentRef.layers;
        for (var i = 0; i < topLayers.length; i++) {
            hoistGuideItemsFromChildLayers(topLayers[i], stats);
        }
    }

    function hoistGuideItemsFromChildLayers(parentLayer, stats) {
        var childLayers = parentLayer.layers;
        for (var i = childLayers.length - 1; i >= 0; i--) {
            var childLayer = childLayers[i];
            hoistGuideItemsFromChildLayers(childLayer, stats);
            moveDirectGuideItemsToLayer(childLayer, parentLayer, stats);
        }
    }

    function moveDirectGuideItemsToLayer(sourceLayer, destinationLayer, stats) {
        var pageItems = sourceLayer.pageItems;
        for (var i = pageItems.length - 1; i >= 0; i--) {
            var item = pageItems[i];
            if (item.parent !== sourceLayer) {
                continue;
            }
            if (!isGuideItem(item)) {
                continue;
            }

            var wasLocked = false;
            try {
                wasLocked = item.locked;
                if (wasLocked) {
                    item.locked = false;
                }
                item.move(destinationLayer, ElementPlacement.PLACEATBEGINNING);
            } catch (e) {
                if (stats) {
                    stats.guideSeparationFailureCount++;
                }
            } finally {
                try {
                    if (wasLocked) {
                        item.locked = true;
                    }
                } catch (restoreError) {
                    if (stats) {
                        stats.itemLockRestoreFailureCount++;
                    }
                }
            }
        }
    }

    /*
     * 指定レイヤー配下を再帰走査し、各レイヤー直下のページアイテムだけを destinationLayer へ移動。
     * Recursively walk sourceLayer and move only the page items directly under each layer to destinationLayer.
     */
    function moveItemsToTargetLayer(sourceLayer, destinationLayer, options, stats) {
        if (sourceLayer === destinationLayer) return false;
        if (options.skipLockedLayers && sourceLayer.locked) return false;
        if (options.skipHiddenLayers && !sourceLayer.visible) return false;
        if (options.skipSlashSlashLayers && isSlashSlashLayer(sourceLayer)) return false;

        var didMove = false;

        var restoreVisibility = false;
        var originalVisibility = true;

        var restoreLayerLock = false;
        var originalLayerLock = false;

        if (!options.skipHiddenLayers && !sourceLayer.visible) {
            try {
                originalVisibility = sourceLayer.visible;
                sourceLayer.visible = true;
                restoreVisibility = true;
            } catch (e0) { }
        }

        if (!options.skipLockedLayers && sourceLayer.locked) {
            try {
                originalLayerLock = sourceLayer.locked;
                sourceLayer.locked = false;
                restoreLayerLock = true;
            } catch (e00) { }
        }

        for (var i = sourceLayer.layers.length - 1; i >= 0; i--) {
            if (moveItemsToTargetLayer(sourceLayer.layers[i], destinationLayer, options, stats)) {
                didMove = true;
            }
        }

        var pageItems = sourceLayer.pageItems;
        for (var j = pageItems.length - 1; j >= 0; j--) {
            var item = pageItems[j];
            if (item.parent !== sourceLayer) {
                continue;
            }

            var wasLocked = false;
            try {
                wasLocked = item.locked;
                if (options.skipLockedObjects && wasLocked) {
                    continue;
                }
                if (shouldSkipHiddenItem(item, options)) {
                    continue;
                }
                if (options.guideMode === 'keep' && isGuideItem(item)) {
                    continue;
                }
                if (wasLocked) item.locked = false;
                item.move(destinationLayer, ElementPlacement.PLACEATBEGINNING);
                didMove = true;
            } catch (error) {
                if (stats) {
                    stats.moveFailureCount++;
                }
                // 移動できない項目は無視 / ignore items that still cannot be moved
            } finally {
                try {
                    if (wasLocked) item.locked = true;
                } catch (e2) {
                    if (stats) {
                        stats.itemLockRestoreFailureCount++;
                    }
                }
            }
        }

        if (restoreVisibility) {
            try {
                sourceLayer.visible = originalVisibility;
            } catch (e3) {
                if (stats) {
                    stats.visibilityRestoreFailureCount++;
                }
            }
        }
        if (restoreLayerLock) {
            try {
                sourceLayer.locked = originalLayerLock;
            } catch (e4) {
                if (stats) {
                    stats.layerLockRestoreFailureCount++;
                }
            }
        }

        return didMove;
    }

    function shouldSkipHiddenItem(item, options) {
        if (!options.skipHiddenObjects) {
            return false;
        }
        if (!item.hidden) {
            return false;
        }
        if (!options.skipHiddenLayers && isHiddenOnlyByLayerVisibility(item)) {
            return false;
        }
        return true;
    }

    function isHiddenOnlyByLayerVisibility(item) {
        var current = item;
        while (current && current.parent) {
            current = current.parent;
            if (current.typename === 'Layer' && !current.visible) {
                return true;
            }
        }
        return false;
    }

    function isSlashSlashLayer(layerRef) {
        if (!layerRef || layerRef.name == null) {
            return false;
        }
        return String(layerRef.name).indexOf('//') === 0;
    }

    /* Exclude-name map for layers (case-insensitive, trimmed) */
    var EXCLUDE = {
        "bg": 1
    };

    /* Check if the layer name is excluded from merging (trim + case-insensitive) */
    function isExcludedLayer(layerName) {
        if (!layerName) return false;
        var key = String(layerName).replace(/^\s+|\s+$/g, '').toLowerCase();
        return EXCLUDE[key] === 1;
    }

    // ==========================
    // 中身が残るサブレイヤーをトップレベルへ移動
    // ==========================
    function promoteSublayersToTop(documentRef, mergedLayer, options, stats) {
        var guard = 0;
        while (guard < 1000) {
            var batch = [];
            collectDirectSublayers(documentRef, batch, mergedLayer, options);
            if (batch.length === 0) {
                break;
            }

            for (var i = 0; i < batch.length; i++) {
                moveLayerToTopLevel(documentRef, batch[i], mergedLayer, stats);
            }
            guard++;
        }
    }

    function collectDirectSublayers(container, resultArray, mergedLayer, options) {
        var layers = container.layers;

        for (var i = layers.length - 1; i >= 0; i--) {
            var currentLayer = layers[i];

            if (currentLayer === mergedLayer) {
                continue;
            }
            if (isExcludedLayer(currentLayer.name)) {
                continue;
            }
            if (options.skipLockedLayers && currentLayer.locked) {
                continue;
            }
            if (options.skipHiddenLayers && !currentLayer.visible) {
                continue;
            }
            if (options.skipSlashSlashLayers && isSlashSlashLayer(currentLayer)) {
                continue;
            }

            if (currentLayer.parent && currentLayer.parent.typename === 'Layer' && layerHasRemainingContent(currentLayer)) {
                resultArray.push(currentLayer);
            }

            if (currentLayer.layers.length > 0) {
                collectDirectSublayers(currentLayer, resultArray, mergedLayer, options);
            }
        }
    }

    function layerHasRemainingContent(layerRef) {
        return hasDirectPageItems(layerRef) || hasDirectGuideItems(layerRef) || layerRef.layers.length > 0;
    }

    function moveLayerToTopLevel(documentRef, layerRef, mergedLayer, stats) {
        if (!layerRef || !layerRef.parent || layerRef.parent.typename !== 'Layer') {
            return;
        }

        var anchor = getTopLevelMoveAnchor(documentRef, mergedLayer, layerRef);
        if (!anchor) {
            return;
        }

        try {
            layerRef.move(anchor, ElementPlacement.PLACEAFTER);
        } catch (e) {
            if (stats) {
                stats.topLevelPromotionFailureCount++;
            }
        }
    }

    function getTopLevelMoveAnchor(documentRef, mergedLayer, movingLayer) {
        var topLayers = documentRef.layers;
        for (var i = topLayers.length - 1; i >= 0; i--) {
            var candidate = topLayers[i];
            if (candidate === movingLayer) {
                continue;
            }
            if (candidate !== mergedLayer) {
                return candidate;
            }
        }
        return mergedLayer || null;
    }

    // ==========================
    // 空レイヤー削除の制御関数 / Control loop for empty-layer deletion
    // - findEmptyLayers() : 現時点で空のレイヤー／サブレイヤーを収集
    // - deleteLayers()    : 収集済みレイヤーを実際に削除
    // - この関数          : 親レイヤーが後から空になるケースに備え、空がなくなるまで反復
    // ==========================
    function deleteEmptyLayersRecursively(documentRef, mergedLayer, options, stats) {
        var guard = 0;
        while (guard < 1000) {
            var emptyLayerList = [];
            findEmptyLayers(documentRef, emptyLayerList, mergedLayer, options);
            if (emptyLayerList.length === 0) {
                break;
            }
            deleteLayers(emptyLayerList, stats);
            guard++;
        }
    }

    // ==========================
    // 空レイヤー収集関数 / Collect empty layers recursively
    // - 削除は行わず、削除候補の収集だけを担当
    // - mergedLayer と除外名レイヤーは対象外
    // - 非表示 / ロック除外はトップレベル・サブレイヤーを問わずここで判定
    // - 子を先に走査し、現時点で空のレイヤーを resultArray に積む
    // ==========================
    function findEmptyLayers(container, resultArray, mergedLayer, options) {
        var layers = container.layers;
        for (var i = 0; i < layers.length; i++) {
            var currentLayer = layers[i];

            if (currentLayer === mergedLayer) {
                continue;
            }
            if (isExcludedLayer(currentLayer.name)) {
                continue;
            }

            var skipCurrentLayerDeletion = false;

            if (options.skipLockedLayers && currentLayer.locked) {
                skipCurrentLayerDeletion = true;
            }
            if (options.skipHiddenLayers && !currentLayer.visible) {
                skipCurrentLayerDeletion = true;
            }
            if (options.skipSlashSlashLayers && isSlashSlashLayer(currentLayer)) {
                skipCurrentLayerDeletion = true;
            }

            if (currentLayer.layers.length > 0) {
                findEmptyLayers(currentLayer, resultArray, mergedLayer, options);
            }

            if (skipCurrentLayerDeletion) {
                continue;
            }

            var hasItems = hasDirectPageItems(currentLayer) || hasDirectGuideItems(currentLayer);
            var hasChildLayers = currentLayer.layers.length > 0;
            if (!hasItems && !hasChildLayers) {
                resultArray.push(currentLayer);
            }
        }
    }

    // 直属アイテム判定専用 / Check only whether the layer itself directly owns pageItems
    // 子サブレイヤー配下のアイテムは含めない。
    function hasDirectPageItems(layerRef) {
        var items = layerRef.pageItems;
        for (var i = 0; i < items.length; i++) {
            if (items[i].parent === layerRef) {
                return true;
            }
        }
        return false;
    }

    function hasDirectGuideItems(layerRef) {
        var items = layerRef.pathItems;
        for (var i = 0; i < items.length; i++) {
            if (items[i].parent === layerRef && isGuideItem(items[i])) {
                return true;
            }
        }
        return false;
    }

    // ==========================
    // 空レイヤー削除実行関数 / Delete listed layers
    // - 収集済みの削除候補だけを削除
    // - 深いサブレイヤーから先に削除
    // - 削除前に親レイヤーを一時的に visible / unlocked にして到達可能にする
    // - 収集ロジックや再試行制御は持たない
    // ==========================
    function deleteLayers(layerArray, stats) {
        var sortedLayers = layerArray.slice(0);
        sortedLayers.sort(function (a, b) {
            return getLayerDepth(b) - getLayerDepth(a);
        });

        for (var i = 0; i < sortedLayers.length; i++) {
            var layerRef = sortedLayers[i];
            var parentStates = null;
            try {
                parentStates = ensureLayerParentsAccessible(layerRef, stats);
                if (!parentStates) {
                    if (stats) {
                        stats.parentAccessFailureCount++;
                    }
                    continue;
                }
                layerRef.visible = true;
                layerRef.locked = false;
                layerRef.remove();
            } catch (e) {
                if (stats) {
                    stats.deleteFailureCount++;
                }
            } finally {
                restoreLayerAccessStates(parentStates, stats);
            }
        }
    }

    function getLayerDepth(layerRef) {
        var depth = 0;
        var current = layerRef;
        while (current && current.parent && current.parent.typename === 'Layer') {
            depth++;
            current = current.parent;
        }
        return depth;
    }

    function ensureLayerParentsAccessible(layerRef, stats) {
        var chain = [];
        var current = layerRef;
        while (current && current.parent && current.parent.typename === 'Layer') {
            current = current.parent;
            chain.unshift(current);
        }

        var states = [];
        for (var i = 0; i < chain.length; i++) {
            try {
                states.push({
                    layer: chain[i],
                    visible: chain[i].visible,
                    locked: chain[i].locked
                });
                chain[i].visible = true;
                chain[i].locked = false;
            } catch (e) {
                restoreLayerAccessStates(states, stats);
                return false;
            }
        }
        return states;
    }

    function restoreLayerAccessStates(states, stats) {
        if (!states || states.length === 0) {
            return;
        }

        for (var i = states.length - 1; i >= 0; i--) {
            try {
                states[i].layer.visible = states[i].visible;
            } catch (e1) {
                if (stats) {
                    stats.visibilityRestoreFailureCount++;
                }
            }

            try {
                states[i].layer.locked = states[i].locked;
            } catch (e2) {
                if (stats) {
                    stats.layerLockRestoreFailureCount++;
                }
            }
        }
    }

    main();

})();
