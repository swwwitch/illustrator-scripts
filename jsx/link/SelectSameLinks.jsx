#target illustrator
#targetengine "SelectSameLinksEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択中のリンク画像と同じファイルを参照する配置画像をドキュメント全体から探し、まとめて選択または削除します。
同一かどうかの判定は、リンクの絶対パスとファイル名のどちらでも行えます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SelectSameLinks.md

### Overview

Searches the whole document for placed images that reference the same file as the selection, then selects or deletes them together.
The match can be made either on the absolute path of the link or on the file name.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SelectSameLinks.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SelectSameLinks";              /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.9";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-05-20";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-01";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SelectSameLinks.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SelectSameLinks.md"; /* README (English) */

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
        tooltip: {
            matchByPath: { ja: "リンク先のフルパスが同じ画像を探します。", en: "Finds images whose full link path is the same." },
            matchByName: { ja: "フォルダーが違っても、ファイル名が同じ画像を探します。", en: "Finds images with the same file name, even in a different folder." },
            actionSelect: { ja: "見つかった画像を選択するだけで、削除はしません。", en: "Only selects the images it finds; nothing is deleted." },
            actionDeleteImage: { ja: "見つかったリンク画像を削除します。クリップグループは残ります。", en: "Deletes the linked images it finds, leaving any clipping group behind." },
            actionDeleteClipGroup: { ja: "見つかったリンク画像を、それを包むクリップグループごと削除します。", en: "Deletes the linked images together with the clipping group around them." }
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
            selected: { ja: "{count}件選択しました。", en: "{count} item(s) selected." },
            deleted: { ja: "{count}件削除しました。", en: "{count} item(s) deleted." }
        }
    };

    // =========================================
    // UIレイアウトの共通設定 / Shared UI layout
    // =========================================

    // UIレイアウト（再利用パーツ） / UI layout (reusable)

    /* ウィンドウ・パネルの余白と間隔 / Window & panel margins and spacing */
    var WINDOW_MARGINS = 16;                 /* ウィンドウ外周の余白 / window margin */
    var WINDOW_SPACING = 12;                 /* ウィンドウ内の要素間隔 / window spacing */
    var PANEL_MARGINS  = [16, 20, 16, 12];   /* パネル余白 [左,上,右,下] / panel margins */
    var PANEL_SPACING  = 12;                 /* パネル内の要素間隔 / panel spacing */
    var COLUMN_SPACING = 12;                 /* 2カラムの間隔 / gap between columns */
    var TAB_MARGINS    = [15, 20, 5, 10];    /* タブ余白 [左,上,右,下] / tab margins */

    /**
     * ウィンドウの共通設定
     * @param {Window} targetWindow - 対象のウィンドウ
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
     * パネルの共通設定（子は幅いっぱい。ボタンは alignment = "left" で広げない）
     * @param {Panel} targetPanel - 対象のパネル
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
     * タブの共通設定
     * @param {Tab} targetTab - 対象のタブ
     * @param {number} [spacing] - 要素間隔（省略時は変えない）
     * @returns {void}
     */
    function setupTab(targetTab, spacing) {
        targetTab.orientation = "column";
        targetTab.alignChildren = "fill";
        targetTab.margins = TAB_MARGINS;
        if (typeof spacing === "number") targetTab.spacing = spacing;
    }

    /**
     * 横並びの行グループの共通設定（ボタン列など）。
     * alignment と alignChildren を対で指定し、中のボタンが横に伸びたり天地がずれたりしないようにする
     * @param {Group} rowGroup - 対象のグループ
     * @param {string|string[]} [rowAlignment] - 横方向の alignment（省略時は "left"）。配列ならそのまま使う
     * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function setupRow(rowGroup, rowAlignment, spacing) {
        rowGroup.orientation = "row";
        rowGroup.alignment = (rowAlignment instanceof Array) ? rowAlignment : [rowAlignment || "left", "center"];
        rowGroup.alignChildren = ["left", "center"];
        rowGroup.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /**
     * ボタンの高さを指定した px だけ詰める（レイアウトが決まったあとに呼ぶ）
     * @param {Button} targetButton - 対象のボタン
     * @param {number} trimPixels - 詰める量（px）
     * @returns {void}
     */
    function trimButtonHeight(targetButton, trimPixels) {
        /* レイアウト前は size が無い / size is not set until the layout runs */
        if (!targetButton.size) return;
        targetButton.size = [targetButton.size.width, targetButton.size.height - trimPixels];
    }

    // UIレイアウト（再利用パーツ）ここまで / End of the reusable UI layout

    /**
     * 見出し付きのパネルを追加する
     * @param {Window|Group} parent - パネルを追加する親
     * @param {string} labelString - パネルの見出し
     * @returns {Panel} 追加したパネル
     */
    function addPanel(parent, labelString) {
        var newPanel = parent.add("panel", undefined, labelString);
        /* ラジオが縦に並ぶパネルなので詰める / Tighter spacing for stacked radio buttons */
        setupPanel(newPanel, 6);
        return newPanel;
    }

    /**
     * ラジオボタンを追加する（パネルのfillを打ち消して左揃えにする）
     * @param {Panel} parentPanel - ラジオボタンを追加するパネル
     * @param {string} labelString - 表示する文言
     * @param {boolean} isSelected - 既定で選択状態にするか
     * @param {string} [tooltipPath] - ツールチップのドット区切りキー
     * @param {boolean} isSelected - 既定で選択状態にするか
     * @param {string} [tooltipPath] - ツールチップのドット区切りキー
     * @param {boolean} isSelected - 初期状態で選択するかどうか
     * @returns {RadioButton} 追加したラジオボタン
     */
    function addRadio(parentPanel, labelString, isSelected, tooltipPath) {
        var radio = parentPanel.add("radiobutton", undefined, labelString);
        radio.alignment = "left";
        radio.value = isSelected;
        if (tooltipPath) radio.helpTip = getLabel(tooltipPath);
        return radio;
    }

    // ボタン行（再利用パーツ） / Button row (reusable)

    var BUTTON_ROW_TOP_MARGIN = 5; /* ボタン行の上の余白 / top margin of the button row */
    var BUTTON_ROW_BOTTOM_MARGIN = 14; /* ボタン行の下の余白。ダイアログの下余白と合わせて約30px（Illustrator 標準のダイアログに合わせる） / bottom margin; with the dialog margin about 30px, like Illustrator's own dialogs */
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
        btnRowGroup.margins = [0, BUTTON_ROW_TOP_MARGIN, 0, BUTTON_ROW_BOTTOM_MARGIN];
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
        addRadio(matchModePanel, getLabel("radio.matchByPath"), initialMatchMode !== MATCH_BY_NAME, "tooltip.matchByPath");
        var matchByNameRadio = addRadio(matchModePanel, getLabel("radio.matchByName"), initialMatchMode === MATCH_BY_NAME, "tooltip.matchByName");

        /* 動作 / Action */
        var actionPanel = addPanel(dialog, getLabel("panel.action"));
        addRadio(actionPanel, getLabel("radio.actionSelect"),
            initialAction !== ACTION_DELETE_IMAGE && initialAction !== ACTION_DELETE_CLIP_GROUP,
            "tooltip.actionSelect");
        var deleteImageRadio = addRadio(actionPanel, getLabel("radio.actionDeleteImage"),
            initialAction === ACTION_DELETE_IMAGE, "tooltip.actionDeleteImage");
        var deleteClipGroupRadio = addRadio(actionPanel, getLabel("radio.actionDeleteClipGroup"),
            initialAction === ACTION_DELETE_CLIP_GROUP, "tooltip.actionDeleteClipGroup");

        /* ボタン / Buttons（Mac 規約：Cancel → OK） */
        var buttonRow = addButtonRow(dialog);
        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });

        var result = null;
        btnOK.onClick = function () {
            var action = ACTION_SELECT;
            if (deleteImageRadio.value) action = ACTION_DELETE_IMAGE;
            else if (deleteClipGroupRadio.value) action = ACTION_DELETE_CLIP_GROUP;
            result = {
                matchMode: matchByNameRadio.value ? MATCH_BY_NAME : MATCH_BY_PATH,
                action: action
            };
            dialog.close(1);
        };

        alignRightOnlyButtonRow(buttonRow);
        prepareDialogWindow(dialog, SCRIPT_NAME);
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
            alert(getLabel("alert.selected", { count: selectedCount }));
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
        alert(getLabel("alert.deleted", { count: deletedCount }));
    }

    main();

})();
