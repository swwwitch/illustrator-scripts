#target illustrator
#targetengine "AverageRulesInRectangleEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択した外枠の長方形と、その内側にある縦罫・横罫を、外枠の中で均等に再配置します。
横罫・縦罫はそれぞれON/OFFでき、プレビューの切り替えにも対応します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AverageRulesInRectangle.md

### Overview

Evenly redistributes the vertical and horizontal rules inside the selected outer rectangle.
Horizontal and vertical rules can be toggled independently, with a live preview switch.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AverageRulesInRectangle.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AverageRulesInRectangle";      /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.5";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-01";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AverageRulesInRectangle.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AverageRulesInRectangle.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /*
        true  : 縦罫の上下、横罫の左右を外枠に合わせる
                Snap rule endpoints to the outer rectangle.
        false : 罫線の長さは変えず、位置だけ平均化する
                Keep rule lengths; redistribute positions only.
    */
    var MATCH_LINES_TO_RECT = true;

    /* 罫線の判定に使う許容値（pt） / Tolerance for detecting rules (pt) */
    var TOLERANCE = 0.5;

    // =========================================
    // レイアウト / Layout
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
        dialog: {
            title: { ja: "罫線を平均化", en: "Distribute Rules" }
        },
        panel: {
            target: { ja: "対象", en: "Target" }
        },
        checkbox: {
            averageHorizontal: { ja: "横罫を平均化", en: "Distribute horizontal rules" },
            averageVertical: { ja: "縦罫を平均化", en: "Distribute vertical rules" },
            matchRuleLengths: { ja: "長さを揃える", en: "Match rule lengths" },
            preview: { ja: "プレビュー", en: "Preview" }
        },
        tooltip: {
            averageHorizontal: {
                ja: "長方形の中にある横罫の間隔を均等にします。",
                en: "Evens out the spacing of the horizontal rules inside the rectangle."
            },
            averageVertical: {
                ja: "長方形の中にある縦罫の間隔を均等にします。",
                en: "Evens out the spacing of the vertical rules inside the rectangle."
            },
            matchRuleLengths: {
                ja: "罫の長さを長方形の辺にそろえます。",
                en: "Matches the length of the rules to the sides of the rectangle."
            },
            preview: {
                ja: "結果を画面で確認します。キャンセルすると元に戻ります。",
                en: "Shows the result on the canvas. Cancel restores the original layout."
            }
        },
        button: {
            ok: { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noSelection: { ja: "外枠の長方形と罫線を選択してください。", en: "Please select the outer rectangle and the rules." },
            noPathInSelection: { ja: "選択内にパスがありません。", en: "No paths found in the selection." },
            outerRectNotFound: { ja: "外枠の長方形が見つかりませんでした。", en: "Outer rectangle not found." },
            noRulesFound: { ja: "縦罫または横罫が見つかりませんでした。", en: "No vertical or horizontal rules were found." }
        }
    };

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択から外枠と罫線を見つけ、ダイアログで平均化を確定する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }

        var doc = app.activeDocument;

        if (doc.selection.length === 0) {
            alert(getLabel("alert.noSelection"));
            return;
        }

        /* ロック中・非表示のものは除き、グループ・複合パスの中のパスも集める / Collect paths inside groups and compound paths, skipping locked and hidden ones */
        var selectedPathItems = collectSelectionPathItems(doc.selection, { skipLocked: true, skipHidden: true });

        if (selectedPathItems.length === 0) {
            alert(getLabel("alert.noPathInSelection"));
            return;
        }

        var outerRect = findOuterRectangle(selectedPathItems);

        if (!outerRect) {
            alert(getLabel("alert.outerRectNotFound"));
            return;
        }

        var rectBounds = getRectBounds(outerRect);
        var rules = classifyRules(selectedPathItems, outerRect, rectBounds);

        if (rules.vertical.length === 0 && rules.horizontal.length === 0) {
            alert(getLabel("alert.noRulesFound"));
            return;
        }

        var originalState = savePathState(selectedPathItems);

        var dialogControls = buildDialog(rules.horizontal.length, rules.vertical.length);
        var horizontalRulesCheckbox = dialogControls.horizontalRulesCheckbox;
        var verticalRulesCheckbox = dialogControls.verticalRulesCheckbox;
        var previewCheckbox = dialogControls.previewCheckbox;

        /**
         * 罫線チェックボックスのクリック処理（Option／Alt キーで横罫・縦罫をまとめて切り替える）
         * @param {Checkbox} clickedCheckbox - クリックされたチェックボックス
         * @returns {void}
         */
        function handleRuleCheckboxClick(clickedCheckbox) {
            if (ScriptUI.environment.keyboardState.altKey) {
                toggleBothRuleCheckboxes(clickedCheckbox.value);
            }
            updatePreview();
        }

        /**
         * 有効な横罫・縦罫のチェックボックスをまとめて同じ状態にする
         * @param {boolean} checkboxValue - 設定する値
         * @returns {void}
         */
        function toggleBothRuleCheckboxes(checkboxValue) {
            if (horizontalRulesCheckbox.enabled) {
                horizontalRulesCheckbox.value = checkboxValue;
            }
            if (verticalRulesCheckbox.enabled) {
                verticalRulesCheckbox.value = checkboxValue;
            }
        }

        /**
         * 元の座標に戻したうえで、チェックボックスの状態どおりに平均化する
         * @returns {void}
         */
        function applyCheckedRules() {
            restorePathState(originalState);
            if (verticalRulesCheckbox.value) {
                redistributeVerticalRules(rules.vertical, rectBounds);
            }
            if (horizontalRulesCheckbox.value) {
                redistributeHorizontalRules(rules.horizontal, rectBounds);
            }
        }

        /**
         * プレビューを更新する（プレビューが OFF のときは元の座標に戻すだけ）
         * @returns {void}
         */
        function updatePreview() {
            if (previewCheckbox.value) {
                applyCheckedRules();
            } else {
                restorePathState(originalState);
            }
            app.redraw();
        }

        horizontalRulesCheckbox.onClick = function () {
            handleRuleCheckboxClick(horizontalRulesCheckbox);
        };

        verticalRulesCheckbox.onClick = function () {
            handleRuleCheckboxClick(verticalRulesCheckbox);
        };

        previewCheckbox.onClick = function () {
            updatePreview();
        };

        dialogControls.btnOK.onClick = function () {
            applyCheckedRules();
            app.redraw();
            dialogControls.rulesDialog.close(1);
        };

        dialogControls.btnCancel.onClick = function () {
            restorePathState(originalState);
            app.redraw();
            dialogControls.rulesDialog.close(0);
        };

        updatePreview();

        prepareDialogWindow(dialogControls.rulesDialog, SCRIPT_NAME);
        dialogControls.rulesDialog.show();
    }

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
     * ダイアログを組み立てる
     * @param {number} horizontalRuleCount - 見つかった横罫の数
     * @param {number} verticalRuleCount - 見つかった縦罫の数
     * @returns {{rulesDialog: Window, horizontalRulesCheckbox: Checkbox, verticalRulesCheckbox: Checkbox, matchRuleLengthsCheckbox: Checkbox, previewCheckbox: Checkbox, btnCancel: Button, btnOK: Button}} ダイアログとコントロール
     */
    function buildDialog(horizontalRuleCount, verticalRuleCount) {
        var rulesDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        setupWindow(rulesDialog);

        var targetPanel = rulesDialog.add("panel", undefined, getLabel("panel.target"));
        setupPanel(targetPanel, 6);

        var horizontalRulesCheckbox = addOptionCheckbox(targetPanel,
            "checkbox.averageHorizontal", "tooltip.averageHorizontal", horizontalRuleCount > 0);
        var verticalRulesCheckbox = addOptionCheckbox(targetPanel,
            "checkbox.averageVertical", "tooltip.averageVertical", verticalRuleCount > 0);
        var matchRuleLengthsCheckbox = addOptionCheckbox(targetPanel,
            "checkbox.matchRuleLengths", "tooltip.matchRuleLengths", horizontalRuleCount > 0 || verticalRuleCount > 0);

        var previewGroup = rulesDialog.add("group");
        previewGroup.orientation = "row";
        previewGroup.alignment = "center";

        var previewCheckbox = previewGroup.add("checkbox", undefined, getLabel("checkbox.preview"));
        previewCheckbox.helpTip = getLabel("tooltip.preview");
        previewCheckbox.value = true;

        var buttonRow = addButtonRow(rulesDialog);
        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"));
        var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"));
        alignRightOnlyButtonRow(buttonRow);

        return {
            rulesDialog: rulesDialog,
            horizontalRulesCheckbox: horizontalRulesCheckbox,
            verticalRulesCheckbox: verticalRulesCheckbox,
            matchRuleLengthsCheckbox: matchRuleLengthsCheckbox,
            previewCheckbox: previewCheckbox,
            btnCancel: btnCancel,
            btnOK: btnOK
        };
    }

    /**
     * OFF で始まるチェックボックスを tooltip 付きで追加する
     * @param {Panel} parentPanel - 追加先のパネル
     * @param {string} labelPath - 表示名の LABELS パス
     * @param {string} tooltipPath - tooltip の LABELS パス
     * @param {boolean} isEnabled - 操作できるかどうか
     * @returns {Checkbox} 追加したチェックボックス
     */
    function addOptionCheckbox(parentPanel, labelPath, tooltipPath, isEnabled) {
        var optionCheckbox = parentPanel.add("checkbox", undefined, getLabel(labelPath));
        optionCheckbox.helpTip = getLabel(tooltipPath);
        optionCheckbox.value = false;
        optionCheckbox.enabled = isEnabled;
        return optionCheckbox;
    }

    // 選択の収集と境界（再利用パーツ） / Selection items and bounds (reusable)

    /* 座標を同じと見なす許容値（pt） / Tolerance for treating coordinates as equal, in points */
    var SELECTION_ITEMS_TOLERANCE = 0.001;

    /**
     * 選択やコレクションを、オブジェクトの配列にそろえる
     * TextRange・PathItem は length を持つので、typename で1個か集まりかを見分ける
     * @param {*} source - doc.selection、配列、DOM のコレクション、または単独のオブジェクト
     * @returns {Array} オブジェクトの配列（空なら []）
     */
    function normalizeSelectionItems(source) {
        var items = [];
        if (!source) return items;
        var typeName = "";
        try { typeName = source.typename || ""; } catch (e) { /* 読めない種類 / unreadable kind */ }
        /* 単数形の typename は1個（PageItems などのコレクションは s で終わる）
           A singular typename is one object (collections such as PageItems end in s) */
        if (typeName && !/s$/.test(typeName)) return [source];
        if (typeof source.length !== "number") return items;
        for (var i = 0; i < source.length; i++) items.push(source[i]);
        return items;
    }

    /**
     * 文字カーソルの選択（TextRange）を、それを含むテキストフレームに読み替える
     * @param {TextRange} textRange - 文字の範囲
     * @returns {TextFrame|null} テキストフレーム（たどれなければ null）
     */
    function resolveTextRangeFrame(textRange) {
        var current = textRange;
        /* parent をたどる（深さは念のため制限） / Walk up the parents, with a safety limit */
        for (var depth = 0; depth < 10 && current; depth++) {
            try {
                if (current.typename === "TextFrame") return current;
                current = current.parent;
            } catch (e) {
                break;
            }
        }
        /* ストーリーの先頭フレームで代用する / Fall back to the first frame of the story */
        try {
            var storyFrames = textRange.story.textFrames;
            if (storyFrames.length > 0) return storyFrames[0];
        } catch (e2) { /* ストーリーを持たない / no story */ }
        return null;
    }

    /**
     * 選択から条件に合うオブジェクトを集める（グループ・レイヤーを再帰でたどり、重複は除く）
     * 条件に合ったオブジェクトの中へは進まない
     * @param {*} source - doc.selection、配列、コレクション、または単独のオブジェクト
     * @param {Object} [options] - 収集の設定
     * @param {function(PageItem): boolean} [options.accept] - 集める条件（既定はグループ・レイヤー以外すべて）
     * @param {boolean} [options.enterGroups] - グループの中をたどる（既定 true）
     * @param {boolean} [options.enterClipGroups] - クリップグループの中をたどる（既定は enterGroups と同じ）
     * @param {boolean} [options.enterCompoundPaths] - 複合パスの中のパスをたどる（既定 false）
     * @param {boolean} [options.textRangeToFrame] - 文字の選択をテキストフレームに読み替える（既定 true）
     * @param {boolean} [options.skipLocked] - ロックされたものを中ごと外す（既定 false）
     * @param {boolean} [options.skipHidden] - 非表示のものを中ごと外す（既定 false）
     * @param {boolean} [options.skipClipMasks] - クリッピングマスクを外す（既定 false）
     * @param {boolean} [options.skipGuides] - ガイドを外す（既定 false）
     * @param {boolean} [options.unique] - 同じ参照を1回だけにする（既定 true。数千件で遅ければ false）
     * @returns {Array} 集めたオブジェクト（前面→背面の順）
     */
    function collectSelectionItems(source, options) {
        var opts = options || {};
        var enterGroups = (opts.enterGroups !== false);
        var enterClipGroups = (opts.enterClipGroups === undefined) ? enterGroups : (opts.enterClipGroups === true);
        var accept = opts.accept || function (item) {
            return item.typename !== "GroupItem" && item.typename !== "Layer";
        };
        var collected = [];

        /**
         * 集めた配列に加える（unique のときは同じ参照を足さない）
         * @param {PageItem} item - 加えるオブジェクト
         * @returns {void}
         */
        function pushItem(item) {
            if (opts.unique !== false) {
                for (var k = 0; k < collected.length; k++) {
                    if (collected[k] === item) return;
                }
            }
            collected.push(item);
        }

        /**
         * 設定に従って外すオブジェクトか判定する
         * @param {PageItem} item - 判定するオブジェクト
         * @returns {boolean} 外すなら true
         */
        function isSkipped(item) {
            try {
                if (item.typename === "Layer") {
                    if (opts.skipLocked && item.locked) return true;
                    if (opts.skipHidden && !item.visible) return true;
                    return false;
                }
                if (opts.skipLocked && item.locked) return true;
                if (opts.skipHidden && item.hidden) return true;
                if (opts.skipGuides && item.guides === true) return true;
                if (opts.skipClipMasks && isClipMaskItem(item)) return true;
            } catch (e) {
                /* 読めないプロパティは「外さない」に倒す / Unreadable properties do not exclude */
            }
            return false;
        }

        /**
         * 1件をたどって集める
         * @param {PageItem} item - 対象のオブジェクト
         * @returns {void}
         */
        function visit(item) {
            if (!item) return;
            var typeName = "";
            try { typeName = item.typename; } catch (e) { return; }

            if (typeName === "TextRange" || typeName === "InsertionPoint") {
                if (opts.textRangeToFrame === false) {
                    if (accept(item)) pushItem(item);
                    return;
                }
                visit(resolveTextRangeFrame(item));
                return;
            }
            if (isSkipped(item)) return;
            if (accept(item)) {
                pushItem(item);
                return;
            }

            var children = null;
            if (typeName === "GroupItem") {
                var isClipped = false;
                try { isClipped = (item.clipped === true); } catch (e2) { }
                if (isClipped ? enterClipGroups : enterGroups) children = item.pageItems;
            } else if (typeName === "CompoundPathItem") {
                if (opts.enterCompoundPaths) children = item.pathItems;
            } else if (typeName === "Layer") {
                /* 重なり順はサブレイヤーとページアイテムで別々なので、ページアイテム→サブレイヤーの順にする
                   Page items and sublayers stack separately; visit page items first, then sublayers */
                walk(item.pageItems);
                walk(item.layers);
                return;
            }
            if (children) walk(children);
        }

        /**
         * 集まりの各要素をたどる
         * @param {*} list - 配列またはコレクション
         * @returns {void}
         */
        function walk(list) {
            var listItems = normalizeSelectionItems(list);
            for (var i = 0; i < listItems.length; i++) visit(listItems[i]);
        }

        walk(source);
        return collected;
    }

    /**
     * テキストフレームの種類を "point" / "area" / "path" で返す
     * @param {TextFrame} textFrame - テキストフレーム
     * @returns {string} 種類のキー（判定できなければ ""）
     */
    function getTextFrameKindKey(textFrame) {
        try {
            if (textFrame.kind === TextType.POINTTEXT) return "point";
            if (textFrame.kind === TextType.AREATEXT) return "area";
            if (textFrame.kind === TextType.PATHTEXT) return "path";
        } catch (e) { /* kind を読めない / kind is unreadable */ }
        return "";
    }

    /**
     * 選択からテキストフレームを集める（グループの中・文字カーソルの選択を含む）
     * @param {*} source - doc.selection など
     * @param {Object} [options] - collectSelectionItems と同じ設定に加えて次を受ける
     * @param {string[]} [options.kinds] - 集める種類（"point" / "area" / "path"。既定はすべて）
     * @returns {TextFrame[]} テキストフレーム（前面→背面の順）
     */
    function collectSelectionTextFrames(source, options) {
        var opts = {};
        var sourceOptions = options || {};
        for (var key in sourceOptions) {
            if (sourceOptions.hasOwnProperty(key)) opts[key] = sourceOptions[key];
        }
        var kindFilter = null;
        if (opts.kinds && opts.kinds.length) {
            kindFilter = {};
            for (var i = 0; i < opts.kinds.length; i++) kindFilter[opts.kinds[i]] = true;
        }
        opts.accept = function (item) {
            if (item.typename !== "TextFrame") return false;
            return !kindFilter || kindFilter[getTextFrameKindKey(item)] === true;
        };
        /* 種類で外したテキストは中をたどらない（accept が false でも子は無い） / Text frames have no children to walk */
        return collectSelectionItems(source, opts);
    }

    /**
     * 選択からパスを集める（グループの中を含む）
     * @param {*} source - doc.selection など
     * @param {Object} [options] - collectSelectionItems と同じ設定に加えて次を受ける
     * @param {string} [options.compoundPaths] - 複合パスの扱い。"children"（中のパス、既定）/ "whole"（複合パスごと）/ "skip"（外す）
     * @returns {Array} PathItem（"whole" のときは CompoundPathItem も）の配列
     */
    function collectSelectionPathItems(source, options) {
        var opts = {};
        var sourceOptions = options || {};
        for (var key in sourceOptions) {
            if (sourceOptions.hasOwnProperty(key)) opts[key] = sourceOptions[key];
        }
        var compoundMode = opts.compoundPaths || "children";
        opts.enterCompoundPaths = (compoundMode === "children");
        opts.accept = function (item) {
            if (item.typename === "PathItem") return true;
            return compoundMode === "whole" && item.typename === "CompoundPathItem";
        };
        return collectSelectionItems(source, opts);
    }

    /**
     * クリッピングマスク（クリップグループの型）か判定する
     * パスは clipping、複合パスは中の先頭パスの clipping、テキストは clipping が無いので「クリップグループの先頭」で見る
     * @param {PageItem} item - 判定するオブジェクト
     * @returns {boolean} マスクなら true
     */
    function isClipMaskItem(item) {
        try {
            if (item.typename === "PathItem") return item.clipping === true;
            if (item.typename === "CompoundPathItem") {
                return item.pathItems.length > 0 && item.pathItems[0].clipping === true;
            }
            if (item.typename === "TextFrame") {
                var parentGroup = item.parent;
                return parentGroup.typename === "GroupItem" && parentGroup.clipped === true &&
                    parentGroup.pageItems.length > 0 && parentGroup.pageItems[0] === item;
            }
        } catch (e) { /* 読めない種類はマスクではない / unreadable kinds are not masks */ }
        return false;
    }

    /**
     * クリップグループの型（マスク）を返す
     * フラグで探し、見つからなければ先頭（pageItems[0]）を返す（型は常に最前面。テキストの型はフラグを持たない）
     * @param {GroupItem} groupItem - 対象のグループ
     * @returns {PageItem|null} マスク（クリップグループでなければ null）
     */
    function getClipMaskItem(groupItem) {
        try {
            if (!groupItem || groupItem.typename !== "GroupItem" || groupItem.clipped !== true) return null;
            var groupChildren = groupItem.pageItems;
            if (groupChildren.length === 0) return null;
            for (var i = 0; i < groupChildren.length; i++) {
                var childType = groupChildren[i].typename;
                if ((childType === "PathItem" || childType === "CompoundPathItem") && isClipMaskItem(groupChildren[i])) {
                    return groupChildren[i];
                }
            }
            return groupChildren[0];
        } catch (e) {
            return null;
        }
    }

    /**
     * グループの中（入れ子を含む）にクリップグループがあるか判定する
     * @param {GroupItem} groupItem - 対象のグループ
     * @returns {boolean} あれば true
     */
    function hasClippedDescendant(groupItem) {
        try {
            var groupChildren = groupItem.pageItems;
            for (var i = 0; i < groupChildren.length; i++) {
                if (groupChildren[i].typename !== "GroupItem") continue;
                if (groupChildren[i].clipped === true || hasClippedDescendant(groupChildren[i])) return true;
            }
        } catch (e) { /* 中を読めない / cannot read the children */ }
        return false;
    }

    /**
     * 環境設定の［プレビュー境界を使用］を読む
     * @returns {boolean} オンなら true（読めなければ false）
     */
    function readUsePreviewBoundsPreference() {
        try {
            return app.preferences.getBooleanPreference("includeStrokeInBounds");
        } catch (e) {
            return false;
        }
    }

    /**
     * 見た目どおりの境界を返す。クリップグループはマスクの境界、
     * 中にクリップグループを含むグループは子の境界を合わせたもの（隠れた部分を含めない）
     * @param {PageItem} item - 対象のオブジェクト
     * @param {boolean} [usePreviewBounds] - true で visibleBounds、false で geometricBounds（省略時は環境設定に従う）
     * @returns {number[]|null} [左, 上, 右, 下] の新しい配列（測れなければ null）
     */
    function getClipAwareBounds(item, usePreviewBounds) {
        var usePreview = (usePreviewBounds === undefined || usePreviewBounds === null) ?
            readUsePreviewBoundsPreference() : (usePreviewBounds === true);
        try {
            var measuredItem = item;
            if (item.typename === "GroupItem") {
                var maskItem = getClipMaskItem(item);
                if (maskItem) {
                    measuredItem = maskItem;
                } else if (hasClippedDescendant(item)) {
                    /* グループ自体の効果（影など）の広がりは含まれなくなる
                       This leaves out the reach of effects applied to the group itself (drop shadows etc.) */
                    var childBounds = getClipAwareUnionBounds(filterMeasurableChildren(item.pageItems), usePreview);
                    if (childBounds) return childBounds;
                }
            }
            var bounds = usePreview ? measuredItem.visibleBounds : measuredItem.geometricBounds;
            return [bounds[0], bounds[1], bounds[2], bounds[3]];
        } catch (e) {
            return null;
        }
    }

    /**
     * 境界の計算に入れる子だけを残す（非表示とガイドを外す）
     * @param {*} childList - 子のコレクション
     * @returns {Array} 残した子
     */
    function filterMeasurableChildren(childList) {
        var childItems = normalizeSelectionItems(childList);
        var measurable = [];
        for (var i = 0; i < childItems.length; i++) {
            try {
                if (childItems[i].hidden === true || childItems[i].guides === true) continue;
            } catch (e) { /* 読めなければ残す / keep when unreadable */ }
            measurable.push(childItems[i]);
        }
        return measurable;
    }

    /**
     * 複数のオブジェクトを囲む外接範囲を返す（クリップグループはマスクで測る）
     * @param {*} items - オブジェクトの配列・コレクション・選択
     * @param {boolean} [usePreviewBounds] - true で visibleBounds、false で geometricBounds（省略時は環境設定に従う）
     * @returns {number[]|null} [左, 上, 右, 下]（測れるものが無ければ null）
     */
    function getClipAwareUnionBounds(items, usePreviewBounds) {
        var usePreview = (usePreviewBounds === undefined || usePreviewBounds === null) ?
            readUsePreviewBoundsPreference() : (usePreviewBounds === true);
        var itemList = normalizeSelectionItems(items);
        var unionBounds = null;
        for (var i = 0; i < itemList.length; i++) {
            var itemBounds = getClipAwareBounds(itemList[i], usePreview);
            if (!itemBounds) continue;
            if (!unionBounds) {
                unionBounds = itemBounds;
                continue;
            }
            if (itemBounds[0] < unionBounds[0]) unionBounds[0] = itemBounds[0];
            if (itemBounds[1] > unionBounds[1]) unionBounds[1] = itemBounds[1];
            if (itemBounds[2] > unionBounds[2]) unionBounds[2] = itemBounds[2];
            if (itemBounds[3] < unionBounds[3]) unionBounds[3] = itemBounds[3];
        }
        return unionBounds;
    }

    /**
     * 2つの座標を許容値つきで比べる
     * @param {number} valueA - 座標A（pt）
     * @param {number} valueB - 座標B（pt）
     * @param {number} [tolerance] - 許容値（pt、既定は SELECTION_ITEMS_TOLERANCE）
     * @returns {boolean} 差が許容値以下なら true
     */
    function isNearlySameCoordinate(valueA, valueB, tolerance) {
        var limit = (typeof tolerance === "number") ? tolerance : SELECTION_ITEMS_TOLERANCE;
        return Math.abs(valueA - valueB) <= limit;
    }

    /**
     * 2つの境界を許容値つきで比べる
     * @param {number[]} boundsA - [左, 上, 右, 下]
     * @param {number[]} boundsB - [左, 上, 右, 下]
     * @param {number} [tolerance] - 許容値（pt、既定は SELECTION_ITEMS_TOLERANCE）
     * @returns {boolean} 4辺とも許容値以内なら true
     */
    function areBoundsNearlyEqual(boundsA, boundsB, tolerance) {
        if (!boundsA || !boundsB) return false;
        for (var i = 0; i < 4; i++) {
            if (!isNearlySameCoordinate(boundsA[i], boundsB[i], tolerance)) return false;
        }
        return true;
    }

    // 選択の収集と境界（再利用パーツ）ここまで / End of the reusable selection items and bounds

    // =========================================
    // 外枠と罫線の判定 / Detecting the frame and rules
    // =========================================

    /**
     * もっとも面積の大きい閉じたパスを外枠とみなす
     * @param {PathItem[]} candidatePathItems - 候補のパス
     * @returns {PathItem|null} 外枠のパス。見つからなければ null
     */
    function findOuterRectangle(candidatePathItems) {
        var outerRectangleCandidate = null;
        var maxArea = 0;

        for (var i = 0; i < candidatePathItems.length; i++) {
            var pathItem = candidatePathItems[i];

            if (!pathItem.closed) {
                continue;
            }

            var geometricBounds = pathItem.geometricBounds;
            var pathWidth = geometricBounds[2] - geometricBounds[0];
            var pathHeight = geometricBounds[1] - geometricBounds[3];

            if (pathWidth <= 0 || pathHeight <= 0) {
                continue;
            }

            var pathArea = pathWidth * pathHeight;

            if (pathArea > maxArea) {
                maxArea = pathArea;
                outerRectangleCandidate = pathItem;
            }
        }

        return outerRectangleCandidate;
    }

    /**
     * 外枠の座標と寸法をまとめて返す
     * @param {PathItem} outerRect - 外枠のパス
     * @returns {{left: number, top: number, right: number, bottom: number, width: number, height: number}} 外枠の範囲
     */
    function getRectBounds(outerRect) {
        var outerBounds = outerRect.geometricBounds;
        return {
            left: outerBounds[0],
            top: outerBounds[1],
            right: outerBounds[2],
            bottom: outerBounds[3],
            width: outerBounds[2] - outerBounds[0],
            height: outerBounds[1] - outerBounds[3]
        };
    }

    /**
     * 外枠の中にある縦罫・横罫を振り分け、縦罫は左から右、横罫は上から下に並べる
     * @param {PathItem[]} candidatePathItems - 候補のパス
     * @param {PathItem} outerRect - 外枠のパス（対象から除く）
     * @param {Object} rectBounds - getRectBounds() の戻り値
     * @returns {{vertical: PathItem[], horizontal: PathItem[]}} 縦罫と横罫
     */
    function classifyRules(candidatePathItems, outerRect, rectBounds) {
        var verticalRules = [];
        var horizontalRules = [];

        for (var i = 0; i < candidatePathItems.length; i++) {
            var pathItem = candidatePathItems[i];

            if (pathItem === outerRect) {
                continue;
            }

            if (isRuleInRect(pathItem, rectBounds, true)) {
                verticalRules.push(pathItem);
            } else if (isRuleInRect(pathItem, rectBounds, false)) {
                horizontalRules.push(pathItem);
            }
        }

        /* 縦罫は左から右へ、横罫は上から下へソート / Sort vertical L→R, horizontal T→B */
        verticalRules.sort(function (firstRule, secondRule) {
            return getCenterX(firstRule) - getCenterX(secondRule);
        });

        horizontalRules.sort(function (firstRule, secondRule) {
            return getCenterY(secondRule) - getCenterY(firstRule);
        });

        return { vertical: verticalRules, horizontal: horizontalRules };
    }

    /**
     * 外枠の中にある縦罫（または横罫）かどうかを判定する
     * 縦罫は「長さ＝高さ、並ぶ方向＝X」、横罫は「長さ＝幅、並ぶ方向＝Y」として同じ条件で見る
     * @param {PathItem} pathItem - 判定するパス
     * @param {Object} rectBounds - getRectBounds() の戻り値
     * @param {boolean} isVertical - true なら縦罫、false なら横罫として判定する
     * @returns {boolean} 罫線とみなせるとき true
     */
    function isRuleInRect(pathItem, rectBounds, isVertical) {
        if (pathItem.closed) {
            return false;
        }

        var geometricBounds = pathItem.geometricBounds;
        var pathWidth = Math.abs(geometricBounds[2] - geometricBounds[0]);
        var pathHeight = Math.abs(geometricBounds[1] - geometricBounds[3]);
        var centerX = (geometricBounds[0] + geometricBounds[2]) / 2;
        var centerY = (geometricBounds[1] + geometricBounds[3]) / 2;

        if (isVertical) {
            return hasRuleShape(pathHeight, pathWidth, rectBounds.height) &&
                isInsideRange(centerX, rectBounds.left, rectBounds.right) &&
                isNearRange(centerY, rectBounds.bottom, rectBounds.top);
        }
        return hasRuleShape(pathWidth, pathHeight, rectBounds.width) &&
            isInsideRange(centerY, rectBounds.bottom, rectBounds.top) &&
            isNearRange(centerX, rectBounds.left, rectBounds.right);
    }

    /**
     * 細長い形で、外枠の辺の 20% 以上の長さがあるかどうか
     * @param {number} ruleLength - 罫線の方向の長さ
     * @param {number} ruleThickness - 罫線と直交する方向の幅
     * @param {number} outerLength - 同じ方向の外枠の長さ
     * @returns {boolean} 罫線の形とみなせるとき true
     */
    function hasRuleShape(ruleLength, ruleThickness, outerLength) {
        if (ruleLength <= 0) {
            return false;
        }
        if (!(ruleThickness <= TOLERANCE || ruleLength > ruleThickness * 5)) {
            return false;
        }
        return !(ruleLength < outerLength * 0.2);
    }

    /**
     * 値が範囲の内側に、許容値より離れて収まっているかどうか（両端ちょうどは外）
     * @param {number} value - 調べる値
     * @param {number} rangeMin - 範囲の最小値
     * @param {number} rangeMax - 範囲の最大値
     * @returns {boolean} 内側なら true
     */
    function isInsideRange(value, rangeMin, rangeMax) {
        return value > rangeMin + TOLERANCE && value < rangeMax - TOLERANCE;
    }

    /**
     * 値が範囲に、許容値の はみ出しまで含めて収まっているかどうか
     * @param {number} value - 調べる値
     * @param {number} rangeMin - 範囲の最小値
     * @param {number} rangeMax - 範囲の最大値
     * @returns {boolean} 収まっていれば true
     */
    function isNearRange(value, rangeMin, rangeMax) {
        return value >= rangeMin - TOLERANCE && value <= rangeMax + TOLERANCE;
    }

    /**
     * パスの外接矩形の中心 X を返す
     * @param {PathItem} pathItem - 対象のパス
     * @returns {number} 中心の X 座標
     */
    function getCenterX(pathItem) {
        var geometricBounds = pathItem.geometricBounds;
        return (geometricBounds[0] + geometricBounds[2]) / 2;
    }

    /**
     * パスの外接矩形の中心 Y を返す
     * @param {PathItem} pathItem - 対象のパス
     * @returns {number} 中心の Y 座標
     */
    function getCenterY(pathItem) {
        var geometricBounds = pathItem.geometricBounds;
        return (geometricBounds[1] + geometricBounds[3]) / 2;
    }

    // =========================================
    // 罫線の再配置 / Redistributing rules
    // =========================================

    /**
     * 縦罫を外枠の幅で等間隔に並べ直す
     * @param {PathItem[]} verticalRules - 左から右に並んだ縦罫
     * @param {Object} rectBounds - getRectBounds() の戻り値
     * @returns {void}
     */
    function redistributeVerticalRules(verticalRules, rectBounds) {
        var ruleCount = verticalRules.length;

        for (var i = 0; i < ruleCount; i++) {
            var verticalRule = verticalRules[i];
            var targetX = rectBounds.left + rectBounds.width * (i + 1) / (ruleCount + 1);

            if (MATCH_LINES_TO_RECT && verticalRule.pathPoints.length === 2) {
                setTwoPointVerticalLine(verticalRule, targetX, rectBounds.top, rectBounds.bottom);
            } else {
                translatePath(verticalRule, targetX - getCenterX(verticalRule), 0);
            }
        }
    }

    /**
     * 横罫を外枠の高さで等間隔に並べ直す
     * @param {PathItem[]} horizontalRules - 上から下に並んだ横罫
     * @param {Object} rectBounds - getRectBounds() の戻り値
     * @returns {void}
     */
    function redistributeHorizontalRules(horizontalRules, rectBounds) {
        var ruleCount = horizontalRules.length;

        for (var i = 0; i < ruleCount; i++) {
            var horizontalRule = horizontalRules[i];
            var targetY = rectBounds.top - rectBounds.height * (i + 1) / (ruleCount + 1);

            if (MATCH_LINES_TO_RECT && horizontalRule.pathPoints.length === 2) {
                setTwoPointHorizontalLine(horizontalRule, targetY, rectBounds.left, rectBounds.right);
            } else {
                translatePath(horizontalRule, 0, targetY - getCenterY(horizontalRule));
            }
        }
    }

    /**
     * パスのすべてのアンカーポイントとハンドルを平行移動する
     * @param {PathItem} pathItem - 対象のパス
     * @param {number} deltaX - X 方向の移動量
     * @param {number} deltaY - Y 方向の移動量
     * @returns {void}
     */
    function translatePath(pathItem, deltaX, deltaY) {
        for (var i = 0; i < pathItem.pathPoints.length; i++) {
            movePathPoint(pathItem.pathPoints[i], deltaX, deltaY);
        }
    }

    /**
     * 2点の縦罫を、指定の X 位置で上端から下端までの線にする
     * @param {PathItem} pathItem - 2点のパス
     * @param {number} x - 配置する X 座標
     * @param {number} top - 上端の Y 座標
     * @param {number} bottom - 下端の Y 座標
     * @returns {void}
     */
    function setTwoPointVerticalLine(pathItem, x, top, bottom) {
        var firstPoint = pathItem.pathPoints[0];
        var secondPoint = pathItem.pathPoints[1];
        var isFirstOnTop = firstPoint.anchor[1] >= secondPoint.anchor[1];

        movePointTo(isFirstOnTop ? firstPoint : secondPoint, x, top);
        movePointTo(isFirstOnTop ? secondPoint : firstPoint, x, bottom);
    }

    /**
     * 2点の横罫を、指定の Y 位置で左端から右端までの線にする
     * @param {PathItem} pathItem - 2点のパス
     * @param {number} y - 配置する Y 座標
     * @param {number} left - 左端の X 座標
     * @param {number} right - 右端の X 座標
     * @returns {void}
     */
    function setTwoPointHorizontalLine(pathItem, y, left, right) {
        var firstPoint = pathItem.pathPoints[0];
        var secondPoint = pathItem.pathPoints[1];
        var isFirstOnLeft = firstPoint.anchor[0] <= secondPoint.anchor[0];

        movePointTo(isFirstOnLeft ? firstPoint : secondPoint, left, y);
        movePointTo(isFirstOnLeft ? secondPoint : firstPoint, right, y);
    }

    /**
     * アンカーポイントを指定の座標へ移動する（ハンドルも同じ量だけ動かす）
     * @param {PathPoint} pathPoint - 対象のアンカーポイント
     * @param {number} targetX - 移動先の X 座標
     * @param {number} targetY - 移動先の Y 座標
     * @returns {void}
     */
    function movePointTo(pathPoint, targetX, targetY) {
        movePathPoint(pathPoint, targetX - pathPoint.anchor[0], targetY - pathPoint.anchor[1]);
    }

    /**
     * アンカーポイントと両ハンドルを同じ量だけ動かす
     * @param {PathPoint} pathPoint - 対象のアンカーポイント
     * @param {number} deltaX - X 方向の移動量
     * @param {number} deltaY - Y 方向の移動量
     * @returns {void}
     */
    function movePathPoint(pathPoint, deltaX, deltaY) {
        pathPoint.anchor = [pathPoint.anchor[0] + deltaX, pathPoint.anchor[1] + deltaY];
        pathPoint.leftDirection = [pathPoint.leftDirection[0] + deltaX, pathPoint.leftDirection[1] + deltaY];
        pathPoint.rightDirection = [pathPoint.rightDirection[0] + deltaX, pathPoint.rightDirection[1] + deltaY];
    }

    // =========================================
    // 座標の保存と復元 / Saving and restoring coordinates
    // =========================================

    /**
     * 編集前のパスの座標を保存する
     * @param {PathItem[]} targetPathItems - 保存するパス
     * @returns {Object[]} パスごとのアンカーポイントの状態
     */
    function savePathState(targetPathItems) {
        var pathStates = [];

        for (var pathIndex = 0; pathIndex < targetPathItems.length; pathIndex++) {
            var pathItem = targetPathItems[pathIndex];
            var pointStates = [];

            for (var pointIndex = 0; pointIndex < pathItem.pathPoints.length; pointIndex++) {
                var pathPoint = pathItem.pathPoints[pointIndex];

                pointStates.push({
                    anchor: [pathPoint.anchor[0], pathPoint.anchor[1]],
                    leftDirection: [pathPoint.leftDirection[0], pathPoint.leftDirection[1]],
                    rightDirection: [pathPoint.rightDirection[0], pathPoint.rightDirection[1]],
                    pointType: pathPoint.pointType
                });
            }

            pathStates.push({ pathItem: pathItem, pointStates: pointStates });
        }

        return pathStates;
    }

    /**
     * 保存しておいた座標へ戻す
     * @param {Object[]} pathStates - savePathState() の戻り値
     * @returns {void}
     */
    function restorePathState(pathStates) {
        for (var pathIndex = 0; pathIndex < pathStates.length; pathIndex++) {
            var pathItem = pathStates[pathIndex].pathItem;
            var pointStates = pathStates[pathIndex].pointStates;

            for (var pointIndex = 0; pointIndex < pointStates.length; pointIndex++) {
                var pathPoint = pathItem.pathPoints[pointIndex];
                var pointState = pointStates[pointIndex];

                pathPoint.anchor = [pointState.anchor[0], pointState.anchor[1]];
                pathPoint.leftDirection = [pointState.leftDirection[0], pointState.leftDirection[1]];
                pathPoint.rightDirection = [pointState.rightDirection[0], pointState.rightDirection[1]];
                pathPoint.pointType = pointState.pointType;
            }
        }
    }

    main();

})();
