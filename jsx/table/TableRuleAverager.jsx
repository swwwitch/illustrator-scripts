#target illustrator
#targetengine "TableRuleAveragerEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択した外枠の長方形を基準に、内部の縦罫・横罫を等間隔に再配置します。
最大の長方形を外枠として判定し、縦罫は左右、横罫は上下方向に均等配置します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/TableRuleAverager.md

### Overview

Evenly redistributes the internal vertical and horizontal rules, using the largest selected rectangle as the outer frame.
Vertical rules are spaced horizontally and horizontal rules vertically.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/TableRuleAverager.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "TableRuleAverager";            /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.5";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-30";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/TableRuleAverager.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/TableRuleAverager.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* 罫線／長方形分類用のしきい値（pt） / Thresholds for rule/rect classification (pt) */
    var RULE_THICKNESS_THRESHOLD = 2;   /* これより細いものを罫線とみなす / Thinner than this counts as a rule */
    var RULE_LENGTH_THRESHOLD = 10;     /* 罫線とみなす最小の長さ / Minimum length of a rule */
    var RECT_SIZE_THRESHOLD = 5;        /* 外枠とみなす最小の幅・高さ / Minimum width and height of the frame */

    // =========================================
    // レイアウト / Layout
    // =========================================

    var DIALOG_MARGINS = 20;                /* ダイアログの余白 / Dialog margins */
    var PANEL_MARGINS = [15, 20, 15, 10];   /* パネル余白 [左,上,右,下] / Panel margins [L,T,R,B] */

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

    // ボタン行（再利用パーツ） / Button row (reusable)

    var BUTTON_ROW_TOP_MARGIN = 5; /* ボタン行の上の余白 / top margin of the button row */
    var BUTTON_ROW_SPACING = 10;   /* ボタンどうしの間隔 / spacing between buttons */

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

    // ボタン行（再利用パーツ）ここまで / End of the reusable button row

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "罫線の均等配置", en: "Even Rule Distribution" }
        },
        panel: {
            averaging: { ja: "均等配置の対象", en: "Equalize Targets" }
        },
        checkbox: {
            vertical: { ja: "縦罫", en: "Vertical" },
            horizontal: { ja: "横罫", en: "Horizontal" },
            preview: { ja: "プレビュー", en: "Preview" }
        },
        tooltip: {
            vertical: { ja: "縦罫の間隔を均等にします。", en: "Evens out the spacing of the vertical rules." },
            horizontal: { ja: "横罫の間隔を均等にします。", en: "Evens out the spacing of the horizontal rules." },
            preview: {
                ja: "結果を画面で確認します。キャンセルすると元に戻ります。",
                en: "Shows the result on the canvas. Cancel restores the original layout."
            }
        },
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" }
        },
        alert: {
            noDocument: { ja: "ドキュメントを開いてください。", en: "Please open a document." },
            selectRectAndRules: { ja: "長方形と罫線を選択してください。", en: "Please select a rectangle and rules." },
            noRect: { ja: "外枠となる長方形が見つかりませんでした。", en: "Could not find a bounding rectangle." }
        }
    };

    // =========================================
    // ダイアログUI / Dialog UI
    // =========================================

    /**
     * ダイアログを組み立て、プレビューとボタンの処理を結び付けて返す
     * @param {Object} boundingRect - 外枠の範囲（bounds / width / height）
     * @param {Object[]} verticalLines - 縦罫の記録（createRuleRecord() の戻り値）
     * @param {Object[]} horizontalLines - 横罫の記録（createRuleRecord() の戻り値）
     * @returns {Window} 組み立てたダイアログ
     */
    function buildDialog(boundingRect, verticalLines, horizontalLines) {

        var averagerDialog = new Window("dialog", getLabel("dialog.title") + ' ' + SCRIPT_VERSION);
        averagerDialog.orientation = "column";
        averagerDialog.alignChildren = ["left", "top"];
        averagerDialog.margins = DIALOG_MARGINS;

        var averagingPanel = averagerDialog.add("panel", undefined, getLabel("panel.averaging"));
        averagingPanel.orientation = "column";
        averagingPanel.alignChildren = ["left", "top"];
        averagingPanel.alignment = ["fill", "top"];
        averagingPanel.margins = PANEL_MARGINS;
        var verticalCheckbox = averagingPanel.add("checkbox", undefined, getLabel("checkbox.vertical"));
        verticalCheckbox.helpTip = getLabel("tooltip.vertical");
        var horizontalCheckbox = averagingPanel.add("checkbox", undefined, getLabel("checkbox.horizontal"));
        horizontalCheckbox.helpTip = getLabel("tooltip.horizontal");

        var previewGroup = averagerDialog.add("group");
        previewGroup.alignment = "center";
        var previewCheckbox = previewGroup.add("checkbox", undefined, getLabel("checkbox.preview"));
        previewCheckbox.helpTip = getLabel("tooltip.preview");

        /* 初期状態：見つかった罫線のチェックをON / Initial state: enable checkboxes for found rules */
        verticalCheckbox.value = (verticalLines.length > 0);
        horizontalCheckbox.value = (horizontalLines.length > 0);
        previewCheckbox.value = true;

        var buttonRow = addButtonRow(averagerDialog, { centered: true });
        var btnCancel = buttonRow.rowGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnOK = buttonRow.rowGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });

        /**
         * 元の位置に戻してから、チェックの状態に合わせて罫線を均等配置する
         * @returns {void}
         */
        function updatePreview() {
            resetRulePositions(verticalLines, horizontalLines);

            if (previewCheckbox.value) {
                if (verticalCheckbox.value) {
                    distributeVerticalRules(boundingRect, verticalLines);
                }
                if (horizontalCheckbox.value) {
                    distributeHorizontalRules(boundingRect, horizontalLines);
                }
            }

            app.redraw();
        }

        /**
         * チェックボックスのクリック処理を設定する。option＋クリックでもう一方を逆の状態にする
         * @param {Checkbox} clickedCheckbox - クリックされるチェックボックス
         * @param {Checkbox} otherCheckbox - 逆の状態にするチェックボックス
         * @returns {void}
         */
        function bindRuleCheckbox(clickedCheckbox, otherCheckbox) {
            clickedCheckbox.onClick = function () {
                if (ScriptUI.environment.keyboardState.altKey) {
                    otherCheckbox.value = !clickedCheckbox.value;
                }
                updatePreview();
            };
        }

        /* option+クリックで縦罫／横罫を互い違いに / Option+click toggles V and H to opposite states */
        bindRuleCheckbox(verticalCheckbox, horizontalCheckbox);
        bindRuleCheckbox(horizontalCheckbox, verticalCheckbox);
        previewCheckbox.onClick = updatePreview;

        btnCancel.onClick = function () {
            /* キャンセル時は元に戻して閉じる / Restore positions on cancel */
            resetRulePositions(verticalLines, horizontalLines);
            app.redraw();
            averagerDialog.close();
        };

        btnOK.onClick = function () {
            /* プレビューOFFでもOK時は最終設定を反映 / Apply final settings even if preview is off */
            previewCheckbox.value = true;
            updatePreview();
            averagerDialog.close();
        };

        /* 初期表示時のプレビュー反映 / Apply preview on initial display */
        updatePreview();

        return averagerDialog;
    }

    // =========================================
    // 罫線の配置 / Rule placement
    // =========================================

    /**
     * すべての罫線を元の位置・サイズに戻す
     * @param {Object[]} verticalLines - 縦罫の記録
     * @param {Object[]} horizontalLines - 横罫の記録
     * @returns {void}
     */
    function resetRulePositions(verticalLines, horizontalLines) {
        var i;
        for (i = 0; i < verticalLines.length; i++) {
            verticalLines[i].pathItem.height = verticalLines[i].originalHeight;
            verticalLines[i].pathItem.left = verticalLines[i].originalLeft;
            verticalLines[i].pathItem.top = verticalLines[i].originalTop;
        }
        for (i = 0; i < horizontalLines.length; i++) {
            horizontalLines[i].pathItem.width = horizontalLines[i].originalWidth;
            horizontalLines[i].pathItem.left = horizontalLines[i].originalLeft;
            horizontalLines[i].pathItem.top = horizontalLines[i].originalTop;
        }
    }

    /**
     * 縦罫を外枠の幅の n+1 分割位置に並べる
     * @param {Object} boundingRect - 外枠の範囲
     * @param {Object[]} verticalLines - 左から右に並んだ縦罫の記録
     * @returns {void}
     */
    function distributeVerticalRules(boundingRect, verticalLines) {
        var verticalSpacing = boundingRect.width / (verticalLines.length + 1);
        for (var i = 0; i < verticalLines.length; i++) {
            var targetX = boundingRect.bounds[0] + verticalSpacing * (i + 1);
            var horizontalDelta = targetX - verticalLines[i].originalX;
            verticalLines[i].pathItem.left = verticalLines[i].originalLeft + horizontalDelta;
        }
    }

    /**
     * 横罫を外枠の高さの n+1 分割位置に並べる
     * @param {Object} boundingRect - 外枠の範囲
     * @param {Object[]} horizontalLines - 上から下に並んだ横罫の記録
     * @returns {void}
     */
    function distributeHorizontalRules(boundingRect, horizontalLines) {
        var horizontalSpacing = boundingRect.height / (horizontalLines.length + 1);
        for (var i = 0; i < horizontalLines.length; i++) {
            var targetY = boundingRect.bounds[1] - horizontalSpacing * (i + 1);
            var verticalDelta = targetY - horizontalLines[i].originalY;
            horizontalLines[i].pathItem.top = horizontalLines[i].originalTop + verticalDelta;
        }
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 罫線の元の位置・サイズを控えた記録を作る
     * @param {PathItem} pathItem - 罫線のパス
     * @param {number[]} itemBounds - pathItem の geometricBounds
     * @returns {Object} 罫線の記録
     */
    function createRuleRecord(pathItem, itemBounds) {
        return {
            pathItem: pathItem,
            originalLeft: pathItem.left,
            originalTop: pathItem.top,
            originalWidth: pathItem.width,
            originalHeight: pathItem.height,
            originalX: itemBounds[0],
            originalY: itemBounds[1]
        };
    }

    /**
     * 選択中のパスを縦罫・横罫・外枠（最大の長方形）に振り分ける
     * @param {PageItem[]} selectedItems - 選択中のアイテム
     * @returns {{boundingRect: Object, verticalLines: Object[], horizontalLines: Object[]}} 振り分けの結果（外枠が無ければ boundingRect は null）
     */
    function classifySelection(selectedItems) {
        var boundingRect = null;
        var verticalLines = [];
        var horizontalLines = [];
        var maxRectArea = 0;

        for (var i = 0; i < selectedItems.length; i++) {
            var pageItem = selectedItems[i];
            if (pageItem.typename !== "PathItem") {
                continue;
            }

            var itemBounds = pageItem.geometricBounds; // [left, top, right, bottom]
            var itemWidth = itemBounds[2] - itemBounds[0];
            var itemHeight = itemBounds[1] - itemBounds[3];

            if (itemWidth < RULE_THICKNESS_THRESHOLD && itemHeight > RULE_LENGTH_THRESHOLD) {
                verticalLines.push(createRuleRecord(pageItem, itemBounds));
            } else if (itemHeight < RULE_THICKNESS_THRESHOLD && itemWidth > RULE_LENGTH_THRESHOLD) {
                horizontalLines.push(createRuleRecord(pageItem, itemBounds));
            } else if (itemWidth > RECT_SIZE_THRESHOLD && itemHeight > RECT_SIZE_THRESHOLD) {
                var itemArea = itemWidth * itemHeight;
                if (itemArea > maxRectArea) {
                    maxRectArea = itemArea;
                    boundingRect = { bounds: itemBounds, width: itemWidth, height: itemHeight };
                }
            }
        }

        return { boundingRect: boundingRect, verticalLines: verticalLines, horizontalLines: horizontalLines };
    }

    /**
     * 選択を振り分けてダイアログを表示する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }

        var doc = app.activeDocument;
        var selectedItems = doc.selection;

        if (selectedItems.length < 2) {
            alert(getLabel("alert.selectRectAndRules"));
            return;
        }

        /* 選択オブジェクトを分類して初期位置を保存 / Classify selection and save initial positions */
        var classified = classifySelection(selectedItems);

        if (!classified.boundingRect) {
            alert(getLabel("alert.noRect"));
            return;
        }

        /* 縦罫はX座標で昇順、横罫はY座標で降順にソート / Sort V rules left-to-right, H rules top-to-bottom */
        classified.verticalLines.sort(function (firstRule, secondRule) {
            return firstRule.originalX - secondRule.originalX;
        });
        classified.horizontalLines.sort(function (firstRule, secondRule) {
            return secondRule.originalY - firstRule.originalY;
        });

        var averagerDialog = buildDialog(classified.boundingRect, classified.verticalLines, classified.horizontalLines);
        prepareDialogWindow(averagerDialog, SCRIPT_NAME);
        averagerDialog.show();
    }

    main();

})();
