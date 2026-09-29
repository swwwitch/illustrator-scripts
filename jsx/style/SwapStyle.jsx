#target illustrator
#targetengine "SwapStyleEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択した2つのオブジェクトの間で、見た目または文字列を交換します。
スタイル交換では、グラフィックスタイル・基本的な塗りや線・文字属性を組み合わせて指定できます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SwapStyle.md

### Overview

Swaps the appearance, or the text content, between two selected objects.
Style swapping can combine the graphic style, the basic fill and stroke, and the character attributes.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SwapStyle.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SwapStyle";                    /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.3";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-05-21";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-29";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SwapStyle.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SwapStyle.md"; /* README (English) */

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

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        // UI labels
        dialogTitle: {
            ja: "スタイル・文字列の交換",
            en: "Swap Style / Content"
        },
        modePanel: {
            ja: "交換方式",
            en: "Swap Mode"
        },
        modeStyle: {
            ja: "スタイル",
            en: "Style"
        },
        modeContent: {
            ja: "文字列",
            en: "Content"
        },
        modeCoordinate: {
            ja: "座標",
            en: "Position"
        },
        formatPanel: {
            ja: "文字書式",
            en: "Text Formatting"
        },
        fontAndStyle: {
            ja: "フォントとスタイル",
            en: "Font & Style"
        },
        fontSize: {
            ja: "フォントサイズ",
            en: "Font Size"
        },
        basicFillStrokePanel: {
            ja: "基本的な塗りや線",
            en: "Basic Fill & Stroke"
        },
        swapFill: {
            ja: "塗り",
            en: "Fill"
        },
        swapStrokeColor: {
            ja: "線のカラー",
            en: "Stroke Color"
        },
        swapStrokeWidth: {
            ja: "線幅",
            en: "Stroke Width"
        },
        graphicStylePanel: {
            ja: "グラフィックスタイル",
            en: "Graphic Style"
        },
        coordinatePanel: {
            ja: "座標",
            en: "Position"
        },
        tipAnchor: {
            ja: "選択した基準点で2つのオブジェクトの座標を交換します。",
            en: "Swap positions of the two objects by the selected reference point."
        },
        axisLabel: {
            ja: "交換する軸",
            en: "Axis"
        },
        axisBoth: {
            ja: "両方",
            en: "Both"
        },
        axisX: {
            ja: "横（X）",
            en: "X"
        },
        axisY: {
            ja: "縦（Y）",
            en: "Y"
        },
        swapZOrder: {
            ja: "重ね順",
            en: "Stacking order"
        },
        swapGraphicStyles: {
            ja: "交換する",
            en: "Swap"
        },
        deleteStyles: {
            ja: "一時スタイルを削除",
            en: "Delete Temporary Styles"
        },
        cancel: {
            ja: "キャンセル",
            en: "Cancel"
        },
        preview: {
            ja: "プレビュー",
            en: "Preview"
        },

        // Tooltips
        tipContentSwap: {
            ja: "文字列交換はテキストオブジェクト同士でのみ実行できます。",
            en: "Content swap is available only between text objects."
        },
        tipDeleteStyles: {
            ja: "削除しても適用済みのアピアランスは各オブジェクトに残ります。",
            en: "Deleting styles does not remove the applied appearance from objects."
        },
        tipGraphicStyle: {
            ja: "現在のアピアランスを一時スタイルとして登録し、相互に適用します。",
            en: "Registers current appearances as temporary styles and applies them crosswise."
        },

        // Alerts
        alertNoDocument: {
            ja: "ドキュメントが開かれていません。",
            en: "No document is open."
        },
        alertSelectTwo: {
            ja: "オブジェクトをちょうど2つ選択してください。",
            en: "Please select exactly two objects."
        },
        alertContentTextOnly: {
            ja: "文字列交換はテキストオブジェクト同士でのみ実行できます。",
            en: "Content swap requires both objects to be text frames."
        },
        alertNoFreePair: {
            ja: "使用できるスタイル名（_swapStyleA〜_swapStyleZ）の空きがありません。",
            en: "No free style name pair available (_swapStyleA–_swapStyleZ)."
        },
        alertSwapFailed: {
            ja: "スタイルの登録に失敗しました。処理を中止します。",
            en: "Graphic style registration failed. Aborting."
        }
    };

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    var PANEL_MARGINS = [15, 20, 15, 10];
    var PANEL_SPACING = 8;

    /* パネルの共通設定 / Common panel setup */
    function setupPanel(panel, spacing) {
        panel.orientation = "column";
        panel.alignChildren = ['fill', 'top'];
        panel.alignment = "fill";
        panel.margins = PANEL_MARGINS;
        panel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /* 交換方式パネルを作成 / Build the swap mode panel */
    function buildModePanel(dialog, bothAreText) {
        var modePanel = dialog.add('panel', undefined, getLabel('modePanel'));
        setupPanel(modePanel);
        modePanel.orientation = "row";
        modePanel.alignChildren = ['center', 'center'];

        var contentSwapRadio = modePanel.add('radiobutton', undefined, getLabel('modeContent'));
        var styleSwapRadio = modePanel.add('radiobutton', undefined, getLabel('modeStyle'));
        var coordinateSwapRadio = modePanel.add('radiobutton', undefined, getLabel('modeCoordinate'));
        contentSwapRadio.enabled = bothAreText;
        contentSwapRadio.helpTip = getLabel('tipContentSwap');
        styleSwapRadio.value = true;

        return {
            panel: modePanel,
            styleSwapRadio: styleSwapRadio,
            contentSwapRadio: contentSwapRadio,
            coordinateSwapRadio: coordinateSwapRadio
        };
    }

    /* 文字書式パネルを作成 / Build the text formatting panel */
    function buildFormatPanel(dialog) {
        var formatPanel = dialog.add('panel', undefined, getLabel('formatPanel'));
        setupPanel(formatPanel);

        var fontAndStyleCheckbox = formatPanel.add('checkbox', undefined, getLabel('fontAndStyle'));
        var fontSizeCheckbox = formatPanel.add('checkbox', undefined, getLabel('fontSize'));
        fontAndStyleCheckbox.value = false;
        fontSizeCheckbox.value = false;

        return {
            panel: formatPanel,
            fontAndStyleCheckbox: fontAndStyleCheckbox,
            fontSizeCheckbox: fontSizeCheckbox
        };
    }

    /* 基本的な塗りや線パネルを作成 / Build the basic fill & stroke panel */
    function buildBasicFillStrokePanel(dialog) {
        var basicFillStrokePanel = dialog.add('panel', undefined, getLabel('basicFillStrokePanel'));
        setupPanel(basicFillStrokePanel);

        var fillCheckbox = basicFillStrokePanel.add('checkbox', undefined, getLabel('swapFill'));
        var strokeColorCheckbox = basicFillStrokePanel.add('checkbox', undefined, getLabel('swapStrokeColor'));
        var strokeWidthCheckbox = basicFillStrokePanel.add('checkbox', undefined, getLabel('swapStrokeWidth'));
        fillCheckbox.value = false;
        strokeColorCheckbox.value = false;
        strokeWidthCheckbox.value = false;

        return {
            panel: basicFillStrokePanel,
            fillCheckbox: fillCheckbox,
            strokeColorCheckbox: strokeColorCheckbox,
            strokeWidthCheckbox: strokeWidthCheckbox
        };
    }

    /* グラフィックスタイルパネルを作成 / Build the graphic style panel */
    function buildGraphicStylePanel(dialog) {
        var graphicStylePanel = dialog.add('panel', undefined, getLabel('graphicStylePanel'));
        setupPanel(graphicStylePanel);
        graphicStylePanel.helpTip = getLabel('tipGraphicStyle');

        var swapGraphicStylesCheckbox = graphicStylePanel.add('checkbox', undefined, getLabel('swapGraphicStyles'));
        var deleteStylesCheckbox = graphicStylePanel.add('checkbox', undefined, getLabel('deleteStyles'));
        swapGraphicStylesCheckbox.value = true;
        deleteStylesCheckbox.value = true;
        deleteStylesCheckbox.helpTip = getLabel('tipDeleteStyles');

        return {
            panel: graphicStylePanel,
            swapGraphicStylesCheckbox: swapGraphicStylesCheckbox,
            deleteStylesCheckbox: deleteStylesCheckbox
        };
    }

    /* 座標パネルを作成（3×3 の基準点ラジオ） / Build the position panel (3x3 reference-point radios) */
    function buildCoordinatePanel(parent) {
        var coordinatePanel = parent.add('panel', undefined, getLabel('coordinatePanel'));
        setupPanel(coordinatePanel);
        coordinatePanel.alignChildren = ['center', 'top'];
        coordinatePanel.helpTip = getLabel('tipAnchor');

        var grid = coordinatePanel.add('group');
        grid.orientation = 'column';
        grid.spacing = 4;
        grid.alignChildren = ['center', 'center'];

        // 行優先（row-major）の 9 個。index 0=左上 … 4=中央 … 8=右下
        var anchorRadios = [];
        for (var row = 0; row < 3; row++) {
            var rowGroup = grid.add('group');
            rowGroup.orientation = 'row';
            rowGroup.spacing = 4;
            for (var col = 0; col < 3; col++) {
                anchorRadios.push(rowGroup.add('radiobutton', undefined, ''));
            }
        }
        anchorRadios[4].value = true; // 中央を初期値

        // 軸ロック（両方／横X／縦Y）。同一グループ内なので自動排他
        var axisGroup = coordinatePanel.add('group');
        axisGroup.orientation = 'column';
        axisGroup.alignChildren = ['left', 'center'];
        axisGroup.spacing = 6;
        axisGroup.add('statictext', undefined, getLabel('axisLabel'));
        var axisBothRadio = axisGroup.add('radiobutton', undefined, getLabel('axisBoth'));
        var axisXRadio = axisGroup.add('radiobutton', undefined, getLabel('axisX'));
        var axisYRadio = axisGroup.add('radiobutton', undefined, getLabel('axisY'));
        axisBothRadio.value = true;

        // 重なり順（前後関係）も交換
        var zOrderCheckbox = coordinatePanel.add('checkbox', undefined, getLabel('swapZOrder'));
        zOrderCheckbox.value = false;

        return {
            panel: coordinatePanel,
            anchorRadios: anchorRadios,
            axisBothRadio: axisBothRadio,
            axisXRadio: axisXRadio,
            axisYRadio: axisYRadio,
            zOrderCheckbox: zOrderCheckbox
        };
    }

    /* 選択中の基準点インデックス（0〜8、未選択時は中央=4） / Selected reference-point index */
    function getSelectedAnchorIndex(anchorRadios) {
        for (var i = 0; i < anchorRadios.length; i++) {
            if (anchorRadios[i].value) return i;
        }
        return 4;
    }

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

    /* フッター行を作成（左：プレビュー／中央：スペーサー／右：ボタン） / Build the footer row (left: preview / center: spacer / right: buttons) */
    function buildFooter(dialog) {
        var buttonRow = addButtonRow(dialog);

        /* 左：プレビュー / Left: preview */
        var previewCheckbox = buttonRow.leftGroup.add('checkbox', undefined, getLabel('preview'));
        previewCheckbox.value = false;

        /* 右：キャンセル／OK / Right: Cancel / OK */
        var btnCancel = buttonRow.rightGroup.add('button', undefined, getLabel('cancel'), { name: 'cancel' });
        var btnOK = buttonRow.rightGroup.add('button', undefined, 'OK', { name: 'ok' });

        return {
            group: buttonRow.rowGroup,
            previewCheckbox: previewCheckbox,
            cancelButton: btnCancel,
            okButton: btnOK
        };
    }

    /* ダイアログの有効／無効状態を更新 / Update dialog enabled states */
    function updateDialogEnabled(ui, bothAreText) {
        var isStyleSwap = ui.mode.styleSwapRadio.value;
        var isCoordinateSwap = ui.mode.coordinateSwapRadio.value;
        var anyBasicOn = ui.basicFillStroke.fillCheckbox.value
            || ui.basicFillStroke.strokeColorCheckbox.value
            || ui.basicFillStroke.strokeWidthCheckbox.value;
        ui.format.panel.enabled = isStyleSwap && bothAreText;
        ui.basicFillStroke.panel.enabled = isStyleSwap;
        ui.graphicStyle.panel.enabled = isStyleSwap && !anyBasicOn;
        // 座標パネルは「座標」選択時のみ有効（それ以外はディム）
        ui.coordinate.panel.enabled = isCoordinateSwap;
        if (isStyleSwap) {
            ui.graphicStyle.deleteStylesCheckbox.enabled = ui.graphicStyle.swapGraphicStylesCheckbox.value;
        }
    }

    /* ダイアログの選択内容を読む / Read dialog choices */
    function readDialogOptions(ui) {
        var anyBasicOn = ui.basicFillStroke.fillCheckbox.value
            || ui.basicFillStroke.strokeColorCheckbox.value
            || ui.basicFillStroke.strokeWidthCheckbox.value;
        return {
            mode: ui.mode.styleSwapRadio.value ? 'style'
                : (ui.mode.coordinateSwapRadio.value ? 'coordinate' : 'content'),
            swapGraphicStyles: ui.graphicStyle.swapGraphicStylesCheckbox.value && !anyBasicOn,
            deleteStyles: ui.graphicStyle.deleteStylesCheckbox.value,
            includeFontAndStyle: ui.format.fontAndStyleCheckbox.value,
            includeFontSize: ui.format.fontSizeCheckbox.value,
            includeFill: ui.basicFillStroke.fillCheckbox.value,
            includeStrokeColor: ui.basicFillStroke.strokeColorCheckbox.value,
            includeStrokeWidth: ui.basicFillStroke.strokeWidthCheckbox.value,
            anchor: getSelectedAnchorIndex(ui.coordinate.anchorRadios),
            axis: ui.coordinate.axisXRadio.value ? 'x'
                : (ui.coordinate.axisYRadio.value ? 'y' : 'both'),
            swapZOrder: ui.coordinate.zOrderCheckbox.value
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

    // ダイアログの位置と不透明度（再利用パーツ）ここまで / End of the reusable dialog position and opacity

    /* オプションダイアログを表示し、プレビューと確定を駆動 / Show the options dialog and drive preview / commit */
    function showOptionsDialog(bothAreText, performSwapFn) {
        var dialog = new Window('dialog', getLabel('dialogTitle') + ' ' + SCRIPT_VERSION);
        dialog.orientation = 'column';
        dialog.alignChildren = 'fill';

        var modeUi = buildModePanel(dialog, bothAreText);

        // 本体を2カラムに：左＝スタイル系、右＝座標 / Two-column body: left = style panels, right = position
        var columns = dialog.add('group');
        columns.orientation = 'row';
        columns.alignment = 'fill';
        columns.alignChildren = ['fill', 'top'];

        var leftColumn = columns.add('group');
        leftColumn.orientation = 'column';
        leftColumn.alignChildren = ['fill', 'top'];

        var rightColumn = columns.add('group');
        rightColumn.orientation = 'column';
        rightColumn.alignChildren = ['fill', 'top'];

        var ui = {
            mode: modeUi,
            format: buildFormatPanel(leftColumn),
            basicFillStroke: buildBasicFillStrokePanel(leftColumn),
            graphicStyle: buildGraphicStylePanel(leftColumn),
            coordinate: buildCoordinatePanel(rightColumn)
        };
        var footer = buildFooter(dialog);
        var previewUi = { previewCheckbox: footer.previewCheckbox };
        var buttons = footer;

        var previewState = { isUndo: false };

        /* UI 変更時：有効状態を更新し、プレビューを再描画 / On UI change: refresh enabled state and re-render preview */
        var refresh = function () {
            updateDialogEnabled(ui, bothAreText);
            runPreview(previewState, function () {
                performSwapFn(readDialogOptions(ui), true);
            }, previewUi.previewCheckbox.value);
        };

        ui.mode.styleSwapRadio.onClick = refresh;
        ui.mode.contentSwapRadio.onClick = refresh;
        ui.mode.coordinateSwapRadio.onClick = refresh;
        ui.graphicStyle.swapGraphicStylesCheckbox.onClick = refresh;
        ui.graphicStyle.deleteStylesCheckbox.onClick = refresh;
        ui.basicFillStroke.fillCheckbox.onClick = refresh;
        ui.basicFillStroke.strokeColorCheckbox.onClick = refresh;
        ui.basicFillStroke.strokeWidthCheckbox.onClick = refresh;
        ui.format.fontAndStyleCheckbox.onClick = refresh;
        ui.format.fontSizeCheckbox.onClick = refresh;
        // 3×3 ラジオは行ごとにグループが分かれ自動排他が効かないため、手動で排他制御
        var anchorRadios = ui.coordinate.anchorRadios;
        for (var ai = 0; ai < anchorRadios.length; ai++) {
            (function (index) {
                anchorRadios[index].onClick = function () {
                    for (var k = 0; k < anchorRadios.length; k++) {
                        anchorRadios[k].value = (k === index);
                    }
                    refresh();
                };
            })(ai);
        }
        ui.coordinate.axisBothRadio.onClick = refresh;
        ui.coordinate.axisXRadio.onClick = refresh;
        ui.coordinate.axisYRadio.onClick = refresh;
        ui.coordinate.zOrderCheckbox.onClick = refresh;
        previewUi.previewCheckbox.onClick = refresh;

        updateDialogEnabled(ui, bothAreText);

        /* OK：プレビュー分を巻き戻してから本番として再実行 / OK: undo preview then commit cleanly */
        buttons.okButton.onClick = function () {
            undoPreview(previewState);
            performSwapFn(readDialogOptions(ui), false);
            dialog.close(1);
        };

        /* キャンセル含むクローズ時：残ったプレビューを巻き戻す / On any close (incl. cancel): undo leftover preview */
        dialog.onClose = function () {
            cleanupPreview(previewState, app.activeDocument);
        };

        prepareDialogWindow(dialog, SCRIPT_NAME);
        return dialog.show() === 1;
    }

    // =========================================
    // グラフィックスタイル関連 / Graphic Style Helpers
    // =========================================

    /* 指定名のグラフィックスタイルが存在するか / Check if a graphic style with the given name exists */
    function hasGraphicStyle(graphicStyles, styleName) {
        try {
            graphicStyles.getByName(styleName);
            return true;
        } catch (e) {
            return false;
        }
    }

    /* 未使用の _swapStyleX/_swapStyleY ペアを返す（A/B → C/D → … → Y/Z） / Return the next unused _swapStyleX/_swapStyleY pair */
    function findNextStyleNamePair(graphicStyles) {
        var suffixLetters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
        for (var i = 0; i < suffixLetters.length - 1; i += 2) {
            var firstStyleName = '_swapStyle' + suffixLetters.charAt(i);
            var secondStyleName = '_swapStyle' + suffixLetters.charAt(i + 1);
            if (!hasGraphicStyle(graphicStyles, firstStyleName) && !hasGraphicStyle(graphicStyles, secondStyleName)) {
                return [firstStyleName, secondStyleName];
            }
        }
        return null;
    }

    // 一時アクション（再利用パーツ） / Temporary action (reusable)

    /**
     * 文字列を UTF-8 のバイト列の16進にする（アクション定義の /name・/localizedName 用）
     * @param {string} sourceText - 変換する文字列
     * @returns {string} 16進の文字列（2文字で1バイト）
     */
    function toActionHex(sourceText) {
        var utf8Text = unescape(encodeURIComponent(String(sourceText)));
        var hexText = "";
        for (var i = 0; i < utf8Text.length; i++) {
            var hexByte = utf8Text.charCodeAt(i).toString(16);
            hexText += (hexByte.length < 2 ? "0" : "") + hexByte;
        }
        return hexText;
    }

    /**
     * アクション定義の「/name [ バイト数 16進 ]」の3行を返す
     * @param {string} indent - 行頭の字下げ（"\t" など）
     * @param {string} nameText - 名前
     * @param {string} [fieldName] - 項目名（既定は "name"。"localizedName" など）
     * @returns {string[]} 3行ぶんの配列
     */
    function buildActionNameLines(indent, nameText, fieldName) {
        var nameHex = toActionHex(nameText);
        return [
            indent + "/" + (fieldName || "name") + " [ " + (nameHex.length / 2),
            indent + "\t" + nameHex,
            indent + "]"
        ];
    }

    /**
     * アクション定義を一時ファイルに書き出してセットを読み込む。読み込んだら一時ファイルは消す
     * （読み込んだ時点で解釈済みなので、以降の失敗でファイルが残らない）
     * @param {string} actionSource - アクション定義のテキスト
     * @param {string} setName - アクションセット名
     * @returns {boolean} 読み込めたら true
     */
    function loadTemporaryActionSet(actionSource, setName) {
        var actionFile = new File(Folder.temp + "/" + setName + "_" + new Date().getTime() + ".aia");
        try {
            actionFile.encoding = "UTF-8";
            if (!actionFile.open("w")) throw new Error("cannot open " + actionFile.fsName);
            actionFile.write(actionSource);
            actionFile.close();
            /* 前回の失敗で同じ名前のセットが残っていれば外す / Remove a same-name set left by an earlier failure */
            unloadTemporaryActionSet(setName);
            app.loadAction(actionFile);
            return true;
        } catch (e) {
            $.writeln("loadTemporaryActionSet: " + e);
            return false;
        } finally {
            try { actionFile.close(); } catch (closeError) { /* 閉じ済み / already closed */ }
            try { actionFile.remove(); } catch (removeError) { /* 消せなくても続ける / keep going */ }
        }
    }

    /**
     * 一時アクションのセットを解除する（読み込まれていなくてもエラーにしない）
     * @param {string} setName - アクションセット名
     * @returns {void}
     */
    function unloadTemporaryActionSet(setName) {
        try {
            app.unloadAction(setName, "");
        } catch (e) {
            /* 読み込まれていない / not loaded */
        }
    }

    /**
     * アクション定義を読み込んで1回実行し、解除する。途中で失敗しても解除は必ず試みる
     * @param {string} actionSource - アクション定義のテキスト
     * @param {string} setName - アクションセット名
     * @param {string} actionName - 実行するアクション名
     * @returns {boolean} 実行できたら true
     */
    function runTemporaryAction(actionSource, setName, actionName) {
        if (!loadTemporaryActionSet(actionSource, setName)) return false;
        try {
            app.doScript(actionName, setName);
            return true;
        } catch (e) {
            $.writeln("runTemporaryAction: " + e);
            return false;
        } finally {
            unloadTemporaryActionSet(setName);
        }
    }

    // 一時アクション（再利用パーツ）ここまで / End of the reusable temporary action

    // 既存のユーザーアクションセットと衝突しないようユニークな名前を使う
    var TEMP_ACTION_SET_NAME = '__SwapStyleGS';
    var TEMP_ACTION_NAME = 'AddNewWithoutName';

    /* 無名グラフィックスタイルを追加するアクション定義を返す / Build the force-new-style action definition */
    function buildForceNewGraphicStyleAction() {
        // hex の `[ 13 5f5f537761705374796c654753 ]` はセット名 "__SwapStyleGS" (UTF-8) を表す
        return '/version 3 /name [ 13 5f5f537761705374796c654753 ] /isOpen 1 /actionCount 1 /action-1 { /name [ 17 4164644e6577576974686f75744e616d65 ] /keyIndex 0 /colorIndex 0 /isOpen 1 /eventCount 1 /event-1 { /useRulersIn1stQuadrant 0 /internalName (ai_plugin_styles) /localizedName [ 30 e382b0e383a9e38395e382a3e38383e382afe382b9e382bfe382a4e383ab ] /isOpen 1 /isOn 1 /hasDialog 1 /showDialog 0 /parameterCount 1 /parameter-1 { /key 1835363957 /showInPalette 4294967295 /type (enumerated) /name [ 36 e696b0e8a68fe382b0e383a9e38395e382a3e38383e382afe382b9e382bfe382a4e383ab ] /value 1 } } }';
    }

    // =========================================
    // 座標の交換 / Position Swap
    // =========================================

    /* geometricBounds と基準点インデックスから基準点座標を返す / Reference point from geometricBounds + anchor index */
    function anchorPoint(bounds, anchorIndex) {
        // geometricBounds = [left, top, right, bottom]、index は行優先（0=左上 … 4=中央 … 8=右下）
        var col = anchorIndex % 3;
        var row = Math.floor(anchorIndex / 3);
        var x = (col === 0) ? bounds[0] : (col === 2 ? bounds[2] : (bounds[0] + bounds[2]) / 2);
        var y = (row === 0) ? bounds[1] : (row === 2 ? bounds[3] : (bounds[1] + bounds[3]) / 2);
        return [x, y];
    }

    /* 2つのオブジェクトの位置を入れ替える（基準点・軸ロックを反映） / Swap positions by reference point with axis lock */
    function swapPositions(itemA, itemB, options) {
        var anchorIndex = (typeof options.anchor === 'number') ? options.anchor : 4;
        var axis = options.axis || 'both';
        var moveX = (axis === 'both' || axis === 'x');
        var moveY = (axis === 'both' || axis === 'y');
        var boundsA = itemA.geometricBounds;
        var boundsB = itemB.geometricBounds;
        var pointA = anchorPoint(boundsA, anchorIndex);
        var pointB = anchorPoint(boundsB, anchorIndex);
        // translate(deltaX, deltaY)：deltaY は上方向が正なので geometricBounds と整合
        itemA.translate(moveX ? (pointB[0] - pointA[0]) : 0, moveY ? (pointB[1] - pointA[1]) : 0);
        itemB.translate(moveX ? (pointA[0] - pointB[0]) : 0, moveY ? (pointA[1] - pointB[1]) : 0);
    }

    /* 2つのオブジェクトの重なり順を入れ替える / Swap the stacking order of two objects */
    function swapZOrder(itemA, itemB) {
        try {
            // 各オブジェクトの直前にマーカーを置いて元位置を記録し、相手側のマーカー位置へ移動
            var markerA = itemA.parent.pathItems.add();
            markerA.move(itemA, ElementPlacement.PLACEBEFORE);
            var markerB = itemB.parent.pathItems.add();
            markerB.move(itemB, ElementPlacement.PLACEBEFORE);
            itemA.move(markerB, ElementPlacement.PLACEBEFORE);
            itemB.move(markerA, ElementPlacement.PLACEBEFORE);
            markerA.remove();
            markerB.remove();
        } catch (e) { }
    }

    // =========================================
    // 基本属性の交換 / Basic Fill & Stroke Swap
    // =========================================

    /* 基本的な塗りや線をクロス入れ替え（TextFrame は characterAttributes 経由） / Cross-swap basic fill & stroke */
    function swapBasicFillStroke(itemA, itemB, options) {
        var bothAreText = itemA.typename === 'TextFrame' && itemB.typename === 'TextFrame';
        if (bothAreText) {
            swapTextCharacterFillStroke(itemA, itemB, options);
        } else {
            swapPageItemFillStroke(itemA, itemB, options);
        }
    }

    /* PageItem 系（PathItem など）の塗り・線をクロス入れ替え / Cross-swap fill & stroke at PageItem level */
    function swapPageItemFillStroke(itemA, itemB, options) {
        if (options.includeFill) {
            try {
                var fillColorA = itemA.fillColor;
                var fillColorB = itemB.fillColor;
                var filledA = itemA.filled;
                var filledB = itemB.filled;
                itemA.fillColor = fillColorB;
                itemB.fillColor = fillColorA;
                itemA.filled = filledB;
                itemB.filled = filledA;
            } catch (e) { }
        }
        if (options.includeStrokeColor) {
            try {
                var strokeColorA = itemA.strokeColor;
                var strokeColorB = itemB.strokeColor;
                var strokedA = itemA.stroked;
                var strokedB = itemB.stroked;
                itemA.strokeColor = strokeColorB;
                itemB.strokeColor = strokeColorA;
                itemA.stroked = strokedB;
                itemB.stroked = strokedA;
            } catch (e) { }
        }
        if (options.includeStrokeWidth) {
            try {
                var strokeWidthA = itemA.strokeWidth;
                var strokeWidthB = itemB.strokeWidth;
                itemA.strokeWidth = strokeWidthB;
                itemB.strokeWidth = strokeWidthA;
            } catch (e) { }
        }
    }

    /* TextFrame の characterAttributes 経由で塗り・線をクロス入れ替え / Cross-swap text fill & stroke via characterAttributes */
    function swapTextCharacterFillStroke(textFrameA, textFrameB, options) {
        var charAttrsA = textFrameA.textRange.characterAttributes;
        var charAttrsB = textFrameB.textRange.characterAttributes;
        if (options.includeFill) {
            try {
                var fillColorA = charAttrsA.fillColor;
                var fillColorB = charAttrsB.fillColor;
                charAttrsA.fillColor = fillColorB;
                charAttrsB.fillColor = fillColorA;
            } catch (e) { }
        }
        if (options.includeStrokeColor) {
            try {
                var strokeColorA = charAttrsA.strokeColor;
                var strokeColorB = charAttrsB.strokeColor;
                charAttrsA.strokeColor = strokeColorB;
                charAttrsB.strokeColor = strokeColorA;
            } catch (e) { }
        }
        if (options.includeStrokeWidth) {
            try {
                // CharacterAttributes は strokeWidth ではなく strokeWeight
                var strokeWeightA = charAttrsA.strokeWeight;
                var strokeWeightB = charAttrsB.strokeWeight;
                charAttrsA.strokeWeight = strokeWeightB;
                charAttrsB.strokeWeight = strokeWeightA;
            } catch (e) { }
        }
    }

    // =========================================
    // プレビュー undo ヘルパー / Preview Undo Helpers
    // =========================================
    /*
     * 設定変更のたびに「仮配置 → undo → 再生成」する典型パターンを
     * 共通化したユーティリティ。
     * Reusable preview/undo pattern for Illustrator dialog scripts.
     *
     * 使い方の概略 / Usage:
     *   var previewState = { isUndo: false };
     *   // UI 変更ごとに: runPreview(previewState, process, isPreview.value);
     *   // OK 時に:       undoPreview(previewState); process(); win.close();
     *   // onClose 時:    cleanupPreview(previewState, app.activeDocument);
     *
     * 注意 / Notes:
     *  - process() 内で graphicStyles.add() / layers.add() など
     *    app.undo() で戻らない副作用を起こす場合、cleanupPreview だけでは
     *    残骸が残る可能性がある（呼び出し側で手動 .remove() を検討）。
     *  - process() は ScriptUI の選択値を参照する純粋な再描画関数として
     *    書くと、内部状態を保持しない実装にしやすい。
     */

    /* プレビューを再描画する / Re-render preview */
    function runPreview(state, processFn, isEnabled) {
        try {
            if (isEnabled) {
                if (state.isUndo) app.undo();
                else state.isUndo = true;
                processFn();
                app.redraw();
            } else if (state.isUndo) {
                app.undo();
                app.redraw();
                state.isUndo = false;
            }
        } catch (err) { }
    }

    /* 確定処理の直前にプレビュー分を巻き戻す / Undo preview before the final commit */
    function undoPreview(state) {
        try {
            if (state.isUndo) app.undo();
        } catch (err) { }
        state.isUndo = false;
    }

    /* ダイアログクローズ時のクリーンアップ / Cleanup on dialog close */
    function cleanupPreview(state, doc, tempLayerName) {
        try {
            if (state.isUndo) app.undo();
            state.isUndo = false;
        } catch (err) { }
        if (tempLayerName) {
            try {
                var tmpLay = doc.layers.getByName(tempLayerName);
                tmpLay.remove();
            } catch (err) { }
        }
    }

    // =========================================
    // 交換処理本体 / Swap Operation
    // =========================================

    /* スタイル／文字列の交換を実行（プレビュー・本番共通） / Run the swap (used for both preview and commit) */
    function performSwap(activeDoc, graphicStyles, targetItems, options, isPreview) {
        // 文字列交換モード：テキストオブジェクト同士で contents を入れ替えるだけ
        if (options.mode === 'content') {
            if (targetItems[0].typename !== 'TextFrame' || targetItems[1].typename !== 'TextFrame') {
                if (!isPreview) alert(getLabel('alertContentTextOnly'));
                return;
            }
            var contentA = targetItems[0].contents;
            targetItems[0].contents = targetItems[1].contents;
            targetItems[1].contents = contentA;
            activeDoc.selection = targetItems;
            return;
        }

        // 座標交換モード：2つのオブジェクトの位置（左上基準）を入れ替える
        if (options.mode === 'coordinate') {
            swapPositions(targetItems[0], targetItems[1], options);
            if (options.swapZOrder) {
                swapZOrder(targetItems[0], targetItems[1]);
            }
            activeDoc.selection = targetItems;
            return;
        }

        if (options.swapGraphicStyles) {
            var tempStyleNamePair = findNextStyleNamePair(graphicStyles);
            if (!tempStyleNamePair) {
                if (!isPreview) alert(getLabel('alertNoFreePair'));
                return;
            }

            // 途中失敗時の掃除のため、実際に登録できた名前だけを別配列で追跡
            var registeredStyleNames = [];
            if (!loadTemporaryActionSet(buildForceNewGraphicStyleAction(), TEMP_ACTION_SET_NAME)) {
                if (!isPreview) alert(getLabel('alertSwapFailed'));
                return;
            }
            try {
                for (var i = 0; i < targetItems.length; i++) {
                    activeDoc.selection = null;
                    targetItems[i].selected = true;
                    var styleCountBefore = graphicStyles.length;
                    app.doScript(TEMP_ACTION_NAME, TEMP_ACTION_SET_NAME, false);
                    // doScript が無音で失敗した場合、最後尾を取ると既存スタイルをリネームしてしまう
                    if (graphicStyles.length <= styleCountBefore) {
                        throw new Error('registration failed');
                    }
                    graphicStyles[graphicStyles.length - 1].name = tempStyleNamePair[i];
                    registeredStyleNames.push(tempStyleNamePair[i]);
                }

                // A に styleB、B に styleA を適用（クロス）
                graphicStyles.getByName(tempStyleNamePair[1]).applyTo(targetItems[0]);
                graphicStyles.getByName(tempStyleNamePair[0]).applyTo(targetItems[1]);

                // ON のとき登録したスタイルを削除（アピアランス自体は各オブジェクトに残る）
                if (options.deleteStyles) {
                    for (var j = 0; j < registeredStyleNames.length; j++) {
                        try { graphicStyles.getByName(registeredStyleNames[j]).remove(); }
                        catch (e) { }
                    }
                    registeredStyleNames = [];
                }
            } catch (mainErr) {
                // 失敗時は登録済みの一時スタイルを掃除してから抜ける
                for (var k = 0; k < registeredStyleNames.length; k++) {
                    try { graphicStyles.getByName(registeredStyleNames[k]).remove(); }
                    catch (e) { }
                }
                if (!isPreview) alert(getLabel('alertSwapFailed'));
                return;
            } finally {
                unloadTemporaryActionSet(TEMP_ACTION_SET_NAME);
            }
        }

        // 基本的な塗りや線のクロス入れ替え
        if (options.includeFill || options.includeStrokeColor || options.includeStrokeWidth) {
            swapBasicFillStroke(targetItems[0], targetItems[1], options);
        }

        // 書式情報のクロス入れ替え（両方ともテキストオブジェクトのとき）
        if ((options.includeFontAndStyle || options.includeFontSize)
            && targetItems[0].typename === 'TextFrame'
            && targetItems[1].typename === 'TextFrame') {
            var charAttrsA = targetItems[0].textRange.characterAttributes;
            var charAttrsB = targetItems[1].textRange.characterAttributes;
            var fontA = charAttrsA.textFont;
            var fontB = charAttrsB.textFont;
            var sizeA = charAttrsA.size;
            var sizeB = charAttrsB.size;

            if (options.includeFontAndStyle) {
                charAttrsA.textFont = fontB;
                charAttrsB.textFont = fontA;
            }
            if (options.includeFontSize) {
                charAttrsA.size = sizeB;
                charAttrsB.size = sizeA;
            }
        }

        // 選択を元の2つに戻す
        activeDoc.selection = targetItems;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    (function () {
        if (app.documents.length === 0) {
            alert(getLabel('alertNoDocument'));
            return;
        }
        var activeDoc = app.activeDocument;
        var graphicStyles = activeDoc.graphicStyles;

        var selectedItems = activeDoc.selection;
        if (selectedItems.length !== 2) {
            alert(getLabel('alertSelectTwo'));
            return;
        }

        // selection は後続処理で変更するため、対象2件を配列として保持 / Keep the two targets as an array because selection changes later
        var targetItems = [selectedItems[0], selectedItems[1]];

        var bothAreText = (targetItems[0].typename === 'TextFrame'
            && targetItems[1].typename === 'TextFrame');

        var performSwapFn = function (options, isPreview) {
            performSwap(activeDoc, graphicStyles, targetItems, options, isPreview);
        };

        showOptionsDialog(bothAreText, performSwapFn);
    })();

})();
