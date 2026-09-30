#target illustrator
#targetengine "SetStrokeAlignmentEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したオブジェクトに、線幅・線端・角の形状・線の位置をまとめて設定します。
DOMからは変更できない線の位置は、一時アクション（ダイナミックアクション）で適用します。

詳細は README を参照してください。

### Overview

Applies the stroke weight, cap, corner join, and alignment to the selected objects at once.
The alignment, which the DOM cannot change, is applied through a temporary action.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SetStrokeAlignment";           /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.6";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-09-11";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-01";                   /* 更新日 / last updated */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================
    var CONFIG = {
        defaultStrokeWidth: "1",          /* 線幅の初期値（pt）/ default stroke weight in pt */
        defaultStrokeCap: "butt",         /* 線端の初期値（butt / round / projecting）/ default cap */
        defaultCornerJoin: "miter",       /* 角の形状の初期値（miter / round / bevel）/ default join */
        defaultStrokeAlignment: "center", /* 線の位置の初期値（center / inside / outside）/ default alignment */
        previewDefault: true              /* プレビューを既定でON / preview on by default */
    };

    // =========================================
    // 一時アクション / Temporary action
    // =========================================
    var ACTION_SET_NAME   = "StrokeAlign";  /* アクションセット名（英数字）/ action set name (ASCII) */
    var ACTION_NAME       = "SetStroke";    /* アクション名（英数字）/ action name (ASCII) */
    var ACTION_PARAM_NAME = "Alignment";    /* パラメーターの表示名 / parameter display name */

    /* 線幅・線端・角の形状はDOMで設定できるので、アクションは線の位置だけを担当する
       / the DOM covers weight, cap and join, so the action only carries the alignment */
    var ACTION_KEY_STROKE_ALIGNMENT = 1634494318;  /* 'algn' 線の位置 / stroke alignment */

    /* 線の位置の識別子 → アクションの enumerated 値 / alignment keys to action enum values */
    var STROKE_ALIGNMENT_VALUES = {
        center: 0,
        inside: 1,
        outside: 2
    };

    // =========================================
    // ローカライズ / Localize
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

    var LABELS = {
        dialog: {
            title: { ja: "線の設定", en: "Stroke Settings" }
        },
        fieldLabel: {
            strokeWidth: { ja: "線幅", en: "Weight" },
            strokeCap: { ja: "線端", en: "Cap" },
            cornerJoin: { ja: "角の形状", en: "Corner" },
            strokeAlignment: { ja: "線の位置", en: "Alignment" }
        },
        radio: {
            buttCap: { ja: "なし", en: "Butt" },
            roundCap: { ja: "丸型", en: "Round" },
            projectingCap: { ja: "突出", en: "Projecting" },
            miterJoin: { ja: "マイター", en: "Miter" },
            roundJoin: { ja: "ラウンド", en: "Round" },
            bevelJoin: { ja: "ベベル", en: "Bevel" },
            centerAlign: { ja: "中央", en: "Center" },
            insideAlign: { ja: "内側", en: "Inside" },
            outsideAlign: { ja: "外側", en: "Outside" }
        },
        // 線パネルのツールチップと同じ表記 / same wording as the Stroke panel tooltips
        tooltip: {
            buttCap: { ja: "線端なし", en: "Butt Cap" },
            roundCap: { ja: "丸型線端", en: "Round Cap" },
            projectingCap: { ja: "突出線端", en: "Projecting Cap" },
            miterJoin: { ja: "マイター結合", en: "Miter Join" },
            roundJoin: { ja: "ラウンド結合", en: "Round Join" },
            bevelJoin: { ja: "ベベル結合", en: "Bevel Join" },
            centerAlign: { ja: "線を中央に揃える", en: "Align Stroke to Center" },
            insideAlign: { ja: "線を内側に揃える", en: "Align Stroke to Inside" },
            outsideAlign: { ja: "線を外側に揃える", en: "Align Stroke to Outside" },
            stepUp: {
                ja: "値を増やす（shift＋クリックで10の倍数へ、option＋クリックで0.1ずつ）",
                en: "Increase (Shift-click to snap to 10s, Option-click by 0.1)"
            },
            stepDown: {
                ja: "値を減らす（shift＋クリックで10の倍数へ、option＋クリックで0.1ずつ）",
                en: "Decrease (Shift-click to snap to 10s, Option-click by 0.1)"
            },
            stepUpInteger: { ja: "値を増やす（shift＋クリックで10の倍数へ）", en: "Increase (Shift-click to snap to 10s)" },
            stepDownInteger: { ja: "値を減らす（shift＋クリックで10の倍数へ）", en: "Decrease (Shift-click to snap to 10s)" }
        },
        checkbox: {
            preview: { ja: "プレビュー", en: "Preview" }
        },
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" }
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noSelection: { ja: "オブジェクトを選択してください。", en: "Select one or more objects." },
            invalidWidth: { ja: "線幅には0より大きい数値を入力してください。", en: "Enter a stroke weight greater than 0." }
        }
    };

    /* ラジオボタンの並び（線パネルと同じ順序）/ Radio button order, matching the Stroke panel */
    var STROKE_CAP_OPTIONS = [
        { key: "butt", label: LABELS.radio.buttCap, tooltip: LABELS.tooltip.buttCap },
        { key: "round", label: LABELS.radio.roundCap, tooltip: LABELS.tooltip.roundCap },
        { key: "projecting", label: LABELS.radio.projectingCap, tooltip: LABELS.tooltip.projectingCap }
    ];
    var CORNER_JOIN_OPTIONS = [
        { key: "miter", label: LABELS.radio.miterJoin, tooltip: LABELS.tooltip.miterJoin },
        { key: "round", label: LABELS.radio.roundJoin, tooltip: LABELS.tooltip.roundJoin },
        { key: "bevel", label: LABELS.radio.bevelJoin, tooltip: LABELS.tooltip.bevelJoin }
    ];
    var STROKE_ALIGNMENT_OPTIONS = [
        { key: "center", label: LABELS.radio.centerAlign, tooltip: LABELS.tooltip.centerAlign },
        { key: "inside", label: LABELS.radio.insideAlign, tooltip: LABELS.tooltip.insideAlign },
        { key: "outside", label: LABELS.radio.outsideAlign, tooltip: LABELS.tooltip.outsideAlign }
    ];

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

    var ROW_SPACING           = 8;                           /* 行内の要素間隔 */
    var LABEL_WIDTH           = (uiLang === "ja") ? 76 : 74; /* 行ラベルの共通幅（コロンまで収まる幅）*/

    /* ラベル付きの行を追加し、コントロールを入れるグループを返す / Add a labeled row and return its control group
       @param {Window|Group} parentContainer - 追加先
       @param {object} labelSet - 行ラベルのラベルセット
       @returns {Group} コントロールを入れるグループ */
    function addControlRow(parentContainer, labelSet) {
        var controlRow = parentContainer.add("group");
        setupRow(controlRow, "left", ROW_SPACING);

        var rowLabel = controlRow.add("statictext", undefined, labelText(labelSet));
        rowLabel.preferredSize.width = LABEL_WIDTH;
        rowLabel.justify = "right";

        var rowControlGroup = controlRow.add("group");
        setupRow(rowControlGroup, "left", ROW_SPACING);
        return rowControlGroup;
    }

    /* ラジオボタンの行を追加する / Add a labeled row of radio buttons
       @param {Window|Group} parentContainer - 追加先
       @param {object} labelSet - 行ラベルのラベルセット
       @param {Array} optionDefinitions - 選択肢の定義（key / label / tooltip）
       @param {string} selectedOptionKey - 初期選択のキー
       @param {function} onSelect - 選択が変わったときに呼ぶ処理
       @returns {Group} 選択状態を selectedOptionKey に持つグループ */
    function addRadioRow(parentContainer, labelSet, optionDefinitions, selectedOptionKey, onSelect) {
        var radioGroup = addControlRow(parentContainer, labelSet);
        radioGroup.selectedOptionKey = selectedOptionKey;
        for (var i = 0; i < optionDefinitions.length; i++) {
            var optionRadio = radioGroup.add("radiobutton", undefined, getLabel(optionDefinitions[i].label));
            optionRadio.helpTip = getLabel(optionDefinitions[i].tooltip);
            optionRadio.optionKey = optionDefinitions[i].key;
            optionRadio.value = (optionDefinitions[i].key === selectedOptionKey);
            optionRadio.onClick = function () {
                radioGroup.selectedOptionKey = this.optionKey;
                onSelect();
            };
        }
        return radioGroup;
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

    // UI の明暗（再利用パーツ） / UI theme (reusable)

    /**
     * UI がダークテーマかどうかを判定する（Illustrator は uiBrightness、InDesign は uiBrightnessPreference）
     * @returns {boolean} ダークなら true。取得できない環境では false（明るいUI扱い）
     */
    function isDarkUI() {
        try {
            if (app.preferences && app.preferences.getRealPreference) {
                return app.preferences.getRealPreference("uiBrightness") <= 0.5; /* Illustrator */
            }
            return app.generalPreferences.uiBrightnessPreference <= 0.5; /* InDesign */
        } catch (e) {
            return false;
        }
    }

    // UI の明暗（再利用パーツ）ここまで / End of the reusable UI theme

    // ステップボタン（再利用パーツ） / Stepper buttons (reusable)

    // -----------------------------------------
    // ステップボタンの寸法・増減量 / Stepper metrics and steps
    // -----------------------------------------
    var STEPPER_BUTTON_WIDTH   = 20;  /* ∧∨ボタンの幅 / button width */
    var STEPPER_BUTTON_HEIGHT  = 11;  /* ∧∨ボタン1つの高さ（2つ重ねた全体の高さは22） / button height (22 for the pair) */
    var STEPPER_CORNER_RADIUS  = 2;   /* 枠の角丸の半径（ScriptUIは円弧を描けないため短い線分で近似） / corner radius, approximated with segments */
    var STEPPER_FIELD_SPACING  = 3;   /* 項目名と∧∨の間隔 / spacing between the label and the stepper */
    var STEPPER_SIDE_MARGIN    = 3;   /* ∧∨の左に足す余白（右は入力欄に突き合わせる） / extra space left of the stepper */
    var STEPPER_SHIFT_MULTIPLE = 10;  /* shift＋クリックでそろえる倍数 / Shift-click snaps to multiples of this */
    var STEPPER_OPTION_STEP    = 0.1; /* option＋クリックの増減量 / Option-click step */

    // -----------------------------------------
    // ステップボタンの配色 / Stepper colors
    // -----------------------------------------
    var STEPPER_UI_DARK           = isDarkUI();
    /* UIの明るさは4段階あり、段階ごとに背景色が違う。どの段階でも背景に対する差で見せるよう、黒・白の半透明を重ねる。
       ダーク側は Illustrator 標準のスピナー（［グリッドに分割］）で実測、明るい側は最も明るい段階（背景 約0.94）から逆算
       UI brightness has four levels with different backgrounds, so colors are translucent overlays that follow the
       dialog background. Dark values are measured from Illustrator's own spinner; light values derived for the lightest level */
    var STEPPER_FILL_COLOR        = STEPPER_UI_DARK ? [0, 0, 0, 0.10]  : [1, 1, 1, 0.50];  /* 地 / background */
    var STEPPER_FRAME_COLOR       = STEPPER_UI_DARK ? [1, 1, 1, 0.07]  : [0, 0, 0, 0.10];  /* 枠線 / frame */
    var STEPPER_PRESSED_COLOR     = STEPPER_UI_DARK ? [1, 1, 1, 0.12]  : [0, 0, 0, 0.13];  /* 押下中 / pressed */
    var STEPPER_CHEVRON_COLOR     = STEPPER_UI_DARK ? [1, 1, 1, 1]     : [0, 0, 0, 0.70];  /* 山形の線 / chevron */
    var STEPPER_DIM_FILL_COLOR    = STEPPER_UI_DARK ? [1, 1, 1, 0.035] : [1, 1, 1, 0.30];  /* 無効時の地 / background when disabled */
    var STEPPER_DIM_FRAME_COLOR   = STEPPER_UI_DARK ? [1, 1, 1, 0.035] : [0, 0, 0, 0.05];  /* 無効時の枠線（ダークは地と同じで見せない） / frame when disabled */
    var STEPPER_DIM_CHEVRON_COLOR = STEPPER_UI_DARK ? [1, 1, 1, 0.20]  : [0, 0, 0, 0.25];  /* 無効時の山形 / chevron when disabled */

    // -----------------------------------------
    // 数値欄を作る（外から呼ぶ関数） / Public API
    // -----------------------------------------
    /**
     * 「項目名・∧∨・入力欄」をひと組にした数値欄を追加する。
     * ↑↓キーでも∧∨と同じように増減する。直接入力した値も、フォーカスが外れたときに
     * 整数化・下限・上限・単位（「20 mm」の形）へそろえ、数値でなければ直前の値に戻す
     * @param {Group|Panel} parent - 追加先
     * @param {Object} fieldOptions - label（コロン込みの項目名）/ labelWidth / text / characters /
     *     step / min / max / integer（true で整数のみ）/ unit / onStep
     * @returns {EditText} 入力欄（項目名は .fieldLabel、∧∨は .stepperGroup で参照できる）
     */
    function addSteppedField(parent, fieldOptions) {
        var fieldRowGroup = parent.add("group");
        fieldRowGroup.orientation = "row";
        fieldRowGroup.alignChildren = ["left", "center"];
        fieldRowGroup.spacing = STEPPER_FIELD_SPACING;

        var fieldLabel = fieldRowGroup.add("statictext", undefined, fieldOptions.label || "");
        if (fieldOptions.labelWidth) {
            fieldLabel.preferredSize.width = fieldOptions.labelWidth;
            fieldLabel.justify = "right";
        }

        /* ∧∨と入力欄は隙間0で突き合わせる / butt the stepper against the field */
        var stepperInputGroup = fieldRowGroup.add("group");
        stepperInputGroup.orientation = "row";
        stepperInputGroup.alignChildren = ["left", "center"];
        stepperInputGroup.spacing = 0;
        stepperInputGroup.margins = 0;

        var numberInput;
        var stepperGroup = addStepper(stepperInputGroup, function () { return numberInput; }, fieldOptions);
        numberInput = stepperInputGroup.add("edittext", undefined, fieldOptions.text || "");
        numberInput.characters = fieldOptions.characters || 6;
        numberInput.fieldLabel = fieldLabel;
        numberInput.stepperGroup = stepperGroup;

        /* ↑↓キーも∧∨と同じ処理で増減する（増減量・下限・上限・単位・修飾キーをそろえる） / arrow keys share the stepper's logic */
        bindSteppedArrowKeys(numberInput, stepperGroup);

        /* 直接入力をそろえる。数値でなければ直前の値に戻す / normalize typed values; revert non-numbers */
        numberInput.lastValidText = numberInput.text;
        numberInput.onChange = function () {
            var value = parseFloat(numberInput.text);
            if (isNaN(value)) {
                numberInput.text = numberInput.lastValidText;
                return;
            }
            writeSteppedValue(numberInput, value, fieldOptions);
        };
        return numberInput;
    }

    /**
     * 数値欄の有効／無効を、項目名・∧∨ごとまとめて切り替える
     * @param {EditText} numberInput - addSteppedField() で作った入力欄
     * @param {boolean} isEnabled - 有効にするなら true
     * @returns {void}
     */
    function setSteppedFieldEnabled(numberInput, isEnabled) {
        numberInput.enabled = isEnabled;
        numberInput.fieldLabel.enabled = isEnabled;
        numberInput.stepperGroup.enabled = isEnabled;
        /* ∧∨は自作描画なので、描き直してディム表示を切り替える / redraw the custom-drawn buttons to update the dimming */
        for (var i = 0; i < numberInput.stepperGroup.children.length; i++) {
            redrawStepperGroup(numberInput.stepperGroup.children[i]);
        }
    }

    /**
     * 入力欄の値を増減する∧∨ボタンを、隙間なく縦に積んで追加する
     * @param {Group|Panel} parent - 追加先
     * @param {Function} getNumberInput - 対象の入力欄を返す関数（入力欄を∧∨より後に作れるよう、クリック時に引く）
     * @param {Object} stepOptions - step（増減量）/ min / max / integer / unit（例 " mm"）/ onStep(numberInput)
     * @returns {Group} ∧∨をまとめた group（.stepBy(direction) で同じ増減を呼べる）
     */
    function addStepper(parent, getNumberInput, stepOptions) {
        var stepperGroup = parent.add("group");
        stepperGroup.orientation = "column";
        stepperGroup.spacing = 0; /* 2つのボタンをつなげて1つの枠に見せる / join the buttons into one frame */
        stepperGroup.margins = [STEPPER_SIDE_MARGIN, 0, 0, 0]; /* 右は入力欄に突き合わせる / butt against the field on the right */
        stepperGroup.alignment = ["left", "center"];

        /**
         * 入力欄の値を増減する（shift を押しながらなら STEPPER_SHIFT_MULTIPLE の倍数へ、option なら STEPPER_OPTION_STEP ずつ。下限・上限で止める）
         * @param {number} direction - 増やすなら 1、減らすなら -1
         * @returns {void}
         */
        function stepBy(direction) {
            var numberInput = getNumberInput();
            if (!isStepperEnabledInTree(numberInput)) return; /* 入力欄か親が無効の間は動かさない */
            var value = parseFloat(numberInput.text);
            if (isNaN(value)) value = 0;
            writeSteppedValue(numberInput, computeSteppedValue(value, direction, stepOptions), stepOptions);
            if (stepOptions.onStep) stepOptions.onStep(numberInput);
        }

        /* 整数の欄では option＋クリックの0.1刻みが効かないので、説明から外す / integer fields have no 0.1 step */
        var upTooltip = stepOptions.integer ? LABELS.tooltip.stepUpInteger : LABELS.tooltip.stepUp;
        var downTooltip = stepOptions.integer ? LABELS.tooltip.stepDownInteger : LABELS.tooltip.stepDown;
        makeStepperChevronButton(stepperGroup, "up", function () { stepBy(1); }).helpTip = getLabel(upTooltip);
        makeStepperChevronButton(stepperGroup, "down", function () { stepBy(-1); }).helpTip = getLabel(downTooltip);
        stepperGroup.stepBy = stepBy; /* ↑↓キーからも同じ処理で増減できるよう公開 / shared with the arrow keys */
        return stepperGroup;
    }

    /**
     * 入力欄の↑↓キーを、∧∨と同じ処理で増減させる。ほかのキーは素通し
     * @param {EditText} numberInput - 対象の入力欄
     * @param {Group} stepperGroup - addStepper() で作った∧∨
     * @returns {void}
     */
    function bindSteppedArrowKeys(numberInput, stepperGroup) {
        numberInput.addEventListener("keydown", function (event) {
            if (event.keyName !== "Up" && event.keyName !== "Down") return;
            stepperGroup.stepBy(event.keyName === "Up" ? 1 : -1);
            event.preventDefault(); /* カーソル移動を止める / keep the caret from moving */
        });
    }

    // -----------------------------------------
    // 値の計算 / Value helpers
    // -----------------------------------------
    /**
     * 押された修飾キーに応じて、1回分増減した値を返す
     * （shift なら STEPPER_SHIFT_MULTIPLE の倍数へ、option なら STEPPER_OPTION_STEP ずつ、それ以外は step の倍数へ（1.5→2、1.5→1）。
     * 整数の欄では option を無視して step の倍数へ）
     * @param {number} value - 元の値
     * @param {number} direction - 増やすなら 1、減らすなら -1
     * @param {Object} stepOptions - step（通常の増減量。省略時は 1）/ integer
     * @returns {number} 増減した値（下限・上限は未適用）
     */
    function computeSteppedValue(value, direction, stepOptions) {
        var keyState = ScriptUI.environment.keyboardState;
        if (keyState.shiftKey) return snapStepperToNextMultiple(value, STEPPER_SHIFT_MULTIPLE, direction);
        if (keyState.altKey && !stepOptions.integer) return value + direction * STEPPER_OPTION_STEP;
        return snapStepperToNextMultiple(value, stepOptions.step || 1, direction);
    }

    /**
     * 値を、指定した方向にある次の倍数へ移す（230→240、232→240、下げるときは 232→230、230→220）
     * @param {number} value - 元の値
     * @param {number} multiple - 倍数の単位（例 10）
     * @param {number} direction - 上げるなら 1、下げるなら -1
     * @returns {number} 移した値
     */
    function snapStepperToNextMultiple(value, multiple, direction) {
        /* 0.29 / 0.01 = 28.999… のような浮動小数の誤差で同じ値に戻らないよう、商を丸めてから切り捨て・切り上げる
           round the quotient first so float error (0.29 / 0.01 = 28.999…) does not step back to the same value */
        var quotient = Math.round(value / multiple * 1e6) / 1e6;
        if (direction > 0) return Math.round((Math.floor(quotient) + 1) * multiple * 1e6) / 1e6;
        return Math.round((Math.ceil(quotient) - 1) * multiple * 1e6) / 1e6;
    }

    /**
     * 値を下限・上限の範囲に収める
     * @param {number} value - 数値
     * @param {Object} rangeOptions - min / max（どちらも省略可）
     * @returns {number} 範囲に収めた値
     */
    function clampSteppedValue(value, rangeOptions) {
        if (rangeOptions.min !== undefined && value < rangeOptions.min) return rangeOptions.min;
        if (rangeOptions.max !== undefined && value > rangeOptions.max) return rangeOptions.max;
        return value;
    }

    /**
     * 値を整数化・下限・上限でそろえ、単位を付けて入力欄に書き込む（直前の正しい値としても控える）
     * @param {EditText} numberInput - 書き込む入力欄
     * @param {number} value - 数値
     * @param {Object} valueOptions - integer / min / max / unit（どれも省略可）
     * @returns {void}
     */
    function writeSteppedValue(numberInput, value, valueOptions) {
        numberInput.text = formatSteppedValue(value, valueOptions);
        numberInput.lastValidText = numberInput.text;
    }

    /**
     * 値を整数化・下限・上限でそろえ、丸めて単位を付けた表示用の文字列にする。
     * 整数化してから下限で止めるので、「整数・下限1」の欄に 0.4 が入っても 1 になる
     * @param {number} value - 数値
     * @param {Object} valueOptions - integer / min / max / unit（どれも省略可）
     * @returns {string} 入力欄に入れる文字列（例 "20 mm"）
     */
    function formatSteppedValue(value, valueOptions) {
        if (valueOptions.integer) value = Math.round(value);
        return formatStepperNumber(clampSteppedValue(value, valueOptions)) + (valueOptions.unit || "");
    }

    /**
     * 小数第2位で丸めた数値を文字列で返す
     * @param {number} value - 数値
     * @returns {string} 表示用の数値文字列
     */
    function formatStepperNumber(value) {
        return String(Math.round(value * 100) / 100);
    }

    // -----------------------------------------
    // ∧∨ボタンの描画 / Drawing
    // -----------------------------------------
    /**
     * 山形（∧／∨）の極小ボタンを作成する。
     * 上下2つを隙間なく積んで1つの枠に見えるよう、枠線は外側の辺だけ描き（上ボタンは上側、下ボタンは下側）、
     * 継ぎ目に線は引かない
     * @param {Group|Panel} parent - 追加先
     * @param {string} direction - "up" または "down"
     * @param {Function} onClickFn - クリック時の処理
     * @returns {Group} ボタンとして使う group
     */
    function makeStepperChevronButton(parent, direction, onClickFn) {
        var buttonWidth = STEPPER_BUTTON_WIDTH;
        var buttonHeight = STEPPER_BUTTON_HEIGHT;
        var isUp = (direction === "up");
        var chevronBox = parent.add("group");
        chevronBox.margins = 0;
        chevronBox.spacing = 0;
        chevronBox.preferredSize = [buttonWidth, buttonHeight];
        chevronBox.minimumSize = [buttonWidth, buttonHeight];
        chevronBox.maximumSize = [buttonWidth, buttonHeight];
        chevronBox.isPressed = false;
        chevronBox.isStepperButton = true; /* redrawSteppersIn() の目印 / marker for redrawSteppersIn() */

        chevronBox.onDraw = function () {
            var boxGraphics = chevronBox.graphics;
            /* 自作描画は自動でディムにならないため、無効なら薄い色で描く。親の無効化は子の enabled に出ないので親も見る
               Custom drawing is not dimmed automatically; the parent's state does not reach the child's enabled */
            var isDimmed = !isStepperEnabledInTree(chevronBox);

            /* 枠線の内側の地（押下中は押下色） / background inside the frame, pressed color while pressed */
            var fillColor = isDimmed ? STEPPER_DIM_FILL_COLOR : (chevronBox.isPressed ? STEPPER_PRESSED_COLOR : STEPPER_FILL_COLOR);
            boxGraphics.newPath();
            boxGraphics.rectPath(1, isUp ? 1 : 0, buttonWidth - 2, buttonHeight - 1);
            boxGraphics.fillPath(boxGraphics.newBrush(boxGraphics.BrushType.SOLID_COLOR, fillColor));

            drawStepperFrame(boxGraphics, buttonWidth, buttonHeight, isUp, isDimmed ? STEPPER_DIM_FRAME_COLOR : STEPPER_FRAME_COLOR);
            drawStepperChevron(boxGraphics, buttonWidth, buttonHeight, isUp, isDimmed ? STEPPER_DIM_CHEVRON_COLOR : STEPPER_CHEVRON_COLOR);
        };

        /**
         * 押下状態を変えて描き直す
         * @param {boolean} isPressed - 押下中なら true
         * @returns {void}
         */
        function repaint(isPressed) {
            if (chevronBox.isPressed === isPressed) return;
            chevronBox.isPressed = isPressed;
            redrawStepperGroup(chevronBox);
        }
        chevronBox.addEventListener("mousedown", function () {
            if (!isStepperEnabledInTree(chevronBox)) return;
            repaint(true);
            if (onClickFn) onClickFn();
        });
        chevronBox.addEventListener("mouseup", function () { repaint(false); });
        /* 押したまま外へ出たときも押下色を残さない / reset when the pointer leaves while pressed */
        chevronBox.addEventListener("mouseout", function () { repaint(false); });
        return chevronBox;
    }

    /**
     * 外側の辺だけの枠を描く（角は丸める）。継ぎ目側は開けておき、上下2つで1つの枠に見せる。
     * ScriptUI は円弧を描けないため、角丸は短い線分で近似する
     * @param {ScriptUIGraphics} boxGraphics - 描画先
     * @param {number} boxWidth - ボタンの幅
     * @param {number} boxHeight - ボタンの高さ
     * @param {boolean} isUp - 上のボタンなら true（上側に枠を描く）
     * @param {number[]} frameColor - [r, g, b, a]
     * @returns {void}
     */
    function drawStepperFrame(boxGraphics, boxWidth, boxHeight, isUp, frameColor) {
        var frameLeft = 0.5;
        var frameRight = boxWidth - 0.5;
        var outerY = isUp ? 0.5 : boxHeight - 0.5;
        var seamY = isUp ? boxHeight : 0;
        var towardSeam = isUp ? 1 : -1; /* 外側の辺から継ぎ目へ向かう向き / direction from the outer edge to the seam */
        var radius = STEPPER_CORNER_RADIUS;
        var arcSteps = 4; /* 角丸1つを何本の線分で近似するか / segments per corner */
        var angle, k;

        boxGraphics.newPath();
        boxGraphics.moveTo(frameLeft, seamY);
        /* 左の角丸 / left corner */
        for (k = 0; k <= arcSteps; k++) {
            angle = (Math.PI / 2) * k / arcSteps;
            boxGraphics.lineTo(frameLeft + radius - radius * Math.cos(angle), outerY + towardSeam * (radius - radius * Math.sin(angle)));
        }
        /* 右の角丸 / right corner */
        for (k = 0; k <= arcSteps; k++) {
            angle = (Math.PI / 2) * k / arcSteps;
            boxGraphics.lineTo(frameRight - radius + radius * Math.sin(angle), outerY + towardSeam * (radius - radius * Math.cos(angle)));
        }
        boxGraphics.lineTo(frameRight, seamY);
        boxGraphics.strokePath(boxGraphics.newPen(boxGraphics.PenType.SOLID_COLOR, frameColor, 1));
    }

    /**
     * 山形（∧／∨）を描く。文字グリフの▲▼は上下で大きさやベースラインが揃わないため、線で描く
     * @param {ScriptUIGraphics} boxGraphics - 描画先
     * @param {number} boxWidth - ボタンの幅
     * @param {number} boxHeight - ボタンの高さ
     * @param {boolean} isUp - ∧なら true、∨なら false
     * @param {number[]} chevronColor - [r, g, b, a]
     * @returns {void}
     */
    function drawStepperChevron(boxGraphics, boxWidth, boxHeight, isUp, chevronColor) {
        var centerX = boxWidth / 2;
        var centerY = isUp ? boxHeight / 2 + 0.5 : boxHeight / 2 - 0.5; /* 継ぎ目から少し離す / nudged away from the seam */
        var halfWidth = 3.6; /* 山形の半幅（高さ1.8に対して開き約127°） / half width of the chevron */
        var tipOffsetY = isUp ? -1.8 : 1.8; /* 頂点の中心からのずれ（上向きは上、下向きは下） */
        boxGraphics.newPath();
        boxGraphics.moveTo(centerX - halfWidth, centerY - tipOffsetY);
        boxGraphics.lineTo(centerX, centerY + tipOffsetY);
        boxGraphics.lineTo(centerX + halfWidth, centerY - tipOffsetY);
        boxGraphics.strokePath(boxGraphics.newPen(boxGraphics.PenType.SOLID_COLOR, chevronColor, 1.2));
    }

    /**
     * コントロールと、その親をたどってすべて有効かを返す（親の無効化は子の enabled に出ない）
     * @param {Object} control - 対象のコントロール
     * @returns {boolean} すべて有効なら true
     */
    function isStepperEnabledInTree(control) {
        for (var node = control; node; node = node.parent) {
            if (!node.enabled) return false;
        }
        return true;
    }

    /**
     * コンテナ以下にある∧∨ボタンをすべて描き直す。行やパネルの enabled を切り替えたあとに呼ぶ
     * @param {Object} container - 行・グループ・パネルなど
     * @returns {void}
     */
    function redrawSteppersIn(container) {
        if (!container.children) return;
        for (var i = 0; i < container.children.length; i++) {
            var child = container.children[i];
            if (child.isStepperButton) redrawStepperGroup(child);
            else redrawSteppersIn(child);
        }
    }

    /**
     * group の onDraw を呼び直す。group には notify() が無いため、隠して再表示して描き直させる
     * @param {Group} targetGroup - 描き直す group
     * @returns {void}
     */
    function redrawStepperGroup(targetGroup) {
        targetGroup.hide();
        targetGroup.show();
    }

    // ステップボタン（再利用パーツ）ここまで / End of the reusable stepper

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
    // プレビュー / Preview
    // =========================================

    /* プレビュー管理（実行した手数を数え、undoで巻き戻す）/ Preview manager that rolls back by counting undo steps */
    function PreviewManager() {
        this.undoDepth = 0;

        /* 変更処理を実行して1手として数える（false を返した場合は数えない）
           / Run a change and count one undo step; a callback returning false is not counted
           @param {function} changeAction - 実行する処理
           @returns {void} */
        this.addStep = function (changeAction) {
            try {
                if (changeAction() === false) return;
                this.undoDepth++;
                app.redraw();
            } catch (e) {
                $.writeln("[PreviewManager] addStep error: " + e);
            }
        };

        /* プレビューによる変更をすべて取り消す / Undo every preview change
           @returns {void} */
        this.rollback = function () {
            while (this.undoDepth > 0) {
                app.undo();
                this.undoDepth--;
            }
            app.redraw();
        };

        /* 巻き戻してから本処理を一度だけ実行する / Roll back, then run the final action once
           @param {function} confirmAction - 巻き戻したあとに実行する処理
           @returns {void} */
        this.confirm = function (confirmAction) {
            this.rollback();
            confirmAction();
            this.undoDepth = 0;
        };
    }

    // =========================================
    // アクションの生成と実行 / Build and run the action
    // =========================================

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

    /* 線の位置のパラメーターブロックを組み立てる / Build the alignment parameter block
       @param {number} alignmentValue - 線の位置（0=中央 / 1=内側 / 2=外側）
       @returns {Array<string>} 行の配列 */
    function buildAlignmentParameterLines(alignmentValue) {
        return [
            "\t\t/parameter-1 {",
            "\t\t\t/key " + ACTION_KEY_STROKE_ALIGNMENT,
            "\t\t\t/showInPalette 4294967295",
            "\t\t\t/type (enumerated)"
        ]
            .concat(buildActionNameLines("\t\t\t", ACTION_PARAM_NAME))
            .concat([
                "\t\t\t/value " + alignmentValue,
                "\t\t}"
            ]);
    }

    /* 線の位置を設定する一時アクションのコードを組み立てる / Build the .aia source that sets the stroke alignment
       @param {number} alignmentValue - 線の位置（0=中央 / 1=内側 / 2=外側）
       @returns {string} .aia のソース */
    function buildStrokeActionCode(alignmentValue) {
        return ["/version 3"]
            .concat(buildActionNameLines("", ACTION_SET_NAME))
            .concat([
                "/isOpen 1",
                "/actionCount 1",
                "/action-1 {"
            ])
            .concat(buildActionNameLines("\t", ACTION_NAME))
            .concat([
                "\t/keyIndex 0",
                "\t/colorIndex 0",
                "\t/isOpen 1",
                "\t/eventCount 1",
                "\t/event-1 {",
                "\t\t/useRulersIn1stQuadrant 0",
                "\t\t/internalName (ai_plugin_setStroke)"
            ])
            .concat(buildActionNameLines("\t\t", ACTION_NAME, "localizedName"))
            .concat([
                "\t\t/isOpen 1",
                "\t\t/isOn 1",
                "\t\t/hasDialog 0",
                "\t\t/parameterCount 1"
            ])
            .concat(buildAlignmentParameterLines(alignmentValue))
            .concat(["\t}", "}"])
            .join("\n");
    }

    /* 一時アクションを読み込み、選択オブジェクトに線の位置を適用する / Apply the stroke alignment via a temporary action
       @param {number} alignmentValue - 線の位置（0=中央 / 1=内側 / 2=外側）
       @returns {void} */
    function applyStrokeByAction(alignmentValue) {
        /* 失敗は従来どおり例外で伝える / Report a failure as an exception, as before */
        if (!runTemporaryAction(buildStrokeActionCode(alignmentValue), ACTION_SET_NAME, ACTION_NAME)) {
            throw new Error("Could not run the action: " + ACTION_SET_NAME + " / " + ACTION_NAME);
        }
    }

    // =========================================
    // 線の適用 / Apply strokes
    // =========================================

    /* 線端・角の形状の識別子 → DOMの列挙値 / option keys to DOM enums */
    var STROKE_CAP_ENUMS = {
        butt: StrokeCap.BUTTENDCAP,
        round: StrokeCap.ROUNDENDCAP,
        projecting: StrokeCap.PROJECTINGENDCAP
    };
    var CORNER_JOIN_ENUMS = {
        miter: StrokeJoin.MITERENDJOIN,
        round: StrokeJoin.ROUNDENDJOIN,
        bevel: StrokeJoin.BEVELENDJOIN
    };

    /* 黒のカラーを取得（スウォッチが無い場合はカラーモードから作る）/ Get black, falling back to the document color space
       @param {Document} targetDocument - 対象ドキュメント
       @returns {Color} 黒のカラー */
    function getBlackColor(targetDocument) {
        try {
            return targetDocument.swatches["[Black]"].color;
        } catch (e) {
            if (targetDocument.documentColorSpace === DocumentColorSpace.RGB) {
                var rgbBlack = new RGBColor();
                rgbBlack.red = 0;
                rgbBlack.green = 0;
                rgbBlack.blue = 0;
                return rgbBlack;
            }
            var cmykBlack = new CMYKColor();
            cmykBlack.cyan = 0;
            cmykBlack.magenta = 0;
            cmykBlack.yellow = 0;
            cmykBlack.black = 100;
            return cmykBlack;
        }
    }

    /* 選択アイテムを再帰的にたどり、線を持てるものに処理を適用する / Walk the selection, applying a callback to every item that can take a stroke
       @param {Array} targetItems - 対象アイテムの配列
       @param {function} applyToItem - 1アイテムに対する処理（false を返すと数えない）
       @returns {number} 処理したアイテム数 */
    function forEachStrokableItem(targetItems, applyToItem) {
        var appliedCount = 0;
        for (var i = 0; i < targetItems.length; i++) {
            var targetItem = targetItems[i];
            if (targetItem.typename === "GroupItem") {
                appliedCount += forEachStrokableItem(targetItem.pageItems, applyToItem);
                continue;
            }
            if (targetItem.typename === "CompoundPathItem") {
                appliedCount += forEachStrokableItem(targetItem.pathItems, applyToItem);
                continue;
            }
            try {
                if (applyToItem(targetItem) !== false) appliedCount++;
            } catch (e) {
                /* 線を持てないオブジェクトはスキップ / skip items that cannot take a stroke */
            }
        }
        return appliedCount;
    }

    /* 閉じたパスが含まれるかを調べる（グループ・複合パスは再帰）/ Does the selection contain a closed path?
       @param {Array} targetItems - 対象アイテムの配列
       @returns {boolean} 閉じたパスがあれば true */
    function hasClosedPath(targetItems) {
        for (var i = 0; i < targetItems.length; i++) {
            var targetItem = targetItems[i];
            if (targetItem.typename === "GroupItem") {
                if (hasClosedPath(targetItem.pageItems)) return true;
                continue;
            }
            if (targetItem.typename === "CompoundPathItem") {
                if (hasClosedPath(targetItem.pathItems)) return true;
                continue;
            }
            if (targetItem.typename === "PathItem" && targetItem.closed) return true;
        }
        return false;
    }

    /* 線が無いアイテムに黒の線を付ける / Give unstroked items a black stroke
       @param {Array} targetItems - 対象アイテムの配列
       @param {Color} blackColor - 適用する黒
       @returns {number} 線を付けたアイテム数 */
    function ensureStrokeColor(targetItems, blackColor) {
        return forEachStrokableItem(targetItems, function (targetItem) {
            /* すでに線があれば色は触らない / existing stroke colors stay */
            if (targetItem.stroked && targetItem.strokeColor.typename !== "NoColor") return false;
            targetItem.stroked = true;
            targetItem.strokeColor = blackColor;
        });
    }

    /* 線幅・線端・角の形状をDOMで設定する / Set weight, cap and join through the DOM
       @param {Array} targetItems - 対象アイテムの配列
       @param {object} strokeSettings - { strokeWidth, strokeCap, cornerJoin }
       @returns {number} 設定できたアイテム数 */
    function applyStrokeProperties(targetItems, strokeSettings) {
        return forEachStrokableItem(targetItems, function (targetItem) {
            targetItem.strokeWidth = strokeSettings.strokeWidth;
            targetItem.strokeCap = STROKE_CAP_ENUMS[strokeSettings.strokeCap];
            targetItem.strokeJoin = CORNER_JOIN_ENUMS[strokeSettings.cornerJoin];
        });
    }

    /* 線の設定を適用する / Apply the stroke settings
       アクションはパラメーターに書いていない項目（線幅・線端・角の形状）を既定値に戻すため、
       線の位置をアクションで決めたあとにDOMで上書きする。
       / the action resets the settings it does not carry, so the DOM pass has to run after it
       @param {Array} targetItems - 対象アイテムの配列
       @param {Color} blackColor - 線が無いアイテムに付ける黒
       @param {object} strokeSettings - { strokeWidth, strokeCap, cornerJoin, strokeAlignment }
       @param {PreviewManager} [previewManager] - 指定するとプレビューの手数として記録する
       @returns {void} */
    function applyStroke(targetItems, blackColor, strokeSettings, previewManager) {
        /* strokeAlignment が null＝線の位置を使えない選択なので、アクションは流さない
           / a null alignment means the selection cannot take one, so the action is skipped */
        var usesAlignment = (strokeSettings.strokeAlignment != null);
        var alignmentValue = usesAlignment ? STROKE_ALIGNMENT_VALUES[strokeSettings.strokeAlignment] : 0;

        if (!previewManager) {
            ensureStrokeColor(targetItems, blackColor);
            app.redraw();
            if (usesAlignment) applyStrokeByAction(alignmentValue);
            applyStrokeProperties(targetItems, strokeSettings);
            return;
        }
        /* 何も変えなかったパスは1手として数えない / a pass that changed nothing is not counted */
        previewManager.addStep(function () {
            return ensureStrokeColor(targetItems, blackColor) > 0;
        });
        if (usesAlignment) {
            previewManager.addStep(function () {
                applyStrokeByAction(alignmentValue);
            });
        }
        previewManager.addStep(function () {
            return applyStrokeProperties(targetItems, strokeSettings) > 0;
        });
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /* 線の設定ダイアログを表示し、OKなら適用する / Show the stroke dialog and apply on OK
       @param {Array} targetItems - 対象アイテムの配列
       @param {Color} blackColor - 線が無いアイテムに付ける黒
       @returns {object} 適用した設定（キャンセル時は null） */
    function showStrokeDialog(targetItems, blackColor) {
        var dialog = new Window("dialog", getLabel(LABELS.dialog.title) + " " + SCRIPT_VERSION);
        setupWindow(dialog);

        /* 線幅 / Stroke weight */
        var strokeWidthGroup = addControlRow(dialog, LABELS.fieldLabel.strokeWidth);
        /* ∧∨と入力欄は隙間0で突き合わせる。線幅は0より大きい値だけ有効なので、下限は0.1にする
           Butt the stepper against the field; the weight must be above 0, so the minimum is 0.1 */
        var strokeWidthStepperGroup = strokeWidthGroup.add("group");
        strokeWidthStepperGroup.orientation = "row";
        strokeWidthStepperGroup.alignChildren = ["left", "center"];
        strokeWidthStepperGroup.spacing = 0;
        strokeWidthStepperGroup.margins = 0;
        var strokeWidthInput;
        var strokeWidthStepper = addStepper(strokeWidthStepperGroup, function () { return strokeWidthInput; }, {
            min: 0.1,
            onStep: function () { updatePreview(); }
        });
        strokeWidthInput = strokeWidthStepperGroup.add("edittext", undefined, CONFIG.defaultStrokeWidth);
        strokeWidthInput.characters = 5;
        /* ↑↓キーも∧∨と同じ処理で増減する / arrow keys share the stepper's logic */
        bindSteppedArrowKeys(strokeWidthInput, strokeWidthStepper);
        strokeWidthGroup.add("statictext", undefined, "pt");

        /* 線端・角の形状・線の位置 / Cap, join, alignment */
        var strokeCapGroup = addRadioRow(dialog, LABELS.fieldLabel.strokeCap,
            STROKE_CAP_OPTIONS, CONFIG.defaultStrokeCap, function () { updatePreview(); });
        var cornerJoinGroup = addRadioRow(dialog, LABELS.fieldLabel.cornerJoin,
            CORNER_JOIN_OPTIONS, CONFIG.defaultCornerJoin, function () { updatePreview(); });
        var strokeAlignmentGroup = addRadioRow(dialog, LABELS.fieldLabel.strokeAlignment,
            STROKE_ALIGNMENT_OPTIONS, CONFIG.defaultStrokeAlignment, function () { updatePreview(); });

        /* 線の位置は閉じたパスにしか効かないので、無ければ行ごとディムにする（ラベルも含めるため親の行を無効化）
           / alignment only works on closed paths; disable the whole row (parent, so the label dims too) */
        var strokeAlignmentRow = strokeAlignmentGroup.parent;
        strokeAlignmentRow.enabled = hasClosedPath(targetItems);

        /* ボタンエリア（左：プレビュー／右：キャンセル・OK）/ Button area */
        var buttonRow = addButtonRow(dialog);
        var previewCheckbox = buttonRow.leftGroup.add("checkbox", undefined, getLabel(LABELS.checkbox.preview));
        previewCheckbox.value = CONFIG.previewDefault;
        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });
        var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.ok), { name: "ok" });

        /* 入力値を読み取る（不正なら null）/ Read the current input, null when invalid
           @returns {object} 線の設定 */
        function readStrokeSettings() {
            var strokeWidth = parseFloat(strokeWidthInput.text);
            if (isNaN(strokeWidth) || strokeWidth <= 0) return null;
            return {
                strokeWidth: strokeWidth,
                strokeCap: strokeCapGroup.selectedOptionKey,
                cornerJoin: cornerJoinGroup.selectedOptionKey,
                strokeAlignment: strokeAlignmentRow.enabled ? strokeAlignmentGroup.selectedOptionKey : null
            };
        }

        var previewManager = new PreviewManager();

        /* Undo履歴を汚さずにプレビューを更新する / Refresh the preview without polluting the undo history
           @returns {void} */
        function updatePreview() {
            previewManager.rollback();
            if (!previewCheckbox.value) return;
            var strokeSettings = readStrokeSettings();
            if (!strokeSettings) return; /* 入力途中はプレビューしない / skip while the input is incomplete */
            applyStroke(targetItems, blackColor, strokeSettings, previewManager);
        }

        strokeWidthInput.onChanging = updatePreview;
        previewCheckbox.onClick = updatePreview;

        btnOK.onClick = function () {
            if (!readStrokeSettings()) {
                alert(getLabel(LABELS.alert.invalidWidth));
                return;
            }
            dialog.close(1);
        };
        btnCancel.onClick = function () {
            dialog.close(0);
        };

        updatePreview();
        strokeWidthInput.active = true;

        alignRightOnlyButtonRow(buttonRow);
        prepareDialogWindow(dialog, SCRIPT_NAME);
        if (dialog.show() !== 1) {
            /* キャンセル：プレビューを巻き戻す / Cancel: roll back the preview */
            previewManager.rollback();
            return null;
        }

        /* プレビューを巻き戻してから、確定として一度だけ適用する / Roll back the preview, then apply once */
        var confirmedSettings = readStrokeSettings();
        previewManager.confirm(function () {
            applyStroke(targetItems, blackColor, confirmedSettings);
            app.redraw();
        });
        return confirmedSettings;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /* 選択オブジェクトに線の設定を適用する / Apply the stroke settings to the selection
       @returns {void} */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel(LABELS.alert.noDocument));
            return;
        }
        var documentRef = app.activeDocument;
        var targetItems = documentRef.selection;
        /* 文字ツールでの文字選択（TextRange）は対象外 / a type-tool selection is not a page-item array */
        if (!(targetItems instanceof Array) || targetItems.length === 0) {
            alert(getLabel(LABELS.alert.noSelection));
            return;
        }

        showStrokeDialog(targetItems, getBlackColor(documentRef));
    }

    main();

})();
