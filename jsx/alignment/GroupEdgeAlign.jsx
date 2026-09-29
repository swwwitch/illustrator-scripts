#target illustrator
#targetengine "GroupEdgeAlignEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択オブジェクトの端または中心を、アクティブアートボードの端・中央、または条件に合うガイドへ整列します。
整列先は3×3の9点から選べるほか、矢印キーでの1段階ずつの送りやファイル名からの自動判定にも対応します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/GroupEdgeAlign.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n4ae0e1e70481

### Overview

Aligns the edges or the center of the selected objects to the edge or the center of the active artboard, or to a matching guide.
The target is picked from a 3x3 grid of nine points, stepped one target at a time with the arrow keys, or derived from the filename.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/GroupEdgeAlign.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "GroupEdgeAlign";               /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.5";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-04-06";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-29";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/GroupEdgeAlign.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/GroupEdgeAlign.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n4ae0e1e70481"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    var SHOW_DIALOG                 = true;     /* ダイアログを表示する / show the dialog */
    var USE_GUIDES                  = true;     /* ガイドも整列先に含める / include guides as targets */
    var DEFAULT_ALIGNMENT_SIDE      = "right";  /* ファイル名から判定できないときの整列先 / fallback target */
    var GUIDE_SEARCH_MODE           = "inside"; /* "inside"＝進行方向の直近 / "nearest"＝最も近い */
    var GUIDE_ORIENTATION_TOLERANCE = 0.01;     /* ガイドの水平・垂直判定に使う許容値 / orientation tolerance */

    // =========================================
    // レイアウト / Layout
    // =========================================

    /* ウィンドウ・パネルの余白と間隔 / Window & panel margins and spacing */
    var WINDOW_MARGINS        = 16;               /* ウィンドウ外周の余白 / window margin */
    var WINDOW_SPACING        = 12;               /* ウィンドウ内の要素間隔 / window spacing */
    var PANEL_MARGINS         = [16, 20, 16, 12]; /* パネル余白 [左,上,右,下] / panel margins */
    var PANEL_SPACING         = 6;                /* パネル内の要素間隔 / panel spacing */
    var ANCHOR_PANEL_MARGINS  = [9, 13, 9, 4];    /* 9軸ウィジェット用の詰めた余白 / tighter padding for the anchor widget */

    // =========================================
    // 整列先の定義 / Alignment targets
    // =========================================

    /* 3×3の各セル（行優先）が担う水平・垂直の整列先とショートカットキー。
       各要素はそのまま整列軸（{ horizontal, vertical }）として使う
       Horizontal/vertical targets and shortcut key per cell of the 3x3 grid (row-major);
       each entry doubles as the { horizontal, vertical } pair used for aligning */
    var ANCHOR_DEFINITIONS = [
        { horizontal: "left",     vertical: "top",      shortcutKey: "W" },
        { horizontal: "CENTER_X", vertical: "top",      shortcutKey: "E" },
        { horizontal: "right",    vertical: "top",      shortcutKey: "R" },
        { horizontal: "left",     vertical: "CENTER_Y", shortcutKey: "S" },
        { horizontal: "CENTER_X", vertical: "CENTER_Y", shortcutKey: "D" },
        { horizontal: "right",    vertical: "CENTER_Y", shortcutKey: "F" },
        { horizontal: "left",     vertical: "bottom",   shortcutKey: "X" },
        { horizontal: "CENTER_X", vertical: "bottom",   shortcutKey: "C" },
        { horizontal: "right",    vertical: "bottom",   shortcutKey: "V" }
    ];

    /* 9軸ウィジェットが未選択のときのインデックス / Index used while no cell is selected */
    var NO_ANCHOR_INDEX = -1;

    /* 矢印キーの向きと、1段階の整列で使う整列先 / Arrow keys mapped to the target of one step */
    var STEP_SIDE_BY_KEY = { "Up": "top", "Down": "bottom", "Left": "left", "Right": "right" };

    /* 整列先ごとの座標の取り出し方。
       axis＝動かす軸、boundsIndex＝境界配列 [左,上,右,下] のインデックス（null は2辺の中点）、
       ahead＝揃える向き（座標が増える向きなら +1、中央揃えは 0＝ガイド吸着の対象外）
       axis is the axis to move, boundsIndex indexes [L,T,R,B] (null means the midpoint of two edges),
       ahead is the sign of the direction aligned toward (0 disables guide snapping) */
    var EDGE_RULES = {
        "left":     { axis: "x", boundsIndex: 0,    ahead: -1 },
        "right":    { axis: "x", boundsIndex: 2,    ahead:  1 },
        "top":      { axis: "y", boundsIndex: 1,    ahead:  1 },
        "bottom":   { axis: "y", boundsIndex: 3,    ahead: -1 },
        "CENTER_X": { axis: "x", boundsIndex: null, ahead:  0 },
        "CENTER_Y": { axis: "y", boundsIndex: null, ahead:  0 }
    };

    /* 整列先1つを水平・垂直の軸に展開した対応表。null はその軸を動かさない
       Each target expanded into horizontal/vertical axes; null leaves that axis alone */
    var AXES_BY_ALIGNMENT_SIDE = {
        "left":     { horizontal: "left",     vertical: null },
        "right":    { horizontal: "right",    vertical: null },
        "top":      { horizontal: null,       vertical: "top" },
        "bottom":   { horizontal: null,       vertical: "bottom" },
        "CENTER_X": { horizontal: "CENTER_X", vertical: null },
        "CENTER_Y": { horizontal: null,       vertical: "CENTER_Y" },
        "CENTER":   { horizontal: "CENTER_X", vertical: "CENTER_Y" }
    };

    /* ファイル名に含まれるキーワードと整列先。CENTERX / CENTERY を CENTER より先に判定する
       Filename keywords mapped to targets; CENTERX/CENTERY are matched before CENTER */
    var ALIGNMENT_SIDE_BY_FILENAME = [
        { keyword: "CENTERX", side: "CENTER_X" },
        { keyword: "CENTERY", side: "CENTER_Y" },
        { keyword: "CENTER",  side: "CENTER" },
        { keyword: "LEFT",    side: "left" },
        { keyword: "RIGHT",   side: "right" },
        { keyword: "TOP",     side: "top" },
        { keyword: "BOTTOM",  side: "bottom" }
    ];

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

    /* カテゴリ分けした日英ラベル定義 / Categorized Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "アートボードに整列", en: "Align to Artboard" }
        },
        panel: {
            alignment: { ja: "整列先", en: "Align To" },
            option:    { ja: "オプション", en: "Options" }
        },
        checkbox: {
            previewBounds: { ja: "プレビュー境界を使用", en: "Use preview bounds" },
            useGuides:     { ja: "ガイドを使用", en: "Use guides" },
            preview:       { ja: "プレビュー", en: "Preview" }
        },
        button: {
            ok:     { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        tooltip: {
            anchorWidget: {
                ja: "アートボードのどこに揃えるかを3×3のマスで選びます。\nW E R／S D F／X C V キーでも選べます。\n矢印キーを押すと、その向きへ1段階ずつ整列します。",
                en: "Pick where on the artboard to align using the 3x3 grid.\nThe keys W E R / S D F / X C V select a cell too.\nAn arrow key aligns one step in that direction."
            },
            previewBounds: {
                ja: "線幅や効果を含めた見た目の端を基準に整列します（B キーで切り替え）。",
                en: "Aligns by the visible edges including strokes and effects (B toggles it)."
            },
            useGuides: {
                ja: "アートボードの端より手前にガイドがあれば、そのガイドに揃えます（G キーで切り替え）。\n3×3で整列先を選んでいる間は使えません。",
                en: "Snaps to a guide when one sits before the artboard edge (G toggles it).\nUnavailable while a cell of the 3x3 grid is selected."
            },
            preview: {
                ja: "結果を画面で確認します。キャンセルすると元の位置に戻ります。",
                en: "Shows the result on the canvas. Cancel restores the original positions."
            }
        },
        alert: {
            noDocument:             { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noSelection:            { ja: "オブジェクトが選択されていません。", en: "No objects are selected." },
            invalidGuideSearchMode: { ja: "GUIDE_SEARCH_MODE の指定が不正です", en: "Invalid GUIDE_SEARCH_MODE" }
        }
    };

    // =========================================
    // UIレイアウト補助 / UI layout helpers
    // =========================================

    /**
     * パネルに共通レイアウトを適用する
     * @param {Panel} targetPanel - 対象パネル
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
     * グループを横並びの行として設定する
     * @param {Group} targetGroup - 対象グループ
     * @param {string} [horizontalAlign] - 横方向の揃え（省略時は "left"）
     * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function setupRow(targetGroup, horizontalAlign, spacing) {
        targetGroup.orientation = "row";
        /* 揃えは横と天地を対で指定し、親の fill 継承を打ち消す / Pair both axes to cancel the parent's fill */
        targetGroup.alignment = [horizontalAlign || "left", "center"];
        targetGroup.alignChildren = ["left", "center"];
        targetGroup.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /**
     * ラベル付きパネルを生成する（共通レイアウト適用）
     * @param {Window|Group} parentContainer - 追加先
     * @param {string} panelTitle - パネルの見出し
     * @returns {Panel} 生成したパネル
     */
    function addPanel(parentContainer, panelTitle) {
        var createdPanel = parentContainer.add("panel", undefined, panelTitle);
        setupPanel(createdPanel);
        return createdPanel;
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

    // 基準点ウィジェット（再利用パーツ） / Anchor widget (reusable)

    // -----------------------------------------
    // 基準点ウィジェットの寸法 / Anchor widget metrics
    // -----------------------------------------
    var ANCHOR_WIDGET_SIZE      = 66;   /* ウィジェット全体の一辺 / overall size of the widget */
    var ANCHOR_WIDGET_CELL_SIZE = 9;    /* □1個の一辺 / size of one square */
    var ANCHOR_WIDGET_CELL_GAP  = 7.5;  /* □どうしの間隔 / gap between squares */
    var ANCHOR_WIDGET_NONE      = -1;   /* 未選択のインデックス / index while nothing is selected */

    /* セルの名前（行優先：上 → 中 → 下、列：左 → 中 → 右）。Transformation の列挙名にそろえる
       Cell names in row-major order, matching the Transformation enumeration */
    var ANCHOR_WIDGET_NAMES = ["topLeft", "top", "topRight", "left", "center", "right", "bottomLeft", "bottom", "bottomRight"];

    /* 中央(4)を除く外周の□どうしをつなぐケイ線 / Rules joining the outer squares (the center stands alone) */
    var ANCHOR_WIDGET_CONNECTIONS = [[0, 1], [1, 2], [6, 7], [7, 8], [0, 3], [3, 6], [2, 5], [5, 8]];

    // -----------------------------------------
    // 基準点ウィジェットの配色 / Anchor widget colors
    // -----------------------------------------
    var ANCHOR_WIDGET_UI_DARK = isDarkUI();
    /* 枠線・ケイ線はグレー、選択セルの塗りはライトで濃いグレー・ダークで明るいグレー（既存スクリプトの配色を踏襲）。
       無効時は同じ色を半透明にして背景へ沈める（不透明の薄いグレーだとダークUIで逆に明るく浮くため）
       Gray rules; the selected fill is dark gray on light UI and light gray on dark UI (as in the existing scripts).
       Disabled colors are translucent versions so they sink into any background */
    var ANCHOR_WIDGET_LINE_COLOR     = ANCHOR_WIDGET_UI_DARK ? [0.55, 0.55, 0.55, 1]   : [0.6, 0.6, 0.6, 1];  /* 枠線・ケイ線 / rules */
    var ANCHOR_WIDGET_FILL_COLOR     = ANCHOR_WIDGET_UI_DARK ? [0.8, 0.8, 0.8, 1]      : [0.4, 0.4, 0.4, 1];  /* 選択セルの塗り / selected fill */
    var ANCHOR_WIDGET_DIM_LINE_COLOR = ANCHOR_WIDGET_UI_DARK ? [0.55, 0.55, 0.55, 0.4] : [0.6, 0.6, 0.6, 0.4];  /* 無効時の枠線 / rules when disabled */
    var ANCHOR_WIDGET_DIM_FILL_COLOR = ANCHOR_WIDGET_UI_DARK ? [0.8, 0.8, 0.8, 0.3]    : [0.4, 0.4, 0.4, 0.3];  /* 無効時の塗り / fill when disabled */

    // -----------------------------------------
    // ウィジェットを作る・読み書きする（外から呼ぶ関数） / Public API
    // -----------------------------------------
    /**
     * 基準点（3×3）を選ぶウィジェットを追加する。クリックしたセルを選び、onChange を呼ぶ
     * @param {Group|Panel} parent - 追加先
     * @param {number|string} initialValue - 最初に選ぶセル（0〜8 か "topLeft" などの名前。allowNone なら -1 も可）
     * @param {Function} [onChange] - クリックで選んだときに呼ぶ関数（引数はセルのインデックスとウィジェット）
     * @param {Object} [widgetOptions] - allowNone（true で未選択 -1 を許す）/ disabledCells（選べないセルの配列）/ size（一辺。既定 66）
     * @returns {Button} ウィジェット（値は getAnchorWidgetIndex() / getAnchorWidgetName() で読む）
     */
    function addAnchorWidget(parent, initialValue, onChange, widgetOptions) {
        var anchorOptions = widgetOptions || {};
        var widgetSize = anchorOptions.size || ANCHOR_WIDGET_SIZE;
        var anchorWidget = parent.add("button", undefined, "");
        anchorWidget.minimumSize = [widgetSize, widgetSize];
        anchorWidget.preferredSize = [widgetSize, widgetSize];
        anchorWidget.maximumSize = [widgetSize, widgetSize];
        anchorWidget.isAnchorWidget = true; /* redrawAnchorWidgetsIn() の目印 / marker for redrawAnchorWidgetsIn() */
        anchorWidget.anchorAllowNone = !!anchorOptions.allowNone;
        anchorWidget.anchorDisabledCells = toAnchorCellFlags(anchorOptions.disabledCells);
        anchorWidget.anchorWidgetIndex = resolveAnchorWidgetIndex(initialValue, anchorWidget.anchorAllowNone);
        anchorWidget.onDraw = function () { drawAnchorWidget(anchorWidget); };
        anchorWidget.onClick = function () {}; /* セルの判定は mousedown で行う / hit-testing happens in mousedown */

        /* クリック座標（コントロール基準）を3分割してセルを判定する / split the control-relative click into thirds */
        anchorWidget.addEventListener("mousedown", function (event) {
            if (!isAnchorWidgetEnabledInTree(anchorWidget)) return;
            var cellIndex = getAnchorCellAt(event.clientX, event.clientY, anchorWidget.size[0], anchorWidget.size[1]);
            if (anchorWidget.anchorDisabledCells[cellIndex]) return;
            anchorWidget.anchorWidgetIndex = cellIndex;
            redrawAnchorWidget(anchorWidget);
            if (onChange) onChange(cellIndex, anchorWidget);
        });
        return anchorWidget;
    }

    /**
     * 選択中のセルのインデックスを返す
     * @param {Button} anchorWidget - addAnchorWidget() で作ったウィジェット
     * @returns {number} 0〜8（行優先）。未選択なら -1
     */
    function getAnchorWidgetIndex(anchorWidget) {
        return anchorWidget.anchorWidgetIndex;
    }

    /**
     * 選択中のセルの名前を返す
     * @param {Button} anchorWidget - addAnchorWidget() で作ったウィジェット
     * @returns {string} "topLeft" など。未選択なら ""
     */
    function getAnchorWidgetName(anchorWidget) {
        return ANCHOR_WIDGET_NAMES[anchorWidget.anchorWidgetIndex] || "";
    }

    /**
     * 選択するセルを変えて描き直す（onChange は呼ばない）
     * @param {Button} anchorWidget - addAnchorWidget() で作ったウィジェット
     * @param {number|string} anchorValue - 0〜8 か名前（allowNone なら -1 も可）
     * @returns {void}
     */
    function setAnchorWidgetValue(anchorWidget, anchorValue) {
        anchorWidget.anchorWidgetIndex = resolveAnchorWidgetIndex(anchorValue, anchorWidget.anchorAllowNone);
        redrawAnchorWidget(anchorWidget);
    }

    /**
     * ウィジェットの有効／無効を切り替えて描き直す（無効の間は薄く描き、クリックも無視する）
     * @param {Button} anchorWidget - addAnchorWidget() で作ったウィジェット
     * @param {boolean} isEnabled - 有効にするなら true
     * @returns {void}
     */
    function setAnchorWidgetEnabled(anchorWidget, isEnabled) {
        anchorWidget.enabled = isEnabled;
        redrawAnchorWidget(anchorWidget);
    }

    /**
     * 選べないセルを指定し直して描き直す（選択中のセルは変えない）
     * @param {Button} anchorWidget - addAnchorWidget() で作ったウィジェット
     * @param {number[]} disabledCells - 選べないセルのインデックス（空配列ですべて選べる）
     * @returns {void}
     */
    function setAnchorWidgetCellsDisabled(anchorWidget, disabledCells) {
        anchorWidget.anchorDisabledCells = toAnchorCellFlags(disabledCells);
        redrawAnchorWidget(anchorWidget);
    }

    /**
     * コンテナ以下にある基準点ウィジェットをすべて描き直す。パネルや行の enabled を切り替えたあとに呼ぶ
     * @param {Object} container - パネル・グループ・ウィンドウなど
     * @returns {void}
     */
    function redrawAnchorWidgetsIn(container) {
        if (container.isAnchorWidget) {
            redrawAnchorWidget(container);
            return;
        }
        if (!container.children) return;
        for (var i = 0; i < container.children.length; i++) {
            redrawAnchorWidgetsIn(container.children[i]);
        }
    }

    // -----------------------------------------
    // 値の変換 / Value helpers
    // -----------------------------------------
    /**
     * セルのインデックスか名前を 0〜8 のインデックスにする。解釈できない値は中央（4）
     * @param {number|string} anchorValue - 0〜8 / -1 / "topLeft" などの名前
     * @param {boolean} [allowNone] - true なら -1（未選択）をそのまま返す
     * @returns {number} 0〜8。allowNone で -1 を渡したときだけ -1
     */
    function resolveAnchorWidgetIndex(anchorValue, allowNone) {
        if (typeof anchorValue === "string") {
            for (var i = 0; i < ANCHOR_WIDGET_NAMES.length; i++) {
                if (ANCHOR_WIDGET_NAMES[i] === anchorValue) return i;
            }
            return 4;
        }
        if (anchorValue === ANCHOR_WIDGET_NONE && allowNone) return ANCHOR_WIDGET_NONE;
        if (typeof anchorValue === "number" && anchorValue >= 0 && anchorValue <= 8 && anchorValue === Math.floor(anchorValue)) {
            return anchorValue;
        }
        return 4;
    }

    /**
     * セルの位置を割合で返す（左・上が 0、中央が 0.5、右・下が 1）
     * @param {number|string} anchorValue - 0〜8 か名前
     * @returns {number[]} [横の割合, 縦の割合]
     */
    function getAnchorRatio(anchorValue) {
        var anchorIndex = resolveAnchorWidgetIndex(anchorValue);
        return [(anchorIndex % 3) / 2, Math.floor(anchorIndex / 3) / 2];
    }

    /**
     * 境界ボックス上の基準点の座標を返す（Illustrator の [左, 上, 右, 下] でも、y 下向きの座標でもそのまま使える）
     * @param {number[]} bounds - [左, 上, 右, 下]（geometricBounds・visibleBounds・artboardRect など）
     * @param {number|string} anchorValue - 0〜8 か名前
     * @returns {number[]} [x, y]
     */
    function getAnchorPointOnBounds(bounds, anchorValue) {
        var anchorRatio = getAnchorRatio(anchorValue);
        return [
            bounds[0] + (bounds[2] - bounds[0]) * anchorRatio[0],
            bounds[1] + (bounds[3] - bounds[1]) * anchorRatio[1]
        ];
    }

    /**
     * resize()・rotate()・transform() に渡す基準点を返す（Illustrator 専用）。
     * 基準は効果を含まない境界（geometricBounds）
     * @param {number|string} anchorValue - 0〜8 か名前
     * @returns {Transformation} Transformation.TOPLEFT など
     */
    function getAnchorTransformation(anchorValue) {
        var transformations = [
            Transformation.TOPLEFT, Transformation.TOP, Transformation.TOPRIGHT,
            Transformation.LEFT, Transformation.CENTER, Transformation.RIGHT,
            Transformation.BOTTOMLEFT, Transformation.BOTTOM, Transformation.BOTTOMRIGHT
        ];
        return transformations[resolveAnchorWidgetIndex(anchorValue)];
    }

    /**
     * symbols.add() に渡す登録点を返す（Illustrator 専用）
     * @param {number|string} anchorValue - 0〜8 か名前
     * @returns {SymbolRegistrationPoint} SymbolRegistrationPoint.SYMBOLTOPLEFTPOINT など
     */
    function getAnchorSymbolRegistrationPoint(anchorValue) {
        var registrationPoints = [
            SymbolRegistrationPoint.SYMBOLTOPLEFTPOINT, SymbolRegistrationPoint.SYMBOLTOPMIDDLEPOINT, SymbolRegistrationPoint.SYMBOLTOPRIGHTPOINT,
            SymbolRegistrationPoint.SYMBOLMIDDLELEFTPOINT, SymbolRegistrationPoint.SYMBOLCENTERPOINT, SymbolRegistrationPoint.SYMBOLMIDDLERIGHTPOINT,
            SymbolRegistrationPoint.SYMBOLBOTTOMLEFTPOINT, SymbolRegistrationPoint.SYMBOLBOTTOMMIDDLEPOINT, SymbolRegistrationPoint.SYMBOLBOTTOMRIGHTPOINT
        ];
        return registrationPoints[resolveAnchorWidgetIndex(anchorValue)];
    }

    /**
     * クリック位置からセルのインデックスを求める（ウィジェットを縦横3等分し、外にはみ出した座標は端のセルに寄せる）
     * @param {number} clickX - コントロール基準の x
     * @param {number} clickY - コントロール基準の y
     * @param {number} widgetWidth - ウィジェットの幅
     * @param {number} widgetHeight - ウィジェットの高さ
     * @returns {number} 0〜8
     */
    function getAnchorCellAt(clickX, clickY, widgetWidth, widgetHeight) {
        var column = Math.min(2, Math.max(0, Math.floor(clickX / (widgetWidth / 3))));
        var row = Math.min(2, Math.max(0, Math.floor(clickY / (widgetHeight / 3))));
        return row * 3 + column;
    }

    /**
     * セルのインデックスの配列を、9個の真偽値に直す
     * @param {number[]} [cellIndexes] - セルのインデックスの配列
     * @returns {boolean[]} 含まれるセルだけ true
     */
    function toAnchorCellFlags(cellIndexes) {
        var cellFlags = [false, false, false, false, false, false, false, false, false];
        if (!cellIndexes) return cellFlags;
        for (var i = 0; i < cellIndexes.length; i++) {
            if (cellIndexes[i] >= 0 && cellIndexes[i] <= 8) cellFlags[cellIndexes[i]] = true;
        }
        return cellFlags;
    }

    // -----------------------------------------
    // 描画 / Drawing
    // -----------------------------------------
    /**
     * ウィジェットを描く（外周の□をケイ線でつなぎ、中央は独立。選択セルだけ塗る）
     * @param {Button} anchorWidget - 描くウィジェット
     * @returns {void}
     */
    function drawAnchorWidget(anchorWidget) {
        var graphics = anchorWidget.graphics;
        var widgetWidth = anchorWidget.size[0];
        var widgetHeight = anchorWidget.size[1];
        var cellSize = ANCHOR_WIDGET_CELL_SIZE;
        var halfCell = cellSize / 2;
        /* 自作描画は自動でディムにならないので、親までたどって判定する / custom drawing is not dimmed automatically */
        var isEnabled = isAnchorWidgetEnabledInTree(anchorWidget);

        /* ボタンの地をコントロールの地色で塗り、パネルに溶け込ませる（backgroundColor が無い環境では例外）
           Paint the control's own background so the widget blends into the panel; throws where backgroundColor is missing */
        try {
            graphics.newPath();
            graphics.rectPath(0, 0, widgetWidth, widgetHeight);
            graphics.fillPath(graphics.backgroundColor);
        } catch (e) {}

        var cellStep = cellSize + ANCHOR_WIDGET_CELL_GAP;
        var gridSize = cellSize * 3 + ANCHOR_WIDGET_CELL_GAP * 2;
        var originX = Math.round((widgetWidth - gridSize) / 2);
        var originY = Math.round((widgetHeight - gridSize) / 2);
        var cellPositions = [];
        var i;
        for (i = 0; i < 9; i++) {
            cellPositions.push([originX + (i % 3) * cellStep, originY + Math.floor(i / 3) * cellStep]);
        }

        var linePen = graphics.newPen(graphics.PenType.SOLID_COLOR, isEnabled ? ANCHOR_WIDGET_LINE_COLOR : ANCHOR_WIDGET_DIM_LINE_COLOR, 1);
        for (i = 0; i < ANCHOR_WIDGET_CONNECTIONS.length; i++) {
            var cellA = cellPositions[ANCHOR_WIDGET_CONNECTIONS[i][0]];
            var cellB = cellPositions[ANCHOR_WIDGET_CONNECTIONS[i][1]];
            graphics.newPath();
            if (ANCHOR_WIDGET_CONNECTIONS[i][1] - ANCHOR_WIDGET_CONNECTIONS[i][0] === 1) {
                /* 横方向：右隣の□へ / horizontal: to the square on the right */
                graphics.moveTo(cellA[0] + cellSize, cellA[1] + halfCell);
                graphics.lineTo(cellB[0], cellB[1] + halfCell);
            } else {
                /* 縦方向：下の□へ / vertical: to the square below */
                graphics.moveTo(cellA[0] + halfCell, cellA[1] + cellSize);
                graphics.lineTo(cellB[0] + halfCell, cellB[1]);
            }
            graphics.strokePath(linePen);
        }

        for (i = 0; i < 9; i++) {
            var isCellEnabled = isEnabled && !anchorWidget.anchorDisabledCells[i];
            drawAnchorWidgetCell(graphics, cellPositions[i][0], cellPositions[i][1], i === anchorWidget.anchorWidgetIndex, isCellEnabled);
        }
    }

    /**
     * □を1つ描く（選択中だけ塗り、枠は塗りの上に重ねる）
     * @param {ScriptUIGraphics} graphics - 描画先
     * @param {number} cellX - 左端
     * @param {number} cellY - 上端
     * @param {boolean} isSelected - 選択中なら true
     * @param {boolean} isEnabled - 選べるセルなら true（false なら薄く描く）
     * @returns {void}
     */
    function drawAnchorWidgetCell(graphics, cellX, cellY, isSelected, isEnabled) {
        var cellSize = ANCHOR_WIDGET_CELL_SIZE;
        /* rectPath の前には毎回 newPath()（呼ばないとパスが累積して塗りが線画になる）
           Always call newPath() before rectPath(), or paths accumulate and fills turn into outlines */
        if (isSelected) {
            graphics.newPath();
            graphics.rectPath(cellX, cellY, cellSize, cellSize);
            graphics.fillPath(graphics.newBrush(graphics.BrushType.SOLID_COLOR, isEnabled ? ANCHOR_WIDGET_FILL_COLOR : ANCHOR_WIDGET_DIM_FILL_COLOR));
        }
        graphics.newPath();
        graphics.rectPath(cellX, cellY, cellSize, cellSize);
        graphics.strokePath(graphics.newPen(graphics.PenType.SOLID_COLOR, isEnabled ? ANCHOR_WIDGET_LINE_COLOR : ANCHOR_WIDGET_DIM_LINE_COLOR, 1));
    }

    /**
     * コントロールと、その親をたどってすべて有効かを返す（親の無効化は子の enabled に出ない）
     * @param {Object} control - 対象のコントロール
     * @returns {boolean} すべて有効なら true
     */
    function isAnchorWidgetEnabledInTree(control) {
        for (var node = control; node; node = node.parent) {
            if (node.enabled === false) return false;
        }
        return true;
    }

    /**
     * ウィジェットの onDraw を呼び直す。notify("onDraw") は環境によって例外や空振りになるため、隠して再表示して描き直させる
     * @param {Button} anchorWidget - 描き直すウィジェット
     * @returns {void}
     */
    function redrawAnchorWidget(anchorWidget) {
        anchorWidget.hide();
        anchorWidget.show();
    }

    // 基準点ウィジェット（再利用パーツ）ここまで / End of the reusable anchor widget

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
    // 境界と整列先の座標 / Bounds and target values
    // =========================================

    /**
     * オブジェクト1つの境界を返す
     * クリップグループはプレビュー境界の設定によらず、マスクの幾何境界を使う。
     * 中にクリップグループを含むグループは、子の境界を合わせて測る（隠れた部分を含めない）
     * @param {PageItem} pageItem - 対象オブジェクト
     * @param {boolean} usePreviewBounds - 線や効果を含めるなら true
     * @returns {number[]} [左, 上, 右, 下]
     */
    function getItemBounds(pageItem, usePreviewBounds) {
        /* クリップグループのマスクは常に幾何境界で測る / Clip-group masks are always measured by geometric bounds */
        var isClipGroup = (getClipMaskItem(pageItem) !== null);
        return getClipAwareBounds(pageItem, isClipGroup ? false : usePreviewBounds);
    }

    /**
     * 選択オブジェクト群を包含する境界を返す
     * @param {PageItem[]} pageItems - 対象オブジェクトの配列
     * @param {boolean} usePreviewBounds - 線や効果を含めるなら true
     * @returns {number[]} [左, 上, 右, 下]
     */
    function computeSelectionBounds(pageItems, usePreviewBounds) {
        var unionBounds = null;
        for (var itemIndex = 0; itemIndex < pageItems.length; itemIndex++) {
            var itemBounds = getItemBounds(pageItems[itemIndex], usePreviewBounds);
            if (!unionBounds) {
                unionBounds = [itemBounds[0], itemBounds[1], itemBounds[2], itemBounds[3]];
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
     * 整列先に対応する座標を境界から取り出す（中央揃えは2辺の中点）
     * @param {number[]} bounds - [左, 上, 右, 下]
     * @param {string} alignmentSide - 整列先（EDGE_RULES のキー）
     * @returns {number} 座標。対応しない整列先なら null
     */
    function getAlignmentValue(bounds, alignmentSide) {
        var edgeRule = EDGE_RULES[alignmentSide];
        if (!edgeRule) return null;
        if (edgeRule.boundsIndex !== null) return bounds[edgeRule.boundsIndex];
        return (edgeRule.axis === "x") ? (bounds[0] + bounds[2]) / 2 : (bounds[1] + bounds[3]) / 2;
    }

    /**
     * 基準の座標より揃える向きの先にあるかを返す
     * @param {number} value - 判定する座標
     * @param {number} referenceValue - 基準の座標
     * @param {string} alignmentSide - 整列先（EDGE_RULES のキー）
     * @returns {boolean} 先にあれば true
     */
    function isAheadOnSide(value, referenceValue, alignmentSide) {
        return (value - referenceValue) * EDGE_RULES[alignmentSide].ahead > 0;
    }

    // =========================================
    // ガイドの探索 / Guide lookup
    // =========================================

    /**
     * 整列方向に対応するガイド座標を返す（向きが合わないガイドは対象外）
     * @param {number[]} guideBounds - ガイドの幾何境界 [左, 上, 右, 下]
     * @param {string} alignmentSide - 整列先（EDGE_RULES のキー）
     * @returns {number} ガイドの座標。対象外なら null
     */
    function getGuideValue(guideBounds, alignmentSide) {
        if (EDGE_RULES[alignmentSide].axis === "x") {
            /* 左右の整列先になるのは縦ガイド（幅ゼロ）だけ / Only vertical guides (zero width) serve left/right */
            return (Math.abs(guideBounds[2] - guideBounds[0]) <= GUIDE_ORIENTATION_TOLERANCE) ? guideBounds[0] : null;
        }
        /* 上下の整列先になるのは横ガイド（高さゼロ）だけ / Only horizontal guides (zero height) serve top/bottom */
        return (Math.abs(guideBounds[1] - guideBounds[3]) <= GUIDE_ORIENTATION_TOLERANCE) ? guideBounds[1] : null;
    }

    /**
     * ガイド座標がアクティブアートボードの内側にあるかを返す
     * @param {number} guideValue - ガイドの座標
     * @param {number[]} artboardRect - アートボードの矩形 [左, 上, 右, 下]
     * @param {string} alignmentSide - 整列先（EDGE_RULES のキー）
     * @returns {boolean} 内側なら true
     */
    function isGuideInsideArtboard(guideValue, artboardRect, alignmentSide) {
        var isHorizontalAxis = (EDGE_RULES[alignmentSide].axis === "x");
        var lowerBound = isHorizontalAxis ? artboardRect[0] : artboardRect[3];
        var upperBound = isHorizontalAxis ? artboardRect[2] : artboardRect[1];
        return guideValue >= lowerBound && guideValue <= upperBound;
    }

    /**
     * アートボード内側のガイドから、整列先として使う座標を探す
     * @param {number} selectionEdge - 選択範囲の境界値
     * @param {string} alignmentSide - 整列先（EDGE_RULES のキー）
     * @param {object} alignContext - 整列コンテキスト
     * @returns {number} 吸着先の座標。見つからなければ null
     */
    function findGuideSnapValue(selectionEdge, alignmentSide, alignContext) {
        var nearestGuideValue = null;
        var nearestGuideDistance = null;
        var guidePathItems = alignContext.documentRef.pathItems;
        var insideOnly = (GUIDE_SEARCH_MODE === "inside");

        for (var guidePathIndex = 0; guidePathIndex < guidePathItems.length; guidePathIndex++) {
            var guidePathItem = guidePathItems[guidePathIndex];
            if (guidePathItem.guides !== true) continue;

            var guideValue = getGuideValue(guidePathItem.geometricBounds, alignmentSide);
            if (guideValue === null) continue;
            if (!isGuideInsideArtboard(guideValue, alignContext.artboardRect, alignmentSide)) continue;
            /* "inside" は揃える向きの先にあるガイドだけを候補にする（"nearest" は向きを問わない）
               "inside" only accepts guides ahead of the selection; "nearest" takes either direction */
            if (insideOnly && !isAheadOnSide(guideValue, selectionEdge, alignmentSide)) continue;

            var guideDistance = Math.abs(guideValue - selectionEdge);
            if (nearestGuideDistance === null || guideDistance < nearestGuideDistance) {
                nearestGuideValue = guideValue;
                nearestGuideDistance = guideDistance;
            }
        }
        return nearestGuideValue;
    }

    // =========================================
    // 整列の適用 / Applying the alignment
    // =========================================

    /**
     * 整列コンテキストを作る
     * @param {Document} documentRef - 対象ドキュメント
     * @param {number[]} artboardRect - アクティブアートボードの矩形 [左, 上, 右, 下]
     * @param {boolean} useGuides - ガイドを整列先に含めるなら true
     * @returns {object} 整列コンテキスト
     */
    function createAlignContext(documentRef, artboardRect, useGuides) {
        return { documentRef: documentRef, artboardRect: artboardRect, useGuides: useGuides };
    }

    /**
     * 1軸分の整列オフセットを計算する
     * @param {string} alignmentSide - 整列先（EDGE_RULES のキー）
     * @param {number[]} selectionBounds - 選択範囲の境界 [左, 上, 右, 下]
     * @param {object} alignContext - 整列コンテキスト
     * @returns {number} 移動量
     */
    function computeAxisOffset(alignmentSide, selectionBounds, alignContext) {
        var selectionValue = getAlignmentValue(selectionBounds, alignmentSide);
        var targetValue = getAlignmentValue(alignContext.artboardRect, alignmentSide);
        if (selectionValue === null || targetValue === null) return 0;

        /* ガイドが吸着先になるのは端揃えのときだけ（中央揃えは ahead が 0）
           Guides only snap for edge alignment; center alignment has ahead 0 */
        if (alignContext.useGuides && EDGE_RULES[alignmentSide].ahead !== 0) {
            var guideValue = findGuideSnapValue(selectionValue, alignmentSide, alignContext);
            if (guideValue !== null) targetValue = guideValue;
        }
        return targetValue - selectionValue;
    }

    /**
     * 選択オブジェクトの現在位置を控える
     * @param {PageItem[]} pageItems - 対象オブジェクトの配列
     * @returns {number[][]} [X, Y] の配列
     */
    function captureItemPositions(pageItems) {
        var capturedPositions = [];
        for (var itemIndex = 0; itemIndex < pageItems.length; itemIndex++) {
            var itemPosition = pageItems[itemIndex].position;
            capturedPositions.push([itemPosition[0], itemPosition[1]]);
        }
        return capturedPositions;
    }

    /**
     * 控えた位置へオブジェクトを戻す
     * @param {PageItem[]} pageItems - 対象オブジェクトの配列
     * @param {number[][]} capturedPositions - captureItemPositions() の戻り値
     * @returns {void}
     */
    function restoreItemPositions(pageItems, capturedPositions) {
        for (var itemIndex = 0; itemIndex < pageItems.length; itemIndex++) {
            pageItems[itemIndex].position = capturedPositions[itemIndex];
        }
    }

    /**
     * 整列を1回分適用する（境界の計測からオフセットの適用まで）
     * @param {PageItem[]} pageItems - 対象オブジェクトの配列
     * @param {object} alignmentAxes - 整列軸 { horizontal, vertical }
     * @param {object} alignContext - 整列コンテキスト
     * @param {boolean} usePreviewBounds - 線や効果を含めるなら true
     * @returns {void}
     */
    function applyAlignment(pageItems, alignmentAxes, alignContext, usePreviewBounds) {
        var selectionBounds = computeSelectionBounds(pageItems, usePreviewBounds);
        var offsetX = alignmentAxes.horizontal ? computeAxisOffset(alignmentAxes.horizontal, selectionBounds, alignContext) : 0;
        var offsetY = alignmentAxes.vertical ? computeAxisOffset(alignmentAxes.vertical, selectionBounds, alignContext) : 0;
        for (var itemIndex = 0; itemIndex < pageItems.length; itemIndex++) {
            pageItems[itemIndex].translate(offsetX, offsetY);
        }
    }

    /**
     * スクリプトのファイル名から整列先を判定する（例: GroupEdgeAlignRIGHT.jsx → "right"）
     * @returns {string} 整列先。判定できない場合は DEFAULT_ALIGNMENT_SIDE
     */
    function detectAlignmentSideFromFileName() {
        var fileNameUpper = File($.fileName).name.toUpperCase();
        for (var keywordIndex = 0; keywordIndex < ALIGNMENT_SIDE_BY_FILENAME.length; keywordIndex++) {
            if (fileNameUpper.indexOf(ALIGNMENT_SIDE_BY_FILENAME[keywordIndex].keyword) !== -1) {
                return ALIGNMENT_SIDE_BY_FILENAME[keywordIndex].side;
            }
        }
        return DEFAULT_ALIGNMENT_SIDE;
    }

    /**
     * プレビューと矢印キーのステップ移動をまとめた整列セッションを作る
     * @param {PageItem[]} pageItems - 対象オブジェクトの配列
     * @param {Document} documentRef - 対象ドキュメント
     * @param {number[]} artboardRect - アクティブアートボードの矩形 [左, 上, 右, 下]
     * @returns {object} preview / step / restoreOriginal をまとめたオブジェクト
     */
    function createAlignmentSession(pageItems, documentRef, artboardRect) {
        /* originalPositions はキャンセル時の完全復元用（不変）、
           basePositions はプレビュー復元の基準で、矢印キーのステップ移動のたびに進む
           originalPositions restores everything on Cancel; basePositions is the preview baseline
           and advances with each arrow-key step */
        var originalPositions = captureItemPositions(pageItems);
        var basePositions = captureItemPositions(pageItems);

        /**
         * 基準位置に戻してからプレビューの整列を適用する
         * @param {object} previewSettings - 整列軸・境界・ガイドの設定。null なら復元のみ
         * @returns {void}
         */
        function preview(previewSettings) {
            restoreItemPositions(pageItems, basePositions);
            if (previewSettings && previewSettings.alignmentAxes) {
                var alignContext = createAlignContext(documentRef, artboardRect, previewSettings.useGuides);
                applyAlignment(pageItems, previewSettings.alignmentAxes, alignContext, previewSettings.usePreviewBounds);
            }
            app.redraw();
        }

        /**
         * 矢印キーによる1段階の整列（スクリプトを1回実行したのと同じ挙動）
         * @param {string} alignmentSide - "left" / "right" / "top" / "bottom"
         * @param {object} stepSettings - 境界とガイドの設定
         * @returns {void}
         */
        function step(alignmentSide, stepSettings) {
            preview({
                alignmentAxes: AXES_BY_ALIGNMENT_SIDE[alignmentSide],
                usePreviewBounds: stepSettings.usePreviewBounds,
                useGuides: stepSettings.useGuides
            });
            /* 動いた先を次の基準にして、押すたびにさらに先へ進めるようにする
               The new position becomes the baseline so each press advances further */
            basePositions = captureItemPositions(pageItems);
        }

        /**
         * 矢印キーでのステップ移動も含めて、実行前の位置へ戻す
         * @returns {void}
         */
        function restoreOriginal() {
            restoreItemPositions(pageItems, originalPositions);
            app.redraw();
        }

        return { preview: preview, step: step, restoreOriginal: restoreOriginal };
    }

    // =========================================
    // ダイアログ UI / Dialog UI
    // =========================================

    /**
     * 整列ダイアログを組み立てる
     * @param {object} initialSettings - 境界とガイドの初期値
     * @param {Function} onCellPicked - 9軸ウィジェットのセルをクリックしたときに呼ぶ関数
     * @returns {{alignDialog: Window, anchorWidget: Button, previewBoundsCheckbox: Checkbox, useGuidesCheckbox: Checkbox, previewCheckbox: Checkbox}} ダイアログと主なコントロール
     */
    function buildAlignmentDialog(initialSettings, onCellPicked) {
        var alignDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        alignDialog.orientation = "column";
        alignDialog.alignChildren = ["fill", "top"];
        alignDialog.margins = WINDOW_MARGINS;
        alignDialog.spacing = WINDOW_SPACING;

        var alignmentPanel = addPanel(alignDialog, getLabel("panel.alignment"));
        alignmentPanel.margins = ANCHOR_PANEL_MARGINS;
        alignmentPanel.alignChildren = ["center", "top"];
        /* 未選択（NO_ANCHOR_INDEX）から始める / Starts with no cell selected */
        var anchorWidget = addAnchorWidget(alignmentPanel, NO_ANCHOR_INDEX, onCellPicked, { allowNone: true });
        anchorWidget.helpTip = getLabel("tooltip.anchorWidget");

        var optionPanel = addPanel(alignDialog, getLabel("panel.option"));
        optionPanel.alignChildren = ["left", "top"];

        var previewBoundsCheckbox = optionPanel.add("checkbox", undefined, getLabel("checkbox.previewBounds"));
        previewBoundsCheckbox.helpTip = getLabel("tooltip.previewBounds");
        previewBoundsCheckbox.value = initialSettings.usePreviewBounds;

        var useGuidesCheckbox = optionPanel.add("checkbox", undefined, getLabel("checkbox.useGuides"));
        useGuidesCheckbox.helpTip = getLabel("tooltip.useGuides");
        useGuidesCheckbox.value = initialSettings.useGuides;

        /* ボタンエリア（左にプレビュー、右にキャンセル・OK）/ Button row (Preview on the left, Cancel / OK on the right) */
        var buttonRow = addButtonRow(alignDialog);
        var previewCheckbox = buttonRow.leftGroup.add("checkbox", undefined, getLabel("checkbox.preview"));
        previewCheckbox.helpTip = getLabel("tooltip.preview");
        previewCheckbox.value = false;
        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });

        return {
            alignDialog: alignDialog,
            anchorWidget: anchorWidget,
            previewBoundsCheckbox: previewBoundsCheckbox,
            useGuidesCheckbox: useGuidesCheckbox,
            previewCheckbox: previewCheckbox
        };
    }

    /**
     * ダイアログの状態を設定オブジェクトにまとめる
     * @param {object} dialogControls - buildAlignmentDialog() の戻り値
     * @returns {object} 整列軸・境界・ガイドの設定（整列先が未選択なら alignmentAxes は null）
     */
    function readAlignSettings(dialogControls) {
        var selectedIndex = getAnchorWidgetIndex(dialogControls.anchorWidget);
        var hasAnchor = (selectedIndex !== NO_ANCHOR_INDEX);
        return {
            /* ANCHOR_DEFINITIONS の要素はそのまま整列軸として使える / Each entry doubles as the axes pair */
            alignmentAxes: hasAnchor ? ANCHOR_DEFINITIONS[selectedIndex] : null,
            usePreviewBounds: dialogControls.previewBoundsCheckbox.value,
            /* 整列先を選んだときはアートボード基準に固定 / A chosen target always aligns to the artboard */
            useGuides: !hasAnchor && dialogControls.useGuidesCheckbox.value
        };
    }

    /**
     * ショートカットキーに対応するセルの番号を返す
     * @param {string} pressedKey - 押されたキー名
     * @returns {number} ANCHOR_DEFINITIONS 上のインデックス。対応が無ければ NO_ANCHOR_INDEX
     */
    function findAnchorIndexByShortcutKey(pressedKey) {
        for (var anchorIndex = 0; anchorIndex < ANCHOR_DEFINITIONS.length; anchorIndex++) {
            if (ANCHOR_DEFINITIONS[anchorIndex].shortcutKey === pressedKey) return anchorIndex;
        }
        return NO_ANCHOR_INDEX;
    }

    /**
     * 整列オプションのダイアログを表示し、選択内容を返す
     * @param {object} initialSettings - 境界とガイドの初期値
     * @param {object} alignmentSession - プレビューとステップ移動を担う整列セッション
     * @returns {object} 整列軸・境界・ガイドの設定。キャンセル時は null
     */
    function showAlignmentDialog(initialSettings, alignmentSession) {
        var dialogControls = buildAlignmentDialog(initialSettings, function (cellIndex) {
            selectAnchorAt(cellIndex);
            triggerPreview();
        });
        var previewBoundsCheckbox = dialogControls.previewBoundsCheckbox;
        var useGuidesCheckbox = dialogControls.useGuidesCheckbox;
        var previewCheckbox = dialogControls.previewCheckbox;

        /**
         * プレビューを更新する（OFF・整列先未選択のときは基準位置へ戻す）
         * @returns {void}
         */
        function triggerPreview() {
            var currentSettings = readAlignSettings(dialogControls);
            /* 整列先が未選択なら戻すだけ（ファイル名由来のフォールバックを抑止）
               With no target selected, only restore; the filename fallback must not kick in here */
            alignmentSession.preview((previewCheckbox.value && currentSettings.alignmentAxes) ? currentSettings : null);
        }

        /**
         * 整列先を選び直し、ウィジェットとガイドのチェックボックスを更新する
         * @param {number} selectedIndex - ANCHOR_DEFINITIONS 上のインデックス（未選択は NO_ANCHOR_INDEX）
         * @returns {void}
         */
        function selectAnchorAt(selectedIndex) {
            setAnchorWidgetValue(dialogControls.anchorWidget, selectedIndex);
            /* 整列先を選ぶとガイドは使わないので、チェックボックスも無効にする
               A chosen target ignores guides, so the checkbox goes dim */
            useGuidesCheckbox.enabled = (selectedIndex === NO_ANCHOR_INDEX);
        }

        /**
         * ダイアログでのキー操作を処理する（矢印キー、W〜V のセル選択、G・B の切り替え）
         * @param {string} keyName - 離したキーの名前
         * @returns {void}
         */
        function handleKeyUp(keyName) {
            var stepSide = STEP_SIDE_BY_KEY[keyName];
            if (stepSide) {
                /* 矢印キー：押すたびに「スクリプトを1回実行」相当のステップ移動。
                   整列先の選択は解除する（OK 時に最終整列が二重適用されないようにするため）
                   Arrow keys step as if the script ran once; the target selection is cleared so
                   the final alignment on OK is not applied twice */
                selectAnchorAt(NO_ANCHOR_INDEX);
                alignmentSession.step(stepSide, readAlignSettings(dialogControls));
                return;
            }
            var anchorIndex = findAnchorIndexByShortcutKey(keyName);
            if (anchorIndex !== NO_ANCHOR_INDEX) {
                selectAnchorAt(anchorIndex);
                triggerPreview();
                return;
            }
            if (keyName === "G" && useGuidesCheckbox.enabled) {
                useGuidesCheckbox.value = !useGuidesCheckbox.value;
                triggerPreview();
                return;
            }
            if (keyName === "B") {
                previewBoundsCheckbox.value = !previewBoundsCheckbox.value;
                triggerPreview();
            }
        }

        previewBoundsCheckbox.onClick = triggerPreview;
        useGuidesCheckbox.onClick = triggerPreview;
        previewCheckbox.onClick = triggerPreview;
        dialogControls.alignDialog.addEventListener("keyup", function (keyEvent) {
            handleKeyUp(keyEvent.keyName);
        });

        prepareDialogWindow(dialogControls.alignDialog, SCRIPT_NAME);
        var dialogShowResult = dialogControls.alignDialog.show();

        /* OK / キャンセルどちらでも、閉じる際はいったん基準位置へ戻す（最終整列は main 側で改めて適用）
           On both OK and Cancel the items go back to the baseline; main re-applies the final alignment */
        alignmentSession.preview(null);
        return (dialogShowResult === 1) ? readAlignSettings(dialogControls) : null;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択オブジェクトをアートボードまたはガイドへ整列する
     * @returns {void}
     */
    function main() {
        if (GUIDE_SEARCH_MODE !== "inside" && GUIDE_SEARCH_MODE !== "nearest") {
            alert(labelText("alert.invalidGuideSearchMode") + GUIDE_SEARCH_MODE);
            return;
        }

        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }

        var documentRef = app.activeDocument;
        var selectedItems = documentRef.selection;
        if (selectedItems.length === 0) {
            alert(getLabel("alert.noSelection"));
            return;
        }

        var artboards = documentRef.artboards;
        var artboardRect = artboards[artboards.getActiveArtboardIndex()].artboardRect;

        var alignSettings = {
            /* ダイアログを出さないときはファイル名から整列先を決める / Without the dialog the filename picks the target */
            alignmentAxes: AXES_BY_ALIGNMENT_SIDE[detectAlignmentSideFromFileName()],
            /* プレビュー境界使用の初期値は環境設定から取得 / Seed the preview-bounds option from the preferences */
            usePreviewBounds: app.preferences.getBooleanPreference("includeStrokeInBounds"),
            useGuides: USE_GUIDES
        };

        if (SHOW_DIALOG) {
            var alignmentSession = createAlignmentSession(selectedItems, documentRef, artboardRect);
            alignSettings = showAlignmentDialog(alignSettings, alignmentSession);
            if (alignSettings === null) {
                /* キャンセル：矢印キーでのステップ移動も含めて完全復元 / Cancel restores the arrow-key steps too */
                alignmentSession.restoreOriginal();
                return;
            }
            /* 整列先が未選択なら、矢印キーでの最終位置をそのまま確定する
               With no target selected, the arrow-key result stands as-is */
            if (!alignSettings.alignmentAxes) return;
        }

        var alignContext = createAlignContext(documentRef, artboardRect, alignSettings.useGuides);
        applyAlignment(selectedItems, alignSettings.alignmentAxes, alignContext, alignSettings.usePreviewBounds);
    }

    main();

})();
