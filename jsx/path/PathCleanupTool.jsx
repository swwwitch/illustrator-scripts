#target illustrator
#targetengine "PathCleanupToolEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したパス（グループ・複合パスの中も含む）のアンカーポイントとハンドルを整理します。
削除・変換の内容を選び、アンカー数とハンドル数の増減を確認してから実行できます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/PathCleanupTool.md

note記事も参照してください。
https://note.com/dtp_tranist/n/nd82f59bf63a8

### Overview

Tidies the anchor points and handles of the selected paths, including those inside groups and compound paths.
You choose what to remove or convert and can see how the anchor and handle counts will change before running it.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/PathCleanupTool.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "PathCleanupTool";              /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.7.4";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-03-01";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-30";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/PathCleanupTool.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/PathCleanupTool.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/nd82f59bf63a8"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    /* 許容誤差の初期値（ダイアログで調整可能） / Default tolerances (adjustable in dialog) */
    var TOL_ANCHOR_COLLINEAR = 0.02; /* 直線上アンカー削除の許容誤差 / Redundant-anchor removal */
    var TOL_HANDLE_COLLINEAR = 0.01; /* 直線区間ハンドル整理の許容誤差 / Straight-segment handle normalization */
    var TOL_SAMEPOINT = 0.02;        /* 同一点判定の許容誤差 / Coincident-point check */

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

    /**
     * 日英ラベル定義（カテゴリ別） / Japanese-English label definitions (by category)
     * getLabel("dialog.title") のようにドット区切りで参照する。短い文言は1行で記述。
     * @type {Object}
     */
    var LABELS = {
        dialog: {
            title: { ja: "パスの最適化", en: "Path Optimization" }
        },
        panel: {
            info:          { ja: "情報", en: "Info" },
            convertPoints: { ja: "アンカーポイントを変換", en: "Convert anchor points" },
            addPoints:     { ja: "アンカーポイントを追加", en: "Add anchor points" },
            pathOps:       { ja: "その他", en: "Other" }
        },
        tab: {
            process: { ja: "削除対象", en: "Removal Targets" },
            other:   { ja: "変換", en: "Transform" }
        },
        checkbox: {
            removeSameAnchors: { ja: "同じ座標のアンカーポイント", en: "Duplicate anchor points" },
            removeAnchors:     { ja: "直線上のアンカーポイント", en: "Collinear anchor points" },
            removeHandles:     { ja: "パスと同じ角度のハンドル", en: "Handles on straight segments" }
        },
        radio: {
            convertSmooth:  { ja: "スムーズポイントに", en: "To smooth points" },
            convertCorner:  { ja: "コーナーポイントに", en: "To corner points" },
            addAnchors:     { ja: "中間に追加", en: "At midpoints" },
            addExtremePoints: { ja: "極点を追加", en: "Add Extreme Points" },
            splitAtAnchors: { ja: "アンカーポイントで分割", en: "Split at anchor points" },
            fillHoles:      { ja: "マド埋め", en: "Fill holes" }
        },
        label: {
            pathCount:   { ja: "パスの数", en: "Paths" },
            anchorCount: { ja: "アンカーポイント数", en: "Anchor points" },
            handleCount: { ja: "ハンドル数", en: "Handles" },
            tolAnchor:   { ja: "許容誤差", en: "Tolerance" },
            tolHandle:   { ja: "許容誤差", en: "Tolerance" }
        },
        button: {
            ok:     { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        alert: {
            noDocument:    { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            needSelection: { ja: "パスを選択してから実行してください。", en: "Please select paths before running." },
            lostSelection: { ja: "選択していたオブジェクトが見つからないため、処理を中止しました。", en: "The selected objects are no longer available, so processing was cancelled." }
        },
        tooltip: {
            stepUp: {
                ja: "値を増やす（shift＋クリックで10の倍数へ、option＋クリックで0.1ずつ）",
                en: "Increase (Shift-click to snap to 10s, Option-click by 0.1)"
            },
            stepDown: {
                ja: "値を減らす（shift＋クリックで10の倍数へ、option＋クリックで0.1ずつ）",
                en: "Decrease (Shift-click to snap to 10s, Option-click by 0.1)"
            },
            stepUpInteger:   { ja: "値を増やす（shift＋クリックで10の倍数へ）", en: "Increase (Shift-click to snap to 10s)" },
            stepDownInteger: { ja: "値を減らす（shift＋クリックで10の倍数へ）", en: "Decrease (Shift-click to snap to 10s)" },
            stepUpTolerance:   { ja: "値を0.01増やす（shift＋クリックで0.1の倍数へ）", en: "Increase by 0.01 (Shift-click to snap to 0.1s)" },
            stepDownTolerance: { ja: "値を0.01減らす（shift＋クリックで0.1の倍数へ）", en: "Decrease by 0.01 (Shift-click to snap to 0.1s)" },
            removeSameAnchors: { ja: "連続して同じ座標にあるアンカーポイントを1つに統合します（離れた位置の同座標は対象外）。", en: "Merges consecutive anchors that share the same coordinates (non-adjacent duplicates are ignored)." },
            removeAnchors:     { ja: "前後のアンカーと一直線上にある冗長なアンカーポイントを削除します。", en: "Removes redundant anchors that lie on a straight line between their neighbors." },
            removeHandles:     { ja: "直線とみなせる区間のハンドルをアンカーに戻します（見た目を変えずにハンドルを整理）。", en: "Resets handles on segments that are effectively straight (tidies handles without changing appearance)." },
            tolAnchor:         { ja: "値が大きいほど、より緩く「直線上」と判定して多くのアンカーを削除します（0.01〜3.00）。", en: "Higher values treat more anchors as collinear and remove more of them (0.01–3.00)." },
            tolHandle:         { ja: "値が大きいほど、より緩く「直線」と判定して多くのハンドルを戻します（0.01〜3.00）。", en: "Higher values treat more segments as straight and reset more handles (0.01–3.00)." },
            addAnchors:        { ja: "各セグメントの中間に1点ずつアンカーポイントを追加します。", en: "Adds one anchor at the midpoint of each segment." },
            addExtremePoints:  { ja: "曲線の上下左右の端（水平・垂直の接線位置＝極点）にアンカーポイントを追加します。", en: "Adds anchors at the curve's extrema (points of horizontal/vertical tangency)." },
            splitAtAnchors:    { ja: "各セグメントを独立したオープンパスに分割します（元のパスは削除）。", en: "Splits each segment into a separate open path (the original path is removed)." },
            fillHoles:         { ja: "複合パスを解除して合体し、穴（マド）を埋めます。選択に複合パスが無い場合は使用できません。", en: "Releases the compound path and unites it to fill holes. Unavailable when the selection has no compound path." }
        }
    };

    // =========================================
    // UIレイアウトの共通設定 / Shared UI layout
    // =========================================

    /* ウィンドウ・パネルの余白と間隔 / Window & panel margins and spacing */
    var WINDOW_MARGINS = 16;               /* ウィンドウ外周の余白 / window margin */
    var WINDOW_SPACING = 12;               /* ウィンドウ内の要素間隔 / window spacing */
    var PANEL_MARGINS  = [16, 20, 16, 12]; /* パネル余白 [左,上,右,下] / panel margins */
    var PANEL_SPACING  = 8;                /* パネル内の要素間隔 / panel spacing */

    /**
     * ウィンドウに共通レイアウトを適用します。
     * @param {Window} win - 対象のダイアログウィンドウ。
     * @param {number} [spacing] - 要素間隔（省略時は WINDOW_SPACING）。
     * @returns {void}
     */
    function setupWindow(win, spacing) {
        win.orientation = "column";
        win.alignChildren = "fill";
        win.margins = WINDOW_MARGINS;
        win.spacing = (typeof spacing === "number") ? spacing : WINDOW_SPACING;
    }

    /**
     * パネル（タブを含む）に共通レイアウトを適用します。
     * @param {Panel} panel - 対象のパネル。
     * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）。
     * @returns {void}
     */
    function setupPanel(panel, spacing) {
        panel.orientation = "column";
        panel.alignChildren = ["fill", "top"];
        panel.alignment = "fill";
        panel.margins = PANEL_MARGINS;
        panel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
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

    (function () {
        // --- shared settings / helpers ---
        /* 共通設定とヘルパー / Shared settings and helpers */
        /* 許容誤差は冒頭「ユーザー設定 / User Settings」で定義（ダイアログで上書き）
           Tolerances are defined in the "User Settings" section at the top (overridden by the dialog). */

        /**
         * パス情報の集計値。
         * @typedef {Object} InfoCounts
         * @property {number} paths - パス数。
         * @property {number} anchors - アンカーポイント数。
         * @property {number} handles - ハンドル数。
         */

        /**
         * シミュレーション用のパスモデル（DOMを書き換えず増減を予測する軽量表現）。
         * @typedef {Object} PathModel
         * @property {boolean} closed - クローズドパスなら true。
         * @property {Array<PathModelPoint>} pts - アンカー点の配列。
         */

        /**
         * パスモデルの1点。各座標は [x, y] の数値配列。
         * @typedef {Object} PathModelPoint
         * @property {Array<number>} a - アンカー座標。
         * @property {Array<number>} l - 左方向線（leftDirection）座標。
         * @property {Array<number>} r - 右方向線（rightDirection）座標。
         * @property {PointType} t - アンカーの種類（書き戻し時に復元）。
         */

        /**
         * 実処理中の例外を最小限ログ出力します（UI位置の保存・復元系は従来どおり silent）。
         * @param {string} context - エラー発生箇所を示すラベル。
         * @param {Error} e - 捕捉した例外。
         * @returns {void}
         */
        function logProcessError(context, e) {
            try {
                $.writeln("[PathCleanupTool] " + context + ": " + e);
            } catch (ignored) {
                // ignore logging failure
            }
        }

        /** ドキュメントが1つ以上開かれていれば true。 */
        function hasDocument() {
            return app.documents.length > 0;
        }

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
         * @param {Object} stepOptions - step（増減量）/ shiftStep（shift のときの倍数）/ min / max / integer / unit（例 " mm"）/
         *     upTooltipKey・downTooltipKey（∧∨の説明の LABELS キー）/ onStep(numberInput)
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
            /* shiftStep を変えた欄は専用の説明（upTooltipKey / downTooltipKey）を使う / fields with their own shiftStep bring their own tooltips */
            var upTooltip = stepOptions.upTooltipKey || (stepOptions.integer ? "tooltip.stepUpInteger" : "tooltip.stepUp");
            var downTooltip = stepOptions.downTooltipKey || (stepOptions.integer ? "tooltip.stepDownInteger" : "tooltip.stepDown");
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
         * @param {Object} stepOptions - step（通常の増減量。省略時は 1）/ shiftStep（shift のときの倍数。省略時は STEPPER_SHIFT_MULTIPLE）/ integer
         * @returns {number} 増減した値（下限・上限は未適用）
         */
        function computeSteppedValue(value, direction, stepOptions) {
            var keyState = ScriptUI.environment.keyboardState;
            if (keyState.shiftKey) return snapStepperToNextMultiple(value, stepOptions.shiftStep || STEPPER_SHIFT_MULTIPLE, direction);
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

        /**
         * ドキュメント・選択の有無を検証し、選択配列を返します。無ければ警告を表示。
         * @returns {?Array<PageItem>} 選択オブジェクト配列。ドキュメント無し／未選択なら null。
         */
        function getSelectionOrAlert() {
            if (!hasDocument()) {
                alert(getLabel('alert.noDocument'));
                return null;
            }
            var doc = app.activeDocument;
            var currentSelection = doc.selection;
            if (!(currentSelection instanceof Array) || currentSelection.length === 0) {
                alert(getLabel('alert.needSelection'));
                return null;
            }
            return currentSelection;
        }

        /**
         * 3点が一直線上にあるかを、点Bから直線A-Cまでの垂直距離で判定します。
         * @param {Array<number>} pointA - 点A [x, y]。
         * @param {Array<number>} pointB - 点B（直線上にあるか判定する点） [x, y]。
         * @param {Array<number>} pointC - 点C [x, y]。
         * @param {number} [tolerance] - 距離の許容誤差（pt、省略時は TOL_ANCHOR_COLLINEAR）。
         * @returns {boolean} 一直線上とみなせれば true。
         */
        function isCollinear(pointA, pointB, pointC, tolerance) {
            /* 0 が既定値に化けないよう != null で判定する */
            tolerance = (tolerance != null) ? tolerance : TOL_ANCHOR_COLLINEAR;
            /* 外積をそのまま比較するとセグメント長に比例して感度が変わり、
               ハンドル側の許容誤差（垂直距離）と意味がずれるため距離で揃える */
            return isPointOnLineByDistance(pointA, pointC, pointB, tolerance);
        }

        /**
         * 2点がほぼ同一座標かを判定します（マンハッタン距離で比較）。
         * @param {Array<number>} pointA - 点A [x, y]。
         * @param {Array<number>} pointB - 点B [x, y]。
         * @param {number} [tolerance] - 許容誤差（省略時は TOL_SAMEPOINT）。
         * @returns {boolean} 同一とみなせれば true。
         */
        function samePoint(pointA, pointB, tolerance) {
            /* 0 が既定値に化けないよう != null で判定する */
            tolerance = (tolerance != null) ? tolerance : TOL_SAMEPOINT;
            return (Math.abs(pointA[0] - pointB[0]) + Math.abs(pointA[1] - pointB[1])) < tolerance;
        }

        /**
         * 点が線分の延長を含む直線上にあるかを、点から直線までの垂直距離で判定します。
         * 線分が極端に短い場合は samePoint にフォールバックします。
         * @param {Array<number>} lineStart - 直線の始点 [x, y]。
         * @param {Array<number>} lineEnd - 直線の終点 [x, y]。
         * @param {Array<number>} testPoint - 判定する点 [x, y]。
         * @param {number} [tolerance] - 距離の許容誤差（pt、省略時は TOL_HANDLE_COLLINEAR）。
         * @returns {boolean} 直線上とみなせれば true。
         */
        function isPointOnLineByDistance(lineStart, lineEnd, testPoint, tolerance) {
            tolerance = (tolerance != null) ? tolerance : TOL_HANDLE_COLLINEAR;

            var abx = lineEnd[0] - lineStart[0];
            var aby = lineEnd[1] - lineStart[1];
            var len = Math.sqrt(abx * abx + aby * aby);

            if (len < 1e-9) {
                // a と b がほぼ同一点
                return samePoint(lineStart, testPoint, Math.max(TOL_SAMEPOINT, tolerance));
            }

            // cross product magnitude / |AB| = perpendicular distance
            var apx = testPoint[0] - lineStart[0];
            var apy = testPoint[1] - lineStart[1];
            var area2 = abx * apy - aby * apx; // signed
            var dist = Math.abs(area2) / len;
            return dist <= tolerance;
        }

        /**
         * オブジェクトがロック／非表示か（親・レイヤーを遡って）を判定します。
         * @param {PageItem} item - 判定対象のオブジェクト。
         * @returns {boolean} ロックまたは非表示なら true。
         */
        function isSkippableItem(item) {
            var currentItem = item;
            while (currentItem) {
                try {
                    // PageItem / GroupItem / PathItem etc.
                    if (currentItem.locked === true) return true;
                    if (currentItem.hidden === true) return true;

                    // Layer
                    if (currentItem.typename === "Layer") {
                        if (currentItem.locked === true) return true;
                        if (currentItem.visible === false) return true;
                    }

                    // If the item has a layer reference, also respect it
                    if (currentItem.layer) {
                        try {
                            if (currentItem.layer.locked === true) return true;
                            if (currentItem.layer.visible === false) return true;
                        } catch (e) { }
                    }
                } catch (e) {
                    // ignore property access errors
                }

                // Walk up
                try {
                    currentItem = currentItem.parent;
                } catch (e) {
                    break;
                }

                // Stop when reaching document-like root
                if (!currentItem || currentItem.typename === "Document") break;
            }
            return false;
        }

        /**
         * 選択配列から処理対象の PathItem 一覧を収集して返します。
         * 収集時に isSkippableItem() でロック／非表示を除外済みのため、
         * targets を受け取る各関数（*OnTargets / *ForTargets）は再判定しない。
         * @param {Array<PageItem>} selection - 選択オブジェクト配列。
         * @returns {Array<PathItem>} ロック／非表示を除外した対象 PathItem 配列。
         */
        function getTargetPathItemsFromSelection(selection) {
            /* グループ（クリップグループを含む）・複合パスの中まで集める / Walk into groups (clip groups too) and compound paths */
            var collectedPaths = collectSelectionPathItems(selection, { skipLocked: true, skipHidden: true, unique: false });
            /* 親やレイヤーがロック／非表示のものも外す / Also drop items whose parent or layer is locked or hidden */
            var pathItems = [];
            for (var i = 0; i < collectedPaths.length; i++) {
                if (!isSkippableItem(collectedPaths[i])) pathItems.push(collectedPaths[i]);
            }
            return pathItems;
        }

        /**
         * オブジェクト（およびグループ内）に、マド（穴）を持つ複合パスが含まれるかを再帰判定します。
         * サブパスが2つ以上の CompoundPathItem を「マドあり」とみなします。ロック／非表示は対象外。
         * @param {PageItem} item - 判定対象のオブジェクト。
         * @returns {boolean} マドを持つ複合パスがあれば true。
         */
        function itemHasHoles(item) {
            if (!item || isSkippableItem(item)) return false;
            try {
                if (item.typename === 'CompoundPathItem') {
                    return item.pathItems.length >= 2;
                }
                if (item.typename === 'GroupItem') {
                    for (var i = 0; i < item.pageItems.length; i++) {
                        if (itemHasHoles(item.pageItems[i])) return true;
                    }
                }
            } catch (e) {
                logProcessError('itemHasHoles', e);
            }
            return false;
        }

        /**
         * 選択内にマド（穴）を持つ複合パスが1つでもあるかを判定します（マド埋めの有効判定に使用）。
         * @param {Array<PageItem>} selection - 選択オブジェクト配列。
         * @returns {boolean} マドを持つ複合パスがあれば true。
         */
        function selectionHasHoles(selection) {
            if (!selection || !selection.length) return false;
            for (var i = 0; i < selection.length; i++) {
                if (itemHasHoles(selection[i])) return true;
            }
            return false;
        }

        /**
         * 指定 targets からパス数・アンカー数・ハンドル数を集計します。
         * @param {Array<PathItem>} targets - 対象の PathItem 配列。
         * @returns {InfoCounts} 集計結果。
         */
        function getInfoCountsFromTargets(targets) {
            var info = { paths: 0, anchors: 0, handles: 0 };
            if (!targets || !targets.length) return info;

            for (var i = 0; i < targets.length; i++) {
                var item = targets[i];
                if (!item) continue;
                info.paths++;

                var pts = item.pathPoints;
                var n = pts.length;
                info.anchors += n;

                for (var k = 0; k < n; k++) {
                    var pt = pts[k];
                    if (!samePoint(pt.leftDirection, pt.anchor, TOL_SAMEPOINT)) info.handles++;
                    if (!samePoint(pt.rightDirection, pt.anchor, TOL_SAMEPOINT)) info.handles++;
                }
            }

            return info;
        }

        // =========================================
        // パスモデル / Path model
        // =========================================
        /* 情報パネルの予測と実処理は、どちらもこのモデル関数を通す。
           アルゴリズムをここ1箇所に集約し、実処理は結果を applyModelToPath() で書き戻す。
           Both the preview and the actual processing go through these model functions. */

        /** PathPoint pt をモデル点 PathModelPoint（{a, l, r, t}）に複製します。 */
        function clonePointModel(pt) {
            return {
                a: [pt.anchor[0], pt.anchor[1]],
                l: [pt.leftDirection[0], pt.leftDirection[1]],
                r: [pt.rightDirection[0], pt.rightDirection[1]],
                t: pt.pointType
            };
        }

        /**
         * PathItem を編集可能なパスモデルに複製します。
         * @param {PathItem} item - 複製元のパス。
         * @returns {PathModel} 複製したパスモデル。
         */
        function clonePathModel(item) {
            var pts = item.pathPoints;
            var out = [];
            for (var i = 0; i < pts.length; i++) {
                out.push(clonePointModel(pts[i]));
            }
            return {
                closed: !!item.closed,
                pts: out
            };
        }

        /** 2つのモデル座標 a, b がほぼ同一（TOL_SAMEPOINT 基準）なら true。 */
        function samePointModel(a, b) {
            return samePoint(a, b, TOL_SAMEPOINT);
        }

        /** モデル点 p にハンドルが無い（直線的：左右ともアンカーと一致）なら true。 */
        function isStraightPointModel(p) {
            return samePointModel(p.a, p.l) && samePointModel(p.a, p.r);
        }

        /**
         * セグメント (p0 → p1) が見た目として直線かを判定します。
         * 両端アンカーを結ぶ直線に p0 右ハンドル・p1 左ハンドルが近ければ直線とみなす。
         * @param {PathModelPoint} p0 - 始点モデル点。
         * @param {PathModelPoint} p1 - 終点モデル点。
         * @param {number} [tolerance] - 距離の許容誤差（pt、省略時は TOL_HANDLE_COLLINEAR）。
         * @returns {boolean} 直線とみなせれば true。
         */
        function isStraightSegmentModel(p0, p1, tolerance) {
            /* 0 が既定値に化けないよう != null で判定する */
            tolerance = (tolerance != null) ? tolerance : TOL_HANDLE_COLLINEAR;
            return isPointOnLineByDistance(p0.a, p1.a, p0.r, tolerance) &&
                isPointOnLineByDistance(p0.a, p1.a, p1.l, tolerance);
        }

        /**
         * 同一座標のアンカーポイント（重複点）を削除します。
         * 連続して同座標のアンカーのみ対象（離れた位置の同座標は対象外）。オープンパスの端点は削除しない。
         * 削除する点の右方向線は残す点へ引き継ぎ、見た目が変わらないようにする。
         * @param {PathModel} pathM - 対象のパスモデル（破壊的に更新）。
         * @returns {number} 削除したアンカー数。
         */
        function removeDuplicateAnchorsModel(pathM) {
            var removed = 0;
            var pts = pathM.pts;
            var isClosed = pathM.closed;
            var n = pts.length;
            if (n < 2) return 0;

            /* 削除でインデックスがずれるのを防ぐため、後ろから走査する
               オープンパスは端点を除外、クローズドパスは全点が対象（最低2点は残す） */
            var start = isClosed ? (n - 1) : (n - 2);
            var end = isClosed ? 0 : 1;

            for (var i = start; i >= end; i--) {
                var curLen = pts.length;
                if (curLen <= 2) break;

                if (i > curLen - 1) i = curLen - 1;
                if (!isClosed && (i <= 0 || i >= curLen - 1)) continue;

                var prevIndex = isClosed ? ((i - 1 + curLen) % curLen) : (i - 1);
                if (prevIndex < 0 || prevIndex > curLen - 1) continue;

                if (samePointModel(pts[prevIndex].a, pts[i].a)) {
                    /* 削除する点が「出ていく側」のハンドルを持っているため、残す点へ引き継ぐ。
                       引き継がないと重複解消のたびに後続セグメントの曲線が直線に潰れる */
                    pts[prevIndex].r = [pts[i].r[0], pts[i].r[1]];
                    /* 入り側と出側で別々のハンドルを持つ点はスムーズでは表現できない */
                    pts[prevIndex].t = PointType.CORNER;
                    pts.splice(i, 1);
                    removed++;
                }
            }

            return removed;
        }

        /**
         * 直線上の冗長なアンカーポイントを削除します。
         * 対象は「その点自身にハンドルが無い」「前後のセグメントがどちらも直線」
         * 「前後アンカーと一直線上にある」の3条件を満たす点のみ。オープンパスの端点は削除しない。
         * @param {PathModel} pathM - 対象のパスモデル（破壊的に更新）。
         * @returns {number} 削除したアンカー数。
         */
        function removeRedundantAnchorsModel(pathM) {
            var removed = 0;
            var pts = pathM.pts;
            var isClosed = pathM.closed;

            if (pts.length < 3) return 0;

            /* オープンパスは端点を削除しない。削除でインデックスがずれるため後ろから走査する */
            var startIndex = isClosed ? (pts.length - 1) : (pts.length - 2);
            var endIndex = isClosed ? 0 : 1;

            for (var i = startIndex; i >= endIndex; i--) {
                var currentLen = pts.length;
                if (currentLen < 3) break;

                // 削除でインデックスが範囲外になった場合に備えて補正
                if (i > currentLen - 1) i = currentLen - 1;
                if (i < endIndex) break;

                var pA = pts[(i - 1 + currentLen) % currentLen];
                var pB = pts[i];
                var pC = pts[(i + 1) % currentLen];

                /* B自身にハンドルが無いだけでは足りない。pA の右ハンドル／pC の左ハンドルが
                   線から外れていると前後のセグメントは曲線なので、B を消すと形が変わる。
                   判定はアンカー側の許容誤差で統一する（ハンドル側スライダーの影響を受けない） */
                if (isStraightPointModel(pB) &&
                    isStraightSegmentModel(pA, pB, TOL_ANCHOR_COLLINEAR) &&
                    isStraightSegmentModel(pB, pC, TOL_ANCHOR_COLLINEAR) &&
                    isCollinear(pA.a, pB.a, pC.a, TOL_ANCHOR_COLLINEAR)) {
                    pts.splice(i, 1);
                    removed++;
                }
            }

            return removed;
        }

        /**
         * 直線になっているベジェ区間のハンドルをアンカーに戻します。
         * オープンパスの最初のセグメントの始点側／最後のセグメントの終点側ハンドルは触らない。
         * @param {PathModel} pathM - 対象のパスモデル（破壊的に更新）。
         * @returns {number} リセットしたハンドル数（左右それぞれ1カウント）。
         */
        function removeRedundantHandlesModel(pathM) {
            var changed = 0;
            var pts = pathM.pts;
            var len = pts.length;
            if (len < 2) return 0;

            var isClosed = pathM.closed;
            var segCount = isClosed ? len : (len - 1);

            for (var i = 0; i < segCount; i++) {
                var p0 = pts[i];
                var p1 = pts[(i + 1) % len];

                if (!isStraightSegmentModel(p0, p1)) continue;

                // オープンパス端点のハンドルは触らない / Do not modify endpoint handles for open paths
                var isOpen = !isClosed;
                var isFirstSeg = isOpen && (i === 0);
                var isLastSeg = isOpen && (i === (segCount - 1));

                /* 片側だけハンドルが無い点はスムーズでは表現できないため、コーナーに落とす
                   （SMOOTH のまま書き戻すと Illustrator 側でハンドルを戻される場合がある） */
                if (!isFirstSeg && !samePointModel(p0.r, p0.a)) {
                    p0.r = [p0.a[0], p0.a[1]];
                    p0.t = PointType.CORNER;
                    changed++;
                }

                if (!isLastSeg && !samePointModel(p1.l, p1.a)) {
                    p1.l = [p1.a[0], p1.a[1]];
                    p1.t = PointType.CORNER;
                    changed++;
                }
            }

            return changed;
        }

        /**
         * モデル上のハンドル数を数えます（左右それぞれ1カウント）。
         * @param {PathModel} pathM - 対象のパスモデル。
         * @returns {number} ハンドル数。
         */
        function countHandlesModel(pathM) {
            var c = 0;
            var pts = pathM.pts;
            for (var i = 0; i < pts.length; i++) {
                var p = pts[i];
                if (!samePointModel(p.l, p.a)) c++;
                if (!samePointModel(p.r, p.a)) c++;
            }
            return c;
        }

        /**
         * モデルにクリーンアップを適用します。
         * 実行順は 重複→冗長→ハンドル→冗長→重複。ハンドルを直線区間から外すと
         * 「ハンドルがあるから消せない」と判定されていたアンカーが新たに対象になるため折り返す。
         * @param {PathModel} pathM - 対象のパスモデル（破壊的に更新）。
         * @param {boolean} doSameAnchors - 重複アンカー削除を含める場合は true。
         * @param {boolean} doAnchors - 直線上の冗長アンカー削除を含める場合は true。
         * @param {boolean} doHandles - 直線区間のハンドル削除を含める場合は true。
         * @returns {number} 削除・リセットした総数（0 なら変化なし）。
         */
        function runCleanupOnModel(pathM, doSameAnchors, doAnchors, doHandles) {
            var changed = 0;
            if (doSameAnchors) changed += removeDuplicateAnchorsModel(pathM);
            if (doAnchors) changed += removeRedundantAnchorsModel(pathM);
            if (doHandles) changed += removeRedundantHandlesModel(pathM);
            if (doAnchors) changed += removeRedundantAnchorsModel(pathM);
            if (doSameAnchors) changed += removeDuplicateAnchorsModel(pathM);
            return changed;
        }

        /**
         * モデルの内容を PathItem へ書き戻します。
         * 余った点を末尾から削除したうえで全点を上書きします。
         * パスを作り直さず同一 PathItem を保持するため、アピアランスは維持されます。
         * @param {PathItem} item - 書き戻し先のパス。
         * @param {PathModel} pathM - 書き戻すパスモデル。
         * @returns {void}
         */
        function applyModelToPath(item, pathM) {
            var modelPoints = pathM.pts;
            if (modelPoints.length < 2) return;

            try {
                var pts = item.pathPoints;
                while (pts.length > modelPoints.length) {
                    pts[pts.length - 1].remove();
                }
                for (var i = 0; i < modelPoints.length; i++) {
                    var pt = pts[i];
                    var m = modelPoints[i];
                    pt.anchor = m.a;
                    pt.pointType = m.t;
                    pt.leftDirection = m.l;
                    pt.rightDirection = m.r;
                }
            } catch (e) {
                logProcessError('applyModelToPath', e);
            }
        }

        /**
         * 指定 targets にクリーンアップを実行します（モデル上で処理してから書き戻す）。
         * @param {Array<PathItem>} targets - 対象の PathItem 配列。
         * @param {boolean} doSameAnchors - 重複アンカー削除を行うか。
         * @param {boolean} doAnchors - 直線上の冗長アンカー削除を行うか。
         * @param {boolean} doHandles - 直線区間のハンドル削除を行うか。
         * @returns {number} 変化のあったパス数。
         */
        function runCleanupOnTargets(targets, doSameAnchors, doAnchors, doHandles) {
            if (!targets || !targets.length) return 0;

            var changedPaths = 0;
            for (var i = 0; i < targets.length; i++) {
                var item = targets[i];
                if (!item) continue;

                try {
                    var pathModel = clonePathModel(item);
                    if (runCleanupOnModel(pathModel, doSameAnchors, doAnchors, doHandles) === 0) continue;
                    applyModelToPath(item, pathModel);
                    changedPaths++;
                } catch (e) {
                    logProcessError('runCleanupOnTargets', e);
                }
            }

            return changedPaths;
        }

        /**
         * クリーンアップ実行後のアンカー数・ハンドル数を予測します。
         * DOM は書き換えず、複製したモデル上で実処理とまったく同じ関数・同じ実行順を通します。
         * @param {Array<PathItem>} targets - 対象の PathItem 配列。
         * @param {boolean} doSameAnchors - 重複アンカー削除を含める場合は true。
         * @param {boolean} doAnchors - 直線上の冗長アンカー削除を含める場合は true。
         * @param {boolean} doHandles - 直線区間のハンドル削除を含める場合は true。
         * @param {InfoCounts} [infoNow] - 実行前の集計値（省略時はここで集計）。ダイアログ表示中は変化しないため使い回す。
         * @returns {{paths: number, anchorsNow: number, anchorsAfter: number, handlesNow: number, handlesAfter: number}} 予測結果。
         */
        function getPredictedInfoCountsForTargets(targets, doSameAnchors, doAnchors, doHandles, infoNow) {
            if (!infoNow) infoNow = getInfoCountsFromTargets(targets);

            var anchorsAfterTotal = 0;
            var handlesAfterTotal = 0;

            for (var t = 0; t < (targets ? targets.length : 0); t++) {
                var item = targets[t];
                if (!item) continue;

                var pathModel = clonePathModel(item);
                runCleanupOnModel(pathModel, doSameAnchors, doAnchors, doHandles);

                anchorsAfterTotal += pathModel.pts.length;
                handlesAfterTotal += countHandlesModel(pathModel);
            }

            return {
                paths: infoNow.paths,
                anchorsNow: infoNow.anchors,
                anchorsAfter: anchorsAfterTotal,
                handlesNow: infoNow.handles,
                handlesAfter: handlesAfterTotal
            };
        }

        /**
         * 「変換」タブ（変換・分割）の予測情報を取得します。
         * corner=全ハンドル削除／smooth=全アンカーに左右ハンドル付与／add=各セグメントに1点追加／
         * extreme=各曲線セグメントの極点数を実測して加算／split=各セグメントを独立パス化／
         * fillHoles=結果を事前算出できないため未確定（"-"）。
         * @param {Array<PathItem>} targets - 対象の PathItem 配列。
         * @param {string} mode - 変換モード（'smooth' | 'corner' | 'add' | 'extreme' | 'split' | 'fillHoles'）。
         * @param {InfoCounts} [infoNow] - 実行前の集計値（省略時はここで集計）。ダイアログ表示中は変化しないため使い回す。
         * @returns {{paths: number, pathsAfter: (number|string), anchorsNow: number, anchorsAfter: (number|string), handlesNow: number, handlesAfter: (number|string)}} 予測結果。
         */
        function getPredictedInfoForConvert(targets, mode, infoNow) {
            if (!infoNow) infoNow = getInfoCountsFromTargets(targets);
            var pathsAfter = infoNow.paths;
            var anchorsAfter = infoNow.anchors;
            var handlesAfter = infoNow.handles;

            if (!targets || !targets.length) {
                return { paths: 0, pathsAfter: 0, anchorsNow: 0, anchorsAfter: 0, handlesNow: 0, handlesAfter: 0 };
            }

            if (mode === 'corner') {
                // 全ハンドル削除
                handlesAfter = 0;
            } else if (mode === 'smooth') {
                /* 全アンカーに左右ハンドルが付く。ただしオープンパスは端点の外側
                   （始点の左・終点の右）にハンドルが付かないため、1本につき2つ引く */
                var openPaths = 0;
                for (var s = 0; s < targets.length; s++) {
                    if (targets[s] && !targets[s].closed) openPaths++;
                }
                handlesAfter = infoNow.anchors * 2 - openPaths * 2;
            } else if (mode === 'add') {
                // 各セグメントに1点追加
                var totalSegs = 0;
                for (var i = 0; i < targets.length; i++) {
                    var item = targets[i];
                    if (!item) continue;
                    var n = item.pathPoints.length;
                    totalSegs += item.closed ? n : (n - 1);
                }
                anchorsAfter = infoNow.anchors + totalSegs;
                handlesAfter = '-';
            } else if (mode === 'split') {
                // 各セグメントが独立パスに
                var totalSegs2 = 0;
                for (var j = 0; j < targets.length; j++) {
                    var item2 = targets[j];
                    if (!item2) continue;
                    var n2 = item2.pathPoints.length;
                    totalSegs2 += item2.closed ? n2 : (n2 - 1);
                }
                pathsAfter = totalSegs2;
                anchorsAfter = totalSegs2 * 2;
                handlesAfter = '-';
            } else if (mode === 'extreme') {
                // 各曲線セグメントの極点数を実測して加算
                anchorsAfter = infoNow.anchors + countExtremaForTargets(targets);
                handlesAfter = '-';
            } else if (mode === 'fillHoles') {
                // パスファインダー結果は事前に正確に出せないため未確定
                pathsAfter = '-';
                anchorsAfter = '-';
                handlesAfter = '-';
            }

            return {
                paths: infoNow.paths,
                pathsAfter: pathsAfter,
                anchorsNow: infoNow.anchors,
                anchorsAfter: anchorsAfter,
                handlesNow: infoNow.handles,
                handlesAfter: handlesAfter
            };
        }

        // --- UI ---
        /**
         * showDialog() の戻り値。
         * @typedef {Object} DialogResult
         * @property {boolean} ok - OK で閉じたら true。
         * @property {string} activeMode - 実行タブ（'process' | 'other'）。
         * @property {boolean} doRemoveSameAnchors - 重複アンカー削除を行うか。
         * @property {boolean} doRemoveAnchors - 直線上の冗長アンカー削除を行うか。
         * @property {boolean} doRemoveHandles - 直線区間のハンドル削除を行うか。
         * @property {string} convertMode - 変換モード（'smooth' | 'corner' | 'add' | 'extreme' | 'split' | 'fillHoles'）。
         */

        /**
         * メインダイアログを表示し、ユーザーの選択結果を返します。
         * @param {Array<PathItem>} frozenTargets - ダイアログ表示時点で固定した対象パス（予測表示に使用）。
         * @param {boolean} hasHoles - 選択内にマド（穴）を持つ複合パスがあるか（false のとき「マド埋め」を無効化）。
         * @returns {DialogResult} 実行内容を表す結果オブジェクト。
         */
        function showDialog(frozenTargets, hasHoles) {
            /* 実行前の集計はダイアログ表示中に変化しないため、1回だけ数えて使い回す
               （スライダー操作のたびに全パスを数え直さないようにするため） */
            var frozenInfoNow = getInfoCountsFromTargets(frozenTargets);

            /**
             * 現在アクティブなタブに対応する実行モードを返します。
             * @returns {string} 'other'（「変換」タブ）または 'process'（「削除対象」タブ）。
             */
            function getActiveMode() {
                return (tabbedPanel.selection === tabOther) ? 'other' : 'process';
            }
            var dlg = new Window('dialog', getLabel('dialog.title') + ' ' + SCRIPT_VERSION);
            setupWindow(dlg);

            var infoPanel = dlg.add('panel', undefined, getLabel('panel.info'));
            setupPanel(infoPanel);

            /**
             * 情報パネルに「ラベル：値」の行を追加します。
             * @param {string} labelKey - ラベルの LABELS ドット区切りキー。
             * @returns {StaticText} 値表示用の StaticText（後で text を更新する）。
             */
            function addInfoRow(labelKey) {
                var infoRow = infoPanel.add('group');
                infoRow.orientation = 'row';
                infoRow.alignChildren = ['left', 'center'];

                var rowLabel = infoRow.add('statictext', undefined, labelText(labelKey));
                rowLabel.characters = 13;
                rowLabel.justify = 'right';

                /* characters は最小幅なので小さく取り、「128 → 96」形式は fill で
                   行の余りを使って表示する（ここでウィンドウ幅を広げないため） */
                var rowValue = infoRow.add('statictext', undefined, '0');
                rowValue.characters = 4;
                rowValue.alignment = ['fill', 'center'];
                return rowValue;
            }

            /**
             * 許容誤差の入力文字列を数値に変換します（角括弧・全角括弧・空白を除去）。
             * @param {string} text - 入力文字列。
             * @returns {number} 変換した数値。空文字なら NaN。
             */
            function parseToleranceText(text) {
                if (!text) return NaN;
                // accept formats like "[0.01]", "0.01", and full-width brackets
                text = String(text).replace(/\[/g, '').replace(/\]/g, '').replace(/［/g, '').replace(/］/g, '').replace(/\s/g, '');
                return parseFloat(text);
            }

            var pathCountValue = addInfoRow('label.pathCount');
            var anchorCountValue = addInfoRow('label.anchorCount');
            var handleCountValue = addInfoRow('label.handleCount');

            /** 「beforeValue → afterValue」の矢印連結文字列を生成します。 */
            function formatArrow(beforeValue, afterValue) {
                return String(beforeValue) + " → " + String(afterValue);
            }

            /**
             * 「削除対象」タブの現在のチェック状態で情報パネルの予測値を更新します。
             * @returns {void}
             */
            function refreshInfoPreview() {
                var predictedInfo = getPredictedInfoCountsForTargets(
                    frozenTargets,
                    removeSameAnchorsCheckbox.value,
                    removeAnchorsCheckbox.value,
                    removeHandlesCheckbox.value,
                    frozenInfoNow);
                pathCountValue.text = String(predictedInfo.paths);
                anchorCountValue.text = formatArrow(predictedInfo.anchorsNow, predictedInfo.anchorsAfter);
                handleCountValue.text = formatArrow(predictedInfo.handlesNow, predictedInfo.handlesAfter);
                updateOkEnabled();
            }

            /**
             * 「変換」タブの予測値で情報パネルを更新します。
             * @param {string} mode - 変換モード（'smooth' | 'corner' | 'add' | 'extreme' | 'split' | 'fillHoles'）。
             * @returns {void}
             */
            function refreshInfoForConvert(mode) {
                var predictedInfo = getPredictedInfoForConvert(frozenTargets, mode, frozenInfoNow);
                pathCountValue.text = formatArrow(predictedInfo.paths, predictedInfo.pathsAfter);
                anchorCountValue.text = formatArrow(predictedInfo.anchorsNow, predictedInfo.anchorsAfter);
                handleCountValue.text = formatArrow(predictedInfo.handlesNow, predictedInfo.handlesAfter);
                updateOkEnabled();
            }

            /**
             * 実行できる内容が選ばれているかに応じて OK ボタンの有効／無効を切り替えます。
             * 「削除対象」タブでチェックが1つも無いと実行しても何も起きないため無効にします。
             * @returns {void}
             */
            function updateOkEnabled() {
                /* ボタン生成前にも予測表示が走るため、その時点では何もしない */
                if (!btnOK) return;
                if (tabbedPanel.selection === tabOther) {
                    btnOK.enabled = true;
                    return;
                }
                btnOK.enabled = removeSameAnchorsCheckbox.value ||
                    removeAnchorsCheckbox.value ||
                    removeHandlesCheckbox.value;
            }

            /**
             * 「変換」タブで選択中のラジオボタンから変換モードを返します。
             * どれも選択されていない場合は、無効化されることのない 'smooth' に倒す
             * （'fillHoles' は選択内容によってディム表示になるためフォールバックに使わない）。
             * @returns {string} 'smooth' | 'corner' | 'add' | 'extreme' | 'split' | 'fillHoles'。
             */
            function getSelectedConvertMode() {
                if (cornerRadio.value) return 'corner';
                if (addAnchorsRadio.value) return 'add';
                if (extremePointsRadio.value) return 'extreme';
                if (splitRadio.value) return 'split';
                if (fillHolesRadio.value) return 'fillHoles';
                return 'smooth';
            }

            /**
             * アクティブなタブに応じて情報パネルの予測表示を切り替えます。
             * @returns {void}
             */
            function refreshByActiveTab() {
                if (tabbedPanel.selection === tabOther) {
                    refreshInfoForConvert(getSelectedConvertMode());
                } else {
                    refreshInfoPreview();
                }
            }

            var tabbedPanel = dlg.add('tabbedpanel');
            tabbedPanel.alignChildren = ['fill', 'top'];

            // --- Tab 1: 削除対象 ---
            var tabProcess = tabbedPanel.add('tab', undefined, getLabel('tab.process'));
            setupPanel(tabProcess);
            /* タブ内の右余白を詰める（PANEL_MARGINS の右16→6）。サブプロパティ代入は反映されないため配列で上書き */
            tabProcess.margins = [16, 20, 6, 12];

            var removeSameAnchorsCheckbox = tabProcess.add('checkbox', undefined, getLabel('checkbox.removeSameAnchors'));
            removeSameAnchorsCheckbox.helpTip = getLabel('tooltip.removeSameAnchors');
            removeSameAnchorsCheckbox.value = true;

            /* 許容誤差のUIはアンカー用とハンドル用で中身が同一のため、1つの生成関数にまとめる
               The anchor and handle tolerance blocks are identical, so one builder creates both. */
            var TOL_MIN = 0.01;
            var TOL_MAX = 3;

            /* アンカー数が多いと onChanging 1回ごとの全パス再計算が重くなるため、
               しきい値を超える選択ではスライダーを離したとき（onChange）だけ予測を更新する */
            var HEAVY_ANCHOR_LIMIT = 2000;
            var isHeavySelection = frozenInfoNow.anchors > HEAVY_ANCHOR_LIMIT;

            /**
             * 許容誤差を有効範囲（0.01〜3.00、小数2桁）に丸めます。
             * @param {number} toleranceValue - 入力値。
             * @param {number} fallbackValue - NaN のときに返す値。
             * @returns {number} 丸めた許容誤差。
             */
            function clampTolerance(toleranceValue, fallbackValue) {
                if (isNaN(toleranceValue)) return fallbackValue;
                if (toleranceValue < TOL_MIN) toleranceValue = TOL_MIN;
                if (toleranceValue > TOL_MAX) toleranceValue = TOL_MAX;
                return Math.round(toleranceValue * 100) / 100;
            }

            /**
             * 「チェックボックス＋許容誤差（ラベル・入力欄・スライダー）」を1組生成します。
             * @param {Object} config - 生成設定。
             * @param {string} config.checkboxKey - チェックボックスの LABELS キー。
             * @param {string} config.checkboxTooltipKey - チェックボックスの tooltip の LABELS キー。
             * @param {string} config.labelKey - 許容誤差ラベルの LABELS キー。
             * @param {string} config.tooltipKey - 許容誤差まわりの tooltip の LABELS キー。
             * @param {number} config.initial - 許容誤差の初期値。
             * @param {Array<number>} [config.margins] - グループ余白 [左,上,右,下]。
             * @param {function(number):void} config.onCommit - 確定した許容誤差を受け取るコールバック。
             * @returns {{checkbox: Checkbox, setEnabled: function(boolean):void}} 生成したUIへのアクセサ。
             */
            function createToleranceBlock(config) {
                var group = tabProcess.add('group');
                group.orientation = 'column';
                group.alignChildren = ['left', 'top'];
                if (config.margins) group.margins = config.margins;

                var checkbox = group.add('checkbox', undefined, getLabel(config.checkboxKey));
                checkbox.helpTip = getLabel(config.checkboxTooltipKey);
                checkbox.value = true;

                var row = group.add('group');
                row.orientation = 'row';
                row.alignChildren = ['left', 'center'];
                row.margins = [20, 0, 0, 0];

                var label = row.add('statictext', undefined, labelText(config.labelKey));
                label.helpTip = getLabel(config.tooltipKey);
                label.characters = 10;

                /* ∧∨と入力欄は隙間0で突き合わせる / butt the stepper against the field */
                var stepperFieldGroup = row.add('group');
                stepperFieldGroup.orientation = 'row';
                stepperFieldGroup.alignChildren = ['left', 'center'];
                stepperFieldGroup.spacing = 0;
                stepperFieldGroup.margins = 0;

                /* 0.01刻み（shift で0.1の倍数へ）・0.01〜3.00。増減後はスライダーと予測表示も揃える
                   steps of 0.01 (Shift: multiples of 0.1), 0.01-3.00; the slider and the prediction follow */
                var stepperGroup = addStepper(stepperFieldGroup, function () { return input; }, {
                    step: 0.01, shiftStep: 0.1, min: TOL_MIN, max: TOL_MAX,
                    upTooltipKey: 'tooltip.stepUpTolerance', downTooltipKey: 'tooltip.stepDownTolerance',
                    onStep: function () { sync(parseToleranceText(input.text), true); }
                });
                var input = stepperFieldGroup.add('edittext', undefined, config.initial.toFixed(2));
                input.helpTip = getLabel(config.tooltipKey);
                input.characters = 6;

                var slider = group.add('slider', undefined, Math.round(config.initial * 100), TOL_MIN * 100, TOL_MAX * 100);
                slider.helpTip = getLabel(config.tooltipKey);
                slider.preferredSize.width = 160;
                slider.indent = 20;

                var currentValue = config.initial;

                /**
                 * 入力欄・スライダー・呼び出し元の値を、丸めた許容誤差に揃えます。
                 * @param {number} toleranceValue - 設定する許容誤差。
                 * @param {boolean} refresh - 予測表示も更新する場合は true。
                 * @returns {void}
                 */
                function sync(toleranceValue, refresh) {
                    currentValue = clampTolerance(toleranceValue, currentValue);
                    input.text = currentValue.toFixed(2);
                    slider.value = Math.round(currentValue * 100);
                    config.onCommit(currentValue);
                    if (refresh) refreshInfoPreview();
                }

                /* ドラッグ中（onChanging）は重い選択では表示更新を省き、離したとき（onChange）に反映する */
                slider.onChanging = function () {
                    sync(slider.value / 100, !isHeavySelection);
                };
                slider.onChange = function () {
                    sync(slider.value / 100, true);
                };
                input.onChange = function () {
                    sync(parseToleranceText(input.text), true);
                };
                /* ↑↓キーも∧∨と同じ処理で増減する / arrow keys share the stepper's logic */
                bindSteppedArrowKeys(input, stepperGroup);

                return {
                    checkbox: checkbox,
                    /* スライダーは入力欄と別グループにあるため、行だけを切り替えると消し忘れる */
                    setEnabled: function (enabled) {
                        row.enabled = enabled;
                        slider.enabled = enabled;
                        redrawSteppersIn(row); /* ∧∨のディム表示を切り替える / update the stepper dimming */
                    }
                };
            }

            var anchorToleranceBlock = createToleranceBlock({
                checkboxKey: 'checkbox.removeAnchors',
                checkboxTooltipKey: 'tooltip.removeAnchors',
                labelKey: 'label.tolAnchor',
                tooltipKey: 'tooltip.tolAnchor',
                initial: TOL_ANCHOR_COLLINEAR,
                margins: [0, 15, 0, 15],
                onCommit: function (toleranceValue) {
                    TOL_ANCHOR_COLLINEAR = toleranceValue;
                }
            });
            var removeAnchorsCheckbox = anchorToleranceBlock.checkbox;

            var handleToleranceBlock = createToleranceBlock({
                checkboxKey: 'checkbox.removeHandles',
                checkboxTooltipKey: 'tooltip.removeHandles',
                labelKey: 'label.tolHandle',
                tooltipKey: 'tooltip.tolHandle',
                initial: TOL_HANDLE_COLLINEAR,
                onCommit: function (toleranceValue) {
                    TOL_HANDLE_COLLINEAR = toleranceValue;
                }
            });
            var removeHandlesCheckbox = handleToleranceBlock.checkbox;

            removeSameAnchorsCheckbox.onClick = function () {
                refreshInfoPreview();
            };
            removeAnchorsCheckbox.onClick = function () {
                anchorToleranceBlock.setEnabled(removeAnchorsCheckbox.value);
                refreshInfoPreview();
            };
            removeHandlesCheckbox.onClick = function () {
                handleToleranceBlock.setEnabled(removeHandlesCheckbox.value);
                refreshInfoPreview();
            };

            // 初回反映
            refreshInfoPreview();
            anchorToleranceBlock.setEnabled(removeAnchorsCheckbox.value);
            handleToleranceBlock.setEnabled(removeHandlesCheckbox.value);

            // --- Tab 2: 変換 ---
            var tabOther = tabbedPanel.add('tab', undefined, getLabel('tab.other'));
            setupPanel(tabOther);
            /* タブ内の右余白を詰める（PANEL_MARGINS の右16→6）。サブプロパティ代入は反映されないため配列で上書き */
            tabOther.margins = [16, 20, 6, 12];

            // パネル1：アンカーポイントを変換（スムーズ／コーナー）
            var convertPointsPanel = tabOther.add('panel', undefined, getLabel('panel.convertPoints'));
            setupPanel(convertPointsPanel);
            var smoothRadio = convertPointsPanel.add('radiobutton', undefined, getLabel('radio.convertSmooth'));
            var cornerRadio = convertPointsPanel.add('radiobutton', undefined, getLabel('radio.convertCorner'));

            // パネル2：アンカーポイントを追加（アンカー追加／極点追加）
            var addPointsPanel = tabOther.add('panel', undefined, getLabel('panel.addPoints'));
            setupPanel(addPointsPanel);
            var addAnchorsRadio = addPointsPanel.add('radiobutton', undefined, getLabel('radio.addAnchors'));
            addAnchorsRadio.helpTip = getLabel('tooltip.addAnchors');
            var extremePointsRadio = addPointsPanel.add('radiobutton', undefined, getLabel('radio.addExtremePoints'));
            extremePointsRadio.helpTip = getLabel('tooltip.addExtremePoints');

            // パネル3：その他（分割／マド埋め）
            var pathOpsPanel = tabOther.add('panel', undefined, getLabel('panel.pathOps'));
            setupPanel(pathOpsPanel);
            var splitRadio = pathOpsPanel.add('radiobutton', undefined, getLabel('radio.splitAtAnchors'));
            splitRadio.helpTip = getLabel('tooltip.splitAtAnchors');
            var fillHolesRadio = pathOpsPanel.add('radiobutton', undefined, getLabel('radio.fillHoles'));
            fillHolesRadio.helpTip = getLabel('tooltip.fillHoles');

            // 選択内にマド（複合パス）が無ければ「マド埋め」は無効化（ディム表示）
            if (!hasHoles) {
                fillHolesRadio.enabled = false;
            }

            smoothRadio.value = true;

            // パネルをまたぐ radiobutton は ScriptUI の自動排他が効かないため、
            // 全ラジオを1グループとして手動で単一選択を維持する
            var convertRadios = [smoothRadio, cornerRadio, addAnchorsRadio, extremePointsRadio, splitRadio, fillHolesRadio];

            /**
             * 指定ラジオだけを選択状態にし、他をすべて解除して予測表示を更新します。
             * @param {RadioButton} selectedRadio - 選択状態にするラジオボタン。
             * @returns {void}
             */
            function selectConvertRadio(selectedRadio) {
                for (var r = 0; r < convertRadios.length; r++) {
                    convertRadios[r].value = (convertRadios[r] === selectedRadio);
                }
                refreshInfoForConvert(getSelectedConvertMode());
            }

            for (var radioIndex = 0; radioIndex < convertRadios.length; radioIndex++) {
                (function (radio) {
                    radio.onClick = function () {
                        selectConvertRadio(radio);
                    };
                })(convertRadios[radioIndex]);
            }

            tabbedPanel.onChange = function () {
                refreshByActiveTab();
            };

            tabbedPanel.selection = 0;

            var buttonRow = addButtonRow(dlg, { centered: true });
            var btnCancel = buttonRow.rowGroup.add('button', undefined, getLabel('button.cancel'), { name: 'cancel' });
            var btnOK = buttonRow.rowGroup.add('button', undefined, getLabel('button.ok'), { name: 'ok' });

            /* ボタン生成前の予測表示では OK の状態を反映できないため、ここで一度反映する */
            updateOkEnabled();

            var dialogResult = {
                ok: false,
                activeMode: 'process',
                doRemoveSameAnchors: false,
                doRemoveAnchors: false,
                doRemoveHandles: false,
                convertMode: 'smooth'
            };

            btnOK.onClick = function () {
                dialogResult.ok = true;
                dialogResult.activeMode = getActiveMode();
                dialogResult.doRemoveSameAnchors = removeSameAnchorsCheckbox.value;
                dialogResult.doRemoveAnchors = removeAnchorsCheckbox.value;
                dialogResult.doRemoveHandles = removeHandlesCheckbox.value;
                dialogResult.convertMode = getSelectedConvertMode();
                dlg.close(1);
            };
            btnCancel.onClick = function () {
                dlg.close(0);
            };

            prepareDialogWindow(dlg, SCRIPT_NAME);
            dlg.show();
            return dialogResult;
        }

        // --- 変換・分割ロジック ---
        /**
         * パスの全アンカーをコーナーポイント化し、ハンドルを除去します。
         * @param {PathItem} pathItem - 対象のパス。
         * @returns {void}
         */
        function convertToCorner(pathItem) {
            var points = pathItem.pathPoints;
            for (var k = 0; k < points.length; k++) {
                var pt = points[k];
                pt.pointType = PointType.CORNER;
                pt.leftDirection = pt.anchor;
                pt.rightDirection = pt.anchor;
            }
        }

        /**
         * パスの全アンカーをスムーズポイント化し、前後アンカーの距離・角度に応じてハンドルを補正します。
         * @param {PathItem} pathItem - 対象のパス。
         * @returns {void}
         */
        function convertToSmooth(pathItem) {
            var points = pathItem.pathPoints;
            for (var k = 0; k < points.length; k++) {
                var pt = points[k];
                pt.pointType = PointType.SMOOTH;

                // クローズドパスは循環参照、オープンパス端点は片側だけを参照する
                var anchor = pt.anchor;
                var isClosed = !!pathItem.closed;
                var prevPt = null;
                var nextPt = null;

                if (isClosed || k > 0) {
                    prevPt = points[(k - 1 + points.length) % points.length];
                }
                if (isClosed || k < points.length - 1) {
                    nextPt = points[(k + 1) % points.length];
                }

                var dPrev = prevPt ? Math.sqrt(
                    Math.pow(prevPt.anchor[0] - anchor[0], 2) +
                    Math.pow(prevPt.anchor[1] - anchor[1], 2)
                ) : 0;
                var dNext = nextPt ? Math.sqrt(
                    Math.pow(nextPt.anchor[0] - anchor[0], 2) +
                    Math.pow(nextPt.anchor[1] - anchor[1], 2)
                ) : 0;

                var lenLeft = dPrev / 3;
                var lenRight = dNext / 3;

                var angleFactor = 1;
                var balanceFactor = 1;

                var prevVecX = prevPt ? (prevPt.anchor[0] - anchor[0]) : 0;
                var prevVecY = prevPt ? (prevPt.anchor[1] - anchor[1]) : 0;
                var nextVecX = nextPt ? (nextPt.anchor[0] - anchor[0]) : 0;
                var nextVecY = nextPt ? (nextPt.anchor[1] - anchor[1]) : 0;

                // 前後アンカーが同一点または極端に近い場合の 0 除算を防ぐ
                var EPS = 1e-6;
                var hasPrev = dPrev > EPS;
                var hasNext = dNext > EPS;

                var dirX = 1;
                var dirY = 0;

                if (hasPrev && hasNext) {
                    var tVecX = nextVecX / dNext - prevVecX / dPrev;
                    var tVecY = nextVecY / dNext - prevVecY / dPrev;
                    var tLen = Math.sqrt(tVecX * tVecX + tVecY * tVecY);

                    if (tLen < 0.0001) {
                        dirX = nextVecX / dNext;
                        dirY = nextVecY / dNext;
                    } else {
                        dirX = tVecX / tLen;
                        dirY = tVecY / tLen;
                    }
                } else if (hasNext) {
                    dirX = nextVecX / dNext;
                    dirY = nextVecY / dNext;
                } else if (hasPrev) {
                    dirX = -prevVecX / dPrev;
                    dirY = -prevVecY / dPrev;
                }

                // 鋭角や前後距離の偏りが大きい場合は、ハンドル長を少しだけ抑える
                if (hasPrev && hasNext) {
                    var inDirX = -prevVecX / dPrev;
                    var inDirY = -prevVecY / dPrev;
                    var outDirX = nextVecX / dNext;
                    var outDirY = nextVecY / dNext;

                    var dot = inDirX * outDirX + inDirY * outDirY;
                    if (dot < -1) dot = -1;
                    if (dot > 1) dot = 1;

                    var turnAngle = Math.acos(dot);
                    var angleNorm = turnAngle / Math.PI; // 0..1
                    angleFactor = 1 - (0.35 * angleNorm);

                    var minLen = Math.min(dPrev, dNext);
                    var maxLen = Math.max(dPrev, dNext);
                    var imbalance = (maxLen > EPS) ? (1 - (minLen / maxLen)) : 0; // 0..1
                    balanceFactor = 1 - (0.25 * imbalance);
                }

                lenLeft *= angleFactor * balanceFactor;
                lenRight *= angleFactor * balanceFactor;

                // 8方向へ丸めず、計算した接線方向をそのまま使う
                pt.leftDirection = [
                    anchor[0] - dirX * lenLeft,
                    anchor[1] - dirY * lenLeft
                ];
                pt.rightDirection = [
                    anchor[0] + dirX * lenRight,
                    anchor[1] + dirY * lenRight
                ];
            }
        }

        /**
         * オブジェクトを再帰的にたどり、含まれるパスへスムーズ／コーナー変換を適用します。
         * @param {PageItem} item - 対象オブジェクト（PathItem／CompoundPathItem／GroupItem）。
         * @param {string} mode - 'smooth' または 'corner'。
         * @returns {void}
         */
        function convertPathItem(item, mode) {
            if (!item || isSkippableItem(item)) return;
            var convertFn = (mode === 'smooth') ? convertToSmooth : convertToCorner;
            if (item.typename === 'PathItem') {
                convertFn(item);
            } else if (item.typename === 'CompoundPathItem') {
                for (var i = 0; i < item.pathItems.length; i++) {
                    if (!isSkippableItem(item.pathItems[i])) {
                        convertFn(item.pathItems[i]);
                    }
                }
            } else if (item.typename === 'GroupItem') {
                /* convertPathItem は入口で isSkippableItem() を見るため、ここでは重ねて判定しない */
                for (var j = 0; j < item.pageItems.length; j++) {
                    convertPathItem(item.pageItems[j], mode);
                }
            }
        }

        /**
         * 分割後のパスを置く位置の基準になるオブジェクトを返します。
         * 通常は元のパス自身ですが、複合パスの子の場合は複合パス自身を返します。
         * 開いたパスは複合パスへ戻せないため、複合パスを基準に move すれば外へ逃がせます。
         * @param {PathItem} pathItem - 分割対象のパス。
         * @returns {PageItem} 重ね順の基準になるオブジェクト。
         */
        function getSplitAnchorItem(pathItem) {
            var anchorItem = pathItem;
            try {
                var parent = pathItem.parent;
                while (parent && parent.typename === 'CompoundPathItem') {
                    anchorItem = parent;
                    parent = parent.parent;
                }
            } catch (e) {
                logProcessError('getSplitAnchorItem', e);
            }
            return anchorItem;
        }

        /**
         * パスを各セグメントごとの独立したオープンパスに分割します（元のパスは削除）。
         * @param {PathItem} pathItem - 分割対象のパス。
         * @returns {void}
         */
        function splitAtAnchors(pathItem) {
            var pts = pathItem.pathPoints;
            if (pts.length < 2) return;

            var anchorItem = getSplitAnchorItem(pathItem);
            var segCount = pathItem.closed ? pts.length : pts.length - 1;
            var previousPath = null;

            for (var s = 0; s < segCount; s++) {
                var startIndex = s;
                var endIndex = (s + 1) % pts.length;
                var startPoint = pts[startIndex];
                var endPoint = pts[endIndex];

                /* duplicate() は効果・グラフィックスタイル・破線・不透明度まで丸ごと複製するため、
                   属性を1つずつコピーするより取りこぼしがない */
                var newPath = pathItem.duplicate();

                /* 元のパスがあった位置へ移す。複合パスの子の複製は複合パス内に落ちるので、
                   この move が「複合パスの外へ逃がす」処理も兼ねる（開いたパスは複合パスへ戻せない）。
                   PLACEBEFORE は基準のすぐ前面、PLACEAFTER はすぐ背面に入るため、
                   1本目だけ基準の前面に置き、2本目以降は直前に作ったパスの背面へ繋げて順序を保つ */
                try {
                    if (previousPath) {
                        newPath.move(previousPath, ElementPlacement.PLACEAFTER);
                    } else {
                        newPath.move(anchorItem, ElementPlacement.PLACEBEFORE);
                    }
                    previousPath = newPath;
                } catch (e) {
                    logProcessError('splitAtAnchors.move', e);
                }

                /* 複合パスから出したあとでないと開けないため、move の後に closed を落とす */
                newPath.closed = false;

                /* 複製は元と同じ点数なので、2点に詰めてからセグメントの座標で上書きする */
                while (newPath.pathPoints.length > 2) {
                    newPath.pathPoints[newPath.pathPoints.length - 1].remove();
                }

                var newStartPoint = newPath.pathPoints[0];
                newStartPoint.anchor = startPoint.anchor;
                newStartPoint.leftDirection = startPoint.anchor;
                newStartPoint.rightDirection = startPoint.rightDirection;
                newStartPoint.pointType = startPoint.pointType;

                var newEndPoint = newPath.pathPoints[1];
                newEndPoint.anchor = endPoint.anchor;
                newEndPoint.leftDirection = endPoint.leftDirection;
                newEndPoint.rightDirection = endPoint.anchor;
                newEndPoint.pointType = endPoint.pointType;
            }

            pathItem.remove();
        }

        /**
         * マド埋め：グループ化 → 複合パス解除 → ライブパスファインダー（合体）
         * → アピアランス展開 → グループ解除を実行します。
         * @returns {void}
         */
        function fillHolesOnSelection() {
            /* DOM で立てた選択が反映される前にコマンドが走らないよう、先に画面を更新する */
            app.redraw();
            /* executeMenuCommand は実行できないコマンドでも例外を投げず何もしないため、
               冒頭の group と対になる ungroup をそのまま呼ぶ */
            app.executeMenuCommand('group');
            app.executeMenuCommand('noCompoundPath');
            app.executeMenuCommand('Live Pathfinder Add');
            app.executeMenuCommand('expandStyle');
            app.executeMenuCommand('ungroup');
        }

        /**
         * オブジェクトを再帰的にたどり、含まれるパスをアンカーで分割します。
         * @param {PageItem} item - 対象オブジェクト（PathItem／CompoundPathItem／GroupItem）。
         * @returns {void}
         */
        function splitItem(item) {
            if (!item || isSkippableItem(item)) return;
            if (item.typename === 'PathItem') {
                splitAtAnchors(item);
            } else if (item.typename === 'CompoundPathItem') {
                var childPaths = [];
                for (var i = 0; i < item.pathItems.length; i++) {
                    if (!isSkippableItem(item.pathItems[i])) {
                        childPaths.push(item.pathItems[i]);
                    }
                }
                for (var ii = 0; ii < childPaths.length; ii++) {
                    splitAtAnchors(childPaths[ii]);
                }
                /* 子をすべて開いたパスとして外へ出したため、空になった複合パスを残さない */
                try {
                    if (item.pathItems.length === 0) item.remove();
                } catch (e) {
                    logProcessError('splitItem.removeEmptyCompoundPath', e);
                }
            } else if (item.typename === 'GroupItem') {
                /* splitAtAnchors がグループへ新しいパスを足すため、先に子をスナップショットする
                   （スキップ判定は再帰先の splitItem が入口で行う） */
                var childItems = [];
                for (var j = 0; j < item.pageItems.length; j++) {
                    childItems.push(item.pageItems[j]);
                }
                for (var jj = 0; jj < childItems.length; jj++) {
                    splitItem(childItems[jj]);
                }
            }
        }

        // --- 極点（Extreme Points）追加ロジック ---
        /** 2点 a, b を t（0〜1）で線形補間した座標 [x, y] を返します。 */
        function lerpPoint(a, b, t) {
            return [a[0] + (b[0] - a[0]) * t, a[1] + (b[1] - a[1]) * t];
        }

        /**
         * 3次ベジェを媒介変数 t で2本に分割します（de Casteljau）。
         * @param {Array<Array<number>>} controlPoints - 制御点4つ [P0, P1, P2, P3]。
         * @param {number} t - 分割位置（0〜1）。
         * @returns {{left: Array<Array<number>>, right: Array<Array<number>>}} 分割後の左右の制御点列。
         */
        function splitCubic(controlPoints, t) {
            var p01 = lerpPoint(controlPoints[0], controlPoints[1], t);
            var p12 = lerpPoint(controlPoints[1], controlPoints[2], t);
            var p23 = lerpPoint(controlPoints[2], controlPoints[3], t);
            var p012 = lerpPoint(p01, p12, t);
            var p123 = lerpPoint(p12, p23, t);
            var p0123 = lerpPoint(p012, p123, t);
            return { left: [controlPoints[0], p01, p012, p0123], right: [p0123, p123, p23, controlPoints[3]] };
        }

        /**
         * 1次元の3次ベジェ f(t) について f'(t)=0 となる t（極値）を求めます。
         * @param {number} v0 - P0 成分。
         * @param {number} v1 - P1 成分。
         * @param {number} v2 - P2 成分。
         * @param {number} v3 - P3 成分。
         * @returns {Array<number>} (0,1) の範囲にある t の配列。
         */
        function extremaParamsOfComponent(v0, v1, v2, v3) {
            var diff0 = v1 - v0;
            var diff1 = v2 - v1;
            var diff2 = v3 - v2;
            var quadA = diff0 - 2 * diff1 + diff2;
            var quadB = 2 * (diff1 - diff0);
            var quadC = diff0;
            var roots = [];
            var EPS = 1e-6;
            if (Math.abs(quadA) < EPS) {
                if (Math.abs(quadB) > EPS) roots.push(-quadC / quadB);
            } else {
                var discriminant = quadB * quadB - 4 * quadA * quadC;
                if (discriminant >= 0) {
                    var sqrtDiscriminant = Math.sqrt(discriminant);
                    roots.push((-quadB + sqrtDiscriminant) / (2 * quadA));
                    roots.push((-quadB - sqrtDiscriminant) / (2 * quadA));
                }
            }
            var paramsInRange = [];
            for (var i = 0; i < roots.length; i++) {
                if (roots[i] > EPS && roots[i] < 1 - EPS) paramsInRange.push(roots[i]);
            }
            return paramsInRange;
        }

        /**
         * セグメント（p0→p1）でハンドルが水平/垂直になる媒介変数 t を求めます。
         * 直線セグメント（両端ハンドルがアンカーと一致）は対象外で空配列を返す。
         * @param {PathPoint} p0 - 始点アンカー。
         * @param {PathPoint} p1 - 終点アンカー。
         * @returns {Array<number>} 昇順・重複除去済みの t 配列。
         */
        function bezierExtremaParams(p0, p1) {
            var isStraight = samePoint(p0.rightDirection, p0.anchor, TOL_SAMEPOINT) &&
                samePoint(p1.leftDirection, p1.anchor, TOL_SAMEPOINT);
            if (isStraight) return [];

            var P0 = p0.anchor, P1 = p0.rightDirection, P2 = p1.leftDirection, P3 = p1.anchor;
            var paramsX = extremaParamsOfComponent(P0[0], P1[0], P2[0], P3[0]); /* 垂直接線 dx/dt=0 */
            var paramsY = extremaParamsOfComponent(P0[1], P1[1], P2[1], P3[1]); /* 水平接線 dy/dt=0 */
            var extremaParams = paramsX.concat(paramsY);

            extremaParams.sort(function (x, y) { return x - y; });
            var uniqueParams = [];
            for (var i = 0; i < extremaParams.length; i++) {
                if (uniqueParams.length === 0 || Math.abs(extremaParams[i] - uniqueParams[uniqueParams.length - 1]) > 1e-4) uniqueParams.push(extremaParams[i]);
            }
            return uniqueParams;
        }

        /**
         * セグメントを複数の t（元セグメントの媒介変数）で分割し、
         * 端点ハンドルの更新値と挿入する内部アンカーを返します。
         * @param {Array<number>} P0 - 始点アンカー座標。
         * @param {Array<number>} P1 - 始点右ハンドル座標。
         * @param {Array<number>} P2 - 終点左ハンドル座標。
         * @param {Array<number>} P3 - 終点アンカー座標。
         * @param {Array<number>} extremaParams - 昇順の分割 t 配列。
         * @returns {{startRight: Array<number>, endLeft: Array<number>, interior: Array<Object>}} 分割結果。
         */
        function splitSegmentAtParams(P0, P1, P2, P3, extremaParams) {
            var currentCurve = [P0, P1, P2, P3];
            var prevParam = 0;
            var interior = [];
            var startRight = P1;
            for (var k = 0; k < extremaParams.length; k++) {
                var localT = (extremaParams[k] - prevParam) / (1 - prevParam);
                var pieces = splitCubic(currentCurve, localT);
                if (k === 0) {
                    startRight = pieces.left[1];
                } else {
                    interior[interior.length - 1].right = pieces.left[1];
                }
                interior.push({
                    anchor: pieces.left[3],
                    left: pieces.left[2],
                    right: pieces.right[1]
                });
                currentCurve = pieces.right;
                prevParam = extremaParams[k];
            }
            return { startRight: startRight, endLeft: currentCurve[2], interior: interior };
        }

        /**
         * パスの各曲線セグメントの極点にアンカーを追加します（同一オブジェクトを維持したまま再構築）。
         * @param {PathItem} pathItem - 対象のパス。
         * @returns {number} 追加したアンカー数。
         */
        function addExtremePointsToPath(pathItem) {
            var pts = pathItem.pathPoints;
            var n = pts.length;
            if (n < 2) return 0;

            var isClosed = !!pathItem.closed;
            var segCount = isClosed ? n : (n - 1);

            /* 既存アンカーの左右ハンドル（更新用）と、各セグメント直後に挿入する内部点 */
            var leftHandles = [];
            var rightHandles = [];
            var interiorAfter = [];
            for (var i = 0; i < n; i++) {
                leftHandles[i] = [pts[i].leftDirection[0], pts[i].leftDirection[1]];
                rightHandles[i] = [pts[i].rightDirection[0], pts[i].rightDirection[1]];
                interiorAfter[i] = [];
            }

            var added = 0;
            for (var s = 0; s < segCount; s++) {
                var j = (s + 1) % n;
                var extremaParams = bezierExtremaParams(pts[s], pts[j]);
                if (!extremaParams.length) continue;

                var splitResult = splitSegmentAtParams(
                    pts[s].anchor, pts[s].rightDirection,
                    pts[j].leftDirection, pts[j].anchor, extremaParams
                );
                rightHandles[s] = splitResult.startRight;
                leftHandles[j] = splitResult.endLeft;
                interiorAfter[s] = splitResult.interior;
                added += splitResult.interior.length;
            }

            if (added === 0) return 0;

            /* 新しい順序で点仕様を組み立てる（既存アンカー＋各セグメント直後の内部点） */
            var newPoints = [];
            for (var a = 0; a < n; a++) {
                newPoints.push({
                    anchor: [pts[a].anchor[0], pts[a].anchor[1]],
                    left: leftHandles[a],
                    right: rightHandles[a],
                    type: pts[a].pointType
                });
                var insertedPoints = interiorAfter[a];
                for (var b = 0; b < insertedPoints.length; b++) {
                    newPoints.push({
                        anchor: insertedPoints[b].anchor,
                        left: insertedPoints[b].left,
                        right: insertedPoints[b].right,
                        type: PointType.SMOOTH
                    });
                }
            }

            /* 同一 PathItem を保持したまま、点数を合わせて全点を上書き（アピアランス維持） */
            while (pathItem.pathPoints.length < newPoints.length) {
                pathItem.pathPoints.add();
            }
            for (var c = 0; c < newPoints.length; c++) {
                var pt = pathItem.pathPoints[c];
                pt.anchor = newPoints[c].anchor;
                pt.pointType = newPoints[c].type;
                pt.leftDirection = newPoints[c].left;
                pt.rightDirection = newPoints[c].right;
            }

            return added;
        }

        /**
         * オブジェクトを再帰的にたどり、含まれるパスの極点にアンカーを追加します。
         * @param {PageItem} item - 対象オブジェクト（PathItem／CompoundPathItem／GroupItem）。
         * @returns {void}
         */
        function addExtremePointsToItem(item) {
            if (!item || isSkippableItem(item)) return;
            if (item.typename === 'PathItem') {
                try {
                    addExtremePointsToPath(item);
                } catch (e) {
                    logProcessError("addExtremePointsToPath", e);
                }
            } else if (item.typename === 'CompoundPathItem') {
                /* addExtremePointsToItem は入口で isSkippableItem() を見るため、ここでは重ねて判定しない */
                for (var i = 0; i < item.pathItems.length; i++) {
                    addExtremePointsToItem(item.pathItems[i]);
                }
            } else if (item.typename === 'GroupItem') {
                for (var j = 0; j < item.pageItems.length; j++) {
                    addExtremePointsToItem(item.pageItems[j]);
                }
            }
        }

        /**
         * 指定 targets に対して追加される極点アンカー数を予測します（実際には変更しない）。
         * @param {Array<PathItem>} targets - 対象の PathItem 配列。
         * @returns {number} 追加されるアンカー総数。
         */
        function countExtremaForTargets(targets) {
            var total = 0;
            if (!targets || !targets.length) return total;
            for (var i = 0; i < targets.length; i++) {
                var item = targets[i];
                if (!item) continue;
                try {
                    var pts = item.pathPoints;
                    var n = pts.length;
                    if (n < 2) continue;
                    var isClosed = !!item.closed;
                    var segCount = isClosed ? n : (n - 1);
                    for (var s = 0; s < segCount; s++) {
                        var j = (s + 1) % n;
                        total += bezierExtremaParams(pts[s], pts[j]).length;
                    }
                } catch (e) {
                    logProcessError("countExtremaForTargets", e);
                }
            }
            return total;
        }

        // ダイアログ表示時点の選択を確定（予測と実行を同一対象に揃える）
        var selectionAtOpen = getSelectionOrAlert();
        if (!selectionAtOpen) return;
        var targetsAtOpen = getTargetPathItemsFromSelection(selectionAtOpen);
        var hasHolesAtOpen = selectionHasHoles(selectionAtOpen);

        var ui = showDialog(targetsAtOpen, hasHolesAtOpen);
        if (!ui.ok) return;

        // タブindex依存だと ScriptUI 環境差で誤判定しうるため、明示状態で分岐する
        /**
         * 選択をダイアログ表示時点のスナップショットに復元します。
         * @param {Array<PageItem>} selectionSnapshot - 復元する選択オブジェクト配列。
         * @returns {number} 復元できたオブジェクト数。
         */
        function restoreSelection(selectionSnapshot) {
            var restoredCount = 0;
            try {
                if (!hasDocument()) return 0;
                var doc = app.activeDocument;
                doc.selection = null;
                for (var i = 0; i < selectionSnapshot.length; i++) {
                    var item = selectionSnapshot[i];
                    if (!item) continue;
                    try {
                        item.selected = true;
                        restoredCount++;
                    } catch (e) {
                        // skip items that can no longer be selected
                    }
                }
            } catch (e) {
                logProcessError("restoreSelection", e);
            }
            return restoredCount;
        }

        /**
         * 選択依存メニューコマンド用に、スキップ対象を除いた PathItem だけを選択復元します。
         * @param {Array<PageItem>} selectionSnapshot - 元の選択スナップショット。
         * @returns {number} 復元できた PathItem 数。
         */
        function restoreSelectableSelection(selectionSnapshot) {
            var restoredCount = 0;
            try {
                if (!hasDocument()) return 0;
                var doc = app.activeDocument;
                var selectablePathItems = getTargetPathItemsFromSelection(selectionSnapshot);

                doc.selection = null;
                for (var i = 0; i < selectablePathItems.length; i++) {
                    var pathItem = selectablePathItems[i];
                    if (!pathItem) continue;
                    try {
                        pathItem.selected = true;
                        restoredCount++;
                    } catch (e) {
                        // skip items that can no longer be selected
                    }
                }
            } catch (e) {
                logProcessError("restoreSelectableSelection", e);
            }
            return restoredCount;
        }

        // 実行前に選択を復元（1つも復元できなければ中止して理由を伝える）
        if (restoreSelection(selectionAtOpen) === 0) {
            alert(getLabel('alert.lostSelection'));
            return;
        }

        if (ui.activeMode === 'other') {
            // --- 「変換」タブ: 変換・分割処理 ---
            /* splitItem / addExtremePointsToItem / convertPathItem はいずれも入口で
               isSkippableItem() を見るため、ここでは重ねて判定しない */
            var selectionSnapshot = selectionAtOpen.slice(0);
            if (ui.convertMode === 'add') {
                if (restoreSelectableSelection(selectionSnapshot) > 0) {
                    /* DOM で立てた選択が反映される前にコマンドが走らないよう、先に画面を更新する */
                    app.redraw();
                    app.executeMenuCommand('Add Anchor Points2');
                }
            } else if (ui.convertMode === 'split') {
                for (var i = 0; i < selectionSnapshot.length; i++) {
                    splitItem(selectionSnapshot[i]);
                }
            } else if (ui.convertMode === 'extreme') {
                for (var j = 0; j < selectionSnapshot.length; j++) {
                    addExtremePointsToItem(selectionSnapshot[j]);
                }
            } else if (ui.convertMode === 'fillHoles') {
                /* マド埋めは複合パス自体に対する操作なので、子パスに分解せず
                   ダイアログ表示時点の選択をそのまま復元する */
                if (restoreSelection(selectionSnapshot) > 0) {
                    fillHolesOnSelection();
                }
            } else {
                for (var k = 0; k < selectionSnapshot.length; k++) {
                    convertPathItem(selectionSnapshot[k], ui.convertMode);
                }
            }
        } else {
            // --- 「削除対象」タブ: クリーンアップ処理 ---
            if (!ui.doRemoveSameAnchors && !ui.doRemoveAnchors && !ui.doRemoveHandles) {
                return;
            }

            // 情報パネルの予測とまったく同じ関数・同じ実行順を通す
            runCleanupOnTargets(targetsAtOpen, ui.doRemoveSameAnchors, ui.doRemoveAnchors, ui.doRemoveHandles);
        }
    })();

})();
