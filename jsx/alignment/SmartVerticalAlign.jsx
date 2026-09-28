#target illustrator
#targetengine "SmartVerticalAlignEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

ポイント文字およびエリア内文字の［字形の境界に整列］を切り替えながら、垂直方向（上・中央・下）に整列します。
チェックボックスやラジオボタンの操作はそのつどプレビューへ反映され、T / M / B キーでも整列位置を切り替えられます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartVerticalAlign.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n9ee716675032

### Overview

Aligns objects vertically — top, center or bottom — while toggling Align to Glyph Bounds for point text and area text.
Every checkbox and radio button refreshes the preview, and the T, M and B keys switch the alignment position.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartVerticalAlign.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SmartVerticalAlign";           /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.4";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-08-04";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-28";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartVerticalAlign.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartVerticalAlign.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n9ee716675032"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* ［プレビュー境界］の初期状態 / Initial state of the Preview Bounds checkbox */
    var DEFAULT_USE_PREVIEW_BOUNDS = false;

    // =========================================
    // 環境設定キー / Preference keys
    // =========================================

    /* 境界にプレビュー境界（線幅・効果）を含めるか / Include stroke and effects in bounds */
    var PREF_KEY_INCLUDE_STROKE_IN_BOUNDS = 'includeStrokeInBounds';

    // =========================================
    // 整列・字形境界の定義 / Alignment and glyph-bounds tables
    // =========================================

    /* 整列位置。ラジオボタン・メニューコマンド・ショートカットキーを1か所で対応付ける
       / Alignment options: radio button, menu command and shortcut key in one table */
    var ALIGN_OPTIONS = [
        { labelKey: 'radio.top',    tooltipKey: 'tooltip.top',    menuCommand: 'Vertical Align Top',    shortcutKey: 'T' },
        { labelKey: 'radio.center', tooltipKey: 'tooltip.center', menuCommand: 'Vertical Align Center', shortcutKey: 'M' },
        { labelKey: 'radio.bottom', tooltipKey: 'tooltip.bottom', menuCommand: 'Vertical Align Bottom', shortcutKey: 'B' }
    ];
    var ALIGN_INDEX_CENTER = 1;
    var ALIGN_INDEX_BOTTOM = 2;

    /* ［字形の境界に整列］のチェックボックスと、対応する環境設定キー・文字種
       / Glyph-bounds checkboxes with their preference key and text kind */
    var GLYPH_BOUNDS_OPTIONS = [
        { labelKey: 'checkbox.pointText', tooltipKey: 'tooltip.pointText', prefKey: 'EnableActualPointTextSpaceAlign', textKind: TextType.POINTTEXT },
        { labelKey: 'checkbox.areaText',  tooltipKey: 'tooltip.areaText',  prefKey: 'EnableActualAreaTextSpaceAlign',  textKind: TextType.AREATEXT }
    ];
    var GLYPH_INDEX_POINT_TEXT = 0;

    // =========================================
    // レイアウト / Layout
    // =========================================

    /* ダイアログの初期位置 / Dialog initial position */
    var DIALOG_OFFSET_X = 300;  /* 右(+)／左(-) / shift right (+) / left (-) */
    var DIALOG_OFFSET_Y = 0;    /* 下(+)／上(-) / shift down (+) / up (-) */

    /* 余白と間隔 / Margins and spacing */
    var PANEL_MARGINS = [15, 20, 15, 15];      /* パネル余白 [左,上,右,下] */
    var RADIO_SPACING = 5;                     /* ラジオボタンの行間 / spacing between radio buttons */
    var PREVIEW_ROW_MARGINS = [15, 0, 15, 0];  /* ［プレビュー境界］行の余白 */
    var BUTTON_WIDTH = 90;

    /**
     * パネルの共通設定
     * @param {Panel} targetPanel - 対象パネル
     * @param {number} [spacing] - 要素間隔（省略時はScriptUIの既定値のまま）
     * @returns {void}
     */
    function setupPanel(targetPanel, spacing) {
        targetPanel.orientation = 'column';
        targetPanel.alignChildren = ['left', 'top'];
        targetPanel.margins = PANEL_MARGINS;
        if (typeof spacing === 'number') {
            targetPanel.spacing = spacing;
        }
    }

    /**
     * ダイアログの表示位置をずらす
     * @param {Window} targetDialog - 対象ダイアログ
     * @param {number} offsetX - 横方向のオフセット
     * @param {number} offsetY - 縦方向のオフセット
     * @returns {void}
     */
    function shiftDialogPosition(targetDialog, offsetX, offsetY) {
        targetDialog.onShow = function () {
            targetDialog.location = [
                targetDialog.location[0] + offsetX,
                targetDialog.location[1] + offsetY
            ];
        };
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

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // ボタン行（再利用パーツ） / Button row (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内（ダイアログを作る関数より前）に貼る。
    //    識別子は BUTTON_ROW_* / addButtonRow
    // 2. ダイアログの最後で行を作り、ボタンは btn 接頭辞の変数で左右のグループに足す（キャンセル → OK の順）
    //      var buttonRow = addButtonRow(dialog);
    //      var btnPreferences = buttonRow.leftGroup.add("button", undefined, getLabel("button.preferences"));
    //      var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
    //      var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
    //    左右中央に並べるときは addButtonRow(dialog, { centered: true }) にして、buttonRow.rowGroup に直接足す
    // 3. 行の上の余白は BUTTON_ROW_TOP_MARGIN で決める。左右の余白はダイアログの margins に任せる
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

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

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // ボタン行（再利用パーツ）ここまで / End of the reusable button row
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    // =========================================
    // ローカライズ / Localization
    // =========================================

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // ローカライズ（再利用パーツ） / Localization (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内のローカライズ節（LABELS の直前）に貼る。
    //    uiLang を使うコード（StepperButtons・LinkToggle の部品など）より前に置く
    // 2. 識別子は uiLang / getCurrentLang / getLabel / labelText / labelValueText / fillLabelPlaceholders。
    //    同じ役割の既存の関数・変数（getCurrentLanguage、currentLanguage、formatLabel など）は消して、これに寄せる
    // 3. 呼び出しはどちらの形でもよい（混ぜてもよい）
    //      getLabel("dialog.title")        … パス
    //      getLabel(LABELS.dialog.title)   … { ja, en } を直接
    //      getLabel("alert.count", { count: 3 })  … "{count} 個" の {count} を差し込む
    //      getLabel("alert.range", [1, 10])       … "%1〜%2" の %1・%2 を差し込む
    //      labelText("fieldLabel.width")   … 末尾にコロン（日本語は全角「：」、英語は半角「:」）
    //      labelValueText("message.count", 5) … 「件数：5」／「Count: 5」（値が続く1行。英語はコロンのあとに空白）
    // 4. 見つからないパスはパスの文字列をそのまま返す（表示で気づけるように）。{ ja, en } が無いときは空文字
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

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

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // ローカライズ（再利用パーツ）ここまで / End of the reusable localization
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    /* ラベル定義（カテゴリ別）/ Label definitions (by category) */
    var LABELS = {
        dialog: {
            title: { ja: "垂直方向の整列", en: "Vertical Alignment" }
        },
        panel: {
            glyphBounds: { ja: "字形の境界に整列", en: "Align to Glyph Bounds" },
            alignment: { ja: "整列位置", en: "Alignment" }
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." }
        },
        tooltip: {
            top:    { ja: "選択範囲の上端にそろえます（T キー）", en: "Align to the top of the selection (T)" },
            center: { ja: "選択範囲の上下中央にそろえます（M キー）", en: "Align to the vertical center of the selection (M)" },
            bottom: { ja: "選択範囲の下端にそろえます（B キー）", en: "Align to the bottom of the selection (B)" },
            pointText: {
                ja: "ポイント文字を、仮想ボディではなく字形の実際の輪郭でそろえます。",
                en: "Aligns point text by the actual glyph outlines instead of the em box."
            },
            areaText: {
                ja: "エリア内文字を、テキストエリアの枠ではなく字形の実際の輪郭でそろえます。",
                en: "Aligns area text by the actual glyph outlines instead of the text area frame."
            },
            previewBounds: {
                ja: "線幅や効果を含めた見た目の端を基準にそろえます（環境設定の［プレビュー境界を使用］を切り替えます）。",
                en: "Aligns by the visible edges including strokes and effects (toggles the Use Preview Bounds preference)."
            }
        },
        checkbox: {
            pointText: { ja: "ポイント文字", en: "Point Text" },
            areaText: { ja: "エリア内文字", en: "Area Text" },
            previewBounds: { ja: "プレビュー境界", en: "Preview Bounds" }
        },
        radio: {
            top: { ja: "上", en: "Top" },
            center: { ja: "中央", en: "Center" },
            bottom: { ja: "下", en: "Bottom" }
        },
        button: {
            ok: { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        }
    };

    // =========================================
    // 選択オブジェクトの取得 / Selection
    // =========================================

    /**
     * 整列の対象になるオブジェクトか判定する
     * @param {PageItem} pageItem - 判定するオブジェクト
     * @returns {boolean} 対象なら true
     */
    function isAlignableItem(pageItem) {
        if (!pageItem) return false;
        return (pageItem.typename === "TextFrame" ||
            pageItem.typename === "PathItem" ||
            pageItem.typename === "GroupItem" ||
            pageItem.typename === "CompoundPathItem");
    }

    /**
     * 選択中の整列対象オブジェクトを配列で取得する
     * @returns {PageItem[]} 整列対象のオブジェクト
     */
    function getAlignableSelection() {
        var currentSelection = app.activeDocument.selection;
        var alignableItems = [];
        for (var i = 0; i < currentSelection.length; i++) {
            if (isAlignableItem(currentSelection[i])) {
                alignableItems.push(currentSelection[i]);
            }
        }
        return alignableItems;
    }

    // =========================================
    // 整列プレビュー / Alignment preview
    // =========================================

    /* ALIGN_OPTIONS と同じ並びのラジオボタン。UI構築時に埋める
       / Radio buttons in ALIGN_OPTIONS order, filled while the dialog is built */
    var alignRadios = [];

    /**
     * 指定したインデックスの整列位置だけを選択状態にする
     * @param {number} optionIndex - ALIGN_OPTIONS のインデックス
     * @returns {void}
     */
    function selectAlignOption(optionIndex) {
        for (var i = 0; i < alignRadios.length; i++) {
            alignRadios[i].value = (i === optionIndex);
        }
    }

    /**
     * 選択中の整列位置のメニューコマンドを返す
     * @returns {string|null} メニューコマンド名（未選択なら null）
     */
    function getSelectedAlignCommand() {
        for (var i = 0; i < alignRadios.length; i++) {
            if (alignRadios[i].value) return ALIGN_OPTIONS[i].menuCommand;
        }
        return null;
    }

    /**
     * 現在の設定で整列を実行し、プレビューへ即時反映する
     * @returns {void}
     */
    function applyPreviewAlignment() {
        var menuCommand = getSelectedAlignCommand();
        if (menuCommand && getAlignableSelection().length > 0) {
            /* 環境設定の変更を境界に反映させてから整列する（OFFに戻したときも効かせるため）
               / Redraw first so the new preference is reflected in the bounds, including when it is turned OFF */
            app.redraw();
            app.executeMenuCommand(menuCommand);
        }
        app.redraw();
    }

    // =========================================
    // UI構築 / Dialog
    // =========================================

    /**
     * ［字形の境界に整列］パネルを追加する
     * @param {Window} targetDialog - 追加先のダイアログ
     * @param {PageItem[]} alignableItems - 選択中の整列対象オブジェクト
     * @returns {Checkbox[]} GLYPH_BOUNDS_OPTIONS と同じ並びのチェックボックス
     */
    function addGlyphBoundsPanel(targetDialog, alignableItems) {
        var glyphBoundsPanel = targetDialog.add('panel', undefined, getLabel('panel.glyphBounds'));
        setupPanel(glyphBoundsPanel);

        /* 選択中の文字種。ポイント文字／エリア内文字のときだけ他方をディムする
           / Kind of the selected text; dims the checkbox for the other kind */
        var selectedTextKind = (alignableItems.length > 0) ? alignableItems[0].kind : null;

        var glyphBoundsCheckboxes = [];
        for (var i = 0; i < GLYPH_BOUNDS_OPTIONS.length; i++) {
            glyphBoundsCheckboxes.push(addGlyphBoundsCheckbox(glyphBoundsPanel, GLYPH_BOUNDS_OPTIONS[i], selectedTextKind));
        }
        return glyphBoundsCheckboxes;
    }

    /**
     * ［字形の境界に整列］のチェックボックスを1つ追加し、環境設定とプレビューに結び付ける
     * @param {Panel} parentPanel - 追加先のパネル
     * @param {object} glyphOption - GLYPH_BOUNDS_OPTIONS の1項目
     * @param {TextType} selectedTextKind - 選択中の文字種（テキスト以外は null）
     * @returns {Checkbox} 追加したチェックボックス
     */
    function addGlyphBoundsCheckbox(parentPanel, glyphOption, selectedTextKind) {
        var glyphBoundsCheckbox = parentPanel.add('checkbox', undefined, getLabel(glyphOption.labelKey));
        glyphBoundsCheckbox.helpTip = getLabel(glyphOption.tooltipKey);
        glyphBoundsCheckbox.value = app.preferences.getBooleanPreference(glyphOption.prefKey);

        /* 選択が別の文字種ならディム / Dim when the selection is the other text kind */
        if (selectedTextKind != null && selectedTextKind !== glyphOption.textKind) {
            glyphBoundsCheckbox.enabled = false;
        }

        /* ON/OFFで環境設定を書き換え、そのままプレビューを更新
           / Write the preference and refresh the preview on every toggle */
        glyphBoundsCheckbox.onClick = function () {
            app.preferences.setBooleanPreference(glyphOption.prefKey, glyphBoundsCheckbox.value === true);
            applyPreviewAlignment();
        };
        return glyphBoundsCheckbox;
    }

    /**
     * ［整列］パネルを追加し、ラジオボタンを alignRadios に登録する
     * @param {Window} targetDialog - 追加先のダイアログ
     * @returns {void}
     */
    function addAlignmentPanel(targetDialog) {
        var alignmentPanel = targetDialog.add('panel', undefined, getLabel('panel.alignment'));
        setupPanel(alignmentPanel, RADIO_SPACING);

        for (var i = 0; i < ALIGN_OPTIONS.length; i++) {
            var alignRadio = alignmentPanel.add('radiobutton', undefined, getLabel(ALIGN_OPTIONS[i].labelKey));
            alignRadio.helpTip = getLabel(ALIGN_OPTIONS[i].tooltipKey);
            alignRadio.onClick = applyPreviewAlignment;
            alignRadios.push(alignRadio);
        }
    }

    /**
     * ［プレビュー境界］の行を追加する
     * @param {Window} targetDialog - 追加先のダイアログ
     * @returns {void}
     */
    function addPreviewBoundsRow(targetDialog) {
        var previewBoundsRow = targetDialog.add('group');
        previewBoundsRow.orientation = 'row';
        previewBoundsRow.alignChildren = ['left', 'center'];
        previewBoundsRow.margins = PREVIEW_ROW_MARGINS;

        var previewBoundsCheckbox = previewBoundsRow.add('checkbox', undefined, getLabel('checkbox.previewBounds'));
        previewBoundsCheckbox.helpTip = getLabel('tooltip.previewBounds');
        previewBoundsCheckbox.value = DEFAULT_USE_PREVIEW_BOUNDS;
        previewBoundsCheckbox.onClick = function () {
            /* ONで線幅・効果を境界に含める / ON: include stroke and effects in the bounds */
            app.preferences.setBooleanPreference(PREF_KEY_INCLUDE_STROKE_IN_BOUNDS, previewBoundsCheckbox.value === true);
            applyPreviewAlignment();
        };
    }

    /**
     * T / M / B キーで整列位置を切り替えるハンドラーを登録する
     * @param {Window} targetDialog - 対象ダイアログ
     * @returns {void}
     */
    function addAlignmentKeyHandler(targetDialog) {
        targetDialog.addEventListener("keydown", function (event) {
            for (var i = 0; i < ALIGN_OPTIONS.length; i++) {
                if (event.keyName !== ALIGN_OPTIONS[i].shortcutKey) continue;
                selectAlignOption(i);
                applyPreviewAlignment();
                event.preventDefault();
                return;
            }
        });
    }

    /**
     * 選択内容に応じて整列位置とポイント文字チェックボックスの初期値を決める
     * @param {PageItem[]} alignableItems - 選択中の整列対象オブジェクト
     * @param {Checkbox} pointTextCheckbox - ポイント文字のチェックボックス
     * @returns {void}
     */
    function applyDefaultAlignment(alignableItems, pointTextCheckbox) {
        var hasTextFrame = false;
        var hasOtherItem = false;
        for (var i = 0; i < alignableItems.length; i++) {
            if (alignableItems[i].typename === "TextFrame") {
                hasTextFrame = true;
            } else {
                hasOtherItem = true;
            }
        }

        if (hasTextFrame && !hasOtherItem) {
            /* テキストのみ：下揃え / Text only: align bottom */
            pointTextCheckbox.value = false;
            selectAlignOption(ALIGN_INDEX_BOTTOM);
            return;
        }

        /* テキストと図形の混在、または図形のみ：中央揃え / Mixed or shapes only: align center */
        if (hasTextFrame) {
            pointTextCheckbox.value = true;
        }
        selectAlignOption(ALIGN_INDEX_CENTER);
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * ダイアログを組み立てて表示する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel('alert.noDocument'));
            return;
        }

        var alignableItems = getAlignableSelection();

        var alignDialog = new Window('dialog');
        alignDialog.text = getLabel('dialog.title') + ' ' + SCRIPT_VERSION;
        alignDialog.orientation = 'column';
        alignDialog.alignChildren = ['fill', 'top'];
        shiftDialogPosition(alignDialog, DIALOG_OFFSET_X, DIALOG_OFFSET_Y);

        var glyphBoundsCheckboxes = addGlyphBoundsPanel(alignDialog, alignableItems);
        addAlignmentPanel(alignDialog);
        addPreviewBoundsRow(alignDialog);
        /* ボタンエリア / Button row */
        var buttonRow = addButtonRow(alignDialog);
        var btnCancel = buttonRow.rightGroup.add('button', undefined, getLabel('button.cancel'), { name: 'cancel' });
        btnCancel.preferredSize.width = BUTTON_WIDTH;
        btnCancel.onClick = function () {
            alignDialog.close();
        };

        var btnOK = buttonRow.rightGroup.add('button', undefined, getLabel('button.ok'), { name: 'ok' });
        btnOK.preferredSize.width = BUTTON_WIDTH;
        btnOK.onClick = function () {
            alignDialog.close();
        };
        addAlignmentKeyHandler(alignDialog);

        applyDefaultAlignment(alignableItems, glyphBoundsCheckboxes[GLYPH_INDEX_POINT_TEXT]);

        prepareDialogWindow(alignDialog, SCRIPT_NAME);
        alignDialog.show();
    }

    main();

})();
