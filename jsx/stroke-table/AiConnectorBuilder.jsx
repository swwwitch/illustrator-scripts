#target illustrator
#targetengine "AiConnectorBuilderEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

キーオブジェクトを起点に、選択した各図形へコネクターを引きます。
直線・ワープ・カギ・分岐・カーブの5種類の経路に、線・線端・矢印をプレビューしながら設定できます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AiConnectorBuilder.md

note記事も参照してください。
https://note.com/dtp_tranist/n/nd0d3486e5f68

### Overview

Draws a connector from the key object to each of the selected objects.
Choose a straight, warped, elbow, branch, or curved route and set the stroke, caps, and arrowheads with a live preview.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AiConnectorBuilder.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AiConnectorBuilder";           /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.3";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-09-05";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-29";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AiConnectorBuilder.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AiConnectorBuilder.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/nd0d3486e5f68"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

/**
 * @discussion 参考、謝辞 / Reference and acknowledgements
 * カーブの作図（中点から振った1点を2次ベジェの制御点として扱う考え方）
 * Egor Chistyakov (@tchegr)
 * https://x.com/tchegr
 */

(function () {
    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    var CONNECTOR_LAYER_NAME = { ja: "コネクター", en: "Connector" };  /* コネクターの作成先レイヤー名 / target layer name */

    /* 線の初期値 / Line defaults */
    var DEFAULT_STROKE_WIDTH  = 1;    /* 線幅（pt） / stroke width */
    var DEFAULT_STROKE_JOIN   = 0;    /* 0=マイター 1=ラウンド 2=ベベル */
    var DEFAULT_STROKE_CAP    = 0;    /* 0=なし 1=丸形 2=突出 */
    var DEFAULT_DASH_STYLE    = 0;    /* 0=なし 1=破線 2=ドット */
    var DEFAULT_DASH_SEGMENTS = 8;    /* 分割数（線分・ドットの数） */
    var DEFAULT_DASH_GAP      = 3;    /* 破線の間隔（pt） */

    /* コネクターの初期値 / Connector defaults */
    var DEFAULT_LINE_SHAPE    = 0;    /* 0=直線 1=ワープ 2=カギ 3=分岐 4=カーブ */
    var DEFAULT_START_POINT   = 0;    /* 開始点 0=各辺の中心 1=等分 2=中心 */
    var DEFAULT_UNIFY_ANCHOR  = 4;    /* 開始点をまとめるときの位置（0〜8。4=中央＝自動） */
    var DEFAULT_WARP_TYPE     = 0;    /* WARP_TYPE_CHOICES のインデックス（0=でこぼこ） */
    var DEFAULT_WARP_AMOUNT   = -80;  /* カーブ（%）：ワープの曲がり具合とカーブのふくらみに共用 */
    var WARP_AMOUNT_MIN       = -100; /* カーブの下限（%） */
    var WARP_AMOUNT_MAX       = 100;  /* カーブの上限（%） */
    var DEFAULT_WARP_AXIS     = 0;    /* 0=自動 1=水平 2=垂直 */
    var DEFAULT_CORNER_RADIUS = 0;    /* カギ・分岐の角丸半径（pt）／0で角丸なし */
    var WARP_DEFORM_H         = 0;    /* 変形・水平方向（%）：XMLで必須のため固定値で渡す */
    var WARP_DEFORM_V         = 0;    /* 変形・垂直方向（%）：同上 */

    /* 矢印の初期値 / Arrowhead defaults */
    var DEFAULT_ARROW_INDEX  = 1;    /* ARROW_CHOICES のインデックス（0=なし） */
    var DEFAULT_ARROW_SCALE  = 100;  /* 矢印の倍率（%）：ARROW_CHOICES に指定がないときの値 */
    var DEFAULT_END_GAP      = 0;    /* 終点と相手の図形とのすき間（pt） */
    var DEFAULT_ARROW_POSITION = 0;  /* 0=終点のみ 1=両端 */
    /* 黒丸に使う矢印番号（［線］パネルの矢印リストに合わせて調整） */
    var ARROW_DOT_FILLED = 21;       /* 黒丸 */
    /* 白丸は黒丸の中心に白い●を重ねて作る */
    var WHITE_DOT_RATIO = 1.5;       /* 白い●の直径＝線幅の何倍か（WHITE_DOT_BASE_SCALE のときの値） */
    var WHITE_DOT_BASE_SCALE = 50;   /* 白い●の基準になる矢印の倍率（%） */

    /* 一時アクション / Temporary action（矢印はDOMから設定できないためアクションで適用） */
    var ACTION_SET_NAME  = "SwwwitchTempConnectorSet";
    var ACTION_NAME      = "SwwwitchTempConnector";

    /* パラメータキー / Parameter keys（記録した .aia から採取） */
    var KEY_STROKE_WIDTH  = 2003072104;      /* 線幅 / stroke width */
    var KEY_CAP           = 1667330094;      /* 線端 / cap */
    var KEY_JOIN          = 1785686382;      /* 角の形状 / join */
    var KEY_DASH_INT      = 1684825454;      /* 破線（整数）/ dash (integer) */
    var KEY_DASH_BOOL     = 1684104298;      /* 破線（真偽）/ dash (boolean) */
    var KEY_ARROW_HEAD_1  = 1634231345;      /* ahd1: 始点の形状 / start arrowhead */
    var KEY_ARROW_HEAD_2  = 1634231346;      /* ahd2: 終点の形状 / end arrowhead */
    var KEY_ARROW_SCALE_1 = 1634951985;      /* asc1: 始点の倍率 / start scale */
    var KEY_ARROW_SCALE_2 = 1634951986;      /* asc2: 終点の倍率 / end scale */
    var KEY_ARROW_ALIGN   = 1634230636;      /* ahal: 矢印の配置 / tip alignment */
    var KEY_ALIGN         = 1634494318;      /* algn: 線の位置 / stroke alignment */
    var UNIT_POINT        = 592476268;       /* ポイント / point（parameter /unit） */

    /* 判定の許容値 / Tolerances */
    var KEY_DETECT_TOLERANCE_PT = 0.001; /* 整列後に「動いていない」とみなす差（pt） */
    var COORD_TOLERANCE_PT      = 0.001; /* 座標が同じとみなす差（pt） */
    var BEND_CLEARANCE_PT       = 4;     /* 折れ位置を選択オブジェクトから離す量（pt） */

    // =========================================
    // レイアウト / Layout
    // =========================================

    var WINDOW_MARGINS = 16;                 /* ウィンドウ外周の余白 */
    var WINDOW_SPACING = 12;                 /* ウィンドウ内の要素間隔 */
    var PANEL_MARGINS  = [16, 20, 16, 12];   /* パネル余白 [左,上,右,下] */
    var PANEL_SPACING  = 8;                  /* パネル内の要素間隔 */
    var LABEL_WIDTH        = 84;             /* コネクターパネルの行ラベル幅（右揃え） */
    var COLUMN_LABEL_WIDTH = 70;             /* 線パネルの行ラベル幅（2カラムなので狭め） */
    var ARROW_LABEL_WIDTH  = { ja: 42, en: 50 }; /* 矢印パネルの行ラベル幅（コロン込み。英語は「Offset:」が入る幅） */
    var FIELD_CHARS    = 4;                  /* 数値欄の文字数 */
    var LIST_WIDTH     = 150;                /* ドロップダウンの幅 */
    var SLIDER_MIN_WIDTH = 60;               /* スライダーの最小幅（余白は fill で伸ばす） */
    var RADIO_COLUMN_SPACING = 4;            /* 縦並びラジオの間隔 */
    var COLUMN_SPACING = 12;                 /* 2カラムの間隔 */
    var PRESET_BUTTON_WIDTH = 60;            /* プリセットの保存・削除ボタンの幅 */
    var ANCHOR_DEFAULT_INDEX = 4;            /* 起点ウィジェットの初期位置（4=中央） */
    var KEY_DIALOG_TEXT_WIDTH = 240;         /* 起点ダイアログの説明文の幅 */
    var PRESET_LIST_WIDTH = 120;             /* プリセットのドロップダウンの幅 */
    var ICON_BUTTON_SIZE = 24;               /* 矢印アイコン1個の大きさ */
    var ICON_BUTTON_SPACING = 2;             /* 矢印アイコンどうしの間隔 */
    var ICON_PADDING = 4;                    /* 矢印アイコンの内側の余白 */
    var ICON_ROW_BOTTOM_MARGIN = 5;          /* 矢印アイコン行の下余白 */

    /* 矢印アイコンの配色。UIの明暗に合わせて initThemeColors() で決める */
    var ICON_COLOR;
    var ICON_SELECTED_COLOR;
    var ICON_BG;
    var ICON_SELECTED_BG;
    var ICON_BORDER_COLOR;

    /* 確定／破棄の判定は show() の戻り値に一本化する（ESCやウィンドウを閉じたときは onClick が発火しないため） */
    var DIALOG_RESULT_OK = 1;
    var DIALOG_RESULT_CANCEL = 2;

    /* 行ラベルの幅。パネルごとに setLabelWidth() で切り替える */
    var currentLabelWidth = LABEL_WIDTH;

    /**
     * 以降に作る行ラベルの幅を切り替える
     * @param {number} width - 行ラベルの幅
     * @returns {void}
     */
    function setLabelWidth(width) {
        currentLabelWidth = width;
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

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // ダイアログの位置と不透明度（再利用パーツ）ここまで / End of the reusable dialog position and opacity
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    /**
     * ウィンドウの共通設定を適用する
     * @param {Window} win - 対象のウィンドウ
     * @returns {void}
     */
    function setupWindow(win) {
        win.orientation = "column";
        win.alignChildren = ["fill", "top"];
        win.margins = WINDOW_MARGINS;
        win.spacing = WINDOW_SPACING;
    }

    /**
     * パネルの共通設定を適用する
     * @param {object} panel - 対象のパネル
     * @param {number} spacing - 要素間隔（省略時は共通値）
     * @returns {void}
     */
    function setupPanel(panel, spacing) {
        panel.orientation = "column";
        panel.alignChildren = ["fill", "top"];
        panel.alignment = "fill";
        panel.margins = PANEL_MARGINS;
        panel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /**
     * 見出し付きのパネルを追加する
     * @param {object} parent - 追加先
     * @param {string} title - パネルの見出し
     * @returns {object} 追加したパネル
     */
    function addPanel(parent, title) {
        var panel = parent.add("panel", undefined, title);
        setupPanel(panel);
        return panel;
    }

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // UI の明暗（再利用パーツ） / UI theme (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内に貼る（StepperButtons・LinkToggle の部品より前）。識別子は isDarkUI
    // 2. 配色を明暗で切り替えるときは isDarkUI() を1回だけ呼んで定数に控える
    //      var MY_UI_DARK = isDarkUI();
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

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

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // UI の明暗（再利用パーツ）ここまで / End of the reusable UI theme
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // ステップボタン（再利用パーツ） / Stepper buttons (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内（ローカライズより前）に貼る。
    //    識別子はすべて STEPPER_* / *Stepper* / *Stepped* の名前なので、既存の名前とはぶつからない
    //    UI の明暗は UITheme 部品の isDarkUI() を使う（先に UITheme の ▼〜▲ も貼っておく）
    // 2. コピー先の LABELS.tooltip に stepUp / stepDown / stepUpInteger / stepDownInteger を足す（このファイルの LABELS から写す）。
    //    getLabel() と uiLang はコピー先のものをそのまま使う
    // 3. 数値欄を addSteppedField() で作る。項目名・∧∨・入力欄がひと組で入り、↑↓キーも∧∨と同じ処理で増減する
    //      var widthInput = addSteppedField(parentPanel, {
    //          label: labelText(LABELS.fieldLabel.width), labelWidth: 60,
    //          text: "210 mm", characters: 8, step: 1, min: 1, unit: " mm",
    //          onStep: function (numberInput) { updatePreview(); }
    //      });
    //    値の種類は options で切り分ける:
    //      小数あり（幅・位置など）   … 指定なし（option＋クリックで0.1ずつ）
    //      整数・1以上（段数・個数など）… integer: true, min: 1（0・小数・負数は受け付けず、option＋クリックも1ずつ）
    //      整数・0以上（間隔の数など）  … integer: true, min: 0
    //      範囲つき（％など）           … min: 0, max: 100, unit: "%"
    // 4. 有効／無効は setSteppedFieldEnabled(widthInput, isEnabled)（∧∨のディム表示も切り替わる）。
    //    行・パネルなど親の enabled を切り替えたときは、そのあとで redrawSteppersIn(親) を呼んで∧∨を描き直す
    //    （∧∨は親をたどって無効を判定し、無効の間はクリックも↑↓キーも効かない）
    // 5. 値は parseFloat(widthInput.text) で読む（unit 付きの欄は「210 mm」の形で入っている）
    // 6. この欄に別の↑↓キー処理を付けない（↑↓キーが二重に効く）
    // 既存の edittext をそのまま使うときは、同じ行の group（spacing 0）に addStepper() → edittext の順で置き、
    // bindSteppedArrowKeys(edittext, stepperGroup) を呼ぶ
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

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

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // ステップボタン（再利用パーツ）ここまで / End of the reusable stepper
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

    var LABELS = {
        dialog: {
            title:     { ja: "コネクターを作成", en: "Build Connectors" },
            keyObject: { ja: "起点にするオブジェクト", en: "Choose the Start Object" }
        },
        message: {
            noKeyObject: {
                ja: "キーオブジェクトが設定されていません。起点にするオブジェクトの位置を選んでください。",
                en: "No key object is set. Pick where the object you want to start from sits."
            }
        },
        panel: {
            connector: { ja: "コネクター", en: "Connector" },
            line:      { ja: "線", en: "Line" },
            arrow:     { ja: "線端と矢印", en: "Ends & Arrowheads" }
        },
        fieldLabel: {
            preset:        { ja: "プリセット", en: "Preset" },
            strokeWidth:   { ja: "線幅", en: "Weight" },
            strokeJoin:    { ja: "角の形状", en: "Corner" },
            strokeCap:     { ja: "線端", en: "Cap" },
            dashStyle:     { ja: "破線", en: "Dashes" },
            dashSegments:  { ja: "分割数", en: "Divisions" },
            dashGap:       { ja: "間隔", en: "Gap" },
            lineShape:     { ja: "形状", en: "Shape" },
            startPoint:    { ja: "開始点", en: "Start point" },
            warpType:      { ja: "種類", en: "Style" },
            warpAmount:    { ja: "カーブ", en: "Bend" },
            warpAxis:      { ja: "方向", en: "Axis" },
            cornerRadius:  { ja: "角丸", en: "Corner radius" },
            arrowScale:    { ja: "倍率", en: "Scale" },
            endGap:        { ja: "余白", en: "Offset" },
            arrowPosition: { ja: "位置", en: "Ends" },
            arrowTip:      { ja: "先端", en: "Tip" }
        },
        radio: {
            joinMiter:      { ja: "マイター", en: "Miter" },
            joinRound:      { ja: "ラウンド", en: "Round" },
            joinBevel:      { ja: "ベベル", en: "Bevel" },
            capButt:        { ja: "なし", en: "Butt" },
            capRound:       { ja: "丸形", en: "Round" },
            capProjecting:  { ja: "突出", en: "Projecting" },
            dashNone:       { ja: "なし", en: "None" },
            dashDashed:     { ja: "破線", en: "Dashed" },
            dashDotted:     { ja: "ドット", en: "Dotted" },
            startCenter:    { ja: "各辺の中心", en: "Edge centers" },
            startDivided:   { ja: "等分", en: "Divided" },
            startKeyCenter: { ja: "中心", en: "Center" },
            shapeStraight:  { ja: "直線", en: "Straight" },
            shapeWarp:      { ja: "ワープ", en: "Warp" },
            shapeElbow:     { ja: "カギ", en: "Elbow" },
            shapeBranch:    { ja: "分岐", en: "Branch" },
            shapeCurve:     { ja: "カーブ", en: "Curve" },
            axisAuto:       { ja: "自動", en: "Auto" },
            axisHorizontal: { ja: "水平", en: "Horizontal" },
            axisVertical:   { ja: "垂直", en: "Vertical" },
            arrowNone:      { ja: "なし", en: "None" },
            arrow8:         { ja: "矢印8", en: "Arrow 8" },
            arrow11:        { ja: "矢印11", en: "Arrow 11" },
            dotFilled:      { ja: "黒丸", en: "Filled dot" },
            dotHollow:      { ja: "白丸", en: "Open dot" },
            arrowEnd:       { ja: "終点", en: "End" },
            arrowBoth:      { ja: "両端", en: "Both ends" },
            tipAtEnd:       { ja: "終点に", en: "At end" },
            tipBeyondEnd:   { ja: "終点から", en: "Beyond end" }
        },
        checkbox: {
            unifyStart: { ja: "開始点をまとめる", en: "Share one start point" }
        },
        dropdown: {
            customPreset: { ja: "（カスタム）", en: "(Custom)" }
        },
        unit: {
            blank:   { ja: "", en: "" },
            pt:      { ja: "pt", en: "pt" },
            percent: { ja: "%", en: "%" }
        },
        button: {
            manualKey:    { ja: "手動で設定", en: "Set Manually" },
            savePreset:   { ja: "保存…", en: "Save…" },
            removePreset: { ja: "削除", en: "Delete" },
            cancel:       { ja: "キャンセル", en: "Cancel" }
        },
        /* アクションに埋め込む Illustrator の表示名（UIには出さない）/ names embedded in the action */
        actionName: {
            arrowNone:    { ja: "[なし]", en: "[None]" },
            arrowPrefix:  { ja: "矢印 ", en: "Arrow " },
            tipAtEnd:     { ja: "パスの終点に配置", en: "Place Arrow Tip At End of Path" },
            tipBeyondEnd: { ja: "パスの終点から配置", en: "Extend Arrow Tip Beyond End of Path" }
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
            stepUpInteger: { ja: "値を増やす（shift＋クリックで10の倍数へ）", en: "Increase (Shift-click to snap to 10s)" },
            stepDownInteger: { ja: "値を減らす（shift＋クリックで10の倍数へ）", en: "Decrease (Shift-click to snap to 10s)" },
            strokeJoin: {
                ja: "カギ・分岐の折れ角の見え方です（［線］パネルの角の形状）。",
                en: "How the elbow corners look (the Stroke panel's corner setting)."
            },
            strokeCap: {
                ja: "線の端の見え方です（［線］パネルの線端）。ドットは丸形にしないと点が出ません。",
                en: "How the line ends look (the Stroke panel's cap). Dots need the round cap to show up."
            },
            dashSegments: {
                ja: "線分の数です。両端が線分で終わるように線分の長さを計算します。",
                en: "Number of dashes; the dash length is solved so both ends finish with a dash."
            },
            dashGap: {
                ja: "線分（ドット）どうしのすき間です。ドットでは両端がドットで乗るよう、指定に近い間隔へそろえます。",
                en: "Gap between dashes (dots). For dots it is nudged to the nearest value that lands a dot on both ends."
            },
            keyObject: {
                ja: "キーオブジェクトが未設定のため、起点にするオブジェクトを位置で選びます。選択範囲の左上〜中央〜右下のうち、押した位置にいちばん近いオブジェクトが起点になります。",
                en: "No key object was set, so pick the start object by position: the object nearest the corner or center you click becomes the start."
            },
            manualKey: {
                ja: "ダイアログを閉じます。選択したうえで基準にしたいオブジェクトをもう一度クリックしてキーオブジェクトにしてから、スクリプトを実行し直してください。",
                en: "Closes the dialog. With the objects selected, click the one you want as the key object, then run the script again."
            },
            unifyStart: {
                ja: "すべてのコネクターをキーオブジェクトの同じ位置から出します。位置は右の9分割で選びます。",
                en: "Runs every connector out of the same point on the key object; pick the point with the 3x3 grid on the right."
            },
            unifyAnchor: {
                ja: "開始点をまとめるときの位置です。中央は本数のいちばん多い辺を自動で選びます。上下の中央は上辺・下辺、左右の中央は左辺・右辺、四隅はその高さの左辺・右辺から出します。",
                en: "Where the shared start point sits. Center picks the edge used by the most connectors; the top and bottom cells use those edges, the left and right cells use theirs, and the corners leave from the left or right edge at that height."
            },
            startPoint: {
                ja: "等分は、同じ辺から出るコネクターの本数＋1でその辺を等分し、起点をずらします。中心は、キーオブジェクトの中心から相手へ向かう向きで、起点を辺の上に置きます。",
                en: "Divided spreads the start points along the key object's edge, splitting it into (connectors + 1) parts; Center aims each connector from the key object's center and starts it where that line meets the edge."
            },
            lineShape: {
                ja: "直線はまっすぐ結び、ワープは直線にワープ効果、カギは直角に折れる線、分岐は折れ位置をそろえて幹を共有し、カーブは弧を描いて結びます。",
                en: "Straight connects directly, Warp adds a warp effect to the straight line, Elbow is a right-angled route, Branch shares a trunk with aligned bends, and Curve bows the line into an arc."
            },
            warpAmount: {
                ja: "ワープの曲がり具合と、カーブのふくらみ（線の長さに対する割合）です。マイナス値で向きが逆になります。",
                en: "How much Warp bends, and how far Curve bows out relative to the line length. A negative value flips the direction."
            },
            warpAxis: {
                ja: "自動は全コネクターをまとめて水平／垂直を選びます（線ごとに変えるとアピアランスが混在するため）。線に沿った向きのワープは曲がりません。",
                en: "Auto picks one axis for all the connectors together (a per-line axis would mix their appearances); warping along the line has no visible effect."
            },
            cornerRadius: {
                ja: "カギ・分岐の角を丸めます（0で角丸なし）。",
                en: "Round the corners of Elbow and Branch routes (0 = square corners)."
            },
            preset: {
                ja: "現在の設定に名前を付けて保存できます。保存先はユーザーの設定フォルダーです。",
                en: "Save the current settings under a name; presets are stored in your user settings folder."
            },
            savePreset: {
                ja: "現在の設定に名前を付けて保存します。同じ名前があれば上書きします。",
                en: "Save the current settings under a name; an existing preset with the same name is overwritten."
            },
            arrowShape: {
                ja: "［線］パネルの矢印を使います。黒丸も矢印の一種です。",
                en: "Uses the Stroke panel arrowheads; the dot is an arrowhead preset too."
            },
            arrowScale: {
                ja: "矢印の大きさ（%）。線幅に対する比率です。",
                en: "Arrowhead size in percent, relative to the stroke weight."
            },
            endGap: {
                ja: "終点と相手の図形とのすき間です。最後の線分より大きい値は無視します。キーオブジェクト側は詰めません。",
                en: "Space left between the end of the connector and the object. Values longer than the final segment are ignored; the key-object end is not inset."
            },
            arrowTip: {
                ja: "矢印の先端をパスの終点に配置するか、パスの終点から配置するかを選びます（［線］パネルの先端位置）。",
                en: "Place the arrow tip at the end of the path, or extend it beyond the end (the Stroke panel's tip alignment)."
            },
            arrowPosition: {
                ja: "終点はキーオブジェクトと反対側、両端は起点にも付けます。",
                en: "End = the far side from the key object; Both ends also marks the start."
            }
        },
        alert: {
            presetName:    { ja: "プリセット名を入力してください。", en: "Enter a preset name." },
            presetRemove:  { ja: "このプリセットを削除しますか？", en: "Delete this preset?" },
            presetFailed:  { ja: "プリセットを保存できませんでした。", en: "Could not save the presets." },
            noDocument:    { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            selectObjects: { ja: "2つ以上のオブジェクトを選択してください。", en: "Select two or more objects." }
        }
    };

    /* 使用する矢印。number=矢印番号、scale=倍率（%）、tip=先端位置（0=終点に 1=終点から） */
    /* scale と tip は、その矢印を選んだときに入れる既定値 */
    var ARROW_CHOICES = [
        { label: LABELS.radio.arrowNone, number: 0,                scale: 100, tip: 0 },
        { label: LABELS.radio.arrow8,    number: 8,                scale: 25,  tip: 0, cap: 0 },
        { label: LABELS.radio.arrow11,   number: 11,               scale: 100, tip: 0, cap: 0 },
        { label: LABELS.radio.dotFilled, number: ARROW_DOT_FILLED, scale: 50,  tip: 1, cap: 1 },
        { label: LABELS.radio.dotHollow, number: ARROW_DOT_FILLED, scale: 50,  tip: 1, cap: 1, innerDot: true }
    ];

    /* 角の形状 / Stroke join */
    var STROKE_JOIN_OPTIONS = [
        { label: LABELS.radio.joinMiter, value: StrokeJoin.MITERENDJOIN },
        { label: LABELS.radio.joinRound, value: StrokeJoin.ROUNDENDJOIN },
        { label: LABELS.radio.joinBevel, value: StrokeJoin.BEVELENDJOIN }
    ];

    var STROKE_CAP_OPTIONS = [
        { label: LABELS.radio.capButt,       value: StrokeCap.BUTTENDCAP },
        { label: LABELS.radio.capRound,      value: StrokeCap.ROUNDENDCAP },
        { label: LABELS.radio.capProjecting, value: StrokeCap.PROJECTINGENDCAP }
    ];

    /* 矢印の先端位置 / Arrow tip alignment（ahal の enumerated 値） */
    var ARROW_TIP_OPTIONS = [
        { label: LABELS.radio.tipAtEnd,     name: LABELS.actionName.tipAtEnd,     value: 0 },
        { label: LABELS.radio.tipBeyondEnd, name: LABELS.actionName.tipBeyondEnd, value: 1 }
    ];

    /* ダイアログで選べるワープの種類 / Warp styles offered in the dialog */
    /* style: ワープ効果の並び順（1始まり）／name: ライブエフェクトXMLで使う名前 */
    var WARP_TYPE_CHOICES = [
        { style: 5,  name: "Bulge",   ja: "でこぼこ", en: "Bulge" },
        { style: 14, name: "Squeeze", ja: "絞り込み", en: "Squeeze" }
    ];

    /**
     * ワープの種類の一覧を、現在の言語の表示名で返す
     * @returns {Array<string>} 表示名の配列
     */
    function getWarpTypeLabels() {
        var labels = [];
        for (var i = 0; i < WARP_TYPE_CHOICES.length; i++) {
            labels.push(getLabel(WARP_TYPE_CHOICES[i]));
        }
        return labels;
    }

    // =========================================
    // 選択オブジェクト / Selection
    // =========================================

    if (app.documents.length === 0) {
        alert(getLabel(LABELS.alert.noDocument));
        return;
    }

    var doc = app.activeDocument;
    if (!doc.selection || doc.selection.length < 2) {
        alert(getLabel(LABELS.alert.selectObjects));
        return;
    }

    var selectedItems = [];
    for (var i = 0; i < doc.selection.length; i++) {
        selectedItems.push(doc.selection[i]);
    }

    // Illustratorの座標系（visibleBounds）：
    // 左 = bounds[0] / 上 = bounds[1] / 右 = bounds[2] / 下 = bounds[3]

    /**
     * 外接矩形の中心を返す
     * @param {Array<number>} bounds - [左, 上, 右, 下]
     * @returns {Array<number>} [X, Y]
     */
    function getBoundsCenter(bounds) {
        return [(bounds[0] + bounds[2]) / 2, (bounds[1] + bounds[3]) / 2];
    }

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // 選択の収集と境界（再利用パーツ） / Selection items and bounds (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内に貼る。使わない関数も消さずに残してよい（互いに呼び合う）。
    //    識別子は SELECTION_ITEMS_TOLERANCE / normalizeSelectionItems / resolveTextRangeFrame /
    //    collectSelectionItems / getTextFrameKindKey / collectSelectionTextFrames / collectSelectionPathItems /
    //    isClipMaskItem / getClipMaskItem / hasClippedDescendant / readUsePreviewBoundsPreference /
    //    getClipAwareBounds / filterMeasurableChildren / getClipAwareUnionBounds / isNearlySameCoordinate / areBoundsNearlyEqual
    // 2. 選択は normalizeSelectionItems(doc.selection) で配列にする。文字カーソルの選択（TextRange）は
    //    配列ではなく1個で返り、しかも .length（文字数）を持つので、length だけで配列と見なさない
    // 3. テキストフレーム:
    //      var frames = collectSelectionTextFrames(doc.selection);                           // 全種類
    //      var frames = collectSelectionTextFrames(doc.selection, { kinds: ["point", "path"] });
    //    パス:
    //      var paths = collectSelectionPathItems(doc.selection);                             // 複合パスは中のパスへ
    //      var paths = collectSelectionPathItems(doc.selection, { compoundPaths: "whole", skipClipMasks: true });
    //    それ以外は collectSelectionItems(source, { accept: function (item) { … } }) で条件を書く
    // 4. 並びは選択と同じ前面→背面（グループの中も pageItems の順）。重なり順を使う処理はこの順を前提にしてよい
    // 5. doc.selection に代入し直す配列は skipLocked / skipHidden を true にする。
    //    ロック・非表示を選択に代入すると例外になり、中の子が選択に残る
    // 6. 境界は getClipAwareBounds(item, usePreviewBounds) / getClipAwareUnionBounds(items, usePreviewBounds)。
    //    usePreviewBounds を省くと環境設定の［プレビュー境界を使用］に従う。返り値は [左, 上, 右, 下] の新しい配列
    //    （書き換えても元のオブジェクトに影響しない）。測れないときは null
    // 7. 座標の一致・前後の判定は isNearlySameCoordinate / areBoundsNearlyEqual で許容値を挟む
    //    （吸着させた辺とガイドは 1e-12 ほどずれる）
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

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

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // 選択の収集と境界（再利用パーツ）ここまで / End of the reusable selection items and bounds
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    // =========================================
    // キーオブジェクトの検出 / Key object detection
    // =========================================

    /**
     * スマートガイドの表示を切り替える（実行前後で同じ状態に戻す）
     * @returns {void}
     */
    function toggleSmartGuides() {
        app.executeMenuCommand("edge");
    }

    /**
     * 控えておいた位置へ戻す（例外の後始末でも呼ぶので、1つ失敗しても残りは戻す）
     * @param {Array<object>} items - 対象のオブジェクト配列
     * @param {Array<Array<number>>} positions - [[left, top], ...] の配列
     * @returns {void}
     */
    function restorePositions(items, positions) {
        for (var i = 0; i < items.length; i++) {
            try {
                items[i].left = positions[i][0];
                items[i].top = positions[i][1];
            } catch (e) {
                $.writeln(SCRIPT_NAME + ": 位置の復元に失敗 / failed to restore position — " + e);
            }
        }
    }

    /**
     * 選択オブジェクトからキーオブジェクトを検出する
     * DOMにキーオブジェクトを示すプロパティは無いため、整列コマンドを実行して
     * 「どの向きに整列しても動かないもの」を実測で特定する。
     * 判定中は app.redraw() を呼ばない。描画するとスクリプトの操作がそこで確定し、
     * 整列がそのつど取り消し履歴に積まれてしまう（描画しなければ位置は同期的に読める）。
     * @param {Array<object>} items - 判定対象のオブジェクト配列
     * @returns {number} キーオブジェクトのインデックス。判定できないときは -1
     */
    function detectKeyObjectIndex(items) {
        var alignCommands = ["Horizontal Align Left", "Horizontal Align Right", "Vertical Align Top", "Vertical Align Bottom"];
        var stayedPut = [];
        var originPositions = [];
        var i;
        for (i = 0; i < items.length; i++) {
            stayedPut.push(true);
            originPositions.push([items[i].left, items[i].top]);
        }

        try {
            for (var commandIndex = 0; commandIndex < alignCommands.length; commandIndex++) {
                // 2回目以降だけ元の位置へ戻す（1回目はまだ動かしていない）
                if (commandIndex > 0) restorePositions(items, originPositions);
                app.executeMenuCommand(alignCommands[commandIndex]);
                for (i = 0; i < items.length; i++) {
                    if (!stayedPut[i]) continue;
                    if (Math.abs(items[i].left - originPositions[i][0]) > KEY_DETECT_TOLERANCE_PT ||
                        Math.abs(items[i].top - originPositions[i][1]) > KEY_DETECT_TOLERANCE_PT) {
                        stayedPut[i] = false;
                    }
                }
            }
        } finally {
            // 判定で動かしたぶんを最後に一度だけ戻す（例外で抜けるときも整列結果を残さない）
            restorePositions(items, originPositions);
        }

        var foundIndex = -1;
        for (i = 0; i < items.length; i++) {
            if (!stayedPut[i]) continue;
            if (foundIndex !== -1) return -1; // 複数残った＝判定不能
            foundIndex = i;
        }
        return foundIndex;
    }

    /**
     * 中心Xがもっとも左にあるオブジェクトのインデックスを返す
     * @param {Array<object>} items - 対象のオブジェクト配列
     * @returns {number} インデックス
     */
    function getLeftmostIndex(items) {
        var leftmostIndex = 0;
        var minCenterX = null;
        for (var i = 0; i < items.length; i++) {
            var centerX = getBoundsCenter(getClipAwareBounds(items[i], true))[0];
            if (minCenterX === null || centerX < minCenterX) {
                minCenterX = centerX;
                leftmostIndex = i;
            }
        }
        return leftmostIndex;
    }

    /**
     * 3×3のどの位置かを指定して、そこにいちばん近いオブジェクトのインデックスを返す
     * @param {Array<object>} items - 対象のオブジェクト
     * @param {number} anchorIndex - 0〜8（左上から右下へ）
     * @returns {number} インデックス
     */
    function getIndexAtAnchor(items, anchorIndex) {
        var union = getClipAwareUnionBounds(items, true);
        var targetX = union[0] + (union[2] - union[0]) * ((anchorIndex % 3) / 2);
        var targetY = union[1] - (union[1] - union[3]) * (Math.floor(anchorIndex / 3) / 2);

        var nearestIndex = 0;
        var minDistance = null;
        for (var i = 0; i < items.length; i++) {
            var center = getBoundsCenter(getClipAwareBounds(items[i], true));
            var dx = center[0] - targetX;
            var dy = center[1] - targetY;
            var distance = dx * dx + dy * dy;
            if (minDistance === null || distance < minDistance) {
                minDistance = distance;
                nearestIndex = i;
            }
        }
        return nearestIndex;
    }

    // =========================================
    // 経路の計算 / Connector geometry
    // =========================================

    /**
     * 2つの外接矩形から、向かい合う辺の中央どうしを結ぶ始点・終点を求める
     * @param {Array<number>} fromBounds - 始点側の visibleBounds
     * @param {Array<number>} toBounds - 終点側の visibleBounds
     * @returns {object} points（始点・終点）と horizontal（左右の辺どうしか）
     */
    function getConnectionPoints(fromBounds, toBounds) {
        var fromCenter = getBoundsCenter(fromBounds);
        var toCenter = getBoundsCenter(toBounds);
        var dx = toCenter[0] - fromCenter[0];
        var dy = toCenter[1] - fromCenter[1];

        // 横のずれが大きければ左右の辺、そうでなければ上下の辺でつなぐ
        var side;
        if (Math.abs(dx) >= Math.abs(dy)) {
            side = (dx >= 0) ? "right" : "left";
        } else {
            side = (dy >= 0) ? "top" : "bottom";
        }
        return getConnectionPointsOnSide(fromBounds, toBounds, side);
    }

    /**
     * 辺を指定して、始点・終点を求める
     * @param {Array<number>} fromBounds - 始点側の visibleBounds
     * @param {Array<number>} toBounds - 終点側の visibleBounds
     * @param {string} side - 使う辺（"right" / "left" / "top" / "bottom"）
     * @returns {object} points（始点・終点）と horizontal（左右の辺どうしか）
     */
    function getConnectionPointsOnSide(fromBounds, toBounds, side) {
        var fromCenter = getBoundsCenter(fromBounds);
        var toCenter = getBoundsCenter(toBounds);
        var fromCenterX = fromCenter[0];
        var fromCenterY = fromCenter[1];
        var toCenterX = toCenter[0];
        var toCenterY = toCenter[1];

        var route;
        if (side === "right") {
            route = { points: [[fromBounds[2], fromCenterY], [toBounds[0], toCenterY]], horizontal: true, side: side };
        } else if (side === "left") {
            route = { points: [[fromBounds[0], fromCenterY], [toBounds[2], toCenterY]], horizontal: true, side: side };
        } else if (side === "top") {
            route = { points: [[fromCenterX, fromBounds[1]], [toCenterX, toBounds[3]]], horizontal: false, side: side };
        } else {
            route = { points: [[fromCenterX, fromBounds[3]], [toCenterX, toBounds[1]]], horizontal: false, side: side };
        }
        // 等分配置で辺に沿って並べ替えるための基準
        route.order = route.horizontal ? toCenterY : toCenterX;
        return route;
    }

    /**
     * いちばん本数の多い辺を返す
     * @param {Array<object>} paths - getConnectionPoints() の結果の配列
     * @returns {string} 辺の名前
     */
    function getMajoritySide(paths) {
        var counts = {};
        var majoritySide = paths[0].side;
        for (var i = 0; i < paths.length; i++) {
            var side = paths[i].side;
            counts[side] = (counts[side] || 0) + 1;
            if (counts[side] > counts[majoritySide]) majoritySide = side;
        }
        return majoritySide;
    }

    /**
     * 外接矩形の中心から指定した点へ向かう線が、矩形の辺と交わる位置を求める
     * @param {Array<number>} bounds - visibleBounds
     * @param {Array<number>} toPoint - 向かう先の座標 [X, Y]
     * @returns {Array<number>} 辺の上の座標 [X, Y]
     */
    function getEdgePoint(bounds, toPoint) {
        var center = getBoundsCenter(bounds);
        var centerX = center[0];
        var centerY = center[1];
        var halfWidth = (bounds[2] - bounds[0]) / 2;
        var halfHeight = (bounds[1] - bounds[3]) / 2;
        var dx = toPoint[0] - centerX;
        var dy = toPoint[1] - centerY;
        if (dx === 0 && dy === 0) return [centerX, centerY];
        // 左右・上下それぞれの辺に届くまでの比率のうち、小さいほうが先に交わる辺
        var ratioX = (dx === 0) ? null : halfWidth / Math.abs(dx);
        var ratioY = (dy === 0) ? null : halfHeight / Math.abs(dy);
        var ratio = (ratioX === null) ? ratioY : ((ratioY === null) ? ratioX : Math.min(ratioX, ratioY));
        return [centerX + dx * ratio, centerY + dy * ratio];
    }

    /**
     * 9分割のどの位置かから、使う辺を求める
     * @param {number} anchorIndex - 0〜8（左上から右下へ）。4=中央は自動
     * @returns {string} 辺の名前。中央のときは null
     */
    function getSideForAnchor(anchorIndex) {
        if (anchorIndex === 1) return "top";
        if (anchorIndex === 7) return "bottom";
        if (anchorIndex === 3 || anchorIndex === 0 || anchorIndex === 6) return "left";
        if (anchorIndex === 5 || anchorIndex === 2 || anchorIndex === 8) return "right";
        return null; // 中央は本数のいちばん多い辺にまかせる
    }

    /**
     * すべてのコネクターを、同じ位置から出し直した経路を返す
     * @param {number} anchorIndex - 開始点の位置（0〜8）
     * @returns {Array<object>} 経路の配列
     */
    function getUnifiedPaths(anchorIndex) {
        var side = getSideForAnchor(anchorIndex);
        if (!side) side = getMajoritySide(connectorPaths);

        // 辺に沿った向きの位置は、9分割の行（左右の辺）または列（上下の辺）で決める
        var horizontal = (side === "left" || side === "right");
        var ratio = (horizontal ? Math.floor(anchorIndex / 3) : (anchorIndex % 3)) / 2;
        var startPerp = horizontal
            ? keyBounds[1] - (keyBounds[1] - keyBounds[3]) * ratio
            : keyBounds[0] + (keyBounds[2] - keyBounds[0]) * ratio;

        var paths = [];
        for (var i = 0; i < selectedItems.length; i++) {
            if (i === keyIndex) continue;
            var path = getConnectionPointsOnSide(keyBounds, getClipAwareBounds(selectedItems[i], true), side);
            path.points[0][horizontal ? 1 : 0] = startPerp;
            paths.push(path);
        }
        return paths;
    }

    /**
     * 経路を、出ていく辺ごとにまとめる
     * @param {Array<object>} routes - 経路の配列
     * @returns {object} 辺の名前をキーにした経路の配列
     */
    function groupRoutesBySide(routes) {
        var routesBySide = {};
        for (var i = 0; i < routes.length; i++) {
            var side = routes[i].side;
            if (!routesBySide[side]) routesBySide[side] = [];
            routesBySide[side].push(routes[i]);
        }
        return routesBySide;
    }

    /**
     * 開始点の指定を反映した経路の複製を返す
     * @param {number} startPoint - 開始点（0=各辺の中心 1=等分 2=中心）
     * @param {boolean} unifyStart - すべてを同じ位置から出すか
     * @param {number} unifyAnchor - まとめるときの位置（0〜8）
     * @returns {Array<object>} points（座標）と horizontal（左右接続か）の配列
     */
    function getConnectorRoutes(startPoint, unifyStart, unifyAnchor) {
        var routes = [];
        var i;
        var paths = unifyStart ? getUnifiedPaths(unifyAnchor) : connectorPaths;
        for (i = 0; i < paths.length; i++) {
            var path = paths[i];
            routes.push({
                points: [[path.points[0][0], path.points[0][1]], [path.points[1][0], path.points[1][1]]],
                horizontal: path.horizontal,
                side: path.side,
                order: path.order
            });
        }
        if (startPoint === 2) {
            // キーオブジェクトの中心から相手へ向かう線を、キーオブジェクトの辺で止める
            for (i = 0; i < routes.length; i++) {
                var start = getEdgePoint(keyBounds, routes[i].points[1]);
                routes[i].points[0][0] = start[0];
                routes[i].points[0][1] = start[1];
            }
            return routes;
        }
        if (startPoint !== 1) return routes;

        // 同じ辺から出るコネクターごとに、その辺を（本数＋1）等分して起点をずらす
        var routesBySide = groupRoutesBySide(routes);
        for (var side in routesBySide) {
            if (!routesBySide.hasOwnProperty(side)) continue;
            var sideRoutes = routesBySide[side];
            // 線が交差しないよう、相手の位置順に辺へ割り当てる
            sideRoutes.sort(function (a, b) {
                return a.order - b.order;
            });
            for (var k = 0; k < sideRoutes.length; k++) {
                var ratio = (k + 1) / (sideRoutes.length + 1);
                if (sideRoutes[k].horizontal) {
                    sideRoutes[k].points[0][1] = keyBounds[3] + (keyBounds[1] - keyBounds[3]) * ratio;
                } else {
                    sideRoutes[k].points[0][0] = keyBounds[0] + (keyBounds[2] - keyBounds[0]) * ratio;
                }
            }
        }
        return routes;
    }

    /**
     * 折れ位置が選択オブジェクトにかからないよう、いちばん近い空き位置へずらす
     * @param {number} bendPosition - もとの折れ位置
     * @param {boolean} horizontal - 左右の辺どうしをつなぐか（true なら折れ位置はX）
     * @param {number} lower - 折れ位置に使える下限
     * @param {number} upper - 折れ位置に使える上限
     * @param {number} perpLower - 折れ線が伸びる向きの下限
     * @param {number} perpUpper - 折れ線が伸びる向きの上限
     * @returns {number} ずらした折れ位置。よけられないときはもとの位置
     */
    function getClearBendPosition(bendPosition, horizontal, lower, upper, perpLower, perpUpper) {
        var clearPosition = bendPosition;
        // 重なった図形の外へ寄せる。図形が並んでいることもあるので選択数だけ繰り返す
        for (var pass = 0; pass < selectedItems.length; pass++) {
            var moved = false;
            for (var i = 0; i < selectedItems.length; i++) {
                var bounds = getClipAwareBounds(selectedItems[i], true);
                var minAxis = horizontal ? bounds[0] : bounds[3];
                var maxAxis = horizontal ? bounds[2] : bounds[1];
                var minPerp = horizontal ? bounds[3] : bounds[0];
                var maxPerp = horizontal ? bounds[1] : bounds[2];

                var before = minAxis - BEND_CLEARANCE_PT;
                var after = maxAxis + BEND_CLEARANCE_PT;
                if (clearPosition <= before || clearPosition >= after) continue;
                if (maxPerp < perpLower || minPerp > perpUpper) continue; // 折れ線が届かない位置

                clearPosition = (clearPosition - before <= after - clearPosition) ? before : after;
                moved = true;
                break;
            }
            if (!moved) break;
        }
        // すき間の外まで押し出されるならよけられない
        if (clearPosition < lower || clearPosition > upper) return bendPosition;
        return clearPosition;
    }

    /**
     * カギ線（直角に折れる経路）の座標列を作る
     * @param {Array<Array<number>>} points - [[始点X, 始点Y], [終点X, 終点Y]]
     * @param {boolean} horizontal - 左右の辺どうしをつなぐか
     * @returns {Array<Array<number>>} 座標の配列
     */
    function getElbowPoints(points, horizontal) {
        var start = points[0];
        var end = points[1];
        var axis = horizontal ? 0 : 1;   /* 折れ位置を決める軸 / axis of the bend position */
        var perpAxis = 1 - axis;         /* 折れ線が伸びる軸 / axis the bend segment runs along */
        if (Math.abs(end[perpAxis] - start[perpAxis]) < COORD_TOLERANCE_PT) return [start, end];

        // 折れ位置は2つの図形のすき間の中央
        var bendPosition = getClearBendPosition((start[axis] + end[axis]) / 2, horizontal,
            Math.min(start[axis], end[axis]), Math.max(start[axis], end[axis]),
            Math.min(start[perpAxis], end[perpAxis]), Math.max(start[perpAxis], end[perpAxis]));
        var bendStart = [start[0], start[1]];
        var bendEnd = [end[0], end[1]];
        bendStart[axis] = bendPosition;
        bendEnd[axis] = bendPosition;
        return [start, bendStart, bendEnd, end];
    }

    /**
     * 分岐用に、同じ辺から出る経路で共有する折れ位置を求める
     * いちばん近い図形とのすき間の中央にそろえ、幹を1本にまとめる
     * @param {Array<object>} routes - getConnectorRoutes() の戻り値
     * @returns {object} 辺をキーにした折れ位置
     */
    function getSharedBendPositions(routes) {
        var nearestRoutes = {};
        var perpRanges = {};
        var i;
        for (i = 0; i < routes.length; i++) {
            var route = routes[i];
            var axis = route.horizontal ? 0 : 1;
            var perpAxis = route.horizontal ? 1 : 0;
            var routeSide = route.side;
            var delta = route.points[1][axis] - route.points[0][axis];
            if (!nearestRoutes[routeSide] || Math.abs(delta) < Math.abs(nearestRoutes[routeSide].delta)) {
                nearestRoutes[routeSide] = { route: route, delta: delta };
            }
            // 折れ線（背骨）が伸びる範囲。よける図形を絞り込むのに使う
            var perpStart = route.points[0][perpAxis];
            var perpEnd = route.points[1][perpAxis];
            if (!perpRanges[routeSide]) {
                perpRanges[routeSide] = [Math.min(perpStart, perpEnd), Math.max(perpStart, perpEnd)];
            } else {
                perpRanges[routeSide][0] = Math.min(perpRanges[routeSide][0], perpStart, perpEnd);
                perpRanges[routeSide][1] = Math.max(perpRanges[routeSide][1], perpStart, perpEnd);
            }
        }

        var bendPositions = {};
        for (var side in nearestRoutes) {
            if (!nearestRoutes.hasOwnProperty(side)) continue;
            var nearest = nearestRoutes[side];
            var bendAxis = nearest.route.horizontal ? 0 : 1;
            var startCoord = nearest.route.points[0][bendAxis];
            var endCoord = nearest.route.points[1][bendAxis];
            bendPositions[side] = getClearBendPosition(startCoord + nearest.delta / 2, nearest.route.horizontal,
                Math.min(startCoord, endCoord), Math.max(startCoord, endCoord),
                perpRanges[side][0], perpRanges[side][1]);
        }
        return bendPositions;
    }

    // =========================================
    // 作図 / Drawing
    // =========================================

    /**
     * 指定名のレイヤーを取得する。無ければ作成する
     * @param {string} name - レイヤー名
     * @returns {object} layer（レイヤー）、existed（既存だったか）、locked／visible（変更前の状態）
     */
    function getOrCreateLayer(name) {
        for (var i = 0; i < doc.layers.length; i++) {
            if (doc.layers[i].name === name) {
                var existingLayer = doc.layers[i];
                var layerState = {
                    layer: existingLayer,
                    existed: true,
                    locked: existingLayer.locked,
                    visible: existingLayer.visible
                };
                existingLayer.locked = false;
                existingLayer.visible = true;
                return layerState;
            }
        }
        var newLayer = doc.layers.add();
        newLayer.name = name;
        return { layer: newLayer, existed: false, locked: false, visible: true };
    }

    /**
     * RGBColorを作る
     * @param {number} red - 赤（0〜255）
     * @param {number} green - 緑（0〜255）
     * @param {number} blue - 青（0〜255）
     * @returns {RGBColor} 生成した色
     */
    function createRGBColor(red, green, blue) {
        var color = new RGBColor();
        color.red = red;
        color.green = green;
        color.blue = blue;
        return color;
    }

    /**
     * ドキュメントのカラーモードに合わせた無彩色を作る
     * @param {number} blackPercent - 黒の割合（0=白、100=黒）
     * @returns {object} CMYKColor または RGBColor
     */
    function createGrayColor(blackPercent) {
        if (doc.documentColorSpace === DocumentColorSpace.CMYK) {
            var cmykColor = new CMYKColor();
            cmykColor.cyan = 0;
            cmykColor.magenta = 0;
            cmykColor.yellow = 0;
            cmykColor.black = blackPercent;
            return cmykColor;
        }
        var level = Math.round(255 * (100 - blackPercent) / 100);
        return createRGBColor(level, level, level);
    }

    /**
     * ライブエフェクトを適用する（失敗しても続行）
     * @param {PathItem} item - 適用対象のパス
     * @param {string} xml - ライブエフェクトのXML
     * @returns {void}
     */
    function applyLiveEffect(item, xml) {
        try {
            item.applyEffect(xml);
        } catch (e) {
            $.writeln(SCRIPT_NAME + ": ライブエフェクトの適用に失敗 / failed to apply effect — " + e);
        }
    }

    /**
     * ワープ効果を適用する
     * @param {PathItem} item - 適用対象のパス
     * @param {object} settings - ダイアログの設定
     * @param {boolean} isVertical - 垂直方向のワープにするか
     * @returns {void}
     */
    function applyWarpEffect(item, settings, isVertical) {
        applyLiveEffect(item, '<LiveEffect name="Adobe Deform"><Dict data="' +
            'S DisplayString Warp:' + settings.warpName +
            ' I DeformStyle ' + settings.warpStyle +
            ' B Rotate ' + (isVertical ? 1 : 0) +
            ' R DeformValue ' + (settings.warpAmount / 100) +
            ' R DeformHoriz ' + (WARP_DEFORM_H / 100) +
            ' R DeformVert ' + (WARP_DEFORM_V / 100) +
            ' "/></LiveEffect>');
    }

    /**
     * 2点のパスを弧にする（中点を線と垂直な向きへふくらませる）
     * 中点から振った1点を2次ベジェの制御点とみなし、両端のハンドルへ置き換える
     * @param {PathItem} pathItem - 対象のパス（アンカーが2点のもの）
     * @param {number} amountPercent - カーブ（%）。マイナスで反対側へふくらむ
     * @returns {void}
     */
    function bendIntoCurve(pathItem, amountPercent) {
        if (!amountPercent || pathItem.pathPoints.length !== 2) return;

        var startPoint = pathItem.pathPoints[0];
        var endPoint = pathItem.pathPoints[1];
        var from = startPoint.anchor;
        var to = endPoint.anchor;
        var dx = to[0] - from[0];
        var dy = to[1] - from[1];
        var length = Math.sqrt(dx * dx + dy * dy);
        if (length <= 0) return;

        // ふくらみは線の長さに比例させ、進行方向の左向き（法線）へ振る
        var sag = length / 2 * (amountPercent / 100);
        var control = [
            (from[0] + to[0]) / 2 - dy / length * sag,
            (from[1] + to[1]) / 2 + dx / length * sag
        ];
        // 2次ベジェの制御点を3次ベジェのハンドルに置き換える比率
        var handleRatio = 2 / 3;
        startPoint.rightDirection = [
            from[0] + (control[0] - from[0]) * handleRatio,
            from[1] + (control[1] - from[1]) * handleRatio
        ];
        endPoint.leftDirection = [
            to[0] + (control[0] - to[0]) * handleRatio,
            to[1] + (control[1] - to[1]) * handleRatio
        ];
    }

    /**
     * 角丸効果を適用する
     * @param {PathItem} item - 適用対象のパス
     * @param {number} radius - 半径（pt）
     * @returns {void}
     */
    function applyRoundCornersEffect(item, radius) {
        applyLiveEffect(item, '<LiveEffect name="Adobe Round Corners"><Dict data="R radius ' + radius + ' "/></LiveEffect>');
    }

    /**
     * パスの長さを返す
     * @param {PathItem} pathItem - 対象のパス
     * @returns {number} 長さ（pt）。取得できないときは0
     */
    function getPathLength(pathItem) {
        var pathLength = pathItem.length;
        return (typeof pathLength === "number" && pathLength > 0) ? pathLength : 0;
    }

    /**
     * パスの長さに合わせた破線の線分・間隔を求める（両端を調整）
     * 線分の本数を n とすると n×線分＋(n−1)×間隔＝全長
     * @param {number} pathLength - パスの長さ（pt）
     * @param {number} segments - 分割数（線分の本数）
     * @param {number} gap - 間隔（pt）
     * @returns {Array<number>} strokeDashes に渡す配列
     */
    function calcFittedDashes(pathLength, segments, gap) {
        if (pathLength <= 0) return [];
        if (segments < 2) return []; // 1本＝実線

        var dash = (pathLength - gap * (segments - 1)) / segments;
        if (dash > 0) return [dash, gap];

        // 間隔が大きすぎて線分が残らないときは、線分と間隔を等分する
        var evenLength = pathLength / (segments * 2 - 1);
        return [evenLength, evenLength];
    }

    /**
     * 指定の間隔に近く、両端にドットが乗る間隔を求める
     * @param {number} pathLength - パスの長さ（pt）
     * @param {number} gap - ドットどうしの間隔（pt）
     * @returns {Array<number>} strokeDashes に渡す配列
     */
    function calcFittedDots(pathLength, gap) {
        if (pathLength <= 0) return [];

        var dotGaps = (gap > 0) ? Math.round(pathLength / gap) : 1;
        if (dotGaps < 1) dotGaps = 1;
        return [0, pathLength / dotGaps];
    }

    /**
     * 角の形状と破線を適用する（アクションで上書きされた後にも呼ぶ）
     * @param {PathItem} pathItem - 対象のパス
     * @param {object} settings - ダイアログの設定
     * @returns {void}
     */
    function applyStrokeStyle(pathItem, settings) {
        pathItem.strokeJoin = STROKE_JOIN_OPTIONS[settings.strokeJoin].value;
        pathItem.strokeCap = STROKE_CAP_OPTIONS[settings.strokeCap].value;
        applyDashStyle(pathItem, settings);
    }

    /**
     * 破線の設定を適用する（線分・間隔はパスの長さに合わせて計算する）
     * @param {PathItem} pathItem - 対象のパス
     * @param {object} settings - ダイアログの設定
     * @returns {void}
     */
    function applyDashStyle(pathItem, settings) {
        var pathLength = getPathLength(pathItem);
        if (settings.dashStyle === 1) {
            pathItem.strokeDashes = calcFittedDashes(pathLength, settings.dashSegments, settings.dashGap);
        } else if (settings.dashStyle === 2) {
            // 線分0＋丸い線端で点線にする
            pathItem.strokeDashes = calcFittedDots(pathLength, settings.dashGap);
        } else {
            pathItem.strokeDashes = [];
        }
    }

    /**
     * コネクターの線を作成し、必要ならライブエフェクトを適用する
     * @param {object} container - 作成先のグループまたはレイヤー
     * @param {Array<Array<number>>} points - 座標の配列
     * @param {object} settings - ダイアログの設定
     * @param {boolean} isVerticalWarp - ワープを垂直方向にするか
     * @returns {PathItem} 作成したパス
     */
    function drawConnectorLine(container, points, settings, isVerticalWarp) {
        var connector = container.pathItems.add();
        connector.setEntirePath(points);
        connector.closed = false;
        connector.filled = false;
        connector.stroked = true;
        connector.strokeWidth = settings.strokeWidth;
        connector.strokeColor = createGrayColor(100);
        // 破線はパスの長さに合わせるので、曲げてから線の設定を入れる
        if (settings.lineShape === 4) {
            bendIntoCurve(connector, settings.warpAmount);
        }
        applyStrokeStyle(connector, settings);

        if (settings.lineShape === 1) {
            applyWarpEffect(connector, settings, isVerticalWarp);
        }

        if (settings.cornerRadius > 0) {
            applyRoundCornersEffect(connector, settings.cornerRadius);
        }
        return connector;
    }

    /**
     * 白丸用に、パスの端点へ白い●を重ねる
     * @param {object} container - 作成先のグループまたはレイヤー
     * @param {Array<number>} point - [X, Y] 端点の座標（●の中心）
     * @param {object} settings - ダイアログの設定
     * @returns {PathItem} 作成したパス。作成できないときは null
     */
    function drawInnerDot(container, point, settings) {
        // 黒丸の矢印は倍率で大きくなるので、白い●も同じ比率で合わせる
        var innerDiameter = settings.strokeWidth * WHITE_DOT_RATIO * (settings.arrowScale / WHITE_DOT_BASE_SCALE);
        if (innerDiameter <= 0) return null;

        var innerRadius = innerDiameter / 2;
        var innerDot = container.pathItems.ellipse(point[1] + innerRadius, point[0] - innerRadius, innerDiameter, innerDiameter);
        innerDot.stroked = false;
        innerDot.filled = true;
        innerDot.fillColor = createGrayColor(0);
        return innerDot;
    }

    // =========================================
    // 矢印（一時アクション）/ Arrowheads via a temporary action
    // =========================================
    // 矢印はDOMから設定できないため、［線］パネルの設定を行うアクションを生成して実行する。
    // Arrowheads cannot be reached from the DOM, so a temporary action is generated and played.

    /**
     * 矢印番号から、アクションに渡す矢印名を作る
     * @param {number} arrowNumber - ［線］パネルの矢印番号
     * @returns {string} 矢印名
     */
    function getArrowName(arrowNumber) {
        return getLabel(LABELS.actionName.arrowPrefix) + arrowNumber;
    }

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // 一時アクション（再利用パーツ） / Temporary action (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内に貼る。
    //    識別子は runTemporaryAction / loadTemporaryActionSet / unloadTemporaryActionSet / toActionHex / buildActionNameLines
    // 2. アクション定義は配列＋join("\n") で組み立てる（''' は ES3 の構文エラー）。
    //    セット名・アクション名は英数字にする。/name [ n 16進 ] は buildActionNameLines で作るとバイト数がずれない
    //      var actionSource = [
    //          "/version 3"
    //      ].concat(buildActionNameLines("", "MySet"), [
    //          "/isOpen 1", "/actionCount 1", "/action-1 {"
    //      ], buildActionNameLines("\t", "myAction"), [ … ]).join("\n");
    // 3. 1回だけ実行するとき:
    //      if (!runTemporaryAction(actionSource, "MySet", "myAction")) alert(getLabel("alert.actionFailed"));
    //    何度も実行するとき（オブジェクトごとなど）は、読み込み・解除を1回ずつにする:
    //      if (!loadTemporaryActionSet(actionSource, "MySet")) { alert(…); return; }
    //      try { for (…) app.doScript("myAction", "MySet"); } finally { unloadTemporaryActionSet("MySet"); }
    // 4. 失敗は例外にせず false で返す（$.writeln に理由を出す）。警告を出すかはコピー先で決める
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

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

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // 一時アクション（再利用パーツ）ここまで / End of the reusable temporary action
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    /**
     * 5 → "5.0" のように必ず小数点を含む文字列にする
     * @param {number} value - 変換する数値
     * @returns {string} 小数点を含む文字列
     */
    function toRealString(value) {
        var realText = String(Number(value));
        if (realText.indexOf(".") === -1 && realText.indexOf("e") === -1) realText += ".0";
        return realText;
    }

    /**
     * アクションセット名・アクション名の /name 行を作る
     * @param {string} name - 名前
     * @returns {string} /name 行
     */
    function buildNameLine(name) {
        var hexText = toActionHex(name);
        return "/name [ " + (hexText.length / 2) + " \n\t" + hexText + "\n]\n";
    }

    /**
     * パラメータブロックの外枠を作る
     * @param {number} index - パラメータ番号
     * @param {number} key - パラメータキー
     * @param {string} body - ブロックの中身
     * @returns {string} パラメータブロック
     */
    function buildParamBlock(index, key, body) {
        return "\t\t/parameter-" + index + " {\n" +
            "\t\t\t/key " + key + "\n" +
            "\t\t\t/showInPalette 4294967295\n" +
            body +
            "\t\t}\n";
    }

    /**
     * 単位付き実数パラメータを作る
     * @param {number} index - パラメータ番号
     * @param {number} key - パラメータキー
     * @param {number} value - 値
     * @param {number} unitCode - 単位コード
     * @returns {string} パラメータブロック
     */
    function buildUnitRealParam(index, key, value, unitCode) {
        return buildParamBlock(index, key,
            "\t\t\t/type (unit real)\n" +
            "\t\t\t/value " + toRealString(value) + "\n" +
            "\t\t\t/unit " + unitCode + "\n");
    }

    /**
     * 列挙パラメータを作る（16進の表示名を直接指定）
     * @param {number} index - パラメータ番号
     * @param {number} key - パラメータキー
     * @param {string} nameHex - 表示名の16進表現
     * @param {number} byteLength - 表示名のバイト数
     * @param {number} value - 値
     * @returns {string} パラメータブロック
     */
    function buildEnumParam(index, key, nameHex, byteLength, value) {
        return buildParamBlock(index, key,
            "\t\t\t/type (enumerated)\n" +
            "\t\t\t/name [ " + byteLength + " \n\t\t\t\t" + nameHex + "\n\t\t\t]\n" +
            "\t\t\t/value " + value + "\n");
    }

    /**
     * 列挙パラメータを表示名から作る
     * @param {number} index - パラメータ番号
     * @param {number} key - パラメータキー
     * @param {string} name - 表示名
     * @param {number} value - 値
     * @returns {string} パラメータブロック
     */
    function buildEnumParamByName(index, key, name, value) {
        var hexText = toActionHex(name);
        return buildEnumParam(index, key, hexText, hexText.length / 2, value);
    }

    /**
     * 整数パラメータを作る
     * @param {number} index - パラメータ番号
     * @param {number} key - パラメータキー
     * @param {number} value - 値
     * @returns {string} パラメータブロック
     */
    function buildIntParam(index, key, value) {
        return buildParamBlock(index, key, "\t\t\t/type (integer)\n\t\t\t/value " + value + "\n");
    }

    /**
     * 真偽値パラメータを作る
     * @param {number} index - パラメータ番号
     * @param {number} key - パラメータキー
     * @param {boolean} value - 値
     * @returns {string} パラメータブロック
     */
    function buildBoolParam(index, key, value) {
        return buildParamBlock(index, key, "\t\t\t/type (boolean)\n\t\t\t/value " + (value ? 1 : 0) + "\n");
    }

    /**
     * 実数パラメータを作る
     * @param {number} index - パラメータ番号
     * @param {number} key - パラメータキー
     * @param {number} value - 値
     * @returns {string} パラメータブロック
     */
    function buildRealParam(index, key, value) {
        return buildParamBlock(index, key, "\t\t\t/type (real)\n\t\t\t/value " + toRealString(value) + "\n");
    }

    /**
     * Unicode文字列パラメータを作る
     * @param {number} index - パラメータ番号
     * @param {number} key - パラメータキー
     * @param {string} value - 値
     * @returns {string} パラメータブロック
     */
    function buildUStrParam(index, key, value) {
        var hexText = toActionHex(value);
        return buildParamBlock(index, key,
            "\t\t\t/type (ustring)\n" +
            "\t\t\t/value [ " + (hexText.length / 2) + " \n\t\t\t\t" + hexText + "\n\t\t\t]\n");
    }

    /**
     * ［線］の設定（矢印を含む）を行うアクションのソースを組み立てる
     * @param {object} settings - ダイアログの設定
     * @returns {string} アクションファイルの内容
     */
    function buildStrokeActionSource(settings) {
        return "/version 3\n" +
            buildNameLine(ACTION_SET_NAME) +
            "/isOpen 1\n" +
            "/actionCount 1\n" +
            "/action-1 {\n" +
            "\t" + buildNameLine(ACTION_NAME) +
            "\t/keyIndex 0\n" +
            "\t/colorIndex 0\n" +
            "\t/isOpen 1\n" +
            "\t/eventCount 1\n" +
            "\t/event-1 {\n" +
            "\t\t/useRulersIn1stQuadrant 0\n" +
            "\t\t/internalName (ai_plugin_setStroke)\n" +
            "\t\t/localizedName [ 0 \n\t\t]\n" +
            "\t\t/isOpen 1\n" +
            "\t\t/isOn 1\n" +
            "\t\t/hasDialog 0\n" +
            "\t\t/parameterCount 11\n" +
            /* 可変 / variable */
            buildUnitRealParam(1, KEY_STROKE_WIDTH, settings.strokeWidth, UNIT_POINT) +
            buildUStrParam(6, KEY_ARROW_HEAD_1, settings.startArrow) +
            buildUStrParam(7, KEY_ARROW_HEAD_2, settings.endArrow) +
            buildRealParam(8, KEY_ARROW_SCALE_1, settings.arrowScale) +
            buildRealParam(9, KEY_ARROW_SCALE_2, settings.arrowScale) +
            buildEnumParamByName(10, KEY_ARROW_ALIGN, getLabel(settings.arrowTip.name), settings.arrowTip.value) +
            /* 以下は記録した .aia のまま / recorded as-is */
            buildEnumParam(2, KEY_CAP, "e4b8b8e59e8be7b79ae7abaf", 12, 1) +             /* 線端: 丸型線端 */
            buildEnumParam(3, KEY_JOIN, "e383a9e382a6e383b3e38389e7b590e59088", 18, 1) + /* 角の形状: ラウンド結合 */
            buildIntParam(4, KEY_DASH_INT, 0) +
            buildBoolParam(5, KEY_DASH_BOOL, 0) +
            buildEnumParam(11, KEY_ALIGN, "e4b8ade5a4ae", 6, 0) +                        /* 線の位置: 中央 */
            "\t}\n" +
            "}\n";
    }

    /**
     * コネクターの線に矢印を設定する（選択を作ってアクションを1回だけ実行）
     * @param {Array<PathItem>} connectorLines - 対象のパス
     * @param {object} settings - ダイアログの設定
     * @returns {void}
     */
    function applyArrowheads(connectorLines, settings) {
        if (!connectorLines.length) return;

        doc.selection = null;
        for (var i = 0; i < connectorLines.length; i++) {
            connectorLines[i].selected = true;
        }
        app.redraw(); // 選択が反映されていないとアクションが空振りする
        if (!runTemporaryAction(buildStrokeActionSource(settings), ACTION_SET_NAME, ACTION_NAME)) {
            // 矢印が付かなくても線は残す / keep the lines even if the arrowheads fail
            $.writeln(SCRIPT_NAME + ": 矢印の設定に失敗 / failed to set arrowheads");
        }
        doc.selection = null;

        // アクションは角の形状・線端・破線も上書きするので、線種を戻す
        for (i = 0; i < connectorLines.length; i++) {
            applyStrokeStyle(connectorLines[i], settings);
        }
    }

    // =========================================
    // コネクターの作成 / Building connectors
    // =========================================

    toggleSmartGuides();

    var keyIndex = -1;
    try {
        keyIndex = detectKeyObjectIndex(selectedItems);
    } catch (e) {
        // 検出できなくても起点ダイアログで選べるので続行する
        $.writeln(SCRIPT_NAME + ": キーオブジェクトの検出に失敗 / key object detection failed — " + e);
    }

    if (keyIndex === -1) {
        if (selectedItems.length > 2) {
            // キーオブジェクトが無いときは、起点を別ダイアログで選んでもらう
            initThemeColors();
            keyIndex = chooseKeyIndex(selectedItems);
            if (keyIndex === -1) {
                toggleSmartGuides();
                return;
            }
        } else {
            keyIndex = getLeftmostIndex(selectedItems); // 2つのときは左側を起点にする
        }
    }

    var keyBounds;
    var connectorPaths;

    /**
     * 起点にするオブジェクトを決め、経路を作り直す
     * @param {number} index - selectedItems のインデックス
     * @returns {void}
     */
    function setKeyIndex(index) {
        keyIndex = index;
        keyBounds = getClipAwareBounds(selectedItems[keyIndex], true);
        connectorPaths = [];
        for (var i = 0; i < selectedItems.length; i++) {
            if (i === keyIndex) continue;
            connectorPaths.push(getConnectionPoints(keyBounds, getClipAwareBounds(selectedItems[i], true)));
        }
    }

    setKeyIndex(keyIndex);

    var connectorLayerState = getOrCreateLayer(getLabel(CONNECTOR_LAYER_NAME));
    var connectorLayer = connectorLayerState.layer;
    var createdConnectors = [];

    /**
     * 作成済みのコネクターを削除する
     * @returns {void}
     */
    function removeConnectors() {
        for (var i = 0; i < createdConnectors.length; i++) {
            createdConnectors[i].remove();
        }
        createdConnectors = [];
    }

    /**
     * 1本ぶんの座標列を、形状の指定に合わせて作る（分岐は getBranchLines() で作る）
     * @param {object} route - getConnectorRoutes() の要素
     * @param {object} settings - ダイアログの設定
     * @returns {Array<Array<number>>} 座標の配列
     */
    function getLinePoints(route, settings) {
        if (settings.lineShape === 2) return getElbowPoints(route.points, route.horizontal);
        return route.points;
    }

    /**
     * 幹の片側にある枝を、幹に近い順に並べて返す
     * @param {Array<object>} sideRoutes - 同じ辺から出る経路
     * @param {number} trunkPosition - 幹の位置
     * @param {boolean} horizontal - 左右の辺どうしをつなぐか
     * @param {number} direction - 1=幹より先、-1=幹より手前（同じ位置は手前に含める）
     * @returns {Array<object>} 並べ替えた経路
     */
    function getBranchOrder(sideRoutes, trunkPosition, horizontal, direction) {
        var axis = horizontal ? 1 : 0;
        var branchRoutes = [];
        for (var i = 0; i < sideRoutes.length; i++) {
            var delta = sideRoutes[i].points[1][axis] - trunkPosition;
            if ((direction > 0) ? (delta > 0) : (delta <= 0)) branchRoutes.push(sideRoutes[i]);
        }
        branchRoutes.sort(function (a, b) {
            return Math.abs(a.points[1][axis] - trunkPosition) - Math.abs(b.points[1][axis] - trunkPosition);
        });
        return branchRoutes;
    }

    /**
     * 幹の片側の枝の座標列を作る。幹からいちばん遠い枝だけが幹から折れる長いパスになり、
     * 手前の枝はその背骨から横に出すだけにして重なりをなくす
     * @param {Array<object>} orderedRoutes - 幹に近い順に並べた経路
     * @param {number} trunkPosition - 幹の位置
     * @param {number} bendPosition - 折れ位置
     * @param {boolean} horizontal - 左右の辺どうしをつなぐか
     * @returns {Array<Array<Array<number>>>} 枝の座標列
     */
    function getSideBranches(orderedRoutes, trunkPosition, bendPosition, horizontal) {
        var branches = [];
        var farthestIndex = orderedRoutes.length - 1;
        for (var i = 0; i < orderedRoutes.length; i++) {
            var end = orderedRoutes[i].points[1];
            var position = horizontal ? end[1] : end[0];
            var isFarthest = (i === farthestIndex);
            var startPosition = isFarthest ? trunkPosition : position;
            var start = horizontal ? [bendPosition, startPosition] : [startPosition, bendPosition];
            if (isFarthest && Math.abs(position - trunkPosition) >= COORD_TOLERANCE_PT) {
                branches.push([start, horizontal ? [bendPosition, position] : [position, bendPosition], end]);
            } else {
                branches.push([start, end]);
            }
        }
        return branches;
    }

    /**
     * 分岐の経路を、幹（キーから折れ位置まで）と枝（折れ位置から相手まで）に分ける
     * @param {Array<object>} routes - getConnectorRoutes() の結果
     * @returns {object} trunks（幹の座標列）と branches（枝の座標列）
     */
    function getBranchLines(routes) {
        var sharedBends = getSharedBendPositions(routes);
        var routesBySide = groupRoutesBySide(routes);
        var trunks = [];
        var branches = [];
        for (var side in routesBySide) {
            if (!routesBySide.hasOwnProperty(side)) continue;
            var sideRoutes = routesBySide[side];
            var horizontal = sideRoutes[0].horizontal;
            var bendPosition = sharedBends[side];

            // 幹は各起点の平均の位置（左右の辺なら高さ、上下の辺なら左右）に置く
            var startSum = 0;
            for (var i = 0; i < sideRoutes.length; i++) {
                startSum += sideRoutes[i].points[0][horizontal ? 1 : 0];
            }
            var trunkPosition = startSum / sideRoutes.length;
            var edgePosition = sideRoutes[0].points[0][horizontal ? 0 : 1];

            // 矢印がキーオブジェクト側の端に付くよう、幹は折れ位置からキーへ向けて引く
            trunks.push(horizontal
                ? [[bendPosition, trunkPosition], [edgePosition, trunkPosition]]
                : [[trunkPosition, bendPosition], [trunkPosition, edgePosition]]);

            // 幹の前後それぞれで、いちばん遠い枝が背骨を兼ねる（手前の枝は横だけなので重ならない）
            branches = branches.concat(getSideBranches(getBranchOrder(sideRoutes, trunkPosition, horizontal, 1), trunkPosition, bendPosition, horizontal));
            branches = branches.concat(getSideBranches(getBranchOrder(sideRoutes, trunkPosition, horizontal, -1), trunkPosition, bendPosition, horizontal));
        }
        return { trunks: trunks, branches: branches };
    }

    /**
     * 終点を進行方向の手前へ戻す（相手の図形とのすき間を作る）
     * @param {Array<Array<number>>} points - 座標の配列
     * @param {number} gap - 空けるすき間（pt）
     * @returns {Array<Array<number>>} 終点をずらした座標の配列
     */
    function applyEndGap(points, gap) {
        if (!(gap > 0) || points.length < 2) return points;

        var lastIndex = points.length - 1;
        var from = points[lastIndex - 1];
        var to = points[lastIndex];
        var dx = to[0] - from[0];
        var dy = to[1] - from[1];
        var length = Math.sqrt(dx * dx + dy * dy);
        if (length <= gap) return points; // 最後の線分より大きいすき間は詰めない

        var shortened = [];
        for (var i = 0; i < lastIndex; i++) {
            shortened.push([points[i][0], points[i][1]]);
        }
        shortened.push([to[0] - dx / length * gap, to[1] - dy / length * gap]);
        return shortened;
    }

    /**
     * 始点側の矢印を外した設定の複製を返す（分岐で合流点に矢印を出さないため）
     * @param {object} settings - ダイアログの設定
     * @returns {object} 複製した設定
     */
    function getEndArrowOnlySettings(settings) {
        var endOnlySettings = {};
        for (var key in settings) {
            if (settings.hasOwnProperty(key)) endOnlySettings[key] = settings[key];
        }
        endOnlySettings.startArrow = getLabel(LABELS.actionName.arrowNone);
        return endOnlySettings;
    }

    /**
     * ワープの軸を決める。自動のときは全線をまとめて1つに決める
     * @param {Array<Array<Array<number>>>} allLinePoints - 全コネクターの座標列
     * @param {number} warpAxis - 0=自動 1=水平 2=垂直
     * @returns {boolean} 垂直方向のワープにするか
     */
    function getWarpAxis(allLinePoints, warpAxis) {
        if (warpAxis === 1) return false;
        if (warpAxis === 2) return true;

        // 線に沿った向きのワープは曲がらないので、いちばん曲がりにくい線でも
        // 幅（高さ）を確保できるほうの軸を選ぶ
        var minWidth = null;
        var minHeight = null;
        for (var i = 0; i < allLinePoints.length; i++) {
            var points = allLinePoints[i];
            var lastIndex = points.length - 1;
            var width = Math.abs(points[lastIndex][0] - points[0][0]);
            var height = Math.abs(points[lastIndex][1] - points[0][1]);
            if (minWidth === null || width < minWidth) minWidth = width;
            if (minHeight === null || height < minHeight) minHeight = height;
        }
        if (minWidth === null) return false;
        return minHeight > minWidth;
    }

    /**
     * 線と白丸をグループにまとめ、作成済みリストの項目をそのグループに差し替える
     * @param {PathItem} line - コネクターの線
     * @param {Array<Array<number>>} dotPoints - 白丸を置く座標
     * @param {object} settings - ダイアログの設定
     * @returns {void}
     */
    function groupWithInnerDots(line, dotPoints, settings) {
        var dotGroup = connectorLayer.groupItems.add();
        line.move(dotGroup, ElementPlacement.PLACEATEND);
        // グループへ追加した順に前面へ入るので、白丸は線より後に作る
        for (var i = 0; i < dotPoints.length; i++) {
            drawInnerDot(dotGroup, dotPoints[i], settings);
        }
        for (var k = 0; k < createdConnectors.length; k++) {
            if (createdConnectors[k] === line) {
                createdConnectors[k] = dotGroup;
                return;
            }
        }
        createdConnectors.push(dotGroup);
    }

    /**
     * 経路から、描く線の座標列を作る（分岐は幹と枝に分ける）
     * @param {Array<object>} routes - getConnectorRoutes() の結果
     * @param {object} settings - ダイアログの設定
     * @returns {object} trunks（幹の座標列。分岐以外は空）と lines（相手側で終わる線の座標列）
     */
    function getConnectorLinePoints(routes, settings) {
        var trunks = [];
        var lines = [];
        var i;
        if (settings.lineShape === 3) {
            // 分岐は幹（キーから折れ位置まで）と枝（折れ位置から相手まで）に分けて作る
            var branchLines = getBranchLines(routes);
            trunks = branchLines.trunks;
            lines = branchLines.branches;
        } else {
            for (i = 0; i < routes.length; i++) {
                lines.push(getLinePoints(routes[i], settings));
            }
        }
        // 幹はキーオブジェクト側で終わるので、すき間は相手側で終わる線だけに入れる
        for (i = 0; i < lines.length; i++) {
            lines[i] = applyEndGap(lines[i], settings.endGap);
        }
        return { trunks: trunks, lines: lines };
    }

    /**
     * 座標列ごとに線を作り、作成済みリストに加える
     * @param {Array<Array<Array<number>>>} linePointsList - 座標列の配列
     * @param {object} settings - ダイアログの設定
     * @param {boolean} isVerticalWarp - ワープを垂直方向にするか
     * @returns {Array<PathItem>} 作成したパス
     */
    function drawConnectorLines(linePointsList, settings, isVerticalWarp) {
        var drawnLines = [];
        for (var i = 0; i < linePointsList.length; i++) {
            var line = drawConnectorLine(connectorLayer, linePointsList[i], settings, isVerticalWarp);
            createdConnectors.push(line);
            drawnLines.push(line);
        }
        return drawnLines;
    }

    /**
     * 現在の設定でコネクターを作り直す（プレビュー兼本番）
     * @param {object} settings - ダイアログの設定
     * @returns {void}
     */
    function buildConnectors(settings) {
        removeConnectors();
        var routes = getConnectorRoutes(settings.startPoint, settings.unifyStart, settings.unifyAnchor);
        var isBranch = (settings.lineShape === 3);
        var isBothEnds = (settings.arrowPosition === 1);
        var linePoints = getConnectorLinePoints(routes, settings);

        // ワープの軸は全線で同じにする（線ごとに変えるとアピアランスが混在する）
        var isVerticalWarp = getWarpAxis(linePoints.lines, settings.warpAxis);
        var trunkLines = drawConnectorLines(linePoints.trunks, settings, isVerticalWarp);
        var connectorLines = drawConnectorLines(linePoints.lines, settings, isVerticalWarp);

        if (settings.hasArrow) {
            if (isBranch) {
                // 枝は終点だけ。両端のときは幹のキー側の端が始点側の矢印になる
                var arrowLines = isBothEnds ? connectorLines.concat(trunkLines) : connectorLines;
                applyArrowheads(arrowLines, getEndArrowOnlySettings(settings));
            } else {
                applyArrowheads(connectorLines, settings);
            }
        }

        if (!settings.arrowInnerDot) return;

        // 白丸は矢印の丸の上に重ねるので、線より後に作り、その線とグループ化する
        var i;
        var points;
        if (isBothEnds) {
            // 両端のときは幹のキー側にも印を付ける
            for (i = 0; i < trunkLines.length; i++) {
                points = linePoints.trunks[i];
                groupWithInnerDots(trunkLines[i], [points[points.length - 1]], settings);
            }
        }
        for (i = 0; i < connectorLines.length; i++) {
            points = linePoints.lines[i];
            var dotPoints = [points[points.length - 1]];
            if (!isBranch && isBothEnds) dotPoints.push(points[0]);
            groupWithInnerDots(connectorLines[i], dotPoints, settings);
        }
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 文字列を数値に変換する。数値にならないときは既定値を返す
     * @param {string} inputText - 入力文字列
     * @param {number} fallback - 既定値
     * @returns {number} 数値
     */
    function parseNumberInput(inputText, fallback) {
        // Number("") は 0 になるため、空欄は既定値に落とす
        var trimmedText = String(inputText).replace(/^\s+|\s+$/g, "");
        if (trimmedText === "") return fallback;
        var parsedValue = Number(trimmedText);
        return isNaN(parsedValue) ? fallback : parsedValue;
    }

    /**
     * 選択されているラジオのインデックスを返す
     * @param {Array<object>} radios - ラジオボタンの配列
     * @returns {number} インデックス
     */
    function getSelectedRadioIndex(radios) {
        for (var i = 0; i < radios.length; i++) {
            if (radios[i].value) return i;
        }
        return 0;
    }

    /**
     * 右揃えラベル付きの1行を追加する
     * @param {object} parent - 追加先のパネル
     * @param {object} labelSet - 行ラベルの定義
     * @returns {object} 追加した行グループ
     */
    function addFieldRow(parent, labelSet) {
        var row = parent.add("group");
        row.orientation = "row";
        row.alignment = ["fill", "top"];
        row.alignChildren = ["left", "center"];
        var rowLabel = row.add("statictext", undefined, labelText(labelSet));
        rowLabel.preferredSize.width = currentLabelWidth;
        rowLabel.justify = "right";
        row.label = rowLabel; // 縦並びのときに上揃えへ変えられるよう控える
        return row;
    }

    /**
     * ラベル付きのラジオボタン行を追加する
     * @param {object} parent - 追加先のパネル
     * @param {object} labelSet - 行ラベルの定義
     * @param {Array<object>} optionSets - 各ラジオのラベル定義
     * @param {number} defaultIndex - 初期選択のインデックス
     * @param {object} tipSet - helpTip の定義（省略可）
     * @param {boolean} vertical - ラジオを縦並びにするか（省略可）
     * @returns {object} row（行グループ）と radios（ラジオの配列）
     */
    function addRadioRow(parent, labelSet, optionSets, defaultIndex, tipSet, vertical) {
        var row = addFieldRow(parent, labelSet);
        var radioParent = row;
        if (vertical) {
            // ラベルは1つ目のラジオに合わせて上揃えにする
            row.alignChildren = ["left", "top"];
            row.label.alignment = ["left", "top"];
            radioParent = row.add("group");
            radioParent.orientation = "column";
            radioParent.alignChildren = ["left", "center"];
            radioParent.spacing = RADIO_COLUMN_SPACING;
        }
        var radios = [];
        for (var i = 0; i < optionSets.length; i++) {
            var radio = radioParent.add("radiobutton", undefined, getLabel(optionSets[i]));
            if (tipSet) radio.helpTip = getLabel(tipSet);
            radios.push(radio);
        }
        radios[defaultIndex].value = true;
        return { row: row, radios: radios };
    }

    /**
     * UIが明るいテーマかを判定する
     * @returns {boolean} 明るいテーマなら true
     */
    function isLightUI() {
        return app.preferences.getRealPreference("uiBrightness") > 0.5;
    }

    /**
     * 表示中のコントロールを描き直す（表示前は notify が例外になるので無視する）
     * @param {object} control - 描き直すコントロール
     * @returns {void}
     */
    function redrawControl(control) {
        try { control.notify("onDraw"); } catch (e) {}
    }

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // 基準点ウィジェット（再利用パーツ） / Anchor widget (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内（ダイアログを作る関数より前）に貼る。
    //    識別子は ANCHOR_WIDGET_* / *AnchorWidget* / getAnchor* の名前。貼る前に、コピー先にある旧版の
    //    ANCHOR_WIDGET_SIZE・ANCHOR_CELL_*・ANCHOR_CONNECTIONS・ANCHOR_*_COLOR / FILL・addAnchorWidget・drawAnchorWidget・
    //    drawAnchorCell・redrawAnchorWidget・initAnchorColors（9軸の配色だけを決めているもの）・clampGridIndex を消す
    //    UI の明暗は UITheme 部品の isDarkUI() を使う（先に UITheme の ▼〜▲ も貼っておく）
    // 2. ウィジェットを作る。初期値は 0〜8（0=左上, 4=中央, 8=右下）か名前（"topLeft" / "top" / "topRight" /
    //    "left" / "center" / "right" / "bottomLeft" / "bottom" / "bottomRight"）
    //      var anchorWidget = addAnchorWidget(anchorPanel, "center", function (anchorIndex) { updatePreview(); });
    //      anchorWidget.helpTip = getLabel(LABELS.tooltip.anchor);
    //    onChange はクリックのたびに呼ぶ（同じセルでも呼ぶ）。setAnchorWidgetValue() からは呼ばない
    //    未選択（-1）を許すときは addAnchorWidget(parent, -1, onChange, { allowNone: true })
    //    選べないセルは { disabledCells: [4] } か setAnchorWidgetCellsDisabled(anchorWidget, [4])（薄く描き、クリックも無視）
    // 3. 値を読む: getAnchorWidgetIndex(anchorWidget) … 0〜8（未選択は -1）/ getAnchorWidgetName(anchorWidget) … "topLeft" など
    //    値を書く: setAnchorWidgetValue(anchorWidget, 2) または setAnchorWidgetValue(anchorWidget, "topRight")（描き直す）
    //    旧版の widget.selectedAnchorIndex / widget.anchorIndex への直接代入は描き直されないので使わない
    // 4. Illustrator の変形に渡す:
    //      pageItem.resize(150, 150, true, true, true, true, 150, getAnchorTransformation(getAnchorWidgetIndex(anchorWidget)));
    //    座標で使うときは getAnchorPointOnBounds(geometricBounds, anchorIndex) → [x, y]、
    //    割合で使うときは getAnchorRatio(anchorIndex) → [0|0.5|1, 0|0.5|1]（左上が [0, 0]）、
    //    シンボル登録は getAnchorSymbolRegistrationPoint(anchorIndex)
    // 5. 有効／無効は setAnchorWidgetEnabled(anchorWidget, isEnabled)（薄い色で描き直し、クリックも無視）。
    //    パネル・行など親の enabled を切り替えたときは、そのあとで redrawAnchorWidgetsIn(親) を呼ぶ
    //    （親の無効化は子の enabled に出ないので、描画とクリックの判定は親までたどる）
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

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

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // 基準点ウィジェット（再利用パーツ）ここまで / End of the reusable anchor widget
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    /**
     * UIの明暗に合わせて、矢印アイコンの配色を決める
     * @returns {void}
     */
    function initThemeColors() {
        var lightUI = isLightUI();
        ICON_COLOR          = lightUI ? [0.25, 0.25, 0.25, 1] : [0.85, 0.85, 0.85, 1];
        ICON_SELECTED_COLOR = lightUI ? [1, 1, 1, 1]          : [0.15, 0.15, 0.15, 1];
        ICON_BG             = lightUI ? [1, 1, 1, 1]          : [0.22, 0.22, 0.22, 1];
        ICON_SELECTED_BG    = lightUI ? [0.4, 0.4, 0.4, 1]    : [0.8, 0.8, 0.8, 1];
        ICON_BORDER_COLOR   = lightUI ? [0.65, 0.65, 0.65, 1] : [0.45, 0.45, 0.45, 1];
    }

    /**
     * 矢印の形状アイコンを描く（線・実線矢印・線矢印・黒丸・白丸）
     * @param {object} iconButton - 描画対象のボタン
     * @returns {void}
     */
    function drawArrowShapeIcon(iconButton) {
        var graphics = iconButton.graphics;
        var width = iconButton.size[0];
        var height = iconButton.size[1];
        var selected = iconButton.isSelected;

        var backColor = selected ? ICON_SELECTED_BG : ICON_BG;
        graphics.rectPath(0, 0, width, height);
        graphics.fillPath(graphics.newBrush(graphics.BrushType.SOLID_COLOR, backColor));
        graphics.rectPath(0, 0, width, height);
        graphics.strokePath(graphics.newPen(graphics.PenType.SOLID_COLOR, ICON_BORDER_COLOR, 1));

        var color = selected ? ICON_SELECTED_COLOR : ICON_COLOR;
        var thickPen = graphics.newPen(graphics.PenType.SOLID_COLOR, color, 3);
        var thinPen = graphics.newPen(graphics.PenType.SOLID_COLOR, color, 1);
        var iconBrush = graphics.newBrush(graphics.BrushType.SOLID_COLOR, color);
        var backBrush = graphics.newBrush(graphics.BrushType.SOLID_COLOR, backColor);

        var left = ICON_PADDING;
        var right = width - ICON_PADDING;
        var centerY = Math.round(height / 2);
        var headSize = 5;
        var dotRadius = 4;

        /**
         * 軸線を引く
         * @param {number} endX - 線の右端
         * @param {object} pen - 使うペン
         * @returns {void}
         */
        function drawShaft(endX, pen) {
            graphics.newPath();
            graphics.moveTo(left, centerY);
            graphics.lineTo(endX, centerY);
            graphics.strokePath(pen);
        }

        if (iconButton.arrowIndex === 0) {
            drawShaft(right, thickPen);
            return;
        }
        if (iconButton.arrowIndex === 1) {
            // 実線の矢印：軸は太く、先端は塗りの三角
            drawShaft(right - headSize - 1, thickPen);
            graphics.newPath();
            graphics.moveTo(right - headSize - 1, centerY - headSize);
            graphics.lineTo(right, centerY);
            graphics.lineTo(right - headSize - 1, centerY + headSize);
            graphics.closePath();
            graphics.fillPath(iconBrush);
            return;
        }
        if (iconButton.arrowIndex === 2) {
            // 線の矢印：軸も先端も細い線
            drawShaft(right, thinPen);
            graphics.newPath();
            graphics.moveTo(right - headSize, centerY - headSize);
            graphics.lineTo(right, centerY);
            graphics.lineTo(right - headSize, centerY + headSize);
            graphics.strokePath(thinPen);
            return;
        }

        // 黒丸・白丸：軸の先に円を置く
        drawShaft(right - dotRadius * 2 + 1, thickPen);
        graphics.ellipsePath(right - dotRadius * 2, centerY - dotRadius, dotRadius * 2, dotRadius * 2);
        graphics.fillPath(iconButton.arrowIndex === 3 ? iconBrush : backBrush);
        if (iconButton.arrowIndex === 4) {
            graphics.ellipsePath(right - dotRadius * 2, centerY - dotRadius, dotRadius * 2, dotRadius * 2);
            graphics.strokePath(graphics.newPen(graphics.PenType.SOLID_COLOR, color, 2));
        }
    }

    /**
     * アイコンボタンの選択状態を返す
     * @param {Array<object>} buttons - アイコンボタンの配列
     * @returns {number} 選択中のインデックス
     */
    function getSelectedIconIndex(buttons) {
        for (var i = 0; i < buttons.length; i++) {
            if (buttons[i].isSelected) return i;
        }
        return 0;
    }

    /**
     * アイコンボタンの選択状態を切り替える
     * @param {Array<object>} buttons - アイコンボタンの配列
     * @param {number} index - 選択するインデックス
     * @returns {void}
     */
    function selectIconButton(buttons, index) {
        if (!(index >= 0) || index >= buttons.length) return;
        for (var i = 0; i < buttons.length; i++) {
            buttons[i].isSelected = (i === index);
            redrawControl(buttons[i]);
        }
    }

    /**
     * 矢印の形状をアイコンで選ぶ行を追加する（行ラベルなし）
     * @param {object} parent - 追加先のパネル
     * @param {number} defaultIndex - 初期選択のインデックス
     * @param {function} onSelect - 選び直したときに呼ぶ処理
     * @returns {object} row（行グループ）と buttons（アイコンボタンの配列）
     */
    function addArrowShapeRow(parent, defaultIndex, onSelect) {
        var row = parent.add("group");
        row.orientation = "row";
        row.alignment = ["center", "top"];
        row.alignChildren = ["left", "center"];
        row.spacing = ICON_BUTTON_SPACING;
        row.margins = [0, 0, 0, ICON_ROW_BOTTOM_MARGIN];

        var buttons = [];
        for (var i = 0; i < ARROW_CHOICES.length; i++) {
            var iconButton = row.add("button", undefined, "");
            iconButton.minimumSize = [ICON_BUTTON_SIZE, ICON_BUTTON_SIZE];
            iconButton.preferredSize = [ICON_BUTTON_SIZE, ICON_BUTTON_SIZE];
            iconButton.maximumSize = [ICON_BUTTON_SIZE, ICON_BUTTON_SIZE];
            iconButton.arrowIndex = i;
            iconButton.isSelected = false;
            iconButton.helpTip = getLabel(ARROW_CHOICES[i].label) + "  —  " + getLabel(LABELS.tooltip.arrowShape);
            iconButton.onDraw = function () {
                drawArrowShapeIcon(this);
            };
            iconButton.onClick = function () {
                selectIconButton(buttons, this.arrowIndex);
                onSelect();
            };
            buttons.push(iconButton);
        }
        selectIconButton(buttons, defaultIndex);
        return { row: row, buttons: buttons };
    }

    /**
     * 起点にするオブジェクトを位置で選ぶダイアログを出す
     * @param {Array<object>} items - 選択したオブジェクト
     * @returns {number} 起点にするインデックス。閉じたときは -1
     */
    function chooseKeyIndex(items) {
        var anchorIndex = ANCHOR_DEFAULT_INDEX;

        var keyDialog = new Window("dialog", getLabel(LABELS.dialog.keyObject));
        setupWindow(keyDialog);

        var messageText = keyDialog.add("statictext", undefined, getLabel(LABELS.message.noKeyObject), { multiline: true });
        messageText.preferredSize.width = KEY_DIALOG_TEXT_WIDTH;

        var widgetGroup = keyDialog.add("group");
        widgetGroup.alignment = ["center", "top"];
        var anchorWidget = addAnchorWidget(widgetGroup, anchorIndex, function (index) {
            anchorIndex = index;
        });
        anchorWidget.helpTip = getLabel(LABELS.tooltip.keyObject);

        /* ボタン行（左：手動で設定、右：キャンセル・OK） / Button row (left: set manually, right: Cancel and OK) */
        var keyButtonRow = addButtonRow(keyDialog);
        var btnManualKey = keyButtonRow.leftGroup.add("button", undefined, getLabel(LABELS.button.manualKey));
        btnManualKey.helpTip = getLabel(LABELS.tooltip.manualKey);
        btnManualKey.onClick = function () {
            // 手動で設定してもらうため、そのまま閉じる
            keyDialog.close(DIALOG_RESULT_CANCEL);
        };
        var btnKeyCancel = keyButtonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });
        btnKeyCancel.onClick = function () {
            keyDialog.close(DIALOG_RESULT_CANCEL);
        };
        var btnKeyOK = keyButtonRow.rightGroup.add("button", undefined, "OK", { name: "ok" });
        btnKeyOK.onClick = function () {
            keyDialog.close(DIALOG_RESULT_OK);
        };

        prepareDialogWindow(keyDialog, SCRIPT_NAME + "_keyObject");
        if (keyDialog.show() !== DIALOG_RESULT_OK) return -1;
        return getIndexAtAnchor(items, anchorIndex);
    }

    /**
     * ラベル付きの数値入力行を追加する
     * @param {object} parent - 追加先のパネル
     * @param {object} labelSet - 行ラベルの定義
     * @param {number} value - 初期値
     * @param {object} unitSet - 単位表示の定義
     * @param {object} tipSet - helpTip の定義（省略可）
     * @param {Array<number>} range - [最小値, 最大値]。渡すとスライダーを付ける（省略可）
     * @param {object} [stepOptions] - ∧∨の増減の条件（integer など。addStepper() の stepOptions）
     * @returns {object} row（行グループ）、input（入力欄）、slider（スライダーまたは null）、stepOptions（∧∨の条件）
     */
    function addNumberRow(parent, labelSet, value, unitSet, tipSet, range, stepOptions) {
        var row = addFieldRow(parent, labelSet);

        /* ∧∨と入力欄は隙間0で突き合わせる / butt the stepper against the field */
        var stepperFieldGroup = row.add("group");
        stepperFieldGroup.orientation = "row";
        stepperFieldGroup.alignChildren = ["left", "center"];
        stepperFieldGroup.spacing = 0;
        stepperFieldGroup.margins = 0;

        /* 下限と増減後の処理は bindPreviewToField() で埋める / min and onStep are filled in by bindPreviewToField() */
        stepOptions = stepOptions || {};
        var input;
        var stepperGroup = addStepper(stepperFieldGroup, function () { return input; }, stepOptions);
        input = stepperFieldGroup.add("edittext", undefined, String(value));
        input.characters = FIELD_CHARS;
        /* ↑↓キーも∧∨と同じ処理で増減する / arrow keys share the stepper's logic */
        bindSteppedArrowKeys(input, stepperGroup);
        row.add("statictext", undefined, getLabel(unitSet));
        if (tipSet) input.helpTip = getLabel(tipSet);

        var slider = null;
        if (range) {
            slider = row.add("slider", undefined, value, range[0], range[1]);
            // 固定幅で広げるとダイアログが太るので、余りを吸わせる
            slider.minimumSize.width = SLIDER_MIN_WIDTH;
            slider.alignment = ["fill", "center"];
            if (tipSet) slider.helpTip = getLabel(tipSet);
        }
        return { row: row, input: input, slider: slider, stepOptions: stepOptions };
    }

    /**
     * 数値欄の行の有効／無効を切り替え、行内の∧∨を描き直す（自作描画は自動でディムにならないため）。
     * 状態が変わらないときは何もしない（プレビューのたびに描き直さない）
     * @param {Group} fieldRow - addNumberRow() で作った行
     * @param {boolean} isEnabled - 有効にするなら true
     * @returns {void}
     */
    function setFieldRowEnabled(fieldRow, isEnabled) {
        if (fieldRow.enabled === isEnabled) return;
        fieldRow.enabled = isEnabled;
        redrawSteppersIn(fieldRow);
    }

    /**
     * 入力欄の値をスライダーへ反映する（範囲外は丸める）
     * @param {object} field - addNumberRow が返したフィールド
     * @returns {void}
     */
    function syncSliderToInput(field) {
        if (!field.slider) return;
        var value = Number(field.input.text);
        if (isNaN(value)) return;
        if (value < field.slider.minvalue) value = field.slider.minvalue;
        if (value > field.slider.maxvalue) value = field.slider.maxvalue;
        field.slider.value = value;
    }

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // 設定の保存（再利用パーツ） / Settings store (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内に貼る。
    //    識別子は SETTINGS_STORE_* / createSettingsStore / readSettingsLegacyFile / readSettingsLegacyPreference / settingsStore*
    // 2. 寿命は今のスクリプトに合わせて選ぶ。
    //      "session"    … $.global に置く。Illustrator を終了するまで残る。#targetengine が必須（無いと毎回消える）
    //      "persistent" … Folder.userData/illustrator-scripts/<storeName>.json に書く。再起動しても残る
    //    storeName はふつう SCRIPT_NAME。ダイアログの位置は DialogPosition の部品が持つので、ここには入れない
    // 3. 既定値を1か所にまとめ、load で受け取る。戻り値は毎回新しいオブジェクト（書き換えても保存されない）
    //      var settingsStore = createSettingsStore(SCRIPT_NAME, "persistent");
    //      var DEFAULT_SETTINGS = { widthPt: 10, addFrame: true, modeKey: "fit", corners: { tl: 0, tr: 0 } };
    //      var dialogSettings = settingsStore.load(DEFAULT_SETTINGS);
    //      …OK で閉じたら…
    //      settingsStore.save({ widthPt: …, addFrame: …, modeKey: …, corners: { tl: …, tr: … } });
    //    型は既定値に合わせる（数値の既定値には "12" も 12 として読む。真偽は "1"/"0"/"true"/"false" も読む）。
    //    合わない値・既定値に無い項目は捨てて既定値を使う。{} と null の既定値は中身を問わずそのまま受け取る
    //    （名前をキーにしたプリセット集など）。配列は配列ならそのまま受け取る
    // 4. 保存できるのは文字列・数値・真偽・null と、その配列・入れ子のオブジェクトだけ。
    //    DOM オブジェクト・File・関数は入れない（パスは fsName の文字列で持つ）。長さは pt で持つ
    // 5. 旧形式の設定を読み継ぐときは、3つ目の引数に legacy 関数を渡す。
    //    新しい保存が1度も無いとき（ファイルが無い・$.global に無い）だけ呼ばれ、戻り値を保存値として既定値と突き合わせる。
    //    旧ファイル・旧キーは消さない。キー名が変わったときは legacy の中で詰め替える
    //      createSettingsStore(SCRIPT_NAME, "persistent", { legacy: function () {
    //          return readSettingsLegacyFile(Folder.userData + "/" + SCRIPT_NAME + "/settings.txt");  … key=value / toSource / JSON を自動判別
    //      } });
    //      createSettingsStore(SCRIPT_NAME, "persistent", { legacy: function () {
    //          return readSettingsLegacyPreference("SmartTextFindReplace/settings");  … app.preferences の文字列
    //      } });
    // 6. clear() は保存を消す。legacy を渡したストアでは空の保存（{}）を書き、旧設定が戻ってこないようにする
    // 7. 失敗は例外にせず、load は既定値、save は false を返す（$.writeln に理由を出す）
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

    var SETTINGS_STORE_FOLDER_NAME = "illustrator-scripts"; /* Folder.userData の下に作るフォルダー / folder created under Folder.userData */
    var SETTINGS_STORE_MAX_DEPTH = 32;                                /* 入れ子の上限（循環参照よけ）/ nesting limit (guards against cycles) */

    /**
     * 設定の保存先を作る。寿命は "session"（Illustrator の終了まで）か "persistent"（ファイルに保存）
     * @param {string} storeName - 保存名（ふつうは SCRIPT_NAME）。ファイル名と $.global のキーに使う
     * @param {string} lifetime - "session" または "persistent"
     * @param {Object} [storeOptions] - { legacy: function () → 旧形式の保存値のオブジェクト|null }
     * @returns {{load: Function, save: Function, clear: Function}} 読み込み・保存・消去の関数
     */
    function createSettingsStore(storeName, lifetime, storeOptions) {
        var isPersistent = (lifetime === "persistent");
        var legacyReader = (storeOptions && typeof storeOptions.legacy === "function") ? storeOptions.legacy : null;
        var safeStoreName = String(storeName).replace(/[\\\/:*?"<>|]/g, "_");
        var sessionKey = "__" + safeStoreName + "_Settings";
        var settingsFile = isPersistent
            ? new File(Folder.userData + "/" + SETTINGS_STORE_FOLDER_NAME + "/" + safeStoreName + ".json")
            : null;

        /**
         * 保存してある文字列を返す
         * @returns {string|null} 保存文字列。1度も保存していなければ null
         */
        function readStoredText() {
            if (!isPersistent) {
                return (typeof $.global[sessionKey] === "string") ? $.global[sessionKey] : null;
            }
            return settingsStoreReadTextFile(settingsFile);
        }

        /**
         * 文字列を保存する
         * @param {string} storedText - 保存する文字列
         * @returns {boolean} 保存できたら true
         */
        function writeStoredText(storedText) {
            if (!isPersistent) {
                $.global[sessionKey] = storedText;
                return true;
            }
            return settingsStoreWriteTextFile(settingsFile, storedText);
        }

        /**
         * 保存値を読み込み、既定値と突き合わせて返す（型の合わない値・知らない項目は捨てる）
         * @param {Object} defaultSettings - 既定値
         * @returns {Object} 設定（毎回新しいオブジェクト）
         */
        function load(defaultSettings) {
            var savedSettings = null;
            try {
                var storedText = readStoredText();
                if (storedText !== null) {
                    savedSettings = settingsStoreParse(storedText);
                } else if (legacyReader) {
                    savedSettings = legacyReader();
                }
            } catch (e) {
                $.writeln("SettingsStore.load(" + storeName + "): " + e);
                savedSettings = null;
            }
            return settingsStoreMerge(defaultSettings, savedSettings);
        }

        /**
         * 設定を保存する
         * @param {Object} settingValues - 保存する値
         * @returns {boolean} 保存できたら true
         */
        function save(settingValues) {
            try {
                return writeStoredText(settingsStoreSerialize(settingValues, "", 0));
            } catch (e) {
                $.writeln("SettingsStore.save(" + storeName + "): " + e);
                return false;
            }
        }

        /**
         * 保存を消す。旧形式を読み継ぐストアでは空の保存を書き、旧設定が戻らないようにする
         * @returns {boolean} 消せたら true
         */
        function clear() {
            if (legacyReader) return writeStoredText("{}");
            if (!isPersistent) {
                try { delete $.global[sessionKey]; } catch (e) { $.global[sessionKey] = undefined; }
                return true;
            }
            try {
                return settingsFile.exists ? settingsFile.remove() : true;
            } catch (e) {
                $.writeln("SettingsStore.clear(" + storeName + "): " + e);
                return false;
            }
        }

        return { load: load, save: save, clear: clear };
    }

    /**
     * 旧形式の設定ファイルを読む（key=value の行 / toSource / JSON を自動判別。eval は使わない）
     * @param {File|string} legacyFileOrPath - 旧ファイルかそのパス
     * @returns {Object|null} 読み込んだ値（key=value は値がすべて文字列）。無い・読めないときは null
     */
    function readSettingsLegacyFile(legacyFileOrPath) {
        try {
            var legacyFile = (legacyFileOrPath instanceof File) ? legacyFileOrPath : new File(legacyFileOrPath);
            var legacyText = settingsStoreReadTextFile(legacyFile);
            return (legacyText === null) ? null : settingsStoreParseLegacyText(legacyText);
        } catch (e) {
            $.writeln("readSettingsLegacyFile: " + e);
            return null;
        }
    }

    /**
     * app.preferences に文字列で保存していた旧設定を読む（形式は readSettingsLegacyFile と同じく自動判別）
     * @param {string} preferenceKey - 環境設定のキー
     * @returns {Object|null} 読み込んだ値。無い・読めないときは null
     */
    function readSettingsLegacyPreference(preferenceKey) {
        try {
            var legacyText = app.preferences.getStringPreference(preferenceKey);
            if (!legacyText) return null;
            return settingsStoreParseLegacyText(String(legacyText));
        } catch (e) {
            $.writeln("readSettingsLegacyPreference: " + e);
            return null;
        }
    }

    /**
     * テキストファイルを UTF-8 で読む
     * @param {File} textFile - 読むファイル
     * @returns {string|null} 中身。ファイルが無ければ null
     */
    function settingsStoreReadTextFile(textFile) {
        if (!textFile.exists) return null;
        textFile.encoding = "UTF-8";
        if (!textFile.open("r")) throw new Error("cannot open " + textFile.fsName);
        try {
            return textFile.read().replace(/^﻿/, "");
        } finally {
            textFile.close();
        }
    }

    /**
     * テキストファイルを UTF-8 で書く（フォルダーが無ければ作る）
     * @param {File} textFile - 書くファイル
     * @param {string} fileText - 中身
     * @returns {boolean} 書けたら true
     */
    function settingsStoreWriteTextFile(textFile, fileText) {
        try {
            var parentFolder = textFile.parent;
            if (!parentFolder.exists && !parentFolder.create()) throw new Error("cannot create " + parentFolder.fsName);
            textFile.encoding = "UTF-8";
            textFile.lineFeed = "Unix";
            if (!textFile.open("w")) throw new Error("cannot open " + textFile.fsName);
            try {
                textFile.write(fileText);
            } finally {
                textFile.close();
            }
            return true;
        } catch (e) {
            $.writeln("SettingsStore write: " + e);
            return false;
        }
    }

    /**
     * 値が配列か
     * @param {*} checkedValue - 調べる値
     * @returns {boolean} 配列なら true
     */
    function settingsStoreIsArray(checkedValue) {
        return Object.prototype.toString.call(checkedValue) === "[object Array]";
    }

    /**
     * 値が素のオブジェクト（{ } で作ったもの）か
     * @param {*} checkedValue - 調べる値
     * @returns {boolean} 素のオブジェクトなら true
     */
    function settingsStoreIsPlainObject(checkedValue) {
        return checkedValue !== null && typeof checkedValue === "object"
            && Object.prototype.toString.call(checkedValue) === "[object Object]"
            && checkedValue.constructor === Object;
    }

    /**
     * 文字列を JSON の文字列リテラルにする（ASCII 以外は \uXXXX にして、文字コードの取り違えに強くする）
     * @param {string} sourceText - 文字列
     * @returns {string} 引用符つきの文字列
     */
    function settingsStoreQuote(sourceText) {
        var quotedText = "\"";
        for (var i = 0; i < sourceText.length; i++) {
            var charCode = sourceText.charCodeAt(i);
            var oneChar = sourceText.charAt(i);
            if (oneChar === "\"" || oneChar === "\\") quotedText += "\\" + oneChar;
            else if (oneChar === "\n") quotedText += "\\n";
            else if (oneChar === "\r") quotedText += "\\r";
            else if (oneChar === "\t") quotedText += "\\t";
            else if (charCode < 0x20 || charCode > 0x7E) quotedText += "\\u" + ("0000" + charCode.toString(16)).slice(-4);
            else quotedText += oneChar;
        }
        return quotedText + "\"";
    }

    /**
     * 値を JSON の文字列にする（オブジェクトは1項目1行、中身が値だけの配列は1行）。
     * undefined・関数・DOM オブジェクトは項目ごと省き、配列の中では null にする。有限でない数値は null
     * @param {*} sourceValue - 値
     * @param {string} indentText - 今の字下げ
     * @param {number} depth - 入れ子の深さ
     * @returns {string|undefined} JSON の文字列。書けない値は undefined
     */
    function settingsStoreSerialize(sourceValue, indentText, depth) {
        if (depth > SETTINGS_STORE_MAX_DEPTH) throw new Error("settings are nested too deeply");
        if (sourceValue === null) return "null";
        var valueType = typeof sourceValue;
        if (valueType === "boolean") return sourceValue ? "true" : "false";
        if (valueType === "number") return isFinite(sourceValue) ? String(sourceValue) : "null";
        if (valueType === "string") return settingsStoreQuote(sourceValue);
        var innerIndent = indentText + "  ";
        var itemTexts = [];
        var i;
        if (settingsStoreIsArray(sourceValue)) {
            var hasNested = false;
            for (i = 0; i < sourceValue.length; i++) {
                var itemText = settingsStoreSerialize(sourceValue[i], innerIndent, depth + 1);
                itemTexts.push(itemText === undefined ? "null" : itemText);
                if (sourceValue[i] !== null && typeof sourceValue[i] === "object") hasNested = true;
            }
            if (!itemTexts.length) return "[]";
            if (!hasNested) return "[" + itemTexts.join(", ") + "]";
            return "[\n" + innerIndent + itemTexts.join(",\n" + innerIndent) + "\n" + indentText + "]";
        }
        if (settingsStoreIsPlainObject(sourceValue)) {
            for (var key in sourceValue) {
                if (!sourceValue.hasOwnProperty(key)) continue;
                var memberText = settingsStoreSerialize(sourceValue[key], innerIndent, depth + 1);
                if (memberText !== undefined) itemTexts.push(settingsStoreQuote(key) + ": " + memberText);
            }
            if (!itemTexts.length) return "{}";
            return "{\n" + innerIndent + itemTexts.join(",\n" + innerIndent) + "\n" + indentText + "}";
        }
        return undefined; /* 関数・DOM オブジェクトなど / functions, DOM objects, etc. */
    }

    /**
     * JSON（と toSource の出力）を読む。eval は使わない。
     * キーの引用符なし・'…' の文字列・全体の ( ) ・末尾のカンマ・(void 0) も受け付ける
     * @param {string} sourceText - 読む文字列
     * @returns {*} 読み込んだ値
     */
    function settingsStoreParse(sourceText) {
        var readPos = 0;
        var textLength = sourceText.length;

        /**
         * 読み取り位置で失敗を知らせる
         * @param {string} reasonText - 理由
         * @returns {void}
         */
        function fail(reasonText) {
            throw new Error("settings parse error at " + readPos + ": " + reasonText);
        }

        /**
         * 空白を読み飛ばす
         * @returns {void}
         */
        function skipSpaces() {
            while (readPos < textLength && /\s/.test(sourceText.charAt(readPos))) readPos++;
        }

        /**
         * 識別子（英数字・_・$）を読む
         * @returns {string} 識別子。無ければ空文字
         */
        function readWord() {
            var startPos = readPos;
            while (readPos < textLength && /[\w$]/.test(sourceText.charAt(readPos))) readPos++;
            return sourceText.substring(startPos, readPos);
        }

        /**
         * 引用符で囲んだ文字列を読む（" と ' のどちらでも）
         * @returns {string} 文字列
         */
        function readString() {
            var quoteChar = sourceText.charAt(readPos++);
            var resultText = "";
            while (readPos < textLength) {
                var oneChar = sourceText.charAt(readPos++);
                if (oneChar === quoteChar) return resultText;
                if (oneChar !== "\\") { resultText += oneChar; continue; }
                var escapeChar = sourceText.charAt(readPos++);
                if (escapeChar === "n") resultText += "\n";
                else if (escapeChar === "r") resultText += "\r";
                else if (escapeChar === "t") resultText += "\t";
                else if (escapeChar === "b") resultText += "\b";
                else if (escapeChar === "f") resultText += "\f";
                else if (escapeChar === "v") resultText += "\v";
                else if (escapeChar === "0") resultText += "\0";
                else if (escapeChar === "u" || escapeChar === "x") {
                    var hexLength = (escapeChar === "u") ? 4 : 2;
                    var hexText = sourceText.substr(readPos, hexLength);
                    if (!new RegExp("^[0-9A-Fa-f]{" + hexLength + "}$").test(hexText)) fail("bad escape");
                    resultText += String.fromCharCode(parseInt(hexText, 16));
                    readPos += hexLength;
                } else resultText += escapeChar;
            }
            fail("unterminated string");
        }

        /**
         * 値を1つ読む
         * @param {number} depth - 入れ子の深さ
         * @returns {*} 値
         */
        function readValue(depth) {
            if (depth > SETTINGS_STORE_MAX_DEPTH) fail("nested too deeply");
            skipSpaces();
            var oneChar = sourceText.charAt(readPos);
            if (oneChar === "{") return readObject(depth);
            if (oneChar === "[") return readArray(depth);
            if (oneChar === "\"" || oneChar === "'") return readString();
            if (oneChar === "(") {
                readPos++;
                var innerValue = readValue(depth + 1);
                skipSpaces();
                if (sourceText.charAt(readPos) !== ")") fail("expected )");
                readPos++;
                return innerValue;
            }
            var numberMatch = /^-?(\d+\.?\d*|\.\d+)([eE][+\-]?\d+)?/.exec(sourceText.substring(readPos, readPos + 64));
            if (numberMatch) {
                readPos += numberMatch[0].length;
                return Number(numberMatch[0]);
            }
            var wordText = readWord();
            if (wordText === "true") return true;
            if (wordText === "false") return false;
            if (wordText === "null") return null;
            if (wordText === "NaN") return NaN;
            if (wordText === "Infinity") return Infinity;
            if (wordText === "void") { readValue(depth + 1); return undefined; } /* toSource の (void 0) */
            fail("unexpected " + (wordText || oneChar || "end of text"));
        }

        /**
         * 配列を読む
         * @param {number} depth - 入れ子の深さ
         * @returns {Array} 配列
         */
        function readArray(depth) {
            var resultArray = [];
            readPos++;
            skipSpaces();
            while (sourceText.charAt(readPos) !== "]") {
                resultArray.push(readValue(depth + 1));
                skipSpaces();
                if (sourceText.charAt(readPos) === ",") { readPos++; skipSpaces(); continue; }
                if (sourceText.charAt(readPos) !== "]") fail("expected , or ]");
            }
            readPos++;
            return resultArray;
        }

        /**
         * オブジェクトを読む（__proto__ のキーは捨てる）
         * @param {number} depth - 入れ子の深さ
         * @returns {Object} オブジェクト
         */
        function readObject(depth) {
            var resultObject = {};
            readPos++;
            skipSpaces();
            while (sourceText.charAt(readPos) !== "}") {
                var keyChar = sourceText.charAt(readPos);
                var memberKey = (keyChar === "\"" || keyChar === "'") ? readString() : readWord();
                if (memberKey === "") fail("expected a key");
                skipSpaces();
                if (sourceText.charAt(readPos) !== ":") fail("expected :");
                readPos++;
                var memberValue = readValue(depth + 1);
                if (memberKey !== "__proto__") resultObject[memberKey] = memberValue;
                skipSpaces();
                if (sourceText.charAt(readPos) === ",") { readPos++; skipSpaces(); continue; }
                if (sourceText.charAt(readPos) !== "}") fail("expected , or }");
            }
            readPos++;
            return resultObject;
        }

        var parsedValue = readValue(0);
        skipSpaces();
        if (readPos < textLength) fail("unexpected text after the value");
        return parsedValue;
    }

    /**
     * 旧形式の文字列を読む。{ [ ( で始まれば JSON / toSource、それ以外は key=value の行とみなす
     * @param {string} legacyText - 旧形式の文字列
     * @returns {Object|null} 読み込んだ値
     */
    function settingsStoreParseLegacyText(legacyText) {
        var trimmedText = legacyText.replace(/^﻿/, "").replace(/^\s+|\s+$/g, "");
        if (trimmedText === "") return null;
        if (/^[\{\[\(]/.test(trimmedText)) return settingsStoreParse(trimmedText);
        var keyValues = {};
        var textLines = trimmedText.split(/\r\n|\r|\n/);
        for (var i = 0; i < textLines.length; i++) {
            var separatorIndex = textLines[i].indexOf("=");
            if (separatorIndex < 1) continue;
            var lineKey = textLines[i].substring(0, separatorIndex).replace(/^\s+|\s+$/g, "");
            if (lineKey !== "" && lineKey !== "__proto__") keyValues[lineKey] = textLines[i].substring(separatorIndex + 1);
        }
        return keyValues;
    }

    /**
     * 値を深くコピーする（素のデータだけ。関数・DOM オブジェクトは null）
     * @param {*} sourceValue - コピー元
     * @returns {*} コピー
     */
    function settingsStoreClone(sourceValue) {
        if (sourceValue === null || typeof sourceValue !== "object") {
            return (typeof sourceValue === "function" || sourceValue === undefined) ? null : sourceValue;
        }
        var i;
        if (settingsStoreIsArray(sourceValue)) {
            var arrayCopy = [];
            for (i = 0; i < sourceValue.length; i++) arrayCopy.push(settingsStoreClone(sourceValue[i]));
            return arrayCopy;
        }
        if (!settingsStoreIsPlainObject(sourceValue)) return null;
        var objectCopy = {};
        for (var key in sourceValue) {
            if (sourceValue.hasOwnProperty(key)) objectCopy[key] = settingsStoreClone(sourceValue[key]);
        }
        return objectCopy;
    }

    /**
     * 保存値を既定値と突き合わせる。型は既定値に合わせ、合わなければ既定値を使う。
     * 既定値が {} か null なら中身を問わず受け取り、配列は配列なら受け取る。既定値に無い項目は捨てる
     * @param {*} defaultValue - 既定値
     * @param {*} savedValue - 保存値
     * @returns {*} 突き合わせた値（新しいオブジェクト）
     */
    function settingsStoreMerge(defaultValue, savedValue) {
        if (defaultValue === null || defaultValue === undefined) {
            return (savedValue === undefined) ? null : settingsStoreClone(savedValue);
        }
        var defaultType = typeof defaultValue;
        var savedType = typeof savedValue;
        if (defaultType === "boolean") {
            if (savedType === "boolean") return savedValue;
            if (savedValue === 1 || savedValue === "1" || savedValue === "true") return true;
            if (savedValue === 0 || savedValue === "0" || savedValue === "false") return false;
            return defaultValue;
        }
        if (defaultType === "number") {
            if (savedType === "number" && isFinite(savedValue)) return savedValue;
            if (savedType === "string" && /\S/.test(savedValue)) {
                var parsedNumber = Number(savedValue);
                if (isFinite(parsedNumber)) return parsedNumber;
            }
            return defaultValue;
        }
        if (defaultType === "string") {
            if (savedType === "string") return savedValue;
            if (savedType === "number" && isFinite(savedValue)) return String(savedValue);
            if (savedType === "boolean") return String(savedValue);
            return defaultValue;
        }
        if (settingsStoreIsArray(defaultValue)) {
            return settingsStoreClone(settingsStoreIsArray(savedValue) ? savedValue : defaultValue);
        }
        if (defaultType === "object") {
            var savedIsObject = settingsStoreIsPlainObject(savedValue);
            var hasDefaultKeys = false;
            var mergedObject = {};
            for (var key in defaultValue) {
                if (!defaultValue.hasOwnProperty(key)) continue;
                hasDefaultKeys = true;
                mergedObject[key] = settingsStoreMerge(defaultValue[key], savedIsObject ? savedValue[key] : undefined);
            }
            /* 既定値が {} なら自由な入れ物として中身ごと受け取る / an empty default {} is a free-form map */
            if (!hasDefaultKeys && savedIsObject) return settingsStoreClone(savedValue);
            return mergedObject;
        }
        return defaultValue;
    }

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // 設定の保存（再利用パーツ）ここまで / End of the reusable settings store
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    // =========================================
    // プリセット / Presets
    // =========================================

    /* ユーザーの設定フォルダーに保存する / Stored in the user settings folder */
    var PRESET_FILE = new File(Folder.userData + "/" + SCRIPT_NAME + "/presets.json");

    /* 直前の設定をIllustratorのセッション中だけ記憶する（終了でリセット） */
    /* Remember the last settings within this Illustrator session (resets on quit) */
    /* 旧キー（$.global.AiConnectorBuilder_lastSettings）が残っていれば1度だけ読み継ぐ / read the old key once if it is still there */
    var sessionSettingsStore = createSettingsStore(SCRIPT_NAME, "session", {
        legacy: function () {
            return $.global[SCRIPT_NAME + "_lastSettings"] || null;
        }
    });

    /* 既定値は {}（getPresetEntries() のキーをそのまま入れる自由な入れ物。空なら何も戻さない）
       The default is {}: a free-form map keyed by getPresetEntries(); empty means nothing to restore */
    var DEFAULT_SESSION_SETTINGS = {};

    /**
     * 保存済みのプリセットを読み込む
     * @returns {Array<object>} プリセットの配列。読み込めないときは空配列
     */
    function loadPresets() {
        if (!PRESET_FILE.exists) return [];
        try {
            PRESET_FILE.encoding = "UTF-8";
            if (!PRESET_FILE.open("r")) return [];
            var savedText = PRESET_FILE.read();
            PRESET_FILE.close();
            var parsedPresets = eval(savedText); // toSource() で書き出した内容
            return (parsedPresets && typeof parsedPresets.length === "number") ? parsedPresets : [];
        } catch (e) {
            $.writeln(SCRIPT_NAME + ": プリセットの読み込みに失敗 / failed to load presets — " + e);
            return [];
        }
    }

    /**
     * プリセットを保存する
     * @param {Array<object>} presetList - 保存するプリセットの配列
     * @returns {boolean} 保存できたか
     */
    function writePresets(presetList) {
        try {
            var presetFolder = PRESET_FILE.parent;
            if (!presetFolder.exists) presetFolder.create();
            PRESET_FILE.encoding = "UTF-8";
            if (!PRESET_FILE.open("w")) return false;
            PRESET_FILE.write(presetList.toSource());
            PRESET_FILE.close();
            return true;
        } catch (e) {
            $.writeln(SCRIPT_NAME + ": プリセットの保存に失敗 / failed to save presets — " + e);
            return false;
        }
    }

    /**
     * 名前でプリセットのインデックスを探す
     * @param {Array<object>} presetList - プリセットの配列
     * @param {string} name - プリセット名
     * @returns {number} インデックス。無ければ -1
     */
    function findPresetIndex(presetList, name) {
        for (var i = 0; i < presetList.length; i++) {
            if (presetList[i].name === name) return i;
        }
        return -1;
    }

    var savedPresets = loadPresets();

    // =========================================
    // メインダイアログ / Main dialog
    // =========================================

    initThemeColors();

    var connectorDialog = new Window("dialog", getLabel(LABELS.dialog.title) + " " + SCRIPT_VERSION);
    setupWindow(connectorDialog);

    /* プリセット / Presets */
    var presetRow = addFieldRow(connectorDialog, LABELS.fieldLabel.preset);
    var presetDropdown = presetRow.add("dropdownlist", undefined, []);
    presetDropdown.preferredSize.width = PRESET_LIST_WIDTH;
    presetDropdown.helpTip = getLabel(LABELS.tooltip.preset);

    /* スペーサー（伸縮）：保存・削除ボタンを右端に寄せる / Spacer so the buttons stay right */
    var presetSpacer = presetRow.add("group");
    presetSpacer.alignment = ["fill", "fill"];
    presetSpacer.minimumSize.width = 0;
    var btnSavePreset = presetRow.add("button", undefined, getLabel(LABELS.button.savePreset));
    btnSavePreset.preferredSize.width = PRESET_BUTTON_WIDTH;
    btnSavePreset.alignment = ["right", "center"];
    btnSavePreset.helpTip = getLabel(LABELS.tooltip.savePreset);
    var btnRemovePreset = presetRow.add("button", undefined, getLabel(LABELS.button.removePreset));
    btnRemovePreset.preferredSize.width = PRESET_BUTTON_WIDTH;
    btnRemovePreset.alignment = ["right", "center"];

    /* コネクター / Connector */
    var connectorPanel = addPanel(connectorDialog, getLabel(LABELS.panel.connector));

    var lineShapeField = addRadioRow(connectorPanel, LABELS.fieldLabel.lineShape, [LABELS.radio.shapeStraight, LABELS.radio.shapeWarp, LABELS.radio.shapeElbow, LABELS.radio.shapeBranch, LABELS.radio.shapeCurve], DEFAULT_LINE_SHAPE, LABELS.tooltip.lineShape);

    var warpTypeRow = addFieldRow(connectorPanel, LABELS.fieldLabel.warpType);
    var warpTypeList = warpTypeRow.add("dropdownlist", undefined, getWarpTypeLabels());
    warpTypeList.preferredSize.width = LIST_WIDTH;
    warpTypeList.selection = DEFAULT_WARP_TYPE;

    var warpAmountField = addNumberRow(connectorPanel, LABELS.fieldLabel.warpAmount, DEFAULT_WARP_AMOUNT, LABELS.unit.percent, LABELS.tooltip.warpAmount, [WARP_AMOUNT_MIN, WARP_AMOUNT_MAX]);
    var warpAxisField = addRadioRow(connectorPanel, LABELS.fieldLabel.warpAxis, [LABELS.radio.axisAuto, LABELS.radio.axisHorizontal, LABELS.radio.axisVertical], DEFAULT_WARP_AXIS, LABELS.tooltip.warpAxis);
    var cornerField = addNumberRow(connectorPanel, LABELS.fieldLabel.cornerRadius, DEFAULT_CORNER_RADIUS, LABELS.unit.pt, LABELS.tooltip.cornerRadius);
    var startPointField = addRadioRow(connectorPanel, LABELS.fieldLabel.startPoint, [LABELS.radio.startCenter, LABELS.radio.startDivided, LABELS.radio.startKeyCenter], DEFAULT_START_POINT, LABELS.tooltip.startPoint, true);

    /* 開始ポイントの右に9分割のウィジェットを置く / 3x3 widget at the right of the row */
    var anchorSpacer = startPointField.row.add("group");
    anchorSpacer.alignment = ["fill", "fill"];
    anchorSpacer.minimumSize.width = 0;
    var unifyAnchorWidget = addAnchorWidget(startPointField.row, DEFAULT_UNIFY_ANCHOR, function () {
        updatePreview();
    });
    unifyAnchorWidget.alignment = ["right", "center"];
    unifyAnchorWidget.helpTip = getLabel(LABELS.tooltip.unifyAnchor);
    unifyAnchorWidget.enabled = false;

    var unifyStartRow = connectorPanel.add("group");
    unifyStartRow.orientation = "row";
    unifyStartRow.alignment = ["fill", "top"];
    unifyStartRow.alignChildren = ["left", "center"];
    /* 行ラベルぶん字下げしてチェックボックスを他の欄にそろえる / Indent by the label width */
    var unifyStartIndent = unifyStartRow.add("statictext", undefined, "");
    unifyStartIndent.preferredSize.width = LABEL_WIDTH;
    var unifyStartCheck = unifyStartRow.add("checkbox", undefined, getLabel(LABELS.checkbox.unifyStart));
    unifyStartCheck.helpTip = getLabel(LABELS.tooltip.unifyStart);

    /* 線・矢印は2カラム。ラベル幅はコネクターパネルより狭くし、矢印はさらに狭くする */
    setLabelWidth(COLUMN_LABEL_WIDTH);

    var lineArrowColumnsGroup = connectorDialog.add("group");
    lineArrowColumnsGroup.orientation = "row";
    lineArrowColumnsGroup.alignChildren = ["fill", "fill"];
    lineArrowColumnsGroup.alignment = ["fill", "top"];
    lineArrowColumnsGroup.spacing = COLUMN_SPACING;

    /* 線 / Line */
    var linePanel = addPanel(lineArrowColumnsGroup, getLabel(LABELS.panel.line));
    var strokeWidthField = addNumberRow(linePanel, LABELS.fieldLabel.strokeWidth, DEFAULT_STROKE_WIDTH, LABELS.unit.pt);
    var strokeJoinField = addRadioRow(linePanel, LABELS.fieldLabel.strokeJoin, [STROKE_JOIN_OPTIONS[0].label, STROKE_JOIN_OPTIONS[1].label, STROKE_JOIN_OPTIONS[2].label], DEFAULT_STROKE_JOIN, LABELS.tooltip.strokeJoin, true);
    var dashStyleField = addRadioRow(linePanel, LABELS.fieldLabel.dashStyle, [LABELS.radio.dashNone, LABELS.radio.dashDashed, LABELS.radio.dashDotted], DEFAULT_DASH_STYLE, null, true);
    var dashSegmentsField = addNumberRow(linePanel, LABELS.fieldLabel.dashSegments, DEFAULT_DASH_SEGMENTS, LABELS.unit.blank, LABELS.tooltip.dashSegments, null, { integer: true });
    var dashGapField = addNumberRow(linePanel, LABELS.fieldLabel.dashGap, DEFAULT_DASH_GAP, LABELS.unit.pt, LABELS.tooltip.dashGap);

    /* 矢印 / Arrowheads */
    setLabelWidth(ARROW_LABEL_WIDTH[uiLang]);
    var arrowPanel = addPanel(lineArrowColumnsGroup, getLabel(LABELS.panel.arrow));
    var arrowShapeField = addArrowShapeRow(arrowPanel, DEFAULT_ARROW_INDEX, function () {
        // 矢印ごとに見え方が違うので、選び直したらその矢印の既定値（倍率・位置・線端）を入れる
        var selectedArrow = ARROW_CHOICES[getSelectedIconIndex(arrowShapeField.buttons)];
        if (selectedArrow.number !== 0) {
            arrowScaleField.input.text = selectedArrow.scale;
            selectRadio(arrowTipField.radios, selectedArrow.tip);
        }
        // 丸は線端を丸形に、矢印はなしにそろえる
        if (selectedArrow.cap !== undefined) selectRadio(strokeCapField.radios, selectedArrow.cap);
        updatePreview();
    });
    // 初期値も選択中の矢印の既定倍率にそろえる
    var arrowScaleField = addNumberRow(arrowPanel, LABELS.fieldLabel.arrowScale, ARROW_CHOICES[DEFAULT_ARROW_INDEX].scale, LABELS.unit.percent, LABELS.tooltip.arrowScale);
    var arrowPositionField = addRadioRow(arrowPanel, LABELS.fieldLabel.arrowPosition, [LABELS.radio.arrowEnd, LABELS.radio.arrowBoth], DEFAULT_ARROW_POSITION, LABELS.tooltip.arrowPosition, true);
    var arrowTipField = addRadioRow(arrowPanel, LABELS.fieldLabel.arrowTip, [ARROW_TIP_OPTIONS[0].label, ARROW_TIP_OPTIONS[1].label], ARROW_CHOICES[DEFAULT_ARROW_INDEX].tip, LABELS.tooltip.arrowTip, true);
    var endGapField = addNumberRow(arrowPanel, LABELS.fieldLabel.endGap, DEFAULT_END_GAP, LABELS.unit.pt, LABELS.tooltip.endGap);
    var strokeCapField = addRadioRow(arrowPanel, LABELS.fieldLabel.strokeCap, [STROKE_CAP_OPTIONS[0].label, STROKE_CAP_OPTIONS[1].label, STROKE_CAP_OPTIONS[2].label], DEFAULT_STROKE_CAP, LABELS.tooltip.strokeCap, true);

    /**
     * ダイアログの入力値を設定として取り出す
     * @returns {object} コネクターの設定
     */
    function getSettings() {
        var warpType = WARP_TYPE_CHOICES[warpTypeList.selection ? warpTypeList.selection.index : DEFAULT_WARP_TYPE];
        var strokeWidth = parseNumberInput(strokeWidthField.input.text, DEFAULT_STROKE_WIDTH);
        if (strokeWidth <= 0) strokeWidth = DEFAULT_STROKE_WIDTH;

        // 矢印はパスの始点＝キーオブジェクト側、終点＝相手の図形側
        var noneName = getLabel(LABELS.actionName.arrowNone);
        var arrowChoice = ARROW_CHOICES[getSelectedIconIndex(arrowShapeField.buttons)];
        var arrowName = (arrowChoice.number === 0) ? noneName : getArrowName(arrowChoice.number);
        var arrowPosition = getSelectedRadioIndex(arrowPositionField.radios);
        var lineShape = getSelectedRadioIndex(lineShapeField.radios);

        return {
            strokeWidth: strokeWidth,
            strokeJoin: getSelectedRadioIndex(strokeJoinField.radios),
            strokeCap: getSelectedRadioIndex(strokeCapField.radios),
            dashStyle: getSelectedRadioIndex(dashStyleField.radios),
            dashSegments: Math.round(parseNumberInput(dashSegmentsField.input.text, DEFAULT_DASH_SEGMENTS)),
            dashGap: parseNumberInput(dashGapField.input.text, DEFAULT_DASH_GAP),
            startPoint: getSelectedRadioIndex(startPointField.radios),
            unifyStart: unifyStartCheck.value,
            unifyAnchor: getAnchorWidgetIndex(unifyAnchorWidget),
            lineShape: lineShape,
            warpName: warpType.name,
            warpStyle: warpType.style,
            warpAmount: parseNumberInput(warpAmountField.input.text, DEFAULT_WARP_AMOUNT),
            warpAxis: getSelectedRadioIndex(warpAxisField.radios),
            // 角丸はカギ・分岐のときだけ / Elbow and Branch only
            cornerRadius: (lineShape === 2 || lineShape === 3) ? parseNumberInput(cornerField.input.text, DEFAULT_CORNER_RADIUS) : 0,
            hasArrow: (arrowChoice.number !== 0),
            arrowInnerDot: (arrowChoice.innerDot === true),
            startArrow: (arrowPosition === 1) ? arrowName : noneName,
            endArrow: arrowName,
            arrowScale: parseNumberInput(arrowScaleField.input.text, DEFAULT_ARROW_SCALE),
            endGap: parseNumberInput(endGapField.input.text, DEFAULT_END_GAP),
            arrowTip: ARROW_TIP_OPTIONS[getSelectedRadioIndex(arrowTipField.radios)],
            arrowPosition: arrowPosition
        };
    }

    /* プリセット適用中は「（カスタム）」へ戻さない */
    var isApplyingPreset = false;

    /**
     * 設定に合わせて、使わない項目をディム表示にする
     * @param {object} settings - ダイアログの設定
     * @returns {void}
     */
    function updateControlStates(settings) {
        warpTypeRow.enabled = (settings.lineShape === 1);
        // カーブのふくらみもこの欄で決める / Curve reuses this field for its bow
        setFieldRowEnabled(warpAmountField.row, (settings.lineShape === 1 || settings.lineShape === 4));
        warpAxisField.row.enabled = (settings.lineShape === 1);
        setFieldRowEnabled(cornerField.row, (settings.lineShape === 2 || settings.lineShape === 3));
        if (unifyAnchorWidget.enabled !== settings.unifyStart) {
            setAnchorWidgetEnabled(unifyAnchorWidget, settings.unifyStart);
        }
        setFieldRowEnabled(dashSegmentsField.row, (settings.dashStyle === 1));
        setFieldRowEnabled(dashGapField.row, (settings.dashStyle !== 0));
        setFieldRowEnabled(arrowScaleField.row, settings.hasArrow);
        arrowPositionField.row.enabled = settings.hasArrow;
        arrowTipField.row.enabled = settings.hasArrow;
    }

    /**
     * 現在の設定でプレビューを更新する
     * @returns {void}
     */
    function updatePreview() {
        // 手で変えたらプリセットの選択を外す
        if (!isApplyingPreset && presetDropdown.selection && presetDropdown.selection.index !== 0) {
            presetDropdown.selection = 0;
        }
        var settings = getSettings();
        updateControlStates(settings);
        buildConnectors(settings);
        app.redraw();

        // 閉じた後はコントロールを読めないので、更新のたびに控えておく
        sessionSettingsStore.save(getPresetFromDialog(""));
    }

    /**
     * ラジオボタンの選択を切り替える
     * @param {Array<object>} radios - ラジオボタンの配列
     * @param {number} index - 選択するインデックス
     * @returns {void}
     */
    function selectRadio(radios, index) {
        if (!(index >= 0) || index >= radios.length) return;
        for (var i = 0; i < radios.length; i++) {
            radios[i].value = (i === index);
        }
    }

    /**
     * プリセットに保存する項目の対応表を返す
     * field=数値入力の行、radios=ラジオの配列、list=ドロップダウン
     * @returns {Array<object>} 保存項目の配列
     */
    function getPresetEntries() {
        return [
            { key: "strokeWidth",   field: strokeWidthField },
            { key: "strokeJoin",    radios: strokeJoinField.radios },
            { key: "strokeCap",     radios: strokeCapField.radios },
            { key: "endGap",        field: endGapField },
            { key: "dashStyle",     radios: dashStyleField.radios },
            { key: "dashSegments",  field: dashSegmentsField },
            { key: "dashGap",       field: dashGapField },
            { key: "lineShape",     radios: lineShapeField.radios },
            { key: "warpType",      list: warpTypeList },
            { key: "warpAmount",    field: warpAmountField },
            { key: "warpAxis",      radios: warpAxisField.radios },
            { key: "cornerRadius",  field: cornerField },
            { key: "startPoint",    radios: startPointField.radios },
            { key: "unifyStart",    check: unifyStartCheck },
            { key: "unifyAnchor",   widget: unifyAnchorWidget },
            { key: "arrowIndex",    icons: arrowShapeField.buttons },
            { key: "arrowScale",    field: arrowScaleField },
            { key: "arrowPosition", radios: arrowPositionField.radios },
            { key: "arrowTip",      radios: arrowTipField.radios }
        ];
    }

    /**
     * 現在のダイアログの状態をプリセットとして取り出す
     * @param {string} name - プリセット名
     * @returns {object} プリセット
     */
    function getPresetFromDialog(name) {
        var entries = getPresetEntries();
        var preset = { name: name };
        for (var i = 0; i < entries.length; i++) {
            var entry = entries[i];
            if (entry.field) {
                preset[entry.key] = entry.field.input.text;
            } else if (entry.radios) {
                preset[entry.key] = getSelectedRadioIndex(entry.radios);
            } else if (entry.icons) {
                preset[entry.key] = getSelectedIconIndex(entry.icons);
            } else if (entry.check) {
                preset[entry.key] = entry.check.value;
            } else if (entry.widget) {
                preset[entry.key] = getAnchorWidgetIndex(entry.widget);
            } else {
                preset[entry.key] = entry.list.selection ? entry.list.selection.index : 0;
            }
        }
        return preset;
    }

    /**
     * プリセットをダイアログへ反映する
     * @param {object} preset - 反映するプリセット
     * @returns {void}
     */
    function applyPresetToDialog(preset) {
        if (!preset) return;

        var entries = getPresetEntries();
        for (var i = 0; i < entries.length; i++) {
            var entry = entries[i];
            var value = preset[entry.key];
            if (value === undefined || value === null) continue;

            if (entry.field) {
                entry.field.input.text = value;
                syncSliderToInput(entry.field);
            } else if (entry.radios) {
                selectRadio(entry.radios, value);
            } else if (entry.icons) {
                selectIconButton(entry.icons, value);
            } else if (entry.check) {
                entry.check.value = (value === true || value === "true");
            } else if (entry.widget) {
                setAnchorWidgetValue(entry.widget, value);
            } else if (value >= 0 && value < entry.list.items.length) {
                entry.list.selection = value;
            }
        }
    }

    /**
     * プリセットのドロップダウンを作り直す
     * @param {string} selectName - 選択状態にするプリセット名（null のときは（カスタム））
     * @returns {void}
     */
    function refreshPresetDropdown(selectName) {
        isApplyingPreset = true;
        presetDropdown.removeAll();
        presetDropdown.add("item", getLabel(LABELS.dropdown.customPreset));
        for (var i = 0; i < savedPresets.length; i++) {
            presetDropdown.add("item", savedPresets[i].name);
        }
        var itemIndex = selectName ? (findPresetIndex(savedPresets, selectName) + 1) : 0;
        presetDropdown.selection = (itemIndex > 0) ? itemIndex : 0;
        isApplyingPreset = false;
    }

    /**
     * ラジオボタンの配列にプレビュー更新を割り当てる
     * @param {Array<object>} radios - ラジオボタンの配列
     * @returns {void}
     */
    function bindPreviewToRadios(radios) {
        for (var i = 0; i < radios.length; i++) {
            radios[i].onClick = updatePreview;
        }
    }

    /**
     * 数値フィールドにプレビュー更新・∧∨と↑↓キーの増減・スライダー連動を割り当てる
     * @param {object} field - addNumberRow が返したフィールド
     * @param {boolean} allowNegative - マイナス値を許可するか
     * @returns {void}
     */
    function bindPreviewToField(field, allowNegative) {
        field.input.onChange = function () {
            syncSliderToInput(field);
            updatePreview();
        };
        /* ∧∨・↑↓キーの増減量の条件と、増減後の反映 / stepping rules and what happens after a step */
        if (!allowNegative) field.stepOptions.min = 0;
        field.stepOptions.onStep = function () { field.input.onChange(); };

        if (!field.slider) return;
        // ドラッグ中は数値の表示だけ更新し、離したところで作り直す
        field.slider.onChanging = function () {
            field.input.text = Math.round(field.slider.value);
        };
        field.slider.onChange = function () {
            field.slider.onChanging();
            updatePreview();
        };
    }

    bindPreviewToRadios(lineShapeField.radios);
    bindPreviewToRadios(startPointField.radios);
    unifyStartCheck.onClick = updatePreview;
    bindPreviewToRadios(strokeJoinField.radios);
    for (i = 0; i < dashStyleField.radios.length; i++) {
        dashStyleField.radios[i].onClick = function () {
            // ドットは丸形の線端でないと点が出ないので、選んだときに丸形へそろえる
            if (getSelectedRadioIndex(dashStyleField.radios) === 2) selectRadio(strokeCapField.radios, 1);
            updatePreview();
        };
    }
    bindPreviewToRadios(strokeCapField.radios);
    bindPreviewToRadios(warpAxisField.radios);
    bindPreviewToRadios(arrowPositionField.radios);
    bindPreviewToRadios(arrowTipField.radios);
    bindPreviewToField(strokeWidthField, false);
    bindPreviewToField(dashSegmentsField, false);
    bindPreviewToField(dashGapField, false);
    bindPreviewToField(warpAmountField, true);
    bindPreviewToField(cornerField, false);
    bindPreviewToField(arrowScaleField, false);
    bindPreviewToField(endGapField, false);
    warpTypeList.onChange = updatePreview;

    /* ボタン行（右：キャンセル・OK） / Button row (right: Cancel and OK) */
    var connectorButtonRow = addButtonRow(connectorDialog);
    var btnCancel = connectorButtonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });
    btnCancel.onClick = function () {
        connectorDialog.close(DIALOG_RESULT_CANCEL);
    };
    var btnOK = connectorButtonRow.rightGroup.add("button", undefined, "OK", { name: "ok" });
    btnOK.onClick = function () {
        connectorDialog.close(DIALOG_RESULT_OK);
    };

    presetDropdown.onChange = function () {
        if (isApplyingPreset) return;
        if (!presetDropdown.selection || presetDropdown.selection.index === 0) return;

        isApplyingPreset = true;
        applyPresetToDialog(savedPresets[presetDropdown.selection.index - 1]);
        isApplyingPreset = false;
        updatePreview();
    };

    btnSavePreset.onClick = function () {
        var selectedName = (presetDropdown.selection && presetDropdown.selection.index > 0) ? presetDropdown.selection.text : "";
        var inputName = prompt(getLabel(LABELS.alert.presetName), selectedName);
        if (inputName === null) return;
        var presetName = String(inputName).replace(/^\s+|\s+$/g, "");
        if (presetName === "") return;

        var preset = getPresetFromDialog(presetName);
        var existingIndex = findPresetIndex(savedPresets, presetName);
        if (existingIndex === -1) {
            savedPresets.push(preset);
        } else {
            savedPresets[existingIndex] = preset; // 同名は上書き / overwrite when the name exists
        }
        if (!writePresets(savedPresets)) {
            alert(getLabel(LABELS.alert.presetFailed));
            return;
        }
        refreshPresetDropdown(presetName);
    };

    btnRemovePreset.onClick = function () {
        if (!presetDropdown.selection || presetDropdown.selection.index === 0) return;
        if (!confirm(getLabel(LABELS.alert.presetRemove))) return;

        savedPresets.splice(presetDropdown.selection.index - 1, 1);
        if (!writePresets(savedPresets)) {
            alert(getLabel(LABELS.alert.presetFailed));
            return;
        }
        refreshPresetDropdown(null);
    };

    // 前回このセッションで閉じたときの設定に戻す
    var lastSessionSettings = sessionSettingsStore.load(DEFAULT_SESSION_SETTINGS);
    isApplyingPreset = true;
    applyPresetToDialog(lastSessionSettings);
    isApplyingPreset = false;
    refreshPresetDropdown(null);
    updatePreview();

    prepareDialogWindow(connectorDialog, SCRIPT_NAME);
    if (connectorDialog.show() !== DIALOG_RESULT_OK) {
        // キャンセル／ESC／ウィンドウを閉じる: プレビューを破棄
        removeConnectors();
        if (!connectorLayerState.existed) {
            connectorLayer.remove();
        } else {
            // 作成のために開いたロック・表示状態を元に戻す
            connectorLayer.locked = connectorLayerState.locked;
            connectorLayer.visible = connectorLayerState.visible;
        }
        app.redraw();
        toggleSmartGuides();
        return;
    }

    // コネクターを選択状態にする
    doc.selection = null;
    for (i = 0; i < createdConnectors.length; i++) {
        createdConnectors[i].selected = true;
    }

    toggleSmartGuides();
})();
