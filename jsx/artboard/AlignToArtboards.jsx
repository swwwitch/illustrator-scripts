#target illustrator
#targetengine "AlignToArtboardsEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したオブジェクトを、アートボード上の指定位置へ整列します。
整列先は3×3の9点から選べ、辺からのマージンも指定できます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AlignToArtboards.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n50aacdeb4908

### Overview

Aligns the selected objects to a chosen position on the artboard.
The target is picked from a 3x3 grid of nine points, with a margin from the edges.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AlignToArtboards.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AlignToArtboards";             /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.2.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-12-17";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-28";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AlignToArtboards.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AlignToArtboards.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n50aacdeb4908"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* ダイアログの初期値。必要に応じて編集 / Dialog defaults; edit as needed */
    var DEFAULT_SETTINGS = {
        useActiveArtboardAsReference: true, /* 「アクティブを基準」を初期選択 / Start in active-artboard reference mode */
        anchorCode: "LT",                   /* 初期の整列先 / Initial anchor */
        marginText: "0",                    /* マージン入力欄の初期値 / Initial margin input */
        linkMargins: true,                  /* 左右の値を上下にも反映 / Mirror horizontal margin to vertical */
        useVisibleBounds: true              /* プレビュー境界で整列 / Align by visible bounds */
    };

    // =========================================
    // レイアウト / Layout
    // =========================================

    /* パネル共通レイアウト / Common panel layout */
    var PANEL_MARGINS = [15, 20, 15, 10];
    /* 整列先パネルは左右・上下とも余白を詰める（9軸ウィジェットの余白調整）
       The anchor panel uses tighter padding on all sides to fit the 9-axis widget */
    var ANCHOR_PANEL_MARGINS = [9, 13, 9, 4];
    var MARGIN_FIELD_CHARACTERS = 4;
    var BOUNDS_OPTION_MARGINS = [4, 4, 4, 4];  /* プレビュー境界の行の余白 / margins of the preview-bounds row */

    /* 9軸ウィジェットの寸法（onDrawで描画） / 9-axis widget metrics (drawn via onDraw) */
    var ANCHOR_WIDGET_SIZE = 66;
    var ANCHOR_CELL_SIZE = 9;
    var ANCHOR_CELL_GAP = 7.5;

    // =========================================
    // 整列先の定義 / Anchor definitions
    // =========================================

    /* 9点アンカーの定義。ratioX/ratioY は矩形内の相対位置（0=左/上, 0.5=中央, 1=右/下）
       Nine anchor points; ratioX/ratioY are relative positions inside a rectangle */
    var ANCHOR_DEFINITIONS = [
        { code: "LT", labelKey: "topLeft",      shortcutKey: "q", ratioX: 0,   ratioY: 0 },
        { code: "TC", labelKey: "topCenter",    shortcutKey: "w", ratioX: 0.5, ratioY: 0 },
        { code: "RT", labelKey: "topRight",     shortcutKey: "e", ratioX: 1,   ratioY: 0 },
        { code: "LM", labelKey: "middleLeft",   shortcutKey: "a", ratioX: 0,   ratioY: 0.5 },
        { code: "C",  labelKey: "center",       shortcutKey: "s", ratioX: 0.5, ratioY: 0.5 },
        { code: "RM", labelKey: "middleRight",  shortcutKey: "d", ratioX: 1,   ratioY: 0.5 },
        { code: "LB", labelKey: "bottomLeft",   shortcutKey: "z", ratioX: 0,   ratioY: 1 },
        { code: "BC", labelKey: "bottomCenter", shortcutKey: "x", ratioX: 0.5, ratioY: 1 },
        { code: "RB", labelKey: "bottomRight",  shortcutKey: "c", ratioX: 1,   ratioY: 1 }
    ];

    /* 中央アンカー（マージンを持たない） / Center anchor (no margin) */
    var CENTER_ANCHOR_CODE = "C";

    // =========================================
    // 単位 / Units
    // =========================================

    /* 単位コードに対応する表示ラベルと、1単位あたりのポイント数
       Unit code -> display label and points per unit */
    var UNITS = [
        { label: "in",    pointsPerUnit: 72 },                /* 0 */
        { label: "mm",    pointsPerUnit: 72 / 25.4 },         /* 1 */
        { label: "pt",    pointsPerUnit: 1 },                 /* 2 */
        { label: "pica",  pointsPerUnit: 12 },                /* 3 */
        { label: "cm",    pointsPerUnit: 72 / 2.54 },         /* 4 */
        { label: "Q",     pointsPerUnit: 72 / 25.4 * 0.25 },  /* 5 */
        { label: "px",    pointsPerUnit: 1 },                 /* 6 */
        { label: "ft/in", pointsPerUnit: 72 * 12 },           /* 7 */
        { label: "m",     pointsPerUnit: 72 / 25.4 * 1000 },  /* 8 */
        { label: "yd",    pointsPerUnit: 72 * 36 },           /* 9 */
        { label: "ft",    pointsPerUnit: 72 * 12 }            /* 10 */
    ];

    /* 単位コード5を「歯（H）」と表示する環境設定キー。文字サイズ（text/units）だけ「級（Q）」
       Preference keys that show unit code 5 as H; only the type size (text/units) shows Q */
    var HA_UNIT_PREF_KEYS = { "rulerType": true, "strokeUnits": true, "text/asianunits": true };

    /**
     * 環境設定キーの単位を返す
     * @param {string} [prefKey] - "rulerType"（既定）/ "strokeUnits" / "text/units" / "text/asianunits"
     * @returns {{code: number, label: string, pointsPerUnit: number}} 単位の情報
     */
    function getUnitInfo(prefKey) {
        var unitKey = prefKey || "rulerType";
        var unitCode = app.preferences.getIntegerPreference(unitKey);
        /* 未知のコードは pt に寄せる / unknown codes fall back to points */
        var unit = UNITS[unitCode] || UNITS[2];
        /* 級（Q）と歯（H）は同じ長さだが、文字サイズは「Q」、距離は「H」と呼び分ける */
        var label = (unitCode === 5 && HA_UNIT_PREF_KEYS[unitKey]) ? "H" : unit.label;
        return { code: unitCode, label: label, pointsPerUnit: unit.pointsPerUnit };
    }

    /* 単位ラベル → UnitValue に渡す単位名 / Ruler unit label to the unit name passed to UnitValue */
    var UNIT_VALUE_NAMES = {
        "in": "in",
        "mm": "mm",
        "pt": "pt",
        "pica": "pc",
        "cm": "cm",
        "px": "px",
        "m": "m",
        "yd": "yd",
        "ft": "ft",
        "ft/in": "in" /* フィートインチ表示のときは入力値をインチとして扱う / Treat input as inches when the ruler shows feet/inches */
    };

    /* 級（Q）・歯（H）は 1単位 = 0.25mm。UnitValue が扱えないので mm 経由で換算
       One Q or H equals 0.25 mm; UnitValue cannot parse it, so convert through mm */
    var MILLIMETERS_PER_Q = 0.25;
    var Q_UNIT_LABEL = "Q/H";

    /* マージンの最小値（内側へのオフセットのみ許可） / Minimum margin (inward offset only) */
    var MINIMUM_MARGIN_VALUE = 0;

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
    // ステップボタン（再利用パーツ） / Stepper buttons (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内（ローカライズより前）に貼る。
    //    識別子はすべて STEPPER_* / *Stepper* / *Stepped* の名前なので、既存の名前とはぶつからない
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
    /**
     * UIがダークテーマかどうかを判定する（Illustrator・InDesign の両方に対応）
     * @returns {boolean} ダークなら true。取得できない環境では false（明るいUI扱い）
     */
    function isDarkStepperUI() {
        try {
            if (app.preferences && app.preferences.getRealPreference) {
                return app.preferences.getRealPreference("uiBrightness") <= 0.5; /* Illustrator */
            }
            return app.generalPreferences.uiBrightnessPreference <= 0.5; /* InDesign */
        } catch (e) {
            return false;
        }
    }

    var STEPPER_UI_DARK           = isDarkStepperUI();
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
        var upTooltip = stepOptions.integer ? "tooltip.stepUpInteger" : "tooltip.stepUp";
        var downTooltip = stepOptions.integer ? "tooltip.stepDownInteger" : "tooltip.stepDown";
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
    // リンクアイコン（再利用パーツ） / Link toggle (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内に貼る。
    //    識別子はすべて LINK_* / *LinkToggle* / *Link* の名前か、描画の下請け関数（buildArcPoints など）
    // 2. アイコンを addLinkToggle(親, 初期値, 切り替え後の関数) で作る。helpTip はコピー先で付ける
    //      var linkToggle = addLinkToggle(fieldsRowGroup, true, function () { syncFields(); });
    //      linkToggle.helpTip = getLabel(LABELS.tooltip.linkToggle);
    // 3. 連動中かは linkToggle.value で読む。コードから変えるときは setLinkToggleValue(linkToggle, true)
    // 4. 有効／無効は setLinkToggleEnabled(linkToggle, isEnabled)（無効の間はクリックが効かず、薄く描く）
    // 5. 2つの入力欄の右に置くときは、行 group の中に「入力欄を縦に積んだ group」とアイコンを並べ、
    //    行 group の alignChildren を ["left", "center"] にすると上下の中央に来る
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

    // -----------------------------------------
    // リンクアイコンの寸法 / Link toggle metrics
    // -----------------------------------------
    var LINK_ICON_SIZE          = [22, 22]; /* アイコンの大きさ / icon size */
    var LINK_ICON_STROKE        = 1.5;      /* 線幅 / stroke width */
    var LINK_CUT_DIRECTION      = [1, 0];   /* 連動中の左辺の切れ目の向き（水平）/ direction of the left-leg cut when linked (horizontal) */
    var LINK_HOOK_CUT_DIRECTION = [0, 1];   /* 連動中の巻き込みの切れ目の向き（垂直）/ direction of the hook cut when linked (vertical) */
    var LINK_STRAND_COUNT       = 4;        /* 切れ目の向きをそろえるための細い線の本数 / strands used to shape the cuts */
    var LINK_SLASH_CLEARANCE    = 2.2;      /* 連動OFFの斜線とフックの間（22px 基準）/ gap between the slash and the hooks when unlinked */

    // -----------------------------------------
    // リンクアイコンの配色 / Link toggle colors
    // -----------------------------------------
    /**
     * UIがダークテーマかどうかを判定する（Illustrator・InDesign の両方に対応）
     * @returns {boolean} ダークなら true。取得できない環境では false（明るいUI扱い）
     */
    function isDarkLinkToggleUI() {
        try {
            if (app.preferences && app.preferences.getRealPreference) {
                return app.preferences.getRealPreference("uiBrightness") <= 0.5; /* Illustrator */
            }
            return app.generalPreferences.uiBrightnessPreference <= 0.5; /* InDesign */
        } catch (e) {
            return false;
        }
    }

    var LINK_UI_DARK = isDarkLinkToggleUI();
    /* ダイアログの地に重ねる半透明の黒・白（UIの明るさの段階に追従する）。値はステップボタンの配色と同じ
       Translucent overlays that follow the dialog background; same values as the stepper buttons */
    var LINK_PRESSED_COLOR  = LINK_UI_DARK ? [1, 1, 1, 0.12] : [0, 0, 0, 0.13]; /* 連動中の地 / background while linked */
    var LINK_FRAME_COLOR    = LINK_UI_DARK ? [1, 1, 1, 0.07] : [0, 0, 0, 0.10]; /* 連動中の枠 / frame while linked */
    var LINK_ICON_COLOR     = LINK_UI_DARK ? [1, 1, 1, 1]    : [0, 0, 0, 0.70]; /* アイコンの線 / icon strokes */
    var LINK_DIM_ICON_COLOR = LINK_UI_DARK ? [1, 1, 1, 0.20] : [0, 0, 0, 0.25]; /* 無効時の線 / strokes when disabled */

    // -----------------------------------------
    // アイコンを作る・切り替える（外から呼ぶ関数） / Public API
    // -----------------------------------------
    /**
     * 連動の ON／OFF を切り替えるリンクアイコンを追加する（onDraw で自作描画）。
     * クリックで切り替わる。連動中は押し込んだボタンのように地と枠を描く。
     * @param {Group} parent - 追加先
     * @param {boolean} initialValue - 連動の初期値
     * @param {Function} onToggle - 切り替えたあとに呼ぶ関数
     * @returns {Group} アイコン（.value で連動中かを読む）
     */
    function addLinkToggle(parent, initialValue, onToggle) {
        var linkToggle = parent.add("group");
        linkToggle.preferredSize = LINK_ICON_SIZE;
        linkToggle.minimumSize = LINK_ICON_SIZE;
        linkToggle.maximumSize = LINK_ICON_SIZE;
        linkToggle.value = initialValue;

        linkToggle.onDraw = function () {
            var iconGraphics = linkToggle.graphics;
            var iconWidth = LINK_ICON_SIZE[0];
            var iconHeight = LINK_ICON_SIZE[1];
            /* 自作描画は自動でディムにならないため、親もたどって判定する / Custom drawing is not dimmed automatically */
            var isDimmed = !isLinkToggleEnabledInTree(linkToggle);
            /* 連動中は押し込んだボタンのように地と枠を描く / While linked, draw it like a pressed button */
            if (linkToggle.value && !isDimmed) {
                iconGraphics.newPath();
                iconGraphics.rectPath(0, 0, iconWidth, iconHeight);
                iconGraphics.fillPath(iconGraphics.newBrush(iconGraphics.BrushType.SOLID_COLOR, LINK_PRESSED_COLOR));
                iconGraphics.newPath();
                iconGraphics.rectPath(0.5, 0.5, iconWidth - 1, iconHeight - 1);
                iconGraphics.strokePath(iconGraphics.newPen(iconGraphics.PenType.SOLID_COLOR, LINK_FRAME_COLOR, 1));
            }
            drawLinkIcon(iconGraphics, iconWidth, iconHeight, linkToggle.value, isDimmed ? LINK_DIM_ICON_COLOR : LINK_ICON_COLOR);
        };

        linkToggle.addEventListener("mousedown", function () {
            if (!isLinkToggleEnabledInTree(linkToggle)) return;
            linkToggle.value = !linkToggle.value;
            redrawLinkToggle(linkToggle);
            if (onToggle) onToggle();
        });
        return linkToggle;
    }

    /**
     * 連動の状態をコードから変えて描き直す（onToggle は呼ばない）
     * @param {Group} linkToggle - addLinkToggle() で作ったアイコン
     * @param {boolean} isLinked - 連動にするなら true
     * @returns {void}
     */
    function setLinkToggleValue(linkToggle, isLinked) {
        if (linkToggle.value === isLinked) return;
        linkToggle.value = isLinked;
        redrawLinkToggle(linkToggle);
    }

    /**
     * アイコンの有効／無効を切り替えて描き直す（変わらないときは描き直さない）
     * @param {Group} linkToggle - addLinkToggle() で作ったアイコン
     * @param {boolean} isEnabled - 有効にするなら true
     * @returns {void}
     */
    function setLinkToggleEnabled(linkToggle, isEnabled) {
        if (linkToggle.enabled === isEnabled) return;
        linkToggle.enabled = isEnabled;
        redrawLinkToggle(linkToggle);
    }

    /**
     * コントロールと親がすべて有効かを判定する（親の無効化は子の enabled に出ないため、親もたどる）
     * @param {Object} control - 判定するコントロール
     * @returns {boolean} すべて有効なら true
     */
    function isLinkToggleEnabledInTree(control) {
        for (var node = control; node; node = node.parent) {
            if (!node.enabled) return false;
        }
        return true;
    }

    /**
     * group の onDraw を呼び直す。group には notify() が無いため、隠して再表示して描き直させる
     * @param {Group} linkToggle - 描き直すアイコン
     * @returns {void}
     */
    function redrawLinkToggle(linkToggle) {
        linkToggle.hide();
        linkToggle.show();
    }

    // -----------------------------------------
    // アイコンの形 / Icon geometry
    // -----------------------------------------
    /**
     * 連動アイコンを描く。Illustrator の［縦横比を固定］に合わせ、連動中は縦につながったチェーン、
     * 連動していないときは上下に分かれたチェーンに斜線を重ねる。座標は 22px 四方を基準に拡大縮小する。
     * @param {ScriptUIGraphics} iconGraphics - 描画先
     * @param {number} iconWidth - 描画範囲の幅
     * @param {number} iconHeight - 描画範囲の高さ
     * @param {boolean} isLinked - 連動中なら true
     * @param {number[]} iconColor - [r, g, b, a]
     * @returns {void}
     */
    function drawLinkIcon(iconGraphics, iconWidth, iconHeight, isLinked, iconColor) {
        var iconScale = Math.min(iconWidth, iconHeight) / 22;
        var offsetX = (iconWidth - 22 * iconScale) / 2;
        var offsetY = (iconHeight - 22 * iconScale) / 2;
        var strokes = isLinked ? buildLinkedChainStrokes() : buildUnlinkedChainStrokes();
        for (var i = 0; i < strokes.length; i++) {
            var strokePoints = strokes[i].points;
            /* newPath() を呼ばないとパスが前の描画に積み重なる / Without newPath() the paths accumulate */
            iconGraphics.newPath();
            for (var j = 0; j < strokePoints.length; j++) {
                var pointX = offsetX + strokePoints[j][0] * iconScale;
                var pointY = offsetY + strokePoints[j][1] * iconScale;
                if (j === 0) iconGraphics.moveTo(pointX, pointY);
                else iconGraphics.lineTo(pointX, pointY);
            }
            iconGraphics.strokePath(iconGraphics.newPen(iconGraphics.PenType.SOLID_COLOR, iconColor, strokes[i].width * iconScale));
        }
    }

    /**
     * 連動中のチェーン（縦に組み合った2つの輪）の線を返す。
     * 上の輪は左辺の途中から上端を回って右辺を下り、下端で内側へ巻き込む。下の輪はそれを180度回したもの。
     * 切れ目の向きをそろえるため、輪を細い線の束にし、両端を延ばしてから直線で切る（左辺は水平、巻き込みは垂直）
     * @returns {Array<{points: Array<number[]>, width: number}>} 線ごとの点列と線幅（22px 四方の座標）
     */
    function buildLinkedChainStrokes() {
        /* 左辺は上端の丸みだけ残して短く切り、下の輪の巻き込みとの間を空ける
           Keep only a stub on the left so it stays clear of the lower ring's hook */
        var upperRing = densifyPoints(buildArcPoints(11, 7, 3.5, 3.5, 180, 360)
            .concat([[14.5, 11.2]])
            .concat(buildArcPoints(11, 11.2, 3.5, 2.3, 0, 115)));
        var ringStart = upperRing[0];
        var ringEnd = upperRing[upperRing.length - 1];
        var extendedRing = extendPolylineEnds(upperRing, LINK_ICON_STROKE);
        /* 延ばした先がどちら側かで、切り捨てる側を決める / The extended tips tell which side to cut away */
        var startOutsideSign = sideOfLine(extendedRing[0], ringStart, LINK_CUT_DIRECTION);
        var endOutsideSign = sideOfLine(extendedRing[extendedRing.length - 1], ringEnd, LINK_HOOK_CUT_DIRECTION);

        var upperStrands = buildStrandStrokes(extendedRing, function (strandPoints) {
            var trimmed = trimPolylineTail(strandPoints, ringEnd, LINK_HOOK_CUT_DIRECTION, endOutsideSign);
            trimmed = trimPolylineTail(trimmed.reverse(), ringStart, LINK_CUT_DIRECTION, startOutsideSign).reverse();
            return [trimmed];
        });
        var strokes = [];
        for (var i = 0; i < upperStrands.length; i++) {
            strokes.push(upperStrands[i]);
            strokes.push({ points: rotatePointsHalfTurn(upperStrands[i].points), width: upperStrands[i].width });
        }
        return strokes;
    }

    /**
     * 中心線を線幅の中で等分した細い線に分け、clipStrand で切った結果を線として返す。
     * @param {Array<number[]>} centerline - 中心線の点列
     * @param {Function} clipStrand - 細い線の点列を受け取り、残す点列の配列を返す関数
     * @returns {Array<{points: Array<number[]>, width: number}>} 細い線ごとの点列と線幅
     */
    function buildStrandStrokes(centerline, clipStrand) {
        var strandWidth = LINK_ICON_STROKE / LINK_STRAND_COUNT;
        var strokes = [];
        for (var k = 0; k < LINK_STRAND_COUNT; k++) {
            /* 線幅の中を等分した位置に細い線を並べる / Lay the strands evenly across the stroke width */
            var strandOffset = -LINK_ICON_STROKE / 2 + strandWidth * (k + 0.5);
            var strandPieces = clipStrand(offsetPolyline(centerline, strandOffset));
            for (var j = 0; j < strandPieces.length; j++) {
                /* 隣の線と少し重ねて隙間を埋める / Overlap neighbours slightly so no seams show */
                if (strandPieces[j].length > 1) strokes.push({ points: strandPieces[j], width: strandWidth * 1.4 });
            }
        }
        return strokes;
    }

    /**
     * 点列の両端を、端の向きのまま length だけ延ばす。
     * @param {Array<number[]>} points - 点列
     * @param {number} length - 延ばす長さ
     * @returns {Array<number[]>} 延ばした点列
     */
    function extendPolylineEnds(points, length) {
        /* from から to の向きへ、to から length 先の点 / point length beyond to, heading from from to to */
        function extendBeyond(from, to) {
            var dx = to[0] - from[0];
            var dy = to[1] - from[1];
            var segmentLength = Math.sqrt(dx * dx + dy * dy) || 1;
            return [to[0] + dx / segmentLength * length, to[1] + dy / segmentLength * length];
        }
        var lastIndex = points.length - 1;
        return [extendBeyond(points[1], points[0])].concat(points, [extendBeyond(points[lastIndex - 1], points[lastIndex])]);
    }

    /**
     * 点が直線のどちら側にあるかを符号で返す。
     * @param {number[]} point - 点
     * @param {number[]} linePoint - 直線上の1点
     * @param {number[]} direction - 直線の向き
     * @returns {number} 正・負で側を表す値
     */
    function sideOfLine(point, linePoint, direction) {
        return direction[0] * (point[1] - linePoint[1]) - direction[1] * (point[0] - linePoint[0]);
    }

    /**
     * 点列の終わり側で、直線より outsideSign の側にはみ出した部分を切り、直線との交点で止める。
     * 輪の別の場所が同じ直線をまたいでも切らないよう、終わりから数点の範囲だけを見る。
     * @param {Array<number[]>} points - 点列
     * @param {number[]} cutPoint - 切る直線上の1点
     * @param {number[]} direction - 切る直線の向き
     * @param {number} outsideSign - 切り捨てる側の符号
     * @returns {Array<number[]>} 切った点列
     */
    function trimPolylineTail(points, cutPoint, direction, outsideSign) {
        var lastIndex = points.length - 1;
        var searchLimit = Math.max(0, lastIndex - 12);
        var index = lastIndex;
        while (index > searchLimit && sideOfLine(points[index], cutPoint, direction) * outsideSign > 0) index--;
        if (index === lastIndex) return points.slice(0);
        var inside = points[index];
        var outside = points[index + 1];
        var insideSide = sideOfLine(inside, cutPoint, direction);
        var ratio = insideSide / (insideSide - sideOfLine(outside, cutPoint, direction));
        return points.slice(0, index + 1).concat([[inside[0] + (outside[0] - inside[0]) * ratio, inside[1] + (outside[1] - inside[1]) * ratio]]);
    }

    /**
     * 連動していないときのチェーン（上下に分かれた輪と斜線）の線を返す。
     * フックは斜線の近くで切る。線の端は進む向きに直角にしか切れないため、フックを細い線の束にして
     * 1本ずつ斜線と平行な境界で切り、切り口が斜線に沿って見えるようにする。
     * @returns {Array<{points: Array<number[]>, width: number}>} 線ごとの点列と線幅（22px 四方の座標）
     */
    function buildUnlinkedChainStrokes() {
        var slashStart = [3.5, 3.5];
        var slashEnd = [18.5, 18.5];
        var upperHook = densifyPoints(buildArcPoints(11, 7, 3.5, 3.5, 180, 360).concat([[14.5, 11.5]]));
        var hooks = [upperHook, rotatePointsHalfTurn(upperHook)];

        /* 斜線の近くの帯を切り取る / Cut away the band around the slash */
        function clipAroundSlash(strandPoints) {
            return clipOutsideBand(strandPoints, slashStart, slashEnd, LINK_SLASH_CLEARANCE);
        }
        var strokes = buildStrandStrokes(hooks[0], clipAroundSlash).concat(buildStrandStrokes(hooks[1], clipAroundSlash));
        strokes.push({ points: [slashStart, slashEnd], width: LINK_ICON_STROKE });
        return strokes;
    }

    /**
     * 点の間隔が 0.5 以下になるよう、線分の間に点を足す。
     * @param {Array<number[]>} points - 点列
     * @returns {Array<number[]>} 細かくした点列
     */
    function densifyPoints(points) {
        var densePoints = [points[0]];
        for (var i = 1; i < points.length; i++) {
            var from = points[i - 1];
            var to = points[i];
            var steps = Math.max(1, Math.ceil(Math.sqrt(Math.pow(to[0] - from[0], 2) + Math.pow(to[1] - from[1], 2)) / 0.5));
            for (var j = 1; j <= steps; j++) {
                densePoints.push([from[0] + (to[0] - from[0]) * j / steps, from[1] + (to[1] - from[1]) * j / steps]);
            }
        }
        return densePoints;
    }

    /**
     * 点列を、進む向きの左側へ offset だけずらした点列を返す（負の値なら右側）。
     * @param {Array<number[]>} points - 点列
     * @param {number} offset - ずらす距離
     * @returns {Array<number[]>} ずらした点列
     */
    function offsetPolyline(points, offset) {
        var shifted = [];
        for (var i = 0; i < points.length; i++) {
            var before = points[Math.max(0, i - 1)];
            var after = points[Math.min(points.length - 1, i + 1)];
            var tangentX = after[0] - before[0];
            var tangentY = after[1] - before[1];
            var tangentLength = Math.sqrt(tangentX * tangentX + tangentY * tangentY) || 1;
            shifted.push([points[i][0] - tangentY / tangentLength * offset, points[i][1] + tangentX / tangentLength * offset]);
        }
        return shifted;
    }

    /**
     * 直線（線分を延長したもの）から clearance 未満の帯に入る部分を切り取り、残りを点列に分けて返す。
     * 帯の境界で線分を補間して切るので、切り口は直線と平行にそろう。
     * @param {Array<number[]>} points - 点列
     * @param {number[]} lineStart - 直線上の1点
     * @param {number[]} lineEnd - 直線上のもう1点
     * @param {number} clearance - 空ける距離
     * @returns {Array<Array<number[]>>} 帯の外側に残った点列（2点未満のものは除く）
     */
    function clipOutsideBand(points, lineStart, lineEnd, clearance) {
        var directionX = lineEnd[0] - lineStart[0];
        var directionY = lineEnd[1] - lineStart[1];
        var directionLength = Math.sqrt(directionX * directionX + directionY * directionY);

        /* 直線からの符号付き距離 / signed distance from the line */
        function signedDistance(point) {
            return (directionX * (point[1] - lineStart[1]) - directionY * (point[0] - lineStart[0])) / directionLength;
        }
        /* 2点の間で、距離が boundary になる点 / point between two points where the distance equals boundary */
        function interpolateAt(from, to, fromDistance, toDistance, boundary) {
            var ratio = (boundary - fromDistance) / (toDistance - fromDistance);
            return [from[0] + (to[0] - from[0]) * ratio, from[1] + (to[1] - from[1]) * ratio];
        }

        var pieces = [];
        var currentPiece = [];
        for (var i = 0; i < points.length; i++) {
            var distance = signedDistance(points[i]);
            var isOutside = Math.abs(distance) >= clearance;
            if (i > 0) {
                var previousDistance = signedDistance(points[i - 1]);
                var wasOutside = Math.abs(previousDistance) >= clearance;
                if (wasOutside && !isOutside) {
                    /* 帯に入る: 境界で止める / entering the band: stop at the boundary */
                    currentPiece.push(interpolateAt(points[i - 1], points[i], previousDistance, distance, previousDistance > 0 ? clearance : -clearance));
                    if (currentPiece.length > 1) pieces.push(currentPiece);
                    currentPiece = [];
                } else if (!wasOutside && isOutside) {
                    /* 帯から出る: 境界から始める / leaving the band: start at the boundary */
                    currentPiece = [interpolateAt(points[i - 1], points[i], previousDistance, distance, distance > 0 ? clearance : -clearance)];
                }
            }
            if (isOutside) currentPiece.push(points[i]);
        }
        if (currentPiece.length > 1) pieces.push(currentPiece);
        return pieces;
    }

    /**
     * 楕円弧の点列を返す（角度は右が0度、下が90度の画面座標）。
     * @param {number} centerX - 中心X
     * @param {number} centerY - 中心Y
     * @param {number} radiusX - 横の半径
     * @param {number} radiusY - 縦の半径
     * @param {number} startDegrees - 開始角度
     * @param {number} endDegrees - 終了角度
     * @returns {Array<number[]>} 点列
     */
    function buildArcPoints(centerX, centerY, radiusX, radiusY, startDegrees, endDegrees) {
        var arcSteps = 12;
        var arcPoints = [];
        for (var i = 0; i <= arcSteps; i++) {
            var angle = (startDegrees + (endDegrees - startDegrees) * i / arcSteps) * Math.PI / 180;
            arcPoints.push([centerX + radiusX * Math.cos(angle), centerY + radiusY * Math.sin(angle)]);
        }
        return arcPoints;
    }

    /**
     * 点列を 22px 四方の中心で180度回す。
     * @param {Array<number[]>} points - 点列
     * @returns {Array<number[]>} 回した点列
     */
    function rotatePointsHalfTurn(points) {
        var rotated = [];
        for (var i = 0; i < points.length; i++) {
            rotated.push([22 - points[i][0], 22 - points[i][1]]);
        }
        return rotated;
    }

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // リンクアイコン（再利用パーツ）ここまで / End of the reusable link toggle
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * Illustrator の UI 言語から表示言語を判定する
     * @returns {string} "ja" または "en"
     */
    function detectUILang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = detectUILang();

    var LABELS = {
        dialog: {
            title: { ja: "各アートボードに整列", en: "Align to Artboards" }
        },
        panel: {
            alignmentBase: { ja: "整列の基準", en: "Align based on" },
            anchor: { ja: "整列先", en: "Target" },
            margin: { ja: "マージン", en: "Margin" }
        },
        radio: {
            eachArtboard: { ja: "すべてのアートボード", en: "All Artboards" },
            activeArtboard: { ja: "アクティブなアートボード", en: "Based on Active Artboard" }
        },
        anchor: {
            topLeft: { ja: "左上", en: "Top-Left" },
            topCenter: { ja: "上中央", en: "Top-Center" },
            topRight: { ja: "右上", en: "Top-Right" },
            middleLeft: { ja: "左中央", en: "Middle-Left" },
            center: { ja: "中央", en: "Center" },
            middleRight: { ja: "右中央", en: "Middle-Right" },
            bottomLeft: { ja: "左下", en: "Bottom-Left" },
            bottomCenter: { ja: "下中央", en: "Bottom-Center" },
            bottomRight: { ja: "右下", en: "Bottom-Right" }
        },
        fieldLabel: {
            marginHorizontal: { ja: "左右", en: "Horizontal" },
            marginVertical: { ja: "上下", en: "Vertical" }
        },
        checkbox: {
            useVisibleBounds: { ja: "プレビュー境界を使用", en: "Use Preview Bounds" }
        },
        button: {
            ok: { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        tooltip: {
            eachArtboard: {
                ja: "選択オブジェクトを中心点が属するアートボードごとに振り分け、各アートボード内の指定位置へ整列します。",
                en: "Groups selected objects by the artboard containing their center point, then aligns them to the selected position on each artboard."
            },
            activeArtboard: {
                ja: "アクティブアートボード上の選択を、相対位置を保ったまま整列先＋マージンの位置へそろえ、他のアートボード上の選択も同じ相対位置へ整列します。",
                en: "Aligns the selection on the active artboard to the target position plus margin while keeping its internal layout, then places selections on other artboards at the same relative position."
            },
            anchor: {
                ja: "整列先の9点を選択します。",
                en: "Choose one of the 9 alignment positions."
            },
            margin: {
                ja: "対応する辺から内側へオフセットします。整列先が中央のときは無効です。",
                en: "Offsets objects inward from the corresponding edge. Disabled when the target is Center."
            },
            linkMargins: {
                ja: "ONのときは左右の値を上下にも連動します。",
                en: "When enabled, the horizontal value is also used for the vertical margin."
            },
            useVisibleBounds: {
                ja: "ON：線幅や効果を含む見た目の境界で整列。OFF：図形本体の幾何境界で整列。",
                en: "On: align by visual bounds including strokes and effects. Off: align by geometric bounds of the object shape."
            },
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
        }
    };

    /**
     * LABELS からドット区切りのパスで表示言語のテキストを取り出す
     * @param {string} labelPath - "panel.anchor" のようなドット区切りのキー
     * @returns {string} 表示言語のテキスト（見つからない場合は labelPath をそのまま返す）
     */
    function getLabel(labelPath) {
        var labelPathKeys = labelPath.split(".");
        var labelNode = LABELS;
        for (var i = 0; i < labelPathKeys.length; i++) {
            labelNode = labelNode[labelPathKeys[i]];
            if (!labelNode) return labelPath;
        }
        return labelNode[uiLang] || labelNode["en"] || labelPath;
    }

    /**
     * コロン付きの項目名を返す（日本語は全角、英語は半角）
     * @param {string} labelPath - ラベルのパス
     * @returns {string} コロン付きの項目名
     */
    function labelText(labelPath) {
        return getLabel(labelPath) + (uiLang === "ja" ? "：" : ":");
    }

    // =========================================
    // 境界取得（プレビュー境界 or 幾何境界） / Bounds resolver
    // プレビュー境界＝ストローク・効果を含む、幾何境界＝図形のみ
    // Preview bounds include strokes/effects; geometric bounds use the object shape only
    // =========================================

    /* ダイアログのチェックボックスで切り替わる整列基準の境界 / Bounds type toggled from the dialog */
    var useVisibleBounds = DEFAULT_SETTINGS.useVisibleBounds;

    /**
     * アイテムの整列用バウンディングボックスを返す
     * @param {PageItem} pageItem - 対象アイテム
     * @returns {number[]} [左, 上, 右, 下]。取得できない場合は null
     */
    function getItemBounds(pageItem) {
        /* 文字編集中の TextRange など、境界を持たない選択は対象外 / Skip selections without bounds (e.g. TextRange) */
        try {
            return useVisibleBounds ? pageItem.visibleBounds : pageItem.geometricBounds;
        } catch (e) {
            return null;
        }
    }

    /**
     * 複数アイテムを包含する最小バウンディングを返す
     * @param {PageItem[]} targetItems - 対象アイテムの配列
     * @returns {number[]} [左, 上, 右, 下]。取得できない場合は null
     */
    function getUnionBounds(targetItems) {
        var unionBounds = null;
        for (var i = 0; i < targetItems.length; i++) {
            var itemBounds = getItemBounds(targetItems[i]);
            if (!itemBounds) continue;
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

    // =========================================
    // 単位（rulerType） / Units (rulerType)
    // =========================================

    /**
     * 入力値（現在のルーラー単位）を pt に変換する
     * @param {number} inputValue - 入力値
     * @param {string} unitLabel - 現在の単位ラベル
     * @returns {number} pt に換算した値
     */
    function convertToPoints(inputValue, unitLabel) {
        if (unitLabel === Q_UNIT_LABEL) {
            return new UnitValue(inputValue * MILLIMETERS_PER_Q, "mm").as("pt");
        }
        /* 対応表にない単位は pt として処理 / Units missing from the table are treated as pt */
        return new UnitValue(inputValue, UNIT_VALUE_NAMES[unitLabel] || "pt").as("pt");
    }

    // =========================================
    // アートボード振り分け / Artboard mapping
    // =========================================

    /**
     * アイテムまたは親階層がロック／非表示か判定する
     * @param {PageItem} pageItem - 対象アイテム
     * @returns {boolean} ロックまたは非表示なら true
     */
    function isLockedOrHidden(pageItem) {
        if (pageItem.locked || pageItem.hidden) return true;
        var ancestor = pageItem.parent;
        while (ancestor && ancestor.typename !== "Document") {
            if (ancestor.typename === "Layer") {
                if (ancestor.locked || !ancestor.visible) return true;
            } else if (ancestor.locked || ancestor.hidden) {
                return true;
            }
            ancestor = ancestor.parent;
        }
        return false;
    }

    /**
     * 座標を含むアートボードのインデックスを返す
     * @param {Document} doc - 対象ドキュメント
     * @param {number} pointX - X座標（pt）
     * @param {number} pointY - Y座標（pt）
     * @returns {number} アートボードのインデックス。該当なしは -1
     */
    function findArtboardIndexByPoint(doc, pointX, pointY) {
        for (var i = 0; i < doc.artboards.length; i++) {
            var artboardRect = doc.artboards[i].artboardRect; /* [左, 上, 右, 下] / [L, T, R, B] */
            if (pointX >= artboardRect[0] && pointX <= artboardRect[2] &&
                pointY <= artboardRect[1] && pointY >= artboardRect[3]) {
                return i;
            }
        }
        return -1;
    }

    /**
     * アイテムを中心点が属するアートボードごとに振り分ける
     * @param {Document} doc - 対象ドキュメント
     * @param {PageItem[]} targetItems - 振り分けるアイテムの配列
     * @returns {Object} アートボードインデックスをキーにしたアイテム配列
     */
    function groupItemsByArtboard(doc, targetItems) {
        var itemsByArtboard = {};
        for (var i = 0; i < targetItems.length; i++) {
            var pageItem = targetItems[i];
            var itemBounds = getItemBounds(pageItem);
            if (!itemBounds) continue;
            if (isLockedOrHidden(pageItem)) continue;

            var centerX = (itemBounds[0] + itemBounds[2]) / 2;
            var centerY = (itemBounds[1] + itemBounds[3]) / 2;
            var artboardIndex = findArtboardIndexByPoint(doc, centerX, centerY);
            if (artboardIndex === -1) continue;

            if (!itemsByArtboard[artboardIndex]) itemsByArtboard[artboardIndex] = [];
            itemsByArtboard[artboardIndex].push(pageItem);
        }
        return itemsByArtboard;
    }

    // =========================================
    // プレビュー状態 / Preview state
    // =========================================

    /**
     * プレビュー状態を作成する（移動量を記録して巻き戻せるようにする）
     * @param {PageItem[]} targetItems - プレビュー対象のアイテム
     * @returns {Object} プレビュー状態
     */
    function createPreviewState(targetItems) {
        var previewState = { items: [], offsetsX: [], offsetsY: [] };
        for (var i = 0; i < targetItems.length; i++) {
            previewState.items.push(targetItems[i]);
            previewState.offsetsX.push(0);
            previewState.offsetsY.push(0);
        }
        return previewState;
    }

    /**
     * プレビュー状態にアイテムの移動量を積算する
     * @param {Object} previewState - プレビュー状態。null のときは何もしない
     * @param {PageItem} pageItem - 移動したアイテム
     * @param {number} dx - X方向の移動量
     * @param {number} dy - Y方向の移動量
     * @returns {void}
     */
    function recordPreviewTranslation(previewState, pageItem, dx, dy) {
        if (!previewState) return;
        for (var i = 0; i < previewState.items.length; i++) {
            if (previewState.items[i] !== pageItem) continue;
            previewState.offsetsX[i] += dx;
            previewState.offsetsY[i] += dy;
            return;
        }
    }

    /**
     * プレビューの移動を巻き戻して元の位置に戻す
     * @param {Object} previewState - プレビュー状態
     * @returns {void}
     */
    function revertPreview(previewState) {
        for (var i = 0; i < previewState.items.length; i++) {
            var dx = previewState.offsetsX[i];
            var dy = previewState.offsetsY[i];
            if (dx === 0 && dy === 0) continue;
            translateItem(previewState.items[i], -dx, -dy, null);
            previewState.offsetsX[i] = 0;
            previewState.offsetsY[i] = 0;
        }
    }

    // =========================================
    // 整列処理 / Alignment
    // =========================================

    /**
     * アンカーコードから定義を引くためのマップを作る / Build the lookup map from anchor code to definition
     * @returns {Object} アンカーコードをキーにした ANCHOR_DEFINITIONS の要素
     */
    function buildAnchorDefinitionMap() {
        var definitionByCode = {};
        for (var i = 0; i < ANCHOR_DEFINITIONS.length; i++) {
            definitionByCode[ANCHOR_DEFINITIONS[i].code] = ANCHOR_DEFINITIONS[i];
        }
        return definitionByCode;
    }
    var anchorDefinitionByCode = buildAnchorDefinitionMap();

    /**
     * 矩形上の9点アンカー座標を返す
     * @param {number[]} bounds - [左, 上, 右, 下]
     * @param {string} anchorCode - アンカーコード（LT, TC, RT, LM, C, RM, LB, BC, RB）
     * @returns {number[]} [X座標, Y座標]
     */
    function getAnchorPoint(bounds, anchorCode) {
        var anchorDefinition = anchorDefinitionByCode[anchorCode] || anchorDefinitionByCode[DEFAULT_SETTINGS.anchorCode];
        var left = bounds[0], top = bounds[1], right = bounds[2], bottom = bounds[3];
        return [left + (right - left) * anchorDefinition.ratioX, top - (top - bottom) * anchorDefinition.ratioY];
    }

    /**
     * 矩形を四辺から内側へ縮める（マージンの適用）
     * @param {number[]} bounds - [左, 上, 右, 下]
     * @param {number} marginX - 左右のマージン（pt）
     * @param {number} marginY - 上下のマージン（pt）
     * @returns {number[]} 縮めた矩形 [左, 上, 右, 下]
     */
    function insetBounds(bounds, marginX, marginY) {
        return [bounds[0] + marginX, bounds[1] - marginY, bounds[2] - marginX, bounds[3] + marginY];
    }

    /**
     * アイテムを移動し、プレビュー中は移動量を記録する
     * @param {PageItem} pageItem - 対象アイテム
     * @param {number} dx - X方向の移動量
     * @param {number} dy - Y方向の移動量
     * @param {Object} previewState - プレビュー状態。確定時は null
     * @returns {void}
     */
    function translateItem(pageItem, dx, dy, previewState) {
        if (dx === 0 && dy === 0) return;
        /* クリッピングマスクなど、移動できないアイテムは無視 / Ignore items that cannot be moved */
        try {
            pageItem.translate(dx, dy);
        } catch (e) {
            return;
        }
        recordPreviewTranslation(previewState, pageItem, dx, dy);
    }

    /**
     * アイテム群をまとめて同じ量だけ移動する
     * @param {PageItem[]} targetItems - 対象アイテムの配列
     * @param {number} dx - X方向の移動量
     * @param {number} dy - Y方向の移動量
     * @param {Object} previewState - プレビュー状態。確定時は null
     * @returns {void}
     */
    function translateItems(targetItems, dx, dy, previewState) {
        for (var i = 0; i < targetItems.length; i++) {
            translateItem(targetItems[i], dx, dy, previewState);
        }
    }

    /**
     * 各アイテムを個別に、アートボード上のアンカーへ整列する
     * @param {PageItem[]} targetItems - 対象アイテムの配列
     * @param {number[]} artboardRect - アートボードの矩形 [左, 上, 右, 下]
     * @param {string} anchorCode - アンカーコード
     * @param {number} marginX - 左右のマージン（pt）
     * @param {number} marginY - 上下のマージン（pt）
     * @param {Object} previewState - プレビュー状態。確定時は null
     * @returns {void}
     */
    function alignItemsToArtboardAnchor(targetItems, artboardRect, anchorCode, marginX, marginY, previewState) {
        var targetPoint = getAnchorPoint(insetBounds(artboardRect, marginX, marginY), anchorCode);
        for (var i = 0; i < targetItems.length; i++) {
            var itemBounds = getItemBounds(targetItems[i]);
            if (!itemBounds) continue;
            var itemAnchor = getAnchorPoint(itemBounds, anchorCode);
            translateItem(targetItems[i], targetPoint[0] - itemAnchor[0], targetPoint[1] - itemAnchor[1], previewState);
        }
    }

    /**
     * アイテム群の相対位置を保ったまま、アートボード上のアンカー＋オフセット位置へ整列する
     * @param {PageItem[]} targetItems - 対象アイテムの配列
     * @param {number[]} artboardRect - アートボードの矩形 [左, 上, 右, 下]
     * @param {string} anchorCode - アンカーコード
     * @param {Object} groupOffset - アンカーからの相対位置 {x, y}
     * @param {Object} previewState - プレビュー状態。確定時は null
     * @returns {void}
     */
    function alignItemGroupToArtboardAnchor(targetItems, artboardRect, anchorCode, groupOffset, previewState) {
        var groupBounds = getUnionBounds(targetItems);
        if (!groupBounds) return;
        var artboardAnchor = getAnchorPoint(artboardRect, anchorCode);
        var groupAnchor = getAnchorPoint(groupBounds, anchorCode);
        translateItems(
            targetItems,
            artboardAnchor[0] + groupOffset.x - groupAnchor[0],
            artboardAnchor[1] + groupOffset.y - groupAnchor[1],
            previewState
        );
    }

    /**
     * アクティブアートボード上の選択を基準に、他のアートボードの選択を同じ相対位置へ整列する
     * @param {Document} doc - 対象ドキュメント
     * @param {Object} itemsByArtboard - アートボードごとに振り分けたアイテム
     * @param {string} anchorCode - アンカーコード
     * @param {number} marginX - 左右のマージン（pt）
     * @param {number} marginY - 上下のマージン（pt）
     * @param {Object} previewState - プレビュー状態。確定時は null
     * @returns {void}
     */
    function alignUsingActiveArtboardReference(doc, itemsByArtboard, anchorCode, marginX, marginY, previewState) {
        var activeArtboardIndex = doc.artboards.getActiveArtboardIndex();
        if (activeArtboardIndex < 0) return;

        var referenceItems = itemsByArtboard[activeArtboardIndex];
        if (!referenceItems) return;

        /* まず基準側を、相対位置を保ったままマージンぶん内側の基準点へ寄せる
           アートボードが1つでも整列が効くのはこの一手があるため
           Move the reference group to the anchor inside the margin first; this is what makes the
           alignment work even when the document has only one artboard */
        var activeArtboardRect = doc.artboards[activeArtboardIndex].artboardRect;
        alignItemGroupToArtboardAnchor(
            referenceItems,
            insetBounds(activeArtboardRect, marginX, marginY),
            anchorCode,
            { x: 0, y: 0 },
            previewState
        );

        /* 寄せたあとの位置を測り直す / Re-measure after the reference group has moved */
        var referenceBounds = getUnionBounds(referenceItems);
        if (!referenceBounds) return;

        /* 基準アートボードのアンカーから見た、基準オブジェクト群の相対位置
           Relative position of the reference group as seen from the artboard anchor */
        var referenceAnchor = getAnchorPoint(activeArtboardRect, anchorCode);
        var groupAnchor = getAnchorPoint(referenceBounds, anchorCode);
        var referenceOffset = { x: groupAnchor[0] - referenceAnchor[0], y: groupAnchor[1] - referenceAnchor[1] };

        for (var artboardIndex in itemsByArtboard) {
            if (!itemsByArtboard.hasOwnProperty(artboardIndex)) continue;
            var targetIndex = parseInt(artboardIndex, 10);
            if (targetIndex === activeArtboardIndex) continue;
            alignItemGroupToArtboardAnchor(
                itemsByArtboard[artboardIndex],
                doc.artboards[targetIndex].artboardRect,
                anchorCode,
                referenceOffset,
                previewState
            );
        }
    }

    /**
     * 設定に従って整列を実行する（プレビューと確定で共用）
     * @param {Document} doc - 対象ドキュメント
     * @param {PageItem[]} targetItems - 整列するアイテムの配列
     * @param {Object} alignSettings - 整列設定 {anchorCode, useActiveArtboardAsReference, marginX, marginY}
     * @param {Object} previewState - プレビュー状態。確定時は null
     * @returns {void}
     */
    function alignItems(doc, targetItems, alignSettings, previewState) {
        var itemsByArtboard = groupItemsByArtboard(doc, targetItems);

        if (alignSettings.useActiveArtboardAsReference) {
            alignUsingActiveArtboardReference(
                doc,
                itemsByArtboard,
                alignSettings.anchorCode,
                alignSettings.marginX,
                alignSettings.marginY,
                previewState
            );
            return;
        }

        for (var artboardIndex in itemsByArtboard) {
            if (!itemsByArtboard.hasOwnProperty(artboardIndex)) continue;
            alignItemsToArtboardAnchor(
                itemsByArtboard[artboardIndex],
                doc.artboards[parseInt(artboardIndex, 10)].artboardRect,
                alignSettings.anchorCode,
                alignSettings.marginX,
                alignSettings.marginY,
                previewState
            );
        }
    }

    // =========================================
    // 入力ユーティリティ / Input utilities
    // =========================================

    /**
     * 文字列を数値に変換する（不正値は既定値）
     * @param {string} inputText - 入力文字列
     * @param {number} defaultValue - 変換できないときに返す値
     * @returns {number} 変換した数値
     */
    function parseNumericInput(inputText, defaultValue) {
        var trimmedText = ("" + inputText).replace(/^\s+|\s+$/g, "");
        if (trimmedText === "") return defaultValue;
        var parsedValue = Number(trimmedText);
        return isNaN(parsedValue) ? defaultValue : parsedValue;
    }

    // =========================================
    // 9軸ウィジェットの描画 / Anchor widget drawing
    // =========================================

    /* 外周の□どうしをつなぐケイ線の組み合わせ（中央は独立） / Rules joining the outer squares (center stands alone) */
    var ANCHOR_CONNECTIONS = [[0, 1], [1, 2], [6, 7], [7, 8], [0, 3], [3, 6], [2, 5], [5, 8]];

    /* UIの明暗に合わせて initAnchorColors() で設定 / Set from the light/dark UI in initAnchorColors() */
    var anchorLineColor = [0.6, 0.6, 0.6, 1];
    var anchorSelectedFillColor = [0.4, 0.4, 0.4, 1];

    /**
     * IllustratorのUIが明るいテーマか判定する
     * @returns {boolean} 明るいテーマなら true
     */
    function isLightUI() {
        return app.preferences.getRealPreference("uiBrightness") > 0.5;
    }

    /**
     * 9軸ウィジェットの配色をUIの明暗に合わせて決める
     * @returns {void}
     */
    function initAnchorColors() {
        var lightUI = isLightUI();
        /* 選択セルの塗り：ライトは濃いグレー、ダークは明るいグレー / Selected-cell fill: dark gray in light UI, bright gray in dark UI */
        anchorLineColor = lightUI ? [0.6, 0.6, 0.6, 1] : [0.55, 0.55, 0.55, 1];
        anchorSelectedFillColor = lightUI ? [0.4, 0.4, 0.4, 1] : [0.8, 0.8, 0.8, 1];
    }

    /**
     * 3×3のグリッド位置を 0〜2 に収める
     * @param {number} gridPosition - 計算した行または列
     * @returns {number} 0〜2 に丸めた値
     */
    function clampGridIndex(gridPosition) {
        if (gridPosition < 0) return 0;
        if (gridPosition > 2) return 2;
        return gridPosition;
    }

    /**
     * アンカーコードから ANCHOR_DEFINITIONS 上のインデックスを求める
     * @param {string} anchorCode - アンカーコード
     * @returns {number} インデックス。見つからない場合は 0
     */
    function findAnchorIndexByCode(anchorCode) {
        for (var i = 0; i < ANCHOR_DEFINITIONS.length; i++) {
            if (ANCHOR_DEFINITIONS[i].code === anchorCode) return i;
        }
        return 0;
    }

    /**
     * 9軸ウィジェットを再描画する
     * @param {Button} anchorWidget - 対象ウィジェット
     * @returns {void}
     */
    function redrawAnchorWidget(anchorWidget) {
        /* notify は環境により例外を投げ得るので保護 / notify can throw in some environments */
        try {
            anchorWidget.notify("onDraw");
        } catch (e) { }
    }

    /**
     * 9軸ウィジェットを描画する（外周の□をケイ線でつなぎ、選択セルを塗る）
     * @param {Button} anchorWidget - 対象ウィジェット
     * @returns {void}
     */
    function drawAnchorWidget(anchorWidget) {
        var graphics = anchorWidget.graphics;
        var widgetWidth = anchorWidget.size[0];
        var widgetHeight = anchorWidget.size[1];

        /* 背景はコントロールの地色で塗り、パネルと同色に見せる / Paint the control background so the widget blends into the panel */
        try {
            graphics.rectPath(0, 0, widgetWidth, widgetHeight);
            graphics.fillPath(graphics.backgroundColor);
        } catch (e) { }

        var cellStep = ANCHOR_CELL_SIZE + ANCHOR_CELL_GAP;
        var gridSize = ANCHOR_CELL_SIZE * 3 + ANCHOR_CELL_GAP * 2;
        var originX = Math.round((widgetWidth - gridSize) / 2);
        var originY = Math.round((widgetHeight - gridSize) / 2);

        /**
         * セルの左上座標を返す
         * @param {number} cellIndex - 0〜8のセル番号（行優先）
         * @returns {number[]} [X座標, Y座標]
         */
        function getCellOrigin(cellIndex) {
            return [
                originX + (cellIndex % 3) * cellStep,
                originY + Math.floor(cellIndex / 3) * cellStep
            ];
        }

        var linePen = graphics.newPen(graphics.PenType.SOLID_COLOR, anchorLineColor, 1);
        for (var i = 0; i < ANCHOR_CONNECTIONS.length; i++) {
            var fromCell = getCellOrigin(ANCHOR_CONNECTIONS[i][0]);
            var toCell = getCellOrigin(ANCHOR_CONNECTIONS[i][1]);
            var isHorizontal = (ANCHOR_CONNECTIONS[i][1] - ANCHOR_CONNECTIONS[i][0] === 1);
            graphics.newPath();
            if (isHorizontal) {
                /* 横方向：右隣の□へ / Horizontal: to the square on the right */
                graphics.moveTo(fromCell[0] + ANCHOR_CELL_SIZE, fromCell[1] + ANCHOR_CELL_SIZE / 2);
                graphics.lineTo(toCell[0], toCell[1] + ANCHOR_CELL_SIZE / 2);
            } else {
                /* 縦方向：下の□へ / Vertical: to the square below */
                graphics.moveTo(fromCell[0] + ANCHOR_CELL_SIZE / 2, fromCell[1] + ANCHOR_CELL_SIZE);
                graphics.lineTo(toCell[0] + ANCHOR_CELL_SIZE / 2, toCell[1]);
            }
            graphics.strokePath(linePen);
        }

        for (var cellIndex = 0; cellIndex < ANCHOR_DEFINITIONS.length; cellIndex++) {
            var cellOrigin = getCellOrigin(cellIndex);
            drawAnchorCell(
                graphics,
                cellOrigin[0],
                cellOrigin[1],
                cellIndex === anchorWidget.selectedAnchorIndex
            );
        }
    }

    /**
     * 9軸ウィジェットのセルを1つ描画する（選択時のみ塗りつぶす）
     * @param {ScriptUIGraphics} graphics - 描画対象のグラフィックス
     * @param {number} cellX - セルの左端
     * @param {number} cellY - セルの上端
     * @param {boolean} isSelected - 選択中なら true
     * @returns {void}
     */
    function drawAnchorCell(graphics, cellX, cellY, isSelected) {
        /**
         * セルの四角形パスを作る
         * @returns {void}
         */
        function addCellPath() {
            graphics.newPath();
            graphics.moveTo(cellX, cellY);
            graphics.lineTo(cellX + ANCHOR_CELL_SIZE, cellY);
            graphics.lineTo(cellX + ANCHOR_CELL_SIZE, cellY + ANCHOR_CELL_SIZE);
            graphics.lineTo(cellX, cellY + ANCHOR_CELL_SIZE);
            graphics.closePath();
        }

        /* 選択中のみ塗る（枠は塗りの上に描く） / Fill only when selected; the border is drawn over the fill */
        if (isSelected) {
            addCellPath();
            graphics.fillPath(graphics.newBrush(graphics.BrushType.SOLID_COLOR, anchorSelectedFillColor));
        }
        addCellPath();
        graphics.strokePath(graphics.newPen(graphics.PenType.SOLID_COLOR, anchorLineColor, 1));
    }

    // =========================================
    // ダイアログ UI / Dialog UI
    // =========================================

    /**
     * パネルの共通レイアウトを適用する
     * @param {Panel} targetPanel - 対象パネル
     * @param {string} orientation - "column" または "row"
     * @returns {void}
     */
    function applyPanelLayout(targetPanel, orientation) {
        targetPanel.orientation = orientation;
        targetPanel.margins = PANEL_MARGINS;
    }

    /**
     * 整列の基準パネル（すべてのアートボード／アクティブを基準）を構築する
     * @param {Window} parentContainer - 追加先のダイアログ
     * @param {number} artboardCount - ドキュメントのアートボード数
     * @param {Function} onSettingsChanged - 選択が変わったときに呼ぶ関数
     * @returns {Object} 基準の取得・設定用インターフェース
     */
    function buildAlignmentBasePanel(parentContainer, artboardCount, onSettingsChanged) {
        var basePanel = parentContainer.add("panel", undefined, getLabel("panel.alignmentBase"));
        applyPanelLayout(basePanel, "column");
        basePanel.alignChildren = ["left", "center"];

        var eachArtboardRadio = basePanel.add("radiobutton", undefined, getLabel("radio.eachArtboard"));
        var activeArtboardRadio = basePanel.add("radiobutton", undefined, getLabel("radio.activeArtboard"));
        eachArtboardRadio.helpTip = getLabel("tooltip.eachArtboard");
        activeArtboardRadio.helpTip = getLabel("tooltip.activeArtboard");

        /* アートボードが1つのときは振り分け先がないので「すべてのアートボード」を選べない
           With a single artboard there is nothing to distribute, so All Artboards stays unavailable */
        var allowsEachArtboard = artboardCount > 1;
        eachArtboardRadio.enabled = allowsEachArtboard;

        /**
         * 整列の基準を切り替える（選べないときはアクティブ基準に固定）
         * @param {boolean} useActiveArtboard - アクティブを基準にするなら true
         * @returns {void}
         */
        function selectAlignmentBase(useActiveArtboard) {
            var usesActive = useActiveArtboard || !allowsEachArtboard;
            eachArtboardRadio.value = !usesActive;
            activeArtboardRadio.value = usesActive;
        }

        selectAlignmentBase(DEFAULT_SETTINGS.useActiveArtboardAsReference);
        eachArtboardRadio.onClick = onSettingsChanged;
        activeArtboardRadio.onClick = onSettingsChanged;

        return {
            selectAlignmentBase: selectAlignmentBase,
            usesActiveArtboard: function () {
                return activeArtboardRadio.value === true;
            }
        };
    }

    /**
     * 整列先パネル（3×3の9軸ウィジェット）を構築する
     * @param {Group} parentContainer - 追加先のグループ
     * @param {Function} onSettingsChanged - 選択が変わったときに呼ぶ関数
     * @returns {Object} パネル参照とアンカー取得・ショートカット処理をまとめたオブジェクト
     */
    function buildAnchorPanel(parentContainer, onSettingsChanged) {
        var anchorPanel = parentContainer.add("panel", undefined, getLabel("panel.anchor"));
        applyPanelLayout(anchorPanel, "column");
        anchorPanel.margins = ANCHOR_PANEL_MARGINS;
        anchorPanel.alignChildren = ["center", "top"];

        var anchorWidget = anchorPanel.add("button", undefined, "");
        anchorWidget.preferredSize = [ANCHOR_WIDGET_SIZE, ANCHOR_WIDGET_SIZE];
        anchorWidget.minimumSize = [ANCHOR_WIDGET_SIZE, ANCHOR_WIDGET_SIZE];
        anchorWidget.maximumSize = [ANCHOR_WIDGET_SIZE, ANCHOR_WIDGET_SIZE];
        anchorWidget.selectedAnchorIndex = findAnchorIndexByCode(DEFAULT_SETTINGS.anchorCode);
        anchorWidget.onDraw = function () {
            drawAnchorWidget(this);
        };

        /**
         * 指定したセルを選択し、ウィジェットとツールチップを更新する
         * @param {number} selectedIndex - ANCHOR_DEFINITIONS 上のインデックス
         * @returns {void}
         */
        function selectAnchorAt(selectedIndex) {
            var anchorDefinition = ANCHOR_DEFINITIONS[selectedIndex];
            anchorWidget.selectedAnchorIndex = selectedIndex;
            anchorWidget.helpTip = getLabel("tooltip.anchor") + "\n" +
                getLabel("anchor." + anchorDefinition.labelKey) + " (" + anchorDefinition.shortcutKey + ")";
            anchorPanel.helpTip = anchorWidget.helpTip;
            redrawAnchorWidget(anchorWidget);
        }

        /* クリックした3×3のセルを整列先にする（座標はウィジェット基準）
           Set the anchor from the clicked 3x3 cell (coordinates are widget-relative) */
        anchorWidget.addEventListener("mousedown", function (event) {
            var cellColumn = clampGridIndex(Math.floor(event.clientX / (anchorWidget.size[0] / 3)));
            var cellRow = clampGridIndex(Math.floor(event.clientY / (anchorWidget.size[1] / 3)));
            selectAnchorAt(cellRow * 3 + cellColumn);
            onSettingsChanged();
        });

        selectAnchorAt(anchorWidget.selectedAnchorIndex);

        return {
            getAnchorCode: function () {
                return ANCHOR_DEFINITIONS[anchorWidget.selectedAnchorIndex].code;
            },
            selectByShortcutKey: function (pressedKey) {
                for (var i = 0; i < ANCHOR_DEFINITIONS.length; i++) {
                    if (ANCHOR_DEFINITIONS[i].shortcutKey !== pressedKey) continue;
                    selectAnchorAt(i);
                    return true;
                }
                return false;
            }
        };
    }

    /**
     * マージンパネル（左右・上下・連動アイコン）を構築する
     * @param {Group} parentContainer - 追加先のグループ
     * @param {string} unitLabel - 現在のルーラー単位ラベル
     * @param {Function} onSettingsChanged - 入力が変わったときに呼ぶ関数
     * @returns {Object} パネル参照とマージン取得・フォーカス制御をまとめたオブジェクト
     */
    function buildMarginPanel(parentContainer, unitLabel, onSettingsChanged) {
        var marginPanel = parentContainer.add("panel", undefined, getLabel("panel.margin") + " (" + unitLabel + ")");
        applyPanelLayout(marginPanel, "row");
        marginPanel.alignChildren = ["fill", "center"];
        marginPanel.helpTip = getLabel("tooltip.margin");

        var fieldColumn = marginPanel.add("group");
        fieldColumn.orientation = "column";
        fieldColumn.alignChildren = ["left", "center"];

        /**
         * ラベル付きのマージン入力欄を作る
         * @param {string} fieldLabelPath - ラベルの LABELS パス
         * @returns {EditText} 追加した入力欄
         */
        function addMarginField(fieldLabelPath) {
            var fieldGroup = fieldColumn.add("group");
            fieldGroup.orientation = "row";
            fieldGroup.alignChildren = ["left", "center"];
            var fieldLabel = fieldGroup.add("statictext", undefined, labelText(fieldLabelPath));
            fieldLabel.helpTip = marginPanel.helpTip;
            /* ∧∨と入力欄は隙間0で突き合わせる / butt the stepper against the field */
            var stepperInputGroup = fieldGroup.add("group");
            stepperInputGroup.orientation = "row";
            stepperInputGroup.alignChildren = ["left", "center"];
            stepperInputGroup.spacing = 0;
            stepperInputGroup.margins = 0;
            var marginField;
            /* 下限は既存の readMarginField() と同じ。増減後は手入力と同じ onChanging を通す
               Same minimum as readMarginField(); run the same onChanging as typing */
            var stepperGroup = addStepper(stepperInputGroup, function () { return marginField; }, {
                min: MINIMUM_MARGIN_VALUE,
                onStep: function (numberInput) { if (numberInput.onChanging) numberInput.onChanging(); }
            });
            marginField = stepperInputGroup.add("edittext", undefined, DEFAULT_SETTINGS.marginText);
            marginField.characters = MARGIN_FIELD_CHARACTERS;
            marginField.helpTip = marginPanel.helpTip;
            marginField.stepperGroup = stepperGroup;
            bindSteppedArrowKeys(marginField, stepperGroup);
            return marginField;
        }

        var horizontalField = addMarginField("fieldLabel.marginHorizontal");
        var verticalField = addMarginField("fieldLabel.marginVertical");

        /* 左右・上下の2行の右に連動アイコンを置く / Put the link icon to the right of the two margin rows */
        var linkToggle = addLinkToggle(marginPanel, DEFAULT_SETTINGS.linkMargins, function () { handleHorizontalMarginChanged(); });
        linkToggle.alignment = ["left", "center"];
        linkToggle.helpTip = getLabel("tooltip.linkMargins");

        /**
         * 入力欄の値を読み取る（マージンは内側へのオフセットなので負値は0として扱う）
         * @param {EditText} marginField - 対象の入力欄
         * @returns {number} 0以上の入力値
         */
        function readMarginField(marginField) {
            var marginValue = parseNumericInput(marginField.text, 0);
            return (marginValue < MINIMUM_MARGIN_VALUE) ? MINIMUM_MARGIN_VALUE : marginValue;
        }

        /**
         * 入力確定時に負値を下限へ丸めて表示に反映する
         * @param {EditText} marginField - 対象の入力欄
         * @returns {void}
         */
        function normalizeMarginField(marginField) {
            var marginValue = readMarginField(marginField);
            if (marginValue !== parseNumericInput(marginField.text, 0)) marginField.text = marginValue;
        }

        /**
         * 連動ONのときは左右の値を上下へ反映し、上下の入力欄を無効化する
         * @returns {void}
         */
        function syncLinkedMargin() {
            var verticalEnabled = !linkToggle.value;
            if (verticalField.enabled !== verticalEnabled) {
                /* ∧∨も合わせて切り替え、ディム表示を描き直す / toggle the stepper too and redraw its dimming */
                verticalField.enabled = verticalEnabled;
                verticalField.stepperGroup.enabled = verticalEnabled;
                redrawSteppersIn(verticalField.stepperGroup);
            }
            if (linkToggle.value) verticalField.text = horizontalField.text;
        }

        /**
         * 左右の入力が変わったときの処理
         * @returns {void}
         */
        function handleHorizontalMarginChanged() {
            syncLinkedMargin();
            onSettingsChanged();
        }

        /**
         * 上下の入力が変わったときの処理（連動ONのときは左右側で処理済み）
         * @returns {void}
         */
        function handleVerticalMarginChanged() {
            if (!linkToggle.value) onSettingsChanged();
        }

        horizontalField.onChanging = handleHorizontalMarginChanged;
        horizontalField.onChange = function () {
            normalizeMarginField(horizontalField);
            handleHorizontalMarginChanged();
        };
        verticalField.onChanging = handleVerticalMarginChanged;
        verticalField.onChange = function () {
            normalizeMarginField(verticalField);
            handleVerticalMarginChanged();
        };
        syncLinkedMargin();

        return {
            panel: marginPanel,
            linkToggle: linkToggle,
            getMarginInPoints: function () {
                var horizontalValue = readMarginField(horizontalField);
                var verticalValue = linkToggle.value ? horizontalValue : readMarginField(verticalField);
                return {
                    x: convertToPoints(horizontalValue, unitLabel),
                    y: convertToPoints(verticalValue, unitLabel)
                };
            },
            focusHorizontalField: function () {
                horizontalField.active = true;
                horizontalField.selection = [0, horizontalField.text.length];
            },
            bindEnterKey: function (btnOK) {
                var marginFields = [horizontalField, verticalField];
                for (var i = 0; i < marginFields.length; i++) {
                    marginFields[i].addEventListener("keydown", function (event) {
                        if (event.keyName !== "Enter" && event.keyName !== "Return") return;
                        btnOK.notify();
                        event.preventDefault();
                    });
                }
            }
        };
    }

    /**
     * プレビュー境界のチェックボックス行を構築する
     * @param {Window} parentContainer - 追加先のダイアログ
     * @param {Function} onSettingsChanged - 切り替え時に呼ぶ関数
     * @returns {void}
     */
    function buildBoundsOptionRow(parentContainer, onSettingsChanged) {
        var boundsOptionGroup = parentContainer.add("group");
        boundsOptionGroup.orientation = "row";
        boundsOptionGroup.alignment = ["left", "top"];
        boundsOptionGroup.margins = BOUNDS_OPTION_MARGINS;

        var boundsCheckbox = boundsOptionGroup.add("checkbox", undefined, getLabel("checkbox.useVisibleBounds"));
        boundsCheckbox.value = useVisibleBounds;
        boundsCheckbox.helpTip = getLabel("tooltip.useVisibleBounds");
        boundsCheckbox.onClick = function () {
            useVisibleBounds = boundsCheckbox.value;
            onSettingsChanged();
        };
    }

    /**
     * OK／キャンセルのボタン行を構築する
     * @param {Window} parentContainer - 追加先のダイアログ
     * @returns {Button} OKボタン
     */
    function buildDialogButtonRow(parentContainer) {
        var btnRowGroup = parentContainer.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.alignChildren = ["center", "center"];
        btnRowGroup.alignment = ["center", "bottom"];

        btnRowGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        return btnRowGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
    }

    /**
     * ダイアログのショートカットキー（1/2＝整列の基準、q〜c＝整列先）を登録する
     * @param {Window} targetDialog - 対象ダイアログ
     * @param {Object} baseControls - 整列の基準パネルのインターフェース
     * @param {Object} anchorControls - 整列先パネルのインターフェース
     * @param {Function} onSettingsChanged - 選択が変わったときに呼ぶ関数
     * @returns {void}
     */
    function registerShortcutKeys(targetDialog, baseControls, anchorControls, onSettingsChanged) {
        targetDialog.addEventListener("keydown", function (event) {
            /* テキスト入力中はショートカットを無効化 / Skip shortcuts while typing in edittext */
            if (event.target && event.target.type === "edittext") return;
            if (!event.keyName) return;

            var pressedKey = ("" + event.keyName).toLowerCase();
            if (pressedKey === "1" || pressedKey === "2") {
                baseControls.selectAlignmentBase(pressedKey === "2");
                event.preventDefault();
                onSettingsChanged();
                return;
            }

            if (!anchorControls.selectByShortcutKey(pressedKey)) return;
            event.preventDefault();
            onSettingsChanged();
        });
    }

    /**
     * ダイアログを構築・表示し、ライブプレビューで整列を反映する
     * @param {Document} doc - 対象ドキュメント
     * @returns {void}
     */
    function showAlignmentDialog(doc) {
        var alignDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        alignDialog.orientation = "column";
        alignDialog.alignChildren = ["fill", "top"];

        var unitLabel = getUnitInfo().label;
        var previewState = createPreviewState(doc.selection);
        initAnchorColors();

        var baseControls = buildAlignmentBasePanel(alignDialog, doc.artboards.length, handleSettingsChanged);

        var contentGroup = alignDialog.add("group");
        contentGroup.orientation = "row";
        contentGroup.alignChildren = ["fill", "top"];

        var anchorColumn = contentGroup.add("group");
        anchorColumn.orientation = "column";
        anchorColumn.alignChildren = ["fill", "top"];
        var anchorControls = buildAnchorPanel(anchorColumn, handleSettingsChanged);

        var marginColumn = contentGroup.add("group");
        marginColumn.orientation = "column";
        marginColumn.alignChildren = ["fill", "top"];
        marginColumn.alignment = ["right", "top"];
        var marginControls = buildMarginPanel(marginColumn, unitLabel, handleSettingsChanged);

        buildBoundsOptionRow(alignDialog, handleSettingsChanged);
        var btnOK = buildDialogButtonRow(alignDialog);
        alignDialog.defaultElement = btnOK;
        marginControls.bindEnterKey(btnOK);
        registerShortcutKeys(alignDialog, baseControls, anchorControls, handleSettingsChanged);

        /**
         * ダイアログの入力から整列設定を読み取る
         * @returns {Object} 整列設定 {anchorCode, useActiveArtboardAsReference, marginX, marginY}
         */
        function readSettingsFromDialog() {
            var anchorCode = anchorControls.getAnchorCode();
            var useActiveArtboardAsReference = baseControls.usesActiveArtboard();
            /* 中央整列時はマージンを使わない / Margins are unused for the center anchor */
            var marginIsAvailable = (anchorCode !== CENTER_ANCHOR_CODE);
            var marginPoints = marginIsAvailable ? marginControls.getMarginInPoints() : { x: 0, y: 0 };
            return {
                anchorCode: anchorCode,
                useActiveArtboardAsReference: useActiveArtboardAsReference,
                marginIsAvailable: marginIsAvailable,
                marginX: marginPoints.x,
                marginY: marginPoints.y
            };
        }

        /**
         * 入力が変わるたびにパネルの有効状態を同期し、プレビューを再適用する
         * @returns {void}
         */
        function handleSettingsChanged() {
            var alignSettings = readSettingsFromDialog();
            if (marginControls.panel.enabled !== alignSettings.marginIsAvailable) {
                marginControls.panel.enabled = alignSettings.marginIsAvailable;
                redrawSteppersIn(marginControls.panel); /* ∧∨のディム表示を切り替える / update the stepper dimming */
            }
            setLinkToggleEnabled(marginControls.linkToggle, alignSettings.marginIsAvailable);

            revertPreview(previewState);
            alignItems(doc, previewState.items, alignSettings, previewState);
            app.redraw();
        }

        alignDialog.onShow = function () {
            marginControls.focusHorizontalField();
        };

        /* 初期状態でプレビューを適用 / Apply the initial preview */
        handleSettingsChanged();

        prepareDialogWindow(alignDialog, SCRIPT_NAME);
        /* OKのときはプレビューの位置をそのまま確定 / OK keeps the previewed positions as the result */
        if (alignDialog.show() === 1) return;

        revertPreview(previewState);
        app.redraw();
    }

    // =========================================
    // メイン / Main
    // =========================================

    if (app.documents.length === 0) return;

    var doc = app.activeDocument;
    if (!doc.selection || doc.selection.length === 0) return;

    showAlignmentDialog(doc);

})();
