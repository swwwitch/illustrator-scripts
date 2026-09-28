#target illustrator
#targetengine "AddPageNumberFromTextSelectionEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択中のポイントテキストを雛形に、すべてのアートボードへページ番号を配置します。
接頭辞・接尾辞・ゼロ埋め・総ページ数表示に対応し、変更は即時プレビューされます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AddPageNumberFromTextSelection.md

note記事も参照してください。
https://note.com/dtp_tranist/n/ndc3d96ffc335

### Overview

Uses the selected point text as a template and places a page number on every artboard.
Prefix, suffix, zero padding and a total-pages display are supported, with an immediate preview.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AddPageNumberFromTextSelection.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AddPageNumberFromTextSelection";  /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v2.2.3";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-06-25";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-29";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AddPageNumberFromTextSelection.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AddPageNumberFromTextSelection.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/ndc3d96ffc335"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    // 連番テキストを配置する対象レイヤー名 / Layer that receives the page-number text
    var PAGENUMBER_LAYER_NAME = "_pagenumber";
    // プレビュー中の雛形を退避する一時レイヤー名 / Temp layer used to back up the template during preview
    var BACKUP_LAYER_NAME = "_pagenumber_preview";

    // =========================================
    // レイアウト / Layout
    // =========================================
    var DIALOG_OFFSET_X = 300;              /* ダイアログの表示位置：右(+)／左(-) / dialog offset to the right */
    var AFFIX_FIELD_CHARS = 10;             /* 接頭辞・接尾辞欄の文字数 / width of the prefix and suffix fields */
    var START_NUMBER_FIELD_CHARS = 6;       /* 開始番号欄の文字数 / width of the start number field */

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

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // キーボードショートカット（再利用パーツ） / Keyboard shortcuts (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内（uiLang の定義より後、ダイアログを作る関数より前）に貼る。
    //    識別子はすべて KEY_SHORTCUT_* / *KeyShortcut* の名前。uiLang はコピー先のものをそのまま使う
    // 2. コントロールをすべて作り、onClick を付けたあとで1回だけ呼ぶ（keydown はウィンドウに1つ）
    //      addKeyShortcuts(dialog, {
    //          "L": alignLeftRadio,                    … ラジオ：選んで onClick
    //          "P": previewCheckbox,                   … チェックボックス：反転して onClick
    //          "Shift+R": btnReset,                    … ボタン：onClick（無ければ notify）
    //          "G": function () { toggleGuides(); },   … 関数：呼ぶだけ
    //          "Escape": { target: function () { palette.close(); }, inFields: true }
    //      }, { numericFields: [widthInput, heightInput], afterKey: updatePreview });
    //    キーは keyName と同じ綴り（"A"〜"Z"・"1"・"Semicolon"・"Escape" など。大小文字は区別しない）。
    //    修飾キーは "Shift+" / "Alt+"（option）/ "Cmd+"（⌘、Windows は Ctrl）を前に付ける
    // 3. 修飾キーは完全一致。"R" は Shift・option・⌘ を押しながらでは効かない（⌘C などを横取りしない）。
    //    Shift＋R に別の動作を付けるときは "Shift+R" を並べる
    // 4. 入力欄（edittext）・ドロップダウン・リストにフォーカスがあるときは効かない（文字は普通に入る）。
    //    数値だけの欄で効かせたいときは options.numericFields に並べる（押した文字は欄に入らない）。
    //    入力中でも効かせたいキーは { target: …, inFields: true } にする（Esc で閉じる、option＋数字など）
    // 5. 無効・非表示のコントロールは、親のパネルやグループが無効なときも含めて何もしない
    //    （親を無効にしても子の enabled は true のまま、のため親までたどる）
    // 6. 関数の戻り値：false はこのキーを使わない（文字をそのまま通す）。コントロールを返すと、そのコントロールを
    //    押したことにする（向きによってラジオが変わるときなど）。それ以外は処理済み
    // 7. ツールチップへのキー表記は options.showInTip: true で「…（L）」「… (L)」を末尾に足す。
    //    LABELS の tooltip にすでにキーを書いてあるスクリプトでは付けない（同じキーが書いてあれば二重には足さない）
    // 8. 既存の keydown 処理（bindKeyboardShortcuts・addAlignKeyHandler など）と入力欄の focus／blur による抑止は消して、これに寄せる。
    //    ↑↓キー（StepperButtons の bindSteppedArrowKeys）はそのまま残す
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

    /* 入力中はショートカットを止めるコントロールの種類 / Control types that swallow keys while focused */
    var KEY_SHORTCUT_TYPING_TYPES = { edittext: true, dropdownlist: true, listbox: true };

    /* 修飾キーの並び順（キーの表記をそろえる）/ Canonical order of modifiers in a key spec */
    var KEY_SHORTCUT_MODIFIERS = ["SHIFT", "ALT", "CMD"];

    /* 修飾キーの別名 / Aliases accepted for the modifiers */
    var KEY_SHORTCUT_MODIFIER_ALIASES = {
        SHIFT: "SHIFT",
        ALT: "ALT", OPTION: "ALT", OPT: "ALT",
        CMD: "CMD", COMMAND: "CMD", META: "CMD", CTRL: "CMD", CONTROL: "CMD"
    };

    /**
     * キーの指定（"Shift+R" など）を、照合用の表記（"SHIFT+R"）にそろえる
     * @param {string} keySpec - キーの指定。修飾キーは "Shift+" / "Alt+" / "Cmd+" を前に付ける
     * @returns {string} 照合用の表記（大文字、修飾キーは SHIFT → ALT → CMD の順）
     */
    function normalizeKeyShortcutSpec(keySpec) {
        var specParts = String(keySpec).split("+");
        var baseKey = specParts.pop().toUpperCase();
        var modifierFlags = {};
        for (var i = 0; i < specParts.length; i++) {
            var modifierName = KEY_SHORTCUT_MODIFIER_ALIASES[specParts[i].toUpperCase()];
            if (modifierName) modifierFlags[modifierName] = true;
        }
        return buildKeyShortcutSpec(modifierFlags, baseKey);
    }

    /**
     * 修飾キーの状態とキー名から照合用の表記を組み立てる
     * @param {Object} modifierFlags - { SHIFT: true, ALT: true, CMD: true } のうち押されているもの
     * @param {string} baseKey - 大文字のキー名
     * @returns {string} 照合用の表記
     */
    function buildKeyShortcutSpec(modifierFlags, baseKey) {
        var specText = "";
        for (var i = 0; i < KEY_SHORTCUT_MODIFIERS.length; i++) {
            if (modifierFlags[KEY_SHORTCUT_MODIFIERS[i]]) specText += KEY_SHORTCUT_MODIFIERS[i] + "+";
        }
        return specText + baseKey;
    }

    /**
     * keydown イベントから照合用の表記を作る。修飾キーはイベントと keyboardState の両方を見る
     * @param {Object} keyEvent - keydown イベント
     * @returns {string} 照合用の表記。キー名が無いときは空文字
     */
    function readKeyShortcutSpec(keyEvent) {
        if (!keyEvent || !keyEvent.keyName) return "";
        var keyboardState = {};
        try { keyboardState = ScriptUI.environment.keyboardState; } catch (e) { }
        var modifierFlags = {
            SHIFT: !!(keyEvent.shiftKey || keyboardState.shiftKey),
            ALT: !!(keyEvent.altKey || keyboardState.altKey),
            CMD: !!(keyEvent.metaKey || keyEvent.ctrlKey || keyboardState.metaKey || keyboardState.ctrlKey)
        };
        return buildKeyShortcutSpec(modifierFlags, String(keyEvent.keyName).toUpperCase());
    }

    /**
     * コントロールが押せる状態か（自分と親がすべて有効で表示中か）を返す
     * @param {Object} control - コントロール
     * @returns {boolean} 押せるなら true
     */
    function isKeyShortcutControlUsable(control) {
        for (var node = control; node; node = node.parent) {
            if (node.enabled === false || node.visible === false) return false;
        }
        return true;
    }

    /**
     * キーを受けたコントロールが、文字を入力する欄か
     * @param {Object} focusedControl - イベントの発生元
     * @param {Object[]} numericFields - 数値だけの欄（ショートカットを効かせる）
     * @returns {boolean} 入力中としてショートカットを止めるなら true
     */
    function isKeyShortcutTypingTarget(focusedControl, numericFields) {
        if (!focusedControl || !KEY_SHORTCUT_TYPING_TYPES[focusedControl.type]) return false;
        for (var i = 0; i < numericFields.length; i++) {
            if (numericFields[i] === focusedControl) return false;
        }
        return true;
    }

    /**
     * コントロールをクリックしたときと同じ動作をする
     * ラジオは同じ親のラジオを外して選び、チェックボックスは反転してから onClick を呼ぶ
     * @param {Object} control - ラジオボタン・チェックボックス・ボタンなど
     * @returns {void}
     */
    function pressKeyShortcutControl(control) {
        if (control.type === "radiobutton") {
            /* 同じ親の直下だけが排他になるので、クリックと同じく兄弟を外す / Clear siblings like a click would */
            var siblings = control.parent ? control.parent.children : [];
            for (var i = 0; i < siblings.length; i++) {
                if (siblings[i] !== control && siblings[i].type === "radiobutton") siblings[i].value = false;
            }
            control.value = true;
        } else if (control.type === "checkbox") {
            control.value = !control.value;
        }
        if (typeof control.onClick === "function") {
            control.onClick.call(control);
        } else if (control.type === "button" && typeof control.notify === "function") {
            /* onClick の無い OK・キャンセルは notify で既定の動作（閉じる）を起こす / Let default buttons close the dialog */
            control.notify("onClick");
        }
    }

    /**
     * 1つのショートカットを実行する
     * @param {Object|Function} shortcutTarget - コントロール、または関数
     * @param {Object} keyEvent - keydown イベント
     * @returns {boolean} キーを使ったなら true（false なら文字をそのまま通す）
     */
    function runKeyShortcutTarget(shortcutTarget, keyEvent) {
        var targetControl = shortcutTarget;
        if (typeof shortcutTarget === "function") {
            var runResult = shortcutTarget(keyEvent);
            if (runResult === false || runResult === null) return false;
            if (!runResult || typeof runResult !== "object" || !runResult.type) return true;
            targetControl = runResult;
        }
        /* 無効なコントロールのキーも使ったことにして、数値欄へ文字を入れない / Consume the key even when disabled */
        if (isKeyShortcutControlUsable(targetControl)) pressKeyShortcutControl(targetControl);
        return true;
    }

    /**
     * キーの指定に修飾キーの表示名を当てて、ツールチップ用の表記にする
     * @param {string} normalizedSpec - 照合用の表記（"SHIFT+R" など）
     * @returns {string} 表示用の表記（"Shift+R" など）
     */
    function formatKeyShortcutLabel(normalizedSpec) {
        var isMac = ($.os.indexOf("Mac") === 0);
        var displayNames = { SHIFT: "Shift", ALT: isMac ? "Option" : "Alt", CMD: isMac ? "Cmd" : "Ctrl" };
        var specParts = normalizedSpec.split("+");
        var baseKey = specParts.pop();
        var labelText = "";
        for (var i = 0; i < specParts.length; i++) labelText += displayNames[specParts[i]] + "+";
        if (baseKey.length > 1) baseKey = baseKey.charAt(0) + baseKey.substring(1).toLowerCase();
        return labelText + baseKey;
    }

    /**
     * コントロールのツールチップの末尾にキーを足す（すでに書いてあれば足さない）
     * @param {Object} control - コントロール
     * @param {string} normalizedSpec - 照合用の表記
     * @returns {void}
     */
    function appendKeyShortcutToTip(control, normalizedSpec) {
        var keyLabel = formatKeyShortcutLabel(normalizedSpec);
        var currentTip = control.helpTip ? String(control.helpTip) : "";
        if (currentTip.indexOf("（" + keyLabel) >= 0 || currentTip.indexOf("(" + keyLabel) >= 0) return;
        var keySuffix = (uiLang === "ja") ? "（" + keyLabel + "）" : " (" + keyLabel + ")";
        control.helpTip = currentTip ? currentTip + keySuffix : keyLabel;
    }

    /**
     * ダイアログ・パレットに文字キーのショートカットを付ける
     * @param {Window} targetWindow - キーを受けるダイアログ・パレット
     * @param {Object} shortcutMap - { "L": ラジオ, "Shift+R": ボタン, "G": 関数, "Escape": { target: 関数, inFields: true } }
     * @param {Object} [shortcutOptions] - numericFields（数値だけの欄の配列）/ afterKey（キーを使ったあとに呼ぶ関数）/ showInTip（ツールチップにキーを足す）
     * @returns {Object} 照合用の表記 → { target, inFields } の表（テスト・デバッグ用）
     */
    function addKeyShortcuts(targetWindow, shortcutMap, shortcutOptions) {
        var shortcutSettings = shortcutOptions || {};
        var numericFields = shortcutSettings.numericFields || [];
        var bindingTable = {};

        for (var keySpec in shortcutMap) {
            if (!shortcutMap.hasOwnProperty(keySpec)) continue;
            var mapEntry = shortcutMap[keySpec];
            if (!mapEntry) continue;
            var isWrapped = (typeof mapEntry === "object" && !mapEntry.type && mapEntry.target);
            var normalizedSpec = normalizeKeyShortcutSpec(keySpec);
            bindingTable[normalizedSpec] = {
                target: isWrapped ? mapEntry.target : mapEntry,
                inFields: !!(isWrapped && mapEntry.inFields)
            };
            var tipControl = bindingTable[normalizedSpec].target;
            if (shortcutSettings.showInTip && typeof tipControl === "object" && tipControl.type) {
                appendKeyShortcutToTip(tipControl, normalizedSpec);
            }
        }

        /* キャプチャで受けて、数値欄に文字が入る前に止める / Capture phase keeps the letter out of numeric fields */
        targetWindow.addEventListener("keydown", function (keyEvent) {
            var binding = bindingTable[readKeyShortcutSpec(keyEvent)];
            if (!binding) return;
            if (!binding.inFields && isKeyShortcutTypingTarget(keyEvent.target, numericFields)) return;
            if (!runKeyShortcutTarget(binding.target, keyEvent)) return;
            if (keyEvent.preventDefault) keyEvent.preventDefault();
            if (typeof shortcutSettings.afterKey === "function") shortcutSettings.afterKey(keyEvent);
        }, true);

        return bindingTable;
    }

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // キーボードショートカット（再利用パーツ）ここまで / End of the reusable keyboard shortcuts
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    // UI 文字列（OK ボタンのラベルは非ローカライズ）/ UI strings (the OK button label is not localized)
    var LABELS = {
        dialog: {
            title: { ja: "ページ番号を一括配置", en: "Place Page Numbers" }
        },
        fieldLabel: {
            prefix: { ja: "接頭辞", en: "Prefix" },
            start: { ja: "開始番号", en: "Start number" },
            suffix: { ja: "接尾辞", en: "Suffix" }
        },
        checkbox: {
            zeroPad: { ja: "ゼロ埋め", en: "Zero padding" },
            showTotal: { ja: "総ページ数を表示", en: "Show total pages" }
        },
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" }
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
            prefix: {
                ja: "番号の前に付ける文字列（例：P.）",
                en: "Text placed before the number (e.g. P.)"
            },
            start: {
                ja: "先頭のアートボードに付ける番号。↑↓キーで増減、Shift+↑↓で10単位。",
                en: "Number for the first artboard. Up/Down to change, Shift+Up/Down for steps of 10."
            },
            suffix: {
                ja: "番号の後ろに付ける文字列（例：ページ）",
                en: "Text placed after the number (e.g. page)"
            },
            zeroPad: {
                ja: "総ページ数の桁数に合わせて0を補います（例：1 → 01）。Zキーで切り替え。",
                en: "Pads numbers with zeros to match the total (e.g. 1 → 01). Press Z to toggle."
            },
            showTotal: {
                ja: "「番号/総ページ数」の形式で表示します（例：3/12）。Aキーで切り替え。",
                en: "Shows the number as \"current/total\" (e.g. 3/12). Press A to toggle."
            },
            cancel: {
                ja: "プレビューを破棄して閉じます。",
                en: "Discard the preview and close."
            },
            ok: {
                ja: "プレビューの内容で確定し、" + PAGENUMBER_LAYER_NAME + " レイヤーに配置します。",
                en: "Commit the preview and place the text on the " + PAGENUMBER_LAYER_NAME + " layer."
            }
        },
        alert: {
            notNumber: {
                ja: "開始番号には数値を入力してください。",
                en: "Enter a number for the start number."
            },
            invalidSelection: {
                ja: "ページ番号の雛形となるポイントテキストを1つ選択してから実行してください。",
                en: "Select a single point text object to use as the page-number template, then run the script again."
            },
            commitFailed: {
                ja: "ページ番号を配置できませんでした。テキストや配置先レイヤーのロック状態を確認してください。",
                en: "Could not place the page numbers. Check whether the text or the destination layer is locked."
            }
        }
    };

    // =========================================
    // 安全実行ヘルパー / Safe Execution Helpers
    // =========================================

    /**
     * 関数を try/catch 内で実行し、action の戻り値を返す。例外時は onError(e) を呼ぶ（省略時は無視）。
     * 例外を握りつぶしてよい処理の共通ヘルパー / Shared helper for operations where ignored exceptions are acceptable
     * @param {Function} action - 実行する関数
     * @param {Function} [onError] - 例外時に呼ぶ関数
     * @returns {*} action の戻り値（例外時は undefined）
     */
    function tryCall(action, onError) {
        try {
            return action ? action() : undefined;
        } catch (e) {
            if (onError) onError(e);
        }
    }

    /**
     * プロパティ代入を例外無視で実行（ロック中・削除済みオブジェクトは代入で例外を出すため）/ Assign a property, ignoring any error (locked or deleted objects throw on assignment)
     * @param {object} targetObject - 代入先（null なら何もしない）
     * @param {string} propertyName - プロパティ名
     * @param {*} propertyValue - 代入する値
     * @returns {void}
     */
    function trySetProperty(targetObject, propertyName, propertyValue) {
        tryCall(function () { if (targetObject) targetObject[propertyName] = propertyValue; });
    }

    /**
     * 関数を実行し、例外時は errorLabel 付きでアラート表示
     * Run a function; on error show an alert prefixed with errorLabel
     * @param {string} errorLabel - アラートの先頭に付ける文字列
     * @param {Function} action - 実行する関数
     * @returns {void}
     */
    function runOrAlert(errorLabel, action) {
        tryCall(action, function (e) { alert(errorLabel + ": " + e); });
    }

    /**
     * 画面を安全に再描画
     * Redraw the screen safely
     * @returns {void}
     */
    function safeRedraw() {
        tryCall(function () { app.redraw(); });
    }

    // =========================================
    // レイヤー操作 / Layer Operations
    // =========================================

    /**
     * 指定名のレイヤーをサブレイヤーまで含めて探す（無ければ null）/ Find a layer by name, including sub-layers (null when it does not exist)
     * @param {Document|Layer} parentContainer - 探す範囲
     * @param {string} layerName - レイヤー名
     * @returns {Layer|null} 見つかったレイヤー
     */
    function findLayerByName(parentContainer, layerName) {
        for (var i = 0; i < parentContainer.layers.length; i++) {
            var childLayer = parentContainer.layers[i];
            if (childLayer.name === layerName) return childLayer;
            // サブレイヤーに同名があっても取りこぼさない / do not miss a nested layer with the same name
            var nestedLayer = findLayerByName(childLayer, layerName);
            if (nestedLayer) return nestedLayer;
        }
        return null;
    }

    /**
     * 指定名のレイヤーを取得、無ければ新規作成して返す
     * Get a layer by name, creating it if it does not exist
     * @param {Document} doc - 対象ドキュメント
     * @param {string} layerName - レイヤー名
     * @returns {Layer} 既存または新規のレイヤー
     */
    function getOrCreateLayer(doc, layerName) {
        var targetLayer = findLayerByName(doc, layerName);
        if (!targetLayer) {
            targetLayer = doc.layers.add();
            targetLayer.name = layerName;
        }
        return targetLayer;
    }

    /**
     * アイテムが属するレイヤーを返す（削除済みなら null）/ Return the layer owning the item, or null if the item is gone
     * @param {PageItem} pageItem - 対象のアイテム
     * @returns {Layer|null} 所属レイヤー
     */
    function getOwnerLayer(pageItem) {
        return tryCall(function () { return pageItem.layer; }) || null;
    }

    /**
     * 指定名のレイヤーを確実に削除（中身のロックを解除してから削除）/ Force-remove a layer by name (unlock its contents first, then remove)
     * @param {Document} doc - 対象ドキュメント
     * @param {string} layerName - レイヤー名
     * @returns {void}
     */
    function forceRemoveLayerByName(doc, layerName) {
        var targetLayer = findLayerByName(doc, layerName);
        if (!targetLayer) return;

        trySetProperty(targetLayer, 'locked', false);
        trySetProperty(targetLayer, 'visible', true);

        // ロックされた中身が削除を妨げるため、先にすべて解除 / locked contents block removal, so unlock them first
        for (var i = 0; i < targetLayer.pageItems.length; i++) trySetProperty(targetLayer.pageItems[i], 'locked', false);
        for (var j = 0; j < targetLayer.layers.length; j++) trySetProperty(targetLayer.layers[j], 'locked', false);

        tryCall(function () { targetLayer.remove(); });
    }

    // =========================================
    // Undo / プレビュー管理 / Undo & Preview Manager
    // =========================================

    /**
     * プレビュー編集をUndoステップとして積み、巻き戻し・確定を一括管理するクラス
     * Manages preview edits as undo steps for batch rollback or commit
     * @constructor
     */
    function PreviewManager() {
        this.undoDepth = 0; // プレビュー中に実行したアクション数 / number of preview actions executed

        // 変更操作を実行し、実際に変更があった場合だけ1ステップとしてカウント / Run an action and count it only when it actually changes the document
        this.runAsStep = function (previewAction) {
            var self = this;
            runOrAlert("Preview Error", function () {
                var changed = (typeof previewAction === "function") ? previewAction() : false;
                if (changed) {
                    self.undoDepth++;
                    safeRedraw();
                }
            });
        };

        // プレビュー分の変更をすべて取り消す / Roll back all preview steps
        this.rollback = function () {
            while (this.undoDepth > 0) {
                try {
                    app.undo();
                } catch (e) {
                    break;
                }
                this.undoDepth--;
            }
            safeRedraw();
        };

        // 確定：プレビュー分を全て取り消してから確定処理を1回実行 / Commit: undo all preview steps, then run the commit action once
        this.commit = function (commitAction) {
            this.rollback();
            if (typeof commitAction === "function") {
                runOrAlert("Commit Error", commitAction);
            }
        };
    }

    // =========================================
    // 選択・型判定 / Selection & Type Guards
    // =========================================

    /**
     * オブジェクトが TextFrame かどうかを判定（削除済み参照は false）/ Return true if the object is a TextFrame (a deleted reference yields false)
     * @param {*} candidate - 判定する値
     * @returns {boolean} TextFrame なら true
     */
    function isTextFrame(candidate) {
        try {
            return !!candidate && candidate.typename === "TextFrame";
        } catch (e) {
            return false;
        }
    }

    /**
     * 選択先頭が TextFrame ならそれを返す（無ければ null）/ Return the selected TextFrame, or null if none is selected
     * @returns {TextFrame|null} 選択中のテキスト
     */
    function getSelectedTextFrame() {
        if (app.documents.length === 0) return null;
        var currentSelection = app.selection;
        if (currentSelection && currentSelection.length > 0 && isTextFrame(currentSelection[0])) {
            return currentSelection[0];
        }
        return null;
    }

    // =========================================
    // _pagenumber レイヤーの状態管理 / Pagenumber Layer State
    // =========================================

    /**
     * 同じ親の中でのレイヤーの重ね順インデックスを返す（無ければ -1）/ Return the stacking-order index of a layer among its siblings (or -1)
     * @param {Layer} targetLayer - 対象のレイヤー
     * @returns {number} 重ね順のインデックス
     */
    function getLayerStackIndex(targetLayer) {
        var siblingLayers = targetLayer.parent.layers;
        for (var i = 0; i < siblingLayers.length; i++) {
            if (siblingLayers[i] === targetLayer) return i;
        }
        return -1;
    }

    /**
     * _pagenumber レイヤーの現在状態（ロック・表示・所属・重ね順）を記録
     * Capture the current state (lock, visibility, parent, stacking order) of the _pagenumber layer
     * @param {Layer} pagenumberLayer - _pagenumber レイヤー
     * @param {boolean} layerExisted - 実行前から存在したか
     * @returns {object} existed / locked / visible / parentContainer / neighborAbove
     */
    function capturePagenumberState(pagenumberLayer, layerExisted) {
        // 親コンテナ（ドキュメントまたは親レイヤー）ごと覚えておく / remember the parent container (document or parent layer) as well
        var parentContainer = pagenumberLayer.parent;
        var stackIndex = getLayerStackIndex(pagenumberLayer);
        return {
            existed: !!layerExisted,
            locked: pagenumberLayer.locked,
            visible: pagenumberLayer.visible,
            parentContainer: parentContainer,
            // ひとつ上（前面側）のレイヤーそのものを復元の基準として記録（同名レイヤーがあっても取り違えない）
            // remember the neighbor layer above itself as a restore anchor, so duplicate layer names cannot confuse it
            neighborAbove: (stackIndex > 0) ? parentContainer.layers[stackIndex - 1] : null
        };
    }

    /**
     * _pagenumber レイヤーを用意し、元状態を記録したうえで作業用に整える
     * Prepare the _pagenumber layer for work and capture its original state
     * @param {Document} doc - 対象ドキュメント
     * @returns {{layer: Layer, originalState: object}} レイヤーと元の状態
     */
    function setupPagenumberLayer(doc) {
        var layerExisted = !!findLayerByName(doc, PAGENUMBER_LAYER_NAME);
        var pagenumberLayer = getOrCreateLayer(doc, PAGENUMBER_LAYER_NAME);
        var originalState = capturePagenumberState(pagenumberLayer, layerExisted);

        // ロック解除・表示・最前面へ / unlock, show, and move to the top
        trySetProperty(pagenumberLayer, 'locked', false);
        trySetProperty(pagenumberLayer, 'visible', true);
        tryCall(function () { pagenumberLayer.move(doc, ElementPlacement.PLACEATBEGINNING); });

        return { layer: pagenumberLayer, originalState: originalState };
    }

    /**
     * capturePagenumberState で記録した状態へ _pagenumber レイヤーを復元
     * Restore the _pagenumber layer to the captured state
     * @param {Document} doc - 対象ドキュメント
     * @param {Layer} pagenumberLayer - _pagenumber レイヤー
     * @param {object} originalState - capturePagenumberState() の戻り値
     * @param {boolean} removeWhenAutoCreated - 自動作成したレイヤーなら削除する（キャンセル時）
     * @returns {void}
     */
    function restorePagenumberState(doc, pagenumberLayer, originalState, removeWhenAutoCreated) {
        if (!pagenumberLayer || !originalState) return;

        // キャンセル時のみ、元々存在しなかった _pagenumber を削除 / remove an auto-created _pagenumber only on Cancel
        if (!originalState.existed && removeWhenAutoCreated) {
            forceRemoveLayerByName(doc, PAGENUMBER_LAYER_NAME);
            return;
        }

        // 所属と重ね順を復元 / restore the parent container and the stacking order
        if (originalState.neighborAbove) {
            tryCall(function () { pagenumberLayer.move(originalState.neighborAbove, ElementPlacement.PLACEAFTER); });
        } else {
            tryCall(function () {
                pagenumberLayer.move(originalState.parentContainer || doc, ElementPlacement.PLACEATBEGINNING);
            });
        }

        // 表示・ロック状態を復元 / restore visibility & lock
        trySetProperty(pagenumberLayer, 'visible', originalState.visible);
        trySetProperty(pagenumberLayer, 'locked', originalState.locked);
    }

    // =========================================
    // アートボードとフレームの探索 / Artboard & Frame Lookup
    // =========================================

    /**
     * 座標 point が矩形 rect 内にあるか判定
     * Return true if the point is inside the rectangle
     * @param {number[]} point - 座標 [x, y]
     * @param {number[]} rect - 矩形 [左, 上, 右, 下]
     * @returns {boolean} 内側なら true
     */
    function isPointInRect(point, rect) {
        return point[0] >= rect[0] && point[0] <= rect[2] && point[1] <= rect[1] && point[1] >= rect[3];
    }

    /**
     * 指定座標が含まれるアートボードのインデックスを返す（無ければ -1）/ Return the index of the artboard containing the given point (or -1)
     * @param {Document} doc - 対象ドキュメント
     * @param {number[]} point - 座標 [x, y]
     * @returns {number} アートボードのインデックス
     */
    function getArtboardIndexByPosition(doc, point) {
        for (var i = 0; i < doc.artboards.length; i++) {
            if (isPointInRect(point, doc.artboards[i].artboardRect)) return i;
        }
        return -1;
    }

    /**
     * いずれかのアートボード上で最初に見つかった TextFrame を返す
     * Return the first TextFrame found on any artboard
     * @param {Document} doc - 対象ドキュメント
     * @param {Layer} targetLayer - 探すレイヤー
     * @returns {TextFrame|null} 見つかったテキスト
     */
    function findTextFrameOnAnyArtboard(doc, targetLayer) {
        for (var i = 0; i < targetLayer.textFrames.length; i++) {
            var textFrame = targetLayer.textFrames[i];
            if (getArtboardIndexByPosition(doc, textFrame.position) >= 0) return textFrame;
        }
        return null;
    }

    /**
     * TextFrame 群をアートボード順に並べた配列を返す（excludedFrame とアートボード外は除外）/ Return the TextFrames sorted by artboard order (excludedFrame and off-artboard frames are skipped)
     * @param {Document} doc - 対象ドキュメント
     * @param {TextFrames|TextFrame[]} textFrames - 並べるテキスト
     * @param {TextFrame|null} excludedFrame - 除外するテキスト
     * @returns {TextFrame[]} アートボード順のテキスト
     */
    function sortFramesByArtboard(doc, textFrames, excludedFrame) {
        var frameEntries = [];
        for (var i = 0; i < textFrames.length; i++) {
            var textFrame = textFrames[i];
            if (excludedFrame && textFrame === excludedFrame) continue;
            var artboardIndex = getArtboardIndexByPosition(doc, textFrame.position);
            // どのアートボードにも乗らないテキストは採番対象外 / text that sits on no artboard is not numbered
            if (artboardIndex < 0) continue;
            frameEntries.push({ frame: textFrame, artboardIndex: artboardIndex });
        }
        frameEntries.sort(function (entryA, entryB) { return entryA.artboardIndex - entryB.artboardIndex; });

        var sortedFrames = [];
        for (var j = 0; j < frameEntries.length; j++) sortedFrames.push(frameEntries[j].frame);
        return sortedFrames;
    }

    // =========================================
    // ページ番号テキストの生成・配置 / Page Number Generation & Placement
    // =========================================

    /**
     * 番号・接頭辞/接尾辞・ゼロ埋め・総ページ表示からページ番号文字列を生成
     * Build the page-number string from the number, prefix/suffix, zero padding, and the optional total
     * @param {number} pageNumber - 番号
     * @param {number} digitCount - ゼロ埋めの桁数
     * @param {object} formatOptions - prefix / suffix / zeroPad / showTotal
     * @param {number} totalPages - 総ページ数として表示する値
     * @returns {string} ページ番号の文字列
     */
    function buildPageNumberText(pageNumber, digitCount, formatOptions, totalPages) {
        var numberText = String(pageNumber);
        if (formatOptions.zeroPad && numberText.length < digitCount) {
            numberText = Array(digitCount - numberText.length + 1).join("0") + numberText; // ES3対応ゼロ埋め / ES3-safe zero pad
        }
        var pageNumberText = formatOptions.prefix + numberText + formatOptions.suffix;
        if (formatOptions.showTotal) pageNumberText += "/" + totalPages;
        return pageNumberText;
    }

    /**
     * レイヤー上のテキストをアートボード順に並べ、連番を流し込む
     * Sort the layer's text frames by artboard and write sequential numbers into them
     * @param {Document} doc - 対象ドキュメント
     * @param {Layer} targetLayer - テキストのあるレイヤー
     * @param {TextFrame|null} excludedFrame - 採番しないテキスト
     * @param {number} startNumber - 開始番号
     * @param {object} formatOptions - prefix / suffix / zeroPad / showTotal
     * @returns {void}
     */
    function numberFramesInOrder(doc, targetLayer, excludedFrame, startNumber, formatOptions) {
        var sortedFrames = sortFramesByArtboard(doc, targetLayer.textFrames, excludedFrame);
        var lastPageNumber = startNumber + doc.artboards.length - 1;
        var digitCount = String(lastPageNumber).length;
        for (var i = 0; i < sortedFrames.length; i++) {
            trySetProperty(sortedFrames[i], 'contents',
                buildPageNumberText(startNumber + i, digitCount, formatOptions, lastPageNumber));
        }
    }

    /**
     * 指定レイヤー上の TextFrame を keptFrame 以外すべて削除
     * Remove every TextFrame on the layer except keptFrame
     * @param {Layer} targetLayer - 対象のレイヤー
     * @param {TextFrame|null} keptFrame - 残すテキスト
     * @returns {void}
     */
    function removeOtherTextFrames(targetLayer, keptFrame) {
        var textFrames = targetLayer.textFrames;
        for (var i = textFrames.length - 1; i >= 0; i--) {
            var textFrame = textFrames[i];
            if (textFrame === keptFrame) continue;
            trySetProperty(textFrame, 'locked', false);
            tryCall(function () { textFrame.remove(); });
        }
    }

    /**
     * 雛形テキストをカットし、全アートボードへ貼り付ける（プレビューと確定で共通）。成功したら true
     * Cut the given text and paste it onto every artboard (shared by preview and commit); returns true on success
     * @param {Document} doc - 対象ドキュメント
     * @param {TextFrame} textFrame - 雛形テキスト
     * @param {Layer} pasteLayer - 貼り付け先レイヤー（null なら元のレイヤー）
     * @param {Function} [beforePaste] - 貼り付け直前に呼ぶ関数
     * @returns {boolean} 成功したら true
     */
    function cutAndPasteToAllArtboards(doc, textFrame, pasteLayer, beforePaste) {
        // 対象と所属レイヤーを一時的にロック解除＆可視化 / temporarily unlock & show the target and its layer
        var sourceLayer = getOwnerLayer(textFrame);
        trySetProperty(textFrame, 'locked', false);
        trySetProperty(sourceLayer, 'locked', false);
        trySetProperty(sourceLayer, 'visible', true);

        // 対象が乗るアートボードをアクティブ化 / activate the artboard the target sits on
        var sourceArtboardIndex = getArtboardIndexByPosition(doc, textFrame.position);
        if (sourceArtboardIndex >= 0) doc.artboards.setActiveArtboardIndex(sourceArtboardIndex);

        // 選択→カット / select -> cut
        app.selection = null;
        trySetProperty(textFrame, 'selected', true);
        var cutSucceeded = tryCall(function () {
            app.cut();
            return true;
        }) === true;

        // カットできていない場合、この先へ進むと既存テキストを消すだけになるため中断
        // Bail out when the cut failed: continuing would only delete the existing text without pasting anything back
        if (!cutSucceeded) return false;

        // 貼り付け直前の後始末（既存テキストの一掃など）/ cleanup right before pasting (e.g. clearing existing text)
        if (beforePaste) beforePaste();

        // 貼り付け先レイヤーをアクティブにして全アートボードへ貼り付け / activate the destination layer, then paste onto all artboards
        var destinationLayer = pasteLayer || sourceLayer;
        if (destinationLayer) trySetProperty(doc, 'activeLayer', destinationLayer);
        return tryCall(function () {
            app.executeMenuCommand('pasteInAllArtboard');
            return true;
        }) === true;
    }

    /**
     * 雛形テキストを開始番号で初期化し、全アートボードへ複製（所属レイヤーの状態は元へ戻す）/ Seed the template text with the start number and duplicate it across all artboards, restoring its layer state afterwards
     * @param {Document} doc - 対象ドキュメント
     * @param {TextFrame} templateText - 雛形テキスト
     * @param {number} startNumber - 開始番号
     * @returns {boolean} 成功したら true
     */
    function seedAndPasteToAllArtboards(doc, templateText, startNumber) {
        trySetProperty(templateText, 'contents', String(startNumber));

        // 所属レイヤーの一時状態を退避 / back up the source layer's state
        var sourceLayer = getOwnerLayer(templateText);
        var originalLocked = sourceLayer ? sourceLayer.locked : null;
        var originalVisible = sourceLayer ? sourceLayer.visible : null;

        var pasteSucceeded = cutAndPasteToAllArtboards(doc, templateText, sourceLayer);

        // レイヤーの一時状態を元へ戻す / restore the layer's temporary state
        if (sourceLayer) {
            trySetProperty(sourceLayer, 'locked', originalLocked);
            trySetProperty(sourceLayer, 'visible', originalVisible);
        }
        return pasteSucceeded;
    }

    // =========================================
    // ライブプレビュー / Live Preview
    // =========================================

    /**
     * 雛形を退避レイヤーへ非表示コピーする（キャンセル時の復元用。常に最新の1つだけ保持）/ Copy the template onto a hidden backup layer for restoring on Cancel, keeping only the latest copy
     * @param {Document} doc - 対象ドキュメント
     * @param {TextFrame} templateText - 雛形テキスト
     * @returns {void}
     */
    function backupTemplateText(doc, templateText) {
        var backupLayer = getOrCreateLayer(doc, BACKUP_LAYER_NAME);
        backupLayer.visible = false;
        backupLayer.locked = false;

        // 前回の退避が残っていると、キャンセル時に雛形が重複して復元されるため先に破棄
        // Leftover backups would be restored on top of each other on Cancel, so discard them first
        for (var i = backupLayer.pageItems.length - 1; i >= 0; i--) {
            var staleItem = backupLayer.pageItems[i];
            trySetProperty(staleItem, 'locked', false);
            tryCall(function () { staleItem.remove(); });
        }

        var backupText = tryCall(function () {
            return templateText.duplicate(backupLayer, ElementPlacement.PLACEATBEGINNING);
        });
        trySetProperty(backupText, 'visible', false);
        trySetProperty(backupText, 'locked', true);
    }

    /**
     * 雛形を退避しつつ、全アートボードへクリーンに複製し直す
     * Back up the template, then cleanly re-duplicate it across every artboard
     * @param {Document} doc - 対象ドキュメント
     * @param {Layer} pagenumberLayer - _pagenumber レイヤー
     * @param {TextFrame} templateText - 雛形テキスト
     * @returns {boolean} 成功したら true
     */
    function rebuildFramesAcrossArtboards(doc, pagenumberLayer, templateText) {
        backupTemplateText(doc, templateText);
        var pasteSucceeded = cutAndPasteToAllArtboards(doc, templateText, pagenumberLayer, function () {
            // 貼り付け前に既存のページ番号を一掃 / clear the existing page numbers before pasting
            removeOtherTextFrames(pagenumberLayer, null);
        });

        // 失敗時は退避レイヤーごと破棄して、中途半端な状態を残さない / on failure, drop the backup layer so no half-finished state remains
        if (!pasteSucceeded) forceRemoveLayerByName(doc, BACKUP_LAYER_NAME);
        return pasteSucceeded;
    }

    /**
     * 選択テキスト（無ければレイヤー上の先頭テキスト）を雛形に、全アートボードへ連番をプレビュー
     * Render a sequential-numbering preview on every artboard, using the selected text (or the first text on the layer) as a template
     * @param {Document} doc - 対象ドキュメント
     * @param {string} layerName - ページ番号のレイヤー名
     * @param {number} startNumber - 開始番号
     * @param {object} formatOptions - prefix / suffix / zeroPad / showTotal
     * @returns {boolean} ドキュメントを変更したら true
     */
    function updatePreview(doc, layerName, startNumber, formatOptions) {
        if (!doc || isNaN(startNumber)) return false;
        var pagenumberLayer = findLayerByName(doc, layerName);
        if (!pagenumberLayer) return false;

        // 雛形を決定：選択 → アートボード上の先頭テキスト → レイヤー上の先頭テキスト
        // pick the template: selection -> first text on an artboard -> first text on the layer
        var templateText = getSelectedTextFrame() ||
            sortFramesByArtboard(doc, pagenumberLayer.textFrames, null)[0] ||
            pagenumberLayer.textFrames[0];
        if (!templateText) return false;

        if (!rebuildFramesAcrossArtboards(doc, pagenumberLayer, templateText)) return false;
        numberFramesInOrder(doc, pagenumberLayer, null, startNumber, formatOptions);

        // 再描画は呼び出し元（PreviewManager）で1回だけ行い、Undoの区切りを1プレビュー＝1ステップに保つ
        // The caller (PreviewManager) redraws once, keeping the undo boundary at one step per preview pass
        return true;
    }

    /**
     * 退避レイヤーに残ったテキストを _pagenumber へ戻し、退避レイヤーを削除
     * Move any text left on the backup layer back to _pagenumber, then remove the backup layer
     * @param {Document} doc - 対象ドキュメント
     * @returns {void}
     */
    function restorePreviewBackupOnCancel(doc) {
        var backupLayer = findLayerByName(doc, BACKUP_LAYER_NAME);
        if (!backupLayer) return;

        var pagenumberLayer = getOrCreateLayer(doc, PAGENUMBER_LAYER_NAME);
        removeOtherTextFrames(pagenumberLayer, null);
        backupLayer.locked = false;
        backupLayer.visible = true;

        for (var i = backupLayer.pageItems.length - 1; i >= 0; i--) {
            var backupItem = backupLayer.pageItems[i];
            trySetProperty(backupItem, 'locked', false);
            tryCall(function () { backupItem.move(pagenumberLayer, ElementPlacement.PLACEATBEGINNING); });
            trySetProperty(backupItem, 'visible', true);
        }
        forceRemoveLayerByName(doc, BACKUP_LAYER_NAME);
        safeRedraw();
    }

    // =========================================
    // UI 構築 / UI Construction
    // =========================================

    /**
     * 親に縦並びカラム（group）を追加して返す
     * @param {Group} parentGroup - 追加先
     * @param {string} [childAlignment] - 子の揃え（省略時は "left"）
     * @returns {Group} 追加したカラム
     */
    function addColumnGroup(parentGroup, childAlignment) {
        var columnGroup = parentGroup.add("group");
        columnGroup.orientation = "column";
        columnGroup.alignChildren = childAlignment || "left";
        return columnGroup;
    }

    /**
     * ラベル付き入力欄を追加し、入力欄（edittext）を返す
     * @param {Group} parentGroup - 追加先
     * @param {string} captionText - 入力欄の上に出す項目名
     * @param {string} initialValue - 入力欄の初期値
     * @param {number} characterWidth - 入力欄の文字数
     * @param {string} tooltipText - 項目名と入力欄に付ける tooltip
     * @param {Object} [stepOptions] - 数値欄のときに指定する∧∨の増減設定（integer / min / onStep など）。省略時は∧∨を付けない
     * @returns {EditText} 追加した入力欄
     */
    function addLabeledEditText(parentGroup, captionText, initialValue, characterWidth, tooltipText, stepOptions) {
        var captionLabel = parentGroup.add("statictext", undefined, captionText);
        var inputField;
        if (stepOptions) {
            /* ∧∨と入力欄は隙間0で突き合わせる。↑↓キーも∧∨と同じ処理で増減する
               Butt the stepper against the field; the arrow keys share its logic */
            var stepperInputGroup = parentGroup.add("group");
            stepperInputGroup.orientation = "row";
            stepperInputGroup.alignChildren = ["left", "center"];
            stepperInputGroup.spacing = 0;
            stepperInputGroup.margins = 0;
            var stepperGroup = addStepper(stepperInputGroup, function () { return inputField; }, stepOptions);
            inputField = stepperInputGroup.add("edittext", undefined, initialValue);
            bindSteppedArrowKeys(inputField, stepperGroup);
        } else {
            inputField = parentGroup.add("edittext", undefined, initialValue);
        }
        inputField.characters = characterWidth;
        // ラベル・入力欄のどちらにマウスを乗せても説明が出るようにする / show the hint from both the caption and the field
        captionLabel.helpTip = tooltipText;
        inputField.helpTip = tooltipText;
        return inputField;
    }

    /**
     * tooltip 付きのチェックボックスを追加して返す
     * @param {Group} parentGroup - 追加先
     * @param {string} checkboxText - チェックボックスの文言
     * @param {string} tooltipText - tooltip
     * @returns {Checkbox} 追加したチェックボックス
     */
    function addCheckbox(parentGroup, checkboxText, tooltipText) {
        var optionCheckbox = parentGroup.add("checkbox", undefined, checkboxText);
        optionCheckbox.helpTip = tooltipText;
        return optionCheckbox;
    }

    /**
     * ダイアログと各UIコントロールを生成し、参照をまとめて返す
     * @returns {object} pageNumberDialog と各入力欄・チェックボックス・ボタン
     */
    function buildPageNumberDialog() {
        // タイトルバーにはバージョンを併記 / show the version in the title bar
        var pageNumberDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        pageNumberDialog.orientation = "column";
        pageNumberDialog.alignChildren = "left";

        // 3カラムレイアウト / 3-column layout
        var columnsGroup = pageNumberDialog.add("group");
        columnsGroup.orientation = "row";
        columnsGroup.alignChildren = "top";

        // 左カラム: 接頭辞 / left column: prefix
        var prefixColumn = addColumnGroup(columnsGroup);
        var prefixField = addLabeledEditText(prefixColumn, getLabel("fieldLabel.prefix"), "", AFFIX_FIELD_CHARS,
            getLabel("tooltip.prefix"));

        // 中央カラム: 開始番号 + ゼロ埋め / center column: start number + zero pad
        var startNumberColumn = addColumnGroup(columnsGroup);
        var startNumberField = addLabeledEditText(startNumberColumn, getLabel("fieldLabel.start"), "1", START_NUMBER_FIELD_CHARS,
            getLabel("tooltip.start"), {
                integer: true,
                min: 0,
                /* ∧∨・↑↓キーで変えたら、確定時と同じくプレビューを更新する / refresh the preview as on commit */
                onStep: function (numberInput) {
                    if (typeof numberInput.onChange === "function") numberInput.onChange();
                }
            });
        var zeroPadCheckbox = addCheckbox(startNumberColumn, getLabel("checkbox.zeroPad"),
            getLabel("tooltip.zeroPad"));

        // 右カラム: 接尾辞 + 総ページ表示 / right column: suffix + show-total
        var suffixColumn = addColumnGroup(columnsGroup);
        var suffixField = addLabeledEditText(suffixColumn, getLabel("fieldLabel.suffix"), "", AFFIX_FIELD_CHARS,
            getLabel("tooltip.suffix"));
        var totalPageCheckbox = addCheckbox(suffixColumn, getLabel("checkbox.showTotal"),
            getLabel("tooltip.showTotal"));

        /* ボタン行（右寄せ：キャンセル → OK）/ Button row (right: Cancel, OK) */
        var buttonRow = addButtonRow(pageNumberDialog);
        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        btnCancel.helpTip = getLabel("tooltip.cancel");

        var btnOK = buttonRow.rightGroup.add("button", undefined, "OK", { name: "ok" });
        btnOK.helpTip = getLabel("tooltip.ok");

        // 初めて開くときの表示位置 / position on first open
        pageNumberDialog.onShow = function () {
            pageNumberDialog.location = [pageNumberDialog.location[0] + DIALOG_OFFSET_X, pageNumberDialog.location[1]];
        };

        return {
            pageNumberDialog: pageNumberDialog,
            prefixField: prefixField,
            startNumberField: startNumberField,
            zeroPadCheckbox: zeroPadCheckbox,
            suffixField: suffixField,
            totalPageCheckbox: totalPageCheckbox,
            btnCancel: btnCancel,
            btnOK: btnOK
        };
    }

    /**
     * 開始番号欄を整数として読む
     * @param {object} dialogControls - buildPageNumberDialog() の戻り値
     * @returns {number} 開始番号（数値でなければ NaN）
     */
    function readStartNumber(dialogControls) {
        return parseInt(dialogControls.startNumberField.text, 10);
    }

    /**
     * 現在の入力値を書式オプションとしてまとめる
     * @param {object} dialogControls - buildPageNumberDialog() の戻り値
     * @returns {{prefix: string, suffix: string, zeroPad: boolean, showTotal: boolean}} 書式オプション
     */
    function readFormatOptions(dialogControls) {
        return {
            prefix: dialogControls.prefixField.text || "",
            suffix: dialogControls.suffixField.text || "",
            zeroPad: !!dialogControls.zeroPadCheckbox.value,
            showTotal: !!dialogControls.totalPageCheckbox.value
        };
    }

    // =========================================
    // 確定処理 / Commit
    // =========================================

    /**
     * OK確定時の雛形テキストを取得（優先候補 → 現在の選択 → レイヤー上の既存テキスト）
     * @param {Document} doc - 対象ドキュメント
     * @param {Layer} pagenumberLayer - _pagenumber レイヤー
     * @param {TextFrame} preferredText - 最優先の候補（実行時に選択していたテキスト）
     * @returns {TextFrame|null} 雛形テキスト
     */
    function resolveTemplateTextForCommit(doc, pagenumberLayer, preferredText) {
        var templateText = isTextFrame(preferredText) ? preferredText : getSelectedTextFrame();
        if (!isTextFrame(templateText)) {
            templateText = findTextFrameOnAnyArtboard(doc, pagenumberLayer);
        }
        return isTextFrame(templateText) ? templateText : null;
    }

    /**
     * 雛形テキストを _pagenumber 上へ移し、他のテキストを除去
     * @param {Layer} pagenumberLayer - _pagenumber レイヤー
     * @param {TextFrame} templateText - 雛形テキスト
     * @returns {boolean} 移せたら true
     */
    function moveTemplateTextToPagenumberLayer(pagenumberLayer, templateText) {
        if (!isTextFrame(templateText)) return false;
        if (templateText.layer.name !== PAGENUMBER_LAYER_NAME) {
            tryCall(function () { templateText.move(pagenumberLayer, ElementPlacement.PLACEATBEGINNING); });
        }
        if (templateText.layer.name !== PAGENUMBER_LAYER_NAME) return false;
        removeOtherTextFrames(pagenumberLayer, templateText);
        return true;
    }

    /**
     * 確定用の連番を全アートボードへ適用
     * @param {Document} doc - 対象ドキュメント
     * @param {Layer} pagenumberLayer - _pagenumber レイヤー
     * @param {TextFrame} templateText - 雛形テキスト
     * @param {number} startNumber - 開始番号
     * @param {object} formatOptions - readFormatOptions() の戻り値
     * @returns {boolean} 配置できたら true
     */
    function applyNumberingToAllArtboards(doc, pagenumberLayer, templateText, startNumber, formatOptions) {
        // 複製に失敗した場合は採番せず、雛形をそのまま残す / when the duplication fails, leave the template as it is
        if (!seedAndPasteToAllArtboards(doc, templateText, startNumber)) return false;
        numberFramesInOrder(doc, pagenumberLayer, templateText, startNumber, formatOptions);
        return true;
    }

    /**
     * OK確定時の本処理：雛形テキストを _pagenumber へ移し、全アートボードへ連番を確定配置
     * @param {Document} doc - 対象ドキュメント
     * @param {object} dialogControls - buildPageNumberDialog() の戻り値
     * @param {TextFrame} originalTemplateText - 実行時に選択していたテキスト
     * @param {object} pagenumberSetup - setupPagenumberLayer() の戻り値（元の状態へ戻すため）
     * @returns {void}
     */
    function commitPageNumbers(doc, dialogControls, originalTemplateText, pagenumberSetup) {
        var startNumber = readStartNumber(dialogControls);
        if (isNaN(startNumber)) {
            alert(getLabel("alert.notNumber"));
            return;
        }

        // プレビュー用の退避レイヤーを破棄 / discard the preview backup layer
        forceRemoveLayerByName(doc, BACKUP_LAYER_NAME);

        var pagenumberLayer = getOrCreateLayer(doc, PAGENUMBER_LAYER_NAME);
        var templateText = resolveTemplateTextForCommit(doc, pagenumberLayer, originalTemplateText);
        if (!templateText || !moveTemplateTextToPagenumberLayer(pagenumberLayer, templateText)) {
            alert(getLabel("alert.invalidSelection"));
            return;
        }

        var placed = applyNumberingToAllArtboards(doc, pagenumberLayer, templateText, startNumber, readFormatOptions(dialogControls));
        safeRedraw();
        restorePagenumberState(doc, pagenumberSetup.layer, pagenumberSetup.originalState, false);
        if (!placed) alert(getLabel("alert.commitFailed"));
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * ダイアログを表示し、選択テキストを雛形に全アートボードへページ番号を配置
     * @returns {void}
     */
    function main() {
        // テキスト未選択なら終了 / Exit if no text is selected
        var originalTemplateText = getSelectedTextFrame();
        if (!originalTemplateText) {
            alert(getLabel("alert.invalidSelection"));
            return;
        }

        var doc = app.activeDocument;
        var pagenumberSetup = setupPagenumberLayer(doc);
        var dialogControls = buildPageNumberDialog();
        var previewManager = new PreviewManager();

        /* 現在の入力値でライブプレビューを更新（前回分を巻き戻し、1ステップとして再実行）
           Refresh the live preview with current input values (roll back the previous one, run as a single step) */
        function refreshPreview() {
            var startNumber = readStartNumber(dialogControls);
            if (isNaN(startNumber)) return;

            var formatOptions = readFormatOptions(dialogControls);
            previewManager.rollback();
            previewManager.runAsStep(function () {
                return updatePreview(doc, PAGENUMBER_LAYER_NAME, startNumber, formatOptions);
            });
        }

        // 入力確定（Tabやフォーカス移動）でプレビューを更新。onChanging は1文字ごとに全アートボードを組み直すため使わない
        // Refresh on commit of the field (Tab or focus change); onChanging would rebuild every artboard on each keystroke
        dialogControls.prefixField.onChange = refreshPreview;
        dialogControls.suffixField.onChange = refreshPreview;
        dialogControls.startNumberField.onChange = refreshPreview;

        // チェックボックスのON/OFFでプレビューを更新（キー操作での切り替えもこの onClick を通る）
        // Update the preview when a checkbox is toggled (key shortcuts go through this onClick too)
        dialogControls.zeroPadCheckbox.onClick = refreshPreview;
        dialogControls.totalPageCheckbox.onClick = refreshPreview;

        // Zキーでゼロ埋め、Aキーで総ページ表示をトグル（接頭辞・接尾辞の欄では文字入力を優先）
        // Z toggles zero-pad, A toggles show-total; typing in the prefix / suffix fields takes precedence
        addKeyShortcuts(dialogControls.pageNumberDialog, {
            "Z": dialogControls.zeroPadCheckbox,
            "A": dialogControls.totalPageCheckbox
        }, { numericFields: [dialogControls.startNumberField] });

        // OK：プレビュー分を全Undoしてから確定処理を1回だけ実行 / OK: undo all preview steps, then run the commit action once
        dialogControls.btnOK.onClick = function () {
            previewManager.commit(function () {
                commitPageNumbers(doc, dialogControls, originalTemplateText, pagenumberSetup);
            });
            forceRemoveLayerByName(doc, BACKUP_LAYER_NAME);
            dialogControls.pageNumberDialog.close(1);
        };

        // キャンセル：プレビューを巻き戻し、退避テキストと _pagenumber 状態を復元 / Cancel: roll back the preview, restore the backed-up text and the _pagenumber state
        dialogControls.btnCancel.onClick = function () {
            previewManager.rollback();
            restorePreviewBackupOnCancel(doc);
            restorePagenumberState(doc, pagenumberSetup.layer, pagenumberSetup.originalState, true);
            dialogControls.pageNumberDialog.close(0);
        };

        dialogControls.startNumberField.active = true;

        // 初回プレビュー / first preview pass
        refreshPreview();

        prepareDialogWindow(dialogControls.pageNumberDialog, SCRIPT_NAME);
        dialogControls.pageNumberDialog.show();
    }

    main();

})();
