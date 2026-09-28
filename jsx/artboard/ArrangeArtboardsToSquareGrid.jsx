#target illustrator
#targetengine "ArrangeArtboardsToSquareGridEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

複数のアートボードを、全体の外形ができるだけ正方形に近づく行列で再配置します。
各アートボード内のアートワークも一緒に移動し、グリッドはカンバス中央に配置します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ArrangeArtboardsToSquareGrid.md

### Overview

Re-lays out every artboard so the whole grid's outline is as close to a square as possible.
Each artboard's artwork moves with it, and the grid is centered on the canvas.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ArrangeArtboardsToSquareGrid.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ArrangeArtboardsToSquareGrid";  /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-28";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ArrangeArtboardsToSquareGrid.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ArrangeArtboardsToSquareGrid.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // レイアウト / Layout
    // =========================================
    var DIALOG_MARGINS = 16;                 /* ダイアログ外周の余白 / dialog margin */
    var DIALOG_SPACING = 12;                 /* ダイアログ内の要素間隔 / dialog spacing */
    var PANEL_MARGINS = 16;                  /* 「整列設定」パネルの余白 / settings panel margin */
    var PANEL_SPACING = 10;                  /* 「整列設定」パネル内の要素間隔 / settings panel spacing */
    var NUMBER_FIELD_CHARS = 4;              /* 間隔・列数の入力欄の文字数 / width of the gap and column fields */
    var AUTO_BUTTON_HEIGHT = 22;             /* ［自動］ボタンの高さ / height of the Auto button */
    var PREVIEW_TEXT_SIZE = [330, 36];       /* 行列・外形サイズの表示欄 / size of the grid summary text */

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

    var LABELS = {
        dialog: {
            title: { ja: "アートボードを正方形に整列", en: "Arrange Artboards to Square" }
        },
        panel: {
            settings: { ja: "整列設定", en: "Layout settings" }
        },
        fieldLabel: {
            artboardCount: { ja: "アートボード数", en: "Artboards" },
            gap: { ja: "間隔", en: "Gap" },
            columns: { ja: "列数", en: "Columns" }
        },
        button: {
            auto: { ja: "自動", en: "Auto" },
            ok: { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        status: {
            recommend: { ja: "推奨", en: "Recommended" },
            columnUnit: { ja: "列", en: "col" },
            rowUnit: { ja: "行", en: "row" }
        },
        tooltip: {
            gap: { ja: "アートボードどうしのあいだにあける間隔です。", en: "Space left between neighbouring artboards." },
            columns: {
                ja: "横に並べるアートボードの数です。行数はこの値から決まります。",
                en: "How many artboards to place in a row. The row count follows from this."
            },
            auto: {
                ja: "アートボード数から、できるだけ正方形に近くなる列数を入れ直します。",
                en: "Fills in the column count that comes closest to a square arrangement."
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
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            needTwo: { ja: "アートボードが 2 つ以上必要です。", en: "At least two artboards are required." },
            badColumns: { ja: "列数は 1 以上の整数で入力してください。", en: "Enter the column count as an integer of 1 or more." },
            badGap: { ja: "間隔は 0 以上の数値で入力してください。", en: "Enter the gap as a number of 0 or more." },
            done: { ja: "整列が完了しました。", en: "Artboards arranged." },
            moveFailed: { ja: " 個のオブジェクトは移動できませんでした（ロック等）。", en: " object(s) could not be moved (locked, etc.)." }
        }
    };

    // =========================================
    // カンバスとグリッド / Canvas & grid
    // =========================================

    /**
     * Illustrator の最大カンバス範囲を取得する（一時レイヤーで原点を測定）
     * Get Illustrator's max canvas bounds (measures the origin via a temp layer)
     * @returns {number[]} [left, top, right, bottom]
     */
    function getCanvasBounds() {
        var CANVAS_MAX_SIZE = 16383;
        var targetDoc = app.activeDocument;
        var wasModified = targetDoc.modified; // 計測前の変更フラグを退避 / remember the modified flag before measuring
        var tempLayer = targetDoc.layers.add();
        var tempTextFrame = tempLayer.textFrames.add();
        var canvasLeft = tempTextFrame.matrix.mValueTX;
        var canvasTop = tempTextFrame.matrix.mValueTY;
        tempLayer.remove();
        targetDoc.modified = wasModified; // 一時レイヤー追加で立った変更フラグを元に戻す / restore the modified flag
        return [canvasLeft, canvasTop, canvasLeft + CANVAS_MAX_SIZE, canvasTop - CANVAS_MAX_SIZE];
    }

    /**
     * 列数からグリッドの行数と外形サイズを求める
     * @param {number} itemCount - アートボード数
     * @param {number} columnCount - 列数
     * @param {number} cellWidth - セルの幅（pt）
     * @param {number} cellHeight - セルの高さ（pt）
     * @param {number} gap - 間隔（pt）
     * @returns {{rows: number, occupiedColumns: number, gridWidth: number, gridHeight: number}} 行数・実際に埋まる列数・外形サイズ
     */
    function computeGridSize(itemCount, columnCount, cellWidth, cellHeight, gap) {
        var rowCount = Math.ceil(itemCount / columnCount);
        var occupiedColumns = (columnCount < itemCount) ? columnCount : itemCount;
        return {
            rows: rowCount,
            occupiedColumns: occupiedColumns,
            gridWidth: occupiedColumns * cellWidth + (occupiedColumns - 1) * gap,
            gridHeight: rowCount * cellHeight + (rowCount - 1) * gap
        };
    }

    /**
     * グリッド全体をカンバスの天地・左右中央へ配置する原点を算出する
     * Compute the top-left origin that centers the whole artboard grid on the canvas.
     * @param {number[]} canvasBounds - カンバスの範囲 [left, top, right, bottom]
     * @param {number} cellWidth - セルの幅（pt）
     * @param {number} cellHeight - セルの高さ（pt）
     * @param {number} gap - 間隔（pt）
     * @param {number} columnCount - 列数
     * @param {number} itemCount - アートボード数
     * @returns {object} { left, top, cols, rows, gridWidth, gridHeight }
     */
    function computeCenteredGridOrigin(canvasBounds, cellWidth, cellHeight, gap, columnCount, itemCount) {
        var gridSize = computeGridSize(itemCount, columnCount, cellWidth, cellHeight, gap);
        var canvasWidth = canvasBounds[2] - canvasBounds[0];
        var canvasHeight = canvasBounds[1] - canvasBounds[3];
        var leftMargin = Math.round((canvasWidth - gridSize.gridWidth) / 2);
        var topMargin = Math.round((canvasHeight - gridSize.gridHeight) / 2);
        return {
            left: canvasBounds[0] + leftMargin,
            top: canvasBounds[1] - topMargin,
            cols: gridSize.occupiedColumns,
            rows: gridSize.rows,
            gridWidth: gridSize.gridWidth,
            gridHeight: gridSize.gridHeight
        };
    }

    /**
     * 全体の外形が最も正方形に近づく列数を求める（セルの縦横比と間隔を考慮する）
     * Find the column count whose overall grid outline is closest to a square (accounts for cell aspect ratio and gap).
     * @param {number} itemCount - アートボード数
     * @param {number} cellWidth - セルの幅（pt）
     * @param {number} cellHeight - セルの高さ（pt）
     * @param {number} gap - 間隔（pt）
     * @returns {number} 列数
     */
    function chooseBestColumnCount(itemCount, cellWidth, cellHeight, gap) {
        var bestColumns = 1, bestScore = null, bestEmptyCells = 0, columnCandidate;
        for (columnCandidate = 1; columnCandidate <= itemCount; columnCandidate++) {
            var gridSize = computeGridSize(itemCount, columnCandidate, cellWidth, cellHeight, gap);
            var aspectRatio = gridSize.gridWidth / gridSize.gridHeight;
            var score = (aspectRatio >= 1) ? aspectRatio : (1 / aspectRatio); // 1 に近いほど正方形 / closer to 1 = squarer
            var emptyCells = gridSize.rows * columnCandidate - itemCount;      // 余りセル数 / unused cells
            if (bestScore === null
                || score < bestScore - 0.0001
                || (Math.abs(score - bestScore) <= 0.0001 && emptyCells < bestEmptyCells)) {
                bestScore = score;
                bestEmptyCells = emptyCells;
                bestColumns = columnCandidate;
            }
        }
        return bestColumns;
    }

    // =========================================
    // ユーティリティ / Utilities
    // =========================================

    /**
     * 全レイヤー（サブレイヤー含む）を再帰的に収集する / Collect all layers recursively, including sublayers
     * @param {Document} doc - 対象ドキュメント
     * @returns {Layer[]} すべてのレイヤー
     */
    function collectAllLayers(doc) {
        var allLayers = [];

        /**
         * レイヤーとその下のサブレイヤーを allLayers へ積む
         * @param {Layers} childLayers - 走査するレイヤー
         * @returns {void}
         */
        function collectLayersRecursive(childLayers) {
            for (var i = 0; i < childLayers.length; i++) {
                allLayers.push(childLayers[i]);
                collectLayersRecursive(childLayers[i].layers);
            }
        }
        collectLayersRecursive(doc.layers);
        return allLayers;
    }

    /**
     * レイヤー直下の最上位アイテムのみ収集する（グループ内は親ごと動かす）/ Collect only top-level items (children move with their parent)
     * @param {Document} doc - 対象ドキュメント
     * @returns {PageItem[]} 最上位のアイテム
     */
    function collectTopLevelItems(doc) {
        var topLevelItems = [], pageItems = doc.pageItems, i;
        for (i = 0; i < pageItems.length; i++) {
            if (pageItems[i].parent && pageItems[i].parent.typename === "Layer") {
                topLevelItems.push(pageItems[i]);
            }
        }
        return topLevelItems;
    }

    /**
     * 矩形 [left, top, right, bottom] の中心点 / Center point of a [left, top, right, bottom] rect
     * @param {number[]} rect - 矩形
     * @returns {number[]} [x, y]
     */
    function getRectCenter(rect) {
        return [(rect[0] + rect[2]) / 2, (rect[1] + rect[3]) / 2];
    }

    /**
     * ロックを一時解除し、復元用の状態を返す / Temporarily unlock; returns the state needed to restore it
     * @param {Document} doc - 対象ドキュメント
     * @returns {{layers: Layer[], items: PageItem[]}} ロックを外したレイヤーとアイテム
     */
    function unlockLayersAndItems(doc) {
        var lockState = { layers: [], items: [] }, i;
        var allLayers = collectAllLayers(doc);
        for (i = 0; i < allLayers.length; i++) {
            if (allLayers[i].locked) { lockState.layers.push(allLayers[i]); allLayers[i].locked = false; }
        }
        var pageItems = doc.pageItems;
        for (i = 0; i < pageItems.length; i++) {
            if (pageItems[i].locked) { lockState.items.push(pageItems[i]); pageItems[i].locked = false; }
        }
        return lockState;
    }

    /**
     * unlockLayersAndItems で外したロックを元に戻す / Re-apply the locks released by unlockLayersAndItems
     * @param {{layers: Layer[], items: PageItem[]}} lockState - unlockLayersAndItems() の戻り値
     * @returns {void}
     */
    function restoreLockState(lockState) {
        var i;
        for (i = 0; i < lockState.items.length; i++) { lockState.items[i].locked = true; }
        for (i = 0; i < lockState.layers.length; i++) { lockState.layers[i].locked = true; }
    }

    /**
     * 小数 1 桁に丸めて文字列化 / Round to 1 decimal place and stringify
     * @param {number} value - 丸める値
     * @returns {string} 表示用の文字列
     */
    function formatNumber(value) {
        return String(Math.round(value * 10) / 10);
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 列数・間隔の入力を読み取り、検証する
     * @param {EditText} columnInput - 列数の入力欄
     * @param {EditText} gapInput - 間隔の入力欄（定規単位）
     * @returns {{columns: number, gapInRulerUnit: number, errorLabelPath: string|null}} 読み取った値と、不正なときのエラー文言のパス
     */
    function readLayoutInputs(columnInput, gapInput) {
        var columns = parseInt(columnInput.text, 10);
        var gapInRulerUnit = parseFloat(gapInput.text);
        var errorLabelPath = null;
        if (isNaN(columns) || columns < 1) {
            errorLabelPath = "alert.badColumns";
        } else if (isNaN(gapInRulerUnit) || gapInRulerUnit < 0) {
            errorLabelPath = "alert.badGap";
        }
        return { columns: columns, gapInRulerUnit: gapInRulerUnit, errorLabelPath: errorLabelPath };
    }

    /**
     * 同じ行に∧∨と入力欄を隙間なく並べて追加する（↑↓キーも∧∨と同じ処理で増減し、増減後は onChanging を通す）
     * @param {Group} parentRow - 追加先の行
     * @param {string} initialText - 入力欄の初期値
     * @param {Object} stepOptions - addStepper() に渡す integer / min など
     * @returns {EditText} 追加した入力欄
     */
    function addStepperInput(parentRow, initialText, stepOptions) {
        var stepperInputGroup = parentRow.add("group");
        stepperInputGroup.orientation = "row";
        stepperInputGroup.alignChildren = ["left", "center"];
        stepperInputGroup.spacing = 0;
        stepperInputGroup.margins = 0;
        var numberInput;
        stepOptions.onStep = function (steppedInput) { if (steppedInput.onChanging) steppedInput.onChanging(); };
        var stepperGroup = addStepper(stepperInputGroup, function () { return numberInput; }, stepOptions);
        numberInput = stepperInputGroup.add("edittext", undefined, initialText);
        bindSteppedArrowKeys(numberInput, stepperGroup);
        return numberInput;
    }

    /**
     * 列数と間隔を指定するダイアログを表示する
     * @param {number} artboardCount - アートボード数
     * @param {number} cellWidth - セルの幅（pt）
     * @param {number} cellHeight - セルの高さ（pt）
     * @param {{label: string, pointsPerUnit: number}} rulerUnit - 定規の単位
     * @returns {{columns: number, gapInPoints: number}|null} 指定した列数と間隔（キャンセルなら null）
     */
    function showGridDialog(artboardCount, cellWidth, cellHeight, rulerUnit) {
        var dialogResult = null;

        var gridDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        gridDialog.orientation = "column";
        gridDialog.alignChildren = "fill";
        gridDialog.margins = DIALOG_MARGINS;
        gridDialog.spacing = DIALOG_SPACING;

        gridDialog.add("statictext", undefined, labelValueText("fieldLabel.artboardCount", artboardCount));

        var settingsPanel = gridDialog.add("panel", undefined, getLabel("panel.settings"));
        settingsPanel.orientation = "column";
        settingsPanel.alignChildren = "left";
        settingsPanel.margins = PANEL_MARGINS;
        settingsPanel.spacing = PANEL_SPACING;

        /* 間隔 / Gap（既定 = アートボード幅の 1/5）/ default = 1/5 of artboard width */
        var defaultGapInPoints = cellWidth / 5;
        var gapGroup = settingsPanel.add("group");
        gapGroup.add("statictext", undefined, labelText("fieldLabel.gap"));
        var gapInput = addStepperInput(gapGroup, formatNumber(defaultGapInPoints / rulerUnit.pointsPerUnit), { min: 0 });
        gapInput.helpTip = getLabel("tooltip.gap");
        gapInput.characters = NUMBER_FIELD_CHARS;
        gapGroup.add("statictext", undefined, rulerUnit.label);

        /* 列数 / Columns */
        var recommendedColumns = chooseBestColumnCount(artboardCount, cellWidth, cellHeight, defaultGapInPoints);
        var columnGroup = settingsPanel.add("group");
        columnGroup.add("statictext", undefined, labelText("fieldLabel.columns"));
        var columnInput = addStepperInput(columnGroup, String(recommendedColumns), { integer: true, min: 1 });
        columnInput.helpTip = getLabel("tooltip.columns");
        columnInput.characters = NUMBER_FIELD_CHARS;
        var btnAuto = columnGroup.add("button", undefined, getLabel("button.auto"));
        btnAuto.helpTip = getLabel("tooltip.auto");
        btnAuto.preferredSize.height = AUTO_BUTTON_HEIGHT;

        /* プレビュー / Preview */
        var previewText = settingsPanel.add("statictext", undefined, "", { multiline: true });
        previewText.preferredSize = PREVIEW_TEXT_SIZE;

        /**
         * 入力値から行列と外形サイズ、推奨列数を表示し直す
         * @returns {void}
         */
        function updatePreview() {
            var layoutInputs = readLayoutInputs(columnInput, gapInput);
            if (layoutInputs.errorLabelPath) { previewText.text = getLabel(layoutInputs.errorLabelPath); return; }
            var columns = layoutInputs.columns;
            var gapInPoints = layoutInputs.gapInRulerUnit * rulerUnit.pointsPerUnit;
            var gridSize = computeGridSize(artboardCount, columns, cellWidth, cellHeight, gapInPoints);
            var bestColumns = chooseBestColumnCount(artboardCount, cellWidth, cellHeight, gapInPoints);
            previewText.text = columns + " " + getLabel("status.columnUnit") + " x " + gridSize.rows + " " + getLabel("status.rowUnit")
                + "  /  " + formatNumber(gridSize.gridWidth / rulerUnit.pointsPerUnit)
                + " x " + formatNumber(gridSize.gridHeight / rulerUnit.pointsPerUnit) + " " + rulerUnit.label
                + "\n" + labelValueText("status.recommend", bestColumns + " " + getLabel("status.columnUnit"));
        }

        gapInput.onChanging = updatePreview;
        columnInput.onChanging = updatePreview;
        btnAuto.onClick = function () {
            var gapInRulerUnit = parseFloat(gapInput.text);
            if (isNaN(gapInRulerUnit) || gapInRulerUnit < 0) gapInRulerUnit = 0;
            columnInput.text = String(chooseBestColumnCount(
                artboardCount, cellWidth, cellHeight, gapInRulerUnit * rulerUnit.pointsPerUnit));
            updatePreview();
        };
        updatePreview();

        /* ボタン / Buttons (Mac 規約: Cancel -> OK) */
        var buttonRow = addButtonRow(gridDialog);
        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
        gridDialog.cancelElement = btnCancel;
        gridDialog.defaultElement = btnOK;

        btnOK.onClick = function () {
            var layoutInputs = readLayoutInputs(columnInput, gapInput);
            if (layoutInputs.errorLabelPath) { alert(getLabel(layoutInputs.errorLabelPath)); return; }
            dialogResult = { columns: layoutInputs.columns, gapInPoints: layoutInputs.gapInRulerUnit * rulerUnit.pointsPerUnit };
            gridDialog.close();
        };

        prepareDialogWindow(gridDialog, SCRIPT_NAME);
        gridDialog.show();
        return dialogResult;
    }

    // =========================================
    // 再配置 / Rearrangement
    // =========================================

    /**
     * 座標系を元へ戻す（控えられなかったときは何もしない）
     * @param {CoordinateSystem|null} savedCoordinateSystem - 控えた座標系
     * @returns {void}
     */
    function restoreCoordinateSystem(savedCoordinateSystem) {
        if (savedCoordinateSystem !== null) {
            try { app.coordinateSystem = savedCoordinateSystem; } catch (e) { }
        }
    }

    /**
     * アートワークを中心点が含まれるアートボードに振り分ける（移動前に呼ぶ）
     * @param {PageItem[]} topLevelItems - 最上位のアイテム
     * @param {number[][]} artboardRects - 移動前のアートボード矩形
     * @returns {PageItem[][]} アートボードごとのアイテム
     */
    function assignItemsToArtboards(topLevelItems, artboardRects) {
        var itemsByArtboard = [];
        for (var i = 0; i < artboardRects.length; i++) itemsByArtboard.push([]);
        for (var j = 0; j < topLevelItems.length; j++) {
            var itemCenter = getRectCenter(topLevelItems[j].geometricBounds); // 中心点で所属を決定 / assign by center point
            for (var k = 0; k < artboardRects.length; k++) {
                var candidateRect = artboardRects[k];
                if (itemCenter[0] >= candidateRect[0] && itemCenter[0] <= candidateRect[2]
                    && itemCenter[1] <= candidateRect[1] && itemCenter[1] >= candidateRect[3]) {
                    itemsByArtboard[k].push(topLevelItems[j]);
                    break;
                }
            }
        }
        return itemsByArtboard;
    }

    /**
     * アートボードをグリッドのセルへ移し、所属アートワークも同じだけ動かす
     * @param {Document} doc - 対象ドキュメント
     * @param {number[][]} originalArtboardRects - 移動前のアートボード矩形
     * @param {PageItem[][]} itemsByArtboard - アートボードごとのアイテム
     * @param {object} gridLayout - gridOrigin / cellWidth / cellHeight / gapInPoints / columnCount
     * @returns {number} 移動できなかったアイテムの数
     */
    function moveArtboardsToGrid(doc, originalArtboardRects, itemsByArtboard, gridLayout) {
        var failedMoveCount = 0;
        for (var i = 0; i < originalArtboardRects.length; i++) {
            var gridRow = Math.floor(i / gridLayout.columnCount);
            var gridColumn = i % gridLayout.columnCount;
            var cellLeft = gridLayout.gridOrigin.left + gridColumn * (gridLayout.cellWidth + gridLayout.gapInPoints);
            var cellTop = gridLayout.gridOrigin.top - gridRow * (gridLayout.cellHeight + gridLayout.gapInPoints);

            var originalRect = originalArtboardRects[i];
            var artboardWidth = originalRect[2] - originalRect[0];
            var artboardHeight = originalRect[1] - originalRect[3];
            /* サイズが異なるアートボードはセル内で中央に置く / center within the cell */
            var newLeft = cellLeft + (gridLayout.cellWidth - artboardWidth) / 2;
            var newTop = cellTop - (gridLayout.cellHeight - artboardHeight) / 2;

            var deltaX = newLeft - originalRect[0];
            var deltaY = newTop - originalRect[1];

            doc.artboards[i].artboardRect = [newLeft, newTop, newLeft + artboardWidth, newTop - artboardHeight];

            var artboardItems = itemsByArtboard[i];
            for (var j = 0; j < artboardItems.length; j++) {
                try { artboardItems[j].translate(deltaX, deltaY); } catch (e) { failedMoveCount++; }
            }
        }
        return failedMoveCount;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * ダイアログで列数と間隔を決め、アートボードを正方形に近いグリッドへ並べ直す
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) { alert(getLabel("alert.noDocument")); return; }
        var doc = app.activeDocument;
        var artboardCount = doc.artboards.length;
        if (artboardCount < 2) { alert(getLabel("alert.needTwo")); return; }

        /* artboardRect と geometricBounds を同じ座標系に揃える / unify the coordinate space */
        var savedCoordinateSystem = null;
        try {
            savedCoordinateSystem = app.coordinateSystem;
            app.coordinateSystem = CoordinateSystem.DOCUMENTCOORDINATESYSTEM;
        } catch (e) { }

        /* 元の矩形を保存し、最大サイズを 1 セルとする / capture original rects; cell = largest artboard */
        var originalArtboardRects = [], cellWidth = 0, cellHeight = 0;
        for (var i = 0; i < artboardCount; i++) {
            var artboardRect = doc.artboards[i].artboardRect; // [left, top, right, bottom]
            originalArtboardRects.push(artboardRect);
            var artboardWidth = artboardRect[2] - artboardRect[0];
            var artboardHeight = artboardRect[1] - artboardRect[3];
            if (artboardWidth > cellWidth) cellWidth = artboardWidth;
            if (artboardHeight > cellHeight) cellHeight = artboardHeight;
        }

        var rulerUnit = getUnitInfo();

        var layoutSettings = showGridDialog(artboardCount, cellWidth, cellHeight, rulerUnit);
        if (!layoutSettings) { restoreCoordinateSystem(savedCoordinateSystem); return; }
        var gapInPoints = layoutSettings.gapInPoints;
        var columnCount = layoutSettings.columns;

        /* アートワークの所属アートボードを判定（移動前に）/ assign artwork to artboards (before any move) */
        var itemsByArtboard = assignItemsToArtboards(collectTopLevelItems(doc), originalArtboardRects);

        /* グリッド原点（カンバス中央）/ grid origin centered on the canvas */
        var canvasBounds = getCanvasBounds();
        var gridOrigin = computeCenteredGridOrigin(canvasBounds, cellWidth, cellHeight, gapInPoints, columnCount, artboardCount);

        /* 再配置 / reposition */
        var lockState = unlockLayersAndItems(doc);
        var failedMoveCount = 0;
        try {
            failedMoveCount = moveArtboardsToGrid(doc, originalArtboardRects, itemsByArtboard, {
                gridOrigin: gridOrigin,
                cellWidth: cellWidth,
                cellHeight: cellHeight,
                gapInPoints: gapInPoints,
                columnCount: columnCount
            });
        } finally {
            restoreLockState(lockState);
            restoreCoordinateSystem(savedCoordinateSystem);
        }

        app.redraw();

        var resultMessage = getLabel("alert.done") + "\n"
            + labelValueText("fieldLabel.artboardCount", artboardCount) + "  /  "
            + columnCount + " " + getLabel("status.columnUnit") + " x " + gridOrigin.rows + " " + getLabel("status.rowUnit");
        if (failedMoveCount > 0) resultMessage += "\n" + failedMoveCount + getLabel("alert.moveFailed");
        alert(resultMessage);
    }

    // =========================================
    // 実行 / Run
    // =========================================
    try {
        main();
    } catch (err) {
        alert("Error: " + err + (err.line ? " (line " + err.line + ")" : ""));
    }

})();
