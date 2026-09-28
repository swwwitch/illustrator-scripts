#target illustrator
#targetengine "DeleteOutsideArtboardEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

アクティブなアートボードの外にあるオブジェクト、またはアートボード内の選択していないオブジェクトを、削除するか保管用レイヤーへ移します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/DeleteOutsideArtboard.md

### Overview

Deletes the objects outside the active artboard, or the unselected objects inside it, or moves them to a backup layer.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/DeleteOutsideArtboard.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "DeleteOutsideArtboard";        /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.4.4";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-07-08";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-28";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/DeleteOutsideArtboard.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/DeleteOutsideArtboard.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* 保管用レイヤーの名前（このレイヤーのオブジェクトは対象外）/ Backup layer name (its objects are never touched) */
    var BACKUP_LAYER_NAME = "// backup";

    // =========================================
    // レイアウト / Layout
    // =========================================

    var PANEL_MARGINS  = [15, 20, 15, 10];  /* パネル余白 [左,上,右,下] */
    var OPTION_MARGINS = [15, 0, 15, 10];   /* オプション欄の余白 [左,上,右,下] */

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

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "オブジェクトを削除", en: "Delete Objects" }
        },
        panel: {
            inside:  { ja: "アートボード内", en: "Inside Artboard" },
            outside: { ja: "アートボード外", en: "Outside Artboard" }
        },
        radio: {
            keepSelected: { ja: "選択オブジェクトを残す", en: "Keep Selected Objects" },
            keepAll:      { ja: "すべて残す", en: "Keep All Objects" },
            remove:       { ja: "削除", en: "Delete" },
            ignore:       { ja: "無視（残す）", en: "Ignore (Keep)" }
        },
        checkbox: {
            includeLocked: { ja: "ロックされたオブジェクトを含む", en: "Include Locked Objects" },
            moveToBackup:  { ja: "保管用レイヤーに移す", en: "Move to Backup Layer" }
        },
        tooltip: {
            keepSelected: {
                ja: "アクティブなアートボード内のオブジェクトのうち、選択しているものだけを残して他を削除します。このときアートボード外の設定は使いません。",
                en: "Inside the active artboard, keeps only the selected objects and deletes the rest. The Outside Artboard setting is not used."
            },
            keepAll: {
                ja: "アートボード内のオブジェクトはすべて残します。",
                en: "Keeps every object inside the artboard."
            },
            remove: {
                ja: "アクティブなアートボードの外にはみ出したオブジェクトを削除します。",
                en: "Deletes the objects that sit outside the active artboard."
            },
            ignore: { ja: "アートボードの外のオブジェクトはそのまま残します。", en: "Leaves the objects outside the artboard untouched." },
            includeLocked: {
                ja: "ロックされたオブジェクトも処理の対象にします。オフのときは触りません。",
                en: "Includes locked objects. They are left alone when this is off."
            },
            moveToBackup: {
                ja: "削除せずに保管用のレイヤーへ移します。あとから戻せます。",
                en: "Moves the objects to a backup layer instead of deleting them, so they can be brought back."
            }
        },
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok:     { ja: "削除", en: "Delete" }
        },
        alert: {
            noDocument:  { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noSelection: { ja: "選択オブジェクトがありません。", en: "No objects selected." },
            noTargets:   { ja: "削除対象のオブジェクトはありません。", en: "No objects to delete." }
        }
    };

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * ラジオボタンを1つ追加し、ツールチップを設定する
     * @param {Group|Panel} parentContainer - 追加先のコンテナ
     * @param {string} labelKey - radio / tooltip 共通のキー名
     * @returns {RadioButton} 追加したラジオボタン
     */
    function addRadio(parentContainer, labelKey) {
        var radioButton = parentContainer.add("radiobutton", undefined, getLabel("radio." + labelKey));
        radioButton.helpTip = getLabel("tooltip." + labelKey);
        return radioButton;
    }

    /**
     * チェックボックスを1つ追加し、ツールチップと初期値を設定する
     * @param {Group|Panel} parentContainer - 追加先のコンテナ
     * @param {string} labelKey - checkbox / tooltip 共通のキー名
     * @param {boolean} initialValue - 初期値
     * @returns {Checkbox} 追加したチェックボックス
     */
    function addCheckbox(parentContainer, labelKey, initialValue) {
        var checkbox = parentContainer.add("checkbox", undefined, getLabel("checkbox." + labelKey));
        checkbox.helpTip = getLabel("tooltip." + labelKey);
        checkbox.value = initialValue;
        return checkbox;
    }

    /**
     * 縦並びのパネルを追加する
     * @param {Window} dialog - 追加先のダイアログ
     * @param {string} labelKey - panel のキー名
     * @returns {Panel} 追加したパネル
     */
    function addPanel(dialog, labelKey) {
        var panel = dialog.add("panel", undefined, getLabel("panel." + labelKey));
        panel.orientation = "column";
        panel.alignChildren = "left";
        panel.margins = PANEL_MARGINS;
        return panel;
    }

    /**
     * ダイアログを表示し、選ばれた設定を返す
     * @param {boolean} hasSelection - 選択オブジェクトがあるか（「選択オブジェクトを残す」の初期値）
     * @returns {{keepSelected: boolean, deleteOutside: boolean, moveToBackup: boolean, includeLocked: boolean}|null} 設定。キャンセル時は null
     */
    function showDialog(hasSelection) {
        var dialog = new Window("dialog", getLabel("dialog.title"));
        dialog.orientation = "column";
        dialog.alignChildren = "fill";

        /* アートボード内パネル / Inside-artboard panel */
        var insidePanel = addPanel(dialog, "inside");
        var keepSelectedRadio = addRadio(insidePanel, "keepSelected");
        var keepAllRadio = addRadio(insidePanel, "keepAll");
        /* 選択があれば「選択オブジェクトを残す」を初期値に / Default to "keep selected" when something is selected */
        keepSelectedRadio.value = hasSelection;
        keepAllRadio.value = !hasSelection;

        /* アートボード外パネル / Outside-artboard panel */
        var outsidePanel = addPanel(dialog, "outside");
        var removeRadio = addRadio(outsidePanel, "remove");
        var ignoreRadio = addRadio(outsidePanel, "ignore");
        ignoreRadio.value = true;

        /* オプション（ロック含む、保管用レイヤー）/ Options (include locked, backup layer) */
        var optionGroup = dialog.add("group");
        optionGroup.orientation = "column";
        optionGroup.alignChildren = "left";
        optionGroup.margins = OPTION_MARGINS;
        var includeLockedCheckbox = addCheckbox(optionGroup, "includeLocked", true);
        var moveToBackupCheckbox = addCheckbox(optionGroup, "moveToBackup", false);

        /* ボタン / Buttons */
        var buttonRow = addButtonRow(dialog, { centered: true });
        var btnCancel = buttonRow.rowGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnOK = buttonRow.rowGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });

        prepareDialogWindow(dialog, SCRIPT_NAME);
        if (dialog.show() !== 1) return null;
        return {
            keepSelected: keepSelectedRadio.value,
            deleteOutside: removeRadio.value,
            moveToBackup: moveToBackupCheckbox.value,
            includeLocked: includeLockedCheckbox.value
        };
    }

    // =========================================
    // 判定と収集 / Hit testing and collection
    // =========================================

    /**
     * オブジェクトがアートボードと重なっているか判定する（visibleBounds の矩形で判定）
     * @param {PageItem} item - 判定するオブジェクト
     * @param {Artboard} artboard - アートボード
     * @returns {boolean} 重なっていれば true
     */
    function isOverlappingArtboard(item, artboard) {
        var itemRect = item.visibleBounds;
        var abRect = artboard.artboardRect;
        var overlapWidth = Math.min(itemRect[2], abRect[2]) - Math.max(itemRect[0], abRect[0]);
        var overlapHeight = Math.min(itemRect[1], abRect[1]) - Math.max(itemRect[3], abRect[3]);
        return overlapWidth > 0 && overlapHeight > 0;
    }

    /**
     * 重なり判定に使うオブジェクトを返す（クリップグループはクリッピングパス）
     * @param {PageItem} item - 対象オブジェクト
     * @returns {PageItem} 判定に使うオブジェクト
     */
    function getHitTestTarget(item) {
        if (item.typename === "GroupItem" && item.clipped) {
            for (var k = 0; k < item.pageItems.length; k++) {
                if (item.pageItems[k].clipping) return item.pageItems[k];
            }
        }
        return item;
    }

    /**
     * 収集の対象にするか判定し、対象ならロックと非表示を解除する
     * @param {PageItem} item - 対象オブジェクト
     * @param {boolean} includeLocked - ロックされたオブジェクトも対象にするか
     * @returns {boolean} 対象なら true
     */
    function prepareCandidate(item, includeLocked) {
        if (item.layer && item.layer.name === BACKUP_LAYER_NAME) return false;
        if (!includeLocked && item.locked) return false;
        if (item.locked) item.locked = false;
        if (!item.visible) item.visible = true;
        return true;
    }

    /**
     * アートボードと重ならないオブジェクトを集める（重なるグループは中身も調べる）
     * @param {PageItems} items - 調べるオブジェクト
     * @param {Artboard} artboard - 基準のアートボード
     * @param {boolean} includeLocked - ロックされたオブジェクトも対象にするか
     * @param {PageItem[]} result - 結果を追加する配列
     * @returns {void}
     */
    function collectOutsideItems(items, artboard, includeLocked, result) {
        for (var i = items.length - 1; i >= 0; i--) {
            var item = items[i];
            if (!prepareCandidate(item, includeLocked)) continue;
            var overlaps = isOverlappingArtboard(getHitTestTarget(item), artboard);
            if (!overlaps) {
                result.push(item);
            } else if (item.typename === "GroupItem") {
                collectOutsideItems(item.pageItems, artboard, includeLocked, result);
            }
        }
    }

    /**
     * アートボードと重なるオブジェクトを集める（重なるグループは中身も調べる）
     * @param {PageItems} items - 調べるオブジェクト
     * @param {Artboard} artboard - 基準のアートボード
     * @param {boolean} includeLocked - ロックされたオブジェクトも対象にするか
     * @param {PageItem[]} result - 結果を追加する配列
     * @returns {void}
     */
    function collectInsideItems(items, artboard, includeLocked, result) {
        for (var i = items.length - 1; i >= 0; i--) {
            var item = items[i];
            if (!prepareCandidate(item, includeLocked)) continue;
            if (!isOverlappingArtboard(getHitTestTarget(item), artboard)) continue;
            result.push(item);
            if (item.typename === "GroupItem") {
                collectInsideItems(item.pageItems, artboard, includeLocked, result);
            }
        }
    }

    /**
     * 選択に含まれないものだけを返す
     * @param {PageItem[]} items - 対象オブジェクト
     * @param {Array} selection - 選択オブジェクト
     * @returns {PageItem[]} 選択されていないオブジェクト
     */
    function excludeSelected(items, selection) {
        var filteredItems = [];
        for (var i = 0; i < items.length; i++) {
            var isSelected = false;
            for (var j = 0; j < selection.length; j++) {
                if (items[i] === selection[j]) {
                    isSelected = true;
                    break;
                }
            }
            if (!isSelected) filteredItems.push(items[i]);
        }
        return filteredItems;
    }

    // =========================================
    // 削除と移動 / Delete and move
    // =========================================

    /**
     * 保管用レイヤーを取得する（無ければ作る）
     * @param {Document} doc - 対象ドキュメント
     * @returns {Layer} 保管用レイヤー
     */
    function getBackupLayer(doc) {
        var backupLayer;
        try {
            backupLayer = doc.layers.getByName(BACKUP_LAYER_NAME);
        } catch (e) {
            /* 見つからないと例外 / getByName throws when missing */
            backupLayer = doc.layers.add();
            backupLayer.name = BACKUP_LAYER_NAME;
        }
        return backupLayer;
    }

    /**
     * オブジェクトと、その中身すべてのロックと非表示を解除する
     * @param {PageItem} target - 対象オブジェクト
     * @returns {void}
     */
    function unlockAndShowAll(target) {
        if (target.locked) target.locked = false;
        if (!target.visible) target.visible = true;
        if (typeof target.pageItems !== "undefined") {
            for (var i = 0; i < target.pageItems.length; i++) {
                unlockAndShowAll(target.pageItems[i]);
            }
        }
    }

    /**
     * オブジェクトを保管用レイヤーへ移し、レイヤーを非表示にする
     * @param {PageItem} item - 移すオブジェクト
     * @param {Document} doc - 対象ドキュメント
     * @returns {void}
     */
    function moveItemToBackupLayer(item, doc) {
        var backupLayer = getBackupLayer(doc);
        if (backupLayer.locked) backupLayer.locked = false;
        if (!backupLayer.visible) backupLayer.visible = true;
        unlockAndShowAll(item);
        item.move(backupLayer, ElementPlacement.PLACEATBEGINNING);
        backupLayer.visible = false;
    }

    /**
     * 対象オブジェクトを削除するか保管用レイヤーへ移す
     * @param {PageItem[]} targetItems - 対象オブジェクト
     * @param {Document} doc - 対象ドキュメント
     * @param {boolean} moveToBackup - 削除せずに保管用レイヤーへ移すか
     * @returns {void}
     */
    function disposeItems(targetItems, doc, moveToBackup) {
        for (var i = 0; i < targetItems.length; i++) {
            var item = targetItems[i];
            /* doc.pageItems はグループの中身も含むため、同じオブジェクトが重複して入り、
               先に処理した親ごと消えたものは例外になる / Duplicates via nested pageItems may already be gone */
            try {
                if (item.locked) item.locked = false;
                if (!item.visible) item.visible = true;
                if (item.layer && item.layer.locked) item.layer.locked = false;
                if (moveToBackup) {
                    moveItemToBackupLayer(item, doc);
                } else {
                    item.remove();
                }
            } catch (e) {
                $.writeln("Skipped invalid object: " + e);
            }
        }
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * メイン処理
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }
        var doc = app.activeDocument;
        var currentSelection = doc.selection;
        var hasSelection = !!(currentSelection && currentSelection.length > 0);

        var deleteOptions = showDialog(hasSelection);
        if (!deleteOptions) return;
        /* 「すべて残す」＋「無視」なら何もしない / Nothing to do when both panels keep everything */
        if (!deleteOptions.keepSelected && !deleteOptions.deleteOutside) return;

        var activeArtboard = doc.artboards[doc.artboards.getActiveArtboardIndex()];
        var targetItems = [];

        if (deleteOptions.keepSelected) {
            /* アートボード内の選択していないオブジェクト / Unselected objects inside the artboard */
            if (!hasSelection) {
                alert(getLabel("alert.noSelection"));
                return;
            }
            var insideItems = [];
            collectInsideItems(doc.pageItems, activeArtboard, deleteOptions.includeLocked, insideItems);
            targetItems = excludeSelected(insideItems, currentSelection);
        } else {
            /* アクティブなアートボードの外のオブジェクト / Objects outside the active artboard */
            collectOutsideItems(doc.pageItems, activeArtboard, deleteOptions.includeLocked, targetItems);
        }

        if (targetItems.length === 0) {
            alert(getLabel("alert.noTargets"));
            return;
        }
        disposeItems(targetItems, doc, deleteOptions.moveToBackup);
    }

    main();

})();
