#target illustrator
#targetengine "AiSetKinsokuEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したテキストに、禁則処理のプリセットを適用します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AiSetKinsoku.md

### Overview

Applies a kinsoku (line-breaking) preset to the selected text.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AiSetKinsoku.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AiSetKinsoku";                 /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-28";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AiSetKinsoku.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AiSetKinsoku.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* 禁則プリセット（表示順）。kinsokuName は paragraphAttributes.kinsoku に渡す値、labelPath は表示名の LABELS キー
       Kinsoku presets in display order: kinsokuName goes to paragraphAttributes.kinsoku, labelPath names the LABELS entry */
    var KINSOKU_PRESETS = [
        { kinsokuName: "None",    labelPath: "radio.none" },
        { kinsokuName: "Hard",    labelPath: "radio.hard" },
        { kinsokuName: "Soft",    labelPath: "radio.soft" },
        { kinsokuName: "Soft_v2", labelPath: "radio.softV2" }
    ];

    // =========================================
    // レイアウト / Layout
    // =========================================
    var PANEL_MARGINS = [16, 20, 16, 12]; /* パネル余白 [左,上,右,下] / panel margins [L,T,R,B] */
    var PANEL_SPACING = 6;                /* パネル内の要素間隔 / spacing inside the panel */
    var BUTTON_SPACING = 8;               /* ボタンの間隔 / spacing between buttons */

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
            title: { ja: "禁則処理の設定", en: "Kinsoku Settings" }
        },
        panel: {
            preset: { ja: "禁則処理", en: "Kinsoku" }
        },
        radio: {
            none: { ja: "なし", en: "None" },
            hard: { ja: "強い禁則", en: "Hard" },
            soft: { ja: "弱い禁則", en: "Soft" },
            softV2: { ja: "弱い禁則 v2", en: "Soft v2" }
        },
        tooltip: {
            preset: {
                ja: "クリックすると、選択中のテキストにすぐ適用します",
                en: "Applies to the selected text as soon as you click"
            },
            close: {
                ja: "適用済みの禁則処理はそのままにして閉じます",
                en: "Closes the dialog and keeps the kinsoku already applied"
            }
        },
        button: {
            close: { ja: "閉じる", en: "Close" },
            ok: { ja: "OK", en: "OK" }
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noSelection: { ja: "テキストオブジェクトを選択してください。", en: "Select text objects." }
        }
    };

    // =========================================
    // 禁則の適用 / Applying kinsoku
    // =========================================

    /**
     * テキストフレームに禁則を適用する
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @param {string} kinsokuName - paragraphAttributes.kinsoku に渡す値
     * @returns {void}
     */
    function applyKinsokuToTextFrame(textFrame, kinsokuName) {
        /* 「なし」= "None" はスクリプトから設定できず例外になるので、そのまま続行する
           "None" cannot be set from a script and throws; carry on */
        try {
            textFrame.textRange.paragraphAttributes.kinsoku = kinsokuName;
        } catch (e) {}
    }

    /**
     * テキストフレームには直接、グループには中身へ再帰して禁則を適用する
     * @param {PageItem} pageItem - 対象のオブジェクト
     * @param {string} kinsokuName - paragraphAttributes.kinsoku に渡す値
     * @returns {void}
     */
    function applyKinsokuToItem(pageItem, kinsokuName) {
        if (pageItem.typename === "TextFrame") {
            applyKinsokuToTextFrame(pageItem, kinsokuName);
        } else if (pageItem.typename === "GroupItem") {
            for (var i = 0; i < pageItem.pageItems.length; i++) {
                applyKinsokuToItem(pageItem.pageItems[i], kinsokuName);
            }
        }
    }

    /**
     * 選択中のオブジェクトすべてに禁則を適用して再描画する
     * @param {Document} doc - 対象のドキュメント
     * @param {string} kinsokuName - paragraphAttributes.kinsoku に渡す値
     * @returns {void}
     */
    function applyKinsokuToSelection(doc, kinsokuName) {
        for (var i = 0; i < doc.selection.length; i++) {
            applyKinsokuToItem(doc.selection[i], kinsokuName);
        }
        app.redraw();
    }

    // =========================================
    // 現在値の読み取り / Reading the current value
    // =========================================

    /**
     * テキストフレームの現在の禁則を返す
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @returns {string} 禁則の値。「なし」のときは ""
     */
    function getKinsoku(textFrame) {
        /* 禁則「なし」の段落は getter が Error 9563 を投げる / The getter throws Error 9563 on "None" paragraphs */
        try {
            return textFrame.textRange.paragraphAttributes.kinsoku;
        } catch (e) {
            return "";
        }
    }

    /**
     * オブジェクト（グループ内を含む）から最初のテキストフレームを探す
     * @param {PageItem} pageItem - 探す対象
     * @returns {TextFrame|null} 見つかったテキストフレーム。無ければ null
     */
    function findFirstTextFrame(pageItem) {
        if (pageItem.typename === "TextFrame") return pageItem;
        if (pageItem.typename === "GroupItem") {
            for (var i = 0; i < pageItem.pageItems.length; i++) {
                var foundTextFrame = findFirstTextFrame(pageItem.pageItems[i]);
                if (foundTextFrame !== null) return foundTextFrame;
            }
        }
        return null;
    }

    /**
     * 選択の中で最初に見つかるテキストフレームの禁則を返す（初期選択の判定用）
     * @param {Document} doc - 対象のドキュメント
     * @returns {string} 禁則の値。テキストフレームが無いか「なし」のときは ""
     */
    function getSelectionKinsoku(doc) {
        for (var i = 0; i < doc.selection.length; i++) {
            var firstTextFrame = findFirstTextFrame(doc.selection[i]);
            if (firstTextFrame !== null) return getKinsoku(firstTextFrame);
        }
        return "";
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

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 禁則プリセットのラジオを並べ、クリックで即適用するようにする
     * @param {Window} kinsokuDialog - 追加先のダイアログ
     * @param {Document} doc - 対象のドキュメント
     * @param {string} currentKinsoku - 初期選択にする禁則の値
     * @returns {RadioButton[]} 作成したラジオボタン
     */
    function addPresetPanel(kinsokuDialog, doc, currentKinsoku) {
        var presetPanel = kinsokuDialog.add("panel", undefined, getLabel("panel.preset"));
        presetPanel.orientation = "column";
        presetPanel.alignChildren = "left";
        presetPanel.alignment = "fill";
        presetPanel.margins = PANEL_MARGINS;
        presetPanel.spacing = PANEL_SPACING;

        var presetRadios = [];
        for (var i = 0; i < KINSOKU_PRESETS.length; i++) {
            var presetRadio = presetPanel.add("radiobutton", undefined, getLabel(KINSOKU_PRESETS[i].labelPath));
            presetRadio.kinsokuName = KINSOKU_PRESETS[i].kinsokuName;
            presetRadio.helpTip = getLabel("tooltip.preset");
            presetRadio.onClick = function () {
                applyKinsokuToSelection(doc, this.kinsokuName);
            };
            presetRadios.push(presetRadio);
        }

        /* 現在の禁則に一致するラジオを選ぶ（一致しなければ先頭＝「なし」）/ Check the matching radio, or the first ("None") */
        var checkedRadio = presetRadios[0];
        for (var j = 0; j < presetRadios.length; j++) {
            if (presetRadios[j].kinsokuName === currentKinsoku) {
                checkedRadio = presetRadios[j];
                break;
            }
        }
        checkedRadio.value = true;

        return presetRadios;
    }

    /**
     * ［閉じる］［OK］のボタン行を追加する
     * @param {Window} kinsokuDialog - 追加先のダイアログ
     * @param {Document} doc - 対象のドキュメント
     * @param {RadioButton[]} presetRadios - プリセットのラジオボタン
     * @returns {void}
     */
    function addButtonRow(kinsokuDialog, doc, presetRadios) {
        var btnRowGroup = kinsokuDialog.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.alignChildren = ["left", "center"];
        btnRowGroup.alignment = "right";
        btnRowGroup.spacing = BUTTON_SPACING;

        var btnClose = btnRowGroup.add("button", undefined, getLabel("button.close"));
        var btnOK = btnRowGroup.add("button", undefined, getLabel("button.ok"));
        btnClose.helpTip = getLabel("tooltip.close");

        btnClose.onClick = function () {
            kinsokuDialog.close(0);
        };

        btnOK.onClick = function () {
            for (var i = 0; i < presetRadios.length; i++) {
                if (presetRadios[i].value) {
                    applyKinsokuToSelection(doc, presetRadios[i].kinsokuName);
                    break;
                }
            }
            kinsokuDialog.close(1);
        };
    }

    /**
     * 禁則処理を選ぶダイアログを表示する
     * @param {Document} doc - 対象のドキュメント
     * @returns {void}
     */
    function showKinsokuDialog(doc) {
        var kinsokuDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        kinsokuDialog.orientation = "column";
        kinsokuDialog.alignChildren = "fill";

        var presetRadios = addPresetPanel(kinsokuDialog, doc, getSelectionKinsoku(doc));
        addButtonRow(kinsokuDialog, doc, presetRadios);

        kinsokuDialog.center();
        prepareDialogWindow(kinsokuDialog, SCRIPT_NAME);
        kinsokuDialog.show();
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * ドキュメントと選択を確かめてダイアログを表示する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }

        var doc = app.activeDocument;
        if (doc.selection.length === 0) {
            alert(getLabel("alert.noSelection"));
            return;
        }

        showKinsokuDialog(doc);
    }

    main();

})();
