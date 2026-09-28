#target illustrator
#targetengine "ReleaseClipMaskEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したクリッピングマスクをモード別に解除します。
単純解除、マスクパス削除、マスク内容削除の3つの方法から選べます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ReleaseClipMask.md

note記事も参照してください。
https://note.com/dtp_tranist/n/nebc832e574f7

### Overview

Releases the selected clipping masks in one of three modes.
You can release them plainly, delete the mask path, or delete the masked contents.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ReleaseClipMask.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ReleaseClipMask";              /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.2.5";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-06-06";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-29";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ReleaseClipMask.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ReleaseClipMask.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/nebc832e574f7"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /**
     * 起動時のデフォルト値。/ Default values on launch.
     *
     * @type {{releaseMode: string, applyFill: boolean}}
     */
    var USER_DEFAULTS = {
        releaseMode: "removePath", /* simple | removePath | removeContent */
        applyFill: true
    };

    /** パスへ適用する塗りの不透明度（%）/ Fill opacity applied to path (%) @type {number} */
    var PATH_FILL_OPACITY = 15;

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

    /** 日英ラベル定義（カテゴリ分け）/ Japanese-English labels (categorized) */
    var LABELS = {
        dialog: {
            title: { ja: "クリッピングマスクの解除", en: "Release Clipping Mask" },
            releaseMode: { ja: "解除方法", en: "Release Mode" }
        },
        radio: {
            simpleRelease: { ja: "単純に解除", en: "Simply release" },
            removeMaskPath: { ja: "マスク内容を残して、パスを削除", en: "Keep content, remove path" },
            removeContent: { ja: "パスを残して、マスク内容を削除", en: "Keep path, remove content" }
        },
        checkbox: {
            applyFill: {
                ja: "解除時、マスクパスにカラーを設定",
                en: "Set fill color to mask path when releasing"
            }
        },
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        tooltip: {
            simpleRelease: {
                ja: "マスクパスと内容の両方を残し、クリッピングマスクだけを解除します。",
                en: "Release the clipping mask while keeping both the mask path and the content."
            },
            removeMaskPath: {
                ja: "マスク用のパスを削除し、マスクされていた内容だけを残します。",
                en: "Delete the mask path and keep only the masked content."
            },
            removeContent: {
                ja: "マスクパス以外（配置画像・埋め込み画像・グループ・テキストなど）をすべて削除し、マスクパスだけを残します。",
                en: "Delete everything except the mask path (images, groups, text, ...) and keep only the mask path."
            },
            applyFill: {
                ja: "残ったマスクパスにK100・不透明度15%の塗りを設定します。",
                en: "Apply a K100 fill at 15% opacity to the remaining mask path."
            }
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noSelection: { ja: "オブジェクトが選択されていません。", en: "No objects selected." },
            noClippingMask: { ja: "選択範囲にクリッピングマスクがありません。", en: "No clipping mask in the selection." },
            releaseFailed: {
                ja: "{count}個のクリッピングマスクの解除に失敗しました。",
                en: "Failed to release {count} clipping mask(s)."
            }
        }
    };

    // =========================================
    // UIレイアウトの共通設定 / Shared UI layout
    // =========================================

    /* ウィンドウ・パネルの余白と間隔 / Window & panel margins and spacing */
    var WINDOW_MARGINS = 16;                 /* ウィンドウ外周の余白 / window margin */
    var WINDOW_SPACING = 12;                 /* ウィンドウ内の要素間隔 / window spacing */
    var PANEL_MARGINS  = [16, 20, 16, 12];   /* パネル余白 [左,上,右,下] / panel margins */
    var PANEL_SPACING  = 8;                  /* パネル内の要素間隔 / panel spacing */

    /**
     * ウィンドウ（ダイアログ）へ共通のレイアウト設定を適用します。/ Apply shared window layout.
     *
     * @param {Window} targetWindow - 設定対象のウィンドウ。
     * @param {number} [spacing] - 要素間隔。省略時は WINDOW_SPACING。
     * @returns {void}
     */
    function setupWindow(targetWindow, spacing) {
        targetWindow.orientation = "column";
        targetWindow.alignChildren = "fill";
        targetWindow.margins = WINDOW_MARGINS;
        targetWindow.spacing = (typeof spacing === "number") ? spacing : WINDOW_SPACING;
    }

    /**
     * パネルへ共通のレイアウト設定を適用します。/ Apply shared panel layout.
     *
     * @param {Panel} panel - 設定対象のパネル。
     * @param {number} [spacing] - 要素間隔。省略時は PANEL_SPACING。
     * @returns {void}
     */
    function setupPanel(panel, spacing) {
        panel.orientation = "column";
        panel.alignChildren = ["fill", "top"];
        panel.alignment = "fill";
        panel.margins = PANEL_MARGINS;
        panel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
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

    // =========================================
    // 選択・マスク判定ヘルパー / Selection & mask helpers
    // =========================================

    /**
     * ドキュメントの選択を安全に配列で取得します。/ Safely get the selection as an array.
     *
     * テキスト選択（TextRange）など配列でない場合は空配列を返します。
     *
     * @param {Document} doc - 対象ドキュメント。
     * @returns {Array<PageItem>} 選択アイテムの配列。取得できない場合は空配列。
     */
    function getSelectionItems(doc) {
        var selection = doc.selection;
        if (!selection || typeof selection.length !== "number" || selection.typename) {
            return [];
        }
        /* ライブ選択を実配列へコピー（以降のDOM変更の影響を受けない）
           / Copy the live selection into a real array so later DOM changes do not affect it */
        var selectionItems = [];
        for (var i = 0; i < selection.length; i++) {
            selectionItems.push(selection[i]);
        }
        return selectionItems;
    }

    /**
     * クリッピングマスクを持つグループかどうかを判定します。/ Check whether the item is a clipping-mask group.
     *
     * @param {PageItem} pageItem - 判定対象のページアイテム。
     * @returns {boolean} クリップグループなら true。
     */
    function isClippingGroup(pageItem) {
        return pageItem.typename === "GroupItem" && pageItem.clipped === true;
    }

    /**
     * 選択からクリップグループだけを抽出して配列で返します。/ Collect only clipping groups from the selection.
     *
     * @param {Array<PageItem>} selectionItems - 選択アイテムの配列。
     * @returns {Array<GroupItem>} クリップグループの配列。
     */
    function collectClippingGroups(selectionItems) {
        var clippingGroups = [];
        for (var i = 0; i < selectionItems.length; i++) {
            if (isClippingGroup(selectionItems[i])) {
                clippingGroups.push(selectionItems[i]);
            }
        }
        return clippingGroups;
    }

    /**
     * 指定アイテムがクリッピングマスクとして機能しているか判定します。/ Check whether the item acts as a clipping mask.
     *
     * PathItem は clipping、CompoundPathItem は内部 pathItems の clipping を確認します。
     *
     * @param {PageItem} pageItem - 判定対象のページアイテム。
     * @returns {boolean} マスクとして機能していれば true。
     */
    function isMaskItem(pageItem) {
        if (pageItem.typename === "PathItem") {
            return pageItem.clipping === true;
        }
        if (pageItem.typename === "CompoundPathItem") {
            var subPaths = pageItem.pathItems;
            for (var i = 0; i < subPaths.length; i++) {
                if (subPaths[i].clipping === true) {
                    return true;
                }
            }
        }
        return false;
    }

    /**
     * クリップグループ内で最初にマスクとして機能しているアイテムを返します。/ Return the first mask item in the group.
     *
     * @param {GroupItem} clippingGroup - 対象のクリップグループ。
     * @returns {PageItem|null} マスクアイテム。見つからなければ null。
     */
    function findMaskItem(clippingGroup) {
        for (var i = 0; i < clippingGroup.pageItems.length; i++) {
            var pageItem = clippingGroup.pageItems[i];
            if (isMaskItem(pageItem)) {
                return pageItem;
            }
        }
        return null;
    }

    /**
     * @typedef {Object} MaskRemovalPlan
     * @property {PageItem|null} maskItem - 最初に見つかったマスクアイテム。無ければ null。
     * @property {Array<PageItem>} itemsToRemove - マスク以外の子（削除対象）。
     */

    /**
     * マスクと削除対象を1回の走査で確定します。/ Determine the mask and removal targets in a single pass.
     *
     * 最初に見つかったマスクを残し、それ以外を削除対象とすることで参照比較（===）を避けます。
     *
     * @param {GroupItem} clippingGroup - 対象のクリップグループ。
     * @returns {MaskRemovalPlan} マスク参照と削除対象アイテムの一覧。
     */
    function collectContentToRemove(clippingGroup) {
        var maskItem = null;
        var itemsToRemove = [];
        for (var i = 0; i < clippingGroup.pageItems.length; i++) {
            var pageItem = clippingGroup.pageItems[i];
            if (!maskItem && isMaskItem(pageItem)) {
                maskItem = pageItem;
            } else {
                itemsToRemove.push(pageItem);
            }
        }
        return { maskItem: maskItem, itemsToRemove: itemsToRemove };
    }

    /**
     * クリップグループ内にマスク以外の削除対象（マスク内容）が存在するか判定します。/ Check for maskable content.
     *
     * @param {Array<GroupItem>} clippingGroups - 対象のクリップグループ配列。
     * @returns {boolean} 1つでもマスク以外の子があれば true。
     */
    function hasMaskedContent(clippingGroups) {
        for (var i = 0; i < clippingGroups.length; i++) {
            var plan = collectContentToRemove(clippingGroups[i]);
            if (plan.maskItem && plan.itemsToRemove.length > 0) {
                return true;
            }
        }
        return false;
    }

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
    // ダイアログ / Dialog
    // =========================================

    /**
     * 処理の起点。ドキュメント・選択・クリップグループを確認し、ダイアログ確定時に解除を実行します。
     * / Entry logic: validate, show dialog, and run release on confirm.
     *
     * @returns {void}
     */
    function main() {
        if (!app.documents.length) {
            alert(getLabel("alert.noDocument"));
            return;
        }

        var doc = app.activeDocument;
        var selectionItems = getSelectionItems(doc);
        if (!selectionItems.length) {
            alert(getLabel("alert.noSelection"));
            return;
        }

        var clippingGroups = collectClippingGroups(selectionItems);
        if (!clippingGroups.length) {
            alert(getLabel("alert.noClippingMask"));
            return;
        }

        var settings = showReleaseDialog(clippingGroups);
        if (!settings) return; /* キャンセル / Cancelled */

        executeRelease(settings.releaseMode, settings.applyFill, clippingGroups);
    }

    /**
     * @typedef {Object} ReleaseSettings
     * @property {string} releaseMode - 解除モード（"simple" | "removePath" | "removeContent"）。
     * @property {boolean} applyFill - 解除時に残ったマスクパスへ塗りを適用する場合は true。
     */

    /**
     * 解除方法・オプションを選ぶダイアログを構築し、選択結果を返します。/ Build the options dialog and return the choice.
     *
     * @param {Array<GroupItem>} clippingGroups - 対象のクリップグループ配列（有効/無効判定に使用）。
     * @returns {ReleaseSettings|null} OK 時は設定オブジェクト、キャンセル・クローズ時は null。
     */
    function showReleaseDialog(clippingGroups) {
        var dialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        setupWindow(dialog);

        /* 解除方法パネル / Release mode panel */
        var releaseModePanel = dialog.add("panel", undefined, getLabel("dialog.releaseMode"));
        setupPanel(releaseModePanel, 6);

        var radioSimpleRelease = releaseModePanel.add("radiobutton", undefined, getLabel("radio.simpleRelease"));
        var radioRemoveMaskPath = releaseModePanel.add("radiobutton", undefined, getLabel("radio.removeMaskPath"));
        var radioRemoveContent = releaseModePanel.add("radiobutton", undefined, getLabel("radio.removeContent"));

        radioSimpleRelease.helpTip = getLabel("tooltip.simpleRelease");
        radioRemoveMaskPath.helpTip = getLabel("tooltip.removeMaskPath");
        radioRemoveContent.helpTip = getLabel("tooltip.removeContent");

        /* デフォルトの解除方法を選択 / Select default release mode */
        if (USER_DEFAULTS.releaseMode === "simple") {
            radioSimpleRelease.value = true;
        } else if (USER_DEFAULTS.releaseMode === "removeContent") {
            radioRemoveContent.value = true;
        } else {
            radioRemoveMaskPath.value = true;
        }

        /* マスク以外の内容が無ければ「マスク内容を削除」を無効化 / Disable content removal when there is nothing to remove */
        if (!hasMaskedContent(clippingGroups)) {
            radioRemoveContent.enabled = false;
            if (radioRemoveContent.value) {
                radioRemoveMaskPath.value = true;
            }
        }

        /* 解除時に塗りカラーを適用するチェックボックス / Checkbox for fill color on release */
        var fillOptionGroup = dialog.add("group");
        fillOptionGroup.orientation = "column";
        fillOptionGroup.alignChildren = "left";
        fillOptionGroup.margins = [20, 5, 0, 10];
        var applyFillCheckbox = fillOptionGroup.add("checkbox", undefined, getLabel("checkbox.applyFill"));
        applyFillCheckbox.helpTip = getLabel("tooltip.applyFill");
        applyFillCheckbox.value = USER_DEFAULTS.applyFill;

        /* パス削除モードでは塗りチェックボックスを無効化 / Disable fill checkbox in remove-path mode */
        function updateFillCheckboxState() {
            applyFillCheckbox.enabled = !radioRemoveMaskPath.value;
        }
        radioSimpleRelease.onClick = updateFillCheckboxState;
        radioRemoveContent.onClick = updateFillCheckboxState;
        radioRemoveMaskPath.onClick = updateFillCheckboxState;
        updateFillCheckboxState();

        /* ボタン行 / Button row */
        var buttonRow = addButtonRow(dialog, { centered: true });
        var btnCancel = buttonRow.rowGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnOK = buttonRow.rowGroup.add("button", undefined, "OK", { name: "ok" });

        prepareDialogWindow(dialog, SCRIPT_NAME);
        /* OK以外（キャンセル・クローズ）は null / Return null unless OK was pressed */
        if (dialog.show() !== 1) return null;

        var selectedMode = "simple";
        if (radioRemoveMaskPath.value) {
            selectedMode = "removePath";
        } else if (radioRemoveContent.value) {
            selectedMode = "removeContent";
        }

        /* 無効時は値が true でも false を返す / Return false when the checkbox is disabled */
        return { releaseMode: selectedMode, applyFill: applyFillCheckbox.enabled && applyFillCheckbox.value };
    }

    // =========================================
    // 解除処理 / Release logic
    // =========================================

    /**
     * 選択モードに従って各クリップグループのマスクを解除します。/ Release masks per mode.
     *
     * @param {string} releaseMode - 解除モード（"simple" | "removePath" | "removeContent"）。
     * @param {boolean} shouldApplyFill - 残ったマスクパスへ塗りを適用する場合は true。
     * @param {Array<GroupItem>} clippingGroups - 対象のクリップグループ配列。
     * @returns {void}
     */
    function executeRelease(releaseMode, shouldApplyFill, clippingGroups) {
        /* グループ単位で try/catch し、途中失敗しても残りを処理して件数を通知
           / Isolate each group so one failure does not abort the rest */
        var failureReasons = [];
        for (var i = 0; i < clippingGroups.length; i++) {
            var clippingGroup = clippingGroups[i];
            try {
                if (releaseMode === "removePath") {
                    releaseKeepContent(clippingGroup);
                } else if (releaseMode === "removeContent") {
                    releaseKeepMaskPath(clippingGroup, shouldApplyFill);
                } else {
                    releaseKeepBoth(clippingGroup, shouldApplyFill);
                }
            } catch (err) {
                /* 失敗理由を保持して後でまとめて通知（ExtendScript は description のことがある）
                   / Keep the reason and report later (ExtendScript may use description) */
                failureReasons.push(err.message || err.description || String(err));
            }
        }
        if (failureReasons.length > 0) {
            var failureMessage = getLabel("alert.releaseFailed").replace("{count}", failureReasons.length);
            alert(failureMessage + "\n" + failureReasons.join("\n"));
        }
    }

    /**
     * マスクパスと内容を両方残し、マスクだけを解除します。/ Release the mask keeping both mask path and content.
     *
     * マスクを特定できない場合でも解除は行い、塗り適用のみスキップします（非破壊）。
     *
     * @param {GroupItem} clippingGroup - 対象のクリップグループ。
     * @param {boolean} shouldApplyFill - マスクパスへ塗りを適用する場合は true。
     * @returns {void}
     */
    function releaseKeepBoth(clippingGroup, shouldApplyFill) {
        var maskItem = findMaskItem(clippingGroup);
        clippingGroup.clipped = false;
        if (shouldApplyFill && maskItem) {
            applyMaskFill(maskItem);
        }
        ungroup(clippingGroup);
    }

    /**
     * マスク内容を残し、マスクパスを削除します。/ Keep masked content, remove the mask path.
     *
     * @param {GroupItem} clippingGroup - 対象のクリップグループ。
     * @returns {void}
     * @throws {Error} マスクを特定できない場合（このグループの処理を中止）。
     */
    function releaseKeepContent(clippingGroup) {
        /* 破壊操作の前にマスクを確定（clipping フラグ喪失を避ける）
           / Identify the mask before any destructive op */
        var maskItem = findMaskItem(clippingGroup);
        if (!maskItem) {
            throw new Error("Clipping mask item not found.");
        }
        clippingGroup.clipped = false;
        maskItem.remove();
        ungroup(clippingGroup);
    }

    /**
     * マスクパスを残し、マスク内容（画像・グループ・テキスト等）を全て削除します。/ Keep the mask path, remove all masked content.
     *
     * @param {GroupItem} clippingGroup - 対象のクリップグループ。
     * @param {boolean} shouldApplyFill - 残ったマスクパスへ塗りを適用する場合は true。
     * @returns {void}
     * @throws {Error} マスクを特定できない場合（全削除の誤動作を防ぐため処理を中止）。
     */
    function releaseKeepMaskPath(clippingGroup, shouldApplyFill) {
        /* 破壊操作の前にマスクと削除対象を確定 / Determine mask & targets before any destructive op */
        var plan = collectContentToRemove(clippingGroup);
        if (!plan.maskItem) {
            throw new Error("Clipping mask item not found.");
        }
        clippingGroup.clipped = false;
        removeItems(plan.itemsToRemove);
        if (shouldApplyFill) {
            applyMaskFill(plan.maskItem);
        }
        ungroup(clippingGroup);
    }

    /**
     * 収集済みのアイテムを順に削除します。/ Remove the collected items in order.
     *
     * @param {Array<PageItem>} items - 削除対象アイテムの配列。
     * @returns {void}
     */
    function removeItems(items) {
        for (var i = 0; i < items.length; i++) {
            items[i].remove();
        }
    }

    /**
     * マスクパスにK100・指定不透明度の塗りを適用します。/ Apply K100 fill at the given opacity to the mask path.
     *
     * PathItem はそのまま、CompoundPathItem は内部 pathItems へ塗りを設定します。
     * それ以外の型（テキストマスク等）は対象外です。
     *
     * @param {PageItem} maskItem - 塗りを適用するマスクアイテム。
     * @returns {void}
     */
    function applyMaskFill(maskItem) {
        if (!maskItem) return;
        if (maskItem.typename === "PathItem") {
            maskItem.filled = true;
            maskItem.fillColor = getK100Black();
            maskItem.opacity = PATH_FILL_OPACITY;
        } else if (maskItem.typename === "CompoundPathItem") {
            maskItem.opacity = PATH_FILL_OPACITY;
            var subPaths = maskItem.pathItems;
            for (var i = 0; i < subPaths.length; i++) {
                subPaths[i].filled = true;
                subPaths[i].fillColor = getK100Black();
            }
        }
    }

    // =========================================
    // ユーティリティ / Utilities
    // =========================================

    /**
     * K100（スミベタ）の CMYK カラーを生成して返します。/ Return a K100 CMYK color.
     *
     * @returns {CMYKColor} K=100、他0の CMYK カラー。
     */
    function getK100Black() {
        var cmykColor = new CMYKColor();
        cmykColor.cyan = 0;
        cmykColor.magenta = 0;
        cmykColor.yellow = 0;
        cmykColor.black = 100;
        return cmykColor;
    }

    /**
     * グループを解除します。子要素をグループの直前へ移動して重なり順を保ちます。/ Ungroup, preserving stacking order.
     *
     * 子を親末尾（PLACEATEND）へ送ると周囲との前後関係が崩れるため、グループ自身を基準に
     * PLACEBEFORE で移動し、元のグループ位置・内部順序を維持します。
     *
     * @param {GroupItem} clippingGroup - 解除対象のグループ。
     * @returns {void}
     */
    function ungroup(clippingGroup) {
        while (clippingGroup.pageItems.length > 0) {
            clippingGroup.pageItems[0].move(clippingGroup, ElementPlacement.PLACEBEFORE);
        }
        clippingGroup.remove();
    }

    // =========================================
    // エントリーポイント / Entry point
    // =========================================

    main();

})();
