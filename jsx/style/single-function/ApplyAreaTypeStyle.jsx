#target illustrator
#targetengine "ApplyAreaTypeStyleEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

ダイアログのラジオボタンでグラフィックスタイルを選び、選択中のオブジェクトへ適用します。
選んだスタイルがドキュメントに無ければ、あらかじめ定義したAIファイルから取り込みます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ApplyAreaTypeStyle.md

### Overview

Picks a graphic style with radio buttons in a dialog and applies it to the selection.
If the chosen style is not in the document, it is imported from a predefined AI file.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ApplyAreaTypeStyle.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ApplyAreaTypeStyle";          /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.6.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-08-14";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-28";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ApplyAreaTypeStyle.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ApplyAreaTypeStyle.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* スタイルを取り込む AI ファイル（環境に合わせて変更）/ AI file to import styles from */
    var TARGET_FILE_PATH = "/Users/takano/sw Dropbox/takano masahiro/sync-setting/ai/styles/StyleForAreaType.ai";

    /* グラフィックスタイル名（AIファイル内の登録名に合わせる）/ Graphic style names as registered in the AI file */
    var STYLE_NAME_WHITE_TEXT = "文字白抜き"; // ラジオ「文字白抜き」/ Radio "White text"
    var STYLE_NAME_FRAME_ONLY = "枠のみ";     // ラジオ「枠のみ」/ Radio "Frame only"

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /* 現在のロケールから言語コードを取得 / Get language code from locale */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var currentLanguage = getCurrentLang();

    var LABELS = {
        /* ダイアログ / Dialog */
        dialog: {
            title: { ja: "スタイルを適用", en: "Apply Style" }
        },
        /* スタイル選択 / Style choices */
        style: {
            panelTitle: { ja: "スタイル", en: "Style" },
            whiteText: { ja: "文字白抜き", en: "White text" },
            frameOnly: { ja: "枠のみ", en: "Frame only" }
        },
        tooltip: {
            whiteText: { ja: "背景を塗りつぶし、文字を白抜きにするスタイルを適用します。", en: "Applies the style that fills the background and knocks the text out in white." },
            frameOnly: { ja: "塗りは付けず、枠線だけのスタイルを適用します。", en: "Applies the style that draws only a frame, with no fill." }
        },
        /* ボタン / Buttons */
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            apply: { ja: "適用", en: "Apply" }
        },
        /* 実行時メッセージ / Runtime messages */
        message: {
            openDocFirst: {
                ja: "元のドキュメントを開いてから実行してください。",
                en: "Please open the destination document first."
            },
            selectObjectFirst: {
                ja: "スタイルを適用するオブジェクトを選択してから実行してください。",
                en: "Please select the object(s) to apply the style to."
            },
            fileNotFoundTitle: { ja: "ファイルが見つかりません", en: "File Not Found" },
            fileNotFoundBody: {
                ja: "指定したファイルが見つかりません：\n",
                en: "The specified file was not found:\n"
            },
            styleNotFound: {
                ja: "指定したグラフィックスタイルが見つかりません：\n",
                en: "The graphic style was not found:\n"
            }
        }
    };

    /* 指定キーのローカライズ文字列を取得（ドット区切りパス対応）/ Resolve localized string by dotted key path */
    function getLabel(key) {
        var parts = key.split(".");
        var node = LABELS;
        for (var i = 0; i < parts.length; i++) {
            if (node == null) break;
            node = node[parts[i]];
        }
        return (node && node[currentLanguage] != null) ? node[currentLanguage] : ("[" + key + "]");
    }

    // =========================================
    // ユーティリティ / Utilities
    // =========================================

    /* 取り込み先レイヤーを取得（無ければ作成）/ Get or create the destination layer "// _imported" */
    function getOrCreateImportLayer(destinationDoc) {
        var importLayerName = "// _imported";
        var importLayer;
        try {
            importLayer = destinationDoc.layers.getByName(importLayerName);
        } catch (e) {
            importLayer = destinationDoc.layers.add();
            importLayer.name = importLayerName;
        }
        importLayer.locked = false;
        importLayer.visible = true;
        return importLayer;
    }

    /* パスから表示用ファイル名を取得（文字化け対策で decodeURI）/ Get a display-friendly filename from a path */
    function getDisplayFileName(filePath) {
        try {
            return decodeURI(new File(filePath).name);
        } catch (e) {
            return filePath;
        }
    }

    /* 名前でグラフィックスタイルを取得（無ければ null）/ Get a graphic style by name (null if none) */
    function findGraphicStyle(destinationDoc, styleName) {
        try {
            return destinationDoc.graphicStyles.getByName(styleName);
        } catch (e) {
            return null;
        }
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

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

    /* スタイルを選ばせ、対応するグラフィックスタイル名を返す（キャンセルで null）/ Ask which style, return its graphic-style name */
    function chooseGraphicStyleName() {
        var styleDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        styleDialog.orientation = "column";
        styleDialog.alignChildren = ["fill", "top"];

        // スタイル選択（パネル＋ラジオ）/ Style panel with radios
        var stylePanel = styleDialog.add("panel", undefined, getLabel("style.panelTitle"));
        stylePanel.orientation = "column";
        stylePanel.alignChildren = ["left", "top"];
        stylePanel.margins = [16, 20, 16, 12];
        stylePanel.spacing = 8;
        var whiteTextRadio = stylePanel.add("radiobutton", undefined, getLabel("style.whiteText"));
        whiteTextRadio.helpTip = getLabel("tooltip.whiteText");
        var frameOnlyRadio = stylePanel.add("radiobutton", undefined, getLabel("style.frameOnly"));
        frameOnlyRadio.helpTip = getLabel("tooltip.frameOnly");
        whiteTextRadio.value = true; // 既定 / Default

        // ボタン行（Mac 規約: キャンセル → 適用）/ Buttons (Mac order: Cancel → Apply)
        var dialogButtonGroup = styleDialog.add("group");
        dialogButtonGroup.alignment = ["right", "bottom"];
        dialogButtonGroup.alignChildren = ["right", "center"];
        dialogButtonGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        dialogButtonGroup.add("button", undefined, getLabel("button.apply"), { name: "ok" });

        prepareDialogWindow(styleDialog, SCRIPT_NAME);
        if (styleDialog.show() !== 1) return null;
        return frameOnlyRadio.value ? STYLE_NAME_FRAME_ONLY : STYLE_NAME_WHITE_TEXT;
    }

    // =========================================
    // スタイル取り込み / Import Styles
    // =========================================

    /* 対象AIを開いてコピー→一時レイヤーへ貼り付け→レイヤーごと削除（アセットのみ登録）/ Import assets from the target AI */
    function importStyles(destinationDoc) {
        var styleFile = new File(encodeURI(TARGET_FILE_PATH));
        if (!styleFile.exists) {
            alert(getLabel("message.fileNotFoundTitle") + "\n" + getLabel("message.fileNotFoundBody") + getDisplayFileName(TARGET_FILE_PATH));
            return false;
        }
        var styleSourceDoc = app.open(styleFile);

        // 作業アートボード内の全てをコピーして保存せず閉じる / Copy in-artboard objects, close without saving
        app.executeMenuCommand("selectallinartboard");
        app.executeMenuCommand("copy");
        styleSourceDoc.close(SaveOptions.DONOTSAVECHANGES);

        // 一時レイヤーへ貼り付け→パネル登録後にレイヤーごと削除 / Paste to temp layer, then remove it
        app.activeDocument = destinationDoc;
        var importLayer = getOrCreateImportLayer(destinationDoc);
        destinationDoc.activeLayer = importLayer;
        app.executeMenuCommand("paste");
        try {
            importLayer.remove();
        } catch (e) {}
        return true;
    }

    /* スタイルを取得（未登録なら取り込んでから取得）／見つからなければ警告して null / Get the style, importing it if needed */
    function ensureGraphicStyle(destinationDoc, styleName) {
        var graphicStyle = findGraphicStyle(destinationDoc, styleName);
        if (graphicStyle) return graphicStyle;

        if (!importStyles(destinationDoc)) return null; // importStyles 側でファイル未検出を警告 / importStyles alerts on missing file
        graphicStyle = findGraphicStyle(destinationDoc, styleName);
        if (!graphicStyle) alert(getLabel("message.styleNotFound") + styleName);
        return graphicStyle;
    }

    // =========================================
    // メイン処理 / Main Process
    // =========================================

    /* スタイル選択→（必要なら取り込み）→選択オブジェクトへ適用 / Choose style, import if needed, apply to selection */
    function main() {
        // ドキュメントの存在を確認 / Check for an open document
        if (app.documents.length === 0) {
            alert(getLabel("message.openDocFirst"));
            return;
        }
        var destinationDoc = app.activeDocument;

        // 選択オブジェクトを確保 / Capture current selection
        var currentSelection = destinationDoc.selection;
        if (!currentSelection || currentSelection.length === 0) {
            alert(getLabel("message.selectObjectFirst"));
            return;
        }
        var selectedItems = [];
        for (var i = 0; i < currentSelection.length; i++) selectedItems.push(currentSelection[i]);

        // 適用するスタイルを選択 / Choose the style to apply
        var styleName = chooseGraphicStyleName();
        if (!styleName) return;

        // スタイルを取得（未登録なら取り込む）/ Get the style (import if needed)
        var graphicStyle = ensureGraphicStyle(destinationDoc, styleName);
        if (!graphicStyle) return;

        // 選択オブジェクトへ適用 / Apply to the selected objects
        for (var j = 0; j < selectedItems.length; j++) {
            try {
                graphicStyle.applyTo(selectedItems[j]);
            } catch (e) {}
        }

        // 選択状態を復帰 / Restore selection
        destinationDoc.selection = selectedItems;
    }

    main();

})();
