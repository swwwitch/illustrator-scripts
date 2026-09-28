#target illustrator
#targetengine "FontSamplerEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

入力したテキストを、インストールされているフォントで並べて表示します。
30個ぶん／アートボードいっぱい／すべてのフォント、の3モードから選べます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/FontSampler.md

### Overview

Lays out the text you type in each of the installed fonts.
Three modes are available: 30 fonts, as many as fit the artboard, or every font.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/FontSampler.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "FontSampler";                  /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.2";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-08-06";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-28";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/FontSampler.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/FontSampler.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * 現在のUI言語を判定する
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = getCurrentLang();

    /* カテゴリ分けした日英ラベル定義 / Categorized Japanese-English label definitions */
    var LABELS = {
        fieldLabel: {
            sampleText: { ja: "テキスト", en: "Text" }
        },
        panel: {
            fontCount: { ja: "フォント数の設定", en: "How many fonts" }
        },
        radio: {
            limit30: { ja: "30個まで", en: "Up to 30" },
            fitBoard: { ja: "アートボードいっぱい", en: "Fill the artboard" },
            all: { ja: "すべて", en: "All" }
        },
        tooltip: {
            sampleText: { ja: "各フォントの見本として並べる文字です。", en: "The text shown as the specimen for each font." },
            limit30: { ja: "先頭から30書体までを並べます。", en: "Lays out the first 30 typefaces." },
            fitBoard: { ja: "アートボードに収まる数だけ並べます。", en: "Lays out as many as fit on the artboard." },
            all: { ja: "環境にあるすべての書体を並べます。数が多いと時間がかかります。", en: "Lays out every typeface on the machine. This can take a while." }
        },
        defaultSampleText: { ja: "山路を登りながら", en: "Handgloves" }
    };

    /**
     * ラベルを取得する（ドット区切りキー）
     * @param {string} labelPath - "panel.fontCount" のようなドット区切りキー
     * @returns {string} 現在のUI言語のラベル（見つからなければキーそのもの）
     */
    function getLabel(labelPath) {
        var pathKeys = String(labelPath).split(".");
        var labelNode = LABELS;
        for (var i = 0; i < pathKeys.length; i++) {
            labelNode = labelNode[pathKeys[i]];
            if (!labelNode) return labelPath;
        }
        return (labelNode[uiLang] != null) ? labelNode[uiLang] : labelPath;
    }

    /**
     * 項目名にコロンを付ける（日本語は全角、英語は半角）
     * @param {string} labelPath - ラベルのドット区切りキー
     * @returns {string} コロン付きの項目名
     */
    function labelText(labelPath) {
        return getLabel(labelPath) + (uiLang === "ja" ? "：" : ": ");
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

    function main() {
        try {
            if (app.documents.length === 0) {
                alert("ドキュメントを開いてください。");
                return;
            }

            var doc = app.activeDocument;

            // ダイアログ作成 / Create dialog
            var dlg = new Window("dialog", "テキストを入力");
            dlg.orientation = "column";
            dlg.alignChildren = ["fill", "top"];

            var inputGroup = dlg.add("group");
            inputGroup.add("statictext", undefined, labelText("fieldLabel.sampleText"));
            var inputText = inputGroup.add("edittext", undefined, getLabel("defaultSampleText"), {
                multiline: false
            });
            inputText.helpTip = getLabel("tooltip.sampleText");
            inputText.characters = 30;

            // 新しいパネルにラジオボタンを追加 / Add radio buttons in new panel
            var optionPanel = dlg.add("panel", undefined, getLabel("panel.fontCount"));
            optionPanel.orientation = "column";
            optionPanel.alignChildren = ["left", "top"];
            optionPanel.margins = [15, 20, 15, 10];
            var rb30 = optionPanel.add("radiobutton", undefined, getLabel("radio.limit30"));
            rb30.helpTip = getLabel("tooltip.limit30");
            var rbBoard = optionPanel.add("radiobutton", undefined, getLabel("radio.fitBoard"));
            rbBoard.helpTip = getLabel("tooltip.fitBoard");
            var rbAll = optionPanel.add("radiobutton", undefined, getLabel("radio.all"));
            rbAll.helpTip = getLabel("tooltip.all");
            rb30.value = true; // デフォルト / Default

            dlg.onShow = function() {
                inputText.active = true;
            };

            var btnGroup = dlg.add("group");
            btnGroup.alignment = "center";
            var cancelBtn = btnGroup.add("button", undefined, "キャンセル");
            var okBtn = btnGroup.add("button", undefined, "OK");

            prepareDialogWindow(dlg, SCRIPT_NAME);
            if (dlg.show() != 1) return; // キャンセル時 / Cancelled

            var textContent = inputText.text;
            if (!textContent || textContent === "") {
                alert("テキストを入力してください。");
                return;
            }

            // 使用可能フォントを取得 / Get available fonts
            var fonts = app.textFonts;

            // アートボードのサイズを取得 / Get artboard size
            var ab = doc.artboards[doc.artboards.getActiveArtboardIndex()].artboardRect;
            var abLeft = ab[0];
            var abTop = ab[1];
            var abRight = ab[2];
            var abBottom = ab[3];

            // レイアウト設定 / Layout settings
            var startX = abLeft + 50;
            var startY = abTop - 50;
            var lineHeight = 30; // 行間 / line height

            // 1つ仮のテキストフレームを作成して幅を測る / Create temp text frame to measure width
            var tempTF = doc.textFrames.add();
            tempTF.contents = textContent;
            tempTF.textRange.characterAttributes.size = 12; // 標準サイズ / standard size
            var sampleWidth = tempTF.width;
            tempTF.remove();

            // 列幅を文字幅に余白を加えて算出 / Calculate column width with margin
            var colWidth = sampleWidth + 50; // 40ptを余白として加算 / add 40pt margin

            var rowsPerCol = Math.floor((abTop - abBottom - 100) / lineHeight);
            var colsPerBoard = 2; // 固定で2カラム
            var maxCells = rowsPerCol * colsPerBoard;

            var mode = rb30.value ? "30" : rbBoard.value ? "board" : "all";

            var maxCount;
            if (mode === "30") {
                maxCount = Math.min(fonts.length, 30);
            } else if (mode === "board") {
                maxCount = Math.min(fonts.length, maxCells);
            } else {
                maxCount = fonts.length;
            }

            var progressWin;
            var progressBar;
            if (mode === "all") {
                progressWin = new Window("palette", "進行状況");
                progressWin.alignChildren = "fill";
                progressBar = progressWin.add("progressbar", undefined, 0, maxCount);
                progressBar.preferredSize = [300, 20];
                progressWin.show();
            }

            var col = 0;
            var row = 0;
            var artIndex = doc.artboards.getActiveArtboardIndex();

            for (var i = 0; i < maxCount; i++) {
                // 新規アートボード作成判定（「すべて」モード） / Check if new artboard needed (mode: All)
                if (mode === "all" && i > 0 && i % maxCells === 0) {
                    artIndex++;
                    // 新規アートボードを追加し位置をずらす / Add new artboard and offset its position
                    var abWidth = abRight - abLeft;
                    var abHeight = abTop - abBottom;
                    var offsetX = (artIndex) * (abWidth + 100); // 100pt間隔で横にずらす / offset horizontally by 100pt
                    var newRect = [abLeft + offsetX, abTop, abRight + offsetX, abBottom];
                    var newAB = doc.artboards.add(newRect);
                    doc.artboards.setActiveArtboardIndex(artIndex);
                    ab = doc.artboards[artIndex].artboardRect;
                    abLeft = ab[0];
                    abTop = ab[1];
                    abRight = ab[2];
                    abBottom = ab[3];
                    startX = abLeft + 50;
                    startY = abTop - 50;
                    col = 0;
                    row = 0;
                }

                var tf = doc.textFrames.add();
                tf.contents = textContent;
                var x = startX + (col * colWidth);
                var y = startY - (row * lineHeight);
                tf.position = [x, y];

                try {
                    tf.textRange.characterAttributes.textFont = fonts[i];
                } catch (e) {}

                row++;
                if (row >= rowsPerCol) {
                    row = 0;
                    col++;
                }
                if (mode === "all" && progressBar) {
                    progressBar.value = i + 1;
                    progressWin.update();
                }
            }

            if (mode === "all" && progressWin) {
                progressWin.close();
            }

        } catch (e) {
            alert("エラー: " + e);
        }
    }

    main();

})();
