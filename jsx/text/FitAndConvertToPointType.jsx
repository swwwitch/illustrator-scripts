#target illustrator
#targetengine "FitAndConvertToPointTypeEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したエリア内文字のうち、あふれているものだけ自動サイズ調整をONにして解消してから、ポイント文字に変換します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/FitAndConvertToPointType.md

### Overview

Turns on auto-sizing for whichever of the selected area texts are overset, resolves the overflow, and then converts them to point text.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/FitAndConvertToPointType.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "FitAndConvertToPointType";     /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-08-20";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-28";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/FitAndConvertToPointType.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/FitAndConvertToPointType.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* 「強制改行を削除」の初期状態 / Initial state of "Remove forced line breaks" */
    var DEFAULT_REMOVE_LINE_BREAKS = false;

    /* 強制改行（ソフトリターン）と見なす文字。段落改行（\r）は残す
       Characters treated as forced line breaks (soft returns); paragraph returns (\r) are kept */
    var FORCED_BREAK_PATTERN = /[\u0003\n\u2028]/;

    // =========================================
    // レイアウト / Layout
    // =========================================

    var WINDOW_MARGINS     = 16;            /* ウィンドウ外周の余白 / window margin */
    var WINDOW_SPACING     = 12;            /* ウィンドウ内の要素間隔 / window spacing */
    var BUTTON_BAR_MARGINS = [0, 10, 0, 0]; /* ボタンバーの余白 / margins of the button bar */
    var BUTTON_BAR_SPACING = 10;            /* ボタンバー内の要素間隔 / spacing inside the button bar */

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
    // ローカライズ / Localization
    // =========================================

    /**
     * Illustrator の UI 言語から表示言語を判定する
     * @returns {string} "ja" または "en"
     */
    function detectUILanguage() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = detectUILanguage();

    /* 日英ラベル定義（カテゴリ別）/ Japanese-English labels grouped by category */
    var LABELS = {
        dialog: {
            title: { ja: "ポイント文字に変換", en: "Convert to Point Type" }
        },
        checkbox: {
            removeLineBreaks: { ja: "強制改行を削除", en: "Remove forced line breaks" }
        },
        tooltip: {
            removeLineBreaks: {
                ja: "変換後に残る強制改行（ソフトリターン）を取り除き、行をつなげます。段落改行は残ります。",
                en: "Strips the forced line breaks (soft returns) left after the conversion, joining the lines. Paragraph returns are kept."
            }
        },
        button: {
            ok:     { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            selectAreaText: { ja: "エリア内文字を選択してください。", en: "Please select area type." },
            noTarget: {
                ja: "変換できるエリア内文字がありませんでした。ロックや非表示になっていないか確認してください。",
                en: "No area type could be converted. Check whether the frames are locked or hidden."
            },
            notSupported: {
                ja: "お使いのIllustratorはポイント文字への変換に対応していません。",
                en: "This version of Illustrator cannot convert area type to point text."
            },
            partialFailure: {
                ja: "{done}件を変換しました。{failed}件は変換できませんでした（連結されたテキストなどは対象外です）。",
                en: "Converted {done}. {failed} could not be converted (threaded text and the like are not supported)."
            }
        }
    };

    /**
     * "category.key" 形式のキーからラベルを取得する
     * @param {string} labelPath - ラベルキー（例: "alert.noDocument"）
     * @returns {string} 現在の言語のラベル文字列（見つからない場合は labelPath をそのまま返す）
     */
    function getLabel(labelPath) {
        var labelPathKeys = labelPath.split(".");
        var labelNode = LABELS;
        for (var i = 0; i < labelPathKeys.length; i++) {
            if (!labelNode) break;
            labelNode = labelNode[labelPathKeys[i]];
        }
        if (!labelNode) return labelPath;
        if (typeof labelNode[uiLang] === "string") return labelNode[uiLang];
        return (typeof labelNode.en === "string") ? labelNode.en : labelPath;
    }

    // =========================================
    // ダイナミックアクション / Dynamic actions
    //   自動サイズ調整はDOMから設定できないため、アクション経由で切り替える
    //   Auto-sizing cannot be set from the DOM, so it is toggled through an action
    // =========================================

    /**
     * 文字列を ASCII 16進に変換する
     * @param {string} text - 変換する文字列
     * @returns {string} 16進文字列
     */
    function asciiToHex(text) {
        var hexText = "";
        for (var i = 0; i < text.length; i++) {
            var hexPair = text.charCodeAt(i).toString(16);
            if (hexPair.length < 2) hexPair = "0" + hexPair;
            hexText += hexPair;
        }
        return hexText;
    }

    /**
     * アクション名ブロック /name [ <len> <hex> ] を生成する
     * @param {string} actionName - アクション名またはセット名
     * @returns {string} 名前ブロックの文字列
     */
    function buildActionNameBlock(actionName) {
        return "/name [ " + actionName.length + " " + asciiToHex(actionName).toUpperCase() + " ]";
    }

    /**
     * 自動サイズ調整アクションセットの定義（.aia 文字列）を組み立てる
     * @param {string} setName - アクションセット名
     * @returns {string} .aia 形式のアクションセット定義
     */
    function buildAutoSizeActionSetAia(setName) {
        return "/version 3" +
            buildActionNameBlock(setName) +
            "/isOpen 1" +
            "/actionCount 1" +
            "/action-1 {" +
            " " + buildActionNameBlock("AutoSizeOn") +
            " /keyIndex 0" +
            " /colorIndex 0" +
            " /isOpen 1" +
            " /eventCount 1" +
            " /event-1 {" +
            " /useRulersIn1stQuadrant 0" +
            " /internalName (adobe_SLOAreaTextDialog)" +
            " /localizedName [ 33 e382a8e383aae382a2e58685e69687e5ad97e382aae38397e382b7e383a7e383b3 ]" +
            " /isOpen 0" +
            " /isOn 1" +
            " /hasDialog 0" +
            " /parameterCount 1" +
            " /parameter-1 {" +
            " /key 1952539754" +
            " /showInPalette 4294967295" +
            " /type (integer)" +
            " /value 1" +
            " }" +
            " }" +
            "}";
    }

    var ACTION_SET_AUTO_SIZE = SCRIPT_NAME + "_AutoSize";

    /**
     * 自動サイズ調整アクションを一時ファイル経由で読み込む（スクリプト開始時に1回。既存があれば先に外す）
     * @returns {void}
     */
    function loadAutoSizeAction() {
        unloadAutoSizeAction();
        var tempFile = new File(Folder.temp + "/" + SCRIPT_NAME + "_action.aia");
        tempFile.open("w");
        tempFile.write(buildAutoSizeActionSetAia(ACTION_SET_AUTO_SIZE));
        tempFile.close();
        app.loadAction(tempFile);
        try { tempFile.remove(); } catch (e) { /* 一時ファイルの削除失敗は無視 / Ignore a failed temp cleanup */ }
    }

    /**
     * 読み込んだアクションセットを破棄する（スクリプト終了時）
     * @returns {void}
     */
    function unloadAutoSizeAction() {
        /* 読み込まれていなければ例外になる / Throws when the set is not loaded */
        try { app.unloadAction(ACTION_SET_AUTO_SIZE, ""); } catch (e) { }
    }

    /**
     * 選択中のエリア内文字に自動サイズ調整をかける（アクションは選択に効く）
     * @returns {void}
     */
    function runAutoSizeAction() {
        try { app.doScript("AutoSizeOn", ACTION_SET_AUTO_SIZE, false); } catch (e) { }
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * グループを横並びの行にする（揃えは横と天地を対で指定する）
     * @param {Group} targetGroup - 対象のグループ
     * @param {string} [horizontalAlign] - 行の横方向の揃え（既定は "left"）
     * @param {number} [spacing] - 要素間隔（既定は WINDOW_SPACING）
     * @returns {void}
     */
    function setupRow(targetGroup, horizontalAlign, spacing) {
        targetGroup.orientation = "row";
        targetGroup.alignment = [horizontalAlign || "left", "center"];
        targetGroup.alignChildren = ["left", "center"];
        targetGroup.spacing = (typeof spacing === "number") ? spacing : WINDOW_SPACING;
    }

    /**
     * オプションダイアログを表示する
     * @returns {{removeLineBreaks: boolean}|null} 設定（キャンセル時は null）
     */
    function showOptionsDialog() {
        var optionsDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        optionsDialog.orientation = "column";
        optionsDialog.alignChildren = ["fill", "top"];
        optionsDialog.spacing = WINDOW_SPACING;
        optionsDialog.margins = WINDOW_MARGINS;

        var optionRow = optionsDialog.add("group");
        setupRow(optionRow);
        var removeLineBreaksCheckbox = optionRow.add("checkbox", undefined, getLabel("checkbox.removeLineBreaks"));
        removeLineBreaksCheckbox.value = DEFAULT_REMOVE_LINE_BREAKS;
        removeLineBreaksCheckbox.helpTip = getLabel("tooltip.removeLineBreaks");

        /* ボタンバー（Mac 規約で キャンセル → OK）/ Button bar (Cancel → OK per macOS) */
        var btnRowGroup = optionsDialog.add("group");
        setupRow(btnRowGroup, "right", BUTTON_BAR_SPACING);
        btnRowGroup.margins = BUTTON_BAR_MARGINS;
        btnRowGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        btnRowGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });

        prepareDialogWindow(optionsDialog, SCRIPT_NAME);
        if (optionsDialog.show() !== 1) return null;
        return { removeLineBreaks: removeLineBreaksCheckbox.value };
    }

    // =========================================
    // 変換 / Conversion
    // =========================================

    /**
     * エリア内文字があふれているかを返す
     * @param {TextFrame} areaTextFrame - 対象のエリア内文字
     * @returns {boolean} あふれていれば true
     */
    function isFrameOverset(areaTextFrame) {
        /* overflows を持たない環境がある / Some versions lack overflows */
        try {
            if (typeof areaTextFrame.overflows !== "undefined") return !!areaTextFrame.overflows;
        } catch (e) { }
        /* overflows が読めない環境では、行に入っている文字数と全文字数を突き合わせる
           Where overflows cannot be read, the characters inside the lines are counted against the total */
        try {
            var visibleCount = 0;
            for (var i = 0; i < areaTextFrame.lines.length; i++) {
                visibleCount += areaTextFrame.lines[i].characters.length;
            }
            return visibleCount < areaTextFrame.characters.length;
        } catch (e0) { return false; }
    }

    /**
     * 変換できるエリア内文字かを返す（ロック・非表示のものは対象外）
     * @param {PageItem} pageItem - 調べるオブジェクト
     * @returns {boolean} 変換できれば true
     */
    function isConvertibleAreaTextFrame(pageItem) {
        try {
            if (!pageItem || pageItem.typename !== "TextFrame" || pageItem.kind !== TextType.AREATEXT) return false;
            if (pageItem.locked || pageItem.hidden) return false;
            var ownerLayer = pageItem.layer;
            if (ownerLayer && (ownerLayer.locked || !ownerLayer.visible)) return false;
        } catch (e) { return false; }
        return true;
    }

    /**
     * 強制改行を1文字ずつ取り除く（contents の一括置換は文字ごとの書式を失うため使わない）
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @returns {void}
     */
    function removeForcedLineBreaks(textFrame) {
        var textCharacters;
        try { textCharacters = textFrame.characters; } catch (e) { return; }
        /* 削除するとインデックスがずれるので後ろから処理する / Deleting shifts the indices, so walk from the back */
        for (var i = textCharacters.length - 1; i >= 0; i--) {
            /* 読めない文字・消せない文字は飛ばす / Skip characters that cannot be read or removed */
            try {
                if (FORCED_BREAK_PATTERN.test(textCharacters[i].contents)) textCharacters[i].remove();
            } catch (eChar) { }
        }
    }

    /**
     * 選択からエリア内文字だけを拾う
     * @param {Array} selectedItems - ドキュメントの選択内容
     * @returns {TextFrame[]} 変換できるエリア内文字
     */
    function collectAreaTextFrames(selectedItems) {
        var areaTextFrames = [];
        for (var i = 0; i < selectedItems.length; i++) {
            if (isConvertibleAreaTextFrame(selectedItems[i])) { areaTextFrames.push(selectedItems[i]); }
        }
        return areaTextFrames;
    }

    /**
     * あふれているフレームだけ自動サイズ調整で解消してから、ポイント文字へ変換する
     * @param {Document} doc - 対象ドキュメント
     * @param {TextFrame[]} areaTextFrames - 変換するエリア内文字
     * @param {boolean} removeLineBreaks - 変換後に強制改行を取り除くか
     * @returns {{converted: TextFrame[], failedCount: number}} 変換できたポイント文字と、変換できなかった数
     */
    function convertAreaTextFramesToPointText(doc, areaTextFrames, removeLineBreaks) {
        /* あふれていないフレームに自動サイズ調整をかけると、テキストの配置が中央・下のときに文字が動くため
           Auto-sizing a frame with no overset would move the text when it is centred or bottom-aligned
           convertAreaObjectToPointObject() は「その場変換・戻り値 null」で、変換後は元の参照が stale になり
           kind を AREATEXT のまま報告することがある。そこで変換前に目印の名前を付け、変換後に
           doc.textFrames を1回だけ走査して回収する
           (the API converts in place and returns null, and the old wrappers can go stale and still report
            AREATEXT, so the frames are tagged with a marker name first and collected afterwards
            by a single fresh scan of doc.textFrames) */
        var markerPrefix = "__" + SCRIPT_NAME + "_marker_";
        var previousNames = [];

        for (var i = 0; i < areaTextFrames.length; i++) {
            var previousName = "";
            try { previousName = areaTextFrames[i].name; } catch (eRead) { }
            previousNames.push(previousName);
            try { areaTextFrames[i].name = markerPrefix + i; } catch (eTag) { }
        }

        /* 変換するとオブジェクトが置き換わるので後ろから処理する
           Each frame is replaced as it goes, so the list is walked from the back */
        for (var j = areaTextFrames.length - 1; j >= 0; j--) {
            /* 1フレームで失敗しても残りを処理できるようにする / One failing frame must not stop the rest */
            try {
                if (isFrameOverset(areaTextFrames[j])) {
                    doc.selection = null;
                    areaTextFrames[j].selected = true;
                    app.redraw(); /* Illustratorに選択状態を確定させる / Let Illustrator settle the selection */
                    runAutoSizeAction();
                }
                areaTextFrames[j].convertAreaObjectToPointObject();
            } catch (e) { }
        }

        /* 目印で回収して名前を戻す（走査は1回だけ）/ Collect by marker and restore the names (a single scan) */
        var convertedFrames = [], failedCount = 0;
        var allTextFrames = doc.textFrames;
        for (var k = 0; k < allTextFrames.length; k++) {
            var frameName = "";
            try { frameName = allTextFrames[k].name; } catch (eName) { continue; }
            if (frameName.indexOf(markerPrefix) !== 0) continue;

            var markerIndex = parseInt(frameName.substring(markerPrefix.length), 10);
            var restoredName = (!isNaN(markerIndex) && typeof previousNames[markerIndex] === "string")
                ? previousNames[markerIndex] : "";
            try { allTextFrames[k].name = restoredName; } catch (eRestore) { }
            if (allTextFrames[k].kind === TextType.POINTTEXT) {
                /* 変換すると折り返し位置に強制改行が残るので、ポイント文字になってから取り除く
                   The conversion leaves a forced break at every wrap, so they are stripped once it is point text */
                if (removeLineBreaks) removeForcedLineBreaks(allTextFrames[k]);
                convertedFrames.push(allTextFrames[k]);
            } else { failedCount++; }
        }
        return { converted: convertedFrames, failedCount: failedCount };
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択を確かめて変換する（ダイナミックアクションは終了時に必ずアンロードする）
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }

        var doc = app.activeDocument;
        var selectedItems = doc.selection;

        /* 文字ツールで文字を選択中は TextRange が返り、length が文字数になるため配列かどうかで判定する
           With the type tool the selection is a TextRange whose length counts characters */
        if (!(selectedItems instanceof Array) || selectedItems.length === 0) {
            alert(getLabel("alert.selectAreaText"));
            return;
        }

        var areaTextFrames = collectAreaTextFrames(selectedItems);
        if (!areaTextFrames.length) {
            alert(getLabel("alert.noTarget"));
            return;
        }

        /* 古いIllustratorには変換APIが無いので、何も触らずに知らせる / Older Illustrator has no conversion API, so nothing is touched */
        var supportsConversion = false;
        try { supportsConversion = !!areaTextFrames[0].convertAreaObjectToPointObject; } catch (eApi) { }
        if (!supportsConversion) {
            alert(getLabel("alert.notSupported"));
            return;
        }

        var dialogOptions = showOptionsDialog();
        if (!dialogOptions) return;

        loadAutoSizeAction();
        var conversionResult;
        try {
            conversionResult = convertAreaTextFramesToPointText(doc, areaTextFrames, dialogOptions.removeLineBreaks);
        } finally {
            unloadAutoSizeAction();
        }

        /* 変換後のポイント文字を選び直す（削除済みの参照を選択に残さない）
           Re-select the resulting point text, so no stale reference lingers in the selection */
        try { doc.selection = conversionResult.converted.length ? conversionResult.converted : null; } catch (eSelect) { }
        app.redraw();

        if (!conversionResult.converted.length) {
            alert(getLabel("alert.noTarget"));
        } else if (conversionResult.failedCount > 0) {
            alert(getLabel("alert.partialFailure")
                .replace("{done}", conversionResult.converted.length)
                .replace("{failed}", conversionResult.failedCount));
        }
    }

    main();

})();
