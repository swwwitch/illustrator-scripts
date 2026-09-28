#target illustrator
#targetengine "AdjustBaselineVerticalCenterEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

指定した文字を、基準文字の中心に合わせてベースラインシフトで上下に動かします。
複数のテキストフレームへまとめて適用できます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AdjustBaselineVerticalCenter.md

note記事も参照してください。
https://note.com/dtp_tranist/n/na7a8c907c68c

### Overview

Shifts the specified characters up or down with a baseline shift so they line up with the center of a reference character.
It can be applied to several text frames at once.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AdjustBaselineVerticalCenter.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AdjustBaselineVerticalCenter"; /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.8";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-07-04";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-28";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AdjustBaselineVerticalCenter.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AdjustBaselineVerticalCenter.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/na7a8c907c68c"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

/**
 * @discussion 参考、謝辞 / Reference and acknowledgements
 * Egor Chistyakov (@tchegr)
 * https://x.com/tchegr
 */

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    var DEFAULT_REFERENCE_CHAR = "0"; /* 基準文字の初期値 / Initial reference character */

    // =========================================
    // レイアウト / Layout
    // =========================================

    var INPUT_GROUP_MARGINS = [15, 5, 15, 5]; /* 入力欄グループの余白 [左,上,右,下] / Input group margins */
    var CHAR_INPUT_CHARACTERS = 5;            /* 文字入力欄の幅（文字数）/ Width of the character fields (in characters) */

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
     * 実行環境の言語を判定する
     * @returns {string} "ja" または "en"
     */
    function detectUILanguage() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = detectUILanguage();

    var LABELS = {
        dialog: {
            title: { ja: "ベースライン調整", en: "Adjust Baseline" },
            description: { ja: "対象文字を縦方向に揃えます。", en: "This will align selected symbol vertically." }
        },
        fieldLabel: {
            targetChar: { ja: "対象文字", en: "Target Character" },
            referenceChar: { ja: "基準文字", en: "Reference Character" }
        },
        tooltip: {
            targetChar: {
                ja: "入力した文字をベースラインシフトで上下に動かします（複数可）。初期値は、選択中のテキストで英数字と空白を除いていちばん多い文字です。",
                en: "Moves these characters up or down with baseline shift (you can enter several). Defaults to the most frequent character in the selection, excluding letters, digits and spaces."
            },
            referenceChar: {
                ja: "対象文字の天地中央を、この文字の天地中央にそろえます（1文字）。",
                en: "Aligns the vertical center of the target characters with the vertical center of this character (one character)."
            }
        },
        button: {
            adjust: { ja: "調整", en: "Adjust" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document open." },
            selectTextFrame: { ja: "テキストフレームを選択してください。", en: "Select one or more text frames." },
            invalidChars: {
                ja: "対象文字は1文字以上、基準文字は1文字を入力してください。",
                en: "Enter at least one target character and exactly one reference character."
            },
            errorPrefix: { ja: "エラー: ", en: "Error: " }
        }
    };

    /**
     * 現在の言語のラベルを取得する
     * @param {object} labelSet - { ja: string, en: string } 形式のラベル
     * @returns {string} ラベル文字列
     */
    function getLabel(labelSet) {
        return labelSet[uiLang] || labelSet.en;
    }

    /**
     * 項目名にコロンを付けて返す（日本語は全角、英語は半角）
     * @param {object} labelSet - { ja: string, en: string } 形式のラベル
     * @returns {string} コロン付きのラベル文字列
     */
    function labelText(labelSet) {
        return getLabel(labelSet) + (uiLang === "ja" ? "：" : ":");
    }

    // =========================================
    // 文字の中心の実測 / Measuring character centers
    // =========================================

    /**
     * アイテムの天地中央のY座標を求める
     * @param {PageItem} pageItem - 対象のアイテム
     * @returns {number} 天地中央のY座標
     */
    function getCenterY(pageItem) {
        var itemBounds = pageItem.geometricBounds;
        return (itemBounds[1] + itemBounds[3]) / 2;
    }

    /**
     * テキストフレームを複製して1文字だけにし、アウトライン化した字形の天地中央を測る
     * @param {TextFrame} textFrame - 書式の元になるテキストフレーム
     * @param {string} character - 測る文字
     * @returns {number} 字形の天地中央のY座標
     */
    function measureCharCenterY(textFrame, character) {
        var tempFrame = textFrame.duplicate();
        var outlineGroup = null;
        try {
            tempFrame.contents = character;
            outlineGroup = tempFrame.createOutline(); /* 複製はここで消費される / The duplicate is consumed here */
            return getCenterY(outlineGroup);
        } finally {
            /* 途中で失敗しても一時オブジェクトを残さない / Never leave the temporary objects behind, even on failure */
            if (outlineGroup) outlineGroup.remove();
            else tempFrame.remove();
        }
    }

    // =========================================
    // 対象の収集 / Collecting targets
    // =========================================

    /**
     * 選択からテキストフレームを取り出す（文字の編集中は対象なし）
     * @param {Array<PageItem>|TextRange} selection - ドキュメントの選択
     * @returns {TextFrame[]} 選択中のテキストフレーム
     */
    function collectTextFrames(selection) {
        var textFrames = [];
        /* 文字の編集中は選択が TextRange になる / While editing text, the selection is a TextRange */
        if (!selection || selection.typename === "TextRange") return textFrames;
        for (var i = 0; i < selection.length; i++) {
            if (selection[i].typename === "TextFrame") textFrames.push(selection[i]);
        }
        return textFrames;
    }

    /**
     * 英数字と空白を除いた文字のうち、いちばん多く出てくる文字を求める（対象文字の初期値）
     * @param {TextFrame[]} textFrames - 対象のテキストフレーム
     * @returns {string} 最も多い文字（見つからなければ空文字）
     */
    function findDefaultTargetChar(textFrames) {
        var charCounts = {};
        for (var i = 0; i < textFrames.length; i++) {
            var frameText = textFrames[i].contents;
            for (var j = 0; j < frameText.length; j++) {
                var character = frameText.charAt(j);
                if (!/^[A-Za-z0-9\s]$/.test(character)) charCounts[character] = (charCounts[character] || 0) + 1;
            }
        }

        var mostFrequentChar = "";
        var maxCount = 0;
        for (var countedChar in charCounts) {
            if (charCounts[countedChar] > maxCount) {
                maxCount = charCounts[countedChar];
                mostFrequentChar = countedChar;
            }
        }
        return mostFrequentChar;
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 項目名＋文字入力欄の1行を追加する
     * @param {Group} parentGroup - 追加先
     * @param {object} labelSet - 項目名のラベル
     * @param {object} tooltipLabelSet - 入力欄の tooltip のラベル
     * @param {string} initialText - 入力欄の初期値
     * @returns {EditText} 追加した入力欄
     */
    function addCharField(parentGroup, labelSet, tooltipLabelSet, initialText) {
        var fieldRow = parentGroup.add("group");
        fieldRow.add("statictext", undefined, labelText(labelSet));
        var charInput = fieldRow.add("edittext", undefined, initialText);
        charInput.characters = CHAR_INPUT_CHARACTERS;
        charInput.helpTip = getLabel(tooltipLabelSet);
        return charInput;
    }

    /**
     * 対象文字と基準文字を入力するダイアログを表示する
     * @param {string} defaultTargetChar - 対象文字の初期値
     * @returns {{targetChars: string, referenceChar: string}|null} 入力された文字（キャンセル時は null）
     */
    function showCharDialog(defaultTargetChar) {
        var baselineDialog = new Window("dialog", getLabel(LABELS.dialog.title));
        baselineDialog.orientation = "column";
        baselineDialog.alignChildren = "left";

        baselineDialog.add("statictext", undefined, getLabel(LABELS.dialog.description));

        var inputGroup = baselineDialog.add("group");
        inputGroup.orientation = "column";
        inputGroup.alignChildren = "left";
        inputGroup.margins = INPUT_GROUP_MARGINS;

        var targetCharInput = addCharField(inputGroup, LABELS.fieldLabel.targetChar, LABELS.tooltip.targetChar, defaultTargetChar);
        targetCharInput.active = true;
        var referenceCharInput = addCharField(inputGroup, LABELS.fieldLabel.referenceChar, LABELS.tooltip.referenceChar, DEFAULT_REFERENCE_CHAR);

        /* ボタンエリア（左右中央）/ Button area (centered) */
        var btnRowGroup = baselineDialog.add("group");
        btnRowGroup.alignment = "center";
        var btnCancel = btnRowGroup.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });
        var btnOK = btnRowGroup.add("button", undefined, getLabel(LABELS.button.adjust), { name: "ok" });

        btnOK.onClick = function () {
            if (targetCharInput.text.length === 0 || referenceCharInput.text.length !== 1) {
                alert(getLabel(LABELS.alert.invalidChars));
                return;
            }
            baselineDialog.close(1);
        };
        btnCancel.onClick = function () { baselineDialog.close(0); };

        prepareDialogWindow(baselineDialog, SCRIPT_NAME);
        if (baselineDialog.show() !== 1) return null;
        return { targetChars: targetCharInput.text, referenceChar: referenceCharInput.text };
    }

    // =========================================
    // ベースラインの調整 / Baseline adjustment
    // =========================================

    /**
     * フレーム内の指定文字すべてにベースラインシフトを設定する
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @param {string} targetChar - 対象の文字
     * @param {number} shiftAmount - ベースラインシフトの値（pt）
     * @returns {void}
     */
    function applyBaselineShift(textFrame, targetChar, shiftAmount) {
        var frameChars = textFrame.textRange.characters;
        for (var i = 0; i < frameChars.length; i++) {
            if (frameChars[i].contents === targetChar) frameChars[i].characterAttributes.baselineShift = shiftAmount;
        }
    }

    /**
     * 1つのフレームで、対象文字それぞれの天地中央を基準文字の天地中央にそろえる
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @param {string} targetChars - 対象文字（複数可）
     * @param {string} referenceChar - 基準文字（1文字）
     * @returns {void}
     */
    function adjustTextFrame(textFrame, targetChars, referenceChar) {
        var frameText = textFrame.contents;

        /* 先にすべて測ってから適用する。測定用の複製は先頭文字の書式を引き継ぐため、
           先頭文字をずらしたあとに測ると基準文字の測定値と食い違う
           Measure everything first: the measuring duplicate inherits the first character's formatting,
           so measuring after that character has been shifted would disagree with the reference measurement */
        var referenceCenterY = null; /* フレームごとに1回だけ測る / Measured once per frame */
        var charShifts = [];
        for (var i = 0; i < targetChars.length; i++) {
            var targetChar = targetChars.charAt(i);
            if (frameText.indexOf(targetChar) === -1) continue;

            if (referenceCenterY === null) referenceCenterY = measureCharCenterY(textFrame, referenceChar);
            charShifts.push({ targetChar: targetChar, shiftAmount: referenceCenterY - measureCharCenterY(textFrame, targetChar) });
        }

        for (var j = 0; j < charShifts.length; j++) {
            applyBaselineShift(textFrame, charShifts[j].targetChar, charShifts[j].shiftAmount);
        }
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 前提を確かめてダイアログを表示し、選択中のテキストフレームを調整する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel(LABELS.alert.noDocument));
            return;
        }

        var textFrames = collectTextFrames(app.activeDocument.selection);
        if (textFrames.length === 0) {
            alert(getLabel(LABELS.alert.selectTextFrame));
            return;
        }

        var charSettings = showCharDialog(findDefaultTargetChar(textFrames));
        if (!charSettings) return;

        /* 字形を持たない文字（スペースなど）はアウトライン化で例外になるため、ここで知らせる
           Characters with no glyph (spaces etc.) throw during outlining, so report it here */
        try {
            for (var i = 0; i < textFrames.length; i++) {
                adjustTextFrame(textFrames[i], charSettings.targetChars, charSettings.referenceChar);
            }
        } catch (e) {
            alert(getLabel(LABELS.alert.errorPrefix) + e);
        }
    }

    main();

})();
