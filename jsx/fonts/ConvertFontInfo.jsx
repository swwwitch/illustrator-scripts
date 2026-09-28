#target illustrator
#targetengine "ConvertFontInfoEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したテキストを、そのテキストで使われているフォント情報の文字列に変換します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ConvertFontInfo.md

### Overview

Replaces the selected text with a string describing the font information it uses.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ConvertFontInfo.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ConvertFontInfo";              /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.4";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-05-09";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-28";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ConvertFontInfo.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ConvertFontInfo.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ローカライズ / Localization
    // =========================================

    function getCurrentLang() {
        return ($.locale && $.locale.indexOf('ja') === 0) ? 'ja' : 'en';
    }

    var LABELS = {
        dialog: {
            title: { ja: "フォント情報に変換", en: "Convert Font Info" }
        },
        panel: {
            title: { ja: "変換形式", en: "Conversion Format" }
        },
        radio: {
            family: { ja: "フォント名", en: "Font Family" },
            style: { ja: "スタイル", en: "Style" },
            familyStyle: { ja: "フォント名＋スタイル", en: "Font Family + Style" },
            postscript: { ja: "PostScript 名", en: "PostScript Name" },
            fullNameSize: { ja: "フルネーム＋サイズ", en: "Full Name + Size" },
            detail: { ja: "詳細", en: "Labeled Detail" }
        },
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" }
        },
        help: {
            fontFamily: { ja: "ショートカット: F", en: "Shortcut: F" },
            style: { ja: "ショートカット: S", en: "Shortcut: S" },
            fontFamilyStyle: { ja: "ショートカット: B", en: "Shortcut: B" },
            postScriptName: { ja: "ショートカット: P", en: "Shortcut: P" },
            fullNameSize: { ja: "ショートカット: M", en: "Shortcut: M" },
            detailLines: { ja: "ショートカット: D", en: "Shortcut: D" }
        },
        detail: {
            familyLabel: { ja: "フォント名", en: "Font Family" },
            styleLabel: { ja: "スタイル", en: "Style" },
            nameLabel: { ja: "PostScript 名", en: "PostScript Name" }
        }
    };

    var currentLanguage = getCurrentLang();

    function getLabel(key) {
        var labelEntry = LABELS;
        var keyParts = key.split('.');
        for (var keyPartIndex = 0; keyPartIndex < keyParts.length; keyPartIndex++) {
            labelEntry = labelEntry && labelEntry[keyParts[keyPartIndex]];
        }
        if (!labelEntry) return key;
        return labelEntry[currentLanguage] || labelEntry.ja || labelEntry.en || key;
    }

    // =========================================
    // 単位 / Units
    // =========================================
    // 単位変換処理（pt → 指定単位、小数第2位）
    function getFontSizeWithUnit(sizePt) {
        var unitPref = app.preferences.getIntegerPreference('text/units');
        var unitLabel = "pt";
        var convertedSize = sizePt;

        switch (unitPref) {
            case 0: unitLabel = "inch"; convertedSize = sizePt / 72; break;
            case 1: unitLabel = "mm"; convertedSize = sizePt * 25.4 / 72; break;
            case 2: unitLabel = "pt"; convertedSize = sizePt; break;
            case 3: unitLabel = "pica"; convertedSize = sizePt / 12; break;
            case 4: unitLabel = "cm"; convertedSize = sizePt * 2.54 / 72; break;
            case 5: unitLabel = "Q"; convertedSize = sizePt * (25.4 / 72) * 4; break;
            case 6: unitLabel = "px"; convertedSize = sizePt * (96 / 72); break;
        }

        return (Math.round(convertedSize * 100) / 100) + " " + unitLabel;
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
    // UI設定 / UI Settings
    // =========================================
    var FORMAT_ITEMS = [
        { labelKey: 'radio.family', format: 'name', helpKey: 'help.fontFamily', shortcut: 'F', showPreview: true },
        { labelKey: 'radio.style', format: 'style', helpKey: 'help.style', shortcut: 'S', showPreview: true },
        { labelKey: 'radio.familyStyle', format: 'family+style', helpKey: 'help.fontFamilyStyle', shortcut: 'B', showPreview: true },
        { labelKey: 'radio.postscript', format: 'postscript', helpKey: 'help.postScriptName', shortcut: 'P', showPreview: true },
        { labelKey: 'radio.fullNameSize', format: 'fullName+size', helpKey: 'help.fullNameSize', shortcut: 'M', showPreview: true },
        { labelKey: 'radio.detail', format: 'detailLines', helpKey: 'help.detailLines', shortcut: 'D', showPreview: false }
    ];

    // 文字ツールで文字を選択すると、app.selection は配列ではなく TextRange を返す
    // （配列でなく .story を持つオブジェクトを TextRange とみなす）
    function isTextRangeSelection(selection) {
        return selection && !(selection instanceof Array) && selection.story;
    }

    // TextRange（文字選択）から、その文字が属する親テキストフレームを取得
    // contents はストーリー単位なので、連結テキストでも先頭フレーム1つで足りる
    function textFramesFromTextRange(textRange) {
        var storyTextFrames = textRange.story && textRange.story.textFrames;
        return (storyTextFrames && storyTextFrames.length >= 1) ? [storyTextFrames[0]] : [];
    }

    // 選択内容から対象テキストフレームの配列を解決する（doc.selection は書き換えない）
    // ・オブジェクト選択：選択内のテキストフレーム
    // ・文字ツールでの部分選択（TextRange）：その文字の親テキストフレーム
    function resolveTargetTextFrames(doc) {
        var selection = doc.selection;
        if (!selection) return [];

        if (isTextRangeSelection(selection)) {
            return textFramesFromTextRange(selection);
        }
        if (!(selection instanceof Array)) return [];

        var textFrames = [];
        for (var selectedIndex = 0; selectedIndex < selection.length; selectedIndex++) {
            var selectedItem = selection[selectedIndex];
            if (selectedItem.typename === "TextFrame") {
                textFrames.push(selectedItem);
            } else if (isTextRangeSelection(selectedItem)) {
                // 配列内に TextRange が入るケースにも対応
                var parentFrames = textFramesFromTextRange(selectedItem);
                for (var parentIndex = 0; parentIndex < parentFrames.length; parentIndex++) {
                    textFrames.push(parentFrames[parentIndex]);
                }
            }
        }
        return textFrames;
    }

    function collectTextFrameInfo(textFrame) {
        try {
            var textRange = textFrame.textRange;
            return {
                textFrame: textFrame,
                font: textRange.characterAttributes.textFont,
                originalSize: textRange.characterAttributes.size,
                originalJustification: textRange.paragraphAttributes.justification,
                originalText: textFrame.contents
            };
        } catch (e) {
            return null;
        }
    }

    function collectSelectedTextFrameInfos(selectedItems) {
        var textFrameInfos = [];

        // 配列風オブジェクト以外（TextRange など）は対象外
        if (!selectedItems || typeof selectedItems.length !== 'number') {
            return textFrameInfos;
        }

        for (var selectedIndex = 0; selectedIndex < selectedItems.length; selectedIndex++) {
            var selectedItem = selectedItems[selectedIndex];
            if (selectedItem.typename !== "TextFrame") continue;

            var textFrameInfo = collectTextFrameInfo(selectedItem);
            if (textFrameInfo) {
                textFrameInfos.push(textFrameInfo);
            }
        }

        return textFrameInfos;
    }

    function restoreTextFrameInfo(originalTextFrameInfo) {
        try {
            var textFrame = originalTextFrameInfo.textFrame;
            textFrame.contents = originalTextFrameInfo.originalText;
            var textRange = textFrame.textRange;
            textRange.characterAttributes.textFont = originalTextFrameInfo.font;
            textRange.characterAttributes.size = originalTextFrameInfo.originalSize;
            textRange.paragraphAttributes.justification = originalTextFrameInfo.originalJustification;
        } catch (e) {
            return false;
        }
        return true;
    }

    function main() {
        if (app.documents.length === 0) return;
        var doc = app.activeDocument;

        var targetTextFrames = resolveTargetTextFrames(doc);
        if (targetTextFrames.length === 0 || targetTextFrames.length >= 1000) return;

        var originalTextFrameInfos = collectSelectedTextFrameInfos(targetTextFrames);
        if (originalTextFrameInfos.length === 0) return;

        showDialog(originalTextFrameInfos);
    }

    function getConvertedFontInfoText(originalTextFrameInfo, format) {
        var sourceFont = originalTextFrameInfo.font;
        var sourceFontSize = originalTextFrameInfo.originalSize;

        if (!sourceFont) return "";

        switch (format) {
            case "style":
                return sourceFont.style;
            case "family+style":
                return sourceFont.family + " " + sourceFont.style;
            case "postscript":
                return sourceFont.name;
            case "fullName+size":
                var fontSizeText = getFontSizeWithUnit(sourceFontSize);
                return (sourceFont.fullName && sourceFont.fullName !== "")
                    ? sourceFont.fullName + "\t" + fontSizeText
                    : sourceFont.family + " " + sourceFont.style + "\t" + fontSizeText;
            default:
                return sourceFont.family;
        }
    }

    function buildDetailLineItems(originalTextFrameInfo) {
        var sourceFont = originalTextFrameInfo.font;

        return [
            { label: getLabel('detail.familyLabel'), value: sourceFont.family },
            { label: getLabel('detail.styleLabel'), value: sourceFont.style },
            { label: getLabel('detail.nameLabel'), value: sourceFont.name }
        ];
    }

    function findTextFontByName(fontName) {
        var textFonts = app.textFonts;
        for (var fontIndex = 0; fontIndex < textFonts.length; fontIndex++) {
            if (textFonts[fontIndex].name === fontName) {
                return textFonts[fontIndex];
            }
        }
        return null;
    }

    // 詳細表示のラベル用フォント（セッション中変わらないので一度だけ走査してキャッシュ）
    var cachedDetailLabelFont; // 未取得は undefined、取得済みで未発見は null
    function getDetailLabelFont() {
        if (cachedDetailLabelFont === undefined) {
            cachedDetailLabelFont = findTextFontByName("HiraginoSans-W3");
        }
        return cachedDetailLabelFont;
    }

    function buildDetailLinesText(originalTextFrameInfo) {
        var lineBreak = String.fromCharCode(13);
        var detailLineItems = buildDetailLineItems(originalTextFrameInfo);
        var detailTextParts = [];

        for (var detailItemIndex = 0; detailItemIndex < detailLineItems.length; detailItemIndex++) {
            detailTextParts.push(detailLineItems[detailItemIndex].label);
            detailTextParts.push(detailLineItems[detailItemIndex].value);
        }

        return detailTextParts.join(lineBreak);
    }

    function applyDetailFontInfoPreview(originalTextFrameInfo, detailLabelFont) {
        try {
            var textFrame = originalTextFrameInfo.textFrame;
            var sourceFont = originalTextFrameInfo.font;
            var sourceFontSize = originalTextFrameInfo.originalSize;

            textFrame.contents = buildDetailLinesText(originalTextFrameInfo);
            textFrame.textRange.paragraphAttributes.justification = Justification.LEFT;

            var previewLines = textFrame.textRange.lines;
            if (previewLines.length > 0 && detailLabelFont) {
                for (var lineIndex = 0; lineIndex < previewLines.length; lineIndex++) {
                    var isLabelLine = (lineIndex % 2 === 0);
                    var lineAttributes = previewLines[lineIndex].characterAttributes;
                    lineAttributes.textFont = isLabelLine ? detailLabelFont : sourceFont;
                    lineAttributes.size = isLabelLine ? 10 : sourceFontSize;
                }
            }
        } catch (e) {
            return false;
        }
        return true;
    }

    function applySimpleFontInfoPreview(originalTextFrameInfo, format) {
        try {
            var textFrame = originalTextFrameInfo.textFrame;
            var sourceFont = originalTextFrameInfo.font;
            var sourceFontSize = originalTextFrameInfo.originalSize;
            var textRange = textFrame.textRange;

            textRange.characterAttributes.textFont = sourceFont;
            textRange.characterAttributes.size = sourceFontSize;
            textFrame.contents = getConvertedFontInfoText(originalTextFrameInfo, format);
        } catch (e) {
            return false;
        }
        return true;
    }

    function updateFontInfoPreview(originalTextFrameInfos, format) {
        var isDetailLinesFormat = (format === "detailLines");
        var detailLabelFont = isDetailLinesFormat ? getDetailLabelFont() : null;

        for (var textFrameIndex = 0; textFrameIndex < originalTextFrameInfos.length; textFrameIndex++) {
            if (isDetailLinesFormat) {
                applyDetailFontInfoPreview(originalTextFrameInfos[textFrameIndex], detailLabelFont);
            } else {
                applySimpleFontInfoPreview(originalTextFrameInfos[textFrameIndex], format);
            }
        }

        app.redraw();
    }

    function restoreOriginalText(originalTextFrameInfos) {
        for (var textFrameIndex = 0; textFrameIndex < originalTextFrameInfos.length; textFrameIndex++) {
            restoreTextFrameInfo(originalTextFrameInfos[textFrameIndex]);
        }

        app.redraw();
    }

    function showDialog(originalTextFrameInfos) {
        var dialog = new Window('dialog', getLabel('dialog.title') + ' ' + SCRIPT_VERSION);

        var dialogContainer = dialog.add("group");
        dialogContainer.orientation = "column";
        dialogContainer.alignChildren = "left";
        dialogContainer.alignment = "fill";

        buildFormatPanel(dialog, dialogContainer, originalTextFrameInfos);
        buildButtonRow(dialog, dialogContainer);

        // 選択中の形式はライブプレビューで既にドキュメントへ適用済み。
        // OK時は再適用せずそのまま確定し（二重適用によるクラッシュを回避）、
        // キャンセル時のみ開始時の内容・フォント・サイズ・行揃えに戻す。
        prepareDialogWindow(dialog, SCRIPT_NAME);
        var dialogResult = dialog.show();
        if (dialogResult !== 1) {
            restoreOriginalText(originalTextFrameInfos);
        }
    }

    // 変換形式パネル（ラジオ生成・選択・キー操作）を組み立て、選択中フォーマットを返す関数を返す
    function buildFormatPanel(dialog, dialogContainer, originalTextFrameInfos) {
        var DEFAULT_FORMAT = 'family+style';
        var previewSourceInfo = originalTextFrameInfos[0];
        var radioColumnWidth = 180;
        var formatRadioButtons = [];

        var formatPanel = dialogContainer.add("panel", undefined, getLabel('panel.title'));
        formatPanel.orientation = "column";
        formatPanel.alignChildren = "left";
        formatPanel.alignment = "fill";
        formatPanel.margins = [15, 20, 15, 10];

        var formatRowsGroup = formatPanel.add("group");
        formatRowsGroup.orientation = "column";
        formatRowsGroup.alignChildren = ["fill", "top"];
        formatRowsGroup.spacing = 6;

        function addFormatRow(labelKey, previewText, helpTipKey) {
            var formatRow = formatRowsGroup.add("group");
            formatRow.orientation = "row";
            formatRow.alignChildren = ["left", "center"];

            var radioColumn = formatRow.add("group");
            radioColumn.preferredSize.width = radioColumnWidth;
            var radioButton = radioColumn.add("radiobutton", undefined, getLabel(labelKey));

            if (helpTipKey) {
                radioButton.helpTip = getLabel(helpTipKey);
            }

            if (previewText) {
                var previewStaticText = formatRow.add("statictext", undefined, previewText);
                previewStaticText.helpTip = previewText;
            }

            return radioButton;
        }

        function selectFormat(format) {
            for (var itemIndex = 0; itemIndex < FORMAT_ITEMS.length; itemIndex++) {
                formatRadioButtons[itemIndex].value = (FORMAT_ITEMS[itemIndex].format === format);
            }
            updateFontInfoPreview(originalTextFrameInfos, format);
        }

        for (var formatIndex = 0; formatIndex < FORMAT_ITEMS.length; formatIndex++) {
            (function (formatItem, itemIndex) {
                var previewText = formatItem.showPreview ? getConvertedFontInfoText(previewSourceInfo, formatItem.format) : "";
                var radioButton = addFormatRow(formatItem.labelKey, previewText, formatItem.helpKey);
                formatRadioButtons[itemIndex] = radioButton;
                radioButton.onClick = function () {
                    selectFormat(formatItem.format);
                };
            })(FORMAT_ITEMS[formatIndex], formatIndex);
        }

        dialog.addEventListener("keydown", function (event) {
            var keyName = String(event.keyName).toUpperCase();
            for (var itemIndex = 0; itemIndex < FORMAT_ITEMS.length; itemIndex++) {
                if (FORMAT_ITEMS[itemIndex].shortcut === keyName) {
                    selectFormat(FORMAT_ITEMS[itemIndex].format);
                    event.preventDefault();
                    return;
                }
            }
        });

        selectFormat(DEFAULT_FORMAT);
    }

    // キャンセル／OK ボタン行を組み立て、OK をデフォルトボタンに設定
    function buildButtonRow(dialog, dialogContainer) {
        var buttonGroup = dialogContainer.add("group");
        buttonGroup.orientation = "row";
        buttonGroup.alignChildren = ["fill", "center"];
        buttonGroup.alignment = "fill";
        buttonGroup.margins = [0, 10, 0, 0];

        var cancelButtonGroup = buttonGroup.add("group");
        cancelButtonGroup.orientation = "row";
        cancelButtonGroup.alignChildren = ["left", "center"];
        cancelButtonGroup.alignment = ["left", "center"];
        cancelButtonGroup.add("button", undefined, getLabel('button.cancel'), { name: "cancel" });

        var centerSpacerGroup = buttonGroup.add("group");
        centerSpacerGroup.alignment = ["fill", "fill"];
        centerSpacerGroup.minimumSize.width = 100;

        var okButtonGroup = buttonGroup.add("group");
        okButtonGroup.orientation = "row";
        okButtonGroup.alignChildren = ["right", "center"];
        okButtonGroup.alignment = ["right", "center"];
        var okButton = okButtonGroup.add("button", undefined, getLabel('button.ok'), { name: "ok" });
        dialog.defaultElement = okButton;
    }

    main();

})();
