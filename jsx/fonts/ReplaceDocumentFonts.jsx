#target illustrator
#targetengine "ReplaceDocumentFontsEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

ドキュメントで使用中のフォントをファミリー／スタイル単位で一覧し、
選んだフォントを別のフォントへまとめて置き換えます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ReplaceDocumentFonts.md

note記事も参照してください。
https://note.com/dtp_tranist/n/ncc9330ba1f7d

### Overview

Lists the fonts used in the document by family and style, and replaces
the selected ones with another font in a single pass.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ReplaceDocumentFonts.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ReplaceDocumentFonts";         /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v2.0.3";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-03-29";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-29";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ReplaceDocumentFonts.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ReplaceDocumentFonts.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/ncc9330ba1f7d"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function() {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* スタイル行の字下げ / Indent used for style rows */
    var STYLE_ROW_INDENT = "　　";

    /* PostScript名表示の初期状態 / Initial state of the PostScript-name display */
    var SHOW_POSTSCRIPT_NAME_DEFAULT = false;

    // =========================================
    // レイアウト / Layout
    // =========================================

    var WINDOW_MARGINS        = 15;   /* ウィンドウ外周の余白 / window margin */
    var WINDOW_SPACING        = 10;   /* ウィンドウ内の要素間隔 / window spacing */
    var COLUMN_SPACING        = 10;   /* 2カラムの間隔 / spacing between the two lists */
    var LIST_LABEL_SPACING    = 6;    /* 見出しとリストの間隔 / spacing between a label and its list */
    var LISTBOX_HEIGHT        = 300;  /* リストの高さ / list height */
    var LISTBOX_WIDTH_MIN     = 200;  /* リスト幅の下限 / minimum list width */
    var LISTBOX_WIDTH_MAX     = 600;  /* リスト幅の上限 / maximum list width */
    var LISTBOX_CHAR_WIDTH    = 9;    /* 1文字あたりの概算幅 / approximate width per character */
    var LISTBOX_WIDTH_PADDING = 60;   /* リスト幅の余裕 / extra width added to the list */

    // ボタン行（再利用パーツ） / Button row (reusable)

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

    // ボタン行（再利用パーツ）ここまで / End of the reusable button row

    // ダイアログの位置と不透明度（再利用パーツ） / Dialog position and opacity (reusable)

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

    // ダイアログの位置と不透明度（再利用パーツ）ここまで / End of the reusable dialog position and opacity

    // =========================================
    // ローカライズ / Localization
    // =========================================

    // ローカライズ（再利用パーツ） / Localization (reusable)

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

    // ローカライズ（再利用パーツ）ここまで / End of the reusable localization

    var LABELS = {
        dialog: {
            title: { ja: "フォント置換", en: "Replace Fonts" }
        },
        fieldLabel: {
            sourceFonts: { ja: "置換元フォント（複数選択可）", en: "Source Fonts (Multiple Selection)" },
            targetFont: { ja: "置換先フォント", en: "Target Font" }
        },
        checkbox: {
            postScriptName: { ja: "PostScript名で表示", en: "Show PostScript names" }
        },
        button: {
            close: { ja: "閉じる", en: "Close" },
            replaceAll: { ja: "全置換", en: "Replace All" },
            replace: { ja: "フォントを置換", en: "Replace Fonts" }
        },
        tooltip: {
            sourceFonts: {
                ja: "置換元のフォントを選びます。ファミリー名の行を選ぶと、そのファミリーのスタイルがすべて選ばれます。",
                en: "Pick the fonts to replace. Selecting a family row selects every style in that family."
            },
            targetFont: {
                ja: "置換先のフォントを選びます。ファミリー名の行は選べません。",
                en: "Pick the font to replace them with. Family rows cannot be selected."
            },
            postScriptName: {
                ja: "ファミリー名とスタイル名の代わりに、PostScript名で一覧します。",
                en: "List the fonts by PostScript name instead of family and style."
            },
            replaceAll: {
                ja: "使用中のすべてのフォントを置換先フォントに置き換えます。置換先を選んでいないときは、置換元の1つ目のフォントに揃えます。",
                en: "Replace every font in use with the target font. With no target selected, the first source font is used instead."
            },
            replace: {
                ja: "選んだ置換元フォントを、置換先フォントに置き換えます。",
                en: "Replace the selected source fonts with the target font."
            }
        },
        alert: {
            noDocument: {
                ja: "ドキュメントが開かれていません。",
                en: "No document is open."
            },
            noFontsFound: {
                ja: "ドキュメント内に使用中のフォントが見つかりません。",
                en: "No fonts in use were found in the document."
            },
            noSourceFont: {
                ja: "置換元フォントを1つ以上選択してください。",
                en: "Please select at least one source font."
            },
            noTargetFont: {
                ja: "置換先フォントを選択してください。",
                en: "Please select a target font."
            },
            selectFonts: {
                ja: "置換先または置換元フォントを選択してください。\n（または、置換元だけ選んで全置換することも可能です）",
                en: "Please select either a target font or source fonts.\n(Alternatively, select only the source fonts to replace all.)"
            },
            targetNotFound: {
                ja: "置換先フォントが見つかりません（%1）。",
                en: "Target font not found (%1)."
            }
        }
    };

    /**
     * 件数を括弧で囲む（日本語は全角括弧、英語は半角括弧）
     * @param {number} count - 件数
     * @returns {string} 括弧付きの件数
     */
    function countSuffix(count) {
        return (uiLang === "ja") ? "（" + count + "）" : " (" + count + ")";
    }

    // =========================================
    // 状態 / State
    // =========================================

    var doc = null;
    var mainDialog = null;
    var sourceFontListBox = null;
    var targetFontListBox = null;
    var postScriptNameCheckbox = null;

    /* リストに並べるフォント（ファミリー見出しを含む）/ Fonts listed in the boxes, family headers included */
    var flatFontList = [];

    /* フォント名ごとのTextRange / Text ranges tagged with their font name */
    var textRangeList = [];

    /* 選択の再帰更新を防ぐフラグ / Guard against recursive selection updates */
    var isUpdatingSelection = false;

    // =========================================
    // フォントの収集 / Collecting fonts
    // =========================================

    /**
     * ドキュメント内のTextRangeを走査し、使用中フォントをファミリー別に集める
     * @returns {object} ファミリー名 → フォント名 → {name, style, family, frameCount} のマップ
     */
    function collectUsedFonts() {
        var usedFontMap = {};
        textRangeList = [];

        var textFrames = doc.textFrames;
        for (var i = 0; i < textFrames.length; i++) {
            var textFrame = textFrames[i];
            var isLocked = textFrame.locked;
            var isHidden = textFrame.hidden;
            var ranges = textFrame.textRanges;
            var fontsInFrame = {};

            for (var j = 0; j < ranges.length; j++) {
                var range = ranges[j];
                if (range.length === 0) continue;

                /* 無効な範囲はフォントを取得できないのでスキップ / Skip ranges whose font cannot be read */
                var font = null;
                try {
                    font = range.characterAttributes.textFont;
                } catch (e) {
                    continue;
                }
                if (!font) continue;

                textRangeList.push({
                    range: range,
                    fontName: font.name,
                    isLocked: isLocked,
                    isHidden: isHidden
                });

                /* 同じTextFrame内では1フォントにつき1回だけ数える / Count a font once per text frame */
                if (fontsInFrame[font.name]) continue;
                fontsInFrame[font.name] = true;

                if (!usedFontMap[font.family]) usedFontMap[font.family] = {};
                if (!usedFontMap[font.family][font.name]) {
                    usedFontMap[font.family][font.name] = {
                        name: font.name,
                        style: font.style,
                        family: font.family,
                        frameCount: 1
                    };
                } else {
                    usedFontMap[font.family][font.name].frameCount++;
                }
            }
        }
        return usedFontMap;
    }

    /**
     * 収集したフォントをリスト表示用の1次元配列にする
     * @param {object} usedFontMap - collectUsedFonts() が返したマップ
     * @returns {Array<object>} ファミリー見出しとスタイル行を並べた配列（PostScript名表示中は見出しなしの1フォント1行）
     */
    function buildFlatFontList(usedFontMap) {
        var showsPostScriptName = postScriptNameCheckbox ? postScriptNameCheckbox.value : SHOW_POSTSCRIPT_NAME_DEFAULT;
        var fontList = [];

        for (var family in usedFontMap) {
            var styles = usedFontMap[family];
            var fontNames = [];
            for (var fontName in styles) {
                fontNames.push(fontName);
            }

            /* PostScript名表示のときは見出しを立てず、1フォント1行で並べる / In PostScript-name mode, list one row per font with no headers */
            if (showsPostScriptName) {
                for (var p = 0; p < fontNames.length; p++) {
                    var psFont = styles[fontNames[p]];
                    fontList.push({
                        label: psFont.name + countSuffix(psFont.frameCount),
                        name: psFont.name,
                        family: psFont.family,
                        style: psFont.style,
                        isHeader: false
                    });
                }
                continue;
            }

            /* スタイルが1つだけのファミリーは見出しを立てず1行で見せる / Show single-style families on one row */
            if (fontNames.length === 1) {
                var onlyFont = styles[fontNames[0]];
                fontList.push({
                    label: onlyFont.family + " " + onlyFont.style + countSuffix(onlyFont.frameCount),
                    name: onlyFont.name,
                    family: onlyFont.family,
                    style: onlyFont.style,
                    isHeader: false
                });
                continue;
            }

            fontList.push({ label: family, family: family, isHeader: true });
            for (var i = 0; i < fontNames.length; i++) {
                var font = styles[fontNames[i]];
                fontList.push({
                    label: STYLE_ROW_INDENT + font.style + countSuffix(font.frameCount),
                    name: font.name,
                    family: font.family,
                    style: font.style,
                    isHeader: false
                });
            }
        }
        return fontList;
    }

    /**
     * フォント一覧を集め直してリストを作り直す（選択はフォント名で引き継ぐ）
     * @returns {void}
     */
    function refreshFontList() {
        var previousSourceFontNames = getSelectedSourceFontNames();
        var previousTargetFontName = getSelectedTargetFontName();

        flatFontList = buildFlatFontList(collectUsedFonts());
        populateFontListBoxes();
        restoreSelection(previousSourceFontNames, previousTargetFontName);
    }

    // =========================================
    // 置換処理 / Replacing fonts
    // =========================================

    /**
     * フォント名から TextFont を取得する
     * @param {string} fontName - フォント名（PostScript名）
     * @returns {TextFont|null} 見つからなければ null
     */
    function findFontByName(fontName) {
        try {
            return app.textFonts.getByName(fontName);
        } catch (e) {
            return null;
        }
    }

    /**
     * 置換元として選択されているフォント名を取り出す（見出し行は除く）
     * @returns {Array<string>} フォント名の配列
     */
    function getSelectedSourceFontNames() {
        var fontNames = [];
        if (!sourceFontListBox || !sourceFontListBox.selection) return fontNames;

        for (var i = 0; i < sourceFontListBox.selection.length; i++) {
            var listEntry = flatFontList[sourceFontListBox.selection[i].index];
            if (!listEntry.isHeader) fontNames.push(listEntry.name);
        }
        return fontNames;
    }

    /**
     * 置換先リストでフォント（見出し行以外）が選ばれているか調べる
     * @returns {boolean} 選ばれていれば true
     */
    function hasTargetFontSelection() {
        if (!targetFontListBox || !targetFontListBox.selection) return false;
        return !flatFontList[targetFontListBox.selection.index].isHeader;
    }

    /**
     * 置換先として選択されているフォント名を取り出す（見出し行は除く）
     * @returns {string} 未選択・見出し行のときは空文字列
     */
    function getSelectedTargetFontName() {
        if (!hasTargetFontSelection()) return "";
        return flatFontList[targetFontListBox.selection.index].name;
    }

    /**
     * 置換先として選択されているフォントを取り出す
     * @returns {TextFont|null} 未選択・見出し行・未インストールの場合は null
     */
    function getSelectedTargetFont() {
        var fontName = getSelectedTargetFontName();
        if (fontName === "") return null;

        var targetFont = findFontByName(fontName);
        if (!targetFont) alert(getLabel(LABELS.alert.targetNotFound, [fontName]));
        return targetFont;
    }

    /**
     * 使用中フォントの名前をすべて集める
     * @param {string} [excludedFontName] - 除外するフォント名
     * @returns {Array<string>} フォント名の配列
     */
    function collectAllFontNames(excludedFontName) {
        var fontNames = [];
        for (var i = 0; i < flatFontList.length; i++) {
            if (flatFontList[i].isHeader) continue;
            if (flatFontList[i].name === excludedFontName) continue;
            fontNames.push(flatFontList[i].name);
        }
        return fontNames;
    }

    /**
     * 指定したフォントを置換先フォントに置き換える（ロック・非表示は対象外）
     * @param {Array<string>} sourceFontNames - 置換元のフォント名
     * @param {TextFont} targetFont - 置換先フォント
     * @returns {void}
     */
    function replaceFonts(sourceFontNames, targetFont) {
        if (!targetFont || sourceFontNames.length === 0) return;

        for (var i = 0; i < textRangeList.length; i++) {
            var entry = textRangeList[i];
            if (entry.isLocked || entry.isHidden) continue;

            for (var j = 0; j < sourceFontNames.length; j++) {
                if (entry.fontName === sourceFontNames[j]) {
                    entry.range.characterAttributes.textFont = targetFont;
                    break;
                }
            }
        }
        refreshFontList();
        app.redraw();
    }

    // =========================================
    // イベントハンドラー / Event handlers
    // =========================================

    /**
     * 置換元リストの選択を整え、該当テキストをハイライトする
     * @returns {void}
     */
    function handleSourceFontSelection() {
        if (isUpdatingSelection || !sourceFontListBox.selection) return;

        /* 見出し行はファミリー内の全スタイルに展開する / Expand a family header to all its styles */
        var expandedSelection = [];
        for (var i = 0; i < sourceFontListBox.selection.length; i++) {
            var selectedIndex = sourceFontListBox.selection[i].index;
            var listEntry = flatFontList[selectedIndex];

            if (!listEntry.isHeader) {
                expandedSelection.push(sourceFontListBox.items[selectedIndex]);
                continue;
            }
            for (var j = 0; j < flatFontList.length; j++) {
                if (!flatFontList[j].isHeader && flatFontList[j].family === listEntry.family) {
                    expandedSelection.push(sourceFontListBox.items[j]);
                }
            }
        }

        isUpdatingSelection = true;
        sourceFontListBox.selection = expandedSelection;
        isUpdatingSelection = false;

        selectTextOfSelectedFonts(expandedSelection);
    }

    /**
     * 選択中のフォントを使っているテキストをドキュメント上で選択する
     * @param {Array<ListItem>} selectedItems - 置換元リストで選択中の項目
     * @returns {void}
     */
    function selectTextOfSelectedFonts(selectedItems) {
        doc.selection = null;

        for (var i = 0; i < selectedItems.length; i++) {
            var fontName = flatFontList[selectedItems[i].index].name;
            for (var j = 0; j < textRangeList.length; j++) {
                if (textRangeList[j].fontName === fontName) {
                    textRangeList[j].range.selected = true;
                }
            }
        }
    }

    /**
     * 置換先リストで見出し行が選ばれたら選択を解除する
     * @returns {void}
     */
    function handleTargetFontSelection() {
        if (isUpdatingSelection || !targetFontListBox.selection) return;
        if (flatFontList[targetFontListBox.selection.index].isHeader) {
            targetFontListBox.selection = null;
        }
    }

    /**
     * ［PostScript名で表示］：表示形式を切り替えてリストを作り直す
     * @returns {void}
     */
    function handleDisplayModeChange() {
        refreshFontList();
    }

    /**
     * ［フォントを置換］：選んだ置換元フォントを置換先フォントに置き換える
     * @returns {void}
     */
    function handleReplaceClick() {
        var sourceFontNames = getSelectedSourceFontNames();
        if (sourceFontNames.length === 0) {
            alert(getLabel(LABELS.alert.noSourceFont));
            return;
        }
        if (!hasTargetFontSelection()) {
            alert(getLabel(LABELS.alert.noTargetFont));
            return;
        }
        replaceFonts(sourceFontNames, getSelectedTargetFont());
    }

    /**
     * ［全置換］：使用中のすべてのフォントを1つのフォントにそろえる
     * @returns {void}
     */
    function handleReplaceAllClick() {
        /* 置換先が選ばれていれば、それにすべてをそろえる / Unify on the target font when one is selected */
        var targetFont = getSelectedTargetFont();
        if (targetFont) {
            replaceFonts(collectAllFontNames(), targetFont);
            return;
        }

        /* 置換先がなければ、置換元の1つ目にそろえる / Otherwise unify on the first source font */
        var sourceFontNames = getSelectedSourceFontNames();
        if (sourceFontNames.length === 0) {
            alert(getLabel(LABELS.alert.selectFonts));
            return;
        }

        var fallbackFont = findFontByName(sourceFontNames[0]);
        if (!fallbackFont) {
            alert(getLabel(LABELS.alert.targetNotFound, [sourceFontNames[0]]));
            return;
        }
        replaceFonts(collectAllFontNames(sourceFontNames[0]), fallbackFont);
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 見出し付きのフォントリストを1カラム分追加する
     * @param {Group} parent - 追加先のグループ
     * @param {object} labelSet - 見出しのラベル（ja/en）
     * @param {object} tooltipSet - ツールチップのラベル（ja/en）
     * @param {boolean} allowsMultiple - 複数選択を許可するか
     * @returns {ListBox} 追加したリストボックス
     */
    function addFontListColumn(parent, labelSet, tooltipSet, allowsMultiple) {
        var columnGroup = parent.add("group");
        columnGroup.orientation = "column";
        columnGroup.alignChildren = ["fill", "top"];
        columnGroup.spacing = LIST_LABEL_SPACING;
        columnGroup.add("statictext", undefined, labelText(labelSet));

        var fontListBox = columnGroup.add("listbox", undefined, [], { multiselect: allowsMultiple });
        fontListBox.preferredSize.height = LISTBOX_HEIGHT;
        fontListBox.tabEnabled = true;
        fontListBox.helpTip = getLabel(tooltipSet);
        return fontListBox;
    }

    /**
     * ダイアログを組み立てる
     * @returns {void}
     */
    function buildDialog() {
        mainDialog = new Window("dialog", getLabel(LABELS.dialog.title) + " " + SCRIPT_VERSION);
        mainDialog.orientation = "column";
        mainDialog.alignChildren = ["fill", "top"];
        mainDialog.margins = WINDOW_MARGINS;
        mainDialog.spacing = WINDOW_SPACING;

        var listGroup = mainDialog.add("group");
        listGroup.orientation = "row";
        listGroup.alignChildren = ["fill", "top"];
        listGroup.spacing = COLUMN_SPACING;

        sourceFontListBox = addFontListColumn(listGroup, LABELS.fieldLabel.sourceFonts, LABELS.tooltip.sourceFonts, true);
        targetFontListBox = addFontListColumn(listGroup, LABELS.fieldLabel.targetFont, LABELS.tooltip.targetFont, false);

        /* 表示オプション / Display options */
        var optionGroup = mainDialog.add("group");
        optionGroup.orientation = "row";
        optionGroup.alignChildren = ["left", "center"];
        postScriptNameCheckbox = optionGroup.add("checkbox", undefined, getLabel(LABELS.checkbox.postScriptName));
        postScriptNameCheckbox.value = SHOW_POSTSCRIPT_NAME_DEFAULT;
        postScriptNameCheckbox.helpTip = getLabel(LABELS.tooltip.postScriptName);

        /* ボタンエリア（右側に閉じる→置換系）/ Button row: Close, then replace buttons, on the right */
        var buttonRow = addButtonRow(mainDialog);
        var btnClose = buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.close), { name: "cancel" });
        var btnReplaceAll = buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.replaceAll));
        var btnReplace = buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.replace), { name: "ok" });
        btnReplaceAll.helpTip = getLabel(LABELS.tooltip.replaceAll);
        btnReplace.helpTip = getLabel(LABELS.tooltip.replace);

        sourceFontListBox.onChange = handleSourceFontSelection;
        targetFontListBox.onChange = handleTargetFontSelection;
        postScriptNameCheckbox.onClick = handleDisplayModeChange;
        btnReplaceAll.onClick = handleReplaceAllClick;
        btnReplace.onClick = handleReplaceClick;
    }

    /**
     * 2つのリストにフォント一覧を流し込み、幅をそろえる
     * @returns {void}
     */
    function populateFontListBoxes() {
        var listBoxWidth = calculateListBoxWidth(flatFontList);

        isUpdatingSelection = true;
        sourceFontListBox.removeAll();
        targetFontListBox.removeAll();
        for (var i = 0; i < flatFontList.length; i++) {
            sourceFontListBox.add("item", flatFontList[i].label);
            targetFontListBox.add("item", flatFontList[i].label);
        }
        isUpdatingSelection = false;

        sourceFontListBox.preferredSize.width = listBoxWidth;
        targetFontListBox.preferredSize.width = listBoxWidth;
    }

    /**
     * フォント名を手がかりに、作り直したリストの選択を元に戻す
     * @param {Array<string>} sourceFontNames - 置換元として選択されていたフォント名
     * @param {string} targetFontName - 置換先として選択されていたフォント名
     * @returns {void}
     */
    function restoreSelection(sourceFontNames, targetFontName) {
        var restoredSelection = [];
        var restoredTargetIndex = -1;

        for (var i = 0; i < flatFontList.length; i++) {
            if (flatFontList[i].isHeader) continue;

            for (var j = 0; j < sourceFontNames.length; j++) {
                if (flatFontList[i].name === sourceFontNames[j]) {
                    restoredSelection.push(sourceFontListBox.items[i]);
                    break;
                }
            }
            if (restoredTargetIndex === -1 && flatFontList[i].name === targetFontName) {
                restoredTargetIndex = i;
            }
        }

        isUpdatingSelection = true;
        sourceFontListBox.selection = restoredSelection;
        targetFontListBox.selection = (restoredTargetIndex === -1) ? null : restoredTargetIndex;
        isUpdatingSelection = false;
    }

    /**
     * いちばん長いラベルからリストの幅を見積もる
     * @param {Array<object>} fontList - リストに並べるフォント
     * @returns {number} リストの幅（px）
     */
    function calculateListBoxWidth(fontList) {
        var maxLength = 0;
        for (var i = 0; i < fontList.length; i++) {
            if (fontList[i].isHeader) continue;
            if (fontList[i].label.length > maxLength) maxLength = fontList[i].label.length;
        }
        var estimatedWidth = maxLength * LISTBOX_CHAR_WIDTH + LISTBOX_WIDTH_PADDING;
        return Math.min(LISTBOX_WIDTH_MAX, Math.max(LISTBOX_WIDTH_MIN, estimatedWidth));
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * スクリプトのエントリーポイント
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel(LABELS.alert.noDocument));
            return;
        }
        doc = app.activeDocument;

        flatFontList = buildFlatFontList(collectUsedFonts());
        if (flatFontList.length === 0) {
            alert(getLabel(LABELS.alert.noFontsFound));
            return;
        }

        buildDialog();
        populateFontListBoxes();

        /* 先頭のフォントを選び、該当テキストをハイライトしておく / Preselect the first font and highlight its text */
        sourceFontListBox.selection = 0;
        targetFontListBox.selection = 0;
        handleSourceFontSelection();

        prepareDialogWindow(mainDialog, SCRIPT_NAME);
        mainDialog.show();
    }

    main();

})();
