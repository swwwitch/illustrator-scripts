#target illustrator
#targetengine "TextGridAlignerEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

テキストフレームを行・列単位で整列またはグループ化します。
行方向と列方向のしきい値を独立して調整でき、行・列のアキを均等に配置するオプションもあります。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/TextGridAligner.md

### Overview

Aligns or groups text frames by row and by column.
The row and column thresholds are tuned independently, and an option evens out the gaps between them.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/TextGridAligner.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "TextGridAligner";              /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.6";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-08-02";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-30";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/TextGridAligner.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/TextGridAligner.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* 行・列の判定に使う隙間のしきい値の初期値（pt）/ initial gap threshold used to group items */
    var DEFAULT_GAP_THRESHOLD = 10;

    // =========================================
    // レイアウト / Layout
    // =========================================
    var PANEL_MARGINS = [15, 20, 15, 10];       /* パネル余白 [左,上,右,下] / panel margins */
    var SLIDER_WIDTH = 150;                     /* しきい値スライダーの幅 / threshold slider width */
    var THRESHOLD_LABEL_CHARS = 5;              /* しきい値表示の文字数 / width of the threshold readout */

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
            /* 文字ツールで文字を選択しているときは TextRange が返り、[0] が無い / Selecting characters with the Type tool returns a TextRange, which has no [0] */
            if (!selectedItems || selectedItems.typename === "TextRange" || !selectedItems.length || !selectedItems[0].visibleBounds) return null;
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

    /**
     * 左のグループにボタンが無い（右のボタンだけの）とき、行を左右中央に並べ直す。
     * ボタンをすべて足したあと、show() の前に呼ぶ。centered で作った行や、左にボタンがある行はそのまま
     * @param {{rowGroup: Group, leftGroup: Group|null, rightGroup: Group|null}} buttonRow - addButtonRow() の戻り値
     * @returns {void}
     */
    function centerButtonRowIfRightOnly(buttonRow) {
        if (!buttonRow.leftGroup || buttonRow.leftGroup.children.length > 0) return;
        var btnRowGroup = buttonRow.rowGroup;
        /* 左のグループとスペーサーを外し、右のグループだけを中央に置く / Drop the left group and the spacer so only the right group remains, centered */
        btnRowGroup.remove(buttonRow.leftGroup);
        btnRowGroup.remove(btnRowGroup.children[0]); /* 左のグループを外すと先頭はスペーサー / the spacer is first once the left group is gone */
        btnRowGroup.alignment = ["center", "bottom"];
        btnRowGroup.alignChildren = ["center", "center"];
        buttonRow.leftGroup = null;
    }

    // ボタン行（再利用パーツ）ここまで / End of the reusable button row

    // =========================================
    // 方向の定義 / Direction constants
    // =========================================
    var DIRECTION_HORIZONTAL = "horizontal"; /* 横並び＝同じ行 / a row */
    var DIRECTION_VERTICAL = "vertical";     /* 縦並び＝同じ列 / a column */

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

    /* カテゴリ分けした日英ラベル定義 / Categorized Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "テキスト整列・グループ化", en: "Text Alignment & Grouping" }
        },
        panel: {
            rows:    { ja: "行", en: "Rows" },
            columns: { ja: "列", en: "Columns" }
        },
        checkbox: {
            alignRows:       { ja: "揃え", en: "Align" },
            groupRows:       { ja: "行をグループ化", en: "Group rows" },
            distributeRows:  { ja: "アキを均等に", en: "Distribute evenly" },
            alignColumns:    { ja: "揃え", en: "Align" },
            groupColumns:    { ja: "列をグループ化", en: "Group columns" },
            distributeColumns: { ja: "アキを均等に", en: "Distribute evenly" }
        },
        tooltip: {
            alignRows: {
                ja: "同じ行とみなしたテキストを、天地中央でそろえます。",
                en: "Vertically centers the text objects that were grouped into the same row."
            },
            groupRows: {
                ja: "同じ行とみなしたテキストを1つのグループにまとめます。行と列は同時にグループ化できません。",
                en: "Groups each detected row into one group. Rows and columns cannot be grouped at the same time."
            },
            distributeRows: {
                ja: "作成した行グループどうしの縦のアキを均等にします。",
                en: "Evens out the vertical gaps between the row groups."
            },
            rowThreshold: {
                ja: "左右の隙間がこの値以内なら、同じ行とみなします。スライダーを動かすと結果がすぐ反映されます。",
                en: "Text objects with a horizontal gap up to this value form one row. The canvas updates as you drag."
            },
            alignColumns: {
                ja: "同じ列とみなしたテキストを、左右中央でそろえます。",
                en: "Horizontally centers the text objects that were grouped into the same column."
            },
            groupColumns: {
                ja: "同じ列とみなしたテキストを1つのグループにまとめます。行と列は同時にグループ化できません。",
                en: "Groups each detected column into one group. Rows and columns cannot be grouped at the same time."
            },
            distributeColumns: {
                ja: "作成した列グループどうしの横のアキを均等にします。",
                en: "Evens out the horizontal gaps between the column groups."
            },
            columnThreshold: {
                ja: "上下の隙間がこの値以内なら、同じ列とみなします。スライダーを動かすと結果がすぐ反映されます。",
                en: "Text objects with a vertical gap up to this value form one column. The canvas updates as you drag."
            }
        },
        button: {
            run:    { ja: "実行", en: "Run" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        alert: {
            noDocument:   { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noTextFrames: { ja: "テキストが選択されていません。", en: "No text object is selected." }
        }
    };

    // =========================================
    // 判定のしきい値 / Detection thresholds
    // =========================================

    /* スライダーで更新される、行・列の判定しきい値 / Updated as the sliders move */
    var rowGapThreshold = DEFAULT_GAP_THRESHOLD;
    var columnGapThreshold = DEFAULT_GAP_THRESHOLD;

    // 選択の収集と境界（再利用パーツ） / Selection items and bounds (reusable)

    /* 座標を同じと見なす許容値（pt） / Tolerance for treating coordinates as equal, in points */
    var SELECTION_ITEMS_TOLERANCE = 0.001;

    /**
     * 選択やコレクションを、オブジェクトの配列にそろえる
     * TextRange・PathItem は length を持つので、typename で1個か集まりかを見分ける
     * @param {*} source - doc.selection、配列、DOM のコレクション、または単独のオブジェクト
     * @returns {Array} オブジェクトの配列（空なら []）
     */
    function normalizeSelectionItems(source) {
        var items = [];
        if (!source) return items;
        var typeName = "";
        try { typeName = source.typename || ""; } catch (e) { /* 読めない種類 / unreadable kind */ }
        /* 単数形の typename は1個（PageItems などのコレクションは s で終わる）
           A singular typename is one object (collections such as PageItems end in s) */
        if (typeName && !/s$/.test(typeName)) return [source];
        if (typeof source.length !== "number") return items;
        for (var i = 0; i < source.length; i++) items.push(source[i]);
        return items;
    }

    /**
     * 文字カーソルの選択（TextRange）を、それを含むテキストフレームに読み替える
     * @param {TextRange} textRange - 文字の範囲
     * @returns {TextFrame|null} テキストフレーム（たどれなければ null）
     */
    function resolveTextRangeFrame(textRange) {
        var current = textRange;
        /* parent をたどる（深さは念のため制限） / Walk up the parents, with a safety limit */
        for (var depth = 0; depth < 10 && current; depth++) {
            try {
                if (current.typename === "TextFrame") return current;
                current = current.parent;
            } catch (e) {
                break;
            }
        }
        /* ストーリーの先頭フレームで代用する / Fall back to the first frame of the story */
        try {
            var storyFrames = textRange.story.textFrames;
            if (storyFrames.length > 0) return storyFrames[0];
        } catch (e2) { /* ストーリーを持たない / no story */ }
        return null;
    }

    /**
     * 選択から条件に合うオブジェクトを集める（グループ・レイヤーを再帰でたどり、重複は除く）
     * 条件に合ったオブジェクトの中へは進まない
     * @param {*} source - doc.selection、配列、コレクション、または単独のオブジェクト
     * @param {Object} [options] - 収集の設定
     * @param {function(PageItem): boolean} [options.accept] - 集める条件（既定はグループ・レイヤー以外すべて）
     * @param {boolean} [options.enterGroups] - グループの中をたどる（既定 true）
     * @param {boolean} [options.enterClipGroups] - クリップグループの中をたどる（既定は enterGroups と同じ）
     * @param {boolean} [options.enterCompoundPaths] - 複合パスの中のパスをたどる（既定 false）
     * @param {boolean} [options.textRangeToFrame] - 文字の選択をテキストフレームに読み替える（既定 true）
     * @param {boolean} [options.skipLocked] - ロックされたものを中ごと外す（既定 false）
     * @param {boolean} [options.skipHidden] - 非表示のものを中ごと外す（既定 false）
     * @param {boolean} [options.skipClipMasks] - クリッピングマスクを外す（既定 false）
     * @param {boolean} [options.skipGuides] - ガイドを外す（既定 false）
     * @param {boolean} [options.unique] - 同じ参照を1回だけにする（既定 true。数千件で遅ければ false）
     * @returns {Array} 集めたオブジェクト（前面→背面の順）
     */
    function collectSelectionItems(source, options) {
        var opts = options || {};
        var enterGroups = (opts.enterGroups !== false);
        var enterClipGroups = (opts.enterClipGroups === undefined) ? enterGroups : (opts.enterClipGroups === true);
        var accept = opts.accept || function (item) {
            return item.typename !== "GroupItem" && item.typename !== "Layer";
        };
        var collected = [];

        /**
         * 集めた配列に加える（unique のときは同じ参照を足さない）
         * @param {PageItem} item - 加えるオブジェクト
         * @returns {void}
         */
        function pushItem(item) {
            if (opts.unique !== false) {
                for (var k = 0; k < collected.length; k++) {
                    if (collected[k] === item) return;
                }
            }
            collected.push(item);
        }

        /**
         * 設定に従って外すオブジェクトか判定する
         * @param {PageItem} item - 判定するオブジェクト
         * @returns {boolean} 外すなら true
         */
        function isSkipped(item) {
            try {
                if (item.typename === "Layer") {
                    if (opts.skipLocked && item.locked) return true;
                    if (opts.skipHidden && !item.visible) return true;
                    return false;
                }
                if (opts.skipLocked && item.locked) return true;
                if (opts.skipHidden && item.hidden) return true;
                if (opts.skipGuides && item.guides === true) return true;
                if (opts.skipClipMasks && isClipMaskItem(item)) return true;
            } catch (e) {
                /* 読めないプロパティは「外さない」に倒す / Unreadable properties do not exclude */
            }
            return false;
        }

        /**
         * 1件をたどって集める
         * @param {PageItem} item - 対象のオブジェクト
         * @returns {void}
         */
        function visit(item) {
            if (!item) return;
            var typeName = "";
            try { typeName = item.typename; } catch (e) { return; }

            if (typeName === "TextRange" || typeName === "InsertionPoint") {
                if (opts.textRangeToFrame === false) {
                    if (accept(item)) pushItem(item);
                    return;
                }
                visit(resolveTextRangeFrame(item));
                return;
            }
            if (isSkipped(item)) return;
            if (accept(item)) {
                pushItem(item);
                return;
            }

            var children = null;
            if (typeName === "GroupItem") {
                var isClipped = false;
                try { isClipped = (item.clipped === true); } catch (e2) { }
                if (isClipped ? enterClipGroups : enterGroups) children = item.pageItems;
            } else if (typeName === "CompoundPathItem") {
                if (opts.enterCompoundPaths) children = item.pathItems;
            } else if (typeName === "Layer") {
                /* 重なり順はサブレイヤーとページアイテムで別々なので、ページアイテム→サブレイヤーの順にする
                   Page items and sublayers stack separately; visit page items first, then sublayers */
                walk(item.pageItems);
                walk(item.layers);
                return;
            }
            if (children) walk(children);
        }

        /**
         * 集まりの各要素をたどる
         * @param {*} list - 配列またはコレクション
         * @returns {void}
         */
        function walk(list) {
            var listItems = normalizeSelectionItems(list);
            for (var i = 0; i < listItems.length; i++) visit(listItems[i]);
        }

        walk(source);
        return collected;
    }

    /**
     * テキストフレームの種類を "point" / "area" / "path" で返す
     * @param {TextFrame} textFrame - テキストフレーム
     * @returns {string} 種類のキー（判定できなければ ""）
     */
    function getTextFrameKindKey(textFrame) {
        try {
            if (textFrame.kind === TextType.POINTTEXT) return "point";
            if (textFrame.kind === TextType.AREATEXT) return "area";
            if (textFrame.kind === TextType.PATHTEXT) return "path";
        } catch (e) { /* kind を読めない / kind is unreadable */ }
        return "";
    }

    /**
     * 選択からテキストフレームを集める（グループの中・文字カーソルの選択を含む）
     * @param {*} source - doc.selection など
     * @param {Object} [options] - collectSelectionItems と同じ設定に加えて次を受ける
     * @param {string[]} [options.kinds] - 集める種類（"point" / "area" / "path"。既定はすべて）
     * @returns {TextFrame[]} テキストフレーム（前面→背面の順）
     */
    function collectSelectionTextFrames(source, options) {
        var opts = {};
        var sourceOptions = options || {};
        for (var key in sourceOptions) {
            if (sourceOptions.hasOwnProperty(key)) opts[key] = sourceOptions[key];
        }
        var kindFilter = null;
        if (opts.kinds && opts.kinds.length) {
            kindFilter = {};
            for (var i = 0; i < opts.kinds.length; i++) kindFilter[opts.kinds[i]] = true;
        }
        opts.accept = function (item) {
            if (item.typename !== "TextFrame") return false;
            return !kindFilter || kindFilter[getTextFrameKindKey(item)] === true;
        };
        /* 種類で外したテキストは中をたどらない（accept が false でも子は無い） / Text frames have no children to walk */
        return collectSelectionItems(source, opts);
    }

    /**
     * 選択からパスを集める（グループの中を含む）
     * @param {*} source - doc.selection など
     * @param {Object} [options] - collectSelectionItems と同じ設定に加えて次を受ける
     * @param {string} [options.compoundPaths] - 複合パスの扱い。"children"（中のパス、既定）/ "whole"（複合パスごと）/ "skip"（外す）
     * @returns {Array} PathItem（"whole" のときは CompoundPathItem も）の配列
     */
    function collectSelectionPathItems(source, options) {
        var opts = {};
        var sourceOptions = options || {};
        for (var key in sourceOptions) {
            if (sourceOptions.hasOwnProperty(key)) opts[key] = sourceOptions[key];
        }
        var compoundMode = opts.compoundPaths || "children";
        opts.enterCompoundPaths = (compoundMode === "children");
        opts.accept = function (item) {
            if (item.typename === "PathItem") return true;
            return compoundMode === "whole" && item.typename === "CompoundPathItem";
        };
        return collectSelectionItems(source, opts);
    }

    /**
     * クリッピングマスク（クリップグループの型）か判定する
     * パスは clipping、複合パスは中の先頭パスの clipping、テキストは clipping が無いので「クリップグループの先頭」で見る
     * @param {PageItem} item - 判定するオブジェクト
     * @returns {boolean} マスクなら true
     */
    function isClipMaskItem(item) {
        try {
            if (item.typename === "PathItem") return item.clipping === true;
            if (item.typename === "CompoundPathItem") {
                return item.pathItems.length > 0 && item.pathItems[0].clipping === true;
            }
            if (item.typename === "TextFrame") {
                var parentGroup = item.parent;
                return parentGroup.typename === "GroupItem" && parentGroup.clipped === true &&
                    parentGroup.pageItems.length > 0 && parentGroup.pageItems[0] === item;
            }
        } catch (e) { /* 読めない種類はマスクではない / unreadable kinds are not masks */ }
        return false;
    }

    /**
     * クリップグループの型（マスク）を返す
     * フラグで探し、見つからなければ先頭（pageItems[0]）を返す（型は常に最前面。テキストの型はフラグを持たない）
     * @param {GroupItem} groupItem - 対象のグループ
     * @returns {PageItem|null} マスク（クリップグループでなければ null）
     */
    function getClipMaskItem(groupItem) {
        try {
            if (!groupItem || groupItem.typename !== "GroupItem" || groupItem.clipped !== true) return null;
            var groupChildren = groupItem.pageItems;
            if (groupChildren.length === 0) return null;
            for (var i = 0; i < groupChildren.length; i++) {
                var childType = groupChildren[i].typename;
                if ((childType === "PathItem" || childType === "CompoundPathItem") && isClipMaskItem(groupChildren[i])) {
                    return groupChildren[i];
                }
            }
            return groupChildren[0];
        } catch (e) {
            return null;
        }
    }

    /**
     * グループの中（入れ子を含む）にクリップグループがあるか判定する
     * @param {GroupItem} groupItem - 対象のグループ
     * @returns {boolean} あれば true
     */
    function hasClippedDescendant(groupItem) {
        try {
            var groupChildren = groupItem.pageItems;
            for (var i = 0; i < groupChildren.length; i++) {
                if (groupChildren[i].typename !== "GroupItem") continue;
                if (groupChildren[i].clipped === true || hasClippedDescendant(groupChildren[i])) return true;
            }
        } catch (e) { /* 中を読めない / cannot read the children */ }
        return false;
    }

    /**
     * 環境設定の［プレビュー境界を使用］を読む
     * @returns {boolean} オンなら true（読めなければ false）
     */
    function readUsePreviewBoundsPreference() {
        try {
            return app.preferences.getBooleanPreference("includeStrokeInBounds");
        } catch (e) {
            return false;
        }
    }

    /**
     * 見た目どおりの境界を返す。クリップグループはマスクの境界、
     * 中にクリップグループを含むグループは子の境界を合わせたもの（隠れた部分を含めない）
     * @param {PageItem} item - 対象のオブジェクト
     * @param {boolean} [usePreviewBounds] - true で visibleBounds、false で geometricBounds（省略時は環境設定に従う）
     * @returns {number[]|null} [左, 上, 右, 下] の新しい配列（測れなければ null）
     */
    function getClipAwareBounds(item, usePreviewBounds) {
        var usePreview = (usePreviewBounds === undefined || usePreviewBounds === null) ?
            readUsePreviewBoundsPreference() : (usePreviewBounds === true);
        try {
            var measuredItem = item;
            if (item.typename === "GroupItem") {
                var maskItem = getClipMaskItem(item);
                if (maskItem) {
                    measuredItem = maskItem;
                } else if (hasClippedDescendant(item)) {
                    /* グループ自体の効果（影など）の広がりは含まれなくなる
                       This leaves out the reach of effects applied to the group itself (drop shadows etc.) */
                    var childBounds = getClipAwareUnionBounds(filterMeasurableChildren(item.pageItems), usePreview);
                    if (childBounds) return childBounds;
                }
            }
            var bounds = usePreview ? measuredItem.visibleBounds : measuredItem.geometricBounds;
            return [bounds[0], bounds[1], bounds[2], bounds[3]];
        } catch (e) {
            return null;
        }
    }

    /**
     * 境界の計算に入れる子だけを残す（非表示とガイドを外す）
     * @param {*} childList - 子のコレクション
     * @returns {Array} 残した子
     */
    function filterMeasurableChildren(childList) {
        var childItems = normalizeSelectionItems(childList);
        var measurable = [];
        for (var i = 0; i < childItems.length; i++) {
            try {
                if (childItems[i].hidden === true || childItems[i].guides === true) continue;
            } catch (e) { /* 読めなければ残す / keep when unreadable */ }
            measurable.push(childItems[i]);
        }
        return measurable;
    }

    /**
     * 複数のオブジェクトを囲む外接範囲を返す（クリップグループはマスクで測る）
     * @param {*} items - オブジェクトの配列・コレクション・選択
     * @param {boolean} [usePreviewBounds] - true で visibleBounds、false で geometricBounds（省略時は環境設定に従う）
     * @returns {number[]|null} [左, 上, 右, 下]（測れるものが無ければ null）
     */
    function getClipAwareUnionBounds(items, usePreviewBounds) {
        var usePreview = (usePreviewBounds === undefined || usePreviewBounds === null) ?
            readUsePreviewBoundsPreference() : (usePreviewBounds === true);
        var itemList = normalizeSelectionItems(items);
        var unionBounds = null;
        for (var i = 0; i < itemList.length; i++) {
            var itemBounds = getClipAwareBounds(itemList[i], usePreview);
            if (!itemBounds) continue;
            if (!unionBounds) {
                unionBounds = itemBounds;
                continue;
            }
            if (itemBounds[0] < unionBounds[0]) unionBounds[0] = itemBounds[0];
            if (itemBounds[1] > unionBounds[1]) unionBounds[1] = itemBounds[1];
            if (itemBounds[2] > unionBounds[2]) unionBounds[2] = itemBounds[2];
            if (itemBounds[3] < unionBounds[3]) unionBounds[3] = itemBounds[3];
        }
        return unionBounds;
    }

    /**
     * 2つの座標を許容値つきで比べる
     * @param {number} valueA - 座標A（pt）
     * @param {number} valueB - 座標B（pt）
     * @param {number} [tolerance] - 許容値（pt、既定は SELECTION_ITEMS_TOLERANCE）
     * @returns {boolean} 差が許容値以下なら true
     */
    function isNearlySameCoordinate(valueA, valueB, tolerance) {
        var limit = (typeof tolerance === "number") ? tolerance : SELECTION_ITEMS_TOLERANCE;
        return Math.abs(valueA - valueB) <= limit;
    }

    /**
     * 2つの境界を許容値つきで比べる
     * @param {number[]} boundsA - [左, 上, 右, 下]
     * @param {number[]} boundsB - [左, 上, 右, 下]
     * @param {number} [tolerance] - 許容値（pt、既定は SELECTION_ITEMS_TOLERANCE）
     * @returns {boolean} 4辺とも許容値以内なら true
     */
    function areBoundsNearlyEqual(boundsA, boundsB, tolerance) {
        if (!boundsA || !boundsB) return false;
        for (var i = 0; i < 4; i++) {
            if (!isNearlySameCoordinate(boundsA[i], boundsB[i], tolerance)) return false;
        }
        return true;
    }

    // 選択の収集と境界（再利用パーツ）ここまで / End of the reusable selection items and bounds

    // =========================================
    // 境界の計測 / Bounds
    // =========================================

    /**
     * 探索方向における2つの境界の隙間を返す
     * 直交する方向にずれている組は同じ行／列とみなさないため、非常に大きい値を返す。
     * @param {number[]} boundsA - 一方の境界 [左, 上, 右, 下]
     * @param {number[]} boundsB - もう一方の境界
     * @param {string} direction - DIRECTION_HORIZONTAL または DIRECTION_VERTICAL
     * @returns {number} 隙間（対象外なら Number.MAX_VALUE）
     */
    function getGapAlongDirection(boundsA, boundsB, direction) {
        var horizontalGap = Math.max(0, Math.max(boundsB[0] - boundsA[2], boundsA[0] - boundsB[2]));
        var verticalGap = Math.max(0, Math.max(boundsB[3] - boundsA[1], boundsA[3] - boundsB[1]));

        if (direction === DIRECTION_HORIZONTAL) {
            return (verticalGap > 0) ? Number.MAX_VALUE : horizontalGap;
        }
        return (horizontalGap > 0) ? Number.MAX_VALUE : verticalGap;
    }

    /**
     * 2つの境界の重なり率（面積が大きい方に対する割合）を返す
     * @param {number[]} boundsA - 一方の境界 [左, 上, 右, 下]
     * @param {number[]} boundsB - もう一方の境界
     * @returns {number} 重なり率（重ならない場合は 0）
     */
    function getOverlapRatio(boundsA, boundsB) {
        var overlapWidth = Math.max(0, Math.min(boundsA[2], boundsB[2]) - Math.max(boundsA[0], boundsB[0]));
        var overlapHeight = Math.max(0, Math.min(boundsA[1], boundsB[1]) - Math.max(boundsA[3], boundsB[3]));
        var overlapArea = overlapWidth * overlapHeight;
        if (overlapArea <= 0) return 0;

        var areaA = (boundsA[2] - boundsA[0]) * (boundsA[1] - boundsA[3]);
        var areaB = (boundsB[2] - boundsB[0]) * (boundsB[1] - boundsB[3]);
        return overlapArea / Math.max(areaA, areaB);
    }

    // =========================================
    // 行・列の抽出 / Detecting rows and columns
    // =========================================

    /**
     * 隣接または重なっているオブジェクトをたどって1グループ分を集める
     * @param {number} startIndex - 起点のインデックス
     * @param {PageItem[]} targetItems - 対象オブジェクト
     * @param {boolean[]} visited - 走査済みフラグ
     * @param {PageItem[]} collectedItems - 集めたオブジェクトの受け皿
     * @param {string} direction - DIRECTION_HORIZONTAL または DIRECTION_VERTICAL
     * @param {number} gapThreshold - 同じ行／列とみなす隙間の上限
     * @returns {void}
     */
    function collectConnectedItems(startIndex, targetItems, visited, collectedItems, direction, gapThreshold) {
        visited[startIndex] = true;
        collectedItems.push(targetItems[startIndex]);

        var boundsA = targetItems[startIndex].visibleBounds;
        for (var j = 0; j < targetItems.length; j++) {
            if (visited[j]) continue;
            var boundsB = targetItems[j].visibleBounds;
            if (getOverlapRatio(boundsA, boundsB) > 0 ||
                getGapAlongDirection(boundsA, boundsB, direction) <= gapThreshold) {
                collectConnectedItems(j, targetItems, visited, collectedItems, direction, gapThreshold);
            }
        }
    }

    /**
     * 指定方向で隣接・重なっているオブジェクトをグループにまとめる
     * @param {PageItem[]} targetItems - 対象オブジェクト
     * @param {string} direction - DIRECTION_HORIZONTAL または DIRECTION_VERTICAL
     * @returns {Array} PageItem[] の配列（1グループ＝1行または1列）
     */
    function getConnectedGroups(targetItems, direction) {
        var gapThreshold = (direction === DIRECTION_HORIZONTAL) ? rowGapThreshold : columnGapThreshold;
        var itemGroups = [];
        var visited = [];

        for (var i = 0; i < targetItems.length; i++) visited[i] = false;

        for (var j = 0; j < targetItems.length; j++) {
            if (visited[j]) continue;
            var collectedItems = [];
            collectConnectedItems(j, targetItems, visited, collectedItems, direction, gapThreshold);
            itemGroups.push(collectedItems);
        }
        return itemGroups;
    }

    // =========================================
    // 整列・分配 / Aligning and distributing
    // =========================================

    /**
     * 行または列ごとに、その方向と直交する軸の中央でそろえる（グループ化はしない）
     * @param {PageItem[]} targetItems - 対象オブジェクト
     * @param {string} direction - DIRECTION_HORIZONTAL（行＝天地中央）または DIRECTION_VERTICAL（列＝左右中央）
     * @returns {void}
     */
    function alignGroupsToCenter(targetItems, direction) {
        if (!targetItems || targetItems.length === 0) return;

        var itemGroups = getConnectedGroups(targetItems, direction);
        for (var i = 0; i < itemGroups.length; i++) {
            if (itemGroups[i].length <= 1) continue;
            centerItemsInGroup(itemGroups[i], direction);
        }
    }

    /**
     * 1グループ分のオブジェクトを、方向と直交する軸の中央にそろえる
     * @param {PageItem[]} groupItems - 1グループ分のオブジェクト
     * @param {string} direction - DIRECTION_HORIZONTAL または DIRECTION_VERTICAL
     * @returns {void}
     */
    function centerItemsInGroup(groupItems, direction) {
        var combinedBounds = getClipAwareUnionBounds(groupItems, false);
        var isRow = (direction === DIRECTION_HORIZONTAL);
        /* 行は天地中央、列は左右中央にそろえる / rows center vertically, columns horizontally */
        var groupCenter = isRow ?
            (combinedBounds[1] + combinedBounds[3]) / 2 :
            (combinedBounds[0] + combinedBounds[2]) / 2;

        for (var i = 0; i < groupItems.length; i++) {
            var itemBounds = groupItems[i].geometricBounds;
            if (isRow) {
                groupItems[i].top += groupCenter - (itemBounds[1] + itemBounds[3]) / 2;
            } else {
                groupItems[i].left += groupCenter - (itemBounds[0] + itemBounds[2]) / 2;
            }
        }
    }

    /**
     * 並んだオブジェクトのアキを均等にする（両端の位置は保つ）
     * @param {PageItem[]} orderedItems - 対象オブジェクト（この関数内で並べ替える）
     * @param {string} direction - DIRECTION_HORIZONTAL（横に均等）または DIRECTION_VERTICAL（縦に均等）
     * @returns {void}
     */
    function distributeSpacingEvenly(orderedItems, direction) {
        if (!orderedItems || orderedItems.length <= 1) return;

        var isHorizontal = (direction === DIRECTION_HORIZONTAL);
        orderedItems.sort(isHorizontal ?
            function (itemA, itemB) { return itemA.geometricBounds[0] - itemB.geometricBounds[0]; } :
            function (itemA, itemB) { return itemB.geometricBounds[1] - itemA.geometricBounds[1]; });

        var totalItemSize = 0;
        for (var i = 0; i < orderedItems.length; i++) {
            totalItemSize += getItemSizeAlong(orderedItems[i], isHorizontal);
        }

        var combinedBounds = getClipAwareUnionBounds(orderedItems, false);
        var availableSpace = isHorizontal ?
            (combinedBounds[2] - combinedBounds[0]) :
            (combinedBounds[1] - combinedBounds[3]);
        var spacing = (availableSpace - totalItemSize) / (orderedItems.length - 1);

        var cursor = isHorizontal ? orderedItems[0].geometricBounds[0] : orderedItems[0].geometricBounds[1];
        for (var j = 0; j < orderedItems.length; j++) {
            var itemSize = getItemSizeAlong(orderedItems[j], isHorizontal);
            if (isHorizontal) {
                orderedItems[j].left = cursor;
                cursor += itemSize + spacing;
            } else {
                orderedItems[j].top = cursor;
                cursor -= itemSize + spacing;
            }
        }
    }

    /**
     * 指定軸方向のオブジェクトの寸法を返す
     * @param {PageItem} pageItem - 対象オブジェクト
     * @param {boolean} isHorizontal - 横方向なら true
     * @returns {number} 幅または高さ
     */
    function getItemSizeAlong(pageItem, isHorizontal) {
        var itemBounds = pageItem.geometricBounds;
        return isHorizontal ? (itemBounds[2] - itemBounds[0]) : (itemBounds[1] - itemBounds[3]);
    }

    // =========================================
    // グループ化 / Grouping
    // =========================================

    /**
     * 指定方向で抽出した行または列を、それぞれ1つのグループにまとめる
     * @param {string} direction - DIRECTION_HORIZONTAL または DIRECTION_VERTICAL
     * @param {boolean} centerBeforeGrouping - まとめる前に中央でそろえるなら true
     * @returns {GroupItem[]} 作成したグループ
     */
    function groupItemsByDirection(direction, centerBeforeGrouping) {
        if (app.documents.length === 0) return [];

        var selectedItems = app.activeDocument.selection;
        if (!selectedItems || selectedItems.length === 0) return [];

        var itemGroups = getConnectedGroups(selectedItems, direction);
        var createdGroups = [];

        for (var i = 0; i < itemGroups.length; i++) {
            var groupItems = itemGroups[i];
            if (groupItems.length <= 1) continue;

            /* グループ化でレイヤーが移るため、元のレイヤーを控えておく
               Grouping can move items between layers, so remember the original one */
            var originalLayer = groupItems[0].layer;

            if (centerBeforeGrouping) centerItemsInGroup(groupItems, direction);

            app.executeMenuCommand('deselectall');
            for (var j = 0; j < groupItems.length; j++) {
                groupItems[j].selected = true;
            }
            app.executeMenuCommand('group');

            var createdGroup = app.activeDocument.selection[0];
            createdGroup.layer = originalLayer;
            createdGroups.push(createdGroup);
        }

        app.redraw();

        app.activeDocument.selection = null;
        for (var k = 0; k < createdGroups.length; k++) {
            createdGroups[k].selected = true;
        }
        return createdGroups;
    }

    /**
     * グループ内のテキストだけを取り出す
     * @param {GroupItem} groupItem - 対象グループ
     * @returns {TextFrame[]} グループ直下のテキスト
     */
    function getTextFramesInGroup(groupItem) {
        var textFrames = [];
        var groupPageItems = groupItem.pageItems;
        for (var i = 0; i < groupPageItems.length; i++) {
            if (groupPageItems[i].typename === "TextFrame") textFrames.push(groupPageItems[i]);
        }
        return textFrames;
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * しきい値の表示用文字列を返す
     * @param {number} thresholdValue - スライダーの値（pt）
     * @returns {string} "10 pt" の形の文字列
     */
    function formatThreshold(thresholdValue) {
        return Math.round(thresholdValue) + " pt";
    }

    /**
     * ラベルと tooltip を LABELS の同じキーから引いてチェックボックスを追加する
     * @param {Panel} parentPanel - 追加先のパネル
     * @param {string} labelKey - checkbox.* と tooltip.* に共通のキー
     * @param {boolean} initialValue - 初期値
     * @returns {Checkbox} 追加したチェックボックス
     */
    function addOptionCheckbox(parentPanel, labelKey, initialValue) {
        var optionCheckbox = parentPanel.add("checkbox", undefined, getLabel("checkbox." + labelKey));
        optionCheckbox.helpTip = getLabel("tooltip." + labelKey);
        optionCheckbox.value = initialValue;
        return optionCheckbox;
    }

    /**
     * 行または列のパネル（揃え・グループ化・アキを均等に・しきい値）を組み立てる
     * @param {Window} parentWindow - 追加先のダイアログ
     * @param {object} axisKeys - LABELS のキー（panel / align / group / distribute / threshold）
     * @param {number} sliderMax - しきい値スライダーの最大値
     * @returns {object} alignCheckbox / groupCheckbox / distributeCheckbox / thresholdSlider / thresholdLabel
     */
    function addAxisPanel(parentWindow, axisKeys, sliderMax) {
        var axisPanel = parentWindow.add("panel", undefined, getLabel("panel." + axisKeys.panel));
        axisPanel.orientation = "column";
        axisPanel.alignChildren = "left";
        axisPanel.margins = PANEL_MARGINS;

        var axisControls = {
            alignCheckbox: addOptionCheckbox(axisPanel, axisKeys.align, true),
            groupCheckbox: addOptionCheckbox(axisPanel, axisKeys.group, false),
            distributeCheckbox: addOptionCheckbox(axisPanel, axisKeys.distribute, false)
        };
        /* 「アキを均等に」は「グループ化」をONにするまで使えない / enabled only while grouping */
        axisControls.distributeCheckbox.enabled = false;

        var thresholdSlider = axisPanel.add("slider", undefined, DEFAULT_GAP_THRESHOLD, 0, sliderMax);
        thresholdSlider.helpTip = getLabel("tooltip." + axisKeys.threshold);
        thresholdSlider.value = Math.min(DEFAULT_GAP_THRESHOLD, sliderMax);
        thresholdSlider.preferredSize.width = SLIDER_WIDTH;

        var thresholdLabel = axisPanel.add("statictext", undefined, formatThreshold(thresholdSlider.value));
        thresholdLabel.alignment = "center";
        thresholdLabel.characters = THRESHOLD_LABEL_CHARS;

        axisControls.thresholdSlider = thresholdSlider;
        axisControls.thresholdLabel = thresholdLabel;

        axisControls.alignCheckbox.onClick = function () {
            syncEnabledState(axisControls);
        };
        return axisControls;
    }

    /**
     * 「揃え」のON/OFFに合わせて、その行または列の他のコントロールをディムする
     * @param {object} axisControls - addAxisPanel() が返すコントロール
     * @returns {void}
     */
    function syncEnabledState(axisControls) {
        var isAligning = axisControls.alignCheckbox.value;
        axisControls.thresholdSlider.enabled = isAligning;
        axisControls.thresholdLabel.enabled = isAligning;
        axisControls.groupCheckbox.enabled = isAligning;
        if (!isAligning) {
            axisControls.groupCheckbox.value = false;
            axisControls.distributeCheckbox.enabled = false;
            axisControls.distributeCheckbox.value = false;
        }
    }

    /**
     * 「グループ化」をONにしたら他方の「グループ化」を外し、「アキを均等に」の有効状態を合わせる
     * 行と列を同時にグループ化はできないため / Rows and columns cannot both be grouped
     * @param {object} ownControls - クリックされた側のコントロール
     * @param {object} otherControls - もう一方（行なら列、列なら行）のコントロール
     * @returns {void}
     */
    function bindExclusiveGrouping(ownControls, otherControls) {
        ownControls.groupCheckbox.onClick = function () {
            var isGrouping = ownControls.groupCheckbox.value;
            if (isGrouping) otherControls.groupCheckbox.value = false;
            ownControls.distributeCheckbox.enabled = isGrouping;
            if (!isGrouping) ownControls.distributeCheckbox.value = false;
        };
    }

    /**
     * ダイアログを開く前の位置を控える（キャンセルで戻す）
     * @param {TextFrame[]} textFrames - 対象のテキスト
     * @returns {number[][]} 各テキストの geometricBounds の複製
     */
    function snapshotBounds(textFrames) {
        var originalBoundsList = [];
        for (var i = 0; i < textFrames.length; i++) {
            originalBoundsList.push(textFrames[i].geometricBounds.slice(0));
        }
        return originalBoundsList;
    }

    /**
     * 控えた位置へテキストを戻す
     * @param {TextFrame[]} textFrames - 対象のテキスト
     * @param {number[][]} originalBoundsList - snapshotBounds() の戻り値
     * @returns {void}
     */
    function restoreBounds(textFrames, originalBoundsList) {
        for (var i = 0; i < textFrames.length; i++) {
            textFrames[i].left = originalBoundsList[i][0];
            textFrames[i].top = originalBoundsList[i][1];
        }
    }

    /**
     * 整列・グループ化のダイアログを表示する
     * @param {TextFrame[]} textFrames - 対象のテキスト
     * @returns {object|null} ［実行］で確定した処理内容（キャンセルなら null）
     */
    function showTextGridDialog(textFrames) {
        /* スライダー操作で動いた分をキャンセルで戻すため / so Cancel can undo slider moves */
        var originalBoundsList = snapshotBounds(textFrames);
        var confirmedOptions = null;

        var combinedBounds = getClipAwareUnionBounds(textFrames, false);
        var totalWidth = combinedBounds[2] - combinedBounds[0];
        var totalHeight = combinedBounds[1] - combinedBounds[3];

        var alignDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        alignDialog.orientation = "column";
        alignDialog.alignChildren = "fill";

        var rowControls = addAxisPanel(alignDialog, {
            panel: "rows", align: "alignRows", group: "groupRows", distribute: "distributeRows", threshold: "rowThreshold"
        }, totalWidth || 100);
        var columnControls = addAxisPanel(alignDialog, {
            panel: "columns", align: "alignColumns", group: "groupColumns", distribute: "distributeColumns", threshold: "columnThreshold"
        }, totalHeight || 100);

        bindExclusiveGrouping(rowControls, columnControls);
        bindExclusiveGrouping(columnControls, rowControls);

        rowControls.thresholdSlider.onChanging = function () {
            rowControls.thresholdLabel.text = formatThreshold(rowControls.thresholdSlider.value);
            rowGapThreshold = rowControls.thresholdSlider.value;
            if (!rowControls.alignCheckbox.value) return;

            if (rowControls.groupCheckbox.value) {
                groupItemsByDirection(DIRECTION_HORIZONTAL, true);
            } else {
                alignGroupsToCenter(textFrames, DIRECTION_HORIZONTAL);
            }
            app.redraw();
        };

        columnControls.thresholdSlider.onChanging = function () {
            columnControls.thresholdLabel.text = formatThreshold(columnControls.thresholdSlider.value);
            columnGapThreshold = columnControls.thresholdSlider.value;
            if (!columnControls.alignCheckbox.value) return;

            if (columnControls.groupCheckbox.value) {
                var createdGroups = groupItemsByDirection(DIRECTION_VERTICAL, false);
                if (columnControls.distributeCheckbox.value) {
                    /* 各列グループの中で、テキストの左右のアキを均等にする
                       Even out the horizontal gaps inside each column group */
                    for (var i = 0; i < createdGroups.length; i++) {
                        distributeSpacingEvenly(getTextFramesInGroup(createdGroups[i]), DIRECTION_HORIZONTAL);
                    }
                }
            } else {
                alignGroupsToCenter(textFrames, DIRECTION_VERTICAL);
            }
            app.redraw();
        };

        /* ボタンエリア / Button row */
        var buttonRow = addButtonRow(alignDialog, { centered: true });

        var btnCancel = buttonRow.rowGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        btnCancel.onClick = function () {
            restoreBounds(textFrames, originalBoundsList);
            alignDialog.close(0);
        };

        var btnRun = buttonRow.rowGroup.add("button", undefined, getLabel("button.run"), { name: "ok" });
        btnRun.onClick = function () {
            confirmedOptions = {
                alignRows: rowControls.alignCheckbox.value,
                groupRows: rowControls.groupCheckbox.value,
                distributeRows: rowControls.distributeCheckbox.value,
                alignColumns: columnControls.alignCheckbox.value,
                groupColumns: columnControls.groupCheckbox.value,
                distributeColumns: columnControls.distributeCheckbox.value
            };
            alignDialog.close(1);
        };

        prepareDialogWindow(alignDialog, SCRIPT_NAME);
        return (alignDialog.show() === 1) ? confirmedOptions : null;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択したテキストを行・列で整列し、必要ならグループ化する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }

        /* 選択の直下にあるテキストだけを対象にする（グループの中・文字の選択は対象外）
           Only text frames selected directly (not inside groups, not a text-cursor selection) */
        var textFrames = collectSelectionTextFrames(app.activeDocument.selection, {
            enterGroups: false,
            textRangeToFrame: false,
            unique: false
        });
        if (textFrames.length === 0) {
            alert(getLabel("alert.noTextFrames"));
            return;
        }

        var confirmedOptions = showTextGridDialog(textFrames);
        if (!confirmedOptions) return;

        applyConfirmedOptions(textFrames, confirmedOptions);
    }

    /**
     * ダイアログで確定した内容を適用する
     * @param {TextFrame[]} textFrames - 対象のテキスト
     * @param {object} confirmedOptions - showTextGridDialog() が返す処理内容
     * @returns {void}
     */
    function applyConfirmedOptions(textFrames, confirmedOptions) {
        if (confirmedOptions.alignRows) {
            if (confirmedOptions.groupRows) {
                var rowGroups = groupItemsByDirection(DIRECTION_HORIZONTAL, true);
                if (confirmedOptions.distributeRows && rowGroups.length > 1) {
                    distributeSpacingEvenly(rowGroups, DIRECTION_VERTICAL);
                }
            } else {
                alignGroupsToCenter(textFrames, DIRECTION_HORIZONTAL);
            }
        }

        if (confirmedOptions.alignColumns) {
            if (confirmedOptions.groupColumns) {
                var columnGroups = groupItemsByDirection(DIRECTION_VERTICAL, false);
                if (confirmedOptions.distributeColumns && columnGroups.length > 1) {
                    distributeSpacingEvenly(columnGroups, DIRECTION_HORIZONTAL);
                }
            } else {
                alignGroupsToCenter(textFrames, DIRECTION_VERTICAL);
            }
        }
    }

    main();

})();
