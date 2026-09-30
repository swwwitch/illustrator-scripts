#target illustrator
#targetengine "FontWeightPickerEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択しているテキストのフォントと、名前が似ているファミリー（新ゴなら UD新ゴ など）のウェイトを一覧表示し、見本を見比べながら選んだフォントを適用します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/FontWeightPicker.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n255437cfdba0

### 注意

一覧は先頭の文字のフォントを基準にします。ウェイトの判定は TypefaceSampler.jsx と同じです。

### Overview

Lists the weights of the selected text's font and of families with similar names (UD Shin Go for Shin Go, etc.), and applies the font you choose while you compare samples.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/FontWeightPicker.md

### Notes

The list is based on the font of the first character. Weights are judged the same way as TypefaceSampler.jsx.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "FontWeightPicker";             /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.6";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-09-28";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-01";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/FontWeightPicker.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/FontWeightPicker.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n255437cfdba0"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* ファミリー名の先頭がこの語（メーカー名など）なら、2語目までを関連ファミリーの判定に使う（小文字で比較）
       When a family name starts with one of these words, the first two words decide related families */
    var FAMILY_PREFIX_WORDS = ["a-otf", "a-cid", "g-otf", "fot", "tb", "dnp", "itc", "adobe", "ff", "lt", "linotype", "monotype", "p22", "noto"];

    /* 中心の名前（「A-OTF 新ゴ Pr6N」なら「新ゴ」）を取り出すときに、前後から外す語（小文字で比較）
       Words stripped from both ends to get the core name ("新ゴ" out of "A-OTF 新ゴ Pr6N") */
    var CORE_NAME_PREFIX_WORDS = ["a", "a-otf", "a-cid", "p-otf", "g-otf", "fot", "tb", "dnp", "itc", "adobe", "ff", "lt", "linotype", "monotype", "p22"];
    var CORE_NAME_SUFFIX_WORDS = ["pro", "pron", "pr5", "pr5n", "pr6", "pr6n", "std", "stdn", "plus", "plusn", "jis2004", "jp", "kr", "sc", "tc"];
    var CORE_NAME_MIN_LENGTH   = 2;  /* 中心の名前で照合する最短の文字数 / shortest core name used for matching */

    /* フォント一覧の表示 / Font list display */
    var SHOW_RELATED_FAMILIES_DEFAULT = true;   /* ［類似フォントを表示］の初期値 / default of the related fonts checkbox */
    var SHOW_POSTSCRIPT_NAME_DEFAULT  = false;  /* ［PostScript名で表示］の初期値 / default of the PostScript name checkbox */

    /* ウェイトの見本（アートボード上の寸法はフレームの高さに対する比）/ Weight samples; sizes are ratios of the frame height */
    var WEIGHT_SAMPLES_DEFAULT  = true;                /* ［ウェイトの見本を右に並べる］の初期値 / default of the weight samples checkbox */
    var SAMPLE_ROW_GAP_RATIO    = 0.3;                 /* 見本の行の間隔 / row gap between the samples */
    var SAMPLE_COLUMN_GAP_RATIO = 1.5;                 /* 選択中のテキストと見本の間隔 / gap between the text and the samples */
    var LABEL_SIZE_RATIO        = 0.3;                 /* 見本の横のウェイト名の文字サイズ / size of the weight names */
    var LABEL_MIN_SIZE          = 6;                   /* ウェイト名の最小の文字サイズ（pt）/ minimum size of the weight names, in pt */
    var LABEL_GAP_RATIO         = 0.5;                 /* 見本とウェイト名の間隔 / gap between the samples and the weight names */
    var PREVIEW_LAYER_NAME      = "// weight-preview"; /* 見本などを置く一時レイヤー名 / temporary layer for the previews */

    /* 背景を白で覆う / Cover the background in white */
    var COVER_BACKGROUND_DEFAULT = true;      /* ［背景を白で覆う］の初期値 / default of the white cover checkbox */
    var COVER_BACKGROUND_OPACITY = 85;        /* 白い長方形の不透明度の初期値（%）/ default opacity of the white cover, in percent */
    var COVER_OPACITY_RANGE      = [0, 100];  /* 不透明度の範囲（%）/ range of the cover opacity */
    var COVER_VIEW_SCALE         = 3;         /* 白い長方形の大きさ（表示範囲に対する倍率）/ size of the white cover relative to the view */

    /* 画面にフィット / Fit to window */
    var FIT_VIEW_DEFAULT = true;       /* ［画面にフィット］の初期値 / default of the fit-to-window checkbox */
    var FIT_VIEW_PERCENT = 65;         /* ［画面にフィット］でプレビューが占める割合の初期値（%）/ default share of the window the previews fill */
    var FIT_VIEW_RANGE   = [10, 100];  /* ［画面にフィット］の割合の範囲（%）/ range of the fit-to-window share */

    /* Illustratorが受け付ける表示倍率の範囲（3.125%〜6400%） / Zoom range Illustrator accepts */
    var VIEW_ZOOM_RANGE = [0.03125, 64];

    // =========================================
    // レイアウト / Layout
    // =========================================

    var DIALOG_MARGINS        = 15;                /* ダイアログの余白 / dialog margins */
    var PANEL_MARGINS         = [15, 20, 15, 10];  /* パネルの余白 / panel margins */
    var FONT_PANEL_MARGINS    = [8, 20, 8, 10];    /* フォントパネルの余白（左右を詰める）/ font panel margins, narrower at the sides */
    var PANEL_SPACING         = 6;                 /* ラジオボタンの間隔 / spacing between radio buttons */
    var FONT_LIST_SIZE        = [220, 260];        /* フォント一覧の寸法 / size of the font list */
    var WEIGHT_RADIO_WIDTH    = 90;                /* ウェイトのラジオボタンの最小幅 / minimum width of the weight radio buttons */
    var RADIO_CHAR_WIDTH      = 7;                 /* ラジオの幅を見積もる1文字あたりの幅 / estimated width per character */
    var RADIO_INDICATOR_WIDTH = 30;                /* ラジオの丸印と余白の幅 / width of the radio indicator and padding */
    var OPTION_ROW_SPACING    = 6;                 /* オプション行の間隔 / spacing in the option rows */

    // ボタン行（再利用パーツ） / Button row (reusable)

    var BUTTON_ROW_TOP_MARGIN = 5; /* ボタン行の上の余白 / top margin of the button row */
    var BUTTON_ROW_SPACING = 10;   /* ボタンどうしの間隔 / spacing between buttons */
    var BUTTON_ROW_CENTER_MAX_WIDTH = 200; /* 右のボタンだけの行を中央に置く、ダイアログの内側の最大幅（px、左右の余白を除く）。広いダイアログは右揃え / max inner dialog width (px, margins excluded) that centers a right-only row; wider dialogs keep it right-aligned */

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
     * 左のグループにボタンが無い（右のボタンだけの）行を、ダイアログの幅に合わせて揃える。
     * 内側の幅（左右の余白を除く）が BUTTON_ROW_CENTER_MAX_WIDTH 以下なら左右中央、それより広ければ右揃えのまま。
     * 幅はレイアウトが決まるまで分からないので、ダイアログを表示した時点（show イベント）で判定する。
     * ボタンをすべて足したあと、show() の前に呼ぶ。centered で作った行や、左にボタンがある行はそのまま
     * @param {{rowGroup: Group, leftGroup: Group|null, rightGroup: Group|null}} buttonRow - addButtonRow() の戻り値
     * @returns {void}
     */
    function alignRightOnlyButtonRow(buttonRow) {
        if (!buttonRow.leftGroup || buttonRow.leftGroup.children.length > 0) return;
        var dialogWindow = buttonRow.rowGroup.window;
        dialogWindow.addEventListener("show", function () {
            if (!buttonRow.leftGroup) return;
            var btnRowGroup = buttonRow.rowGroup;
            /* 行の幅＝ダイアログの内側の幅（左右の余白を除く）/ The row spans the dialog's inner width (margins excluded) */
            if (!btnRowGroup.size || btnRowGroup.size.width > BUTTON_ROW_CENTER_MAX_WIDTH) return;
            /* 左のグループとスペーサーを外し、右のグループだけを中央に置く / Drop the left group and the spacer so only the right group remains, centered */
            btnRowGroup.remove(buttonRow.leftGroup);
            btnRowGroup.remove(btnRowGroup.children[0]); /* 左のグループを外すと先頭はスペーサー / the spacer is first once the left group is gone */
            btnRowGroup.alignment = ["center", "bottom"];
            btnRowGroup.alignChildren = ["center", "center"];
            buttonRow.leftGroup = null;
            dialogWindow.layout.layout(true);
        });
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

    // =========================================
    // ウェイト語句の定義 / Weight term definitions
    // （TypefaceSampler.jsx と同じ定義 / Shared with TypefaceSampler.jsx）
    // =========================================

    /* ウェイト語句の並び（インデックスが大きいほど太い）/ Weight terms ordered from thin to bold */
    var WEIGHT_GROUPS = [
        ["hairline", "hair"], // +0
        ["ultra thin", "ultrathin", "ut"], // +1
        ["thin", "th"], // +2
        ["default"], // +3
        ["ultralight", "ultra light", "ultlt", "ul"], // +4
        ["extralight", "extra light", "el", "xlight", "xl"], // +5
        ["lightsemi"], // +6
        ["light", "lt", "lite", "l"], // +7
        ["lb"], // +8
        ["book", "bk"], // +9
        ["n", "normal"], // +10
        ["middle"], // +11
        ["regular", "roman", "レギュラー", "r"], // +12
        ["rb"], // +13
        ["medium", "md", "ミディアム", "m"], // +14
        ["semibold", "semi bold", "sb"], // +15
        ["demibold", "demi bold", "db", "デミボールド", "demi", "d", "demixtra"], // +16
        ["bold", "bd", "ボールド", "b"], // +17
        ["extrabold", "extra bold", "xbold", "エクストラボールド", "e", "eb", "xb"], // +18
        ["heavy", "h"], // +19
        ["black"], // +20
        ["xblack", "extra black", "extrablack"], // +21
        ["ultra", "u", "ub", "ultra black", "ultrablack"] // +22
    ];

    /* 単独で使われたら Regular 扱いする装飾語句 / Decoration-only styles treated as Regular */
    var DECORATION_ONLY_STYLES = [
        "display", "compressed", "comp", "compact", "expanded", "extended", "semiextended",
        "ultracondensed", "extracondensed", "semicondensed", "cond", "condensed", "wide",
        "headline", "text", "low", "micro", "extra compressed",
        "semi expanded", "semiexpanded"
    ];

    /* 幅を表す複合語（Ultra Condensed など）。ultra / extra をウェイト語と取り違えないよう照合前に除く
       Width compounds such as "ultra condensed"; removed first so "ultra" is not read as a weight */
    var WIDTH_COMPOUND_PATTERN = /(^|\s)(ultra|extra|semi)\s+(condensed|cond|compressed|comp|expanded|extended)(?=\s|$)/g;

    /* WEIGHT_GROUPS における Regular のインデックス / Index of "regular" in WEIGHT_GROUPS */
    var REGULAR_GROUP_INDEX = (function() {
        for (var i = 0; i < WEIGHT_GROUPS.length; i++) {
            for (var j = 0; j < WEIGHT_GROUPS[i].length; j++) {
                if (WEIGHT_GROUPS[i][j] === "regular") return i;
            }
        }
        return 12; /* fallback */
    })();

    /* 複合語一致用に、長い語から順に並べた照合テーブル / Match table sorted by term length */
    var WEIGHT_TERM_PATTERNS = (function() {
        var termPatterns = [];
        for (var i = 0; i < WEIGHT_GROUPS.length; i++) {
            for (var j = 0; j < WEIGHT_GROUPS[i].length; j++) {
                var weightTerm = WEIGHT_GROUPS[i][j];
                /* \b は和文の前後で効かないため、英数字以外を境界とみなす / \b fails next to Japanese, so treat any non-alphanumeric as a boundary */
                termPatterns.push({
                    term: weightTerm,
                    groupIndex: i,
                    pattern: new RegExp("(?:^|[^a-z0-9])" + weightTerm.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&') + "(?=[^a-z0-9]|$)")
                });
            }
        }
        /* 同じ長さなら細い方を先に（並べ替えの結果を環境で変えない）/ Break ties by group so the order is deterministic */
        termPatterns.sort(function(a, b) {
            return (b.term.length - a.term.length) || (a.groupIndex - b.groupIndex);
        });
        return termPatterns;
    })();

    /**
     * スタイル文字列を照合用に正規化する
     * @param {string} rawStyle - font.style の値
     * @returns {string} 小文字化し、区切り記号を空白に置き換えた文字列
     */
    function normalizeStyle(rawStyle) {
        return (rawStyle || "").toLowerCase().replace(/[_\-]+/g, " ").replace(/^\s+|\s+$/g, "");
    }

    /**
     * スタイル文字列に一致する WEIGHT_GROUPS のインデックスを返す
     * 幅の複合語を除いてから、完全一致を優先し、なければ長い語から順に語の境界つきで照合する
     * @param {string} normalizedStyle - 正規化済みのスタイル文字列
     * @returns {number} 一致したインデックス。見つからない場合は -1
     */
    function getWeightGroupIndex(normalizedStyle) {
        var weightStyle = normalizedStyle.replace(WIDTH_COMPOUND_PATTERN, " ").replace(/\s+/g, " ").replace(/^\s+|\s+$/g, "");
        if (weightStyle === "") return -1;

        var i, j;
        for (i = 0; i < WEIGHT_GROUPS.length; i++) {
            for (j = 0; j < WEIGHT_GROUPS[i].length; j++) {
                if (weightStyle === WEIGHT_GROUPS[i][j]) return i;
            }
        }

        for (i = 0; i < WEIGHT_TERM_PATTERNS.length; i++) {
            if (WEIGHT_TERM_PATTERNS[i].pattern.test(weightStyle)) return WEIGHT_TERM_PATTERNS[i].groupIndex;
        }

        return -1;
    }

    /**
     * W3・W600・25 Ultra Light のような数値スタイルを読み取る
     * @param {string} normalizedStyle - 正規化済みのスタイル文字列
     * @returns {object|null} value（数値）と digitCount（桁数）。数値スタイルでなければ null
     */
    function getNumericWeight(normalizedStyle) {
        var numericMatch = normalizedStyle.match(/^w?(\d{1,3})(?=\D|$)/);
        if (!numericMatch) return null;
        return { value: parseInt(numericMatch[1], 10), digitCount: numericMatch[1].length };
    }

    // =========================================
    // ウェイト評価 / Weight scoring
    // （TypefaceSampler.jsx の並べ替え評価を、ウェイトと装飾語の加点に分けたもの
    //   TypefaceSampler.jsx's sort score, split into weight and decoration offset）
    // =========================================

    /**
     * スタイル文字列に対する基本ウェイトスコアを取得する
     * @param {string} normalizedStyle - 正規化済みのスタイル文字列
     * @param {string} postscriptName - 小文字化した PostScript 名
     * @param {string} familyName - 小文字化したファミリー名
     * @returns {number} ウェイトの評価値（小さいほど細い）
     */
    function getBaseWeightScore(normalizedStyle, postscriptName, familyName) {
        var styleWords = normalizedStyle.split(/\s+/);
        var i;

        var applyFrutigerCorrection = (/frutiger/i.test(familyName) && /ultralight/.test(normalizedStyle));

        /* W0〜W9、W000〜W999、先頭数値（例：25 Ultra Light）/ Numeric styles */
        var numericWeight = getNumericWeight(normalizedStyle);
        if (numericWeight) return numericWeight.value;

        /* 特例：HelveticaNeue, Tazugane, UniversNextPro + Ultra Light → 999 */
        if (
            (
                /helveticaneue/i.test(postscriptName) ||
                /tazugane/i.test(postscriptName) ||
                /universnextpro/i.test(postscriptName)
            ) &&
            /ultralight|ultra light|ultlt/i.test(normalizedStyle)
        ) {
            return 999;
        }

        /* 単独語が italic / oblique / wide → Regular 扱い */
        if (styleWords.length === 1 && /^(italic|oblique|it|wide)$/.test(styleWords[0])) {
            return 1000 + REGULAR_GROUP_INDEX;
        }

        /* 装飾語だけなら Regular 扱い */
        if (styleWords.length === 1) {
            for (i = 0; i < DECORATION_ONLY_STYLES.length; i++) {
                if (styleWords[0] === DECORATION_ONLY_STYLES[i]) return 1000 + REGULAR_GROUP_INDEX;
            }
        }

        /* 完全一致・複合語一致（長い語優先）/ Exact match, then longest-term match */
        var groupIndex = getWeightGroupIndex(normalizedStyle);
        if (groupIndex !== -1) {
            var weightScore = 1000 + groupIndex;
            if (applyFrutigerCorrection && groupIndex === 4) weightScore -= 5;
            return weightScore;
        }

        /* fallbackScore：Regular 扱い */
        var fallbackScore = 1000 + REGULAR_GROUP_INDEX;
        if (applyFrutigerCorrection) fallbackScore -= 5;
        return fallbackScore;
    }

    /**
     * 字幅・用途・イタリックなど、ウェイト以外の装飾語に対する加点を取得する
     * 同じ加点のフォント同士を「同じ系列」とみなす
     * @param {string} normalizedStyle - 正規化済みのスタイル文字列
     * @returns {number} 装飾語の加点（装飾なしは 0）
     */
    function getDecorationOffset(normalizedStyle) {
        var decorationOffset = 0;
        var styleWords = normalizedStyle.split(/\s+/);

        /* 装飾フラグ初期化 / Initialize decoration flags */
        var decorationFlags = {
            hasText: false,
            hasHeadline: false,
            hasCondensed: false,
            hasCn: false,
            hasExpanded: false,
            hasExtended: false,
            hasUltraCondensed: false,
            hasExtraCondensed: false,
            hasSemiCondensed: false,
            hasCompressed: false,
            hasExtraCompressed: false,
            hasCompact: false,
            hasDisplay: false,
            hasMicro: false,
            hasLow: false,
            hasWide: false
        };

        /* 装飾キーワードに応じたフラグ設定 / Set a flag for each decoration keyword */
        for (var i = 0; i < styleWords.length; i++) {
            var styleWord = styleWords[i];
            if (styleWord === "text") decorationFlags.hasText = true;
            if (styleWord === "headline") decorationFlags.hasHeadline = true;
            if (styleWord === "cond" || styleWord === "condensed") decorationFlags.hasCondensed = true;
            if (styleWord === "cn") decorationFlags.hasCn = true;
            if (styleWord === "expanded") decorationFlags.hasExpanded = true;
            if (styleWord === "extended") decorationFlags.hasExtended = true;
            if (styleWord === "semiextended" || (styleWord === "semi" && styleWords[i + 1] === "extended")) decorationFlags.hasExtended = true;
            if (styleWord === "semiexpanded" || (styleWord === "semi" && styleWords[i + 1] === "expanded")) decorationFlags.hasExpanded = true;
            if (styleWord === "ultracondensed" || (styleWord === "ultra" && styleWords[i + 1] === "condensed")) decorationFlags.hasUltraCondensed = true;
            if (styleWord === "extracondensed" || (styleWord === "extra" && styleWords[i + 1] === "condensed")) decorationFlags.hasExtraCondensed = true;
            if (styleWord === "semicondensed" || (styleWord === "semi" && styleWords[i + 1] === "condensed")) decorationFlags.hasSemiCondensed = true;
            if (styleWord === "compressed" || styleWord === "comp") decorationFlags.hasCompressed = true;
            if (styleWord === "extra" && styleWords[i + 1] === "compressed") decorationFlags.hasExtraCompressed = true;
            if (styleWord === "compact") decorationFlags.hasCompact = true;
            if (styleWord === "display") decorationFlags.hasDisplay = true;
            if (styleWord === "micro") decorationFlags.hasMicro = true;
            if (styleWord === "low") decorationFlags.hasLow = true;
            if (styleWord === "wide") decorationFlags.hasWide = true;
        }

        /* Italic 判定（全体 styleName に対して）/ Detect italic across the whole styleName */
        var isItalic = /italic|oblique|slanted|inclined|kursiv|\bit\b/.test(normalizedStyle);

        /* 加点処理（100刻み + 特例あり）/ Offsets in steps of 100, with exceptions */
        if (decorationFlags.hasDisplay) decorationOffset += 100;
        if (decorationFlags.hasCompressed) decorationOffset += 200;
        if (decorationFlags.hasCompact) decorationOffset += 300;
        if (decorationFlags.hasExpanded) decorationOffset += 400;
        if (decorationFlags.hasExtended) decorationOffset += 500;
        if (decorationFlags.hasUltraCondensed) decorationOffset += 600;
        if (decorationFlags.hasExtraCondensed) decorationOffset += 700;
        if (decorationFlags.hasSemiCondensed) decorationOffset += 850;

        /* Condensed系代表加点（複数条件一致でも一度のみ）/ Applied once even on multiple matches */
        if (
            decorationFlags.hasCondensed ||
            decorationFlags.hasCn ||
            decorationFlags.hasWide ||
            decorationFlags.hasSemiCondensed ||
            decorationFlags.hasExtraCompressed
        ) {
            decorationOffset += 900;
        }

        if (decorationFlags.hasHeadline) decorationOffset += 1000;
        if (decorationFlags.hasText) decorationOffset += 1100;
        if (decorationFlags.hasLow) decorationOffset += 1200;
        if (decorationFlags.hasMicro) decorationOffset += 1250;
        if (decorationFlags.hasWide) decorationOffset += 1275;
        if (decorationFlags.hasExtraCompressed) decorationOffset += 150; /* 特別加点 / Extra offset */
        if (isItalic) decorationOffset += 1300;

        return decorationOffset;
    }

    /**
     * フォントのウェイトと系列（装飾語の加点）を返す
     * @param {TextFont} textFont - 対象フォント
     * @returns {{weight: number, variant: number}|null} 判定結果。style が空（合成フォント・置換用の仮エントリ）は null
     */
    function getWeightInfo(textFont) {
        if (!textFont.style) return null;

        var styleName = normalizeStyle(textFont.style);
        var postscriptName = textFont.name.toLowerCase();

        /* 特例：PostScript名が「FuturaPT-Heavy」なら 1015 固定（加点処理なし）/ Fixed rank, no offsets */
        if (postscriptName === "futurapt-heavy") return { weight: 1015, variant: 0 };

        return {
            weight: getBaseWeightScore(styleName, postscriptName, textFont.family.toLowerCase()),
            variant: getDecorationOffset(styleName)
        };
    }

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
    // 対象テキストの収集 / Collect target text
    // =========================================

    /**
     * 選択から対象のテキスト範囲を集める（文字選択中はその範囲、オブジェクト選択ではグループ内も含む）
     * @param {Object} docSelection - doc.selection
     * @returns {TextRange[]} テキスト範囲の配列
     */
    function collectTextRanges(docSelection) {
        var textRanges = [];
        if (!docSelection) return textRanges;

        /* 文字選択中は TextRange が単独で返る / A TextRange is returned while editing text */
        if (docSelection.typename === "TextRange") {
            textRanges.push(docSelection);
            return textRanges;
        }

        /* グループの中のテキストも集める / Include text frames inside groups */
        var textFrames = collectSelectionTextFrames(docSelection);
        for (var i = 0; i < textFrames.length; i++) {
            textRanges.push(textFrames[i].textRange);
        }
        return textRanges;
    }

    /**
     * テキスト範囲の文字ごとに現在のフォントを控える
     * @param {TextRange[]} textRanges - 対象のテキスト範囲
     * @returns {Object[]} 文字と元フォントの組 {character, font}
     */
    function collectCharacterFonts(textRanges) {
        var characterEntries = [];
        for (var i = 0; i < textRanges.length; i++) {
            var characters = textRanges[i].characters;
            for (var j = 0; j < characters.length; j++) {
                characterEntries.push({ character: characters[j], font: characters[j].characterAttributes.textFont });
            }
        }
        return characterEntries;
    }

    /**
     * 選択から基準のテキストフレーム（先頭の文字を含むもの）を返す
     * @param {Object} docSelection - doc.selection
     * @returns {TextFrame|null} テキストフレーム。無ければ null
     */
    function findBaseTextFrame(docSelection) {
        if (!docSelection) return null;

        /* 文字選択中は所属するストーリーの先頭のフレーム / While editing text, the story's first frame */
        if (docSelection.typename === "TextRange") {
            var storyFrames = docSelection.story.textFrames;
            return (storyFrames.length > 0) ? storyFrames[0] : null;
        }

        /* グループの中も含めて最初のテキストフレーム / The first text frame, including inside groups */
        var textFrames = collectSelectionTextFrames(docSelection);
        return (textFrames.length > 0) ? textFrames[0] : null;
    }

    // =========================================
    // 関連ファミリーとウェイトの一覧 / Related families and weights
    // =========================================

    /**
     * 関連ファミリーを判定するためのキーを返す（先頭の語。メーカー名などで始まるときは2語目まで）
     * @param {string} familyName - ファミリー名
     * @returns {string} 小文字のキー
     */
    function getRelatedFamilyKey(familyName) {
        var familyWords = familyName.toLowerCase().split(/\s+/);
        for (var i = 0; i < FAMILY_PREFIX_WORDS.length; i++) {
            if (familyWords[0] === FAMILY_PREFIX_WORDS[i] && familyWords.length > 1) {
                return familyWords[0] + " " + familyWords[1];
            }
        }
        return familyWords[0];
    }

    /**
     * 前後のメーカー名・文字セット名を外した、ファミリーの中心の名前を返す（「A P-OTF 新ゴ ProN」→「新ゴ」）
     * @param {string} familyName - ファミリー名
     * @returns {string} 小文字の中心の名前（残らなければ空文字）
     */
    function getCoreFamilyName(familyName) {
        var familyWords = familyName.toLowerCase().replace(/^fot-/, "").split(/\s+/);
        while (familyWords.length > 1 && isListedWord(familyWords[0], CORE_NAME_PREFIX_WORDS)) familyWords.shift();
        while (familyWords.length > 1 && isListedWord(familyWords[familyWords.length - 1], CORE_NAME_SUFFIX_WORDS)) familyWords.pop();
        return familyWords.join(" ");
    }

    /**
     * 語が一覧に含まれるかを返す
     * @param {string} word - 調べる語（小文字）
     * @param {string[]} wordList - 一覧
     * @returns {boolean} 含まれれば true
     */
    function isListedWord(word, wordList) {
        for (var i = 0; i < wordList.length; i++) {
            if (wordList[i] === word) return true;
        }
        return false;
    }

    /**
     * 基準のファミリーと関連するかを返す（先頭の語が同じか、基準の中心の名前を含むもの。新ゴなら UD新ゴ も含む）
     * @param {string} familyName - 調べるファミリー名
     * @param {string} relatedKey - 基準の getRelatedFamilyKey()
     * @param {string} baseCoreName - 基準の getCoreFamilyName()
     * @returns {boolean} 関連すれば true
     */
    function isRelatedFamily(familyName, relatedKey, baseCoreName) {
        if (getRelatedFamilyKey(familyName) === relatedKey) return true;
        if (baseCoreName.length < CORE_NAME_MIN_LENGTH) return false;
        return getCoreFamilyName(familyName).indexOf(baseCoreName) !== -1;
    }

    /**
     * 関連ファミリーのフォントを app.textFonts から一度の走査で集める
     * @param {string} baseFamilyName - 基準のファミリー名
     * @returns {{familyNames: string[], membersByFamily: Object}} 並べ替えたファミリー名と、ファミリー名 → [{font, weight, variant}]
     */
    function collectRelatedFamilies(baseFamilyName) {
        var relatedKey = getRelatedFamilyKey(baseFamilyName);
        var baseCoreName = getCoreFamilyName(baseFamilyName);
        var familyNames = [];
        var membersByFamily = {};
        var isRelatedByFamily = {};
        var textFonts = app.textFonts;
        for (var i = 0; i < textFonts.length; i++) {
            var textFont = textFonts[i];
            var family = textFont.family;
            /* 判定はファミリーごとに1回だけ / Judge each family once */
            if (!isRelatedByFamily.hasOwnProperty(family)) {
                isRelatedByFamily[family] = isRelatedFamily(family, relatedKey, baseCoreName);
                if (isRelatedByFamily[family]) {
                    membersByFamily[family] = [];
                    familyNames.push(family);
                }
            }
            if (!isRelatedByFamily[family]) continue;

            /* 合成フォント・置換用の仮エントリは除く / Skip composite fonts and placeholders */
            var weightInfo = getWeightInfo(textFont);
            if (weightInfo) membersByFamily[family].push({ font: textFont, weight: weightInfo.weight, variant: weightInfo.variant });
        }

        /* ウェイトを持たないファミリー（合成フォントなど）は一覧に出さない / Drop families without weights */
        var listedFamilyNames = [];
        for (var j = 0; j < familyNames.length; j++) {
            if (membersByFamily[familyNames[j]].length > 0) listedFamilyNames.push(familyNames[j]);
        }
        listedFamilyNames.sort();
        return { familyNames: listedFamilyNames, membersByFamily: membersByFamily };
    }

    /**
     * ファミリーから右に並べるウェイトを選ぶ（基準と同じ系列、無ければ装飾なし、それも無ければ最小の系列）
     * @param {Object[]} familyMembers - ファミリーのフォント [{font, weight, variant}]
     * @param {number} baseVariant - 基準の系列
     * @returns {Object[]} 細い順に並べたフォント [{font, weight, variant}]
     */
    function pickSeriesMembers(familyMembers, baseVariant) {
        var targetVariant = null;
        var smallestVariant = null;
        for (var i = 0; i < familyMembers.length; i++) {
            var memberVariant = familyMembers[i].variant;
            if (memberVariant === baseVariant) targetVariant = baseVariant;
            if (smallestVariant === null || memberVariant < smallestVariant) smallestVariant = memberVariant;
        }
        /* 装飾なし（0）があれば最小なのでそれが選ばれる / Plain (0) wins when present because it is the smallest */
        if (targetVariant === null) targetVariant = smallestVariant;

        var seriesMembers = [];
        for (var j = 0; j < familyMembers.length; j++) {
            if (familyMembers[j].variant === targetVariant) seriesMembers.push(familyMembers[j]);
        }
        /* TypefaceSampler.jsx の sortFontGroups() と同じく、同点は PostScript 名順（件数が少ないので比較関数つきで足りる）
           Ties broken by PostScript name, as in TypefaceSampler.jsx's sortFontGroups() */
        seriesMembers.sort(function (a, b) {
            if (a.weight !== b.weight) return a.weight - b.weight;
            return (a.font.name < b.font.name) ? -1 : ((a.font.name > b.font.name) ? 1 : 0);
        });
        return seriesMembers;
    }

    /**
     * 選択中の文字で使われているフォント名を集める
     * @param {Object[]} characterEntries - collectCharacterFonts() の結果
     * @returns {Object} フォント名をキーにした集合
     */
    function collectUsedFontNames(characterEntries) {
        var usedFontNames = {};
        for (var i = 0; i < characterEntries.length; i++) {
            usedFontNames[characterEntries[i].font.name] = true;
        }
        return usedFontNames;
    }

    /**
     * 初期選択の番号を返す
     * 基準のフォントを含む一覧では一つ太いもの（無ければ最も近い細いもの）、ほかのファミリーでは基準のウェイトに最も近いもの
     * @param {Object[]} seriesMembers - pickSeriesMembers() の結果
     * @param {Object} usedFontNames - 使用中のフォント名の集合
     * @param {TextFont} baseFont - 基準のフォント
     * @param {number} baseWeight - 基準のウェイト
     * @returns {number} 選択する番号。選べるものが無ければ -1
     */
    function findInitialIndex(seriesMembers, usedFontNames, baseFont, baseWeight) {
        var baseIndex = -1;
        for (var i = 0; i < seriesMembers.length; i++) {
            if (seriesMembers[i].font.name === baseFont.name) { baseIndex = i; break; }
        }

        if (baseIndex !== -1) {
            for (var j = baseIndex + 1; j < seriesMembers.length; j++) {
                if (!usedFontNames[seriesMembers[j].font.name]) return j;
            }
            for (var k = baseIndex - 1; k >= 0; k--) {
                if (!usedFontNames[seriesMembers[k].font.name]) return k;
            }
            return -1;
        }

        var nearestIndex = -1;
        for (var n = 0; n < seriesMembers.length; n++) {
            if (usedFontNames[seriesMembers[n].font.name]) continue;
            if (nearestIndex === -1 || Math.abs(seriesMembers[n].weight - baseWeight) < Math.abs(seriesMembers[nearestIndex].weight - baseWeight)) {
                nearestIndex = n;
            }
        }
        return nearestIndex;
    }

    // UI の明暗（再利用パーツ） / UI theme (reusable)

    /**
     * UI がダークテーマかどうかを判定する（Illustrator は uiBrightness、InDesign は uiBrightnessPreference）
     * @returns {boolean} ダークなら true。取得できない環境では false（明るいUI扱い）
     */
    function isDarkUI() {
        try {
            if (app.preferences && app.preferences.getRealPreference) {
                return app.preferences.getRealPreference("uiBrightness") <= 0.5; /* Illustrator */
            }
            return app.generalPreferences.uiBrightnessPreference <= 0.5; /* InDesign */
        } catch (e) {
            return false;
        }
    }

    // UI の明暗（再利用パーツ）ここまで / End of the reusable UI theme

    // ステップボタン（再利用パーツ） / Stepper buttons (reusable)

    // -----------------------------------------
    // ステップボタンの寸法・増減量 / Stepper metrics and steps
    // -----------------------------------------
    var STEPPER_BUTTON_WIDTH   = 20;  /* ∧∨ボタンの幅 / button width */
    var STEPPER_BUTTON_HEIGHT  = 11;  /* ∧∨ボタン1つの高さ（2つ重ねた全体の高さは22） / button height (22 for the pair) */
    var STEPPER_CORNER_RADIUS  = 2;   /* 枠の角丸の半径（ScriptUIは円弧を描けないため短い線分で近似） / corner radius, approximated with segments */
    var STEPPER_FIELD_SPACING  = 3;   /* 項目名と∧∨の間隔 / spacing between the label and the stepper */
    var STEPPER_SIDE_MARGIN    = 3;   /* ∧∨の左に足す余白（右は入力欄に突き合わせる） / extra space left of the stepper */
    var STEPPER_SHIFT_MULTIPLE = 10;  /* shift＋クリックでそろえる倍数 / Shift-click snaps to multiples of this */
    var STEPPER_OPTION_STEP    = 0.1; /* option＋クリックの増減量 / Option-click step */

    // -----------------------------------------
    // ステップボタンの配色 / Stepper colors
    // -----------------------------------------
    var STEPPER_UI_DARK           = isDarkUI();
    /* UIの明るさは4段階あり、段階ごとに背景色が違う。どの段階でも背景に対する差で見せるよう、黒・白の半透明を重ねる。
       ダーク側は Illustrator 標準のスピナー（［グリッドに分割］）で実測、明るい側は最も明るい段階（背景 約0.94）から逆算
       UI brightness has four levels with different backgrounds, so colors are translucent overlays that follow the
       dialog background. Dark values are measured from Illustrator's own spinner; light values derived for the lightest level */
    var STEPPER_FILL_COLOR        = STEPPER_UI_DARK ? [0, 0, 0, 0.10]  : [1, 1, 1, 0.50];  /* 地 / background */
    var STEPPER_FRAME_COLOR       = STEPPER_UI_DARK ? [1, 1, 1, 0.07]  : [0, 0, 0, 0.10];  /* 枠線 / frame */
    var STEPPER_PRESSED_COLOR     = STEPPER_UI_DARK ? [1, 1, 1, 0.12]  : [0, 0, 0, 0.13];  /* 押下中 / pressed */
    var STEPPER_CHEVRON_COLOR     = STEPPER_UI_DARK ? [1, 1, 1, 1]     : [0, 0, 0, 0.70];  /* 山形の線 / chevron */
    var STEPPER_DIM_FILL_COLOR    = STEPPER_UI_DARK ? [1, 1, 1, 0.035] : [1, 1, 1, 0.30];  /* 無効時の地 / background when disabled */
    var STEPPER_DIM_FRAME_COLOR   = STEPPER_UI_DARK ? [1, 1, 1, 0.035] : [0, 0, 0, 0.05];  /* 無効時の枠線（ダークは地と同じで見せない） / frame when disabled */
    var STEPPER_DIM_CHEVRON_COLOR = STEPPER_UI_DARK ? [1, 1, 1, 0.20]  : [0, 0, 0, 0.25];  /* 無効時の山形 / chevron when disabled */

    // -----------------------------------------
    // 数値欄を作る（外から呼ぶ関数） / Public API
    // -----------------------------------------
    /**
     * 「項目名・∧∨・入力欄」をひと組にした数値欄を追加する。
     * ↑↓キーでも∧∨と同じように増減する。直接入力した値も、フォーカスが外れたときに
     * 整数化・下限・上限・単位（「20 mm」の形）へそろえ、数値でなければ直前の値に戻す
     * @param {Group|Panel} parent - 追加先
     * @param {Object} fieldOptions - label（コロン込みの項目名）/ labelWidth / text / characters /
     *     step / min / max / integer（true で整数のみ）/ unit / onStep
     * @returns {EditText} 入力欄（項目名は .fieldLabel、∧∨は .stepperGroup で参照できる）
     */
    function addSteppedField(parent, fieldOptions) {
        var fieldRowGroup = parent.add("group");
        fieldRowGroup.orientation = "row";
        fieldRowGroup.alignChildren = ["left", "center"];
        fieldRowGroup.spacing = STEPPER_FIELD_SPACING;

        var fieldLabel = fieldRowGroup.add("statictext", undefined, fieldOptions.label || "");
        if (fieldOptions.labelWidth) {
            fieldLabel.preferredSize.width = fieldOptions.labelWidth;
            fieldLabel.justify = "right";
        }

        /* ∧∨と入力欄は隙間0で突き合わせる / butt the stepper against the field */
        var stepperInputGroup = fieldRowGroup.add("group");
        stepperInputGroup.orientation = "row";
        stepperInputGroup.alignChildren = ["left", "center"];
        stepperInputGroup.spacing = 0;
        stepperInputGroup.margins = 0;

        var numberInput;
        var stepperGroup = addStepper(stepperInputGroup, function () { return numberInput; }, fieldOptions);
        numberInput = stepperInputGroup.add("edittext", undefined, fieldOptions.text || "");
        numberInput.characters = fieldOptions.characters || 6;
        numberInput.fieldLabel = fieldLabel;
        numberInput.stepperGroup = stepperGroup;

        /* ↑↓キーも∧∨と同じ処理で増減する（増減量・下限・上限・単位・修飾キーをそろえる） / arrow keys share the stepper's logic */
        bindSteppedArrowKeys(numberInput, stepperGroup);

        /* 直接入力をそろえる。数値でなければ直前の値に戻す / normalize typed values; revert non-numbers */
        numberInput.lastValidText = numberInput.text;
        numberInput.onChange = function () {
            var value = parseFloat(numberInput.text);
            if (isNaN(value)) {
                numberInput.text = numberInput.lastValidText;
                return;
            }
            writeSteppedValue(numberInput, value, fieldOptions);
        };
        return numberInput;
    }

    /**
     * 数値欄の有効／無効を、項目名・∧∨ごとまとめて切り替える
     * @param {EditText} numberInput - addSteppedField() で作った入力欄
     * @param {boolean} isEnabled - 有効にするなら true
     * @returns {void}
     */
    function setSteppedFieldEnabled(numberInput, isEnabled) {
        numberInput.enabled = isEnabled;
        numberInput.fieldLabel.enabled = isEnabled;
        numberInput.stepperGroup.enabled = isEnabled;
        /* ∧∨は自作描画なので、描き直してディム表示を切り替える / redraw the custom-drawn buttons to update the dimming */
        for (var i = 0; i < numberInput.stepperGroup.children.length; i++) {
            redrawStepperGroup(numberInput.stepperGroup.children[i]);
        }
    }

    /**
     * 入力欄の値を増減する∧∨ボタンを、隙間なく縦に積んで追加する
     * @param {Group|Panel} parent - 追加先
     * @param {Function} getNumberInput - 対象の入力欄を返す関数（入力欄を∧∨より後に作れるよう、クリック時に引く）
     * @param {Object} stepOptions - step（増減量）/ min / max / integer / unit（例 " mm"）/ onStep(numberInput)
     * @returns {Group} ∧∨をまとめた group（.stepBy(direction) で同じ増減を呼べる）
     */
    function addStepper(parent, getNumberInput, stepOptions) {
        var stepperGroup = parent.add("group");
        stepperGroup.orientation = "column";
        stepperGroup.spacing = 0; /* 2つのボタンをつなげて1つの枠に見せる / join the buttons into one frame */
        stepperGroup.margins = [STEPPER_SIDE_MARGIN, 0, 0, 0]; /* 右は入力欄に突き合わせる / butt against the field on the right */
        stepperGroup.alignment = ["left", "center"];

        /**
         * 入力欄の値を増減する（shift を押しながらなら STEPPER_SHIFT_MULTIPLE の倍数へ、option なら STEPPER_OPTION_STEP ずつ。下限・上限で止める）
         * @param {number} direction - 増やすなら 1、減らすなら -1
         * @returns {void}
         */
        function stepBy(direction) {
            var numberInput = getNumberInput();
            if (!isStepperEnabledInTree(numberInput)) return; /* 入力欄か親が無効の間は動かさない */
            var value = parseFloat(numberInput.text);
            if (isNaN(value)) value = 0;
            writeSteppedValue(numberInput, computeSteppedValue(value, direction, stepOptions), stepOptions);
            if (stepOptions.onStep) stepOptions.onStep(numberInput);
        }

        /* 整数の欄では option＋クリックの0.1刻みが効かないので、説明から外す / integer fields have no 0.1 step */
        var upTooltip = stepOptions.integer ? LABELS.tooltip.stepUpInteger : LABELS.tooltip.stepUp;
        var downTooltip = stepOptions.integer ? LABELS.tooltip.stepDownInteger : LABELS.tooltip.stepDown;
        makeStepperChevronButton(stepperGroup, "up", function () { stepBy(1); }).helpTip = getLabel(upTooltip);
        makeStepperChevronButton(stepperGroup, "down", function () { stepBy(-1); }).helpTip = getLabel(downTooltip);
        stepperGroup.stepBy = stepBy; /* ↑↓キーからも同じ処理で増減できるよう公開 / shared with the arrow keys */
        return stepperGroup;
    }

    /**
     * 入力欄の↑↓キーを、∧∨と同じ処理で増減させる。ほかのキーは素通し
     * @param {EditText} numberInput - 対象の入力欄
     * @param {Group} stepperGroup - addStepper() で作った∧∨
     * @returns {void}
     */
    function bindSteppedArrowKeys(numberInput, stepperGroup) {
        numberInput.addEventListener("keydown", function (event) {
            if (event.keyName !== "Up" && event.keyName !== "Down") return;
            stepperGroup.stepBy(event.keyName === "Up" ? 1 : -1);
            event.preventDefault(); /* カーソル移動を止める / keep the caret from moving */
        });
    }

    // -----------------------------------------
    // 値の計算 / Value helpers
    // -----------------------------------------
    /**
     * 押された修飾キーに応じて、1回分増減した値を返す
     * （shift なら STEPPER_SHIFT_MULTIPLE の倍数へ、option なら STEPPER_OPTION_STEP ずつ、それ以外は step の倍数へ（1.5→2、1.5→1）。
     * 整数の欄では option を無視して step の倍数へ）
     * @param {number} value - 元の値
     * @param {number} direction - 増やすなら 1、減らすなら -1
     * @param {Object} stepOptions - step（通常の増減量。省略時は 1）/ integer
     * @returns {number} 増減した値（下限・上限は未適用）
     */
    function computeSteppedValue(value, direction, stepOptions) {
        var keyState = ScriptUI.environment.keyboardState;
        if (keyState.shiftKey) return snapStepperToNextMultiple(value, STEPPER_SHIFT_MULTIPLE, direction);
        if (keyState.altKey && !stepOptions.integer) return value + direction * STEPPER_OPTION_STEP;
        return snapStepperToNextMultiple(value, stepOptions.step || 1, direction);
    }

    /**
     * 値を、指定した方向にある次の倍数へ移す（230→240、232→240、下げるときは 232→230、230→220）
     * @param {number} value - 元の値
     * @param {number} multiple - 倍数の単位（例 10）
     * @param {number} direction - 上げるなら 1、下げるなら -1
     * @returns {number} 移した値
     */
    function snapStepperToNextMultiple(value, multiple, direction) {
        /* 0.29 / 0.01 = 28.999… のような浮動小数の誤差で同じ値に戻らないよう、商を丸めてから切り捨て・切り上げる
           round the quotient first so float error (0.29 / 0.01 = 28.999…) does not step back to the same value */
        var quotient = Math.round(value / multiple * 1e6) / 1e6;
        if (direction > 0) return Math.round((Math.floor(quotient) + 1) * multiple * 1e6) / 1e6;
        return Math.round((Math.ceil(quotient) - 1) * multiple * 1e6) / 1e6;
    }

    /**
     * 値を下限・上限の範囲に収める
     * @param {number} value - 数値
     * @param {Object} rangeOptions - min / max（どちらも省略可）
     * @returns {number} 範囲に収めた値
     */
    function clampSteppedValue(value, rangeOptions) {
        if (rangeOptions.min !== undefined && value < rangeOptions.min) return rangeOptions.min;
        if (rangeOptions.max !== undefined && value > rangeOptions.max) return rangeOptions.max;
        return value;
    }

    /**
     * 値を整数化・下限・上限でそろえ、単位を付けて入力欄に書き込む（直前の正しい値としても控える）
     * @param {EditText} numberInput - 書き込む入力欄
     * @param {number} value - 数値
     * @param {Object} valueOptions - integer / min / max / unit（どれも省略可）
     * @returns {void}
     */
    function writeSteppedValue(numberInput, value, valueOptions) {
        numberInput.text = formatSteppedValue(value, valueOptions);
        numberInput.lastValidText = numberInput.text;
    }

    /**
     * 値を整数化・下限・上限でそろえ、丸めて単位を付けた表示用の文字列にする。
     * 整数化してから下限で止めるので、「整数・下限1」の欄に 0.4 が入っても 1 になる
     * @param {number} value - 数値
     * @param {Object} valueOptions - integer / min / max / unit（どれも省略可）
     * @returns {string} 入力欄に入れる文字列（例 "20 mm"）
     */
    function formatSteppedValue(value, valueOptions) {
        if (valueOptions.integer) value = Math.round(value);
        return formatStepperNumber(clampSteppedValue(value, valueOptions)) + (valueOptions.unit || "");
    }

    /**
     * 小数第2位で丸めた数値を文字列で返す
     * @param {number} value - 数値
     * @returns {string} 表示用の数値文字列
     */
    function formatStepperNumber(value) {
        return String(Math.round(value * 100) / 100);
    }

    // -----------------------------------------
    // ∧∨ボタンの描画 / Drawing
    // -----------------------------------------
    /**
     * 山形（∧／∨）の極小ボタンを作成する。
     * 上下2つを隙間なく積んで1つの枠に見えるよう、枠線は外側の辺だけ描き（上ボタンは上側、下ボタンは下側）、
     * 継ぎ目に線は引かない
     * @param {Group|Panel} parent - 追加先
     * @param {string} direction - "up" または "down"
     * @param {Function} onClickFn - クリック時の処理
     * @returns {Group} ボタンとして使う group
     */
    function makeStepperChevronButton(parent, direction, onClickFn) {
        var buttonWidth = STEPPER_BUTTON_WIDTH;
        var buttonHeight = STEPPER_BUTTON_HEIGHT;
        var isUp = (direction === "up");
        var chevronBox = parent.add("group");
        chevronBox.margins = 0;
        chevronBox.spacing = 0;
        chevronBox.preferredSize = [buttonWidth, buttonHeight];
        chevronBox.minimumSize = [buttonWidth, buttonHeight];
        chevronBox.maximumSize = [buttonWidth, buttonHeight];
        chevronBox.isPressed = false;
        chevronBox.isStepperButton = true; /* redrawSteppersIn() の目印 / marker for redrawSteppersIn() */

        chevronBox.onDraw = function () {
            var boxGraphics = chevronBox.graphics;
            /* 自作描画は自動でディムにならないため、無効なら薄い色で描く。親の無効化は子の enabled に出ないので親も見る
               Custom drawing is not dimmed automatically; the parent's state does not reach the child's enabled */
            var isDimmed = !isStepperEnabledInTree(chevronBox);

            /* 枠線の内側の地（押下中は押下色） / background inside the frame, pressed color while pressed */
            var fillColor = isDimmed ? STEPPER_DIM_FILL_COLOR : (chevronBox.isPressed ? STEPPER_PRESSED_COLOR : STEPPER_FILL_COLOR);
            boxGraphics.newPath();
            boxGraphics.rectPath(1, isUp ? 1 : 0, buttonWidth - 2, buttonHeight - 1);
            boxGraphics.fillPath(boxGraphics.newBrush(boxGraphics.BrushType.SOLID_COLOR, fillColor));

            drawStepperFrame(boxGraphics, buttonWidth, buttonHeight, isUp, isDimmed ? STEPPER_DIM_FRAME_COLOR : STEPPER_FRAME_COLOR);
            drawStepperChevron(boxGraphics, buttonWidth, buttonHeight, isUp, isDimmed ? STEPPER_DIM_CHEVRON_COLOR : STEPPER_CHEVRON_COLOR);
        };

        /**
         * 押下状態を変えて描き直す
         * @param {boolean} isPressed - 押下中なら true
         * @returns {void}
         */
        function repaint(isPressed) {
            if (chevronBox.isPressed === isPressed) return;
            chevronBox.isPressed = isPressed;
            redrawStepperGroup(chevronBox);
        }
        chevronBox.addEventListener("mousedown", function () {
            if (!isStepperEnabledInTree(chevronBox)) return;
            repaint(true);
            if (onClickFn) onClickFn();
        });
        chevronBox.addEventListener("mouseup", function () { repaint(false); });
        /* 押したまま外へ出たときも押下色を残さない / reset when the pointer leaves while pressed */
        chevronBox.addEventListener("mouseout", function () { repaint(false); });
        return chevronBox;
    }

    /**
     * 外側の辺だけの枠を描く（角は丸める）。継ぎ目側は開けておき、上下2つで1つの枠に見せる。
     * ScriptUI は円弧を描けないため、角丸は短い線分で近似する
     * @param {ScriptUIGraphics} boxGraphics - 描画先
     * @param {number} boxWidth - ボタンの幅
     * @param {number} boxHeight - ボタンの高さ
     * @param {boolean} isUp - 上のボタンなら true（上側に枠を描く）
     * @param {number[]} frameColor - [r, g, b, a]
     * @returns {void}
     */
    function drawStepperFrame(boxGraphics, boxWidth, boxHeight, isUp, frameColor) {
        var frameLeft = 0.5;
        var frameRight = boxWidth - 0.5;
        var outerY = isUp ? 0.5 : boxHeight - 0.5;
        var seamY = isUp ? boxHeight : 0;
        var towardSeam = isUp ? 1 : -1; /* 外側の辺から継ぎ目へ向かう向き / direction from the outer edge to the seam */
        var radius = STEPPER_CORNER_RADIUS;
        var arcSteps = 4; /* 角丸1つを何本の線分で近似するか / segments per corner */
        var angle, k;

        boxGraphics.newPath();
        boxGraphics.moveTo(frameLeft, seamY);
        /* 左の角丸 / left corner */
        for (k = 0; k <= arcSteps; k++) {
            angle = (Math.PI / 2) * k / arcSteps;
            boxGraphics.lineTo(frameLeft + radius - radius * Math.cos(angle), outerY + towardSeam * (radius - radius * Math.sin(angle)));
        }
        /* 右の角丸 / right corner */
        for (k = 0; k <= arcSteps; k++) {
            angle = (Math.PI / 2) * k / arcSteps;
            boxGraphics.lineTo(frameRight - radius + radius * Math.sin(angle), outerY + towardSeam * (radius - radius * Math.cos(angle)));
        }
        boxGraphics.lineTo(frameRight, seamY);
        boxGraphics.strokePath(boxGraphics.newPen(boxGraphics.PenType.SOLID_COLOR, frameColor, 1));
    }

    /**
     * 山形（∧／∨）を描く。文字グリフの▲▼は上下で大きさやベースラインが揃わないため、線で描く
     * @param {ScriptUIGraphics} boxGraphics - 描画先
     * @param {number} boxWidth - ボタンの幅
     * @param {number} boxHeight - ボタンの高さ
     * @param {boolean} isUp - ∧なら true、∨なら false
     * @param {number[]} chevronColor - [r, g, b, a]
     * @returns {void}
     */
    function drawStepperChevron(boxGraphics, boxWidth, boxHeight, isUp, chevronColor) {
        var centerX = boxWidth / 2;
        var centerY = isUp ? boxHeight / 2 + 0.5 : boxHeight / 2 - 0.5; /* 継ぎ目から少し離す / nudged away from the seam */
        var halfWidth = 3.6; /* 山形の半幅（高さ1.8に対して開き約127°） / half width of the chevron */
        var tipOffsetY = isUp ? -1.8 : 1.8; /* 頂点の中心からのずれ（上向きは上、下向きは下） */
        boxGraphics.newPath();
        boxGraphics.moveTo(centerX - halfWidth, centerY - tipOffsetY);
        boxGraphics.lineTo(centerX, centerY + tipOffsetY);
        boxGraphics.lineTo(centerX + halfWidth, centerY - tipOffsetY);
        boxGraphics.strokePath(boxGraphics.newPen(boxGraphics.PenType.SOLID_COLOR, chevronColor, 1.2));
    }

    /**
     * コントロールと、その親をたどってすべて有効かを返す（親の無効化は子の enabled に出ない）
     * @param {Object} control - 対象のコントロール
     * @returns {boolean} すべて有効なら true
     */
    function isStepperEnabledInTree(control) {
        for (var node = control; node; node = node.parent) {
            if (!node.enabled) return false;
        }
        return true;
    }

    /**
     * コンテナ以下にある∧∨ボタンをすべて描き直す。行やパネルの enabled を切り替えたあとに呼ぶ
     * @param {Object} container - 行・グループ・パネルなど
     * @returns {void}
     */
    function redrawSteppersIn(container) {
        if (!container.children) return;
        for (var i = 0; i < container.children.length; i++) {
            var child = container.children[i];
            if (child.isStepperButton) redrawStepperGroup(child);
            else redrawSteppersIn(child);
        }
    }

    /**
     * group の onDraw を呼び直す。group には notify() が無いため、隠して再表示して描き直させる
     * @param {Group} targetGroup - 描き直す group
     * @returns {void}
     */
    function redrawStepperGroup(targetGroup) {
        targetGroup.hide();
        targetGroup.show();
    }

    // ステップボタン（再利用パーツ）ここまで / End of the reusable stepper

    // =========================================
    // 画面にフィット / Fit to window
    // （_templates/FitViewToItems.jsx から移植 / Ported from _templates/FitViewToItems.jsx）
    // =========================================

    /**
     * 数値を範囲に収める。数値として読めないときは既定値を返す
     * @param {string|number} value - 入力値
     * @param {number[]} range - [下限, 上限]
     * @param {number} fallbackValue - 読めないときの既定値
     * @returns {number} 範囲内の数値
     */
    function clampNumber(value, range, fallbackValue) {
        var numberValue = Number(value);
        if (isNaN(numberValue) || (typeof value === "string" && !/\S/.test(value))) numberValue = fallbackValue;
        return Math.min(range[1], Math.max(range[0], numberValue));
    }

    /**
     * 対象が指定の割合でウィンドウに収まるよう、中心を合わせて表示倍率を変える
     * @param {Document} doc - 対象ドキュメント
     * @param {PageItem[]} targetItems - 対象アイテム
     * @param {number} fillRatio - ウィンドウに対して占める割合（1でいっぱい）
     * @returns {void}
     */
    function fitViewToItems(doc, targetItems, fillRatio) {
        /* 効果を含まない geometricBounds で測る / Measure geometricBounds, without effects */
        var bounds = getClipAwareUnionBounds(targetItems, false);
        if (bounds === null) return;

        var itemWidth = bounds[2] - bounds[0];
        var itemHeight = bounds[1] - bounds[3];
        var activeView = doc.activeView;
        activeView.centerPoint = [(bounds[0] + bounds[2]) / 2, (bounds[1] + bounds[3]) / 2];
        if (itemWidth <= 0 || itemHeight <= 0) return;

        /* 中心をそろえたあとの表示範囲を基準に倍率を求める / Scale from the view bounds after the center has moved */
        var viewBounds = activeView.bounds;
        var scale = Math.min(
            (viewBounds[2] - viewBounds[0]) / itemWidth,
            (viewBounds[1] - viewBounds[3]) / itemHeight
        ) * fillRatio;
        activeView.zoom = clampNumber(activeView.zoom * scale, VIEW_ZOOM_RANGE, 1);
    }

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
            title: { ja: "フォントとウェイトを選択", en: "Choose Font and Weight" }
        },
        checkbox: {
            weightSamples: { ja: "ウェイトの見本を右に並べる", en: "Show Weight Samples on the Right" },
            coverBackground: { ja: "背景を白で覆う", en: "Cover Background in White" },
            showRelatedFamilies: { ja: "類似フォントを表示", en: "Show Related Fonts" },
            showPostScriptName: { ja: "PostScript名で表示", en: "Show PostScript Names" },
            fitView:      { ja: "画面にフィット", en: "Fit to Window" }
        },
        panel: {
            font:   { ja: "フォント", en: "Font" },
            weight: { ja: "ウェイト", en: "Weight" },
            display: { ja: "表示", en: "Display" },
            preview: { ja: "プレビュー", en: "Preview" }
        },
        tooltip: {
            fitView:        { ja: "選択中のテキストと右に並べた見本が収まるよう表示倍率を合わせます。", en: "Zooms so the selected text and the weight samples fit in the window." },
            fitViewPercent: { ja: "ウィンドウに対するプレビューの大きさ（100%でいっぱい）", en: "Size of the previews relative to the window; 100% fills it" },
            stepUp: {
                ja: "値を増やす（shift＋クリックで10の倍数へ、option＋クリックで0.1ずつ）",
                en: "Increase (Shift-click to snap to 10s, Option-click by 0.1)"
            },
            stepDown: {
                ja: "値を減らす（shift＋クリックで10の倍数へ、option＋クリックで0.1ずつ）",
                en: "Decrease (Shift-click to snap to 10s, Option-click by 0.1)"
            },
            stepUpInteger:   { ja: "値を増やす（shift＋クリックで10の倍数へ）", en: "Increase (Shift-click to snap to 10s)" },
            stepDownInteger: { ja: "値を減らす（shift＋クリックで10の倍数へ）", en: "Decrease (Shift-click to snap to 10s)" },
            showRelatedFamilies: { ja: "フォントの一覧に、名前が似ているファミリー（「新ゴ」なら UD新ゴ など）も並べます。オフでは選択中のファミリーだけです。", en: "Also lists families with similar names (UD Shin Go for Shin Go, etc.). When off, only the selected family is listed." },
            showPostScriptName: { ja: "フォントの一覧を、ファミリー名の代わりに PostScript 名（BebasNeuePro など）で表示します。", en: "Shows PostScript names (such as BebasNeuePro) instead of family names in the font list." },
            coverOpacity:   { ja: "白い長方形の不透明度（100%で完全に隠す）", en: "Opacity of the white rectangle; 100% hides everything behind it" },
            coverBackground: { ja: "見本の後ろに半透明の白い長方形を敷き、ほかのオブジェクトを薄くします。選択中のテキストは黒の複製を上に重ねます。閉じると消えます。", en: "Places a translucent white rectangle behind the samples to fade other objects; a black copy of the selected text sits on top. Removed when the dialog closes." },
            weightSamples:  { ja: "選択中のテキストの右に、ウェイトごとの見本を黒で縦に並べます。今のウェイトを横に並べ、細いものを上、太いものを下に置きます。閉じると消えます。", en: "Lists a sample of each weight in black to the right of the selected text: the current weight level with it, lighter ones above and heavier ones below. They are removed when the dialog closes." },
            fontList:      { ja: "選択中のフォントと先頭の語が同じファミリーと、「新ゴ」のような中心の名前を含むファミリーです。", en: "Families that share the first word with the selected font or contain its core name (such as Shin Go)." },
            currentWeight: { ja: "選択中のテキストで使われているウェイトです。", en: "This weight is used in the selected text." }
        },
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok:     { ja: "OK", en: "OK" }
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noText:     { ja: "テキストが選択されていません。", en: "No text is selected." },
            noWeights: {
                ja: "このフォントには、ほかに選べるウェイトがありません。",
                en: "This font has no other weights to choose from."
            }
        }
    };

    // =========================================
    // プレビュー / Preview
    // （対象の文字に直接適用し、キャンセル時は控えた元のフォントに戻す
    //   Applied to the characters directly; the recorded fonts are restored on cancel）
    // =========================================

    /**
     * 文字にフォントを適用して再描画する
     * @param {Object[]} characterEntries - 対象の文字と元フォントの組
     * @param {TextFont} textFont - 適用するフォント
     * @returns {void}
     */
    function applyFontToEntries(characterEntries, textFont) {
        for (var i = 0; i < characterEntries.length; i++) {
            characterEntries[i].character.characterAttributes.textFont = textFont;
        }
        app.redraw();
    }

    /**
     * 文字を控えておいた元のフォントに戻して再描画する
     * @param {Object[]} characterEntries - 対象の文字と元フォントの組
     * @returns {void}
     */
    function restoreOriginalFonts(characterEntries) {
        for (var i = 0; i < characterEntries.length; i++) {
            characterEntries[i].character.characterAttributes.textFont = characterEntries[i].font;
        }
        app.redraw();
    }

    // =========================================
    // ウェイトの見本 / Weight samples
    // （基準のフレームを一時レイヤーに複製し、右に縦一列で並べる。細いウェイトを上・太いウェイトを下
    //   Copies of the base frame in one column on the right: lighter above, heavier below）
    // =========================================

    /**
     * 並べた複製を置く一時レイヤーを消す
     * @param {Document} doc - 対象ドキュメント
     * @returns {void}
     */
    function removePreviewLayer(doc) {
        for (var i = doc.layers.length - 1; i >= 0; i--) {
            if (doc.layers[i].name === PREVIEW_LAYER_NAME) {
                doc.layers[i].locked = false;
                doc.layers[i].remove();
            }
        }
    }

    /**
     * 見本を置く一時レイヤーを返す（無ければ最前面に作る）
     * @param {Document} doc - 対象ドキュメント
     * @returns {Layer} 一時レイヤー
     */
    function getPreviewLayer(doc) {
        for (var i = 0; i < doc.layers.length; i++) {
            if (doc.layers[i].name === PREVIEW_LAYER_NAME) return doc.layers[i];
        }
        var previewLayer = doc.layers.add();
        previewLayer.name = PREVIEW_LAYER_NAME;
        previewLayer.zOrder(ZOrderMethod.BRINGTOFRONT);
        return previewLayer;
    }

    /**
     * 基準のフレームを複製し、ファミリーのウェイトの見本を右に縦一列で並べる
     * 基準のフォントは基準と同じ高さ、細いものは上、太いものは下に置く
     * @param {Document} doc - 対象ドキュメント
     * @param {TextFrame} baseFrame - 基準のテキストフレーム
     * @param {Object[]} seriesMembers - 並べるフォント [{font, weight, variant}]（細い順）
     * @param {TextFont} baseFont - 基準のフォント
     * @param {number} baseWeight - 基準のウェイト
     * @returns {PageItem[]} 並べた見本とウェイト名
     */
    function buildWeightSamples(doc, baseFrame, seriesMembers, baseFont, baseWeight) {
        removePreviewLayer(doc);

        var lighterMembers = [];
        var heavierMembers = [];
        var hasBaseFont = false;
        for (var i = 0; i < seriesMembers.length; i++) {
            if (seriesMembers[i].font.name === baseFont.name) hasBaseFont = true;
            else if (seriesMembers[i].weight < baseWeight) lighterMembers.push(seriesMembers[i]);
            else heavierMembers.push(seriesMembers[i]);
        }
        var sampleItems = [];
        if (seriesMembers.length === 0) return sampleItems;

        var previewLayer = getPreviewLayer(doc);

        var frameBounds = baseFrame.geometricBounds;
        var frameHeight = frameBounds[1] - frameBounds[3];
        var sampleRowStep = frameHeight * (1 + SAMPLE_ROW_GAP_RATIO);
        var columnOffsetX = (frameBounds[2] - frameBounds[0]) + frameHeight * SAMPLE_COLUMN_GAP_RATIO;

        /* 基準のフォントは基準と同じ高さ。無いとき（別ファミリー）は太い側の先頭をその高さに置く
           The base font sits level with the text; without it, the first heavier one takes that row */
        var sampleFonts = [];
        if (hasBaseFont) {
            sampleItems.push(addWeightSample(baseFrame, previewLayer, baseFont, columnOffsetX, 0));
            sampleFonts.push(baseFont);
        }
        var heavierFirstRow = hasBaseFont ? 1 : 0;

        /* 細いものは基準に近い順に上へ、太いものは近い順に下へ / Nearest weights sit closest to the base row */
        for (var j = 0; j < lighterMembers.length; j++) {
            var lighterFont = lighterMembers[lighterMembers.length - 1 - j].font;
            sampleItems.push(addWeightSample(baseFrame, previewLayer, lighterFont, columnOffsetX, sampleRowStep * (j + 1)));
            sampleFonts.push(lighterFont);
        }
        for (var k = 0; k < heavierMembers.length; k++) {
            sampleItems.push(addWeightSample(baseFrame, previewLayer, heavierMembers[k].font, columnOffsetX, -sampleRowStep * (k + heavierFirstRow)));
            sampleFonts.push(heavierMembers[k].font);
        }

        return sampleItems.concat(addWeightLabels(previewLayer, sampleItems, sampleFonts, frameHeight));
    }

    /**
     * 見本の右に、そろえた位置でウェイト名（スタイル名）を置く。縦は各見本の中央に合わせる
     * @param {Layer} previewLayer - 置き先のレイヤー
     * @param {TextFrame[]} sampleFrames - 見本
     * @param {TextFont[]} sampleFonts - 見本ごとのフォント（sampleFrames と同じ順）
     * @param {number} frameHeight - 元のテキストフレームの高さ（pt）
     * @returns {TextFrame[]} 作ったウェイト名
     */
    function addWeightLabels(previewLayer, sampleFrames, sampleFonts, frameHeight) {
        /* 見本は字幅がウェイトごとに違うので、いちばん右の端にそろえる / Align to the rightmost sample edge */
        var labelLeft = null;
        for (var i = 0; i < sampleFrames.length; i++) {
            var sampleRight = sampleFrames[i].geometricBounds[2];
            if (labelLeft === null || sampleRight > labelLeft) labelLeft = sampleRight;
        }
        labelLeft += frameHeight * LABEL_GAP_RATIO;

        var labelSize = Math.max(LABEL_MIN_SIZE, frameHeight * LABEL_SIZE_RATIO);
        var weightLabels = [];
        for (var j = 0; j < sampleFrames.length; j++) {
            var weightLabel = previewLayer.textFrames.add();
            weightLabel.contents = sampleFonts[j].style;
            weightLabel.textRange.characterAttributes.size = labelSize;
            paintTextBlack(weightLabel);

            var sampleBounds = sampleFrames[j].geometricBounds;
            var labelBounds = weightLabel.geometricBounds;
            var sampleCenterY = (sampleBounds[1] + sampleBounds[3]) / 2;
            var labelCenterY = (labelBounds[1] + labelBounds[3]) / 2;
            weightLabel.translate(labelLeft - labelBounds[0], sampleCenterY - labelCenterY);
            weightLabels.push(weightLabel);
        }
        return weightLabels;
    }

    /**
     * 基準のフレームを一時レイヤーに複製し、フォントを変えて移動する
     * @param {TextFrame} baseFrame - 基準のテキストフレーム
     * @param {Layer} previewLayer - 置き先のレイヤー
     * @param {TextFont} textFont - 適用するフォント
     * @param {number} offsetX - 横の移動量（pt、右が正）
     * @param {number} offsetY - 縦の移動量（pt、上が正）
     * @returns {TextFrame} 作った複製
     */
    function addWeightSample(baseFrame, previewLayer, textFont, offsetX, offsetY) {
        var sampleFrame = baseFrame.duplicate(previewLayer, ElementPlacement.PLACEATEND);
        sampleFrame.hidden = false;
        sampleFrame.locked = false;
        sampleFrame.textRange.characterAttributes.textFont = textFont;
        paintTextBlack(sampleFrame);
        sampleFrame.translate(offsetX, offsetY);
        return sampleFrame;
    }

    /**
     * 元のテキストを一時レイヤーの同じ位置に黒で複製する（白い長方形の上でも見えるように）
     * @param {TextFrame} baseFrame - 元のテキストフレーム
     * @param {Document} doc - 対象ドキュメント
     * @returns {TextFrame} 作った複製
     */
    function addBaseTextCopy(baseFrame, doc) {
        var baseTextCopy = baseFrame.duplicate(getPreviewLayer(doc), ElementPlacement.PLACEATBEGINNING);
        baseTextCopy.hidden = false;
        baseTextCopy.locked = false;
        paintTextBlack(baseTextCopy);
        return baseTextCopy;
    }

    /**
     * テキストの塗りを黒、線をなしにする
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @returns {void}
     */
    function paintTextBlack(textFrame) {
        var characterAttributes = textFrame.textRange.characterAttributes;
        characterAttributes.fillColor = createNeutralColor(app.activeDocument, 100);
        characterAttributes.strokeColor = new NoColor();
    }

    // =========================================
    // 背景を白で覆う / Cover the background in white
    // =========================================

    /**
     * 一時レイヤーのいちばん奥に、表示範囲より広い半透明の白い長方形を置く
     * @param {Document} doc - 対象ドキュメント
     * @param {number} coverOpacity - 不透明度（%）
     * @returns {PathItem} 作った長方形
     */
    function addBackgroundCover(doc, coverOpacity) {
        var viewBounds = doc.activeView.bounds;
        var viewWidth = viewBounds[2] - viewBounds[0];
        var viewHeight = viewBounds[1] - viewBounds[3];
        var marginX = viewWidth * (COVER_VIEW_SCALE - 1) / 2;
        var marginY = viewHeight * (COVER_VIEW_SCALE - 1) / 2;

        var coverRect = getPreviewLayer(doc).pathItems.rectangle(
            viewBounds[1] + marginY, viewBounds[0] - marginX, viewWidth * COVER_VIEW_SCALE, viewHeight * COVER_VIEW_SCALE
        );
        coverRect.stroked = false;
        coverRect.filled = true;
        coverRect.fillColor = createNeutralColor(doc, 0);
        coverRect.opacity = coverOpacity;
        coverRect.zOrder(ZOrderMethod.SENDTOBACK);
        return coverRect;
    }

    /**
     * ドキュメントのカラーモードに合わせた無彩色を返す（0で白、100で黒。CMYKはK版だけ）
     * @param {Document} doc - 対象ドキュメント
     * @param {number} blackPercent - 黒の濃さ（%）
     * @returns {CMYKColor|RGBColor} 色
     */
    function createNeutralColor(doc, blackPercent) {
        if (doc.documentColorSpace === DocumentColorSpace.CMYK) {
            var cmykColor = new CMYKColor();
            cmykColor.cyan = 0;
            cmykColor.magenta = 0;
            cmykColor.yellow = 0;
            cmykColor.black = blackPercent;
            return cmykColor;
        }
        var rgbLevel = Math.round(255 * (100 - blackPercent) / 100);
        var rgbColor = new RGBColor();
        rgbColor.red = rgbLevel;
        rgbColor.green = rgbLevel;
        rgbColor.blue = rgbLevel;
        return rgbColor;
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * ファミリーごとに右に並べるウェイトを決め、ラジオの数と幅を見積もる
     * @param {Object} relatedFamilies - collectRelatedFamilies() の結果
     * @param {number} baseVariant - 基準の系列
     * @returns {{seriesByFamily: Object, maxSeriesCount: number, radioWidth: number}} ファミリー名 → ウェイト、ラジオの数、ラジオの幅
     */
    function prepareSeriesLayout(relatedFamilies, baseVariant) {
        var seriesByFamily = {};
        var maxSeriesCount = 0;
        /* ラジオは作成後に広がらないので、最長のスタイル名に合わせた幅を見積もる / Radios do not grow after creation */
        var radioWidth = WEIGHT_RADIO_WIDTH;
        for (var i = 0; i < relatedFamilies.familyNames.length; i++) {
            var familyName = relatedFamilies.familyNames[i];
            var seriesMembers = pickSeriesMembers(relatedFamilies.membersByFamily[familyName], baseVariant);
            seriesByFamily[familyName] = seriesMembers;
            if (seriesMembers.length > maxSeriesCount) maxSeriesCount = seriesMembers.length;
            for (var j = 0; j < seriesMembers.length; j++) {
                radioWidth = Math.max(radioWidth, seriesMembers[j].font.style.length * RADIO_CHAR_WIDTH + RADIO_INDICATOR_WIDTH);
            }
        }
        return { seriesByFamily: seriesByFamily, maxSeriesCount: maxSeriesCount, radioWidth: radioWidth };
    }

    /**
     * 「□チェック ∧∨［数値］%」の行を追加する（∧∨・↑↓キーで増減し、直接入力は範囲と整数にそろえる）
     * @param {Group|Panel} parent - 追加先
     * @param {Object} rowOptions - label / tooltip / checked / percent / range / fieldTooltip / onUpdate（値やチェックが変わったとき）
     * @returns {{checkbox: Checkbox, percentInput: EditText}} 作ったチェックボックスと入力欄
     */
    function addPercentOptionRow(parent, rowOptions) {
        var percentRow = parent.add("group");
        percentRow.orientation = "row";
        percentRow.alignment = ["left", "center"];
        percentRow.alignChildren = ["left", "center"];
        percentRow.spacing = OPTION_ROW_SPACING;

        var optionCheckbox = percentRow.add("checkbox", undefined, rowOptions.label);
        optionCheckbox.helpTip = rowOptions.tooltip;
        optionCheckbox.value = rowOptions.checked;

        var stepOptions = {
            step: 1, min: rowOptions.range[0], max: rowOptions.range[1], integer: true,
            onStep: function () { rowOptions.onUpdate(); }
        };
        /* ∧∨と入力欄は隙間0で突き合わせる / Butt the stepper against the field */
        var percentGroup = percentRow.add("group");
        percentGroup.orientation = "row";
        percentGroup.alignChildren = ["left", "center"];
        percentGroup.spacing = 0;
        percentGroup.margins = 0;
        var percentInput;
        var percentStepper = addStepper(percentGroup, function () { return percentInput; }, stepOptions);
        percentInput = percentGroup.add("edittext", undefined, String(rowOptions.percent));
        percentInput.characters = 3;
        percentInput.helpTip = rowOptions.fieldTooltip;
        percentInput.stepperGroup = percentStepper;
        percentInput.fieldLabel = percentRow.add("statictext", undefined, "%");
        bindSteppedArrowKeys(percentInput, percentStepper);

        /* 直接入力は範囲と整数にそろえ、数値でなければ既定値に戻す / Normalize typed values */
        percentInput.onChange = function () {
            writeSteppedValue(percentInput, clampNumber(percentInput.text, rowOptions.range, rowOptions.percent), stepOptions);
            rowOptions.onUpdate();
        };

        /* 割合の欄はチェックがONのときだけ使える / The field works only while checked */
        var updatePercentEnabled = function () {
            setSteppedFieldEnabled(percentInput, optionCheckbox.value && optionCheckbox.enabled);
        };
        optionCheckbox.onClick = function () {
            updatePercentEnabled();
            rowOptions.onUpdate();
        };
        percentInput.updateEnabled = updatePercentEnabled;
        updatePercentEnabled();

        return { checkbox: optionCheckbox, percentInput: percentInput };
    }

    /**
     * 左に関連ファミリー、右にウェイトを表示し、選んだフォントをプレビューする（使用中のウェイトは無効表示）
     * OK で適用したまま閉じ、キャンセルで元のフォントに戻す
     * @param {Object} relatedFamilies - collectRelatedFamilies() の結果
     * @param {TextFont} baseFont - 基準のフォント
     * @param {Object} baseWeightInfo - 基準のフォントの getWeightInfo()
     * @param {Object[]} characterEntries - プレビューを反映する文字と元フォントの組
     * @param {Document} doc - 対象ドキュメント
     * @param {TextFrame|null} baseFrame - 見本を並べる基準のテキストフレーム
     * @returns {boolean} OK で閉じたら true
     */
    function showWeightDialog(relatedFamilies, baseFont, baseWeightInfo, characterEntries, doc, baseFrame) {
        var usedFontNames = collectUsedFontNames(characterEntries);
        var familyNames = relatedFamilies.familyNames;
        var seriesLayout = prepareSeriesLayout(relatedFamilies, baseWeightInfo.variant);
        var seriesByFamily = seriesLayout.seriesByFamily;

        var weightDialog = new Window("dialog", getLabel(LABELS.dialog.title));
        weightDialog.orientation = "column";
        weightDialog.alignChildren = ["fill", "top"];
        weightDialog.margins = DIALOG_MARGINS;

        var listColumnsGroup = weightDialog.add("group");
        listColumnsGroup.orientation = "row";
        listColumnsGroup.alignChildren = ["fill", "fill"];

        /* 左：関連ファミリー / Left: related families */
        var fontPanel = listColumnsGroup.add("panel", undefined, getLabel(LABELS.panel.font));
        fontPanel.margins = FONT_PANEL_MARGINS;
        var fontList = fontPanel.add("listbox", undefined, []);
        fontList.preferredSize = FONT_LIST_SIZE;
        fontList.helpTip = getLabel(LABELS.tooltip.fontList);

        /* 右：ウェイト（ラジオは同じ親に入れて排他にする）/ Right: weights, radios kept in one parent */
        var weightPanel = listColumnsGroup.add("panel", undefined, getLabel(LABELS.panel.weight));
        weightPanel.orientation = "column";
        weightPanel.alignChildren = ["left", "top"];
        weightPanel.margins = PANEL_MARGINS;
        weightPanel.spacing = PANEL_SPACING;

        /* ラジオは件数の最も多いファミリーに合わせて作り、ファミリーを替えるたびに文字と表示を入れ替える
           As many radios as the longest list; their labels and visibility change with the family */
        var weightRadios = [];
        for (var i = 0; i < seriesLayout.maxSeriesCount; i++) {
            var weightRadio = weightPanel.add("radiobutton", undefined, "");
            weightRadio.preferredSize.width = seriesLayout.radioWidth;
            weightRadio.onClick = createPreviewHandler(i);
            weightRadios.push(weightRadio);
        }

        var shownMembers = [];
        var sampleItems = [];
        var initialView = { centerPoint: doc.activeView.centerPoint, zoom: doc.activeView.zoom };
        var backgroundCover = null;
        var baseTextCopy = null;
        var shownFamilyName = null;
        var listedFamilyNames = [];

        /**
         * ラジオを選んだときにプレビューを反映するハンドラーを作る
         * @param {number} radioIndex - ラジオの番号
         * @returns {Function} onClick ハンドラー
         */
        function createPreviewHandler(radioIndex) {
            return function () {
                previewFont(shownMembers[radioIndex].font);
            };
        }

        /**
         * 元のテキストにフォントを反映し、一時レイヤーの黒い複製も作り直す
         * @param {TextFont} textFont - 反映するフォント
         * @returns {void}
         */
        function previewFont(textFont) {
            applyFontToEntries(characterEntries, textFont);
            refreshBaseTextCopy();
        }

        /**
         * 元のテキストの黒い複製を作り直す（見本か白い長方形があるときだけ）
         * @returns {void}
         */
        function refreshBaseTextCopy() {
            if (baseTextCopy) {
                baseTextCopy.remove();
                baseTextCopy = null;
            }
            if (baseFrame && (weightSamplesCheckbox.value || coverBackgroundCheckbox.value)) {
                baseTextCopy = addBaseTextCopy(baseFrame, doc);
            }
            app.redraw();
        }

        /**
         * 右のラジオを指定ファミリーのウェイトに入れ替え、初期選択をプレビューする
         * @param {string} familyName - ファミリー名
         * @returns {void}
         */
        function showFamilyWeights(familyName) {
            shownFamilyName = familyName;
            shownMembers = seriesByFamily[familyName];
            var initialIndex = findInitialIndex(shownMembers, usedFontNames, baseFont, baseWeightInfo.weight);
            for (var k = 0; k < weightRadios.length; k++) {
                var member = shownMembers[k];
                weightRadios[k].visible = !!member;
                weightRadios[k].text = member ? member.font.style : "";
                weightRadios[k].enabled = !!member && !usedFontNames[member.font.name];
                weightRadios[k].helpTip = (member && usedFontNames[member.font.name]) ? getLabel(LABELS.tooltip.currentWeight) : "";
                weightRadios[k].value = (k === initialIndex);
            }
            refreshWeightSamples();
            if (initialIndex !== -1) previewFont(shownMembers[initialIndex].font);
        }

        /**
         * チェックボックスに応じて、右に並べた見本を作り直すか消す
         * @returns {void}
         */
        function refreshWeightSamples() {
            /* 見本を作り直すとレイヤーごと消えるので、白い長方形と元のテキストの複製もあとで作り直す / They go with the layer */
            backgroundCover = null;
            baseTextCopy = null;
            if (weightSamplesCheckbox.value && baseFrame) {
                sampleItems = buildWeightSamples(doc, baseFrame, shownMembers, baseFont, baseWeightInfo.weight);
            } else {
                removePreviewLayer(doc);
                sampleItems = [];
            }
            applyFitView();
        }

        /**
         * ［画面にフィット］がONのとき、選択中のテキストと並べたプレビューすべてが収まるよう表示倍率を合わせる
         * @returns {void}
         */
        function applyFitView() {
            if (fitViewCheckbox.value && baseFrame) {
                fitViewToItems(doc, [baseFrame].concat(sampleItems), clampNumber(fitViewPercentInput.text, FIT_VIEW_RANGE, FIT_VIEW_PERCENT) / 100);
            }
            /* 表示範囲が変わるので白い長方形も合わせ直す / The view may have moved, so refit the cover */
            refreshBackgroundCover();
        }

        /**
         * ［背景を白で覆う］に合わせて、白い長方形を表示範囲に作り直すか消す
         * @returns {void}
         */
        function refreshBackgroundCover() {
            if (backgroundCover) {
                backgroundCover.remove();
                backgroundCover = null;
            }
            if (coverBackgroundCheckbox.value) backgroundCover = addBackgroundCover(doc, getCoverOpacity());
            refreshBaseTextCopy();
        }

        /**
         * フォントの一覧に出すファミリーの表示名を返す（［PostScript名で表示］がONなら PostScript 名のファミリー部分）
         * @param {string} familyName - ファミリー名
         * @returns {string} 表示名（「BebasNeuePro-Bold」なら「BebasNeuePro」）
         */
        function getFamilyLabel(familyName) {
            if (!showPostScriptNameCheckbox.value) return familyName;
            /* スタイル部分は最後のハイフン以降 / The style follows the last hyphen */
            var postScriptName = seriesByFamily[familyName][0].font.name;
            return (postScriptName.indexOf("-") !== -1) ? postScriptName.replace(/-[^-]*$/, "") : postScriptName;
        }

        fontList.onChange = function () {
            if (!fontList.selection) return;
            var familyName = listedFamilyNames[fontList.selection.index];
            /* 表示名の切り替えで同じファミリーが選び直されたときは、選択とプレビューをそのままにする / Keep the state when only labels changed */
            if (familyName !== shownFamilyName) showFamilyWeights(familyName);
        };

        /**
         * ［類似フォントを表示］［PostScript名で表示］に合わせてフォントの一覧を作り直し、表示中のファミリー（無ければ基準）を選ぶ
         * @returns {void}
         */
        function fillFontList() {
            var selectedFamily = shownFamilyName || baseFont.family;
            listedFamilyNames = showRelatedFamiliesCheckbox.value ? familyNames : [baseFont.family];

            fontList.removeAll();
            var selectedIndex = 0;
            for (var i = 0; i < listedFamilyNames.length; i++) {
                fontList.add("item", getFamilyLabel(listedFamilyNames[i]));
                if (listedFamilyNames[i] === selectedFamily) selectedIndex = i;
            }
            fontList.selection = selectedIndex;
            /* 選択の代入で onChange が呼ばれなかったときの保険 / In case assigning the selection did not fire onChange */
            if (shownFamilyName !== listedFamilyNames[selectedIndex]) showFamilyWeights(listedFamilyNames[selectedIndex]);
        }

        /* オプション：左は一覧の表示、右はプレビュー / Options: list display on the left, preview on the right */
        var optionsGroup = weightDialog.add("group");
        optionsGroup.orientation = "row";
        optionsGroup.alignChildren = ["fill", "fill"];

        var displayPanel = optionsGroup.add("panel", undefined, getLabel(LABELS.panel.display));
        displayPanel.orientation = "column";
        displayPanel.alignChildren = ["left", "top"];
        displayPanel.margins = PANEL_MARGINS;
        displayPanel.spacing = OPTION_ROW_SPACING;

        var showRelatedFamiliesCheckbox = displayPanel.add("checkbox", undefined, getLabel(LABELS.checkbox.showRelatedFamilies));
        showRelatedFamiliesCheckbox.helpTip = getLabel(LABELS.tooltip.showRelatedFamilies);
        showRelatedFamiliesCheckbox.value = SHOW_RELATED_FAMILIES_DEFAULT;
        showRelatedFamiliesCheckbox.onClick = fillFontList;

        var showPostScriptNameCheckbox = displayPanel.add("checkbox", undefined, getLabel(LABELS.checkbox.showPostScriptName));
        showPostScriptNameCheckbox.helpTip = getLabel(LABELS.tooltip.showPostScriptName);
        showPostScriptNameCheckbox.value = SHOW_POSTSCRIPT_NAME_DEFAULT;
        showPostScriptNameCheckbox.onClick = fillFontList;

        var previewPanel = optionsGroup.add("panel", undefined, getLabel(LABELS.panel.preview));
        previewPanel.orientation = "column";
        previewPanel.alignChildren = ["left", "top"];
        previewPanel.margins = PANEL_MARGINS;
        previewPanel.spacing = OPTION_ROW_SPACING;

        var weightSamplesCheckbox = previewPanel.add("checkbox", undefined, getLabel(LABELS.checkbox.weightSamples));
        weightSamplesCheckbox.helpTip = getLabel(LABELS.tooltip.weightSamples);
        weightSamplesCheckbox.value = WEIGHT_SAMPLES_DEFAULT;
        weightSamplesCheckbox.enabled = !!baseFrame;
        weightSamplesCheckbox.onClick = refreshWeightSamples;

        /* ［背景を白で覆う］と不透明度 / Cover the background with its opacity field */
        var coverBackgroundRow = addPercentOptionRow(previewPanel, {
            label: getLabel(LABELS.checkbox.coverBackground), tooltip: getLabel(LABELS.tooltip.coverBackground),
            checked: COVER_BACKGROUND_DEFAULT, percent: COVER_BACKGROUND_OPACITY, range: COVER_OPACITY_RANGE,
            fieldTooltip: getLabel(LABELS.tooltip.coverOpacity),
            onUpdate: function () { updateBackgroundCover(); }
        });
        var coverBackgroundCheckbox = coverBackgroundRow.checkbox;
        var coverOpacityInput = coverBackgroundRow.percentInput;

        /**
         * 不透明度の欄の値を返す
         * @returns {number} 不透明度（%）
         */
        function getCoverOpacity() {
            return clampNumber(coverOpacityInput.text, COVER_OPACITY_RANGE, COVER_BACKGROUND_OPACITY);
        }

        /**
         * 白い長方形の有無・不透明度を反映する（ある長方形は不透明度だけ書き換える）
         * @returns {void}
         */
        function updateBackgroundCover() {
            if (coverBackgroundCheckbox.value && backgroundCover) {
                backgroundCover.opacity = getCoverOpacity();
                app.redraw();
                return;
            }
            refreshBackgroundCover();
        }

        /* ［画面にフィット］と割合 / Fit to window with its share field */
        var fitViewRow = addPercentOptionRow(previewPanel, {
            label: getLabel(LABELS.checkbox.fitView), tooltip: getLabel(LABELS.tooltip.fitView),
            checked: FIT_VIEW_DEFAULT, percent: FIT_VIEW_PERCENT, range: FIT_VIEW_RANGE,
            fieldTooltip: getLabel(LABELS.tooltip.fitViewPercent),
            onUpdate: function () { applyFitView(); }
        });
        var fitViewCheckbox = fitViewRow.checkbox;
        var fitViewPercentInput = fitViewRow.percentInput;
        if (!baseFrame) {
            fitViewCheckbox.enabled = false;
            fitViewPercentInput.updateEnabled();
        }

        var buttonRow = addButtonRow(weightDialog);
        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });
        var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.ok), { name: "ok" });

        /* 開いた時点で基準のファミリーを選び、初期選択をプレビュー / Select the base family on open */
        weightDialog.onShow = fillFontList;

        alignRightOnlyButtonRow(buttonRow);
        prepareDialogWindow(weightDialog, SCRIPT_NAME);
        var isConfirmed = (weightDialog.show() === 1);
        removePreviewLayer(doc);
        if (!isConfirmed) {
            restoreOriginalFonts(characterEntries);
            /* キャンセル時は表示位置と倍率も開いたときに戻す / Restore the view on cancel */
            doc.activeView.centerPoint = initialView.centerPoint;
            doc.activeView.zoom = initialView.zoom;
        }
        app.redraw();
        return isConfirmed;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    if (app.documents.length === 0) {
        alert(getLabel(LABELS.alert.noDocument));
        return;
    }

    var doc = app.activeDocument;
    var characterEntries = collectCharacterFonts(collectTextRanges(doc.selection));
    if (characterEntries.length === 0) {
        alert(getLabel(LABELS.alert.noText));
        return;
    }

    /* 先頭の文字のフォントを基準にする / Base the list on the first character's font */
    var baseFont = characterEntries[0].font;
    var baseWeightInfo = getWeightInfo(baseFont);
    if (!baseWeightInfo) {
        alert(getLabel(LABELS.alert.noWeights));
        return;
    }

    var relatedFamilies = collectRelatedFamilies(baseFont.family);
    var baseSeries = pickSeriesMembers(relatedFamilies.membersByFamily[baseFont.family], baseWeightInfo.variant);
    if (relatedFamilies.familyNames.length === 1 && findInitialIndex(baseSeries, collectUsedFontNames(characterEntries), baseFont, baseWeightInfo.weight) === -1) {
        alert(getLabel(LABELS.alert.noWeights));
        return;
    }

    showWeightDialog(relatedFamilies, baseFont, baseWeightInfo, characterEntries, doc, findBaseTextFrame(doc.selection));

})();
