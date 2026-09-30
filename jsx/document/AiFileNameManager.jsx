#target illustrator
#targetengine "Ai-FileNameManagerEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

アクティブなドキュメントのファイル名を、ベース・サブテキスト・ステータス・タイムスタンプ・連番・バージョン番号のセグメント単位で組み立て直し、リネーム／別名で保存／コピーを保存のいずれかで付け直します。
保存形式は常に Illustrator 形式（.ai）なので、.ai 以外の書類では「別名で保存」だけが使えます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AiFileNameManager.md

note記事も参照してください。
https://note.com/dtp_tranist/n/nc88dd887eb1c

### Overview

Reassembles the active document's filename from its segments — base, title, status, timestamp, sequence and version — and then renames it, saves it under a new name, or saves a copy.
The output is always Illustrator format (.ai), so only "Save As" is available for a non-.ai document.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AiFileNameManager.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AiFileNameManager";            /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.3.16";                      /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-05-27";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-01";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AiFileNameManager.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AiFileNameManager.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/nc88dd887eb1c"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // ユーザー設定 / User Settings
    // =========================================

    /* 整形機能のオン/オフ。false にすると該当 UI とロジックを丸ごとスキップ
       / Feature switches. Set to false to omit the UI and skip the logic */
    var FEATURE_STATUS = true;        // ステータス dropdown と検出 / Status dropdown + detection
    var FEATURE_SORT = true;          // ソートパネル + 並び順カスタマイズ / Sort panel + custom segment order
    var FEATURE_SEPARATOR = '-';      // 区切り記号統一: '-' / '_' で有効化＋既定値、false で無効 / '-' or '_' enables with that default; false disables
    var FEATURE_DOT_NORMALIZE = true; // "." を区切り記号にあわせて置換（要 FEATURE_SEPARATOR） / "." normalization with the chosen separator
    var FEATURE_NFC = true;           // 濁点・半濁点の NFC 結合 / Combine separated dakuten/handakuten (NFC)
    var FEATURE_CLEAN = true;         // クリーンなファイル名（OS 禁止文字 + 絵文字・機種依存文字の処理） / Clean filename (OS-invalid + emoji + platform-dependent)
    var FEATURE_TRANSLITERATE = true; // 法人略記・丸数字などを ASCII 相当に変換 / Transliterate corp abbrev / circled numbers / etc.
    var FEATURE_HALFWIDTH_KANA = true; // 半角カナ → 全角カナ変換（クリーンが '-' / '_' のときのみ有効） / Halfwidth-kana → fullwidth (active only when Clean is '-' or '_')
    var FEATURE_PAGE = true;          // 連番（pageNN）セグメント / Page-number segment (pageNN)

    /* ファイル名（拡張子込み）の UTF-8 バイト長の上限。超過時は確認ダイアログを出して続行可
       / Max UTF-8 byte length of the basename + extension; exceeding it triggers a confirm dialog */
    var FEATURE_MAX_FILENAME_BYTES = 240;

    /* rename モードで元ファイルを削除するときに、即削除ではなく ~/.Trash に移動
       / In rename mode, move the original to ~/.Trash instead of removing it outright */
    var FEATURE_USE_TRASH = true;

    /* 保存時の「PDF互換ファイルを作成」。書類が元々どう保存されていたかは DOM から読めないため、
       リネーム・別名保存の結果はここの値で決まる（既定は Illustrator の初期値と同じ true）
       / "Create PDF Compatible File" on save. The document's own setting is not readable from the
         DOM, so every save uses this value (true matches Illustrator's own default) */
    var FEATURE_PDF_COMPATIBLE = true;

    /* 文字 → 置換文字列のマップ。法人略記、丸数字（白・黒・括弧）、ローマ数字、略記号、単位、ダッシュ類
       / Char → replacement map: corporate abbrev, circled (white/black/parenthesized), Roman numerals, abbrev symbols, units, dashes */
    var TRANSLITERATE_MAP = {
        // 法人略記
        '㈱': '株', '㈲': '有', '㈹': '代', '㈳': '社',
        '㈵': '特', '㈶': '財', '㈻': '学', '㍿': '株式会社',
        // 白丸数字 ①-⑳ (U+2460-U+2473)
        '①': '1', '②': '2', '③': '3', '④': '4', '⑤': '5',
        '⑥': '6', '⑦': '7', '⑧': '8', '⑨': '9', '⑩': '10',
        '⑪': '11', '⑫': '12', '⑬': '13', '⑭': '14', '⑮': '15',
        '⑯': '16', '⑰': '17', '⑱': '18', '⑲': '19', '⑳': '20',
        // 黒丸数字 ❶-❿ (U+2776-U+277F) + ⓫-⓴ (U+24EB-U+24F4)
        '❶': '1', '❷': '2', '❸': '3', '❹': '4', '❺': '5',
        '❻': '6', '❼': '7', '❽': '8', '❾': '9', '❿': '10',
        '⓫': '11', '⓬': '12', '⓭': '13', '⓮': '14', '⓯': '15',
        '⓰': '16', '⓱': '17', '⓲': '18', '⓳': '19', '⓴': '20',
        // 括弧数字 ⑴-⒇ (U+2474-U+2487)
        '⑴': '1', '⑵': '2', '⑶': '3', '⑷': '4', '⑸': '5',
        '⑹': '6', '⑺': '7', '⑻': '8', '⑼': '9', '⑽': '10',
        '⑾': '11', '⑿': '12', '⒀': '13', '⒁': '14', '⒂': '15',
        '⒃': '16', '⒄': '17', '⒅': '18', '⒆': '19', '⒇': '20',
        // ローマ数字 大文字 Ⅰ-Ⅻ (U+2160-U+216B)
        'Ⅰ': '1', 'Ⅱ': '2', 'Ⅲ': '3', 'Ⅳ': '4', 'Ⅴ': '5',
        'Ⅵ': '6', 'Ⅶ': '7', 'Ⅷ': '8', 'Ⅸ': '9', 'Ⅹ': '10',
        'Ⅺ': '11', 'Ⅻ': '12',
        // ローマ数字 小文字 ⅰ-ⅻ (U+2170-U+217B)
        'ⅰ': '1', 'ⅱ': '2', 'ⅲ': '3', 'ⅳ': '4', 'ⅴ': '5',
        'ⅵ': '6', 'ⅶ': '7', 'ⅷ': '8', 'ⅸ': '9', 'ⅹ': '10',
        'ⅺ': '11', 'ⅻ': '12',
        // 略記号
        '℡': 'TEL', '№': 'No',
        // 単位記号
        '㎜': 'mm', '㎝': 'cm', '㎞': 'km',
        '㎎': 'mg', '㎏': 'kg',
        '㎡': 'm2', '㎥': 'm3',
        // チルダ類
        '〜': '~', '～': '~',  // 〜 WAVE DASH / ～ FULLWIDTH TILDE
        // ハイフン・ダッシュ類
        '－': '-',  // － FULLWIDTH HYPHEN-MINUS
        '‐': '-',  // ‐ HYPHEN
        '‑': '-',  // ‑ NON-BREAKING HYPHEN
        '‒': '-',  // ‒ FIGURE DASH
        '–': '-',  // – EN DASH
        '—': '-',  // — EM DASH
        '―': '-',  // ― HORIZONTAL BAR
        '−': '-'   // − MINUS SIGN
    };

    /* 出力時のセグメント順序。base / title / status / timestamp / page / version。
       FEATURE_STATUS=false なら status を、FEATURE_PAGE=false なら page を除外
       / Output segment order; "status" / "page" are dropped when their FEATURE_* flags are false */
    var SEGMENT_ORDER = (function () {
        var order = ['base', 'title'];
        if (FEATURE_STATUS) order.push('status');
        order.push('timestamp');
        if (FEATURE_PAGE) order.push('page');
        order.push('version');
        return order;
    })();

    /* ステータス選択肢。value がファイル名に入り、ja / en が UI 表示。先頭は「なし」。
       ja が '---' の項目は dropdown 上の区切り線として表示（選択不可）
       / Status choices: `value` is what enters the filename; `ja` / `en` are UI labels.
       Items with `ja === '---'` render as an unselectable divider. */
    var STATUS_ITEMS = [
        { value: '', ja: 'なし', en: 'None' },
        { value: 'wip', ja: 'wip：作業中・仕掛かり中', en: 'wip: Work in progress' },
        { value: 'draft', ja: 'draft：下書き・ラフ・素案', en: 'draft: Draft / Rough' },
        { value: 'review', ja: 'review：レビュー待ち', en: 'review: Awaiting review' },
        { value: 'revised', ja: 'revised：改訂版・修正反映版', en: 'revised: Revised' },
        { value: 'updated', ja: 'updated：更新版', en: 'updated: Updated' },
        { value: 'fixed', ja: 'fixed：修正完了', en: 'fixed: Fix completed' },
        { value: 'approved', ja: 'approved：承認済み・確定', en: 'approved: Approved / Final' },
        { value: 'rejected', ja: 'rejected：ボツ・不採用案', en: 'rejected: Rejected' },
        { value: 'archived', ja: 'archived：保管用', en: 'archived: Archive' },
        { value: '', ja: '---', en: '---' },
        { value: 'flattened', ja: 'flattened：レイヤー結合済み（.psd）', en: 'flattened: Flattened (.psd)' },
        { value: 'outlined', ja: 'outlined：アウトライン済み（.ai）', en: 'outlined: Outlined (.ai)' }
    ];

    /**
     * 区切り線エントリかどうか / Whether a STATUS_ITEMS entry is a divider
     * @param {object} item STATUS_ITEMS の要素
     * @returns {boolean} 区切り線なら true
     */
    function isStatusDivider(item) {
        return item && item.ja === '---';
    }

    var NEW_NAME_FIELD_WIDTH = 250;

    // ==============================
    // レイアウト / Layout
    // ==============================

    // UIレイアウト（再利用パーツ） / UI layout (reusable)

    /* ウィンドウ・パネルの余白と間隔 / Window & panel margins and spacing */
    var WINDOW_MARGINS = 16;                 /* ウィンドウ外周の余白 / window margin */
    var WINDOW_SPACING = 12;                 /* ウィンドウ内の要素間隔 / window spacing */
    var PANEL_MARGINS  = [16, 20, 16, 12];   /* パネル余白 [左,上,右,下] / panel margins */
    var PANEL_SPACING  = 12;                 /* パネル内の要素間隔 / panel spacing */
    var COLUMN_SPACING = 12;                 /* 2カラムの間隔 / gap between columns */
    var TAB_MARGINS    = [15, 20, 5, 10];    /* タブ余白 [左,上,右,下] / tab margins */

    /**
     * ウィンドウの共通設定
     * @param {Window} targetWindow - 対象のウィンドウ
     * @param {number} [spacing] - 要素間隔（省略時は WINDOW_SPACING）
     * @returns {void}
     */
    function setupWindow(targetWindow, spacing) {
        targetWindow.orientation = "column";
        targetWindow.alignChildren = "fill";
        targetWindow.margins = WINDOW_MARGINS;
        targetWindow.spacing = (typeof spacing === "number") ? spacing : WINDOW_SPACING;
    }

    /**
     * パネルの共通設定（子は幅いっぱい。ボタンは alignment = "left" で広げない）
     * @param {Panel} targetPanel - 対象のパネル
     * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function setupPanel(targetPanel, spacing) {
        targetPanel.orientation = "column";
        targetPanel.alignChildren = ["fill", "top"];
        targetPanel.alignment = "fill";
        targetPanel.margins = PANEL_MARGINS;
        targetPanel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /**
     * タブの共通設定
     * @param {Tab} targetTab - 対象のタブ
     * @param {number} [spacing] - 要素間隔（省略時は変えない）
     * @returns {void}
     */
    function setupTab(targetTab, spacing) {
        targetTab.orientation = "column";
        targetTab.alignChildren = "fill";
        targetTab.margins = TAB_MARGINS;
        if (typeof spacing === "number") targetTab.spacing = spacing;
    }

    /**
     * 横並びの行グループの共通設定（ボタン列など）。
     * alignment と alignChildren を対で指定し、中のボタンが横に伸びたり天地がずれたりしないようにする
     * @param {Group} rowGroup - 対象のグループ
     * @param {string|string[]} [rowAlignment] - 横方向の alignment（省略時は "left"）。配列ならそのまま使う
     * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function setupRow(rowGroup, rowAlignment, spacing) {
        rowGroup.orientation = "row";
        rowGroup.alignment = (rowAlignment instanceof Array) ? rowAlignment : [rowAlignment || "left", "center"];
        rowGroup.alignChildren = ["left", "center"];
        rowGroup.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /**
     * ボタンの高さを指定した px だけ詰める（レイアウトが決まったあとに呼ぶ）
     * @param {Button} targetButton - 対象のボタン
     * @param {number} trimPixels - 詰める量（px）
     * @returns {void}
     */
    function trimButtonHeight(targetButton, trimPixels) {
        /* レイアウト前は size が無い / size is not set until the layout runs */
        if (!targetButton.size) return;
        targetButton.size = [targetButton.size.width, targetButton.size.height - trimPixels];
    }

    // UIレイアウト（再利用パーツ）ここまで / End of the reusable UI layout

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

    /* 日英ラベル定義 / Japanese-English label definitions */

    var LABELS = {
        dialog: {
            title: { ja: "ファイル名を変更して保存", en: "Rename and Save" }
        },
        panel: {
            mode: { ja: "動作", en: "Mode" },
            opMode: { ja: "モード", en: "Scope" },
            filename: { ja: "ファイル名プレビュー", en: "File Name Preview" },
            options: { ja: "ファイル名の設定", en: "Filename Settings" },
            sort: { ja: "構成要素の順序", en: "Segment Order" }
        },
        radio: {
            rename: { ja: "元ファイルをリネーム", en: "Rename Original" },
            saveAs: { ja: "別名で保存", en: "Save As" },
            saveCopy: { ja: "コピー（複製）を保存", en: "Save a Copy" },
            opVersionOnly: { ja: "バージョンのみ", en: "Version Only" },
            opFull: { ja: "全体", en: "Full" },
            noChange: { ja: "変更しない", en: "No Change" },
            titleNone: { ja: "なし", en: "None" },
            titleParent: { ja: "親フォルダー", en: "Parent Folder" },
            titleGrandparent: { ja: "2 階層上のフォルダー", en: "Grandparent Folder" },
            titleCustom: { ja: "指定", en: "Custom" },
            timestampNone: { ja: "なし", en: "None" },
            timestampDate: { ja: "YYYYMMDD", en: "YYYYMMDD" },
            timestampDateDash: { ja: "YYYY-MM-DD", en: "YYYY-MM-DD" },
            timestampWithTime: { ja: "時刻も付与", en: "Append HHMM" },
            versionNone: { ja: "なし", en: "None" },
            versionShort: { ja: "v1, v2…", en: "v1, v2…" },
            versionPadded: { ja: "v01, v02…", en: "v01, v02…" },
            versionPaddedWide: { ja: "v001, v002…", en: "v001, v002…" },
            changeToDash: { ja: "-に変更", en: "Change to -" },
            changeToUnderscore: { ja: "_に変更", en: "Change to _" },
            nfcCombine: { ja: "結合する", en: "Combine" },
            cleanRemove: { ja: "削除する", en: "Remove" },
            translitRemove: { ja: "削除する", en: "Remove" },
            translitConvert: { ja: "変換する", en: "Convert" },
            halfwidthKanaConvert: { ja: "半角カナ → 全角", en: "Half-width → Full-width Kana" },
            pageEnable: { ja: "連番を付与", en: "Append Sequence" },
            pagePad2: { ja: "01, 02…", en: "01, 02…" },
            pagePad3: { ja: "001, 002…", en: "001, 002…" },
            sortOff: { ja: "標準順", en: "Default" },
            sortCurrent: { ja: "現在のファイル名に準じる", en: "Match Current" },
            sortOn: { ja: "カスタム順", en: "Custom" }
        },
        label: {
            currentName: { ja: "現在", en: "Current" },
            finalName: { ja: "保存後の名前", en: "Saved Name" },
            base: { ja: "ベース", en: "Base" },
            title: { ja: "サブテキスト", en: "Project Name" },
            status: { ja: "ステータス", en: "Status" },
            timestamp: { ja: "タイムスタンプ", en: "Timestamp" },
            version: { ja: "バージョン番号", en: "Version" },
            page: { ja: "連番", en: "Sequence" },
            separator: { ja: "区切り記号", en: "Separator" },
            nfc: { ja: "濁点・半濁点の正規化", en: "NFC Normalization" },
            clean: { ja: "クリーンなファイル名", en: "Clean Filename" },
            translit: { ja: "丸数字や法人略記など", en: "Symbols" }
        },
        tip: {
            rename: {
                ja: "新しい名前で保存したあと、元ファイルをゴミ箱に移します（実質的にリネーム）。元ファイルを参照している他のドキュメント（配置 .ai／InDesign のリンクなど）はリンク切れになります。",
                en: "Saves with the new name, then moves the original to the Trash (effectively a rename). Documents that reference the original file (placed .ai or InDesign links) will lose the link."
            },
            nonNativeUnsupported: {
                ja: "保存されるのは Illustrator 形式（.ai）だけなので、.ai 以外の書類ではリネーム・コピーを使えません。「別名で保存」を使ってください。",
                en: "Only Illustrator (.ai) is written, so Rename and Save a Copy are unavailable for a non-.ai document. Use \"Save As\" instead."
            },
            saveAs: {
                ja: "新しい名前で保存します。元ファイルは残り、作業中のドキュメントが新ファイルに切り替わります。",
                en: "Saves with the new name. The original file is kept, and the active document switches to the new file."
            },
            saveCopy: {
                ja: "元ファイルに上書き保存したうえで、別名のコピーを作成します。作業中のドキュメントは元ファイルのまま残ります。",
                en: "Saves the original, then creates a copy with the new name. The active document remains the original file."
            },
            base: {
                ja: "ファイル名の先頭部分。空欄にすると省略されます。",
                en: "Leading part of the filename. Leave empty to omit."
            },
            title: {
                ja: "ファイル名に追加する案件名・補足テキスト。「指定」で入力欄の文字列を使用します。",
                en: "Choose how to set the project name in the filename. With \"Custom\", the entered text is used."
            },
            status: {
                ja: "ファイル名に挿入する制作ステータスを選択します。「：」より前の文字列が入ります。",
                en: "Choose a production status to insert. Only the text before \":\" is used in the filename."
            },
            timestamp: {
                ja: "タイムスタンプの形式を選択。「なし」で元の日付があっても削除します。",
                en: "Choose timestamp format. \"None\" removes any existing date."
            },
            timestampWithTime: {
                ja: "タイムスタンプの末尾に時刻（HHMM）を付加します。1 日に複数版を出すときに便利。",
                en: "Append the current time (HHMM) to the timestamp. Useful for multiple versions per day."
            },
            version: {
                ja: "バージョン番号の形式。v1/v2 はパディング無し、v01/v02 は 2 桁、v001/v002 は 3 桁ゼロ埋め。既存の v 番号は +1、無い場合は v1/v01/v001 を付与。「なし」で削除。",
                en: "Version format. v1/v2 has no padding, v01/v02 is 2-digit, v001/v002 is 3-digit zero-padded. An existing v-number is bumped by +1; otherwise v1/v01/v001 is added. \"None\" removes."
            },
            page: {
                ja: "「page01」「page001」のような連番をファイル名に追加します。プレフィックス（page 等）と桁数（01 / 001）を選択。保存時に同フォルダ内の最大連番 +1 へ自動繰り上げ。",
                en: "Append a sequence such as page01 / page001. Configure prefix (e.g. page) and width (01 / 001). On save, bumps to (max in folder) + 1 if collisions exist."
            },
            separator: {
                ja: "ファイル名全体の区切り記号の扱いを選択します。「-」「_」「.」が対象。YYYY-MM-DD のタイムスタンプは保護されます。",
                en: "Choose how separators in the filename are handled. Targets `-`, `_`, and `.`. YYYY-MM-DD timestamps are preserved."
            },
            sort: {
                ja: "並び順を選択。「現在のファイル名に準じる」は検出された要素のみ、「カスタム順」は［編集］で並び替え。",
                en: "Choose order. \"Match Current\" uses only detected segments; \"Custom\" enables [Edit]."
            },
            nfc: {
                ja: "ファイル名に分離した濁点・半濁点が含まれる場合、結合済み文字（NFC）に正規化します。",
                en: "Normalize separated dakuten/handakuten to combined characters (NFC)."
            },
            clean: {
                ja: "OS で使えない文字（\\ / : * ? \" < > |）、絵文字・機種依存文字（㈱・①・㌔ など）、スペースの扱いをまとめて選択します。連続するスペースは 1 つにまとめてから処理。",
                en: "How to handle OS-invalid characters (\\ / : * ? \" < > |), emoji, platform-dependent chars (㈱, ①, ㌔), and spaces. Consecutive spaces collapse to one before replacement."
            },
            translit: {
                ja: "法人略記（㈱→株）、丸数字（①→1）、ローマ数字（Ⅰ→1）、TEL/No、単位記号（㎜→mm）、ダッシュ類を ASCII 相当に変換します。",
                en: "Convert corporate abbreviations (㈱→株), circled numbers (①→1), Roman numerals, TEL/No, units (㎜→mm), and dashes to ASCII equivalents."
            },
            halfwidthKana: {
                ja: "「-」または「_」を選んだとき、半角カタカナを全角カタカナに変換します（濁点・半濁点付きは結合）。",
                en: "When `-` or `_` is selected, convert half-width katakana to full-width (voiced/semi-voiced marks are merged)."
            }
        },
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            sort: { ja: "編集", en: "Edit" }
        },
        sort: {
            title: { ja: "ファイル名の並び順", en: "Filename Order" },
            hint: { ja: "選択中の項目を ↑↓ で並び替え", en: "Select an item and use ↑↓ to reorder" }
        },
        message: {
            noDoc: { ja: "ドキュメントが開かれていません", en: "No document is open." },
            emptyName: { ja: "ファイル名が空です", en: "File name is empty." },
            chooseDestination: { ja: "保存先フォルダを指定", en: "Choose destination folder" },
            confirmOverwrite: {
                ja: "同名ファイルが存在します。上書きしますか？",
                en: "A file with the same name exists. Overwrite?"
            },
            confirmTooLong: {
                ja: "ファイル名が長すぎる可能性があります（{bytes} バイト / 上限 {limit} バイト）。このまま続行しますか？",
                en: "The filename may be too long ({bytes} bytes / limit {limit}). Continue anyway?"
            },
            saveFailed: { ja: "保存に失敗しました", en: "Failed to save" }
        }
    };

    // ボタン行（再利用パーツ） / Button row (reusable)

    var BUTTON_ROW_TOP_MARGIN = 5; /* ボタン行の上の余白 / top margin of the button row */
    var BUTTON_ROW_BOTTOM_MARGIN = 14; /* ボタン行の下の余白。ダイアログの下余白と合わせて約30px（Illustrator 標準のダイアログに合わせる） / bottom margin; with the dialog margin about 30px, like Illustrator's own dialogs */
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
        btnRowGroup.margins = [0, BUTTON_ROW_TOP_MARGIN, 0, BUTTON_ROW_BOTTOM_MARGIN];
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

    // =========================================
    // ヘルパー / Helpers
    // =========================================

    var VERSION_TOKEN_RE = /^[vV]\d+$/;   // v123 / V123
    var VERSION_TOKEN_SOURCE = '[vV]';     // 採番パターンを組み立てるときの v 部分
    var PAGE_TOKEN_RE = /^(page)(\d+)$/i;  // page01 / PAGE003（既定プレフィックス page + 数字）

    /**
     * 実在する暦日か（うるう年・月末を含む）を Date で判定。
     * year/month/day は文字列でも数値でも可。存在しない日付（例: 2026-02-31）は false
     * Whether year/month/day is a real calendar date (leap years / month-ends), validated via Date
     * @param {string} yearStr 年（数値でも可）
     * @param {string} monthStr 月（数値でも可）
     * @param {string} dayStr 日（数値でも可）
     * @returns {boolean} 実在する日付なら true
     */
    function isRealDate(yearStr, monthStr, dayStr) {
        var y = parseInt(yearStr, 10);
        var m = parseInt(monthStr, 10);
        var d = parseInt(dayStr, 10);
        if (isNaN(y) || m < 1 || m > 12 || d < 1 || d > 31) return false;
        var dt = new Date(y, m - 1, d);
        // Date が桁あふれ（2/31 → 3/3 など）を補正するので、元の Y/M/D と一致するかで実在日を判定
        return dt.getFullYear() === y && dt.getMonth() === (m - 1) && dt.getDate() === d;
    }

    /**
     * YYYYMMDD（8 桁）または YYMMDD（6 桁）の日付トークンか。実在する暦日のみ true。
     * 6 桁の年は 2000 年代として扱い、うるう年判定を行う
     * Whether a token is a real YYYYMMDD or YYMMDD date (6-digit year assumed 20xx for leap checks)
     * @param {string} token 判定するトークン
     * @returns {boolean} 日付トークンなら true
     */
    function isDateToken(token) {
        var s = String(token);
        if (/^\d{8}$/.test(s)) return isRealDate(s.substring(0, 4), s.substring(4, 6), s.substring(6, 8));
        if (/^\d{6}$/.test(s)) return isRealDate('20' + s.substring(0, 2), s.substring(2, 4), s.substring(4, 6));
        return false;
    }

    /**
     * "v123" 形式のバージョントークンか / Whether a token is a v+digits version
     * @param {string} token 判定するトークン
     * @returns {boolean} バージョントークンなら true
     */
    function isVersionToken(token) {
        return VERSION_TOKEN_RE.test(String(token));
    }

    /**
     * "page01" / "page003" 形式の連番トークンか（既定プレフィックス page、大文字小文字無視）
     * Whether a token is a page-number token ("page" + digits, case-insensitive)
     * @param {string} token 判定するトークン
     * @returns {boolean} 連番トークンなら true
     */
    function isPageToken(token) {
        return PAGE_TOKEN_RE.test(String(token));
    }

    /**
     * STATUS_ITEMS のいずれかの value と完全一致（大文字小文字無視）すれば true。一致した正規 value を返す。divider は除外
     * Match against STATUS_ITEMS values (case-insensitive); returns the canonical value or '' (skips dividers)
     * @param {string} token 判定するトークン
     * @returns {string} 一致したステータス値（無ければ空文字）
     */
    function matchStatusToken(token) {
        var lowerToken = String(token).toLowerCase();
        for (var i = 1; i < STATUS_ITEMS.length; i++) {
            if (isStatusDivider(STATUS_ITEMS[i])) continue;
            if (STATUS_ITEMS[i].value === lowerToken) return STATUS_ITEMS[i].value;
        }
        return '';
    }

    /**
     * ファイル名を順序付きセグメント配列に分解。各 segment は前にあった区切り文字も保持
     * Decompose into ordered segments; each segment stores the preceding separator
     * @param {string} name 拡張子を除いたファイル名
     * @returns {array} {kind, value, sep} を要素とするセグメント配列
     */
    function parseFileName(name) {
        var splitTokens = String(name).split(/([-_])/); // ["handout","-","Adobe","-","20260422"]
        var segments = [];
        var textBuffer = [];      // token と区切りを交互に蓄積（join('') で結合）
        var textLeadingSep = '';  // テキスト segment の直前に置く区切り
        var hasDate = false, hasVersion = false, hasStatus = false, hasPage = false;
        var currentSep = '';

        /**
         * ためていたテキストトークンを 1 つの text セグメントとして確定する
         * @returns {void}
         */
        function flushText() {
            if (!textBuffer.length) return;
            segments.push({
                kind: 'text',
                value: textBuffer.join(''),
                sep: textLeadingSep
            });
            textBuffer = [];
            textLeadingSep = '';
        }

        var i = 0;
        while (i < splitTokens.length) {
            if (i % 2 === 1) {
                currentSep = splitTokens[i];
                i++;
                continue;
            }
            var token = splitTokens[i];
            // ステータス判定は同一トークンで 1 回だけ評価し、結果を再利用する
            var statusMatch = (FEATURE_STATUS && !hasStatus) ? matchStatusToken(token) : '';

            // YYYY-MM-DD / YYYY_MM_DD（3 トークン + 同じ区切り 2 つ）の日付パターン
            if (!hasDate && i > 0
                && /^\d{4}$/.test(token)
                && i + 4 < splitTokens.length
                && (splitTokens[i + 1] === '-' || splitTokens[i + 1] === '_')
                && /^\d{2}$/.test(splitTokens[i + 2])
                && splitTokens[i + 3] === splitTokens[i + 1]
                && /^\d{2}$/.test(splitTokens[i + 4])
                && isRealDate(token, splitTokens[i + 2], splitTokens[i + 4])) {
                flushText();
                var compositeDate = token + splitTokens[i + 1] + splitTokens[i + 2] + splitTokens[i + 3] + splitTokens[i + 4];
                segments.push({ kind: 'date', value: compositeDate, sep: currentSep });
                hasDate = true;
                i += 5; // 3 tokens + 2 separators
                continue;
            }

            if (i === 0) {
                // 仕様: 先頭トークンは常に base として扱う。
                // したがって "20260712-catalog" のような先頭日付は date ではなく base になる
                segments.push({ kind: 'base', value: token, sep: '' });
            } else if (!hasDate && isDateToken(token)) {
                flushText();
                segments.push({ kind: 'date', value: token, sep: currentSep });
                hasDate = true;
            } else if (!hasVersion && isVersionToken(token)) {
                flushText();
                segments.push({ kind: 'version', value: token, sep: currentSep });
                hasVersion = true;
            } else if (FEATURE_PAGE && !hasPage && isPageToken(token)) {
                flushText();
                segments.push({ kind: 'page', value: token, sep: currentSep });
                hasPage = true;
            } else if (statusMatch) {
                flushText();
                segments.push({ kind: 'status', value: statusMatch, sep: currentSep });
                hasStatus = true;
            } else {
                if (!textBuffer.length) {
                    textLeadingSep = currentSep;
                } else {
                    textBuffer.push(currentSep);
                }
                textBuffer.push(token);
            }
            i++;
        }
        flushText();
        return segments;
    }

    /**
     * 分散した text セグメント（間に date/status/page/version が挟まる）を 1 つの text にまとめる。
     * ただし認識済みセグメント（date/status/page/version）はタイトルへ吸収せず、順序を保って残す。
     * text の値どうしのみを結合し、最初の text の位置に置く
     * Merge fragmented text values into one (at the first text's position), while keeping recognized
     * segments (date/status/page/version) as their own segments — they are NOT folded into the title
     * @param {array} segments セグメント配列
     * @returns {array} text をまとめたセグメント配列
     */
    function mergeFragmentedText(segments) {
        var textIndices = [];
        for (var i = 0; i < segments.length; i++) {
            if (segments[i].kind === 'text') textIndices.push(i);
        }
        if (textIndices.length <= 1) return segments;
        var firstText = textIndices[0];
        // text 値のみを（各 text の直前区切りで）連結。認識済みセグメントの値は含めない
        var combinedText = segments[firstText].value;
        for (var j = 1; j < textIndices.length; j++) {
            var textSegment = segments[textIndices[j]];
            combinedText += textSegment.sep + textSegment.value;
        }
        var mergedSegments = [];
        for (var k = 0; k < segments.length; k++) {
            if (segments[k].kind === 'text') {
                // 最初の text だけを結合値で置き換え、以降の text は破棄（結合済み）
                if (k === firstText) {
                    mergedSegments.push({ kind: 'text', value: combinedText, sep: segments[firstText].sep });
                }
            } else {
                mergedSegments.push(segments[k]);
            }
        }
        return mergedSegments;
    }

    /**
     * parseFileName の kind を SEGMENT_ORDER の kind に変換しつつ、現在のファイル名で
     * 実際に出現した要素の並びを返す。重複・FEATURE_STATUS=false の status は除外
     * Map parsed segment kinds to SEGMENT_ORDER kinds in appearance order;
     * dedupe and drop 'status' when FEATURE_STATUS is false
     * @param {array} segments セグメント配列
     * @returns {array} 出現順の並び（SEGMENT_ORDER の kind）
     */
    function deriveOrderFromSegments(segments) {
        var kindToOrder = {
            base: 'base',
            text: 'title',
            status: 'status',
            date: 'timestamp',
            page: 'page',
            version: 'version'
        };
        var order = [];
        var seen = {};
        for (var i = 0; i < segments.length; i++) {
            var orderKind = kindToOrder[segments[i].kind];
            if (!orderKind || seen[orderKind]) continue;
            if (orderKind === 'status' && !FEATURE_STATUS) continue;
            if (orderKind === 'page' && !FEATURE_PAGE) continue;
            order.push(orderKind);
            seen[orderKind] = true;
        }
        return order;
    }

    /**
     * baseOrder（検出された並び）に含まれない SEGMENT_ORDER の kind を、正規位置に補完して返す。
     * 連番（page）は元ファイル名に現れないため常に欠落し、ステータスも未検出だと欠落する。
     * これらを正規位置へ挿入することで、有効化されているのに出力から漏れるのを防ぐ
     * （無効な要素は buildFinalName が空値としてスキップするので害はない）
     * Fill in SEGMENT_ORDER kinds missing from a detected order, inserting each at its
     * canonical position so enabled-but-undetected segments (page/status) aren't dropped.
     * Inactive kinds are harmless because buildFinalName skips empty values.
     * @param {array} baseOrder 検出された並び
     * @returns {array} 欠落分を補完した並び
     */
    function fillMissingSegmentOrder(baseOrder) {
        var filledOrder = baseOrder.slice();
        for (var i = 0; i < SEGMENT_ORDER.length; i++) {
            var kind = SEGMENT_ORDER[i];
            if (arrayContains(filledOrder, kind)) continue;
            // SEGMENT_ORDER 上で直前にあり、かつ filledOrder に既にある kind の直後へ挿入
            var insertAt = filledOrder.length;
            for (var j = i - 1; j >= 0; j--) {
                var foundIndex = -1;
                for (var k = 0; k < filledOrder.length; k++) {
                    if (filledOrder[k] === SEGMENT_ORDER[j]) { foundIndex = k; break; }
                }
                if (foundIndex !== -1) { insertAt = foundIndex + 1; break; }
            }
            filledOrder.splice(insertAt, 0, kind);
        }
        return filledOrder;
    }

    /**
     * segments から最初に出現する kind の value を取得 / Get the value of the first segment of the given kind
     * @param {array} segments セグメント配列
     * @param {string} kind 取得する kind
     * @returns {string} 最初に見つかった値（無ければ空文字）
     */
    function getFirstSegmentValue(segments, kind) {
        for (var i = 0; i < segments.length; i++) {
            if (segments[i].kind === kind) return segments[i].value;
        }
        return '';
    }

    /**
     * 指定 kind の segment が存在するか / Whether a segment of the given kind exists
     * @param {array} segments セグメント配列
     * @param {string} kind 探す kind
     * @returns {boolean} 存在すれば true
     */
    function hasSegmentKind(segments, kind) {
        for (var i = 0; i < segments.length; i++) {
            if (segments[i].kind === kind) return true;
        }
        return false;
    }

    /**
     * 文字列を左 0 パディング / Left-pad a string with zeros to the given width
     * @param {string} text 対象の文字列
     * @param {number} width 必要な桁数
     * @returns {string} 0 で左詰めした文字列
     */
    function padLeft(text, width) {
        while (text.length < width) text = '0' + text;
        return text;
    }

    /**
     * 今日の日付を返す。partSeparator で日付内区切り、withTime=true で末尾に "-HHMM" を付加
     * Today's date; with `partSeparator` between Y/M/D parts; `withTime=true` appends "-HHMM"
     * @param {string} partSeparator 年月日の間に挟む文字（省略可）
     * @param {boolean} withTime true なら末尾に "-HHMM" を付ける
     * @returns {string} 今日の日付文字列
     */
    function todayTimestamp(partSeparator, withTime) {
        partSeparator = partSeparator || '';
        var now = new Date();
        var dateText = String(now.getFullYear()) + partSeparator +
            padLeft(String(now.getMonth() + 1), 2) + partSeparator +
            padLeft(String(now.getDate()), 2);
        if (!withTime) return dateText;
        return dateText + '-' + padLeft(String(now.getHours()), 2) + padLeft(String(now.getMinutes()), 2);
    }

    /**
     * 現在のファイル名（拡張子なし）の v 番号だけを桁数維持で +1。
     * v 番号が無ければ末尾に "-v2" を付与
     * Bump the v-number inside the current basename in place (preserving digit width).
     * Appends "-v2" if no v-number exists.
     * @param {string} currentBaseName 拡張子を除いた現在のファイル名
     * @returns {string} v 番号を 1 つ繰り上げたファイル名
     */
    function bumpVersionInPlace(currentBaseName) {
        var versionParts = extractNumberedParts(currentBaseName, VERSION_TOKEN_SOURCE);
        if (!versionParts) return currentBaseName + '-v2';
        var nextNumber = parseInt(versionParts.digits, 10) + 1;
        var newDigits = padLeft(String(nextNumber), versionParts.digits.length);
        return versionParts.prefix + versionParts.token + newDigits + versionParts.suffix;
    }

    /**
     * バージョン文字列を +1。mode='padded' で 2 桁ゼロ埋め、'paddedWide' で 3 桁ゼロ埋め。
     * 元バージョンが無ければ新規付与（v1 / v01 / v001）
     * Bump the version string by +1. 'padded'=min 2 digits, 'paddedWide'=min 3. Returns v1/v01/v001 if no original
     * @param {string} originalVersion 元のバージョン文字列（無ければ空）
     * @param {string} mode 桁数モード（'padded' / 'paddedWide' / それ以外）
     * @returns {string} 繰り上げたバージョン文字列
     */
    function formatVersion(originalVersion, mode) {
        var versionMatch = String(originalVersion || '').match(/^([vV])(\d+)$/);
        var letter = versionMatch ? versionMatch[1] : 'v';
        var nextNumber = versionMatch ? (parseInt(versionMatch[2], 10) + 1) : 1;
        if (mode === 'padded') {
            var width = versionMatch ? Math.max(versionMatch[2].length, 2) : 2;
            return letter + padLeft(String(nextNumber), width);
        }
        if (mode === 'paddedWide') {
            var wideWidth = versionMatch ? Math.max(versionMatch[2].length, 3) : 3;
            return letter + padLeft(String(nextNumber), wideWidth);
        }
        return letter + String(nextNumber);
    }

    /* ダイアログ表示中はフォルダーの内容が変わらない前提で、ファイル名一覧を 1 度だけ列挙して使い回す。
       採番は 1 打鍵ごとに走るため、毎回 getFiles() すると数千ファイルのフォルダーで目に見えて重くなる
       / Cache the folder listing: numbering runs on every keystroke and getFiles() is expensive */
    var folderFileNamesCache = {};

    /**
     * フォルダー内のファイル名一覧を返す（同じフォルダーの 2 回目以降はキャッシュを返す）
     * @param {Folder} folder 走査するフォルダー
     * @returns {array} % デコード済みのファイル名の配列
     */
    function getFolderFileNames(folder) {
        if (!folder) return [];
        var key = folder.fsName;
        if (folderFileNamesCache[key]) return folderFileNamesCache[key];
        var folderFiles;
        try { folderFiles = folder.getFiles(); } catch (e) { return []; }
        var names = [];
        for (var i = 0; i < folderFiles.length; i++) {
            if (!(folderFiles[i] instanceof File)) continue;
            names.push(decodePercentEncoded(folderFiles[i].name));
        }
        folderFileNamesCache[key] = names;
        return names;
    }

    /**
     * フォルダー列挙のキャッシュを破棄する。ダイアログを閉じたあと、実保存前に呼んで
     * 表示中に増えたファイルを採番に反映させる
     * @returns {void}
     */
    function resetFolderFileNamesCache() {
        folderFileNamesCache = {};
    }

    /**
     * baseName 内の最初の「{tokenSource}+数字」セグメントを抽出する。マッチしなければ null。
     * 区切り（- _ .）か文字列端に挟まれたものだけを対象にするため、"rev1-catalog-v02" では
     * "rev1" ではなく "v02" を拾う（部分文字列で拾うと無関係な語が繰り上がる）
     * Extract the first {token}+digits *segment* (delimited by -, _, . or the string ends)
     * @param {string} baseName 拡張子を除いたファイル名
     * @param {string} tokenSource 数字の直前に来るトークンの正規表現ソース（"[vV]" など）
     * @returns {object} {prefix, token, digits, suffix}（マッチしなければ null）
     */
    function extractNumberedParts(baseName, tokenSource) {
        var pattern = new RegExp('^(|.*?[-_.])(' + tokenSource + ')(\\d+)(?=$|[-_.])(.*)$', 'i');
        var match = String(baseName).match(pattern);
        if (!match) return null;
        return { prefix: match[1], token: match[2], digits: match[3], suffix: match[4] };
    }

    /**
     * 正規表現エスケープ / Escape for use in RegExp
     * @param {string} text エスケープする文字列
     * @returns {string} 正規表現で安全に使える文字列
     */
    function escapeRegExp(text) {
        return String(text).replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&');
    }

    /**
     * baseName と同じ prefix / suffix を持つファイルを folder から探し、最大の番号を返す
     * Scan the folder for files matching baseName's numbering pattern; return the max number or null
     * @param {string} baseName 照合するベース名
     * @param {Folder} folder 走査するフォルダー
     * @param {string} extension 拡張子（".ai" など）
     * @param {string} tokenSource 数字の直前に来るトークンの正規表現ソース
     * @returns {number} 最大の番号（見つからなければ null）
     */
    function findMaxNumberedInFolder(baseName, folder, extension, tokenSource) {
        if (!folder) return null;
        var parts = extractNumberedParts(baseName, tokenSource);
        if (!parts) return null;
        var pattern = new RegExp(
            '^' + escapeRegExp(parts.prefix) + '(?:' + tokenSource + ')(\\d+)'
            + escapeRegExp(parts.suffix) + escapeRegExp(extension) + '$',
            'i'
        );
        var fileNames = getFolderFileNames(folder);
        var maxNumber = -1;
        for (var i = 0; i < fileNames.length; i++) {
            var fileMatch = fileNames[i].match(pattern);
            if (!fileMatch) continue;
            var number = parseInt(fileMatch[1], 10);
            if (number > maxNumber) maxNumber = number;
        }
        return maxNumber >= 0 ? maxNumber : null;
    }

    /**
     * baseName 内の番号を、親フォルダー内の同パターンの最大番号 +1 に置き換える。
     * 無い、または +1 が現在値以下のときは baseName のまま返す。桁数は維持
     * Replace the number with (folder max + 1); preserves digit width
     * @param {string} baseName 拡張子を除いたファイル名
     * @param {Folder} folder 走査するフォルダー
     * @param {string} extension 拡張子
     * @param {string} tokenSource 数字の直前に来るトークンの正規表現ソース
     * @returns {string} 衝突しない番号に置き換えたファイル名
     */
    function nextAvailableNumberedName(baseName, folder, extension, tokenSource) {
        var maxNumber = findMaxNumberedInFolder(baseName, folder, extension, tokenSource);
        if (maxNumber === null) return baseName;
        var parts = extractNumberedParts(baseName, tokenSource);
        if (!parts) return baseName;
        var currentNumber = parseInt(parts.digits, 10);
        var targetNumber = maxNumber + 1;
        if (targetNumber <= currentNumber) return baseName;
        var width = Math.max(parts.digits.length, String(targetNumber).length);
        return parts.prefix + parts.token + padLeft(String(targetNumber), width) + parts.suffix;
    }

    /**
     * v 番号を親フォルダー内の最大値 +1 に繰り上げる
     * @param {string} baseName 拡張子を除いたファイル名
     * @param {Folder} folder 走査するフォルダー
     * @param {string} extension 拡張子
     * @returns {string} 衝突しない v 番号に置き換えたファイル名
     */
    function nextAvailableVersionName(baseName, folder, extension) {
        return nextAvailableNumberedName(baseName, folder, extension, VERSION_TOKEN_SOURCE);
    }

    /**
     * 連番を親フォルダー内の最大値 +1 に繰り上げる
     * @param {string} baseName 拡張子を除いたファイル名
     * @param {Folder} folder 走査するフォルダー
     * @param {string} extension 拡張子
     * @param {string} prefix 連番のプレフィックス（空なら何もしない）
     * @returns {string} 衝突しない連番に置き換えたファイル名
     */
    function nextAvailablePageName(baseName, folder, extension, prefix) {
        var pagePrefix = String(prefix || '');
        if (!pagePrefix) return baseName;
        return nextAvailableNumberedName(baseName, folder, extension, escapeRegExp(pagePrefix));
    }

    /**
     * 連番プレフィックスにも本体と同じ整形を掛ける。ベース名は整形後の文字列で照合されるため、
     * 生の入力（"page " など）のままだとパターンが一致せず採番が空振りする。
     * 区切りの圧縮（collapseAndTrimSeparators）は末尾の区切りを落としてしまうので掛けない
     * @param {string} prefix 連番のプレフィックス
     * @param {object} uiState UI の状態
     * @returns {string} 整形後のプレフィックス
     */
    function transformPagePrefix(prefix, uiState) {
        var text = String(prefix || '');
        if (!text) return text;
        if (FEATURE_NFC && uiState.nfc === 'combine') text = normalizeNFC(text);
        if (FEATURE_HALFWIDTH_KANA && uiState.halfwidthKana === 'convert'
            && (uiState.clean === 'dash' || uiState.clean === 'underscore')) {
            text = convertHalfwidthKana(text);
        }
        if (FEATURE_TRANSLITERATE) text = transliterate(text, uiState.translit);
        if (FEATURE_CLEAN) text = cleanFilenameChars(text, uiState.clean);
        return text;
    }

    /**
     * ベース名に整形と採番をまとめて適用する。
     * プレビューと実保存の両方がこの 1 本を通ることで、ダイアログの表示と
     * 実際に書き出される名前がずれないようにしている
     * @param {string} baseName 整形前のベース名（拡張子なし）
     * @param {object} uiState UI の状態
     * @param {Folder} folder 採番で走査するフォルダー
     * @param {string} extension 拡張子（".ai"）
     * @returns {string} 整形・採番済みのベース名
     */
    function applyNameTransforms(baseName, uiState, folder, extension) {
        var name = baseName;
        if (FEATURE_NFC && uiState.nfc === 'combine') {
            name = normalizeNFC(name);
        }
        // 半角カナ → 全角カナ（clean が '-' / '_' のときだけ）。translit より先に行う
        if (FEATURE_HALFWIDTH_KANA && uiState.halfwidthKana === 'convert'
            && (uiState.clean === 'dash' || uiState.clean === 'underscore')) {
            name = convertHalfwidthKana(name);
        }
        // translit はクリーンより先に行う（㈱→株 などを残すため）
        if (FEATURE_TRANSLITERATE) {
            name = transliterate(name, uiState.translit);
        }
        if (FEATURE_CLEAN) {
            name = cleanFilenameChars(name, uiState.clean);
        }
        name = collapseAndTrimSeparators(name, uiState.separator);
        // 採番は整形後の名前で行う。フォルダー内の既存ファイル名は整形済みで保存されているため、
        // clean / translit / 区切り統一のあとに走らせないと prefix/suffix が一致せず最大値を取りこぼす
        if (uiState.version === 'short' || uiState.version === 'padded' || uiState.version === 'paddedWide') {
            name = nextAvailableVersionName(name, folder, extension);
        }
        if (uiState.pageEnable === 'yes') {
            name = nextAvailablePageName(name, folder, extension, transformPagePrefix(uiState.pagePrefix, uiState));
        }
        return escapeWindowsReserved(name);
    }

    /* Windows 予約名（拡張子の有無を問わず使用不可）。一致したら末尾に "_" を足してエスケープ
       / Windows-reserved basenames (regardless of extension); append "_" to escape them */
    var WINDOWS_RESERVED_NAMES = {
        CON: 1, PRN: 1, AUX: 1, NUL: 1,
        COM1: 1, COM2: 1, COM3: 1, COM4: 1, COM5: 1,
        COM6: 1, COM7: 1, COM8: 1, COM9: 1,
        LPT1: 1, LPT2: 1, LPT3: 1, LPT4: 1, LPT5: 1,
        LPT6: 1, LPT7: 1, LPT8: 1, LPT9: 1
    };

    /**
     * baseName が Windows 予約名と衝突するなら末尾に "_" を足して回避（大文字小文字無視）
     * If baseName matches a Windows reserved name, append "_" to avoid the collision
     * @param {string} baseName 拡張子を除いたファイル名
     * @returns {string} 予約名なら "_" を足した名前
     */
    function escapeWindowsReserved(baseName) {
        var upperName = String(baseName).toUpperCase();
        return (WINDOWS_RESERVED_NAMES[upperName] === 1) ? baseName + '_' : baseName;
    }

    /**
     * % エンコードを 1 回デコード / Decode a percent-encoded string once (best-effort)
     * @param {string} text 対象の文字列
     * @returns {string} デコードした文字列（失敗時は元の文字列）
     */
    function decodePercentEncoded(text) {
        text = String(text);
        if (text.indexOf('%') === -1) return text;
        try {
            return decodeURIComponent(text);
        } catch (e) {
            return text;
        }
    }

    /**
     * 拡張子を除去 / Strip a trailing file extension
     * @param {string} name 対象のファイル名
     * @returns {string} 拡張子を除いた名前
     */
    function stripExtension(name) {
        var dotIndex = name.lastIndexOf('.');
        return (dotIndex > 0) ? name.substring(0, dotIndex) : name;
    }

    /**
     * パス比較用に畳んだキー（% デコード + 濁点の NFC 結合 + 小文字化）。
     * macOS は大文字小文字を区別せず、ディスク上の日本語名は NFD で保持されるため、
     * 生の文字列比較では同じファイルを別ファイルと誤判定する
     * Fold a path for comparison: percent-decode, compose kana marks, lower-case
     * @param {string} path 対象のパス
     * @returns {string} 比較用に畳んだパス
     */
    function foldPathForCompare(path) {
        return normalizeNFC(decodePercentEncoded(String(path))).toLowerCase();
    }

    /**
     * 2 つのパスが同じファイルの綴り違い（大文字小文字 / 濁点の合成違い）かどうか
     * Whether two paths are just different spellings of the same file
     * @param {string} pathA 比較するパス
     * @param {string} pathB 比較するパス
     * @returns {boolean} 同じファイルを指していれば true
     */
    function isSamePathSpelling(pathA, pathB) {
        if (!pathA || !pathB) return false;
        if (pathA === pathB) return true;
        return foldPathForCompare(pathA) === foldPathForCompare(pathB);
    }

    /**
     * ファイルの更新日時（ミリ秒）。存在しない・読めない場合は null
     * File modification time in ms; null when missing or unreadable
     * @param {string} path 対象のパス
     * @returns {number} 更新日時（存在しない・読めない場合は null）
     */
    function fileModifiedTime(path) {
        if (!path) return null;
        var file = File(path);
        if (!file.exists) return null;
        try {
            return file.modified ? file.modified.getTime() : null;
        } catch (e) {
            return null;
        }
    }

    /**
     * 連続する区切り（- _ .）を 1 つに圧縮し、先頭・末尾の区切りと空白をトリム。
     * sep は最終的な区切り文字（'-' / '_' / '' のいずれか、'' のときは混在をそのまま残す）
     * Collapse runs of separators (-, _, .) to one and trim leading/trailing separators + spaces
     * @param {string} text 対象の文字列
     * @param {string} separator 最終的な区切り文字（'-' / '_' / ''）
     * @returns {string} 整えた文字列
     */
    function collapseAndTrimSeparators(text, separator) {
        var collapsed = String(text);
        if (separator === '-' || separator === '_') {
            // 異種の混在も含めて連続した区切りを separator 1 つに統一（YYYY-MM-DD 等は呼び出し前に通過済み想定）
            collapsed = collapsed.replace(/[-_.]{2,}/g, separator);
        } else {
            // 区切り未指定でも同種連続だけは圧縮（-- → -, __ → _, .. → .）
            collapsed = collapsed.replace(/--+/g, '-').replace(/__+/g, '_').replace(/\.\.+/g, '.');
        }
        // 先頭末尾の区切り・空白を除去
        collapsed = collapsed.replace(/^[-_.\s]+|[-_.\s]+$/g, '');
        return collapsed;
    }

    /**
     * 文字列の UTF-8 バイト長（拡張子込みのファイル名サイズ判定用）。ES3 / ExtendScript 向け
     * UTF-8 byte length of a string (for file-name size checks); ES3-safe
     * @param {string} text 対象の文字列
     * @returns {number} UTF-8 でのバイト長
     */
    function byteLengthUTF8(text) {
        text = String(text);
        var byteCount = 0;
        for (var i = 0; i < text.length; i++) {
            var code = text.charCodeAt(i);
            if (code < 0x80) byteCount += 1;
            else if (code < 0x800) byteCount += 2;
            else if (code >= 0xD800 && code <= 0xDBFF) { byteCount += 4; i++; } // サロゲートペア（4 バイト）
            else byteCount += 3;
        }
        return byteCount;
    }

    /**
     * 前後空白をトリム。FEATURE_CLEAN=false のときは OS 禁止文字も除去（後段で処理されないため）
     * Trim whitespace. When FEATURE_CLEAN is off, also strip OS-invalid chars (no post-process safety net)
     * @param {string} text 対象の文字列
     * @returns {string} ファイル名に使える文字列
     */
    function sanitizeFilename(text) {
        var trimmed = String(text).replace(/^\s+|\s+$/g, '');
        if (FEATURE_CLEAN) return trimmed;
        return trimmed.replace(/[\\\/:*?"<>|]/g, '');
    }

    /* 濁点（U+3099）結合マップ: 合成可能な base → 合成済み濁音
       / Base → voiced composed form (combining dakuten U+3099) */
    var NFC_DAKUTEN_MAP = {
        // ひらがな
        'か': 'が', 'き': 'ぎ', 'く': 'ぐ', 'け': 'げ', 'こ': 'ご',
        'さ': 'ざ', 'し': 'じ', 'す': 'ず', 'せ': 'ぜ', 'そ': 'ぞ',
        'た': 'だ', 'ち': 'ぢ', 'つ': 'づ', 'て': 'で', 'と': 'ど',
        'は': 'ば', 'ひ': 'び', 'ふ': 'ぶ', 'へ': 'べ', 'ほ': 'ぼ',
        'う': 'ゔ',
        // カタカナ
        'カ': 'ガ', 'キ': 'ギ', 'ク': 'グ', 'ケ': 'ゲ', 'コ': 'ゴ',
        'サ': 'ザ', 'シ': 'ジ', 'ス': 'ズ', 'セ': 'ゼ', 'ソ': 'ゾ',
        'タ': 'ダ', 'チ': 'ヂ', 'ツ': 'ヅ', 'テ': 'デ', 'ト': 'ド',
        'ハ': 'バ', 'ヒ': 'ビ', 'フ': 'ブ', 'ヘ': 'ベ', 'ホ': 'ボ',
        'ウ': 'ヴ', 'ワ': 'ヷ', 'ヰ': 'ヸ', 'ヱ': 'ヹ', 'ヲ': 'ヺ'
    };

    /* 半濁点（U+309A）結合マップ: 合成可能な base → 合成済み半濁音（ハ行のみ）
       / Base → semi-voiced composed form (combining handakuten U+309A; ha-row only) */
    var NFC_HANDAKUTEN_MAP = {
        // ひらがな
        'は': 'ぱ', 'ひ': 'ぴ', 'ふ': 'ぷ', 'へ': 'ぺ', 'ほ': 'ぽ',
        // カタカナ
        'ハ': 'パ', 'ヒ': 'ピ', 'フ': 'プ', 'ヘ': 'ペ', 'ホ': 'ポ'
    };

    /**
     * HFS+/APFS 由来の NFD 文字列（か + ゛など）を NFC（が）に結合。
     * 合成可能な組み合わせのみ変換し、合成不能な濁点・半濁点（例: あ゙）は元の文字列を維持する
     * Normalize NFD-style kana + combining marks to composed NFC form (map-based).
     * Only combinable pairs are converted; non-combinable sequences are kept as-is
     * @param {string} text 対象の文字列
     * @returns {string} 濁点・半濁点を結合した文字列
     */
    function normalizeNFC(text) {
        return String(text).replace(/(.)([゙゚])/g, function (_, base, mark) {
            var markMap = (mark === '゙') ? NFC_DAKUTEN_MAP : NFC_HANDAKUTEN_MAP;
            var composed = markMap[base];
            return (typeof composed === 'string') ? composed : base + mark;
        });
    }

    /**
     * ファイル名で安全に使える「標準」文字か。
     * OS 禁止文字（\ / : * ? " < > |）と半角・全角スペースは ASCII / CJK 範囲だが除外。
     * ASCII printable / CJK 記号 / 仮名 / 漢字 / 全角 ASCII / 半角カタカナ を許容
     * Whether a code point is "standard" for filenames; OS-invalid chars and spaces are excluded
     * @param {number} code 判定する文字コード
     * @returns {boolean} 標準的な文字なら true
     */
    function isStandardFilenameChar(code) {
        // OS 禁止文字: \ 5C / 2F : 3A * 2A ? 3F " 22 < 3C > 3E | 7C
        if (code === 0x5C || code === 0x2F || code === 0x3A || code === 0x2A
            || code === 0x3F || code === 0x22 || code === 0x3C
            || code === 0x3E || code === 0x7C) return false;
        // 半角・全角スペース
        if (code === 0x20 || code === 0x3000) return false;
        if (code >= 0x21 && code <= 0x7E) return true;     // ASCII printable（space を除く）
        if (code >= 0x3001 && code <= 0x303F) return true; // CJK 記号と句読点（、。「」など / 全角スペースを除く）
        if (code >= 0x3040 && code <= 0x309F) return true; // ひらがな
        if (code >= 0x30A0 && code <= 0x30FF) return true; // カタカナ
        if (code >= 0x3400 && code <= 0x4DBF) return true; // CJK 拡張 A
        if (code >= 0x4E00 && code <= 0x9FFF) return true; // CJK 基本
        if (code >= 0xAC00 && code <= 0xD7A3) return true; // ハングル音節
        if (code >= 0xFF00 && code <= 0xFF5E) return true; // 全角 ASCII
        if (code >= 0xFF61 && code <= 0xFF9F) return true; // 半角カタカナ
        return false;
    }

    /**
     * OS 禁止文字 + 絵文字 + 機種依存文字 + スペースを mode に応じて置換／削除。
     * 連続するスペース（タブ・改行含む）は事前に 1 つに畳んでから処理。サロゲートペアは 2 文字単位
     * mode: 'remove'（削除） / 'dash'（-）/ 'underscore'（_）
     * Clean OS-invalid + emoji + platform-dependent chars + spaces per mode
     * @param {string} text 対象の文字列
     * @param {string} mode 処理モード（'remove' / 'dash' / 'underscore'）
     * @returns {string} 整えた文字列
     */
    function cleanFilenameChars(text, mode) {
        var replacement = (mode === 'dash') ? '-' : (mode === 'underscore') ? '_' : '';
        // 連続する空白を 1 つにまとめてから per-char で置換（mode='dash' で "  " が "--" にならないように）
        var spaceCollapsed = String(text).replace(/\s+/g, ' ');
        var cleaned = '';
        var i = 0;
        while (i < spaceCollapsed.length) {
            var code = spaceCollapsed.charCodeAt(i);
            if (code >= 0xD800 && code <= 0xDBFF && i + 1 < spaceCollapsed.length) {
                // 高サロゲート + 低サロゲートは常に非標準（BMP 外の絵文字など）
                cleaned += replacement;
                i += 2;
                continue;
            }
            if (isStandardFilenameChar(code)) {
                cleaned += spaceCollapsed.charAt(i);
            } else {
                cleaned += replacement;
            }
            i++;
        }
        return cleaned;
    }

    /* 半角カタカナ → 全角カタカナのマップ（清音）/ Half-width → full-width katakana (unvoiced) */
    var HALFWIDTH_KANA_MAP = {
        'ｦ': 'ヲ', 'ｧ': 'ァ', 'ｨ': 'ィ', 'ｩ': 'ゥ', 'ｪ': 'ェ', 'ｫ': 'ォ',
        'ｬ': 'ャ', 'ｭ': 'ュ', 'ｮ': 'ョ', 'ｯ': 'ッ', 'ｰ': 'ー',
        'ｱ': 'ア', 'ｲ': 'イ', 'ｳ': 'ウ', 'ｴ': 'エ', 'ｵ': 'オ',
        'ｶ': 'カ', 'ｷ': 'キ', 'ｸ': 'ク', 'ｹ': 'ケ', 'ｺ': 'コ',
        'ｻ': 'サ', 'ｼ': 'シ', 'ｽ': 'ス', 'ｾ': 'セ', 'ｿ': 'ソ',
        'ﾀ': 'タ', 'ﾁ': 'チ', 'ﾂ': 'ツ', 'ﾃ': 'テ', 'ﾄ': 'ト',
        'ﾅ': 'ナ', 'ﾆ': 'ニ', 'ﾇ': 'ヌ', 'ﾈ': 'ネ', 'ﾉ': 'ノ',
        'ﾊ': 'ハ', 'ﾋ': 'ヒ', 'ﾌ': 'フ', 'ﾍ': 'ヘ', 'ﾎ': 'ホ',
        'ﾏ': 'マ', 'ﾐ': 'ミ', 'ﾑ': 'ム', 'ﾒ': 'メ', 'ﾓ': 'モ',
        'ﾔ': 'ヤ', 'ﾕ': 'ユ', 'ﾖ': 'ヨ',
        'ﾗ': 'ラ', 'ﾘ': 'リ', 'ﾙ': 'ル', 'ﾚ': 'レ', 'ﾛ': 'ロ',
        'ﾜ': 'ワ', 'ﾝ': 'ン'
    };

    /* 半角カナ + ﾞ の結合マップ（濁音）/ Half-width + dakuten → voiced full-width */
    var HALFWIDTH_KANA_VOICED = {
        'ｶ': 'ガ', 'ｷ': 'ギ', 'ｸ': 'グ', 'ｹ': 'ゲ', 'ｺ': 'ゴ',
        'ｻ': 'ザ', 'ｼ': 'ジ', 'ｽ': 'ズ', 'ｾ': 'ゼ', 'ｿ': 'ゾ',
        'ﾀ': 'ダ', 'ﾁ': 'ヂ', 'ﾂ': 'ヅ', 'ﾃ': 'デ', 'ﾄ': 'ド',
        'ﾊ': 'バ', 'ﾋ': 'ビ', 'ﾌ': 'ブ', 'ﾍ': 'ベ', 'ﾎ': 'ボ',
        'ｳ': 'ヴ'
    };

    /* 半角カナ + ﾟ の結合マップ（半濁音）/ Half-width + handakuten → semi-voiced full-width */
    var HALFWIDTH_KANA_SEMI_VOICED = {
        'ﾊ': 'パ', 'ﾋ': 'ピ', 'ﾌ': 'プ', 'ﾍ': 'ペ', 'ﾎ': 'ポ'
    };

    /**
     * 半角カタカナを全角カタカナへ変換。ｶﾞ → ガ、ﾊﾟ → パ のように濁点・半濁点も結合
     * Convert half-width katakana to full-width; merges trailing dakuten / handakuten
     * @param {string} text 対象の文字列
     * @returns {string} 半角カナを全角にした文字列
     */
    function convertHalfwidthKana(text) {
        text = String(text);
        var converted = '';
        var i = 0;
        while (i < text.length) {
            var currentChar = text.charAt(i);
            var nextChar = (i + 1 < text.length) ? text.charAt(i + 1) : '';
            if (nextChar === 'ﾞ' && typeof HALFWIDTH_KANA_VOICED[currentChar] === 'string') {
                converted += HALFWIDTH_KANA_VOICED[currentChar];
                i += 2;
                continue;
            }
            if (nextChar === 'ﾟ' && typeof HALFWIDTH_KANA_SEMI_VOICED[currentChar] === 'string') {
                converted += HALFWIDTH_KANA_SEMI_VOICED[currentChar];
                i += 2;
                continue;
            }
            var fullwidthKana = HALFWIDTH_KANA_MAP[currentChar];
            if (typeof fullwidthKana === 'string') {
                converted += fullwidthKana;
            } else if (currentChar === 'ﾞ') {
                converted += '゛';
            } else if (currentChar === 'ﾟ') {
                converted += '゜';
            } else {
                converted += currentChar;
            }
            i++;
        }
        return converted;
    }

    /**
     * TRANSLITERATE_MAP に従って 1 文字ずつ処理。mode='convert' は変換、'remove' は削除、'keep' は無処理。
     * 値が文字列のときのみ対象（toString などの prototype プロパティ衝突を回避）
     * Per-char transform via TRANSLITERATE_MAP: 'convert' replaces, 'remove' drops, 'keep' returns as-is
     * @param {string} text 対象の文字列
     * @param {string} mode 処理モード（'remove' / 'convert' / それ以外）
     * @returns {string} 変換後の文字列
     */
    function transliterate(text, mode) {
        if (mode === 'keep') return String(text);
        text = String(text);
        var converted = '';
        for (var i = 0; i < text.length; i++) {
            var currentChar = text.charAt(i);
            var replacement = TRANSLITERATE_MAP[currentChar];
            if (typeof replacement === 'string') {
                converted += (mode === 'remove') ? '' : replacement;
            } else {
                converted += currentChar;
            }
        }
        return converted;
    }

    /**
     * 保存先ダイアログ。.ai 拡張子を補正して File を返す。キャンセルは null / Show save dialog, normalize to .ai; null on cancel
     * @param {string} promptLabel ダイアログに表示する説明
     * @returns {File} 選択されたファイル（キャンセル時は null）
     */
    function pickAiDestination(promptLabel) {
        var chosenFile = File.saveDialog(promptLabel, '*.ai');
        if (!chosenFile) return null;
        var aiFile = File(chosenFile.fsName);
        if (!/\.ai$/i.test(aiFile.name)) {
            aiFile = File(aiFile.parent.fsName + '/' + stripExtension(aiFile.name) + '.ai');
        }
        return aiFile;
    }

    // 設定の保存（再利用パーツ） / Settings store (reusable)

    var SETTINGS_STORE_FOLDER_NAME = "illustrator-scripts"; /* Folder.userData の下に作るフォルダー / folder created under Folder.userData */
    var SETTINGS_STORE_MAX_DEPTH = 32;                                /* 入れ子の上限（循環参照よけ）/ nesting limit (guards against cycles) */

    /**
     * 設定の保存先を作る。寿命は "session"（Illustrator の終了まで）か "persistent"（ファイルに保存）
     * @param {string} storeName - 保存名（ふつうは SCRIPT_NAME）。ファイル名と $.global のキーに使う
     * @param {string} lifetime - "session" または "persistent"
     * @param {Object} [storeOptions] - { legacy: function () → 旧形式の保存値のオブジェクト|null }
     * @returns {{load: Function, save: Function, clear: Function}} 読み込み・保存・消去の関数
     */
    function createSettingsStore(storeName, lifetime, storeOptions) {
        var isPersistent = (lifetime === "persistent");
        var legacyReader = (storeOptions && typeof storeOptions.legacy === "function") ? storeOptions.legacy : null;
        var safeStoreName = String(storeName).replace(/[\\\/:*?"<>|]/g, "_");
        var sessionKey = "__" + safeStoreName + "_Settings";
        var settingsFile = isPersistent
            ? new File(Folder.userData + "/" + SETTINGS_STORE_FOLDER_NAME + "/" + safeStoreName + ".json")
            : null;

        /**
         * 保存してある文字列を返す
         * @returns {string|null} 保存文字列。1度も保存していなければ null
         */
        function readStoredText() {
            if (!isPersistent) {
                return (typeof $.global[sessionKey] === "string") ? $.global[sessionKey] : null;
            }
            return settingsStoreReadTextFile(settingsFile);
        }

        /**
         * 文字列を保存する
         * @param {string} storedText - 保存する文字列
         * @returns {boolean} 保存できたら true
         */
        function writeStoredText(storedText) {
            if (!isPersistent) {
                $.global[sessionKey] = storedText;
                return true;
            }
            return settingsStoreWriteTextFile(settingsFile, storedText);
        }

        /**
         * 保存値を読み込み、既定値と突き合わせて返す（型の合わない値・知らない項目は捨てる）
         * @param {Object} defaultSettings - 既定値
         * @returns {Object} 設定（毎回新しいオブジェクト）
         */
        function load(defaultSettings) {
            var savedSettings = null;
            try {
                var storedText = readStoredText();
                if (storedText !== null) {
                    savedSettings = settingsStoreParse(storedText);
                } else if (legacyReader) {
                    savedSettings = legacyReader();
                }
            } catch (e) {
                $.writeln("SettingsStore.load(" + storeName + "): " + e);
                savedSettings = null;
            }
            return settingsStoreMerge(defaultSettings, savedSettings);
        }

        /**
         * 設定を保存する
         * @param {Object} settingValues - 保存する値
         * @returns {boolean} 保存できたら true
         */
        function save(settingValues) {
            try {
                return writeStoredText(settingsStoreSerialize(settingValues, "", 0));
            } catch (e) {
                $.writeln("SettingsStore.save(" + storeName + "): " + e);
                return false;
            }
        }

        /**
         * 保存を消す。旧形式を読み継ぐストアでは空の保存を書き、旧設定が戻らないようにする
         * @returns {boolean} 消せたら true
         */
        function clear() {
            if (legacyReader) return writeStoredText("{}");
            if (!isPersistent) {
                try { delete $.global[sessionKey]; } catch (e) { $.global[sessionKey] = undefined; }
                return true;
            }
            try {
                return settingsFile.exists ? settingsFile.remove() : true;
            } catch (e) {
                $.writeln("SettingsStore.clear(" + storeName + "): " + e);
                return false;
            }
        }

        return { load: load, save: save, clear: clear };
    }

    /**
     * 旧形式の設定ファイルを読む（key=value の行 / toSource / JSON を自動判別。eval は使わない）
     * @param {File|string} legacyFileOrPath - 旧ファイルかそのパス
     * @returns {Object|null} 読み込んだ値（key=value は値がすべて文字列）。無い・読めないときは null
     */
    function readSettingsLegacyFile(legacyFileOrPath) {
        try {
            var legacyFile = (legacyFileOrPath instanceof File) ? legacyFileOrPath : new File(legacyFileOrPath);
            var legacyText = settingsStoreReadTextFile(legacyFile);
            return (legacyText === null) ? null : settingsStoreParseLegacyText(legacyText);
        } catch (e) {
            $.writeln("readSettingsLegacyFile: " + e);
            return null;
        }
    }

    /**
     * app.preferences に文字列で保存していた旧設定を読む（形式は readSettingsLegacyFile と同じく自動判別）
     * @param {string} preferenceKey - 環境設定のキー
     * @returns {Object|null} 読み込んだ値。無い・読めないときは null
     */
    function readSettingsLegacyPreference(preferenceKey) {
        try {
            var legacyText = app.preferences.getStringPreference(preferenceKey);
            if (!legacyText) return null;
            return settingsStoreParseLegacyText(String(legacyText));
        } catch (e) {
            $.writeln("readSettingsLegacyPreference: " + e);
            return null;
        }
    }

    /**
     * テキストファイルを UTF-8 で読む
     * @param {File} textFile - 読むファイル
     * @returns {string|null} 中身。ファイルが無ければ null
     */
    function settingsStoreReadTextFile(textFile) {
        if (!textFile.exists) return null;
        textFile.encoding = "UTF-8";
        if (!textFile.open("r")) throw new Error("cannot open " + textFile.fsName);
        try {
            return textFile.read().replace(/^\uFEFF/, "");
        } finally {
            textFile.close();
        }
    }

    /**
     * テキストファイルを UTF-8 で書く（フォルダーが無ければ作る）
     * @param {File} textFile - 書くファイル
     * @param {string} fileText - 中身
     * @returns {boolean} 書けたら true
     */
    function settingsStoreWriteTextFile(textFile, fileText) {
        try {
            var parentFolder = textFile.parent;
            if (!parentFolder.exists && !parentFolder.create()) throw new Error("cannot create " + parentFolder.fsName);
            textFile.encoding = "UTF-8";
            textFile.lineFeed = "Unix";
            if (!textFile.open("w")) throw new Error("cannot open " + textFile.fsName);
            try {
                textFile.write(fileText);
            } finally {
                textFile.close();
            }
            return true;
        } catch (e) {
            $.writeln("SettingsStore write: " + e);
            return false;
        }
    }

    /**
     * 値が配列か
     * @param {*} checkedValue - 調べる値
     * @returns {boolean} 配列なら true
     */
    function settingsStoreIsArray(checkedValue) {
        return Object.prototype.toString.call(checkedValue) === "[object Array]";
    }

    /**
     * 値が素のオブジェクト（{ } で作ったもの）か
     * @param {*} checkedValue - 調べる値
     * @returns {boolean} 素のオブジェクトなら true
     */
    function settingsStoreIsPlainObject(checkedValue) {
        return checkedValue !== null && typeof checkedValue === "object"
            && Object.prototype.toString.call(checkedValue) === "[object Object]"
            && checkedValue.constructor === Object;
    }

    /**
     * 文字列を JSON の文字列リテラルにする（ASCII 以外は \uXXXX にして、文字コードの取り違えに強くする）
     * @param {string} sourceText - 文字列
     * @returns {string} 引用符つきの文字列
     */
    function settingsStoreQuote(sourceText) {
        var quotedText = "\"";
        for (var i = 0; i < sourceText.length; i++) {
            var charCode = sourceText.charCodeAt(i);
            var oneChar = sourceText.charAt(i);
            if (oneChar === "\"" || oneChar === "\\") quotedText += "\\" + oneChar;
            else if (oneChar === "\n") quotedText += "\\n";
            else if (oneChar === "\r") quotedText += "\\r";
            else if (oneChar === "\t") quotedText += "\\t";
            else if (charCode < 0x20 || charCode > 0x7E) quotedText += "\\u" + ("0000" + charCode.toString(16)).slice(-4);
            else quotedText += oneChar;
        }
        return quotedText + "\"";
    }

    /**
     * 値を JSON の文字列にする（オブジェクトは1項目1行、中身が値だけの配列は1行）。
     * undefined・関数・DOM オブジェクトは項目ごと省き、配列の中では null にする。有限でない数値は null
     * @param {*} sourceValue - 値
     * @param {string} indentText - 今の字下げ
     * @param {number} depth - 入れ子の深さ
     * @returns {string|undefined} JSON の文字列。書けない値は undefined
     */
    function settingsStoreSerialize(sourceValue, indentText, depth) {
        if (depth > SETTINGS_STORE_MAX_DEPTH) throw new Error("settings are nested too deeply");
        if (sourceValue === null) return "null";
        var valueType = typeof sourceValue;
        if (valueType === "boolean") return sourceValue ? "true" : "false";
        if (valueType === "number") return isFinite(sourceValue) ? String(sourceValue) : "null";
        if (valueType === "string") return settingsStoreQuote(sourceValue);
        var innerIndent = indentText + "  ";
        var itemTexts = [];
        var i;
        if (settingsStoreIsArray(sourceValue)) {
            var hasNested = false;
            for (i = 0; i < sourceValue.length; i++) {
                var itemText = settingsStoreSerialize(sourceValue[i], innerIndent, depth + 1);
                itemTexts.push(itemText === undefined ? "null" : itemText);
                if (sourceValue[i] !== null && typeof sourceValue[i] === "object") hasNested = true;
            }
            if (!itemTexts.length) return "[]";
            if (!hasNested) return "[" + itemTexts.join(", ") + "]";
            return "[\n" + innerIndent + itemTexts.join(",\n" + innerIndent) + "\n" + indentText + "]";
        }
        if (settingsStoreIsPlainObject(sourceValue)) {
            for (var key in sourceValue) {
                if (!sourceValue.hasOwnProperty(key)) continue;
                var memberText = settingsStoreSerialize(sourceValue[key], innerIndent, depth + 1);
                if (memberText !== undefined) itemTexts.push(settingsStoreQuote(key) + ": " + memberText);
            }
            if (!itemTexts.length) return "{}";
            return "{\n" + innerIndent + itemTexts.join(",\n" + innerIndent) + "\n" + indentText + "}";
        }
        return undefined; /* 関数・DOM オブジェクトなど / functions, DOM objects, etc. */
    }

    /**
     * JSON（と toSource の出力）を読む。eval は使わない。
     * キーの引用符なし・'…' の文字列・全体の ( ) ・末尾のカンマ・(void 0) も受け付ける
     * @param {string} sourceText - 読む文字列
     * @returns {*} 読み込んだ値
     */
    function settingsStoreParse(sourceText) {
        var readPos = 0;
        var textLength = sourceText.length;

        /**
         * 読み取り位置で失敗を知らせる
         * @param {string} reasonText - 理由
         * @returns {void}
         */
        function fail(reasonText) {
            throw new Error("settings parse error at " + readPos + ": " + reasonText);
        }

        /**
         * 空白を読み飛ばす
         * @returns {void}
         */
        function skipSpaces() {
            while (readPos < textLength && /\s/.test(sourceText.charAt(readPos))) readPos++;
        }

        /**
         * 識別子（英数字・_・$）を読む
         * @returns {string} 識別子。無ければ空文字
         */
        function readWord() {
            var startPos = readPos;
            while (readPos < textLength && /[\w$]/.test(sourceText.charAt(readPos))) readPos++;
            return sourceText.substring(startPos, readPos);
        }

        /**
         * 引用符で囲んだ文字列を読む（" と ' のどちらでも）
         * @returns {string} 文字列
         */
        function readString() {
            var quoteChar = sourceText.charAt(readPos++);
            var resultText = "";
            while (readPos < textLength) {
                var oneChar = sourceText.charAt(readPos++);
                if (oneChar === quoteChar) return resultText;
                if (oneChar !== "\\") { resultText += oneChar; continue; }
                var escapeChar = sourceText.charAt(readPos++);
                if (escapeChar === "n") resultText += "\n";
                else if (escapeChar === "r") resultText += "\r";
                else if (escapeChar === "t") resultText += "\t";
                else if (escapeChar === "b") resultText += "\b";
                else if (escapeChar === "f") resultText += "\f";
                else if (escapeChar === "v") resultText += "\v";
                else if (escapeChar === "0") resultText += "\0";
                else if (escapeChar === "u" || escapeChar === "x") {
                    var hexLength = (escapeChar === "u") ? 4 : 2;
                    var hexText = sourceText.substr(readPos, hexLength);
                    if (!new RegExp("^[0-9A-Fa-f]{" + hexLength + "}$").test(hexText)) fail("bad escape");
                    resultText += String.fromCharCode(parseInt(hexText, 16));
                    readPos += hexLength;
                } else resultText += escapeChar;
            }
            fail("unterminated string");
        }

        /**
         * 値を1つ読む
         * @param {number} depth - 入れ子の深さ
         * @returns {*} 値
         */
        function readValue(depth) {
            if (depth > SETTINGS_STORE_MAX_DEPTH) fail("nested too deeply");
            skipSpaces();
            var oneChar = sourceText.charAt(readPos);
            if (oneChar === "{") return readObject(depth);
            if (oneChar === "[") return readArray(depth);
            if (oneChar === "\"" || oneChar === "'") return readString();
            if (oneChar === "(") {
                readPos++;
                var innerValue = readValue(depth + 1);
                skipSpaces();
                if (sourceText.charAt(readPos) !== ")") fail("expected )");
                readPos++;
                return innerValue;
            }
            var numberMatch = /^-?(\d+\.?\d*|\.\d+)([eE][+\-]?\d+)?/.exec(sourceText.substring(readPos, readPos + 64));
            if (numberMatch) {
                readPos += numberMatch[0].length;
                return Number(numberMatch[0]);
            }
            var wordText = readWord();
            if (wordText === "true") return true;
            if (wordText === "false") return false;
            if (wordText === "null") return null;
            if (wordText === "NaN") return NaN;
            if (wordText === "Infinity") return Infinity;
            if (wordText === "void") { readValue(depth + 1); return undefined; } /* toSource の (void 0) */
            fail("unexpected " + (wordText || oneChar || "end of text"));
        }

        /**
         * 配列を読む
         * @param {number} depth - 入れ子の深さ
         * @returns {Array} 配列
         */
        function readArray(depth) {
            var resultArray = [];
            readPos++;
            skipSpaces();
            while (sourceText.charAt(readPos) !== "]") {
                resultArray.push(readValue(depth + 1));
                skipSpaces();
                if (sourceText.charAt(readPos) === ",") { readPos++; skipSpaces(); continue; }
                if (sourceText.charAt(readPos) !== "]") fail("expected , or ]");
            }
            readPos++;
            return resultArray;
        }

        /**
         * オブジェクトを読む（__proto__ のキーは捨てる）
         * @param {number} depth - 入れ子の深さ
         * @returns {Object} オブジェクト
         */
        function readObject(depth) {
            var resultObject = {};
            readPos++;
            skipSpaces();
            while (sourceText.charAt(readPos) !== "}") {
                var keyChar = sourceText.charAt(readPos);
                var memberKey = (keyChar === "\"" || keyChar === "'") ? readString() : readWord();
                if (memberKey === "") fail("expected a key");
                skipSpaces();
                if (sourceText.charAt(readPos) !== ":") fail("expected :");
                readPos++;
                var memberValue = readValue(depth + 1);
                if (memberKey !== "__proto__") resultObject[memberKey] = memberValue;
                skipSpaces();
                if (sourceText.charAt(readPos) === ",") { readPos++; skipSpaces(); continue; }
                if (sourceText.charAt(readPos) !== "}") fail("expected , or }");
            }
            readPos++;
            return resultObject;
        }

        var parsedValue = readValue(0);
        skipSpaces();
        if (readPos < textLength) fail("unexpected text after the value");
        return parsedValue;
    }

    /**
     * 旧形式の文字列を読む。{ [ ( で始まれば JSON / toSource、それ以外は key=value の行とみなす
     * @param {string} legacyText - 旧形式の文字列
     * @returns {Object|null} 読み込んだ値
     */
    function settingsStoreParseLegacyText(legacyText) {
        var trimmedText = legacyText.replace(/^\uFEFF/, "").replace(/^\s+|\s+$/g, "");
        if (trimmedText === "") return null;
        if (/^[\{\[\(]/.test(trimmedText)) return settingsStoreParse(trimmedText);
        var keyValues = {};
        var textLines = trimmedText.split(/\r\n|\r|\n/);
        for (var i = 0; i < textLines.length; i++) {
            var separatorIndex = textLines[i].indexOf("=");
            if (separatorIndex < 1) continue;
            var lineKey = textLines[i].substring(0, separatorIndex).replace(/^\s+|\s+$/g, "");
            if (lineKey !== "" && lineKey !== "__proto__") keyValues[lineKey] = textLines[i].substring(separatorIndex + 1);
        }
        return keyValues;
    }

    /**
     * 値を深くコピーする（素のデータだけ。関数・DOM オブジェクトは null）
     * @param {*} sourceValue - コピー元
     * @returns {*} コピー
     */
    function settingsStoreClone(sourceValue) {
        if (sourceValue === null || typeof sourceValue !== "object") {
            return (typeof sourceValue === "function" || sourceValue === undefined) ? null : sourceValue;
        }
        var i;
        if (settingsStoreIsArray(sourceValue)) {
            var arrayCopy = [];
            for (i = 0; i < sourceValue.length; i++) arrayCopy.push(settingsStoreClone(sourceValue[i]));
            return arrayCopy;
        }
        if (!settingsStoreIsPlainObject(sourceValue)) return null;
        var objectCopy = {};
        for (var key in sourceValue) {
            if (sourceValue.hasOwnProperty(key)) objectCopy[key] = settingsStoreClone(sourceValue[key]);
        }
        return objectCopy;
    }

    /**
     * 保存値を既定値と突き合わせる。型は既定値に合わせ、合わなければ既定値を使う。
     * 既定値が {} か null なら中身を問わず受け取り、配列は配列なら受け取る。既定値に無い項目は捨てる
     * @param {*} defaultValue - 既定値
     * @param {*} savedValue - 保存値
     * @returns {*} 突き合わせた値（新しいオブジェクト）
     */
    function settingsStoreMerge(defaultValue, savedValue) {
        if (defaultValue === null || defaultValue === undefined) {
            return (savedValue === undefined) ? null : settingsStoreClone(savedValue);
        }
        var defaultType = typeof defaultValue;
        var savedType = typeof savedValue;
        if (defaultType === "boolean") {
            if (savedType === "boolean") return savedValue;
            if (savedValue === 1 || savedValue === "1" || savedValue === "true") return true;
            if (savedValue === 0 || savedValue === "0" || savedValue === "false") return false;
            return defaultValue;
        }
        if (defaultType === "number") {
            if (savedType === "number" && isFinite(savedValue)) return savedValue;
            if (savedType === "string" && /\S/.test(savedValue)) {
                var parsedNumber = Number(savedValue);
                if (isFinite(parsedNumber)) return parsedNumber;
            }
            return defaultValue;
        }
        if (defaultType === "string") {
            if (savedType === "string") return savedValue;
            if (savedType === "number" && isFinite(savedValue)) return String(savedValue);
            if (savedType === "boolean") return String(savedValue);
            return defaultValue;
        }
        if (settingsStoreIsArray(defaultValue)) {
            return settingsStoreClone(settingsStoreIsArray(savedValue) ? savedValue : defaultValue);
        }
        if (defaultType === "object") {
            var savedIsObject = settingsStoreIsPlainObject(savedValue);
            var hasDefaultKeys = false;
            var mergedObject = {};
            for (var key in defaultValue) {
                if (!defaultValue.hasOwnProperty(key)) continue;
                hasDefaultKeys = true;
                mergedObject[key] = settingsStoreMerge(defaultValue[key], savedIsObject ? savedValue[key] : undefined);
            }
            /* 既定値が {} なら自由な入れ物として中身ごと受け取る / an empty default {} is a free-form map */
            if (!hasDefaultKeys && savedIsObject) return settingsStoreClone(savedValue);
            return mergedObject;
        }
        return defaultValue;
    }

    // 設定の保存（再利用パーツ）ここまで / End of the reusable settings store

    /* プリセットの保存先は Folder.userData/illustrator-scripts/Ai-FileNameManager.json。
       旧名のまま据え置き（改名で既存のプリセットを失わないため）。旧版の FileNameManager-prefs.txt は最初の1回だけ読み継ぐ
       Presets live in illustrator-scripts/Ai-FileNameManager.json; the old FileNameManager-prefs.txt is read once */
    var settingsStore = createSettingsStore("Ai-FileNameManager", "persistent", {
        legacy: function () {
            return readSettingsLegacyFile(Folder.userData.fsName + '/FileNameManager-prefs.txt');
        }
    });

    /**
     * 前回保存したプリセットを読み込み（無ければ空オブジェクト） / Load previously saved prefs (or empty if none)
     * 機能の有無で項目が変わり、無い項目は各所の既定値に任せるので、既定値は {}（中身をそのまま受け取る）
     * @returns {object} 保存されていた設定（値はすべて文字列。無ければ空オブジェクト）
     */
    function loadPrefs() {
        return settingsStore.load({});
    }

    /**
     * プリセットを保存（失敗時は黙って継続） / Save prefs (silently ignore failure)
     * @param {object} prefs 保存する設定
     * @returns {void}
     */
    function savePrefs(prefs) {
        /* 旧形式と同じく値は文字列で持つ / Values stay strings, as in the old key=value file */
        var prefValues = {};
        for (var key in prefs) {
            if (prefs.hasOwnProperty(key)) prefValues[key] = String(prefs[key]);
        }
        settingsStore.save(prefValues);
    }

    /**
     * ドキュメントから現在のファイル情報を収集。Illustrator は未保存ドキュメントでも fullName を返すため、
     * fullName.exists で保存済みかを判定する
     * Collect file info from the active document; in Illustrator fullName exists even when unsaved,
     * so use fullName.exists to detect a saved-to-disk file
     * @param {Document} doc 対象のドキュメント
     * @returns {object} ファイル名・パス・親フォルダー名などの情報
     */
    function gatherDocumentInfo(doc) {
        var fullName;
        var savedToDisk = false;
        try {
            fullName = doc.fullName;
            savedToDisk = !!(fullName && fullName.exists);
        } catch (e) { /* unsaved document */ }
        var rawName = (fullName && fullName.name) ? fullName.name : doc.name;
        var currentName = decodePercentEncoded(rawName);
        var parentFolderName = '';
        var grandparentFolderName = '';
        if (savedToDisk && fullName.parent) {
            parentFolderName = decodePercentEncoded(fullName.parent.name);
            if (fullName.parent.parent) {
                grandparentFolderName = decodePercentEncoded(fullName.parent.parent.name);
            }
        }
        return {
            currentName: currentName,
            baseName: stripExtension(currentName),
            fsPath: savedToDisk ? fullName.fsName : null,
            folder: savedToDisk ? fullName.parent : null,
            parentFolderName: parentFolderName,
            grandparentFolderName: grandparentFolderName
        };
    }

    /**
     * 書類が Illustrator ネイティブ（.ai）かどうか。保存されるのは常に .ai なので、
     * .ai 以外の書類でリネームすると形式変換したうえで原本を失う。未保存なら消す原本が無いので許可。
     * コピーも同様で、元ファイルをそのままバイトコピーするため中身が PDF のまま拡張子だけ .ai になる
     * Whether the document is a native .ai file. The output is always .ai, so renaming a non-.ai
     * document would convert the format and destroy the original, and "save a copy" would produce
     * a file whose contents do not match its .ai extension. Unsaved documents have nothing to lose
     * @param {object} documentInfo gatherDocumentInfo が返すファイル情報
     * @returns {boolean} .ai（または未保存）なら true
     */
    function isNativeDocument(documentInfo) {
        if (!documentInfo.fsPath) return true;
        return /\.ai$/i.test(documentInfo.currentName);
    }

    /**
     * 保存先フォルダを確定（未保存なら saveDialog で取得）。キャンセル時は null / Resolve the target folder (prompt if needed); null on cancel
     * @param {Folder} folder 既に判明しているフォルダー（無ければ null）
     * @returns {Folder} 保存先フォルダー（キャンセル時は null）
     */
    function ensureTargetFolder(folder) {
        if (folder) return folder;
        var pickedFile = pickAiDestination(getLabel('message.chooseDestination'));
        return pickedFile ? pickedFile.parent : null;
    }

    /**
     * 既存同名（自分自身を除く）の上書き確認。OK なら true / Confirm overwrite for an existing file (excluding self); true if approved
     * @param {File} destFile 保存先のファイル
     * @param {string} originalFsPath 元ファイルのパス
     * @returns {boolean} 続行してよければ true
     */
    function confirmOverwriteIfExists(destFile, originalFsPath) {
        if (!destFile.exists) return true;
        // 大文字小文字・濁点の合成違いだけのリネームは自分自身への上書きなので確認しない
        if (isSamePathSpelling(destFile.fsName, originalFsPath)) return true;
        return confirm(getLabel('message.confirmOverwrite') + '\n\n' + destFile.fsName);
    }

    /**
     * file を ~/.Trash に移動。同名衝突時は連番でユニーク化。成功で true、失敗で false
     * Move file to ~/.Trash, disambiguating by appending a counter. Returns true on success
     * @param {File} sourceFile 移動するファイル
     * @returns {boolean} 移動できたら true
     */
    function moveToTrash(sourceFile) {
        if (!sourceFile || !sourceFile.exists) return false;
        var trashFolder = Folder("~/.Trash");
        if (!trashFolder.exists) return false;
        var originalName = sourceFile.name;
        var dotIndex = originalName.lastIndexOf('.');
        var stemName = (dotIndex > 0) ? originalName.substring(0, dotIndex) : originalName;
        var extension = (dotIndex > 0) ? originalName.substring(dotIndex) : '';
        var trashDest = File(trashFolder.fsName + '/' + originalName);
        var counter = 1;
        while (trashDest.exists) {
            trashDest = File(trashFolder.fsName + '/' + stemName + ' ' + counter + extension);
            counter++;
            if (counter > 9999) return false; // 暴走防止
        }
        try {
            if (!sourceFile.copy(trashDest)) return false;
            if (sourceFile.remove()) return true;
            // オリジナルを消せなかった場合はゴミ箱側のコピーを後始末（二重化を防ぐ）
            try { trashDest.remove(); } catch (removeError) { /* 後始末に失敗しても続行 */ }
            return false;
        } catch (e) {
            return false;
        }
    }

    /**
     * 旧ファイルを削除（FEATURE_USE_TRASH=true ならゴミ箱に移動、失敗時は file.remove() にフォールバック）
     * Remove the original file: move to ~/.Trash when enabled, else (or on failure) hard-remove
     * @param {string} originalFsPath 元ファイルのパス
     * @param {string} destFsPath 保存先のパス
     * @param {number} modifiedBeforeSave 保存前の元ファイルの更新日時（ミリ秒、無ければ null）
     * @returns {void}
     */
    function removeOriginalFile(originalFsPath, destFsPath, modifiedBeforeSave) {
        if (!originalFsPath || isSamePathSpelling(originalFsPath, destFsPath)) return;
        var originalFile = File(originalFsPath);
        if (!originalFile.exists) return;
        // 綴り違いを取りこぼしても保存したてのファイルを消さないための最終防衛線。
        // 直前の saveAs は保存先しか触らないので、元ファイルの更新日時が動いていたら
        // それは保存先と同一実体（大文字小文字・NFD/NFC 違いのパス）である
        var modifiedNow = fileModifiedTime(originalFsPath);
        if (modifiedBeforeSave !== null && modifiedNow !== null && modifiedNow !== modifiedBeforeSave) return;
        if (FEATURE_USE_TRASH && moveToTrash(originalFile)) return;
        try { originalFile.remove(); } catch (e) { /* 削除できない場合は黙って継続 */ }
    }

    /**
     * モード別の出力処理（rename / saveAs / copy） / Execute the output according to the selected mode
     * @param {Document} doc 対象のドキュメント
     * @param {File} destFile 保存先のファイル
     * @param {string} mode 動作モード（'rename' / 'saveAs' / 'copy'）
     * @param {string} originalFsPath 元ファイルのパス（未保存なら null）
     * @returns {void}
     */
    function executeOutput(doc, destFile, mode, originalFsPath) {
        if (mode === 'copy' && originalFsPath) {
            // 現在の変更を元ファイルへ保存してから、物理ファイルとしてコピー
            if (!doc.saved) doc.save();
            var originalFile = File(originalFsPath);
            if (!originalFile.exists) {
                throw new Error(getLabel('message.saveFailed') + '\n' + destFile.fsName);
            }
            // File.copy() は既存ファイルを上書きしない。上書きは確認済みなので先に取り除く
            if (destFile.exists && !destFile.remove()) {
                throw new Error(getLabel('message.saveFailed') + '\n' + destFile.fsName);
            }
            if (!originalFile.copy(destFile)) {
                throw new Error(getLabel('message.saveFailed') + '\n' + destFile.fsName);
            }
            return;
        }
        // rename / saveAs / 未保存ドキュメントの copy: 新名で保存
        var modifiedBeforeSave = fileModifiedTime(originalFsPath);
        var saveOptions = new IllustratorSaveOptions();
        saveOptions.pdfCompatible = FEATURE_PDF_COMPATIBLE;
        doc.saveAs(destFile, saveOptions);
        if (mode === 'rename') {
            removeOriginalFile(originalFsPath, destFile.fsName, modifiedBeforeSave);
        }
    }

    /**
     * segments と UI 状態から最終ファイル名（拡張子なし）を構築。SEGMENT_ORDER に従う
     * Build the final filename from segments and UI state, following SEGMENT_ORDER
     * @param {array} segments セグメント配列
     * @param {object} uiState UI の状態
     * @returns {string} 拡張子を除いた最終ファイル名
     */
    function buildFinalName(segments, uiState) {
        // 区切り記号: 明示選択があればそれを、無ければ元のファイル名で優勢な区切りを使う
        var defaultSeparator;
        if (uiState.separator === '-' || uiState.separator === '_') {
            defaultSeparator = uiState.separator;
        } else {
            defaultSeparator = dominantSeparator(segments);
        }

        /**
         * セグメントの kind ごとに、UI 状態から出力する値を求める
         * @param {string} kind セグメントの kind
         * @returns {string} その kind の値（空文字なら出力しない）
         */
        function valueForKind(kind) {
            if (kind === 'base') {
                return sanitizeFilename(uiState.baseText || '');
            }
            if (kind === 'title') {
                if (uiState.titleMode === 'none') return '';
                if (uiState.titleMode === 'parent') return sanitizeFilename(uiState.parentFolderName);
                if (uiState.titleMode === 'grandparent') return sanitizeFilename(uiState.grandparentFolderName);
                return sanitizeFilename(uiState.titleText);
            }
            if (kind === 'status') {
                return uiState.status || '';
            }
            if (kind === 'timestamp') {
                var withTime = (uiState.timestampTime === 'hhmm');
                if (uiState.timestamp === 'date') return todayTimestamp('', withTime);
                if (uiState.timestamp === 'dateDash') return todayTimestamp('-', withTime);
                return '';
            }
            if (kind === 'version') {
                if (uiState.version === 'none') return '';
                return formatVersion(getFirstSegmentValue(segments, 'version'), uiState.version);
            }
            if (kind === 'page') {
                if (uiState.pageEnable !== 'yes') return '';
                var pageWidth = (uiState.pagePad === '3') ? 3 : 2;
                // 既存の連番があれば「既存 +1」を初期候補にする（無ければ 1）。
                // フォルダー内の最大連番 +1 への繰り上げは後段の nextAvailablePageName が担う
                var pageStart = 1;
                var existingPage = getFirstSegmentValue(segments, 'page');
                if (existingPage) {
                    var existingPageMatch = String(existingPage).match(PAGE_TOKEN_RE);
                    if (existingPageMatch) pageStart = parseInt(existingPageMatch[2], 10) + 1;
                }
                return (uiState.pagePrefix || '') + padLeft(String(pageStart), pageWidth);
            }
            return '';
        }

        // ピースごとに区切り統一（YYYY-MM-DD のタイムスタンプだけは内部 "-" を保護）
        var order = (uiState.segmentOrder && uiState.segmentOrder.length) ? uiState.segmentOrder : SEGMENT_ORDER;
        var nameParts = [];
        for (var i = 0; i < order.length; i++) {
            var kind = order[i];
            var segmentValue = valueForKind(kind);
            if (!segmentValue) continue;
            // YYYY-MM-DD（および YYYY-MM-DD-HHMM、YYYYMMDD-HHMM）は内部の "-" を保護
            var isProtectedDate = (kind === 'timestamp')
                && (uiState.timestamp === 'dateDash' || uiState.timestampTime === 'hhmm');
            if (!isProtectedDate && FEATURE_SEPARATOR) {
                // 区切り記号統一。FEATURE_DOT_NORMALIZE が true なら "." も対象
                if (uiState.separator === '-') {
                    segmentValue = FEATURE_DOT_NORMALIZE ? segmentValue.replace(/[_.]/g, '-') : segmentValue.replace(/_/g, '-');
                } else if (uiState.separator === '_') {
                    segmentValue = FEATURE_DOT_NORMALIZE ? segmentValue.replace(/[-.]/g, '_') : segmentValue.replace(/-/g, '_');
                }
            }
            nameParts.push(segmentValue);
        }
        var joinedName = nameParts.join(defaultSeparator);
        // スペース（半角・全角・タブ・改行）の処理は後段の cleanFilenameChars に集約
        return joinedName;
    }

    /**
     * segments から優勢な区切り記号を返す。同数なら "-" / Dominant separator across segments (defaults to "-")
     * @param {array} segments セグメント配列
     * @returns {string} 優勢な区切り文字
     */
    function dominantSeparator(segments) {
        var dashes = 0, underscores = 0;
        for (var i = 0; i < segments.length; i++) {
            if (segments[i].sep === '-') dashes++;
            else if (segments[i].sep === '_') underscores++;
        }
        return (underscores > dashes) ? '_' : '-';
    }

    // =========================================
    // ダイアログビルダー / Dialog builders
    // =========================================

    /**
     * ソートパネルを構築（標準順 / 現在のファイル名に準じる / カスタム順 + [順序を編集] ボタン）。
     * ボタンは「カスタム順」のときだけ有効。currentOrderAvailable=false なら「現在...」は無効化
     * Build the sort panel: Default / Match Current / Custom + [Edit order] button.
     * The button is enabled only for "Custom". When currentOrderAvailable=false, "Match Current" is disabled
     * @param {Group} parent 追加先のコンテナ
     * @param {boolean} currentOrderAvailable 「現在のファイル名に準じる」を選べるか
     * @returns {object} パネル内のコントロール
     */
    function buildSortPanel(parent, currentOrderAvailable) {
        var panel = parent.add('panel', undefined, getLabel('panel.sort'));
        setupPanel(panel, 6);
        var sortOffRadio = panel.add('radiobutton', undefined, getLabel('radio.sortOff'));
        sortOffRadio.helpTip = getLabel('tip.sort');
        var sortCurrentRadio = panel.add('radiobutton', undefined, getLabel('radio.sortCurrent'));
        sortCurrentRadio.helpTip = getLabel('tip.sort');
        if (!currentOrderAvailable) sortCurrentRadio.enabled = false;
        // 「カスタム順」ラジオと「順序を編集...」ボタンを同じ行に並べる
        var customRow = panel.add('group');
        customRow.orientation = 'row';
        customRow.alignment = ['fill', 'top'];
        customRow.alignChildren = ['left', 'center'];
        customRow.spacing = 8;
        var sortOnRadio = customRow.add('radiobutton', undefined, getLabel('radio.sortOn'));
        sortOnRadio.helpTip = getLabel('tip.sort');
        var sortButton = customRow.add('button', undefined, getLabel('button.sort'));
        // 並び順の初期値は prefs を見ず、常に「現在のファイル名に準じる」（不可なら「標準順」）に固定
        var initialSort = currentOrderAvailable ? 'current' : 'off';
        /**
         * 3 つのラジオを明示的に排他制御する。「カスタム順」だけが customRow の中にいるため、
         * ScriptUI の自動排他（同じ親コンテナ内でのみ働く）が 3 つ揃っては効かない
         * Select one of the three radios explicitly: "Custom" lives inside customRow, so
         * ScriptUI's built-in exclusivity (same parent only) does not span all three
         * @param {string} mode 選択する並び順（'off' / 'current' / 'on'）
         * @returns {void}
         */
        function selectSortMode(mode) {
            sortOffRadio.value = (mode === 'off');
            sortCurrentRadio.value = (mode === 'current');
            sortOnRadio.value = (mode === 'on');
            sortButton.enabled = (mode === 'on');
        }
        selectSortMode(initialSort);
        // onClick は呼び出し側で wire（refreshPreviews と組み合わせるため）
        return {
            panel: panel,
            sortOffRadio: sortOffRadio,
            sortCurrentRadio: sortCurrentRadio,
            sortOnRadio: sortOnRadio,
            sortButton: sortButton,
            /* 'off' / 'current' / 'on' */
            getSortMode: function () {
                if (sortOnRadio.value) return 'on';
                if (sortCurrentRadio.value) return 'current';
                return 'off';
            },
            isSortOn: function () { return sortOnRadio.value; },
            selectSortMode: selectSortMode
        };
    }

    /**
     * 「モード」パネルを構築（バージョン番号のみ / 全体）。
     * defaultVersionOnly=true なら「バージョンのみ」、それ以外は「全体」を初期選択
     * Build the scope panel (Version Only / Full); defaults to "Version Only" when defaultVersionOnly is true, else "Full"
     * @param {Group} parent 追加先のコンテナ
     * @param {boolean} defaultVersionOnly 「バージョンのみ」を初期選択にするか
     * @returns {object} パネル内のコントロール
     */
    function buildOpModePanel(parent, defaultVersionOnly) {
        var panel = parent.add('panel', undefined, getLabel('panel.opMode'));
        setupPanel(panel, 6);
        var versionOnlyRadio = panel.add('radiobutton', undefined, getLabel('radio.opVersionOnly'));
        var fullRadio = panel.add('radiobutton', undefined, getLabel('radio.opFull'));
        versionOnlyRadio.value = !!defaultVersionOnly;
        fullRadio.value = !defaultVersionOnly;
        return {
            panel: panel,
            versionOnlyRadio: versionOnlyRadio,
            fullRadio: fullRadio,
            isVersionOnly: function () { return versionOnlyRadio.value; },
            getOpMode: function () { return versionOnlyRadio.value ? 'versionOnly' : 'full'; }
        };
    }

    /**
     * モード選択パネルを構築（リネーム / 別名で保存 / コピーを保存） / Build the mode panel (Rename / Save As / Save a Copy)
     * @param {Group} parent 追加先のコンテナ
     * @param {boolean} isNative 書類が .ai ネイティブか
     * @returns {object} パネル内のコントロール
     */
    function buildModePanel(parent, isNative) {
        var panel = parent.add('panel', undefined, getLabel('panel.mode'));
        setupPanel(panel, 6);
        var renameRadio = panel.add('radiobutton', undefined, getLabel('radio.rename'));
        // .ai 以外の書類は保存が形式変換になるため、元ファイルを消すリネームは選ばせない
        renameRadio.enabled = !!isNative;
        renameRadio.helpTip = isNative ? getLabel('tip.rename') : getLabel('tip.nonNativeUnsupported');
        var saveAsRadio = panel.add('radiobutton', undefined, getLabel('radio.saveAs'));
        saveAsRadio.helpTip = getLabel('tip.saveAs');
        var saveCopyRadio = panel.add('radiobutton', undefined, getLabel('radio.saveCopy'));
        // コピーは元ファイルのバイトコピーなので、.ai 以外だと中身と拡張子が食い違う
        saveCopyRadio.enabled = !!isNative;
        saveCopyRadio.helpTip = isNative ? getLabel('tip.saveCopy') : getLabel('tip.nonNativeUnsupported');
        // 初期選択は常に「別名で保存」
        renameRadio.value = false;
        saveAsRadio.value = true;
        saveCopyRadio.value = false;
        return {
            panel: panel,
            renameRadio: renameRadio,
            saveAsRadio: saveAsRadio,
            saveCopyRadio: saveCopyRadio,
            /* 現在選択中のモード（'rename' / 'saveAs' / 'copy'） / Currently selected mode */
            getMode: function () {
                if (renameRadio.value) return 'rename';
                if (saveCopyRadio.value) return 'copy';
                return 'saveAs';
            }
        };
    }

    /**
     * ファイル名パネルを構築（現在名・変更後名のみ） / Build the file-name panel (current / final only)
     * @param {Group} parent 追加先のコンテナ
     * @param {string} currentName 現在のファイル名（拡張子込み）
     * @returns {object} パネル内のコントロールとラベル参照
     */
    function buildFilenamePanel(parent, currentName) {
        var panel = parent.add('panel', undefined, getLabel('panel.filename'));
        setupPanel(panel, 6);

        var currentNameRow = panel.add('group');
        currentNameRow.orientation = 'row';
        var currentNameLabel = currentNameRow.add('statictext', undefined, labelText('label.currentName'), { justify: 'right' });
        currentNameRow.add('statictext', undefined, currentName);

        var finalNameRow = panel.add('group');
        finalNameRow.orientation = 'row';
        var finalNameLabel = finalNameRow.add('statictext', undefined, labelText('label.finalName'), { justify: 'right' });
        // 「変更後：」は statictext のためレイアウト後にサイズ固定。
        // 現在のファイル名と「入力フィールド + 余白」の大きい方を確保しておく
        var finalNameValue = finalNameRow.add('statictext', undefined, currentName);
        var currentNameWidth = panel.graphics.measureString(currentName).width;
        finalNameValue.preferredSize.width = Math.max(currentNameWidth + 20, 340);

        // 個別整列はせず、ラベル参照を呼び出し側に返し、後段で全パネル統一整列する
        return {
            panel: panel,
            finalNameValue: finalNameValue,
            labels: [currentNameLabel, finalNameLabel],
            labelTexts: [labelText('label.currentName'), labelText('label.finalName')]
        };
    }

    /**
     * オプションパネルを構築（ベース表示・タイトル選択・タイムスタンプ・バージョン番号・区切り） / Build the options panel
     * @param {Group} parent 追加先のコンテナ
     * @param {array} segments セグメント配列
     * @param {object} prefs 保存しておいた設定
     * @param {string} parentFolderName 親フォルダー名
     * @param {string} grandparentFolderName 2 階層上のフォルダー名
     * @returns {object} パネル内のコントロールと取得関数
     */
    function buildOptionsPanel(parent, segments, prefs, parentFolderName, grandparentFolderName) {
        var panel = parent.add('panel', undefined, getLabel('panel.options'));
        setupPanel(panel, 6);

        // ベース: 検出値を初期表示する入力欄。空欄可、prefs には保存しない
        var detectedBase = getFirstSegmentValue(segments, 'base');
        var baseRow = panel.add('group');
        baseRow.orientation = 'row';
        var baseLabel = baseRow.add('statictext', undefined, labelText('label.base'), { justify: 'right' });
        baseLabel.helpTip = getLabel('tip.base');
        var baseField = baseRow.add('edittext', undefined, detectedBase);
        baseField.preferredSize.width = NEW_NAME_FIELD_WIDTH;
        baseField.helpTip = getLabel('tip.base');

        // 元ファイル名のテキスト部（base / status / date / version 以外）をサブテキストとして検出
        var detectedTitle = getFirstSegmentValue(segments, 'text');

        // サブテキスト: 1 行目 = ラベル + 3 ラジオ、2 行目 = 「指定」用の入力欄
        var titleSection = panel.add('group');
        titleSection.orientation = 'column';
        titleSection.alignChildren = ['fill', 'top'];
        titleSection.spacing = 4;

        var titleRow = addRadioRow(titleSection, 'label.title', 'tip.title', [
            { key: 'none', text: getLabel('radio.titleNone') },
            { key: 'custom', text: getLabel('radio.titleCustom') },
            { key: 'parent', text: getLabel('radio.titleParent') },
            { key: 'grandparent', text: getLabel('radio.titleGrandparent') }
        ]);
        titleRow.group.alignment = ['left', 'top'];
        // 親/2 階層上のフォルダー名がある場合は helpTip にフォルダ名を追記、無ければ無効化
        if (parentFolderName) titleRow.radios.parent.helpTip = getLabel('tip.title') + ' (' + parentFolderName + ')';
        else titleRow.radios.parent.enabled = false;
        if (grandparentFolderName) titleRow.radios.grandparent.helpTip = getLabel('tip.title') + ' (' + grandparentFolderName + ')';
        else titleRow.radios.grandparent.enabled = false;

        // 「指定」用の入力欄は次の行（ラベル列幅だけ左に余白を入れて radios に揃える）
        var titleFieldRow = titleSection.add('group');
        titleFieldRow.orientation = 'row';
        titleFieldRow.alignment = ['left', 'top'];
        var titleFieldSpacer = titleFieldRow.add('statictext', undefined, '');
        var titleField = titleFieldRow.add('edittext', undefined, detectedTitle);
        titleField.preferredSize.width = NEW_NAME_FIELD_WIDTH;
        titleField.helpTip = getLabel('tip.title');

        // 初期モード: 元ファイル名にサブテキストが無ければ「なし」を強制。
        // 検出できた場合は prefs.titleMode を優先、無ければ 'custom'（parent/grandparent は対応フォルダ名必須）
        var initialTitleMode = !detectedTitle ? 'none' :
            pickPref(prefs, 'titleMode', ['none', 'parent', 'grandparent', 'custom'], 'custom');
        if (initialTitleMode === 'parent' && !parentFolderName) {
            initialTitleMode = detectedTitle ? 'custom' : 'none';
        }
        if (initialTitleMode === 'grandparent' && !grandparentFolderName) {
            initialTitleMode = detectedTitle ? 'custom' : 'none';
        }
        titleRow.radios.none.value = (initialTitleMode === 'none');
        titleRow.radios.parent.value = (initialTitleMode === 'parent');
        titleRow.radios.grandparent.value = (initialTitleMode === 'grandparent');
        titleRow.radios.custom.value = (initialTitleMode === 'custom');
        titleField.enabled = (initialTitleMode === 'custom');

        /**
         * サブテキスト入力欄の有効／無効を選択中のラジオに合わせる
         * @returns {void}
         */
        function syncTitleFieldEnabled() {
            titleField.enabled = titleRow.radios.custom.value;
        }

        // ステータス（dropdown。先頭は「なし」）
        var statusLabel = null, statusDropdown = null;
        if (FEATURE_STATUS) {
            var statusRow = panel.add('group');
            statusRow.orientation = 'row';
            statusLabel = statusRow.add('statictext', undefined, labelText('label.status'), { justify: 'right' });
            statusLabel.helpTip = getLabel('tip.status');
            statusDropdown = statusRow.add('dropdownlist', undefined, undefined);
            statusDropdown.helpTip = getLabel('tip.status');
            var parsedStatus = getFirstSegmentValue(segments, 'status');
            var initialStatus = parsedStatus || ((prefs && prefs.status) ? prefs.status : '');
            var initialStatusIndex = -1;
            for (var statusIndex = 0; statusIndex < STATUS_ITEMS.length; statusIndex++) {
                var statusItem = STATUS_ITEMS[statusIndex];
                var dropdownItem = statusDropdown.add('item', (uiLang === "ja" ? statusItem.ja : statusItem.en));
                if (isStatusDivider(statusItem)) {
                    dropdownItem.enabled = false;
                    continue;
                }
                if (initialStatusIndex === -1 && statusItem.value === initialStatus) {
                    initialStatusIndex = statusIndex;
                }
            }
            if (initialStatusIndex === -1) initialStatusIndex = 0;
            statusDropdown.selection = initialStatusIndex;
        }

        // タイムスタンプ（なし / YYYYMMDD / YYYY-MM-DD。デフォルト YYYYMMDD）
        // 末尾に「時刻も付与」(HHMM) チェックボックスを同居
        var timestampRow = addRadioRow(panel, 'label.timestamp', 'tip.timestamp', [
            { key: 'none', text: getLabel('radio.timestampNone') },
            { key: 'date', text: getLabel('radio.timestampDate') },
            { key: 'dateDash', text: getLabel('radio.timestampDateDash') }
        ]);
        var initialTimestamp = pickPref(prefs, 'timestamp', ['none', 'date', 'dateDash'], 'date');
        timestampRow.radios.none.value = (initialTimestamp === 'none');
        timestampRow.radios.date.value = (initialTimestamp === 'date');
        timestampRow.radios.dateDash.value = (initialTimestamp === 'dateDash');

        var timestampHHMMCheckbox = timestampRow.group.add('checkbox', undefined, getLabel('radio.timestampWithTime'));
        timestampHHMMCheckbox.helpTip = getLabel('tip.timestampWithTime');
        timestampHHMMCheckbox.value = pickPref(prefs, 'timestampTime', ['no', 'hhmm'], 'no') === 'hhmm';

        /**
         * 「時刻も付与」が ON なのにタイムスタンプが「なし」だと時刻はファイル名に出ない。
         * ON にした時点で YYYYMMDD を選び直す（ユーザーが明示的に「なし」へ戻すのは妨げない）
         * Turning on "append HHMM" while the timestamp is "none" would print nothing,
         * so select YYYYMMDD at that moment
         * @returns {void}
         */
        function coerceTimestampForHHMM() {
            if (!timestampHHMMCheckbox.value) return;
            if (!timestampRow.radios.none.value) return;
            timestampRow.radios.none.value = false;
            timestampRow.radios.date.value = true;
        }
        coerceTimestampForHHMM();

        // 連番（チェックボックス + プレフィックス入力 + 桁数ラジオ）。デフォルト OFF / "page" / 2 桁
        // SEGMENT_ORDER 内では timestamp の後・version の前に配置
        var pageRow = null;
        var pageCheckbox = null;
        var pagePrefixField = null;
        var pagePadRadio2 = null;
        var pagePadRadio3 = null;
        if (FEATURE_PAGE) {
            // 元ファイル名から連番セグメント（page03 など）を検出。あればプレフィックス・桁数を初期値に反映し、
            // 既定で ON にして既存の連番が出力から消えないようにする（無ければ prefs / 既定値）
            var detectedPage = getFirstSegmentValue(segments, 'page');
            var detectedPageMatch = detectedPage ? String(detectedPage).match(PAGE_TOKEN_RE) : null;
            var detectedPagePrefix = detectedPageMatch ? detectedPageMatch[1] : '';
            var detectedPagePad = detectedPageMatch ? (detectedPageMatch[2].length >= 3 ? '3' : '2') : '';

            pageRow = panel.add('group');
            pageRow.orientation = 'row';
            var pageLabel = pageRow.add('statictext', undefined, labelText('label.page'), { justify: 'right' });
            pageLabel.helpTip = getLabel('tip.page');
            pageCheckbox = pageRow.add('checkbox', undefined, getLabel('radio.pageEnable'));
            pageCheckbox.helpTip = getLabel('tip.page');
            // 既存の連番を検出したら prefs より優先して必ず ON（既存 pageNN を消さない）
            pageCheckbox.value = detectedPage
                ? true
                : (pickPref(prefs, 'pageEnable', ['no', 'yes'], 'no') === 'yes');
            pagePrefixField = pageRow.add('edittext', undefined, '');
            pagePrefixField.preferredSize.width = 60;
            pagePrefixField.helpTip = getLabel('tip.page');
            pagePrefixField.text = detectedPagePrefix
                ? detectedPagePrefix
                : ((prefs && typeof prefs.pagePrefix === 'string') ? prefs.pagePrefix : 'page');
            pagePadRadio2 = pageRow.add('radiobutton', undefined, getLabel('radio.pagePad2'));
            pagePadRadio2.helpTip = getLabel('tip.page');
            pagePadRadio3 = pageRow.add('radiobutton', undefined, getLabel('radio.pagePad3'));
            pagePadRadio3.helpTip = getLabel('tip.page');
            var initialPagePad = detectedPagePad ? detectedPagePad : pickPref(prefs, 'pagePad', ['2', '3'], '2');
            pagePadRadio2.value = (initialPagePad === '2');
            pagePadRadio3.value = (initialPagePad === '3');
            pageRow.label = pageLabel; // alignLabelWidths 用に統一形にしておく
        }

        /**
         * 連番のプレフィックス欄と桁数ラジオの有効／無効をチェックボックスに合わせる
         * @returns {void}
         */
        function syncPageControlsEnabled() {
            if (!pageCheckbox) return;
            var enabled = pageCheckbox.value;
            if (pagePrefixField) pagePrefixField.enabled = enabled;
            if (pagePadRadio2) pagePadRadio2.enabled = enabled;
            if (pagePadRadio3) pagePadRadio3.enabled = enabled;
        }
        syncPageControlsEnabled();

        // バージョン番号（なし / v1, v2… / v01, v02… / v001, v002…。初期値は「なし」）
        // ES3 で 'short' を裸キーにすると予約語エラーになるため、ローカル変数名は short_ にする
        var versionRow = addRadioRow(panel, 'label.version', 'tip.version', [
            { key: 'none', text: getLabel('radio.versionNone') },
            { key: 'short_', text: getLabel('radio.versionShort') },
            { key: 'padded', text: getLabel('radio.versionPadded') },
            { key: 'paddedWide', text: getLabel('radio.versionPaddedWide') }
        ]);
        // デフォルトは「なし」。元ファイルに v 番号がある場合は prefs を優先（無ければ「なし」）
        var hasOriginalVersion = !!getFirstSegmentValue(segments, 'version');
        var initialVersion = hasOriginalVersion
            ? pickPref(prefs, 'version', ['none', 'short', 'padded', 'paddedWide'], 'none')
            : 'none';
        versionRow.radios.none.value = (initialVersion === 'none');
        versionRow.radios.short_.value = (initialVersion === 'short');
        versionRow.radios.padded.value = (initialVersion === 'padded');
        versionRow.radios.paddedWide.value = (initialVersion === 'paddedWide');

        // 区切り記号（変更しない / - / _ 横並び。デフォルト "-"）
        var separatorRow = null;
        if (FEATURE_SEPARATOR) {
            separatorRow = addRadioRow(panel, 'label.separator', 'tip.separator', [
                { key: 'noChange', text: getLabel('radio.noChange') },
                { key: 'dash', text: '-' },
                { key: 'underscore', text: '_' }
            ]);
            // prefs があればそれを優先、無ければ FEATURE_SEPARATOR の指定文字（'-' or '_'）を初期値に
            var defaultSeparator = (FEATURE_SEPARATOR === '_') ? '_' : '-';
            var initialSeparator = pickPref(prefs, 'separator', ['', '-', '_'], defaultSeparator);
            separatorRow.radios.noChange.value = (initialSeparator === '');
            separatorRow.radios.dash.value = (initialSeparator === '-');
            separatorRow.radios.underscore.value = (initialSeparator === '_');
        }

        // 濁点処理（変更しない / 結合する。デフォルト "結合する"）
        var nfcRow = null;
        if (FEATURE_NFC) {
            nfcRow = addRadioRow(panel, 'label.nfc', 'tip.nfc', [
                { key: 'keep', text: getLabel('radio.noChange') },
                { key: 'combine', text: getLabel('radio.nfcCombine') }
            ]);
            var initialNfc = pickPref(prefs, 'nfc', ['keep', 'combine'], 'combine');
            nfcRow.radios.keep.value = (initialNfc === 'keep');
            nfcRow.radios.combine.value = (initialNfc === 'combine');
        }

        // クリーンなファイル名（削除する / -に変更 / _に変更。デフォルト "-に変更"）
        // OS 禁止文字（\/:*?"<>|）と絵文字・機種依存文字をまとめて処理。
        // 末尾に「半角カナ → 全角」チェックボックスを同居（'-' / '_' のときだけ有効）
        var cleanRow = null;
        var halfwidthKanaCheckbox = null;
        if (FEATURE_CLEAN) {
            cleanRow = addRadioRow(panel, 'label.clean', 'tip.clean', [
                { key: 'remove', text: getLabel('radio.cleanRemove') },
                { key: 'dash', text: getLabel('radio.changeToDash') },
                { key: 'underscore', text: getLabel('radio.changeToUnderscore') }
            ]);
            var initialClean = pickPref(prefs, 'clean', ['remove', 'dash', 'underscore'], 'dash');
            cleanRow.radios.remove.value = (initialClean === 'remove');
            cleanRow.radios.dash.value = (initialClean === 'dash');
            cleanRow.radios.underscore.value = (initialClean === 'underscore');

            if (FEATURE_HALFWIDTH_KANA) {
                halfwidthKanaCheckbox = cleanRow.group.add('checkbox', undefined, getLabel('radio.halfwidthKanaConvert'));
                halfwidthKanaCheckbox.helpTip = getLabel('tip.halfwidthKana');
                halfwidthKanaCheckbox.value = pickPref(prefs, 'halfwidthKana', ['keep', 'convert'], 'convert') === 'convert';
            }
        }

        // 法人略記・丸数字（変更しない / 削除する / 変換する。デフォルト "変換する"）
        var translitRow = null;
        if (FEATURE_TRANSLITERATE) {
            translitRow = addRadioRow(panel, 'label.translit', 'tip.translit', [
                { key: 'keep', text: getLabel('radio.noChange') },
                { key: 'remove', text: getLabel('radio.translitRemove') },
                { key: 'convert', text: getLabel('radio.translitConvert') }
            ]);
            var initialTranslit = pickPref(prefs, 'translit', ['keep', 'remove', 'convert'], 'convert');
            translitRow.radios.keep.value = (initialTranslit === 'keep');
            translitRow.radios.remove.value = (initialTranslit === 'remove');
            translitRow.radios.convert.value = (initialTranslit === 'convert');
        }

        /**
         * クリーンが '-' / '_' のときだけ「半角カナ → 全角」を有効化
         * @returns {void}
         */
        function syncHalfwidthKanaEnabled() {
            if (!halfwidthKanaCheckbox || !cleanRow) return;
            halfwidthKanaCheckbox.enabled = (cleanRow.radios.dash.value || cleanRow.radios.underscore.value);
        }
        syncHalfwidthKanaEnabled();

        // ラベル幅を統一（FEATURE で UI 非表示の行は除外）
        var labelTexts = [labelText('label.base'), labelText('label.title')];
        var labelControls = [baseLabel, titleRow.label];
        if (FEATURE_STATUS) { labelTexts.push(labelText('label.status')); labelControls.push(statusLabel); }
        labelTexts.push(labelText('label.timestamp')); labelControls.push(timestampRow.label);
        if (FEATURE_PAGE) { labelTexts.push(labelText('label.page')); labelControls.push(pageRow.label); }
        labelTexts.push(labelText('label.version')); labelControls.push(versionRow.label);
        if (FEATURE_SEPARATOR) { labelTexts.push(labelText('label.separator')); labelControls.push(separatorRow.label); }
        if (FEATURE_NFC) { labelTexts.push(labelText('label.nfc')); labelControls.push(nfcRow.label); }
        if (FEATURE_CLEAN) { labelTexts.push(labelText('label.clean')); labelControls.push(cleanRow.label); }
        if (FEATURE_TRANSLITERATE) { labelTexts.push(labelText('label.translit')); labelControls.push(translitRow.label); }
        // 個別整列はせず、呼び出し側に渡してファイル名パネルと統一整列する
        // titleFieldSpacer の幅は createDialog 側で整列後に設定する

        return {
            panel: panel,
            baseField: baseField,
            titleRow: titleRow,
            titleField: titleField,
            titleFieldSpacer: titleFieldSpacer,
            syncTitleFieldEnabled: syncTitleFieldEnabled,
            labels: labelControls,
            labelTexts: labelTexts,
            statusDropdown: statusDropdown,
            timestampRow: timestampRow,
            timestampHHMMCheckbox: timestampHHMMCheckbox,
            coerceTimestampForHHMM: coerceTimestampForHHMM,
            pageCheckbox: pageCheckbox,
            pagePrefixField: pagePrefixField,
            pagePadRadio2: pagePadRadio2,
            pagePadRadio3: pagePadRadio3,
            syncPageControlsEnabled: syncPageControlsEnabled,
            versionRow: versionRow,
            separatorRow: separatorRow,
            nfcRow: nfcRow,
            cleanRow: cleanRow,
            halfwidthKanaCheckbox: halfwidthKanaCheckbox,
            translitRow: translitRow,
            syncHalfwidthKanaEnabled: syncHalfwidthKanaEnabled,
            /* '' = 変更しない、'-' / '_' = 統一。FEATURE_SEPARATOR=false なら常に '' */
            getSeparator: function () {
                if (!separatorRow) return '';
                if (separatorRow.radios.noChange.value) return '';
                if (separatorRow.radios.dash.value) return '-';
                return '_';
            },
            /* 'keep' = そのまま、'combine' = 濁点・半濁点を NFC 結合。FEATURE_NFC=false なら常に 'keep' */
            getNfc: function () {
                if (!nfcRow) return 'keep';
                return nfcRow.radios.keep.value ? 'keep' : 'combine';
            },
            /* 'remove' / 'dash' / 'underscore'。FEATURE_CLEAN=false なら常に '' を返す（無処理） */
            getClean: function () {
                if (!cleanRow) return '';
                if (cleanRow.radios.dash.value) return 'dash';
                if (cleanRow.radios.underscore.value) return 'underscore';
                return 'remove';
            },
            /* 'keep' = そのまま、'remove' = 削除、'convert' = TRANSLITERATE_MAP で変換。FEATURE_TRANSLITERATE=false なら常に 'keep' */
            getTranslit: function () {
                if (!translitRow) return 'keep';
                if (translitRow.radios.convert.value) return 'convert';
                if (translitRow.radios.remove.value) return 'remove';
                return 'keep';
            },
            /* 'keep' = そのまま、'convert' = 半角カナを全角カナに。FEATURE_HALFWIDTH_KANA=false なら常に 'keep' */
            getHalfwidthKana: function () {
                if (!halfwidthKanaCheckbox) return 'keep';
                return halfwidthKanaCheckbox.value ? 'convert' : 'keep';
            },
            /* 'none' / 'date' / 'dateDash' */
            getTimestamp: function () {
                if (timestampRow.radios.none.value) return 'none';
                if (timestampRow.radios.dateDash.value) return 'dateDash';
                return 'date';
            },
            /* 'no' / 'hhmm'（時刻 HHMM をタイムスタンプ末尾に付与するか）*/
            getTimestampTime: function () {
                return timestampHHMMCheckbox.value ? 'hhmm' : 'no';
            },
            /* 'no' / 'yes'（ページ番号セグメントを付与するか）。FEATURE_PAGE=false なら常に 'no' */
            getPageEnable: function () {
                if (!pageCheckbox) return 'no';
                return pageCheckbox.value ? 'yes' : 'no';
            },
            /* '2' / '3'（連番のゼロ埋め桁数）。FEATURE_PAGE=false なら '2' */
            getPagePad: function () {
                if (!pagePadRadio3) return '2';
                return pagePadRadio3.value ? '3' : '2';
            },
            /* 連番のプレフィックス（例: "page"）。FEATURE_PAGE=false なら '' */
            getPagePrefix: function () {
                if (!pagePrefixField) return '';
                return pagePrefixField.text;
            },
            /* STATUS_ITEMS の value（'' = なし）。FEATURE_STATUS=false なら常に '' */
            getStatus: function () {
                if (!FEATURE_STATUS || !statusDropdown) return '';
                var selectedIndex = statusDropdown.selection ? statusDropdown.selection.index : 0;
                var selectedItem = STATUS_ITEMS[selectedIndex];
                if (!selectedItem || isStatusDivider(selectedItem)) return '';
                return selectedItem.value;
            },
            /* 'none' / 'short' / 'padded' / 'paddedWide' */
            getVersion: function () {
                if (versionRow.radios.none.value) return 'none';
                if (versionRow.radios.short_.value) return 'short';
                if (versionRow.radios.padded.value) return 'padded';
                return 'paddedWide';
            },
            /* 'none' / 'parent' / 'grandparent' / 'custom' */
            getTitleMode: function () {
                if (titleRow.radios.none.value) return 'none';
                if (titleRow.radios.parent.value) return 'parent';
                if (titleRow.radios.grandparent.value) return 'grandparent';
                return 'custom';
            },
            /* ベース入力欄の現在値 / Current value of the base input */
            getBase: function () {
                return baseField.text;
            }
        };
    }

    /**
     * 並び順を ↑↓ で編集するサブダイアログ。OK で新しい順序を返す。キャンセルで null
     * Sub-dialog to reorder segments with ↑↓; returns the new order, or null on cancel
     * @param {array} initialOrder 編集前の並び
     * @returns {array} 編集後の並び（キャンセル時は null）
     */
    function openSortDialog(initialOrder) {
        var sortDialog = new Window('dialog', getLabel('sort.title'));
        setupWindow(sortDialog);

        var hintText = sortDialog.add('statictext', undefined, getLabel('sort.hint'));
        hintText.alignment = 'left';

        var bodyGroup = sortDialog.add('group');
        bodyGroup.orientation = 'row';
        bodyGroup.alignChildren = ['fill', 'fill'];

        var order = initialOrder.slice();
        var orderList = bodyGroup.add('listbox', undefined, []);
        orderList.preferredSize = [200, 140];
        for (var i = 0; i < order.length; i++) {
            orderList.add('item', getLabel('label.' + order[i]));
        }
        orderList.selection = 0;

        var buttonColumn = bodyGroup.add('group');
        buttonColumn.orientation = 'column';
        buttonColumn.alignChildren = 'fill';
        var upButton = buttonColumn.add('button', undefined, '↑');
        upButton.preferredSize = [36, 24];
        var downButton = buttonColumn.add('button', undefined, '↓');
        downButton.preferredSize = [36, 24];

        /**
         * リストボックスを現在の並びで作り直し、指定位置を選択する
         * @param {number} newIndex 選択し直す位置
         * @returns {void}
         */
        function refreshList(newIndex) {
            for (var i = 0; i < order.length; i++) {
                orderList.items[i].text = getLabel('label.' + order[i]);
            }
            orderList.selection = newIndex;
        }
        upButton.onClick = function () {
            var selectedIndex = orderList.selection ? orderList.selection.index : -1;
            if (selectedIndex <= 0) return;
            var swapTemp = order[selectedIndex - 1];
            order[selectedIndex - 1] = order[selectedIndex];
            order[selectedIndex] = swapTemp;
            refreshList(selectedIndex - 1);
        };
        downButton.onClick = function () {
            var selectedIndex = orderList.selection ? orderList.selection.index : -1;
            if (selectedIndex < 0 || selectedIndex >= order.length - 1) return;
            var swapTemp = order[selectedIndex + 1];
            order[selectedIndex + 1] = order[selectedIndex];
            order[selectedIndex] = swapTemp;
            refreshList(selectedIndex + 1);
        };

        var buttonRow = addButtonRow(sortDialog);
        var btnCancel = buttonRow.rightGroup.add('button', undefined, getLabel('button.cancel'), { name: 'cancel' });
        var btnOK = buttonRow.rightGroup.add('button', undefined, 'OK', { name: 'ok' });
        alignRightOnlyButtonRow(buttonRow);

        prepareDialogWindow(sortDialog, SCRIPT_NAME + "_sort");
        if (sortDialog.show() !== 1) return null;
        return order;
    }

    /**
     * ES3 互換: 配列に値が含まれるか / ES3-safe array contains check
     * @param {array} array 探す配列
     * @param {string} value 探す値
     * @returns {boolean} 含まれていれば true
     */
    function arrayContains(array, value) {
        for (var i = 0; i < array.length; i++) {
            if (array[i] === value) return true;
        }
        return false;
    }

    /**
     * prefs[key] が validValues 内ならそれを、そうでなければ fallback を返す
     * Returns prefs[key] if it's in validValues, else fallback
     * @param {object} prefs 保存しておいた設定
     * @param {string} key 取り出すキー
     * @param {array} validValues 許可する値の一覧
     * @param {string} fallback 一覧に無いときの既定値
     * @returns {string} 採用する値
     */
    function pickPref(prefs, key, validValues, fallback) {
        var value = prefs && prefs[key];
        return arrayContains(validValues, value) ? value : fallback;
    }

    /**
     * 保存された並び順文字列を妥当性チェックして配列で返す。不正なら null
     * Parse a comma-separated segment order; returns null if invalid
     * @param {string} value カンマ区切りの並び順文字列
     * @returns {array} 並び順の配列（不正なら null）
     */
    function parseSegmentOrderPref(value) {
        if (!value) return null;
        var parts = String(value).split(',');
        if (parts.length !== SEGMENT_ORDER.length) return null;
        var seen = {};
        for (var i = 0; i < parts.length; i++) {
            var kind = parts[i];
            if (!arrayContains(SEGMENT_ORDER, kind) || seen[kind]) return null;
            seen[kind] = true;
        }
        return parts;
    }

    /**
     * labelTexts の最大幅に controls の幅を揃える。checkbox/radiobutton が含まれていれば
     * インジケーター分（+20px）を全コントロールに加算して右端を揃える
     * Align controls' widths to the widest label. If any control is a checkbox/radiobutton,
     * add indicator width (+20px) to all so their right edges align
     * @param {Panel} panel 幅の計測に使うパネル
     * @param {array} labelTexts ラベル文字列の配列
     * @param {array} controls 幅を揃えるコントロールの配列
     * @returns {void}
     */
    function alignLabelWidths(panel, labelTexts, controls) {
        var graphics = panel.graphics;
        var maxWidth = 0;
        for (var i = 0; i < labelTexts.length; i++) {
            var width = graphics.measureString(labelTexts[i]).width;
            if (width > maxWidth) maxWidth = width;
        }
        maxWidth += 12;
        var needsIndicator = false;
        for (var k = 0; k < controls.length; k++) {
            if (controls[k].type === 'checkbox' || controls[k].type === 'radiobutton') {
                needsIndicator = true;
                break;
            }
        }
        if (needsIndicator) maxWidth += 20;
        for (var j = 0; j < controls.length; j++) {
            controls[j].preferredSize = [maxWidth, controls[j].preferredSize.height || 20];
        }
    }

    /**
     * ラベル + ラジオ群を 1 行追加して { group, label, radios: { key: radio, ... } } を返す。
     * radioDefs: [{ key, text }, ...]。すべてのコントロールに getLabel(tipKey) の helpTip を設定
     * Add a label + radio row; returns { group, label, radios }
     * @param {Panel} panel 追加先のパネル
     * @param {string} labelKey ラベルのキー
     * @param {string} tipKey ヘルプチップのキー
     * @param {array} radioDefs {key, text} を要素とするラジオ定義
     * @returns {object} {group, label, radios}
     */
    function addRadioRow(panel, labelKey, tipKey, radioDefs) {
        var row = panel.add('group');
        row.orientation = 'row';
        var tipText = getLabel(tipKey);
        // ラベルは右揃え（alignLabelWidths で固定列幅になるため右端が入力欄側に揃う）
        var label = row.add('statictext', undefined, labelText(labelKey), { justify: 'right' });
        label.helpTip = tipText;
        var radios = {};
        for (var i = 0; i < radioDefs.length; i++) {
            var radioDef = radioDefs[i];
            var radio = row.add('radiobutton', undefined, radioDef.text);
            radio.helpTip = tipText;
            radios[radioDef.key] = radio;
        }
        return { group: row, label: label, radios: radios };
    }

    /**
     * 各コントロールの種類に応じた変更イベントに callback を割り当てる。null/undefined はスキップ。
     * radiobutton/checkbox → onClick、edittext → onChanging、dropdownlist → onChange
     * Wire a refresh callback to each control's appropriate change event
     * @param {function} callback 変更時に呼ぶコールバック
     * @param {array} controls 対象のコントロール配列（null は読み飛ばす）
     * @returns {void}
     */
    function wireRefresh(callback, controls) {
        for (var i = 0; i < controls.length; i++) {
            var control = controls[i];
            if (!control) continue;
            var controlType = control.type;
            if (controlType === 'radiobutton' || controlType === 'checkbox') control.onClick = callback;
            else if (controlType === 'edittext') control.onChanging = callback;
            else if (controlType === 'dropdownlist') control.onChange = callback;
        }
    }

    /**
     * ダイアログ全体を組み立て、イベント配線とプレビューを行う / Compose the full dialog, wire events, and run live preview
     * @param {array} segments セグメント配列
     * @param {string} currentName 現在のファイル名
     * @param {object} prefs 保存しておいた設定
     * @param {string} parentFolderName 親フォルダー名
     * @param {string} grandparentFolderName 2 階層上のフォルダー名
     * @param {Folder} folder 保存先フォルダー
     * @param {boolean} isNative 書類が .ai ネイティブか
     * @returns {object} {dialog, getUIState, getMode}
     */
    function createDialog(segments, currentName, prefs, parentFolderName, grandparentFolderName, folder, isNative) {
        var dialog = new Window('dialog', getLabel('dialog.title') + ' ' + SCRIPT_VERSION);
        setupWindow(dialog);

        // 上部: 動作パネル + モードパネル + ソートパネル（FEATURE_SORT=false ならソートパネル無し）
        var topRow = dialog.add('group');
        topRow.orientation = 'row';
        topRow.alignChildren = ['fill', 'top'];
        var mode = buildModePanel(topRow, isNative);
        // 元ファイル名が「タイムスタンプ無し」かつ「バージョン番号あり（-vN）」なら、初期モードを「バージョンのみ」に
        var defaultVersionOnly = !hasSegmentKind(segments, 'date') && hasSegmentKind(segments, 'version');
        var opMode = buildOpModePanel(topRow, defaultVersionOnly);
        // 現在のファイル名から検出した並び順（要素ゼロなら「現在に準じる」は無効化）
        var currentOrder = FEATURE_SORT ? deriveOrderFromSegments(segments) : [];
        var sort = FEATURE_SORT ? buildSortPanel(topRow, currentOrder.length > 0) : null;

        // 並び順カスタム値（prefs に保存されたものを採用、不正・未保存ならデフォルト）
        var customOrder = FEATURE_SORT
            ? (parseSegmentOrderPref(prefs && prefs.segmentOrder) || SEGMENT_ORDER.slice())
            : SEGMENT_ORDER.slice();

        var filename = buildFilenamePanel(dialog, currentName);
        var options = buildOptionsPanel(dialog, segments, prefs, parentFolderName, grandparentFolderName);

        // 2 つのパネルのラベル幅を統一整列（測定基準は options.panel）
        var combinedTexts = filename.labelTexts.concat(options.labelTexts);
        var combinedLabels = filename.labels.concat(options.labels);
        alignLabelWidths(options.panel, combinedTexts, combinedLabels);
        // サブテキスト 2 行目「指定」入力欄の左余白を統一後のラベル列幅に合わせる
        options.titleFieldSpacer.preferredSize = [options.titleRow.label.preferredSize.width, 1];

        // ---- ライブプレビュー ----

        /**
         * ダイアログ上の各コントロールから、名前の組み立てに必要な状態をまとめて取り出す
         * @returns {object} 現在の UI 状態
         */
        function currentUIState() {
            var sortMode = (FEATURE_SORT && sort) ? sort.getSortMode() : 'off';
            var segmentOrder;
            // 'current' は検出順のみだと連番/未検出ステータスが欠落するため正規位置に補完する
            if (sortMode === 'on') segmentOrder = customOrder;
            else if (sortMode === 'current') segmentOrder = fillMissingSegmentOrder(currentOrder);
            else segmentOrder = SEGMENT_ORDER;
            return {
                opMode: opMode.getOpMode(),
                baseText: options.getBase(),
                titleMode: options.getTitleMode(),
                titleText: options.titleField.text,
                parentFolderName: parentFolderName,
                grandparentFolderName: grandparentFolderName,
                status: options.getStatus(),
                timestamp: options.getTimestamp(),
                timestampTime: options.getTimestampTime(),
                pageEnable: options.getPageEnable(),
                pagePad: options.getPagePad(),
                pagePrefix: options.getPagePrefix(),
                version: options.getVersion(),
                separator: options.getSeparator(),
                nfc: options.getNfc(),
                translit: options.getTranslit(),
                clean: options.getClean(),
                halfwidthKana: options.getHalfwidthKana(),
                sort: sortMode,
                customSegmentOrder: customOrder.slice(),
                segmentOrder: segmentOrder
            };
        }

        /**
         * 「バージョン番号のみ」モード時に隠す UI（ソート + ファイル名の設定）
         * @returns {void}
         */
        function syncOpModeVisibility() {
            var versionOnly = opMode.isVersionOnly();
            if (sort) sort.panel.visible = !versionOnly;
            options.panel.visible = !versionOnly;
            dialog.layout.layout(true);
            dialog.layout.resize();
        }

        /**
         * 現在の入力内容でファイル名のプレビューを更新する
         * @returns {void}
         */
        function refreshPreviews() {
            options.syncTitleFieldEnabled();
            options.syncHalfwidthKanaEnabled();
            options.syncPageControlsEnabled();
            // 「バージョンのみ」モードでは UI 整形を一切かけず、元ファイル名の v 番号だけ更新
            if (opMode.isVersionOnly()) {
                var versionOnlyBase = bumpVersionInPlace(stripExtension(currentName));
                versionOnlyBase = nextAvailableVersionName(versionOnlyBase, folder, '.ai');
                filename.finalNameValue.text = versionOnlyBase + '.ai';
                return;
            }
            var uiState = currentUIState();
            var finalBase = applyNameTransforms(buildFinalName(segments, uiState), uiState, folder, '.ai');
            filename.finalNameValue.text = finalBase ? (finalBase + '.ai') : getLabel('message.emptyName');
        }

        opMode.versionOnlyRadio.onClick = function () {
            syncOpModeVisibility();
            refreshPreviews();
        };
        opMode.fullRadio.onClick = function () {
            syncOpModeVisibility();
            refreshPreviews();
        };

        // ファイル名設定パネル内のすべての入力（無効化中の FEATURE は row が null で来るので skip される）
        var titleRow = options.titleRow, timestampRow = options.timestampRow, versionRow = options.versionRow;
        var separatorRow = options.separatorRow, nfcRow = options.nfcRow, cleanRow = options.cleanRow, translitRow = options.translitRow;

        /**
         * 構成要素（ベース／サブテキスト／ステータス／タイムスタンプ／バージョン）を変更したら、
         * 「現在のファイル名に準じる」は前提が崩れるので「標準順」に降格させる
         * @returns {void}
         */
        function demoteSortToDefault() {
            if (!sort) return;
            if (!sort.sortCurrentRadio.value) return;
            sort.selectSortMode('off');
        }
        /**
         * 構成要素を変える操作用。並び順を降格させてからプレビューを更新する
         * @returns {void}
         */
        function refreshAndDemoteSort() {
            demoteSortToDefault();
            refreshPreviews();
        }

        // 構成要素を変える操作（「現在のファイル名に準じる」は解除）
        wireRefresh(refreshAndDemoteSort, [
            titleRow.radios.none, titleRow.radios.parent, titleRow.radios.grandparent, titleRow.radios.custom,
            options.baseField, options.titleField,
            options.statusDropdown,
            timestampRow.radios.none, timestampRow.radios.date, timestampRow.radios.dateDash,
            options.pageCheckbox, options.pagePrefixField, options.pagePadRadio2, options.pagePadRadio3,
            versionRow.radios.none, versionRow.radios.short_, versionRow.radios.padded, versionRow.radios.paddedWide
        ]);
        // 「時刻も付与」だけは、ON にしたときタイムスタンプを YYYYMMDD に引き上げてから更新する
        options.timestampHHMMCheckbox.onClick = function () {
            options.coerceTimestampForHHMM();
            refreshAndDemoteSort();
        };

        // 整形のみ変える操作（並び順には影響しないので「現在のファイル名に準じる」を維持）
        wireRefresh(refreshPreviews, [
            separatorRow && separatorRow.radios.noChange, separatorRow && separatorRow.radios.dash, separatorRow && separatorRow.radios.underscore,
            nfcRow && nfcRow.radios.keep, nfcRow && nfcRow.radios.combine,
            cleanRow && cleanRow.radios.remove, cleanRow && cleanRow.radios.dash, cleanRow && cleanRow.radios.underscore,
            options.halfwidthKanaCheckbox,
            translitRow && translitRow.radios.keep, translitRow && translitRow.radios.remove, translitRow && translitRow.radios.convert
        ]);

        // ソートパネルの ON/OFF とサブダイアログ起動（FEATURE_SORT のとき）
        if (FEATURE_SORT && sort) {
            var onSortToggle = function (mode) {
                return function () {
                    sort.selectSortMode(mode);
                    refreshPreviews();
                };
            };
            sort.sortOffRadio.onClick = onSortToggle('off');
            sort.sortCurrentRadio.onClick = onSortToggle('current');
            sort.sortOnRadio.onClick = onSortToggle('on');
            sort.sortButton.onClick = function () {
                var newOrder = openSortDialog(customOrder);
                if (newOrder) {
                    customOrder = newOrder;
                    refreshPreviews();
                }
            };
        }

        refreshPreviews();

        // ---- ボタン（右寄せ Cancel / OK） ----
        var buttonRow = addButtonRow(dialog);
        var btnCancel = buttonRow.rightGroup.add('button', undefined, getLabel('button.cancel'), { name: 'cancel' });
        var btnOK = buttonRow.rightGroup.add('button', undefined, 'OK', { name: 'ok' });
        alignRightOnlyButtonRow(buttonRow);

        // 初期モードが「バージョンのみ」のときは該当 UI を隠した状態で表示
        if (opMode.isVersionOnly()) syncOpModeVisibility();

        return {
            dialog: dialog,
            getUIState: currentUIState,
            getMode: mode.getMode
        };
    }

    // =========================================
    // メイン / Main
    // =========================================

    /**
     * エントリポイント。ダイアログを開き、選択モードに応じた出力を実行 / Entry point: open the dialog and execute the selected mode
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel('message.noDoc'));
            return;
        }

        var doc = app.activeDocument;
        var documentInfo = gatherDocumentInfo(doc);
        var segments = mergeFragmentedText(parseFileName(documentInfo.baseName));
        var prefs = loadPrefs();

        // 保存先はダイアログ表示前に確定する。未保存ドキュメントはここで保存先フォルダを選択。
        // 同じフォルダをプレビュー（連番/バージョンの繰り上げ）と実処理の両方で使い、両者を一致させる
        var targetFolder = ensureTargetFolder(documentInfo.folder);
        if (!targetFolder) return; // キャンセル

        var dialogUI = createDialog(segments, documentInfo.currentName, prefs, documentInfo.parentFolderName, documentInfo.grandparentFolderName, targetFolder, isNativeDocument(documentInfo));
        prepareDialogWindow(dialogUI.dialog, SCRIPT_NAME);
        if (dialogUI.dialog.show() !== 1) return; // キャンセル

        // プレビュー中はフォルダー列挙をキャッシュしている。実保存の採番は取り直した一覧で行う
        resetFolderFileNamesCache();

        var uiState = dialogUI.getUIState();
        var newBaseName = (uiState.opMode === 'versionOnly')
            ? bumpVersionInPlace(documentInfo.baseName)
            : buildFinalName(segments, uiState);

        if (uiState.opMode === 'versionOnly') {
            // 「バージョンのみ」モードでは UI 整形をスキップし、元ファイル名の書式を尊重したまま v 番号だけ繰り上げ
            newBaseName = nextAvailableVersionName(newBaseName, targetFolder, '.ai');
        } else {
            newBaseName = applyNameTransforms(newBaseName, uiState, targetFolder, '.ai');
        }

        // 整形前だけでなく整形後も判定する。記号だけ・絵文字だけの名前は
        // clean / collapse を通ると空になり、そのままでは「.ai」という不可視ファイルになる
        if (!newBaseName) {
            alert(getLabel('message.emptyName'));
            return;
        }

        var destFile = File(targetFolder.fsName + '/' + newBaseName + '.ai');

        // 長さチェック（拡張子込み）: 上限超過なら確認ダイアログを出して続行可
        var fullByteLength = byteLengthUTF8(newBaseName + '.ai');
        if (fullByteLength > FEATURE_MAX_FILENAME_BYTES) {
            var warningMessage = getLabel('message.confirmTooLong')
                .replace('{bytes}', String(fullByteLength))
                .replace('{limit}', String(FEATURE_MAX_FILENAME_BYTES));
            if (!confirm(warningMessage + '\n\n' + newBaseName + '.ai')) return;
        }

        if (!confirmOverwriteIfExists(destFile, documentInfo.fsPath)) return;

        try {
            executeOutput(doc, destFile, dialogUI.getMode(), documentInfo.fsPath);
            // 成功したら今回の選択をプリセットとして保存（versionOnly モードでは保存しない）
            if (uiState.opMode !== 'versionOnly') {
                var prefsToSave = {
                    titleMode: uiState.titleMode,
                    timestamp: uiState.timestamp,
                    timestampTime: uiState.timestampTime,
                    version: uiState.version
                };
                if (FEATURE_PAGE) {
                    prefsToSave.pageEnable = uiState.pageEnable;
                    prefsToSave.pagePad = uiState.pagePad;
                    prefsToSave.pagePrefix = uiState.pagePrefix;
                }
                if (FEATURE_STATUS) prefsToSave.status = uiState.status;
                if (FEATURE_SEPARATOR) prefsToSave.separator = uiState.separator;
                if (FEATURE_NFC) prefsToSave.nfc = uiState.nfc;
                if (FEATURE_CLEAN) prefsToSave.clean = uiState.clean;
                if (FEATURE_HALFWIDTH_KANA && FEATURE_CLEAN) prefsToSave.halfwidthKana = uiState.halfwidthKana;
                if (FEATURE_TRANSLITERATE) prefsToSave.translit = uiState.translit;
                if (FEATURE_SORT) {
                    // sort モード自体は復元しない（毎回「現在のファイル名に準じる」を初期値）
                    prefsToSave.segmentOrder = uiState.customSegmentOrder.join(',');
                }
                savePrefs(prefsToSave);
            }
        } catch (e) {
            alert(getLabel('message.saveFailed') + '\n' + e);
        }
    }

    main();

})();
