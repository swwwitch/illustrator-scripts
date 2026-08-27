#target illustrator

/*

### 概要

あらかじめ登録した複数のフォルダーから .ai/.svg ファイルをキーワードで絞り込み、選んだファイルをその場で開くファインダーです。
フォルダーとファイル名を左右のリストに分けて表示し、一度作った索引をキャッシュして次回以降の起動を早くします。

詳細は README を参照してください。

### Overview

A finder that filters .ai/.svg files across several registered folders by keyword and opens the selected file on the spot.
Folders and file names are shown in two side-by-side lists, and the index is cached so later launches start quickly.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AiFileFinder";                 /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-08-27";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-08-28";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AiFileFinder.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AiFileFinder.md
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/xxxxxxxx"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

/**
 * @discussion 参考 / Reference
 * 「Finderで表示」の仕組み（Automatorアプリとの連携）
 * 自分用メモ (@mute_racoon3631)「真の「Finderで表示」をイラレでも」
 * https://note.com/mute_racoon3631/n/n9e0e08f5d5f7
 */

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* 一覧に載せる拡張子と、絞り込みチェックボックスの初期状態。チェックボックスもこの順に並ぶ
       / Extensions to list and whether each checkbox starts checked; the checkboxes follow this order */
    var FILE_EXTENSIONS = [
        { ext: "ai",  isChecked: true },
        { ext: "svg", isChecked: false }
    ];

    /* 走査に使う拡張子の判定。FILE_EXTENSIONS から組み立てる / The scan test, built from FILE_EXTENSIONS */
    var FILE_EXT_RE = makeFileExtRegExp(FILE_EXTENSIONS);

    /* 検索フォルダーの初期値。環境設定で追加・削除できる / Default search folders, editable in the preferences */
    var SEARCH_FOLDER_DEFAULTS = [
        "~/sw Dropbox/takano masahiro/Dropbox-shared/DTPTransit_neta2026",
        "~/sw Dropbox/takano masahiro/Dropbox-shared/DTPTransit_neta2024",
        "~/sw Dropbox/takano masahiro/Dropbox-shared/DTPTransit_neta2021",
        "~/sw Dropbox/takano masahiro/Dropbox-shared/DTPTransit_neta2019",
        "~/sw Dropbox/takano masahiro/Dropbox-shared/DTPTransit_neta2017",
        "~/sw Dropbox/takano masahiro/Dropbox-shared/DTPTransit_neta",
        "~/sw Dropbox/takano masahiro/Dropbox-shared/AT-doc"
    ];

    /* キーワードボタンに並べる語。環境設定で編集できる / Keyword buttons, editable in the preferences */
    var KEYWORD_PRESET_DEFAULTS = ["icon", "logo", "font", "keyboard", "アイコン", "ロゴ", "カラー", "フォント", "アップデート", "ツール", "パネル", "線"];

    /* 一覧から外す語。ファイル名かフォルダー名に含むものを落とす / Files whose name or folder contains one of these are hidden */
    var EXCLUDE_KEYWORD_DEFAULTS = ["note-cover-", "backup", "_old", "_outlined", "test", "copy", "名称未設定"];

    /* 索引キャッシュを作り直すまでの時間。0 にすると毎回スキャンする / Hours before the cached index is rebuilt */
    var CACHE_MAX_AGE_HOURS = 24;

    /* キーワードの判定の初期状態。true でAND（すべて含む）、false でOR（どれか含む）
       / Initial keyword match, AND when true */
    var MATCH_ALL_DEFAULT = true;

    /* 並び順の初期状態。true で更新日、false で名前 / Initial sort key, by modified date when true */
    var SORT_BY_MODIFIED_DEFAULT = true;

    /* 並び順の向きの初期状態。true で降順（更新日なら新しい順） / Initial direction, descending when true */
    var SORT_DESCENDING_DEFAULT = true;

    /* 一覧のファイル名に拡張子を付けるか。種類は拡張子のチェックボックスで絞れるので既定では省く
       / Whether the file list shows the extension; the checkboxes already filter by type */
    var SHOW_FILE_EXTENSION = false;

    /* ファイル名リストに一度に並べる上限。これを超えた分はキーワードで絞り込む / Maximum rows in the file list */
    var FILE_LIST_MAX_ITEMS = 300;

    /* Finder表示に使うAutomatorアプリと、パスを受け渡す一時ファイル / Automator app used to reveal a file */
    /* 一時ファイル名はアプリ内のAppleScriptが読む固定名。アプリを旧名 IllustratorRevealLink.app から
       改名した名残で綴りが揃っていないが、アプリ側を直すまでここは変えない
       / The temp file name is hard-coded in the app; leave it until the app itself is updated */
    var REVEAL_APP_PATH  = "/Applications/RevealInFinder.app";
    var REVEAL_PATH_FILE = "/tmp/illustrator_reveal_path.txt";

    // =========================================
    // 設定の保存 / Stored settings
    // =========================================

    /* Illustratorの環境設定に保存する。再起動しても残り、余分なファイルを作らない / Stored in Illustrator preferences */
    var PREF_KEY_FOLDERS  = "AiFileFinder.searchFolders";
    var PREF_KEY_KEYWORDS = "AiFileFinder.keywordPresets";
    var PREF_KEY_EXCLUDES = "AiFileFinder.excludeKeywords";

    /* 一覧はどれも1つの文字列にまとめて保存する。改行はIllustratorの設定ファイルを壊しかねないのでタブで区切る
       / Lists are stored as one joined string; tabs avoid putting newlines into the preferences file */
    var SETTING_LIST_SEPARATOR = "\t";

    /* 保存済みの目印。これが無ければ未設定とみなし、初期値に戻す
       / A stored list always starts with this tag, so "cleared" is not mistaken for "never set" */
    var SETTING_LIST_TAG = "v1\t";

    /**
     * 文字列を検索フォルダーの配列に変換する
     * @param {Array<string>} pathList - フォルダーのパス
     * @returns {Array<Folder>} 実在するフォルダーだけを並び順のまま返す
     */
    function toExistingFolders(pathList) {
        var folders = [];
        for (var i = 0; i < pathList.length; i++) {
            var path = trimWhitespace(pathList[i]);
            if (path === "") continue;

            var folder = new Folder(path);
            if (folder.exists) folders.push(folder);
        }
        return folders;
    }

    /**
     * 検索フォルダーを読み出す
     * @returns {Array<Folder>} 記録がなければ初期値のフォルダー
     */
    function readSearchFolders() {
        var savedPaths = trimWhitespace(app.preferences.getStringPreference(PREF_KEY_FOLDERS));
        if (savedPaths === "") return toExistingFolders(SEARCH_FOLDER_DEFAULTS);
        return toExistingFolders(savedPaths.split(SETTING_LIST_SEPARATOR));
    }

    /**
     * 検索フォルダーを次回起動用に記録する
     * @param {Array<Folder>} folders - 記録するフォルダー
     * @returns {void}
     */
    function saveSearchFolders(folders) {
        var paths = [];
        for (var i = 0; i < folders.length; i++) paths.push(folders[i].fsName);
        app.preferences.setStringPreference(PREF_KEY_FOLDERS, paths.join(SETTING_LIST_SEPARATOR));
    }

    /**
     * 改行またはタブ区切りの文字列を語の配列に変換する
     * 前後の空白を落とし、空行と重複は取り除く
     * @param {string} storedText - 入力欄の文字列、または保存してある文字列
     * @returns {Array<string>} 語の並び
     */
    function toWordList(storedText) {
        var wordLines = String(storedText).split(/[\t\r\n]+/);
        var words = [];
        var seenWords = {};

        for (var i = 0; i < wordLines.length; i++) {
            var word = trimWhitespace(wordLines[i]);
            if (word === "") continue;

            /* ハッシュのキー衝突を避けるため接頭辞を付けて既出判定する / Prefix the key to avoid collisions */
            var wordKey = "#" + word;
            if (seenWords[wordKey]) continue;
            seenWords[wordKey] = true;
            words.push(word);
        }
        return words;
    }

    /**
     * 記録済みの語の一覧を読み出す
     * @param {string} prefKey - 環境設定のキー
     * @param {Array<string>} defaultWords - 未設定のときに返す初期値
     * @returns {Array<string>} 記録がなければ初期値
     */
    function readWordList(prefKey, defaultWords) {
        var savedValue = String(app.preferences.getStringPreference(prefKey) || "");

        /* 目印が無いのは未設定。全部消した設定を初期値で上書きしないための判定 / No tag means "never stored" */
        if (savedValue.indexOf(SETTING_LIST_TAG) !== 0) return defaultWords.slice(0);
        return toWordList(savedValue.substring(SETTING_LIST_TAG.length));
    }

    /**
     * 語の一覧を記録する
     * @param {string} prefKey - 環境設定のキー
     * @param {Array<string>} words - 記録する語
     * @returns {void}
     */
    function saveWordList(prefKey, words) {
        app.preferences.setStringPreference(prefKey, SETTING_LIST_TAG + words.join(SETTING_LIST_SEPARATOR));
    }

    // =========================================
    // レイアウト / Layout
    // =========================================

    /* ウィンドウ・パネルの余白と間隔 / Window & panel margins and spacing */
    var WINDOW_MARGINS = 16;                 /* ウィンドウ外周の余白 / window margin */
    var WINDOW_SPACING = 12;                 /* ウィンドウ内の要素間隔 / window spacing */
    var PANEL_MARGINS  = [16, 20, 16, 12];   /* パネル余白 [左,上,右,下] / panel margins */
    var PANEL_SPACING  = 12;                 /* パネル内の要素間隔 / panel spacing */
    var COLUMN_SPACING = 12;                 /* 2カラムの間隔 / gap between columns */

    /* 密なパネル・行の間隔 / Spacing for dense panels and rows */
    var DENSE_SPACING = 8;                   /* パネル内を詰めるときの間隔 / dense panel spacing */

    /* リストの寸法 / List sizes */
    var FOLDER_LIST_SIZE = [170, 360];       /* フォルダーリストの寸法 [幅,高さ] / folder list size */
    var FILE_LIST_HEIGHT = 360;              /* ファイル名リストの高さ。幅は列幅から決まる / file list height; the width comes from the columns */

    /* 列幅の合計はリスト幅からスクロールバーぶんを引いた値に合わせる。余ると空の列に見える
       / Columns fill the list minus the scrollbar; leftover space reads as an empty column */
    var LIST_SCROLLBAR_WIDTH = 20;           /* 列幅から差し引くスクロールバーの幅 / scrollbar width */
    var FOLDER_COLUMN_WIDTH = FOLDER_LIST_SIZE[0] - LIST_SCROLLBAR_WIDTH;  /* フォルダー列の幅 / folder column width */
    var NAME_COLUMN_WIDTH   = 230;           /* ファイル名列の幅 / file name column width */
    var DATE_COLUMN_WIDTH   = 230;           /* 更新日列の幅 / modified date column width */

    /* ファイル名リストの寸法 [幅,高さ]。左右のリスト幅がダイアログ幅を決める / file list size */
    var FILE_LIST_SIZE = [NAME_COLUMN_WIDTH + DATE_COLUMN_WIDTH + LIST_SCROLLBAR_WIDTH, FILE_LIST_HEIGHT];

    /* ボタンの寸法と余白 / Button sizes and margins */
    var BUTTON_HEIGHT        = 28;           /* ボタンの高さ / button height */
    var DIALOG_BUTTON_WIDTH  = 92;           /* 開く・キャンセルの幅 / dialog button width */
    var WIDE_BUTTON_WIDTH    = 110;          /* 文言の長いボタンの幅（Finderで表示・初期値に戻す）/ wide button width */
    var SETTINGS_BUTTON_WIDTH = 92;          /* 環境設定内のボタンの幅 / preferences button width */
    var BUTTON_ROW_TOP_MARGIN = 10;          /* ボタン列の上余白 / top margin above the button row */
    var ROW_TOP_MARGIN        = 8;           /* リスト下の行の上余白 / top margin above a row under the lists */
    var PRESET_TOP_MARGIN     = 5;           /* キーワードボタンの上余白 / top margin above the preset buttons */

    /* ダイアログの中身の幅。左右のリストから決まる / Content width, driven by the two lists */
    var CONTENT_WIDTH    = FOLDER_LIST_SIZE[0] + COLUMN_SPACING + FILE_LIST_SIZE[0];
    var PRESET_ROW_WIDTH = CONTENT_WIDTH - PANEL_MARGINS[0] - PANEL_MARGINS[2];

    /* キーワードボタンは小ぶりにする / Keyword preset buttons are smaller */
    var PRESET_BUTTON_HEIGHT  = 22;          /* キーワードボタンの高さ / preset button height */
    var PRESET_BUTTON_PADDING = 16;          /* 文字幅に足す左右の余白 / horizontal padding */
    var PRESET_CHAR_WIDTH     = 7;           /* 実測できない環境用の半角1文字の概算幅 / fallback char width */

    /* 環境設定ダイアログ / Preferences dialog */
    var SETTINGS_LIST_SIZE    = [440, 180];  /* 検索フォルダーリストの寸法 [幅,高さ] / search folder list size */
    var SETTINGS_COLUMN_WIDTH = 214;         /* 2カラムに割ったパネルの幅 / width of one of the two columns */
    var SETTINGS_KEYWORD_ROWS = 9;           /* キーワード入力欄の行数 / rows in the keyword field */
    var SETTINGS_KEYWORD_ROW_HEIGHT = 18;    /* 1行の高さ / height of one row */
    var SETTINGS_KEYWORD_SIZE = [182, SETTINGS_KEYWORD_ROWS * SETTINGS_KEYWORD_ROW_HEIGHT];  /* キーワード入力欄の寸法 [幅,高さ] / keyword field size */

    /* 年別・並び順・期間 / Year, sort order and period */
    var YEAR_DROPDOWN_WIDTH = 74;            /* 年別ドロップダウンの幅 / year dropdown width */
    var PERIOD_YEAR_WIDTH   = 62;            /* 期間の年ポップアップの幅。5文字ぶん / period year dropdown width */
    var PERIOD_MONTH_WIDTH  = 62;            /* 期間の月ポップアップの幅 / period month dropdown width */
    var PERIOD_LEFT_SPACE   = 28;            /* 年別と期間の間の余白。全角2文字ぶん / gap before the period filter */
    var PERIOD_PAIR_SPACING = 4;             /* 年と月の間隔 / gap between year and month */

    /* 昇順・降順ボタン / Sort order buttons */
    var SORT_ORDER_BUTTON_SIZE = [32, 22];   /* ボタンの寸法 [幅,高さ] / button size */
    var SORT_ORDER_BUTTON_SPACING = 2;       /* ▲と▼の間隔 / gap between the two buttons */

    /* キーワード欄のクリアボタン（自前描画） / Clear button next to the keyword field (custom drawn) */
    var CLEAR_BUTTON_SIZE  = 20;             /* クリアボタンの一辺 / clear button size */
    var CLEAR_CIRCLE_INSET = 2;              /* 円と外周の間隔 / inset of the circle */
    var CLEAR_GLYPH_INSET  = 6;              /* ×と外周の間隔 / inset of the × glyph */
    var CLEAR_STROKE_WIDTH = 1.5;            /* 円と×の線幅 / stroke width */

    /* 進行状況ウィンドウ / Progress window */
    var PROGRESS_BAR_SIZE  = [320, 12];      /* プログレスバーの寸法 [幅,高さ] / progress bar size */
    var PROGRESS_TEXT_WIDTH = 320;           /* 走査中のフォルダー名の幅 / progress caption width */

    /**
     * ウィンドウの共通設定を適用する
     * @param {Window} win - 対象ウィンドウ
     * @param {number} [spacing] - 要素間隔。省略時は WINDOW_SPACING
     * @returns {void}
     */
    function setupWindow(win, spacing) {
        win.orientation = "column";
        win.alignChildren = "fill";
        win.margins = WINDOW_MARGINS;
        win.spacing = (typeof spacing === "number") ? spacing : WINDOW_SPACING;
    }

    /**
     * パネルの共通設定を適用する
     * @param {Panel} panel - 対象パネル
     * @param {number} [spacing] - 要素間隔。省略時は PANEL_SPACING
     * @returns {void}
     */
    function setupPanel(panel, spacing) {
        panel.orientation = "column";
        panel.alignChildren = ["fill", "top"];
        panel.alignment = "fill";
        panel.margins = PANEL_MARGINS;
        panel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /**
     * 行グループの共通設定を適用する（ボタン列など）
     * @param {Group} group - 対象グループ
     * @param {string} [alignment] - 配置。省略時は "left"
     * @param {number} [spacing] - 要素間隔。省略時は PANEL_SPACING
     * @returns {void}
     */
    function setupRow(group, alignment, spacing) {
        group.orientation = "row";
        group.alignment = [alignment || "left", "center"];
        group.alignChildren = ["left", "center"];
        group.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /**
     * ボタンの寸法をそろえる
     * @param {Button} button - 対象ボタン
     * @param {number} width - ボタンの幅
     * @returns {void}
     */
    function applyButtonSize(button, width) {
        button.preferredSize = [width, BUTTON_HEIGHT];
        button.minimumSize = [width, BUTTON_HEIGHT];
    }

    /**
     * 伸縮するスペーサーを足す
     * @param {Group} parent - 追加先のグループ
     * @returns {Group} スペーサー
     */
    function addSpacer(parent) {
        /* スペーサー（伸縮）/ Spacer (stretchable) */
        var spacer = parent.add("group");
        spacer.alignment = ["fill", "fill"];
        spacer.minimumSize.width = 0;
        return spacer;
    }

    /**
     * ダイアログ下部の、左右に分けたボタン列を作る
     * @param {Window} parent - 追加先のウィンドウ
     * @returns {{left: Group, right: Group}} 左寄せと右寄せのグループ
     */
    function addButtonRow(parent) {
        /* メイングループ（横並び） / Main group (horizontal layout) */
        var btnRowGroup = parent.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.margins = [0, BUTTON_ROW_TOP_MARGIN, 0, 0];
        btnRowGroup.alignment = ["fill", "bottom"];

        /* 左側グループ / Left-side button group */
        var btnLeftGroup = btnRowGroup.add("group");
        btnLeftGroup.alignChildren = ["left", "center"];

        addSpacer(btnRowGroup);

        /* 右側グループ / Right-side button group */
        var btnRightGroup = btnRowGroup.add("group");
        btnRightGroup.alignChildren = ["right", "center"];

        return { left: btnLeftGroup, right: btnRightGroup };
    }

    /**
     * ボタンに並べた文字の幅を概算する
     * @param {string} text - 対象の文字列
     * @returns {number} 概算の幅（px）
     */
    function estimateTextWidth(text) {
        var width = 0;
        for (var i = 0; i < text.length; i++) {
            /* 日本語は半角のおよそ2倍の幅で見積もる / Assume double width for non-ASCII */
            width += (text.charCodeAt(i) > 0xFF) ? PRESET_CHAR_WIDTH * 2 : PRESET_CHAR_WIDTH;
        }
        return width;
    }

    /**
     * 部品に並べた文字の幅を測る
     * @param {Object} control - 文字を持つ部品
     * @param {string} text - 測る文字列
     * @returns {number} 文字の幅（px）。実測できない環境では文字数からの概算
     */
    function measureTextWidth(control, text) {
        var textWidth = estimateTextWidth(text);
        try {
            var measured = control.graphics.measureString(text);
            var measuredWidth = (measured.width !== undefined) ? measured.width : measured[0];
            /* 環境によっては値が取れずNaNになる。その場合は概算のままにする / Keep the estimate if unusable */
            if (!isNaN(measuredWidth) && measuredWidth > 0) textWidth = measuredWidth;
        } catch (e) {}
        return textWidth;
    }

    /**
     * キーワードボタンの寸法を文字幅に合わせて詰める
     * @param {Button} button - 対象ボタン
     * @returns {void}
     */
    function applyPresetButtonSize(button) {
        var textWidth = measureTextWidth(button, button.text);

        var buttonWidth = Math.ceil(textWidth) + PRESET_BUTTON_PADDING;
        button.preferredSize = [buttonWidth, PRESET_BUTTON_HEIGHT];
        button.minimumSize = [buttonWidth, PRESET_BUTTON_HEIGHT];
    }

    /**
     * IllustratorのUIが明るい設定かどうかを返す
     * @returns {boolean} 明るいUIなら true
     */
    function isLightUI() {
        try {
            return app.preferences.getRealPreference("uiBrightness") > 0.5;
        } catch (e) {
            return false;
        }
    }

    /**
     * クリアボタンの線の色を決める
     * @param {Button} clearButton - 対象ボタン
     * @returns {Array<number>} RGBAの並び
     */
    function pickClearGlyphColor(clearButton) {
        var useLightUI = isLightUI();

        /* キーワードが空のときは無効にしてあるので、いちばん淡くする / Dimmest while disabled */
        if (clearButton.enabled === false) return useLightUI ? [0.78, 0.78, 0.78, 1] : [0.42, 0.42, 0.42, 1];

        /* 押している間は濃くして手応えを出す / Darker while pressed */
        if (clearButton.pressed) return useLightUI ? [0.20, 0.20, 0.20, 1] : [0.95, 0.95, 0.95, 1];

        /* 主役ではないので、通常時は少し薄めに描く / Slightly muted at rest */
        return useLightUI ? [0.45, 0.45, 0.45, 1] : [0.72, 0.72, 0.72, 1];
    }

    /**
     * クリアボタンの丸と×を描く
     * @param {Button} clearButton - 対象ボタン
     * @returns {void}
     */
    function drawClearButton(clearButton) {
        var graphics = clearButton.graphics;
        var glyphColor = pickClearGlyphColor(clearButton);

        /* 最初の onDraw では size が取れないことがある。0 のまま描くと何も出ない
           / size can be unavailable on the first onDraw; drawing at zero would paint nothing */
        var buttonSize = (clearButton.size && clearButton.size[0]) ? clearButton.size[0] : CLEAR_BUTTON_SIZE;

        /* 前回の描画を消すため、まず親と同じ色で塗りつぶす
           backgroundColor が取れない環境では塗りを飛ばす。ここで投げると円も×も描かれない
           / Repaint the background first; skip it where backgroundColor is unavailable,
             since throwing here would leave the whole button blank */
        try {
            graphics.newPath();
            graphics.rectPath(0, 0, buttonSize, buttonSize);
            graphics.fillPath(graphics.backgroundColor);
        } catch (e) {}

        var glyphPen = graphics.newPen(graphics.PenType.SOLID_COLOR, glyphColor, CLEAR_STROKE_WIDTH);
        var circleSize = buttonSize - CLEAR_CIRCLE_INSET * 2;
        graphics.newPath();
        graphics.ellipsePath(CLEAR_CIRCLE_INSET, CLEAR_CIRCLE_INSET, circleSize, circleSize);
        graphics.strokePath(glyphPen);

        var glyphStart = CLEAR_GLYPH_INSET;
        var glyphEnd = buttonSize - CLEAR_GLYPH_INSET;
        graphics.newPath();
        graphics.moveTo(glyphStart, glyphStart);
        graphics.lineTo(glyphEnd, glyphEnd);
        graphics.strokePath(glyphPen);
        graphics.newPath();
        graphics.moveTo(glyphEnd, glyphStart);
        graphics.lineTo(glyphStart, glyphEnd);
        graphics.strokePath(glyphPen);
    }

    /**
     * 昇順・降順ボタンに出す三角を返す
     * 選んでいる側は塗りつぶし、選んでいない側は白抜きにする
     * @param {boolean} isDescending - 降順なら true
     * @param {boolean} isSelected - 選んでいるなら true
     * @returns {string} ▲▽などの三角1文字
     */
    function getSortOrderGlyph(isDescending, isSelected) {
        if (isDescending) return isSelected ? "▼" : "▽";
        return isSelected ? "▲" : "△";
    }

    /**
     * 昇順・降順のボタンを1つ作る
     * 自前描画にするとリストを組み直したときに絵が消えるので、文字で三角を出す
     * @param {Group} parent - 追加先のグループ
     * @param {boolean} isDescending - 降順（▼）なら true
     * @param {boolean} isSelected - 選んだ状態で始めるなら true
     * @returns {Button} 三角を1文字だけ載せたボタン
     */
    function addSortOrderButton(parent, isDescending, isSelected) {
        var orderButton = parent.add("button", undefined, getSortOrderGlyph(isDescending, isSelected));
        orderButton.helpTip = getLabel(isDescending ? LABELS.button.descending : LABELS.button.ascending);
        orderButton.preferredSize = SORT_ORDER_BUTTON_SIZE;
        orderButton.minimumSize = SORT_ORDER_BUTTON_SIZE;
        orderButton.isDescending = isDescending;
        orderButton.selected = isSelected;
        return orderButton;
    }

    /**
     * キーワードを消すクリアボタンを作る
     * @param {Group} parent - 追加先のグループ
     * @returns {Button} 丸に×を自前描画したボタン
     */
    function addClearButton(parent) {
        var clearButton = parent.add("button", undefined, "");
        clearButton.helpTip = getLabel(LABELS.button.clearKeyword);

        /* 余った幅は入力欄が取るが、端数が残ることがある。右端に貼り付けて隙間を作らない
           / The field takes the slack, but a remainder can linger; pin the button to the right edge */
        clearButton.alignment = ["right", "center"];
        clearButton.preferredSize = [CLEAR_BUTTON_SIZE, CLEAR_BUTTON_SIZE];
        clearButton.minimumSize = [CLEAR_BUTTON_SIZE, CLEAR_BUTTON_SIZE];
        clearButton.maximumSize = [CLEAR_BUTTON_SIZE, CLEAR_BUTTON_SIZE];
        clearButton.pressed = false;

        /* キーワード欄は空の状態で開くので、ディム表示から始める / Start dimmed for the empty field */
        clearButton.enabled = false;

        clearButton.onDraw = function () {
            drawClearButton(this);
        };
        clearButton.addEventListener("mousedown", function () {
            this.pressed = true;
            this.notify("onDraw");
        });
        clearButton.addEventListener("mouseup", function () {
            this.pressed = false;
            this.notify("onDraw");
        });
        /* ボタンの外でマウスを離したときも押下表示を戻す / Reset the pressed look on mouseout */
        clearButton.addEventListener("mouseout", function () {
            if (!this.pressed) return;
            this.pressed = false;
            this.notify("onDraw");
        });

        return clearButton;
    }

    // =========================================
    // ラベル定義 / Labels
    // =========================================

    /**
     * UI言語を判定する
     * @returns {string} "ja" または "en"
     */
    function getCurrentUILang() {
        return ($.locale && $.locale.indexOf("ja") === 0) ? "ja" : "en";
    }

    var uiLang = getCurrentUILang();

    var LABELS = {
        dialog: {
            title:        { ja: "ファイルファインダー", en: "File Finder" },
            preferences:  { ja: "環境設定", en: "Preferences" },
            selectFolder: { ja: "検索対象に加えるフォルダーを選択してください", en: "Select a folder to add to the search" },
            scanning:     { ja: "検索中", en: "Scanning" }
        },
        panel: {
            filter:        { ja: "絞り込み（%1件）", en: "Filter (%1)" },
            filterLimited: { ja: "絞り込み（%1件中 %2 件を表示）", en: "Filter (showing %2 of %1)" },
            searchFolders:  { ja: "検索フォルダー", en: "Search Folders" },
            keywordButtons: { ja: "キーワードボタン", en: "Keyword Buttons" },
            excludeRules:   { ja: "除外条件", en: "Exclusions" }
        },
        listCaption: {
            folder:   { ja: "フォルダー", en: "Folder" },
            fileName: { ja: "ファイル名", en: "File Name" },
            modified: { ja: "更新日", en: "Modified" }
        },
        folderRow: {
            allFolders: { ja: "（すべて）", en: "(All)" }
        },
        fieldLabel: {
            keyword:   { ja: "キーワード", en: "Keyword" },
            year:      { ja: "年別", en: "Year" },
            period:    { ja: "期間", en: "Period" },
            periodTo:  { ja: "〜", en: "–" },
            sortBy:    { ja: "並び順", en: "Sort by" }
        },
        hint: {
            onePerLine: { ja: "1行に1語", en: "One per line" },
            keywordButtons: {
                ja: "絞り込みパネルにボタンとして並びます。option＋クリックで語を足せます。",
                en: "Shown as buttons in the filter panel. Option-click adds the word to the keyword."
            },
            excludeRules: {
                ja: "ファイル名かフォルダー名にこの語を含むファイルを、一覧から外します。",
                en: "Files whose name or folder contains one of these words are hidden."
            },
            periodYear: {
                ja: "年を選ぶと期間で絞り込みます。「なし」でその側の制限を外します",
                en: "Pick a year to filter by period. \"None\" turns that end off"
            },
            resetSettings: {
                ja: "検索フォルダー・キーワードボタン・除外条件を、すべて初期値に戻します",
                en: "Reset the search folders, the keyword buttons, and the exclusions to their defaults"
            },
            periodMonth: {
                ja: "月。開始側はその月の1日から、終了側はその月の末日までを含みます",
                en: "Month. The start begins on day 1, the end runs through the last day"
            }
        },
        dropdown: {
            allYears: { ja: "すべて", en: "All" },
            noLimit:  { ja: "なし", en: "None" }
        },
        radio: {
            matchAll:       { ja: "AND", en: "AND" },
            matchAny:       { ja: "OR", en: "OR" },
            sortByModified: { ja: "更新日", en: "Modified" },
            sortByName:     { ja: "名前", en: "Name" }
        },
        button: {
            preferences:   { ja: "環境設定", en: "Preferences" },
            rescan:        { ja: "再スキャン", en: "Rescan" },
            ascending:     { ja: "昇順", en: "Ascending" },
            descending:    { ja: "降順", en: "Descending" },
            clearKeyword:  { ja: "キーワードをクリア", en: "Clear keyword" },
            appendKeyword: { ja: "option＋クリックで語を足す", en: "Option-click to add the word" },
            addFolder:     { ja: "追加", en: "Add" },
            removeFolder:  { ja: "削除", en: "Remove" },
            resetSettings: { ja: "初期値に戻す", en: "Reset" },
            cancel:        { ja: "キャンセル", en: "Cancel" },
            open:          { ja: "開く", en: "Open" },
            ok:            { ja: "OK", en: "OK" }
        },
        progress: {
            scanning: { ja: "%1 / %2 フォルダー", en: "%1 / %2 folders" }
        },
        alert: {
            noFiles: {
                ja: "検索フォルダーに対象のファイルが見つかりませんでした。",
                en: "No matching files were found in the search folders."
            },
            missingFile: {
                ja: "選択したファイルが見つかりません。索引が古い可能性があるので、再スキャンしてください。\n\n%1",
                en: "The selected file could not be found. The index may be out of date, so try rescanning.\n\n%1"
            },
            openFailed: {
                ja: "ファイルを開けませんでした。\n\n%1\n\nエラー: %2",
                en: "The file could not be opened.\n\n%1\n\nError: %2"
            }
        }
    };

    /**
     * ラベル定義から現在のUI言語の文字列を取り出す
     * @param {{ja: string, en: string}} labelSet - 言語別のラベル定義
     * @returns {string} 現在のUI言語の文字列
     */
    function getLabel(labelSet) {
        return labelSet[uiLang] || labelSet.en;
    }

    /**
     * 項目名にコロンを付ける（日本語は全角、英語は半角）
     * @param {{ja: string, en: string}} labelSet - 言語別のラベル定義
     * @returns {string} コロン付きの項目名
     */
    function labelText(labelSet) {
        return getLabel(labelSet) + (uiLang === "ja" ? "：" : ":");
    }

    /**
     * ラベル内のプレースホルダー（%1, %2 …）を値で置き換える
     * @param {string} template - プレースホルダーを含む文字列
     * @param {Array<string>} values - 差し込む値
     * @returns {string} 置き換え後の文字列
     */
    function formatLabel(template, values) {
        var text = template;
        for (var i = 0; i < values.length; i++) {
            text = text.split("%" + (i + 1)).join(String(values[i]));
        }
        return text;
    }

    // =========================================
    // 検索キー / Search keys
    // =========================================

    /**
     * ファイル1件分の情報
     * @typedef {object} FileEntry
     * @property {File} file - 対象ファイル
     * @property {string} fileName - ファイル名
     * @property {string} folderPath - 検索フォルダー名から始まる表示用のフォルダーパス
     * @property {number} rootIndex - 何番目の検索フォルダーに属するか
     * @property {number} modifiedTime - 更新日時（ミリ秒）
     * @property {number} modifiedYear - 更新年。更新日時が取れなかったものは 0
     * @property {string} sortKey - 名前順に並べるためのキー
     * @property {string} normalizedSearchText - 正規化済みの検索キー
     * @property {boolean} isExcluded - 除外条件に当たるなら true
     */

    /**
     * 前後の空白を取り除く
     * @param {string} value - 対象の文字列
     * @returns {string} 前後の空白を除いた文字列
     */
    function trimWhitespace(value) {
        return String(value).replace(/^\s+|\s+$/g, "");
    }

    /* 半角カナを全角カタカナへ置き換える並び。U+FF61 から順に対応する / Half-width kana in code point order */
    var KANA_HALFWIDTH_START = 0xFF61;
    var KANA_HALFWIDTH_TABLE = "。「」、・ヲァィゥェォャュョッーアイウエオカキクケコサシスセソタチツテトナニヌネノハヒフヘホマミムメモヤユヨラリルレロワン";
    var KANA_HALFWIDTH_END   = KANA_HALFWIDTH_START + KANA_HALFWIDTH_TABLE.length - 1;

    /**
     * カタカナを濁点・半濁点・小書きのない基音へ寄せる
     * 変換表は持たず、Unicodeの並び方から計算する
     * @param {number} code - カタカナのコードポイント
     * @returns {number} 基音のコードポイント。カタカナ以外はそのまま返す
     */
    function foldKatakana(code) {
        /* ヴ と ヷヸヹヺ は並びから外れるので個別に扱う / These sit outside the regular rows */
        if (code === 0x30F4) return 0x30A6;
        if (code >= 0x30F7 && code <= 0x30FA) return code - 8;

        /* カ〜ヂ、ツ〜ド は「清音・濁音」の2つ並び。ッ を挟んで並びが一度切れる / Pairs, broken by ッ */
        if (code >= 0x30AB && code <= 0x30C2 && (code - 0x30AB) % 2 === 1) code -= 1;
        else if (code >= 0x30C4 && code <= 0x30C9 && (code - 0x30C4) % 2 === 1) code -= 1;

        /* ハ〜ポ は「清音・濁音・半濁音」の3つ並び / The ha row runs in threes */
        else if (code >= 0x30CF && code <= 0x30DD) code -= (code - 0x30CF) % 3;

        /* 小書きは対応する大書きの1つ手前にある / Each small kana sits right before its large form */
        if (code === 0x30A1 || code === 0x30A3 || code === 0x30A5 || code === 0x30A7 || code === 0x30A9 ||
            code === 0x30C3 || code === 0x30E3 || code === 0x30E5 || code === 0x30E7 || code === 0x30EE) {
            code += 1;
        }
        return code;
    }

    /**
     * 比較用の検索キーへ変換する
     * 大文字小文字・全角半角・かなの種類・濁点や小書きの違いを無視して一致させる
     * @param {string} value - 変換前の文字列
     * @returns {string} 正規化した文字列
     */
    function normalizeSearchKey(value) {
        var source = String(value).toLowerCase();
        var normalized = "";

        for (var i = 0; i < source.length; i++) {
            var currentChar = source.charAt(i);
            if (/\s/.test(currentChar)) continue;

            var code = source.charCodeAt(i);

            /* 単独の濁点・半濁点は落とす。濁りは基音へ寄せるので不要 / Standalone voiced marks are dropped */
            if (code === 0x3099 || code === 0x309A || code === 0x309B || code === 0x309C ||
                code === 0xFF9E || code === 0xFF9F) continue;

            /* 全角の英数記号は半角へ / Full-width ASCII to half-width */
            if (code >= 0xFF01 && code <= 0xFF5E) {
                normalized += String.fromCharCode(code - 0xFEE0).toLowerCase();
                continue;
            }

            /* 半角カナは全角カタカナへ / Half-width kana to full-width katakana */
            if (code >= KANA_HALFWIDTH_START && code <= KANA_HALFWIDTH_END) {
                code = KANA_HALFWIDTH_TABLE.charCodeAt(code - KANA_HALFWIDTH_START);
            }

            /* ひらがなはカタカナへ寄せる / Hiragana to katakana */
            if (code >= 0x3041 && code <= 0x3096) code += 0x60;

            normalized += String.fromCharCode(foldKatakana(code));
        }
        return normalized;
    }

    /**
     * 検索語のどれかを含むかどうかを判定する（OR検索）
     * @param {string} searchTarget - 正規化済みの検索対象
     * @param {Array<string>} searchTerms - 正規化済みの検索語
     * @returns {boolean} どれかを含むなら true
     */
    function matchesAnyTerm(searchTarget, searchTerms) {
        for (var i = 0; i < searchTerms.length; i++) {
            if (searchTarget.indexOf(searchTerms[i]) !== -1) return true;
        }
        return false;
    }

    /**
     * キーワードに一致するかどうかを判定する
     * @param {string} searchTarget - 正規化済みの検索対象
     * @param {Array<string>} searchTerms - 正規化済みの検索語
     * @param {boolean} isMatchAll - true ですべて含む（AND）、false でどれか含む（OR）
     * @returns {boolean} 一致するなら true
     */
    function matchesKeyword(searchTarget, searchTerms, isMatchAll) {
        /* 語が無いときは絞り込まない。ORのままだと1件も残らなくなる / No terms means no filtering */
        if (searchTerms.length === 0) return true;
        if (!isMatchAll) return matchesAnyTerm(searchTarget, searchTerms);

        for (var i = 0; i < searchTerms.length; i++) {
            if (searchTarget.indexOf(searchTerms[i]) === -1) return false;
        }
        return true;
    }

    /**
     * 語の並びを比較用の検索語に変換する
     * @param {Array<string>} keywords - 変換する語
     * @returns {Array<string>} 正規化した検索語。空の語は含まない
     */
    function toNormalizedTerms(keywords) {
        var searchTerms = [];
        for (var i = 0; i < keywords.length; i++) {
            var term = normalizeSearchKey(keywords[i]);
            if (term !== "") searchTerms.push(term);
        }
        return searchTerms;
    }

    /**
     * 入力文字列を空白区切りの検索語に分解する
     * @param {string} value - キーワード欄の文字列
     * @returns {Array<string>} 正規化した検索語。空の語は含まない
     */
    function splitSearchTerms(value) {
        return toNormalizedTerms(String(value).split(/\s+/));
    }

    /**
     * 1桁の数値を2桁にそろえる
     * @param {number} value - 対象の数値
     * @returns {string} 2桁の文字列
     */
    function padZero(value) {
        return (value < 10 ? "0" : "") + value;
    }

    /**
     * その月の日数を求める
     * @param {number} year - 年
     * @param {number} month - 月（1〜12）
     * @returns {number} 末日の日付
     */
    function getDaysInMonth(year, month) {
        /* 翌月の0日はその月の末日になる / Day zero of the next month is the last day of this one */
        return new Date(year, month, 0).getDate();
    }

    /**
     * 期間の片側のポップアップを読み取って、境目の時刻にする
     * 年が「（指定なし）」のときは、その側の制限なしとして null を返す
     * @param {{yearDropdown: DropDownList, monthDropdown: DropDownList}} periodPair - 年と月のポップアップ
     * @param {boolean} isEndOfPeriod - 終了側なら true
     * @returns {number|null} 時刻（ミリ秒）。制限なしのときは null
     */
    function readPeriodBoundary(periodPair, isEndOfPeriod) {
        var year = periodPair.yearDropdown.selection ? periodPair.yearDropdown.selection.year : undefined;

        /* 年を選ばないうちは月だけ動かしても意味がないので、月は触れないようにしておく
           / The month is meaningless until a year is picked */
        periodPair.monthDropdown.enabled = (year !== undefined);
        if (year === undefined) return null;

        var month = periodPair.monthDropdown.selection ? periodPair.monthDropdown.selection.month : 1;

        /* 開始はその月の頭から、終了はその月の末日の終わりまでを含める / Whole months on both ends */
        return isEndOfPeriod
            ? new Date(year, month - 1, getDaysInMonth(year, month), 23, 59, 59, 999).getTime()
            : new Date(year, month - 1, 1, 0, 0, 0, 0).getTime();
    }

    /**
     * 一覧に出すファイル名を作る
     * @param {string} fileName - 拡張子付きのファイル名
     * @returns {string} 設定に応じて拡張子を落としたファイル名
     */
    function toDisplayName(fileName) {
        return SHOW_FILE_EXTENSION ? fileName : fileName.replace(/\.[^.]+$/, "");
    }

    /**
     * ファイル名から拡張子を取り出す
     * @param {string} fileName - 拡張子付きのファイル名
     * @returns {string} 小文字にした拡張子。拡張子が無いときは空文字
     */
    function toFileExtension(fileName) {
        var extensionMatch = /\.([^.]+)$/.exec(fileName);
        return extensionMatch ? extensionMatch[1].toLowerCase() : "";
    }

    /**
     * 対象の拡張子に当たるかを見る正規表現を組み立てる
     * チェックが外れていても索引には残すので、走査はすべての拡張子を通す
     * @param {Array<{ext: string, isChecked: boolean}>} fileExtensions - 対象の拡張子
     * @returns {RegExp} 拡張子の判定
     */
    function makeFileExtRegExp(fileExtensions) {
        var extensions = [];
        for (var i = 0; i < fileExtensions.length; i++) {
            extensions.push(fileExtensions[i].ext);
        }
        return new RegExp("\\.(" + extensions.join("|") + ")$", "i");
    }

    /**
     * 更新日時を表示用の文字列にする
     * @param {number} modifiedTime - 更新日時（ミリ秒）
     * @returns {string} "YYYY-MM-DD HH:MM" 形式の文字列。日時が無いときは空文字
     */
    function formatModifiedTime(modifiedTime) {
        if (!modifiedTime) return "";

        var date = new Date(modifiedTime);
        return date.getFullYear() + "-" + padZero(date.getMonth() + 1) + "-" + padZero(date.getDate()) +
            " " + padZero(date.getHours()) + ":" + padZero(date.getMinutes());
    }

    // =========================================
    // ファイルの収集 / File collection
    // =========================================

    /**
     * 検索フォルダーの情報
     * @typedef {object} RootFolderInfo
     * @property {Folder} folder - 検索フォルダー
     * @property {string} label - リストに出すフォルダー名
     * @property {string} path - フォルダーの絶対パス
     * @property {number} index - 何番目の検索フォルダーか
     */

    /**
     * 検索フォルダーの情報をまとめる
     * @param {Array<Folder>} folders - 検索フォルダー
     * @returns {Array<RootFolderInfo>} 収集とパス判定に使う情報
     */
    function makeRootFolderInfoList(folders) {
        var rootInfoList = [];
        for (var i = 0; i < folders.length; i++) {
            var folderPath = folders[i].fsName;
            rootInfoList.push({
                folder: folders[i],
                label: folderPath.replace(/^.*[\/\\]/, ""),
                path: folderPath,
                index: i
            });
        }
        return rootInfoList;
    }

    /**
     * パスが検索フォルダーの中にあるかどうかを判定する
     * 先頭一致だけでは DTPTransit_neta が DTPTransit_neta2024 にも一致してしまうため、区切りまで確かめる
     * @param {string} targetPath - 判定するパス
     * @param {string} folderPath - 検索フォルダーの絶対パス
     * @returns {boolean} フォルダー自身か、その中にあれば true
     */
    function isInsideFolder(targetPath, folderPath) {
        if (targetPath.indexOf(folderPath) !== 0) return false;

        var rest = targetPath.substring(folderPath.length);
        return rest === "" || rest.charAt(0) === "/" || rest.charAt(0) === "\\";
    }

    /**
     * ファイル1件分の情報を作る
     * @param {File} file - 対象ファイル
     * @param {number} modifiedTime - 更新日時（ミリ秒）
     * @param {RootFolderInfo} rootInfo - 属している検索フォルダー
     * @param {string} [cachedSearchKey] - 索引キャッシュに残しておいた検索キー。省略時は作り直す
     * @returns {FileEntry} 検索キーまで用意した1件分の情報
     */
    function makeFileEntry(file, modifiedTime, rootInfo, cachedSearchKey) {
        /* fsName は復号済みの絶対パスなので、そこから名前とフォルダーを切り出す / fsName is the decoded absolute path */
        var fullPath = file.fsName;
        var fileName = fullPath.replace(/^.*[\/\\]/, "");
        var parentPath = fullPath.replace(/[\/\\][^\/\\]*$/, "");

        var relativeFolder = "";
        if (isInsideFolder(parentPath, rootInfo.path)) {
            relativeFolder = parentPath.substring(rootInfo.path.length).replace(/^[\/\\]+/, "").split("\\").join("/");
        }

        var folderPath = (relativeFolder === "") ? rootInfo.label : rootInfo.label + "/" + relativeFolder;

        return {
            file: file,
            fileName: fileName,
            folderPath: folderPath,
            rootIndex: rootInfo.index,
            modifiedTime: modifiedTime,
            modifiedYear: modifiedTime ? new Date(modifiedTime).getFullYear() : 0,
            sortKey: (folderPath + "/" + fileName).toLowerCase(),
            extension: toFileExtension(fileName),
            normalizedSearchText: cachedSearchKey || normalizeSearchKey(folderPath + " " + fileName),
            isExcluded: false
        };
    }

    /**
     * フォルダーを再帰的にたどって対象ファイルを集める
     * @param {Folder} targetFolder - 走査するフォルダー
     * @param {RootFolderInfo} rootInfo - 相対パスの基準になる検索フォルダー
     * @param {Array<FileEntry>} collectedEntries - 収集結果の追加先
     * @returns {void}
     */
    function collectFilesInFolder(targetFolder, rootInfo, collectedEntries) {
        var childItems = null;
        try {
            childItems = targetFolder.getFiles();
        } catch (e) {}

        /* 読めないフォルダーは null が返る。ここで止めないと length を読んだ時点で落ちる / getFiles() can return null */
        if (!childItems) return;

        for (var i = 0; i < childItems.length; i++) {
            var childItem = childItems[i];

            if (childItem instanceof Folder) {
                /* エイリアスは循環の元になるのでたどらない / Aliases can loop back into an ancestor */
                if (!childItem.alias && !/^\./.test(childItem.name)) collectFilesInFolder(childItem, rootInfo, collectedEntries);
                continue;
            }

            if (!(childItem instanceof File) || !FILE_EXT_RE.test(childItem.name)) continue;

            /* 更新日時が取れないファイルもあるので、その場合は0として扱う / Fall back to 0 when unavailable */
            var modifiedDate = childItem.modified;
            collectedEntries.push(makeFileEntry(childItem, modifiedDate ? modifiedDate.getTime() : 0, rootInfo));
        }
    }

    /**
     * 除外条件に当たるファイルへ印を付ける
     * 打鍵のたびに全件を照合し直さないよう、除外条件が変わったときだけ数える
     * @param {Array<FileEntry>} fileEntries - 対象のファイル（破壊的に変更する）
     * @param {Array<string>} excludeTerms - 正規化済みの除外語
     * @returns {void}
     */
    function applyExcludeFlags(fileEntries, excludeTerms) {
        for (var i = 0; i < fileEntries.length; i++) {
            fileEntries[i].isExcluded = matchesAnyTerm(fileEntries[i].normalizedSearchText, excludeTerms);
        }
    }

    /**
     * 検索語に打ち消されずに残る除外語を返す
     * 除外語をキーワードに入れたときは、その語での除外をやめて対象に含める
     * @param {Array<string>} excludeTerms - 正規化済みの除外語
     * @param {Array<string>} searchTerms - 正規化済みの検索語
     * @returns {Array<string>} まだ効いている除外語
     */
    function collectActiveExcludeTerms(excludeTerms, searchTerms) {
        if (searchTerms.length === 0) return excludeTerms;

        var activeTerms = [];
        for (var i = 0; i < excludeTerms.length; i++) {
            var excludeTerm = excludeTerms[i];
            var isLifted = false;

            for (var j = 0; j < searchTerms.length; j++) {
                /* どちらかがどちらかを含めば、その除外語を狙って打たれたものとみなす
                   / Either containing the other means the term was typed to reach the excluded files */
                if (excludeTerm.indexOf(searchTerms[j]) !== -1 || searchTerms[j].indexOf(excludeTerm) !== -1) {
                    isLifted = true;
                    break;
                }
            }
            if (!isLifted) activeTerms.push(excludeTerm);
        }
        return activeTerms;
    }

    /**
     * 年別ドロップダウンに並べる年を求める
     * @param {Array<FileEntry>} fileEntries - 集計対象のファイル
     * @returns {Array<number>} 新しい年から順に並べた更新年
     */
    function collectModifiedYears(fileEntries) {
        var seenYears = {};
        var years = [];

        for (var i = 0; i < fileEntries.length; i++) {
            /* 更新日時が取れなかったものは年で絞り込めないので数えない / Skip entries without a date */
            var year = fileEntries[i].modifiedYear;
            if (!year) continue;

            /* ハッシュのキー衝突を避けるため接頭辞を付けて既出判定する / Prefix the key to avoid collisions */
            var yearKey = "#" + year;
            if (seenYears[yearKey]) continue;
            seenYears[yearKey] = true;
            years.push(year);
        }

        years.sort(function (yearA, yearB) { return yearB - yearA; });
        return years;
    }

    /**
     * ファイルを昇順に並べ替える
     * 降順は並びを逆からたどって表すので、ここでは向きを扱わない
     * @param {Array<FileEntry>} fileEntries - 並べ替える配列（破壊的に変更する）
     * @param {boolean} sortByModified - true で更新日順、false で名前順
     * @returns {void}
     */
    function sortFileEntries(fileEntries, sortByModified) {
        /* 更新日が同じものは名前で決める / Ties on the date fall back to the name */
        fileEntries.sort(function (entryA, entryB) {
            if (sortByModified && entryA.modifiedTime !== entryB.modifiedTime) {
                return entryA.modifiedTime < entryB.modifiedTime ? -1 : 1;
            }
            if (entryA.sortKey === entryB.sortKey) return 0;
            return entryA.sortKey < entryB.sortKey ? -1 : 1;
        });
    }

    // =========================================
    // 索引キャッシュ / Index cache
    // =========================================

    /* キャッシュファイルと、書式が変わったときに古い内容を捨てるための目印 / Cache file and its format tag */
    var CACHE_FILE_PATH  = Folder.userData + "/AiFileFinder-index.txt";
    var CACHE_FORMAT_TAG = "AiFileFinder-index/2";

    /**
     * ファイルやウィンドウを閉じる。閉じられなくても処理は続ける
     * @param {File|Window} target - 対象
     * @returns {void}
     */
    function closeQuietly(target) {
        try { target.close(); } catch (e) {}
    }

    /**
     * キャッシュの2行目に書く、検索フォルダーの並びを表す文字列を作る
     * @param {Array<Folder>} folders - 検索フォルダー
     * @returns {string} タブ区切りのパス
     */
    function makeFolderSignature(folders) {
        var paths = [];
        for (var i = 0; i < folders.length; i++) paths.push(folders[i].fsName);
        return paths.join("\t");
    }

    /**
     * 索引キャッシュのファイルを1行ずつ読み出す
     * 古すぎるものと読めないものは、無かったものとして扱う
     * @returns {Array<string>|null} 読めなければ null
     */
    function readCacheLines() {
        if (CACHE_MAX_AGE_HOURS <= 0) return null;

        var cacheFile = new File(CACHE_FILE_PATH);
        if (!cacheFile.exists) return null;

        /* 作ってから時間が経ったものは信用しない / Ignore an index that has gone stale */
        var cacheModified = cacheFile.modified;
        if (!cacheModified) return null;
        if (((new Date()).getTime() - cacheModified.getTime()) / 3600000 > CACHE_MAX_AGE_HOURS) return null;

        try {
            cacheFile.encoding = "UTF-8";
            if (!cacheFile.open("r")) return null;
            return cacheFile.read().split(/\r\n|\r|\n/);
        } catch (e) {
            return null;
        } finally {
            closeQuietly(cacheFile);
        }
    }

    /**
     * 索引キャッシュを読み出す
     * @param {Array<Folder>} folders - 現在の検索フォルダー
     * @param {Array<RootFolderInfo>} rootInfoList - 検索フォルダーの情報
     * @returns {Array<FileEntry>|null} 使えるキャッシュがなければ null
     */
    function readIndexCache(folders, rootInfoList) {
        var cacheLines = readCacheLines();

        /* 書式と検索フォルダーの並びが一致しないキャッシュは作り直す / Rebuild when the tag or folder list differs */
        if (!cacheLines || cacheLines.length < 2) return null;
        if (trimWhitespace(cacheLines[0]) !== CACHE_FORMAT_TAG) return null;
        if (cacheLines[1] !== makeFolderSignature(folders)) return null;

        var fileEntries = [];
        for (var i = 2; i < cacheLines.length; i++) {
            if (cacheLines[i] === "") continue;

            var cacheColumns = cacheLines[i].split("\t");
            if (cacheColumns.length < 2) continue;

            /* パスはURI表記で持つ。%や記号を含む名前でも File に戻したときに崩れない / URI notation round-trips safely */
            var cachedFile = new File(cacheColumns[0]);
            var cachedFilePath = cachedFile.fsName;

            for (var j = 0; j < rootInfoList.length; j++) {
                if (!isInsideFolder(cachedFilePath, rootInfoList[j].path)) continue;
                fileEntries.push(makeFileEntry(cachedFile, Number(cacheColumns[1]) || 0, rootInfoList[j], cacheColumns[2]));
                break;
            }
        }
        return fileEntries;
    }

    /**
     * 索引キャッシュを書き出す
     * @param {Array<Folder>} folders - 現在の検索フォルダー
     * @param {Array<FileEntry>} fileEntries - 書き出すファイル
     * @returns {boolean} 書き出せたら true
     */
    function writeIndexCache(folders, fileEntries) {
        var cacheFile = new File(CACHE_FILE_PATH);
        try {
            cacheFile.encoding = "UTF-8";
            cacheFile.lineFeed = "Unix";
            if (!cacheFile.open("w")) return false;

            cacheFile.writeln(CACHE_FORMAT_TAG);
            cacheFile.writeln(makeFolderSignature(folders));
            for (var i = 0; i < fileEntries.length; i++) {
                /* 検索キーも一緒に残す。全件を正規化し直すと、キャッシュから開いても待たされる
                   / Keep the search key too, so a cached start does not re-normalize every name */
                cacheFile.writeln(fileEntries[i].file.fullName + "\t" + fileEntries[i].modifiedTime +
                    "\t" + fileEntries[i].normalizedSearchText);
            }
            return true;
        } catch (e) {
            return false;
        } finally {
            closeQuietly(cacheFile);
        }
    }

    /**
     * 走査中の進行状況を出すパレットを作る
     * @param {number} totalCount - 走査するフォルダーの数
     * @returns {{update: function(number, string): void, close: function(): void}} 更新と後始末の関数
     */
    function createScanProgress(totalCount) {
        /* 0 のままだとプログレスバーの値域が作れない / A zero range would be invalid */
        if (totalCount < 1) totalCount = 1;

        var progressWindow = null;
        var progressBar = null;
        var progressText = null;

        try {
            progressWindow = new Window("palette", getLabel(LABELS.dialog.scanning));
            setupWindow(progressWindow, DENSE_SPACING);

            progressText = progressWindow.add("statictext", undefined, "", { truncate: "middle" });
            progressText.preferredSize.width = PROGRESS_TEXT_WIDTH;

            /* 生成時の引数だけでは値域の解釈が環境で揺れるので、作ってから明示する / Set the range explicitly */
            progressBar = progressWindow.add("progressbar", undefined, 0, totalCount);
            progressBar.minvalue = 0;
            progressBar.maxvalue = totalCount;
            progressBar.value = 0;
            progressBar.preferredSize = PROGRESS_BAR_SIZE;

            progressWindow.center();
            progressWindow.show();
        } catch (e) {
            progressWindow = null;
        }

        return {
            /**
             * 進行状況を書き換える
             * @param {number} doneCount - 走査し終えたフォルダー数
             * @param {string} folderLabel - いま走査しているフォルダー名
             * @returns {void}
             */
            update: function (doneCount, folderLabel) {
                if (!progressWindow) return;

                progressText.text = folderLabel + "  " +
                    formatLabel(getLabel(LABELS.progress.scanning), [doneCount, totalCount]);
                progressBar.value = doneCount;
                progressWindow.update();
            },
            /**
             * パレットを閉じる
             * @returns {void}
             */
            close: function () {
                if (!progressWindow) return;

                closeQuietly(progressWindow);
                progressWindow = null;
            }
        };
    }

    /**
     * 検索フォルダーのファイルを集める。使えるキャッシュがあればそれを使う
     * @param {Array<Folder>} folders - 検索フォルダー
     * @param {boolean} forceScan - true でキャッシュを無視して走査する
     * @returns {Array<FileEntry>} 更新日の古い順に並べたファイル
     */
    function loadFileEntries(folders, forceScan) {
        var rootInfoList = makeRootFolderInfoList(folders);

        if (!forceScan) {
            var cachedEntries = readIndexCache(folders, rootInfoList);
            if (cachedEntries && cachedEntries.length > 0) {
                sortFileEntries(cachedEntries, true);
                return cachedEntries;
            }
        }

        var fileEntries = [];
        var progress = createScanProgress(rootInfoList.length);
        try {
            for (var i = 0; i < rootInfoList.length; i++) {
                progress.update(i, rootInfoList[i].label);
                collectFilesInFolder(rootInfoList[i].folder, rootInfoList[i], fileEntries);
            }
            progress.update(rootInfoList.length, "");
        } finally {
            progress.close();
        }

        writeIndexCache(folders, fileEntries);
        sortFileEntries(fileEntries, true);
        return fileEntries;
    }

    // =========================================
    // 環境設定ダイアログ / Preferences dialog
    // =========================================

    /**
     * 検索フォルダーのパネルを組み立てる
     * 一覧の中身はこのパネルの中で持ち、外へは読み出しと初期化の窓口だけを返す
     * @param {Window} parent - 追加先のウィンドウ
     * @param {Array<Folder>} initialFolders - 最初に表示するフォルダー
     * @returns {{getFolders: function(): Array<Folder>, reset: function(): void}} 編集結果の窓口
     */
    function buildSearchFolderPanel(parent, initialFolders) {
        var editedFolders = initialFolders.slice(0);

        var folderPanel = parent.add("panel", undefined, getLabel(LABELS.panel.searchFolders));
        setupPanel(folderPanel, DENSE_SPACING);

        var folderListBox = folderPanel.add("listbox", undefined, [], { multiselect: true });
        folderListBox.preferredSize = SETTINGS_LIST_SIZE;

        var folderButtonRow = folderPanel.add("group");
        setupRow(folderButtonRow, "left", DENSE_SPACING);
        var btnAddFolder = folderButtonRow.add("button", undefined, getLabel(LABELS.button.addFolder));
        var btnRemoveFolder = folderButtonRow.add("button", undefined, getLabel(LABELS.button.removeFolder));
        applyButtonSize(btnAddFolder, SETTINGS_BUTTON_WIDTH);
        applyButtonSize(btnRemoveFolder, SETTINGS_BUTTON_WIDTH);

        /**
         * フォルダーリストを今の内容に合わせて組み直す
         * @returns {void}
         */
        function refreshFolderListBox() {
            folderListBox.removeAll();
            for (var i = 0; i < editedFolders.length; i++) {
                folderListBox.add("item", editedFolders[i].fsName);
            }
            btnRemoveFolder.enabled = editedFolders.length > 0;
        }

        /**
         * すでに登録済みのフォルダーかどうかを判定する
         * @param {Folder} folder - 判定するフォルダー
         * @returns {boolean} 登録済みなら true
         */
        function isRegisteredFolder(folder) {
            for (var i = 0; i < editedFolders.length; i++) {
                if (editedFolders[i].fsName === folder.fsName) return true;
            }
            return false;
        }

        btnAddFolder.onClick = function () {
            var startFolder = (editedFolders.length > 0) ? editedFolders[editedFolders.length - 1] : Folder.myDocuments;
            var pickedFolder = startFolder.selectDlg(getLabel(LABELS.dialog.selectFolder));
            if (!pickedFolder || isRegisteredFolder(pickedFolder)) return;

            editedFolders.push(pickedFolder);
            refreshFolderListBox();
        };

        btnRemoveFolder.onClick = function () {
            var selectedItems = folderListBox.selection;
            if (!selectedItems) return;

            /* 後ろから消さないと、消したぶんだけ以降の添字がずれる / Remove from the end so the indexes stay valid */
            var selectedIndexes = [];
            for (var i = 0; i < selectedItems.length; i++) selectedIndexes.push(selectedItems[i].index);
            selectedIndexes.sort(function (indexA, indexB) { return indexB - indexA; });
            for (var j = 0; j < selectedIndexes.length; j++) editedFolders.splice(selectedIndexes[j], 1);

            refreshFolderListBox();
        };

        refreshFolderListBox();

        return {
            /**
             * いま並んでいるフォルダーを返す
             * @returns {Array<Folder>} 編集後のフォルダー
             */
            getFolders: function () {
                return editedFolders;
            },
            /**
             * 初期値のフォルダーに戻す
             * @returns {void}
             */
            reset: function () {
                editedFolders = toExistingFolders(SEARCH_FOLDER_DEFAULTS);
                refreshFolderListBox();
            }
        };
    }

    /**
     * 「1行に1語」で入力するパネルを1つ作る
     * @param {Group} parent - 追加先のグループ
     * @param {{ja: string, en: string}} titleSet - パネルの見出し
     * @param {{ja: string, en: string}} helpSet - 入力欄のツールチップ
     * @param {Array<string>} words - 初期表示する語
     * @returns {EditText} 追加した複数行の入力欄
     */
    function addWordListPanel(parent, titleSet, helpSet, words) {
        var wordPanel = parent.add("panel", undefined, getLabel(titleSet));
        setupPanel(wordPanel, DENSE_SPACING);

        /* 見出しの長さでカラム幅がずれないよう、幅は決め打ちにする / Fix the width so the columns stay even */
        wordPanel.preferredSize.width = SETTINGS_COLUMN_WIDTH;
        wordPanel.add("statictext", undefined, getLabel(LABELS.hint.onePerLine));

        var wordInput = wordPanel.add("edittext", undefined, words.join("\n"), { multiline: true });
        wordInput.preferredSize = SETTINGS_KEYWORD_SIZE;
        wordInput.helpTip = getLabel(helpSet);
        return wordInput;
    }

    /**
     * 検索フォルダー・キーワードボタン・除外条件を編集するダイアログを表示する
     * @param {Array<Folder>} currentFolders - 現在の検索フォルダー
     * @param {Array<string>} currentKeywords - 現在のキーワード
     * @param {Array<string>} currentExcludes - 現在の除外条件
     * @returns {{folders: Array<Folder>, keywords: Array<string>, excludes: Array<string>, rescan: boolean}|null} 編集後の設定。取り消し時は null
     */
    function showPreferencesDialog(currentFolders, currentKeywords, currentExcludes) {
        var settingsDialog = new Window("dialog", getLabel(LABELS.dialog.preferences));
        setupWindow(settingsDialog, DENSE_SPACING);

        var folderPanelUI = buildSearchFolderPanel(settingsDialog, currentFolders);

        /* キーワードボタンと除外条件は、どちらも「1行に1語」の同じ形なので横に並べる
           / Both lists take one word per line, so they sit side by side */
        var wordPanelRow = settingsDialog.add("group");
        setupRow(wordPanelRow, "fill", COLUMN_SPACING);
        wordPanelRow.alignChildren = ["fill", "fill"];

        var keywordListInput = addWordListPanel(wordPanelRow, LABELS.panel.keywordButtons, LABELS.hint.keywordButtons, currentKeywords);
        var excludeListInput = addWordListPanel(wordPanelRow, LABELS.panel.excludeRules, LABELS.hint.excludeRules, currentExcludes);

        var buttonRow = addButtonRow(settingsDialog);
        var btnRescan = buttonRow.left.add("button", undefined, getLabel(LABELS.button.rescan));
        var btnReset = buttonRow.left.add("button", undefined, getLabel(LABELS.button.resetSettings));
        btnReset.helpTip = getLabel(LABELS.hint.resetSettings);

        var btnCancel = buttonRow.right.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });
        var btnOk = buttonRow.right.add("button", undefined, getLabel(LABELS.button.ok), { name: "ok" });

        applyButtonSize(btnRescan, DIALOG_BUTTON_WIDTH);
        applyButtonSize(btnReset, WIDE_BUTTON_WIDTH);
        applyButtonSize(btnCancel, DIALOG_BUTTON_WIDTH);
        applyButtonSize(btnOk, DIALOG_BUTTON_WIDTH);

        /* 3つの設定をまとめて初期値に戻す。ここで書き戻すだけで、保存はOKを押したとき
           / Puts the defaults back into the fields; nothing is stored until OK */
        btnReset.onClick = function () {
            folderPanelUI.reset();
            keywordListInput.text = KEYWORD_PRESET_DEFAULTS.join("\n");
            excludeListInput.text = EXCLUDE_KEYWORD_DEFAULTS.join("\n");
        };

        /* 再スキャンは、いま入力してある設定を活かしたまま索引を作り直す / Rescan applies the edits, then rebuilds */
        var requestedRescan = false;
        btnRescan.onClick = function () {
            requestedRescan = true;
            settingsDialog.close(1);
        };

        settingsDialog.center();
        if (settingsDialog.show() !== 1) return null;

        return {
            folders: folderPanelUI.getFolders(),
            keywords: toWordList(keywordListInput.text),
            excludes: toWordList(excludeListInput.text),
            rescan: requestedRescan
        };
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 絞り込みパネルを組み立てる
     * @param {Window} parent - 追加先のウィンドウ
     * @param {Array<number>} years - 年別に並べる更新年
     * @returns {{filterPanel: Panel, keywordInput: EditText, clearButton: Button, matchAll: RadioButton, matchAny: RadioButton, extensionCheckboxes: Array<Checkbox>, yearDropdown: DropDownList, periodFrom: object, periodTo: object, presetContainer: Group}} パネルの部品
     */
    function buildFilterPanel(parent, years) {
        var filterPanel = parent.add("panel", undefined, formatLabel(getLabel(LABELS.panel.filter), [0]));
        setupPanel(filterPanel, DENSE_SPACING);

        var keywordRow = filterPanel.add("group");
        setupRow(keywordRow, "fill", DENSE_SPACING);

        var keywordLabel = keywordRow.add("statictext", undefined, labelText(LABELS.fieldLabel.keyword));

        /* 項目名の幅を測って固定し、下の行を入力欄の左端にそろえる
           / Fix the label width, then line the row below up with the field */
        var keywordLabelWidth = Math.ceil(measureTextWidth(keywordLabel, keywordLabel.text));
        keywordLabel.preferredSize.width = keywordLabelWidth;

        var keywordInput = keywordRow.add("edittext", undefined, "");
        keywordInput.alignment = ["fill", "center"];

        var clearButton = addClearButton(keywordRow);

        /* 判定は左、拡張子は右。間はスペーサーで押し広げる / Match mode on the left, extensions on the right */
        var matchRow = filterPanel.add("group");
        setupRow(matchRow, "fill", DENSE_SPACING);
        matchRow.margins = [keywordLabelWidth + DENSE_SPACING, 0, 0, 0];

        /* 語を空白で区切ったときの扱い。ANDは全部含むもの、ORはどれか含むもの
           / How space-separated terms combine: AND needs them all, OR any one */
        var matchPair = addRadioPair(matchRow, LABELS.radio.matchAll, LABELS.radio.matchAny, MATCH_ALL_DEFAULT);

        addSpacer(matchRow);

        /* チェックの入っている拡張子だけを一覧に出す / Only the checked extensions are listed */
        var extensionCheckboxes = addExtensionCheckboxes(matchRow);

        var dateFilterUI = buildDateFilterRow(filterPanel, years);

        /* キーワードボタンはパネルの一番下に置く。数で行数が変わるので、置き場だけ先に用意する
           / The preset buttons sit at the bottom; only the area is reserved, since the rows are rebuilt */
        var presetContainer = filterPanel.add("group");
        presetContainer.orientation = "column";
        presetContainer.alignChildren = ["left", "top"];
        presetContainer.alignment = ["fill", "top"];
        presetContainer.spacing = DENSE_SPACING;
        presetContainer.margins = [0, PRESET_TOP_MARGIN, 0, 0];

        return {
            filterPanel: filterPanel,
            keywordInput: keywordInput,
            clearButton: clearButton,
            matchAll: matchPair.first,
            matchAny: matchPair.second,
            extensionCheckboxes: extensionCheckboxes,
            yearDropdown: dateFilterUI.yearDropdown,
            periodFrom: dateFilterUI.from,
            periodTo: dateFilterUI.to,
            presetContainer: presetContainer
        };
    }

    /**
     * 左右2本のリストを組み立てる
     * @param {Window} parent - 追加先のウィンドウ
     * @returns {{folderListBox: ListBox, fileListBox: ListBox}} リスト部品
     */
    function buildListColumns(parent) {
        /* 左でフォルダーを選び、右にそのフォルダー内のファイルだけを並べる / Folder on the left, files on the right */
        var listRow = parent.add("group");
        setupRow(listRow, "fill", COLUMN_SPACING);
        listRow.alignChildren = ["fill", "fill"];

        var folderListBox = listRow.add("listbox", undefined, [], {
            numberOfColumns: 1,
            showHeaders: true,
            columnTitles: [getLabel(LABELS.listCaption.folder)],
            columnWidths: [FOLDER_COLUMN_WIDTH]
        });
        folderListBox.preferredSize = FOLDER_LIST_SIZE;

        var fileListBox = listRow.add("listbox", undefined, [], {
            numberOfColumns: 2,
            showHeaders: true,
            columnTitles: [getLabel(LABELS.listCaption.fileName), getLabel(LABELS.listCaption.modified)],
            columnWidths: [NAME_COLUMN_WIDTH, DATE_COLUMN_WIDTH]
        });
        fileListBox.preferredSize = FILE_LIST_SIZE;

        return { folderListBox: folderListBox, fileListBox: fileListBox };
    }

    /**
     * 更新年を並べたポップアップを作る
     * @param {Group} parent - 追加先のグループ
     * @param {Array<number>} years - 並べる更新年
     * @param {{ja: string, en: string}} placeholderSet - 先頭に置く「絞り込まない」項目の名前
     * @param {number} width - ポップアップの幅
     * @returns {DropDownList} 追加したポップアップ
     */
    function addYearDropdown(parent, years, placeholderSet, width) {
        var yearDropdown = parent.add("dropdownlist", undefined, []);
        yearDropdown.preferredSize.width = width;
        yearDropdown.add("item", getLabel(placeholderSet));

        for (var i = 0; i < years.length; i++) {
            var yearItem = yearDropdown.add("item", String(years[i]));
            yearItem.year = years[i];
        }
        yearDropdown.selection = 0;
        return yearDropdown;
    }

    /**
     * 期間の片側（年と月のポップアップ）を組み立てる
     * @param {Group} parent - 追加先のグループ
     * @param {Array<number>} years - 年に並べる更新年
     * @param {number} defaultMonth - 最初に選んでおく月
     * @returns {{yearDropdown: DropDownList, monthDropdown: DropDownList}} 年と月のポップアップ
     */
    function addYearMonthPair(parent, years, defaultMonth) {
        var pairGroup = parent.add("group");
        setupRow(pairGroup, "left", PERIOD_PAIR_SPACING);

        var yearDropdown = addYearDropdown(pairGroup, years, LABELS.dropdown.noLimit, PERIOD_YEAR_WIDTH);
        yearDropdown.helpTip = getLabel(LABELS.hint.periodYear);

        var monthDropdown = pairGroup.add("dropdownlist", undefined, []);
        monthDropdown.preferredSize.width = PERIOD_MONTH_WIDTH;
        monthDropdown.helpTip = getLabel(LABELS.hint.periodMonth);
        for (var month = 1; month <= 12; month++) {
            var monthItem = monthDropdown.add("item", padZero(month));
            monthItem.month = month;
        }
        monthDropdown.selection = defaultMonth - 1;

        /* 年が「なし」の間は月を触れないようにしておく / The month waits until a year is picked */
        monthDropdown.enabled = false;

        return { yearDropdown: yearDropdown, monthDropdown: monthDropdown };
    }

    /**
     * 年別と期間の行を組み立てる
     * @param {Panel} parent - 追加先のパネル
     * @param {Array<number>} years - 年に並べる更新年
     * @returns {{yearDropdown: DropDownList, from: object, to: object}} 年別と、期間の開始側・終了側
     */
    function buildDateFilterRow(parent, years) {
        var dateFilterRow = parent.add("group");
        setupRow(dateFilterRow, "left", DENSE_SPACING);

        dateFilterRow.add("statictext", undefined, labelText(LABELS.fieldLabel.year));
        var yearDropdown = addYearDropdown(dateFilterRow, years, LABELS.dropdown.allYears, YEAR_DROPDOWN_WIDTH);

        /* 年別と地続きに見えないよう、期間の前を空ける / Keep the period filter from reading as part of the year one */
        var periodSpacer = dateFilterRow.add("group");
        periodSpacer.preferredSize.width = PERIOD_LEFT_SPACE;
        periodSpacer.maximumSize.width = PERIOD_LEFT_SPACE;

        dateFilterRow.add("statictext", undefined, labelText(LABELS.fieldLabel.period));

        /* 月の初期値は、開始側が1月、終了側が12月。年を選ぶだけでその年の全体が指定できる
           / January on the left and December on the right, so picking a year covers it whole */
        var fromPair = addYearMonthPair(dateFilterRow, years, 1);
        dateFilterRow.add("statictext", undefined, getLabel(LABELS.fieldLabel.periodTo));
        var toPair = addYearMonthPair(dateFilterRow, years, 12);

        return { yearDropdown: yearDropdown, from: fromPair, to: toPair };
    }

    /**
     * ラジオボタンを2つ並べた組を作る
     * ラジオボタンは親のグループ単位で排他になるので、組ごとにグループで囲む
     * @param {Group} parent - 追加先のグループ
     * @param {{ja: string, en: string}} firstSet - 1つめの項目名
     * @param {{ja: string, en: string}} secondSet - 2つめの項目名
     * @param {boolean} isFirstSelected - 1つめを選んだ状態で始めるなら true
     * @returns {{first: RadioButton, second: RadioButton}} 2つのラジオボタン
     */
    function addRadioPair(parent, firstSet, secondSet, isFirstSelected) {
        var radioGroup = parent.add("group");
        setupRow(radioGroup, "left", DENSE_SPACING);

        var firstRadio = radioGroup.add("radiobutton", undefined, getLabel(firstSet));
        var secondRadio = radioGroup.add("radiobutton", undefined, getLabel(secondSet));
        firstRadio.value = isFirstSelected;
        secondRadio.value = !isFirstSelected;

        return { first: firstRadio, second: secondRadio };
    }

    /**
     * 拡張子で絞り込むチェックボックスを並べる
     * @param {Group} parent - 追加先のグループ
     * @returns {Array<Checkbox>} 拡張子を持たせたチェックボックス。FILE_EXTENSIONS の順に並ぶ
     */
    function addExtensionCheckboxes(parent) {
        var extensionGroup = parent.add("group");
        setupRow(extensionGroup, "left", DENSE_SPACING);

        var extensionCheckboxes = [];
        for (var i = 0; i < FILE_EXTENSIONS.length; i++) {
            var extensionCheckbox = extensionGroup.add("checkbox", undefined, FILE_EXTENSIONS[i].ext);
            extensionCheckbox.value = FILE_EXTENSIONS[i].isChecked;
            extensionCheckbox.extension = FILE_EXTENSIONS[i].ext;
            extensionCheckboxes.push(extensionCheckbox);
        }
        return extensionCheckboxes;
    }

    /**
     * リスト下の並び順の行を組み立てる
     * @param {Window} parent - 追加先のウィンドウ
     * @returns {{byModified: RadioButton, byName: RadioButton, ascending: RadioButton, descending: RadioButton}} 並び順のラジオボタン
     */
    function buildSortRow(parent) {
        /* 行そのものを右寄せにすると中身ぴったりの幅になり、右のリストの右端にそろう
           / A right-aligned row hugs its contents and lines up with the file list */
        var sortRow = parent.add("group");
        setupRow(sortRow, "right", DENSE_SPACING);
        sortRow.margins = [0, ROW_TOP_MARGIN, 0, 0];

        sortRow.add("statictext", undefined, labelText(LABELS.fieldLabel.sortBy));

        var sortKeyPair = addRadioPair(sortRow, LABELS.radio.sortByModified, LABELS.radio.sortByName, SORT_BY_MODIFIED_DEFAULT);

        /* ▲▼は見れば分かるので項目名は付けない / The arrows speak for themselves */
        var orderGroup = sortRow.add("group");
        setupRow(orderGroup, "left", SORT_ORDER_BUTTON_SPACING);
        var ascendingButton = addSortOrderButton(orderGroup, false, !SORT_DESCENDING_DEFAULT);
        var descendingButton = addSortOrderButton(orderGroup, true, SORT_DESCENDING_DEFAULT);

        return {
            byModified: sortKeyPair.first,
            byName: sortKeyPair.second,
            ascending: ascendingButton,
            descending: descendingButton
        };
    }

    /**
     * ダイアログ下部のボタン列を組み立てる
     * @param {Window} parent - 追加先のウィンドウ
     * @returns {{preferences: Button, cancel: Button, open: Button}} ボタン
     */
    function buildDialogButtons(parent) {
        var buttonRow = addButtonRow(parent);
        var btnPreferences = buttonRow.left.add("button", undefined, getLabel(LABELS.button.preferences));

        var btnCancel = buttonRow.right.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });
        var btnOpen = buttonRow.right.add("button", undefined, getLabel(LABELS.button.open), { name: "ok" });
        btnOpen.enabled = false;

        applyButtonSize(btnPreferences, DIALOG_BUTTON_WIDTH);
        applyButtonSize(btnCancel, DIALOG_BUTTON_WIDTH);
        applyButtonSize(btnOpen, DIALOG_BUTTON_WIDTH);

        return {
            preferences: btnPreferences,
            cancel: btnCancel,
            open: btnOpen
        };
    }

    /**
     * ファインダーのダイアログを表示する
     * @param {Array<FileEntry>} fileEntries - 検索対象のファイル
     * @param {Array<Folder>} searchFolders - 現在の検索フォルダー
     * @returns {{action: string, entry: FileEntry|null, folders: Array<Folder>|null}} 操作結果（"open" / "rescan" / "cancel"）
     */
    function showFinderDialog(fileEntries, searchFolders) {
        var finderDialog = new Window("dialog", getLabel(LABELS.dialog.title) + " " + SCRIPT_VERSION);
        setupWindow(finderDialog);

        /* 年の顔ぶれは絞り込みで変えない。打つたびに選択肢が消えると選びにくい
           / The year list is built once so the choices do not shift while typing */
        var modifiedYears = collectModifiedYears(fileEntries);

        /* 検索欄を最初の操作部品にして、起動時のフォーカスを安定させる / Keep the keyword field first */
        var filterUI = buildFilterPanel(finderDialog, modifiedYears);
        var filterPanel = filterUI.filterPanel;
        var keywordInput = filterUI.keywordInput;
        var keywordClearButton = filterUI.clearButton;
        var keywordMatchAll = filterUI.matchAll;
        var keywordMatchAny = filterUI.matchAny;
        var extensionCheckboxes = filterUI.extensionCheckboxes;
        var yearDropdown = filterUI.yearDropdown;
        var periodFromPair = filterUI.periodFrom;
        var periodToPair = filterUI.periodTo;
        var keywordPresetContainer = filterUI.presetContainer;

        var listBoxUI = buildListColumns(finderDialog);
        var folderListBox = listBoxUI.folderListBox;
        var fileListBox = listBoxUI.fileListBox;

        var sortUI = buildSortRow(finderDialog);

        var dialogButtons = buildDialogButtons(finderDialog);
        var btnPreferences = dialogButtons.preferences;
        var btnCancel = dialogButtons.cancel;
        var btnOpen = dialogButtons.open;

        /* 検索フォルダーは登録順のまま左のリストに並べる / The folder list follows the configured order */
        var rootInfoList = makeRootFolderInfoList(searchFolders);

        var keywordPresets = readWordList(PREF_KEY_KEYWORDS, KEYWORD_PRESET_DEFAULTS);
        var keywordPresetRows = [];

        var excludeKeywords = readWordList(PREF_KEY_EXCLUDES, EXCLUDE_KEYWORD_DEFAULTS);
        var excludeTerms = toNormalizedTerms(excludeKeywords);
        applyExcludeFlags(fileEntries, excludeTerms);

        /* 並べ替え済みの並びを覚えておく。並び順を変えるたびに数千件を並べ替え直すと待たされる
           / Cached orders; re-sorting thousands of entries on every toggle would stall the dialog */
        var sortedByModified = fileEntries;
        var sortedByName = null;

        /* 期間の絞り込み。null は制限なし / Period filter, null means no limit */
        var periodFrom = null;
        var periodTo = null;

        var filteredEntries = [];
        var listedEntries = [];
        var dialogResult = { action: "cancel", entry: null, folders: null };

        /* 組み直し中の選択変更でファイルリストが何度も再構築されるのを防ぐ / Suppress cascaded rebuilds */
        var isRebuildingFolderList = false;

        /**
         * 年別で選択中の年を返す
         * @returns {number|null} 選択中の年。「すべて」選択時は null
         */
        function selectedFilterYear() {
            if (!yearDropdown.selection) return null;

            var year = yearDropdown.selection.year;
            return (year === undefined) ? null : year;
        }

        /**
         * いま選ばれている並び順の、昇順で並べた配列を返す
         * 一度作った並びは使い回し、名前順は最初に選ばれたときだけ作る
         * @returns {Array<FileEntry>} 昇順に並べたファイル
         */
        function getOrderedEntries() {
            if (sortUI.byModified.value) return sortedByModified;

            if (!sortedByName) {
                sortedByName = fileEntries.slice(0);
                sortFileEntries(sortedByName, false);
            }
            return sortedByName;
        }

        /**
         * 並び順の指定に合わせて一覧を組み直す
         * @returns {void}
         */
        function handleSortChanged() {
            refreshFolderList();
        }

        /**
         * 昇順・降順を切り替える
         * ラジオボタンと違って排他にならないので、選択の付け替えと三角の差し替えは自分で行う
         * @param {boolean} isDescending - 降順にするなら true
         * @returns {void}
         */
        function setSortOrder(isDescending) {
            if (sortUI.descending.selected === isDescending) return;

            sortUI.ascending.selected = !isDescending;
            sortUI.descending.selected = isDescending;
            sortUI.ascending.text = getSortOrderGlyph(false, !isDescending);
            sortUI.descending.text = getSortOrderGlyph(true, isDescending);
            handleSortChanged();
        }

        /**
         * 期間のポップアップを読み直して一覧を組み直す
         * @returns {void}
         */
        function handlePeriodChanged() {
            periodFrom = readPeriodBoundary(periodFromPair, false);
            periodTo = readPeriodBoundary(periodToPair, true);
            refreshFolderList();
        }

        /**
         * 選択中の検索フォルダーの番号を返す
         * @returns {number|null} 検索フォルダーの番号。「すべて」選択時と未選択時は null
         */
        function selectedRootIndex() {
            if (!folderListBox.selection) return null;

            var rootIndex = folderListBox.selection.rootIndex;
            return (rootIndex === undefined) ? null : rootIndex;
        }

        /**
         * キーワード・年別・期間・除外条件で絞り込み、左のフォルダーリストを組み直す
         * サブフォルダーの中身も含めて、検索フォルダー単位でまとめる
         * @returns {void}
         */
        function refreshFolderList() {
            var previousRootIndex = selectedRootIndex();
            isRebuildingFolderList = true;

            rebuildFolderListBox(collectFilteredEntries(), previousRootIndex);

            isRebuildingFolderList = false;
            refreshFileList();
        }

        /**
         * チェックの入っている拡張子を引く
         * @returns {Object} 拡張子をキーにしたハッシュ。キー衝突を避けるため接頭辞を付ける
         */
        function collectAllowedExtensions() {
            var allowedExtensions = {};
            for (var i = 0; i < extensionCheckboxes.length; i++) {
                if (extensionCheckboxes[i].value) allowedExtensions["#" + extensionCheckboxes[i].extension] = true;
            }
            return allowedExtensions;
        }

        /**
         * 条件に合うファイルを filteredEntries に集める
         * @returns {Array<number>} 集まったファイルが属する検索フォルダーの番号
         */
        function collectFilteredEntries() {
            var searchTerms = splitSearchTerms(keywordInput.text);
            var isMatchAll = keywordMatchAll.value;
            var filterYear = selectedFilterYear();
            filteredEntries = [];

            var allowedExtensions = collectAllowedExtensions();

            /* キーワードに入れた除外語はここで抜ける。残った語だけを印の付いたファイルへ当て直す
               / Exclusion terms typed as keywords drop out; the rest are re-tested on the flagged files */
            var activeExcludeTerms = collectActiveExcludeTerms(excludeTerms, searchTerms);
            var isExcludeLifted = activeExcludeTerms.length !== excludeTerms.length;

            /* 昇順に並べた配列を、降順なら後ろからたどる。並べ替え直さずに向きを変えられる
               / Walk the ascending order backwards for descending; no re-sorting needed */
            var orderedEntries = getOrderedEntries();
            var isDescending = sortUI.descending.selected;
            var startIndex = isDescending ? orderedEntries.length - 1 : 0;
            var step = isDescending ? -1 : 1;

            var seenRootIndexes = {};
            var rootIndexes = [];
            for (var i = startIndex; i >= 0 && i < orderedEntries.length; i += step) {
                var entry = orderedEntries[i];
                if (!allowedExtensions["#" + entry.extension]) continue;
                if (entry.isExcluded &&
                    (!isExcludeLifted || matchesAnyTerm(entry.normalizedSearchText, activeExcludeTerms))) continue;
                if (filterYear !== null && entry.modifiedYear !== filterYear) continue;
                if (periodFrom !== null && entry.modifiedTime < periodFrom) continue;
                if (periodTo !== null && entry.modifiedTime > periodTo) continue;
                if (!matchesKeyword(entry.normalizedSearchText, searchTerms, isMatchAll)) continue;

                filteredEntries.push(entry);

                /* ハッシュのキー衝突を避けるため接頭辞を付けて既出判定する / Prefix the key to avoid collisions */
                var rootKey = "#" + entry.rootIndex;
                if (!seenRootIndexes[rootKey]) {
                    seenRootIndexes[rootKey] = true;
                    rootIndexes.push(entry.rootIndex);
                }
            }

            /* 環境設定で並べた順のまま見せる / Keep the configured folder order */
            rootIndexes.sort(function (indexA, indexB) { return indexA - indexB; });
            return rootIndexes;
        }

        /**
         * 左のフォルダーリストを組み直す
         * @param {Array<number>} rootIndexes - 並べる検索フォルダーの番号
         * @param {number|null} previousRootIndex - 組み直す前に選んでいた検索フォルダーの番号
         * @returns {void}
         */
        function rebuildFolderListBox(rootIndexes, previousRootIndex) {
            folderListBox.removeAll();
            folderListBox.add("item", getLabel(LABELS.folderRow.allFolders));
            for (var i = 0; i < rootIndexes.length; i++) {
                var folderItem = folderListBox.add("item", rootInfoList[rootIndexes[i]].label);
                folderItem.rootIndex = rootIndexes[i];
            }

            /* 絞り込み前に選んでいたフォルダーが残っていれば選択を引き継ぐ / Keep the previous folder selection */
            folderListBox.selection = 0;
            if (previousRootIndex === null) return;

            for (var j = 1; j < folderListBox.items.length; j++) {
                if (folderListBox.items[j].rootIndex !== previousRootIndex) continue;
                folderListBox.selection = j;
                return;
            }
        }

        /**
         * 選択中フォルダーに合わせて右のファイルリストを組み直す
         * @returns {void}
         */
        function refreshFileList() {
            var rootIndex = selectedRootIndex();
            fileListBox.removeAll();
            listedEntries = [];

            var matchCount = 0;
            for (var i = 0; i < filteredEntries.length; i++) {
                if (rootIndex !== null && filteredEntries[i].rootIndex !== rootIndex) continue;
                matchCount++;

                /* 一度に並べる数を抑える。全件を並べるとリストの組み直しだけで待たされる / Cap the rows to keep typing responsive */
                if (listedEntries.length >= FILE_LIST_MAX_ITEMS) continue;

                listedEntries.push(filteredEntries[i]);
                var fileItem = fileListBox.add("item", toDisplayName(filteredEntries[i].fileName));
                fileItem.subItems[0].text = formatModifiedTime(filteredEntries[i].modifiedTime);
                fileItem.entryIndex = listedEntries.length - 1;
            }

            /* 件数はパネルのタイトルに出す。打ち切った場合は表示件数も添える / Show the match count in the panel title */
            filterPanel.text = (matchCount > listedEntries.length)
                ? formatLabel(getLabel(LABELS.panel.filterLimited), [matchCount, listedEntries.length])
                : formatLabel(getLabel(LABELS.panel.filter), [matchCount]);

            if (fileListBox.items.length > 0) fileListBox.selection = 0;
            updateSelectionState();
        }

        /**
         * 右のリストで選択中のファイルを返す
         * @returns {FileEntry|null} 選択中のファイル。未選択なら null
         */
        function selectedFileEntry() {
            if (!fileListBox.selection) return null;

            var selectedIndex = fileListBox.selection.entryIndex;
            if (selectedIndex === undefined) return null;
            return listedEntries[selectedIndex] || null;
        }

        /**
         * 選択に合わせて［開く］の有効・無効を切り替える
         * @returns {void}
         */
        function updateSelectionState() {
            btnOpen.enabled = !!selectedFileEntry();
        }

        /**
         * 選択中のファイルを開く対象に確定してダイアログを閉じる
         * @returns {void}
         */
        function openSelectedFile() {
            var selectedEntry = selectedFileEntry();
            if (!selectedEntry) return;

            dialogResult.action = "open";
            dialogResult.entry = selectedEntry;
            finderDialog.close(1);
        }

        /**
         * 選択中のファイルをFinderで表示する（ダイアログは開いたまま）
         * @returns {void}
         */
        function revealSelectedFile() {
            var selectedEntry = selectedFileEntry();
            if (selectedEntry) revealFile(selectedEntry.file);
        }

        /**
         * 左のリストで選択中のフォルダーをFinderで開く（ダイアログは開いたまま）
         * 「すべて」のときは何もしない
         * @returns {void}
         */
        function openSelectedFolder() {
            var rootIndex = selectedRootIndex();
            if (rootIndex === null) return;

            rootInfoList[rootIndex].folder.execute();
        }

        /**
         * キーワードボタンを組み立て直す
         * @returns {void}
         */
        function refreshKeywordPresetButtons() {
            for (var i = 0; i < keywordPresetRows.length; i++) {
                keywordPresetContainer.remove(keywordPresetRows[i]);
            }
            keywordPresetRows = [];

            /* 行に収まらなくなったら次の行へ送る。ダイアログを広げないため幅で折り返す
               / Wrap by width so the dialog never has to grow sideways */
            var presetRow = null;
            var rowWidth = 0;
            for (var j = 0; j < keywordPresets.length; j++) {
                var buttonWidth = estimateTextWidth(keywordPresets[j]) + PRESET_BUTTON_PADDING;
                var filledWidth = (rowWidth === 0) ? buttonWidth : rowWidth + DENSE_SPACING + buttonWidth;

                if (presetRow === null || filledWidth > PRESET_ROW_WIDTH) {
                    presetRow = keywordPresetContainer.add("group");
                    setupRow(presetRow, "left", DENSE_SPACING);
                    keywordPresetRows.push(presetRow);
                    filledWidth = buttonWidth;
                }

                var presetButton = presetRow.add("button", undefined, keywordPresets[j]);
                presetButton.helpTip = getLabel(LABELS.button.appendKeyword);
                applyPresetButtonSize(presetButton);
                presetButton.onClick = makeKeywordPresetHandler(keywordPresets[j]);
                rowWidth = filledWidth;
            }
        }

        /**
         * キーワードボタン用の onClick ハンドラーを作る
         * @param {string} presetKeyword - ボタンに割り当てるキーワード
         * @returns {function(): void} 通常クリックで置き換え、option+クリックで追加するハンドラー
         */
        function makeKeywordPresetHandler(presetKeyword) {
            return function () {
                /* option+クリックは空白を挟んで語を足し、AND検索で絞り込む / Option-click appends the word */
                var currentText = trimWhitespace(keywordInput.text);
                var isAppend = ScriptUI.environment.keyboardState.altKey && currentText !== "";
                keywordInput.text = isAppend ? currentText + " " + presetKeyword : presetKeyword;

                /* 続けて打ち足せるよう、フォーカスは入力欄へ移す / Put the focus in the field */
                keywordInput.active = true;
                handleKeywordChanged();
            };
        }

        /* 検索キーは作成済みなので入力時は比較だけを行う。onChangingで即時反映し、onChangeでIME確定も拾う */
        var lastNormalizedQuery = null;

        /**
         * キーワードが変わったときだけリストを組み直す
         * @returns {void}
         */
        function refreshListIfQueryChanged() {
            var normalizedQuery = splitSearchTerms(keywordInput.text).join(" ");
            if (normalizedQuery === lastNormalizedQuery) return;

            lastNormalizedQuery = normalizedQuery;
            refreshFolderList();
        }

        /**
         * キーワードの有無に合わせてクリアボタンのディム表示を切り替える
         * @returns {void}
         */
        function updateClearButtonState() {
            var hasKeyword = trimWhitespace(keywordInput.text) !== "";
            if (keywordClearButton.enabled === hasKeyword) return;

            keywordClearButton.enabled = hasKeyword;
            keywordClearButton.notify("onDraw");
        }

        /**
         * キーワードが変わったときの処理をまとめる
         * @returns {void}
         */
        function handleKeywordChanged() {
            updateClearButtonState();
            refreshListIfQueryChanged();
        }

        /**
         * 環境設定を開き、変わった設定を反映する
         * 検索フォルダーが変わったときだけ索引を作り直すため、ダイアログを閉じる
         * @returns {void}
         */
        function handlePreferences() {
            var changedSettings = showPreferencesDialog(searchFolders, keywordPresets, excludeKeywords);
            if (!changedSettings) return;

            var isFolderChanged = makeFolderSignature(changedSettings.folders) !== makeFolderSignature(searchFolders);
            var isKeywordChanged = changedSettings.keywords.join("\n") !== keywordPresets.join("\n");
            var isExcludeChanged = changedSettings.excludes.join("\n") !== excludeKeywords.join("\n");

            if (isKeywordChanged) {
                keywordPresets = changedSettings.keywords;
                saveWordList(PREF_KEY_KEYWORDS, keywordPresets);
            }
            if (isExcludeChanged) {
                excludeKeywords = changedSettings.excludes;
                excludeTerms = toNormalizedTerms(excludeKeywords);
                applyExcludeFlags(fileEntries, excludeTerms);
                saveWordList(PREF_KEY_EXCLUDES, excludeKeywords);
            }

            /* 検索フォルダーが変わったときと再スキャンを押されたときは、索引を作り直すため開き直す
               / A changed folder list or an explicit rescan reopens the dialog with a fresh index */
            if (isFolderChanged || changedSettings.rescan) {
                dialogResult.action = "rescan";
                dialogResult.folders = isFolderChanged ? changedSettings.folders : null;
                finderDialog.close(2);
                return;
            }

            /* 除外条件は索引を作り直さずに効く。一覧だけ組み直す / Exclusions apply without rebuilding the index */
            if (isExcludeChanged) refreshFolderList();
            if (!isKeywordChanged) return;

            /* ボタンの数で行数が変わるので、ダイアログ全体を組み直す / The row count changes, so re-layout */
            refreshKeywordPresetButtons();
            finderDialog.layout.layout(true);
        }

        /**
         * Enterならファイルを開き、イベントをそこで止める
         * @param {Object} keyEvent - keydown イベント
         * @returns {boolean} 開く操作として扱ったら true
         */
        function handleEnterToOpen(keyEvent) {
            if (keyEvent.keyName !== "Enter" && keyEvent.keyName !== "Return") return false;

            openSelectedFile();
            keyEvent.preventDefault();
            if (keyEvent.stopPropagation) keyEvent.stopPropagation();
            return true;
        }

        /**
         * 画面の操作を、対応する処理に結び付ける
         * @returns {void}
         */
        function bindEventHandlers() {
            bindKeywordHandlers();
            bindDateAndSortHandlers();
            bindListHandlers();
            bindKeyHandlers();
        }

        /**
         * キーワード欄まわりの操作を結び付ける
         * @returns {void}
         */
        function bindKeywordHandlers() {
            keywordInput.onChanging = handleKeywordChanged;
            keywordInput.onChange = handleKeywordChanged;

            /* AND / OR の切り替えはキーワードが同じままなので、組み直しを直に呼ぶ
               / The keyword does not change here, so rebuild the list directly */
            keywordMatchAll.onClick = refreshFolderList;
            keywordMatchAny.onClick = refreshFolderList;

            /* 拡張子のチェックもキーワードは変わらないので、同じく直に組み直す
               / The extension checkboxes leave the keyword alone too */
            for (var i = 0; i < extensionCheckboxes.length; i++) {
                extensionCheckboxes[i].onClick = refreshFolderList;
            }

            keywordClearButton.onClick = function () {
                keywordInput.text = "";
                /* 続けて打ち直せるよう、フォーカスは入力欄へ戻す / Put the focus back in the field */
                keywordInput.active = true;
                handleKeywordChanged();
            };
        }

        /**
         * 更新日での絞り込みと並び順の操作を結び付ける
         * @returns {void}
         */
        function bindDateAndSortHandlers() {
            sortUI.byModified.onClick = handleSortChanged;
            sortUI.byName.onClick = handleSortChanged;
            sortUI.ascending.onClick = function () { setSortOrder(false); };
            sortUI.descending.onClick = function () { setSortOrder(true); };

            yearDropdown.onChange = refreshFolderList;
            periodFromPair.yearDropdown.onChange = handlePeriodChanged;
            periodFromPair.monthDropdown.onChange = handlePeriodChanged;
            periodToPair.yearDropdown.onChange = handlePeriodChanged;
            periodToPair.monthDropdown.onChange = handlePeriodChanged;
        }

        /**
         * 左右のリストと、下のボタンの操作を結び付ける
         * @returns {void}
         */
        function bindListHandlers() {
            folderListBox.onChange = function () {
                if (isRebuildingFolderList) return;
                refreshFileList();
            };
            folderListBox.onDoubleClick = openSelectedFolder;

            fileListBox.onChange = updateSelectionState;
            fileListBox.onDoubleClick = function () {
                /* option+ダブルクリックは開かずにFinderで場所を出す / Option-double-click reveals instead of opening */
                if (ScriptUI.environment.keyboardState.altKey) {
                    revealSelectedFile();
                    return;
                }
                openSelectedFile();
            };

            btnPreferences.onClick = handlePreferences;
            btnOpen.onClick = openSelectedFile;
            btnCancel.onClick = function () {
                dialogResult.action = "cancel";
                finderDialog.close();
            };
        }

        /**
         * キー操作を結び付ける
         * @returns {void}
         */
        function bindKeyHandlers() {
            keywordInput.addEventListener("keydown", function (event) {
                /* ↓でファイルリストへ移る。1件も無いときは入力欄に留まる / Down moves into the file list */
                if (event.keyName === "Down" && fileListBox.items.length > 0) {
                    fileListBox.active = true;
                    if (!fileListBox.selection) fileListBox.selection = 0;
                    event.preventDefault();
                    return;
                }

                /* 絞り込みは onChanging / onChange で済んでいる。ここで組み直すと選択が先頭へ戻る */
                handleEnterToOpen(event);
            });

            fileListBox.addEventListener("keydown", handleEnterToOpen);

            /* テンキーEnterを含め、ダイアログ内のどこにフォーカスがあっても開く / Enter opens from anywhere */
            finderDialog.addEventListener("keydown", function (event) {
                /* ボタンにフォーカスがあるときは、そのボタン自身の動作に任せる */
                if (event.target && event.target.type === "button") return;

                handleEnterToOpen(event);
            });
        }

        bindEventHandlers();
        refreshKeywordPresetButtons();
        refreshFolderList();
        lastNormalizedQuery = splitSearchTerms(keywordInput.text).join(" ");
        finderDialog.center();

        /* 検索欄を先頭の操作部品にしたうえで、表示時にも明示的にフォーカスする / Focus the keyword field on show */
        keywordInput.active = true;
        finderDialog.onShow = function () {
            keywordInput.active = true;
            keywordInput.selection = [0, 0];
        };

        finderDialog.show();
        return dialogResult;
    }

    // =========================================
    // 実行 / Run
    // =========================================

    /**
     * Finder表示用のAutomatorアプリを探す
     * .app は実体がフォルダーなので File.exists が false を返す環境がある。Folder でも確認する
     * @returns {File|null} 実在するアプリ。無ければ null
     */
    function findRevealApp() {
        if ($.os.indexOf("Macintosh") === -1) return null;
        if (!new Folder(REVEAL_APP_PATH).exists && !new File(REVEAL_APP_PATH).exists) return null;
        return new File(REVEAL_APP_PATH);
    }

    /**
     * Finder表示用アプリに渡すパスを一時ファイルへ書き出す
     * @param {File} targetFile - 対象のファイル
     * @returns {boolean} 書き出せたら true
     */
    function writeRevealPath(targetFile) {
        var pathFile = new File(REVEAL_PATH_FILE);
        try {
            pathFile.encoding = "UTF-8";
            pathFile.lineFeed = "Unix";
            if (!pathFile.open("w")) return false;

            /* fsName で ~ ではなく絶対パスを渡す / fsName gives the absolute POSIX path */
            pathFile.write(targetFile.fsName);
            return true;
        } catch (e) {
            return false;
        } finally {
            closeQuietly(pathFile);
        }
    }

    /**
     * ファイルの場所をFinderで開く
     * @param {File} targetFile - 対象のファイル
     * @returns {void}
     */
    function revealFile(targetFile) {
        /* アプリが無い環境では囲みフォルダーを開くだけにとどめる / Fall back to the enclosing folder */
        var revealApp = findRevealApp();
        if (!revealApp) {
            targetFile.parent.execute();
            return;
        }

        /* アプリは一時ファイルからパスを読むので、書き出してから起動する / The app reads the path from a temp file */
        if (writeRevealPath(targetFile) && revealApp.execute()) return;

        /* 起動できなかった場合も何も起きないままにはしない / Never end up doing nothing */
        targetFile.parent.execute();
    }

    /**
     * すでに開いているドキュメントを探す
     * @param {File} targetFile - 対象のファイル
     * @returns {Document|null} 開いていればそのドキュメント。無ければ null
     */
    function findOpenDocument(targetFile) {
        var targetPath = targetFile.fsName;
        for (var i = 0; i < app.documents.length; i++) {
            var openDocument = app.documents[i];
            try {
                if (openDocument.fullName.fsName === targetPath) return openDocument;
            } catch (e) {}
        }
        return null;
    }

    /**
     * 選ばれたファイルを開く。すでに開いていればそのドキュメントを前面に出す
     * @param {FileEntry} fileEntry - 開くファイル
     * @returns {void}
     */
    function openFileEntry(fileEntry) {
        var targetFile = fileEntry.file;
        if (!targetFile.exists) {
            alert(formatLabel(getLabel(LABELS.alert.missingFile), [targetFile.fsName]), getLabel(LABELS.dialog.title));
            return;
        }

        var openDocument = findOpenDocument(targetFile);
        if (openDocument) {
            openDocument.activate();
            return;
        }

        try {
            app.open(targetFile);
        } catch (e) {
            alert(formatLabel(getLabel(LABELS.alert.openFailed), [targetFile.fsName, e.message]), getLabel(LABELS.dialog.title));
        }
    }

    /**
     * 環境設定で編集した設定をまとめて記録する
     * @param {{folders: Array<Folder>, keywords: Array<string>, excludes: Array<string>}} settings - 編集後の設定
     * @returns {void}
     */
    function saveSettings(settings) {
        saveSearchFolders(settings.folders);
        saveWordList(PREF_KEY_KEYWORDS, settings.keywords);
        saveWordList(PREF_KEY_EXCLUDES, settings.excludes);
    }

    /**
     * 検索フォルダーの確認からファイルを開くまでを進める
     * @returns {void}
     */
    function main() {
        var searchFolders = readSearchFolders();

        /* 検索フォルダーが未登録のうちは索引も作れない。0件のまま繰り返しへ入り、環境設定へ促す
           / With no folders registered there is nothing to index; the loop opens the preferences */
        var fileEntries = (searchFolders.length > 0) ? loadFileEntries(searchFolders, false) : [];

        while (true) {
            if (fileEntries.length === 0) {
                /* 走査した結果が0件のときだけ知らせる。未登録では出しても意味がない
                   / Only report an empty scan; there is nothing to report before the first folder */
                if (searchFolders.length > 0) alert(getLabel(LABELS.alert.noFiles), getLabel(LABELS.dialog.title));

                var changedSettings = showPreferencesDialog(searchFolders,
                    readWordList(PREF_KEY_KEYWORDS, KEYWORD_PRESET_DEFAULTS),
                    readWordList(PREF_KEY_EXCLUDES, EXCLUDE_KEYWORD_DEFAULTS));
                /* 選び直しを取り消したら終了する。対象が空のまま繰り返すと抜けられなくなる
                   / Cancelling ends the script; looping with no folders would never finish */
                if (!changedSettings || changedSettings.folders.length === 0) return;

                searchFolders = changedSettings.folders;
                saveSettings(changedSettings);
                fileEntries = loadFileEntries(searchFolders, true);
                continue;
            }

            var finderResult = showFinderDialog(fileEntries, searchFolders);

            if (finderResult.action === "rescan") {
                /* 空の配列も真なので、件数で確かめてから差し替える / An empty array is truthy, so check the length */
                if (finderResult.folders && finderResult.folders.length > 0) {
                    searchFolders = finderResult.folders;
                    saveSearchFolders(searchFolders);
                }
                fileEntries = loadFileEntries(searchFolders, true);
                continue;
            }

            if (finderResult.action === "open") openFileEntry(finderResult.entry);
            break;
        }
    }

    main();

})();
