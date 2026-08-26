#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

指定したフォルダー内の .jsx / .js / .jsxbin をキーワードで絞り込み、選んだスクリプトをその場で実行するランチャーです。
フォルダーとファイル名を左右のリストに分けて表示し、絞り込んだ結果によく出てくる語をワンクリックのボタンとして自動で並べます。

詳細は README を参照してください。

### Overview

A launcher that filters .jsx / .js / .jsxbin files in a chosen folder by keyword and runs the selected script on the spot.
Folders and file names are shown in two side-by-side lists, and words that appear often in the filtered results become one-click buttons automatically.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AiScriptLauncher";             /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.4.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-11-13";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-08-27";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AiScriptLauncher.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AiScriptLauncher.md
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n86fe7e6251ec"; /* 紹介記事 / article URL */

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

    /* 一覧に載せる拡張子 / Script file extensions to list */
    var SCRIPT_EXT_RE = /\.(jsx|js|jsxbin)$/i;

    /* スクリプト名から自動で作るキーワードボタンの条件 / Rules for the auto keyword buttons */
    var KEYWORD_MIN_WORD_LENGTH    = 3;      /* ボタンにする語の最小文字数 / minimum word length */
    var KEYWORD_PRESET_MIN_COUNT   = 4;      /* ボタンにする最小出現ファイル数 / minimum number of files */
    var KEYWORD_PRESET_MAX_BUTTONS = 10;     /* ボタンの最大個数 / maximum number of buttons */
    var KEYWORD_PRESET_LIMIT       = 40;     /* 環境設定で指定できる上限 / ceiling for the preferences dialog */
    var KEYWORD_PRESETS_PER_ROW    = 5;      /* 1行に並べる個数 / buttons per row */

    /* ボタンにしない語。件数は多いが検索の役に立たない接続語 / Words never turned into buttons */
    var KEYWORD_STOP_WORDS = ["and", "the", "for", "with", "from", "into", "その他"];

    /* Finder表示に使うAutomatorアプリと、パスを受け渡す一時ファイル / Automator app used to reveal a file */
    /* 一時ファイル名はアプリ内のAppleScriptが読む固定名。アプリを旧名 IllustratorRevealLink.app から
       改名した名残で綴りが揃っていないが、アプリ側を直すまでここは変えない
       / The temp file name is hard-coded in the app; leave it until the app itself is updated */
    var REVEAL_APP_PATH  = "/Applications/RevealInFinder.app";
    var REVEAL_PATH_FILE = "/tmp/illustrator_reveal_path.txt";

    /* 「サブディレクトリを含む」の初期状態 / Initial state of the subfolder checkbox */
    var INCLUDE_SUBFOLDERS_DEFAULT = false;

    /* 「フルパス」の初期状態。OFFではホームフォルダーを ~ に略す / Initial state of the full path checkbox */
    var SHOW_FULL_PATH_DEFAULT = false;

    /* サブディレクトリOFF時に残す階層の深さ（対象フォルダー直下のフォルダーまで） / Folder depth kept when subfolders are excluded */
    var NESTED_FOLDER_DEPTH_LIMIT = 1;

    // =========================================
    // 設定の保存 / Stored settings
    // =========================================

    /* Illustratorの環境設定に保存する。再起動しても残り、余分なファイルを作らない / Stored in Illustrator preferences */
    var PREF_KEY_FOLDER      = "AiScriptLauncher.targetFolder";
    var PREF_KEY_MIN_COUNT   = "AiScriptLauncher.keywordMinCount";
    var PREF_KEY_MAX_BUTTONS = "AiScriptLauncher.keywordMaxButtons";

    /**
     * 前回使った対象フォルダーを読み出す
     * @returns {Folder|null} 実在するフォルダー。記録がなければ null
     */
    function readSavedScriptFolder() {
        var savedPath = trimWhitespace(app.preferences.getStringPreference(PREF_KEY_FOLDER));
        if (!savedPath) return null;

        var savedFolder = new Folder(savedPath);
        return savedFolder.exists ? savedFolder : null;
    }

    /**
     * 対象フォルダーを次回起動用に記録する
     * @param {Folder} targetFolder - 記録するフォルダー
     * @returns {void}
     */
    function saveScriptFolder(targetFolder) {
        app.preferences.setStringPreference(PREF_KEY_FOLDER, targetFolder.fsName);
    }

    /**
     * 記録済みのキーワード設定を読み出す
     * @returns {{minCount: number, maxButtons: number}} 記録がなければユーザー設定の初期値
     */
    function readKeywordSettings() {
        /* 未登録のキーは 0 が返るので、そのときは初期値を使う / Missing keys come back as 0 */
        var minCount = app.preferences.getIntegerPreference(PREF_KEY_MIN_COUNT);
        var maxButtons = app.preferences.getIntegerPreference(PREF_KEY_MAX_BUTTONS);

        return {
            minCount: (minCount > 0) ? minCount : KEYWORD_PRESET_MIN_COUNT,
            maxButtons: (maxButtons > 0) ? maxButtons : KEYWORD_PRESET_MAX_BUTTONS
        };
    }

    /**
     * キーワード設定を記録する
     * @param {number} minCount - 出現数
     * @param {number} maxButtons - キーワード数
     * @returns {void}
     */
    function saveKeywordSettings(minCount, maxButtons) {
        app.preferences.setIntegerPreference(PREF_KEY_MIN_COUNT, minCount);
        app.preferences.setIntegerPreference(PREF_KEY_MAX_BUTTONS, maxButtons);
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
    var DENSE_SPACING      = 8;              /* パネル内を詰めるときの間隔 / dense panel spacing */
    var LIST_LABEL_SPACING = 4;              /* 見出しとリストの間隔 / gap under a list caption */

    /* リストと入力欄の寸法 / List and field sizes */
    var FOLDER_LIST_SIZE  = [165, 360];      /* フォルダーリストの寸法 [幅,高さ] / folder list size */
    var SCRIPT_LIST_SIZE  = [280, 360];      /* ファイル名リストの寸法 [幅,高さ]。ダイアログ幅＝キーワード欄の長さを決める / script list size */

    /* ボタンの寸法と余白 / Button sizes and margins */
    var BUTTON_HEIGHT           = 28;        /* ボタンの高さ / button height */
    var FOLDER_BUTTON_WIDTH     = 120;       /* フォルダー変更ボタンの幅 / change folder button width */
    var DIALOG_BUTTON_WIDTH     = 92;        /* 実行・キャンセルの幅 / dialog button width */
    var SETTINGS_BUTTON_WIDTH   = 100;       /* 環境設定ボタンの幅 / preferences button width */

    /* 環境設定ダイアログ / Preferences dialog */
    var SETTINGS_LABEL_WIDTH = 100;          /* 項目名の幅 / field label width */
    var SETTINGS_INPUT_WIDTH = 60;           /* 数値入力欄の幅 / number field width */
    var FOLDER_PATH_WIDTH    = 360;          /* 対象フォルダーのパス表示の幅 / target folder path width */
    var BUTTON_ROW_TOP_MARGIN   = 10;        /* ボタン列の上余白 / top margin above the button row */
    var CHECKBOX_TOP_MARGIN     = 10;        /* リスト下のチェックボックスの上余白 / top margin above a checkbox under a list */

    /* キーワードボタンは小ぶりにする / Keyword preset buttons are smaller */
    var PRESET_BUTTON_HEIGHT  = 22;          /* キーワードボタンの高さ / preset button height */
    var PRESET_BUTTON_PADDING = 16;          /* 文字幅に足す左右の余白 / horizontal padding */
    var PRESET_CHAR_WIDTH     = 7;           /* 実測できない環境用の1文字概算幅 / fallback char width */

    /* キーワード欄のクリアボタン（自前描画） / Clear button next to the keyword field (custom drawn) */
    var CLEAR_BUTTON_SIZE    = 20;           /* クリアボタンの一辺 / clear button size */
    var CLEAR_CIRCLE_INSET   = 2;            /* 円と外周の間隔 / inset of the circle */
    var CLEAR_GLYPH_INSET    = 6;            /* ×と外周の間隔 / inset of the × glyph */
    var CLEAR_STROKE_WIDTH   = 1.5;          /* 円と×の線幅 / stroke width */

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
     * 見出し付きのリスト用カラムを作る
     * @param {Group} parent - 追加先のグループ
     * @param {string} captionText - リストの見出し
     * @returns {Group} 見出しを追加済みのカラムグループ
     */
    function addListColumn(parent, captionText) {
        var column = parent.add("group");
        column.orientation = "column";
        column.alignChildren = ["fill", "top"];
        column.spacing = LIST_LABEL_SPACING;
        column.add("statictext", undefined, captionText);
        return column;
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
     * キーワードボタンの寸法を文字幅に合わせて詰める
     * @param {Button} button - 対象ボタン
     * @returns {void}
     */
    function applyPresetButtonSize(button) {
        /* 実測できればそれを使い、駄目なら文字数から概算する / Measure if possible, else estimate */
        var textWidth = String(button.text).length * PRESET_CHAR_WIDTH;
        try {
            var measured = button.graphics.measureString(button.text);
            var measuredWidth = (measured.width !== undefined) ? measured.width : measured[0];
            /* 環境によっては値が取れずNaNになる。その場合は概算のままにする / Keep the estimate if unusable */
            if (!isNaN(measuredWidth) && measuredWidth > 0) textWidth = measuredWidth;
        } catch (e) {}

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
     * クリアボタンの丸と×を描く
     * @param {Button} button - 対象ボタン
     * @returns {void}
     */
    function drawClearButton(button) {
        var graphics = button.graphics;
        var size = button.size[0];
        var useLightUI = isLightUI();

        /* 主役ではないので、通常時も少し薄めに描く。押している間は濃くして手応えを出す / Slightly muted, darker while pressed */
        var glyphColor = useLightUI ? [0.45, 0.45, 0.45, 1] : [0.72, 0.72, 0.72, 1];
        if (button.pressed) glyphColor = useLightUI ? [0.20, 0.20, 0.20, 1] : [0.95, 0.95, 0.95, 1];

        /* キーワードが空のときは無効にしてあるので、さらに淡くする / Dim further while disabled */
        if (button.enabled === false) glyphColor = useLightUI ? [0.78, 0.78, 0.78, 1] : [0.42, 0.42, 0.42, 1];

        /* 前回の描画を消すため、まず親と同じ色で塗りつぶす / Repaint the background first */
        try {
            graphics.newPath();
            graphics.rectPath(0, 0, size, size);
            graphics.fillPath(graphics.backgroundColor);
        } catch (e) {}

        var pen = graphics.newPen(graphics.PenType.SOLID_COLOR, glyphColor, CLEAR_STROKE_WIDTH);
        var circleSize = size - CLEAR_CIRCLE_INSET * 2;
        graphics.newPath();
        graphics.ellipsePath(CLEAR_CIRCLE_INSET, CLEAR_CIRCLE_INSET, circleSize, circleSize);
        graphics.strokePath(pen);

        var glyphStart = CLEAR_GLYPH_INSET;
        var glyphEnd = size - CLEAR_GLYPH_INSET;
        graphics.newPath();
        graphics.moveTo(glyphStart, glyphStart);
        graphics.lineTo(glyphEnd, glyphEnd);
        graphics.strokePath(pen);
        graphics.newPath();
        graphics.moveTo(glyphEnd, glyphStart);
        graphics.lineTo(glyphStart, glyphEnd);
        graphics.strokePath(pen);
    }

    /**
     * キーワードを消すクリアボタンを作る
     * @param {Group} parent - 追加先のグループ
     * @returns {Button} 丸に×を自前描画したボタン
     */
    function addClearButton(parent) {
        var button = parent.add("button", undefined, "");
        button.helpTip = getLabel(LABELS.button.clearKeyword);
        button.alignment = ["right", "center"];
        button.preferredSize = [CLEAR_BUTTON_SIZE, CLEAR_BUTTON_SIZE];
        button.minimumSize = [CLEAR_BUTTON_SIZE, CLEAR_BUTTON_SIZE];
        button.maximumSize = [CLEAR_BUTTON_SIZE, CLEAR_BUTTON_SIZE];
        button.pressed = false;

        /* キーワード欄は空の状態で開くので、ディム表示から始める / Start dimmed for the empty field */
        button.enabled = false;

        button.onDraw = function () {
            drawClearButton(this);
        };
        button.addEventListener("mousedown", function () {
            this.pressed = true;
            this.notify("onDraw");
        });
        button.addEventListener("mouseup", function () {
            this.pressed = false;
            this.notify("onDraw");
        });
        /* ボタンの外でマウスを離したときも押下表示を戻す / Reset the pressed look on mouseout */
        button.addEventListener("mouseout", function () {
            if (!this.pressed) return;
            this.pressed = false;
            this.notify("onDraw");
        });

        return button;
    }

    /**
     * テキストフィールドに↑↓キーでの値の増減を組み込む
     * ↑↓で±1、shift併用で±10（10の倍数にスナップ）、option併用で±0.1
     * @param {EditText} editText - 対象のテキストフィールド
     * @returns {void}
     */
    function changeValueByArrowKey(editText) {
        editText.addEventListener("keydown", function(event) {
            // ↑↓以外では何もしない。これがないと入力中の文字が毎回書き換わる
            if (!event || (event.keyName !== "Up" && event.keyName !== "Down")) return;

            var value = Number(editText.text);
            if (isNaN(value)) return;

            var keyboard = ScriptUI.environment.keyboardState;
            var delta = 1;

            if (keyboard.shiftKey) {
                delta = 10;
                // Shiftキー押下時は10の倍数にスナップ
                if (event.keyName === "Up") {
                    value = Math.ceil((value + 1) / delta) * delta;
                    event.preventDefault();
                } else if (event.keyName === "Down") {
                    value = Math.floor((value - 1) / delta) * delta;
                    if (value < 0) value = 0;
                    event.preventDefault();
                }
            } else if (keyboard.altKey) {
                delta = 0.1;
                // Optionキー押下時は0.1単位で増減
                if (event.keyName === "Up") {
                    value += delta;
                    event.preventDefault();
                } else if (event.keyName === "Down") {
                    value -= delta;
                    event.preventDefault();
                }
            } else {
                delta = 1;
                if (event.keyName === "Up") {
                    value += delta;
                    event.preventDefault();
                } else if (event.keyName === "Down") {
                    value -= delta;
                    if (value < 0) value = 0;
                    event.preventDefault();
                }
            }

            if (keyboard.altKey) {
                // 小数第1位までに丸め
                value = Math.round(value * 10) / 10;
            } else {
                // 整数に丸め
                value = Math.round(value);
            }

            editText.text = value;
        });
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
            title:           { ja: "スクリプトランチャー", en: "Script Launcher" },
            selectFolder:    { ja: "検索対象のスクリプトフォルダーを選択してください", en: "Select the script folder to search" },
            preferences:     { ja: "環境設定", en: "Preferences" }
        },
        panel: {
            keyword:        { ja: "絞り込み（%1件）", en: "Filter (%1)" },
            scriptFolder:   { ja: "対象フォルダー", en: "Target Folder" },
            keywordButtons: { ja: "キーワードボタン", en: "Keyword Buttons" }
        },
        listCaption: {
            folder:   { ja: "フォルダー", en: "Folder" },
            fileName: { ja: "ファイル名", en: "File Name" }
        },
        folderRow: {
            allFolders: { ja: "（すべて）", en: "(All)" },
            rootFolder: { ja: "（ルート）", en: "(Root)" }
        },
        fieldLabel: {
            keyword:    { ja: "キーワード", en: "Keyword" },
            minCount:   { ja: "出現数", en: "Occurrences" },
            maxButtons: { ja: "キーワード数", en: "Keywords" }
        },
        checkbox: {
            includeSubfolders: { ja: "サブディレクトリを含む", en: "Include subdirectories" },
            showFullPath:      { ja: "フルパス", en: "Full path" }
        },
        button: {
            changeFolder:    { ja: "フォルダー変更", en: "Change Folder" },
            clearKeyword:    { ja: "キーワードをクリア", en: "Clear keyword" },
            preferences:     { ja: "環境設定", en: "Preferences" },
            cancel:          { ja: "キャンセル", en: "Cancel" },
            run:             { ja: "実行", en: "Run" },
            ok:              { ja: "OK", en: "OK" }
        },
        alert: {
            noScripts: {
                ja: "対象フォルダー内に .jsx / .js / .jsxbin ファイルが見つかりませんでした。",
                en: "No .jsx / .js / .jsxbin files were found in the target folder."
            },
            missingScript: {
                ja: "選択したスクリプトが見つかりません。",
                en: "The selected script could not be found."
            },
            runFailed: {
                ja: "スクリプトの実行中にエラーが発生しました。\n\n%1\n\nエラー: %2",
                en: "An error occurred while running the script.\n\n%1\n\nError: %2"
            },
            errorLine: {
                ja: "\n行: %1",
                en: "\nLine: %1"
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
    // スクリプトの収集と検索キー / Script collection and search keys
    // =========================================

    /**
     * スクリプト1件分の情報
     * @typedef {object} ScriptEntry
     * @property {File} file - スクリプトファイル
     * @property {string} fileName - デコード済みのファイル名
     * @property {string} localizedFileName - OSの表示名
     * @property {string} relativePath - 対象フォルダーからの相対パス
     * @property {string} folderPath - 相対パスのフォルダー部分
     * @property {number} folderDepth - 対象フォルダーからの階層の深さ
     * @property {string} searchText - 検索対象を連結した文字列
     * @property {string} normalizedSearchText - 正規化済みの検索キー
     * @property {Array<string>} nameWords - ファイル名から取り出した語（キーワードボタン用）
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
     * 入力文字列を空白区切りの検索語に分解する
     * @param {string} value - キーワード欄の文字列
     * @returns {Array<string>} 正規化した検索語。空の語は含まない
     */
    function splitSearchTerms(value) {
        var rawTerms = String(value).split(/\s+/);
        var searchTerms = [];
        for (var i = 0; i < rawTerms.length; i++) {
            var term = normalizeSearchKey(rawTerms[i]);
            if (term !== "") searchTerms.push(term);
        }
        return searchTerms;
    }

    /**
     * 検索語をすべて含むかどうかを判定する（AND検索）
     * @param {string} searchTarget - 正規化済みの検索対象
     * @param {Array<string>} searchTerms - 正規化済みの検索語
     * @returns {boolean} すべて含むなら true
     */
    function matchesSearchTerms(searchTarget, searchTerms) {
        for (var i = 0; i < searchTerms.length; i++) {
            if (searchTarget.indexOf(searchTerms[i]) === -1) return false;
        }
        return true;
    }

    /**
     * 対象フォルダーを選ばせる
     * @param {Folder|null} currentFolder - 初期表示に使うフォルダー
     * @returns {Folder|null} 選ばれたフォルダー。取り消し時は null
     */
    function chooseScriptFolder(currentFolder) {
        var startFolder = currentFolder && currentFolder.exists ? currentFolder : Folder.myDocuments;
        return startFolder.selectDlg(getLabel(LABELS.dialog.selectFolder));
    }

    /**
     * 対象フォルダーのパスを表示用に整える
     * @param {Folder} targetFolder - 表示するフォルダー
     * @param {boolean} showFullPath - true でフルパス、false でホームフォルダーを ~ に略す
     * @returns {string} 表示用のパス文字列
     */
    function formatFolderPath(targetFolder, showFullPath) {
        var fullPath = targetFolder.fsName;
        if (showFullPath) return fullPath;

        var homePath = Folder("~").fsName;
        if (homePath && fullPath.indexOf(homePath) === 0) {
            /* 区切りの手前で切れているか確かめる。/Users/tak が /Users/takano に一致しないように */
            var rest = fullPath.substring(homePath.length);
            if (rest === "" || rest.charAt(0) === "/" || rest.charAt(0) === "\\") return "~" + rest;
        }
        return fullPath;
    }

    /**
     * decodeURI に失敗しても元の文字列を返す
     * @param {string} value - デコードする文字列
     * @returns {string} デコード結果。失敗時は元の文字列
     */
    function decodeSafely(value) {
        try {
            return decodeURI(value);
        } catch (e) {
            return value;
        }
    }

    /**
     * フォルダーを再帰的にたどってスクリプトを集める
     * @param {Folder} targetFolder - 走査するフォルダー
     * @param {Array<ScriptEntry>} collected - 収集結果の追加先
     * @param {Folder} rootFolder - 相対パスの基準になる対象フォルダー
     * @returns {void}
     */
    function collectScriptFiles(targetFolder, collected, rootFolder) {
        var entries;
        try {
            entries = targetFolder.getFiles();
        } catch (e) {
            return;
        }

        for (var i = 0; i < entries.length; i++) {
            var entry = entries[i];

            if (entry instanceof Folder) {
                /* エイリアスは循環の元になるのでたどらない / Aliases can loop back into an ancestor */
                if (!entry.alias && !/^\./.test(entry.name)) collectScriptFiles(entry, collected, rootFolder);
                continue;
            }

            if (!(entry instanceof File) || !SCRIPT_EXT_RE.test(entry.name)) continue;

            /* 対象フォルダー内にこのランチャー自身がある場合は一覧に載せない / Skip this launcher itself */
            if (File($.fileName).fsName === entry.fsName) continue;

            var relativePath = entry.fsName;
            var rootPath = rootFolder.fsName;
            if (relativePath.indexOf(rootPath) === 0) {
                relativePath = relativePath.substring(rootPath.length).replace(/^[\\\/]+/, "");
            }

            var decodedFileName = decodeSafely(entry.name);
            var localizedFileName = entry.displayName || decodedFileName;
            var decodedRelativePath = decodeSafely(relativePath);

            /* リストは左右2本なので、フォルダー部分とファイル名を分けて持つ / Split folder and file name for the two lists */
            var decodedFolderPath = decodedRelativePath.replace(/[\\\/][^\\\/]*$/, "");
            if (decodedFolderPath === decodedRelativePath) decodedFolderPath = "";

            /* 対象フォルダーからの階層の深さ。直下は0、artboard/backup は2 / Depth below the root folder */
            var folderDepth = (decodedFolderPath === "") ? 0 : decodedFolderPath.split(/[\\\/]/).length;

            collected.push({
                file: entry,
                fileName: decodedFileName,
                localizedFileName: localizedFileName,
                relativePath: decodedRelativePath,
                folderPath: decodedFolderPath,
                folderDepth: folderDepth,
                searchText: decodedRelativePath + " " + localizedFileName
            });
        }
    }

    /**
     * 相対パス順にスクリプトを並べ替える
     * @param {Array<ScriptEntry>} scriptEntries - 並べ替える配列（破壊的に変更する）
     * @returns {void}
     */
    function sortScriptEntries(scriptEntries) {
        scriptEntries.sort(function (entryA, entryB) {
            var pathA = entryA.relativePath.toLowerCase();
            var pathB = entryB.relativePath.toLowerCase();
            return pathA < pathB ? -1 : (pathA > pathB ? 1 : 0);
        });
    }

    /**
     * 入力文字列を1以上の整数として読み取る
     * @param {string} value - 入力文字列
     * @param {number} fallbackValue - 数値として読めないときに返す値
     * @param {number} [maxValue] - 上限。超えた場合はこの値に丸める
     * @returns {number} 1以上の整数
     */
    function parsePositiveInt(value, fallbackValue, maxValue) {
        var parsed = parseInt(trimWhitespace(value), 10);
        if (isNaN(parsed) || parsed < 1) return fallbackValue;
        if (typeof maxValue === "number" && parsed > maxValue) return maxValue;
        return parsed;
    }

    /**
     * ボタンにしない語かどうかを判定する
     * @param {string} word - 小文字化した語
     * @returns {boolean} 除外語なら true
     */
    function isStopWord(word) {
        for (var i = 0; i < KEYWORD_STOP_WORDS.length; i++) {
            if (KEYWORD_STOP_WORDS[i] === word) return true;
        }
        return false;
    }

    /**
     * ファイル名を単語に分解する（キャメルケース・ハイフン・アンダースコア区切り）
     * @param {string} fileName - 拡張子付きのファイル名
     * @returns {Array<string>} 小文字化した単語の配列
     */
    function extractNameWords(fileName) {
        var baseName = String(fileName).replace(/\.[^.]+$/, "");

        /* 区切り文字と数字を空白にし、続けてキャメルケースの境目を空ける / Split separators, then camel case */
        var separated = baseName.replace(/[-_\s\d]+/g, " ");
        separated = separated.replace(/([a-z])([A-Z])/g, "$1 $2");
        separated = separated.replace(/([A-Z]+)([A-Z][a-z])/g, "$1 $2");

        var words = [];
        var parts = separated.split(" ");
        for (var i = 0; i < parts.length; i++) {
            var word = parts[i].toLowerCase();
            /* 英字だけの語に限る。日本語名や記号混じりはボタンにしない / ASCII words only */
            if (!/^[a-z]+$/.test(word)) continue;
            if (word.length < KEYWORD_MIN_WORD_LENGTH) continue;
            if (isStopWord(word)) continue;
            words.push(word);
        }
        return words;
    }

    /**
     * すでに入力済みの検索語で覆われている語かどうかを判定する
     * @param {string} word - 判定する語
     * @param {Array<string>} searchTerms - 正規化済みの検索語
     * @returns {boolean} いずれかの検索語を含むなら true
     */
    function isCoveredByTerms(word, searchTerms) {
        for (var i = 0; i < searchTerms.length; i++) {
            if (word.indexOf(searchTerms[i]) !== -1) return true;
        }
        return false;
    }

    /**
     * スクリプト名によく出てくる語を求める
     * @param {Array<ScriptEntry>} scriptEntries - 集計対象のスクリプト
     * @param {number} minCount - ボタンにする最小出現ファイル数
     * @param {number} maxButtons - 返す語の最大個数
     * @param {Array<string>} searchTerms - 入力済みの検索語。これを含む語は除く
     * @returns {Array<string>} 出現ファイル数の多い順に並べた語
     */
    function collectFrequentWords(scriptEntries, minCount, maxButtons, searchTerms) {
        var wordCounts = {};
        for (var i = 0; i < scriptEntries.length; i++) {
            var words = scriptEntries[i].nameWords;

            /* 同じ語が1つのファイル名に複数あっても1件と数える / Count each file once per word */
            var seenWords = {};
            for (var j = 0; j < words.length; j++) {
                var wordKey = "#" + words[j];
                if (seenWords[wordKey]) continue;
                seenWords[wordKey] = true;
                wordCounts[wordKey] = (wordCounts[wordKey] || 0) + 1;
            }
        }

        var frequentWords = [];
        for (var countKey in wordCounts) {
            if (!wordCounts.hasOwnProperty(countKey)) continue;
            if (wordCounts[countKey] < minCount) continue;

            /* 入力済みの語をボタンにしても絞り込めないので外す / A word the query already covers is a no-op */
            var word = countKey.substring(1);
            if (isCoveredByTerms(word, searchTerms)) continue;
            frequentWords.push({ word: word, count: wordCounts[countKey] });
        }

        frequentWords.sort(function (wordA, wordB) {
            if (wordA.count !== wordB.count) return wordB.count - wordA.count;
            return wordA.word < wordB.word ? -1 : (wordA.word > wordB.word ? 1 : 0);
        });

        var presetWords = [];
        for (var k = 0; k < frequentWords.length && k < maxButtons; k++) {
            presetWords.push(frequentWords[k].word);
        }
        return presetWords;
    }

    /**
     * 語の先頭を大文字にする（ボタン表示用）
     * @param {string} word - 小文字の語
     * @returns {string} 先頭だけ大文字にした語
     */
    function capitalizeWord(word) {
        return word.charAt(0).toUpperCase() + word.substring(1);
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * キーワードボタン置き場の高さを固定する（絞り込みのたびにダイアログが伸縮しないようにする）
     * @param {Group} container - ボタン行を入れるグループ
     * @param {number} maxButtons - 並べうるボタンの最大個数
     * @returns {void}
     */
    function reserveKeywordPresetHeight(container, maxButtons) {
        var rowCount = Math.ceil(maxButtons / KEYWORD_PRESETS_PER_ROW);
        var height = rowCount * PRESET_BUTTON_HEIGHT + (rowCount - 1) * DENSE_SPACING;
        container.preferredSize.height = height;
        container.minimumSize.height = height;
    }

    /**
     * 項目名付きの数値入力欄を1行追加する
     * @param {Window} parent - 追加先のウィンドウ
     * @param {{ja: string, en: string}} labelSet - 項目名のラベル定義
     * @param {number} value - 初期値
     * @returns {EditText} 追加した数値入力欄
     */
    function addNumberField(parent, labelSet, value) {
        var row = parent.add("group");
        setupRow(row, "left", DENSE_SPACING);

        var fieldLabel = row.add("statictext", undefined, labelText(labelSet));
        fieldLabel.preferredSize.width = SETTINGS_LABEL_WIDTH;

        var input = row.add("edittext", undefined, String(value));
        input.preferredSize.width = SETTINGS_INPUT_WIDTH;
        changeValueByArrowKey(input);
        return input;
    }

    /**
     * 環境設定（キーワードボタンの出現条件）を編集するダイアログを表示する
     * @param {Folder} targetFolder - 現在の対象フォルダー
     * @param {number} minCount - 現在の出現数
     * @param {number} maxButtons - 現在のキーワード数
     * @returns {{folder: Folder, minCount: number, maxButtons: number}|null} 入力された値。取り消し時は null
     */
    function showPreferencesDialog(targetFolder, minCount, maxButtons) {
        var settingsDialog = new Window("dialog", getLabel(LABELS.dialog.preferences));
        setupWindow(settingsDialog, DENSE_SPACING);

        var selectedFolder = targetFolder;
        var folderUI = buildScriptFolderPanel(settingsDialog, selectedFolder);
        folderUI.changeButton.onClick = function () {
            var pickedFolder = chooseScriptFolder(selectedFolder);
            if (!pickedFolder) return;
            selectedFolder = pickedFolder;
            folderUI.refreshPath(selectedFolder);
        };

        var keywordButtonPanel = settingsDialog.add("panel", undefined, getLabel(LABELS.panel.keywordButtons));
        setupPanel(keywordButtonPanel, DENSE_SPACING);
        var minCountInput = addNumberField(keywordButtonPanel, LABELS.fieldLabel.minCount, minCount);
        var maxButtonsInput = addNumberField(keywordButtonPanel, LABELS.fieldLabel.maxButtons, maxButtons);

        var settingsButtonRow = settingsDialog.add("group");
        setupRow(settingsButtonRow, "right", DENSE_SPACING);
        settingsButtonRow.margins = [0, BUTTON_ROW_TOP_MARGIN, 0, 0];
        var settingsCancelButton = settingsButtonRow.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });
        var settingsOkButton = settingsButtonRow.add("button", undefined, getLabel(LABELS.button.ok), { name: "ok" });
        applyButtonSize(settingsCancelButton, DIALOG_BUTTON_WIDTH);
        applyButtonSize(settingsOkButton, DIALOG_BUTTON_WIDTH);

        settingsDialog.center();
        minCountInput.active = true;
        if (settingsDialog.show() !== 1) return null;

        return {
            folder: selectedFolder,
            minCount: parsePositiveInt(minCountInput.text, minCount),
            maxButtons: parsePositiveInt(maxButtonsInput.text, maxButtons, KEYWORD_PRESET_LIMIT)
        };
    }

    /**
     * 対象フォルダーのスクリプトを集め、検索キーまで用意する
     * @param {Folder} targetFolder - 検索対象のフォルダー
     * @returns {Array<ScriptEntry>} 相対パス順に並べ、検索キーを付けたスクリプト
     */
    function loadScriptEntries(targetFolder) {
        var scriptEntries = [];
        collectScriptFiles(targetFolder, scriptEntries, targetFolder);
        sortScriptEntries(scriptEntries);

        /* 検索キーは起動時に一度だけ作ってキャッシュし、入力のたびの再計算を避ける / Cache search keys once */
        for (var i = 0; i < scriptEntries.length; i++) {
            scriptEntries[i].normalizedSearchText = normalizeSearchKey(scriptEntries[i].searchText);
            scriptEntries[i].nameWords = extractNameWords(scriptEntries[i].fileName);
        }
        return scriptEntries;
    }

    /**
     * 絞り込みパネルを組み立てる
     * @param {Window} parent - 追加先のウィンドウ
     * @returns {{panel: Panel, input: EditText, clearButton: Button, presetContainer: Group}} パネル・入力欄・クリアボタン・ボタン置き場
     */
    function buildKeywordPanel(parent) {
        var panel = parent.add("panel", undefined, formatLabel(getLabel(LABELS.panel.keyword), [0]));
        setupPanel(panel, DENSE_SPACING);

        var row = panel.add("group");
        setupRow(row, "fill", DENSE_SPACING);
        row.add("statictext", undefined, labelText(LABELS.fieldLabel.keyword));

        var input = row.add("edittext", undefined, "");
        input.alignment = ["fill", "center"];

        var clearButton = addClearButton(row);

        /* ボタンは絞り込みのたびに作り直すので、置き場だけ先に用意する / Reserve the area for the preset buttons */
        var presetContainer = panel.add("group");
        presetContainer.orientation = "column";
        presetContainer.alignChildren = ["left", "top"];
        presetContainer.alignment = ["fill", "top"];
        presetContainer.spacing = DENSE_SPACING;

        return { panel: panel, input: input, clearButton: clearButton, presetContainer: presetContainer };
    }

    /**
     * 対象フォルダーパネルを組み立てる
     * @param {Window} parent - 追加先のウィンドウ
     * @param {Folder} targetFolder - 最初に表示するフォルダー
     * @returns {{changeButton: Button, refreshPath: function(Folder): void}} 変更ボタンと表示更新関数
     */
    function buildScriptFolderPanel(parent, targetFolder) {
        var currentFolder = targetFolder;

        var panel = parent.add("panel", undefined, getLabel(LABELS.panel.scriptFolder));
        setupPanel(panel, DENSE_SPACING);

        var row = panel.add("group");
        setupRow(row, "fill", DENSE_SPACING);

        /* 入力欄に見せないよう statictext で表示し、長いパスは中央を省略する / Plain text, truncated in the middle */
        var pathText = row.add("statictext", undefined, formatFolderPath(currentFolder, SHOW_FULL_PATH_DEFAULT), { truncate: "middle" });
        pathText.preferredSize.width = FOLDER_PATH_WIDTH;

        var changeButton = row.add("button", undefined, getLabel(LABELS.button.changeFolder));
        changeButton.alignment = ["right", "center"];
        applyButtonSize(changeButton, FOLDER_BUTTON_WIDTH);

        var fullPathCheckbox = panel.add("checkbox", undefined, getLabel(LABELS.checkbox.showFullPath));
        fullPathCheckbox.alignment = "left";
        fullPathCheckbox.value = SHOW_FULL_PATH_DEFAULT;

        /**
         * パス表示を今のフォルダーと「フルパス」の状態に合わせる
         * @param {Folder} [folder] - 新しいフォルダー。省略時は表示だけ更新する
         * @returns {void}
         */
        function refreshPath(folder) {
            if (folder) currentFolder = folder;
            pathText.text = formatFolderPath(currentFolder, fullPathCheckbox.value);
        }

        fullPathCheckbox.onClick = function () { refreshPath(); };
        return { changeButton: changeButton, refreshPath: refreshPath };
    }

    /**
     * 左右2本のリストを組み立てる
     * @param {Window} parent - 追加先のウィンドウ
     * @returns {{folderListBox: ListBox, subfoldersCheckbox: Checkbox, scriptListBox: ListBox}} リスト部品
     */
    function buildListColumns(parent) {
        /* 左でフォルダーを選び、右にそのフォルダー内のファイル名だけを並べる / Folder on the left, file names on the right */
        var row = parent.add("group");
        setupRow(row, "fill", COLUMN_SPACING);
        row.alignChildren = ["fill", "fill"];

        var folderColumn = addListColumn(row, getLabel(LABELS.listCaption.folder));
        var folderListBox = folderColumn.add("listbox", undefined, [], { multiselect: false });
        folderListBox.preferredSize = FOLDER_LIST_SIZE;

        var checkboxRow = folderColumn.add("group");
        setupRow(checkboxRow, "left", 0);
        checkboxRow.margins = [0, CHECKBOX_TOP_MARGIN, 0, 0];
        var subfoldersCheckbox = checkboxRow.add("checkbox", undefined, getLabel(LABELS.checkbox.includeSubfolders));
        subfoldersCheckbox.value = INCLUDE_SUBFOLDERS_DEFAULT;

        var scriptColumn = addListColumn(row, getLabel(LABELS.listCaption.fileName));
        var scriptListBox = scriptColumn.add("listbox", undefined, [], { multiselect: false });
        scriptListBox.preferredSize = SCRIPT_LIST_SIZE;

        return { folderListBox: folderListBox, subfoldersCheckbox: subfoldersCheckbox, scriptListBox: scriptListBox };
    }

    /**
     * ダイアログ下部のボタン列を組み立てる
     * @param {Window} parent - 追加先のウィンドウ
     * @returns {{settings: Button, cancel: Button, run: Button}} 3つのボタン
     */
    function buildDialogButtons(parent) {
        /* メイングループ（横並び） / Main group (horizontal layout) */
        var btnRowGroup = parent.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.margins = [0, BUTTON_ROW_TOP_MARGIN, 0, 0];
        btnRowGroup.alignment = ["fill", "bottom"];

        /* 左側グループ / Left-side button group */
        var btnLeftGroup = btnRowGroup.add("group");
        btnLeftGroup.alignChildren = ["left", "center"];
        var btnPreferences = btnLeftGroup.add("button", undefined, getLabel(LABELS.button.preferences));

        /* スペーサー（伸縮）/ Spacer (stretchable) */
        var spacer = btnRowGroup.add("group");
        spacer.alignment = ["fill", "fill"];
        spacer.minimumSize.width = 0;

        /* 右側グループ / Right-side button group */
        var btnRightGroup = btnRowGroup.add("group");
        btnRightGroup.alignChildren = ["right", "center"];
        var btnCancel = btnRightGroup.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });
        var btnRun = btnRightGroup.add("button", undefined, getLabel(LABELS.button.run), { name: "ok" });
        btnRun.enabled = false;

        applyButtonSize(btnPreferences, SETTINGS_BUTTON_WIDTH);
        applyButtonSize(btnCancel, DIALOG_BUTTON_WIDTH);
        applyButtonSize(btnRun, DIALOG_BUTTON_WIDTH);

        return { settings: btnPreferences, cancel: btnCancel, run: btnRun };
    }

    /**
     * ランチャーのダイアログを表示する
     * @param {Folder} targetFolder - 検索対象のフォルダー
     * @returns {{action: string, file: File|null, folder: Folder|null}} 操作結果（"run" / "changeFolder" / "cancel"）
     */
    function showLauncherDialog(targetFolder) {
        var scriptEntries = loadScriptEntries(targetFolder);
        if (scriptEntries.length === 0) {
            alert(getLabel(LABELS.alert.noScripts), getLabel(LABELS.dialog.title));
            return { action: "changeFolder", file: null, folder: null };
        }

        var launcherDialog = new Window("dialog", getLabel(LABELS.dialog.title) + " " + SCRIPT_VERSION);
        setupWindow(launcherDialog);

        /* 検索欄を最初の操作部品にして、起動時のフォーカスを安定させる / Keep the keyword field first */
        var keywordUI = buildKeywordPanel(launcherDialog);
        var keywordPanel = keywordUI.panel;
        var keywordInput = keywordUI.input;
        var keywordClearButton = keywordUI.clearButton;
        var keywordPresetContainer = keywordUI.presetContainer;

        /* 絞り込み結果によく出る語をワンクリックで入れる / One-click presets from the filtered results */
        var storedSettings = readKeywordSettings();
        var presetMinCount = storedSettings.minCount;
        var presetMaxButtons = storedSettings.maxButtons;
        var keywordPresetRows = [];
        var currentPresetWords = null;
        var isDialogShown = false;
        reserveKeywordPresetHeight(keywordPresetContainer, presetMaxButtons);

        var listUI = buildListColumns(launcherDialog);
        var folderListBox = listUI.folderListBox;
        var includeSubfoldersCheckbox = listUI.subfoldersCheckbox;
        var scriptListBox = listUI.scriptListBox;

        var dialogButtons = buildDialogButtons(launcherDialog);
        var btnPreferences = dialogButtons.settings;
        var btnCancel = dialogButtons.cancel;
        var btnRun = dialogButtons.run;

        var filteredScripts = [];
        var listedScripts = [];
        var dialogResult = { action: "cancel", file: null, folder: null };

        /* 組み直し中の選択変更でファイル名リストが何度も再構築されるのを防ぐ / Suppress cascaded rebuilds */
        var isRebuildingFolderList = false;

        /**
         * 絞り込み結果に合わせてキーワードボタンを作り直す
         * @param {Array<ScriptEntry>} entries - 集計対象のスクリプト（現在の絞り込み結果）
         * @param {Array<string>} searchTerms - 入力済みの検索語。これを含む語はボタンにしない
         * @returns {void}
         */
        function refreshKeywordPresetButtons(entries, searchTerms) {
            var presetWords = collectFrequentWords(entries, presetMinCount, presetMaxButtons, searchTerms);

            /* 顔ぶれが変わらないなら作り直さない。1打鍵ごとの再構築を避ける / Skip the rebuild when nothing changed */
            var presetWordsKey = presetWords.join(" ");
            if (presetWordsKey === currentPresetWords) return;
            currentPresetWords = presetWordsKey;

            for (var i = 0; i < keywordPresetRows.length; i++) {
                keywordPresetContainer.remove(keywordPresetRows[i]);
            }
            keywordPresetRows = [];

            var presetRow = null;
            for (var j = 0; j < presetWords.length; j++) {
                if (j % KEYWORD_PRESETS_PER_ROW === 0) {
                    presetRow = keywordPresetContainer.add("group");
                    setupRow(presetRow, "left", DENSE_SPACING);
                    keywordPresetRows.push(presetRow);
                }
                var presetButton = presetRow.add("button", undefined, capitalizeWord(presetWords[j]));
                applyPresetButtonSize(presetButton);
                presetButton.onClick = makeKeywordPresetHandler(presetWords[j]);
            }

            /* 置き場の高さは固定してあるので、ここだけ組み直せばダイアログは動かない / Fixed height keeps the dialog still */
            if (isDialogShown) keywordPresetContainer.layout.layout(true);
        }

        /**
         * プリセットボタン用の onClick ハンドラーを作る
         * @param {string} presetKeyword - ボタンに割り当てるキーワード
         * @returns {function(): void} 通常クリックで置き換え、option+クリックで追加するハンドラー
         */
        function makeKeywordPresetHandler(presetKeyword) {
            return function () {
                /* option+クリックは空白を挟んで語を足し、AND検索で絞り込む / Option-click appends the word */
                var currentText = trimWhitespace(keywordInput.text);
                var isAppend = ScriptUI.environment.keyboardState.altKey && currentText !== "";
                keywordInput.text = isAppend ? currentText + " " + presetKeyword : presetKeyword;
                lastNormalizedQuery = splitSearchTerms(keywordInput.text).join(" ");

                /* 押されたボタンは組み直しで消える。先にフォーカスを入力欄へ移しておく */
                keywordInput.active = true;
                updateClearButtonState();
                refreshFolderList();
            };
        }

        /**
         * 選択中フォルダーの相対パスを返す
         * @returns {string|null} 相対パス。「すべて」選択時と未選択時は null
         */
        function selectedFolderPath() {
            if (!folderListBox.selection) return null;
            var folderPath = folderListBox.selection.folderPath;
            return (folderPath === undefined) ? null : folderPath;
        }

        /**
         * キーワードとサブディレクトリ設定で絞り込み、左のフォルダーリストを組み直す
         * @returns {void}
         */
        function refreshFolderList() {
            var searchTerms = splitSearchTerms(keywordInput.text);
            var previousFolderPath = selectedFolderPath();
            var depthLimit = includeSubfoldersCheckbox.value ? null : NESTED_FOLDER_DEPTH_LIMIT;
            isRebuildingFolderList = true;
            filteredScripts = [];

            var seenFolderPaths = {};
            var folderPaths = [];
            for (var i = 0; i < scriptEntries.length; i++) {
                var searchTarget = scriptEntries[i].normalizedSearchText;
                if (!matchesSearchTerms(searchTarget, searchTerms)) continue;
                if (depthLimit !== null && scriptEntries[i].folderDepth > depthLimit) continue;

                filteredScripts.push(scriptEntries[i]);

                /* ハッシュのキー衝突を避けるため接頭辞を付けて既出判定する / Prefix the key to avoid collisions */
                var folderKey = "#" + scriptEntries[i].folderPath;
                if (!seenFolderPaths[folderKey]) {
                    seenFolderPaths[folderKey] = true;
                    folderPaths.push(scriptEntries[i].folderPath);
                }
            }
            folderPaths.sort();
            refreshKeywordPresetButtons(filteredScripts, searchTerms);

            folderListBox.removeAll();
            folderListBox.add("item", getLabel(LABELS.folderRow.allFolders));
            for (var j = 0; j < folderPaths.length; j++) {
                var folderItem = folderListBox.add("item", folderPaths[j] === "" ? getLabel(LABELS.folderRow.rootFolder) : folderPaths[j]);
                folderItem.folderPath = folderPaths[j];
            }

            /* 絞り込み前に選んでいたフォルダーが残っていれば選択を引き継ぐ / Keep the previous folder selection */
            folderListBox.selection = 0;
            if (previousFolderPath !== null) {
                for (var k = 1; k < folderListBox.items.length; k++) {
                    if (folderListBox.items[k].folderPath === previousFolderPath) {
                        folderListBox.selection = k;
                        break;
                    }
                }
            }

            isRebuildingFolderList = false;
            refreshScriptList();
        }

        /**
         * 選択中フォルダーに合わせて右のファイル名リストを組み直す
         * @returns {void}
         */
        function refreshScriptList() {
            var folderPath = selectedFolderPath();
            scriptListBox.removeAll();
            listedScripts = [];

            for (var i = 0; i < filteredScripts.length; i++) {
                if (folderPath !== null && filteredScripts[i].folderPath !== folderPath) continue;
                listedScripts.push(filteredScripts[i]);
                var scriptItem = scriptListBox.add("item", filteredScripts[i].fileName);
                scriptItem.scriptIndex = listedScripts.length - 1;
            }

            /* 件数はパネルのタイトルに出す / Show the match count in the panel title */
            keywordPanel.text = formatLabel(getLabel(LABELS.panel.keyword), [listedScripts.length]);
            if (scriptListBox.items.length > 0) {
                scriptListBox.selection = 0;
                btnRun.enabled = true;
            } else {
                btnRun.enabled = false;
            }
        }

        /**
         * 右のリストで選択中のスクリプトを返す
         * @returns {ScriptEntry|null} 選択中のスクリプト。未選択なら null
         */
        function selectedScriptEntry() {
            if (!scriptListBox.selection) return null;
            var selectedIndex = scriptListBox.selection.scriptIndex;
            if (selectedIndex === undefined) return null;
            return listedScripts[selectedIndex] || null;
        }

        /**
         * 選択中のスクリプトを実行対象に確定してダイアログを閉じる
         * @returns {void}
         */
        function runSelectedScript() {
            var selectedScript = selectedScriptEntry();
            if (!selectedScript) return;
            dialogResult.action = "run";
            dialogResult.file = selectedScript.file;
            launcherDialog.close(1);
        }

        /**
         * 左のリストで選択中のフォルダーをFinderで開く（ダイアログは開いたまま）
         * 「すべて」と「ルート」は対象フォルダーそのものを開く
         * @returns {void}
         */
        function openSelectedFolder() {
            var folderPath = selectedFolderPath();
            var folderToOpen = targetFolder;

            /* パス文字列を組み立てず、そのフォルダーにあるスクリプトの親を使う / Use a real Folder object */
            if (folderPath) {
                for (var i = 0; i < filteredScripts.length; i++) {
                    if (filteredScripts[i].folderPath === folderPath) {
                        folderToOpen = filteredScripts[i].file.parent;
                        break;
                    }
                }
            }
            folderToOpen.execute();
        }

        /**
         * 選択中のスクリプトをFinderで表示する（ダイアログは開いたまま）
         * @returns {void}
         */
        function revealSelectedScript() {
            var selectedScript = selectedScriptEntry();
            if (selectedScript) revealScriptFile(selectedScript.file);
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

        keywordInput.onChanging = handleKeywordChanged;
        keywordInput.onChange = handleKeywordChanged;
        keywordClearButton.onClick = function () {
            keywordInput.text = "";
            /* 続けて打ち直せるよう、フォーカスは入力欄へ戻す / Put the focus back in the field */
            keywordInput.active = true;
            handleKeywordChanged();
        };
        btnPreferences.onClick = function () {
            var settings = showPreferencesDialog(targetFolder, presetMinCount, presetMaxButtons);
            if (!settings) return;
            presetMinCount = settings.minCount;
            presetMaxButtons = settings.maxButtons;
            saveKeywordSettings(presetMinCount, presetMaxButtons);

            /* フォルダーが変わったら一覧を作り直すため、ダイアログを開き直す / Reopen the dialog to rescan */
            if (settings.folder.fsName !== targetFolder.fsName) {
                dialogResult.action = "changeFolder";
                dialogResult.folder = settings.folder;
                launcherDialog.close(2);
                return;
            }

            /* 置き場の高さが変わるので、ダイアログ全体を組み直す / The reserved height changes, so re-layout the dialog */
            reserveKeywordPresetHeight(keywordPresetContainer, presetMaxButtons);
            currentPresetWords = null;
            refreshFolderList();
            launcherDialog.layout.layout(true);
        };
        includeSubfoldersCheckbox.onClick = refreshFolderList;
        folderListBox.onChange = function () {
            if (isRebuildingFolderList) return;
            refreshScriptList();
        };
        scriptListBox.onChange = function () {
            btnRun.enabled = !!scriptListBox.selection;
        };
        folderListBox.onDoubleClick = openSelectedFolder;
        scriptListBox.onDoubleClick = function () {
            /* option+ダブルクリックは実行せずFinderで場所を開く / Option-double-click reveals instead of running */
            if (ScriptUI.environment.keyboardState.altKey) {
                revealSelectedScript();
                return;
            }
            runSelectedScript();
        };
        btnRun.onClick = runSelectedScript;
        btnCancel.onClick = function () {
            dialogResult.action = "cancel";
            launcherDialog.close();
        };
        keywordInput.addEventListener("keydown", function (event) {
            if (event.keyName === "Down" && scriptListBox.items.length > 0) {
                scriptListBox.active = true;
                if (!scriptListBox.selection) scriptListBox.selection = 0;
                event.preventDefault();
            } else if (event.keyName === "Enter" || event.keyName === "Return") {
                /* 絞り込みは onChanging / onChange で済んでいる。ここで組み直すと選択が先頭へ戻る */
                runSelectedScript();
                event.preventDefault();
                if (event.stopPropagation) event.stopPropagation();
            }
        });

        scriptListBox.addEventListener("keydown", function (event) {
            if (event.keyName === "Enter" || event.keyName === "Return") {
                runSelectedScript();
                event.preventDefault();
                if (event.stopPropagation) event.stopPropagation();
            }
        });

        /* テンキーEnterを含め、ダイアログ内のどこにフォーカスがあっても実行する / Enter runs from anywhere */
        launcherDialog.addEventListener("keydown", function (event) {
            /* ボタンにフォーカスがあるときは、そのボタン自身の動作に任せる */
            if (event.target && event.target.type === "button") return;
            if (event.keyName === "Enter" || event.keyName === "Return") {
                runSelectedScript();
                event.preventDefault();
            }
        });

        refreshFolderList();
        lastNormalizedQuery = splitSearchTerms(keywordInput.text).join(" ");
        launcherDialog.center();

        /* 検索欄を先頭の操作部品にしたうえで、表示時にも明示的にフォーカスする / Focus the keyword field on show */
        keywordInput.active = true;
        launcherDialog.onShow = function () {
            keywordInput.active = true;
            keywordInput.selection = [0, 0];
        };

        isDialogShown = true;
        launcherDialog.show();
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
     * @param {File} scriptFile - 対象のスクリプトファイル
     * @returns {boolean} 書き出せたら true
     */
    function writeRevealPath(scriptFile) {
        var pathFile = new File(REVEAL_PATH_FILE);
        try {
            pathFile.encoding = "UTF-8";
            pathFile.lineFeed = "Unix";
            if (!pathFile.open("w")) return false;

            /* fsName で ~ ではなく絶対パスを渡す / fsName gives the absolute POSIX path */
            pathFile.write(scriptFile.fsName);
            return true;
        } catch (e) {
            return false;
        } finally {
            try { pathFile.close(); } catch (e) {}
        }
    }

    /**
     * スクリプトの場所をFinderで開く
     * @param {File} scriptFile - 対象のスクリプトファイル
     * @returns {void}
     */
    function revealScriptFile(scriptFile) {
        /* アプリが無い環境では囲みフォルダーを開くだけにとどめる / Fall back to the enclosing folder */
        var revealApp = findRevealApp();
        if (!revealApp) {
            scriptFile.parent.execute();
            return;
        }

        /* アプリは一時ファイルからパスを読むので、書き出してから起動する / The app reads the path from a temp file */
        if (writeRevealPath(scriptFile) && revealApp.execute()) return;

        /* 起動できなかった場合も何も起きないままにはしない / Never end up doing nothing */
        scriptFile.parent.execute();
    }

    /**
     * 選ばれたスクリプトを実行する
     * @param {File} scriptFile - 実行するスクリプトファイル
     * @returns {void}
     */
    function runScriptFile(scriptFile) {
        if (!scriptFile || !scriptFile.exists) {
            alert(getLabel(LABELS.alert.missingScript), getLabel(LABELS.dialog.title));
            return;
        }

        try {
            $.evalFile(scriptFile);
        } catch (e) {
            var message = formatLabel(getLabel(LABELS.alert.runFailed), [scriptFile.fsName, e.message]);
            if (e.line) message += formatLabel(getLabel(LABELS.alert.errorLine), [e.line]);
            alert(message, getLabel(LABELS.dialog.title));
        }
    }

    /**
     * 対象フォルダーの確認からスクリプト実行までを進める
     * @returns {void}
     */
    function main() {
        var scriptFolder = readSavedScriptFolder();

        while (true) {
            if (!scriptFolder || !scriptFolder.exists) {
                scriptFolder = chooseScriptFolder(scriptFolder);
                if (!scriptFolder) return;
                saveScriptFolder(scriptFolder);
            }

            var launcherResult = showLauncherDialog(scriptFolder);

            if (launcherResult.action === "changeFolder") {
                /* 環境設定で選び直した場合はそのフォルダーを使い、見つからなかった場合だけ選ばせる */
                var changedFolder = launcherResult.folder || chooseScriptFolder(scriptFolder);

                /* 選び直しを取り消したら終了する。対象が空のまま繰り返すと抜けられなくなる */
                if (!changedFolder) return;
                scriptFolder = changedFolder;
                saveScriptFolder(scriptFolder);
                continue;
            }

            if (launcherResult.action === "run") runScriptFile(launcherResult.file);
            break;
        }
    }

    main();

})();
