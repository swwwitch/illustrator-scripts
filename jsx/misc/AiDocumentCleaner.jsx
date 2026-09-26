#target illustrator
#targetengine "AiDocumentCleanerSession"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

ドキュメント内の不要な要素（未使用のパネル項目、孤立点や空のテキスト、空のグループ・レイヤー、ガイド、アートボード外のオブジェクト、属性パネルのメモなど）をまとめて削除します。
処理対象は、最前面のドキュメント／開いているすべてのドキュメント／指定フォルダー内の .ai ファイルから選べます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AiDocumentCleaner.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n0d70178f0f65

### Overview

Removes the clutter from a document — unused panel entries, stray points and empty text, empty groups and layers, guides, objects outside the artboards, Attributes panel notes, and more.
The scope can be the frontmost document, every open document, or the .ai files in a folder you choose.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AiDocumentCleaner.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AiDocumentCleaner";            /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.2";                      /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-06-27";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-22";                  /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AiDocumentCleaner.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AiDocumentCleaner.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n0d70178f0f65"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* Dropboxのローカルマウントパス。空文字にするとホーム直下から自動検出 / Local Dropbox mount path ("" = auto detect) */
    var DROPBOX_MOUNT_PATH = "";

    /* システム管理レイヤー名（空でも削除しない）/ System-managed layer names (kept even when empty) */
    var PROTECTED_LAYER_NAMES = {
        "_guide": true,
        "_pasteboard": true
    };

    /* 初期状態でOFFにするチェックボックス（誤削除を招きやすい積極的な項目）/ Checkboxes that start unchecked (aggressive options prone to false positives) */
    var UNCHECKED_BY_DEFAULT = {
        hiddenObjects: true,
        brokenLink: true,
        notes: true,
        outsideAllArtboards: true,
        outsideActiveArtboard: true
    };

    // =========================================
    // レイアウト / Layout
    // ダイアログの見た目の寸法。ScriptUI は生成後に幅が伸びないため、事前に確保しておく値を含む
    // Dialog metrics; ScriptUI controls don't grow after creation, so some widths are reserved up front
    // =========================================

    /* パネルの余白と間隔 / Panel margins and spacing */
    var PANEL_MARGINS = [16, 20, 16, 12];
    var PANEL_SPACING = 12;

    /* 「使用中のパネル項目も削除」の上に置く区切り線の、上下の余白（px）/ Space above and below the divider that sits over the force option */
    var FORCE_DIVIDER_MARGIN = 8;

    /* 「処理対象」パネルで選択中のフォルダーパスを表示する幅と、ラジオのラベル位置に合わせる字下げ（px）
       Width of the chosen-folder path label in the Scope panel, and the indent that lines it up with the radio's label */
    var FOLDER_PATH_WIDTH = 280;
    var FOLDER_PATH_INDENT = 18;

    /* 「フォルダー指定」ラジオの幅（px）。選択後に件数を付け足してもラベルが切れないよう先に確保する
       Width of the folder radio; reserved up front so appending the file count after selection doesn't clip the label */
    var FOLDER_RADIO_WIDTH = 160;

    /* セットのポップアップの幅（px）/ Width of the preset popup */
    var PRESET_DROPDOWN_WIDTH = 140;

    /* 復元したダイアログ位置を採用するかの判定に使う、画面端からの余裕（px）。タイトルバーが画面外に出ないようにする
       Margin from a screen edge required to reuse a restored dialog position, so the title bar never lands off-screen */
    var ON_SCREEN_MARGIN = 60;

    /**
     * パネルの共通設定を適用する
     * @param {Panel} sectionPanel - 設定するパネル
     * @param {number} [spacing] - 子の間隔。省略時は PANEL_SPACING
     * @returns {void}
     */
    function setupPanel(sectionPanel, spacing) {
        sectionPanel.orientation = "column";
        sectionPanel.alignChildren = ["fill", "top"];
        sectionPanel.alignment = "fill";
        sectionPanel.margins = PANEL_MARGINS;
        sectionPanel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    // =========================================
    // セッション記憶 / Session memory
    // #targetengine で確保した永続エンジンの $.global に置くため、スクリプトを再実行しても
    // Illustrator を終了するまで保持される（環境設定には書き込まない）
    // Kept on $.global of the persistent engine declared by #targetengine, so it survives re-runs
    // until Illustrator quits. Nothing is written to the application preferences.
    // =========================================

    var SESSION_STATE_KEY = "__aiDocumentCleanerSession";
    if (typeof $.global[SESSION_STATE_KEY] === "undefined") {
        $.global[SESSION_STATE_KEY] = {
            choices: null,       /* 前回の選択内容 / the previous selection */
            folderPath: null,    /* 前回の対象フォルダー（パス文字列で保持）/ the previous target folder, kept as a path string */
            dialogLocation: null /* 前回のダイアログ位置 / the previous dialog position */
        };
    }
    var sessionState = $.global[SESSION_STATE_KEY];

    // =========================================
    // 一時アクション設定 / Temporary action settings
    // =========================================

    var ACTION_SET_NAME = "TemporaryActionSet";
    var ACTION_NAME = "TemporaryActionName";
    /* 一時ファイルはホーム直下ではなく OS の一時フォルダへ / Write the temp file to the OS temp folder, not the home directory */
    var ACTION_FILE_NAME = Folder.temp.fsName + "/AiDocumentCleaner_TemporaryAction.aia";

    /* パネルごとの録画値（internalName・Select All Unused 値・Delete 値・各ラベルの16進）
       localizedNameHex は日本語UIのパネル名を録画したもの。英語版など他言語UIでの再生は未検証で、
       失敗した場合は完了メッセージに「アクションを実行できませんでした」と出る（強制ONならDOM側の削除は動く）
       Recorded per-panel values (internalName, Select All Unused value, Delete value, label hex).
       localizedNameHex holds the panel names as recorded on a Japanese UI; playback on other UI languages is
       unverified and, if it fails, the summary reports the action as unplayable (force mode still works via the DOM) */
    var PRUNE_SPECS = {
        swatch: {
            internalName: "ai_plugin_swatches",
            localizedNameHex: "e382b9e382a6e382a9e38383e38381",
            selectValue: 11,
            deleteValue: 3,
            deleteNameHex: "44656c65746520537761746368"
        },
        graphicstyle: {
            internalName: "ai_plugin_styles",
            localizedNameHex: "e382b0e383a9e38395e382a3e38383e382afe382b9e382bfe382a4e383ab",
            selectValue: 14,
            deleteValue: 3,
            deleteNameHex: "44656c657465205374796c65"
        },
        symbol: {
            internalName: "ai_plugin_symbol_palette",
            localizedNameHex: "e382b7e383b3e3839ce383ab",
            selectValue: 12,
            deleteValue: 5,
            deleteNameHex: "44656c6574652053796d626f6c"
        },
        brush: {
            internalName: "ai_plugin_brush",
            localizedNameHex: "e38396e383a9e382b7",
            selectValue: 8,
            deleteValue: 3,
            deleteNameHex: "44656c657465204272757368"
        }
    };

    /* 「Select All Unused」コマンド名の16進（全パネル共通）/ Hex for the "Select All Unused" command name (shared by all panels) */
    var SELECT_ALL_UNUSED_HEX = "53656c65637420416c6c20556e75736564";

    // =========================================
    // パス表示 / Path display
    // 移植元 / Ported from: LinkedImageManagerPalette.jsx
    // findDropboxFolder / findSingleSubFolder / resolveDropboxPrefix は移植元と同じ実装を保つ
    // findDropboxFolder / findSingleSubFolder / resolveDropboxPrefix stay in step with the source
    // =========================================

    /**
     * 失敗しうる取得を試み、例外時は代替値を返す
     * @param {Function} getValue - 値を返す関数
     * @param {*} fallback - 例外時に返す値
     * @returns {*} 取得した値、または代替値
     */
    function tryGet(getValue, fallback) {
        try {
            return getValue();
        } catch (e) {
            return fallback;
        }
    }

    /**
     * ホーム直下から「Dropbox」を含むフォルダーを探す。
     * チームフォルダー（「sw Dropbox」など）を優先し、個人用の「Dropbox」は「~/Dropbox/」として短縮できるため優先度を下げる
     * @returns {Folder|null} 見つかったフォルダー。なければ null
     */
    function findDropboxFolder() {
        var homeFolder = Folder("~");
        if (!homeFolder.exists) {
            return null;
        }

        var entryList = tryGet(function() {
            return homeFolder.getFiles();
        }, []);
        var personalFolder = null;
        var teamFolder = null;

        for (var i = 0; i < entryList.length; i++) {
            var entry = entryList[i];
            if (!(entry instanceof Folder)) {
                continue;
            }

            var entryName = tryGet(function() {
                return decodeURI(String(entry.name));
            }, String(entry.name));
            if (entryName.charAt(0) === ".") {
                continue;
            }
            if (entryName.indexOf("Dropbox") === -1) {
                continue;
            }

            if (entryName === "Dropbox") {
                if (!personalFolder) {
                    personalFolder = entry;
                }
            } else if (!teamFolder) {
                teamFolder = entry;
            }
        }
        return teamFolder ? teamFolder : personalFolder;
    }

    /**
     * フォルダー直下にサブフォルダーが1つだけあるとき、そのフォルダーを返す。
     * チームDropboxのメンバーフォルダー（「takano masahiro」など）の判定に使う
     * @param {Folder} parentFolder - 探索するフォルダー
     * @returns {Folder|null} 唯一のサブフォルダー。0個または2個以上のときは null
     */
    function findSingleSubFolder(parentFolder) {
        var entryList = tryGet(function() {
            return parentFolder.getFiles();
        }, []);
        var foundFolder = null;

        for (var i = 0; i < entryList.length; i++) {
            var entry = entryList[i];
            if (!(entry instanceof Folder)) {
                continue;
            }

            var entryName = tryGet(function() {
                return decodeURI(String(entry.name));
            }, String(entry.name));
            if (entryName.charAt(0) === ".") {
                continue;
            }

            if (foundFolder) {
                return null;
            }
            foundFolder = entry;
        }
        return foundFolder;
    }

    /**
     * Dropboxのローカルマウントパスを決める。
     * 手動指定が空のときはホーム直下の「Dropbox」を含むフォルダーを探し、メンバーフォルダーが1つだけあればそこまでをプレフィックスにする
     * @param {string} manualPath - 手動で指定するパス。空文字なら自動検出
     * @returns {string} 末尾に「/」を付けたプレフィックス。見つからない場合は空文字
     */
    function resolveDropboxPrefix(manualPath) {
        if (manualPath) {
            return (manualPath.charAt(manualPath.length - 1) === "/") ? manualPath : manualPath + "/";
        }

        var dropboxFolder = findDropboxFolder();
        if (!dropboxFolder) {
            return "";
        }

        var memberFolder = findSingleSubFolder(dropboxFolder);
        return (memberFolder ? memberFolder : dropboxFolder).fsName + "/";
    }

    var DROPBOX_PREFIX = resolveDropboxPrefix(DROPBOX_MOUNT_PATH);

    /**
     * ホームフォルダー以下のパスを ~ 表記に短縮する
     * @param {string} fsPath - 短縮するパス
     * @returns {string} 短縮したパス。ホーム以下でなければそのまま
     */
    function abbreviateHomePath(fsPath) {
        var homePath = tryGet(function() {
            return Folder("~").fsName;
        }, "");
        if (!homePath) {
            return fsPath;
        }
        if (fsPath === homePath) {
            return "~";
        }
        if (fsPath.indexOf(homePath) === 0) {
            var pathAfterHome = fsPath.substring(homePath.length);
            /* 直後が区切り文字のときだけ短縮（同名の別フォルダーを誤判定しない）/ Abbreviate only on a separator boundary, so a similarly named folder isn't mistaken for the home folder */
            if (pathAfterHome.charAt(0) === "/" || pathAfterHome.charAt(0) === "\\") {
                return "~" + pathAfterHome;
            }
        }
        return fsPath;
    }

    /**
     * 表示用にパスを短縮する。Dropbox配下ならプレフィックスを落とし、そうでなければ ~ 表記にする
     * @param {string} fsPath - 表示するパス
     * @returns {string} 表示用のパス
     */
    function formatDisplayPath(fsPath) {
        if (!fsPath) {
            return fsPath;
        }
        if (DROPBOX_PREFIX && fsPath.indexOf(DROPBOX_PREFIX) === 0) {
            return fsPath.substring(DROPBOX_PREFIX.length);
        }
        return abbreviateHomePath(fsPath);
    }

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

    var LABELS = {
        dialog: {
            title: { ja: "不要な要素を一括削除", en: "Clean Up Documents" }
        },
        panel: {
            target: { ja: "処理対象", en: "Scope" },
            panelItems: { ja: "パネル項目", en: "Panel items" },
            object: { ja: "パス／オブジェクト", en: "Paths / Objects" },
            container: { ja: "グループ／レイヤー", en: "Groups / Layers" },
            guide: { ja: "ガイド", en: "Guides" },
            artboard: { ja: "アートボード", en: "Artboards" }
        },
        radio: {
            frontmostDocument: { ja: "最前面のドキュメント", en: "Frontmost document" },
            allOpenDocuments: { ja: "開いているすべてのドキュメント（{0}）", en: "All open documents ({0})" },
            targetFolder: { ja: "フォルダー指定", en: "Folder" },
            targetFolderWithCount: { ja: "フォルダー指定（{0}）", en: "Folder ({0})" },
            guidesNone: { ja: "削除しない", en: "Don't delete" },
            clearGuides: { ja: "ガイドを消去（ロック分は残る）", en: "Clear Guides (locked ones remain)" },
            guides: { ja: "すべてのガイド（ロックも解除）", en: "All guides (unlocks everything)" },
            guidesOutsideActiveArtboard: { ja: "アクティブなアートボード以外", en: "Outside the active artboard" }
        },
        checkbox: {
            swatches: { ja: "スウォッチ", en: "Swatches" },
            graphicStyles: { ja: "グラフィックスタイル", en: "Graphic styles" },
            symbols: { ja: "シンボル", en: "Symbols" },
            brushes: { ja: "ブラシ", en: "Brushes" },
            paragraphStyles: { ja: "段落スタイル", en: "Paragraph styles" },
            characterStyles: { ja: "文字スタイル", en: "Character styles" },
            strayPoints: { ja: "孤立点", en: "Stray points" },
            emptyText: { ja: "空のテキスト", en: "Empty text frames" },
            noPaintPath: { ja: "塗りも線もないパス", en: "Paths with no fill or stroke" },
            zeroOpacity: { ja: "不透明度0%のオブジェクト", en: "Objects at 0% opacity" },
            hiddenObjects: { ja: "非表示オブジェクト", en: "Hidden objects" },
            brokenLink: { ja: "リンク切れの配置画像", en: "Broken-link placed images" },
            notes: { ja: "メモ", en: "Notes" },
            outsideAllArtboards: { ja: "アートボード外のオブジェクト", en: "Objects outside all artboards" },
            outsideActiveArtboard: { ja: "アクティブなアートボード外のオブジェクト", en: "Objects outside the active artboard" },
            emptyGroup: { ja: "空のグループ", en: "Empty groups" },
            emptyLayer: { ja: "空のレイヤー／サブレイヤー", en: "Empty layers / sublayers" },
            artboards: { ja: "空のアートボード", en: "Empty artboards" },
            force: { ja: "使用中のパネル項目も削除", en: "Delete panel items even if in use" }
        },
        fieldLabel: {
            preset: { ja: "セット", en: "Preset" }
        },
        dropdown: {
            presetBasic: { ja: "基本", en: "Default" },
            presetAllOff: { ja: "すべてOFF", en: "All off" },
            presetAllOn: { ja: "すべてON", en: "All on" },
            presetPanelItemsOnly: { ja: "パネル項目のみ", en: "Panel items only" },
            presetCustom: { ja: "カスタム", en: "Custom" }
        },
        status: {
            noFolderChosen: { ja: "（未指定）", en: "(none chosen)" }
        },
        button: {
            chooseFolder: { ja: "指定...", en: "Choose..." },
            cancel: { ja: "キャンセル", en: "Cancel" },
            run: { ja: "実行", en: "Run" }
        },
        result: {
            swatches: { ja: "スウォッチ", en: "Swatches" },
            graphicStyles: { ja: "グラフィックスタイル", en: "Graphic styles" },
            symbols: { ja: "シンボル", en: "Symbols" },
            brushes: { ja: "ブラシ", en: "Brushes" },
            paragraphStyles: { ja: "段落スタイル", en: "Paragraph styles" },
            characterStyles: { ja: "文字スタイル", en: "Character styles" },
            strayPoints: { ja: "孤立点", en: "Stray points" },
            emptyText: { ja: "空のテキスト", en: "Empty text frames" },
            noPaintPath: { ja: "塗りも線もないパス", en: "Paths with no fill or stroke" },
            zeroOpacity: { ja: "不透明度0%のオブジェクト", en: "Objects at 0% opacity" },
            hiddenObjects: { ja: "非表示オブジェクト", en: "Hidden objects" },
            brokenLink: { ja: "リンク切れの配置画像", en: "Broken-link placed images" },
            notes: { ja: "メモ", en: "Notes" },
            outsideAllArtboards: { ja: "アートボード外のオブジェクト", en: "Objects outside all artboards" },
            outsideActiveArtboard: { ja: "アクティブなアートボード外のオブジェクト", en: "Objects outside the active artboard" },
            emptyGroup: { ja: "空のグループ", en: "Empty groups" },
            clearGuides: { ja: "ガイドを消去（ロック分は残る）", en: "Clear Guides (locked ones remain)" },
            guides: { ja: "すべてのガイド（ロックも解除）", en: "All guides (unlocks everything)" },
            guidesOutsideActiveArtboard: { ja: "アクティブなアートボード以外のガイド", en: "Guides outside the active artboard" },
            emptyLayer: { ja: "空のレイヤー／サブレイヤー", en: "Empty layers / sublayers" },
            artboards: { ja: "空のアートボード", en: "Empty artboards" }
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            done: { ja: "削除が完了しました。", en: "Deletion complete." },
            noTarget: { ja: "削除対象はありませんでした。", en: "No items to delete." },
            actionFailed: {
                ja: "次の項目は「未使用をすべて選択」アクションを実行できず、削除をスキップしました：",
                en: "The following items were skipped because the \"Select All Unused\" action could not be played:"
            },
            folderNotChosen: {
                ja: "対象フォルダーが指定されていません。［指定...］ボタンから選んでください。",
                en: "No target folder is chosen. Pick one with the Choose button."
            },
            folderNoFiles: {
                ja: "指定したフォルダーに .ai ファイルがありません。\n\n{0}\n（フォルダー直下の {1} 項目を確認しました。サブフォルダーは対象外です）",
                en: "The chosen folder contains no .ai files.\n\n{0}\n({1} item(s) checked directly inside the folder; subfolders are not included.)"
            },
            folderConfirm: {
                ja: "「{1}」内の {0} 個の .ai ファイルを開いて処理し、上書き保存します。\n\n元のファイルは復元できません。実行しますか？",
                en: "{0} .ai file(s) in \"{1}\" will be opened, cleaned, and saved over the originals.\n\nThis cannot be undone. Continue?"
            },
            batchDoneFiles: { ja: "{0} 個のファイルを処理しました。", en: "Processed {0} file(s)." },
            batchDoneDocuments: {
                ja: "{0} 個のドキュメントを処理しました（保存はしていません）。",
                en: "Processed {0} document(s) (not saved)."
            },
            batchFailed: {
                ja: "次のファイル／ドキュメントは処理できませんでした：",
                en: "The following files/documents could not be processed:"
            }
        },
        prompt: {
            selectFolder: {
                ja: "処理対象の .ai ファイルが入ったフォルダーを選択",
                en: "Choose the folder containing the .ai files to process"
            }
        },
        tooltip: {
            optionClickToggleAll: {
                ja: "option（Alt）＋クリックで、このパネル内のチェックをまとめてON/OFFできます。",
                en: "Option/Alt-click any checkbox to turn every option in this panel on or off."
            },
            preset: {
                ja: "削除対象をまとめて切り替えます。「基本」は初期状態に戻し、「すべてON」はガイドも「すべてのガイド」に、「パネル項目のみ」はパネル項目だけをONにします。手で変えると「カスタム」になります。「使用中のパネル項目も削除」は変わりません。",
                en: "Switches every deletion target at once. Default restores the initial state, All on also picks the all-guides option, and Panel items only leaves just the panel items checked. Changing anything by hand switches this to Custom. The force option is left alone."
            },
            target: {
                ja: "どのドキュメントを処理するかを選びます。",
                en: "Choose which documents to process."
            },
            chooseFolder: {
                ja: "処理対象の .ai ファイルが入ったフォルダーを選びます。サブフォルダーは対象外です。",
                en: "Pick the folder holding the .ai files to process. Subfolders are not included."
            },
            run: {
                ja: "フォルダー指定のときは、各ファイルを上書き保存します。実行前に確認ダイアログが出ます。",
                en: "With the folder scope, every file is saved over the original. A confirmation appears before the run."
            },
            guidesNone: {
                ja: "ガイドには手を触れません。",
                en: "Leaves every guide untouched."
            },
            frontmostDocument: {
                ja: "現在いちばん手前にあるドキュメントだけを処理します。保存はしません。",
                en: "Processes only the frontmost document. Nothing is saved."
            },
            allOpenDocuments: {
                ja: "開いているすべてのドキュメントを処理します。保存はしないので、結果を確認してから保存してください。",
                en: "Processes every open document. Nothing is saved, so review the results before saving."
            },
            targetFolder: {
                ja: "指定したフォルダー直下の .ai ファイルを順に開いて処理し、上書き保存して閉じます。サブフォルダーは対象外です。",
                en: "Opens each .ai file directly inside the chosen folder, cleans it, saves over the original, and closes it. Subfolders are not included."
            },
            swatches: {
                ja: "未使用のスウォッチをアクションで削除します（プロセスカラーも判定）。強制ONで保護対象以外をすべて削除します。空になったスウォッチグループも片付けます。",
                en: "Prunes unused swatches via an action (process colors included). Force mode removes all but protected ones. Swatch groups left empty are cleared out too."
            },
            graphicStyles: {
                ja: "未使用のグラフィックスタイルをアクションで削除します。強制ONで既定（最後の1つ）以外をすべて削除します。",
                en: "Prunes unused graphic styles via an action. Force mode removes all but the default (the last one)."
            },
            symbols: {
                ja: "未使用のシンボルをアクションで削除します。強制ONでは削除できるものをすべて削除します（使用中は Illustrator が削除を拒むため残ります）。",
                en: "Prunes unused symbols via an action. Force mode removes every removable one; symbols still in use are refused by Illustrator and remain."
            },
            brushes: {
                ja: "未使用のブラシをアクションで削除します。強制ONでは使用中・基本ブラシ以外をすべて削除します。",
                en: "Prunes unused brushes via an action. Force mode removes all but in-use and basic brushes."
            },
            strayPoints: {
                ja: "アンカーが1点だけで長さを持たないパス（孤立点）を削除します。",
                en: "Removes stray points (single-anchor paths with no length)."
            },
            emptyText: {
                ja: "文字が入っていない空のテキストを削除します（ポイント文字／塗り・線のないエリア内・パス上文字）。",
                en: "Removes empty text frames (point text, and area/path text whose path has no fill or stroke)."
            },
            noPaintPath: {
                ja: "塗りも線もない（画面に見えない）パスを削除します。ガイドとクリッピングパスは対象外です。",
                en: "Removes paths with no fill and no stroke (invisible). Guides and clipping paths are excluded."
            },
            zeroOpacity: {
                ja: "不透明度が0%（完全に透明）の個々のオブジェクトを削除します（グループ内の項目も対象）。グループ自体に設定した不透明度0%、ガイド、クリッピングパス、コンパウンドパスの構成パスは対象外です。",
                en: "Removes individual objects at 0% opacity (fully transparent), including items inside groups. A group's own 0% opacity, guides, clipping paths, and compound-path members are all excluded."
            },
            hiddenObjects: {
                ja: "非表示（隠した）オブジェクトを削除します。非表示グループはその中身ごと削除されます。",
                en: "Removes hidden objects. A hidden group is removed together with its contents."
            },
            brokenLink: {
                ja: "リンク先ファイルが見つからない配置画像（リンク切れ）を削除します。埋め込み画像や正常なリンクは対象外です。",
                en: "Removes placed images whose linked file is missing (broken links). Embedded images and valid links are kept."
            },
            notes: {
                ja: "属性パネルの「メモ」を削除します（グループ内の項目も対象。オブジェクト自体は残ります）。スクリプトがメモに保存した情報（アウトライン化したテキストの復元用など）も失われます。",
                en: "Deletes the Attributes panel note on each object, including items inside groups; the objects themselves stay. Any data a script keeps in a note, such as what's needed to restore outlined text, is lost too."
            },
            outsideAllArtboards: {
                ja: "どのアートボードにも載っていないオブジェクトを削除します。",
                en: "Removes objects that sit on none of the artboards."
            },
            outsideActiveArtboard: {
                ja: "アクティブなアートボードに載っていないオブジェクトを削除します。",
                en: "Removes objects that do not sit on the active artboard."
            },
            emptyGroup: {
                ja: "中身のない空のグループを削除します。クリップグループは、マスク以外が塗り・線なしのパスだけの場合も対象です。",
                en: "Removes empty groups with no contents. Clip groups also count when their non-mask contents are only paths with no fill or stroke."
            },
            clearGuides: {
                ja: "メニュー「ガイドを消去」を実行します。ロックされたレイヤー上のガイドは残ることがあります（その場合は「すべてのガイド（ロックも解除）」を使用）。",
                en: "Runs the Clear Guides menu command. Guides on locked layers may remain (use \"All guides (force)\" for those)."
            },
            guides: {
                ja: "ドキュメント内のすべてのガイドを削除します（ロックされたレイヤー・サブレイヤー・ガイドも一時的に解除）。",
                en: "Removes all guides in the document (temporarily unlocking locked layers, sublayers, and the guides themselves)."
            },
            guidesOutsideActiveArtboard: {
                ja: "アクティブなアートボード上にないガイドを削除し、そのアートボード上のガイドだけを残します。",
                en: "Removes guides that are not on the active artboard, keeping only that artboard's guides."
            },
            emptyLayer: {
                ja: "中身が空のレイヤー／サブレイヤーを再帰的に削除します。ガイドだけが残っているレイヤーは残し、トップレベルは最低1つ残します。",
                en: "Recursively removes empty layers and sublayers. Layers holding only guides are kept, and at least one top-level layer remains."
            },
            paragraphStyles: {
                ja: "使用状況を取得できないため、強制ON時のみ [標準段落スタイル] 以外を削除します。",
                en: "Usage can't be detected, so all but [Normal Paragraph Style] are removed only in force mode."
            },
            characterStyles: {
                ja: "使用状況を取得できないため、強制ON時のみ [標準文字スタイル] 以外を削除します。",
                en: "Usage can't be detected, so all but [Normal Character Style] are removed only in force mode."
            },
            artboards: {
                ja: "アートワークが載っていない空のアートボードを削除します（最低1つは残します）。中身のあるアートボードは削除しません。",
                en: "Removes empty artboards with no artwork (keeps at least one). Artboards holding artwork are never removed."
            },
            force: {
                ja: "パネル項目だけに効きます。OFF では各パネルの「未使用を選択」で未使用のみ削除。ON ではドキュメント内で使用されていても削除し、段落・文字スタイルの削除も有効になります（保護対象・既定は残ります）。シンボルとブラシは使用中だと Illustrator が削除を拒むため残ります。アートボードには影響しません。",
                en: "Applies to panel items only. When off, each panel's Select All Unused prunes unused items. When on, items are removed even when used in the document and paragraph/character style removal is enabled (protected and default items remain). Symbols and brushes still in use are refused by Illustrator and remain. Artboards are unaffected."
            }
        }
    };

    /**
     * LABELS からドット区切りのパスで表示言語のテキストを取り出す
     * @param {string} labelPath - "panel.target" のようなドット区切りのキー
     * @returns {string} 表示言語のテキスト（見つからない場合は labelPath をそのまま返す）
     */
    function getLabel(labelPath) {
        var labelPathKeys = labelPath.split(".");
        var labelNode = LABELS;
        for (var i = 0; i < labelPathKeys.length; i++) {
            if (labelNode == null) {
                return labelPath;
            }
            labelNode = labelNode[labelPathKeys[i]];
        }
        if (labelNode == null) {
            return labelPath;
        }
        return labelNode[uiLang] || labelNode.ja || labelPath;
    }

    /**
     * ラベル内の {0} {1} … を値で置き換える
     * @param {string} labelTemplate - プレースホルダーを含むラベル
     * @param {Array} placeholderValues - {0} から順に差し込む値
     * @returns {string} 置き換えたテキスト
     */
    function fillPlaceholders(labelTemplate, placeholderValues) {
        var filledText = labelTemplate;
        for (var i = 0; i < placeholderValues.length; i++) {
            filledText = filledText.replace("{" + i + "}", placeholderValues[i]);
        }
        return filledText;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================
    (function() {
        /* 種類キーごとの削除処理。handler(doc, force) で件数を返す（force を使わない処理は無視する）
           Handler per type key; handler(doc, force) returns the count (handlers that don't use force ignore it) */
        var TARGET_RUNNERS = {
            swatches: deleteUnusedSwatches,
            graphicStyles: deleteUnusedGraphicStyles,
            symbols: deleteUnusedSymbols,
            brushes: deleteUnusedBrushes,
            paragraphStyles: deleteUnusedParagraphStyles,
            characterStyles: deleteUnusedCharacterStyles,
            strayPoints: deleteStrayPoints,
            emptyText: deleteEmptyTextFrames,
            noPaintPath: deleteUnpaintedPaths,
            zeroOpacity: deleteZeroOpacityObjects,
            hiddenObjects: deleteHiddenObjects,
            brokenLink: deleteBrokenLinkImages,
            notes: clearNotes,
            outsideAllArtboards: deleteObjectsOutsideAllArtboards,
            outsideActiveArtboard: deleteObjectsOutsideActiveArtboard,
            emptyGroup: deleteEmptyGroups,
            clearGuides: clearGuides,
            guides: deleteAllGuides,
            guidesOutsideActiveArtboard: deleteGuidesOutsideActiveArtboard,
            emptyLayer: deleteEmptyLayers,
            artboards: deleteUnusedArtboards
        };

        /* 削除対象プリセットのポップアップ項目。「カスタム」は選ぶものではなく、手でチェックを変えたことを示す状態
           Items in the deletion-target preset popup; Custom isn't meant to be picked, it reports a hand-edited state */
        var PRESET_BASIC = 0;
        var PRESET_ALL_OFF = 1;
        var PRESET_ALL_ON = 2;
        var PRESET_PANEL_ITEMS_ONLY = 3;
        var PRESET_CUSTOM = 4;

        /* 対象の指定方法 / How the target is chosen */
        var TARGET_FRONTMOST = "frontmost";
        var TARGET_ALL_OPEN = "allOpen";
        var TARGET_FOLDER = "folder";

        /* ダイアログの構成（トップパネル → サブパネル/チェックボックス）。force はパネル項目パネルの末尾に置く / Dialog layout; the force option sits at the bottom of the panel-items panel */
        var DIALOG_LAYOUT = [{
                row: [{
                        column: [{
                                titleKey: "panel.panelItems",
                                force: true,
                                /* 「パネル項目のみ」プリセットで ON にするパネル / The panel the Panel-items-only preset turns on */
                                panelItemsGroup: true,
                                keys: ["swatches", "symbols", "brushes", "graphicStyles", "paragraphStyles", "characterStyles"]
                            },
                            {
                                titleKey: "panel.container",
                                keys: ["emptyGroup", "emptyLayer"]
                            }
                        ]
                    },
                    {
                        column: [{
                                titleKey: "panel.object",
                                keys: ["strayPoints", "emptyText", "noPaintPath", "zeroOpacity", "hiddenObjects", "brokenLink", "notes"]
                            },
                            {
                                titleKey: "panel.guide",
                                radio: true,
                                noneKey: "guidesNone",
                                /* 「すべてON」プリセットで選ぶ項目 / The option the All-on preset picks */
                                allOnKey: "guides",
                                keys: ["clearGuides", "guides", "guidesOutsideActiveArtboard"]
                            }
                        ]
                    }
                ]
            },
            {
                titleKey: "panel.artboard",
                keys: ["outsideAllArtboards", "outsideActiveArtboard", "artboards"]
            }
        ];

        /* レイアウト順の全キー / Every key in layout order */
        var ALL_KEYS = [];
        /* ラジオで排他選択するパネル（ガイド）の情報。プリセット適用と復元で「削除しない」を含めて明示的に設定するために使う
           The radio-selected panel (guides); presets and restoring need to set its options explicitly, none option included */
        var GUIDE_RADIO_GROUP = null;
        /* 「パネル項目のみ」プリセットで ON にするキーの集合 / The set of keys the Panel-items-only preset turns on */
        var PANEL_ITEM_KEY_SET = {};

        forEachLayoutPanel(DIALOG_LAYOUT, function(panelSpec) {
            ALL_KEYS = ALL_KEYS.concat(panelSpec.keys);
            if (panelSpec.radio && !GUIDE_RADIO_GROUP) {
                GUIDE_RADIO_GROUP = {
                    noneKey: panelSpec.noneKey,
                    allOnKey: panelSpec.allOnKey,
                    keys: panelSpec.keys
                };
            }
            if (panelSpec.panelItemsGroup) {
                for (var i = 0; i < panelSpec.keys.length; i++) {
                    PANEL_ITEM_KEY_SET[panelSpec.keys[i]] = true;
                }
            }
        });

        /* 空グループ・空レイヤーの掃除は他の削除の後に回す（他の削除で空になった親も同じ実行で消せるように）/ Run container cleanup after the other deletions so parents emptied by them are removed in the same pass */
        var CONTAINER_CLEANUP_KEYS = {
            emptyGroup: true,
            emptyLayer: true
        };
        /* メモの削除はオブジェクトを残すので、さらに最後に回す（あとで削除されるオブジェクトを件数に含めないように）
           Note removal keeps the objects, so it runs at the very end; objects deleted afterwards would otherwise inflate its count */
        var FINAL_KEYS = {
            notes: true
        };
        var EXECUTION_KEYS = (function(layoutOrderedKeys) {
            var earlierKeys = [];
            var containerCleanupKeys = [];
            var finalKeys = [];
            for (var i = 0; i < layoutOrderedKeys.length; i++) {
                var optionKey = layoutOrderedKeys[i];
                if (FINAL_KEYS[optionKey]) {
                    finalKeys.push(optionKey);
                } else if (CONTAINER_CLEANUP_KEYS[optionKey]) {
                    containerCleanupKeys.push(optionKey);
                } else {
                    earlierKeys.push(optionKey);
                }
            }
            return earlierKeys.concat(containerCleanupKeys, finalKeys);
        })(ALL_KEYS);

        /* 削除対象を選ぶダイアログを表示 / Show the dialog for choosing what to delete */
        var dialogChoices = showDeleteDialog();
        if (!dialogChoices) {
            return;
        }

        /* アクションの再生に失敗した種類キー。完了メッセージで「未使用0件」と区別して警告する / Type keys whose action failed to play; warned in the summary so they aren't mistaken for a genuine zero */
        var actionFailureKeys = [];

        /* 対象の指定方法に応じて実行 / Run according to the chosen target */
        if (dialogChoices.targetMode === TARGET_FOLDER) {
            runFolderBatch(dialogChoices.targetFolder, dialogChoices);
        } else if (dialogChoices.targetMode === TARGET_ALL_OPEN) {
            runOpenDocumentsBatch(dialogChoices);
        } else {
            runFrontmostDocument(dialogChoices);
        }

        /**
         * レイアウト定義を順に辿り、パネル1枚ぶんの定義（row / column 以外のノード）ごとに関数を呼ぶ
         * @param {Object[]} layoutNodes - レイアウト定義の配列
         * @param {Function} visitPanel - パネルの定義を受け取る関数
         * @returns {void}
         */
        function forEachLayoutPanel(layoutNodes, visitPanel) {
            for (var i = 0; i < layoutNodes.length; i++) {
                var layoutNode = layoutNodes[i];
                if (layoutNode.row) {
                    forEachLayoutPanel(layoutNode.row, visitPanel);
                } else if (layoutNode.column) {
                    forEachLayoutPanel(layoutNode.column, visitPanel);
                } else {
                    visitPanel(layoutNode);
                }
            }
        }

        // ==================================================
        // 実行 / Execution
        // ==================================================

        /**
         * 1ドキュメントに対して選択された種類をすべて実行し、種類キーごとの件数を返す（実行順は EXECUTION_KEYS）
         * @param {Document} targetDoc - 処理するドキュメント
         * @param {Object} choices - ダイアログの選択結果
         * @returns {Object} 種類キーごとの削除件数
         */
        function runCleanup(targetDoc, choices) {
            /* アクションとメニューコマンドは最前面のドキュメントに効くため、対象を必ず前面にする
               Actions and menu commands act on the frontmost document, so bring the target forward first */
            app.activeDocument = targetDoc;

            var deletedCounts = {};
            for (var i = 0; i < EXECUTION_KEYS.length; i++) {
                var optionKey = EXECUTION_KEYS[i];
                if (choices[optionKey]) {
                    deletedCounts[optionKey] = TARGET_RUNNERS[optionKey](targetDoc, choices.force);
                }
            }
            return deletedCounts;
        }

        /**
         * 最前面のドキュメントだけを処理する（保存はしない）
         * @param {Object} choices - ダイアログの選択結果
         * @returns {void}
         */
        function runFrontmostDocument(choices) {
            if (app.documents.length === 0) {
                alert(getLabel('alert.noDocument'));
                return;
            }

            var deletedCounts = runCleanup(app.activeDocument, choices);
            app.redraw();
            alert(buildSingleResultMessage(deletedCounts, choices));
        }

        /**
         * 開いているすべてのドキュメントを処理する（保存はしない）
         * @param {Object} choices - ダイアログの選択結果
         * @returns {void}
         */
        function runOpenDocumentsBatch(choices) {
            if (app.documents.length === 0) {
                alert(getLabel('alert.noDocument'));
                return;
            }

            /* 処理中にアクティブドキュメントが切り替わるため、参照を先に控える / The active document changes while running, so snapshot the references first */
            var targetDocs = [];
            for (var i = 0; i < app.documents.length; i++) {
                targetDocs.push(app.documents[i]);
            }
            var originallyActiveDoc = app.activeDocument;

            var deletedTotals = {};
            var processedCount = 0;
            var failedNames = [];
            for (var j = 0; j < targetDocs.length; j++) {
                try {
                    addCounts(deletedTotals, runCleanup(targetDocs[j], choices));
                    processedCount++;
                } catch (e) {
                    failedNames.push(getDocumentName(targetDocs[j]));
                }
            }

            /* 元の作業ドキュメントを前面に戻す / Bring the originally active document back to the front */
            try {
                app.activeDocument = originallyActiveDoc;
            } catch (e2) {
                /* 閉じられていた場合などは無視 / Ignore when it is no longer available */
            }
            app.redraw();

            alert(fillPlaceholders(getLabel('alert.batchDoneDocuments'), [processedCount]) + "\n\n" +
                buildBatchResultMessage(deletedTotals, choices, failedNames));
        }

        /**
         * 指定フォルダー直下の .ai ファイルを順に開いて処理し、上書き保存して閉じる
         * @param {Folder} targetFolder - 対象フォルダー
         * @param {Object} choices - ダイアログの選択結果
         * @returns {void}
         */
        function runFolderBatch(targetFolder, choices) {
            var aiFiles = collectAiFiles(targetFolder);
            if (aiFiles.length === 0) {
                /* 対象フォルダーと走査した項目数を出して、フォルダー違いと絞り込み漏れを切り分けられるようにする
                   Report the folder and how many entries were scanned, so a wrong folder can be told apart from a filtering problem */
                alert(fillPlaceholders(getLabel('alert.folderNoFiles'),
                    [formatDisplayPath(targetFolder.fsName), targetFolder.getFiles().length]));
                return;
            }

            /* 元ファイルを上書きして元に戻せないため、実行前に必ず確認する / The originals are overwritten irreversibly, so always confirm first */
            if (!confirm(fillPlaceholders(getLabel('alert.folderConfirm'), [aiFiles.length, targetFolder.fsName]))) {
                return;
            }

            /* 一括処理中はフォント欠落などのダイアログで止まらないようにする / Keep blocking dialogs (missing fonts, etc.) out of the way during the batch */
            var originalInteractionLevel = app.userInteractionLevel;
            app.userInteractionLevel = UserInteractionLevel.DONTDISPLAYALERTS;

            var deletedTotals = {};
            var processedCount = 0;
            var failedNames = [];
            try {
                for (var i = 0; i < aiFiles.length; i++) {
                    var openedDoc = null;
                    try {
                        openedDoc = app.open(aiFiles[i]);
                        /* 保存に失敗した変更は破棄されるので、集計への加算は上書き保存が終わってから
                           Changes are discarded when the save fails, so only fold the counts in once the file is safely saved */
                        var fileCounts = runCleanup(openedDoc, choices);
                        openedDoc.close(SaveOptions.SAVECHANGES);
                        openedDoc = null;
                        addCounts(deletedTotals, fileCounts);
                        processedCount++;
                    } catch (e) {
                        failedNames.push(getEntryName(aiFiles[i]));
                        /* 失敗したファイルは保存せずに閉じ、壊れた状態を書き戻さない / Close a failed file without saving so a broken state isn't written back */
                        if (openedDoc !== null) {
                            try {
                                openedDoc.close(SaveOptions.DONOTSAVECHANGES);
                            } catch (e2) {
                                /* 閉じられない場合は無視 / Ignore when it can't be closed */
                            }
                        }
                    }
                }
            } finally {
                /* 例外時もユーザー操作レベルを必ず元に戻す / Always restore the interaction level, even on error */
                app.userInteractionLevel = originalInteractionLevel;
            }

            alert(fillPlaceholders(getLabel('alert.batchDoneFiles'), [processedCount]) + "\n\n" +
                buildBatchResultMessage(deletedTotals, choices, failedNames));
        }

        /**
         * ファイル名を取得する。displayName は OS の表示名なので拡張子の判定には使わず、実ファイル名の name を復号して使う
         * @param {File} fileEntry - ファイル
         * @returns {string} 復号したファイル名
         */
        function getEntryName(fileEntry) {
            return decodeURI(fileEntry.name);
        }

        /**
         * フォルダー直下の .ai ファイルを名前順に集める（サブフォルダー・不可視ファイルは対象外）
         * @param {Folder} targetFolder - 対象フォルダー
         * @returns {File[]} .ai ファイルの配列
         */
        function collectAiFiles(targetFolder) {
            var folderEntries = targetFolder.getFiles();
            var aiFiles = [];
            for (var i = 0; i < folderEntries.length; i++) {
                var entry = folderEntries[i];
                /* サブフォルダーを除く（File 判定ではなく Folder 判定にして、取りこぼしを防ぐ）/ Exclude subfolders; testing for Folder rather than File avoids dropping everything if the check misbehaves */
                if (entry instanceof Folder) {
                    continue;
                }
                var entryName = getEntryName(entry);
                /* 不可視ファイルは対象外 / Skip invisible files */
                if (entryName.charAt(0) === ".") {
                    continue;
                }
                if (entryName.length > 3 && entryName.substring(entryName.length - 3).toLowerCase() === ".ai") {
                    aiFiles.push(entry);
                }
            }
            aiFiles.sort(function(entryA, entryB) {
                var nameA = getEntryName(entryA);
                var nameB = getEntryName(entryB);
                if (nameA === nameB) {
                    return 0;
                }
                return (nameA < nameB) ? -1 : 1;
            });
            return aiFiles;
        }

        /**
         * 件数の集計に1ドキュメントぶんの結果を足し込む
         * @param {Object} deletedTotals - 種類キーごとの合計（書き換える）
         * @param {Object} deletedCounts - 1ドキュメントぶんの件数
         * @returns {void}
         */
        function addCounts(deletedTotals, deletedCounts) {
            for (var i = 0; i < ALL_KEYS.length; i++) {
                var optionKey = ALL_KEYS[i];
                if (deletedCounts[optionKey] > 0) {
                    deletedTotals[optionKey] = (deletedTotals[optionKey] || 0) + deletedCounts[optionKey];
                }
            }
        }

        /**
         * ドキュメント名を取得する（取得できない場合は空文字）
         * @param {Document} targetDoc - ドキュメント
         * @returns {string} ドキュメント名
         */
        function getDocumentName(targetDoc) {
            try {
                return targetDoc.name;
            } catch (e) {
                return "";
            }
        }

        // ==================================================
        // 結果メッセージ / Result message
        // ==================================================

        /**
         * 集計1行を組み立てる（日本語は「件」を付ける）
         * @param {string} labelPath - 種類名のラベルのパス
         * @param {number} deletedCount - 件数
         * @returns {string} 改行付きの1行
         */
        function formatResultLine(labelPath, deletedCount) {
            if (uiLang === "ja") {
                return getLabel(labelPath) + ": " + deletedCount + " 件\n";
            }
            return getLabel(labelPath) + ": " + deletedCount + "\n";
        }

        /**
         * 種類ごとの削除件数を行にまとめる（0件の種類は省略）
         * @param {Object} deletedCounts - 種類キーごとの件数
         * @param {Object} choices - ダイアログの選択結果
         * @returns {string} 件数の行。1件もなければ空文字
         */
        function buildCountLines(deletedCounts, choices) {
            var countLines = "";
            for (var i = 0; i < ALL_KEYS.length; i++) {
                var optionKey = ALL_KEYS[i];
                if (choices[optionKey] && deletedCounts[optionKey] > 0) {
                    countLines += formatResultLine('result.' + optionKey, deletedCounts[optionKey]);
                }
            }
            return countLines;
        }

        /**
         * 見出しと項目名の一覧を組み立てる
         * @param {string} headingPath - 見出しのラベルのパス
         * @param {string[]} itemNames - 並べる項目名
         * @returns {string} 一覧のテキスト。該当なしなら空文字
         */
        function buildNoticeLines(headingPath, itemNames) {
            if (itemNames.length === 0) {
                return "";
            }
            var noticeLines = "\n" + getLabel(headingPath) + "\n";
            for (var i = 0; i < itemNames.length; i++) {
                noticeLines += "- " + itemNames[i] + "\n";
            }
            return noticeLines;
        }

        /**
         * アクションを再生できなかった種類の一覧を組み立てる（0件ではなく失敗として明示する）
         * @returns {string} 一覧のテキスト。該当なしなら空文字
         */
        function buildActionFailureLines() {
            var failedTypeNames = [];
            for (var i = 0; i < actionFailureKeys.length; i++) {
                failedTypeNames.push(getLabel('result.' + actionFailureKeys[i]));
            }
            return buildNoticeLines('alert.actionFailed', failedTypeNames);
        }

        /**
         * 単一ドキュメント用の完了メッセージを組み立てる
         * @param {Object} deletedCounts - 種類キーごとの件数
         * @param {Object} choices - ダイアログの選択結果
         * @returns {string} 完了メッセージ
         */
        function buildSingleResultMessage(deletedCounts, choices) {
            var countLines = buildCountLines(deletedCounts, choices);
            /* 1件も削除されなければ「対象なし」だけ表示 / If nothing was deleted, show only the no-target message */
            var message = (countLines === "") ? getLabel('alert.noTarget') : getLabel('alert.done') + "\n\n" + countLines;
            return message + buildActionFailureLines();
        }

        /**
         * 一括処理用の完了メッセージを組み立てる（合計件数と失敗したファイル名）
         * @param {Object} deletedTotals - 種類キーごとの合計
         * @param {Object} choices - ダイアログの選択結果
         * @param {string[]} failedNames - 処理できなかったファイル／ドキュメント名
         * @returns {string} 完了メッセージ
         */
        function buildBatchResultMessage(deletedTotals, choices, failedNames) {
            var countLines = buildCountLines(deletedTotals, choices);
            var message = (countLines === "") ? getLabel('alert.noTarget') : countLines;
            return message + buildActionFailureLines() + buildNoticeLines('alert.batchFailed', failedNames);
        }

        // ==================================================
        // ダイアログ生成 / Build the dialog
        // ==================================================

        /**
         * 削除対象を選択するダイアログを表示し、選択結果を返す
         * @returns {Object|null} 対象の指定と各項目の選択状態。キャンセル時は null
         */
        function showDeleteDialog() {
            var cleanerDialog = new Window("dialog", getLabel('dialog.title') + " " + SCRIPT_VERSION);
            cleanerDialog.orientation = "column";
            cleanerDialog.alignChildren = ["fill", "top"];

            var targetControls = buildTargetPanel(cleanerDialog);
            var optionState = buildOptionPanels(cleanerDialog);
            buildButtonRow(cleanerDialog, targetControls, optionState);

            restoreSessionChoices(targetControls, optionState);
            /* 復元後の状態に合わせてプリセット表示を決める / Pick the preset display that matches the restored state */
            refreshPresetSelection(optionState);

            /* 前回の位置を再現する。画面構成が変わって画面外になる位置は使わない
               Reuse the previous position, unless a display change would leave it off-screen */
            cleanerDialog.onShow = function() {
                if (sessionState.dialogLocation && isLocationOnScreen(sessionState.dialogLocation)) {
                    cleanerDialog.location = sessionState.dialogLocation;
                }
            };

            if (cleanerDialog.show() !== 1) {
                return null;
            }

            var choices = readDialogChoices(targetControls, optionState);

            /* 次回の起動で再現できるようセッションに控える。Folder オブジェクトは持ち越さず、パス文字列だけを残す
               Stash it in the session for the next run; the Folder object is dropped and only its path is kept */
            sessionState.choices = copyChoicesForSession(choices);
            sessionState.folderPath = targetControls.selectedFolder ? targetControls.selectedFolder.fsName : null;

            return choices;
        }

        /**
         * 最上部の「処理対象」パネル（最前面／すべて開いている／フォルダー指定）を作る
         * @param {Window} parentDialog - 親ダイアログ
         * @returns {Object} ラジオ・選択中のフォルダー・操作関数をまとめたオブジェクト
         */
        function buildTargetPanel(parentDialog) {
            var targetPanel = parentDialog.add("panel", undefined, getLabel('panel.target'));
            setupPanel(targetPanel, 6);
            targetPanel.helpTip = getLabel('tooltip.target');

            var openDocumentCount = app.documents.length;

            var frontmostRadio = targetPanel.add("radiobutton", undefined, getLabel('radio.frontmostDocument'));
            frontmostRadio.helpTip = getLabel('tooltip.frontmostDocument');
            /* ドキュメントが1つも開いていなければ選べない / Not selectable when no document is open */
            frontmostRadio.enabled = (openDocumentCount > 0);

            /* 対象になるドキュメント数をラベルに添える / Show how many documents the option covers */
            var allOpenRadio = targetPanel.add("radiobutton", undefined,
                fillPlaceholders(getLabel('radio.allOpenDocuments'), [openDocumentCount]));
            allOpenRadio.helpTip = getLabel('tooltip.allOpenDocuments');
            /* 1つ以下のときは「最前面のドキュメント」と変わらないのでディム表示 / Dimmed at one document or fewer, where it would do the same as the frontmost option */
            allOpenRadio.enabled = (openDocumentCount > 1);

            /* フォルダー指定はラジオと［指定］ボタンを1行に並べる / The folder option puts the radio and the Choose button on one row */
            var folderRow = targetPanel.add("group");
            folderRow.orientation = "row";
            folderRow.alignment = ["fill", "top"];
            folderRow.alignChildren = ["left", "center"];
            folderRow.spacing = 8;

            var folderRadio = folderRow.add("radiobutton", undefined, getLabel('radio.targetFolder'));
            folderRadio.helpTip = getLabel('tooltip.targetFolder');
            /* ラジオも生成後に伸びないので、件数を付けた状態の幅を先に確保する / A radiobutton doesn't grow after creation either, so reserve the width the counted label needs */
            folderRadio.preferredSize.width = FOLDER_RADIO_WIDTH;
            var btnChooseFolder = folderRow.add("button", undefined, getLabel('button.chooseFolder'));
            btnChooseFolder.helpTip = getLabel('tooltip.chooseFolder');

            /* 選択中のパスは次の行に、ラジオのラベル位置に合わせて字下げして表示 / The chosen path goes on the next line, indented to line up with the radio's label */
            var folderPathRow = targetPanel.add("group");
            folderPathRow.orientation = "row";
            folderPathRow.alignment = ["fill", "top"];
            folderPathRow.alignChildren = ["left", "center"];
            folderPathRow.margins = [FOLDER_PATH_INDENT, 0, 0, 0];

            /* statictext は生成後に伸びないので、長いパスを収める幅を先に確保する / A statictext never grows after creation, so reserve enough width for a long path up front */
            var folderPathText = folderPathRow.add("statictext", undefined, getLabel('status.noFolderChosen'), {
                truncate: "middle"
            });
            folderPathText.preferredSize.width = FOLDER_PATH_WIDTH;
            folderPathText.alignment = ["fill", "center"];

            var targetControls = {
                frontmostRadio: frontmostRadio,
                allOpenRadio: allOpenRadio,
                folderRadio: folderRadio,
                /* 選んだフォルダー（未指定なら null）/ The chosen folder (null when none) */
                selectedFolder: null,
                selectRadio: selectTargetRadio,
                applyChosenFolder: applyChosenFolder
            };

            /* 3つのラジオのうち1つだけを選択状態にする。フォルダー指定だけ横並び用の別グループにいて ScriptUI の自動排他が効かないため手動で揃える
               Select exactly one of the three radios by hand: the folder radio sits in its own row group, so ScriptUI won't treat the three as one exclusive set */
            function selectTargetRadio(chosenRadio) {
                frontmostRadio.value = (chosenRadio === frontmostRadio);
                allOpenRadio.value = (chosenRadio === allOpenRadio);
                folderRadio.value = (chosenRadio === folderRadio);
            }

            /* 使えないラジオは選ばない。ドキュメントが1つも開いていなければフォルダー指定から始める
               Never land on a disabled radio; with no document open, start on the folder option */
            function selectDefaultTargetRadio() {
                selectTargetRadio(frontmostRadio.enabled ? frontmostRadio : folderRadio);
            }

            /* 選んだフォルダーを保持して表示に反映する（選択ダイアログ経由でもセッション復元でも通る）
               Hold the chosen folder and reflect it in the display; used both by the picker and by the session restore */
            function applyChosenFolder(chosenFolder) {
                targetControls.selectedFolder = chosenFolder;
                /* 表示は Dropbox 配下ならプレフィックスを落とし、そうでなければ ~ に短縮。幅に収まらないぶんは中央が省略されるため、末尾のフォルダー名は残る
                   Drop the Dropbox prefix when it applies, otherwise show the ~ form; middle truncation keeps the trailing folder name readable */
                folderPathText.text = formatDisplayPath(chosenFolder.fsName);
                folderPathText.helpTip = chosenFolder.fsName;
                /* 処理対象になる .ai ファイル数をラベルに添える（0件なら実行前に気づける）/ Show how many .ai files the run will cover, so a zero is obvious before running */
                folderRadio.text = fillPlaceholders(getLabel('radio.targetFolderWithCount'), [collectAiFiles(chosenFolder).length]);
            }

            /* フォルダー選択ダイアログを開き、選んだフォルダーを表示に反映する（キャンセル時は false）
               Open the folder picker and reflect the choice in the display; returns false when cancelled */
            function chooseTargetFolder() {
                var chosenFolder = Folder.selectDialog(getLabel('prompt.selectFolder'), targetControls.selectedFolder);
                if (!chosenFolder) {
                    return false;
                }
                applyChosenFolder(chosenFolder);
                return true;
            }

            selectDefaultTargetRadio();

            frontmostRadio.onClick = function() {
                selectTargetRadio(frontmostRadio);
            };

            allOpenRadio.onClick = function() {
                selectTargetRadio(allOpenRadio);
            };

            btnChooseFolder.onClick = function() {
                /* ［指定］を押したらフォルダー指定モードに切り替える / Pressing Choose switches the target to folder mode */
                if (chooseTargetFolder()) {
                    selectTargetRadio(folderRadio);
                }
            };

            folderRadio.onClick = function() {
                /* 未指定のままフォルダー指定を選んだら、その場で選択させる。キャンセル時はディム表示のラジオに戻さない
                   Picking folder mode with nothing chosen opens the picker right away; cancelling must not land on a dimmed radio */
                if (targetControls.selectedFolder === null && !chooseTargetFolder()) {
                    selectDefaultTargetRadio();
                    return;
                }
                selectTargetRadio(folderRadio);
            };

            return targetControls;
        }

        /**
         * レイアウト定義どおりに削除対象のパネル・サブパネル・チェックボックス／ラジオを作る
         * @param {Window} parentDialog - 親ダイアログ
         * @returns {Object} 項目のコントロールとプリセットの状態をまとめたオブジェクト
         */
        function buildOptionPanels(parentDialog) {
            var optionState = {
                /* 種類キー → チェックボックス／ラジオ / Type key -> checkbox or radio */
                optionControls: {},
                /* プリセットが値を書き込むチェックボックスの一覧。{ checkbox, key, initialValue } の組で持ち、
                   ScriptUI ウィジェットに独自プロパティを生やさない。強制オプションは含めない（ガイドのラジオは GUIDE_RADIO_GROUP 経由）
                   The checkboxes a preset writes to, held as { checkbox, key, initialValue } records rather than as custom
                   properties on the ScriptUI widgets; the force option is never included (guide radios go through GUIDE_RADIO_GROUP) */
                checkboxEntries: [],
                forceCheckbox: null,
                presetDropdown: null,
                /* 表示を合わせ直している最中かどうか。ScriptUI は selection への代入でも onChange を呼ぶため、
                   そのままだと手で変えた直後にプリセットが再適用されてしまう
                   Whether the display is being resynced; ScriptUI fires onChange on assignment to selection too,
                   which would otherwise re-apply the preset right after a hand-made change */
                isSyncingPreset: false
            };

            /* 手で選択を変えたらプリセット表示を今の状態に合わせ直す（多くは「カスタム」になる）
               After a hand-made change, resync the preset display with the current state (usually landing on Custom) */
            function onManualChange() {
                refreshPresetSelection(optionState);
            }

            for (var i = 0; i < DIALOG_LAYOUT.length; i++) {
                var layoutNode = DIALOG_LAYOUT[i];
                if (layoutNode.row) {
                    /* 複数パネルを横並び / Lay multiple panels side by side */
                    var panelRow = parentDialog.add("group");
                    panelRow.orientation = "row";
                    panelRow.alignChildren = ["fill", "top"];
                    panelRow.alignment = "fill";
                    panelRow.spacing = PANEL_SPACING;
                    for (var j = 0; j < layoutNode.row.length; j++) {
                        buildPanelNode(panelRow, layoutNode.row[j], optionState, onManualChange);
                    }
                } else {
                    buildPanelNode(parentDialog, layoutNode, optionState, onManualChange);
                }
            }
            return optionState;
        }

        /**
         * レイアウト定義1件ぶんのパネルを親コンテナに作る
         * @param {Object} parentContainer - 親のダイアログまたはグループ
         * @param {Object} layoutNode - レイアウト定義のノード
         * @param {Object} optionState - buildOptionPanels() の状態（書き換える）
         * @param {Function} onManualChange - 手で変えたときに呼ぶ関数
         * @returns {void}
         */
        function buildPanelNode(parentContainer, layoutNode, optionState, onManualChange) {
            /* column ノードは複数パネルを縦積みするグループ / A column node stacks several panels vertically */
            if (layoutNode.column) {
                var panelStack = parentContainer.add("group");
                panelStack.orientation = "column";
                panelStack.alignChildren = ["fill", "top"];
                panelStack.alignment = "fill";
                panelStack.spacing = PANEL_SPACING;
                for (var k = 0; k < layoutNode.column.length; k++) {
                    buildPanelNode(panelStack, layoutNode.column[k], optionState, onManualChange);
                }
                return;
            }

            var sectionPanel = parentContainer.add("panel", undefined, getLabel(layoutNode.titleKey));
            setupPanel(sectionPanel, 6);
            if (layoutNode.radio) {
                addRadioGroup(sectionPanel, layoutNode, optionState.optionControls, onManualChange);
            } else {
                /* option＋クリックの一括切り替えが効くのはチェックボックスのパネルだけ / Only checkbox panels support the option-click bulk toggle */
                sectionPanel.helpTip = getLabel('tooltip.optionClickToggleAll');
                optionState.checkboxEntries = optionState.checkboxEntries.concat(
                    addCheckboxes(sectionPanel, layoutNode.keys, optionState.optionControls, onManualChange));
            }

            /* 強制オプションは対象パネル（パネル項目）の末尾に追加 / The force option sits at the bottom of its panel (panel items) */
            if (layoutNode.force) {
                optionState.forceCheckbox = addForceCheckbox(sectionPanel);
            }
        }

        /**
         * 区切り線と「使用中のパネル項目も削除」のチェックボックスをパネルの末尾に追加する
         * @param {Panel} sectionPanel - 追加先のパネル
         * @returns {Checkbox} 追加したチェックボックス
         */
        function addForceCheckbox(sectionPanel) {
            /* 使用中削除オプションの上に区切り線。上下の余白は区切り線側にまとめて持たせる / Divider above the force option; the space above and below it belongs to the divider */
            var forceDividerWrap = sectionPanel.add("group");
            forceDividerWrap.orientation = "column";
            forceDividerWrap.alignChildren = ["fill", "top"];
            forceDividerWrap.alignment = ["fill", "top"];
            forceDividerWrap.margins = [0, FORCE_DIVIDER_MARGIN, 0, FORCE_DIVIDER_MARGIN];
            forceDividerWrap.spacing = 0;
            var forceDivider = forceDividerWrap.add("panel");
            forceDivider.alignment = ["fill", "top"];
            forceDivider.minimumSize.height = forceDivider.maximumSize.height = 1;

            var forceGroup = sectionPanel.add("group");
            forceGroup.orientation = "column";
            forceGroup.alignChildren = ["left", "top"];
            forceGroup.alignment = "left";
            forceGroup.margins = [0, 0, 0, 0];
            var forceCheckbox = forceGroup.add("checkbox", undefined, getLabel('checkbox.force'));
            forceCheckbox.value = false;
            forceCheckbox.helpTip = getLabel('tooltip.force');
            return forceCheckbox;
        }

        /**
         * ボタンエリアを作る。左（削除対象プリセット）／中央スペーサー／右（キャンセル・実行）の3カラム
         * @param {Window} parentDialog - 親ダイアログ
         * @param {Object} targetControls - buildTargetPanel() の戻り値
         * @param {Object} optionState - buildOptionPanels() の戻り値（presetDropdown を書き込む）
         * @returns {void}
         */
        function buildButtonRow(parentDialog, targetControls, optionState) {
            var btnRowGroup = parentDialog.add("group");
            btnRowGroup.orientation = "row";
            btnRowGroup.alignment = "fill";
            btnRowGroup.alignChildren = ["fill", "center"];

            var btnLeftGroup = btnRowGroup.add("group");
            btnLeftGroup.orientation = "row";
            btnLeftGroup.alignment = ["left", "center"];
            btnLeftGroup.add("statictext", undefined, getLabel('fieldLabel.preset'));
            var presetDropdown = btnLeftGroup.add("dropdownlist", undefined, [
                getLabel('dropdown.presetBasic'),
                getLabel('dropdown.presetAllOff'),
                getLabel('dropdown.presetAllOn'),
                getLabel('dropdown.presetPanelItemsOnly'),
                getLabel('dropdown.presetCustom')
            ]);
            presetDropdown.helpTip = getLabel('tooltip.preset');
            presetDropdown.preferredSize.width = PRESET_DROPDOWN_WIDTH;
            /* 初期状態は「基本」そのもの / The initial state is exactly the Default preset */
            presetDropdown.selection = PRESET_BASIC;
            presetDropdown.onChange = function() {
                /* 表示を合わせ直しただけの代入では、プリセットを適用し直さない / An assignment that only resyncs the display must not re-apply the preset */
                if (optionState.isSyncingPreset) {
                    return;
                }
                if (this.selection) {
                    applyCheckboxPreset(optionState.checkboxEntries, optionState.optionControls, this.selection.index);
                }
            };
            optionState.presetDropdown = presetDropdown;

            /* 中央のスペーサーが余白を吸収して左右を両端に寄せる / The center spacer absorbs slack, pushing the two sides apart */
            var spacer = btnRowGroup.add("group");
            spacer.alignment = ["fill", "center"];
            spacer.minimumSize.width = 1;

            var btnRightGroup = btnRowGroup.add("group");
            btnRightGroup.orientation = "row";
            btnRightGroup.alignment = ["right", "center"];
            var btnCancel = btnRightGroup.add("button", undefined, getLabel('button.cancel'), {
                name: "cancel"
            });
            /* キャンセルでも位置だけは覚える（onClick を付けたので明示的に閉じる）
               Remember the position even on cancel; onClick replaces the default close, so close explicitly */
            btnCancel.onClick = function() {
                rememberDialogLocation(parentDialog);
                parentDialog.close(2);
            };

            var btnRun = btnRightGroup.add("button", undefined, getLabel('button.run'), {
                name: "ok"
            });
            btnRun.helpTip = getLabel('tooltip.run');

            /* フォルダー未指定のまま実行させない（onClick を付けたので明示的に閉じる）
               Don't let the run start without a folder; onClick replaces the default close, so close explicitly */
            btnRun.onClick = function() {
                if (targetControls.folderRadio.value && targetControls.selectedFolder === null) {
                    alert(getLabel('alert.folderNotChosen'));
                    return;
                }
                rememberDialogLocation(parentDialog);
                parentDialog.close(1);
            };
        }

        /**
         * ダイアログの位置をセッションに控える
         * @param {Window} cleanerDialog - ダイアログ
         * @returns {void}
         */
        function rememberDialogLocation(cleanerDialog) {
            sessionState.dialogLocation = tryGet(function() {
                return [cleanerDialog.location[0], cleanerDialog.location[1]];
            }, null);
        }

        /**
         * 手で選択を変えたあと、プリセット表示を今の状態に合わせ直す（多くは「カスタム」になる）
         * @param {Object} optionState - buildOptionPanels() の戻り値
         * @returns {void}
         */
        function refreshPresetSelection(optionState) {
            if (!optionState.presetDropdown) {
                return;
            }
            optionState.isSyncingPreset = true;
            try {
                optionState.presetDropdown.selection = detectPreset(optionState.checkboxEntries, optionState.optionControls);
            } finally {
                optionState.isSyncingPreset = false;
            }
        }

        /**
         * 前回の選択をセッションから復元する（記憶がなければ何もしない）
         * @param {Object} targetControls - buildTargetPanel() の戻り値
         * @param {Object} optionState - buildOptionPanels() の戻り値
         * @returns {void}
         */
        function restoreSessionChoices(targetControls, optionState) {
            var savedChoices = sessionState.choices;
            if (!savedChoices) {
                return;
            }

            /* チェックボックス（ガイドのラジオはこの後まとめて設定する）/ Checkboxes; the guide radios are set together below */
            var checkboxEntries = optionState.checkboxEntries;
            for (var i = 0; i < checkboxEntries.length; i++) {
                checkboxEntries[i].checkbox.value = (savedChoices[checkboxEntries[i].key] === true);
            }

            /* ガイドは1つだけONにする / Exactly one guide option is on */
            if (GUIDE_RADIO_GROUP) {
                var chosenGuideKey = null;
                for (var j = 0; j < GUIDE_RADIO_GROUP.keys.length; j++) {
                    if (savedChoices[GUIDE_RADIO_GROUP.keys[j]]) {
                        chosenGuideKey = GUIDE_RADIO_GROUP.keys[j];
                        break;
                    }
                }
                selectGuideRadio(optionState.optionControls, chosenGuideKey);
            }

            if (optionState.forceCheckbox) {
                optionState.forceCheckbox.value = (savedChoices.force === true);
            }

            /* 処理対象。今のドキュメント数で選べない項目と、消えたフォルダーは復元しない
               The scope; an option the current document count disables, or a folder that is gone, is not restored */
            if (savedChoices.targetMode === TARGET_ALL_OPEN && targetControls.allOpenRadio.enabled) {
                targetControls.selectRadio(targetControls.allOpenRadio);
            } else if (savedChoices.targetMode === TARGET_FOLDER && sessionState.folderPath) {
                var savedFolder = new Folder(sessionState.folderPath);
                if (savedFolder.exists) {
                    targetControls.applyChosenFolder(savedFolder);
                    targetControls.selectRadio(targetControls.folderRadio);
                }
            }
        }

        /**
         * 対象の指定と各項目の選択状態を種類キーごとにまとめる
         * @param {Object} targetControls - buildTargetPanel() の戻り値
         * @param {Object} optionState - buildOptionPanels() の戻り値
         * @returns {Object} 選択結果（force / targetMode / targetFolder と種類キーごとの真偽値）
         */
        function readDialogChoices(targetControls, optionState) {
            var choices = {
                force: optionState.forceCheckbox ? optionState.forceCheckbox.value : false,
                targetMode: targetControls.folderRadio.value ? TARGET_FOLDER : (targetControls.allOpenRadio.value ? TARGET_ALL_OPEN : TARGET_FRONTMOST),
                targetFolder: targetControls.selectedFolder
            };
            for (var i = 0; i < ALL_KEYS.length; i++) {
                choices[ALL_KEYS[i]] = optionState.optionControls[ALL_KEYS[i]].value;
            }
            return choices;
        }

        /**
         * セッションに残す用に選択結果を複製する。targetFolder（Folder オブジェクト）は folderPath に置き換わるので持ち越さない
         * @param {Object} choices - ダイアログの選択結果
         * @returns {Object} 複製した選択結果
         */
        function copyChoicesForSession(choices) {
            var sessionChoices = {
                force: choices.force,
                targetMode: choices.targetMode
            };
            for (var i = 0; i < ALL_KEYS.length; i++) {
                sessionChoices[ALL_KEYS[i]] = choices[ALL_KEYS[i]];
            }
            return sessionChoices;
        }

        /**
         * 保存した位置がいずれかの画面に収まるか判定する
         * @param {number[]} dialogLocation - [x, y]
         * @returns {boolean} 収まれば true
         */
        function isLocationOnScreen(dialogLocation) {
            var screenList = tryGet(function() {
                return $.screens;
            }, null);
            if (!screenList || !screenList.length) {
                return false;
            }
            for (var i = 0; i < screenList.length; i++) {
                var screenBounds = screenList[i];
                if (dialogLocation[0] >= screenBounds.left && dialogLocation[0] <= screenBounds.right - ON_SCREEN_MARGIN &&
                    dialogLocation[1] >= screenBounds.top && dialogLocation[1] <= screenBounds.bottom - ON_SCREEN_MARGIN) {
                    return true;
                }
            }
            return false;
        }

        /**
         * キー配列ぶんのチェックボックスを親に追加し、参照を記録して { checkbox, key, initialValue } の配列を返す。
         * initialValue は「基本」プリセットで元に戻すために控える
         * @param {Object} parentContainer - 追加先のパネル
         * @param {string[]} optionKeys - 種類キーの配列
         * @param {Object} optionControls - 種類キー → コントロール（書き換える）
         * @param {Function} onManualChange - 手で変えたときに呼ぶ関数
         * @returns {Object[]} { checkbox, key, initialValue } の配列
         */
        function addCheckboxes(parentContainer, optionKeys, optionControls, onManualChange) {
            var panelEntries = [];
            for (var i = 0; i < optionKeys.length; i++) {
                var optionCheckbox = parentContainer.add("checkbox", undefined, getLabel('checkbox.' + optionKeys[i]));
                optionCheckbox.value = !UNCHECKED_BY_DEFAULT[optionKeys[i]];
                optionCheckbox.helpTip = getLabel('tooltip.' + optionKeys[i]);
                optionControls[optionKeys[i]] = optionCheckbox;
                panelEntries.push({
                    checkbox: optionCheckbox,
                    key: optionKeys[i],
                    initialValue: optionCheckbox.value
                });
            }
            enableOptionClickToggleAll(panelEntries, onManualChange);
            return panelEntries;
        }

        /**
         * option（Alt）キーが押されているか判定する
         * @returns {boolean} 押されていれば true
         */
        function isOptionKeyDown() {
            return tryGet(function() {
                return ScriptUI.environment.keyboardState.altKey === true;
            }, false);
        }

        /**
         * option（Alt）＋クリックで、渡したチェックボックス全部をクリック先の状態に揃える。修飾キーなしのクリックは通常どおり
         * @param {Object[]} panelEntries - { checkbox, key, initialValue } の配列
         * @param {Function} onManualChange - 手で変えたときに呼ぶ関数
         * @returns {void}
         */
        function enableOptionClickToggleAll(panelEntries, onManualChange) {
            for (var i = 0; i < panelEntries.length; i++) {
                panelEntries[i].checkbox.onClick = function() {
                    if (isOptionKeyDown()) {
                        /* onClick の時点でクリックされた本人の値は反映済み / By the time onClick runs, the clicked checkbox already holds its new value */
                        var newValue = this.value;
                        for (var j = 0; j < panelEntries.length; j++) {
                            panelEntries[j].checkbox.value = newValue;
                        }
                    }
                    /* 手で変えた時点でプリセットは「カスタム」に落ちる / A hand-made change drops the preset to Custom */
                    onManualChange();
                };
            }
        }

        /**
         * ガイドのラジオで選ばれているキーを返す
         * @param {Object} optionControls - 種類キー → コントロール
         * @returns {string|null} 選ばれているキー。「削除しない」のときは null
         */
        function getSelectedGuideKey(optionControls) {
            if (!GUIDE_RADIO_GROUP) {
                return null;
            }
            for (var i = 0; i < GUIDE_RADIO_GROUP.keys.length; i++) {
                if (optionControls[GUIDE_RADIO_GROUP.keys[i]].value) {
                    return GUIDE_RADIO_GROUP.keys[i];
                }
            }
            return null;
        }

        /**
         * ガイドのラジオを1つだけONにする
         * @param {Object} optionControls - 種類キー → コントロール
         * @param {string|null} chosenKey - ONにするキー。null なら「削除しない」
         * @returns {void}
         */
        function selectGuideRadio(optionControls, chosenKey) {
            if (!GUIDE_RADIO_GROUP) {
                return;
            }
            optionControls[GUIDE_RADIO_GROUP.noneKey].value = (chosenKey === null);
            for (var i = 0; i < GUIDE_RADIO_GROUP.keys.length; i++) {
                optionControls[GUIDE_RADIO_GROUP.keys[i]].value = (GUIDE_RADIO_GROUP.keys[i] === chosenKey);
            }
        }

        /**
         * プリセットに合わせてチェックとガイドのラジオを一括設定する（「カスタム」は現状維持）
         * @param {Object[]} checkboxEntries - { checkbox, key, initialValue } の配列
         * @param {Object} optionControls - 種類キー → コントロール
         * @param {number} presetIndex - プリセットの番号（PRESET_*）
         * @returns {void}
         */
        function applyCheckboxPreset(checkboxEntries, optionControls, presetIndex) {
            if (presetIndex === PRESET_CUSTOM) {
                return;
            }
            for (var i = 0; i < checkboxEntries.length; i++) {
                var checkboxEntry = checkboxEntries[i];
                if (presetIndex === PRESET_ALL_ON) {
                    checkboxEntry.checkbox.value = true;
                } else if (presetIndex === PRESET_ALL_OFF) {
                    checkboxEntry.checkbox.value = false;
                } else if (presetIndex === PRESET_PANEL_ITEMS_ONLY) {
                    /* すべてOFFにしてから、パネル項目だけをONにする / Everything off, then just the panel items back on */
                    checkboxEntry.checkbox.value = (PANEL_ITEM_KEY_SET[checkboxEntry.key] === true);
                } else {
                    checkboxEntry.checkbox.value = checkboxEntry.initialValue;
                }
            }
            /* 「すべてON」はガイドも「すべてのガイド」に。それ以外は既定の「削除しない」
               All-on also picks the all-guides option; the others fall back to "Don't delete" */
            var chosenGuideKey = (GUIDE_RADIO_GROUP && presetIndex === PRESET_ALL_ON) ? GUIDE_RADIO_GROUP.allOnKey : null;
            selectGuideRadio(optionControls, chosenGuideKey);
        }

        /**
         * 現在の選択がどのプリセットに当たるかを判定する
         * @param {Object[]} checkboxEntries - { checkbox, key, initialValue } の配列
         * @param {Object} optionControls - 種類キー → コントロール
         * @returns {number} プリセットの番号（PRESET_*）
         */
        function detectPreset(checkboxEntries, optionControls) {
            var matchesInitial = true;
            var allOn = true;
            var allOff = true;
            var panelItemsOnly = true;

            for (var i = 0; i < checkboxEntries.length; i++) {
                var checkboxEntry = checkboxEntries[i];
                if (checkboxEntry.checkbox.value !== checkboxEntry.initialValue) {
                    matchesInitial = false;
                }
                if (checkboxEntry.checkbox.value) {
                    allOff = false;
                } else {
                    allOn = false;
                }
                if (checkboxEntry.checkbox.value !== (PANEL_ITEM_KEY_SET[checkboxEntry.key] === true)) {
                    panelItemsOnly = false;
                }
            }

            /* ガイドの既定は「削除しない」なので、選ばれていれば初期状態ではない / The guide default is "Don't delete", so any choice means it isn't the initial state */
            if (GUIDE_RADIO_GROUP) {
                var guideKey = getSelectedGuideKey(optionControls);
                if (guideKey !== null) {
                    matchesInitial = false;
                    allOff = false;
                    panelItemsOnly = false;
                }
                if (guideKey !== GUIDE_RADIO_GROUP.allOnKey) {
                    allOn = false;
                }
            }

            if (matchesInitial) {
                return PRESET_BASIC;
            }
            if (allOn) {
                return PRESET_ALL_ON;
            }
            if (allOff) {
                return PRESET_ALL_OFF;
            }
            if (panelItemsOnly) {
                return PRESET_PANEL_ITEMS_ONLY;
            }
            return PRESET_CUSTOM;
        }

        /**
         * ラジオグループを作る。先頭に「削除しない」（既定で選択）を置き、各キーのラジオを記録する。
         * プリセット判定と復元で使うため、「削除しない」も noneKey で記録する
         * @param {Panel} parentPanel - 追加先のパネル
         * @param {Object} layoutNode - レイアウト定義のノード（noneKey と keys を持つ）
         * @param {Object} optionControls - 種類キー → コントロール（書き換える）
         * @param {Function} onManualChange - 手で変えたときに呼ぶ関数
         * @returns {void}
         */
        function addRadioGroup(parentPanel, layoutNode, optionControls, onManualChange) {
            var panelRadios = [];
            var noneRadio = parentPanel.add("radiobutton", undefined, getLabel('radio.' + layoutNode.noneKey));
            noneRadio.helpTip = getLabel('tooltip.' + layoutNode.noneKey);
            noneRadio.value = true;
            optionControls[layoutNode.noneKey] = noneRadio;
            panelRadios.push(noneRadio);

            for (var i = 0; i < layoutNode.keys.length; i++) {
                var optionRadio = parentPanel.add("radiobutton", undefined, getLabel('radio.' + layoutNode.keys[i]));
                optionRadio.helpTip = getLabel('tooltip.' + layoutNode.keys[i]);
                optionRadio.value = false;
                optionControls[layoutNode.keys[i]] = optionRadio;
                panelRadios.push(optionRadio);
            }

            /* ガイドを手で切り替えたときもプリセット表示を追従させる / Keep the preset display in step when a guide option is picked by hand */
            for (var j = 0; j < panelRadios.length; j++) {
                panelRadios[j].onClick = onManualChange;
            }
        }

        // ==================================================
        // 一時アクション / Temporary action
        // ==================================================

        /**
         * ASCII 文字列を16進に変換する
         * @param {string} asciiText - 変換する文字列
         * @returns {string} 16進の文字列
         */
        function asciiToHex(asciiText) {
            var hexText = "";
            for (var i = 0; i < asciiText.length; i++) {
                var hexByte = asciiText.charCodeAt(i).toString(16);
                if (hexByte.length < 2) {
                    hexByte = "0" + hexByte;
                }
                hexText += hexByte;
            }
            return hexText;
        }

        /**
         * メニューコマンド1件のイベントブロックを組み立てる
         * @param {number} eventIndex - イベント番号
         * @param {string} internalName - プラグインの内部名
         * @param {string} localizedNameHex - パネル名の16進
         * @param {string} commandNameHex - コマンド名の16進
         * @param {number} commandValue - コマンドの値
         * @param {boolean} hasDialog - ダイアログを持つコマンドか
         * @returns {string} イベントブロックのテキスト
         */
        function buildMenuEventBlock(eventIndex, internalName, localizedNameHex, commandNameHex, commandValue, hasDialog) {
            var eventLines = [
                "\t/event-" + eventIndex + " {",
                "\t\t/useRulersIn1stQuadrant 1",
                "\t\t/internalName (" + internalName + ")",
                "\t\t/localizedName [ " + (localizedNameHex.length / 2),
                "\t\t\t" + localizedNameHex,
                "\t\t]",
                "\t\t/isOpen 0",
                "\t\t/isOn 1",
                "\t\t/hasDialog " + (hasDialog ? "1" : "0")
            ];
            if (hasDialog) {
                eventLines.push("\t\t/showDialog 0");
            }
            eventLines.push(
                "\t\t/parameterCount 1",
                "\t\t/parameter-1 {",
                "\t\t\t/key 1835363957",
                "\t\t\t/showInPalette 1",
                "\t\t\t/type (enumerated)",
                "\t\t\t/name [ " + (commandNameHex.length / 2),
                "\t\t\t\t" + commandNameHex,
                "\t\t\t]",
                "\t\t\t/value " + commandValue,
                "\t\t}",
                "\t}"
            );
            return eventLines.join("\n");
        }

        /**
         * 「未使用をすべて選択 → 削除」の一時アクション定義を録画値から組み立てる
         * @param {string} setName - アクションセット名
         * @param {string} actionName - アクション名
         * @param {Object} pruneSpec - PRUNE_SPECS の1件
         * @returns {string} アクション定義のテキスト
         */
        function buildActionSource(setName, actionName, pruneSpec) {
            return [
                "/version 3",
                "/name [ " + setName.length,
                "\t" + asciiToHex(setName),
                "]",
                "/isOpen 1",
                "/actionCount 1",
                "/action-1 {",
                "\t/name [ " + actionName.length,
                "\t\t" + asciiToHex(actionName),
                "\t]",
                "\t/keyIndex 0",
                "\t/colorIndex 0",
                "\t/isOpen 1",
                "\t/eventCount 2",
                buildMenuEventBlock(1, pruneSpec.internalName, pruneSpec.localizedNameHex, SELECT_ALL_UNUSED_HEX, pruneSpec.selectValue, false),
                buildMenuEventBlock(2, pruneSpec.internalName, pruneSpec.localizedNameHex, pruneSpec.deleteNameHex, pruneSpec.deleteValue, true),
                "}"
            ].join("\n");
        }

        /**
         * 失敗しても無視してよい後始末を実行する
         * @param {Function} cleanupStep - 後始末の処理
         * @returns {void}
         */
        function ignoringErrors(cleanupStep) {
            try {
                cleanupStep();
            } catch (e) {
                /* 後始末の失敗は無視 / Ignore cleanup failures */
            }
        }

        /**
         * アクションを一時ファイルに書き出して再生する。close/unload/remove は finally で必ず試みる
         * @param {string} actionSource - アクション定義のテキスト
         * @param {string} setName - アクションセット名
         * @param {string} actionName - アクション名
         * @param {string} fileName - 一時ファイルのパス
         * @returns {boolean} 再生できたら true
         */
        function playTemporaryAction(actionSource, setName, actionName, fileName) {
            var actionFile = new File(fileName);
            var played = false;
            try {
                actionFile.encoding = "UTF-8";
                if (!actionFile.open("w")) {
                    return false;
                }
                actionFile.write(actionSource);
                actionFile.close();

                /* loadAction 前に同名セットを解放 / Unload any same-name set before loading */
                ignoringErrors(function() {
                    app.unloadAction(setName, "");
                });

                app.loadAction(actionFile);
                app.doScript(actionName, setName);
                played = true;
            } catch (e) {
                /* 書き出し・読み込み・再生に失敗 / Failed to write, load, or play */
            } finally {
                ignoringErrors(function() {
                    actionFile.close();
                });
                ignoringErrors(function() {
                    app.unloadAction(setName, "");
                });
                ignoringErrors(function() {
                    actionFile.remove();
                });
            }
            return played;
        }

        /**
         * 指定コレクションの未使用項目を、対応するアクションを再生して削除し、件数を返す。
         * 再生に失敗した種類は resultKey を控えて完了時に警告する（「未使用0件」と区別するため）
         * @param {Object} collection - スウォッチなどのコレクション
         * @param {Object} pruneSpec - PRUNE_SPECS の1件
         * @param {string} resultKey - 種類キー
         * @returns {number} 削除した件数
         */
        function pruneUnusedViaAction(collection, pruneSpec, resultKey) {
            var countBefore = collection.length;
            var actionSource = buildActionSource(ACTION_SET_NAME, ACTION_NAME, pruneSpec);
            if (!playTemporaryAction(actionSource, ACTION_SET_NAME, ACTION_NAME, ACTION_FILE_NAME)) {
                /* 一括処理ではファイルごとに同じ失敗が起きるため、種類は1回だけ控える / A batch run hits the same failure per file, so record each type only once */
                var alreadyRecorded = false;
                for (var i = 0; i < actionFailureKeys.length; i++) {
                    if (actionFailureKeys[i] === resultKey) {
                        alreadyRecorded = true;
                        break;
                    }
                }
                if (!alreadyRecorded) {
                    actionFailureKeys.push(resultKey);
                }
                return 0;
            }
            return Math.max(0, countBefore - collection.length);
        }

        // ==================================================
        // 矩形の判定 / Rectangle tests
        // ==================================================

        /**
         * アートワークの外接矩形を一度だけ集める（ガイドは対象外）。アートボードごとに全 pageItems を走査し直さないため
         * @param {Document} doc - 対象ドキュメント
         * @returns {number[][]} 外接矩形 [left, top, right, bottom] の配列
         */
        function collectArtworkBounds(doc) {
            var artworkBounds = [];
            for (var i = 0; i < doc.pageItems.length; i++) {
                var pageItem = doc.pageItems[i];
                try {
                    /* ガイドはアートワークとみなさない / Guides don't count as artwork */
                    if (pageItem.typename === "PathItem" && pageItem.guides) {
                        continue;
                    }
                    artworkBounds.push(pageItem.visibleBounds);
                } catch (e) {
                    /* 取得できないものは数えない / Skip items whose bounds we can't read */
                }
            }
            return artworkBounds;
        }

        /**
         * 2つの矩形 [left, top, right, bottom] が重なるか判定する（Illustrator は上が大きいY）
         * @param {number[]} rectA - 矩形A
         * @param {number[]} rectB - 矩形B
         * @returns {boolean} 重なれば true
         */
        function rectsIntersect(rectA, rectB) {
            return !(
                rectB[2] < rectA[0] || /* B 右端が A 左端より左 / B right is left of A left */
                rectB[0] > rectA[2] || /* B 左端が A 右端より右 / B left is right of A right */
                rectB[3] > rectA[1] || /* B 下端が A 上端より上 / B bottom is above A top */
                rectB[1] < rectA[3]    /* B 上端が A 下端より下 / B top is below A bottom */
            );
        }

        /**
         * 矩形がいずれかの矩形と重なるか判定する
         * @param {number[]} bounds - 調べる矩形
         * @param {number[][]} rects - 比べる矩形の配列
         * @returns {boolean} 1つでも重なれば true
         */
        function intersectsAnyRect(bounds, rects) {
            for (var i = 0; i < rects.length; i++) {
                if (rectsIntersect(rects[i], bounds)) {
                    return true;
                }
            }
            return false;
        }

        /**
         * すべてのアートボードの矩形を取得する
         * @param {Document} doc - 対象ドキュメント
         * @returns {number[][]} アートボード矩形の配列
         */
        function getAllArtboardRects(doc) {
            var artboardRects = [];
            for (var i = 0; i < doc.artboards.length; i++) {
                artboardRects.push(doc.artboards[i].artboardRect);
            }
            return artboardRects;
        }

        // ==================================================
        // 孤立点・空テキスト削除 / Stray points and empty text
        // 移植元 / Ported from: 不要なアイテムを削除.jsx (c) 2020 Toshiyuki Takahashi, MIT License
        // ==================================================

        /**
         * 孤立点（アンカー1点・長さ0のパス）を削除する
         * @param {Document} doc - 対象ドキュメント
         * @returns {number} 削除した件数
         */
        function deleteStrayPoints(doc) {
            var removedCount = 0;
            for (var i = doc.pageItems.length - 1; i >= 0; i--) {
                var pageItem = doc.pageItems[i];
                if (pageItem.typename !== "PathItem") {
                    continue;
                }
                try {
                    if (pageItem.pathPoints.length < 2 && pageItem.length <= 0) {
                        pageItem.remove();
                        removedCount++;
                    }
                } catch (e) {
                    /* 削除不可 / Not removable */
                }
            }
            return removedCount;
        }

        /**
         * 文字のない空テキストを削除する（エリア内文字／パス上文字は塗り・線のないパスのときだけ）
         * @param {Document} doc - 対象ドキュメント
         * @returns {number} 削除した件数
         */
        function deleteEmptyTextFrames(doc) {
            var removedCount = 0;
            for (var i = doc.pageItems.length - 1; i >= 0; i--) {
                var textFrame = doc.pageItems[i];
                if (textFrame.typename !== "TextFrame") {
                    continue;
                }
                try {
                    if (textFrame.contents.length >= 1) {
                        continue;
                    }
                    var shouldRemove = false;
                    if (textFrame.kind === TextType.POINTTEXT) {
                        shouldRemove = true;
                    } else if (textFrame.kind === TextType.AREATEXT || textFrame.kind === TextType.PATHTEXT) {
                        /* テキストパスが塗り・線なしのときだけ削除 / Remove only when the text path has no fill/stroke */
                        if (!textFrame.textPath.stroked && !textFrame.textPath.filled) {
                            shouldRemove = true;
                        }
                    }
                    if (shouldRemove) {
                        textFrame.remove();
                        removedCount++;
                    }
                } catch (e) {
                    /* 削除不可 / Not removable */
                }
            }
            return removedCount;
        }

        // ==================================================
        // オブジェクト削除 / Object deletion
        // ==================================================

        /**
         * 単体で削除してはいけないパスか判定する（ガイド・クリッピングパス・コンパウンドパスの構成パス）
         * @param {PathItem} pathItem - 調べるパス
         * @returns {boolean} 削除対象から外すなら true
         */
        function isStructuralPath(pathItem) {
            /* ガイドは「ガイド」セクションでのみ削除する / Guides are only removed by the guide section */
            if (pathItem.guides) {
                return true;
            }
            /* マスクを消すとクリップが解除され、隠れていた中身が現れてしまう / Removing the mask releases the clip and reveals what it was hiding */
            if (pathItem.clipping) {
                return true;
            }
            /* コンパウンドパスの構成パス（穴など）は単体で削除しない / Don't delete a compound path's member paths (holes, etc.) */
            return !!(pathItem.parent && pathItem.parent.typename === "CompoundPathItem");
        }

        /**
         * 塗りも線もない（不可視の）パスを削除する。ガイド・クリッピングパス・コンパウンドパスの構成パスは除外
         * @param {Document} doc - 対象ドキュメント
         * @returns {number} 削除した件数
         */
        function deleteUnpaintedPaths(doc) {
            var removedCount = 0;
            var pathItems = doc.pathItems;
            for (var i = pathItems.length - 1; i >= 0; i--) {
                var pathItem = pathItems[i];
                try {
                    if (isStructuralPath(pathItem)) {
                        continue;
                    }
                    if (!pathItem.filled && !pathItem.stroked) {
                        pathItem.remove();
                        removedCount++;
                    }
                } catch (e) {
                    /* 削除不可 / Not removable */
                }
            }
            return removedCount;
        }

        /**
         * 不透明度が0%のオブジェクトを削除する（グループ内も対象）。
         * 単体で消すと構造が壊れるもの、および別オプションで扱うガイドは除外する
         * @param {Document} doc - 対象ドキュメント
         * @returns {number} 削除した件数
         */
        function deleteZeroOpacityObjects(doc) {
            var removedCount = 0;

            /* doc.pageItems はグループ内も含む平坦なコレクション / doc.pageItems is a flat collection that includes items inside groups */
            for (var i = doc.pageItems.length - 1; i >= 0; i--) {
                var pageItem = doc.pageItems[i];
                if (pageItem.typename === "GroupItem") {
                    continue;
                }
                try {
                    if (pageItem.typename === "PathItem" && isStructuralPath(pageItem)) {
                        continue;
                    }
                    if (pageItem.opacity === 0) {
                        pageItem.remove();
                        removedCount++;
                    }
                } catch (e) {
                    /* 削除不可 / Not removable */
                }
            }
            return removedCount;
        }

        /**
         * 非表示オブジェクトを削除する（非表示グループは中身ごと）
         * @param {Document} doc - 対象ドキュメント
         * @returns {number} 削除した件数
         */
        function deleteHiddenObjects(doc) {
            /* まず参照だけを収集（この間はコレクションを変更しない）。非表示グループを消すと子のインデックスがずれ、末尾からの走査でも取りこぼすため
               Collect references first without mutating the collection; removing a hidden group shifts child indices, so even a reverse scan would skip items */
            var hiddenItems = [];
            var pageItems = doc.pageItems;
            for (var i = 0; i < pageItems.length; i++) {
                try {
                    if (pageItems[i].hidden) {
                        hiddenItems.push(pageItems[i]);
                    }
                } catch (e) {
                    /* 判定不可はスキップ / Skip items we can't test */
                }
            }

            var removedCount = 0;
            for (var j = 0; j < hiddenItems.length; j++) {
                try {
                    hiddenItems[j].remove();
                    removedCount++;
                } catch (e) {
                    /* 親ごと削除済み、または削除不可 / Already removed with its parent, or not removable */
                }
            }
            return removedCount;
        }

        /**
         * リンク切れ（リンク先が見つからない）の配置画像を削除する。埋め込み・正常リンクは対象外
         * @param {Document} doc - 対象ドキュメント
         * @returns {number} 削除した件数
         */
        function deleteBrokenLinkImages(doc) {
            var removedCount = 0;
            var placedItems = doc.placedItems;
            for (var i = placedItems.length - 1; i >= 0; i--) {
                var placedItem = placedItems[i];
                var isBroken = false;
                try {
                    var linkedFile = placedItem.file;
                    isBroken = (!linkedFile || !linkedFile.exists);
                } catch (e) {
                    /* .file 取得で例外＝リンク切れ扱い / A throwing .file access means a missing link */
                    isBroken = true;
                }
                if (isBroken) {
                    try {
                        placedItem.remove();
                        removedCount++;
                    } catch (e2) {
                        /* 削除不可 / Not removable */
                    }
                }
            }
            return removedCount;
        }

        /**
         * 属性パネルの「メモ」を空にする（グループ内も対象。オブジェクト自体は残す）
         * @param {Document} doc - 対象ドキュメント
         * @returns {number} メモを空にした件数
         */
        function clearNotes(doc) {
            var clearedCount = 0;
            var pageItems = doc.pageItems;
            for (var i = 0; i < pageItems.length; i++) {
                try {
                    if (pageItems[i].note) {
                        pageItems[i].note = "";
                        clearedCount++;
                    }
                } catch (e) {
                    /* ロック等で変更不可 / Can't be changed (locked, etc.) */
                }
            }
            return clearedCount;
        }

        /**
         * どのアートボードにも載っていないオブジェクトを削除する
         * @param {Document} doc - 対象ドキュメント
         * @returns {number} 削除した件数
         */
        function deleteObjectsOutsideAllArtboards(doc) {
            return deleteObjectsOutsideRects(doc, getAllArtboardRects(doc));
        }

        /**
         * アクティブなアートボードに載っていないオブジェクトを削除する
         * @param {Document} doc - 対象ドキュメント
         * @returns {number} 削除した件数
         */
        function deleteObjectsOutsideActiveArtboard(doc) {
            var activeIndex = doc.artboards.getActiveArtboardIndex();
            return deleteObjectsOutsideRects(doc, [doc.artboards[activeIndex].artboardRect]);
        }

        /**
         * 指定矩形のいずれにも重ならないトップレベルオブジェクトを削除する
         * @param {Document} doc - 対象ドキュメント
         * @param {number[][]} keepRects - この矩形に重なるものは残す
         * @returns {number} 削除した件数
         */
        function deleteObjectsOutsideRects(doc, keepRects) {
            var removedCount = 0;

            /* トップレベル（レイヤー直下）のオブジェクトだけをスナップショット。ガイドは専用オプションで扱うため除外 / Snapshot only top-level objects (direct children of a layer); guides are excluded (handled by the dedicated guide option) */
            var topLevelItems = [];
            for (var i = 0; i < doc.pageItems.length; i++) {
                var pageItem = doc.pageItems[i];
                if (pageItem.typename === "PathItem" && pageItem.guides) {
                    continue;
                }
                if (pageItem.parent && pageItem.parent.typename === "Layer") {
                    topLevelItems.push(pageItem);
                }
            }

            for (var j = topLevelItems.length - 1; j >= 0; j--) {
                var topLevelItem = topLevelItems[j];
                var itemBounds;
                try {
                    itemBounds = topLevelItem.visibleBounds;
                } catch (e) {
                    continue;
                }
                if (!intersectsAnyRect(itemBounds, keepRects)) {
                    try {
                        topLevelItem.remove();
                        removedCount++;
                    } catch (e2) {
                        /* 削除不可（ロック等）/ Not removable (locked, etc.) */
                    }
                }
            }
            return removedCount;
        }

        /**
         * 空のグループ（通常グループ・クリップグループ）を削除する
         * @param {Document} doc - 対象ドキュメント
         * @returns {number} 削除した件数
         */
        function deleteEmptyGroups(doc) {
            var removedCount = 0;
            for (var i = 0; i < doc.layers.length; i++) {
                removedCount += removeEmptyGroupsIn(doc.layers[i]);
            }
            return removedCount;
        }

        /**
         * コンテナ内を再帰的に探索し、空のグループを削除する。子を先に掃除するので、空になった親も同じパスで削除できる
         * @param {Layer|GroupItem} parentContainer - 探索するレイヤーまたはグループ
         * @returns {number} 削除した件数
         */
        function removeEmptyGroupsIn(parentContainer) {
            var removedCount = 0;
            /* サブレイヤーの中身は layer.pageItems に含まれないため、先に再帰する / Sublayer contents aren't in layer.pageItems, so recurse into sublayers first */
            if (parentContainer.typename === "Layer") {
                for (var j = parentContainer.layers.length - 1; j >= 0; j--) {
                    removedCount += removeEmptyGroupsIn(parentContainer.layers[j]);
                }
            }
            /* 削除でインデックスがずれるため末尾から / Iterate from the end because removal shifts indices */
            for (var i = parentContainer.pageItems.length - 1; i >= 0; i--) {
                var pageItem = parentContainer.pageItems[i];
                if (pageItem.typename === "GroupItem") {
                    /* 先に中を掃除してから自身の空判定 / Clean inside first, then test this group */
                    removedCount += removeEmptyGroupsIn(pageItem);
                    if (isEmptyGroup(pageItem)) {
                        try {
                            pageItem.remove();
                            removedCount++;
                        } catch (e) {
                            /* 削除不可 / Not removable */
                        }
                    }
                }
            }
            return removedCount;
        }

        /**
         * グループが空か判定する。子が無いグループ、またはマスク以外が塗り・線なしのパスだけのクリップグループを空とみなす
         * @param {PageItem} groupItem - 調べるオブジェクト
         * @returns {boolean} 空なら true
         */
        function isEmptyGroup(groupItem) {
            if (groupItem.typename !== "GroupItem") {
                return false;
            }

            var childItems = groupItem.pageItems;
            /* 子のないグループは空 / A group with no children is empty */
            if (childItems.length === 0) {
                return true;
            }

            /* クリップグループのみ、中身が塗り・線なしパスだけなら空とみなす / Only for clip groups: empty when all contents are unpainted paths */
            if (groupItem.clipped !== true) {
                return false;
            }

            for (var i = 0; i < childItems.length; i++) {
                var childItem = childItems[i];
                if (childItem.typename !== "PathItem") {
                    return false;
                }
                if (childItem.filled === true && childItem.fillColor.typename !== "NoColor") {
                    return false;
                }
                if (childItem.stroked === true && childItem.strokeColor.typename !== "NoColor") {
                    return false;
                }
            }
            return true;
        }

        // ==================================================
        // ガイド削除 / Guide deletion
        // ==================================================

        /**
         * ガイド属性を持つパスの数を数える
         * @param {Document} doc - 対象ドキュメント
         * @returns {number} ガイドの数
         */
        function countGuidePaths(doc) {
            var guideCount = 0;
            var pathItems = doc.pathItems;
            for (var i = 0; i < pathItems.length; i++) {
                if (pathItems[i].guides) {
                    guideCount++;
                }
            }
            return guideCount;
        }

        /**
         * メニューコマンド「ガイドを消去」でガイドを削除する。ロック済みガイドは残す（一時解除しない）
         * @param {Document} doc - 対象ドキュメント
         * @returns {number} 削除した件数
         */
        function clearGuides(doc) {
            var countBefore = countGuidePaths(doc);
            app.executeMenuCommand("clearguide");
            return Math.max(0, countBefore - countGuidePaths(doc));
        }

        /**
         * サブレイヤーを含む全レイヤーを親→子の順に集める
         * @param {Layers} layerCollection - 集めるレイヤーのコレクション
         * @param {Layer[]} collectedLayers - 集めた結果（書き足す）
         * @returns {Layer[]} collectedLayers
         */
        function collectLayersDeep(layerCollection, collectedLayers) {
            for (var i = 0; i < layerCollection.length; i++) {
                var currentLayer = layerCollection[i];
                collectedLayers.push(currentLayer);
                collectLayersDeep(currentLayer.layers, collectedLayers);
            }
            return collectedLayers;
        }

        /**
         * レイヤーのロック状態を控えてから、親から順に解除する
         * @param {Layer[]} allLayers - 親→子の順のレイヤー
         * @returns {boolean[]} 元のロック状態
         */
        function unlockLayers(allLayers) {
            var lockStates = [];
            for (var i = 0; i < allLayers.length; i++) {
                lockStates[i] = false;
                try {
                    lockStates[i] = allLayers[i].locked;
                    if (lockStates[i]) {
                        allLayers[i].locked = false;
                    }
                } catch (e) {
                    /* 解除できないレイヤーはそのまま / Leave layers we can't unlock */
                }
            }
            return lockStates;
        }

        /**
         * レイヤーのロック状態を元に戻す。子から順に戻して親のロックに邪魔されないようにする
         * @param {Layer[]} allLayers - 親→子の順のレイヤー
         * @param {boolean[]} lockStates - unlockLayers() で控えたロック状態
         * @returns {void}
         */
        function restoreLayerLocks(allLayers, lockStates) {
            for (var j = allLayers.length - 1; j >= 0; j--) {
                try {
                    allLayers[j].locked = lockStates[j];
                } catch (e) {
                    /* 復元できない場合は無視 / Ignore when it can't be restored */
                }
            }
        }

        /**
         * ガイドのロック・レイヤーロック（サブレイヤー含む）・ガイド自体のロックを一時解除し、判定関数が真のガイドを削除する。
         * ロック状態は finally で必ず復元する
         * @param {Document} doc - 対象ドキュメント
         * @param {Function} shouldRemove - ガイドのパスを受け取り、削除するなら true を返す関数
         * @returns {number} 削除した件数
         */
        function removeGuidesWhere(doc, shouldRemove) {
            var removedCount = 0;

            /* サブレイヤーまで含めてロック状態を保存し、親から順に解除 / Save every lock state down to sublayers and clear them parents-first */
            var allLayers = collectLayersDeep(doc.layers, []);
            var lockStates = unlockLayers(allLayers);

            var guidesWereLocked = doc.guidesLocked;

            try {
                doc.guidesLocked = false;

                var pathItems = doc.pathItems;
                for (var k = pathItems.length - 1; k >= 0; k--) {
                    var guidePath = pathItems[k];
                    if (!guidePath.guides) {
                        continue;
                    }
                    /* ガイド自体がロックされていると削除できないため一時解除 / A locked guide can't be removed, so clear its lock first */
                    var pathWasLocked = false;
                    try {
                        pathWasLocked = guidePath.locked;
                        if (pathWasLocked) {
                            guidePath.locked = false;
                        }
                        if (shouldRemove(guidePath)) {
                            guidePath.remove();
                            removedCount++;
                        } else if (pathWasLocked) {
                            /* 残すガイドはロックを戻す / Restore the lock on guides we keep */
                            guidePath.locked = true;
                        }
                    } catch (e) {
                        /* 判定不可・削除不可はロックを戻してスキップ / Restore the lock and skip guides we can't test or remove */
                        if (pathWasLocked) {
                            try {
                                guidePath.locked = true;
                            } catch (e2) {
                                /* 復元できない場合は無視 / Ignore when it can't be restored */
                            }
                        }
                    }
                }
            } finally {
                /* 例外時もロック状態を必ず元に戻す / Always restore the lock states, even on error */
                restoreLayerLocks(allLayers, lockStates);
                doc.guidesLocked = guidesWereLocked;
            }

            return removedCount;
        }

        /**
         * すべてのガイドを削除する
         * @param {Document} doc - 対象ドキュメント
         * @returns {number} 削除した件数
         */
        function deleteAllGuides(doc) {
            return removeGuidesWhere(doc, function() {
                return true;
            });
        }

        /**
         * 現在（アクティブ）のアートボード上にないガイドを削除する
         * @param {Document} doc - 対象ドキュメント
         * @returns {number} 削除した件数
         */
        function deleteGuidesOutsideActiveArtboard(doc) {
            var activeRect = doc.artboards[doc.artboards.getActiveArtboardIndex()].artboardRect;
            return removeGuidesWhere(doc, function(guidePath) {
                return !rectsIntersect(activeRect, guidePath.geometricBounds);
            });
        }

        // ==================================================
        // 空のレイヤー削除 / Delete empty layers
        // ==================================================

        /**
         * 中身が空（pageItems もサブレイヤーも無い）のレイヤーを削除する。
         * ガイドは pageItems に含まれるためガイドのみのレイヤーは残る。トップレベルは最低1つ残す
         * @param {Document} doc - 対象ドキュメント
         * @returns {number} 削除した件数
         */
        function deleteEmptyLayers(doc) {
            var removedCount = 0;
            for (var i = doc.layers.length - 1; i >= 0; i--) {
                var topLayer = doc.layers[i];
                /* 先に空のサブレイヤーを削除 / Remove empty sublayers first */
                removedCount += removeEmptySublayers(topLayer);
                if (PROTECTED_LAYER_NAMES[topLayer.name]) {
                    continue;
                }
                /* トップレベルは最低1つ必要 / At least one top-level layer must remain */
                if (topLayer.pageItems.length === 0 && topLayer.layers.length === 0 && doc.layers.length > 1) {
                    removedCount += removeLayerUnlocked(topLayer);
                }
            }
            return removedCount;
        }

        /**
         * サブレイヤーを再帰的に処理し、空のものを削除する
         * @param {Layer} parentLayer - 親レイヤー
         * @returns {number} 削除した件数
         */
        function removeEmptySublayers(parentLayer) {
            var removedCount = 0;
            for (var i = parentLayer.layers.length - 1; i >= 0; i--) {
                var subLayer = parentLayer.layers[i];
                removedCount += removeEmptySublayers(subLayer);
                if (PROTECTED_LAYER_NAMES[subLayer.name]) {
                    continue;
                }
                if (subLayer.pageItems.length === 0 && subLayer.layers.length === 0) {
                    removedCount += removeLayerUnlocked(subLayer);
                }
            }
            return removedCount;
        }

        /**
         * ロックを一時解除してレイヤーを削除する。失敗した場合はロック状態を元に戻す
         * @param {Layer} targetLayer - 削除するレイヤー
         * @returns {number} 削除できた数（0 または 1）
         */
        function removeLayerUnlocked(targetLayer) {
            var wasLocked = false;
            try {
                wasLocked = targetLayer.locked;
                targetLayer.locked = false;
                targetLayer.remove();
                return 1;
            } catch (e) {
                /* 削除できなかったのでロックを戻す（意図しないロック解除を残さない）/ Removal failed, so restore the lock instead of leaving it cleared */
                if (wasLocked) {
                    try {
                        targetLayer.locked = true;
                    } catch (e2) {
                        /* 復元できない場合は無視 / Ignore when it can't be restored */
                    }
                }
                return 0;
            }
        }

        // ==================================================
        // パネル項目の削除 / Delete panel items
        // ==================================================

        /**
         * コレクションの項目を末尾から削除する（先頭の keepLeadingCount 件は残す）。削除できないものは飛ばす
         * @param {Object} collection - シンボル・ブラシ・スタイルなどのコレクション
         * @param {number} keepLeadingCount - 残す先頭の件数
         * @returns {number} 削除した件数
         */
        function removeCollectionItems(collection, keepLeadingCount) {
            var removedCount = 0;
            for (var i = collection.length - 1; i >= keepLeadingCount; i--) {
                try {
                    collection[i].remove();
                    removedCount++;
                } catch (e) {
                    /* 使用中・既定など削除不可 / In use, default, or otherwise not removable */
                }
            }
            return removedCount;
        }

        /**
         * 未使用スウォッチを削除する。通常はアクションで未使用のみ、強制時は保護対象以外を全削除
         * @param {Document} doc - 対象ドキュメント
         * @param {boolean} force - 使用中も削除するか
         * @returns {number} 削除した件数
         */
        function deleteUnusedSwatches(doc, force) {
            var removedCount;
            if (!force) {
                removedCount = pruneUnusedViaAction(doc.swatches, PRUNE_SPECS.swatch, "swatches");
                removeEmptySwatchGroups(doc);
                return removedCount;
            }

            /* 強制時：削除してはいけない既定スウォッチ以外を総当たり削除 / Force: remove everything except the built-in swatches */
            var protectedSwatchNames = {
                "[None]": true,
                "[Registration]": true,
                "[Black]": true,
                "[White]": true
            };

            removedCount = 0;
            for (var i = doc.swatches.length - 1; i >= 0; i--) {
                var swatch = doc.swatches[i];
                if (protectedSwatchNames[swatch.name]) {
                    continue;
                }
                try {
                    swatch.remove();
                    removedCount++;
                } catch (e) {
                    /* 削除不可 / Not removable */
                }
            }
            removeEmptySwatchGroups(doc);
            return removedCount;
        }

        /**
         * スウォッチを消したあとに中身が空になったスウォッチグループを片付ける。スウォッチそのものではないので件数には数えない
         * @param {Document} doc - 対象ドキュメント
         * @returns {void}
         */
        function removeEmptySwatchGroups(doc) {
            for (var i = doc.swatchGroups.length - 1; i >= 0; i--) {
                try {
                    if (doc.swatchGroups[i].getAllSwatches().length === 0) {
                        doc.swatchGroups[i].remove();
                    }
                } catch (e) {
                    /* 空判定も削除もできない場合はそのまま残す / Leave it alone when it can't be tested or removed */
                }
            }
        }

        /**
         * 未使用グラフィックスタイルを削除する。通常はアクションで未使用のみ、強制時は既定（最後の1つ）以外を全削除
         * @param {Document} doc - 対象ドキュメント
         * @param {boolean} force - 使用中も削除するか
         * @returns {number} 削除した件数
         */
        function deleteUnusedGraphicStyles(doc, force) {
            if (!force) {
                return pruneUnusedViaAction(doc.graphicStyles, PRUNE_SPECS.graphicstyle, "graphicStyles");
            }

            var removedCount = 0;
            for (var i = doc.graphicStyles.length - 1; i >= 0; i--) {
                /* 最後の1つ（既定スタイル）は削除不可 / The last one (default style) cannot be removed */
                if (doc.graphicStyles.length <= 1) {
                    break;
                }
                try {
                    doc.graphicStyles[i].remove();
                    removedCount++;
                } catch (e) {
                    /* 削除不可 / Not removable */
                }
            }
            return removedCount;
        }

        /**
         * 未使用シンボルを削除する。通常はアクションで未使用のみ、強制時はすべて削除（使用中は Illustrator が拒むため残る）
         * @param {Document} doc - 対象ドキュメント
         * @param {boolean} force - 使用中も削除するか
         * @returns {number} 削除した件数
         */
        function deleteUnusedSymbols(doc, force) {
            if (!force) {
                return pruneUnusedViaAction(doc.symbols, PRUNE_SPECS.symbol, "symbols");
            }
            return removeCollectionItems(doc.symbols, 0);
        }

        /**
         * 未使用ブラシを削除する。通常はアクションで未使用のみ、強制時は削除できるものをすべて削除（使用中・基本ブラシは不可）
         * @param {Document} doc - 対象ドキュメント
         * @param {boolean} force - 使用中も削除するか
         * @returns {number} 削除した件数
         */
        function deleteUnusedBrushes(doc, force) {
            if (!force) {
                return pruneUnusedViaAction(doc.brushes, PRUNE_SPECS.brush, "brushes");
            }
            return removeCollectionItems(doc.brushes, 0);
        }

        /**
         * 段落スタイルを削除する。使用情報を取得できないため強制時のみ、既定（先頭の [標準段落スタイル]）以外を削除
         * @param {Document} doc - 対象ドキュメント
         * @param {boolean} force - 使用中も削除するか
         * @returns {number} 削除した件数
         */
        function deleteUnusedParagraphStyles(doc, force) {
            if (!force) {
                return 0;
            }
            return removeCollectionItems(doc.paragraphStyles, 1);
        }

        /**
         * 文字スタイルを削除する。使用情報を取得できないため強制時のみ、既定（先頭の [標準文字スタイル]）以外を削除
         * @param {Document} doc - 対象ドキュメント
         * @param {boolean} force - 使用中も削除するか
         * @returns {number} 削除した件数
         */
        function deleteUnusedCharacterStyles(doc, force) {
            if (!force) {
                return 0;
            }
            return removeCollectionItems(doc.characterStyles, 1);
        }

        // ==================================================
        // 空のアートボード削除 / Delete empty artboards
        // ==================================================

        /**
         * 空のアートボードを削除する（最低1つは残す）
         * @param {Document} doc - 対象ドキュメント
         * @returns {number} 削除した件数
         */
        function deleteUnusedArtboards(doc) {
            var removedCount = 0;
            /* アートボードを消してもアートワークは変わらないので、外接矩形は最初に1回だけ集める
               Removing artboards doesn't touch the artwork, so the bounds are collected once up front */
            var artworkBounds = collectArtworkBounds(doc);

            for (var i = doc.artboards.length - 1; i >= 0; i--) {
                /* アートボードは最低1つ必要 / At least one artboard must remain */
                if (doc.artboards.length <= 1) {
                    break;
                }

                /* 「使用中のパネル項目も削除」はパネル項目だけの設定なので、アートボードは常に空のものだけを削除する
                   The force option covers panel items only, so artboards are always limited to the empty ones */
                if (intersectsAnyRect(doc.artboards[i].artboardRect, artworkBounds)) {
                    continue;
                }

                try {
                    doc.artboards.remove(i);
                    removedCount++;
                } catch (e) {
                    /* 削除不可 / Not removable */
                }
            }

            return removedCount;
        }

    })();

})();
