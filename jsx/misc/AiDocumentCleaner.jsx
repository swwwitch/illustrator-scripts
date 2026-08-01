#target illustrator
#targetengine "AiDocumentCleanerSession"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
### 概要

- ドキュメント内の不要な要素（未使用のパネル項目、孤立点や空のテキスト、空のグループ・レイヤー、ガイド、アートボード外のオブジェクトなど）をまとめて削除します。
- 処理対象は「最前面のドキュメント」「開いているすべてのドキュメント」「指定フォルダー内の .ai ファイル（上書き保存）」から選べます。
- ダイアログで削除対象を選び、実行後は種類ごとの削除件数を表示します。
- 誤削除しやすい項目は初期OFF、ガイドは既定で「削除しない」です。
- 詳細な機能・オプションはREADMEを参照してください。

### Overview

- Removes unneeded elements from a document in one pass: unused panel items, stray points and empty text frames, empty groups and layers, guides, objects outside the artboards, and more.
- Runs against the frontmost document, every open document, or all .ai files in a chosen folder (saved over the originals).
- Choose the targets in the dialog; the deleted count per type is reported afterwards.
- Deletion-prone options start unchecked, and guides default to "Don't delete".
- See the README for the full feature and option list.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AiDocumentCleaner";            /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-06-27";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-08-02";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AiDocumentCleaner.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AiDocumentCleaner.md
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n0d70178f0f65"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

// =========================================
// ユーザー設定 / User Settings
// 動作を変えたいときに触る値 / Values to tweak when you want different behavior
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

/* パネルの共通設定 / Apply shared panel layout */
function setupPanel(panel, spacing) {
    panel.orientation = "column";
    panel.alignChildren = ["fill", "top"];
    panel.alignment = "fill";
    panel.margins = PANEL_MARGINS;
    panel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
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
// 移植元 / Ported from: LinkedImageManager.jsx
// =========================================

/* 失敗しうる取得を試み、例外時は代替値を返す / Try a lookup that may throw, falling back to a default */
function tryGet(getValue, fallback) {
    try {
        return getValue();
    } catch (e) {
        return fallback;
    }
}

/* ホーム直下から「Dropbox」を含むフォルダーを探す。チームフォルダー（「sw Dropbox」など）を優先し、
   個人用の「Dropbox」は「~/Dropbox/」として短縮できるため優先度を下げる
   Find a folder named like "Dropbox" directly under the home folder; a team folder ("sw Dropbox") wins over
   the personal "Dropbox", which the ~ form already shortens well */
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

/* フォルダー直下にサブフォルダーが1つだけあるときそれを返す。チームDropboxのメンバーフォルダー（「takano masahiro」など）の判定に使う
   Return the only subfolder directly inside a folder; used to spot a team Dropbox member folder. Null when there are zero or several */
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

/* Dropboxのローカルマウントパスを決める。手動指定が空のときはホーム直下の「Dropbox」を含むフォルダーを探し、
   メンバーフォルダーが1つだけあればそこまでをプレフィックスにする。見つからない場合は空文字
   Resolve the local Dropbox mount path; with no manual setting, look for a "Dropbox" folder under home and
   extend the prefix to its single member folder when there is one. Empty string when nothing is found */
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

/* ホームフォルダー以下のパスを ~ 表記に短縮する / Abbreviate a path under the home folder with ~ */
function abbreviateHomePath(fsPath) {
    var home = tryGet(function() {
        return Folder("~").fsName;
    }, "");
    if (!home) {
        return fsPath;
    }
    if (fsPath === home) {
        return "~";
    }
    if (fsPath.indexOf(home) === 0) {
        var rest = fsPath.substring(home.length);
        /* 直後が区切り文字のときだけ短縮（同名の別フォルダーを誤判定しない）/ Abbreviate only on a separator boundary, so a similarly named folder isn't mistaken for the home folder */
        if (rest.charAt(0) === "/" || rest.charAt(0) === "\\") {
            return "~" + rest;
        }
    }
    return fsPath;
}

/* 表示用にパスを短縮する。Dropbox配下ならプレフィックスを落とし、そうでなければ ~ 表記にする
   Shorten a path for display: drop the Dropbox prefix when it applies, otherwise fall back to the ~ form */
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

/* 現在の UI 言語を取得（ja / en）/ Get current UI language (ja / en) */
function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var currentLanguage = getCurrentLang();

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
        outsideAllArtboards: { ja: "アートボード外のオブジェクト", en: "Objects outside all artboards" },
        outsideActiveArtboard: { ja: "アクティブなアートボード外のオブジェクト", en: "Objects outside the active artboard" },
        emptyGroup: { ja: "空のグループ", en: "Empty groups" },
        guidesNone: { ja: "削除しない", en: "Don't delete" },
        clearGuides: { ja: "ガイドを消去（ロック分は残る）", en: "Clear Guides (locked ones remain)" },
        guides: { ja: "すべてのガイド（ロックも解除）", en: "All guides (unlocks everything)" },
        guidesOutsideActiveArtboard: { ja: "アクティブなアートボード以外", en: "Outside the active artboard" },
        emptyLayer: { ja: "空のレイヤー／サブレイヤー", en: "Empty layers / sublayers" },
        artboards: { ja: "空のアートボード", en: "Empty artboards" },
        force: { ja: "使用中のパネル項目も削除", en: "Delete panel items even if in use" }
    },
    target: {
        frontmostDocument: { ja: "最前面のドキュメント", en: "Frontmost document" },
        allOpenDocuments: { ja: "開いているすべてのドキュメント（{0}）", en: "All open documents ({0})" },
        targetFolder: { ja: "フォルダー指定", en: "Folder" },
        targetFolderWithCount: { ja: "フォルダー指定（{0}）", en: "Folder ({0})" },
        noFolderChosen: { ja: "（未指定）", en: "(none chosen)" }
    },
    button: {
        chooseFolder: { ja: "指定...", en: "Choose..." },
        cancel: { ja: "キャンセル", en: "Cancel" },
        run: { ja: "実行", en: "Run" }
    },
    preset: {
        label: { ja: "セット", en: "Preset" },
        basic: { ja: "基本", en: "Default" },
        allOff: { ja: "すべてOFF", en: "All off" },
        allOn: { ja: "すべてON", en: "All on" },
        panelItemsOnly: { ja: "パネル項目のみ", en: "Panel items only" },
        custom: { ja: "カスタム", en: "Custom" }
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
            ja: "次の項目は「未使用をすべて選択」アクションを実行できず、削除をスキップしました:",
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
            ja: "次のファイル／ドキュメントは処理できませんでした:",
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

/* ドット区切りのキーからローカライズ文字列を取得 / Resolve a localized string from a dotted key path */
function getLabel(labelKey) {
    var keyParts = labelKey.split(".");
    var labelNode = LABELS;
    for (var i = 0; i < keyParts.length; i++) {
        if (labelNode == null) {
            return labelKey;
        }
        labelNode = labelNode[keyParts[i]];
    }
    if (labelNode == null) {
        return labelKey;
    }
    return labelNode[currentLanguage] || labelNode.ja || labelKey;
}

/* ラベル内の {0} {1} … を値で置き換える / Replace {0}, {1}, … placeholders in a label */
function fillPlaceholders(template, values) {
    var filledText = template;
    for (var i = 0; i < values.length; i++) {
        filledText = filledText.replace("{" + i + "}", values[i]);
    }
    return filledText;
}

// =========================================
// メイン処理 / Main
// =========================================
(function() {
    /* 種類キーごとの削除処理。run(doc, force) で件数を返す / Handler per type key; run(doc, force) returns the count */
    var TARGET_RUNNERS = {
        swatches: function(doc, force) {
            return deleteUnusedSwatches(doc, force);
        },
        graphicStyles: function(doc, force) {
            return deleteUnusedGraphicStyles(doc, force);
        },
        symbols: function(doc, force) {
            return deleteUnusedSymbols(doc, force);
        },
        brushes: function(doc, force) {
            return deleteUnusedBrushes(doc, force);
        },
        paragraphStyles: function(doc, force) {
            return deleteUnusedParagraphStyles(doc, force);
        },
        characterStyles: function(doc, force) {
            return deleteUnusedCharacterStyles(doc, force);
        },
        strayPoints: function(doc) {
            return deleteStrayPoints(doc);
        },
        emptyText: function(doc) {
            return deleteEmptyTextFrames(doc);
        },
        noPaintPath: function(doc) {
            return deleteUnpaintedPaths(doc);
        },
        zeroOpacity: function(doc) {
            return deleteZeroOpacityObjects(doc);
        },
        hiddenObjects: function(doc) {
            return deleteHiddenObjects(doc);
        },
        brokenLink: function(doc) {
            return deleteBrokenLinkImages(doc);
        },
        outsideAllArtboards: function(doc) {
            return deleteObjectsOutsideAllArtboards(doc);
        },
        outsideActiveArtboard: function(doc) {
            return deleteObjectsOutsideActiveArtboard(doc);
        },
        emptyGroup: function(doc) {
            return deleteEmptyGroups(doc);
        },
        clearGuides: function(doc) {
            return clearGuides(doc);
        },
        guides: function(doc) {
            return deleteAllGuides(doc);
        },
        guidesOutsideActiveArtboard: function(doc) {
            return deleteGuidesOutsideActiveArtboard(doc);
        },
        emptyLayer: function(doc) {
            return deleteEmptyLayers(doc);
        },
        artboards: function(doc) {
            return deleteUnusedArtboards(doc);
        }
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
                            keys: ["strayPoints", "emptyText", "noPaintPath", "zeroOpacity", "hiddenObjects", "brokenLink"]
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

    /* レイアウトを順に辿って全キーを取り出す / Flatten all keys in layout order */
    var ALL_KEYS = (function(layout) {
        var collectedKeys = [];

        function collectKeysFromNode(layoutNode) {
            if (layoutNode.row) {
                for (var j = 0; j < layoutNode.row.length; j++) {
                    collectKeysFromNode(layoutNode.row[j]);
                }
            } else if (layoutNode.column) {
                for (var k = 0; k < layoutNode.column.length; k++) {
                    collectKeysFromNode(layoutNode.column[k]);
                }
            } else {
                collectedKeys = collectedKeys.concat(layoutNode.keys);
            }
        }
        for (var i = 0; i < layout.length; i++) {
            collectKeysFromNode(layout[i]);
        }
        return collectedKeys;
    })(DIALOG_LAYOUT);

    /* ラジオで排他選択するパネル（ガイド）の情報。プリセット適用と復元で「削除しない」を含めて明示的に設定するために使う
       The radio-selected panel (guides); presets and restoring need to set its options explicitly, none option included */
    var RADIO_GROUP = (function(layout) {
        var found = null;

        function scanNode(layoutNode) {
            if (layoutNode.row) {
                for (var j = 0; j < layoutNode.row.length; j++) {
                    scanNode(layoutNode.row[j]);
                }
            } else if (layoutNode.column) {
                for (var k = 0; k < layoutNode.column.length; k++) {
                    scanNode(layoutNode.column[k]);
                }
            } else if (layoutNode.radio && !found) {
                found = {
                    noneKey: layoutNode.noneKey,
                    allOnKey: layoutNode.allOnKey,
                    keys: layoutNode.keys
                };
            }
        }
        for (var i = 0; i < layout.length; i++) {
            scanNode(layout[i]);
        }
        return found;
    })(DIALOG_LAYOUT);

    /* 「パネル項目のみ」プリセットで ON にするキーの集合 / The set of keys the Panel-items-only preset turns on */
    var PANEL_ITEM_KEY_SET = (function(layout) {
        var keySet = {};

        function scanNode(layoutNode) {
            if (layoutNode.row) {
                for (var j = 0; j < layoutNode.row.length; j++) {
                    scanNode(layoutNode.row[j]);
                }
            } else if (layoutNode.column) {
                for (var k = 0; k < layoutNode.column.length; k++) {
                    scanNode(layoutNode.column[k]);
                }
            } else if (layoutNode.panelItemsGroup) {
                for (var m = 0; m < layoutNode.keys.length; m++) {
                    keySet[layoutNode.keys[m]] = true;
                }
            }
        }
        for (var i = 0; i < layout.length; i++) {
            scanNode(layout[i]);
        }
        return keySet;
    })(DIALOG_LAYOUT);

    /* 空グループ・空レイヤーの掃除は最後に回す（他の削除で空になった親も同じ実行で消せるように）/ Run container cleanup last so parents emptied by other deletions are removed in the same pass */
    var CONTAINER_CLEANUP_KEYS = {
        emptyGroup: true,
        emptyLayer: true
    };
    var EXECUTION_KEYS = (function(layoutOrderedKeys) {
        var earlierKeys = [];
        var containerCleanupKeys = [];
        for (var i = 0; i < layoutOrderedKeys.length; i++) {
            (CONTAINER_CLEANUP_KEYS[layoutOrderedKeys[i]] ? containerCleanupKeys : earlierKeys).push(layoutOrderedKeys[i]);
        }
        return earlierKeys.concat(containerCleanupKeys);
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

    // ==================================================
    // 実行 / Execution
    // ==================================================

    /* 1ドキュメントに対して選択された種類をすべて実行し、種類キーごとの件数を返す（実行順は EXECUTION_KEYS）
       Run every selected type against one document and return the per-type counts (in EXECUTION_KEYS order) */
    function runCleanup(targetDoc, choices) {
        /* アクションとメニューコマンドは最前面のドキュメントに効くため、対象を必ず前面にする
           Actions and menu commands act on the frontmost document, so bring the target forward first */
        app.activeDocument = targetDoc;

        var counts = {};
        for (var i = 0; i < EXECUTION_KEYS.length; i++) {
            var key = EXECUTION_KEYS[i];
            if (choices[key]) {
                counts[key] = TARGET_RUNNERS[key](targetDoc, choices.force);
            }
        }
        return counts;
    }

    /* 最前面のドキュメントだけを処理する（保存はしない）/ Process only the frontmost document (nothing is saved) */
    function runFrontmostDocument(choices) {
        if (app.documents.length === 0) {
            alert(getLabel('alert.noDocument'));
            return;
        }

        var counts = runCleanup(app.activeDocument, choices);
        app.redraw();
        alert(buildSingleResultMessage(counts, choices));
    }

    /* 開いているすべてのドキュメントを処理する（保存はしない）/ Process every open document (nothing is saved) */
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

        var totals = {};
        var processedCount = 0;
        var failedNames = [];
        for (var j = 0; j < targetDocs.length; j++) {
            try {
                addCounts(totals, runCleanup(targetDocs[j], choices));
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
            buildBatchResultMessage(totals, choices, failedNames));
    }

    /* 指定フォルダー直下の .ai ファイルを順に開いて処理し、上書き保存して閉じる / Open each .ai file directly inside the folder, clean it, save over the original, and close it */
    function runFolderBatch(folder, choices) {
        var aiFiles = collectAiFiles(folder);
        if (aiFiles.length === 0) {
            /* 対象フォルダーと走査した項目数を出して、フォルダー違いと絞り込み漏れを切り分けられるようにする
               Report the folder and how many entries were scanned, so a wrong folder can be told apart from a filtering problem */
            alert(fillPlaceholders(getLabel('alert.folderNoFiles'),
                [formatDisplayPath(folder.fsName), folder.getFiles().length]));
            return;
        }

        /* 元ファイルを上書きして元に戻せないため、実行前に必ず確認する / The originals are overwritten irreversibly, so always confirm first */
        if (!confirm(fillPlaceholders(getLabel('alert.folderConfirm'), [aiFiles.length, folder.fsName]))) {
            return;
        }

        /* 一括処理中はフォント欠落などのダイアログで止まらないようにする / Keep blocking dialogs (missing fonts, etc.) out of the way during the batch */
        var originalInteractionLevel = app.userInteractionLevel;
        app.userInteractionLevel = UserInteractionLevel.DONTDISPLAYALERTS;

        var totals = {};
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
                    addCounts(totals, fileCounts);
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
            buildBatchResultMessage(totals, choices, failedNames));
    }

    /* ファイル名を取得。displayName は OS の表示名なので拡張子の判定には使わず、実ファイル名の name を復号して使う
       Get a file's name from `name` (URI-decoded): `displayName` is the OS display name and isn't reliable for extension matching */
    function getEntryName(entry) {
        return decodeURI(entry.name);
    }

    /* フォルダー直下の .ai ファイルを名前順に集める（サブフォルダー・不可視ファイルは対象外）
       Collect .ai files directly inside the folder, sorted by name (subfolders and invisible files are skipped) */
    function collectAiFiles(folder) {
        var folderEntries = folder.getFiles();
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

    /* 件数の集計に1ドキュメントぶんの結果を足し込む / Add one document's counts into the running totals */
    function addCounts(totals, counts) {
        for (var i = 0; i < ALL_KEYS.length; i++) {
            var key = ALL_KEYS[i];
            if (counts[key] > 0) {
                totals[key] = (totals[key] || 0) + counts[key];
            }
        }
    }

    /* ドキュメント名を取得（取得できない場合は空文字）/ Get a document's name (empty string when unavailable) */
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

    /* 集計1行を組み立て（日本語は「件」を付ける）/ Build one result line (JA appends a counter word) */
    function formatResultLine(key, count) {
        if (currentLanguage === "ja") {
            return getLabel(key) + ": " + count + " 件\n";
        }
        return getLabel(key) + ": " + count + "\n";
    }

    /* 種類ごとの削除件数を行にまとめる（0件の種類は省略）/ Collect the per-type counts into lines (zero-count types are omitted) */
    function buildCountLines(counts, choices) {
        var lines = "";
        for (var i = 0; i < ALL_KEYS.length; i++) {
            var key = ALL_KEYS[i];
            if (choices[key] && counts[key] > 0) {
                lines += formatResultLine('result.' + key, counts[key]);
            }
        }
        return lines;
    }

    /* 見出しと項目名の一覧を組み立て（該当なしなら空文字）/ Build a heading followed by a list of names (empty when there is nothing to list) */
    function buildNoticeLines(headingKey, names) {
        if (names.length === 0) {
            return "";
        }
        var lines = "\n" + getLabel(headingKey) + "\n";
        for (var i = 0; i < names.length; i++) {
            lines += "- " + names[i] + "\n";
        }
        return lines;
    }

    /* アクションを再生できなかった種類の一覧（0件ではなく失敗として明示する）/ List the types whose action couldn't be played, so they aren't reported as a genuine zero */
    function buildActionFailureLines() {
        var names = [];
        for (var i = 0; i < actionFailureKeys.length; i++) {
            names.push(getLabel('result.' + actionFailureKeys[i]));
        }
        return buildNoticeLines('alert.actionFailed', names);
    }

    /* 単一ドキュメント用の完了メッセージ / Completion message for a single document */
    function buildSingleResultMessage(counts, choices) {
        var lines = buildCountLines(counts, choices);
        /* 1件も削除されなければ「対象なし」だけ表示 / If nothing was deleted, show only the no-target message */
        var message = (lines === "") ? getLabel('alert.noTarget') : getLabel('alert.done') + "\n\n" + lines;
        return message + buildActionFailureLines();
    }

    /* 一括処理用の完了メッセージ（合計件数と失敗したファイル名）/ Completion message for a batch run (totals plus the names that failed) */
    function buildBatchResultMessage(totals, choices, failedNames) {
        var lines = buildCountLines(totals, choices);
        var message = (lines === "") ? getLabel('alert.noTarget') : lines;
        return message + buildActionFailureLines() + buildNoticeLines('alert.batchFailed', failedNames);
    }

    // ==================================================
    // ダイアログ生成 / Build the dialog
    // ==================================================

    /* 削除対象を選択するダイアログを構築し、選択結果を返す / Build the dialog and return the user's selection */
    function showDeleteDialog() {
        var dialog = new Window("dialog", getLabel('dialog.title') + " " + SCRIPT_VERSION);
        dialog.orientation = "column";
        dialog.alignChildren = ["fill", "top"];

        /* 最上部に「対象」パネル（最前面／すべて開いている／フォルダー指定）/ The Target panel (frontmost / all open / folder) sits at the top */
        var targetPanel = dialog.add("panel", undefined, getLabel('panel.target'));
        setupPanel(targetPanel, 6);
        targetPanel.helpTip = getLabel('tooltip.target');

        var openDocumentCount = app.documents.length;

        var frontmostRadio = targetPanel.add("radiobutton", undefined, getLabel('target.frontmostDocument'));
        frontmostRadio.helpTip = getLabel('tooltip.frontmostDocument');
        /* ドキュメントが1つも開いていなければ選べない / Not selectable when no document is open */
        frontmostRadio.enabled = (openDocumentCount > 0);

        /* 対象になるドキュメント数をラベルに添える / Show how many documents the option covers */
        var allOpenRadio = targetPanel.add("radiobutton", undefined,
            fillPlaceholders(getLabel('target.allOpenDocuments'), [openDocumentCount]));
        allOpenRadio.helpTip = getLabel('tooltip.allOpenDocuments');
        /* 1つ以下のときは「最前面のドキュメント」と変わらないのでディム表示 / Dimmed at one document or fewer, where it would do the same as the frontmost option */
        allOpenRadio.enabled = (openDocumentCount > 1);

        /* フォルダー指定はラジオと［指定］ボタンを1行に並べる / The folder option puts the radio and the Choose button on one row */
        var folderRow = targetPanel.add("group");
        folderRow.orientation = "row";
        folderRow.alignment = ["fill", "top"];
        folderRow.alignChildren = ["left", "center"];
        folderRow.spacing = 8;

        var folderRadio = folderRow.add("radiobutton", undefined, getLabel('target.targetFolder'));
        folderRadio.helpTip = getLabel('tooltip.targetFolder');
        /* ラジオも生成後に伸びないので、件数を付けた状態の幅を先に確保する / A radiobutton doesn't grow after creation either, so reserve the width the counted label needs */
        folderRadio.preferredSize.width = FOLDER_RADIO_WIDTH;
        var chooseFolderButton = folderRow.add("button", undefined, getLabel('button.chooseFolder'));
        chooseFolderButton.helpTip = getLabel('tooltip.chooseFolder');

        /* 選択中のパスは次の行に、ラジオのラベル位置に合わせて字下げして表示 / The chosen path goes on the next line, indented to line up with the radio's label */
        var folderPathRow = targetPanel.add("group");
        folderPathRow.orientation = "row";
        folderPathRow.alignment = ["fill", "top"];
        folderPathRow.alignChildren = ["left", "center"];
        folderPathRow.margins = [FOLDER_PATH_INDENT, 0, 0, 0];

        /* statictext は生成後に伸びないので、長いパスを収める幅を先に確保する / A statictext never grows after creation, so reserve enough width for a long path up front */
        var folderPathText = folderPathRow.add("statictext", undefined, getLabel('target.noFolderChosen'), {
            truncate: "middle"
        });
        folderPathText.preferredSize.width = FOLDER_PATH_WIDTH;
        folderPathText.alignment = ["fill", "center"];

        var selectedFolder = null;

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
        selectDefaultTargetRadio();

        /* 選んだフォルダーを保持して表示に反映する（選択ダイアログ経由でもセッション復元でも通る）
           Hold the chosen folder and reflect it in the display; used both by the picker and by the session restore */
        function applyChosenFolder(folder) {
            selectedFolder = folder;
            /* 表示は Dropbox 配下ならプレフィックスを落とし、そうでなければ ~ に短縮。幅に収まらないぶんは中央が省略されるため、末尾のフォルダー名は残る
               Drop the Dropbox prefix when it applies, otherwise show the ~ form; middle truncation keeps the trailing folder name readable */
            folderPathText.text = formatDisplayPath(folder.fsName);
            folderPathText.helpTip = folder.fsName;
            /* 処理対象になる .ai ファイル数をラベルに添える（0件なら実行前に気づける）/ Show how many .ai files the run will cover, so a zero is obvious before running */
            folderRadio.text = fillPlaceholders(getLabel('target.targetFolderWithCount'), [collectAiFiles(folder).length]);
        }

        /* フォルダー選択ダイアログを開き、選んだフォルダーを表示に反映する（キャンセル時は false）
           Open the folder picker and reflect the choice in the display; returns false when cancelled */
        function chooseTargetFolder() {
            var folder = Folder.selectDialog(getLabel('prompt.selectFolder'), selectedFolder);
            if (!folder) {
                return false;
            }
            applyChosenFolder(folder);
            return true;
        }

        frontmostRadio.onClick = function() {
            selectTargetRadio(frontmostRadio);
        };

        allOpenRadio.onClick = function() {
            selectTargetRadio(allOpenRadio);
        };

        chooseFolderButton.onClick = function() {
            /* ［指定］を押したらフォルダー指定モードに切り替える / Pressing Choose switches the target to folder mode */
            if (chooseTargetFolder()) {
                selectTargetRadio(folderRadio);
            }
        };

        folderRadio.onClick = function() {
            /* 未指定のままフォルダー指定を選んだら、その場で選択させる。キャンセル時はディム表示のラジオに戻さない
               Picking folder mode with nothing chosen opens the picker right away; cancelling must not land on a dimmed radio */
            if (selectedFolder === null && !chooseTargetFolder()) {
                selectDefaultTargetRadio();
                return;
            }
            selectTargetRadio(folderRadio);
        };

        /* レイアウト定義どおりにパネル・サブパネル・チェックボックス／ラジオを生成 / Build panels, sub-panels, and their checkboxes/radios per the layout */
        var optionControls = {};
        var forceCheckbox = null;
        /* プリセットが値を書き込むチェックボックスの一覧。{ checkbox, key, initialValue } の組で持ち、
           ScriptUI ウィジェットに独自プロパティを生やさない。強制オプションは含めない（ガイドのラジオは RADIO_GROUP 経由）
           The checkboxes a preset writes to, held as { checkbox, key, initialValue } records rather than as custom
           properties on the ScriptUI widgets; the force option is never included (guide radios go through RADIO_GROUP) */
        var checkboxEntries = [];
        var presetDropdown = null;
        /* 表示を合わせ直している最中かどうか。ScriptUI は selection への代入でも onChange を呼ぶため、
           そのままだと手で変えた直後にプリセットが再適用されてしまう
           Whether the display is being resynced; ScriptUI fires onChange on assignment to selection too,
           which would otherwise re-apply the preset right after a hand-made change */
        var isSyncingPreset = false;

        /* 手で選択を変えたらプリセット表示を今の状態に合わせ直す（多くは「カスタム」になる）
           After a hand-made change, resync the preset display with the current state (usually landing on Custom) */
        function refreshPresetSelection() {
            if (!presetDropdown) {
                return;
            }
            isSyncingPreset = true;
            try {
                presetDropdown.selection = detectPreset(checkboxEntries, optionControls);
            } finally {
                isSyncingPreset = false;
            }
        }

        /* レイアウト定義1件ぶんのパネルを親コンテナに生成 / Build the panel for one layout node inside a parent container */
        function buildPanelNode(parentContainer, layoutNode) {
            /* column ノードは複数パネルを縦積みするグループ / A column node stacks several panels vertically */
            if (layoutNode.column) {
                var panelStack = parentContainer.add("group");
                panelStack.orientation = "column";
                panelStack.alignChildren = ["fill", "top"];
                panelStack.alignment = "fill";
                panelStack.spacing = PANEL_SPACING;
                for (var k = 0; k < layoutNode.column.length; k++) {
                    buildPanelNode(panelStack, layoutNode.column[k]);
                }
                return;
            }

            var sectionPanel = parentContainer.add("panel", undefined, getLabel(layoutNode.titleKey));
            setupPanel(sectionPanel, 6);
            if (layoutNode.radio) {
                addRadioGroup(sectionPanel, layoutNode, optionControls, refreshPresetSelection);
            } else {
                /* option＋クリックの一括切り替えが効くのはチェックボックスのパネルだけ / Only checkbox panels support the option-click bulk toggle */
                sectionPanel.helpTip = getLabel('tooltip.optionClickToggleAll');
                checkboxEntries = checkboxEntries.concat(
                    addCheckboxes(sectionPanel, layoutNode.keys, optionControls, refreshPresetSelection));
            }

            /* 強制オプションはグループに入れて対象パネル（パネル項目）の末尾に追加 / The force option sits in a group at the bottom of its panel (panel items) */
            if (layoutNode.force) {
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
                forceCheckbox = forceGroup.add("checkbox", undefined, getLabel('checkbox.force'));
                forceCheckbox.value = false;
                forceCheckbox.helpTip = getLabel('tooltip.force');
            }
        }

        for (var i = 0; i < DIALOG_LAYOUT.length; i++) {
            var layoutNode = DIALOG_LAYOUT[i];
            if (layoutNode.row) {
                /* 複数パネルを横並び / Lay multiple panels side by side */
                var panelRow = dialog.add("group");
                panelRow.orientation = "row";
                panelRow.alignChildren = ["fill", "top"];
                panelRow.alignment = "fill";
                panelRow.spacing = PANEL_SPACING;
                for (var j = 0; j < layoutNode.row.length; j++) {
                    buildPanelNode(panelRow, layoutNode.row[j]);
                }
            } else {
                buildPanelNode(dialog, layoutNode);
            }
        }

        /* ボタンエリアは 左（削除対象プリセット）／中央スペーサー／右（キャンセル・実行）の3カラム
           Button area is three columns: left (deletion-target preset), center spacer, right (cancel / run) */
        var buttonRow = dialog.add("group");
        buttonRow.orientation = "row";
        buttonRow.alignment = "fill";
        buttonRow.alignChildren = ["fill", "center"];

        var buttonLeft = buttonRow.add("group");
        buttonLeft.orientation = "row";
        buttonLeft.alignment = ["left", "center"];
        buttonLeft.add("statictext", undefined, getLabel('preset.label'));
        presetDropdown = buttonLeft.add("dropdownlist", undefined, [
            getLabel('preset.basic'),
            getLabel('preset.allOff'),
            getLabel('preset.allOn'),
            getLabel('preset.panelItemsOnly'),
            getLabel('preset.custom')
        ]);
        presetDropdown.helpTip = getLabel('tooltip.preset');
        presetDropdown.preferredSize.width = PRESET_DROPDOWN_WIDTH;
        /* 初期状態は「基本」そのもの / The initial state is exactly the Default preset */
        presetDropdown.selection = PRESET_BASIC;
        presetDropdown.onChange = function() {
            /* 表示を合わせ直しただけの代入では、プリセットを適用し直さない / An assignment that only resyncs the display must not re-apply the preset */
            if (isSyncingPreset) {
                return;
            }
            if (this.selection) {
                applyCheckboxPreset(checkboxEntries, optionControls, this.selection.index);
            }
        };

        /* 中央のスペーサーが余白を吸収して左右を両端に寄せる / The center spacer absorbs slack, pushing the two sides apart */
        var buttonSpacer = buttonRow.add("group");
        buttonSpacer.alignment = ["fill", "center"];
        buttonSpacer.minimumSize.width = 1;

        var buttonRight = buttonRow.add("group");
        buttonRight.orientation = "row";
        buttonRight.alignment = ["right", "center"];
        var cancelButton = buttonRight.add("button", undefined, getLabel('button.cancel'), {
            name: "cancel"
        });
        /* キャンセルでも位置だけは覚える（onClick を付けたので明示的に閉じる）
           Remember the position even on cancel; onClick replaces the default close, so close explicitly */
        cancelButton.onClick = function() {
            rememberDialogLocation();
            dialog.close(2);
        };

        var runButton = buttonRight.add("button", undefined, getLabel('button.run'), {
            name: "ok"
        });
        runButton.helpTip = getLabel('tooltip.run');

        /* フォルダー未指定のまま実行させない（onClick を付けたので明示的に閉じる）
           Don't let the run start without a folder; onClick replaces the default close, so close explicitly */
        runButton.onClick = function() {
            if (folderRadio.value && selectedFolder === null) {
                alert(getLabel('alert.folderNotChosen'));
                return;
            }
            rememberDialogLocation();
            dialog.close(1);
        };

        /* 前回の位置を覚えておく / Remember where the dialog was left */
        function rememberDialogLocation() {
            sessionState.dialogLocation = tryGet(function() {
                return [dialog.location[0], dialog.location[1]];
            }, null);
        }

        /* 前回の選択をセッションから復元する（記憶がなければ何もしない）/ Restore the previous selection from the session (a no-op when there is none) */
        function restoreSessionChoices() {
            var saved = sessionState.choices;
            if (!saved) {
                return;
            }

            /* チェックボックス（ガイドのラジオはこの後まとめて設定する）/ Checkboxes; the guide radios are set together below */
            for (var i = 0; i < checkboxEntries.length; i++) {
                checkboxEntries[i].checkbox.value = (saved[checkboxEntries[i].key] === true);
            }

            /* ガイドは1つだけONにする / Exactly one guide option is on */
            if (RADIO_GROUP) {
                var chosenGuideKey = null;
                for (var j = 0; j < RADIO_GROUP.keys.length; j++) {
                    if (saved[RADIO_GROUP.keys[j]]) {
                        chosenGuideKey = RADIO_GROUP.keys[j];
                        break;
                    }
                }
                selectGuideRadio(optionControls, chosenGuideKey);
            }

            if (forceCheckbox) {
                forceCheckbox.value = (saved.force === true);
            }

            /* 処理対象。今のドキュメント数で選べない項目と、消えたフォルダーは復元しない
               The scope; an option the current document count disables, or a folder that is gone, is not restored */
            if (saved.targetMode === TARGET_ALL_OPEN && allOpenRadio.enabled) {
                selectTargetRadio(allOpenRadio);
            } else if (saved.targetMode === TARGET_FOLDER && sessionState.folderPath) {
                var savedFolder = new Folder(sessionState.folderPath);
                if (savedFolder.exists) {
                    applyChosenFolder(savedFolder);
                    selectTargetRadio(folderRadio);
                }
            }
        }

        restoreSessionChoices();
        /* 復元後の状態に合わせてプリセット表示を決める / Pick the preset display that matches the restored state */
        refreshPresetSelection();

        /* 前回の位置を再現する。画面構成が変わって画面外になる位置は使わない
           Reuse the previous position, unless a display change would leave it off-screen */
        dialog.onShow = function() {
            if (sessionState.dialogLocation && isLocationOnScreen(sessionState.dialogLocation)) {
                dialog.location = sessionState.dialogLocation;
            }
        };

        if (dialog.show() !== 1) {
            return null;
        }

        /* 対象の指定と各項目の選択状態を種類キーごとにまとめて返す / Collect the chosen target and every option's state, keyed by type */
        var choices = {
            force: forceCheckbox ? forceCheckbox.value : false,
            targetMode: folderRadio.value ? TARGET_FOLDER : (allOpenRadio.value ? TARGET_ALL_OPEN : TARGET_FRONTMOST),
            targetFolder: selectedFolder
        };
        for (i = 0; i < ALL_KEYS.length; i++) {
            choices[ALL_KEYS[i]] = optionControls[ALL_KEYS[i]].value;
        }

        /* 次回の起動で再現できるようセッションに控える。Folder オブジェクトは持ち越さず、パス文字列だけを残す
           Stash it in the session for the next run; the Folder object is dropped and only its path is kept */
        sessionState.choices = copyChoicesForSession(choices);
        sessionState.folderPath = selectedFolder ? selectedFolder.fsName : null;

        return choices;
    }

    /* セッションに残す用の複製。targetFolder（Folder オブジェクト）は folderPath に置き換わるので持ち越さない
       Copy for the session; targetFolder (a Folder object) is left out because folderPath replaces it */
    function copyChoicesForSession(choices) {
        var copy = {
            force: choices.force,
            targetMode: choices.targetMode
        };
        for (var i = 0; i < ALL_KEYS.length; i++) {
            copy[ALL_KEYS[i]] = choices[ALL_KEYS[i]];
        }
        return copy;
    }

    /* 保存した位置がいずれかの画面に収まるか / Whether a saved position still lands on one of the screens */
    function isLocationOnScreen(location) {
        var screens = tryGet(function() {
            return $.screens;
        }, null);
        if (!screens || !screens.length) {
            return false;
        }
        for (var i = 0; i < screens.length; i++) {
            var screen = screens[i];
            if (location[0] >= screen.left && location[0] <= screen.right - ON_SCREEN_MARGIN &&
                location[1] >= screen.top && location[1] <= screen.bottom - ON_SCREEN_MARGIN) {
                return true;
            }
        }
        return false;
    }

    /* キー配列ぶんのチェックボックスを親に追加し、参照を記録して { checkbox, key, initialValue } の配列を返す。
       initialValue は「基本」プリセットで元に戻すために控える
       Add a checkbox per key, record the reference, and return { checkbox, key, initialValue } records;
       initialValue is kept so the Default preset can restore it */
    function addCheckboxes(parentContainer, keys, optionControls, onManualChange) {
        var panelEntries = [];
        for (var i = 0; i < keys.length; i++) {
            var checkbox = parentContainer.add("checkbox", undefined, getLabel('checkbox.' + keys[i]));
            checkbox.value = !UNCHECKED_BY_DEFAULT[keys[i]];
            checkbox.helpTip = getLabel('tooltip.' + keys[i]);
            optionControls[keys[i]] = checkbox;
            panelEntries.push({
                checkbox: checkbox,
                key: keys[i],
                initialValue: checkbox.value
            });
        }
        enableOptionClickToggleAll(panelEntries, onManualChange);
        return panelEntries;
    }

    /* option（Alt）キーが押されているか / Whether the option/alt key is held down */
    function isOptionKeyDown() {
        return tryGet(function() {
            return ScriptUI.environment.keyboardState.altKey === true;
        }, false);
    }

    /* option（Alt）＋クリックで、渡したチェックボックス全部をクリック先の状態に揃える。修飾キーなしのクリックは通常どおり
       Option/alt-click sets every given checkbox to the clicked one's new state; an unmodified click toggles just that one */
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

    /* ガイドのラジオで選ばれているキーを返す。「削除しない」のときは null / The selected guide key, or null when "Don't delete" is on */
    function getSelectedGuideKey(optionControls) {
        if (!RADIO_GROUP) {
            return null;
        }
        for (var i = 0; i < RADIO_GROUP.keys.length; i++) {
            if (optionControls[RADIO_GROUP.keys[i]].value) {
                return RADIO_GROUP.keys[i];
            }
        }
        return null;
    }

    /* ガイドのラジオを1つだけONにする。null なら「削除しない」/ Turn on exactly one guide radio; null means "Don't delete" */
    function selectGuideRadio(optionControls, chosenKey) {
        if (!RADIO_GROUP) {
            return;
        }
        optionControls[RADIO_GROUP.noneKey].value = (chosenKey === null);
        for (var i = 0; i < RADIO_GROUP.keys.length; i++) {
            optionControls[RADIO_GROUP.keys[i]].value = (RADIO_GROUP.keys[i] === chosenKey);
        }
    }

    /* プリセットに合わせてチェックとガイドのラジオを一括設定（「カスタム」は現状維持）
       Apply a preset to every checkbox and to the guide radios (Custom leaves them as they are) */
    function applyCheckboxPreset(checkboxEntries, optionControls, presetIndex) {
        if (presetIndex === PRESET_CUSTOM) {
            return;
        }
        for (var i = 0; i < checkboxEntries.length; i++) {
            var entry = checkboxEntries[i];
            if (presetIndex === PRESET_ALL_ON) {
                entry.checkbox.value = true;
            } else if (presetIndex === PRESET_ALL_OFF) {
                entry.checkbox.value = false;
            } else if (presetIndex === PRESET_PANEL_ITEMS_ONLY) {
                /* すべてOFFにしてから、パネル項目だけをONにする / Everything off, then just the panel items back on */
                entry.checkbox.value = (PANEL_ITEM_KEY_SET[entry.key] === true);
            } else {
                entry.checkbox.value = entry.initialValue;
            }
        }
        /* 「すべてON」はガイドも「すべてのガイド」に。それ以外は既定の「削除しない」
           All-on also picks the all-guides option; the others fall back to "Don't delete" */
        var chosenGuideKey = (RADIO_GROUP && presetIndex === PRESET_ALL_ON) ? RADIO_GROUP.allOnKey : null;
        selectGuideRadio(optionControls, chosenGuideKey);
    }

    /* 現在の選択がどのプリセットに当たるかを判定する / Work out which preset the current selection matches */
    function detectPreset(checkboxEntries, optionControls) {
        var matchesInitial = true;
        var allOn = true;
        var allOff = true;
        var panelItemsOnly = true;

        for (var i = 0; i < checkboxEntries.length; i++) {
            var entry = checkboxEntries[i];
            if (entry.checkbox.value !== entry.initialValue) {
                matchesInitial = false;
            }
            if (entry.checkbox.value) {
                allOff = false;
            } else {
                allOn = false;
            }
            if (entry.checkbox.value !== (PANEL_ITEM_KEY_SET[entry.key] === true)) {
                panelItemsOnly = false;
            }
        }

        /* ガイドの既定は「削除しない」なので、選ばれていれば初期状態ではない / The guide default is "Don't delete", so any choice means it isn't the initial state */
        if (RADIO_GROUP) {
            var guideKey = getSelectedGuideKey(optionControls);
            if (guideKey !== null) {
                matchesInitial = false;
                allOff = false;
                panelItemsOnly = false;
            }
            if (guideKey !== RADIO_GROUP.allOnKey) {
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

    /* ラジオグループを生成。先頭に「削除しない」（既定で選択）を置き、各キーのラジオを記録。
       プリセット判定と復元で使うため、「削除しない」も noneKey で記録する
       Build a radio group with a "Don't delete" option (selected by default) first and record each key's radio;
       the none option is recorded under noneKey too, since presets and restoring need it */
    function addRadioGroup(parentPanel, layoutNode, optionControls, onManualChange) {
        var radios = [];
        var noneRadio = parentPanel.add("radiobutton", undefined, getLabel('checkbox.' + layoutNode.noneKey));
        noneRadio.helpTip = getLabel('tooltip.' + layoutNode.noneKey);
        noneRadio.value = true;
        optionControls[layoutNode.noneKey] = noneRadio;
        radios.push(noneRadio);

        for (var i = 0; i < layoutNode.keys.length; i++) {
            var radio = parentPanel.add("radiobutton", undefined, getLabel('checkbox.' + layoutNode.keys[i]));
            radio.helpTip = getLabel('tooltip.' + layoutNode.keys[i]);
            radio.value = false;
            optionControls[layoutNode.keys[i]] = radio;
            radios.push(radio);
        }

        /* ガイドを手で切り替えたときもプリセット表示を追従させる / Keep the preset display in step when a guide option is picked by hand */
        for (var j = 0; j < radios.length; j++) {
            radios[j].onClick = onManualChange;
        }
    }

    // ==================================================
    // 一時アクション / Temporary action
    // ==================================================

    /* ASCII 文字列を16進に変換 / Convert an ASCII string to hex */
    function asciiToHex(text) {
        var hex = "";
        for (var i = 0; i < text.length; i++) {
            var code = text.charCodeAt(i).toString(16);
            if (code.length < 2) {
                code = "0" + code;
            }
            hex += code;
        }
        return hex;
    }

    /* メニューコマンド1件のイベントブロックを組み立て / Build one menu-command event block */
    function buildMenuEventBlock(eventIndex, internalName, localizedNameHex, commandNameHex, value, hasDialog) {
        var lines = [
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
            lines.push("\t\t/showDialog 0");
        }
        lines.push(
            "\t\t/parameterCount 1",
            "\t\t/parameter-1 {",
            "\t\t\t/key 1835363957",
            "\t\t\t/showInPalette 1",
            "\t\t\t/type (enumerated)",
            "\t\t\t/name [ " + (commandNameHex.length / 2),
            "\t\t\t\t" + commandNameHex,
            "\t\t\t]",
            "\t\t\t/value " + value,
            "\t\t}",
            "\t}"
        );
        return lines.join("\n");
    }

    /* 「未使用をすべて選択 → 削除」の一時アクション定義を録画値から組み立て / Build the temporary "Select All Unused -> Delete" action from recorded values */
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

    /* 失敗しても無視してよい後始末を実行 / Run best-effort cleanup, ignoring any failure */
    function ignoringErrors(action) {
        try {
            action();
        } catch (e) {
            /* 後始末の失敗は無視 / Ignore cleanup failures */
        }
    }

    /* アクションを一時ファイルに書き出して再生。close/unload/remove は finally で必ず試みる / Write, load, and play the action; close/unload/remove are always attempted in finally */
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

    /* 指定コレクションの未使用項目を、対応するアクションを再生して削除し、件数を返す。再生に失敗した種類は resultKey を控えて完了時に警告する（「未使用0件」と区別するため）
       Prune a collection's unused items via its action and return the count; a failed playback records resultKey so the summary can warn about it (instead of looking like a genuine zero) */
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
    // アートボードの空判定 / Artboard emptiness check
    // ==================================================

    /* アートワークの外接矩形を一度だけ集める（ガイドは対象外）。アートボードごとに全 pageItems を走査し直さないため
       Collect artwork bounds once (guides excluded), so each artboard doesn't rescan every pageItem */
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

    /* 集めておいた外接矩形を使い、アートボード矩形にアートワークが載っているか判定
       Using the collected bounds, determine whether any artwork sits on the artboard rectangle */
    function isArtboardEmpty(artworkBounds, artboardRect) {
        for (var i = 0; i < artworkBounds.length; i++) {
            if (rectsIntersect(artboardRect, artworkBounds[i])) {
                return false;
            }
        }
        return true;
    }

    /* 2つの矩形 [left, top, right, bottom] が重なるか判定（Illustrator は上が大きいY）/ Whether two [left, top, right, bottom] rects overlap (Illustrator: top has the larger Y) */
    function rectsIntersect(rectA, rectB) {
        if (rectB[2] < rectA[0]) {
            return false;
        } /* B 右端が A 左端より左 / B right is left of A left */
        if (rectB[0] > rectA[2]) {
            return false;
        } /* B 左端が A 右端より右 / B left is right of A right */
        if (rectB[3] > rectA[1]) {
            return false;
        } /* B 下端が A 上端より上 / B bottom is above A top */
        if (rectB[1] < rectA[3]) {
            return false;
        } /* B 上端が A 下端より下 / B top is below A bottom */
        return true;
    }

    /* 矩形がいずれかの矩形と重なるか / Whether a rect overlaps any of the rects */
    function intersectsAnyRect(bounds, rects) {
        for (var i = 0; i < rects.length; i++) {
            if (rectsIntersect(rects[i], bounds)) {
                return true;
            }
        }
        return false;
    }

    /* すべてのアートボードの矩形を取得 / Get the rectangles of all artboards */
    function getAllArtboardRects(doc) {
        var rects = [];
        for (var i = 0; i < doc.artboards.length; i++) {
            rects.push(doc.artboards[i].artboardRect);
        }
        return rects;
    }

    // ==================================================
    // 孤立点・空テキスト削除 / Stray points and empty text
    // 移植元 / Ported from: 不要なアイテムを削除.jsx (c) 2020 Toshiyuki Takahashi, MIT License
    // ==================================================

    /* 孤立点（アンカー1点・長さ0のパス）を削除し、件数を返す / Remove stray points (single-anchor, zero-length paths), return the count */
    function deleteStrayPoints(doc) {
        var removedCount = 0;
        for (var i = doc.pageItems.length - 1; i >= 0; i--) {
            var item = doc.pageItems[i];
            if (item.typename !== "PathItem") {
                continue;
            }
            try {
                if (item.pathPoints.length < 2 && item.length <= 0) {
                    item.remove();
                    removedCount++;
                }
            } catch (e) {
                /* 削除不可 / Not removable */
            }
        }
        return removedCount;
    }

    /* 文字のない空テキストを削除し、件数を返す（エリア内/パス上は塗り・線のないパスのみ）/ Remove empty text frames, return the count (area/path text only when the path has no fill/stroke) */
    function deleteEmptyTextFrames(doc) {
        var removedCount = 0;
        for (var i = doc.pageItems.length - 1; i >= 0; i--) {
            var item = doc.pageItems[i];
            if (item.typename !== "TextFrame") {
                continue;
            }
            try {
                if (item.contents.length >= 1) {
                    continue;
                }
                var shouldRemove = false;
                if (item.kind === TextType.POINTTEXT) {
                    shouldRemove = true;
                } else if (item.kind === TextType.AREATEXT || item.kind === TextType.PATHTEXT) {
                    /* テキストパスが塗り・線なしのときだけ削除 / Remove only when the text path has no fill/stroke */
                    if (!item.textPath.stroked && !item.textPath.filled) {
                        shouldRemove = true;
                    }
                }
                if (shouldRemove) {
                    item.remove();
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

    /* 塗りも線もない（不可視の）パスを削除し、件数を返す。ガイド・クリッピングパスは除外 / Remove paths with no fill and no stroke (invisible) and return the count; guides and clipping paths are excluded */
    function deleteUnpaintedPaths(doc) {
        var removedCount = 0;
        var paths = doc.pathItems;
        for (var i = paths.length - 1; i >= 0; i--) {
            var pathItem = paths[i];
            try {
                /* ガイドは別オプションで扱う / Guides are handled by a separate option */
                if (pathItem.guides) {
                    continue;
                }
                /* クリッピングパスはグループの一部なので残す / Keep clipping paths (part of a clip group) */
                if (pathItem.clipping) {
                    continue;
                }
                /* コンパウンドパスの構成パス（穴など）は単体で削除しない / Don't delete a compound path's member paths (holes, etc.) */
                if (pathItem.parent && pathItem.parent.typename === "CompoundPathItem") {
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

    /* 不透明度が0%のオブジェクトを削除し、件数を返す（グループ内も対象）。単体で消すと構造が壊れるもの、
       および別オプションで扱うガイドは除外する
       Remove objects at 0% opacity and return the count (including inside groups); items whose individual removal
       would break a structure, and guides (handled by their own option), are excluded */
    function deleteZeroOpacityObjects(doc) {
        var removedCount = 0;

        /* doc.pageItems はグループ内も含む平坦なコレクション / doc.pageItems is a flat collection that includes items inside groups */
        for (var i = doc.pageItems.length - 1; i >= 0; i--) {
            var item = doc.pageItems[i];
            if (item.typename === "GroupItem") {
                continue;
            }
            try {
                if (item.typename === "PathItem") {
                    /* ガイドは「ガイド」セクションでのみ削除する / Guides are only removed by the guide section */
                    if (item.guides) {
                        continue;
                    }
                    /* マスクを消すとクリップが解除され、隠れていた中身が現れてしまう / Removing the mask releases the clip and reveals what it was hiding */
                    if (item.clipping) {
                        continue;
                    }
                    /* コンパウンドパスの構成パス（穴など）は単体で削除しない / Don't delete a compound path's member paths (holes, etc.) */
                    if (item.parent && item.parent.typename === "CompoundPathItem") {
                        continue;
                    }
                }
                if (item.opacity === 0) {
                    item.remove();
                    removedCount++;
                }
            } catch (e) {
                /* 削除不可 / Not removable */
            }
        }
        return removedCount;
    }

    /* 非表示オブジェクトを削除し、件数を返す（非表示グループは中身ごと）/ Remove hidden objects and return the count (hidden groups go with their contents) */
    function deleteHiddenObjects(doc) {
        /* まず参照だけを収集（この間はコレクションを変更しない）。非表示グループを消すと子のインデックスがずれ、末尾からの走査でも取りこぼすため
           Collect references first without mutating the collection; removing a hidden group shifts child indices, so even a reverse scan would skip items */
        var hiddenItems = [];
        var items = doc.pageItems;
        for (var i = 0; i < items.length; i++) {
            try {
                if (items[i].hidden) {
                    hiddenItems.push(items[i]);
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

    /* リンク切れ（リンク先が見つからない）の配置画像を削除し、件数を返す。埋め込み・正常リンクは対象外 / Remove placed images with a missing link and return the count; embedded images and valid links are kept */
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

    /* どのアートボードにも載っていないオブジェクトを削除し、件数を返す / Remove objects that sit on none of the artboards, return the count */
    function deleteObjectsOutsideAllArtboards(doc) {
        return deleteObjectsOutsideRects(doc, getAllArtboardRects(doc));
    }

    /* アクティブなアートボードに載っていないオブジェクトを削除し、件数を返す / Remove objects not on the active artboard, return the count */
    function deleteObjectsOutsideActiveArtboard(doc) {
        var activeIndex = doc.artboards.getActiveArtboardIndex();
        return deleteObjectsOutsideRects(doc, [doc.artboards[activeIndex].artboardRect]);
    }

    /* 指定矩形のいずれにも重ならないトップレベルオブジェクトを削除し、件数を返す / Remove top-level objects overlapping none of the given rects, return the count */
    function deleteObjectsOutsideRects(doc, rects) {
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
            var bounds;
            try {
                bounds = topLevelItem.visibleBounds;
            } catch (e) {
                continue;
            }
            if (!intersectsAnyRect(bounds, rects)) {
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

    /* 空のグループ（通常グループ・クリップグループ）を削除し、件数を返す / Remove empty groups (ordinary and clip groups) and return the count */
    function deleteEmptyGroups(doc) {
        var removedCount = 0;
        for (var i = 0; i < doc.layers.length; i++) {
            removedCount += removeEmptyGroupsIn(doc.layers[i]);
        }
        return removedCount;
    }

    /* コンテナ内を再帰的に探索し、空のグループを削除して件数を返す。子を先に掃除するので、空になった親も同じパスで削除できる / Recurse a container removing empty groups; children are cleaned first so a parent that becomes empty is removed in the same pass */
    function removeEmptyGroupsIn(container) {
        var removed = 0;
        /* サブレイヤーの中身は layer.pageItems に含まれないため、先に再帰する / Sublayer contents aren't in layer.pageItems, so recurse into sublayers first */
        if (container.typename === "Layer") {
            for (var j = container.layers.length - 1; j >= 0; j--) {
                removed += removeEmptyGroupsIn(container.layers[j]);
            }
        }
        /* 削除でインデックスがずれるため末尾から / Iterate from the end because removal shifts indices */
        for (var i = container.pageItems.length - 1; i >= 0; i--) {
            var item = container.pageItems[i];
            if (item.typename === "GroupItem") {
                /* 先に中を掃除してから自身の空判定 / Clean inside first, then test this group */
                removed += removeEmptyGroupsIn(item);
                if (isEmptyGroup(item)) {
                    try {
                        item.remove();
                        removed++;
                    } catch (e) {
                        /* 削除不可 / Not removable */
                    }
                }
            }
        }
        return removed;
    }

    /* グループが空か判定。子が無いグループ、またはマスク以外が塗り・線なしのパスだけのクリップグループ / Whether a group is empty: no children, or a clip group whose non-mask contents are only paths with no fill/stroke */
    function isEmptyGroup(item) {
        if (item.typename !== "GroupItem") {
            return false;
        }

        var children = item.pageItems;
        /* 子のないグループは空 / A group with no children is empty */
        if (children.length === 0) {
            return true;
        }

        /* クリップグループのみ、中身が塗り・線なしパスだけなら空とみなす / Only for clip groups: empty when all contents are unpainted paths */
        if (item.clipped !== true) {
            return false;
        }

        for (var i = 0; i < children.length; i++) {
            var child = children[i];
            if (child.typename !== "PathItem") {
                return false;
            }
            if (child.filled === true && child.fillColor.typename !== "NoColor") {
                return false;
            }
            if (child.stroked === true && child.strokeColor.typename !== "NoColor") {
                return false;
            }
        }
        return true;
    }

    /* ガイド属性を持つパスの数を数える / Count paths flagged as guides */
    function countGuidePaths(doc) {
        var count = 0;
        var paths = doc.pathItems;
        for (var i = 0; i < paths.length; i++) {
            if (paths[i].guides) {
                count++;
            }
        }
        return count;
    }

    /* メニューコマンド「ガイドを消去」でガイドを削除し、件数を返す。ロック済みガイドは残す（一時解除しない）/ Remove guides via the Clear Guides menu command and return the count; locked guides are kept (no temporary unlock) */
    function clearGuides(doc) {
        var countBefore = countGuidePaths(doc);
        app.executeMenuCommand("clearguide");
        return Math.max(0, countBefore - countGuidePaths(doc));
    }

    /* サブレイヤーを含む全レイヤーを親→子の順に集める / Collect every layer and sublayer, parents before children */
    function collectLayersDeep(layerCollection, collectedLayers) {
        for (var i = 0; i < layerCollection.length; i++) {
            var layer = layerCollection[i];
            collectedLayers.push(layer);
            collectLayersDeep(layer.layers, collectedLayers);
        }
        return collectedLayers;
    }

    /* ガイドのロック・レイヤーロック（サブレイヤー含む）・ガイド自体のロックを一時解除し、判定関数が真のガイドを削除して件数を返す。ロック状態は finally で必ず復元
       Temporarily clear guide locks, layer locks (sublayers included), and each guide's own lock; remove guides for which the predicate is true and return the count. Locks are always restored in finally */
    function removeGuidesWhere(doc, shouldRemove) {
        var removedCount = 0;

        /* サブレイヤーまで含めてロック状態を保存し、親から順に解除 / Save every lock state down to sublayers and clear them parents-first */
        var layers = collectLayersDeep(doc.layers, []);
        var lockStates = [];
        for (var i = 0; i < layers.length; i++) {
            lockStates[i] = false;
            try {
                lockStates[i] = layers[i].locked;
                if (lockStates[i]) {
                    layers[i].locked = false;
                }
            } catch (e) {
                /* 解除できないレイヤーはそのまま / Leave layers we can't unlock */
            }
        }

        var guidesWereLocked = doc.guidesLocked;

        try {
            doc.guidesLocked = false;

            var paths = doc.pathItems;
            for (var k = paths.length - 1; k >= 0; k--) {
                var guidePath = paths[k];
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
            /* 例外時もロック状態を必ず元に戻す。子から順に戻して親のロックに邪魔されないようにする
               Always restore the lock states, even on error; restore children first so a re-locked parent doesn't block them */
            for (var j = layers.length - 1; j >= 0; j--) {
                try {
                    layers[j].locked = lockStates[j];
                } catch (e3) {
                    /* 復元できない場合は無視 / Ignore when it can't be restored */
                }
            }
            doc.guidesLocked = guidesWereLocked;
        }

        return removedCount;
    }

    /* すべてのガイドを削除し、件数を返す / Remove all guides and return the count */
    function deleteAllGuides(doc) {
        return removeGuidesWhere(doc, function() {
            return true;
        });
    }

    /* 現在（アクティブ）のアートボード上にないガイドを削除し、件数を返す / Remove guides not on the active artboard and return the count */
    function deleteGuidesOutsideActiveArtboard(doc) {
        var activeRect = doc.artboards[doc.artboards.getActiveArtboardIndex()].artboardRect;
        return removeGuidesWhere(doc, function(guidePath) {
            return !rectsIntersect(activeRect, guidePath.geometricBounds);
        });
    }

    // ==================================================
    // 関数：空のレイヤー削除 / Delete empty layers
    // ==================================================

    /* 中身が空（pageItems もサブレイヤーも無い）のレイヤーを削除し、件数を返す。ガイドは pageItems に含まれるためガイドのみのレイヤーは残る。トップレベルは最低1つ残す
       Remove empty layers (no pageItems and no sublayers) and return the count; guides count as pageItems so guide-only layers stay, and at least one top-level layer remains */
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

    /* サブレイヤーを再帰的に処理し、空のものを削除して件数を返す / Recurse sublayers, removing empty ones, return the count */
    function removeEmptySublayers(parentLayer) {
        var removed = 0;
        for (var i = parentLayer.layers.length - 1; i >= 0; i--) {
            var subLayer = parentLayer.layers[i];
            removed += removeEmptySublayers(subLayer);
            if (PROTECTED_LAYER_NAMES[subLayer.name]) {
                continue;
            }
            if (subLayer.pageItems.length === 0 && subLayer.layers.length === 0) {
                removed += removeLayerUnlocked(subLayer);
            }
        }
        return removed;
    }

    /* ロックを一時解除してレイヤーを削除し、削除できた数（0 または 1）を返す。失敗した場合はロック状態を元に戻す
       Unlock a layer, remove it, and return how many were removed (0 or 1); the lock is restored when removal fails */
    function removeLayerUnlocked(layer) {
        var wasLocked = false;
        try {
            wasLocked = layer.locked;
            layer.locked = false;
            layer.remove();
            return 1;
        } catch (e) {
            /* 削除できなかったのでロックを戻す（意図しないロック解除を残さない）/ Removal failed, so restore the lock instead of leaving it cleared */
            if (wasLocked) {
                try {
                    layer.locked = true;
                } catch (e2) {
                    /* 復元できない場合は無視 / Ignore when it can't be restored */
                }
            }
            return 0;
        }
    }

    // ==================================================
    // 関数：未使用スウォッチ削除 / Delete unused swatches
    // ==================================================

    /* 通常はアクションで未使用のみ、強制時は保護対象以外を全削除。件数を返す / Normally prune unused via action; force mode removes all but protected ones. Returns the count */
    function deleteUnusedSwatches(doc, force) {
        var removedCount;
        if (!force) {
            removedCount = pruneUnusedViaAction(doc.swatches, PRUNE_SPECS.swatch, "swatches");
            removeEmptySwatchGroups(doc);
            return removedCount;
        }

        /* 強制時：削除してはいけない既定スウォッチ以外を総当たり削除 / Force: remove everything except the built-in swatches */
        var protectedNames = {
            "[None]": true,
            "[Registration]": true,
            "[Black]": true,
            "[White]": true
        };

        removedCount = 0;
        for (var i = doc.swatches.length - 1; i >= 0; i--) {
            var swatch = doc.swatches[i];
            if (protectedNames[swatch.name]) {
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

    /* スウォッチを消したあとに中身が空になったスウォッチグループを片付ける。スウォッチそのものではないので件数には数えない
       Clear out swatch groups left empty after swatches were removed; they aren't swatches, so they don't count toward the total */
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

    // ==================================================
    // 関数：未使用グラフィックスタイル削除 / Delete unused graphic styles
    // ==================================================

    /* 通常はアクションで未使用のみ、強制時は既定（最後の1つ）以外を全削除。件数を返す / Normally prune unused via action; force mode removes all but the default. Returns the count */
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

    // ==================================================
    // 関数：未使用シンボル削除 / Delete unused symbols
    // ==================================================

    /* 通常はアクションで未使用のみ、強制時はすべて削除。件数を返す / Normally prune unused via action; force mode removes them all. Returns the count */
    function deleteUnusedSymbols(doc, force) {
        if (!force) {
            return pruneUnusedViaAction(doc.symbols, PRUNE_SPECS.symbol, "symbols");
        }

        var removedCount = 0;
        for (var i = doc.symbols.length - 1; i >= 0; i--) {
            try {
                doc.symbols[i].remove();
                removedCount++;
            } catch (e) {
                /* 使用中、または削除不可 / In use or not removable */
            }
        }
        return removedCount;
    }

    // ==================================================
    // 関数：未使用ブラシ削除 / Delete unused brushes
    // ==================================================

    /* 通常はアクションで未使用のみ、強制時は削除できるものをすべて削除（使用中・基本ブラシは不可）。件数を返す / Normally prune unused via action; force mode removes every removable brush (in-use and basic brushes can't be removed). Returns the count */
    function deleteUnusedBrushes(doc, force) {
        if (!force) {
            return pruneUnusedViaAction(doc.brushes, PRUNE_SPECS.brush, "brushes");
        }

        var removedCount = 0;
        for (var i = doc.brushes.length - 1; i >= 0; i--) {
            try {
                doc.brushes[i].remove();
                removedCount++;
            } catch (e) {
                /* 使用中・基本ブラシなど削除不可 / In use, basic brush, or otherwise not removable */
            }
        }
        return removedCount;
    }

    // ==================================================
    // 関数：未使用段落スタイル削除 / Delete unused paragraph styles
    // ==================================================

    /* 使用情報を取得できないため強制時のみ、既定（先頭）以外を削除。件数を返す / No usage info, so only force mode removes all but the default (first). Returns the count */
    function deleteUnusedParagraphStyles(doc, force) {
        if (!force) {
            return 0;
        }

        var removedCount = 0;
        /* インデックス0は [標準段落スタイル] なので残す / Index 0 is [Normal Paragraph Style], keep it */
        for (var i = doc.paragraphStyles.length - 1; i >= 1; i--) {
            try {
                doc.paragraphStyles[i].remove();
                removedCount++;
            } catch (e) {
                /* 削除不可 / Not removable */
            }
        }
        return removedCount;
    }

    // ==================================================
    // 関数：未使用文字スタイル削除 / Delete unused character styles
    // ==================================================

    /* 使用情報を取得できないため強制時のみ、既定（先頭）以外を削除。件数を返す / No usage info, so only force mode removes all but the default (first). Returns the count */
    function deleteUnusedCharacterStyles(doc, force) {
        if (!force) {
            return 0;
        }

        var removedCount = 0;
        /* インデックス0は [標準文字スタイル] なので残す / Index 0 is [Normal Character Style], keep it */
        for (var i = doc.characterStyles.length - 1; i >= 1; i--) {
            try {
                doc.characterStyles[i].remove();
                removedCount++;
            } catch (e) {
                /* 削除不可 / Not removable */
            }
        }
        return removedCount;
    }

    // ==================================================
    // 関数：未使用アートボード削除 / Delete unused artboards
    // ==================================================

    /* 空のアートボードを削除し、最低1つは残す。件数を返す / Remove empty artboards keeping at least one, return the count */
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
            var artboard = doc.artboards[i];
            if (!isArtboardEmpty(artworkBounds, artboard.artboardRect)) {
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