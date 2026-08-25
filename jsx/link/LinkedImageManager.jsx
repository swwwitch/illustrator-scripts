#target illustrator
#targetengine "LinkedImageManager"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

ドキュメント内の配置画像（リンク画像・埋め込み画像）を解析して一覧表示し、ソート・絞り込み・再リンク・リネーム・削除・埋め込み／解除までを一元化する常駐パレットです。
一覧とカンバスの選択は相互に連動します。

詳細は README を参照してください。

### Overview

A persistent palette that lists every placed image in the document — linked and embedded — and centralizes sorting, filtering, relinking, renaming, deleting, and embedding or unembedding.
The list and the canvas selection stay in sync with each other.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "LinkedImageManager";           /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.5.3";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-04-24";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-08-02";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/LinkedImageManager.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/LinkedImageManager.md
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/na66732d2056a"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // ユーザー設定 / User configuration
    // =========================================

    /* Dropboxのローカルマウントパス。空文字にするとホーム直下から自動検出 / Local Dropbox mount path ("" = auto detect) */
    var DROPBOX_PREFIX = resolveDropboxPrefix("");

    // 「ファイル名不明」を worker → palette 間で運ぶためのセンチネル（JSON では ￾ がエスケープされ往復する）
    var UNKNOWN_NAME = "￾UNKNOWN";

    // =========================================
    // ローカライズ / Localization
    // =========================================

    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var currentLanguage = getCurrentLang();

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {

        dialog: {
            main: { ja: "リンク画像の管理", en: "Linked Image Manager" },
            changeExt: { ja: "拡張子の変更", en: "Change Extension" },
            clipGroupDelete: { ja: "クリップグループ内の画像", en: "Image inside a clip group" }
        },

        panel: {
            sort: { ja: "ソート", en: "Sort" },
            sameFile: { ja: "同一ファイル", en: "Same Files" },
            status: { ja: "ステータス", en: "Status" },
            artboard: { ja: "アートボード", en: "Artboard" },
            displayColumn: { ja: "表示列", en: "Display Options" },
            path: { ja: "ファイルパス", en: "File Path" },
            extension: { ja: "拡張子", en: "Extension" },
            destFolder: { ja: "変更先のフォルダー", en: "Destination Folder" }
        },

        sort: {
            by: { ja: "並び順", en: "Sort by" },
            fileName: { ja: "ファイル名", en: "File Name" },
            fileSize: { ja: "サイズ", en: "Size" },
            fileCount: { ja: "使用数", en: "Usage Count" },
            artboard: { ja: "アートボード", en: "Artboard" },
            width: { ja: "幅", en: "Width" },
            height: { ja: "高さ", en: "Height" },
            scale: { ja: "スケール", en: "Scale" },
            ppi: { ja: "PPI", en: "PPI" },
            status: { ja: "ステータス", en: "Status" },
            colorSpace: { ja: "カラースペース", en: "Color Space" },
            asc: { ja: "昇順", en: "Ascending" },
            desc: { ja: "降順", en: "Descending" }
        },

        column: {
            fileName: { ja: "ファイル名", en: "File Name" },
            fileSizeMb: { ja: "サイズ(MB)", en: "Size (MB)" },
            fileSize: { ja: "サイズ", en: "Size" },
            fileCount: { ja: "使用数", en: "Usage Count" },
            widthMm: { ja: "幅(mm)", en: "Width (mm)" },
            heightMm: { ja: "高さ(mm)", en: "Height (mm)" },
            scale: { ja: "スケール", en: "Scale" },
            ppi: { ja: "PPI", en: "PPI" },
            artboards: { ja: "アートボード", en: "Artboards" },
            colorSpace: { ja: "カラースペース", en: "Color Space" }
        },

        checkbox: {
            dedup: { ja: "同一ファイルをまとめる", en: "Group Same Files" },
            unit: { ja: "単位表示「MB」", en: "Use MB" },
            displaySize: { ja: "サイズ", en: "Size" },
            displayFileCount: { ja: "使用数を表示", en: "Show Usage Count" },
            displayDimScalePpi: { ja: "サイズ、%、PPI", en: "Dimensions, Scale, PPI" },
            displayColorSpace: { ja: "カラースペース", en: "Color Space" },
            fullPath: { ja: "フルパス", en: "Full path" },
            dropbox: { ja: "Dropboxパスを短縮", en: "Shorten Dropbox Path" },
            fileName: { ja: "ファイル名", en: "Name" },
            showOnCanvas: { ja: "選択時にズーム表示", en: "Zoom to Selection" },
            filterOk: { ja: "✓ リンク正常", en: "✓ Link OK" },
            filterBroken: { ja: "⚠ リンク切れ", en: "⚠ Broken Link" },
            filterUpdate: { ja: "⟳ 更新が必要", en: "⟳ Needs Update" },
            filterEmbedded: { ja: "▣ 埋め込み", en: "▣ Embedded" },
            collectAfterRelink: { ja: "再リンク後に収集", en: "Collect after relinking" }
        },

        /* 埋め込み解除の失敗理由（worker はコードのみ返す）*/
        reason: {
            LOCKED_LAYER: { ja: "レイヤーがロックされています", en: "The layer is locked" },
            HIDDEN_LAYER: { ja: "レイヤーが非表示です", en: "The layer is hidden" },
            LOCKED_ITEM: { ja: "画像または親グループがロックされています", en: "The image or its parent group is locked" },
            HIDDEN_ITEM: { ja: "画像または親グループが非表示です", en: "The image or its parent group is hidden" },
            SELECT_FAILED: { ja: "対象の画像を選択できませんでした", en: "Could not select the target image" },
            ACTION_RESULT_MISSING: { ja: "置換結果を取得できませんでした", en: "Could not get the replaced image" },
            ACTION_FILE_FAILED: { ja: "一時アクションファイルを作成できませんでした", en: "Could not create the temporary action file" },
            UNSUPPORTED_COLORSPACE: { ja: "未対応のカラースペース", en: "Unsupported color space" },
            SCALE_UNAVAILABLE: { ja: "拡大率を取得できませんでした", en: "Could not get the scale" },
            EXPORT_FAILED: { ja: "PSDを書き出せませんでした", en: "Could not export the PSD" },
            DOC_NOT_SAVED: { ja: "ドキュメントが保存されていません", en: "The document has not been saved" },
            LINKS_FOLDER_FAILED: { ja: "「Links」フォルダーを作成できませんでした", en: "Could not create the \"Links\" folder" },
            COPY_FAILED: { ja: "リンクファイルを複製できませんでした", en: "Could not copy the linked file" }
        },

        button: {
            rename: { ja: "リネーム", en: "Rename" },
            open: { ja: "開く", en: "Open" },
            "delete": { ja: "削除", en: "Delete" },
            deleteImageOnly: { ja: "画像のみを削除", en: "Delete image only" },
            deleteWithClipGroup: { ja: "クリップグループごと削除", en: "Delete entire clip group" },
            copyFileName: { ja: "ファイル名をコピー", en: "Copy File Name" },
            openFolder: { ja: "開く", en: "Open" },
            relinkSelected: { ja: "再リンク", en: "Relink Selected" },
            relinkAll: { ja: "一括再リンク", en: "Relink All" },
            relinkFolder: { ja: "フォルダー再リンク", en: "Relink Folder" },
            embed: { ja: "埋め込み", en: "Embed" },
            unembed: { ja: "埋め込み解除", en: "Unembed" },
            changeExtension: { ja: "拡張子の変更", en: "Change Extension" },
            chooseFolder: { ja: "フォルダー指定", en: "Choose Folder" },
            collectLinks: { ja: "リンクを収集", en: "Collect Links" },
            openLinksPanel: { ja: "［リンク］パネルを開く", en: "Open Links Panel" },
            reload: { ja: "更新", en: "Reload" },
            close: { ja: "閉じる", en: "Close" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },

        message: {
            noDocument: { ja: "ドキュメントが開いていません。", en: "No document is open." },
            noPlacedItems: {
                ja: "配置されているリンク画像が見つかりませんでした。",
                en: "No linked placed images were found."
            },
            selectItem: { ja: "リストからアイテムを選択してください。", en: "Please select an item from the list." },
            noValidPath: { ja: "有効なファイルパスがありません。", en: "No valid file path is available." },
            openFolderFailed: { ja: "フォルダを開けませんでした：", en: "Could not open the folder: " },
            selectLinkedFolder: { ja: "リンクフォルダを選択してください。", en: "Please select a linked folder." },
            linkFileNotFound: { ja: "リンクファイルが見つかりません：", en: "Linked file not found: " },
            clipGroupDelete: {
                ja: "選択した画像はクリップグループ内にあります。どのように削除しますか？",
                en: "The selected image is inside a clip group. How would you like to delete it?"
            },
            confirmDeleteLinks: {
                ja: "選択したリンクファイルをドキュメントからすべて削除します。",
                en: "Remove the selected file(s) from the document?"
            },
            deleteDone: { ja: "削除完了", en: "Delete Complete" },
            copyFileNameDone: { ja: "ファイル名をコピーしました", en: "File name copied to clipboard" },
            copyFileNameFailed: { ja: "ファイル名のコピーに失敗しました", en: "Failed to copy file name" },
            promptNewFileName: {
                ja: "新しいファイル名を入力してください。\n（拡張子 {ext} は自動で保持されます）\n\n現在のファイル名：{name}",
                en: "Enter the new file name.\n(Extension {ext} will be kept automatically)\n\nCurrent file name: {name}"
            },
            invalidFileName: { ja: "ファイル名に / や \\ は使用できません。", en: "File name must not contain / or \\." },
            confirmOverwrite: {
                ja: "同名のファイルが既に存在します。\n上書きしますか？\n\n",
                en: "A file with the same name already exists.\nOverwrite?\n\n"
            },
            renameFailed: { ja: "ファイルのリネームに失敗しました。", en: "Failed to rename the file." },
            renameDone: { ja: "リネームして再リンクしました", en: "Renamed and relinked" },
            nameUnchanged: {
                ja: "ファイル名が変更されていません。処理を中断します。",
                en: "The file name has not changed. Aborting."
            },
            confirmBatchRelink: {
                ja: "複数のリンクを一括で再リンクします。よろしいですか？",
                en: "Multiple links will be relinked at once. Continue?"
            },
            relinkDone: { ja: "リンク更新完了", en: "Relink Complete" },
            confirmBatchEmbed: {
                ja: "選択したリンクの配置をすべて埋め込み画像に変換します。よろしいですか？",
                en: "All placements of the selected link will be embedded. Continue?"
            },
            embedDone: { ja: "埋め込み完了", en: "Embed Complete" },
            unembedDone: { ja: "埋め込み解除完了", en: "Unembed Complete" },
            unembedFailedDetail: { ja: "埋め込み解除できなかった画像があります。", en: "Some images could not be unembedded." },
            changeExtDone: { ja: "拡張子を変更しました", en: "Extension changed" },
            docNotSaved: {
                ja: "ドキュメントが保存されていません。保存してからもう一度お試しください。",
                en: "The document has not been saved. Please save the document and try again."
            },
            createLinksFolderFailed: { ja: "Linksフォルダーを作成できませんでした。", en: "Could not create the Links folder." },
            collectLinksDone: { ja: "リンクのコピーと再リンクが完了しました", en: "Copy and Relink Complete" }
        },

        label: {
            statusOk: { ja: "✓ リンク正常", en: "✓ Link OK" },
            statusBroken: { ja: "⚠ リンク切れ", en: "⚠ Broken Link" },
            statusUpdate: { ja: "⟳ 更新が必要", en: "⟳ Needs Update" },
            statusEmbedded: { ja: "▣ 埋め込み", en: "▣ Embedded" },
            artboardAll: { ja: "すべて", en: "All" },
            artboardFallback: { ja: "アートボード", en: "Artboard" },
            prevArtboardTip: { ja: "前のアートボード", en: "Previous Artboard" },
            nextArtboardTip: { ja: "次のアートボード", en: "Next Artboard" },
            pathPlaceholder: {
                ja: "リストからアイテムを選択してください",
                en: "Select an item from the list."
            },
            pathHelpTip: { ja: "パスの表示", en: "File path" },
            items: { ja: "件", en: "item(s)" },
            target: { ja: "対象", en: "Target" },
            success: { ja: "成功", en: "Succeeded" },
            failed: { ja: "失敗", en: "Failed" },
            skipped: { ja: "スキップ", en: "Skipped" },
            copied: { ja: "コピー", en: "Copied" },
            noExt: { ja: "(なし)", en: "(none)" },
            fileNameUnknown: { ja: "(ファイル名不明)", en: "(Unknown File Name)" },
            selectNewLinkFile: { ja: "新しいリンクファイルを選択", en: "Select a new linked file" },
            selectAltFolder: { ja: "代替フォルダを選択", en: "Select a replacement folder" },
            linkedFolders: { ja: "リンクフォルダー一覧", en: "Linked Folders" },
            selectExtensionReferenceFolder: {
                ja: "拡張子変更で参照するフォルダーを選択してください",
                en: "Select the folder to search for files with the new extension"
            },
            extensionReferenceFolderPlaceholder: { ja: "参照フォルダー未指定", en: "No reference folder selected" }
        },

        status: {
            ready: { ja: "準備完了", en: "Ready" },
            loaded: { ja: "読み込み完了", en: "Loaded" },
            loadFailed: { ja: "読み込みに失敗しました", en: "Failed to load" },
            noDocument: { ja: "ドキュメントが開いていません", en: "No document is open" },
            noPlaced: { ja: "配置されているリンク画像がありません", en: "No linked images found" },
            busy: { ja: "処理中です…", en: "Working…" },
            notWired: {
                ja: "この操作は次のステップで有効化されます（現在は準備中）",
                en: "This action will be enabled in a later step."
            }
        }
    };

    // ドット区切りキー（例 'panel.sort'）で LABELS を辿る。見つからなければキー文字列を返す
    function L(key) {
        var node = LABELS;
        var parts = key.split(".");
        for (var i = 0; i < parts.length; i++) {
            if (node && typeof node === "object" && node.hasOwnProperty(parts[i])) {
                node = node[parts[i]];
            } else {
                return key;
            }
        }
        return (node && node[currentLanguage]) ? node[currentLanguage] : key;
    }

    function labelText(key) {
        return L(key) + (currentLanguage === 'ja' ? '：' : ':');
    }

    // 数値＋単位をローカライズ付きで整形
    function withUnit(value, unitKey) {
        return (currentLanguage === 'ja')
            ? (value + L(unitKey))
            : (value + " " + L(unitKey));
    }

    // alert / status 用に「ラベル：値+単位」を 1 行で組み立てる
    function kvLine(labelKey, value, unitKey) {
        var sep = (currentLanguage === 'ja' ? '：' : ': ');
        var valueText = unitKey ? withUnit(value, unitKey) : String(value);
        return L(labelKey) + sep + valueText;
    }

    // ステータスコードから表示用ラベル・アイコン・正常判定を返す（worker はコードのみ返す）
    function statusDisplay(statusCode) {
        if (statusCode === "broken") return { status: L('label.statusBroken'), statusIcon: "⚠", isLinkOk: false };
        if (statusCode === "update") return { status: L('label.statusUpdate'), statusIcon: "⟳", isLinkOk: false };
        if (statusCode === "embedded") return { status: L('label.statusEmbedded'), statusIcon: "▣", isLinkOk: true };
        return { status: L('label.statusOk'), statusIcon: "✓", isLinkOk: true };
    }

    // =========================================
    // UIレイアウトの共通設定 / Shared UI layout
    // =========================================

    /* ウィンドウ・パネルの余白と間隔 / Window & panel margins and spacing */
    var WINDOW_MARGINS = 16;                 /* ウィンドウ外周の余白 / window margin */
    var WINDOW_SPACING = 12;                 /* ウィンドウ内の要素間隔 / window spacing */
    var PANEL_MARGINS  = [16, 20, 16, 12];   /* パネル余白 [左,上,右,下] / panel margins */
    var PANEL_SPACING  = 12;                 /* パネル内の要素間隔 / panel spacing */
    var COLUMN_SPACING = 12;                 /* 2カラムの間隔 / gap between columns */

    /* ウィンドウの共通設定 / Apply shared window layout */
    function setupWindow(win, spacing) {
        win.orientation = "column";
        win.alignChildren = ["fill", "top"];
        win.margins = WINDOW_MARGINS;
        win.spacing = (typeof spacing === "number") ? spacing : WINDOW_SPACING;
    }

    /* パネルの共通設定 / Apply shared panel layout */
    function setupPanel(panel, spacing) {
        panel.orientation = "column";
        panel.alignChildren = ["fill", "top"];
        panel.alignment = "fill";
        panel.margins = PANEL_MARGINS;
        panel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /* 行グループの共通設定（ボタン列など） / Apply a horizontal row group */
    function setupRow(group, alignment, spacing) {
        group.orientation = "row";
        group.alignment = alignment || "left";
        group.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    function setupGroup(group, orientation, spacing) {
        group.orientation = orientation || "column";
        group.alignChildren = (group.orientation === "row") ? ["left", "center"] : ["fill", "top"];
        group.alignment = "fill";
        group.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    // =========================================
    // パレット側 純粋ヘルパー / Palette-side pure helpers（DOM を触らない）
    // =========================================

    function safeRelayout() {
        for (var i = 0; i < arguments.length; i++) {
            var comp = arguments[i];
            if (!comp) continue;
            ignoreError(function () { comp.layout.layout(true); });
        }
    }

    // UI の明暗を判定（uiBrightness > 0.5 で明。取得失敗は暗い側にフォールバック）
    function isLightUI() {
        var b = tryGet(function () { return app.preferences.getRealPreference("uiBrightness"); }, 0);
        return b > 0.5;
    }

    /* ◀ ▶ ボタンの描画色。三角（mark）より枠（frame）を薄くして主張を抑える */
    var ARROW_COLORS = {
        light:         { mark: [0.40, 0.40, 0.40, 1], frame: [0.60, 0.60, 0.60, 1] }, /* #666 / #999 */
        lightDisabled: { mark: [0.72, 0.72, 0.72, 1], frame: [0.82, 0.82, 0.82, 1] },
        dark:          { mark: [0.65, 0.65, 0.65, 1], frame: [0.45, 0.45, 0.45, 1] },
        darkDisabled:  { mark: [0.38, 0.38, 0.38, 1], frame: [0.30, 0.30, 0.30, 1] }
    };

    // ボタン面に三角形（◀ / ▶）を onDraw で自前描画する。
    // フォントグリフに依存せず、枠＋塗り三角をベクターで描く。direction: -1=左, +1=右
    function attachArrowDraw(btn, direction) {
        btn.text = "";
        btn.onDraw = function () {
            var g = this.graphics;
            var w = this.size[0], h = this.size[1];
            if (!w || !h) return;
            var scheme = isLightUI()
                ? (this.enabled ? ARROW_COLORS.light : ARROW_COLORS.lightDisabled)
                : (this.enabled ? ARROW_COLORS.dark : ARROW_COLORS.darkDisabled);
            var pen = g.newPen(g.PenType.SOLID_COLOR, scheme.frame, 1);
            var brush = g.newBrush(g.BrushType.SOLID_COLOR, scheme.mark);

            // 枠
            g.newPath();
            g.moveTo(0.5, 0.5);
            g.lineTo(w - 0.5, 0.5);
            g.lineTo(w - 0.5, h - 0.5);
            g.lineTo(0.5, h - 0.5);
            g.closePath();
            g.strokePath(pen);

            // 三角
            var cx = w / 2, cy = h / 2;
            var half = Math.min(w, h) * 0.26;
            g.newPath();
            if (direction < 0) {
                g.moveTo(cx - half, cy);
                g.lineTo(cx + half, cy - half);
                g.lineTo(cx + half, cy + half);
            } else {
                g.moveTo(cx + half, cy);
                g.lineTo(cx - half, cy - half);
                g.lineTo(cx - half, cy + half);
            }
            g.closePath();
            g.fillPath(brush);
        };
    }

    function tryGet(fn, fallback) {
        try {
            return fn();
        } catch (e) {
            return fallback;
        }
    }

    function ignoreError(fn) {
        try { fn(); } catch (e) { }
    }

    function lastPathSep(path) {
        return Math.max(path.lastIndexOf("/"), path.lastIndexOf("\\"));
    }

    function pathParent(path) {
        if (!path || path === "---") return "";
        var sep = lastPathSep(path);
        return (sep > 0) ? path.substring(0, sep) : "";
    }

    function pathBaseName(path) {
        if (!path || path === "---") return "";
        var sep = lastPathSep(path);
        return (sep >= 0) ? path.substring(sep + 1) : path;
    }

    function toFolderOnly(path) {
        if (!path || path === "---") return path;
        var sep = lastPathSep(path);
        if (sep < 0) return path;
        return path.substring(0, sep + 1);
    }

    /**
     * ホーム直下から「Dropbox」を含むフォルダーを探す。
     * チームフォルダー（「sw Dropbox」など）を優先する。
     * 個人用の「Dropbox」は「~/Dropbox/」として短縮できるため、優先度を下げている。
     * @returns {Folder|null} 見つかったフォルダー。なければnull
     */
    function findDropboxFolder() {
        var homeFolder = Folder("~");
        if (!homeFolder.exists) return null;

        var entryList = tryGet(function () { return homeFolder.getFiles(); }, []);
        var personalFolder = null;
        var teamFolder     = null;

        for (var i = 0; i < entryList.length; i++) {
            var entry = entryList[i];
            if (!(entry instanceof Folder)) continue;

            var entryName = tryGet(function () { return decodeURI(String(entry.name)); }, String(entry.name));
            if (entryName.charAt(0) === ".") continue;
            if (entryName.indexOf("Dropbox") === -1) continue;

            if (entryName === "Dropbox") {
                if (!personalFolder) personalFolder = entry;
            } else if (!teamFolder) {
                teamFolder = entry;
            }
        }
        return teamFolder ? teamFolder : personalFolder;
    }

    /**
     * フォルダー直下に表示用のサブフォルダーが1つだけあるとき、そのフォルダーを返す。
     * チームDropboxのメンバーフォルダー（「takano masahiro」など）の判定に使う。
     * @param {Folder} parentFolder - 探索するフォルダー
     * @returns {Folder|null} 唯一のサブフォルダー。0個または2個以上のときはnull
     */
    function findSingleSubFolder(parentFolder) {
        var entryList = tryGet(function () { return parentFolder.getFiles(); }, []);
        var foundFolder = null;

        for (var i = 0; i < entryList.length; i++) {
            var entry = entryList[i];
            if (!(entry instanceof Folder)) continue;

            var entryName = tryGet(function () { return decodeURI(String(entry.name)); }, String(entry.name));
            if (entryName.charAt(0) === ".") continue;

            if (foundFolder) return null;
            foundFolder = entry;
        }
        return foundFolder;
    }

    /**
     * Dropboxのローカルマウントパスを決める。
     * 手動指定が空のときは、ホーム直下の「Dropbox」を含むフォルダーを探し、
     * その中にメンバーフォルダーが1つだけあれば、そこまでをプレフィックスとする。
     * @param {string} manualPath - 手動で指定するパス。空文字なら自動検出
     * @returns {string} 末尾に「/」を付けたプレフィックス。見つからない場合は空文字
     */
    function resolveDropboxPrefix(manualPath) {
        if (manualPath) {
            return (manualPath.charAt(manualPath.length - 1) === "/") ? manualPath : manualPath + "/";
        }

        var dropboxFolder = findDropboxFolder();
        if (!dropboxFolder) return "";

        var memberFolder = findSingleSubFolder(dropboxFolder);
        return (memberFolder ? memberFolder : dropboxFolder).fsName + "/";
    }

    function toTildePath(path) {
        if (!path || path === "---") return path;
        var home = tryGet(function () { return Folder("~").fsName; }, "");
        if (home && home.length > 0) {
            if (path === home) return "~";
            if (path.indexOf(home + "/") === 0) return "~" + path.substring(home.length);
            if (path.indexOf(home + "\\") === 0) return "~" + path.substring(home.length);
        }
        return path;
    }

    function formatDisplayPath(absPath, useTilde, useDropbox) {
        if (!absPath || absPath === "---") return absPath;
        if (useDropbox && DROPBOX_PREFIX && absPath.indexOf(DROPBOX_PREFIX) === 0) {
            return absPath.substring(DROPBOX_PREFIX.length);
        }
        if (useTilde) {
            return toTildePath(absPath);
        }
        return absPath;
    }

    function splitFileName(name) {
        var dotIdx = name.lastIndexOf(".");
        if (dotIdx <= 0) {
            return { base: name, ext: "" };
        }
        return {
            base: name.substring(0, dotIdx),
            ext: name.substring(dotIdx)
        };
    }

    function getRealFileName(file) {
        return file.fsName.split(/[\\\/]/).pop();
    }

    function promptNewFileName(originalName, message) {
        var parts = splitFileName(originalName);
        var ext = parts.ext;
        var base = parts.base;

        var msg = message || L('message.promptNewFileName')
            .replace("{ext}", ext || L('label.noExt'))
            .replace("{name}", originalName);
        var input = prompt(msg, base);
        if (input === null) return null;
        input = input.replace(/^\s+|\s+$/g, "");
        if (input === "") return null;

        if (ext && input.length > ext.length &&
            input.substr(input.length - ext.length).toLowerCase() === ext.toLowerCase()) {
            input = input.slice(0, -ext.length);
        }

        return input + ext;
    }

    function formatFileSize(bytes) {
        if (bytes === null || bytes === undefined || isNaN(bytes) || bytes < 0) return "-";
        return (bytes / (1024 * 1024)).toFixed(2);
    }

    function formatFileSizeAuto(bytes) {
        if (bytes === null || bytes === undefined || isNaN(bytes) || bytes < 0) return "-";
        if (bytes < 1024) return bytes + "B";
        if (bytes < 1024 * 1024) return Math.round(bytes / 1024) + "KB";
        if (bytes < 1024 * 1024 * 1024) return (bytes / (1024 * 1024)).toFixed(1) + "MB";
        return (bytes / (1024 * 1024 * 1024)).toFixed(2) + "GB";
    }

    function normalizeFolderPathForCompare(path) {
        if (!path) return "";
        path = String(path).replace(/\\/g, "/");
        path = path.replace(/\/+$/g, "");
        return path;
    }

    // File 操作（存在確認・rename・copy・getFiles）は常駐エンジンでも動くためパレット側で完結できる。
    // DOM を触るのは placedItem.file への代入のみで、そこだけ worker（relinkPairs）へ委譲する。

    // 上書き先がいる場合の処理（確認＋削除）。true=rename 続行可、false=中止
    function prepareRenameOverwrite(oldFile, newFile) {
        if (!newFile.exists) return true;
        if (newFile.fsName.toLowerCase() === oldFile.fsName.toLowerCase()) return true;
        if (!confirm(L('message.confirmOverwrite') + newFile.fsName)) return false;
        var removed = tryGet(function () { return newFile.remove(); }, false);
        if (!removed) {
            setStatus(L('message.renameFailed'));
            return false;
        }
        return true;
    }

    // 参照フォルダー内で baseName＋指定拡張子（primary→fallback）のファイルを大小文字無視で探す
    function findReplacementFileByExtension(folder, baseName, primaryExt, fallbackExt) {
        if (!folder || !baseName || !primaryExt) return null;
        var candidateExts = [String(primaryExt).toLowerCase()];
        if (fallbackExt) candidateExts.push(String(fallbackExt).toLowerCase());
        var baseLower = String(baseName).toLowerCase();
        var files;
        try {
            files = folder.getFiles(function (f) { return !(f instanceof Folder); });
        } catch (e) {
            return null;
        }
        if (!files) return null;
        for (var extIdx = 0; extIdx < candidateExts.length; extIdx++) {
            var targetExt = candidateExts[extIdx];
            for (var i = 0; i < files.length; i++) {
                var parts = splitFileName(decodeURI(files[i].name));
                if (!parts.ext) continue;
                if (parts.base.toLowerCase() === baseLower && parts.ext.toLowerCase() === targetExt) {
                    return files[i];
                }
            }
        }
        return null;
    }

    // =========================================
    // BridgeTalk 委譲インフラ / Delegation infrastructure
    // =========================================

    // 再入防止：委譲中に別の委譲を走らせない
    var isBusy = false;
    // worker 常駐済みフラグ（メインエンジンに __LIM を一度だけ送る）
    var workerPersisted = false;

    // 文字列を worker 呼び出し式に安全に埋め込むための JS 文字列リテラル化
    function jsString(s) {
        if (s === null || s === undefined) s = "";
        s = String(s);
        var out = '"';
        for (var i = 0; i < s.length; i++) {
            var c = s.charAt(i);
            var code = s.charCodeAt(i);
            if (c === '"') out += '\\"';
            else if (c === '\\') out += '\\\\';
            else if (c === '\n') out += '\\n';
            else if (c === '\r') out += '\\r';
            else if (c === '\t') out += '\\t';
            else if (code < 32 || code > 126) {
                var h = code.toString(16);
                while (h.length < 4) h = "0" + h;
                out += '\\u' + h;
            } else out += c;
        }
        return out + '"';
    }

    // 整数配列を JS 配列リテラルに
    function jsIntArray(arr) {
        var parts = [];
        for (var i = 0; i < arr.length; i++) parts.push(String(parseInt(arr[i], 10)));
        return "[" + parts.join(",") + "]";
    }

    // worker 関数群を 1 つの IIFE にまとめてメインエンジンへ送る本文を組み立てる。
    // 全ヘルパーを IIFE クロージャ内に置き、$.global.__LIM にエントリだけ公開することで、
    // グローバル関数宣言の永続化に依存せず確実に常駐させる。
    // toString() は関数本体の前後に周辺コメントの断片（閉じ「*/」を欠いたもの）を含めて返すことがある。
    // そのままつなぐと未終端コメントが次の関数を飲み込むため、宣言行から閉じ括弧行までを行単位で抜き出す。
    function sliceFunctionSource(rawSource) {
        var lines = rawSource.replace(/\r\n?/g, "\n").split("\n");
        var first = 0;
        while (first < lines.length && lines[first].indexOf("function ") !== 0) first++;
        if (first >= lines.length) return rawSource; /* 想定外の形式：そのまま返す */
        var last = lines.length - 1;
        while (last > first && !/^\s*\}\s*$/.test(lines[last])) last--;
        return lines.slice(first, last + 1).join("\n");
    }

    function buildWorkerBundle() {
        var src = "$.global.__LIM = (function () {\n";
        for (var i = 0; i < WORKER_FUNCS.length; i++) {
            src += sliceFunctionSource(WORKER_FUNCS[i].toString()) + "\n";
        }
        src += "return { analyze: w_analyze, currentIndex: w_currentIndex, select: w_select, fitArtboard: w_fitArtboard, openLinksPanel: w_openLinksPanel, relinkPairs: w_relinkPairs, embed: w_embed, unembed: w_unembed, probeDelete: w_probeDelete, del: w_del, docFolder: w_docFolder, copyText: w_copyText };\n";
        src += "})();";
        return src;
    }

    // メインエンジンへ 1 文を同期送信。結果 body を返す。エラー／タイムアウトは null
    // timeoutSec 省略時は 10 秒（PSD 書き出しなど長い処理は呼び出し側で延ばす）
    // 直近の送信失敗理由（エラー本文／タイムアウト）。失敗時のステータス表示に使う
    var lastBridgeError = "";

    function sendRaw(bodyExpr, timeoutSec) {
        var holder = {};
        var bridge = new BridgeTalk();
        bridge.target = "illustrator";
        bridge.body = "eval(decodeURIComponent(\"" + encodeURIComponent(bodyExpr) + "\"))";
        bridge.onResult = function (msg) { holder.value = msg.body; };
        bridge.onError = function (msg) { holder.error = msg.body; };
        lastBridgeError = "";
        bridge.send(timeoutSec || 10);
        if (holder.hasOwnProperty("error")) {
            lastBridgeError = String(holder.error);
            return null;
        }
        if (!holder.hasOwnProperty("value")) {
            lastBridgeError = "timeout";
            return null;
        }
        return holder.value;
    }

    // worker が未常駐なら常駐させる
    function ensureWorker() {
        if (workerPersisted) return true;
        var r = sendRaw(buildWorkerBundle());
        if (r === null) return false;
        workerPersisted = true;
        return true;
    }

    // 委譲呼び出し：常駐確認 → 呼び出し式送信 → 失敗時は再常駐して 1 度だけリトライ
    function delegate(callExpr) {
        if (!ensureWorker()) return null;
        var res = sendRaw(callExpr);
        if (res === null) {
            workerPersisted = false;
            if (!ensureWorker()) return null;
            res = sendRaw(callExpr);
        }
        return res;
    }

    // 委譲呼び出し（リトライなし）。書き込みを伴う長時間処理は再送すると二重実行になるためこちらを使う
    function delegateOnce(callExpr, timeoutSec) {
        if (!ensureWorker()) return null;
        return sendRaw(callExpr, timeoutSec);
    }

    // "OK\n<json>" / "NODOC" / "ERR\n<msg>" を分解 (最初の \n が区切り)
    function parseWorkerResult(res) {
        if (res === null || res === undefined) return { marker: "ERR", body: lastBridgeError || "no result" };
        var sep = res.indexOf("\n");
        if (sep < 0) return { marker: res, body: "" };
        return { marker: res.substring(0, sep), body: res.substring(sep + 1) };
    }

    // =========================================
    // WORKER 関数（メインエンジンで実行）/ Worker functions (run in main engine)
    // toString で連結して送るため：// 行コメント禁止・/* */ のみ・必ずセミコロンで終える。
    // 相互参照は w_ プレフィックス名で行う（同一 IIFE クロージャ内に置かれるため確実）。
    // =========================================

    /* 値を ASCII のみの JSON 文字列に符号化（非 ASCII・制御文字は \\u エスケープ） */
    function w_jsonStr(v) {
        if (v === null || v === undefined) return "null";
        var t = typeof v;
        if (t === "number") { if (isNaN(v) || !isFinite(v)) return "null"; return String(v); }
        if (t === "boolean") { return v ? "true" : "false"; }
        if (t === "string") {
            var out = "\"";
            for (var i = 0; i < v.length; i++) {
                var c = v.charAt(i);
                var code = v.charCodeAt(i);
                if (c === "\"") { out += "\\\""; }
                else if (c === "\\") { out += "\\\\"; }
                else if (code < 32 || code > 126) {
                    var h = code.toString(16);
                    while (h.length < 4) { h = "0" + h; }
                    out += "\\u" + h;
                } else { out += c; }
            }
            return out + "\"";
        }
        if (v instanceof Array) {
            var parts = [];
            for (var j = 0; j < v.length; j++) { parts.push(w_jsonStr(v[j])); }
            return "[" + parts.join(",") + "]";
        }
        var op = [];
        for (var k in v) { if (v.hasOwnProperty(k)) { op.push(w_jsonStr(String(k)) + ":" + w_jsonStr(v[k])); } }
        return "{" + op.join(",") + "}";
    }

    function w_tryGet(fn, fallback) {
        try { return fn(); } catch (e) { return fallback; }
    }

    function w_safeExists(target) {
        return w_tryGet(function () { return !!(target && target.exists); }, false);
    }

    function w_safeProp(obj, key, fallback) {
        return w_tryGet(function () {
            var value = obj[key];
            return (value !== undefined && value !== null && value !== "") ? value : fallback;
        }, fallback);
    }

    function w_readU16BE(bytes, offset) {
        return ((bytes.charCodeAt(offset) & 0xFF) << 8) | (bytes.charCodeAt(offset + 1) & 0xFF);
    }

    function w_readU32BE(bytes, offset) {
        return ((bytes.charCodeAt(offset) & 0xFF) * 16777216) +
            ((bytes.charCodeAt(offset + 1) & 0xFF) << 16) +
            ((bytes.charCodeAt(offset + 2) & 0xFF) << 8) +
            (bytes.charCodeAt(offset + 3) & 0xFF);
    }

    /* ICC プロファイルのバイナリから desc タグ（プロファイル名）を取り出す */
    function w_readIccDesc(iccBuffer) {
        if (!iccBuffer || iccBuffer.length < 132) return "";
        return w_tryGet(function () {
            var tagCount = w_readU32BE(iccBuffer, 128);
            for (var i = 0; i < tagCount; i++) {
                var tagOffset = 132 + i * 12;
                if (tagOffset + 12 > iccBuffer.length) break;
                var tagSignature = iccBuffer.substr(tagOffset, 4);
                if (tagSignature === "desc") {
                    var dataOffset = w_readU32BE(iccBuffer, tagOffset + 4);
                    var dataSize = w_readU32BE(iccBuffer, tagOffset + 8);
                    if (dataOffset + dataSize > iccBuffer.length) return "";
                    var dataType = iccBuffer.substr(dataOffset, 4);
                    if (dataType === "desc") {
                        var asciiCount = w_readU32BE(iccBuffer, dataOffset + 8);
                        if (asciiCount > 0) {
                            var asciiString = iccBuffer.substr(dataOffset + 12, asciiCount);
                            return asciiString.replace(/\0+$/g, "").replace(/\0.*$/g, "");
                        }
                    } else if (dataType === "mluc") {
                        var recordCount = w_readU32BE(iccBuffer, dataOffset + 8);
                        if (recordCount > 0) {
                            var recordOffset = dataOffset + 16;
                            var stringLength = w_readU32BE(iccBuffer, recordOffset + 4);
                            var stringOffset = w_readU32BE(iccBuffer, recordOffset + 8);
                            if (dataOffset + stringOffset + stringLength <= iccBuffer.length && stringLength > 0) {
                                var rawBytes = iccBuffer.substr(dataOffset + stringOffset, stringLength);
                                var decoded = "";
                                for (var byteIdx = 0; byteIdx + 1 < rawBytes.length; byteIdx += 2) {
                                    var codeUnit = ((rawBytes.charCodeAt(byteIdx) & 0xFF) << 8) | (rawBytes.charCodeAt(byteIdx + 1) & 0xFF);
                                    if (codeUnit === 0) break;
                                    decoded += String.fromCharCode(codeUnit);
                                }
                                return decoded;
                            }
                        }
                    }
                    break;
                }
            }
            return "";
        }, "");
    }

    /* PNG IHDR / iCCP / sRGB から幅高さ・カラーモード・ICC 名 */
    function w_readPngImageInfo(binaryFile) {
        var info = { width: null, height: null, colorMode: "", iccDesc: "" };
        binaryFile.seek(16);
        var ihdrBytes = binaryFile.read(10);
        if (ihdrBytes && ihdrBytes.length === 10) {
            info.width = w_readU32BE(ihdrBytes, 0);
            info.height = w_readU32BE(ihdrBytes, 4);
            var colorType = ihdrBytes.charCodeAt(9) & 0xFF;
            if (colorType === 0 || colorType === 4) { info.colorMode = "Grayscale"; }
            else if (colorType === 3) { info.colorMode = "Indexed"; }
            else { info.colorMode = "RGB"; }
        }
        binaryFile.seek(8 + 4 + 4 + 13 + 4);
        for (var chunkIdx = 0; chunkIdx < 32; chunkIdx++) {
            var pngChunkLengthBytes = binaryFile.read(4);
            if (!pngChunkLengthBytes || pngChunkLengthBytes.length < 4) break;
            var chunkLength = w_readU32BE(pngChunkLengthBytes, 0);
            var chunkType = binaryFile.read(4);
            if (!chunkType || chunkType.length < 4) break;
            if (chunkType === "iCCP") {
                var chunkData = binaryFile.read(chunkLength);
                if (chunkData) {
                    var nullIndex = chunkData.indexOf("\0");
                    if (nullIndex > 0) { info.iccDesc = chunkData.substring(0, nullIndex); }
                }
                break;
            } else if (chunkType === "sRGB") {
                info.iccDesc = "sRGB IEC61966-2.1";
                break;
            } else if (chunkType === "IDAT" || chunkType === "IEND") {
                break;
            } else {
                binaryFile.seek(binaryFile.tell() + chunkLength + 4);
            }
        }
        return info;
    }

    /* JPEG SOFn / APP2(ICC) から幅高さ・コンポーネント数・ICC */
    function w_readJpegImageInfo(binaryFile) {
        var info = { width: null, height: null, colorMode: "", iccDesc: "" };
        binaryFile.seek(2);
        var iccPieces = {};
        var iccTotalPieces = 0;
        var componentCount = 0;
        for (var markerIdx = 0; markerIdx < 64; markerIdx++) {
            var markerBytes = binaryFile.read(2);
            if (!markerBytes || markerBytes.length < 2) break;
            if ((markerBytes.charCodeAt(0) & 0xFF) !== 0xFF) break;
            var markerCode = markerBytes.charCodeAt(1) & 0xFF;
            if (markerCode === 0xD9 || markerCode === 0xDA) break;
            var segmentLengthBytes = binaryFile.read(2);
            if (!segmentLengthBytes || segmentLengthBytes.length < 2) break;
            var segmentLength = w_readU16BE(segmentLengthBytes, 0);
            var segmentDataLength = segmentLength - 2;
            if (segmentDataLength < 0) break;
            var segmentStart = binaryFile.tell();
            var isSof = (markerCode >= 0xC0 && markerCode <= 0xCF) && markerCode !== 0xC4 && markerCode !== 0xC8 && markerCode !== 0xCC;
            if (isSof) {
                var sofBytes = binaryFile.read(6);
                if (sofBytes && sofBytes.length === 6) {
                    info.height = w_readU16BE(sofBytes, 1);
                    info.width = w_readU16BE(sofBytes, 3);
                    componentCount = sofBytes.charCodeAt(5) & 0xFF;
                }
                binaryFile.seek(segmentStart + segmentDataLength);
            } else if (markerCode === 0xE2) {
                var appSegmentHeader = binaryFile.read(14);
                if (appSegmentHeader && appSegmentHeader.length === 14 && appSegmentHeader.substring(0, 12) === "ICC_PROFILE\0") {
                    var sequenceNumber = appSegmentHeader.charCodeAt(12) & 0xFF;
                    var totalPieces = appSegmentHeader.charCodeAt(13) & 0xFF;
                    iccTotalPieces = totalPieces;
                    var pieceLength = segmentDataLength - 14;
                    if (pieceLength > 0) { iccPieces[sequenceNumber] = binaryFile.read(pieceLength); }
                    else { binaryFile.seek(segmentStart + segmentDataLength); }
                } else {
                    binaryFile.seek(segmentStart + segmentDataLength);
                }
            } else {
                binaryFile.seek(segmentStart + segmentDataLength);
            }
        }
        if (componentCount === 1) { info.colorMode = "Grayscale"; }
        else if (componentCount === 4) { info.colorMode = "CMYK"; }
        else if (componentCount > 0) { info.colorMode = "RGB"; }
        if (iccTotalPieces > 0) {
            var iccData = "";
            var allPiecesReceived = true;
            for (var pieceIdx = 1; pieceIdx <= iccTotalPieces; pieceIdx++) {
                if (!iccPieces[pieceIdx]) { allPiecesReceived = false; break; }
                iccData += iccPieces[pieceIdx];
            }
            if (allPiecesReceived && iccData.length > 0) { info.iccDesc = w_readIccDesc(iccData); }
        }
        return info;
    }

    /* PSD ヘッダ / Image Resources(0x040F) から幅高さ・modeCode・ICC */
    function w_readPsdImageInfo(binaryFile) {
        var info = { width: null, height: null, colorMode: "", iccDesc: "" };
        binaryFile.seek(14);
        var psdDimsBytes = binaryFile.read(8);
        if (psdDimsBytes && psdDimsBytes.length === 8) {
            info.height = w_readU32BE(psdDimsBytes, 0);
            info.width = w_readU32BE(psdDimsBytes, 4);
        }
        var depthModeBytes = binaryFile.read(4);
        if (depthModeBytes && depthModeBytes.length === 4) {
            var modeCode = w_readU16BE(depthModeBytes, 2);
            if (modeCode === 0) { info.colorMode = "Bitmap"; }
            else if (modeCode === 1) { info.colorMode = "Grayscale"; }
            else if (modeCode === 2) { info.colorMode = "Indexed"; }
            else if (modeCode === 3) { info.colorMode = "RGB"; }
            else if (modeCode === 4) { info.colorMode = "CMYK"; }
            else if (modeCode === 7) { info.colorMode = "Multichannel"; }
            else if (modeCode === 8) { info.colorMode = "Duotone"; }
            else if (modeCode === 9) { info.colorMode = "Lab"; }
        }
        var colorModeDataLengthBytes = binaryFile.read(4);
        if (!colorModeDataLengthBytes || colorModeDataLengthBytes.length < 4) return info;
        var colorModeDataLength = w_readU32BE(colorModeDataLengthBytes, 0);
        var imageResourceStart = 30 + colorModeDataLength;
        binaryFile.seek(imageResourceStart);
        var imageResourceLengthBytes = binaryFile.read(4);
        if (!imageResourceLengthBytes || imageResourceLengthBytes.length < 4) return info;
        var imageResourceLength = w_readU32BE(imageResourceLengthBytes, 0);
        var imageResourceEnd = imageResourceStart + 4 + imageResourceLength;
        while (binaryFile.tell() < imageResourceEnd) {
            var positionBefore = binaryFile.tell();
            var resourceSigBytes = binaryFile.read(4);
            if (!resourceSigBytes || resourceSigBytes.length < 4 || resourceSigBytes !== "8BIM") break;
            var resourceIdBytes = binaryFile.read(2);
            if (!resourceIdBytes || resourceIdBytes.length < 2) break;
            var resourceId = w_readU16BE(resourceIdBytes, 0);
            var nameLengthByte = binaryFile.read(1);
            if (!nameLengthByte) break;
            var nameLength = nameLengthByte.charCodeAt(0) & 0xFF;
            var paddedNameLength = nameLength + 1;
            if (paddedNameLength % 2 !== 0) paddedNameLength++;
            binaryFile.seek(binaryFile.tell() + (paddedNameLength - 1));
            var dataSizeBytes = binaryFile.read(4);
            if (!dataSizeBytes || dataSizeBytes.length < 4) break;
            var dataSize = w_readU32BE(dataSizeBytes, 0);
            var paddedDataSize = (dataSize % 2 === 0) ? dataSize : dataSize + 1;
            if (resourceId === 0x040F) {
                var iccProfileData = binaryFile.read(dataSize);
                if (iccProfileData && iccProfileData.length > 0) { info.iccDesc = w_readIccDesc(iccProfileData); }
                break;
            } else {
                binaryFile.seek(binaryFile.tell() + paddedDataSize);
            }
            if (binaryFile.tell() <= positionBefore) break;
        }
        return info;
    }

    /* 画像ファイルを 1 度開いて寸法・カラーモード・ICC 名を取得。失敗時 null */
    function w_readImageInfo(file) {
        if (!file) return null;
        if (!w_safeExists(file)) return null;
        var binaryFile = new File(file.fsName);
        binaryFile.encoding = "BINARY";
        if (!binaryFile.open("r")) return null;
        var info = null;
        try {
            var signature = binaryFile.read(8);
            if (signature && signature.length >= 4) {
                var b0 = signature.charCodeAt(0) & 0xFF;
                var b1 = signature.charCodeAt(1) & 0xFF;
                var b2 = signature.charCodeAt(2) & 0xFF;
                var b3 = signature.charCodeAt(3) & 0xFF;
                if (b0 === 0x89 && b1 === 0x50 && b2 === 0x4E && b3 === 0x47) { info = w_readPngImageInfo(binaryFile); }
                else if (b0 === 0xFF && b1 === 0xD8) { info = w_readJpegImageInfo(binaryFile); }
                else if (b0 === 0x38 && b1 === 0x42 && b2 === 0x50 && b3 === 0x53) { info = w_readPsdImageInfo(binaryFile); }
            }
        } catch (e) {
            info = null;
        } finally {
            binaryFile.close();
        }
        return info;
    }

    function w_lastPathSep(path) {
        return Math.max(path.lastIndexOf("/"), path.lastIndexOf("\\"));
    }

    function w_pathBaseName(path) {
        if (!path || path === "---") return "";
        var sep = w_lastPathSep(path);
        return (sep >= 0) ? path.substring(sep + 1) : path;
    }

    function w_formatFileSize(bytes) {
        if (bytes === null || bytes === undefined || isNaN(bytes) || bytes < 0) return "-";
        return (bytes / (1024 * 1024)).toFixed(2);
    }

    function w_getEffectivePPI(item, pixelSize) {
        if (!pixelSize) return null;
        var placedWidthPt = w_tryGet(function () { return item.width; }, null);
        var placedHeightPt = w_tryGet(function () { return item.height; }, null);
        if (!placedWidthPt || !placedHeightPt) return null;
        var ppiX = pixelSize.width * 72 / placedWidthPt;
        var ppiY = pixelSize.height * 72 / placedHeightPt;
        return Math.round((ppiX + ppiY) / 2);
    }

    /* XMP の ISO 8601 日付文字列を Date に変換 */
    function w_parseXmpDate(dateString) {
        if (!dateString) return null;
        var m = String(dateString).match(
            /^\s*(\d{4})-(\d{2})-(\d{2})(?:[T ](\d{2}):(\d{2})(?::(\d{2}))?(?:\.\d+)?\s*(Z|[+\-]\d{2}:?\d{2})?)?/
        );
        if (!m) return null;
        var year = parseInt(m[1], 10);
        var month = parseInt(m[2], 10) - 1;
        var day = parseInt(m[3], 10);
        var hour = m[4] ? parseInt(m[4], 10) : 0;
        var minute = m[5] ? parseInt(m[5], 10) : 0;
        var second = m[6] ? parseInt(m[6], 10) : 0;
        var tz = m[7];
        var parsed;
        if (tz) {
            var offsetMinutes = 0;
            if (tz !== "Z") {
                var sign = (tz.charAt(0) === "-") ? -1 : 1;
                var tzDigits = tz.substring(1).replace(":", "");
                offsetMinutes = sign * (parseInt(tzDigits.substring(0, 2), 10) * 60 + parseInt(tzDigits.substring(2, 4), 10));
            }
            parsed = new Date(Date.UTC(year, month, day, hour, minute, second) - offsetMinutes * 60000);
        } else {
            parsed = new Date(year, month, day, hour, minute, second);
        }
        if (isNaN(parsed.getTime())) return null;
        return parsed;
    }

    function w_isLinkOutOfDate(placedFile, xmpLastModifyDate) {
        if (!w_safeExists(placedFile) || !xmpLastModifyDate) return false;
        var storedDate = w_parseXmpDate(xmpLastModifyDate);
        var currentDate = w_tryGet(function () { return placedFile.modified; }, null);
        if (!storedDate || !currentDate) return false;
        return (currentDate.getTime() - storedDate.getTime()) > 5000;
    }

    /* リンク状態を判定（コードのみ返す。ローカライズはパレット側）*/
    function w_resolveLinkStatus(placedFile, xmpLastModifyDate) {
        if (!w_safeExists(placedFile)) { return { statusCode: "broken", statusIcon: "⚠", isLinkOk: false }; }
        if (w_isLinkOutOfDate(placedFile, xmpLastModifyDate)) { return { statusCode: "update", statusIcon: "⟳", isLinkOk: false }; }
        return { statusCode: "ok", statusIcon: "✓", isLinkOk: true };
    }

    function w_getScaleInfo(item) {
        var matrix = w_tryGet(function () { return item.matrix; }, null);
        if (!matrix) return { scalePct: null, scaleText: "---" };
        var scaleX = Math.sqrt(matrix.mValueA * matrix.mValueA + matrix.mValueB * matrix.mValueB);
        var scaleY = Math.sqrt(matrix.mValueC * matrix.mValueC + matrix.mValueD * matrix.mValueD);
        var scaleXPct = scaleX * 100;
        var scaleYPct = scaleY * 100;
        var averagePct = (scaleXPct + scaleYPct) / 2;
        var scaleText = (Math.abs(scaleXPct - scaleYPct) < 0.1)
            ? scaleXPct.toFixed(1) + "%"
            : scaleXPct.toFixed(1) + "% × " + scaleYPct.toFixed(1) + "%";
        return { scalePct: averagePct, scaleText: scaleText };
    }

    function w_getDimensionsMm(item) {
        var widthPt = w_tryGet(function () { return item.width; }, null);
        var heightPt = w_tryGet(function () { return item.height; }, null);
        if (!widthPt || !heightPt) return { widthMm: null, heightMm: null };
        return { widthMm: widthPt * 25.4 / 72, heightMm: heightPt * 25.4 / 72 };
    }

    function w_getArtboardNumber(item, doc) {
        var itemBounds = w_tryGet(function () { return item.visibleBounds; }, null);
        if (!itemBounds) return null;
        var centerX = (itemBounds[0] + itemBounds[2]) / 2;
        var centerY = (itemBounds[1] + itemBounds[3]) / 2;
        var artboards = w_tryGet(function () { return doc.artboards; }, null);
        if (!artboards) return null;
        for (var i = 0; i < artboards.length; i++) {
            var artboardRect = w_tryGet(function () { return artboards[i].artboardRect; }, null);
            if (!artboardRect) continue;
            var xMin = Math.min(artboardRect[0], artboardRect[2]);
            var xMax = Math.max(artboardRect[0], artboardRect[2]);
            var yMin = Math.min(artboardRect[1], artboardRect[3]);
            var yMax = Math.max(artboardRect[1], artboardRect[3]);
            if (centerX >= xMin && centerX <= xMax && centerY >= yMin && centerY <= yMax) { return i + 1; }
        }
        return null;
    }

    /* XMP メタデータからリンク参照情報を収集 */
    function w_collectXmpLinkedRefs(doc) {
        var refs = [];
        var xmp = w_tryGet(function () { return new XML(doc.XMPString); }, null);
        if (!xmp) return refs;
        var paths = w_tryGet(function () { return xmp.xpath("//stRef:filePath"); }, null);
        var dates = w_tryGet(function () { return xmp.xpath("//stRef:lastModifyDate"); }, null);
        if (!paths) return refs;
        for (var i = 0; i < paths.length(); i++) {
            var filePath = paths[i].toString();
            var lastModifyDate = (dates && i < dates.length()) ? dates[i].toString() : "";
            refs.push({ filePath: filePath, fileName: filePath.replace(/^.*[\/\\]/, ""), lastModifyDate: lastModifyDate });
        }
        return refs;
    }

    /* File.name は URI エンコードされているため表示前にデコードする（失敗時は元の文字列）*/
    function w_decodeName(name) {
        if (name === null || name === undefined || name === "") return name;
        return w_tryGet(function () { return decodeURI(String(name)); }, String(name));
    }

    function w_xmpNameKey(name) {
        if (!name) return "";
        var decoded = w_tryGet(function () { return decodeURI(String(name)); }, String(name));
        return decoded.toLowerCase();
    }

    function w_resolveXmpRef(item, positionalRef, xmpByName) {
        var linkedFile = w_tryGet(function () { return item.file; }, null);
        if (linkedFile) {
            var key = w_xmpNameKey(w_safeProp(linkedFile, "name", ""));
            if (key && xmpByName[key]) return xmpByName[key];
        }
        return positionalRef;
    }

    /* 配置アイテムのファイル名を取得（リンク切れ時は XMP 由来名を優先。不明はセンチネル）*/
    function w_getPlacedItemFileName(placedItem, fallbackName, defaultName) {
        var UNKNOWN = "￾UNKNOWN";
        var fileName = defaultName || UNKNOWN;
        var placedFile = w_tryGet(function () { return placedItem.file; }, null);
        if (placedFile) { fileName = w_decodeName(w_safeProp(placedFile, "name", fileName)); }
        if (fileName === UNKNOWN && fallbackName) { fileName = fallbackName; }
        if (fileName === UNKNOWN) { fileName = w_safeProp(placedItem, "name", fileName); }
        return fileName;
    }

    function w_getPlacementFileBasics(item, xmpRef) {
        var UNKNOWN = "￾UNKNOWN";
        var basics = {
            linkedFile: w_tryGet(function () { return item.file; }, null),
            filePath: "---",
            fileName: UNKNOWN,
            fileSize: "---",
            fileSizeBytes: -1,
            statusCode: "broken",
            statusIcon: "⚠",
            isLinkOk: false
        };
        if (basics.linkedFile) {
            basics.fileName = w_decodeName(w_safeProp(basics.linkedFile, "name", basics.fileName));
            basics.filePath = w_safeProp(basics.linkedFile, "fsName", basics.filePath);
            var resolved = w_resolveLinkStatus(basics.linkedFile, xmpRef ? xmpRef.lastModifyDate : "");
            basics.statusCode = resolved.statusCode;
            basics.statusIcon = resolved.statusIcon;
            basics.isLinkOk = resolved.isLinkOk;
            if (w_safeExists(basics.linkedFile)) {
                var byteLength = w_tryGet(function () { return basics.linkedFile.length; }, -1);
                if (byteLength >= 0) {
                    basics.fileSizeBytes = byteLength;
                    basics.fileSize = w_formatFileSize(basics.fileSizeBytes);
                }
            }
        }
        basics.fileName = w_getPlacedItemFileName(item, xmpRef ? xmpRef.fileName : "", basics.fileName);
        if (basics.fileName === UNKNOWN && basics.filePath !== "---") {
            var derivedName = w_pathBaseName(basics.filePath);
            if (derivedName) basics.fileName = derivedName;
        }
        return basics;
    }

    function w_getPlacementGeometry(item, doc) {
        var dimensions = w_getDimensionsMm(item);
        var widthText = (dimensions.widthMm !== null) ? dimensions.widthMm.toFixed(1) : "---";
        var heightText = (dimensions.heightMm !== null) ? dimensions.heightMm.toFixed(1) : "---";
        var scaleInfo = w_getScaleInfo(item);
        return {
            artboardNum: w_getArtboardNumber(item, doc),
            widthMm: dimensions.widthMm,
            heightMm: dimensions.heightMm,
            widthText: widthText,
            heightText: heightText,
            scalePct: scaleInfo.scalePct,
            scaleText: scaleInfo.scaleText
        };
    }

    function w_getPlacementImageMeta(item, linkedFile) {
        var ppi = null;
        var colorSpace = "";
        if (linkedFile) {
            var info = w_tryGet(function () { return w_readImageInfo(linkedFile); }, null);
            if (info) {
                if (info.width && info.height) { ppi = w_getEffectivePPI(item, { width: info.width, height: info.height }); }
                if (info.colorMode) {
                    colorSpace = info.iccDesc ? info.colorMode + "（" + info.iccDesc + "）" : info.colorMode;
                }
            }
        }
        return { ppi: ppi, ppiText: (ppi !== null) ? String(ppi) : "---", colorSpace: colorSpace };
    }

    function w_buildPlacementEntry(item, itemIndex, xmpRef, doc) {
        var basics = w_getPlacementFileBasics(item, xmpRef);
        var geometry = w_getPlacementGeometry(item, doc);
        var imageMeta = w_getPlacementImageMeta(item, basics.linkedFile);
        return {
            kind: "placed",
            itemIndex: itemIndex,
            index: itemIndex + 1,
            fileName: basics.fileName,
            filePath: basics.filePath,
            fileSize: basics.fileSize,
            fileSizeBytes: basics.fileSizeBytes,
            statusIcon: basics.statusIcon,
            statusCode: basics.statusCode,
            isLinkOk: basics.isLinkOk,
            artboardNum: geometry.artboardNum,
            artboards: (geometry.artboardNum !== null) ? String(geometry.artboardNum) : "-",
            widthMm: geometry.widthMm,
            heightMm: geometry.heightMm,
            widthText: geometry.widthText,
            heightText: geometry.heightText,
            scalePct: geometry.scalePct,
            scaleText: geometry.scaleText,
            ppi: imageMeta.ppi,
            ppiText: imageMeta.ppiText,
            colorSpace: imageMeta.colorSpace,
            itemIndices: [itemIndex]
        };
    }

    /* 埋め込みラスターの itemIndex を placedItems と衝突しない値へずらすためのオフセット。
       これ以上なら rasterItems、未満なら placedItems の添字として解決する（w_itemByIndex）*/
    function w_rasterIndexBase() {
        return 1000000;
    }

    /* エンコード済み itemIndex から実アイテムを取得（placedItems / rasterItems を自動判別）*/
    function w_itemByIndex(doc, itemIndex) {
        var base = w_rasterIndexBase();
        if (itemIndex >= base) return doc.rasterItems[itemIndex - base];
        return doc.placedItems[itemIndex];
    }

    /* RasterItem.imageColorSpace を表示用文字列へ（リンク側の colorMode 表記に合わせる）*/
    function w_rasterColorSpaceText(item) {
        var raw = w_tryGet(function () { return String(item.imageColorSpace); }, "");
        if (!raw) return "";
        var name = raw.replace(/^.*\./, "");
        if (/^grayscale$/i.test(name)) return "Grayscale";
        if (/^lab$/i.test(name)) return "Lab";
        if (/^rgb$/i.test(name)) return "RGB";
        if (/^cmyk$/i.test(name)) return "CMYK";
        return name;
    }

    /* 埋め込みラスター 1 件分のエントリを組み立てる。
       パス・ファイルサイズ・PPI は埋め込み後に取得手段が無いため "---" 固定 */
    function w_buildEmbeddedEntry(item, rasterIndex, displayIndex, doc) {
        var UNKNOWN = "￾UNKNOWN";
        var itemIndex = w_rasterIndexBase() + rasterIndex;
        var geometry = w_getPlacementGeometry(item, doc);
        var fileName = UNKNOWN;
        var sourceFile = w_tryGet(function () { return item.file; }, null);
        if (sourceFile) fileName = w_decodeName(w_safeProp(sourceFile, "name", fileName));
        /* 埋め込み後は item.file を辿れないため、拡張子つきの名前を優先して拾う */
        if (fileName === UNKNOWN) fileName = w_getImageNameFromItem(item) || fileName;
        return {
            kind: "raster",
            itemIndex: itemIndex,
            index: displayIndex,
            fileName: fileName,
            filePath: "---",
            fileSize: "---",
            fileSizeBytes: -1,
            statusIcon: "▣",
            statusCode: "embedded",
            isLinkOk: true,
            artboardNum: geometry.artboardNum,
            artboards: (geometry.artboardNum !== null) ? String(geometry.artboardNum) : "-",
            widthMm: geometry.widthMm,
            heightMm: geometry.heightMm,
            widthText: geometry.widthText,
            heightText: geometry.heightText,
            scalePct: geometry.scalePct,
            scaleText: geometry.scaleText,
            ppi: null,
            ppiText: "---",
            colorSpace: w_rasterColorSpaceText(item),
            itemIndices: [itemIndex]
        };
    }

    /* ドキュメント内の埋め込みラスターを収集（リンク画像は PlacedItem 側で扱う）*/
    function w_collectEmbeddedEntries(doc, displayIndexOffset) {
        var entries = [];
        var rasters = w_tryGet(function () { return doc.rasterItems; }, null);
        if (!rasters) return entries;
        for (var i = 0; i < rasters.length; i++) {
            var raster = rasters[i];
            var isEmbedded = w_tryGet(function () { return raster.embedded; }, true);
            if (isEmbedded === false) continue;
            entries.push(w_buildEmbeddedEntry(raster, i, displayIndexOffset + entries.length + 1, doc));
        }
        return entries;
    }

    function w_getLinkGroupKey(info) {
        if (info.kind === "raster") return "__embedded__#" + info.itemIndex;
        if (info.filePath && info.filePath !== "---") return info.filePath;
        return "__broken__#" + info.itemIndex + "#" + (info.fileName || "");
    }

    function w_assignFileCounts(linkInfoList) {
        var countMap = {};
        for (var i = 0; i < linkInfoList.length; i++) {
            var key = w_getLinkGroupKey(linkInfoList[i]);
            countMap[key] = (countMap[key] || 0) + 1;
        }
        for (var j = 0; j < linkInfoList.length; j++) {
            linkInfoList[j].fileCount = countMap[w_getLinkGroupKey(linkInfoList[j])];
        }
    }

    function w_dedupeByFile(linkInfoList) {
        var uniqueList = [];
        var keyToEntry = {};
        for (var di = 0; di < linkInfoList.length; di++) {
            var dedupInfo = linkInfoList[di];
            var dedupKey = w_getLinkGroupKey(dedupInfo);
            if (keyToEntry[dedupKey]) {
                keyToEntry[dedupKey].itemIndices.push(dedupInfo.itemIndex);
                if (dedupInfo.artboardNum !== null) { keyToEntry[dedupKey].artboardSet[dedupInfo.artboardNum] = true; }
            } else {
                var artboardSet = {};
                if (dedupInfo.artboardNum !== null) artboardSet[dedupInfo.artboardNum] = true;
                var entry = {
                    index: uniqueList.length + 1,
                    kind: dedupInfo.kind,
                    fileName: dedupInfo.fileName,
                    filePath: dedupInfo.filePath,
                    fileSize: dedupInfo.fileSize,
                    fileSizeBytes: dedupInfo.fileSizeBytes,
                    statusIcon: dedupInfo.statusIcon,
                    statusCode: dedupInfo.statusCode,
                    isLinkOk: dedupInfo.isLinkOk,
                    fileCount: dedupInfo.fileCount,
                    artboardNum: dedupInfo.artboardNum,
                    artboardSet: artboardSet,
                    widthMm: dedupInfo.widthMm,
                    heightMm: dedupInfo.heightMm,
                    widthText: dedupInfo.widthText,
                    heightText: dedupInfo.heightText,
                    scalePct: dedupInfo.scalePct,
                    scaleText: dedupInfo.scaleText,
                    ppi: dedupInfo.ppi,
                    ppiText: dedupInfo.ppiText,
                    colorSpace: dedupInfo.colorSpace,
                    itemIndices: [dedupInfo.itemIndex]
                };
                keyToEntry[dedupKey] = entry;
                uniqueList.push(entry);
            }
        }
        for (var ui = 0; ui < uniqueList.length; ui++) {
            var uniqueEntry = uniqueList[ui];
            var artboardNumbers = [];
            for (var abKey in uniqueEntry.artboardSet) {
                if (uniqueEntry.artboardSet.hasOwnProperty(abKey)) artboardNumbers.push(parseInt(abKey, 10));
            }
            artboardNumbers.sort(function (a, b) { return a - b; });
            uniqueEntry.artboards = (artboardNumbers.length > 0) ? artboardNumbers.join(", ") : "-";
            uniqueEntry.artboardNum = (artboardNumbers.length > 0) ? artboardNumbers[0] : null;
        }
        return uniqueList;
    }

    function w_collectLinkInfo(doc, placedItems) {
        var xmpRefs = w_collectXmpLinkedRefs(doc);
        var xmpByName = {};
        for (var r = 0; r < xmpRefs.length; r++) {
            var nameKey = w_xmpNameKey(xmpRefs[r].fileName);
            if (nameKey && !xmpByName[nameKey]) xmpByName[nameKey] = xmpRefs[r];
        }
        var linkInfoList = [];
        for (var i = 0; i < placedItems.length; i++) {
            var effectiveRef = w_resolveXmpRef(placedItems[i], xmpRefs[i] || null, xmpByName);
            linkInfoList.push(w_buildPlacementEntry(placedItems[i], i, effectiveRef, doc));
        }
        /* 埋め込み画像はリンクの後ろに続けて並べる（以降の集計・重複統合は同じ経路を通す）*/
        var embeddedEntries = w_collectEmbeddedEntries(doc, linkInfoList.length);
        for (var e = 0; e < embeddedEntries.length; e++) linkInfoList.push(embeddedEntries[e]);
        w_assignFileCounts(linkInfoList);
        var uniqueList = w_dedupeByFile(linkInfoList);
        return { linkInfoList: linkInfoList, uniqueList: uniqueList };
    }

    /* selection から単一 PlacedItem を抽出（クリップグループ配下に 1 つだけの場合も対応）*/
    function w_pickSinglePlacedItem(selection) {
        if (!selection || selection.length !== 1) return null;
        var topItem = selection[0];
        if (!topItem) return null;
        var typeName = w_tryGet(function () { return topItem.typename; }, "");
        if (typeName === "PlacedItem") return topItem;
        if (typeName !== "GroupItem") return null;
        var found = null;
        var multiple = false;
        function visit(group) {
            if (multiple) return;
            var children = w_tryGet(function () { return group.pageItems; }, null);
            if (!children) return;
            for (var i = 0; i < children.length; i++) {
                var child = children[i];
                var childTypeName = w_tryGet(function () { return child.typename; }, "");
                if (childTypeName === "PlacedItem") {
                    if (found) { multiple = true; return; }
                    found = child;
                } else if (childTypeName === "GroupItem") {
                    visit(child);
                    if (multiple) return;
                }
            }
        }
        visit(topItem);
        return (found && !multiple) ? found : null;
    }

    /* 選択オブジェクトを画面いっぱいにズーム・センタリング */
    function w_zoomToSelection(doc) {
        var selection = doc.selection;
        if (!selection || selection.length === 0) return;
        var firstBounds = selection[0].visibleBounds;
        var minL = firstBounds[0], maxT = firstBounds[1], maxR = firstBounds[2], minB = firstBounds[3];
        for (var i = 1; i < selection.length; i++) {
            var bounds = selection[i].visibleBounds;
            if (bounds[0] < minL) minL = bounds[0];
            if (bounds[1] > maxT) maxT = bounds[1];
            if (bounds[2] > maxR) maxR = bounds[2];
            if (bounds[3] < minB) minB = bounds[3];
        }
        var centerX = (minL + maxR) / 2;
        var centerY = (maxT + minB) / 2;
        var selectionWidth = Math.max(maxR - minL, 1);
        var selectionHeight = Math.max(maxT - minB, 1);
        var view = doc.views[0];
        var viewBounds = view.bounds;
        var viewWidth = Math.abs(viewBounds[2] - viewBounds[0]);
        var viewHeight = Math.abs(viewBounds[1] - viewBounds[3]);
        var margin = 1.2;
        var zoomX = view.zoom * viewWidth / (selectionWidth * margin);
        var zoomY = view.zoom * viewHeight / (selectionHeight * margin);
        var newZoom = Math.min(zoomX, zoomY);
        if (newZoom < 0.03125) newZoom = 0.03125;
        if (newZoom > 64) newZoom = 64;
        view.zoom = newZoom;
        view.centerPoint = [centerX, centerY];
    }

    /* 指定アートボードをアクティブにしてビューをフィット */
    function w_fitViewToArtboard(doc, abIndex) {
        doc.artboards.setActiveArtboardIndex(abIndex);
        var rect = doc.artboards[abIndex].artboardRect;
        var centerX = (rect[0] + rect[2]) / 2;
        var centerY = (rect[1] + rect[3]) / 2;
        var view = doc.views[0];
        view.centerPoint = [centerX, centerY];
        var viewWidth = Math.abs(view.bounds[2] - view.bounds[0]);
        var viewHeight = Math.abs(view.bounds[1] - view.bounds[3]);
        var artboardWidth = Math.abs(rect[2] - rect[0]);
        var artboardHeight = Math.abs(rect[1] - rect[3]);
        var zoomX = view.zoom * viewWidth / artboardWidth;
        var zoomY = view.zoom * viewHeight / artboardHeight;
        view.zoom = Math.min(zoomX, zoomY) * 0.9;
    }

    /* ===== 委譲エントリポイント / Delegated entry points ===== */

    /* 全配置画像を解析し、entries / unique / artboards / preIndex を JSON で返す */
    /* 単独選択されている配置画像／埋め込み画像の itemIndex を返す。該当なしは -1 */
    function w_findSelectedItemIndex(doc) {
        var placed = doc.placedItems;
        var pre = -1;
        var sel = w_tryGet(function () { return doc.selection; }, null);
        var single = w_pickSinglePlacedItem(sel);
        if (single) {
            for (var s = 0; s < placed.length; s++) { if (placed[s] === single) { pre = s; break; } }
            if (pre < 0) {
                var sb = w_tryGet(function () { return single.geometricBounds; }, null);
                if (sb) {
                    for (var t = 0; t < placed.length; t++) {
                        var cb = w_tryGet(function () { return placed[t].geometricBounds; }, null);
                        if (cb && cb[0] === sb[0] && cb[1] === sb[1] && cb[2] === sb[2] && cb[3] === sb[3]) { pre = t; break; }
                    }
                }
            }
        }
        /* 埋め込みラスターが単独選択されている場合はそちらを対象にする */
        if (pre < 0 && sel && sel.length === 1) {
            var selTypeName = w_tryGet(function () { return sel[0].typename; }, "");
            if (selTypeName === "RasterItem") {
                var rasters = w_tryGet(function () { return doc.rasterItems; }, null);
                if (rasters) {
                    for (var r = 0; r < rasters.length; r++) {
                        if (rasters[r] === sel[0]) { pre = w_rasterIndexBase() + r; break; }
                    }
                }
            }
        }
        return pre;
    }

    function w_analyze() {
        try {
            if (app.documents.length === 0) return "NODOC";
            var doc = app.activeDocument;
            var collected = w_collectLinkInfo(doc, doc.placedItems);
            var abs = [];
            for (var a = 0; a < doc.artboards.length; a++) { abs.push(w_safeProp(doc.artboards[a], "name", "")); }
            var payload = {
                docOpen: true,
                artboards: abs,
                preIndex: w_findSelectedItemIndex(doc),
                entries: collected.linkInfoList,
                unique: collected.uniqueList
            };
            return "OK\n" + w_jsonStr(payload);
        } catch (e) {
            return "ERR\n" + e;
        }
    }

    /* カンバスで単独選択されている画像の itemIndex を返す（カンバス→一覧の逆同期用） */
    function w_currentIndex() {
        try {
            if (app.documents.length === 0) return "NODOC";
            return "OK\n" + w_jsonStr({ index: w_findSelectedItemIndex(app.activeDocument) });
        } catch (e) {
            return "ERR\n" + e;
        }
    }

    /* itemIndex 配列に対応する配置画像／埋め込み画像をカンバス上で選択＆ズーム */
    function w_select(indices) {
        try {
            if (app.documents.length === 0) return "NODOC";
            var doc = app.activeDocument;
            doc.selection = null;
            for (var i = 0; i < indices.length; i++) {
                w_tryGet(function () { w_itemByIndex(doc, indices[i]).selected = true; return 1; }, 0);
            }
            w_zoomToSelection(doc);
            app.redraw();
            return "OK";
        } catch (e) {
            return "ERR\n" + e;
        }
    }

    /* 0 始まりアートボード番号にビューをフィット */
    function w_fitArtboard(ab0) {
        try {
            if (app.documents.length === 0) return "NODOC";
            var doc = app.activeDocument;
            if (ab0 < 0 || ab0 >= doc.artboards.length) return "OK";
            w_fitViewToArtboard(doc, ab0);
            app.redraw();
            return "OK";
        } catch (e) {
            return "ERR\n" + e;
        }
    }

    /* ［リンク］パネルを開く */
    function w_openLinksPanel() {
        try {
            app.executeMenuCommand("Adobe LinkPalette Menu Item");
            return "OK";
        } catch (e) {
            return "ERR\n" + e;
        }
    }

    /* PlacedItem の祖先から最も近いクリップグループ（clipped===true）を返す。無ければ null */
    function w_findEnclosingClipGroup(item) {
        var ancestor = w_tryGet(function () { return item.parent; }, null);
        while (ancestor) {
            var typeName = w_tryGet(function () { return ancestor.typename; }, "");
            if (typeName !== "GroupItem") return null;
            var clipped = w_tryGet(function () { return ancestor.clipped; }, false);
            if (clipped) return ancestor;
            ancestor = w_tryGet(function () { return ancestor.parent; }, null);
        }
        return null;
    }

    /* [index, path] のペア配列に沿って placedItem のリンク先を差し替え。counts を返す。
       埋め込み画像（エンコード済み index）は再リンクできないため失敗として数える */
    function w_relinkPairs(pairs) {
        try {
            if (app.documents.length === 0) return "NODOC";
            var placed = app.activeDocument.placedItems;
            var success = 0, failed = 0;
            for (var i = 0; i < pairs.length; i++) {
                if (pairs[i][0] >= w_rasterIndexBase()) { failed++; continue; }
                try {
                    placed[pairs[i][0]].file = new File(pairs[i][1]);
                    success++;
                } catch (e) {
                    failed++;
                }
            }
            app.redraw();
            return "OK\n" + w_jsonStr({ success: success, failed: failed, total: pairs.length });
        } catch (e2) {
            return "ERR\n" + e2;
        }
    }

    /*
     * art item 配下の RasterItem を再帰的に集める（クリップグループの中も辿る）。
     * @param {object} item - PageItem（RasterItem / GroupItem など）
     * @param {Array<RasterItem>} results - 収集先の配列
     * @returns {void}
     */
    function w_collectRasterItems(item, results) {
        var typeName = w_tryGet(function () { return item.typename; }, "");
        if (typeName === "RasterItem") { results.push(item); return; }
        if (typeName !== "GroupItem") return;
        var children = w_tryGet(function () { return item.pageItems; }, null);
        if (!children) return;
        for (var i = 0; i < children.length; i++) w_collectRasterItems(children[i], results);
    }

    /*
     * 埋め込み直後（選択中）のラスターに元リンクのファイル名を拡張子つきで付ける。
     * Illustrator の既定名は拡張子を落とすため、一覧表示と埋め込み解除の名前復元がこれに依存する。
     * ただし親グループが既に同じ名前を持つ場合は付けない（レイヤーパネルで同名が親子に並ぶため）。
     * ラスターが複数に分かれた場合は取り違えるため何もしない。
     */
    function w_nameEmbeddedRaster(doc, linkName) {
        if (!linkName) return;
        var rasters = [];
        w_tryGet(function () {
            var selection = doc.selection;
            for (var i = 0; i < selection.length; i++) w_collectRasterItems(selection[i], rasters);
            return 1;
        }, 0);
        if (rasters.length !== 1) return;
        var raster = rasters[0];
        if (w_safeProp(raster, "name", "") === linkName) return;
        var parentItem = w_tryGet(function () { return raster.parent; }, null);
        var parentIsGroup = !!parentItem && w_tryGet(function () { return parentItem.typename; }, "") === "GroupItem";
        /* 親グループ名は w_getImageNameFromItem が拾うため、ラスターが1つだけのグループなら付け直さない */
        if (parentIsGroup &&
            w_safeProp(parentItem, "name", "") === linkName &&
            w_tryGet(function () { return parentItem.rasterItems.length; }, 0) === 1) return;
        w_tryGet(function () { raster.name = linkName; return 1; }, 0);
    }

    /* グループ内のクリッピングパス（マスク）を返す。無ければ null */
    function w_findClippingPath(group) {
        var items = w_tryGet(function () { return group.pageItems; }, null);
        if (!items) return null;
        for (var i = 0; i < items.length; i++) {
            var item = items[i];
            if (w_tryGet(function () { return item.typename; }, "") !== "PathItem") continue;
            if (w_tryGet(function () { return item.clipping; }, false)) return item;
        }
        return null;
    }

    /* マスクがグループ内の他アイテムを覆いきっている（＝実際にはトリミングしていない）か */
    function w_clipCoversItems(group, maskPath) {
        var maskBounds = w_tryGet(function () { return maskPath.geometricBounds; }, null);
        if (!maskBounds) return false;
        var items = w_tryGet(function () { return group.pageItems; }, null);
        if (!items) return false;
        var tolerance = 0.5;
        for (var i = 0; i < items.length; i++) {
            var item = items[i];
            if (item === maskPath) continue;
            var itemBounds = w_tryGet(function () { return item.geometricBounds; }, null);
            if (!itemBounds) return false;
            if (itemBounds[0] < maskBounds[0] - tolerance) return false;
            if (itemBounds[1] > maskBounds[1] + tolerance) return false;
            if (itemBounds[2] > maskBounds[2] + tolerance) return false;
            if (itemBounds[3] < maskBounds[3] - tolerance) return false;
        }
        return true;
    }

    /* 埋め込みで生じたグループを DOM 操作で解除し、中身を選択状態にする。
       Illustrator は埋め込んだラスターをロックするため executeMenuCommand("ungroup") は効かない。
       クリップグループの場合、マスクが実際にトリミングしているなら見た目が変わるのでグループのまま残す。*/
    function w_releaseEmbeddedGroup(doc, group) {
        if (!group || w_tryGet(function () { return group.typename; }, "") !== "GroupItem") return false;
        var maskPath = null;
        if (w_tryGet(function () { return group.clipped; }, false)) {
            maskPath = w_findClippingPath(group);
            if (!maskPath || !w_clipCoversItems(group, maskPath)) return false;
            w_tryGet(function () { group.clipped = false; return 1; }, 0);
            w_tryGet(function () { maskPath.remove(); return 1; }, 0);
        }
        var moved = [];
        var guard = 0;
        while (w_tryGet(function () { return group.pageItems.length; }, 0) > 0 && guard++ < 1000) {
            var child = group.pageItems[0];
            w_tryGet(function () { child.locked = false; return 1; }, 0);
            w_tryGet(function () { child.hidden = false; return 1; }, 0);
            if (!w_tryGet(function () { child.move(group, ElementPlacement.PLACEBEFORE); return 1; }, 0)) break;
            moved.push(child);
        }
        if (moved.length === 0) return false;
        w_tryGet(function () { group.remove(); return 1; }, 0);
        w_tryGet(function () { doc.selection = null; return 1; }, 0);
        for (var m = 0; m < moved.length; m++) {
            var movedItem = moved[m];
            w_tryGet(function () { movedItem.selected = true; return 1; }, 0);
        }
        return true;
    }

    /* indices の placedItem を埋め込み画像へ変換。PSD はクリップグループ外のときのみグループ解除。counts を返す */
    function w_embed(indices) {
        try {
            if (app.documents.length === 0) return "NODOC";
            var doc = app.activeDocument;
            /* embed() で placedItems の並びが変わるため、先に参照と付随情報を確保しておく */
            var targets = [];
            for (var i = 0; i < indices.length; i++) {
                if (indices[i] >= w_rasterIndexBase()) continue; /* 既に埋め込み済み */
                w_tryGet(function () {
                    var pr = doc.placedItems[indices[i]];
                    var linkName = w_decodeName(w_tryGet(function () { return String(pr.file.name); }, ""));
                    targets.push({
                        placed: pr,
                        linkName: linkName,
                        isPsd: /\.psd$/i.test(linkName),
                        inClipGroup: (w_findEnclosingClipGroup(pr) !== null)
                    });
                    return 1;
                }, 0);
            }
            var success = 0, failed = 0;
            for (var k = 0; k < targets.length; k++) {
                var target = targets[k];
                try {
                    w_tryGet(function () { doc.selection = null; return 1; }, 0);
                    w_tryGet(function () { target.placed.selected = true; return 1; }, 0);
                    target.placed.embed();
                    /* PSD は埋め込み時にレイヤーがグループ化されるため、クリップグループ外なら解除する */
                    if (target.isPsd && !target.inClipGroup) {
                        w_tryGet(function () {
                            var sel = doc.selection;
                            if (sel && sel.length > 0) w_releaseEmbeddedGroup(doc, sel[0]);
                            return 1;
                        }, 0);
                    }
                    /* クリップグループを残した場合、Illustrator が付けたロックのままだと選択も埋め込み解除もできない */
                    w_tryGet(function () {
                        var embedded = [];
                        var sel = doc.selection;
                        for (var s = 0; s < sel.length; s++) w_collectRasterItems(sel[s], embedded);
                        for (var e = 0; e < embedded.length; e++) embedded[e].locked = false;
                        return 1;
                    }, 0);
                    /* 拡張子つきの名前を残す（グループ解除で親グループ名が失われるため）*/
                    w_nameEmbeddedRaster(doc, target.linkName);
                    success++;
                } catch (e) {
                    failed++;
                }
            }
            w_tryGet(function () { doc.selection = null; return 1; }, 0);
            app.redraw();
            return "OK\n" + w_jsonStr({ success: success, failed: failed });
        } catch (e2) {
            return "ERR\n" + e2;
        }
    }

    /* ===== 埋め込み解除（UnembedToLinks.jsx のロジックを worker へ移植）===== */

    /* 埋め込み解除で使う定数。worker へは関数しか転送しないため定数も関数から返す */
    function w_unembedConfig() {
        return {
            actionSetName: "LinkedImageManagerTempSet",
            actionName: "LinkedImageManagerPlace",
            actionFilePath: "~/LinkedImageManagerTemp.aia",
            linksFolderName: "Links",
            exportResolution: 72
        };
    }

    /* adobe_placeDocument のパラメーター定義（記録した .aia から採取した値）。
       ustring（ファイルパス）だけ実行時に差し込む */
    function w_placeParameters() {
        return [
            { key: 1851878757, type: "ustring", value: null },
            { key: 1818848875, type: "boolean", value: "1" },
            { key: 1919970403, type: "boolean", value: "1" },
            { key: 1953329260, type: "boolean", value: "0" },
            { key: 1768779887, type: "boolean", value: "0" },
            { key: 1885828462, type: "boolean", value: "0" },
            { key: 1935895653, type: "real", value: "1.0" },
            { key: 1953656440, type: "real", value: "0.0" },
            { key: 1953656441, type: "real", value: "0.0" }
        ];
    }

    /* 文字列を UTF-8 バイト列へ */
    function w_stringToUtf8Bytes(sourceText) {
        var encodedText = encodeURIComponent(sourceText);
        var byteList = [];
        for (var i = 0; i < encodedText.length; i++) {
            var currentChar = encodedText.charAt(i);
            if (currentChar === "%") {
                byteList.push(parseInt(encodedText.substr(i + 1, 2), 16));
                i += 2;
            } else {
                byteList.push(currentChar.charCodeAt(0));
            }
        }
        return byteList;
    }

    /* バイト列を 16 進文字列にし、記録された .aia と同じく 32 バイトごとに改行する */
    function w_bytesToHexLines(byteList, indentText) {
        var hexText = "";
        for (var i = 0; i < byteList.length; i++) {
            hexText += (byteList[i] < 16 ? "0" : "") + byteList[i].toString(16);
        }
        var lineList = [];
        for (var start = 0; start < hexText.length; start += 64) {
            lineList.push(indentText + hexText.substr(start, 64));
        }
        return lineList.join("\n");
    }

    /* .aia のテキスト値（[ バイト数 16進 ]）を組み立てる */
    function w_buildTextValue(sourceText, indentText) {
        var byteList = w_stringToUtf8Bytes(sourceText);
        if (byteList.length === 0) return "[ 0 ]";
        return "[ " + byteList.length + "\n"
            + w_bytesToHexLines(byteList, indentText + "\t") + "\n"
            + indentText + "]";
    }

    /* パラメーター 1 件分の .aia ブロック */
    function w_buildParameterBlock(index, parameter, filePath) {
        var indentText = "\t\t\t";
        var valueText = (parameter.value !== null) ? parameter.value : w_buildTextValue(filePath, indentText);
        return "\t\t/parameter-" + index + " {\n"
            + indentText + "/key " + parameter.key + "\n"
            + indentText + "/showInPalette 4294967295\n"
            + indentText + "/type (" + parameter.type + ")\n"
            + indentText + "/value " + valueText + "\n"
            + "\t\t}\n";
    }

    /* 配置（adobe_placeDocument）を実行する一時アクションのソースを生成 */
    function w_buildActionSource(setName, actionName, filePath) {
        var parameters = w_placeParameters();
        var parameterText = "";
        for (var i = 0; i < parameters.length; i++) {
            parameterText += w_buildParameterBlock(i + 1, parameters[i], filePath);
        }
        return ""
            + "/version 3\n"
            + "/name " + w_buildTextValue(setName, "") + "\n"
            + "/isOpen 1\n"
            + "/actionCount 1\n"
            + "/action-1 {\n"
            + "\t/name " + w_buildTextValue(actionName, "\t") + "\n"
            + "\t/keyIndex 0\n"
            + "\t/colorIndex 0\n"
            + "\t/isOpen 1\n"
            + "\t/eventCount 1\n"
            + "\t/event-1 {\n"
            + "\t\t/useRulersIn1stQuadrant 0\n"
            + "\t\t/internalName (adobe_placeDocument)\n"
            + "\t\t/localizedName [ 0 ]\n"
            + "\t\t/isOpen 1\n"
            + "\t\t/isOn 1\n"
            + "\t\t/hasDialog 1\n"
            + "\t\t/showDialog 0\n"
            + "\t\t/parameterCount " + parameters.length + "\n"
            + parameterText
            + "\t}\n"
            + "}\n";
    }

    /* 一時アクションを書き出して読み込み、実行後に必ず後片付けする */
    function w_playTemporaryAction(actionSource, setName, actionName, actionFilePath) {
        var actionFile = new File(actionFilePath);
        w_tryGet(function () { app.unloadAction(setName, ""); return 1; }, 0);
        try {
            actionFile.lineFeed = "Unix";
            if (!actionFile.open("w")) throw new Error("ACTION_FILE_FAILED");
            actionFile.write(actionSource);
            actionFile.close();
            app.loadAction(actionFile);
            app.doScript(actionName, setName, false);
        } finally {
            actionFile.close();
            if (actionFile.exists) actionFile.remove();
            w_tryGet(function () { app.unloadAction(setName, ""); return 1; }, 0);
        }
    }

    /* 選択できない理由コードを返す。選択できる場合は空文字（祖先のロック・非表示も見る）*/
    /* 選択を妨げるロックを一時的に外す。Illustrator は埋め込んだラスターを自動でロックするため。
       戻り値は復元対象の配列 */
    function w_unlockForSelection(item) {
        var unlockedNodes = [];
        var node = item;
        while (node && w_tryGet(function () { return node.typename; }, "Document") !== "Document") {
            var target = node;
            if (w_tryGet(function () { return target.locked; }, false)) {
                if (w_tryGet(function () { target.locked = false; return 1; }, 0)) unlockedNodes.push(target);
            }
            node = w_tryGet(function () { return target.parent; }, null);
        }
        return unlockedNodes;
    }

    /* w_unlockForSelection で外したロックを戻す（差し替えで消えたアイテムは無視される）*/
    function w_restoreLocks(unlockedNodes) {
        for (var i = 0; i < unlockedNodes.length; i++) {
            var target = unlockedNodes[i];
            w_tryGet(function () { target.locked = true; return 1; }, 0);
        }
    }

    function w_findSelectionBlocker(item) {
        var node = item;
        while (node && node.typename !== "Document") {
            if (node.typename === "Layer") {
                if (node.locked) return "LOCKED_LAYER";
                if (!node.visible) return "HIDDEN_LAYER";
            } else {
                if (node.locked) return "LOCKED_ITEM";
                if (node.hidden) return "HIDDEN_ITEM";
            }
            node = node.parent;
        }
        return "";
    }

    /* 2 つのアイテムが同じアートオブジェクトを指しているか（uuid を参照できない場合は判定しない）*/
    function w_isSameArtItem(itemA, itemB) {
        return w_tryGet(function () { return itemA.uuid === itemB.uuid; }, true);
    }

    /* 埋め込み画像を一時アクション経由でリンク画像に置き換える。
       「選択オブジェクトと置換」で配置するため、位置・サイズ・回転・重ね順はアクション側が引き継ぐ */
    function w_relinkByAction(doc, item, targetFile, status) {
        /* 埋め込み時に付いたロックは選択できない原因になるだけなので、置換の間だけ外す */
        var unlockedNodes = w_unlockForSelection(item);
        try {
            var blocker = w_findSelectionBlocker(item);
            if (blocker) throw new Error(blocker);
            var config = w_unembedConfig();
            doc.activeLayer = item.layer;
            doc.selection = null;
            item.selected = true;
            var currentSelection = doc.selection;
            if (!currentSelection || currentSelection.length !== 1 || !w_isSameArtItem(currentSelection[0], item)) {
                doc.selection = null;
                throw new Error("SELECT_FAILED");
            }
            var actionSource = w_buildActionSource(config.actionSetName, config.actionName, targetFile.fsName);
            /* ここから先はアクションが実行済みとして扱う */
            status.actionPlayed = true;
            w_playTemporaryAction(actionSource, config.actionSetName, config.actionName, config.actionFilePath);
            var newSelection = doc.selection;
            if (!newSelection || newSelection.length === 0 || newSelection[0].typename !== "PlacedItem") {
                throw new Error("ACTION_RESULT_MISSING");
            }
            return newSelection[0];
        } finally {
            w_restoreLocks(unlockedNodes);
        }
    }

    /* 2 つのファイルを同一とみなせるか */
    function w_isSameFile(fileA, fileB) {
        if (fileA.fsName === fileB.fsName) return true;
        return fileA.length === fileB.length && fileA.modified.getTime() === fileB.modified.getTime();
    }

    /* 収集先ファイルを決める。同名で内容が異なる場合は連番を付ける */
    function w_resolveCollectDestination(linksFolder, sourceFile) {
        var fileName = w_decodeName(sourceFile.name);
        var dotIndex = fileName.lastIndexOf(".");
        var baseName = (dotIndex > 0) ? fileName.substring(0, dotIndex) : fileName;
        var extension = (dotIndex > 0) ? fileName.substring(dotIndex) : "";
        var destFile = new File(linksFolder.fsName + "/" + fileName);
        var counter = 1;
        while (destFile.exists && !w_isSameFile(destFile, sourceFile)) {
            destFile = new File(linksFolder.fsName + "/" + baseName + "-" + counter + extension);
            counter++;
        }
        return destFile;
    }

    /* ドキュメントと同階層の「Links」フォルダーを返す（無ければ作成）*/
    function w_getLinksFolder(doc) {
        var config = w_unembedConfig();
        var docFile = w_tryGet(function () { return doc.fullName; }, null);
        if (!docFile || !docFile.exists) throw new Error("DOC_NOT_SAVED");
        var linksFolder = new Folder(docFile.parent.fsName + "/" + config.linksFolderName);
        if (!linksFolder.exists && !linksFolder.create()) throw new Error("LINKS_FOLDER_FAILED");
        return linksFolder;
    }

    /* リンク先を「Links」フォルダーへ複製してリンクを張り替える */
    function w_collectLink(doc, placedItem, sourceFile) {
        var linksFolder = w_getLinksFolder(doc);
        var destFile = w_resolveCollectDestination(linksFolder, sourceFile);
        if (!destFile.exists && !sourceFile.copy(destFile.fsName)) throw new Error("COPY_FAILED");
        placedItem.file = destFile;
        return destFile;
    }

    /* 埋め込み画像の元ファイル（参照自体が失敗することがあるため tryGet 経由）*/
    function w_getEmbeddedSourceFile(item) {
        return w_tryGet(function () {
            return (item.file && item.file.exists) ? item.file : null;
        }, null);
    }

    /*
     * 埋め込み画像自身が持つ名前を返す。
     * Illustrator は埋め込み時、RasterItem にはレイヤー名（拡張子なし）を、
     * 親グループに拡張子つきの元ファイル名を付けるため、拡張子つきの名前を優先する。
     * 親グループ名は、そのグループにラスターが 1 つだけのときのみ採用する（取り違え防止）。
     * @param {RasterItem} item - 埋め込み画像
     * @returns {string} 名前（取得できないときは空文字）
     */
    function w_getImageNameFromItem(item) {
        var hasExtension = /\.[a-z][a-z0-9]{1,4}\s*$/i;
        var itemName = w_safeProp(item, "name", "");
        if (itemName && hasExtension.test(itemName)) return itemName;
        var parentItem = w_tryGet(function () { return item.parent; }, null);
        if (!parentItem || w_tryGet(function () { return parentItem.typename; }, "") !== "GroupItem") return itemName;
        if (w_tryGet(function () { return parentItem.rasterItems.length; }, 0) !== 1) return itemName;
        var parentName = w_safeProp(parentItem, "name", "");
        return hasExtension.test(parentName) ? parentName : itemName;
    }

    /* XMP マニフェストから埋め込み前の元ファイル名を出現順に取得 */
    function w_getManifestFileNames(doc) {
        var nameList = [];
        var foundNames = {};
        var filePaths = w_tryGet(function () {
            var documentXMP = new XML(doc.XMPString);
            var refs = documentXMP.xpath("//stMfs:reference/stRef:filePath");
            if (refs == null || refs.length() === 0) refs = documentXMP.xpath("//stRef:filePath");
            return refs;
        }, null);
        if (filePaths == null) return nameList;
        for (var i = 0; i < filePaths.length(); i++) {
            var fileName = w_decodeName(String(filePaths[i])).replace(/^.*[\/\\]/, "");
            var nameKey = fileName.toLowerCase();
            if (fileName === "" || foundNames[nameKey]) continue;
            foundNames[nameKey] = true;
            nameList.push(fileName);
        }
        return nameList;
    }

    /* 名前が取れなかった画像を XMP マニフェストの元ファイル名で補う。
       候補数と名前未定の件数が一致するときだけ割り当て、取り違えを避ける */
    function w_fillNamesFromManifest(doc, nameList) {
        var manifestNames = w_getManifestFileNames(doc);
        if (manifestNames.length === 0) return;
        var knownNames = {};
        var missingCount = 0;
        var i;
        for (i = 0; i < nameList.length; i++) {
            if (nameList[i] === "") { missingCount++; continue; }
            knownNames[nameList[i].toLowerCase()] = true;
        }
        if (missingCount === 0) return;
        var remainingNames = [];
        for (i = 0; i < manifestNames.length; i++) {
            if (!knownNames[manifestNames[i].toLowerCase()]) remainingNames.push(manifestNames[i]);
        }
        if (remainingNames.length !== missingCount) return;
        var nameIndex = 0;
        for (i = 0; i < nameList.length; i++) {
            if (nameList[i] === "") { nameList[i] = remainingNames[nameIndex]; nameIndex++; }
        }
    }

    /* 拡張子と使用できない文字を除いてファイル名として整える */
    function w_toSafeBaseName(fileName, fallbackName) {
        var baseName = (fileName || "").replace(/\.[a-z][a-z0-9]{1,4}$/i, "").replace(/^\s+|\s+$/g, "");
        baseName = baseName.replace(/[\\\/:*?"<>|]/g, "_");
        return (baseName === "") ? fallbackName : baseName;
    }

    /* 埋め込み画像ごとの書き出し名（拡張子なし）を uuid をキーに決める */
    function w_buildExportNameMap(doc) {
        var allItems = doc.rasterItems;
        var nameList = [];
        var i;
        for (i = 0; i < allItems.length; i++) { nameList.push(w_getImageNameFromItem(allItems[i])); }
        w_fillNamesFromManifest(doc, nameList);
        var nameMap = {};
        for (i = 0; i < allItems.length; i++) {
            var uuid = w_tryGet(function () { return allItems[i].uuid; }, "");
            if (uuid !== "") nameMap[uuid] = w_toSafeBaseName(nameList[i], "image" + (i + 1));
        }
        return nameMap;
    }

    /* PSD 書き出しに対応するカラースペースか（CMYK / RGB / グレースケール）*/
    function w_isSupportedColorSpace(imageColorSpace) {
        return imageColorSpace == ImageColorSpace.CMYK
            || imageColorSpace == ImageColorSpace.RGB
            || imageColorSpace == ImageColorSpace.GrayScale;
    }

    /* 画像に蓄積された回転角（度）*/
    function w_getAccumulatedRotation(item) {
        var itemTags = w_tryGet(function () { return item.tags; }, null);
        if (!itemTags) return 0;
        for (var i = 0; i < itemTags.length; i++) {
            if (itemTags[i].name === "BBAccumRotation") return itemTags[i].value * 180 / Math.PI;
        }
        return 0;
    }

    /* 画像の拡大率と回転角（RasterItem は行列の Y 方向が反転している）*/
    function w_getScaleAndRotation(item) {
        var placedItemFlip = (item.typename === "PlacedItem") ? 1 : -1;
        var rotationAngle = w_getAccumulatedRotation(item);
        var unrotatedMatrix = app.concatenateRotationMatrix(item.matrix, rotationAngle * placedItemFlip);
        return {
            scaleX: unrotatedMatrix.mValueA * 100,
            scaleY: unrotatedMatrix.mValueD * -100 * placedItemFlip,
            rotation: rotationAngle
        };
    }

    /* 既存ファイルを上書きしないパス（連番の接尾辞を付ける）*/
    function w_getNonOverwritingFilePath(filePath) {
        var suffixIndex = 1;
        var pathParts = filePath.split(/(\.[^\.]+)$/);
        while (File(filePath).exists) {
            suffixIndex++;
            filePath = pathParts[0] + "(" + suffixIndex + ")" + pathParts[1];
        }
        return filePath;
    }

    /* 書き出し用の新規ドキュメントを作成 */
    function w_createExportDocument(documentTitle, imageColorSpace) {
        var documentPreset = new DocumentPreset();
        var documentPresetType;
        documentPreset.title = documentTitle;
        documentPreset.width = 1000;
        documentPreset.height = 1000;
        if (imageColorSpace == ImageColorSpace.RGB) {
            documentPresetType = DocumentPresetType.BasicRGB;
            documentPreset.colorMode = DocumentColorSpace.RGB;
        } else {
            documentPresetType = DocumentPresetType.BasicCMYK;
            documentPreset.colorMode = DocumentColorSpace.CMYK;
        }
        return app.documents.addDocument(documentPresetType, documentPreset);
    }

    /* ドキュメントを PSD として書き出す */
    function w_exportDocumentAsPSD(exportDocument, exportFilePath, imageColorSpace, resolution) {
        var exportedFile = File(exportFilePath);
        var psdOptions = new ExportOptionsPhotoshop();
        psdOptions.antiAliasing = false;
        psdOptions.artBoardClipping = true;
        psdOptions.imageColorSpace = imageColorSpace;
        psdOptions.editableText = false;
        psdOptions.flatten = true;
        psdOptions.maximumEditability = false;
        psdOptions.resolution = (resolution || 72);
        psdOptions.warnings = false;
        psdOptions.writeLayers = false;
        exportDocument.exportFile(exportedFile, ExportType.PHOTOSHOP, psdOptions);
        return exportedFile;
    }

    /* 元ファイルが不明な埋め込み画像を一時ドキュメントへ複製し、等倍・回転なしで PSD に書き出す。
       再配置はアクションの置換が行うため、ここでは変形を戻した素の画像だけを書き出す */
    function w_exportEmbeddedImageAsPSD(doc, item, baseName) {
        var config = w_unembedConfig();
        var linksFolder = w_getLinksFolder(doc);
        var imageColorSpace = item.imageColorSpace;
        var scaleAndRotation = w_getScaleAndRotation(item);
        /* 0 で割ると変形行列が壊れるため、等倍に戻せない画像はここで中止 */
        if (!scaleAndRotation.scaleX || !scaleAndRotation.scaleY) throw new Error("SCALE_UNAVAILABLE");
        var previousInteractionLevel = app.userInteractionLevel;
        app.userInteractionLevel = UserInteractionLevel.DONTDISPLAYALERTS;
        var exportDocument = w_createExportDocument(baseName, imageColorSpace);
        var exportedFile;
        try {
            var workingImage = item.duplicate(exportDocument.layers[0], ElementPlacement.PLACEATBEGINNING);
            var transformMatrix = app.getRotationMatrix(-scaleAndRotation.rotation);
            transformMatrix = app.concatenateScaleMatrix(transformMatrix,
                100 / scaleAndRotation.scaleX * 100,
                100 / scaleAndRotation.scaleY * 100);
            workingImage.transform(transformMatrix, true, true, true, true, true);
            workingImage.position = [0, workingImage.height];
            exportDocument.artboards[0].artboardRect = [0, workingImage.height, workingImage.width, 0];
            var exportFilePath = w_getNonOverwritingFilePath(linksFolder.fsName + "/" + baseName + ".psd");
            exportedFile = w_exportDocumentAsPSD(exportDocument, exportFilePath, imageColorSpace, config.exportResolution);
        } finally {
            exportDocument.close(SaveOptions.DONOTSAVECHANGES);
            app.userInteractionLevel = previousInteractionLevel;
            /* 一時ドキュメントを閉じたあと、確実に元のドキュメントへ戻す */
            app.activeDocument = doc;
        }
        if (!exportedFile.exists) throw new Error("EXPORT_FAILED");
        return exportedFile;
    }

    /* indices の埋め込み画像をリンク画像へ戻す。collect が true なら「Links」フォルダーへ収集。
       counts と失敗理由（コード）を返す */
    function w_unembed(indices, collect) {
        try {
            if (app.documents.length === 0) return "NODOC";
            var doc = app.activeDocument;
            /* 置換で rasterItems の並びが変わるため、先に参照を確保しておく */
            var targets = [];
            for (var i = 0; i < indices.length; i++) {
                if (indices[i] < w_rasterIndexBase()) continue;
                w_tryGet(function () { targets.push(w_itemByIndex(doc, indices[i])); return 1; }, 0);
            }
            var success = 0, skipped = 0, failed = 0;
            var details = [];
            var placedList = [];
            var nameMap = w_tryGet(function () { return w_buildExportNameMap(doc); }, {});
            for (var k = 0; k < targets.length; k++) {
                var item = targets[k];
                var itemName = w_safeProp(item, "name", "") || ("#" + (k + 1));
                var targetFile = w_getEmbeddedSourceFile(item);
                var isExported = false;
                if (!targetFile) {
                    if (!w_isSupportedColorSpace(w_tryGet(function () { return item.imageColorSpace; }, null))) {
                        skipped++;
                        details.push({ name: itemName, code: "UNSUPPORTED_COLORSPACE" });
                        continue;
                    }
                    var baseName = w_tryGet(function () { return nameMap[item.uuid]; }, "") || ("image" + (k + 1));
                    try {
                        targetFile = w_exportEmbeddedImageAsPSD(doc, item, baseName);
                        isExported = true;
                    } catch (eExport) {
                        failed++;
                        details.push({ name: itemName, code: String(eExport.message) });
                        continue;
                    }
                }
                var status = { actionPlayed: false };
                try {
                    var placedItem = w_relinkByAction(doc, item, targetFile, status);
                    placedList.push(placedItem);
                    /* 書き出した PSD はすでに収集先にあるため、収集は不要 */
                    if (collect && !isExported) {
                        try {
                            w_collectLink(doc, placedItem, targetFile);
                        } catch (eCollect) {
                            details.push({ name: itemName, code: String(eCollect.message) });
                        }
                    }
                    success++;
                } catch (eRelink) {
                    /* アクション実行前に失敗した場合だけ、未使用の PSD を片付ける */
                    if (isExported && !status.actionPlayed) {
                        w_tryGet(function () { if (targetFile.exists) targetFile.remove(); return 1; }, 0);
                    }
                    failed++;
                    details.push({ name: itemName, code: String(eRelink.message) });
                }
            }
            w_tryGet(function () { doc.selection = null; return 1; }, 0);
            for (var p = 0; p < placedList.length; p++) {
                w_tryGet(function () { placedList[p].selected = true; return 1; }, 0);
            }
            app.redraw();
            return "OK\n" + w_jsonStr({ success: success, skipped: skipped, failed: failed, details: details });
        } catch (e2) {
            return "ERR\n" + e2;
        }
    }

    /* 削除対象にクリップグループ内のものが含まれるか。CLIP / PLAIN を返す */
    function w_probeDelete(indices) {
        try {
            if (app.documents.length === 0) return "NODOC";
            var doc = app.activeDocument;
            for (var i = 0; i < indices.length; i++) {
                var g = w_tryGet(function () { return w_findEnclosingClipGroup(w_itemByIndex(doc, indices[i])); }, null);
                if (g) return "CLIP";
            }
            return "PLAIN";
        } catch (e) {
            return "ERR\n" + e;
        }
    }

    /* indices を一括削除（clipMode==='group' なら囲むクリップグループごと）。counts を返す */
    function w_del(indices, clipMode) {
        try {
            if (app.documents.length === 0) return "NODOC";
            var doc = app.activeDocument;
            var refs = [];
            for (var i = 0; i < indices.length; i++) {
                w_tryGet(function () {
                    var pr = w_itemByIndex(doc, indices[i]);
                    refs.push({ placed: pr, clipGroup: w_findEnclosingClipGroup(pr) });
                    return 1;
                }, 0);
            }
            var success = 0, failed = 0;
            for (var k = 0; k < refs.length; k++) {
                try {
                    if (refs[k].clipGroup && clipMode === "group") { refs[k].clipGroup.remove(); }
                    else { refs[k].placed.remove(); }
                    success++;
                } catch (eRm) {
                    failed++;
                }
            }
            app.redraw();
            return "OK\n" + w_jsonStr({ success: success, failed: failed });
        } catch (e2) {
            return "ERR\n" + e2;
        }
    }

    /* アクティブドキュメントの親フォルダー fsName を返す。未保存は body 空 */
    function w_docFolder() {
        try {
            if (app.documents.length === 0) return "NODOC";
            var doc = app.activeDocument;
            var f = w_tryGet(function () { return doc.fullName; }, null);
            if (!f || !f.parent || !f.parent.exists) return "OK\n";
            return "OK\n" + f.parent.fsName;
        } catch (e) {
            return "ERR\n" + e;
        }
    }

    /* 文字列をクリップボードへコピー（一時 TextFrame + app.copy）*/
    function w_copyText(text) {
        try {
            if (text === null || text === undefined) text = "";
            text = String(text);
            if (app.documents.length === 0) return "ERR\nnodoc";
            var doc = app.activeDocument;
            var prev = w_tryGet(function () { return [].slice.call(doc.selection); }, null);
            var tmp = null;
            try {
                tmp = doc.textFrames.add();
                tmp.contents = text;
                doc.selection = null;
                tmp.selected = true;
                app.copy();
                tmp.remove();
                tmp = null;
                w_tryGet(function () { doc.selection = null; return 1; }, 0);
                if (prev && prev.length) {
                    for (var i = 0; i < prev.length; i++) {
                        w_tryGet(function () { prev[i].selected = true; return 1; }, 0);
                    }
                }
                return "OK";
            } catch (e) {
                w_tryGet(function () { if (tmp) tmp.remove(); return 1; }, 0);
                return "ERR\n" + e;
            }
        } catch (e2) {
            return "ERR\n" + e2;
        }
    }

    // worker 関数は全てここに登録（追加漏れ＝委譲側だけ無言で壊れる）
    var WORKER_FUNCS = [
        w_jsonStr, w_tryGet, w_safeExists, w_safeProp, w_readU16BE, w_readU32BE, w_readIccDesc,
        w_readPngImageInfo, w_readJpegImageInfo, w_readPsdImageInfo, w_readImageInfo,
        w_lastPathSep, w_pathBaseName, w_formatFileSize, w_getEffectivePPI, w_parseXmpDate,
        w_isLinkOutOfDate, w_resolveLinkStatus, w_getScaleInfo, w_getDimensionsMm, w_getArtboardNumber,
        w_collectXmpLinkedRefs, w_decodeName, w_xmpNameKey, w_resolveXmpRef, w_getPlacedItemFileName,
        w_getPlacementFileBasics, w_getPlacementGeometry, w_getPlacementImageMeta, w_buildPlacementEntry,
        w_rasterIndexBase, w_itemByIndex, w_rasterColorSpaceText, w_buildEmbeddedEntry, w_collectEmbeddedEntries,
        w_getLinkGroupKey, w_assignFileCounts, w_dedupeByFile, w_collectLinkInfo, w_pickSinglePlacedItem,
        w_zoomToSelection, w_fitViewToArtboard, w_findEnclosingClipGroup,
        w_findSelectedItemIndex, w_analyze, w_currentIndex, w_select, w_fitArtboard, w_openLinksPanel,
        w_relinkPairs, w_collectRasterItems, w_nameEmbeddedRaster,
        w_findClippingPath, w_clipCoversItems, w_releaseEmbeddedGroup, w_embed,
        w_unembedConfig, w_placeParameters, w_stringToUtf8Bytes, w_bytesToHexLines, w_buildTextValue,
        w_buildParameterBlock, w_buildActionSource, w_playTemporaryAction,
        w_unlockForSelection, w_restoreLocks, w_findSelectionBlocker,
        w_isSameArtItem, w_relinkByAction, w_isSameFile, w_resolveCollectDestination, w_getLinksFolder,
        w_collectLink, w_getEmbeddedSourceFile, w_getImageNameFromItem, w_getManifestFileNames,
        w_fillNamesFromManifest, w_toSafeBaseName, w_buildExportNameMap, w_isSupportedColorSpace,
        w_getAccumulatedRotation, w_getScaleAndRotation, w_getNonOverwritingFilePath,
        w_createExportDocument, w_exportDocumentAsPSD, w_exportEmbeddedImageAsPSD, w_unembed,
        w_probeDelete, w_del, w_docFolder, w_copyText
    ];

    // =========================================
    // パレット側データ状態 / Palette-side data state
    // =========================================

    // worker から受け取ったシリアライズ済みデータ（DOM 参照は保持しない）
    var allPlacementEntries = [];
    var uniqueFileEntries = [];
    var artboardNames = [];
    var preIndex = -1;
    // 下部ステータス表示（showPalette で代入）
    var statusText = null;
    // 直近のステータス文言。statusText 生成前（初回 loadData）のメッセージを取りこぼさないため保持する
    var lastStatusMessage = "";

    function setStatus(msg) {
        lastStatusMessage = msg;
        if (statusText) statusText.text = msg;
    }

    // worker が返す statusCode / センチネルをローカライズ済み表示へ変換（破壊的）
    function localizeEntry(e) {
        var disp = statusDisplay(e.statusCode);
        e.status = disp.status;
        if (e.statusIcon === undefined || e.statusIcon === null || e.statusIcon === "") e.statusIcon = disp.statusIcon;
        e.isLinkOk = disp.isLinkOk;
        if (e.fileName === UNKNOWN_NAME) e.fileName = L('label.fileNameUnknown');
    }

    // 保持しているデータ状態をすべて初期化（読み込み失敗時に古い内容を残さない）
    function clearLoadedData() {
        allPlacementEntries = [];
        uniqueFileEntries = [];
        artboardNames = [];
        preIndex = -1;
    }

    // 解析を委譲してデータを読み込む。成功なら true / 失敗なら false / 実行中なら null
    function loadData() {
        if (isBusy) return null;
        isBusy = true;
        try {
            var parsed = parseWorkerResult(delegate("$.global.__LIM.analyze()"));
            if (parsed.marker === "NODOC") {
                clearLoadedData();
                setStatus(L('status.noDocument'));
                return false;
            }
            if (parsed.marker !== "OK") {
                clearLoadedData();
                setStatus(L('status.loadFailed') + "（" + parsed.body + "）");
                return false;
            }
            var data = tryGet(function () { return eval("(" + parsed.body + ")"); }, null);
            if (!data) {
                clearLoadedData();
                setStatus(L('status.loadFailed'));
                return false;
            }
            allPlacementEntries = data.entries || [];
            uniqueFileEntries = data.unique || [];
            artboardNames = data.artboards || [];
            preIndex = (typeof data.preIndex === "number") ? data.preIndex : -1;
            for (var i = 0; i < allPlacementEntries.length; i++) localizeEntry(allPlacementEntries[i]);
            for (var j = 0; j < uniqueFileEntries.length; j++) localizeEntry(uniqueFileEntries[j]);
            if (allPlacementEntries.length === 0) setStatus(L('status.noPlaced'));
            else setStatus(L('status.loaded') + "：" + withUnit(allPlacementEntries.length, 'label.items'));
            return true;
        } finally {
            isBusy = false;
        }
    }

    // カンバスで単独選択されている画像の itemIndex を取得（該当なし・失敗時は -1）
    function delegateCurrentIndex() {
        if (isBusy) return -1;
        isBusy = true;
        try {
            var res = parseWorkerResult(delegate("$.global.__LIM.currentIndex()"));
            if (res.marker !== "OK") return -1;
            var data = tryGet(function () { return eval("(" + res.body + ")"); }, null);
            return (data && typeof data.index === "number") ? data.index : -1;
        } finally {
            isBusy = false;
        }
    }

    // カンバス選択＆ズームを委譲
    function delegateSelect(indices) {
        if (isBusy) return;
        isBusy = true;
        try {
            parseWorkerResult(delegate("$.global.__LIM.select(" + jsIntArray(indices) + ")"));
        } finally {
            isBusy = false;
        }
    }

    // アートボードへのフィットを委譲（ab0 は 0 始まり）
    function delegateFitArtboard(ab0) {
        if (isBusy) return;
        isBusy = true;
        try {
            parseWorkerResult(delegate("$.global.__LIM.fitArtboard(" + parseInt(ab0, 10) + ")"));
        } finally {
            isBusy = false;
        }
    }

    // ［リンク］パネルを開く委譲
    function delegateOpenLinksPanel() {
        if (isBusy) return;
        isBusy = true;
        try {
            parseWorkerResult(delegate("$.global.__LIM.openLinksPanel()"));
        } finally {
            isBusy = false;
        }
    }

    // [index, path] ペア配列を worker 呼び出し式へ
    function pairsExpr(pairs) {
        var parts = [];
        for (var i = 0; i < pairs.length; i++) {
            parts.push("[" + parseInt(pairs[i][0], 10) + "," + jsString(pairs[i][1]) + "]");
        }
        return "[" + parts.join(",") + "]";
    }

    // 再リンク（ペア一括）を委譲。counts オブジェクト or null
    function delegateRelinkPairs(pairs) {
        var res = parseWorkerResult(delegate("$.global.__LIM.relinkPairs(" + pairsExpr(pairs) + ")"));
        if (res.marker !== "OK") return null;
        return tryGet(function () { return eval("(" + res.body + ")"); }, null);
    }

    // 埋め込み変換を委譲。counts オブジェクト or null
    function delegateEmbed(indices) {
        var res = parseWorkerResult(delegate("$.global.__LIM.embed(" + jsIntArray(indices) + ")"));
        if (res.marker !== "OK") return null;
        return tryGet(function () { return eval("(" + res.body + ")"); }, null);
    }

    // 埋め込み解除を委譲。PSD 書き出しとアクション実行を含むためタイムアウトを長く取り、リトライしない
    function delegateUnembed(indices, collect) {
        var res = parseWorkerResult(delegateOnce(
            "$.global.__LIM.unembed(" + jsIntArray(indices) + "," + (collect ? "true" : "false") + ")", 300));
        if (res.marker !== "OK") return null;
        return tryGet(function () { return eval("(" + res.body + ")"); }, null);
    }

    // worker が返す失敗理由コードを表示用テキストへ（未知のコードはそのまま返す）
    function unembedReasonText(code) {
        var key = 'reason.' + code;
        var text = L(key);
        return (text === key) ? String(code) : text;
    }

    // 拡張子変更ダイアログ（参照フォルダー＋拡張子を選ぶモーダル）
    // 戻り値: { referenceFolder, primaryExt, fallbackExt } or null
    function showChangeExtensionDialog() {
        var extdialog = new Window("dialog", L('dialog.changeExt'));
        setupWindow(extdialog);

        var folderPanel = extdialog.add("panel", undefined, L('panel.destFolder'));
        setupPanel(folderPanel);

        var chooseFolderBtn = folderPanel.add("button", undefined, L('button.chooseFolder'));
        chooseFolderBtn.alignment = ["left", "top"];

        var folderLabel = folderPanel.add("statictext", undefined, L('label.extensionReferenceFolderPlaceholder'));
        folderLabel.alignment = ["fill", "top"];
        folderLabel.preferredSize = [300, 20];

        var extPanel = extdialog.add("panel", undefined, L('panel.extension'));
        setupPanel(extPanel);

        var referenceFolder = null;
        var okBtn = null;

        chooseFolderBtn.onClick = function () {
            var selectedFolder = Folder.selectDialog(L('label.selectExtensionReferenceFolder'));
            if (!selectedFolder) return;
            referenceFolder = selectedFolder;
            folderLabel.text = selectedFolder.fsName;
            if (okBtn) okBtn.enabled = true;
            safeRelayout(extdialog);
        };

        var radios = [];
        var radioRow = extPanel.add("group");
        radioRow.orientation = "row";
        radioRow.alignment = ["fill", "top"];
        radioRow.alignChildren = ["fill", "top"];

        function addRadioColumn(parent) {
            var col = parent.add("group");
            col.orientation = "column";
            col.alignment = ["fill", "top"];
            col.alignChildren = ["left", "top"];
            col.preferredSize.width = 110;
            return col;
        }

        function addRadio(parent, label, ext, alt, isDefault) {
            var radioBtn = parent.add("radiobutton", undefined, label);
            if (isDefault) radioBtn.value = true;
            var entry = { ui: radioBtn, ext: ext, alt: alt || "" };
            radios.push(entry);
            radioBtn.onClick = function () {
                for (var i = 0; i < radios.length; i++) {
                    if (radios[i].ui !== radioBtn) radios[i].ui.value = false;
                }
                radioBtn.value = true;
            };
        }

        var radioColumns = [
            [
                { label: "png", ext: ".png" },
                { label: "jpg / jpeg", ext: ".jpg", alt: ".jpeg" },
                { label: "psd", ext: ".psd", isDefault: true }
            ],
            [
                { label: "tiff", ext: ".tif", alt: ".tiff" },
                { label: "webp", ext: ".webp" },
                { label: "avif", ext: ".avif" }
            ],
            [
                { label: "gif", ext: ".gif" },
                { label: "ai", ext: ".ai" },
                { label: "pdf", ext: ".pdf" }
            ]
        ];
        for (var colIdx = 0; colIdx < radioColumns.length; colIdx++) {
            var column = addRadioColumn(radioRow);
            for (var entryIdx = 0; entryIdx < radioColumns[colIdx].length; entryIdx++) {
                var spec = radioColumns[colIdx][entryIdx];
                addRadio(column, spec.label, spec.ext, spec.alt || null, !!spec.isDefault);
            }
        }

        function findSelectedRadioSpec() {
            for (var radioIdx = 0; radioIdx < radios.length; radioIdx++) {
                if (radios[radioIdx].ui.value) return radios[radioIdx];
            }
            return null;
        }

        var btnRow = extdialog.add("group");
        setupRow(btnRow, ["right", "top"]);
        var cancelBtn = btnRow.add("button", undefined, L('button.cancel'), { name: "cancel" });
        okBtn = btnRow.add("button", undefined, "OK", { name: "ok" });
        okBtn.enabled = false;
        cancelBtn.onClick = function () { extdialog.close(0); };
        okBtn.onClick = function () { if (!referenceFolder) return; extdialog.close(1); };

        if (extdialog.show() !== 1) return null;
        var selected = findSelectedRadioSpec();
        if (!selected) return null;
        return { referenceFolder: referenceFolder, primaryExt: selected.ext, fallbackExt: selected.alt };
    }

    // クリップグループ内削除の確認＋モード選択。戻り値: 'image' / 'group' / null
    function askDeleteModeWithConfirm(targetCount) {
        var deleteDialog = new Window("dialog", L('dialog.clipGroupDelete'));
        setupWindow(deleteDialog);

        var msgText = L('message.confirmDeleteLinks') + "\n" +
            kvLine('label.target', targetCount, 'label.items') + "\n\n" +
            L('message.clipGroupDelete');
        var msg = deleteDialog.add("statictext", undefined, msgText, { multiline: true });
        msg.preferredSize.width = 360;

        var btnRow = deleteDialog.add("group");
        setupRow(btnRow, ["right", "center"]);
        var cancelBtn = btnRow.add("button", undefined, L('button.cancel'), { name: "cancel" });
        var imageOnlyBtn = btnRow.add("button", undefined, L('button.deleteImageOnly'));
        var withGroupBtn = btnRow.add("button", undefined, L('button.deleteWithClipGroup'), { name: "ok" });

        cancelBtn.onClick = function () { deleteDialog.close(0); };
        imageOnlyBtn.onClick = function () { deleteDialog.close(1); };
        withGroupBtn.onClick = function () { deleteDialog.close(2); };

        var dialogResult = deleteDialog.show();
        if (dialogResult === 1) return 'image';
        if (dialogResult === 2) return 'group';
        return null;
    }

    // =========================================
    // パレット構築 / Build palette
    // =========================================

    function showPalette() {
        var MAIN_LISTBOX_SIZE = [450, 190];
        var FOLDER_LISTBOX_SIZE = [450, 120];

        var palette = new Window("palette", L('dialog.main') + " " + SCRIPT_VERSION, undefined, { resizeable: false });
        setupWindow(palette);
        palette.preferredSize.width = 450;

        // 多重起動防止：既存パレットがあれば閉じる
        ignoreError(function () {
            if ($.global.__LIM_paletteWindow && $.global.__LIM_paletteWindow !== palette) {
                $.global.__LIM_paletteWindow.close();
            }
        });
        $.global.__LIM_paletteWindow = palette;

        var sourceEntries = uniqueFileEntries;
        var filteredEntries = sourceEntries;
        var suppressCanvasOnce = false;
        var pendingSelectionItemIndex = -1;

        // ボタン操作の再入防止。BridgeTalk の同期送信中も ScriptUI のイベントは回るため、
        // 委譲を伴うハンドラはこれで包んで二重実行を防ぐ（内部の isBusy とは別レイヤー）。
        var isActionRunning = false;
        function guardAction(handler) {
            return function () {
                if (isActionRunning) return;
                isActionRunning = true;
                try {
                    return handler.apply(this, arguments);
                } finally {
                    isActionRunning = false;
                }
            };
        }

        var topRow = palette.add("group");
        topRow.orientation = "row";
        topRow.alignChildren = ["fill", "top"];
        topRow.spacing = COLUMN_SPACING;

        var leftCol = topRow.add("group");
        leftCol.orientation = "column";
        leftCol.alignChildren = ["fill", "top"];

        var sortDropdown, ascRadio, descRadio;
        var currentVisibleSpecs = [];

        function createSortPanel(parent) {
            var sortPanel = parent.add("panel", undefined, L('panel.sort'));
            setupPanel(sortPanel, 6);
            var sortKeyRow = sortPanel.add("group");
            sortKeyRow.orientation = "row";
            sortKeyRow.alignChildren = ["left", "center"];
            sortKeyRow.add("statictext", undefined, labelText('sort.by'));
            sortDropdown = sortKeyRow.add("dropdownlist", undefined, []);
            var orderRow = sortPanel.add("group");
            orderRow.orientation = "row";
            orderRow.alignChildren = ["left", "center"];
            ascRadio = orderRow.add("radiobutton", undefined, L('sort.asc'));
            descRadio = orderRow.add("radiobutton", undefined, L('sort.desc'));
            descRadio.value = true;
        }
        createSortPanel(leftCol);

        var optPanel = leftCol.add("panel", undefined, L('panel.sameFile'));
        setupPanel(optPanel, 6);
        var dedupCheck = optPanel.add("checkbox", undefined, L('checkbox.dedup'));
        dedupCheck.value = true;
        dedupCheck.helpTip = (currentLanguage === 'ja')
            ? "ON：同じリンクファイルを1行にまとめます。\nOFF：配置ごとに個別表示します。"
            : "ON: Group same linked files into one row.\nOFF: Each placement is listed separately.";
        var countColCheck = optPanel.add("checkbox", undefined, L('checkbox.displayFileCount'));
        countColCheck.value = true;

        // 選択時にズーム表示：「同一ファイル」パネルの直下にグループで配置（左右中央揃え）
        var showOnCanvasGroup = leftCol.add("group");
        showOnCanvasGroup.orientation = "row";
        showOnCanvasGroup.alignChildren = ["center", "center"];
        showOnCanvasGroup.alignment = ["fill", "top"];
        var showOnCanvasCheck = showOnCanvasGroup.add("checkbox", undefined, L('checkbox.showOnCanvas'));
        showOnCanvasCheck.value = true;

        var otherPanel = topRow.add("group");
        otherPanel.orientation = "column";
        otherPanel.alignChildren = ["fill", "top"];

        var otherTopRow = otherPanel.add("group");
        otherTopRow.orientation = "row";
        otherTopRow.alignChildren = ["fill", "top"];

        var sizeColCheck, unitCheck, dimScalePpiCheck, colorSpaceColCheck;

        function createDisplayOptionsPanel(parent) {
            var optionPanel = parent.add("panel", undefined, L('panel.displayColumn'));
            setupPanel(optionPanel, 6);
            var sizeRow = optionPanel.add("group");
            sizeRow.orientation = "row";
            sizeRow.alignChildren = ["left", "center"];
            sizeColCheck = sizeRow.add("checkbox", undefined, L('checkbox.displaySize'));
            unitCheck = optionPanel.add("checkbox", undefined, L('checkbox.unit'));
            dimScalePpiCheck = optionPanel.add("checkbox", undefined, L('checkbox.displayDimScalePpi'));
            colorSpaceColCheck = optionPanel.add("checkbox", undefined, L('checkbox.displayColorSpace'));
            sizeColCheck.value = false;
            unitCheck.value = true;
            unitCheck.enabled = sizeColCheck.value;
            dimScalePpiCheck.value = false;
            colorSpaceColCheck.value = false;
        }
        createDisplayOptionsPanel(otherTopRow);

        var okCheck, brokenCheck, updateCheck, embeddedCheck;

        function createStatusFilterPanel(parent) {
            var filterPanel = parent.add("panel", undefined, L('panel.status'));
            setupPanel(filterPanel, 6);
            var statusGroup = filterPanel.add("group");
            statusGroup.orientation = "column";
            statusGroup.alignChildren = ["left", "top"];
            statusGroup.spacing = 6;
            okCheck = statusGroup.add("checkbox", undefined, L('checkbox.filterOk'));
            brokenCheck = statusGroup.add("checkbox", undefined, L('checkbox.filterBroken'));
            updateCheck = statusGroup.add("checkbox", undefined, L('checkbox.filterUpdate'));
            embeddedCheck = statusGroup.add("checkbox", undefined, L('checkbox.filterEmbedded'));
            okCheck.value = true;
            brokenCheck.value = true;
            updateCheck.value = true;
            embeddedCheck.value = true;
        }
        createStatusFilterPanel(otherTopRow);

        var abFilterDropdown, abPrevBtn, abNextBtn;
        // 再構築中の onChange 発火を抑止（selection への代入でも onChange が走るため）
        var suppressArtboardChange = false;

        // アートボードドロップダウンの項目・ディム状態を現在のデータから作り直す。
        // 選べる項目を選んでいた場合は選択を維持し、そうでなければ「すべて」に戻す。
        function populateArtboardDropdown() {
            var prevIndex = abFilterDropdown.selection ? abFilterDropdown.selection.index : 0;
            var artboardsWithImages = {};
            for (var entryIdx = 0; entryIdx < allPlacementEntries.length; entryIdx++) {
                var abNum = allPlacementEntries[entryIdx].artboardNum;
                if (abNum !== null) artboardsWithImages[abNum] = true;
            }
            var artboardSep = (currentLanguage === 'ja') ? '：' : ': ';

            suppressArtboardChange = true;
            try {
                abFilterDropdown.removeAll();
                abFilterDropdown.add("item", L('label.artboardAll'));
                for (var artboardIndex = 0; artboardIndex < artboardNames.length; artboardIndex++) {
                    var artboardName = artboardNames[artboardIndex] || "";
                    abFilterDropdown.add("item", (artboardIndex + 1) + artboardSep + (artboardName || L('label.artboardFallback') + (artboardIndex + 1)));
                }

                var enabledCount = 0;
                for (var di = 0; di < artboardNames.length; di++) {
                    var isEnabled = !!artboardsWithImages[di + 1];
                    abFilterDropdown.items[di + 1].enabled = isEnabled;
                    if (isEnabled) enabledCount++;
                }

                var keepPrev = (prevIndex > 0 && prevIndex < abFilterDropdown.items.length && abFilterDropdown.items[prevIndex].enabled);
                abFilterDropdown.selection = keepPrev ? prevIndex : 0;
                abPrevBtn.enabled = (enabledCount > 1);
                abNextBtn.enabled = (enabledCount > 1);
            } finally {
                suppressArtboardChange = false;
            }
        }

        function createArtboardFilterPanel(parent) {
            var abPanel = parent.add("panel", undefined, L('panel.artboard'));
            abPanel.orientation = "row";
            abPanel.alignChildren = ["left", "center"];
            abPanel.alignment = ["fill", "top"];
            abPanel.margins = PANEL_MARGINS;

            abFilterDropdown = abPanel.add("dropdownlist", undefined, []);
            abFilterDropdown.preferredSize.width = 200;

            abPrevBtn = abPanel.add("button", undefined, "");
            abPrevBtn.preferredSize = [22, 22];
            abPrevBtn.helpTip = L('label.prevArtboardTip');
            attachArrowDraw(abPrevBtn, -1);
            abNextBtn = abPanel.add("button", undefined, "");
            abNextBtn.preferredSize = [22, 22];
            abNextBtn.helpTip = L('label.nextArtboardTip');
            attachArrowDraw(abNextBtn, 1);

            populateArtboardDropdown();
        }
        createArtboardFilterPanel(otherPanel);

        var listHolder = palette.add("group");
        listHolder.orientation = "column";
        listHolder.alignChildren = ["fill", "top"];

        var listBox = null;
        var selectedFilePath = "";
        var selectedEntry = null;
        var restoreListSelectionByItemIndex = null;

        function getColumnSpec() {
            var cols = [];
            cols.push({ key: "statusIcon", title: "", width: 30 });
            cols.push({ key: "fileName", title: L('column.fileName'), width: 210 });
            if (sizeColCheck.value) {
                cols.push({ key: "fileSize", title: unitCheck.value ? L('column.fileSizeMb') : L('column.fileSize'), width: 65 });
            }
            if (countColCheck.value) {
                cols.push({ key: "fileCount", title: L('column.fileCount'), width: 45 });
            }
            if (dimScalePpiCheck.value) {
                cols.push({ key: "widthText", title: L('column.widthMm'), width: 60 });
                cols.push({ key: "heightText", title: L('column.heightMm'), width: 60 });
                cols.push({ key: "scaleText", title: L('column.scale'), width: 60 });
                cols.push({ key: "ppiText", title: L('column.ppi'), width: 50 });
            }
            if (colorSpaceColCheck.value) {
                cols.push({ key: "colorSpace", title: L('column.colorSpace'), width: 160 });
            }
            var shouldShowArtboardColumn = !abFilterDropdown.selection || abFilterDropdown.selection.index === 0;
            if (shouldShowArtboardColumn) {
                cols.push({ key: "artboards", title: L('column.artboards'), width: 70 });
            }
            return cols;
        }

        function createListBox() {
            if (listBox) {
                ignoreError(function () { listHolder.remove(listBox); });
            }
            var columns = getColumnSpec();
            var titles = [], widths = [];
            for (var colIdx = 0; colIdx < columns.length; colIdx++) {
                titles.push(columns[colIdx].title);
                widths.push(columns[colIdx].width);
            }
            listBox = listHolder.add("listbox", undefined, [], {
                numberOfColumns: titles.length,
                showHeaders: true,
                columnTitles: titles,
                columnWidths: widths,
                multiselect: false
            });
            listBox.preferredSize = MAIN_LISTBOX_SIZE;
            listBox.alignment = ["fill", "fill"];
        }

        function getStatusSortValue(statusCode) {
            if (statusCode === "broken") return 1;
            if (statusCode === "update") return 2;
            if (statusCode === "ok") return 3;
            if (statusCode === "embedded") return 4;
            return 9;
        }
        var SORT_SPECS = [
            { key: 'fileName', labelKey: 'sort.fileName', preferDesc: false, getValue: function (i) { return i.fileName; }, isVisible: function () { return true; } },
            { key: 'fileSize', labelKey: 'sort.fileSize', preferDesc: true, getValue: function (i) { return i.fileSizeBytes; }, isVisible: function () { return sizeColCheck.value; } },
            { key: 'fileCount', labelKey: 'sort.fileCount', preferDesc: true, getValue: function (i) { return i.fileCount; }, isVisible: function () { return countColCheck.value; } },
            { key: 'artboard', labelKey: 'sort.artboard', preferDesc: false, getValue: function (i) { return (i.artboardNum === null) ? null : i.artboardNum; }, isVisible: function () { return !abFilterDropdown.selection || abFilterDropdown.selection.index === 0; } },
            { key: 'width', labelKey: 'sort.width', preferDesc: true, getValue: function (i) { return i.widthMm; }, isVisible: function () { return dimScalePpiCheck.value; } },
            { key: 'height', labelKey: 'sort.height', preferDesc: true, getValue: function (i) { return i.heightMm; }, isVisible: function () { return dimScalePpiCheck.value; } },
            { key: 'scale', labelKey: 'sort.scale', preferDesc: true, getValue: function (i) { return i.scalePct; }, isVisible: function () { return dimScalePpiCheck.value; } },
            { key: 'ppi', labelKey: 'sort.ppi', preferDesc: true, getValue: function (i) { return i.ppi; }, isVisible: function () { return dimScalePpiCheck.value; } },
            { key: 'status', labelKey: 'sort.status', preferDesc: false, getValue: function (i) { return getStatusSortValue(i.statusCode); }, isVisible: function () { return true; } },
            { key: 'colorSpace', labelKey: 'sort.colorSpace', preferDesc: false, getValue: function (i) { return i.colorSpace || ""; }, isVisible: function () { return colorSpaceColCheck.value; } }
        ];

        function rebuildSortDropdown() {
            var prevKey = null;
            if (sortDropdown.selection && currentVisibleSpecs.length > 0) {
                var ps = currentVisibleSpecs[sortDropdown.selection.index];
                if (ps) prevKey = ps.key;
            }
            currentVisibleSpecs = [];
            for (var si = 0; si < SORT_SPECS.length; si++) {
                if (SORT_SPECS[si].isVisible()) currentVisibleSpecs.push(SORT_SPECS[si]);
            }
            sortDropdown.removeAll();
            var nextIdx = 0;
            for (var vi = 0; vi < currentVisibleSpecs.length; vi++) {
                sortDropdown.add("item", L(currentVisibleSpecs[vi].labelKey));
                if (prevKey && currentVisibleSpecs[vi].key === prevKey) nextIdx = vi;
            }
            if (currentVisibleSpecs.length > 0) sortDropdown.selection = nextIdx;
        }

        function getSortValue(info, sortBy) {
            var spec = currentVisibleSpecs[sortBy];
            if (!spec) return info.fileName;
            return spec.getValue(info);
        }

        function isMissing(v) {
            return v === null || v === undefined || (typeof v === "number" && isNaN(v));
        }

        function rebuildList() {
            var selectionItemIndexToRestore = pendingSelectionItemIndex;
            pendingSelectionItemIndex = -1;
            if (selectionItemIndexToRestore < 0 && selectedEntry && selectedEntry.itemIndices && selectedEntry.itemIndices.length > 0) {
                selectionItemIndexToRestore = selectedEntry.itemIndices[0];
            }

            var sortBy = sortDropdown.selection ? sortDropdown.selection.index : 0;
            var desc = descRadio.value;

            var selectedArtboardIndex = abFilterDropdown.selection ? abFilterDropdown.selection.index : 0;
            var hasArtboardFilter = selectedArtboardIndex > 0;
            var targetArtboardNumber = selectedArtboardIndex;
            var allowOk = okCheck.value;
            var allowBroken = brokenCheck.value;
            var allowUpdate = updateCheck.value;
            var allowEmbedded = embeddedCheck.value;
            var hasStatusFlt = !(allowOk && allowBroken && allowUpdate && allowEmbedded);

            if (hasStatusFlt || hasArtboardFilter) {
                filteredEntries = [];
                for (var sourceIdx = 0; sourceIdx < sourceEntries.length; sourceIdx++) {
                    var entry = sourceEntries[sourceIdx];
                    if (hasStatusFlt) {
                        var statusCode = entry.statusCode;
                        if (statusCode === "ok" && !allowOk) continue;
                        if (statusCode === "broken" && !allowBroken) continue;
                        if (statusCode === "update" && !allowUpdate) continue;
                        if (statusCode === "embedded" && !allowEmbedded) continue;
                    }
                    if (hasArtboardFilter) {
                        var matchesArtboard = entry.artboardSet
                            ? !!entry.artboardSet[targetArtboardNumber]
                            : (entry.artboardNum === targetArtboardNumber);
                        if (!matchesArtboard) continue;
                    }
                    filteredEntries.push(entry);
                }
            } else {
                filteredEntries = sourceEntries.slice();
            }

            filteredEntries.sort(function (a, b) {
                var valA = getSortValue(a, sortBy);
                var valB = getSortValue(b, sortBy);
                var aMissing = isMissing(valA), bMissing = isMissing(valB);
                if (aMissing && bMissing) return a.index - b.index;
                if (aMissing) return 1;
                if (bMissing) return -1;
                if (valA < valB) return desc ? 1 : -1;
                if (valA > valB) return desc ? -1 : 1;
                return a.index - b.index;
            });

            var columns = getColumnSpec();
            var useUnifiedSize = unitCheck.value;
            listBox.removeAll();
            for (var rowIdx = 0; rowIdx < filteredEntries.length; rowIdx++) {
                var info = filteredEntries[rowIdx];
                var firstKey = columns[0].key;
                var firstText = info[firstKey];
                var row = listBox.add("item", (firstText === undefined || firstText === null) ? "" : String(firstText));
                for (var colIdx = 1; colIdx < columns.length; colIdx++) {
                    var key = columns[colIdx].key;
                    var cellText;
                    if (key === "fileSize") {
                        cellText = useUnifiedSize ? formatFileSize(info.fileSizeBytes) : formatFileSizeAuto(info.fileSizeBytes);
                    } else if (key === "fileCount") {
                        cellText = String(info.fileCount);
                    } else {
                        cellText = info[key];
                    }
                    row.subItems[colIdx - 1].text = (cellText === undefined || cellText === null) ? "" : cellText;
                }
            }

            if (restoreListSelectionByItemIndex) {
                restoreListSelectionByItemIndex(selectionItemIndexToRestore);
            }
        }

        function onSortChange() {
            var idx = sortDropdown.selection ? sortDropdown.selection.index : 0;
            var spec = currentVisibleSpecs[idx];
            if (spec && spec.preferDesc) {
                descRadio.value = true;
                ascRadio.value = false;
            }
            rebuildList();
        }

        function onDedupClick() {
            pendingSelectionItemIndex = (selectedEntry && selectedEntry.itemIndices && selectedEntry.itemIndices.length > 0)
                ? selectedEntry.itemIndices[0]
                : -1;
            sourceEntries = dedupCheck.value ? uniqueFileEntries : allPlacementEntries;
            rebuildList();
        }

        function applyArtboardFilter() {
            if (suppressArtboardChange) return;
            var selectedArtboardIndex = abFilterDropdown.selection ? abFilterDropdown.selection.index : 0;
            if (selectedArtboardIndex > 0) {
                delegateFitArtboard(selectedArtboardIndex - 1);
            }
            createListBox();
            bindListBoxEvents();
            rebuildSortDropdown();
            palette.layout.layout(true);
            rebuildList();
        }

        function stepArtboard(direction) {
            var itemCount = abFilterDropdown.items.length;
            if (itemCount <= 1) return;
            var currentIdx = abFilterDropdown.selection ? abFilterDropdown.selection.index : 0;
            var next = currentIdx;
            for (var step = 0; step < itemCount; step++) {
                next = next + direction;
                if (next < 1) next = itemCount - 1;
                else if (next >= itemCount) next = 1;
                if (abFilterDropdown.items[next].enabled) {
                    abFilterDropdown.selection = next;
                    applyArtboardFilter();
                    return;
                }
            }
        }

        function recreateListBoxAndRebuildList() {
            createListBox();
            bindListBoxEvents();
            rebuildSortDropdown();
            palette.layout.layout(true);
            rebuildList();
        }

        function onSizeColClick() {
            unitCheck.enabled = sizeColCheck.value;
            recreateListBoxAndRebuildList();
        }

        var pathPanel, pathStaticText, fullPathCheck, dropboxCheck, fileNameCheck;

        function createPathPanel(parent) {
            pathPanel = parent.add("panel", undefined, L('panel.path'));
            setupPanel(pathPanel);

            var pathRow = pathPanel.add("group");
            pathRow.orientation = "row";
            pathRow.alignChildren = ["fill", "center"];

            pathStaticText = pathRow.add("statictext", undefined, L('label.pathPlaceholder'), { multiline: true });
            pathStaticText.alignment = ["fill", "fill"];
            pathStaticText.preferredSize = [450, 20];
            pathStaticText.helpTip = L('label.pathHelpTip');

            var pathOptRow = pathPanel.add("group");
            pathOptRow.orientation = "row";
            pathOptRow.alignment = "fill";
            pathOptRow.alignChildren = ["fill", "center"];

            var pathOptLeft = pathOptRow.add("group");
            pathOptLeft.orientation = "row";
            pathOptLeft.alignChildren = ["left", "center"];
            pathOptLeft.alignment = ["left", "center"];
            fullPathCheck = pathOptLeft.add("checkbox", undefined, L('checkbox.fullPath'));
            if (DROPBOX_PREFIX) {
                dropboxCheck = pathOptLeft.add("checkbox", undefined, L('checkbox.dropbox'));
                dropboxCheck.value = true;
            } else {
                dropboxCheck = { value: false, enabled: false };
            }
            fileNameCheck = pathOptLeft.add("checkbox", undefined, L('checkbox.fileName'));
            fullPathCheck.value = false;
            fileNameCheck.value = false;

            var pathOptSpacer = pathOptRow.add("group");
            pathOptSpacer.alignment = ["fill", "fill"];

            var pathOptRight = pathOptRow.add("group");
            pathOptRight.orientation = "row";
            pathOptRight.alignChildren = ["right", "center"];
            pathOptRight.alignment = ["right", "center"];
        }
        createPathPanel(palette);

        function requireSelectedEntry(handler) {
            return function () {
                if (!selectedEntry) {
                    setStatus(L('message.selectItem'));
                    return;
                }
                return handler.apply(this, arguments);
            };
        }

        function buildDisplayedPath(absPath) {
            if (!absPath || absPath === "---") return absPath;
            var displayPath = fileNameCheck.value ? absPath : toFolderOnly(absPath);
            if (fullPathCheck.value) return displayPath;
            return formatDisplayPath(displayPath, true, dropboxCheck.value);
        }

        function updateSelectedEntryDisplay(info) {
            selectedEntry = info;
            selectedFilePath = info.filePath;
            pathStaticText.text = buildDisplayedPath(info.filePath);
            updateActionButtonStates();
            highlightFolderFor(info.filePath);
        }

        function handleListSelectionChange() {
            if (listBox.selection === null) return;
            var info = filteredEntries[listBox.selection.index];
            updateSelectedEntryDisplay(info);
            if (suppressCanvasOnce) {
                suppressCanvasOnce = false;
                return;
            }
            if (showOnCanvasCheck.value) selectPlacedItemsOnCanvas(info);
        }

        function bindListBoxEvents() {
            if (!listBox) return;
            listBox.onChange = handleListSelectionChange;
        }

        restoreListSelectionByItemIndex = function (itemIndexToRestore) {
            if (itemIndexToRestore < 0 || !filteredEntries || filteredEntries.length === 0) return;
            for (var i = 0; i < filteredEntries.length; i++) {
                var entry = filteredEntries[i];
                if (!entry || !entry.itemIndices) continue;
                for (var j = 0; j < entry.itemIndices.length; j++) {
                    if (entry.itemIndices[j] === itemIndexToRestore) {
                        suppressCanvasOnce = true;
                        updateSelectedEntryDisplay(entry);
                        try {
                            var targetItem = listBox.items[i];
                            listBox.selection = targetItem;
                            if (typeof listBox.revealItem === "function") {
                                listBox.revealItem(targetItem);
                            }
                        } catch (e) {
                            ignoreError(function () { listBox.selection = i; });
                        }
                        return;
                    }
                }
            }
        };

        function updatePathDisplay() {
            if (!selectedFilePath) return;
            pathStaticText.text = buildDisplayedPath(selectedFilePath);
        }

        function onPathOptionChange() {
            updatePathDisplay();
            populateFoldersList();
        }

        createListBox();
        bindListBoxEvents();
        rebuildSortDropdown();
        for (var initSi = 0; initSi < currentVisibleSpecs.length; initSi++) {
            if (currentVisibleSpecs[initSi].key === 'fileCount') {
                sortDropdown.selection = initSi;
                break;
            }
        }
        rebuildList();

        function updateFullPathEnable() {
            fullPathCheck.enabled = !dropboxCheck.value;
            if (!fullPathCheck.enabled) fullPathCheck.value = false;
        }

        function onDropboxClick() {
            updateFullPathEnable();
            onPathOptionChange();
        }

        function onFileNameClick() {
            updatePathDisplay();
            pathPanel.layout.layout(true);
            palette.layout.layout(true);
        }
        updateFullPathEnable();

        function bindPaletteEvents() {
            sortDropdown.onChange = onSortChange;
            ascRadio.onClick = rebuildList;
            descRadio.onClick = rebuildList;
            dedupCheck.onClick = onDedupClick;
            okCheck.onClick = rebuildList;
            brokenCheck.onClick = rebuildList;
            updateCheck.onClick = rebuildList;
            embeddedCheck.onClick = rebuildList;
            abFilterDropdown.onChange = applyArtboardFilter;
            abPrevBtn.onClick = function () { stepArtboard(-1); };
            abNextBtn.onClick = function () { stepArtboard(1); };
            unitCheck.onClick = recreateListBoxAndRebuildList;
            sizeColCheck.onClick = onSizeColClick;
            countColCheck.onClick = recreateListBoxAndRebuildList;
            dimScalePpiCheck.onClick = recreateListBoxAndRebuildList;
            colorSpaceColCheck.onClick = recreateListBoxAndRebuildList;
            fullPathCheck.onClick = onPathOptionChange;
            dropboxCheck.onClick = onDropboxClick;
            fileNameCheck.onClick = onFileNameClick;
        }
        bindPaletteEvents();

        // エントリが埋め込み画像（rasterItems 由来）かどうか
        function isEmbeddedEntry(entry) {
            return !!(entry && entry.kind === "raster");
        }

        // --- アクションボタン行（M1 では削除/リネーム/再リンク/コピーはスタブ）---
        var actionBtnRow = pathPanel.add("group");
        actionBtnRow.orientation = "row";
        actionBtnRow.alignment = ["fill", "top"];
        actionBtnRow.alignChildren = ["fill", "center"];

        var actionBtnLeft = actionBtnRow.add("group");
        actionBtnLeft.orientation = "row";
        actionBtnLeft.alignChildren = ["left", "center"];
        actionBtnLeft.alignment = ["left", "center"];

        var actionBtnSpacer = actionBtnRow.add("group");
        actionBtnSpacer.alignment = ["fill", "fill"];

        var actionBtnRight = actionBtnRow.add("group");
        actionBtnRight.orientation = "row";
        actionBtnRight.alignChildren = ["right", "center"];
        actionBtnRight.alignment = ["right", "center"];

        var openFileBtn = actionBtnRight.add("button", undefined, L('button.open'));
        openFileBtn.preferredSize = [50, 24];
        openFileBtn.onClick = requireSelectedEntry(function () {
            var absPath = selectedEntry.filePath;
            if (!absPath || absPath === "---") { setStatus(L('message.noValidPath')); return; }
            var fileToOpen = new File(absPath);
            if (!fileToOpen.exists) { setStatus(L('message.linkFileNotFound') + absPath); return; }
            fileToOpen.execute();
        });

        var deleteLinkBtn = actionBtnRight.add("button", undefined, L('button.delete'));
        deleteLinkBtn.preferredSize = [50, 24];
        deleteLinkBtn.onClick = requireSelectedEntry(guardAction(handleDeleteSelected));

        var renameLinkBtn = actionBtnRight.add("button", undefined, L('button.rename'));
        renameLinkBtn.preferredSize = [70, 24];
        renameLinkBtn.onClick = requireSelectedEntry(guardAction(handleRenameSelected));

        var copyFileNameBtn = actionBtnLeft.add("button", undefined, L('button.copyFileName'));
        copyFileNameBtn.onClick = requireSelectedEntry(guardAction(handleCopyFileName));

        var reloadOneBtn = actionBtnRight.add("button", undefined, L('button.relinkSelected'));
        reloadOneBtn.preferredSize = [94, 24];
        reloadOneBtn.onClick = requireSelectedEntry(guardAction(handleRelinkSelected));

        // --- 埋め込み／埋め込み解除行（アクションボタン行と同じく、対象はリストで選択中の画像）---
        var embedRow = pathPanel.add("group");
        embedRow.orientation = "row";
        embedRow.alignment = "fill";
        embedRow.alignChildren = ["left", "center"];

        var embedBtn = embedRow.add("button", undefined, L('button.embed'));
        embedBtn.helpTip = (currentLanguage === 'ja')
            ? "選択中のリンクを埋め込み画像に変換します。\nPSD はクリップグループ外のときグループ解除します。"
            : "Embed the selected link.\nPSD files are ungrouped unless they are inside a clip group.";
        embedBtn.onClick = requireSelectedEntry(guardAction(handleEmbedSelected));

        var unembedBtn = embedRow.add("button", undefined, L('button.unembed'));
        unembedBtn.helpTip = (currentLanguage === 'ja')
            ? "選択中の埋め込み画像をリンク画像に戻します。\n元ファイルが不明な場合は「Links」フォルダーへ PSD を書き出してリンクします。"
            : "Turn the selected embedded image back into a link.\nWhen the original file is unknown, a PSD is exported into the \"Links\" folder.";
        unembedBtn.onClick = requireSelectedEntry(guardAction(handleUnembedSelected));

        var collectAfterRelinkCheck = embedRow.add("checkbox", undefined, L('checkbox.collectAfterRelink'));
        collectAfterRelinkCheck.value = true;
        collectAfterRelinkCheck.helpTip = (currentLanguage === 'ja')
            ? "埋め込み解除後、リンク先をドキュメントと同階層の「Links」フォルダーへコピーします。"
            : "After unembedding, copy the linked file into the \"Links\" folder next to the document.";

        function updateRelinkButtonLabel() {
            var placementCount = (selectedEntry && selectedEntry.itemIndices) ? selectedEntry.itemIndices.length : 0;
            var useBatchLabel = dedupCheck.value && placementCount > 1;
            var nextLabel = useBatchLabel ? L('button.relinkAll') : L('button.relinkSelected');
            var nextWidth = 94, nextHeight = 24;
            reloadOneBtn.text = nextLabel;
            reloadOneBtn.preferredSize = [nextWidth, nextHeight];
            reloadOneBtn.size = [nextWidth, nextHeight];
            ignoreError(function () {
                reloadOneBtn.bounds = [reloadOneBtn.bounds[0], reloadOneBtn.bounds[1], reloadOneBtn.bounds[0] + nextWidth, reloadOneBtn.bounds[1] + nextHeight];
            });
            safeRelayout(actionBtnRight, actionBtnRow, pathPanel, palette);
        }

        function updateActionButtonStates() {
            var hasSelection = (selectedEntry !== null);
            // 埋め込み画像はリンク元ファイルを持たないため、ファイル操作系は対象外
            var canUseLinkFile = hasSelection && !isEmbeddedEntry(selectedEntry);
            var canUnembed = hasSelection && isEmbeddedEntry(selectedEntry);
            embedBtn.enabled = canUseLinkFile;
            unembedBtn.enabled = canUnembed;
            // ［再リンク後に収集］は［埋め込み解除］のオプションなので、使えないときは一緒にディムにする
            collectAfterRelinkCheck.enabled = canUnembed;
            reloadOneBtn.enabled = canUseLinkFile;
            renameLinkBtn.enabled = canUseLinkFile;
            openFileBtn.enabled = canUseLinkFile;
            deleteLinkBtn.enabled = hasSelection;
            copyFileNameBtn.enabled = hasSelection;
            updateRelinkButtonLabel();
        }
        updateActionButtonStates();

        // --- リンクフォルダ一覧 ---
        var linkedFolderPaths = [];
        function rebuildFolderList() {
            var linkedFolderMap = {};
            linkedFolderPaths = [];
            for (var fi = 0; fi < allPlacementEntries.length; fi++) {
                var filePath = allPlacementEntries[fi].filePath;
                if (filePath && filePath !== "---") {
                    var linkedFolderPath = pathParent(filePath);
                    if (linkedFolderPath && !linkedFolderMap[linkedFolderPath]) {
                        linkedFolderMap[linkedFolderPath] = true;
                        linkedFolderPaths.push(linkedFolderPath);
                    }
                }
            }
            linkedFolderPaths.sort();
        }
        rebuildFolderList();

        var folderCountLabel = pathPanel.add("statictext", undefined, L('label.linkedFolders') + " (" + withUnit(linkedFolderPaths.length, 'label.items') + ")");
        var foldersListBox = pathPanel.add("listbox", undefined, [], { multiselect: false });
        foldersListBox.preferredSize = FOLDER_LISTBOX_SIZE;
        foldersListBox.alignment = ["fill", "fill"];

        function populateFoldersList() {
            foldersListBox.removeAll();
            for (var ii = 0; ii < linkedFolderPaths.length; ii++) {
                foldersListBox.add("item", formatDisplayPath(linkedFolderPaths[ii], !fullPathCheck.value, dropboxCheck.value));
            }
        }
        populateFoldersList();

        // ドキュメントから再読込してリストを更新（[更新] ボタン・将来のミューテーション後に使用）
        function refreshFromDoc() {
            var ok = loadData();
            if (ok === null) return; /* 別の委譲が実行中：現在の表示を壊さない */
            updateEmptyStateVisibility();
            populateArtboardDropdown();
            if (!ok) {
                sourceEntries = [];
                filteredEntries = [];
                if (listBox) listBox.removeAll();
                selectedEntry = null;
                selectedFilePath = "";
                pathStaticText.text = L('label.pathPlaceholder');
                rebuildFolderList();
                folderCountLabel.text = L('label.linkedFolders') + " (" + withUnit(linkedFolderPaths.length, 'label.items') + ")";
                populateFoldersList();
                updateActionButtonStates();
                return;
            }
            sourceEntries = dedupCheck.value ? uniqueFileEntries : allPlacementEntries;

            if (selectedEntry) {
                var anchor = (selectedEntry.itemIndices && selectedEntry.itemIndices.length > 0) ? selectedEntry.itemIndices[0] : -1;
                var rematched = null;
                if (anchor >= 0) {
                    for (var srcIdx = 0; srcIdx < sourceEntries.length; srcIdx++) {
                        var candidateEntry = sourceEntries[srcIdx];
                        for (var idxIdx = 0; idxIdx < candidateEntry.itemIndices.length; idxIdx++) {
                            if (candidateEntry.itemIndices[idxIdx] === anchor) { rematched = candidateEntry; break; }
                        }
                        if (rematched) break;
                    }
                }
                selectedEntry = rematched;
                selectedFilePath = rematched ? rematched.filePath : "";
                pathStaticText.text = selectedFilePath ? buildDisplayedPath(selectedFilePath) : L('label.pathPlaceholder');
                updateActionButtonStates();
            }

            rebuildFolderList();
            folderCountLabel.text = L('label.linkedFolders') + " (" + withUnit(linkedFolderPaths.length, 'label.items') + ")";
            populateFoldersList();
            /* アートボード絞り込みが解除されると列構成が変わるためリストごと作り直す */
            recreateListBoxAndRebuildList();
            if (selectedEntry) highlightFolderFor(selectedEntry.filePath);
        }

        foldersListBox.onDoubleClick = function () {
            if (foldersListBox.selection === null) return;
            var idx = foldersListBox.selection.index;
            try {
                var folderPath = linkedFolderPaths[idx];
                var ff = new Folder(folderPath);
                if (ff.exists) ff.execute();
            } catch (e) {
                setStatus(L('message.openFolderFailed') + e.message);
            }
        };

        // --- フォルダ操作行（M1 では再リンク/拡張子変更/収集はスタブ）---
        var folderActionRow = pathPanel.add("group");
        folderActionRow.orientation = "row";
        folderActionRow.alignment = "fill";
        folderActionRow.alignChildren = ["fill", "center"];

        var folderActionLeft = folderActionRow.add("group");
        folderActionLeft.orientation = "row";
        folderActionLeft.alignChildren = ["left", "center"];
        folderActionLeft.alignment = ["left", "center"];

        var openFolderBtn = folderActionLeft.add("button", undefined, L('button.openFolder'));
        openFolderBtn.onClick = function () {
            if (foldersListBox.selection === null) { setStatus(L('message.selectLinkedFolder')); return; }
            try {
                var folderPath = linkedFolderPaths[foldersListBox.selection.index];
                var folderToOpen = new Folder(folderPath);
                if (folderToOpen.exists) folderToOpen.execute();
            } catch (e) {
                setStatus(L('message.openFolderFailed') + e.message);
            }
        };

        var folderActionSpacer = folderActionRow.add("group");
        folderActionSpacer.alignment = ["fill", "fill"];

        var folderActionRight = folderActionRow.add("group");
        folderActionRight.orientation = "row";
        folderActionRight.alignChildren = ["right", "center"];
        folderActionRight.alignment = ["right", "center"];

        var reloadFolderBtn = folderActionLeft.add("button", undefined, L('button.relinkFolder'));
        reloadFolderBtn.onClick = guardAction(handleRelinkFolder);
        reloadFolderBtn.enabled = false;

        var changeExtensionBtn = folderActionLeft.add("button", undefined, L('button.changeExtension'));
        changeExtensionBtn.onClick = guardAction(handleChangeExtension);
        changeExtensionBtn.enabled = false;

        var collectLinksBtn = folderActionRight.add("button", undefined, L('button.collectLinks'));
        collectLinksBtn.onClick = guardAction(handleCollectLinks);

        foldersListBox.onChange = function () {
            var hasSelection = (foldersListBox.selection !== null);
            reloadFolderBtn.enabled = hasSelection;
            openFolderBtn.enabled = hasSelection;
            changeExtensionBtn.enabled = hasSelection;
        };

        function highlightFolderFor(absPath) {
            if (!absPath || absPath === "---") { foldersListBox.selection = null; return; }
            var folder = pathParent(absPath);
            if (!folder) return;
            for (var folderIdx = 0; folderIdx < linkedFolderPaths.length; folderIdx++) {
                if (linkedFolderPaths[folderIdx] === folder) {
                    ignoreError(function () {
                        foldersListBox.selection = foldersListBox.items[folderIdx];
                        if (typeof foldersListBox.revealItem === "function") {
                            foldersListBox.revealItem(foldersListBox.items[folderIdx]);
                        }
                    });
                    return;
                }
            }
            foldersListBox.selection = null;
        }

        function selectPlacedItemsOnCanvas(info) {
            if (!info) return;
            delegateSelect(info.itemIndices);
        }

        // ---- ミューテーション系ハンドラ（File 操作はパレット側、リンク差替/削除は worker 委譲）----

        function handleDeleteSelected() {
            var indices = selectedEntry.itemIndices || [];
            if (indices.length === 0) return;
            var probe = parseWorkerResult(delegate("$.global.__LIM.probeDelete(" + jsIntArray(indices) + ")"));
            if (probe.marker === "NODOC") { setStatus(L('status.noDocument')); return; }
            if (probe.marker !== "CLIP" && probe.marker !== "PLAIN") { setStatus(L('status.loadFailed')); return; }
            var clipMode = 'image';
            if (probe.marker === "CLIP") {
                clipMode = askDeleteModeWithConfirm(indices.length);
                if (clipMode === null) return;
            } else {
                var confirmMessage = L('message.confirmDeleteLinks') + "\n" + kvLine('label.target', indices.length, 'label.items');
                if (!confirm(confirmMessage)) return;
            }
            var res = parseWorkerResult(delegate("$.global.__LIM.del(" + jsIntArray(indices) + "," + jsString(clipMode) + ")"));
            var counts = (res.marker === "OK") ? tryGet(function () { return eval("(" + res.body + ")"); }, { success: 0, failed: 0 }) : { success: 0, failed: 0 };
            selectedEntry = null;
            selectedFilePath = "";
            refreshFromDoc();
            updateActionButtonStates();
            setStatus(L('message.deleteDone') + "／" + kvLine('label.success', counts.success, 'label.items') + "／" + kvLine('label.failed', counts.failed, 'label.items'));
        }

        function handleRenameSelected() {
            var absPath = selectedEntry.filePath;
            if (!absPath || absPath === "---") { setStatus(L('message.noValidPath')); return; }
            var oldFile = new File(absPath);
            if (!oldFile.exists) { setStatus(L('message.linkFileNotFound') + absPath); return; }
            var oldName = getRealFileName(oldFile);
            var oldFolder = oldFile.parent;
            var newName = promptNewFileName(oldName);
            if (newName === null) return;
            if (newName === oldName) { setStatus(L('message.nameUnchanged')); return; }
            if (/[\/\\]/.test(newName)) { setStatus(L('message.invalidFileName')); return; }
            var newFile = new File(oldFolder.fsName + "/" + newName);
            if (!prepareRenameOverwrite(oldFile, newFile)) return;
            var renamed = tryGet(function () { return oldFile.rename(newName); }, false);
            if (!renamed) { setStatus(L('message.renameFailed')); return; }
            var pairs = [];
            var idxs = selectedEntry.itemIndices || [];
            for (var i = 0; i < idxs.length; i++) pairs.push([idxs[i], newFile.fsName]);
            var counts = delegateRelinkPairs(pairs) || { success: 0, failed: 0 };
            refreshFromDoc();
            setStatus(L('message.renameDone') + "：" + oldName + " → " + newName + "／" + kvLine('label.success', counts.success, 'label.items'));
        }

        function handleCopyFileName() {
            var name = selectedEntry.fileName || "";
            var res = parseWorkerResult(delegate("$.global.__LIM.copyText(" + jsString(name) + ")"));
            if (res.marker === "OK") setStatus(L('message.copyFileNameDone') + "：" + name);
            else setStatus(L('message.copyFileNameFailed'));
        }

        function handleRelinkSelected() {
            var idxs = selectedEntry.itemIndices || [];
            if (idxs.length > 1) {
                var confirmMessage = L('message.confirmBatchRelink') + "\n" + kvLine('label.target', idxs.length, 'label.items');
                if (!confirm(confirmMessage)) return;
            }
            var picked = File.openDialog(L('label.selectNewLinkFile'));
            if (!picked) return;
            var pairs = [];
            for (var i = 0; i < idxs.length; i++) pairs.push([idxs[i], picked.fsName]);
            var counts = delegateRelinkPairs(pairs) || { success: 0, failed: 0 };
            refreshFromDoc();
            setStatus(L('message.relinkDone') + "／" + kvLine('label.success', counts.success, 'label.items') + "／" + kvLine('label.failed', counts.failed, 'label.items'));
        }

        function handleEmbedSelected() {
            if (isEmbeddedEntry(selectedEntry)) return; /* 既に埋め込み済み */
            var idxs = selectedEntry.itemIndices || [];
            if (idxs.length === 0) { setStatus(L('message.selectItem')); return; }
            if (idxs.length > 1) {
                var confirmMessage = L('message.confirmBatchEmbed') + "\n" + kvLine('label.target', idxs.length, 'label.items');
                if (!confirm(confirmMessage)) return;
            }
            var counts = delegateEmbed(idxs) || { success: 0, failed: 0 };
            refreshFromDoc();
            setStatus(L('message.embedDone') + "／" + kvLine('label.success', counts.success, 'label.items') + "／" + kvLine('label.failed', counts.failed, 'label.items'));
        }

        function handleUnembedSelected() {
            if (!isEmbeddedEntry(selectedEntry)) return; /* 対象は埋め込み画像のみ */
            var idxs = selectedEntry.itemIndices || [];
            if (idxs.length === 0) { setStatus(L('message.selectItem')); return; }
            var counts = delegateUnembed(idxs, collectAfterRelinkCheck.value);
            if (!counts) { setStatus(L('status.loadFailed')); return; }
            selectedEntry = null;
            selectedFilePath = "";
            refreshFromDoc();
            setStatus(L('message.unembedDone')
                + "／" + kvLine('label.success', counts.success, 'label.items')
                + "／" + kvLine('label.skipped', counts.skipped, 'label.items')
                + "／" + kvLine('label.failed', counts.failed, 'label.items'));
            var details = counts.details || [];
            if (details.length > 0) {
                var lines = [L('message.unembedFailedDetail'), ""];
                for (var i = 0; i < details.length; i++) {
                    lines.push(details[i].name + "：" + unembedReasonText(details[i].code));
                }
                alert(lines.join("\n"));
            }
        }

        function handleRelinkFolder() {
            if (foldersListBox.selection === null) { setStatus(L('message.selectLinkedFolder')); return; }
            var oldFolder = linkedFolderPaths[foldersListBox.selection.index];
            var newFolder = Folder.selectDialog(L('label.selectAltFolder'));
            if (!newFolder) return;
            var pairs = [], total = 0, missing = 0;
            for (var k = 0; k < allPlacementEntries.length; k++) {
                var ent = allPlacementEntries[k];
                if (!ent.filePath || ent.filePath === "---") continue;
                if (pathParent(ent.filePath) === oldFolder) {
                    total++;
                    var nf = new File(newFolder.fsName + "/" + pathBaseName(ent.filePath));
                    if (nf.exists) pairs.push([ent.itemIndex, nf.fsName]);
                    else missing++;
                }
            }
            var counts = delegateRelinkPairs(pairs) || { success: 0, failed: 0 };
            refreshFromDoc();
            setStatus(L('message.relinkDone') + "／" + kvLine('label.target', total, 'label.items') + "／" + kvLine('label.success', counts.success, 'label.items') + "／" + kvLine('label.failed', (counts.failed + missing), 'label.items'));
        }

        function handleChangeExtension() {
            if (foldersListBox.selection === null) { setStatus(L('message.selectLinkedFolder')); return; }
            var sourceFolderPath = normalizeFolderPathForCompare(linkedFolderPaths[foldersListBox.selection.index]);
            var extPrefs = showChangeExtensionDialog();
            if (!extPrefs) return;
            var pairs = [], total = 0, skipped = 0, failed = 0;
            for (var i = 0; i < allPlacementEntries.length; i++) {
                var entry = allPlacementEntries[i];
                if (!entry || !entry.filePath || entry.filePath === "---") continue;
                if (normalizeFolderPathForCompare(pathParent(entry.filePath)) !== sourceFolderPath) continue;
                total++;
                /* fileName は worker 側でデコード済み */
                var sourceFileName = entry.fileName || pathBaseName(entry.filePath);
                var baseName = splitFileName(sourceFileName).base;
                if (!baseName) { failed++; continue; }
                var repl = findReplacementFileByExtension(extPrefs.referenceFolder, baseName, extPrefs.primaryExt, extPrefs.fallbackExt);
                if (!repl) { failed++; continue; }
                if (repl.fsName === entry.filePath) { skipped++; continue; }
                pairs.push([entry.itemIndex, repl.fsName]);
            }
            var counts = delegateRelinkPairs(pairs) || { success: 0, failed: 0 };
            failed += counts.failed;
            refreshFromDoc();
            setStatus(L('message.changeExtDone') + "／" + kvLine('label.target', total, 'label.items') + "／" + kvLine('label.success', counts.success, 'label.items') + "／" + kvLine('label.skipped', skipped, 'label.items') + "／" + kvLine('label.failed', failed, 'label.items'));
        }

        function handleCollectLinks() {
            var df = parseWorkerResult(delegate("$.global.__LIM.docFolder()"));
            if (df.marker === "NODOC") { setStatus(L('status.noDocument')); return; }
            if (df.marker !== "OK" || !df.body) { setStatus(L('message.docNotSaved')); return; }
            var linksFolder = new Folder(df.body + "/Links");
            if (!linksFolder.exists && !linksFolder.create()) { setStatus(L('message.createLinksFolderFailed')); return; }
            var pairs = [], total = 0, copied = 0, skipped = 0, failed = 0;
            for (var k = 0; k < allPlacementEntries.length; k++) {
                var ent = allPlacementEntries[k];
                if (!ent.filePath || ent.filePath === "---") continue;
                total++;
                var srcFile = new File(ent.filePath);
                if (!srcFile.exists) { failed++; continue; }
                var destFile = new File(linksFolder.fsName + "/" + getRealFileName(srcFile));
                if (destFile.fsName === srcFile.fsName) { skipped++; continue; }
                if (destFile.exists) {
                    skipped++;
                } else {
                    if (!srcFile.copy(destFile.fsName)) { failed++; continue; }
                    copied++;
                }
                pairs.push([ent.itemIndex, destFile.fsName]);
            }
            var counts = delegateRelinkPairs(pairs) || { success: 0, failed: 0 };
            failed += counts.failed;
            refreshFromDoc();
            setStatus(L('message.collectLinksDone') + "／" + kvLine('label.target', total, 'label.items') + "／" + kvLine('label.copied', copied, 'label.items') + "／" + kvLine('label.skipped', skipped, 'label.items') + "／" + kvLine('label.failed', failed, 'label.items'));
        }

        openFolderBtn.enabled = false;

        // --- 下部：ボタン行 + ステータス表示 ---
        var btnGroup = palette.add("group");
        btnGroup.orientation = "row";
        btnGroup.alignment = "fill";
        btnGroup.alignChildren = ["fill", "center"];

        var openLinksPanelBtn = btnGroup.add("button", undefined, L('button.openLinksPanel'));
        openLinksPanelBtn.alignment = ["left", "center"];
        openLinksPanelBtn.onClick = function () { delegateOpenLinksPanel(); };

        var spacer = btnGroup.add("group");
        spacer.alignment = ["fill", "fill"];

        var reloadBtn = btnGroup.add("button", undefined, L('button.reload'));
        reloadBtn.alignment = ["right", "center"];
        reloadBtn.helpTip = (currentLanguage === 'ja') ? "ドキュメントから再読み込み" : "Reload from document";
        reloadBtn.onClick = guardAction(function () { refreshFromDoc(); });

        statusText = palette.add("statictext", undefined, "", { truncate: "middle" });
        statusText.alignment = ["fill", "bottom"];
        // loadData が出した理由（未オープン／読み込み失敗）を優先し、無い場合だけ件数表示にフォールバック
        if (lastStatusMessage) statusText.text = lastStatusMessage;
        else if (allPlacementEntries.length === 0) setStatus(L('status.noPlaced'));
        else setStatus(L('status.loaded') + "：" + withUnit(allPlacementEntries.length, 'label.items'));

        // 実行前に選択していた PlacedItem があれば対応行を初期選択（カンバスは触らない）
        function applyInitialSelection() {
            if (typeof preIndex !== "number" || preIndex < 0) return;
            for (var rowIdx = 0; rowIdx < filteredEntries.length; rowIdx++) {
                var ent = filteredEntries[rowIdx];
                var found = false;
                for (var indexIdx = 0; indexIdx < ent.itemIndices.length; indexIdx++) {
                    if (ent.itemIndices[indexIdx] === preIndex) { found = true; break; }
                }
                if (found) {
                    suppressCanvasOnce = true;
                    selectedEntry = ent;
                    selectedFilePath = ent.filePath;
                    pathStaticText.text = buildDisplayedPath(ent.filePath);
                    updateActionButtonStates();
                    highlightFolderFor(ent.filePath);
                    try {
                        var targetItem = listBox.items[rowIdx];
                        listBox.selection = targetItem;
                        if (typeof listBox.revealItem === "function") listBox.revealItem(targetItem);
                    } catch (e) {
                        ignoreError(function () { listBox.selection = rowIdx; });
                    }
                    break;
                }
            }
        }
        applyInitialSelection();

        // 空状態（ドキュメント未オープン／配置画像0件）は簡易表示にする。
        // オプション各パネル・リスト・パス欄を隠し、メッセージと［リンク］パネル／更新ボタン行だけ残す。
        // ScriptUI は visible = false でもレイアウト上の高さを保持するため、
        // maximumSize.height を 0 に潰して初めてウィンドウが縮む。
        function setSectionCollapsed(section, collapsed) {
            if (!section) return;
            if (section.__limMaxHeight === undefined) {
                section.__limMaxHeight = tryGet(function () { return section.maximumSize.height; }, 10000);
            }
            section.visible = !collapsed;
            ignoreError(function () {
                section.maximumSize.height = collapsed ? 0 : section.__limMaxHeight;
            });
        }

        function updateEmptyStateVisibility() {
            var isEmpty = (allPlacementEntries.length === 0);
            setSectionCollapsed(topRow, isEmpty);
            setSectionCollapsed(listHolder, isEmpty);
            setSectionCollapsed(pathPanel, isEmpty);
            ignoreError(function () {
                palette.layout.layout(true);
                palette.layout.resize();
            });
        }
        updateEmptyStateVisibility();

        // Esc で閉じる（× でも閉じられる）
        palette.addEventListener("keydown", function (k) {
            if (k.keyName === "Escape" || k.keyName === "Esc") { palette.close(); }
        });

        // カンバス側で選択した画像に一覧の行を合わせる。
        // Illustrator は選択変更を通知しないため、パレットがアクティブになった時点で取りに行く。
        function syncSelectionFromCanvas() {
            if (isActionRunning || isBusy) return;
            var canvasIndex = delegateCurrentIndex();
            if (canvasIndex < 0) return;
            if (selectedEntry && selectedEntry.itemIndices) {
                for (var i = 0; i < selectedEntry.itemIndices.length; i++) {
                    if (selectedEntry.itemIndices[i] === canvasIndex) return; /* すでに一致：触らない */
                }
            }
            restoreListSelectionByItemIndex(canvasIndex);
        }
        palette.onActivate = function () { syncSelectionFromCanvas(); };

        palette.onClose = function () {
            $.global.__LIM_paletteWindow = null;
            return true;
        };

        palette.center();
        palette.show();
    }

    // =========================================
    // メインエントリ / Main entry
    // =========================================

    // 多重起動防止：既存パレットがあれば閉じてから新規表示
    ignoreError(function () {
        if ($.global.__LIM_paletteWindow) {
            $.global.__LIM_paletteWindow.close();
            $.global.__LIM_paletteWindow = null;
        }
    });

    // 初回データ読み込み（ドキュメント未オープンでもパレットは開き、[更新] で再取得できる）
    loadData();
    showPalette();

})();
