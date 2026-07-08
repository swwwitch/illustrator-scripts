#target illustrator
#targetengine "LinkedImageManager"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

LinkedImageManager.jsx（常駐パレット版 / Persistent palette edition）

ドキュメント内の配置画像（リンク画像）を解析し、
一覧表示・絞り込み・重複管理・再リンク・リネーム・削除を一元化するユーティリティ。

Illustrator標準機能では行いにくい
「全体把握・状態確認・重複管理・再リンク作業」を効率化することを目的とする。

### パレット構成メモ / Palette architecture

・targetengine による常駐エンジンで動作する。常駐エンジンの app は表示中に DOM 接続を
  失うため、選択取得・配置解析・移動・属性変更など DOM を触る処理はすべて worker 関数に
  まとめ、押下のたびに BridgeTalk でメインエンジン（target=illustrator）へ委譲する。
・worker 関数は WORKER_FUNCS に全登録し、初回に一度だけメインエンジンへ常駐させる
  （$.global.__LIM）。以降は呼び出し式だけを送る。
・worker 関数内は行コメント禁止（toString が改行を消すため）・ブロックコメントのみ・
  必ずセミコロンで終える。
・戻り値はマーカー方式（OK+json / NODOC / ERR+msg）。

### 主な機能

・ファイル名／サイズ／使用数／配置寸法／スケール／PPI／カラースペースの一覧表示
・カラースペース表示（カラーモード + ICC プロファイル名／PNG・JPEG・PSD 対応）
・リンク状態（正常／リンク切れ／更新が必要）の判定とフィルタリング
・アートボード単位での絞り込みとビュー連動（選択時に移動・ズーム）
・同一ファイルの重複検出・統合表示・使用数カウント
・任意列でのソート／行選択に連動したカンバス上での選択・ズーム表示
・単一ファイル／フォルダー単位での再リンク、参照フォルダー内の別拡張子への一括再リンク
・リンクファイルのリネーム（物理ファイル名変更＋再リンクを同時実行）
・配置画像の削除（クリップグループ内は「画像のみ」「クリップグループごと」を選択）
・ファイル名のクリップボードコピー／リンクファイルのコピーと再リンク（Links フォルダーへ収集）

### 設定

・DROPBOX_PREFIX：Dropbox のローカルマウントパスを指定すると、
  リンクパス表示時に接頭辞を省略できる。使わない場合は "" を指定する。

### 参考記事

https://note.com/dtp_tranist/n/na66732d2056a

*/

(function () {

    // =========================================
    // バージョン / Version
    // =========================================

    var SCRIPT_VERSION = "v1.4.1";

    // =========================================
    // ユーザー設定 / User configuration
    // =========================================

    var DROPBOX_PREFIX = "/Users/takano/sw Dropbox/takano masahiro/";

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
            filterUpdate: { ja: "⟳ 更新が必要", en: "⟳ Needs Update" }
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
        return { status: L('label.statusOk'), statusIcon: "✓", isLinkOk: true };
    }

    // =========================================
    // パレット側 純粋ヘルパー / Palette-side pure helpers（DOM を触らない）
    // =========================================

    var PANEL_MARGINS = [10, 14, 10, 11];
    var PANEL_SPACING = 8;

    function setupPanel(panel, spacing) {
        panel.orientation = "column";
        panel.alignChildren = ["fill", "top"];
        panel.alignment = "fill";
        panel.margins = PANEL_MARGINS;
        panel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    function setupGroup(group, orientation, spacing) {
        group.orientation = orientation || "column";
        group.alignChildren = (group.orientation === "row") ? ["left", "center"] : ["fill", "top"];
        group.alignment = "fill";
        group.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

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

    // ボタン面に三角形（◀ / ▶）を onDraw で自前描画する。
    // フォントグリフに依存せず、枠＋塗り三角をベクターで描く。direction: -1=左, +1=右
    function attachArrowDraw(btn, direction) {
        btn.text = "";
        btn.onDraw = function () {
            var g = this.graphics;
            var w = this.size[0], h = this.size[1];
            if (!w || !h) return;
            var light = isLightUI();
            var color = this.enabled
                ? (light ? [0.13, 0.13, 0.13, 1] : [0.85, 0.85, 0.85, 1])
                : (light ? [0.60, 0.60, 0.60, 1] : [0.40, 0.40, 0.40, 1]);
            var pen = g.newPen(g.PenType.SOLID_COLOR, color, 1);
            var brush = g.newBrush(g.BrushType.SOLID_COLOR, color);

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
    function buildWorkerBundle() {
        var src = "$.global.__LIM = (function () {\n";
        for (var i = 0; i < WORKER_FUNCS.length; i++) {
            src += WORKER_FUNCS[i].toString() + "\n";
        }
        src += "return { analyze: w_analyze, select: w_select, fitArtboard: w_fitArtboard, openLinksPanel: w_openLinksPanel, relinkPairs: w_relinkPairs, probeDelete: w_probeDelete, del: w_del, docFolder: w_docFolder, copyText: w_copyText };\n";
        src += "})();";
        return src;
    }

    // メインエンジンへ 1 文を同期送信。結果 body を返す。エラー／タイムアウトは null
    function sendRaw(bodyExpr) {
        var holder = {};
        var bridge = new BridgeTalk();
        bridge.target = "illustrator";
        bridge.body = "eval(decodeURIComponent(\"" + encodeURIComponent(bodyExpr) + "\"))";
        bridge.onResult = function (msg) { holder.value = msg.body; };
        bridge.onError = function (msg) { holder.error = msg.body; };
        bridge.send(10);
        if (holder.hasOwnProperty("error")) return null;
        return holder.hasOwnProperty("value") ? holder.value : null;
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

    // "OK\n<json>" / "NODOC" / "ERR\n<msg>" を分解 (最初の \n が区切り)
    function parseWorkerResult(res) {
        if (res === null || res === undefined) return { marker: "ERR", body: "no result" };
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
        if (placedFile) { fileName = w_safeProp(placedFile, "name", fileName); }
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
            basics.fileName = w_safeProp(basics.linkedFile, "name", basics.fileName);
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

    function w_getLinkGroupKey(info) {
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
    function w_analyze() {
        try {
            if (app.documents.length === 0) return "NODOC";
            var doc = app.activeDocument;
            var placed = doc.placedItems;
            var collected = w_collectLinkInfo(doc, placed);
            var abs = [];
            for (var a = 0; a < doc.artboards.length; a++) { abs.push(w_safeProp(doc.artboards[a], "name", "")); }
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
            var payload = { docOpen: true, artboards: abs, preIndex: pre, entries: collected.linkInfoList, unique: collected.uniqueList };
            return "OK\n" + w_jsonStr(payload);
        } catch (e) {
            return "ERR\n" + e;
        }
    }

    /* itemIndex 配列に対応する配置画像をカンバス上で選択＆ズーム */
    function w_select(indices) {
        try {
            if (app.documents.length === 0) return "NODOC";
            var doc = app.activeDocument;
            var placed = doc.placedItems;
            doc.selection = null;
            for (var i = 0; i < indices.length; i++) {
                w_tryGet(function () { placed[indices[i]].selected = true; return 1; }, 0);
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

    /* [index, path] のペア配列に沿って placedItem のリンク先を差し替え。counts を返す */
    function w_relinkPairs(pairs) {
        try {
            if (app.documents.length === 0) return "NODOC";
            var placed = app.activeDocument.placedItems;
            var success = 0, failed = 0;
            for (var i = 0; i < pairs.length; i++) {
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

    /* 削除対象にクリップグループ内のものが含まれるか。CLIP / PLAIN を返す */
    function w_probeDelete(indices) {
        try {
            if (app.documents.length === 0) return "NODOC";
            var placed = app.activeDocument.placedItems;
            for (var i = 0; i < indices.length; i++) {
                var g = w_tryGet(function () { return w_findEnclosingClipGroup(placed[indices[i]]); }, null);
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
            var placed = app.activeDocument.placedItems;
            var refs = [];
            for (var i = 0; i < indices.length; i++) {
                w_tryGet(function () {
                    var pr = placed[indices[i]];
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
        w_collectXmpLinkedRefs, w_xmpNameKey, w_resolveXmpRef, w_getPlacedItemFileName,
        w_getPlacementFileBasics, w_getPlacementGeometry, w_getPlacementImageMeta, w_buildPlacementEntry,
        w_getLinkGroupKey, w_assignFileCounts, w_dedupeByFile, w_collectLinkInfo, w_pickSinglePlacedItem,
        w_zoomToSelection, w_fitViewToArtboard, w_findEnclosingClipGroup,
        w_analyze, w_select, w_fitArtboard, w_openLinksPanel,
        w_relinkPairs, w_probeDelete, w_del, w_docFolder, w_copyText
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

    function setStatus(msg) {
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

    // 解析を委譲してデータを読み込む。成功なら true。marker で状況を setStatus
    function loadData() {
        if (isBusy) return false;
        isBusy = true;
        try {
            var parsed = parseWorkerResult(delegate("$.global.__LIM.analyze()"));
            if (parsed.marker === "NODOC") {
                allPlacementEntries = [];
                uniqueFileEntries = [];
                setStatus(L('status.noDocument'));
                return false;
            }
            if (parsed.marker !== "OK") {
                setStatus(L('status.loadFailed') + "（" + parsed.body + "）");
                return false;
            }
            var data = tryGet(function () { return eval("(" + parsed.body + ")"); }, null);
            if (!data) {
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

    // 拡張子変更ダイアログ（参照フォルダー＋拡張子を選ぶモーダル）
    // 戻り値: { referenceFolder, primaryExt, fallbackExt } or null
    function showChangeExtensionDialog() {
        var extdialog = new Window("dialog", L('dialog.changeExt'));
        extdialog.orientation = "column";
        extdialog.alignChildren = ["fill", "top"];

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
        btnRow.alignment = ["right", "top"];
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
        deleteDialog.orientation = "column";
        deleteDialog.alignChildren = ["fill", "top"];
        deleteDialog.margins = 16;

        var msgText = L('message.confirmDeleteLinks') + "\n" +
            kvLine('label.target', targetCount, 'label.items') + "\n\n" +
            L('message.clipGroupDelete');
        var msg = deleteDialog.add("statictext", undefined, msgText, { multiline: true });
        msg.preferredSize.width = 360;

        var btnRow = deleteDialog.add("group");
        btnRow.orientation = "row";
        btnRow.alignment = ["right", "center"];
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
        palette.orientation = "column";
        palette.alignChildren = ["fill", "top"];
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

        var topRow = palette.add("group");
        topRow.orientation = "row";
        topRow.alignChildren = ["fill", "top"];

        var leftCol = topRow.add("group");
        leftCol.orientation = "column";
        leftCol.alignChildren = ["fill", "top"];

        var sortDropdown, ascRadio, descRadio;
        var currentVisibleSpecs = [];

        function createSortPanel(parent) {
            var sortPanel = parent.add("panel", undefined, L('panel.sort'));
            setupPanel(sortPanel);
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
        setupPanel(optPanel);
        var dedupCheck = optPanel.add("checkbox", undefined, L('checkbox.dedup'));
        dedupCheck.value = true;
        dedupCheck.helpTip = (currentLanguage === 'ja')
            ? "ON：同じリンクファイルを1行にまとめます。\nOFF：配置ごとに個別表示します。"
            : "ON: Group same linked files into one row.\nOFF: Each placement is listed separately.";
        var countColCheck = optPanel.add("checkbox", undefined, L('checkbox.displayFileCount'));
        countColCheck.value = true;

        // 選択時にズーム表示：「同一ファイル」パネルの直下に配置
        var showOnCanvasCheck = leftCol.add("checkbox", undefined, L('checkbox.showOnCanvas'));
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
            setupPanel(optionPanel);
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

        var okCheck, brokenCheck, updateCheck;

        function createStatusFilterPanel(parent) {
            var filterPanel = parent.add("panel", undefined, L('panel.status'));
            setupPanel(filterPanel);
            var statusGroup = filterPanel.add("group");
            statusGroup.orientation = "column";
            statusGroup.alignChildren = ["left", "top"];
            okCheck = statusGroup.add("checkbox", undefined, L('checkbox.filterOk'));
            brokenCheck = statusGroup.add("checkbox", undefined, L('checkbox.filterBroken'));
            updateCheck = statusGroup.add("checkbox", undefined, L('checkbox.filterUpdate'));
            okCheck.value = true;
            brokenCheck.value = true;
            updateCheck.value = true;
        }
        createStatusFilterPanel(otherTopRow);

        var abFilterDropdown, abPrevBtn, abNextBtn;

        function createArtboardFilterPanel(parent) {
            var abPanel = parent.add("panel", undefined, L('panel.artboard'));
            abPanel.orientation = "row";
            abPanel.alignChildren = ["left", "center"];
            abPanel.alignment = ["fill", "top"];
            abPanel.margins = PANEL_MARGINS;
            var artboardsWithImages = {};
            for (var entryIdx = 0; entryIdx < allPlacementEntries.length; entryIdx++) {
                var abNum = allPlacementEntries[entryIdx].artboardNum;
                if (abNum !== null) artboardsWithImages[abNum] = true;
            }

            var artboardDropdownItems = [L('label.artboardAll')];
            var artboardSep = (currentLanguage === 'ja') ? '：' : ': ';
            for (var artboardIndex = 0; artboardIndex < artboardNames.length; artboardIndex++) {
                var artboardName = artboardNames[artboardIndex] || "";
                artboardDropdownItems.push((artboardIndex + 1) + artboardSep + (artboardName || L('label.artboardFallback') + (artboardIndex + 1)));
            }
            abFilterDropdown = abPanel.add("dropdownlist", undefined, artboardDropdownItems);

            for (var di = 0; di < artboardNames.length; di++) {
                if (!artboardsWithImages[di + 1]) {
                    abFilterDropdown.items[di + 1].enabled = false;
                }
            }

            abFilterDropdown.selection = 0;
            abFilterDropdown.preferredSize.width = 200;

            abPrevBtn = abPanel.add("button", undefined, "");
            abPrevBtn.preferredSize = [22, 22];
            abPrevBtn.helpTip = L('label.prevArtboardTip');
            attachArrowDraw(abPrevBtn, -1);
            abNextBtn = abPanel.add("button", undefined, "");
            abNextBtn.preferredSize = [22, 22];
            abNextBtn.helpTip = L('label.nextArtboardTip');
            attachArrowDraw(abNextBtn, 1);

            var enabledCount = 0;
            for (var i = 1; i < abFilterDropdown.items.length; i++) {
                if (abFilterDropdown.items[i].enabled) enabledCount++;
            }
            if (enabledCount <= 1) {
                abPrevBtn.enabled = false;
                abNextBtn.enabled = false;
            }
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
            var hasStatusFlt = !(allowOk && allowBroken && allowUpdate);

            if (hasStatusFlt || hasArtboardFilter) {
                filteredEntries = [];
                for (var sourceIdx = 0; sourceIdx < sourceEntries.length; sourceIdx++) {
                    var entry = sourceEntries[sourceIdx];
                    if (hasStatusFlt) {
                        var statusCode = entry.statusCode;
                        if (statusCode === "ok" && !allowOk) continue;
                        if (statusCode === "broken" && !allowBroken) continue;
                        if (statusCode === "update" && !allowUpdate) continue;
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
        deleteLinkBtn.onClick = requireSelectedEntry(handleDeleteSelected);

        var renameLinkBtn = actionBtnRight.add("button", undefined, L('button.rename'));
        renameLinkBtn.preferredSize = [70, 24];
        renameLinkBtn.onClick = requireSelectedEntry(handleRenameSelected);

        var copyFileNameBtn = actionBtnLeft.add("button", undefined, L('button.copyFileName'));
        copyFileNameBtn.onClick = requireSelectedEntry(handleCopyFileName);

        var reloadOneBtn = actionBtnRight.add("button", undefined, L('button.relinkSelected'));
        reloadOneBtn.preferredSize = [94, 24];
        reloadOneBtn.onClick = requireSelectedEntry(handleRelinkSelected);

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
            reloadOneBtn.enabled = (selectedEntry !== null);
            renameLinkBtn.enabled = (selectedEntry !== null);
            openFileBtn.enabled = (selectedEntry !== null);
            deleteLinkBtn.enabled = (selectedEntry !== null);
            copyFileNameBtn.enabled = (selectedEntry !== null);
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
            updateEmptyStateVisibility();
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
            rebuildList();
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
        reloadFolderBtn.onClick = handleRelinkFolder;
        reloadFolderBtn.enabled = false;

        var changeExtensionBtn = folderActionLeft.add("button", undefined, L('button.changeExtension'));
        changeExtensionBtn.onClick = handleChangeExtension;
        changeExtensionBtn.enabled = false;

        var collectLinksBtn = folderActionRight.add("button", undefined, L('button.collectLinks'));
        collectLinksBtn.onClick = handleCollectLinks;

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
                var sourceFileName = entry.fileName || pathBaseName(entry.filePath);
                sourceFileName = tryGet(function () { return decodeURI(String(sourceFileName)); }, String(sourceFileName));
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
        reloadBtn.onClick = function () { refreshFromDoc(); };

        statusText = palette.add("statictext", undefined, "", { truncate: "middle" });
        statusText.alignment = ["fill", "bottom"];
        if (allPlacementEntries.length === 0) setStatus(L('status.noPlaced'));
        else setStatus(L('status.loaded') + "：" + withUnit(allPlacementEntries.length, 'label.items'));

        // 実行前に選択していた PlacedItem があれば対応行を初期選択（カンバスは触らない）
        if (typeof preIndex === "number" && preIndex >= 0) {
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

        // 空状態（ドキュメント未オープン／配置画像0件）は簡易表示にする。
        // オプション各パネル・リスト・パス欄を隠し、メッセージと［リンク］パネル／更新ボタン行だけ残す。
        function updateEmptyStateVisibility() {
            var isEmpty = (allPlacementEntries.length === 0);
            topRow.visible = !isEmpty;
            listHolder.visible = !isEmpty;
            pathPanel.visible = !isEmpty;
            ignoreError(function () { palette.layout.layout(true); });
        }
        updateEmptyStateVisibility();

        // Esc で閉じる（× でも閉じられる）
        palette.addEventListener("keydown", function (k) {
            if (k.keyName === "Escape" || k.keyName === "Esc") { palette.close(); }
        });

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
