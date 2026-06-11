#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

LinkedImageManager.jsx

ドキュメント内の配置画像（リンク画像）を解析し、
一覧表示・絞り込み・重複管理・再リンク・リネーム・削除を一元化するユーティリティ。

Illustrator標準機能では行いにくい
「全体把握・状態確認・重複管理・再リンク作業」を効率化することを目的とする。

### 主な機能

・ファイル名／サイズ／使用数／配置寸法／スケール／PPI／カラースペースの一覧表示
・カラースペース表示（カラーモード + ICC プロファイル名／PNG・JPEG・PSD 対応）
・ファイルサイズの単位切替（自動単位 B/KB/MB/GB ↔ MB 統一）
・リンク状態（正常／リンク切れ／更新が必要）の判定とフィルタリング
・リンク切れ時のファイル名フォールバック取得（XMPメタデータ参照）
・アートボード単位での絞り込みとビュー連動（選択時に移動・ズーム）
・同一ファイルの重複検出・統合表示・使用数カウント
・任意列でのソート（昇順／降順、数値系は自動降順）
・行選択に連動したカンバス上での選択・ズーム表示
・パス表示の最適化（~/ 表示／Dropbox パス短縮［未使用時は自動非表示］／ファイル名表示切替）
・リンク先フォルダの一覧表示／選択フォルダの直接オープン
・単一ファイル／フォルダ単位での再リンク
・参照フォルダ内の別拡張子ファイルへの一括再リンク（拡張子の変更）
・リンクファイルのリネーム（物理ファイル名変更＋再リンクを同時実行）
・配置画像の削除（クリップグループ内は「画像のみ」「クリップグループごと」を選択）
・ファイル名のクリップボードコピー
・リンクファイルのコピーと再リンク（Linksフォルダーへの収集）

### 設定

・DROPBOX_PREFIX：Dropbox のローカルマウントパスを指定すると、
  リンクパス表示時に接頭辞を省略できる。
  使わない場合は "" を指定すると、関連チェックボックスが非表示になる。

### 参考記事

https://note.com/dtp_tranist/n/na66732d2056a

*/

(function () {

    // =========================================
    // バージョン / Version
    // =========================================

    var SCRIPT_VERSION = "v1.3.0";

    // =========================================
    // ユーザー設定 / User configuration
    // =========================================

    // Dropbox のローカルマウントパス接頭辞（ここを自分の環境に書き換えてください）
    // Dropbox を使わない場合は "" を指定すると、パネルの「Dropboxパスを短縮」チェックボックス自体が非表示になります
    var DROPBOX_PREFIX = "/Users/takano/sw Dropbox/takano masahiro/";

    // =========================================
    // ローカライズ / Localization
    // =========================================

    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var currentLanguage = getCurrentLang();

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {

        // ダイアログタイトル / Dialog titles
        dialog: {
            main: { ja: "リンク画像の管理", en: "Linked Image Manager" },
            changeExt: { ja: "拡張子の変更", en: "Change Extension" },
            clipGroupDelete: { ja: "クリップグループ内の画像", en: "Image inside a clip group" }
        },

        // パネルタイトル / Panel titles
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

        // ソート / Sort options
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

        // 一覧の列見出し / List column headers
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

        // チェックボックス / Checkboxes
        checkbox: {
            dedup: { ja: "同一ファイルをまとめる", en: "Group Same Files" },
            unit: { ja: "単位を「MB」で統一", en: "Use MB" },
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

        // ボタン / Buttons
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
            relinkFolder: { ja: "フォルダーに再リンク", en: "Relink Folder" },
            changeExtension: { ja: "拡張子の変更", en: "Change Extension" },
            chooseFolder: { ja: "フォルダー指定", en: "Choose Folder" },
            collectLinks: { ja: "リンクを収集", en: "Collect Links" },
            openLinksPanel: { ja: "［リンク］パネルを開く", en: "Open Links Panel" },
            close: { ja: "閉じる", en: "Close" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },

        // アラート・確認・プロンプト / Alerts, confirmations, prompts
        message: {
            clipGroupDelete: {
                ja: "選択した画像はクリップグループ内にあります。どのように削除しますか？",
                en: "The selected image is inside a clip group. How would you like to delete it?"
            },
            noDocument: { ja: "ドキュメントが開いていません。", en: "No document is open." },
            noPlacedItems: {
                ja: "配置されているリンク画像が見つかりませんでした。",
                en: "No linked placed images were found."
            },
            selectItem: { ja: "リストからアイテムを選択してください。", en: "Please select an item from the list." },
            noValidPath: { ja: "有効なファイルパスがありません。", en: "No valid file path is available." },
            openFolderFailed: { ja: "フォルダを開けませんでした：", en: "Could not open the folder: " },
            selectLinkedFolder: { ja: "リンクフォルダを選択してください。", en: "Please select a linked folder." },
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
            linkFileNotFound: { ja: "リンクファイルが見つかりません：", en: "Linked file not found: " },
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

        // ステータス・その他ラベル / Status and miscellaneous labels
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

    // alert / confirm 用に「ラベル：値+単位」を 1 行で組み立てる
    // 例: kvLine('label.success', 3, 'label.items') → "成功：3件" / "Succeeded: 3 item(s)"
    function kvLine(labelKey, value, unitKey) {
        var sep = (currentLanguage === 'ja' ? '：' : ': ');
        var valueText = unitKey ? withUnit(value, unitKey) : String(value);
        return L(labelKey) + sep + valueText;
    }

    // =========================================
    // パネル・グループ共通レイアウト / Shared panel & group layout
    // =========================================

    /* パネルの余白と間隔 / Panel margins and spacing */
    var PANEL_MARGINS = [16, 20, 16, 12];
    var PANEL_SPACING = 8;

    /* パネルの共通設定 / Apply shared panel layout */
    function setupPanel(panel, spacing) {
        panel.orientation = "column";
        panel.alignChildren = ["fill", "top"];
        panel.alignment = "fill";
        panel.margins = PANEL_MARGINS;
        panel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /* グループの共通設定（orientation は呼び出し側で指定）/ Apply shared group layout (orientation passed in) */
    function setupGroup(group, orientation, spacing) {
        group.orientation = orientation || "column";
        group.alignChildren = ["left", "center"];
        group.alignment = "fill";
        group.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    // ScriptUI のレイアウト再計算を例外抑止付きで呼ぶ。引数は可変（複数コンテナを順に再計算）
    function safeRelayout() {
        for (var i = 0; i < arguments.length; i++) {
            var comp = arguments[i];
            if (!comp) continue;
            try { comp.layout.layout(true); } catch (e) { }
        }
    }

    // パスから最後のセパレータ（/ または \）の位置を返す。なければ -1
    function lastPathSep(path) {
        return Math.max(path.lastIndexOf("/"), path.lastIndexOf("\\"));
    }

    // パスから親フォルダ部分を返す（末尾セパレータなし）。判別不能なら ""
    function pathParent(path) {
        if (!path || path === "---") return "";
        var sep = lastPathSep(path);
        return (sep > 0) ? path.substring(0, sep) : "";
    }

    // パスからファイル名部分（最後のセパレータ以降）を返す
    function pathBaseName(path) {
        if (!path || path === "---") return "";
        var sep = lastPathSep(path);
        return (sep >= 0) ? path.substring(sep + 1) : path;
    }

    // ファイルパスからファイル名を除いた親フォルダ（末尾セパレータ付き）を返す
    function toFolderOnly(path) {
        if (!path || path === "---") return path;
        var sep = lastPathSep(path);
        if (sep < 0) return path;
        return path.substring(0, sep + 1);
    }

    // 表示用にパスを整形：Dropbox 短縮が ON のときは接頭辞を削って返し、それ以外で ~ 短縮が ON のときだけ ~/... 形式にする
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

    // 文字列をクリップボードへコピー：一時テキストフレームを作って app.copy() でクリップボードに送る
    // pbcopy 経由（app.system）は modern Illustrator で非同期動作になり、finally の tmpFile 削除が先に走って失敗するため使わない
    function copyTextToClipboard(text) {
        if (text === null || text === undefined) text = "";
        text = String(text);
        var activeDoc = tryGet(function () { return app.activeDocument; }, null);
        if (!activeDoc) return false;

        var prevSelection = tryGet(function () { return activeDoc.selection; }, null);
        var tempFrame = null;
        try {
            tempFrame = activeDoc.textFrames.add();
            tempFrame.contents = text;
            activeDoc.selection = null;
            tempFrame.selected = true;
            app.copy();
            tempFrame.remove();
            tempFrame = null;

            // 元の選択を可能な限り復元
            try { activeDoc.selection = null; } catch (e0) { }
            if (prevSelection && prevSelection.length) {
                for (var i = 0; i < prevSelection.length; i++) {
                    try { prevSelection[i].selected = true; } catch (ei) { }
                }
            }
            return true;
        } catch (e) {
            try { if (tempFrame) tempFrame.remove(); } catch (e3) { }
            return false;
        }
    }

    // ファイル名を base（拡張子なし）と ext（"."付き、なしは ""）に分割
    // 「.」が先頭にしかない隠しファイル（.DS_Store など）は拡張子なし扱い
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

    // File から fsName ベースで実ファイル名を取り出す
    // displayName は Finder の「拡張子を隠す」設定で拡張子が欠落することがあるため使わない
    function getRealFileName(file) {
        return file.fsName.split(/[\\\/]/).pop();
    }

    // 拡張子を保持したまま新しいファイル名をユーザーに入力させる
    // ・prompt の初期値は拡張子なしのベース名
    // ・誤って末尾に拡張子まで入力した場合は除去（大小文字無視）
    // ・キャンセル／空は null
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

    // ホームディレクトリ配下のパスを "~/..." 形式に短縮（範囲外はそのまま）
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

    // バイト数を MB 単位（小数 2 桁）の数値文字列に変換。列表示では単位ラベルを付けない
    function formatFileSize(bytes) {
        if (bytes === null || bytes === undefined || isNaN(bytes) || bytes < 0) return "-";
        return (bytes / (1024 * 1024)).toFixed(2);
    }

    // バイト数を自動単位（B / KB / MB / GB）付きで整形
    function formatFileSizeAuto(bytes) {
        if (bytes === null || bytes === undefined || isNaN(bytes) || bytes < 0) return "-";
        if (bytes < 1024) return bytes + "B";
        if (bytes < 1024 * 1024) return Math.round(bytes / 1024) + "KB";
        if (bytes < 1024 * 1024 * 1024) return (bytes / (1024 * 1024)).toFixed(1) + "MB";
        return (bytes / (1024 * 1024 * 1024)).toFixed(2) + "GB";
    }

    // 現在選択中のオブジェクトが画面いっぱいに収まるようズーム・センタリング
    function zoomToSelection(doc) {
        var selection = doc.selection;
        if (!selection || selection.length === 0) return;

        var firstBounds = selection[0].visibleBounds; // [left, top, right, bottom]
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
        var viewBounds = view.bounds; // 現在の表示領域（ドキュメント座標）
        var viewWidth = Math.abs(viewBounds[2] - viewBounds[0]);
        var viewHeight = Math.abs(viewBounds[1] - viewBounds[3]);

        var margin = 1.2; // 上下左右に 20% の余白
        var zoomX = view.zoom * viewWidth / (selectionWidth * margin);
        var zoomY = view.zoom * viewHeight / (selectionHeight * margin);
        var newZoom = Math.min(zoomX, zoomY);

        // Illustrator のズーム上限/下限に合わせて安全側にクランプ
        if (newZoom < 0.03125) newZoom = 0.03125; //  3.125%
        if (newZoom > 64) newZoom = 64;      // 6400%

        view.zoom = newZoom;
        view.centerPoint = [centerX, centerY];
    }

    // バイナリ文字列から big-endian の 16bit / 32bit 整数を取り出す
    function readU16BE(bytes, offset) {
        return ((bytes.charCodeAt(offset) & 0xFF) << 8) | (bytes.charCodeAt(offset + 1) & 0xFF);
    }
    function readU32BE(bytes, offset) {
        return ((bytes.charCodeAt(offset) & 0xFF) * 16777216) +
            ((bytes.charCodeAt(offset + 1) & 0xFF) << 16) +
            ((bytes.charCodeAt(offset + 2) & 0xFF) << 8) +
            (bytes.charCodeAt(offset + 3) & 0xFF);
    }

    // 例外を無視して評価し、失敗時は fallback を返す
    function tryGet(fn, fallback) {
        try {
            return fn();
        } catch (e) {
            return fallback;
        }
    }

    // File / Folder などの .exists を安全に取得
    function safeExists(target) {
        return tryGet(function () { return !!(target && target.exists); }, false);
    }

    // doc.selection から「単一の PlacedItem」を抜き出す
    // ・PlacedItem そのもの → そのまま返す
    // ・GroupItem（クリップグループ含む）配下に PlacedItem が 1 つだけ → その PlacedItem を返す
    // ・上記以外（複数選択、PlacedItem 非含有、複数の PlacedItem を含むグループなど）→ null
    function pickSinglePlacedItem(selection) {
        if (!selection || selection.length !== 1) return null;
        var topItem = selection[0];
        if (!topItem) return null;
        var typeName = tryGet(function () { return topItem.typename; }, "");
        if (typeName === "PlacedItem") return topItem;
        if (typeName !== "GroupItem") return null;

        var found = null;
        var multiple = false;
        function visit(group) {
            if (multiple) return;
            var children = tryGet(function () { return group.pageItems; }, null);
            if (!children) return;
            for (var i = 0; i < children.length; i++) {
                var child = children[i];
                var childTypeName = tryGet(function () { return child.typename; }, "");
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

    // PlacedItem の祖先を辿り、最も近い「クリップグループ（GroupItem.clipped === true）」を返す
    // 見つからなければ null
    function findEnclosingClipGroup(item) {
        var ancestor = tryGet(function () { return item.parent; }, null);
        while (ancestor) {
            var typeName = tryGet(function () { return ancestor.typename; }, "");
            if (typeName !== "GroupItem") return null; // GroupItem 以外（Layer / Document）に出たら終了
            var clipped = tryGet(function () { return ancestor.clipped; }, false);
            if (clipped) return ancestor;
            ancestor = tryGet(function () { return ancestor.parent; }, null);
        }
        return null;
    }

    // オブジェクトのプロパティ値を安全に取得
    function safeProp(obj, key, fallback) {
        return tryGet(function () {
            var value = obj[key];
            return (value !== undefined && value !== null && value !== "") ? value : fallback;
        }, fallback);
    }

    // PNG IHDR から幅/高さ・カラーモード、iCCP/sRGB チャンクから ICC プロファイル名を 1 度の走査でまとめて取得
    function readPngImageInfo(binaryFile) {
        var info = { width: null, height: null, colorMode: "", iccDesc: "" };
        // 8(sig) + 4(chunk length) + 4("IHDR") = 16 から IHDR データ。先頭 10 バイトに width/height/bitDepth/colorType まで含む
        binaryFile.seek(16);
        var ihdrBytes = binaryFile.read(10);
        if (ihdrBytes && ihdrBytes.length === 10) {
            info.width = readU32BE(ihdrBytes, 0);
            info.height = readU32BE(ihdrBytes, 4);
            var colorType = ihdrBytes.charCodeAt(9) & 0xFF;
            if (colorType === 0 || colorType === 4) info.colorMode = "Grayscale";
            else if (colorType === 3) info.colorMode = "Indexed";
            else info.colorMode = "RGB";
        }
        // iCCP / sRGB チャンクを探す（PNG sig + IHDR length + "IHDR" + IHDR data + CRC を飛ばす）
        binaryFile.seek(8 + 4 + 4 + 13 + 4);
        for (var chunkIdx = 0; chunkIdx < 32; chunkIdx++) {
            var pngChunkLengthBytes = binaryFile.read(4);
            if (!pngChunkLengthBytes || pngChunkLengthBytes.length < 4) break;
            var chunkLength = readU32BE(pngChunkLengthBytes, 0);
            var chunkType = binaryFile.read(4);
            if (!chunkType || chunkType.length < 4) break;
            if (chunkType === "iCCP") {
                var chunkData = binaryFile.read(chunkLength);
                if (chunkData) {
                    var nullIndex = chunkData.indexOf("\0");
                    if (nullIndex > 0) info.iccDesc = chunkData.substring(0, nullIndex);
                }
                break;
            } else if (chunkType === "sRGB") {
                info.iccDesc = "sRGB IEC61966-2.1";
                break;
            } else if (chunkType === "IDAT" || chunkType === "IEND") {
                break;
            } else {
                binaryFile.seek(binaryFile.tell() + chunkLength + 4); // data + CRC
            }
        }
        return info;
    }

    // JPEG マーカーを 1 度走査して SOFn から幅/高さ・コンポーネント数、APP2 から ICC プロファイルを集約
    function readJpegImageInfo(binaryFile) {
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
            var segmentLength = readU16BE(segmentLengthBytes, 0);
            var segmentDataLength = segmentLength - 2;
            if (segmentDataLength < 0) break;
            var segmentStart = binaryFile.tell();

            // SOFn（0xC0〜0xCF, ただし 0xC4/0xC8/0xCC は除く）
            var isSof = (markerCode >= 0xC0 && markerCode <= 0xCF) && markerCode !== 0xC4 && markerCode !== 0xC8 && markerCode !== 0xCC;
            if (isSof) {
                // SOFn segment data: precision(1) + height(2) + width(2) + Nf(1) + ...
                var sofBytes = binaryFile.read(6);
                if (sofBytes && sofBytes.length === 6) {
                    info.height = readU16BE(sofBytes, 1);
                    info.width = readU16BE(sofBytes, 3);
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
                    if (pieceLength > 0) iccPieces[sequenceNumber] = binaryFile.read(pieceLength);
                    else binaryFile.seek(segmentStart + segmentDataLength);
                } else {
                    binaryFile.seek(segmentStart + segmentDataLength);
                }
            } else {
                binaryFile.seek(segmentStart + segmentDataLength);
            }
        }
        if (componentCount === 1) info.colorMode = "Grayscale";
        else if (componentCount === 4) info.colorMode = "CMYK";
        else if (componentCount > 0) info.colorMode = "RGB";
        if (iccTotalPieces > 0) {
            var iccData = "";
            var allPiecesReceived = true;
            for (var pieceIdx = 1; pieceIdx <= iccTotalPieces; pieceIdx++) {
                if (!iccPieces[pieceIdx]) { allPiecesReceived = false; break; }
                iccData += iccPieces[pieceIdx];
            }
            if (allPiecesReceived && iccData.length > 0) info.iccDesc = readIccDesc(iccData);
        }
        return info;
    }

    // PSD ヘッダから幅/高さ・modeCode、Image Resources の resourceId=0x040F から ICC を取り出す
    function readPsdImageInfo(binaryFile) {
        var info = { width: null, height: null, colorMode: "", iccDesc: "" };
        // PSD header: sig(4) + ver(2) + reserved(6) + channels(2) + rows(4) + cols(4) + depth(2) + mode(2)
        binaryFile.seek(14);
        var psdDimsBytes = binaryFile.read(8);
        if (psdDimsBytes && psdDimsBytes.length === 8) {
            info.height = readU32BE(psdDimsBytes, 0);
            info.width = readU32BE(psdDimsBytes, 4);
        }
        var depthModeBytes = binaryFile.read(4); // depth(2) + mode(2)
        if (depthModeBytes && depthModeBytes.length === 4) {
            var modeCode = readU16BE(depthModeBytes, 2);
            if (modeCode === 0) info.colorMode = "Bitmap";
            else if (modeCode === 1) info.colorMode = "Grayscale";
            else if (modeCode === 2) info.colorMode = "Indexed";
            else if (modeCode === 3) info.colorMode = "RGB";
            else if (modeCode === 4) info.colorMode = "CMYK";
            else if (modeCode === 7) info.colorMode = "Multichannel";
            else if (modeCode === 8) info.colorMode = "Duotone";
            else if (modeCode === 9) info.colorMode = "Lab";
        }
        var colorModeDataLengthBytes = binaryFile.read(4);
        if (!colorModeDataLengthBytes || colorModeDataLengthBytes.length < 4) return info;
        var colorModeDataLength = readU32BE(colorModeDataLengthBytes, 0);
        var imageResourceStart = 30 + colorModeDataLength;
        binaryFile.seek(imageResourceStart);
        var imageResourceLengthBytes = binaryFile.read(4);
        if (!imageResourceLengthBytes || imageResourceLengthBytes.length < 4) return info;
        var imageResourceLength = readU32BE(imageResourceLengthBytes, 0);
        var imageResourceEnd = imageResourceStart + 4 + imageResourceLength;
        // リソースセクション終端まで走査。壊れたファイルで binaryFile.tell() が進まない場合は無限ループを避けて break
        while (binaryFile.tell() < imageResourceEnd) {
            var positionBefore = binaryFile.tell();
            var resourceSigBytes = binaryFile.read(4);
            if (!resourceSigBytes || resourceSigBytes.length < 4 || resourceSigBytes !== "8BIM") break;
            var resourceIdBytes = binaryFile.read(2);
            if (!resourceIdBytes || resourceIdBytes.length < 2) break;
            var resourceId = readU16BE(resourceIdBytes, 0);
            var nameLengthByte = binaryFile.read(1);
            if (!nameLengthByte) break;
            var nameLength = nameLengthByte.charCodeAt(0) & 0xFF;
            var paddedNameLength = nameLength + 1;
            if (paddedNameLength % 2 !== 0) paddedNameLength++;
            binaryFile.seek(binaryFile.tell() + (paddedNameLength - 1));
            var dataSizeBytes = binaryFile.read(4);
            if (!dataSizeBytes || dataSizeBytes.length < 4) break;
            var dataSize = readU32BE(dataSizeBytes, 0);
            var paddedDataSize = (dataSize % 2 === 0) ? dataSize : dataSize + 1;
            if (resourceId === 0x040F) {
                var iccProfileData = binaryFile.read(dataSize);
                if (iccProfileData && iccProfileData.length > 0) info.iccDesc = readIccDesc(iccProfileData);
                break;
            } else {
                binaryFile.seek(binaryFile.tell() + paddedDataSize);
            }
            // 進行していない（seek が無効）なら壊れたファイル。無限ループ回避
            if (binaryFile.tell() <= positionBefore) break;
        }
        return info;
    }

    // 画像ファイルを 1 度だけ開いてピクセル寸法・カラーモード・ICC プロファイル名をまとめて取得
    // 旧 readImagePixelSize / readImageColorSpace を統合したもの。戻り値: { width, height, colorMode, iccDesc } もしくは null
    function readImageInfo(file) {
        if (!file) return null;
        if (!safeExists(file)) return null;

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

                if (b0 === 0x89 && b1 === 0x50 && b2 === 0x4E && b3 === 0x47) {
                    info = readPngImageInfo(binaryFile);
                } else if (b0 === 0xFF && b1 === 0xD8) {
                    info = readJpegImageInfo(binaryFile);
                } else if (b0 === 0x38 && b1 === 0x42 && b2 === 0x50 && b3 === 0x53) {
                    info = readPsdImageInfo(binaryFile);
                }
            }
        } catch (e) { }
        binaryFile.close();
        return info;
    }

    // ICC プロファイルのバイナリから desc タグ（プロファイル名）を取り出す
    function readIccDesc(iccBuffer) {
        if (!iccBuffer || iccBuffer.length < 132) return "";
        try {
            var tagCount = readU32BE(iccBuffer, 128);
            for (var i = 0; i < tagCount; i++) {
                var tagOffset = 132 + i * 12;
                if (tagOffset + 12 > iccBuffer.length) break;
                var tagSignature = iccBuffer.substr(tagOffset, 4);
                if (tagSignature === "desc") {
                    var dataOffset = readU32BE(iccBuffer, tagOffset + 4);
                    var dataSize = readU32BE(iccBuffer, tagOffset + 8);
                    if (dataOffset + dataSize > iccBuffer.length) return "";
                    var dataType = iccBuffer.substr(dataOffset, 4);
                    if (dataType === "desc") {
                        // ICC v2: reserved(4) + asciiCount(4) + ascii...
                        var asciiCount = readU32BE(iccBuffer, dataOffset + 8);
                        if (asciiCount > 0) {
                            var asciiString = iccBuffer.substr(dataOffset + 12, asciiCount);
                            return asciiString.replace(/\0+$/g, "").replace(/\0.*$/g, "");
                        }
                    } else if (dataType === "mluc") {
                        // ICC v4: reserved(4) + recordCount(4) + recordSize(4) + records
                        var recordCount = readU32BE(iccBuffer, dataOffset + 8);
                        if (recordCount > 0) {
                            var recordOffset = dataOffset + 16;
                            var stringLength = readU32BE(iccBuffer, recordOffset + 4);
                            var stringOffset = readU32BE(iccBuffer, recordOffset + 8);
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
        } catch (e) { }
        return "";
    }

    // 配置サイズ（pt）とピクセル寸法から実効 PPI を算出
    function getEffectivePPI(item, pixelSize) {
        if (!pixelSize) return null;
        var placedWidthPt = tryGet(function () { return item.width; }, null);
        var placedHeightPt = tryGet(function () { return item.height; }, null);
        if (!placedWidthPt || !placedHeightPt) return null;
        var ppiX = pixelSize.width * 72 / placedWidthPt;
        var ppiY = pixelSize.height * 72 / placedHeightPt;
        return Math.round((ppiX + ppiY) / 2);
    }

    // XMP の日付文字列を Date に変換 / Convert XMP date string to Date
    function parseXmpDate(dateString) {
        if (!dateString) return null;
        var normalized = String(dateString).replace(/Z$/, "+00:00");
        var parsed = new Date(normalized);
        if (isNaN(parsed.getTime())) return null;
        return parsed;
    }

    // 「更新が必要」判定 / Detect whether the linked file is newer than the stored XMP timestamp
    function isLinkOutOfDate(placedFile, xmpLastModifyDate) {
        if (!safeExists(placedFile) || !xmpLastModifyDate) return false;

        var storedDate = parseXmpDate(xmpLastModifyDate);
        var currentDate = tryGet(function () { return placedFile.modified; }, null);
        if (!storedDate || !currentDate) return false;

        // ファイルシステムやXMPの丸め誤差を避けるため、5秒以上新しい場合のみ更新扱い
        return (currentDate.getTime() - storedDate.getTime()) > 5000;
    }

    // リンク状態を判定 / Resolve linked file status
    function resolveLinkStatus(placedFile, item, doc, xmpLastModifyDate) {
        if (!safeExists(placedFile)) {
            return {
                statusCode: "broken",
                status: L('label.statusBroken'),
                statusIcon: "⚠",
                isLinkOk: false
            };
        }

        if (isLinkOutOfDate(placedFile, xmpLastModifyDate)) {
            return {
                statusCode: "update",
                status: L('label.statusUpdate'),
                statusIcon: "⟳",
                isLinkOk: false
            };
        }

        return {
            statusCode: "ok",
            status: L('label.statusOk'),
            statusIcon: "✓",
            isLinkOk: true
        };
    }

    // 配置行列から拡大縮小率（%）を算出。X/Y が異なる場合は "X% × Y%"、同じなら単一値
    function getScaleInfo(item) {
        var matrix = tryGet(function () { return item.matrix; }, null);
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

    // 配置サイズ（pt）を mm に変換して {widthMm, heightMm} を返す。不明なら {null, null}
    function getDimensionsMm(item) {
        var widthPt = tryGet(function () { return item.width; }, null);   // pt
        var heightPt = tryGet(function () { return item.height; }, null);
        if (!widthPt || !heightPt) return { widthMm: null, heightMm: null };
        return {
            widthMm: widthPt * 25.4 / 72,
            heightMm: heightPt * 25.4 / 72
        };
    }

    // 指定アイテムの中心が含まれるアートボード番号（1 始まり）を返す。無ければ null
    function getArtboardNumber(item, doc) {
        var itemBounds = tryGet(function () { return item.visibleBounds; }, null); // [left, top, right, bottom]
        if (!itemBounds) return null;
        var centerX = (itemBounds[0] + itemBounds[2]) / 2;
        var centerY = (itemBounds[1] + itemBounds[3]) / 2;
        var artboards = tryGet(function () { return doc.artboards; }, null);
        if (!artboards) return null;
        for (var i = 0; i < artboards.length; i++) {
            var artboardRect = tryGet(function () { return artboards[i].artboardRect; }, null);
            if (!artboardRect) continue;
            var xMin = Math.min(artboardRect[0], artboardRect[2]);
            var xMax = Math.max(artboardRect[0], artboardRect[2]);
            var yMin = Math.min(artboardRect[1], artboardRect[3]);
            var yMax = Math.max(artboardRect[1], artboardRect[3]);
            if (centerX >= xMin && centerX <= xMax && centerY >= yMin && centerY <= yMax) {
                return i + 1;
            }
        }
        return null;
    }

    // XMPメタデータからリンク参照情報を収集 / Collect linked file references from XMP metadata
    function collectXmpLinkedRefs(doc) {
        var refs = [];
        var xmp = tryGet(function () { return new XML(doc.XMPString); }, null);
        if (!xmp) return refs;

        var paths = tryGet(function () { return xmp.xpath('//stRef:filePath'); }, null);
        var dates = tryGet(function () { return xmp.xpath('//stRef:lastModifyDate'); }, null);
        if (!paths) return refs;

        for (var i = 0; i < paths.length(); i++) {
            var filePath = paths[i].toString();
            var lastModifyDate = (dates && i < dates.length()) ? dates[i].toString() : "";
            refs.push({
                filePath: filePath,
                fileName: filePath.replace(/^.*[\/\\]/, ""),
                lastModifyDate: lastModifyDate
            });
        }
        return refs;
    }

    // 配置アイテムからファイル名を取得。リンク切れ時は XMP 由来の fallbackName を優先 / Get file name from placed item with XMP fallback
    function getPlacedItemFileName(placedItem, fallbackName, defaultName) {
        var fileName = defaultName || L('label.fileNameUnknown');

        var placedFile = tryGet(function () { return placedItem.file; }, null);
        if (placedFile) {
            fileName = safeProp(placedFile, "name", fileName);
        }

        if (fileName === L('label.fileNameUnknown') && fallbackName) {
            fileName = fallbackName;
        }

        if (fileName === L('label.fileNameUnknown')) {
            fileName = safeProp(placedItem, "name", fileName);
        }

        return fileName;
    }

    // 配置画像のファイル参照・リンク状態・サイズ・ファイル名を解決する
    function getPlacementFileBasics(item, xmpRef, doc) {
        var basics = {
            linkedFile: tryGet(function () { return item.file; }, null),
            filePath: "---",
            fileName: L('label.fileNameUnknown'),
            fileSize: "---",
            fileSizeBytes: -1,
            status: L('label.statusBroken'),
            statusCode: "broken",
            statusIcon: "⚠",
            isLinkOk: false
        };

        if (basics.linkedFile) {
            basics.fileName = safeProp(basics.linkedFile, "name", basics.fileName);
            basics.filePath = safeProp(basics.linkedFile, "fsName", basics.filePath);
            var resolvedStatus = resolveLinkStatus(basics.linkedFile, item, doc, xmpRef ? xmpRef.lastModifyDate : "");
            basics.statusCode = resolvedStatus.statusCode;
            basics.status = resolvedStatus.status;
            basics.statusIcon = resolvedStatus.statusIcon;
            basics.isLinkOk = resolvedStatus.isLinkOk;
            if (safeExists(basics.linkedFile)) {
                var byteLength = tryGet(function () { return basics.linkedFile.length; }, -1);
                if (byteLength >= 0) {
                    basics.fileSizeBytes = byteLength;
                    basics.fileSize = formatFileSize(basics.fileSizeBytes);
                }
            }
        }

        basics.fileName = getPlacedItemFileName(item, xmpRef ? xmpRef.fileName : "", basics.fileName);

        if (basics.fileName === L('label.fileNameUnknown') && basics.filePath !== "---") {
            var derivedName = pathBaseName(basics.filePath);
            if (derivedName) basics.fileName = derivedName;
        }

        return basics;
    }

    // 配置オブジェクトの寸法（mm）・スケール・所属アートボードを返す
    function getPlacementGeometry(item, doc) {
        var dimensions = getDimensionsMm(item);
        var widthText = (dimensions.widthMm !== null) ? dimensions.widthMm.toFixed(1) : "---";
        var heightText = (dimensions.heightMm !== null) ? dimensions.heightMm.toFixed(1) : "---";
        var scaleInfo = getScaleInfo(item);
        return {
            artboardNum: getArtboardNumber(item, doc),
            widthMm: dimensions.widthMm,
            heightMm: dimensions.heightMm,
            widthText: widthText,
            heightText: heightText,
            scalePct: scaleInfo.scalePct,
            scaleText: scaleInfo.scaleText
        };
    }

    // リンクファイルから実効 PPI とカラースペース表示を組み立てる
    // 旧実装では幅高さ取得用とカラースペース取得用に同じファイルを 2 回開いていたが、readImageInfo に統合済み
    function getPlacementImageMeta(item, linkedFile) {
        var ppi = null;
        var colorSpace = "";
        if (linkedFile) {
            var info = tryGet(function () { return readImageInfo(linkedFile); }, null);
            if (info) {
                if (info.width && info.height) {
                    ppi = getEffectivePPI(item, { width: info.width, height: info.height });
                }
                if (info.colorMode) {
                    colorSpace = info.iccDesc
                        ? info.colorMode + "（" + info.iccDesc + "）"
                        : info.colorMode;
                }
            }
        }
        return {
            ppi: ppi,
            ppiText: (ppi !== null) ? String(ppi) : "---",
            colorSpace: colorSpace
        };
    }

    // placedItems からリンク情報を収集し、フラットリストと重複排除済みリストを返す
    // 1 つの placedItem からファイル情報・寸法・PPI・カラースペース等を抽出して 1 行分の info を作る
    function buildPlacementEntry(item, itemIndex, xmpRef, doc) {
        var basics = getPlacementFileBasics(item, xmpRef, doc);
        var geometry = getPlacementGeometry(item, doc);
        var imageMeta = getPlacementImageMeta(item, basics.linkedFile);

        return {
            itemIndex: itemIndex,
            index: itemIndex + 1,
            fileName: basics.fileName,
            filePath: basics.filePath,
            fileSize: basics.fileSize,
            fileSizeBytes: basics.fileSizeBytes,
            status: basics.status,
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

    // 集約用キー：filePath が取れていればそれで集約、リンク切れ（"---"）は itemIndex 付きで個別扱い
    // 同名でも元ファイルが別物の可能性があるため、XMP 由来のファイル名で安易にまとめない
    function getLinkGroupKey(info) {
        if (info.filePath && info.filePath !== "---") return info.filePath;
        return "__broken__#" + info.itemIndex + "#" + (info.fileName || "");
    }

    // info の同一ファイル数を集計し、各 info に fileCount を付加する（破壊的）
    function assignFileCounts(linkInfoList) {
        var countMap = {};
        for (var i = 0; i < linkInfoList.length; i++) {
            var key = getLinkGroupKey(linkInfoList[i]);
            countMap[key] = (countMap[key] || 0) + 1;
        }
        for (var j = 0; j < linkInfoList.length; j++) {
            linkInfoList[j].fileCount = countMap[getLinkGroupKey(linkInfoList[j])];
        }
    }

    // 同一ファイルでまとめた重複排除済みリストを返す。artboardSet / artboards / artboardNum も集約。
    function dedupeByFile(linkInfoList) {
        var uniqueList = [];
        var keyToEntry = {};
        for (var di = 0; di < linkInfoList.length; di++) {
            var dedupInfo = linkInfoList[di];
            var dedupKey = getLinkGroupKey(dedupInfo);
            if (keyToEntry[dedupKey]) {
                keyToEntry[dedupKey].itemIndices.push(dedupInfo.itemIndex);
                if (dedupInfo.artboardNum !== null) {
                    keyToEntry[dedupKey].artboardSet[dedupInfo.artboardNum] = true;
                }
            } else {
                var artboardSet = {};
                if (dedupInfo.artboardNum !== null) artboardSet[dedupInfo.artboardNum] = true;
                var entry = {
                    index: uniqueList.length + 1,
                    fileName: dedupInfo.fileName,
                    filePath: dedupInfo.filePath,
                    fileSize: dedupInfo.fileSize,
                    fileSizeBytes: dedupInfo.fileSizeBytes,
                    status: dedupInfo.status,
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

    function collectLinkInfo(doc, placedItems) {
        var xmpRefs = collectXmpLinkedRefs(doc);
        var linkInfoList = [];
        for (var i = 0; i < placedItems.length; i++) {
            linkInfoList.push(buildPlacementEntry(placedItems[i], i, xmpRefs[i] || null, doc));
        }
        assignFileCounts(linkInfoList);
        var uniqueList = dedupeByFile(linkInfoList);
        return { linkInfoList: linkInfoList, uniqueList: uniqueList };
    }

    // ドキュメントが開いているか確認
    if (app.documents.length === 0) {
        alert(L('message.noDocument'));
        return;
    }

    var doc = app.activeDocument;
    var placedItems = doc.placedItems; // 配置画像（リンク画像）を取得

    // 配置画像が存在するか確認
    if (placedItems.length === 0) {
        alert(L('message.noPlacedItems'));
        return;
    }

    // ---- リンク情報を収集 / Collect link information ----
    var collected = collectLinkInfo(doc, placedItems);
    var linkInfoList = collected.linkInfoList;
    var uniqueList = collected.uniqueList;

    // ---- 実行前の選択状態チェック：PlacedItem 1 つだけなら初期選択行として渡す / Capture initial selection before dialog ----
    // PlacedItem 直接選択のほか、クリップグループなど GroupItem 配下に PlacedItem が 1 つだけ含まれる場合にも対応
    var preselectedItemIndex = -1;
    var selection = tryGet(function () { return doc.selection; }, null);
    var selectedPlacedItem = pickSinglePlacedItem(selection);
    if (selectedPlacedItem) {
        // まず参照同一性で照合
        for (var si = 0; si < placedItems.length; si++) {
            if (placedItems[si] === selectedPlacedItem) {
                preselectedItemIndex = si;
                break;
            }
        }
        // === が DOM プロキシ差で失敗するケースに備え、geometricBounds で再照合
        if (preselectedItemIndex < 0) {
            var selectedBounds = tryGet(function () { return selectedPlacedItem.geometricBounds; }, null);
            if (selectedBounds) {
                for (var sj = 0; sj < placedItems.length; sj++) {
                    var candidateBounds = tryGet(function () { return placedItems[sj].geometricBounds; }, null);
                    if (candidateBounds && candidateBounds[0] === selectedBounds[0] && candidateBounds[1] === selectedBounds[1] && candidateBounds[2] === selectedBounds[2] && candidateBounds[3] === selectedBounds[3]) {
                        preselectedItemIndex = sj;
                        break;
                    }
                }
            }
        }
    }

    // ---- ダイアログ表示 / Show dialog ----
    showLinkDialog(linkInfoList, uniqueList, placedItems, doc, preselectedItemIndex);

    // ---- ダイアログ構築 / Build dialog ----
    function showLinkDialog(allPlacementEntries, uniqueFileEntries, placedItems, doc, preselectedItemIndex) {
        var MAIN_LISTBOX_SIZE = [500, 190]; // メイン一覧のサイズ
        var FOLDER_LISTBOX_SIZE = [500, 120]; // リンクフォルダ一覧のサイズ

        var dialog = new Window("dialog", L('dialog.main') + " " + SCRIPT_VERSION);
        dialog.orientation = "column";
        dialog.alignChildren = ["fill", "top"];
        dialog.preferredSize.width = 500;

        // ソート・フィルタ前の元データ（重複チェックボックスで切り替え）
        var sourceEntries = uniqueFileEntries;
        // 実際にリストに表示しているエントリ群（フィルタ後）— listBox.onChange などから参照
        var filteredEntries = sourceEntries;
        // プログラムから選択を入れる際に、カンバス側の再選択＆ズームを 1 回だけ抑制するフラグ
        var suppressCanvasOnce = false;
        // リスト再構築後に選択を復元するための配置アイテム index
        var pendingSelectionItemIndex = -1;

        // --- 上部 2 カラム：左（ソート + オプション） / 右（その他） ---
        var topRow = dialog.add("group");
        topRow.orientation = "row";
        topRow.alignChildren = ["fill", "top"];

        // 左カラム：ソート + オプション を縦に
        var leftCol = topRow.add("group");
        leftCol.orientation = "column";
        leftCol.alignChildren = ["fill", "top"];

        // ソート
        var sortPanel = leftCol.add("panel", undefined, L('panel.sort'));
        setupPanel(sortPanel);
        var sortKeyRow = sortPanel.add("group");
        sortKeyRow.orientation = "row";
        sortKeyRow.alignChildren = ["left", "center"];
        sortKeyRow.add("statictext", undefined, labelText('sort.by'));
        var sortDropdown = sortKeyRow.add("dropdownlist", undefined, []);
        // ソート対象は列の表示状態と連動（rebuildSortDropdown で動的に構築）
        var currentVisibleSpecs = [];
        var orderRow = sortPanel.add("group");
        orderRow.orientation = "row";
        orderRow.alignChildren = ["left", "center"];
        var ascRadio = orderRow.add("radiobutton", undefined, L('sort.asc'));
        var descRadio = orderRow.add("radiobutton", undefined, L('sort.desc'));
        descRadio.value = true;

        // オプション（重複 / ファイルサイズの単位を統一）
        var optPanel = leftCol.add("panel", undefined, L('panel.sameFile'));
        setupPanel(optPanel);
        var dedupCheck = optPanel.add("checkbox", undefined, L('checkbox.dedup'));
        dedupCheck.value = true; // ON：同一ファイルは 1 行にまとめる
        dedupCheck.helpTip = (currentLanguage === 'ja')
            ? "ON：同じリンクファイルを1行にまとめ、再リンクは“すべて一括”で実行されます（複数変更に注意）。\nOFF：配置ごとに個別表示され、再リンクも1件ずつ実行されます。"
            : "ON: Group same linked files into one row; relink applies to ALL placements (affects multiple items).\nOFF: Each placement is listed separately; relink applies individually.";
        var countColCheck = optPanel.add("checkbox", undefined, L('checkbox.displayFileCount'));
        countColCheck.value = true;

        // 右カラム：表示列・ステータス・アートボードを配置
        var otherPanel = topRow.add("group");
        otherPanel.orientation = "column";
        otherPanel.alignChildren = ["fill", "top"];
        // otherPanel.margins = PANEL_MARGINS;

        var otherTopRow = otherPanel.add("group");
        otherTopRow.orientation = "row";
        otherTopRow.alignChildren = ["fill", "top"];

        // 左：表示列の表示／非表示を切り替える
        var optionPanel = otherTopRow.add("panel", undefined, L('panel.displayColumn'));
        setupPanel(optionPanel);
        var sizeRow = optionPanel.add("group");
        sizeRow.orientation = "row";
        sizeRow.alignChildren = ["left", "center"];
        var sizeColCheck = sizeRow.add("checkbox", undefined, L('checkbox.displaySize'));

        var unitCheck = optionPanel.add("checkbox", undefined, L('checkbox.unit'));

        var dimScalePpiCheck = optionPanel.add("checkbox", undefined, L('checkbox.displayDimScalePpi'));
        var colorSpaceColCheck = optionPanel.add("checkbox", undefined, L('checkbox.displayColorSpace'));
        sizeColCheck.value = false;
        unitCheck.value = true; // ON：全行 MB で統一表示、OFF：B/KB/MB/GB の自動単位
        unitCheck.enabled = sizeColCheck.value;
        dimScalePpiCheck.value = false;
        colorSpaceColCheck.value = false;

        // 右：ステータスフィルター
        var filterPanel = otherTopRow.add("panel", undefined, L('panel.status'));
        setupPanel(filterPanel);
        var statusGroup = filterPanel.add("group");
        statusGroup.orientation = "column";
        statusGroup.alignChildren = ["left", "top"];
        var okCheck = statusGroup.add("checkbox", undefined, L('checkbox.filterOk'));
        var brokenCheck = statusGroup.add("checkbox", undefined, L('checkbox.filterBroken'));
        var updateCheck = statusGroup.add("checkbox", undefined, L('checkbox.filterUpdate'));
        okCheck.value = true;
        brokenCheck.value = true;
        updateCheck.value = true;

        // 下段：アートボード（右カラム内で左右貫通）
        var abPanel = otherPanel.add("panel", undefined, L('panel.artboard'));
        abPanel.orientation = "row";
        abPanel.alignChildren = ["left", "center"];
        abPanel.alignment = ["fill", "top"];
        abPanel.margins = PANEL_MARGINS;
        // 配置画像のあるアートボード番号を収集（無い番号はドロップダウンでディム表示）
        var artboardsWithImages = {};
        for (var entryIdx = 0; entryIdx < allPlacementEntries.length; entryIdx++) {
            var abNum = allPlacementEntries[entryIdx].artboardNum;
            if (abNum !== null) artboardsWithImages[abNum] = true;
        }

        var artboardDropdownItems = [L('label.artboardAll')];
        var artboardSep = (currentLanguage === 'ja') ? '：' : ': ';
        for (var artboardIndex = 0; artboardIndex < doc.artboards.length; artboardIndex++) {
            var artboardName = "";
            artboardName = safeProp(doc.artboards[artboardIndex], "name", artboardName);
            artboardDropdownItems.push((artboardIndex + 1) + artboardSep + (artboardName || L('label.artboardFallback') + (artboardIndex + 1)));
        }
        var abFilterDropdown = abPanel.add("dropdownlist", undefined, artboardDropdownItems);

        // 配置画像がないアートボードはディム表示（↑↓ボタンでもスキップ）
        for (var di = 0; di < doc.artboards.length; di++) {
            if (!artboardsWithImages[di + 1]) {
                abFilterDropdown.items[di + 1].enabled = false;
            }
        }

        abFilterDropdown.selection = 0;
        abFilterDropdown.preferredSize.width = 200;

        var abPrevBtn = abPanel.add("button", undefined, "◀");
        abPrevBtn.preferredSize = [30, 22];
        abPrevBtn.helpTip = L('label.prevArtboardTip');
        var abNextBtn = abPanel.add("button", undefined, "▶");
        abNextBtn.preferredSize = [30, 22];
        abNextBtn.helpTip = L('label.nextArtboardTip');

        var enabledCount = 0;
        for (var i = 1; i < abFilterDropdown.items.length; i++) {
            if (abFilterDropdown.items[i].enabled) enabledCount++;
        }
        if (enabledCount <= 1) {
            abPrevBtn.enabled = false;
            abNextBtn.enabled = false;
        }

        // --- リストボックス（アートボード列の有無で再生成するため、コンテナ経由で配置） ---
        var listHolder = dialog.add("group");
        listHolder.orientation = "column";
        listHolder.alignChildren = ["fill", "top"];

        var listBox = null;
        // 行選択中の絶対パス＋エントリ（ボタンや整形切替で使用）
        // Absolute path and entry for the currently selected row.
        var selectedFilePath = "";
        var selectedEntry = null;
        // rebuildList() から呼ぶ選択復元処理。path UI 側の依存関数定義後に実体を代入する。
        // Selection-restore handler called from rebuildList(); assigned after path-UI dependent functions are defined.
        var restoreListSelectionByItemIndex = null;

        // 各列の有効/無効・キー・タイトル・幅を表示オプションに応じて返す
        function getColumnSpec() {
            var cols = [];
            cols.push({ key: "statusIcon", title: "", width: 30 });
            cols.push({ key: "fileName", title: L('column.fileName'), width: 210 });
            if (sizeColCheck.value) {
                cols.push({
                    key: "fileSize",
                    title: unitCheck.value ? L('column.fileSizeMb') : L('column.fileSize'),
                    width: 65
                });
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
            // アートボードドロップダウンの選択に応じて列を表示／非表示
            // Show/hide the Artboard column based on the artboard dropdown selection.
            var shouldShowArtboardColumn = !abFilterDropdown.selection || abFilterDropdown.selection.index === 0;
            if (shouldShowArtboardColumn) {
                cols.push({ key: "artboards", title: L('column.artboards'), width: 70 });
            }
            return cols;
        }

        function createListBox() {
            if (listBox) {
                try { listHolder.remove(listBox); } catch (e) { }
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

        // ソートキー仕様：listBox で非表示の項目はドロップダウンから自動的に除外
        // Hidden listBox columns are automatically removed from the sort dropdown.
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

        // ソートキーの値を取得（currentVisibleSpecs のインデックス経由）
        // Resolve the sort value via the currentVisibleSpecs index.
        function getSortValue(info, sortBy) {
            var spec = currentVisibleSpecs[sortBy];
            if (!spec) return info.fileName;
            return spec.getValue(info);
        }

        function isMissing(v) {
            return v === null || v === undefined || (typeof v === "number" && isNaN(v));
        }

        // ソート＋フィルタ適用 → リスト再構築
        function rebuildList() {
            var selectionItemIndexToRestore = pendingSelectionItemIndex;
            pendingSelectionItemIndex = -1;
            if (selectionItemIndexToRestore < 0 && selectedEntry && selectedEntry.itemIndices && selectedEntry.itemIndices.length > 0) {
                selectionItemIndexToRestore = selectedEntry.itemIndices[0];
            }

            var sortBy = sortDropdown.selection ? sortDropdown.selection.index : 0;
            var desc = descRadio.value;

            // フィルタ：ステータス別チェック／アートボード指定（組み合わせ）
            var selectedArtboardIndex = abFilterDropdown.selection ? abFilterDropdown.selection.index : 0;
            var hasArtboardFilter = selectedArtboardIndex > 0;
            var targetArtboardNumber = selectedArtboardIndex; // ドロップダウン index はそのまま 1 始まりのアートボード番号
            var allowOk = okCheck.value;
            var allowBroken = brokenCheck.value;
            var allowUpdate = updateCheck.value;
            var hasStatusFlt = !(allowOk && allowBroken && allowUpdate); // 全て ON なら素通し

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
                filteredEntries = sourceEntries;
            }

            filteredEntries.sort(function (a, b) {
                var valA = getSortValue(a, sortBy);
                var valB = getSortValue(b, sortBy);

                // 欠損値は常に末尾
                // Missing values are always placed at the end.
                var aMissing = isMissing(valA), bMissing = isMissing(valB);
                if (aMissing && bMissing) return a.index - b.index;
                if (aMissing) return 1;
                if (bMissing) return -1;

                if (valA < valB) return desc ? 1 : -1;
                if (valA > valB) return desc ? -1 : 1;

                // ExtendScript の sort 不安定対策：最後は元 index で固定
                // Stabilize ExtendScript sort by falling back to original index.
                return a.index - b.index;
            });

            // 列スペックに沿って動的にセル値を流し込む
            var columns = getColumnSpec();
            var useUnifiedSize = unitCheck.value;
            listBox.removeAll();
            for (var rowIdx = 0; rowIdx < filteredEntries.length; rowIdx++) {
                var info = filteredEntries[rowIdx];
                // 先頭列は .text、残りは subItems にセット
                var firstKey = columns[0].key;
                var firstText = info[firstKey];
                var row = listBox.add("item", (firstText === undefined || firstText === null) ? "" : String(firstText));
                for (var colIdx = 1; colIdx < columns.length; colIdx++) {
                    var key = columns[colIdx].key;
                    var cellText;
                    if (key === "fileSize") {
                        cellText = useUnifiedSize
                            ? formatFileSize(info.fileSizeBytes)
                            : formatFileSizeAuto(info.fileSizeBytes);
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


        // 数値系のソートに切り替えたら自動で降順 ON（spec.preferDesc による）
        sortDropdown.onChange = function () {
            var idx = sortDropdown.selection ? sortDropdown.selection.index : 0;
            var spec = currentVisibleSpecs[idx];
            if (spec && spec.preferDesc) {
                descRadio.value = true;
                ascRadio.value = false;
            }
            rebuildList();
        };
        ascRadio.onClick = rebuildList;
        descRadio.onClick = rebuildList;

        // 「重複」チェックボックス：ON で重複をまとめる、OFF で全配置を個別表示
        dedupCheck.onClick = function () {
            pendingSelectionItemIndex = (selectedEntry && selectedEntry.itemIndices && selectedEntry.itemIndices.length > 0)
                ? selectedEntry.itemIndices[0]
                : -1;
            sourceEntries = dedupCheck.value ? uniqueFileEntries : allPlacementEntries;
            rebuildList();
        };

        // ステータスフィルター：いずれかが変わったら再描画
        okCheck.onClick = rebuildList;
        brokenCheck.onClick = rebuildList;
        updateCheck.onClick = rebuildList;

        // アートボードフィルター：特定アートボードを選んだら、冗長なアートボード列は自動で OFF
        function applyArtboardFilter() {
            var selectedArtboardIndex = abFilterDropdown.selection ? abFilterDropdown.selection.index : 0;

            // アートボードを選択した場合（0 は「すべて」なので除外）
            if (selectedArtboardIndex > 0) {
                try {
                    var abIndex = selectedArtboardIndex - 1; // 0-based に変換
                    doc.artboards.setActiveArtboardIndex(abIndex);

                    // ビューをそのアートボードにフィット
                    var rect = doc.artboards[abIndex].artboardRect;
                    var cx = (rect[0] + rect[2]) / 2;
                    var cy = (rect[1] + rect[3]) / 2;

                    var view = doc.views[0];
                    view.centerPoint = [cx, cy];

                    // ズーム調整（軽めにフィット）
                    var vw = Math.abs(view.bounds[2] - view.bounds[0]);
                    var vh = Math.abs(view.bounds[1] - view.bounds[3]);
                    var aw = Math.abs(rect[2] - rect[0]);
                    var ah = Math.abs(rect[1] - rect[3]);

                    var zoomX = view.zoom * vw / aw;
                    var zoomY = view.zoom * vh / ah;
                    view.zoom = Math.min(zoomX, zoomY) * 0.9; // 少し余白を持たせる

                } catch (e) {
                    // ignore
                }
            }

            createListBox();
            bindListBoxEvents();
            rebuildSortDropdown();
            dialog.layout.layout(true);
            rebuildList();
        }
        abFilterDropdown.onChange = applyArtboardFilter;

        // direction: -1 で前、+1 で次のアートボードへ移動。enabled な項目だけを巡回
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

        abPrevBtn.onClick = function () { stepArtboard(-1); };
        abNextBtn.onClick = function () { stepArtboard(1); };

        // 列の表示/非表示・ヘッダ更新時の共通処理（listBox を作り直し、ソート候補も更新）
        function recreateListBoxAndRebuildList() {
            createListBox();
            bindListBoxEvents();
            rebuildSortDropdown();
            dialog.layout.layout(true);
            rebuildList();
        }
        unitCheck.onClick = recreateListBoxAndRebuildList;
        sizeColCheck.onClick = function () {
            unitCheck.enabled = sizeColCheck.value;
            recreateListBoxAndRebuildList();
        };
        countColCheck.onClick = recreateListBoxAndRebuildList;
        dimScalePpiCheck.onClick = recreateListBoxAndRebuildList;
        colorSpaceColCheck.onClick = recreateListBoxAndRebuildList;

        // 初期表示（createListBox / rebuildSortDropdown / rebuildList）は
        // pathStaticText など UI 構築後にまとめて実行する
        // The initial display calls are batched after the path UI is built.

        // --- パスパネル直上：選択時の挙動オプション ---
        var viewOptRow = dialog.add("group");
        viewOptRow.orientation = "row";
        viewOptRow.alignment = ["fill", "top"];
        viewOptRow.alignChildren = ["left", "center"];
        var showOnCanvasCheck = viewOptRow.add("checkbox", undefined, L('checkbox.showOnCanvas'));
        showOnCanvasCheck.value = true; // ON：行選択でカンバス上の該当画像をフィット表示

        // --- パスパネル ---
        var pathPanel = dialog.add("panel", undefined, L('panel.path'));
        setupPanel(pathPanel);

        // 1 行目：選択中の親フォルダ表示
        var pathRow = pathPanel.add("group");
        pathRow.orientation = "row";
        pathRow.alignChildren = ["fill", "center"];

        var pathStaticText = pathRow.add(
            "statictext", undefined, L('label.pathPlaceholder'),
            { multiline: true }
        );
        pathStaticText.alignment = ["fill", "fill"];
        pathStaticText.preferredSize = [450, 20];
        pathStaticText.helpTip = L('label.pathHelpTip');

        // 2 行目：整形オプション（アクションボタンはさらに下の独立行に配置）
        var pathOptRow = pathPanel.add("group");
        pathOptRow.orientation = "row";
        pathOptRow.alignment = "fill";
        pathOptRow.alignChildren = ["fill", "center"];

        var pathOptLeft = pathOptRow.add("group");
        pathOptLeft.orientation = "row";
        pathOptLeft.alignChildren = ["left", "center"];
        pathOptLeft.alignment = ["left", "center"];
        var fullPathCheck = pathOptLeft.add("checkbox", undefined, L('checkbox.fullPath'));
        // DROPBOX_PREFIX が空文字なら Dropbox チェックボックスは追加せず、参照だけ満たすスタブを使う
        var dropboxCheck;
        if (DROPBOX_PREFIX) {
            dropboxCheck = pathOptLeft.add("checkbox", undefined, L('checkbox.dropbox'));
            dropboxCheck.value = true;
        } else {
            dropboxCheck = { value: false, enabled: false };
        }
        var fileNameCheck = pathOptLeft.add("checkbox", undefined, L('checkbox.fileName'));
        fullPathCheck.value = false; // ON：フルパスをそのまま表示。OFF：~/ や Dropbox 接頭辞で短縮
        fileNameCheck.value = false; // ON：パスにファイル名まで含める。OFF：親フォルダまで表示

        var pathOptSpacer = pathOptRow.add("group");
        pathOptSpacer.alignment = ["fill", "fill"];

        var pathOptRight = pathOptRow.add("group");
        pathOptRight.orientation = "row";
        pathOptRight.alignChildren = ["right", "center"];
        pathOptRight.alignment = ["right", "center"];

        // selectedEntry 未選択ガード付きで onClick ハンドラを生成
        function requireSelectedEntry(handler) {
            return function () {
                if (!selectedEntry) {
                    alert(L('message.selectItem'));
                    return;
                }
                return handler.apply(this, arguments);
            };
        }

        // listBox.onChange / rebuildList() から呼ばれるため、初期表示前に定義しておく
        // Defined before the initial list build because listBox.onChange / rebuildList() depend on these functions.

        // フルパス ON：そのまま表示。OFF：~ 短縮や Dropbox 短縮を適用する（ファイル名 ON でも短縮は効く）
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

        // リスト再構築後、配置アイテム index を基準に該当行を再選択する
        // Restore the row selection by the placed item index after rebuilding the list.
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
                            try { listBox.selection = i; } catch (e2) { }
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

        // 初期表示：listBox とソートドロップダウンを構築し、使用数をデフォルト選択
        // Initial display: build listBox / sort dropdown and default to fileCount.
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

        // Dropbox 対応が ON のときは Dropbox 接頭辞短縮が優先されるので、フルパスチェックをディム
        // ディムするだけだと .value が残って buildDisplayedPath / populateFoldersList で挙動がズレるため、
        // 無効化と同時に value も false に倒して状態を一本化する
        function updateFullPathEnable() {
            fullPathCheck.enabled = !dropboxCheck.value;
            if (!fullPathCheck.enabled) fullPathCheck.value = false;
        }

        fullPathCheck.onClick = onPathOptionChange;
        dropboxCheck.onClick = function () {
            updateFullPathEnable();
            onPathOptionChange();
        };
        fileNameCheck.onClick = function () {
            updatePathDisplay();
            pathPanel.layout.layout(true);
            dialog.layout.layout(true);
        }; // フォルダ一覧には影響しない
        updateFullPathEnable(); // 初期化

        // 選択エントリに対応する itemIndices のリンク先を newFile に差し替え
        function relinkIndicesTo(indices, newFile) {
            var success = 0, failed = 0;
            for (var i = 0; i < indices.length; i++) {
                try {
                    placedItems[indices[i]].file = newFile;
                    success++;
                } catch (e) {
                    failed++;
                }
            }
            return { success: success, failed: failed };
        }

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

            // 110px 幅の縦カラムを 1 つ追加して返す
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

                var entry = {
                    ui: radioBtn,
                    ext: ext,
                    alt: alt || ""
                };
                radios.push(entry);

                radioBtn.onClick = function () {
                    // 他のラジオをすべてOFFにする（グループを跨いで排他にする）
                    for (var i = 0; i < radios.length; i++) {
                        if (radios[i].ui !== radioBtn) {
                            radios[i].ui.value = false;
                        }
                    }
                    radioBtn.value = true;
                };
            }

            // 3 カラム × 3 行のラジオを定義どおりに配置
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

            var cancelBtn = btnRow.add("button", undefined, L('button.cancel'));
            okBtn = btnRow.add("button", undefined, "OK");
            okBtn.enabled = false;

            cancelBtn.onClick = function () {
                extdialog.close(0);
            };

            okBtn.onClick = function () {
                if (!referenceFolder) return;
                extdialog.close(1);
            };

            if (extdialog.show() !== 1) return null;

            var selected = findSelectedRadioSpec();
            if (!selected) return null;
            return {
                referenceFolder: referenceFolder,
                primaryExt: selected.ext,
                fallbackExt: selected.alt
            };
        }

        // 参照フォルダー内を列挙し、baseName と拡張子を大小文字無視で比較する。
        // ファイルシステムが case-sensitive（APFS の一部構成など）でも
        // 拡張子の揺れ（.PNG / .Jpg など）を拾えるようにするため。
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

            // primary → fallback の優先順で探す
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

        function normalizeFolderPathForCompare(path) {
            if (!path) return "";
            path = String(path).replace(/\\/g, "/");
            path = path.replace(/\/+$/g, "");
            return path;
        }

        // oldFolder 配下に配置された全ファイルを newFolder 内の同名ファイルに差し替え
        // （InDesign の「フォルダーに再リンク」相当）
        function relinkFolder(oldFolder, newFolder) {
            var success = 0, failed = 0, total = 0;
            for (var k = 0; k < allPlacementEntries.length; k++) {
                var ent = allPlacementEntries[k];
                if (!ent.filePath || ent.filePath === "---") continue;
                if (pathParent(ent.filePath) === oldFolder) {
                    total++;
                    var fname = pathBaseName(ent.filePath);
                    try {
                        var nf = new File(newFolder.fsName + "/" + fname);
                        if (nf.exists) {
                            placedItems[ent.itemIndex].file = nf;
                            success++;
                        } else {
                            failed++;
                        }
                    } catch (e) { failed++; }
                }
            }
            return { success: success, failed: failed, total: total };
        }

        // アクションボタンは pathPanel 直下の独立行に 3 カラム（左：ファイル名コピー / 中央：spacer / 右：開く・削除・リネーム・再リンク）
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
            if (!absPath || absPath === "---") {
                alert(L('message.noValidPath'));
                return;
            }
            var fileToOpen = new File(absPath);
            if (!fileToOpen.exists) {
                alert(L('message.linkFileNotFound') + absPath);
                return;
            }
            fileToOpen.execute();
        });

        // インデックス列から PlacedItem 参照と内包クリップグループのペアを集める
        function collectDeleteRefs(indices) {
            var refs = [];
            for (var di = 0; di < indices.length; di++) {
                try {
                    var placedRef = placedItems[indices[di]];
                    refs.push({ placed: placedRef, clipGroup: findEnclosingClipGroup(placedRef) });
                } catch (eRef) { }
            }
            return refs;
        }

        // refs の中にクリップグループ内のものが 1 つでもあるか
        function hasClippedRef(refs) {
            for (var dc = 0; dc < refs.length; dc++) {
                if (refs[dc].clipGroup) return true;
            }
            return false;
        }

        // refs を一括削除（clipMode='group' なら囲むクリップグループごと削除）
        function deleteRefs(refs, clipMode) {
            var success = 0, failed = 0;
            for (var dk = 0; dk < refs.length; dk++) {
                var entry = refs[dk];
                try {
                    if (entry.clipGroup && clipMode === 'group') {
                        entry.clipGroup.remove();
                    } else {
                        entry.placed.remove();
                    }
                    success++;
                } catch (eRm) {
                    failed++;
                }
            }
            return { success: success, failed: failed };
        }

        function handleDeleteSelected() {
            var indices = selectedEntry.itemIndices || [];
            if (indices.length === 0) return;

            // インデックスを伴う削除は collection の再採番で破綻するため、先に PageItem 参照を集める
            var refs = collectDeleteRefs(indices);

            // 確認ダイアログ：
            // ・クリップグループを含まない → 通常の確認ダイアログ（はい/いいえ）
            // ・含む                      → 確認と削除モード選択を 1 つにまとめた 3 ボタンダイアログ
            var clipMode = 'image';
            if (hasClippedRef(refs)) {
                clipMode = askDeleteModeWithConfirm(indices.length);
                if (clipMode === null) return;
            } else {
                var confirmMessage = L('message.confirmDeleteLinks') + "\n" +
                    kvLine('label.target', indices.length, 'label.items');
                if (!confirm(confirmMessage)) return;
            }

            var res = deleteRefs(refs, clipMode);

            // 選択は削除済みなので解除しておく（refreshFromDoc の itemIndex 復元に巻き込まれないように）
            selectedEntry = null;
            selectedFilePath = "";

            app.redraw();
            refreshFromDoc();
            updateActionButtonStates();
            alert(
                L('message.deleteDone') + "\n" +
                kvLine('label.success', res.success, 'label.items') + "\n" +
                kvLine('label.failed', res.failed, 'label.items')
            );
        }

        var deleteLinkBtn = actionBtnRight.add("button", undefined, L('button.delete'));
        deleteLinkBtn.preferredSize = [50, 24];
        deleteLinkBtn.onClick = requireSelectedEntry(handleDeleteSelected);

        // クリップグループ内の画像が含まれる場合の、確認 + 削除モード選択を兼ねたダイアログ
        // 戻り値: 'image' / 'group' / null（キャンセル）
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

        // newName と oldFile を比較し、上書き先がいる場合の処理（確認＋削除）を行う
        // 戻り値: true = rename を続行してよい、false = 中止
        function prepareRenameOverwrite(oldFile, newFile) {
            if (!newFile.exists) return true;
            // 大小文字だけの違いは macOS の case-insensitive FS で同一ファイルを指すため
            // 削除対象にしない（rename は OS 側で正しく処理される）
            if (newFile.fsName.toLowerCase() === oldFile.fsName.toLowerCase()) return true;
            if (!confirm(L('message.confirmOverwrite') + newFile.fsName)) return false;
            // File.rename は既存ファイルがあると失敗するので、先に削除してから rename する
            var removed = tryGet(function () { return newFile.remove(); }, false);
            if (!removed) {
                alert(L('message.renameFailed'));
                return false;
            }
            return true;
        }

        function handleRenameSelected() {
            var absPath = selectedEntry.filePath;
            if (!absPath || absPath === "---") {
                alert(L('message.noValidPath'));
                return;
            }
            var oldFile = new File(absPath);
            if (!oldFile.exists) {
                alert(L('message.linkFileNotFound') + absPath);
                return;
            }

            var oldName = getRealFileName(oldFile);
            var oldFolder = oldFile.parent;

            var newName = promptNewFileName(oldName);
            if (newName === null) return;
            if (newName === oldName) {
                alert(L('message.nameUnchanged'));
                return;
            }
            if (/[\/\\]/.test(newName)) {
                alert(L('message.invalidFileName'));
                return;
            }

            var newFile = new File(oldFolder.fsName + "/" + newName);
            if (!prepareRenameOverwrite(oldFile, newFile)) return;

            // 物理リネーム（失敗時は再リンクしない）
            var renamed = tryGet(function () { return oldFile.rename(newName); }, false);
            if (!renamed) {
                alert(L('message.renameFailed'));
                return;
            }

            var res = relinkIndicesTo(selectedEntry.itemIndices, newFile);
            app.redraw();
            refreshFromDoc();
            alert(
                L('message.renameDone') + "\n" +
                oldName + " → " + newName + "\n" +
                kvLine('label.success', res.success, 'label.items') + "\n" +
                kvLine('label.failed', res.failed, 'label.items')
            );
        }

        var renameLinkBtn = actionBtnRight.add("button", undefined, L('button.rename'));
        renameLinkBtn.preferredSize = [70, 24];
        renameLinkBtn.onClick = requireSelectedEntry(handleRenameSelected);
        var copyFileNameBtn = actionBtnLeft.add("button", undefined, L('button.copyFileName'));
        copyFileNameBtn.onClick = requireSelectedEntry(function () {
            var name = selectedEntry.fileName || "";
            if (copyTextToClipboard(name)) {
                alert(L('message.copyFileNameDone') + "\n" + name);
            } else {
                alert(L('message.copyFileNameFailed'));
            }
        });
        var reloadOneBtn = actionBtnRight.add("button", undefined, L('button.relinkSelected'));
        reloadOneBtn.preferredSize = [80, 24];

        reloadOneBtn.onClick = requireSelectedEntry(function () {
            var relinkCount = (selectedEntry.itemIndices && selectedEntry.itemIndices.length)
                ? selectedEntry.itemIndices.length
                : 0;
            if (relinkCount > 1) {
                var confirmMessage = L('message.confirmBatchRelink') + "\n" +
                    kvLine('label.target', relinkCount, 'label.items');
                if (!confirm(confirmMessage)) return;
            }

            var picked = File.openDialog(L('label.selectNewLinkFile'));
            if (!picked) return;
            var res = relinkIndicesTo(selectedEntry.itemIndices, picked);
            app.redraw();
            refreshFromDoc();
            alert(
                L('message.relinkDone') + "\n" +
                kvLine('label.success', res.success, 'label.items') + "\n" +
                kvLine('label.failed', res.failed, 'label.items')
            );
        });

        // 「同一ファイルをまとめる」の状態に応じて、再リンクボタンの意味を明示する
        // ただし、まとめ対象の使用数が 1 の場合は「一括」ではなく単数ラベルを使う
        function updateRelinkButtonLabel() {
            var placementCount = (selectedEntry && selectedEntry.itemIndices) ? selectedEntry.itemIndices.length : 0;
            var useBatchLabel = dedupCheck.value && placementCount > 1;
            var nextLabel = useBatchLabel ? L('button.relinkAll') : L('button.relinkSelected');
            var nextWidth = 94;
            var nextHeight = 24;

            reloadOneBtn.text = nextLabel;

            // ScriptUI は text 変更後に preferredSize だけでは反映されない場合があるため、
            // size / bounds も明示的に更新する。
            reloadOneBtn.preferredSize = [nextWidth, nextHeight];
            reloadOneBtn.size = [nextWidth, nextHeight];
            try {
                reloadOneBtn.bounds = [
                    reloadOneBtn.bounds[0],
                    reloadOneBtn.bounds[1],
                    reloadOneBtn.bounds[0] + nextWidth,
                    reloadOneBtn.bounds[1] + nextHeight
                ];
            } catch (e0) { }

            safeRelayout(actionBtnRight, actionBtnRow, pathPanel, dialog);
        }

        // パスが「---」（不明）のときは再リンクボタンを無効化
        function updateActionButtonStates() {
            reloadOneBtn.enabled = (selectedEntry !== null);
            renameLinkBtn.enabled = (selectedEntry !== null);
            openFileBtn.enabled = (selectedEntry !== null);
            deleteLinkBtn.enabled = (selectedEntry !== null);
            copyFileNameBtn.enabled = (selectedEntry !== null);
            updateRelinkButtonLabel();
        }

        // 初期状態：まだ選択がないので各ボタンを無効化し、再リンクボタンの表示も整える
        updateActionButtonStates();

        // --- リンクフォルダ一覧（パスパネル内、重複排除） ---
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
                foldersListBox.add("item",
                    formatDisplayPath(linkedFolderPaths[ii], !fullPathCheck.value, dropboxCheck.value)
                );
            }
        }
        populateFoldersList();

        // リンク後に doc から再収集してリストを更新
        function refreshFromDoc() {
            var collected = collectLinkInfo(doc, placedItems);
            allPlacementEntries = collected.linkInfoList;
            uniqueFileEntries = collected.uniqueList;
            sourceEntries = dedupCheck.value ? uniqueFileEntries : allPlacementEntries;

            // 選択中エントリはソース側から再検索して再紐付け（filePath が変わっていても itemIndices で拾う）
            if (selectedEntry) {
                var anchor = (selectedEntry.itemIndices && selectedEntry.itemIndices.length > 0)
                    ? selectedEntry.itemIndices[0] : -1;
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

        // 行ダブルクリックで Finder/Explorer を開く
        foldersListBox.onDoubleClick = function () {
            if (foldersListBox.selection === null) return;
            var idx = foldersListBox.selection.index;
            try {
                var folderPath = linkedFolderPaths[idx];
                var ff = new Folder(folderPath);
                if (ff.exists) ff.execute();
            } catch (e) {
                alert(L('message.openFolderFailed') + e.message);
            }
        };

        // --- リンクフォルダ一覧の直下：3カラム（左：一括再リンク / 中央：spacer / 右：リンク収集） ---
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
            if (foldersListBox.selection === null) {
                alert(L('message.selectLinkedFolder'));
                return;
            }
            try {
                var folderPath = linkedFolderPaths[foldersListBox.selection.index];
                var folderToOpen = new Folder(folderPath);
                if (folderToOpen.exists) folderToOpen.execute();
            } catch (e) {
                alert(L('message.openFolderFailed') + e.message);
            }
        };

        var folderActionSpacer = folderActionRow.add("group");
        folderActionSpacer.alignment = ["fill", "fill"];

        var folderActionRight = folderActionRow.add("group");
        folderActionRight.orientation = "row";
        folderActionRight.alignChildren = ["right", "center"];
        folderActionRight.alignment = ["right", "center"];

        var reloadFolderBtn = folderActionLeft.add("button", undefined, L('button.relinkFolder'));
        reloadFolderBtn.onClick = function () {
            if (foldersListBox.selection === null) {
                alert(L('message.selectLinkedFolder'));
                return;
            }
            var oldFolder = linkedFolderPaths[foldersListBox.selection.index];
            var newFolder = Folder.selectDialog(L('label.selectAltFolder'));
            if (!newFolder) return;
            var res = relinkFolder(oldFolder, newFolder);
            app.redraw();
            refreshFromDoc();
            alert(
                L('message.relinkDone') + "\n" +
                kvLine('label.target', res.total, 'label.items') + "\n" +
                kvLine('label.success', res.success, 'label.items') + "\n" +
                kvLine('label.failed', res.failed, 'label.items') + (currentLanguage === 'ja' ? '（同名ファイルなし）' : ' (same-name file not found)')
            );
        };
        reloadFolderBtn.enabled = false;
        openFolderBtn.enabled = false;

        // entry に対し referenceFolder 内の同名・指定拡張子ファイルへ再リンク
        // 戻り値: 'success' / 'failed' / 'skipped'
        function applyExtensionChangeToEntry(entry, referenceFolder, extPrefs) {
            var placedItem = placedItems[entry.itemIndex];
            // entry.fileName は File.name 由来で URI エンコード済みの可能性があるため、
            // 比較側（findReplacementFileByExtension で decodeURI したファイル名）と揃えるためデコードする
            var sourceFileName = entry.fileName || pathBaseName(entry.filePath);
            sourceFileName = tryGet(function () { return decodeURI(String(sourceFileName)); }, String(sourceFileName));
            var baseName = splitFileName(sourceFileName).base;
            if (!baseName) return 'failed';

            var replacementFile = findReplacementFileByExtension(
                referenceFolder, baseName, extPrefs.primaryExt, extPrefs.fallbackExt
            );
            if (!replacementFile) return 'failed';

            // 置換先が現在のリンクと同一ファイルなら差し替え不要
            try {
                var currentPath = placedItem.file ? placedItem.file.fsName : null;
                if (currentPath && currentPath === replacementFile.fsName) return 'skipped';
            } catch (e) { }

            try {
                placedItem.file = replacementFile;
                return 'success';
            } catch (e2) {
                return 'failed';
            }
        }

        function handleChangeExtension() {
            if (foldersListBox.selection === null) {
                alert(L('message.selectLinkedFolder'));
                return;
            }

            var sourceFolderPath = normalizeFolderPathForCompare(
                linkedFolderPaths[foldersListBox.selection.index]
            );

            // 1. 参照フォルダー＋拡張子を選択
            var extPrefs = showChangeExtensionDialog();
            if (!extPrefs) return;

            // 2. 実行
            var success = 0, failed = 0, skipped = 0, total = 0;
            for (var i = 0; i < allPlacementEntries.length; i++) {
                var entry = allPlacementEntries[i];
                if (!entry || !entry.filePath || entry.filePath === "---") continue;
                if (normalizeFolderPathForCompare(pathParent(entry.filePath)) !== sourceFolderPath) continue;

                total++;
                var outcome = applyExtensionChangeToEntry(entry, extPrefs.referenceFolder, extPrefs);
                if (outcome === 'success') success++;
                else if (outcome === 'skipped') skipped++;
                else failed++;
            }

            app.redraw();
            refreshFromDoc();

            alert(
                L('message.changeExtDone') + "\n" +
                kvLine('label.target', total, 'label.items') + "\n" +
                kvLine('label.success', success, 'label.items') + "\n" +
                kvLine('label.skipped', skipped, 'label.items') + "\n" +
                kvLine('label.failed', failed, 'label.items')
            );
        }

        var changeExtensionBtn = folderActionLeft.add("button", undefined, L('button.changeExtension'));
        changeExtensionBtn.onClick = handleChangeExtension;
        changeExtensionBtn.enabled = false;

        // 保存済みドキュメントの隣に Links フォルダーを用意する。失敗時は alert を出して null
        function ensureLinksFolder() {
            var docFile = tryGet(function () { return doc.fullName; }, null);
            if (!docFile || !docFile.parent || !docFile.parent.exists) {
                alert(L('message.docNotSaved'));
                return null;
            }
            var linksFolder = new Folder(docFile.parent.fsName + "/Links");
            if (!linksFolder.exists && !linksFolder.create()) {
                alert(L('message.createLinksFolderFailed'));
                return null;
            }
            return linksFolder;
        }

        // 1 件分のコピー＆再リンク。戻り値: 'copied' / 'skipped' / 'failed'
        function collectLinkForEntry(ent, linksFolder) {
            try {
                var srcFile = new File(ent.filePath);
                if (!srcFile.exists) return 'failed';

                var destFile = new File(linksFolder.fsName + "/" + getRealFileName(srcFile));
                // 既に収集先に同名ファイルがある場合はコピーをスキップ（同一パスなら再リンクも不要）
                if (destFile.fsName === srcFile.fsName) return 'skipped';

                var outcome;
                if (destFile.exists) {
                    outcome = 'skipped';
                } else {
                    if (!srcFile.copy(destFile.fsName)) return 'failed';
                    outcome = 'copied';
                }

                placedItems[ent.itemIndex].file = destFile;
                return outcome;
            } catch (e) {
                return 'failed';
            }
        }

        function handleCollectLinks() {
            var linksFolder = ensureLinksFolder();
            if (!linksFolder) return;

            var copied = 0, skipped = 0, failed = 0, total = 0;
            for (var k = 0; k < allPlacementEntries.length; k++) {
                var ent = allPlacementEntries[k];
                if (!ent.filePath || ent.filePath === "---") continue;
                total++;
                var outcome = collectLinkForEntry(ent, linksFolder);
                if (outcome === 'copied') copied++;
                else if (outcome === 'skipped') skipped++;
                else failed++;
            }

            app.redraw();
            refreshFromDoc();

            alert(
                L('message.collectLinksDone') + "\n" +
                kvLine('label.target', total, 'label.items') + "\n" +
                kvLine('label.copied', copied, 'label.items') + "\n" +
                kvLine('label.skipped', skipped, 'label.items') + "\n" +
                kvLine('label.failed', failed, 'label.items')
            );
        }

        var collectLinksBtn = folderActionRight.add("button", undefined, L('button.collectLinks'));
        collectLinksBtn.onClick = handleCollectLinks;

        // フォルダ選択が変わったらボタンを enable/disable
        foldersListBox.onChange = function () {
            var hasSelection = (foldersListBox.selection !== null);
            reloadFolderBtn.enabled = hasSelection;
            openFolderBtn.enabled = hasSelection;
            changeExtensionBtn.enabled = hasSelection;
        };

        // メインリスト選択に合わせてフォルダ一覧の該当行をハイライト
        function highlightFolderFor(absPath) {
            if (!absPath || absPath === "---") {
                foldersListBox.selection = null;
                return;
            }
            var folder = pathParent(absPath);
            if (!folder) return;
            for (var folderIdx = 0; folderIdx < linkedFolderPaths.length; folderIdx++) {
                if (linkedFolderPaths[folderIdx] === folder) {
                    try {
                        foldersListBox.selection = foldersListBox.items[folderIdx];
                        if (typeof foldersListBox.revealItem === "function") {
                            foldersListBox.revealItem(foldersListBox.items[folderIdx]);
                        }
                    } catch (e) { }
                    return;
                }
            }
            foldersListBox.selection = null;
        }

        // 指定エントリに対応する配置を Illustrator 上で選択＆ズーム
        function selectPlacedItemsOnCanvas(info) {
            if (!info) return;
            var indices = info.itemIndices;
            doc.selection = null;
            for (var idx = 0; idx < indices.length; idx++) {
                placedItems[indices[idx]].selected = true;
            }
            zoomToSelection(doc);
        }

        // --- ボタングループ（左：リンクパネル起動 / 右：キャンセル） ---
        var btnGroup = dialog.add("group");
        btnGroup.orientation = "row";
        btnGroup.alignment = "fill";
        btnGroup.alignChildren = ["fill", "center"];

        var openLinksPanelBtn = btnGroup.add("button", undefined, L('button.openLinksPanel'));
        openLinksPanelBtn.alignment = ["left", "center"];
        openLinksPanelBtn.onClick = function () {
            try { app.executeMenuCommand('Adobe LinkPalette Menu Item'); } catch (e) { }
            dialog.close();
        };

        var spacer = btnGroup.add("group");
        spacer.alignment = ["fill", "fill"];

        var closeBtn = btnGroup.add("button", undefined, L('button.close'), { name: "ok" });
        closeBtn.alignment = ["right", "center"];

        // 「閉じる」ボタン
        closeBtn.onClick = function () {
            dialog.close();
        };

        // 実行前に選択していた PlacedItem があれば、対応行を初期選択（カンバス側は触らない）
        if (typeof preselectedItemIndex === "number" && preselectedItemIndex >= 0) {
            for (var rowIdx = 0; rowIdx < filteredEntries.length; rowIdx++) {
                var ent = filteredEntries[rowIdx];
                var found = false;
                for (var indexIdx = 0; indexIdx < ent.itemIndices.length; indexIdx++) {
                    if (ent.itemIndices[indexIdx] === preselectedItemIndex) { found = true; break; }
                }
                if (found) {
                    suppressCanvasOnce = true;
                    selectedEntry = ent;
                    selectedFilePath = ent.filePath;
                    pathStaticText.text = buildDisplayedPath(ent.filePath);
                    updateActionButtonStates();
                    highlightFolderFor(ent.filePath);
                    // ScriptUI の一部環境で数値代入が無視されるため ListItem で選択し revealItem で表示位置も合わせる
                    try {
                        var targetItem = listBox.items[rowIdx];
                        listBox.selection = targetItem;
                        if (typeof listBox.revealItem === "function") {
                            listBox.revealItem(targetItem);
                        }
                    } catch (e) {
                        try { listBox.selection = rowIdx; } catch (e2) { }
                    }
                    break;
                }
            }
        }

        dialog.show();
    }

})();