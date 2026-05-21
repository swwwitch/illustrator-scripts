#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

LinkedImageManager.jsx

ドキュメント内の配置画像（リンク画像）を解析し、
一覧表示・絞り込み・重複管理・再リンクを一元化するユーティリティ。

Illustrator標準機能では行いにくい
「全体把握・状態確認・重複管理・再リンク作業」を効率化することを目的とする。

### 主な機能

・ファイル名／サイズ／配置寸法／スケール／PPI の一覧表示
・リンク状態（正常／リンク切れ／更新が必要）の判定とフィルタリング
・リンク切れ時のファイル名フォールバック取得（XMPメタデータ参照）
・アートボード単位での絞り込みとビュー連動（選択時に移動・ズーム）
・同一ファイルの重複検出・統合表示・使用数カウント
・任意列でのソート（昇順／降順）
・行選択に連動したカンバス上での選択・ズーム表示
・パス表示の最適化（~/ 表示／Dropbox パス短縮／ファイル名表示切替）
・リンク先フォルダの一覧表示／選択フォルダの直接オープン
・単一ファイル／フォルダ単位での再リンク
・リンクファイルのコピーと再リンク（Linksフォルダーへの収集）

### 参考記事

https://note.com/dtp_tranist/n/na66732d2056a

*/

(function () {

    // =========================================
    // バージョン / Version
    // =========================================

    var SCRIPT_VERSION = "v1.2.8";

    // =========================================
    // ユーザー設定 / User configuration
    // =========================================

    // Dropbox のローカルマウントパス接頭辞（ここを自分の環境に書き換えてください）
    var DROPBOX_PREFIX = "/Users/takano/sw Dropbox/takano masahiro/";

    // =========================================
    // ローカライズ / Localization
    // =========================================

    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var lang = getCurrentLang();

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {

        // UI labels

        dialogTitle: { ja: "リンク画像の管理", en: "Linked Image Manager" },
        sortPanelTitle: { ja: "ソート", en: "Sort" },
        sortLabel: { ja: "並び順", en: "Sort by" },
        sortIndex: { ja: "No.", en: "No." },
        sortFileName: { ja: "ファイル名", en: "File Name" },
        sortFileSize: { ja: "サイズ", en: "Size" },
        sortFileCount: { ja: "使用数", en: "Usage Count" },
        sortArtboard: { ja: "アートボード", en: "Artboard" },
        sortWidth: { ja: "幅", en: "Width" },
        sortHeight: { ja: "高さ", en: "Height" },
        sortScale: { ja: "スケール", en: "Scale" },
        sortPpi: { ja: "PPI", en: "PPI" },
        sortStatus: { ja: "ステータス", en: "Status" },
        sortColorSpace: { ja: "カラースペース", en: "Color Space" },
        ascOrder: { ja: "昇順", en: "Ascending" },
        descOrder: { ja: "降順", en: "Descending" },
        optionPanelTitle: { ja: "同一ファイル", en: "Same Files" },
        otherPanelTitle: { ja: "その他", en: "Other" },
        dedupCheck: { ja: "同一ファイルをまとめる", en: "Group Same Files" },
        unitCheck: { ja: "単位を「MB」で統一", en: "Use MB" },
        filterPanelTitle: { ja: "ステータス", en: "Filters" },
        filterOk: { ja: "✓ リンク正常", en: "✓ Link OK" },
        filterBroken: { ja: "⚠ リンク切れ", en: "⚠ Broken Link" },
        filterUpdate: { ja: "⟳ 更新が必要", en: "⟳ Needs Update" },
        artboardPanelTitle: { ja: "アートボード", en: "Artboard" },
        artboardAll: { ja: "すべて", en: "All" },
        artboardFallback: { ja: "アートボード", en: "Artboard" },
        prevArtboardTip: { ja: "前のアートボード", en: "Previous Artboard" },
        nextArtboardTip: { ja: "次のアートボード", en: "Next Artboard" },
        displayOptionPanelTitle: { ja: "表示列", en: "Display Options" },
        displaySize: { ja: "サイズ", en: "Size" },
        displayFileCount: { ja: "使用数を表示", en: "Show Usage Count" },
        displayDimScalePpi: { ja: "サイズ、%、PPI", en: "Dimensions, Scale, PPI" },
        displayColorSpace: { ja: "カラースペース", en: "Color Space" },
        colIndex: { ja: "No.", en: "No." },
        colFileName: { ja: "ファイル名", en: "File Name" },
        colFileSizeMb: { ja: "サイズ(MB)", en: "Size (MB)" },
        colFileSize: { ja: "サイズ", en: "Size" },
        colFileCount: { ja: "使用数", en: "Usage Count" },
        colWidthMm: { ja: "幅(mm)", en: "Width (mm)" },
        colHeightMm: { ja: "高さ(mm)", en: "Height (mm)" },
        colScale: { ja: "スケール", en: "Scale" },
        colPpi: { ja: "PPI", en: "PPI" },
        colArtboards: { ja: "アートボード", en: "Artboards" },
        colColorSpace: { ja: "カラースペース", en: "Color Space" },
        pathPanelTitle: { ja: "ファイルパス", en: "File Path" },
        pathPlaceholder: {
            ja: "リストからアイテムを選択してください",
            en: "Select an item from the list."
        },
        pathHelpTip: { ja: "パスの表示", en: "File path" },
        tildeCheck: { ja: "フルパス", en: "Full path" },
        renameLinkBtn: { ja: "リネーム", en: "Rename" },
        openFileBtn: { ja: "開く", en: "Open" },
        deleteLinkBtn: { ja: "削除", en: "Delete" },
        clipGroupDeleteTitle: { ja: "クリップグループ内の画像", en: "Image inside a clip group" },
        clipGroupDeleteMessage: {
            ja: "選択した画像はクリップグループ内にあります。どのように削除しますか？",
            en: "The selected image is inside a clip group. How would you like to delete it?"
        },
        deleteImageOnlyBtn: { ja: "画像のみを削除", en: "Delete image only" },
        deleteWithClipGroupBtn: { ja: "クリップグループごと削除", en: "Delete entire clip group" },
        copyFileNameBtn: { ja: "ファイル名をコピー", en: "Copy File Name" },
        labelNoExt: { ja: "(なし)", en: "(none)" },
        dropboxCheck: { ja: "Dropboxパスを短縮", en: "Shorten Dropbox Path" },
        fileNameCheck: { ja: "ファイル名", en: "Name" },
        openFolderBtn: { ja: "開く", en: "Open" },
        reloadOneBtnSingle: { ja: "再リンク", en: "Relink Selected" },
        reloadOneBtnBatch: { ja: "一括再リンク", en: "Relink All" },
        linkedFolderListLabel: { ja: "リンクフォルダー一覧", en: "Linked Folders" },
        reloadFolderBtn: { ja: "フォルダーに再リンク", en: "Relink Folder" },
        changeExtensionBtn: { ja: "拡張子の変更", en: "Change Extension" },
        changeExtDialogTitle: { ja: "拡張子の変更", en: "Change Extension" },
        changeExtPanelTitle: { ja: "拡張子", en: "Extension" },
        changeExtFolderPanelTitle: { ja: "変更先のフォルダー", en: "Destination Folder" },
        selectExtensionReferenceFolder: {
            ja: "拡張子変更で参照するフォルダーを選択してください",
            en: "Select the folder to search for files with the new extension"
        },
        chooseExtensionReferenceFolderBtn: { ja: "フォルダー指定", en: "Choose Folder" },
        extensionReferenceFolderPlaceholder: { ja: "参照フォルダー未指定", en: "No reference folder selected" },
        collectLinksBtn: { ja: "リンクを収集", en: "Collect Links" },
        labelItems: { ja: "件", en: "item(s)" },
        labelCopied: { ja: "コピー", en: "Copied" },
        labelSkipped: { ja: "スキップ", en: "Skipped" },
        openLinksPanelBtn: { ja: "［リンク］パネルを開く", en: "Open Links Panel" },
        showOnCanvasCheck: { ja: "選択時にズーム表示", en: "Zoom to Selection" },
        closeBtn: { ja: "閉じる", en: "Close" },
        cancelBtn: { ja: "キャンセル", en: "Cancel" },
        okBtn: { ja: "実行", en: "Run" },
        fileNameUnknown: { ja: "(ファイル名不明)", en: "(Unknown File Name)" },
        statusBroken: { ja: "⚠ リンク切れ", en: "⚠ Broken Link" },
        statusUpdate: { ja: "⟳ 更新が必要", en: "⟳ Needs Update" },
        statusOk: { ja: "✓ リンク正常", en: "✓ Link OK" },
        selectNewLinkFile: { ja: "新しいリンクファイルを選択", en: "Select a new linked file" },
        selectAltFolder: { ja: "代替フォルダを選択", en: "Select a replacement folder" },
        labelSuccess: { ja: "成功", en: "Succeeded" },
        labelFailed: { ja: "失敗", en: "Failed" },
        labelTarget: { ja: "対象", en: "Target" },

        // Alerts

        alertNoDocument: { ja: "ドキュメントが開いていません。", en: "No document is open." },
        alertNoPlacedItems: {
            ja: "配置されているリンク画像が見つかりませんでした。",
            en: "No linked placed images were found."
        },
        alertSelectItem: { ja: "リストからアイテムを選択してください。", en: "Please select an item from the list." },
        alertNoValidPath: { ja: "有効なファイルパスがありません。", en: "No valid file path is available." },
        alertOpenFolderFailed: { ja: "フォルダを開けませんでした：", en: "Could not open the folder: " },
        alertSelectLinkedFolder: { ja: "リンクフォルダを選択してください。", en: "Please select a linked folder." },
        confirmDeleteLinks: {
            ja: "選択したリンクファイルをドキュメントからすべて削除します。",
            en: "The selected file(s) will be removed from the document.?"
        },
        alertDeleteDone: { ja: "削除完了", en: "Delete Complete" },
        alertCopyFileNameDone: { ja: "ファイル名をコピーしました", en: "File name copied to clipboard" },
        alertCopyFileNameFailed: { ja: "ファイル名のコピーに失敗しました", en: "Failed to copy file name" },
        promptNewFileName: {
            ja: "新しいファイル名を入力してください。\n（拡張子 {ext} は自動で保持されます）\n\n現在のファイル名：{name}",
            en: "Enter the new file name.\n(Extension {ext} will be kept automatically)\n\nCurrent file name: {name}"
        },
        alertInvalidFileName: { ja: "ファイル名に / や \\ は使用できません。", en: "File name must not contain / or \\." },
        alertLinkFileNotFound: { ja: "リンクファイルが見つかりません：", en: "Linked file not found: " },
        confirmOverwrite: {
            ja: "同名のファイルが既に存在します。\n上書きしますか？\n\n",
            en: "A file with the same name already exists.\nOverwrite?\n\n"
        },
        alertRenameFailed: { ja: "ファイルのリネームに失敗しました。", en: "Failed to rename the file." },
        alertRenameDone: { ja: "リネームして再リンクしました", en: "Renamed and relinked" },
        alertNameUnchanged: {
            ja: "ファイル名が変更されていません。処理を中断します。",
            en: "The file name has not changed. Aborting."
        },
        confirmBatchRelink: {
            ja: "複数のリンクを一括で再リンクします。よろしいですか？",
            en: "Multiple links will be relinked at once. Continue?"
        },
        alertRelinkDone: { ja: "リンク更新完了", en: "Relink Complete" },
        alertChangeExtDone: { ja: "拡張子を変更しました", en: "Extension changed" },
        alertDocNotSaved: {
            ja: "ドキュメントが保存されていません。保存してからもう一度お試しください。",
            en: "The document has not been saved. Please save the document and try again."
        },
        alertCreateLinksFolderFailed: { ja: "Linksフォルダーを作成できませんでした。", en: "Could not create the Links folder." },
        alertCollectLinksDone: { ja: "リンクのコピーと再リンクが完了しました", en: "Copy and Relink Complete" }
    };

    function L(key) {
        return (LABELS[key] && LABELS[key][lang]) ? LABELS[key][lang] : key;
    }

    function labelText(key) {
        return L(key) + (lang === 'ja' ? '：' : ':');
    }

    // 数値＋単位をローカライズ付きで整形
    function withUnit(value, unitKey) {
        return (lang === 'ja')
            ? (value + L(unitKey))
            : (value + " " + L(unitKey));
    }

    // alert / confirm 用に「ラベル：値+単位」を 1 行で組み立てる
    // 例: kvLine('labelSuccess', 3, 'labelItems') → "成功：3件" / "Succeeded: 3 item(s)"
    function kvLine(labelKey, value, unitKey) {
        var sep = (lang === 'ja' ? '：' : ': ');
        var valueText = unitKey ? withUnit(value, unitKey) : String(value);
        return L(labelKey) + sep + valueText;
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

        var msg = message || L('promptNewFileName')
            .replace("{ext}", ext || L('labelNoExt'))
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

    // JPEG / PNG / PSD ファイルヘッダからピクセル寸法 {width, height} を読み取る。取得不可なら null
    function readImagePixelSize(file) {
        if (!file) return null;
        if (!safeExists(file)) return null;

        var binaryFile = new File(file.fsName);
        binaryFile.encoding = "BINARY";
        if (!binaryFile.open("r")) return null;

        var result = null;
        try {
            var signature = binaryFile.read(8);
            if (signature && signature.length >= 4) {
                var b0 = signature.charCodeAt(0) & 0xFF;
                var b1 = signature.charCodeAt(1) & 0xFF;
                var b2 = signature.charCodeAt(2) & 0xFF;
                var b3 = signature.charCodeAt(3) & 0xFF;

                // PNG: 89 50 4E 47 0D 0A 1A 0A
                if (b0 === 0x89 && b1 === 0x50 && b2 === 0x4E && b3 === 0x47) {
                    // 8(sig) + 4(chunk length) + 4("IHDR") = 16 から幅/高さ
                    binaryFile.seek(16);
                    var ihdrBytes = binaryFile.read(8);
                    if (ihdrBytes && ihdrBytes.length === 8) {
                        result = { width: readU32BE(ihdrBytes, 0), height: readU32BE(ihdrBytes, 4) };
                    }
                }
                // JPEG: FF D8
                else if (b0 === 0xFF && b1 === 0xD8) {
                    binaryFile.seek(2);
                    while (true) {
                        var markerBytes = binaryFile.read(2);
                        if (!markerBytes || markerBytes.length < 2) break;
                        if ((markerBytes.charCodeAt(0) & 0xFF) !== 0xFF) break;
                        var markerCode = markerBytes.charCodeAt(1) & 0xFF;
                        // SOFn（0xC0〜0xCF, ただし 0xC4/0xC8/0xCC は除く）
                        if ((markerCode >= 0xC0 && markerCode <= 0xC3) ||
                            (markerCode >= 0xC5 && markerCode <= 0xC7) ||
                            (markerCode >= 0xC9 && markerCode <= 0xCB) ||
                            (markerCode >= 0xCD && markerCode <= 0xCF)) {
                            binaryFile.read(3); // length(2) + precision(1)
                            var heightBytes = binaryFile.read(2);
                            var widthBytes = binaryFile.read(2);
                            if (heightBytes && widthBytes && heightBytes.length === 2 && widthBytes.length === 2) {
                                result = { width: readU16BE(widthBytes, 0), height: readU16BE(heightBytes, 0) };
                            }
                            break;
                        }
                        var segmentLengthBytes = binaryFile.read(2);
                        if (!segmentLengthBytes || segmentLengthBytes.length < 2) break;
                        var segmentLength = readU16BE(segmentLengthBytes, 0);
                        binaryFile.seek(binaryFile.tell() + segmentLength - 2);
                    }
                }
                // PSD: "8BPS"
                else if (b0 === 0x38 && b1 === 0x42 && b2 === 0x50 && b3 === 0x53) {
                    // PSD header: sig(4) + ver(2) + reserved(6) + channels(2) + rows(4) + cols(4)
                    // → rows は offset 14、cols は offset 18
                    binaryFile.seek(14);
                    var psdDimsBytes = binaryFile.read(8);
                    if (psdDimsBytes && psdDimsBytes.length === 8) {
                        result = { width: readU32BE(psdDimsBytes, 4), height: readU32BE(psdDimsBytes, 0) };
                    }
                }
            }
        } catch (e) { }
        binaryFile.close();
        return result;
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

    // 画像ファイルヘッダからカラーモードと ICC プロファイル名を読み取り、"RGB（Adobe RGB (1998)）" 形式で返す
    function readImageColorSpace(file) {
        if (!file) return null;
        if (!safeExists(file)) return null;

        var binaryFile = new File(file.fsName);
        binaryFile.encoding = "BINARY";
        if (!binaryFile.open("r")) return null;

        var colorMode = "";
        var iccDesc = "";

        try {
            var signature = binaryFile.read(8);
            if (!signature || signature.length < 4) { binaryFile.close(); return null; }
            var b0 = signature.charCodeAt(0) & 0xFF;
            var b1 = signature.charCodeAt(1) & 0xFF;
            var b2 = signature.charCodeAt(2) & 0xFF;
            var b3 = signature.charCodeAt(3) & 0xFF;

            // PNG: 89 50 4E 47
            if (b0 === 0x89 && b1 === 0x50 && b2 === 0x4E && b3 === 0x47) {
                binaryFile.seek(16);
                var ihdrBytes = binaryFile.read(10);
                if (ihdrBytes && ihdrBytes.length === 10) {
                    var colorType = ihdrBytes.charCodeAt(9) & 0xFF;
                    if (colorType === 0 || colorType === 4) colorMode = "Grayscale";
                    else if (colorType === 3) colorMode = "Indexed";
                    else colorMode = "RGB";
                }
                // iCCP / sRGB チャンクを探す
                binaryFile.seek(8 + 4 + 4 + 13 + 4); // PNG sig + IHDR length + "IHDR" + IHDR data + CRC
                for (var pi = 0; pi < 32; pi++) {
                    var pngChunkLengthBytes = binaryFile.read(4);
                    if (!pngChunkLengthBytes || pngChunkLengthBytes.length < 4) break;
                    var chunkLength = readU32BE(pngChunkLengthBytes, 0);
                    var chunkType = binaryFile.read(4);
                    if (!chunkType || chunkType.length < 4) break;
                    if (chunkType === "iCCP") {
                        var chunkData = binaryFile.read(chunkLength);
                        if (chunkData) {
                            var nullIndex = chunkData.indexOf("\0");
                            if (nullIndex > 0) {
                                iccDesc = chunkData.substring(0, nullIndex);
                            }
                        }
                        break;
                    } else if (chunkType === "sRGB") {
                        iccDesc = "sRGB IEC61966-2.1";
                        break;
                    } else if (chunkType === "IDAT" || chunkType === "IEND") {
                        break;
                    } else {
                        binaryFile.seek(binaryFile.tell() + chunkLength + 4); // data + CRC
                    }
                }
            }
            // JPEG: FF D8
            else if (b0 === 0xFF && b1 === 0xD8) {
                binaryFile.seek(2);
                var iccPieces = {};
                var iccTotalPieces = 0;
                var componentCount = 0;
                for (var ji = 0; ji < 64; ji++) {
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

                    var isSof = (markerCode >= 0xC0 && markerCode <= 0xCF) && markerCode !== 0xC4 && markerCode !== 0xC8 && markerCode !== 0xCC;
                    if (isSof) {
                        var sofBytes = binaryFile.read(6);
                        if (sofBytes && sofBytes.length === 6) {
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
                if (componentCount === 1) colorMode = "Grayscale";
                else if (componentCount === 4) colorMode = "CMYK";
                else if (componentCount > 0) colorMode = "RGB";
                if (iccTotalPieces > 0) {
                    var iccData = "";
                    var allPiecesReceived = true;
                    for (var pieceIdx = 1; pieceIdx <= iccTotalPieces; pieceIdx++) {
                        if (!iccPieces[pieceIdx]) { allPiecesReceived = false; break; }
                        iccData += iccPieces[pieceIdx];
                    }
                    if (allPiecesReceived && iccData.length > 0) iccDesc = readIccDesc(iccData);
                }
            }
            // PSD: "8BPS"
            else if (b0 === 0x38 && b1 === 0x42 && b2 === 0x50 && b3 === 0x53) {
                binaryFile.seek(24);
                var modeBytes = binaryFile.read(2);
                if (modeBytes && modeBytes.length === 2) {
                    var modeCode = readU16BE(modeBytes, 0);
                    if (modeCode === 0) colorMode = "Bitmap";
                    else if (modeCode === 1) colorMode = "Grayscale";
                    else if (modeCode === 2) colorMode = "Indexed";
                    else if (modeCode === 3) colorMode = "RGB";
                    else if (modeCode === 4) colorMode = "CMYK";
                    else if (modeCode === 7) colorMode = "Multichannel";
                    else if (modeCode === 8) colorMode = "Duotone";
                    else if (modeCode === 9) colorMode = "Lab";
                }
                binaryFile.seek(26);
                var colorModeDataLengthBytes = binaryFile.read(4);
                if (colorModeDataLengthBytes && colorModeDataLengthBytes.length === 4) {
                    var colorModeDataLength = readU32BE(colorModeDataLengthBytes, 0);
                    var imageResourceStart = 30 + colorModeDataLength;
                    binaryFile.seek(imageResourceStart);
                    var imageResourceLengthBytes = binaryFile.read(4);
                    if (imageResourceLengthBytes && imageResourceLengthBytes.length === 4) {
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
                                if (iccProfileData && iccProfileData.length > 0) iccDesc = readIccDesc(iccProfileData);
                                break;
                            } else {
                                binaryFile.seek(binaryFile.tell() + paddedDataSize);
                            }
                            // 進行していない（seek が無効）なら壊れたファイル。無限ループ回避
                            if (binaryFile.tell() <= positionBefore) break;
                        }
                    }
                }
            }
        } catch (e) { }
        binaryFile.close();

        if (!colorMode) return null;
        if (iccDesc) return colorMode + "（" + iccDesc + "）";
        return colorMode;
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
                status: L('statusBroken'),
                statusIcon: "⚠",
                isLinkOk: false
            };
        }

        if (isLinkOutOfDate(placedFile, xmpLastModifyDate)) {
            return {
                statusCode: "update",
                status: L('statusUpdate'),
                statusIcon: "⟳",
                isLinkOk: false
            };
        }

        return {
            statusCode: "ok",
            status: L('statusOk'),
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
        var fileName = defaultName || L('fileNameUnknown');

        var placedFile = tryGet(function () { return placedItem.file; }, null);
        if (placedFile) {
            fileName = safeProp(placedFile, "name", fileName);
        }

        if (fileName === L('fileNameUnknown') && fallbackName) {
            fileName = fallbackName;
        }

        if (fileName === L('fileNameUnknown')) {
            fileName = safeProp(placedItem, "name", fileName);
        }

        return fileName;
    }

    // placedItems からリンク情報を収集し、フラットリストと重複排除済みリストを返す
    function collectLinkInfo(doc, placedItems) {
        var linkInfoList = [];

        var xmpRefs = collectXmpLinkedRefs(doc);

        for (var i = 0; i < placedItems.length; i++) {
            var item = placedItems[i];
            var xmpRef = xmpRefs[i] || null;
            var filePath = "---";
            var fileName = L('fileNameUnknown');
            var fileSize = "---";
            var fileSizeBytes = -1; // ソート用（-1 は不明）
            var status = L('statusBroken');

            var linkedFile = tryGet(function () { return item.file; }, null);

            var isLinkOk = false;
            var statusCode = "broken";
            var statusIcon = "⚠";

            if (linkedFile) {
                fileName = safeProp(linkedFile, "name", fileName);
                filePath = safeProp(linkedFile, "fsName", filePath);
                var resolvedStatus = resolveLinkStatus(linkedFile, item, doc, xmpRef ? xmpRef.lastModifyDate : "");
                statusCode = resolvedStatus.statusCode;
                status = resolvedStatus.status;
                statusIcon = resolvedStatus.statusIcon;
                isLinkOk = resolvedStatus.isLinkOk;
                if (safeExists(linkedFile)) {
                    var byteLength = tryGet(function () { return linkedFile.length; }, -1);
                    if (byteLength >= 0) {
                        fileSizeBytes = byteLength;
                        fileSize = formatFileSize(fileSizeBytes);
                    }
                }
            }

            fileName = getPlacedItemFileName(item, xmpRef ? xmpRef.fileName : "", fileName);

            if (fileName === L('fileNameUnknown') && filePath !== "---") {
                var derivedName = pathBaseName(filePath);
                if (derivedName) fileName = derivedName;
            }

            var artboardNum = getArtboardNumber(item, doc);

            var dimensions = getDimensionsMm(item);
            var widthText = (dimensions.widthMm !== null) ? dimensions.widthMm.toFixed(1) : "---";
            var heightText = (dimensions.heightMm !== null) ? dimensions.heightMm.toFixed(1) : "---";

            var scaleInfo = getScaleInfo(item);

            var ppi = null;
            var pixelSize = null;
            var colorSpace = "";
            if (linkedFile) {
                pixelSize = tryGet(function () { return readImagePixelSize(linkedFile); }, null);
                ppi = getEffectivePPI(item, pixelSize);
                colorSpace = tryGet(function () { return readImageColorSpace(linkedFile); }, "") || "";
            }
            var ppiText = (ppi !== null) ? String(ppi) : "---";

            linkInfoList.push({
                itemIndex: i,
                index: i + 1,
                fileName: fileName,
                filePath: filePath,
                fileSize: fileSize,
                fileSizeBytes: fileSizeBytes,
                status: status,
                statusIcon: statusIcon,
                statusCode: statusCode,
                isLinkOk: isLinkOk,
                artboardNum: artboardNum,
                artboards: (artboardNum !== null) ? String(artboardNum) : "-",
                widthMm: dimensions.widthMm,
                heightMm: dimensions.heightMm,
                widthText: widthText,
                heightText: heightText,
                scalePct: scaleInfo.scalePct,
                scaleText: scaleInfo.scaleText,
                ppi: ppi,
                ppiText: ppiText,
                colorSpace: colorSpace,
                itemIndices: [i]
            });
        }

        // 同一ファイルの配置数をカウント
        var countMap = {};
        for (var ci = 0; ci < linkInfoList.length; ci++) {
            var countInfo = linkInfoList[ci];
            var countKey = (countInfo.filePath !== "---") ? countInfo.filePath : countInfo.fileName;
            countMap[countKey] = (countMap[countKey] || 0) + 1;
        }
        for (var ai = 0; ai < linkInfoList.length; ai++) {
            var assignInfo = linkInfoList[ai];
            var assignKey = (assignInfo.filePath !== "---") ? assignInfo.filePath : assignInfo.fileName;
            assignInfo.fileCount = countMap[assignKey];
        }

        // 重複排除
        var uniqueList = [];
        var keyToEntry = {};
        for (var di = 0; di < linkInfoList.length; di++) {
            var dedupInfo = linkInfoList[di];
            var dedupKey = (dedupInfo.filePath !== "---") ? dedupInfo.filePath : dedupInfo.fileName;
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

        return { linkInfoList: linkInfoList, uniqueList: uniqueList };
    }

    // ドキュメントが開いているか確認
    if (app.documents.length === 0) {
        alert(L('alertNoDocument'));
        return;
    }

    var doc = app.activeDocument;
    var placedItems = doc.placedItems; // 配置画像（リンク画像）を取得

    // 配置画像が存在するか確認
    if (placedItems.length === 0) {
        alert(L('alertNoPlacedItems'));
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
        var PANEL_MARGINS = [10, 20, 10, 10]; // 全パネル共通のマージン
        var MAIN_LISTBOX_SIZE = [500, 190]; // メイン一覧のサイズ
        var FOLDER_LISTBOX_SIZE = [500, 120]; // リンクフォルダ一覧のサイズ

        var dlg = new Window("dialog", L('dialogTitle') + " " + SCRIPT_VERSION);
        dlg.orientation = "column";
        dlg.alignChildren = ["fill", "top"];
        dlg.preferredSize.width = 500;

        // ソート・フィルタ前の元データ（重複チェックボックスで切り替え）
        var sourceEntries = uniqueFileEntries;
        // 実際にリストに表示しているエントリ群（フィルタ後）— listBox.onChange などから参照
        var filteredEntries = sourceEntries;
        // プログラムから選択を入れる際に、カンバス側の再選択＆ズームを 1 回だけ抑制するフラグ
        var suppressCanvasOnce = false;
        // リスト再構築後に選択を復元するための配置アイテム index
        var pendingSelectionItemIndex = -1;

        // --- 上部 2 カラム：左（ソート + オプション） / 右（その他） ---
        var topRow = dlg.add("group");
        topRow.orientation = "row";
        topRow.alignChildren = ["fill", "top"];

        // 左カラム：ソート + オプション を縦に
        var leftCol = topRow.add("group");
        leftCol.orientation = "column";
        leftCol.alignChildren = ["fill", "top"];

        // ソート
        var sortPanel = leftCol.add("panel", undefined, L('sortPanelTitle'));
        sortPanel.orientation = "column";
        sortPanel.alignChildren = ["left", "top"];
        sortPanel.margins = PANEL_MARGINS;
        var sortKeyRow = sortPanel.add("group");
        sortKeyRow.orientation = "row";
        sortKeyRow.alignChildren = ["left", "center"];
        sortKeyRow.add("statictext", undefined, labelText('sortLabel'));
        var sortDropdown = sortKeyRow.add("dropdownlist", undefined, []);
        // ソート対象は列の表示状態と連動（rebuildSortDropdown で動的に構築）
        var currentVisibleSpecs = [];
        var orderRow = sortPanel.add("group");
        orderRow.orientation = "row";
        orderRow.alignChildren = ["left", "center"];
        var ascRadio = orderRow.add("radiobutton", undefined, L('ascOrder'));
        var descRadio = orderRow.add("radiobutton", undefined, L('descOrder'));
        descRadio.value = true;

        // オプション（重複 / ファイルサイズの単位を統一）
        var optPanel = leftCol.add("panel", undefined, L('optionPanelTitle'));
        optPanel.orientation = "column";
        optPanel.alignChildren = ["left", "top"];
        optPanel.margins = PANEL_MARGINS;
        var dedupCheck = optPanel.add("checkbox", undefined, L('dedupCheck'));
        dedupCheck.value = true; // ON：同一ファイルは 1 行にまとめる
        dedupCheck.helpTip = (lang === 'ja')
            ? "ON：同じリンクファイルを1行にまとめ、再リンクは“すべて一括”で実行されます（複数変更に注意）。\nOFF：配置ごとに個別表示され、再リンクも1件ずつ実行されます。"
            : "ON: Group same linked files into one row; relink applies to ALL placements (affects multiple items).\nOFF: Each placement is listed separately; relink applies individually.";
        var countColCheck = optPanel.add("checkbox", undefined, L('displayFileCount'));
        countColCheck.value = true;

        // 右カラム：その他（表示オプション + フィルター / アートボード）
        var otherPanel = topRow.add("group");
        otherPanel.orientation = "column";
        otherPanel.alignChildren = ["fill", "top"];
        // otherPanel.margins = PANEL_MARGINS;

        var otherTopRow = otherPanel.add("group");
        otherTopRow.orientation = "row";
        otherTopRow.alignChildren = ["fill", "top"];

        // 左：表示オプション（列の表示/非表示を切替）
        var optionPanel = otherTopRow.add("panel", undefined, L('displayOptionPanelTitle'));
        optionPanel.orientation = "column";
        optionPanel.alignChildren = ["left", "top"];
        optionPanel.margins = PANEL_MARGINS;
        var sizeRow = optionPanel.add("group");
        sizeRow.orientation = "row";
        sizeRow.alignChildren = ["left", "center"];
        var sizeColCheck = sizeRow.add("checkbox", undefined, L('displaySize'));

        var unitCheck = optionPanel.add("checkbox", undefined, L('unitCheck'));

        var dimScalePpiCheck = optionPanel.add("checkbox", undefined, L('displayDimScalePpi'));
        var colorSpaceColCheck = optionPanel.add("checkbox", undefined, L('displayColorSpace'));
        sizeColCheck.value = false;
        unitCheck.value = true; // ON：全行 MB で統一表示、OFF：B/KB/MB/GB の自動単位
        unitCheck.enabled = sizeColCheck.value;
        dimScalePpiCheck.value = false;
        colorSpaceColCheck.value = false;

        // 右：フィルター
        var filterPanel = otherTopRow.add("panel", undefined, L('filterPanelTitle'));
        filterPanel.orientation = "column";
        filterPanel.alignChildren = ["left", "top"];
        filterPanel.margins = PANEL_MARGINS;
        var statusGroup = filterPanel.add("group");
        statusGroup.orientation = "column";
        statusGroup.alignChildren = ["left", "top"];
        var okCheck = statusGroup.add("checkbox", undefined, L('filterOk'));
        var brokenCheck = statusGroup.add("checkbox", undefined, L('filterBroken'));
        var updateCheck = statusGroup.add("checkbox", undefined, L('filterUpdate'));
        okCheck.value = true;
        brokenCheck.value = true;
        updateCheck.value = true;

        // 下段：アートボード（右カラム内で左右貫通）
        var abPanel = otherPanel.add("panel", undefined, L('artboardPanelTitle'));
        abPanel.orientation = "row";
        abPanel.alignChildren = ["left", "center"];
        abPanel.alignment = ["fill", "top"];
        abPanel.margins = PANEL_MARGINS;
        // 配置画像のあるアートボード番号を収集（無い番号はドロップダウンでディム表示）
        var artboardsWithImages = {};
        for (var pi = 0; pi < allPlacementEntries.length; pi++) {
            var abNum = allPlacementEntries[pi].artboardNum;
            if (abNum !== null) artboardsWithImages[abNum] = true;
        }

        var artboardDropdownItems = [L('artboardAll')];
        for (var artboardIndex = 0; artboardIndex < doc.artboards.length; artboardIndex++) {
            var artboardName = "";
            artboardName = safeProp(doc.artboards[artboardIndex], "name", artboardName);
            artboardDropdownItems.push((artboardIndex + 1) + "：" + (artboardName || L('artboardFallback') + (artboardIndex + 1)));
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
        abPrevBtn.helpTip = L('prevArtboardTip');
        var abNextBtn = abPanel.add("button", undefined, "▶");
        abNextBtn.preferredSize = [30, 22];
        abNextBtn.helpTip = L('nextArtboardTip');

        var enabledCount = 0;
        for (var i = 1; i < abFilterDropdown.items.length; i++) {
            if (abFilterDropdown.items[i].enabled) enabledCount++;
        }
        if (enabledCount <= 1) {
            abPrevBtn.enabled = false;
            abNextBtn.enabled = false;
        }

        // --- リストボックス（アートボード列の有無で再生成するため、コンテナ経由で配置） ---
        var listHolder = dlg.add("group");
        listHolder.orientation = "column";
        listHolder.alignChildren = ["fill", "top"];

        var listBox = null;
        // 各列の有効/無効・キー・タイトル・幅を表示オプションに応じて返す
        function getColumnSpec() {
            var cols = [];
            cols.push({ key: "statusIcon", title: "", width: 30 });
            cols.push({ key: "fileName", title: L('colFileName'), width: 210 });
            if (sizeColCheck.value) {
                cols.push({
                    key: "fileSize",
                    title: unitCheck.value ? L('colFileSizeMb') : L('colFileSize'),
                    width: 65
                });
            }
            if (countColCheck.value) {
                cols.push({ key: "fileCount", title: L('colFileCount'), width: 45 });
            }
            if (dimScalePpiCheck.value) {
                cols.push({ key: "widthText", title: L('colWidthMm'), width: 60 });
                cols.push({ key: "heightText", title: L('colHeightMm'), width: 60 });
                cols.push({ key: "scaleText", title: L('colScale'), width: 60 });
                cols.push({ key: "ppiText", title: L('colPpi'), width: 50 });
            }
            if (colorSpaceColCheck.value) {
                cols.push({ key: "colorSpace", title: L('colColorSpace'), width: 160 });
            }
            // Show/hide Artboard column based on artboard dropdown selection
            var shouldShowArtboardColumn = !abFilterDropdown.selection || abFilterDropdown.selection.index === 0;
            if (shouldShowArtboardColumn) {
                cols.push({ key: "artboards", title: L('colArtboards'), width: 70 });
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

            // 行選択時：パス表示を更新。「カンバス上で表示」が ON なら該当画像を選択＆フィット
            listBox.onChange = function () {
                if (listBox.selection === null) return;
                var info = filteredEntries[listBox.selection.index];
                selectedEntry = info;
                selectedFilePath = info.filePath;
                pathStaticText.text = buildDisplayedPath(info.filePath);
                updateActionButtonStates();
                highlightFolderFor(info.filePath);
                if (suppressCanvasOnce) {
                    suppressCanvasOnce = false;
                    return;
                }
                if (showOnCanvasCheck.value) selectPlacedItemsOnCanvas(info);
            };
        }
        createListBox();

        // ソートキー仕様：listbox で非表示の項目はドロップダウンから自動的に除外される
        var SORT_SPECS = [
            { key: 'fileName', labelKey: 'sortFileName', preferDesc: false, getValue: function (i) { return i.fileName; }, isVisible: function () { return true; } },
            { key: 'fileSize', labelKey: 'sortFileSize', preferDesc: true, getValue: function (i) { return i.fileSizeBytes; }, isVisible: function () { return sizeColCheck.value; } },
            { key: 'fileCount', labelKey: 'sortFileCount', preferDesc: true, getValue: function (i) { return i.fileCount; }, isVisible: function () { return countColCheck.value; } },
            { key: 'artboard', labelKey: 'sortArtboard', preferDesc: false, getValue: function (i) { return (i.artboardNum === null) ? null : i.artboardNum; }, isVisible: function () { return !abFilterDropdown.selection || abFilterDropdown.selection.index === 0; } },
            { key: 'width', labelKey: 'sortWidth', preferDesc: true, getValue: function (i) { return i.widthMm; }, isVisible: function () { return dimScalePpiCheck.value; } },
            { key: 'height', labelKey: 'sortHeight', preferDesc: true, getValue: function (i) { return i.heightMm; }, isVisible: function () { return dimScalePpiCheck.value; } },
            { key: 'scale', labelKey: 'sortScale', preferDesc: true, getValue: function (i) { return i.scalePct; }, isVisible: function () { return dimScalePpiCheck.value; } },
            { key: 'ppi', labelKey: 'sortPpi', preferDesc: true, getValue: function (i) { return i.ppi; }, isVisible: function () { return dimScalePpiCheck.value; } },
            { key: 'status', labelKey: 'sortStatus', preferDesc: false, getValue: function (i) { return i.status; }, isVisible: function () { return true; } },
            { key: 'colorSpace', labelKey: 'sortColorSpace', preferDesc: false, getValue: function (i) { return i.colorSpace || ""; }, isVisible: function () { return colorSpaceColCheck.value; } }
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

        // ソートキーの値取り出し（currentVisibleSpecs のインデックス経由）
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
                var aMissing = isMissing(valA), bMissing = isMissing(valB);
                if (aMissing && bMissing) return 0;
                if (aMissing) return 1;
                if (bMissing) return -1;
                if (valA < valB) return desc ? 1 : -1;
                if (valA > valB) return desc ? -1 : 1;
                return 0;
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

            restoreListSelectionByItemIndex(selectionItemIndexToRestore);
        }

        // リスト再構築後、配置アイテム index を基準に該当行を再選択する
        function restoreListSelectionByItemIndex(itemIndexToRestore) {
            if (itemIndexToRestore < 0 || !filteredEntries || filteredEntries.length === 0) return;

            for (var i = 0; i < filteredEntries.length; i++) {
                var entry = filteredEntries[i];
                if (!entry || !entry.itemIndices) continue;

                for (var j = 0; j < entry.itemIndices.length; j++) {
                    if (entry.itemIndices[j] === itemIndexToRestore) {
                        suppressCanvasOnce = true;
                        selectedEntry = entry;
                        selectedFilePath = entry.filePath;
                        pathStaticText.text = buildDisplayedPath(entry.filePath);
                        updateActionButtonStates();
                        highlightFolderFor(entry.filePath);
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
            rebuildSortDropdown();
            dlg.layout.layout(true);
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
            rebuildSortDropdown();
            dlg.layout.layout(true);
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

        // ソートドロップダウンの初期構築：使用数をデフォルト選択
        rebuildSortDropdown();
        for (var initSi = 0; initSi < currentVisibleSpecs.length; initSi++) {
            if (currentVisibleSpecs[initSi].key === 'fileCount') {
                sortDropdown.selection = initSi;
                break;
            }
        }

        // 初期表示
        rebuildList();

        // --- パスパネル直上：選択時の挙動オプション ---
        var viewOptRow = dlg.add("group");
        viewOptRow.orientation = "row";
        viewOptRow.alignment = ["fill", "top"];
        viewOptRow.alignChildren = ["left", "center"];
        var showOnCanvasCheck = viewOptRow.add("checkbox", undefined, L('showOnCanvasCheck'));
        showOnCanvasCheck.value = true; // ON：行選択でカンバス上の該当画像をフィット表示

        // --- パスパネル ---
        var pathPanel = dlg.add("panel", undefined, L('pathPanelTitle'));
        pathPanel.orientation = "column";
        pathPanel.alignChildren = ["fill", "top"];
        pathPanel.margins = PANEL_MARGINS;

        // 1 行目：選択中の親フォルダ表示
        var pathRow = pathPanel.add("group");
        pathRow.orientation = "row";
        pathRow.alignChildren = ["fill", "center"];

        var pathStaticText = pathRow.add(
            "statictext", undefined, L('pathPlaceholder'),
            { multiline: true }
        );
        pathStaticText.alignment = ["fill", "fill"];
        pathStaticText.preferredSize = [450, 20];
        pathStaticText.helpTip = L('pathHelpTip');

        // 2 行目：整形オプション（アクションボタンはさらに下の独立行に配置）
        var pathOptRow = pathPanel.add("group");
        pathOptRow.orientation = "row";
        pathOptRow.alignment = "fill";
        pathOptRow.alignChildren = ["fill", "center"];

        var pathOptLeft = pathOptRow.add("group");
        pathOptLeft.orientation = "row";
        pathOptLeft.alignChildren = ["left", "center"];
        pathOptLeft.alignment = ["left", "center"];
        var tildeCheck = pathOptLeft.add("checkbox", undefined, L('tildeCheck'));
        var dropboxCheck = pathOptLeft.add("checkbox", undefined, L('dropboxCheck'));
        var fileNameCheck = pathOptLeft.add("checkbox", undefined, L('fileNameCheck'));
        // フルパス表示（ON）時は ~ 短縮を無効化する（意味は内部的に反転）
        tildeCheck.value = false;
        dropboxCheck.value = true;
        fileNameCheck.value = false; // ON：パスにファイル名まで含める。OFF：親フォルダまで表示

        var pathOptSpacer = pathOptRow.add("group");
        pathOptSpacer.alignment = ["fill", "fill"];

        var pathOptRight = pathOptRow.add("group");
        pathOptRight.orientation = "row";
        pathOptRight.alignChildren = ["right", "center"];
        pathOptRight.alignment = ["right", "center"];

        // 行選択中の絶対パス＋エントリ（ボタンや整形切替で使用）
        var selectedFilePath = "";
        var selectedEntry = null;

        // selectedEntry 未選択ガード付きで onClick ハンドラを生成
        function requireSelectedEntry(handler) {
            return function () {
                if (!selectedEntry) {
                    alert(L('alertSelectItem'));
                    return;
                }
                return handler.apply(this, arguments);
            };
        }

        // ファイル名ON時も ~ 短縮や Dropbox 短縮は適用する（フルパスONのときだけ生パスを表示）
        function buildDisplayedPath(absPath) {
            if (!absPath || absPath === "---") return absPath;

            var pathToFormat = fileNameCheck.value ? absPath : toFolderOnly(absPath);
            return formatDisplayPath(pathToFormat, !tildeCheck.value, dropboxCheck.value);
        }
        function updatePathDisplay() {
            if (!selectedFilePath) return;
            pathStaticText.text = buildDisplayedPath(selectedFilePath);
        }
        function onPathOptionChange() {
            updatePathDisplay();
            populateFoldersList();
        }
        // Dropbox 対応が ON のときは、~ 置換が上書きされる前提なので ~ チェックをディム
        function updateTildeEnable() {
            tildeCheck.enabled = !dropboxCheck.value;
        }
        tildeCheck.onClick = onPathOptionChange;
        dropboxCheck.onClick = function () {
            updateTildeEnable();
            onPathOptionChange();
        };
        fileNameCheck.onClick = function () {
            updatePathDisplay();
            pathPanel.layout.layout(true);
            dlg.layout.layout(true);
        }; // フォルダ一覧には影響しない
        updateTildeEnable(); // 初期化

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
            var extDlg = new Window("dialog", L('changeExtDialogTitle'));
            extDlg.orientation = "column";
            extDlg.alignChildren = ["fill", "top"];

            var folderPanel = extDlg.add("panel", undefined, L('changeExtFolderPanelTitle'));
            folderPanel.orientation = "column";
            folderPanel.alignChildren = ["fill", "top"];
            folderPanel.margins = PANEL_MARGINS;

            var chooseFolderBtn = folderPanel.add("button", undefined, L('chooseExtensionReferenceFolderBtn'));
            chooseFolderBtn.alignment = ["left", "top"];

            var folderLabel = folderPanel.add("statictext", undefined, L('extensionReferenceFolderPlaceholder'));
            folderLabel.alignment = ["fill", "top"];
            folderLabel.preferredSize = [300, 20];

            var extPanel = extDlg.add("panel", undefined, L('changeExtPanelTitle'));
            extPanel.orientation = "column";
            extPanel.alignChildren = ["fill", "top"];
            extPanel.margins = PANEL_MARGINS;

            var referenceFolder = null;
            var okBtn = null;

            chooseFolderBtn.onClick = function () {
                var selectedFolder = Folder.selectDialog(L('selectExtensionReferenceFolder'));
                if (!selectedFolder) return;

                referenceFolder = selectedFolder;
                folderLabel.text = selectedFolder.fsName;
                if (okBtn) okBtn.enabled = true;
                safeRelayout(extDlg);
            };

            var radios = [];

            var radioRow = extPanel.add("group");
            radioRow.orientation = "row";
            radioRow.alignment = ["fill", "top"];
            radioRow.alignChildren = ["fill", "top"];

            var colLeft = radioRow.add("group");
            colLeft.orientation = "column";
            colLeft.alignment = ["fill", "top"];
            colLeft.alignChildren = ["left", "top"];
            colLeft.preferredSize.width = 110;

            var colCenter = radioRow.add("group");
            colCenter.orientation = "column";
            colCenter.alignment = ["fill", "top"];
            colCenter.alignChildren = ["left", "top"];
            colCenter.preferredSize.width = 110;

            var colRight = radioRow.add("group");
            colRight.orientation = "column";
            colRight.alignment = ["fill", "top"];
            colRight.alignChildren = ["left", "top"];
            colRight.preferredSize.width = 110;

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

            // 左カラム
            addRadio(colLeft, "png", ".png");
            addRadio(colLeft, "jpg / jpeg", ".jpg", ".jpeg");
            addRadio(colLeft, "psd", ".psd", null, true);

            // 中央カラム
            addRadio(colCenter, "tiff", ".tif", ".tiff");
            addRadio(colCenter, "webp", ".webp");
            addRadio(colCenter, "avif", ".avif");

            // 右カラム
            addRadio(colRight, "gif", ".gif");
            addRadio(colRight, "ai", ".ai");
            addRadio(colRight, "pdf", ".pdf");

            var btnRow = extDlg.add("group");
            btnRow.alignment = ["right", "top"];

            var cancelBtn = btnRow.add("button", undefined, L('cancelBtn'));
            okBtn = btnRow.add("button", undefined, L('okBtn'));
            okBtn.enabled = false;

            cancelBtn.onClick = function () {
                extDlg.close(0);
            };

            okBtn.onClick = function () {
                if (!referenceFolder) return;
                extDlg.close(1);
            };

            if (extDlg.show() !== 1) return null;

            for (var radioIdx = 0; radioIdx < radios.length; radioIdx++) {
                if (radios[radioIdx].ui.value) {
                    return {
                        referenceFolder: referenceFolder,
                        primaryExt: radios[radioIdx].ext,
                        fallbackExt: radios[radioIdx].alt
                    };
                }
            }

            return null;
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

        // アクションボタンは pathPanel 直下の独立行に 3 カラム（左：ファイル名コピー / 中央：spacer / 右：リネーム・再リンク）
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

        var openFileBtn = actionBtnRight.add("button", undefined, L('openFileBtn'));
        openFileBtn.preferredSize = [50, 24];
        openFileBtn.onClick = requireSelectedEntry(function () {
            var absPath = selectedEntry.filePath;
            if (!absPath || absPath === "---") {
                alert(L('alertNoValidPath'));
                return;
            }
            var fileToOpen = new File(absPath);
            if (!fileToOpen.exists) {
                alert(L('alertLinkFileNotFound') + absPath);
                return;
            }
            fileToOpen.execute();
        });

        var deleteLinkBtn = actionBtnRight.add("button", undefined, L('deleteLinkBtn'));
        deleteLinkBtn.preferredSize = [50, 24];
        deleteLinkBtn.onClick = requireSelectedEntry(function () {
            var indices = selectedEntry.itemIndices || [];
            if (indices.length === 0) return;

            // インデックスを伴う削除は collection の再採番で破綻するため、先に PageItem 参照を集める
            // 各対象について、内包しているクリップグループ（あれば）も併せて保持
            var refs = [];
            for (var di = 0; di < indices.length; di++) {
                try {
                    var placedRef = placedItems[indices[di]];
                    refs.push({ placed: placedRef, clipGroup: findEnclosingClipGroup(placedRef) });
                } catch (eRef) { }
            }

            var hasClipped = false;
            for (var dc = 0; dc < refs.length; dc++) {
                if (refs[dc].clipGroup) { hasClipped = true; break; }
            }

            // 確認ダイアログ：
            // ・クリップグループを含まない → 通常の確認ダイアログ（はい/いいえ）
            // ・含む                      → 確認と削除モード選択を 1 つにまとめた 3 ボタンダイアログ
            var clipMode = 'image';
            if (hasClipped) {
                clipMode = askDeleteModeWithConfirm(indices.length);
                if (clipMode === null) return;
            } else {
                var confirmMessage = L('confirmDeleteLinks') + "\n" +
                    kvLine('labelTarget', indices.length, 'labelItems');
                if (!confirm(confirmMessage)) return;
            }

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

            // 選択は削除済みなので解除しておく（refreshFromDoc の itemIndex 復元に巻き込まれないように）
            selectedEntry = null;
            selectedFilePath = "";

            app.redraw();
            refreshFromDoc();
            updateActionButtonStates();
            alert(
                L('alertDeleteDone') + "\n" +
                kvLine('labelSuccess', success, 'labelItems') + "\n" +
                kvLine('labelFailed', failed, 'labelItems')
            );
        });

        // クリップグループ内の画像が含まれる場合の、確認 + 削除モード選択を兼ねたダイアログ
        // 戻り値: 'image' / 'group' / null（キャンセル）
        function askDeleteModeWithConfirm(targetCount) {
            var deleteDialog = new Window("dialog", L('clipGroupDeleteTitle'));
            deleteDialog.orientation = "column";
            deleteDialog.alignChildren = ["fill", "top"];
            deleteDialog.margins = 16;

            var msgText = L('confirmDeleteLinks') + "\n" +
                kvLine('labelTarget', targetCount, 'labelItems') + "\n\n" +
                L('clipGroupDeleteMessage');
            var msg = deleteDialog.add("statictext", undefined, msgText, { multiline: true });
            msg.preferredSize.width = 360;

            var btnRow = deleteDialog.add("group");
            btnRow.orientation = "row";
            btnRow.alignment = ["right", "center"];
            var cancelBtn = btnRow.add("button", undefined, L('cancelBtn'), { name: "cancel" });
            var imageOnlyBtn = btnRow.add("button", undefined, L('deleteImageOnlyBtn'));
            var withGroupBtn = btnRow.add("button", undefined, L('deleteWithClipGroupBtn'), { name: "ok" });

            cancelBtn.onClick = function () { deleteDialog.close(0); };
            imageOnlyBtn.onClick = function () { deleteDialog.close(1); };
            withGroupBtn.onClick = function () { deleteDialog.close(2); };

            var dialogResult = deleteDialog.show();
            if (dialogResult === 1) return 'image';
            if (dialogResult === 2) return 'group';
            return null;
        }

        var renameLinkBtn = actionBtnRight.add("button", undefined, L('renameLinkBtn'));
        renameLinkBtn.preferredSize = [70, 24];
        renameLinkBtn.onClick = requireSelectedEntry(function () {
            var absPath = selectedEntry.filePath;
            if (!absPath || absPath === "---") {
                alert(L('alertNoValidPath'));
                return;
            }
            var oldFile = new File(absPath);
            if (!oldFile.exists) {
                alert(L('alertLinkFileNotFound') + absPath);
                return;
            }

            var oldName = getRealFileName(oldFile);
            var oldFolder = oldFile.parent;

            var newName = promptNewFileName(oldName);
            if (newName === null) return;
            if (newName === oldName) {
                alert(L('alertNameUnchanged'));
                return;
            }
            if (/[\/\\]/.test(newName)) {
                alert(L('alertInvalidFileName'));
                return;
            }

            var newFile = new File(oldFolder.fsName + "/" + newName);
            if (newFile.exists) {
                // 大小文字だけの違いは macOS の case-insensitive FS で同一ファイルを指すため
                // 削除対象にしない（rename は OS 側で正しく処理される）
                var caseOnlyDiff = (newFile.fsName.toLowerCase() === oldFile.fsName.toLowerCase());
                if (!caseOnlyDiff) {
                    if (!confirm(L('confirmOverwrite') + newFile.fsName)) return;
                    // File.rename は既存ファイルがあると失敗するので、先に削除してから rename する
                    var removed = tryGet(function () { return newFile.remove(); }, false);
                    if (!removed) {
                        alert(L('alertRenameFailed'));
                        return;
                    }
                }
            }

            // 物理リネーム（失敗時は再リンクしない）
            var renamed = tryGet(function () { return oldFile.rename(newName); }, false);
            if (!renamed) {
                alert(L('alertRenameFailed'));
                return;
            }

            var res = relinkIndicesTo(selectedEntry.itemIndices, newFile);
            app.redraw();
            refreshFromDoc();
            alert(
                L('alertRenameDone') + "\n" +
                oldName + " → " + newName + "\n" +
                kvLine('labelSuccess', res.success, 'labelItems') + "\n" +
                kvLine('labelFailed', res.failed, 'labelItems')
            );
        });
        var copyFileNameBtn = actionBtnLeft.add("button", undefined, L('copyFileNameBtn'));
        copyFileNameBtn.onClick = requireSelectedEntry(function () {
            var name = selectedEntry.fileName || "";
            if (copyTextToClipboard(name)) {
                alert(L('alertCopyFileNameDone') + "\n" + name);
            } else {
                alert(L('alertCopyFileNameFailed'));
            }
        });
        var reloadOneBtn = actionBtnRight.add("button", undefined, L('reloadOneBtnSingle'));
        reloadOneBtn.preferredSize = [80, 24];

        reloadOneBtn.onClick = requireSelectedEntry(function () {
            var relinkCount = (selectedEntry.itemIndices && selectedEntry.itemIndices.length)
                ? selectedEntry.itemIndices.length
                : 0;
            if (relinkCount > 1) {
                var confirmMessage = L('confirmBatchRelink') + "\n" +
                    kvLine('labelTarget', relinkCount, 'labelItems');
                if (!confirm(confirmMessage)) return;
            }

            var picked = File.openDialog(L('selectNewLinkFile'));
            if (!picked) return;
            var res = relinkIndicesTo(selectedEntry.itemIndices, picked);
            app.redraw();
            refreshFromDoc();
            alert(
                L('alertRelinkDone') + "\n" +
                kvLine('labelSuccess', res.success, 'labelItems') + "\n" +
                kvLine('labelFailed', res.failed, 'labelItems')
            );
        });

        // 「同一ファイルをまとめる」の状態に応じて、再リンクボタンの意味を明示する
        // ただし、まとめ対象の使用数が 1 の場合は「一括」ではなく単数ラベルを使う
        function updateRelinkButtonLabel() {
            var placementCount = (selectedEntry && selectedEntry.itemIndices) ? selectedEntry.itemIndices.length : 0;
            var useBatchLabel = dedupCheck.value && placementCount > 1;
            var nextLabel = useBatchLabel ? L('reloadOneBtnBatch') : L('reloadOneBtnSingle');
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

            safeRelayout(actionBtnRight, actionBtnRow, pathPanel, dlg);
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

        var folderCountLabel = pathPanel.add("statictext", undefined, L('linkedFolderListLabel') + " (" + withUnit(linkedFolderPaths.length, 'labelItems') + ")");
        var foldersListBox = pathPanel.add("listbox", undefined, [], { multiselect: false });
        foldersListBox.preferredSize = FOLDER_LISTBOX_SIZE;
        foldersListBox.alignment = ["fill", "fill"];

        function populateFoldersList() {
            foldersListBox.removeAll();
            for (var ii = 0; ii < linkedFolderPaths.length; ii++) {
                foldersListBox.add("item",
                    formatDisplayPath(linkedFolderPaths[ii], !tildeCheck.value, dropboxCheck.value)
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
                pathStaticText.text = selectedFilePath ? buildDisplayedPath(selectedFilePath) : L('pathPlaceholder');
                updateActionButtonStates();
            }

            rebuildFolderList();
            folderCountLabel.text = L('linkedFolderListLabel') + " (" + withUnit(linkedFolderPaths.length, 'labelItems') + ")";
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
                alert(L('alertOpenFolderFailed') + e.message);
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

        var openFolderBtn = folderActionLeft.add("button", undefined, L('openFolderBtn'));
        openFolderBtn.onClick = function () {
            if (foldersListBox.selection === null) {
                alert(L('alertSelectLinkedFolder'));
                return;
            }
            try {
                var folderPath = linkedFolderPaths[foldersListBox.selection.index];
                var folderToOpen = new Folder(folderPath);
                if (folderToOpen.exists) folderToOpen.execute();
            } catch (e) {
                alert(L('alertOpenFolderFailed') + e.message);
            }
        };

        var folderActionSpacer = folderActionRow.add("group");
        folderActionSpacer.alignment = ["fill", "fill"];

        var folderActionRight = folderActionRow.add("group");
        folderActionRight.orientation = "row";
        folderActionRight.alignChildren = ["right", "center"];
        folderActionRight.alignment = ["right", "center"];

        var reloadFolderBtn = folderActionLeft.add("button", undefined, L('reloadFolderBtn'));
        reloadFolderBtn.onClick = function () {
            if (foldersListBox.selection === null) {
                alert(L('alertSelectLinkedFolder'));
                return;
            }
            var oldFolder = linkedFolderPaths[foldersListBox.selection.index];
            var newFolder = Folder.selectDialog(L('selectAltFolder'));
            if (!newFolder) return;
            var res = relinkFolder(oldFolder, newFolder);
            app.redraw();
            refreshFromDoc();
            alert(
                L('alertRelinkDone') + "\n" +
                kvLine('labelTarget', res.total, 'labelItems') + "\n" +
                kvLine('labelSuccess', res.success, 'labelItems') + "\n" +
                kvLine('labelFailed', res.failed, 'labelItems') + (lang === 'ja' ? '（同名ファイルなし）' : ' (same-name file not found)')
            );
        };
        reloadFolderBtn.enabled = false;
        openFolderBtn.enabled = false;

        var changeExtensionBtn = folderActionLeft.add("button", undefined, L('changeExtensionBtn'));

        changeExtensionBtn.onClick = function () {
            if (foldersListBox.selection === null) {
                alert(L('alertSelectLinkedFolder'));
                return;
            }

            var sourceFolderPath = normalizeFolderPathForCompare(
                linkedFolderPaths[foldersListBox.selection.index]
            );

            // 1. 参照フォルダー＋拡張子を選択
            var extPrefs = showChangeExtensionDialog();
            if (!extPrefs) return;

            var referenceFolder = extPrefs.referenceFolder;

            // 2. 実行
            var success = 0;
            var failed = 0;
            var skipped = 0;
            var total = 0;

            for (var i = 0; i < allPlacementEntries.length; i++) {
                var entry = allPlacementEntries[i];
                if (!entry || !entry.filePath || entry.filePath === "---") continue;

                var entryFolderPath = normalizeFolderPathForCompare(pathParent(entry.filePath));

                if (entryFolderPath !== sourceFolderPath) continue;

                total++;

                var placedItem = placedItems[entry.itemIndex];
                var sourceFileName = entry.fileName || pathBaseName(entry.filePath);
                var baseName = splitFileName(String(sourceFileName)).base;

                if (!baseName) {
                    failed++;
                    continue;
                }

                var replacementFile = findReplacementFileByExtension(
                    referenceFolder,
                    baseName,
                    extPrefs.primaryExt,
                    extPrefs.fallbackExt
                );

                if (!replacementFile) {
                    failed++;
                    continue;
                }

                // 置換先が現在のリンクと同一ファイルなら差し替え不要
                try {
                    var currentPath = placedItem.file ? placedItem.file.fsName : null;
                    if (currentPath && currentPath === replacementFile.fsName) {
                        skipped++;
                        continue;
                    }
                } catch (e) { }

                try {
                    placedItem.file = replacementFile;
                    success++;
                } catch (e) {
                    failed++;
                }
            }

            app.redraw();
            refreshFromDoc();

            alert(
                L('alertChangeExtDone') + "\n" +
                kvLine('labelTarget', total, 'labelItems') + "\n" +
                kvLine('labelSuccess', success, 'labelItems') + "\n" +
                kvLine('labelSkipped', skipped, 'labelItems') + "\n" +
                kvLine('labelFailed', failed, 'labelItems')
            );
        };

        changeExtensionBtn.enabled = false;

        var collectLinksBtn = folderActionRight.add("button", undefined, L('collectLinksBtn'));
        collectLinksBtn.onClick = function () {
            var docFile = tryGet(function () { return doc.fullName; }, null);
            if (!docFile || !docFile.parent || !docFile.parent.exists) {
                alert(L('alertDocNotSaved'));
                return;
            }

            var linksFolder = new Folder(docFile.parent.fsName + "/Links");
            if (!linksFolder.exists) {
                if (!linksFolder.create()) {
                    alert(L('alertCreateLinksFolderFailed'));
                    return;
                }
            }

            var copied = 0, skipped = 0, failed = 0, total = 0;
            for (var k = 0; k < allPlacementEntries.length; k++) {
                var ent = allPlacementEntries[k];
                if (!ent.filePath || ent.filePath === "---") continue;
                total++;
                try {
                    var srcFile = new File(ent.filePath);
                    if (!srcFile.exists) { failed++; continue; }

                    var destFile = new File(linksFolder.fsName + "/" + srcFile.name);
                    // 既に収集先に同名ファイルがある場合はコピーをスキップ（同一パスなら再リンクも不要）
                    if (destFile.fsName === srcFile.fsName) {
                        skipped++;
                        continue;
                    }
                    if (destFile.exists) {
                        skipped++;
                    } else {
                        if (!srcFile.copy(destFile.fsName)) { failed++; continue; }
                        copied++;
                    }

                    placedItems[ent.itemIndex].file = destFile;
                } catch (e) {
                    failed++;
                }
            }

            app.redraw();
            refreshFromDoc();

            alert(
                L('alertCollectLinksDone') + "\n" +
                kvLine('labelTarget', total, 'labelItems') + "\n" +
                kvLine('labelCopied', copied, 'labelItems') + "\n" +
                kvLine('labelSkipped', skipped, 'labelItems') + "\n" +
                kvLine('labelFailed', failed, 'labelItems')
            );
        };

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
        var btnGroup = dlg.add("group");
        btnGroup.orientation = "row";
        btnGroup.alignment = "fill";
        btnGroup.alignChildren = ["fill", "center"];

        var openLinksPanelBtn = btnGroup.add("button", undefined, L('openLinksPanelBtn'));
        openLinksPanelBtn.alignment = ["left", "center"];
        openLinksPanelBtn.onClick = function () {
            try { app.executeMenuCommand('Adobe LinkPalette Menu Item'); } catch (e) { }
            dlg.close();
        };

        var spacer = btnGroup.add("group");
        spacer.alignment = ["fill", "fill"];

        var closeBtn = btnGroup.add("button", undefined, L('closeBtn'), { name: "cancel" });
        closeBtn.alignment = ["right", "center"];

        // 「閉じる」ボタン
        closeBtn.onClick = function () {
            dlg.close();
        };

        // 実行前に選択していた PlacedItem があれば、対応行を初期選択（カンバス側は触らない）
        if (typeof preselectedItemIndex === "number" && preselectedItemIndex >= 0) {
            for (var pi = 0; pi < filteredEntries.length; pi++) {
                var ent = filteredEntries[pi];
                var found = false;
                for (var pj = 0; pj < ent.itemIndices.length; pj++) {
                    if (ent.itemIndices[pj] === preselectedItemIndex) { found = true; break; }
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
                        var targetItem = listBox.items[pi];
                        listBox.selection = targetItem;
                        if (typeof listBox.revealItem === "function") {
                            listBox.revealItem(targetItem);
                        }
                    } catch (e) {
                        try { listBox.selection = pi; } catch (e2) { }
                    }
                    break;
                }
            }
        }

        dlg.show();
    }

})();