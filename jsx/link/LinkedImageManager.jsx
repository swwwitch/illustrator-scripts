#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

(function () {

    /*
        ============================================================
        LinkedImageManager.jsx
        ------------------------------------------------------------
        ドキュメント内の配置画像（リンク画像）を解析し、
        一覧表示・絞り込み・重複管理・再リンクを一元化するユーティリティ。

        Illustrator標準機能では行いにくい
        「全体把握・状態確認・重複管理・再リンク作業」を一元的に行えるようにすることを目的とする。

        ### 主な機能

        ・ファイル名／サイズ／配置寸法／拡大率／PPI の一覧表示
        ・リンク状態（正常／リンク切れ／更新が必要）の判定とフィルタリング
        ・リンク切れ時のファイル名フォールバック取得（XMPメタデータ参照）
        ・配置先アートボードの特定とアートボード単位での絞り込み
        ・アートボード選択に連動したビュー移動・ズーム表示
        ・同一ファイルの重複検出・統合表示・使用数カウント
        ・任意列でのソート（昇順／降順）
        ・行選択に連動したカンバス上での選択・ズーム表示
        ・パス表示の最適化（~/ 表示／Dropbox パス短縮）
        ・リンク先フォルダの表示とフォルダ一覧からのアクセス
        ・単一ファイル／フォルダ単位での再リンク
        ・リンクファイルのコピーと再リンク（Linksフォルダーへの収集）

        ============================================================
    */


    // =========================================
    // バージョンとローカライズ / Version & Localization
    // =========================================

    var SCRIPT_VERSION = "v1.0";

    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var lang = getCurrentLang();

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialogTitle: {
            ja: "リンク画像の管理",
            en: "Linked Image Manager"
        },
        alertNoDocument: {
            ja: "ドキュメントが開いていません。",
            en: "No document is open."
        },
        alertNoPlacedItems: {
            ja: "配置されているリンク画像が見つかりませんでした。",
            en: "No linked placed images were found."
        },
        sortPanelTitle: {
            ja: "ソート",
            en: "Sort"
        },
        sortLabel: {
            ja: "並び順",
            en: "Sort by"
        },
        sortIndex: {
            ja: "No.",
            en: "No."
        },
        sortFileName: {
            ja: "ファイル名",
            en: "File Name"
        },
        sortFileSize: {
            ja: "サイズ",
            en: "Size"
        },
        sortFileCount: {
            ja: "使用数",
            en: "Usage Count"
        },
        sortArtboard: {
            ja: "アートボード",
            en: "Artboard"
        },
        sortWidth: {
            ja: "幅",
            en: "Width"
        },
        sortHeight: {
            ja: "高さ",
            en: "Height"
        },
        sortScale: {
            ja: "スケール",
            en: "Scale"
        },
        sortPpi: {
            ja: "PPI",
            en: "PPI"
        },
        sortStatus: {
            ja: "ステータス",
            en: "Status"
        },
        ascOrder: {
            ja: "昇順",
            en: "Ascending"
        },
        descOrder: {
            ja: "降順",
            en: "Descending"
        },
        optionPanelTitle: {
            ja: "オプション",
            en: "Options"
        },
        otherPanelTitle: {
            ja: "その他",
            en: "Other"
        },
        dedupCheck: {
            ja: "同一ファイルをまとめる",
            en: "Group Same Files"
        },
        unitCheck: {
            ja: "サイズをMBで統一",
            en: "Show File Size in MB"
        },
        filterPanelTitle: {
            ja: "フィルター",
            en: "Filters"
        },
        filterOk: {
            ja: "✓ リンク正常",
            en: "✓ Link OK"
        },
        filterBroken: {
            ja: "⚠ リンク切れ",
            en: "⚠ Broken Link"
        },
        filterUpdate: {
            ja: "⟳ 更新が必要",
            en: "⟳ Needs Update"
        },
        artboardPanelTitle: {
            ja: "アートボード",
            en: "Artboard"
        },
        artboardAll: {
            ja: "すべて",
            en: "All"
        },
        artboardFallback: {
            ja: "アートボード",
            en: "Artboard"
        },
        prevArtboardTip: {
            ja: "前のアートボード",
            en: "Previous Artboard"
        },
        nextArtboardTip: {
            ja: "次のアートボード",
            en: "Next Artboard"
        },
        displayOptionPanelTitle: {
            ja: "表示列",
            en: "Display Options"
        },
        displaySize: {
            ja: "サイズ",
            en: "Size"
        },
        displayFileCount: {
            ja: "使用数",
            en: "Usage Count"
        },
        displayDimScalePpi: {
            ja: "サイズ、%、PPI",
            en: "Dimensions, Scale, PPI"
        },
        colIndex: {
            ja: "No.",
            en: "No."
        },
        colFileName: {
            ja: "ファイル名",
            en: "File Name"
        },
        colFileSizeMb: {
            ja: "サイズ(MB)",
            en: "Size (MB)"
        },
        colFileSize: {
            ja: "サイズ",
            en: "Size"
        },
        colFileCount: {
            ja: "使用数",
            en: "Usage Count"
        },
        colWidthMm: {
            ja: "幅(mm)",
            en: "Width (mm)"
        },
        colHeightMm: {
            ja: "高さ(mm)",
            en: "Height (mm)"
        },
        colScale: {
            ja: "スケール",
            en: "Scale"
        },
        colPpi: {
            ja: "PPI",
            en: "PPI"
        },
        colArtboards: {
            ja: "アートボード",
            en: "Artboards"
        },
        pathPanelTitle: {
            ja: "パス",
            en: "Path"
        },
        pathPlaceholder: {
            ja: "リストからアイテムを選択してください",
            en: "Select an item from the list."
        },
        tildeCheck: {
            ja: "~",
            en: "~"
        },
        dropboxCheck: {
            ja: "Dropboxパスを短縮",
            en: "Shorten Dropbox Path"
        },
        fileNameCheck: {
            ja: "ファイル名",
            en: "Name"
        },
        openFolderBtn: {
            ja: "開く",
            en: "Open"
        },
        reloadOneBtn: {
            ja: "再リンク",
            en: "Relink Selected File"
        },
        linkedFolderListLabel: {
            ja: "リンクフォルダ一覧",
            en: "Linked Folders"
        },
        reloadFolderBtn: {
            ja: "選択フォルダーへのリンクを変更",
            en: "Relink Folder Links"
        },
        collectLinksBtn: {
            ja: "収集（リンクをコピー＋再リンク）",
            en: "Collect (Copy + Relink)"
        },
        alertDocNotSaved: {
            ja: "ドキュメントが保存されていません。保存してからもう一度お試しください。",
            en: "The document has not been saved. Please save the document and try again."
        },
        alertCreateLinksFolderFailed: {
            ja: "Linksフォルダーを作成できませんでした。",
            en: "Could not create the Links folder."
        },
        alertCollectLinksDone: {
            ja: "リンクのコピーと再リンクが完了しました",
            en: "Copy and Relink Complete"
        },
        labelCopied: {
            ja: "コピー",
            en: "Copied"
        },
        labelSkipped: {
            ja: "スキップ",
            en: "Skipped"
        },
        showOnCanvasCheck: {
            ja: "選択時にズーム表示",
            en: "Zoom to Selection"
        },
        closeBtn: {
            ja: "閉じる",
            en: "Close"
        },
        fileNameUnknown: {
            ja: "(ファイル名不明)",
            en: "(Unknown File Name)"
        },
        statusBroken: {
            ja: "⚠ リンク切れ",
            en: "⚠ Broken Link"
        },
        statusUpdate: {
            ja: "⟳ 更新が必要",
            en: "⟳ Needs Update"
        },
        statusOk: {
            ja: "✓ リンク正常",
            en: "✓ Link OK"
        },
        alertNoValidPath: {
            ja: "有効なファイルパスがありません。",
            en: "No valid file path is available."
        },
        alertOpenFolderFailed: {
            ja: "フォルダを開けませんでした：",
            en: "Could not open the folder: "
        },
        alertSelectItem: {
            ja: "リストからアイテムを選択してください。",
            en: "Please select an item from the list."
        },
        selectNewLinkFile: {
            ja: "新しいリンクファイルを選択",
            en: "Select a new linked file"
        },
        alertRelinkDone: {
            ja: "リンク更新完了",
            en: "Relink Complete"
        },
        labelSuccess: {
            ja: "成功",
            en: "Succeeded"
        },
        labelFailed: {
            ja: "失敗",
            en: "Failed"
        },
        alertSelectLinkedFolder: {
            ja: "リンクフォルダを選択してください。",
            en: "Please select a linked folder."
        },
        selectAltFolder: {
            ja: "代替フォルダを選択",
            en: "Select a replacement folder"
        },
        labelTarget: {
            ja: "対象",
            en: "Target"
        }
    };

    function L(key) {
        return (LABELS[key] && LABELS[key][lang]) ? LABELS[key][lang] : key;
    }

    function labelText(key) {
        return L(key) + (lang === 'ja' ? '：' : ':');
    }

    // Dropbox のローカルマウントパス接頭辞（ここを自分の環境に書き換えてください）
    var DROPBOX_PREFIX = "/Users/takano/sw Dropbox/takano masahiro/";

    // ファイルパスからファイル名を除いた親フォルダ（末尾セパレータ付き）を返す
    function toFolderOnly(path) {
        if (!path || path === "---") return path;
        var sep = Math.max(path.lastIndexOf("/"), path.lastIndexOf("\\"));
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
        var sel = doc.selection;
        if (!sel || sel.length === 0) return;

        var first = sel[0].visibleBounds; // [left, top, right, bottom]
        var minL = first[0], maxT = first[1], maxR = first[2], minB = first[3];
        for (var i = 1; i < sel.length; i++) {
            var b = sel[i].visibleBounds;
            if (b[0] < minL) minL = b[0];
            if (b[1] > maxT) maxT = b[1];
            if (b[2] > maxR) maxR = b[2];
            if (b[3] < minB) minB = b[3];
        }

        var cx = (minL + maxR) / 2;
        var cy = (maxT + minB) / 2;
        var w = Math.max(maxR - minL, 1);
        var h = Math.max(maxT - minB, 1);

        var view = doc.views[0];
        var vb = view.bounds; // 現在の表示領域（ドキュメント座標）
        var vw = Math.abs(vb[2] - vb[0]);
        var vh = Math.abs(vb[1] - vb[3]);

        var margin = 1.2; // 上下左右に 20% の余白
        var zx = view.zoom * vw / (w * margin);
        var zy = view.zoom * vh / (h * margin);
        var newZoom = Math.min(zx, zy);

        // Illustrator のズーム上限/下限に合わせて安全側にクランプ
        if (newZoom < 0.03125) newZoom = 0.03125; //  3.125%
        if (newZoom > 64) newZoom = 64;      // 6400%

        view.zoom = newZoom;
        view.centerPoint = [cx, cy];
    }

    // バイナリ文字列から big-endian の 16bit / 32bit 整数を取り出す
    function readU16BE(s, o) {
        return ((s.charCodeAt(o) & 0xFF) << 8) | (s.charCodeAt(o + 1) & 0xFF);
    }
    function readU32BE(s, o) {
        return ((s.charCodeAt(o) & 0xFF) * 16777216) +
            ((s.charCodeAt(o + 1) & 0xFF) << 16) +
            ((s.charCodeAt(o + 2) & 0xFF) << 8) +
            (s.charCodeAt(o + 3) & 0xFF);
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

    // オブジェクトのプロパティ値を安全に取得
    function safeProp(obj, key, fallback) {
        return tryGet(function () {
            var value = obj[key];
            return (value !== undefined && value !== null && value !== "") ? value : fallback;
        }, fallback);
    }

    // JPEG / PNG ファイルヘッダからピクセル寸法 {width, height} を読み取る。取得不可なら null
    function readImagePixelSize(file) {
        if (!file) return null;
        if (!safeExists(file)) return null;

        var f = new File(file.fsName);
        f.encoding = "BINARY";
        if (!f.open("r")) return null;

        var result = null;
        try {
            var sig = f.read(8);
            if (sig && sig.length >= 4) {
                var b0 = sig.charCodeAt(0) & 0xFF;
                var b1 = sig.charCodeAt(1) & 0xFF;
                var b2 = sig.charCodeAt(2) & 0xFF;
                var b3 = sig.charCodeAt(3) & 0xFF;

                // PNG: 89 50 4E 47 0D 0A 1A 0A
                if (b0 === 0x89 && b1 === 0x50 && b2 === 0x4E && b3 === 0x47) {
                    // 8(sig) + 4(chunk length) + 4("IHDR") = 16 から幅/高さ
                    f.seek(16);
                    var ihdr = f.read(8);
                    if (ihdr && ihdr.length === 8) {
                        result = { width: readU32BE(ihdr, 0), height: readU32BE(ihdr, 4) };
                    }
                }
                // JPEG: FF D8
                else if (b0 === 0xFF && b1 === 0xD8) {
                    f.seek(2);
                    while (true) {
                        var mark = f.read(2);
                        if (!mark || mark.length < 2) break;
                        if ((mark.charCodeAt(0) & 0xFF) !== 0xFF) break;
                        var m = mark.charCodeAt(1) & 0xFF;
                        // SOFn（0xC0〜0xCF, ただし 0xC4/0xC8/0xCC は除く）
                        if ((m >= 0xC0 && m <= 0xC3) ||
                            (m >= 0xC5 && m <= 0xC7) ||
                            (m >= 0xC9 && m <= 0xCB) ||
                            (m >= 0xCD && m <= 0xCF)) {
                            f.read(3); // length(2) + precision(1)
                            var hh = f.read(2);
                            var ww = f.read(2);
                            if (hh && ww && hh.length === 2 && ww.length === 2) {
                                result = { width: readU16BE(ww, 0), height: readU16BE(hh, 0) };
                            }
                            break;
                        }
                        var lenBytes = f.read(2);
                        if (!lenBytes || lenBytes.length < 2) break;
                        var len = readU16BE(lenBytes, 0);
                        f.seek(f.tell() + len - 2);
                    }
                }
            }
        } catch (e) { }
        f.close();
        return result;
    }

    // 配置サイズ（pt）とピクセル寸法から実効 PPI を算出
    function getEffectivePPI(item, pixelSize) {
        if (!pixelSize) return null;
        var w = tryGet(function () { return item.width; }, null);
        var h = tryGet(function () { return item.height; }, null);
        if (!w || !h) return null;
        var ppiX = pixelSize.width * 72 / w;
        var ppiY = pixelSize.height * 72 / h;
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
        var m = tryGet(function () { return item.matrix; }, null);
        if (!m) return { scalePct: null, scaleText: "---" };
        var sx = Math.sqrt(m.mValueA * m.mValueA + m.mValueB * m.mValueB);
        var sy = Math.sqrt(m.mValueC * m.mValueC + m.mValueD * m.mValueD);
        var sxPct = sx * 100;
        var syPct = sy * 100;
        var avg = (sxPct + syPct) / 2;
        var text = (Math.abs(sxPct - syPct) < 0.1)
            ? sxPct.toFixed(1) + "%"
            : sxPct.toFixed(1) + "% × " + syPct.toFixed(1) + "%";
        return { scalePct: avg, scaleText: text };
    }

    // 配置サイズ（pt）を mm に変換して {widthMm, heightMm} を返す。不明なら {null, null}
    function getDimensionsMm(item) {
        var w = tryGet(function () { return item.width; }, null);   // pt
        var h = tryGet(function () { return item.height; }, null);
        if (!w || !h) return { widthMm: null, heightMm: null };
        return {
            widthMm: w * 25.4 / 72,
            heightMm: h * 25.4 / 72
        };
    }

    // 指定アイテムの中心が含まれるアートボード番号（1 始まり）を返す。無ければ null
    function getArtboardNumber(item, doc) {
        var b = tryGet(function () { return item.visibleBounds; }, null); // [left, top, right, bottom]
        if (!b) return null;
        var cx = (b[0] + b[2]) / 2;
        var cy = (b[1] + b[3]) / 2;
        var artboards = tryGet(function () { return doc.artboards; }, null);
        if (!artboards) return null;
        for (var i = 0; i < artboards.length; i++) {
            var r = tryGet(function () { return artboards[i].artboardRect; }, null);
            if (!r) continue;
            var xMin = Math.min(r[0], r[2]);
            var xMax = Math.max(r[0], r[2]);
            var yMin = Math.min(r[1], r[3]);
            var yMax = Math.max(r[1], r[3]);
            if (cx >= xMin && cx <= xMax && cy >= yMin && cy <= yMax) {
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

            var f = tryGet(function () { return item.file; }, null);

            var isLinkOk = false;
            var statusCode = "broken";
            var statusIcon = "⚠";

            if (f) {
                fileName = safeProp(f, "name", fileName);
                filePath = safeProp(f, "fsName", filePath);
                var resolvedStatus = resolveLinkStatus(f, item, doc, xmpRef ? xmpRef.lastModifyDate : "");
                statusCode = resolvedStatus.statusCode;
                status = resolvedStatus.status;
                statusIcon = resolvedStatus.statusIcon;
                isLinkOk = resolvedStatus.isLinkOk;
                if (safeExists(f)) {
                    var len = tryGet(function () { return f.length; }, -1);
                    if (len >= 0) {
                        fileSizeBytes = len;
                        fileSize = formatFileSize(fileSizeBytes);
                    }
                }
            }

            fileName = getPlacedItemFileName(item, xmpRef ? xmpRef.fileName : "", fileName);

            if (fileName === L('fileNameUnknown') && filePath !== "---") {
                var slash = filePath.lastIndexOf("/");
                var bslash = filePath.lastIndexOf("\\");
                var cut = Math.max(slash, bslash);
                if (cut >= 0 && cut < filePath.length - 1) {
                    fileName = filePath.substring(cut + 1);
                }
            }

            var artboardNum = getArtboardNumber(item, doc);

            var dim = getDimensionsMm(item);
            var widthText = (dim.widthMm !== null) ? dim.widthMm.toFixed(1) : "---";
            var heightText = (dim.heightMm !== null) ? dim.heightMm.toFixed(1) : "---";

            var scaleInfo = getScaleInfo(item);

            var ppi = null;
            var pxSize = null;
            if (f) {
                pxSize = tryGet(function () { return readImagePixelSize(f); }, null);
                ppi = getEffectivePPI(item, pxSize);
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
                widthMm: dim.widthMm,
                heightMm: dim.heightMm,
                widthText: widthText,
                heightText: heightText,
                scalePct: scaleInfo.scalePct,
                scaleText: scaleInfo.scaleText,
                ppi: ppi,
                ppiText: ppiText,
                itemIndices: [i]
            });
        }

        // 同一ファイルの配置数をカウント
        var countMap = {};
        for (var k = 0; k < linkInfoList.length; k++) {
            var kInfo = linkInfoList[k];
            var kKey = (kInfo.filePath !== "---") ? kInfo.filePath : kInfo.fileName;
            countMap[kKey] = (countMap[kKey] || 0) + 1;
        }
        for (var m = 0; m < linkInfoList.length; m++) {
            var mInfo = linkInfoList[m];
            var mKey = (mInfo.filePath !== "---") ? mInfo.filePath : mInfo.fileName;
            mInfo.fileCount = countMap[mKey];
        }

        // 重複排除
        var uniqueList = [];
        var keyToEntry = {};
        for (var n = 0; n < linkInfoList.length; n++) {
            var nInfo = linkInfoList[n];
            var nKey = (nInfo.filePath !== "---") ? nInfo.filePath : nInfo.fileName;
            if (keyToEntry[nKey]) {
                keyToEntry[nKey].itemIndices.push(nInfo.itemIndex);
                if (nInfo.artboardNum !== null) {
                    keyToEntry[nKey].artboardSet[nInfo.artboardNum] = true;
                }
            } else {
                var abSet = {};
                if (nInfo.artboardNum !== null) abSet[nInfo.artboardNum] = true;
                var entry = {
                    index: uniqueList.length + 1,
                    fileName: nInfo.fileName,
                    filePath: nInfo.filePath,
                    fileSize: nInfo.fileSize,
                    fileSizeBytes: nInfo.fileSizeBytes,
                    status: nInfo.status,
                    statusIcon: nInfo.statusIcon,
                    statusCode: nInfo.statusCode,
                    isLinkOk: nInfo.isLinkOk,
                    fileCount: nInfo.fileCount,
                    artboardNum: nInfo.artboardNum,
                    artboardSet: abSet,
                    widthMm: nInfo.widthMm,
                    heightMm: nInfo.heightMm,
                    widthText: nInfo.widthText,
                    heightText: nInfo.heightText,
                    scalePct: nInfo.scalePct,
                    scaleText: nInfo.scaleText,
                    ppi: nInfo.ppi,
                    ppiText: nInfo.ppiText,
                    itemIndices: [nInfo.itemIndex]
                };
                keyToEntry[nKey] = entry;
                uniqueList.push(entry);
            }
        }

        for (var q = 0; q < uniqueList.length; q++) {
            var u = uniqueList[q];
            var nums = [];
            for (var abKey in u.artboardSet) {
                if (u.artboardSet.hasOwnProperty(abKey)) nums.push(parseInt(abKey, 10));
            }
            nums.sort(function (a, b) { return a - b; });
            u.artboards = (nums.length > 0) ? nums.join(", ") : "-";
            u.artboardNum = (nums.length > 0) ? nums[0] : null;
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
    var preselectedItemIndex = -1;
    var selection = tryGet(function () { return doc.selection; }, null);
    if (selection && selection.length === 1) {
        var selItem = selection[0];
        if (selItem && selItem.typename === "PlacedItem") {
            for (var si = 0; si < placedItems.length; si++) {
                if (placedItems[si] === selItem) {
                    preselectedItemIndex = si;
                    break;
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
        var sortDropdown = sortKeyRow.add("dropdownlist", undefined,
            [L('sortIndex'), L('sortFileName'), L('sortFileSize'), L('sortFileCount'), L('sortArtboard'), L('sortWidth'), L('sortHeight'), L('sortScale'), L('sortPpi'), L('sortStatus')]
        );
        sortDropdown.selection = 2; // サイズ
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
        var unitCheck = optPanel.add("checkbox", undefined, L('unitCheck'));
        unitCheck.value = true; // ON：全行 MB で統一表示、OFF：B/KB/MB/GB の自動単位

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
        var sizeColCheck = optionPanel.add("checkbox", undefined, L('displaySize'));
        var countColCheck = optionPanel.add("checkbox", undefined, L('displayFileCount'));
        var dimScalePpiCheck = optionPanel.add("checkbox", undefined, L('displayDimScalePpi'));
        sizeColCheck.value = true;
        countColCheck.value = true;
        dimScalePpiCheck.value = false;

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

        var abPrevBtn = abPanel.add("button", undefined, "↑");
        abPrevBtn.preferredSize = [30, 22];
        abPrevBtn.helpTip = L('prevArtboardTip');
        var abNextBtn = abPanel.add("button", undefined, "↓");
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
            var cols = getColumnSpec();
            var titles = [], widths = [];
            for (var cc = 0; cc < cols.length; cc++) {
                titles.push(cols[cc].title);
                widths.push(cols[cc].width);
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
                pathText.text = buildDisplayedPath(info.filePath);
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

        // ソートキーの値取り出し
        function getSortValue(info, sortBy) {
            switch (sortBy) {
                case 0: return info.index;
                case 1: return info.fileName;
                case 2: return info.fileSizeBytes;
                case 3: return info.fileCount;
                case 4: return (info.artboardNum === null) ? null : info.artboardNum;
                case 5: return info.widthMm;
                case 6: return info.heightMm;
                case 7: return info.scalePct;
                case 8: return info.ppi;
                case 9: return info.status;
            }
            return info.index;
        }

        function isMissing(v) {
            return v === null || v === undefined || (typeof v === "number" && isNaN(v));
        }

        // ソート＋フィルタ適用 → リスト再構築
        function rebuildList() {
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
                for (var k = 0; k < sourceEntries.length; k++) {
                    var entry = sourceEntries[k];
                    if (hasStatusFlt) {
                        var sc = entry.statusCode;
                        if (sc === "ok" && !allowOk) continue;
                        if (sc === "broken" && !allowBroken) continue;
                        if (sc === "update" && !allowUpdate) continue;
                    }
                    if (hasArtboardFilter) {
                        var match = entry.artboardSet
                            ? !!entry.artboardSet[targetArtboardNumber]
                            : (entry.artboardNum === targetArtboardNumber);
                        if (!match) continue;
                    }
                    filteredEntries.push(entry);
                }
            } else {
                filteredEntries = sourceEntries;
            }

            filteredEntries.sort(function (a, b) {
                var va = getSortValue(a, sortBy);
                var vb = getSortValue(b, sortBy);
                // 欠損値は常に末尾
                var am = isMissing(va), bm = isMissing(vb);
                if (am && bm) return 0;
                if (am) return 1;
                if (bm) return -1;
                if (va < vb) return desc ? 1 : -1;
                if (va > vb) return desc ? -1 : 1;
                return 0;
            });

            // 列スペックに沿って動的にセル値を流し込む
            var cols = getColumnSpec();
            var unified = unitCheck.value;
            listBox.removeAll();
            for (var j = 0; j < filteredEntries.length; j++) {
                var info = filteredEntries[j];
                // 先頭列は .text、残りは subItems にセット
                var firstKey = cols[0].key;
                var firstText = info[firstKey];
                var row = listBox.add("item", (firstText === undefined || firstText === null) ? "" : String(firstText));
                for (var cj = 1; cj < cols.length; cj++) {
                    var key = cols[cj].key;
                    var txt;
                    if (key === "fileSize") {
                        txt = unified
                            ? formatFileSize(info.fileSizeBytes)
                            : formatFileSizeAuto(info.fileSizeBytes);
                    } else if (key === "fileCount") {
                        txt = String(info.fileCount);
                    } else {
                        txt = info[key];
                    }
                    row.subItems[cj - 1].text = (txt === undefined || txt === null) ? "" : txt;
                }
            }
        }

        // 並び順を「サイズ」「使用数」「幅」「高さ」「スケール」「PPI」に変えたら自動で降順 ON
        sortDropdown.onChange = function () {
            var idx = sortDropdown.selection ? sortDropdown.selection.index : 0;
            if (idx === 2 || idx === 3 || idx === 5 || idx === 6 || idx === 7 || idx === 8) {
                descRadio.value = true;
                ascRadio.value = false;
            }
            rebuildList();
        };
        ascRadio.onClick = rebuildList;
        descRadio.onClick = rebuildList;

        // 「重複」チェックボックス：ON で重複をまとめる、OFF で全配置を個別表示
        dedupCheck.onClick = function () {
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
            dlg.layout.layout(true);
            rebuildList();
        }
        abFilterDropdown.onChange = applyArtboardFilter;

        abPrevBtn.onClick = function () {
            var n = abFilterDropdown.items.length;
            if (n <= 1) return;
            var cur = abFilterDropdown.selection ? abFilterDropdown.selection.index : 0;
            var next = cur;
            for (var step = 0; step < n; step++) {
                next = next - 1;
                if (next < 1) next = n - 1;
                if (abFilterDropdown.items[next].enabled) {
                    abFilterDropdown.selection = next;
                    applyArtboardFilter();
                    return;
                }
            }
        };

        abNextBtn.onClick = function () {
            var n = abFilterDropdown.items.length;
            if (n <= 1) return;
            var cur = abFilterDropdown.selection ? abFilterDropdown.selection.index : 0;
            var next = cur;
            for (var step = 0; step < n; step++) {
                next = next + 1;
                if (next >= n) next = 1;
                if (abFilterDropdown.items[next].enabled) {
                    abFilterDropdown.selection = next;
                    applyArtboardFilter();
                    return;
                }
            }
        };

        // 列の表示/非表示・ヘッダ更新時の共通処理（listBox を作り直す）
        function recreateListBoxAndRebuildList() {
            createListBox();
            dlg.layout.layout(true);
            rebuildList();
        }
        unitCheck.onClick = recreateListBoxAndRebuildList;
        sizeColCheck.onClick = recreateListBoxAndRebuildList;
        countColCheck.onClick = recreateListBoxAndRebuildList;
        dimScalePpiCheck.onClick = recreateListBoxAndRebuildList;

        // 初期表示
        rebuildList();

        // --- パスパネル ---
        var pathPanel = dlg.add("panel", undefined, L('pathPanelTitle'));
        pathPanel.orientation = "column";
        pathPanel.alignChildren = ["fill", "top"];
        pathPanel.margins = PANEL_MARGINS;

        // 1 行目：選択中の親フォルダ表示
        var pathRow = pathPanel.add("group");
        pathRow.orientation = "row";
        pathRow.alignChildren = ["fill", "center"];

        var pathText = pathRow.add(
            "edittext", undefined, L('pathPlaceholder'),
            { readonly: true, multiline: false }
        );
        pathText.preferredSize = [500, 30];

        // 2 行目：3 カラム（左：整形オプション / 中央：spacer / 右：アクションボタン）
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
        tildeCheck.value = true;
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

        // ファイル名チェックに応じて「フォルダのみ」「ファイル名まで」を切替
        function buildDisplayedPath(absPath) {
            var base = fileNameCheck.value ? absPath : toFolderOnly(absPath);
            return formatDisplayPath(base, tildeCheck.value, dropboxCheck.value);
        }
        function updatePathDisplay() {
            if (!selectedFilePath) return;
            pathText.text = buildDisplayedPath(selectedFilePath);
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
        fileNameCheck.onClick = updatePathDisplay; // フォルダ一覧には影響しない
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

        // oldFolder 配下に配置された全ファイルを newFolder 内の同名ファイルに差し替え
        // （InDesign の「フォルダーに再リンク」相当）
        function relinkFolder(oldFolder, newFolder) {
            var success = 0, failed = 0, total = 0;
            for (var k = 0; k < allPlacementEntries.length; k++) {
                var ent = allPlacementEntries[k];
                if (!ent.filePath || ent.filePath === "---") continue;
                var sp = Math.max(ent.filePath.lastIndexOf("/"), ent.filePath.lastIndexOf("\\"));
                if (sp > 0 && ent.filePath.substring(0, sp) === oldFolder) {
                    total++;
                    var fname = ent.filePath.substring(sp + 1);
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

        // アクションボタンは右カラム（pathOptRight）に追加
        var openFolderBtn = pathOptRight.add("button", undefined, L('openFolderBtn'));
        openFolderBtn.onClick = function () {
            if (!selectedFilePath || selectedFilePath === "---") {
                alert(L('alertNoValidPath'));
                return;
            }
            try {
                var f = new File(selectedFilePath);
                if (f.parent) f.parent.execute();
            } catch (e) {
                alert(L('alertOpenFolderFailed') + e.message);
            }
        };

        var reloadOneBtn = pathOptRight.add("button", undefined, L('reloadOneBtn'));
        reloadOneBtn.onClick = function () {
            if (!selectedEntry) {
                alert(L('alertSelectItem'));
                return;
            }
            var picked = File.openDialog(L('selectNewLinkFile'));
            if (!picked) return;
            var res = relinkIndicesTo(selectedEntry.itemIndices, picked);
            app.redraw();
            refreshFromDoc();
            alert(
                L('alertRelinkDone') + "\n" +
                L('labelSuccess') + (lang === 'ja' ? '：' : ': ') + res.success + (lang === 'ja' ? ' 件' : '') + "\n" +
                L('labelFailed') + (lang === 'ja' ? '：' : ': ') + res.failed + (lang === 'ja' ? ' 件' : '')
            );
        };

        // パスが「---」（不明）のときはフォルダ操作系ボタンを無効化
        function updateActionButtonStates() {
            var hasPath = !!selectedFilePath && selectedFilePath !== "---";
            openFolderBtn.enabled = hasPath;
            reloadOneBtn.enabled = (selectedEntry !== null);
        }

        // 初期状態：まだ選択がないのでフォルダ系ボタンを無効化
        updateActionButtonStates();

        // --- リンクフォルダ一覧（パスパネル内、重複排除） ---
        var linkedFolderPaths = [];
        function rebuildFolderList() {
            var linkedFolderMap = {};
            linkedFolderPaths = [];
            for (var fi = 0; fi < allPlacementEntries.length; fi++) {
                var filePath = allPlacementEntries[fi].filePath;
                if (filePath && filePath !== "---") {
                    var folderSeparatorIndex = Math.max(filePath.lastIndexOf("/"), filePath.lastIndexOf("\\"));
                    if (folderSeparatorIndex > 0) {
                        var linkedFolderPath = filePath.substring(0, folderSeparatorIndex);
                        if (!linkedFolderMap[linkedFolderPath]) {
                            linkedFolderMap[linkedFolderPath] = true;
                            linkedFolderPaths.push(linkedFolderPath);
                        }
                    }
                }
            }
            linkedFolderPaths.sort();
        }
        rebuildFolderList();

        var folderCountLabel = pathPanel.add("statictext", undefined, L('linkedFolderListLabel') + " (" + linkedFolderPaths.length + (lang === 'ja' ? ' 件)' : ' items)'));
        var foldersListBox = pathPanel.add("listbox", undefined, [], { multiselect: false });
        foldersListBox.preferredSize = FOLDER_LISTBOX_SIZE;
        foldersListBox.alignment = ["fill", "fill"];

        function populateFoldersList() {
            foldersListBox.removeAll();
            for (var ii = 0; ii < linkedFolderPaths.length; ii++) {
                foldersListBox.add("item",
                    formatDisplayPath(linkedFolderPaths[ii], tildeCheck.value, dropboxCheck.value)
                );
            }
        }
        populateFoldersList();

        // リリンク後に doc から再収集してリストを更新
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
                    for (var r = 0; r < sourceEntries.length; r++) {
                        var re = sourceEntries[r];
                        for (var rj = 0; rj < re.itemIndices.length; rj++) {
                            if (re.itemIndices[rj] === anchor) { rematched = re; break; }
                        }
                        if (rematched) break;
                    }
                }
                selectedEntry = rematched;
                selectedFilePath = rematched ? rematched.filePath : "";
                pathText.text = selectedFilePath ? buildDisplayedPath(selectedFilePath) : L('pathPlaceholder');
                updateActionButtonStates();
            }

            rebuildFolderList();
            folderCountLabel.text = L('linkedFolderListLabel') + " (" + linkedFolderPaths.length + (lang === 'ja' ? ' 件)' : ' items)');
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
                L('labelTarget') + (lang === 'ja' ? '：' : ': ') + res.total + (lang === 'ja' ? ' 件' : '') + "\n" +
                L('labelSuccess') + (lang === 'ja' ? '：' : ': ') + res.success + (lang === 'ja' ? ' 件' : '') + "\n" +
                L('labelFailed') + (lang === 'ja' ? '：' : ': ') + res.failed + (lang === 'ja' ? ' 件（同名ファイルなし）' : ' item(s) (same-name file not found)')
            );
        };
        reloadFolderBtn.enabled = false;

        var collectLinksBtn = folderActionRight.add("button", undefined, L('collectLinksBtn'));
        collectLinksBtn.onClick = function () {
            var docFile = null;
            try { docFile = doc.fullName; } catch (e) { docFile = null; }
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

            var colon = (lang === 'ja' ? '：' : ': ');
            var unit = (lang === 'ja' ? ' 件' : '');
            alert(
                L('alertCollectLinksDone') + "\n" +
                L('labelTarget') + colon + total + unit + "\n" +
                L('labelCopied') + colon + copied + unit + "\n" +
                L('labelSkipped') + colon + skipped + unit + "\n" +
                L('labelFailed') + colon + failed + unit
            );
        };

        // フォルダ選択が変わったらボタンを enable/disable
        foldersListBox.onChange = function () {
            reloadFolderBtn.enabled = (foldersListBox.selection !== null);
        };

        // メインリスト選択に合わせてフォルダ一覧の該当行をハイライト
        function highlightFolderFor(absPath) {
            if (!absPath || absPath === "---") {
                foldersListBox.selection = null;
                return;
            }
            var sep = Math.max(absPath.lastIndexOf("/"), absPath.lastIndexOf("\\"));
            if (sep <= 0) return;
            var folder = absPath.substring(0, sep);
            for (var f = 0; f < linkedFolderPaths.length; f++) {
                if (linkedFolderPaths[f] === folder) {
                    try {
                        foldersListBox.selection = foldersListBox.items[f];
                        if (typeof foldersListBox.revealItem === "function") {
                            foldersListBox.revealItem(foldersListBox.items[f]);
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
            for (var t = 0; t < indices.length; t++) {
                placedItems[indices[t]].selected = true;
            }
            zoomToSelection(doc);
        }

        // --- ボタングループ（左：カンバス上で表示 / 右：キャンセル） ---
        var btnGroup = dlg.add("group");
        btnGroup.orientation = "row";
        btnGroup.alignment = "fill";
        btnGroup.alignChildren = ["fill", "center"];

        var showOnCanvasCheck = btnGroup.add("checkbox", undefined, L('showOnCanvasCheck'));
        showOnCanvasCheck.value = true; // ON：行選択でカンバス上の該当画像をフィット表示
        showOnCanvasCheck.alignment = ["left", "center"];

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
                    pathText.text = buildDisplayedPath(ent.filePath);
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
