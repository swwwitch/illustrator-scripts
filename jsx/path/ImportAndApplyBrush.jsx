#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

このスクリプトと同じ場所に置いたブラシライブラリー（.ai）からブラシを取り込み、選択したパスに適用します。
書類の［ブラシ］パネルに無いブラシは、取り込んでから適用します。

詳細は README を参照してください。

### Overview

Imports a brush from the brush library (.ai) kept beside this script and applies it to the selected paths.
A brush the document's Brushes panel does not carry yet is imported first, then applied.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ImportAndApplyBrush";          /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-09-20";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-20";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ImportAndApplyBrush.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ImportAndApplyBrush.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* ブラシライブラリー（このスクリプトと同じ場所に置く）。ドロップダウンにはこのファイルのブラシが並ぶ
       Brush library, kept beside this script; its brushes are what the dropdown lists */
    var BRUSH_LIBRARY_NAME = "brushlibrary.ai";

    /* ドロップダウンに出さないブラシ名。どの書類にも最初から入っている既定のブラシを、日本語UI・英語UIの名前で外す
       Brush names left out of the dropdown: the default brush every document already carries, in Japanese and English */
    var EXCLUDED_BRUSH_NAMES = ["カリグラフィブラシをタッチ", "Touch Calligraphic Brush"];

    var DEFAULT_BRUSH = ""; /* 初期のブラシ名（空でライブラリーの先頭）/ initial brush name ("" selects the library's first) */

    // =========================================
    // レイアウト / Layout
    // =========================================

    var WINDOW_MARGINS     = 16;            /* ウィンドウ外周の余白 / window margin */
    var WINDOW_SPACING     = 12;            /* ウィンドウ内の要素間隔 / window spacing */
    var FIELD_ROW_SPACING  = 6;             /* ラベルと入力欄の間隔 / gap inside a labeled row */
    var LABEL_WIDTH        = 48;            /* 行ラベルの共通幅 / shared width of row labels */
    var DROPDOWN_WIDTH     = 180;           /* ドロップダウンの幅（長い項目でダイアログを広げないよう固定）/ fixed width, so a long item cannot widen the dialog */
    var BUTTON_BAR_MARGINS = [0, 10, 0, 0]; /* ボタンバーの余白 / margins of the bottom button bar */
    var BUTTON_BAR_SPACING = 10;            /* ボタンバー内の要素間隔 / spacing inside the button bar */

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * 現在の表示言語を取得する
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        var localeText = ($.locale || "") + ""; /* 文字列化して扱う / Ensure a string */
        /* "ja" で始まるロケール（ja, ja_JP など）は日本語扱い / Treat "ja*" locales as Japanese */
        if (localeText.indexOf("ja") === 0) {
            return "ja";
        }
        return "en";
    }
    var uiLang = getCurrentLang();

    /* カテゴリ分けした日英ラベル定義 / Categorized Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "ブラシの適用", en: "Apply Brush" }
        },
        fieldLabel: {
            brush: { ja: "ブラシ", en: "Brush" }
        },
        tooltip: {
            brush:  { ja: "ブラシライブラリーのブラシを、選択したパスに適用します", en: "Applies a brush from the brush library to the selected paths" },
            apply:  { ja: "選択したパスにブラシを適用して閉じます", en: "Applies the brush to the selected paths and closes" },
            cancel: { ja: "適用せずに閉じます", en: "Closes without applying" }
        },
        button: {
            apply:  { ja: "適用", en: "Apply" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        alert: {
            noDocument:  { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            lockedLayer: { ja: "アクティブレイヤーがロックまたは非表示です。", en: "The active layer is locked or hidden." },
            noSelection: { ja: "ブラシを適用するパスを選択してください。", en: "Select the paths you want the brush applied to." },
            noLibrary:   { ja: "ブラシライブラリーが見つかりません。このスクリプトと同じ場所に置いてください：", en: "The brush library was not found. Keep it beside this script:" },
            noBrush:     { ja: "ブラシライブラリーに、適用できるブラシがありません。", en: "The brush library has no brush to apply." },
            applyFailed: { ja: "ブラシを適用できませんでした。", en: "The brush could not be applied." }
        }
    };

    /**
     * LABELS からカテゴリを辿って現在の言語のラベルを取得する（例: getLabel('button','apply')）
     * @param {...string} keys - LABELS を辿るキー列
     * @returns {string} 該当するラベル（見つからない場合は空文字）
     */
    function getLabel() {
        var labelNode = LABELS;
        for (var i = 0; i < arguments.length; i++) {
            if (labelNode == null) break;
            labelNode = labelNode[arguments[i]];
        }
        return (labelNode && labelNode[uiLang] != null) ? labelNode[uiLang] : "";
    }

    /**
     * コロン付きの項目名を返す（日本語は全角、英語は半角）
     * @param {...string} keys - LABELS を辿るキー列
     * @returns {string} コロンを付けたラベル
     */
    function labelText() {
        return getLabel.apply(null, arguments) + (uiLang === "ja" ? "：" : ":");
    }

    // =========================================
    // UIレイアウト補助 / UI layout helpers
    // =========================================

    /**
     * ダイアログ全体の並びと余白を設定する
     * @param {Window} targetWindow - 対象ウィンドウ
     * @returns {void}
     */
    function setupWindow(targetWindow) {
        targetWindow.orientation = "column";
        targetWindow.alignChildren = ["fill", "top"];
        targetWindow.margins = WINDOW_MARGINS;
        targetWindow.spacing = WINDOW_SPACING;
    }

    /**
     * グループを横並びの行として設定する
     * @param {Group} targetGroup - 対象グループ
     * @param {string} [horizontalAlign] - 横方向の揃え（省略時は "left"）
     * @param {number} [spacing] - 要素間隔（省略時は FIELD_ROW_SPACING）
     * @returns {void}
     */
    function setupRow(targetGroup, horizontalAlign, spacing) {
        targetGroup.orientation = "row";
        /* 揃えは横と天地を対で指定し、親の fill 継承を打ち消す / Pair both axes to cancel the parent's fill */
        targetGroup.alignment = [horizontalAlign || "left", "center"];
        targetGroup.alignChildren = ["left", "center"];
        targetGroup.spacing = (typeof spacing === "number") ? spacing : FIELD_ROW_SPACING;
    }

    /**
     * 共通幅で右揃えの行ラベルを追加する
     * @param {Group} parentRow - 追加先の行グループ
     * @param {string} rowLabelText - 表示するラベル
     * @returns {StaticText} 生成したラベル
     */
    function addRowLabel(parentRow, rowLabelText) {
        var rowLabel = parentRow.add("statictext", undefined, rowLabelText);
        rowLabel.preferredSize.width = LABEL_WIDTH;
        rowLabel.justify = "right";
        return rowLabel;
    }

    // =========================================
    // ブラシ / Brush
    // =========================================

    /* ライブラリーから読んだブラシ名を同一セッション内だけ覚えておく $.global 上のキー / Key on $.global that remembers the library's brush names within this session */
    var BRUSH_LIBRARY_CACHE_KEY = "importAndApplyBrushLibrary";

    /**
     * ブラシライブラリーのファイルを返す（スクリプトと同じ場所）
     * @returns {File|null} ライブラリーのファイル（無ければ null）
     */
    function getBrushLibraryFile() {
        /* スクリプトの場所が取れない実行方法もある / Some ways of running a script leave no path to work from */
        var scriptFolder = File($.fileName).parent;
        if (!scriptFolder) return null;

        var libraryFile = new File(scriptFolder.fsName + "/" + BRUSH_LIBRARY_NAME);
        return libraryFile.exists ? libraryFile : null;
    }

    /**
     * 書類の［ブラシ］パネルに、その名前のブラシがあるか
     * @param {Document} doc - 対象ドキュメント
     * @param {string} brushName - 探すブラシ名
     * @returns {boolean} あれば true
     */
    function hasBrush(doc, brushName) {
        for (var i = 0; i < doc.brushes.length; i++) {
            if (doc.brushes[i].name === brushName) return true;
        }
        return false;
    }

    /**
     * ブラシライブラリーを開いて処理に渡し、開いたぶんは閉じて元の書類に戻る
     * @param {Document} doc - 元の書類（処理後にアクティブへ戻す）
     * @param {function} handleLibrary - ライブラリーのドキュメントを受け取る処理
     * @returns {*} handleLibrary の戻り値（ライブラリーが無ければ null）
     */
    function withBrushLibrary(doc, handleLibrary) {
        var libraryFile = getBrushLibraryFile();
        if (!libraryFile) return null;

        /* すでに開いているライブラリーを開き直すと、開くのではなく手前に来るだけで書類数が変わらない。
           これを目印にすれば、パス文字列の比較（正規化やケース違いで外しうる）に頼らず判定できる
           Reopening an already open library only brings it to front, leaving the document count unchanged;
           that is a safer signal than comparing path strings, which normalization or case can defeat */
        var documentCountBefore = app.documents.length;
        var libraryDoc = app.open(libraryFile);
        var wasLibraryOpen = (app.documents.length === documentCountBefore);

        try {
            return handleLibrary(libraryDoc);
        } finally {
            /* 開いたのがこのスクリプトなら閉じる。もともと開いていたものは触らない / Close only what this script opened; leave what was already open */
            if (!wasLibraryOpen) libraryDoc.close(SaveOptions.DONOTSAVECHANGES);
            app.activeDocument = doc;
        }
    }

    /**
     * ドロップダウンに出さないブラシか
     * @param {string} brushName - 調べるブラシ名
     * @returns {boolean} 除外するなら true
     */
    function isExcludedBrush(brushName) {
        for (var i = 0; i < EXCLUDED_BRUSH_NAMES.length; i++) {
            if (EXCLUDED_BRUSH_NAMES[i] === brushName) return true;
        }
        return false;
    }

    /**
     * ライブラリーのブラシ名を並び順で返す（除外するブラシは飛ばす）
     * @param {Document} doc - 対象ドキュメント
     * @returns {string[]} ブラシ名（ライブラリーが無ければ空）
     */
    function getLibraryBrushNames(doc) {
        var libraryFile = getBrushLibraryFile();
        if (!libraryFile) return [];

        /* 一度読んだ名前は覚えておき、ライブラリーを開き直さない（更新されていれば読み直す）
           Remembered names save reopening the library; a newer file is read again */
        var libraryCache = $.global[BRUSH_LIBRARY_CACHE_KEY];
        var libraryStamp = libraryFile.fsName + "\t" + libraryFile.modified.getTime();
        if (libraryCache && libraryCache.stamp === libraryStamp) return libraryCache.names;

        var libraryNames = withBrushLibrary(doc, function (libraryDoc) {
            var names = [];
            for (var i = 0; i < libraryDoc.brushes.length; i++) {
                var libraryBrushName = libraryDoc.brushes[i].name;
                if (!isExcludedBrush(libraryBrushName)) names.push(libraryBrushName);
            }
            return names;
        }) || [];

        $.global[BRUSH_LIBRARY_CACHE_KEY] = { stamp: libraryStamp, names: libraryNames };
        return libraryNames;
    }

    /* 取り込めなかったブラシ名（実行中だけ覚える）/ Brushes that could not be imported, remembered for this run */
    var failedBrushNames = {};

    /**
     * ブラシを書類の［ブラシ］パネルに用意する（無ければライブラリーから取り込む）
     * ブラシを適用したパスを複製すると、そのパスを消してもブラシはパネルに残る
     * @param {Document} doc - 取り込み先のドキュメント
     * @param {string} brushName - 用意したいブラシ名
     * @returns {boolean} 使える状態になったら true
     */
    function ensureBrush(doc, brushName) {
        if (hasBrush(doc, brushName)) return true;
        /* 一度取り込めなかったブラシは、同じ実行でライブラリーを開き直さない / A brush that failed once must not reopen the library again in this run */
        if (failedBrushNames[brushName] === true) return false;

        var carrierCopy = withBrushLibrary(doc, function (libraryDoc) {
            var carrierPath = null;
            var copiedPath = null;
            try {
                var libraryBrush = libraryDoc.brushes.getByName(brushName);
                carrierPath = libraryDoc.pathItems.add();
                carrierPath.setEntirePath([[0, 0], [10, 0]]);
                carrierPath.filled = false;
                carrierPath.stroked = true;
                libraryBrush.applyTo(carrierPath);
                copiedPath = carrierPath.duplicate(doc.activeLayer, ElementPlacement.PLACEATEND);
            } catch (eImport) {
                copiedPath = null;
            }
            if (carrierPath) carrierPath.remove();
            return copiedPath;
        });

        /* 取り込みに使った複製は、ライブラリーを閉じて戻ってから消す / The carrier copy goes once the library is closed and we are back */
        if (carrierCopy) carrierCopy.remove();

        var isBrushReady = hasBrush(doc, brushName);
        if (!isBrushReady) failedBrushNames[brushName] = true;
        return isBrushReady;
    }

    /**
     * 書類のブラシをパスに適用する
     * @param {Document} doc - 対象ドキュメント
     * @param {PathItem} targetPath - 適用先のパス
     * @param {string} brushName - 適用するブラシの名前
     * @returns {boolean} 適用できたら true
     */
    function applyBrush(doc, targetPath, brushName) {
        try {
            if (!ensureBrush(doc, brushName)) return false;
            doc.brushes.getByName(brushName).applyTo(targetPath);
            return true;
        } catch (eBrush) {
            /* 1つ適用できなくても、残りのパスは続けて処理する / One failure must not stop the remaining paths */
            return false;
        }
    }

    // =========================================
    // 適用対象 / Target paths
    // =========================================

    /**
     * 選択からブラシを適用できるパスを集める（グループ・複合パスは中まで辿る）
     * @param {Array|PageItems} items - 走査するアイテム
     * @param {PathItem[]} collectedPaths - 集めたパスの受け皿
     * @returns {PathItem[]} 集めたパス
     */
    function collectTargetPaths(items, collectedPaths) {
        for (var i = 0; i < items.length; i++) {
            var item = items[i];
            if (item.locked || item.hidden) continue;

            if (item.typename === "PathItem") {
                collectedPaths.push(item);
            } else if (item.typename === "CompoundPathItem") {
                collectTargetPaths(item.pathItems, collectedPaths);
            } else if (item.typename === "GroupItem") {
                collectTargetPaths(item.pageItems, collectedPaths);
            }
        }
        return collectedPaths;
    }

    // =========================================
    // 設定の記憶 / Remembered settings
    // =========================================

    /* 前回選んだブラシ名を同一セッション内だけ覚えておく $.global 上のキー / Key on $.global that remembers the last brush within this session */
    var LAST_BRUSH_KEY = "importAndApplyBrushLastBrush";

    /**
     * 前回選んだブラシ名を読む
     * @returns {string} 覚えていたブラシ名（無ければ空文字）
     */
    function loadLastBrushName() {
        var lastBrushName = $.global[LAST_BRUSH_KEY];
        return (lastBrushName === undefined || lastBrushName === null) ? "" : String(lastBrushName);
    }

    /**
     * 選んだブラシ名を覚えておく
     * @param {string} brushName - 覚えるブラシ名
     * @returns {void}
     */
    function saveLastBrushName(brushName) {
        $.global[LAST_BRUSH_KEY] = brushName;
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * ブラシ選択ダイアログを構築する
     * @param {Document} doc - 対象ドキュメント
     * @param {string[]} libraryBrushNames - ドロップダウンに並べるブラシ名
     * @param {PathItem[]} targetPaths - ブラシを適用するパス
     * @returns {Window} 構築済みのダイアログ
     */
    function createBrushDialog(doc, libraryBrushNames, targetPaths) {
        var brushDialog = new Window("dialog", getLabel('dialog', 'title') + " " + SCRIPT_VERSION);
        setupWindow(brushDialog);

        var brushRow = brushDialog.add("group");
        setupRow(brushRow, "fill", FIELD_ROW_SPACING);
        addRowLabel(brushRow, labelText('fieldLabel', 'brush'));

        var brushDropdown = brushRow.add("dropdownlist", undefined, libraryBrushNames);
        brushDropdown.helpTip = getLabel('tooltip', 'brush');
        /* 名前が長くてもダイアログを広げない / A long brush name must not widen the dialog */
        brushDropdown.preferredSize.width = DROPDOWN_WIDTH;
        brushDropdown.alignment = ["fill", "center"];

        /* 覚えていたブラシがライブラリーから消えていることもあるので、無ければ先頭に戻す / A remembered brush may be gone from the library, so fall back to the first one */
        var initialBrushName = loadLastBrushName() || DEFAULT_BRUSH;
        brushDropdown.selection = 0;
        for (var i = 0; i < libraryBrushNames.length; i++) {
            if (libraryBrushNames[i] === initialBrushName) brushDropdown.selection = i;
        }

        /* ボタンエリア：左側に置くものがないので行ごと右寄せ / Button row: nothing sits on the left, so the row itself is right aligned */
        var btnRowGroup = brushDialog.add("group");
        setupRow(btnRowGroup, "right", BUTTON_BAR_SPACING);
        btnRowGroup.margins = BUTTON_BAR_MARGINS;
        var btnCancel = btnRowGroup.add("button", undefined, getLabel('button', 'cancel'), { name: "cancel" });
        btnCancel.helpTip = getLabel('tooltip', 'cancel');
        var btnApply = btnRowGroup.add("button", undefined, getLabel('button', 'apply'), { name: "ok" });
        btnApply.helpTip = getLabel('tooltip', 'apply');

        /* ［適用］：選んだブラシを選択したパスすべてに適用する / Apply: put the chosen brush on every selected path */
        btnApply.onClick = function () {
            var selectedBrushName = brushDropdown.selection ? brushDropdown.selection.text : "";
            if (!selectedBrushName) {
                brushDialog.close();
                return;
            }

            var appliedCount = 0;
            for (var i = 0; i < targetPaths.length; i++) {
                if (applyBrush(doc, targetPaths[i], selectedBrushName)) appliedCount++;
            }

            saveLastBrushName(selectedBrushName);
            if (appliedCount === 0) alert(getLabel('alert', 'applyFailed'));

            brushDialog.close();
        };

        return brushDialog;
    }

    // =========================================
    // メイン / Main
    // =========================================

    /**
     * 選択したパスにライブラリーのブラシを適用する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel('alert', 'noDocument'));
            return;
        }
        var doc = app.activeDocument;

        /* 取り込みの複製をアクティブレイヤーに置くので、そこが使えないなら始めない / The carrier copy lands on the active layer, so do not start when it cannot take one */
        var activeLayer = doc.activeLayer;
        if (activeLayer.locked || !activeLayer.visible) {
            alert(getLabel('alert', 'lockedLayer'));
            return;
        }

        if (!getBrushLibraryFile()) {
            alert(getLabel('alert', 'noLibrary') + "\n" + BRUSH_LIBRARY_NAME);
            return;
        }

        /* 文字を編集中の選択は TextRange で返るので、アイテムの配列のときだけ見る / A text-editing selection comes back as a TextRange, so only an array of items is examined */
        var selectedItems = (doc.selection instanceof Array) ? doc.selection : [];
        var targetPaths = collectTargetPaths(selectedItems, []);
        if (targetPaths.length === 0) {
            alert(getLabel('alert', 'noSelection'));
            return;
        }

        var libraryBrushNames = getLibraryBrushNames(doc);
        if (libraryBrushNames.length === 0) {
            alert(getLabel('alert', 'noBrush'));
            return;
        }

        createBrushDialog(doc, libraryBrushNames, targetPaths).show();
    }

    main();

})();
