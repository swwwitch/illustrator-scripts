#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

このスクリプトと同じ場所に置いたライブラリー（.ai）からグラフィックスタイルを取り込み、選択したオブジェクトに適用します。
書類の［グラフィックスタイル］パネルに無いスタイルは、取り込んでから適用します。

詳細は README を参照してください。

### Overview

Imports graphic styles from the library (.ai) kept beside this script and applies one to the selected objects.
A style the document's Graphic Styles panel does not carry yet is imported first, then applied.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ImportAndApplyGraphicLibary";  /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-09-20";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-20";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ImportAndApplyGraphicLibary.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ImportAndApplyGraphicLibary.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* グラフィックスタイルのライブラリー（このスクリプトと同じ場所に置く）。ドロップダウンにはこのファイルのスタイルが並ぶ
       Graphic style library, kept beside this script; its styles are what the dropdown lists */
    var GRAPHIC_LIBRARY_NAME = "graphiclibraryzoomin.ai";

    /* 取り込みに使う一時レイヤーの名前（貼り付けたオブジェクトごと削除する）/ Name of the temporary layer used for the import (removed with what was pasted) */
    var IMPORT_LAYER_NAME = "__import_graphic_styles__";

    var DEFAULT_STYLE = ""; /* 初期のスタイル名（空でライブラリーの先頭）/ initial style name ("" selects the library's first) */

    // =========================================
    // レイアウト / Layout
    // =========================================

    var WINDOW_MARGINS     = 16;            /* ウィンドウ外周の余白 / window margin */
    var WINDOW_SPACING     = 12;            /* ウィンドウ内の要素間隔 / window spacing */
    var FIELD_ROW_SPACING  = 6;             /* ラベルと入力欄の間隔 / gap inside a labeled row */
    var LABEL_WIDTH        = 60;            /* 行ラベルの共通幅 / shared width of row labels */
    var DROPDOWN_WIDTH     = 200;           /* ドロップダウンの幅（長い項目でダイアログを広げないよう固定）/ fixed width, so a long item cannot widen the dialog */
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
            title: { ja: "グラフィックスタイルの適用", en: "Apply Graphic Style" }
        },
        fieldLabel: {
            style: { ja: "スタイル", en: "Style" }
        },
        tooltip: {
            style:  { ja: "ライブラリーのグラフィックスタイルを、選択したオブジェクトに適用します（初回の適用でライブラリーのスタイルが書類に取り込まれます）", en: "Applies a graphic style from the library to the selected objects (the first apply imports the library's styles into the document)" },
            apply:  { ja: "選択したオブジェクトにスタイルを適用して閉じます", en: "Applies the style to the selected objects and closes" },
            cancel: { ja: "適用せずに閉じます", en: "Closes without applying" }
        },
        button: {
            apply:  { ja: "適用", en: "Apply" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        alert: {
            noDocument:  { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            lockedLayer: { ja: "アクティブレイヤーがロックまたは非表示です。", en: "The active layer is locked or hidden." },
            noSelection: { ja: "グラフィックスタイルを適用するオブジェクトを選択してください。", en: "Select the objects you want the graphic style applied to." },
            noLibrary:   { ja: "ライブラリーが見つかりません。このスクリプトと同じ場所に置いてください：", en: "The library was not found. Keep it beside this script:" },
            noStyle:     { ja: "ライブラリーに、適用できるグラフィックスタイルがありません。", en: "The library has no graphic style to apply." },
            applyFailed: { ja: "グラフィックスタイルを適用できませんでした。", en: "The graphic style could not be applied." }
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
    // グラフィックスタイル / Graphic styles
    // =========================================

    /* ライブラリーから読んだスタイル名を同一セッション内だけ覚えておく $.global 上のキー / Key on $.global that remembers the library's style names within this session */
    var STYLE_LIBRARY_CACHE_KEY = "importAndApplyGraphicLibrary";

    /**
     * グラフィックスタイルのライブラリーを返す（スクリプトと同じ場所）
     * @returns {File|null} ライブラリーのファイル（無ければ null）
     */
    function getGraphicLibraryFile() {
        /* スクリプトの場所が取れない実行方法もある / Some ways of running a script leave no path to work from */
        var scriptFolder = File($.fileName).parent;
        if (!scriptFolder) return null;

        var libraryFile = new File(scriptFolder.fsName + "/" + GRAPHIC_LIBRARY_NAME);
        return libraryFile.exists ? libraryFile : null;
    }

    /**
     * 書類の［グラフィックスタイル］パネルに、その名前のスタイルがあるか
     * @param {Document} doc - 対象ドキュメント
     * @param {string} styleName - 探すスタイル名
     * @returns {boolean} あれば true
     */
    function hasGraphicStyle(doc, styleName) {
        for (var i = 0; i < doc.graphicStyles.length; i++) {
            if (doc.graphicStyles[i].name === styleName) return true;
        }
        return false;
    }

    /**
     * ライブラリーを開いて処理に渡し、開いたぶんは閉じて元の書類に戻る
     * @param {Document} doc - 元の書類（処理後にアクティブへ戻す）
     * @param {function} handleLibrary - ライブラリーのドキュメントを受け取る処理
     * @returns {*} handleLibrary の戻り値（ライブラリーが無ければ null）
     */
    function withGraphicLibrary(doc, handleLibrary) {
        var libraryFile = getGraphicLibraryFile();
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
     * ライブラリーのスタイル名を並び順で返す
     * @param {Document} doc - 対象ドキュメント
     * @returns {string[]} スタイル名（ライブラリーが無ければ空）
     */
    function getLibraryStyleNames(doc) {
        var libraryFile = getGraphicLibraryFile();
        if (!libraryFile) return [];

        /* 一度読んだ名前は覚えておき、ライブラリーを開き直さない（更新されていれば読み直す）
           Remembered names save reopening the library; a newer file is read again */
        var libraryCache = $.global[STYLE_LIBRARY_CACHE_KEY];
        var libraryStamp = libraryFile.fsName + "\t" + libraryFile.modified.getTime();
        if (libraryCache && libraryCache.stamp === libraryStamp) return libraryCache.names;

        var libraryNames = withGraphicLibrary(doc, function (libraryDoc) {
            var names = [];
            /* 先頭はどの書類にも入っている既定のスタイルなので外す / The first entry is the default style every document carries, so it is left out */
            for (var i = 1; i < libraryDoc.graphicStyles.length; i++) {
                names.push(libraryDoc.graphicStyles[i].name);
            }
            return names;
        }) || [];

        $.global[STYLE_LIBRARY_CACHE_KEY] = { stamp: libraryStamp, names: libraryNames };
        return libraryNames;
    }

    /* ライブラリーの取り込みを済ませたか（実行中だけ覚える）/ Whether the library has been imported, remembered for this run */
    var isLibraryImported = false;

    /**
     * ライブラリーのグラフィックスタイルを書類に取り込む
     * ライブラリーのオブジェクトを一時レイヤーへ貼り付けると、そのレイヤーを捨ててもスタイルはパネルに残る
     * @param {Document} doc - 取り込み先のドキュメント
     * @returns {void}
     */
    function importLibraryStyles(doc) {
        /* コピーはメニューコマンドで行う（app.copy() は黙って失敗することがある）/ A menu command does the copy; app.copy() can fail silently */
        var isCopied = withGraphicLibrary(doc, function () {
            app.redraw();
            app.executeMenuCommand("selectallinartboard");
            app.executeMenuCommand("copy");
            return true;
        });
        if (!isCopied) return;

        /* 貼り付け前に選択を解除する（解除しないと選択中のオブジェクトが置き換わる）/ Clear the selection first, or the paste replaces what is selected */
        doc.selection = null;

        var importLayer = doc.layers.add();
        importLayer.name = IMPORT_LAYER_NAME;
        doc.activeLayer = importLayer;
        app.executeMenuCommand("paste");

        /* スタイルはパネルに残るので、貼り付けたものはレイヤーごと捨てる / The styles stay in the panel, so what was pasted goes with the layer */
        try {
            importLayer.remove();
        } catch (eLayer) {
            /* レイヤーを消せなくても、取り込み自体は済んでいる / The import is done even when the layer cannot be removed */
        }
    }

    /**
     * スタイルを書類の［グラフィックスタイル］パネルに用意する（無ければライブラリーから取り込む）
     * @param {Document} doc - 取り込み先のドキュメント
     * @param {string} styleName - 用意したいスタイル名
     * @returns {boolean} 使える状態になったら true
     */
    function ensureGraphicStyle(doc, styleName) {
        if (hasGraphicStyle(doc, styleName)) return true;
        /* 取り込みは一度で全スタイルが入るので、ライブラリーを開き直さない / One import brings every style, so the library is never reopened */
        if (isLibraryImported) return false;

        importLibraryStyles(doc);
        isLibraryImported = true;
        return hasGraphicStyle(doc, styleName);
    }

    /**
     * 書類のグラフィックスタイルをオブジェクトに適用する
     * @param {Document} doc - 対象ドキュメント
     * @param {PageItem} targetItem - 適用先のオブジェクト
     * @param {string} styleName - 適用するスタイルの名前
     * @returns {boolean} 適用できたら true
     */
    function applyGraphicStyle(doc, targetItem, styleName) {
        try {
            if (!ensureGraphicStyle(doc, styleName)) return false;
            doc.graphicStyles.getByName(styleName).applyTo(targetItem);
            return true;
        } catch (eStyle) {
            /* 1つ適用できなくても、残りのオブジェクトは続けて処理する / One failure must not stop the remaining objects */
            return false;
        }
    }

    // =========================================
    // 設定の記憶 / Remembered settings
    // =========================================

    /* 前回選んだスタイル名を同一セッション内だけ覚えておく $.global 上のキー / Key on $.global that remembers the last style within this session */
    var LAST_STYLE_KEY = "importAndApplyGraphicLibraryLastStyle";

    /**
     * 前回選んだスタイル名を読む
     * @returns {string} 覚えていたスタイル名（無ければ空文字）
     */
    function loadLastStyleName() {
        var lastStyleName = $.global[LAST_STYLE_KEY];
        return (lastStyleName === undefined || lastStyleName === null) ? "" : String(lastStyleName);
    }

    /**
     * 選んだスタイル名を覚えておく
     * @param {string} styleName - 覚えるスタイル名
     * @returns {void}
     */
    function saveLastStyleName(styleName) {
        $.global[LAST_STYLE_KEY] = styleName;
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * スタイル選択ダイアログを構築する
     * @param {Document} doc - 対象ドキュメント
     * @param {string[]} libraryStyleNames - ドロップダウンに並べるスタイル名
     * @param {PageItem[]} targetItems - スタイルを適用するオブジェクト
     * @returns {Window} 構築済みのダイアログ
     */
    function createGraphicStyleDialog(doc, libraryStyleNames, targetItems) {
        var styleDialog = new Window("dialog", getLabel('dialog', 'title') + " " + SCRIPT_VERSION);
        setupWindow(styleDialog);

        var styleRow = styleDialog.add("group");
        setupRow(styleRow, "fill", FIELD_ROW_SPACING);
        addRowLabel(styleRow, labelText('fieldLabel', 'style'));

        var styleDropdown = styleRow.add("dropdownlist", undefined, libraryStyleNames);
        styleDropdown.helpTip = getLabel('tooltip', 'style');
        /* 名前が長くてもダイアログを広げない / A long style name must not widen the dialog */
        styleDropdown.preferredSize.width = DROPDOWN_WIDTH;
        styleDropdown.alignment = ["fill", "center"];

        /* 覚えていたスタイルがライブラリーから消えていることもあるので、無ければ先頭に戻す / A remembered style may be gone from the library, so fall back to the first one */
        var initialStyleName = loadLastStyleName() || DEFAULT_STYLE;
        styleDropdown.selection = 0;
        for (var i = 0; i < libraryStyleNames.length; i++) {
            if (libraryStyleNames[i] === initialStyleName) styleDropdown.selection = i;
        }

        /* ボタンエリア：左側に置くものがないので行ごと右寄せ / Button row: nothing sits on the left, so the row itself is right aligned */
        var btnRowGroup = styleDialog.add("group");
        setupRow(btnRowGroup, "right", BUTTON_BAR_SPACING);
        btnRowGroup.margins = BUTTON_BAR_MARGINS;
        var btnCancel = btnRowGroup.add("button", undefined, getLabel('button', 'cancel'), { name: "cancel" });
        btnCancel.helpTip = getLabel('tooltip', 'cancel');
        var btnApply = btnRowGroup.add("button", undefined, getLabel('button', 'apply'), { name: "ok" });
        btnApply.helpTip = getLabel('tooltip', 'apply');

        /* ［適用］：選んだスタイルを選択したオブジェクトすべてに適用する / Apply: put the chosen style on every selected object */
        btnApply.onClick = function () {
            var selectedStyleName = styleDropdown.selection ? styleDropdown.selection.text : "";
            if (!selectedStyleName) {
                styleDialog.close();
                return;
            }

            var appliedCount = 0;
            for (var i = 0; i < targetItems.length; i++) {
                if (applyGraphicStyle(doc, targetItems[i], selectedStyleName)) appliedCount++;
            }

            /* 取り込みで解除した選択を元に戻す / Restore the selection the import had to clear */
            doc.selection = targetItems;

            saveLastStyleName(selectedStyleName);
            if (appliedCount === 0) alert(getLabel('alert', 'applyFailed'));

            styleDialog.close();
        };

        return styleDialog;
    }

    // =========================================
    // メイン / Main
    // =========================================

    /**
     * 選択したオブジェクトにライブラリーのグラフィックスタイルを適用する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel('alert', 'noDocument'));
            return;
        }
        var doc = app.activeDocument;

        /* 取り込みの一時レイヤーを作るので、アクティブレイヤーが使えないなら始めない / A temporary layer is created for the import, so do not start when the active layer cannot take one */
        var activeLayer = doc.activeLayer;
        if (activeLayer.locked || !activeLayer.visible) {
            alert(getLabel('alert', 'lockedLayer'));
            return;
        }

        if (!getGraphicLibraryFile()) {
            alert(getLabel('alert', 'noLibrary') + "\n" + GRAPHIC_LIBRARY_NAME);
            return;
        }

        /* 文字を編集中の選択は TextRange で返るので、アイテムの配列のときだけ見る / A text-editing selection comes back as a TextRange, so only an array of items is examined */
        var selectedItems = (doc.selection instanceof Array) ? doc.selection : [];
        /* 取り込みで選択が解除されても使えるよう、選択を写しておく / Copy the selection, so it survives the import clearing it */
        var targetItems = [];
        for (var i = 0; i < selectedItems.length; i++) {
            targetItems.push(selectedItems[i]);
        }
        if (targetItems.length === 0) {
            alert(getLabel('alert', 'noSelection'));
            return;
        }

        var libraryStyleNames = getLibraryStyleNames(doc);
        if (libraryStyleNames.length === 0) {
            alert(getLabel('alert', 'noStyle'));
            return;
        }

        createGraphicStyleDialog(doc, libraryStyleNames, targetItems).show();
    }

    main();

})();
