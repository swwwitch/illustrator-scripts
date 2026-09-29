#target illustrator
#targetengine "ImportAndApplyGraphicLibaryEngine"
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
var SCRIPT_VERSION  = "v1.0.3";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-09-20";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-29";                   /* 更新日 / last updated */

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

    // =========================================
    // ローカライズ / Localization
    // =========================================

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // ローカライズ（再利用パーツ） / Localization (reusable)
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

    /**
     * UI の言語を返す（"ja" で始まるロケールは日本語、それ以外は英語）
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        return (String($.locale || "").indexOf("ja") === 0) ? "ja" : "en";
    }

    var uiLang = getCurrentLang();

    /**
     * LABELS から今の UI 言語の文言を取り出す。
     * @param {string|Object} labelRef - "dialog.title" のようなパス、または { ja, en }
     * @param {Object|Array} [placeholderValues] - { name: 値 } なら {name} を、[値, …] なら %1, %2 … を差し込む
     * @returns {string} 文言。パスが見つからなければパスの文字列、{ ja, en } が無ければ空文字
     */
    function getLabel(labelRef, placeholderValues) {
        var labelEntry = labelRef;
        if (typeof labelRef === "string") {
            var labelPathKeys = labelRef.split(".");
            labelEntry = LABELS;
            for (var i = 0; i < labelPathKeys.length && labelEntry != null; i++) {
                labelEntry = labelEntry[labelPathKeys[i]];
            }
        }
        var labelString;
        if (typeof labelEntry === "string") labelString = labelEntry;
        else if (labelEntry != null && labelEntry[uiLang] != null) labelString = labelEntry[uiLang];
        else if (labelEntry != null && labelEntry.en != null) labelString = labelEntry.en;
        else return (typeof labelRef === "string") ? labelRef : "";
        return fillLabelPlaceholders(String(labelString), placeholderValues);
    }

    /**
     * 項目名の文言の末尾にコロンを付ける（日本語は全角「：」、英語は半角「:」）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {Object|Array} [placeholderValues] - getLabel と同じ
     * @returns {string} コロン付きの文言
     */
    function labelText(labelRef, placeholderValues) {
        return getLabel(labelRef, placeholderValues) + (uiLang === "ja" ? "：" : ":");
    }

    /**
     * 「項目名：値」の1行を返す（日本語は「件数：5」、英語は「Count: 5」とコロンのあとに空白を入れる）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {string|number} value - コロンのあとに続ける値
     * @returns {string} 項目名と値をつないだ文字列
     */
    function labelValueText(labelRef, value) {
        return labelText(labelRef) + (uiLang === "ja" ? "" : " ") + value;
    }

    /**
     * 文言の {name} や %1 に値を差し込む
     * @param {string} labelString - 文言
     * @param {Object|Array} [placeholderValues] - { name: 値 } または [値, …]
     * @returns {string} 差し込んだ文言
     */
    function fillLabelPlaceholders(labelString, placeholderValues) {
        if (placeholderValues == null) return labelString;
        if (placeholderValues instanceof Array) {
            /* 大きい番号から置き換え、%1 が %10 の一部を置き換えないようにする / Replace from the highest index so %1 does not eat into %10 */
            for (var i = placeholderValues.length; i >= 1; i--) {
                labelString = labelString.split("%" + i).join(String(placeholderValues[i - 1]));
            }
            return labelString;
        }
        for (var placeholderKey in placeholderValues) {
            if (!placeholderValues.hasOwnProperty(placeholderKey)) continue;
            labelString = labelString.split("{" + placeholderKey + "}").join(String(placeholderValues[placeholderKey]));
        }
        return labelString;
    }

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // ローカライズ（再利用パーツ）ここまで / End of the reusable localization
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

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

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // ボタン行（再利用パーツ） / Button row (reusable)
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

    var BUTTON_ROW_TOP_MARGIN = 5; /* ボタン行の上の余白 / top margin of the button row */
    var BUTTON_ROW_SPACING = 10;   /* ボタンどうしの間隔 / spacing between buttons */

    /**
     * ダイアログ下部のボタン行を作る。
     * 通常は「左のグループ・伸びるスペーサー・右のグループ」、centered なら行そのものを左右中央に置く
     * @param {Window|Group|Panel} parent - 行を足す先（ふつうはダイアログ）
     * @param {Object} [rowOptions] - { centered: true } で左右中央に並べる
     * @returns {{rowGroup: Group, leftGroup: Group|null, rightGroup: Group|null}} 行と左右のグループ（centered のときは左右が null）
     */
    function addButtonRow(parent, rowOptions) {
        var isCentered = !!(rowOptions && rowOptions.centered);
        var btnRowGroup = parent.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.margins = [0, BUTTON_ROW_TOP_MARGIN, 0, 0];
        btnRowGroup.spacing = BUTTON_ROW_SPACING;

        if (isCentered) {
            btnRowGroup.alignment = ["center", "bottom"];
            btnRowGroup.alignChildren = ["center", "center"];
            return { rowGroup: btnRowGroup, leftGroup: null, rightGroup: null };
        }

        btnRowGroup.alignment = ["fill", "bottom"];

        var btnLeftGroup = btnRowGroup.add("group");
        btnLeftGroup.alignChildren = ["left", "center"];
        btnLeftGroup.spacing = BUTTON_ROW_SPACING;

        /* 余りの幅を吸って、右のグループを右端に寄せる / Absorbs the extra width so the right group sits at the right edge */
        var spacer = btnRowGroup.add("group");
        spacer.alignment = ["fill", "fill"];
        spacer.minimumSize.width = 0;

        var btnRightGroup = btnRowGroup.add("group");
        btnRightGroup.alignChildren = ["right", "center"];
        btnRightGroup.spacing = BUTTON_ROW_SPACING;

        return { rowGroup: btnRowGroup, leftGroup: btnLeftGroup, rightGroup: btnRightGroup };
    }

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // ボタン行（再利用パーツ）ここまで / End of the reusable button row
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

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

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // ダイアログの位置と不透明度（再利用パーツ） / Dialog position and opacity (reusable)
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

    var DIALOG_OPACITY = 0.98;       /* ダイアログの不透明度 / dialog opacity */
    var DIALOG_AVOID_MARGIN = 60;    /* 選択範囲の推定位置の両側に取る余裕（px）/ margin on each side of the estimated selection (px) */
    var DIALOG_AVOID_MAX_ITEMS = 100; /* 選択範囲を測るオブジェクトの上限 / max items measured for the selection bounds */

    /**
     * ダイアログの不透明度を設定し、前回閉じた位置で開いて、動かした位置を記録するようにする。
     * 開く位置が選択中のオブジェクトに重なりそうなときは、左右の反対側へずらす（Illustrator のみ）。
     * 既存の onShow / onMove / onClose は先に呼んでから、位置の復元・記録を行う。
     * @param {Window} dialog - 対象のダイアログ
     * @param {string} storageKey - 位置を覚えるキー（ふつうは SCRIPT_NAME）
     * @returns {void}
     */
    function prepareDialogWindow(dialog, storageKey) {
        /* 同じダイアログを開き直すときは、選択範囲を測り直すだけにする（ハンドラーを重ねない）
           When the same dialog is shown again, only re-measure the selection (don't stack handlers) */
        if (dialog.dialogWindowState) {
            dialog.dialogWindowState.selectionSpan = getSelectionViewSpan();
            dialog.dialogWindowState.avoidedLocation = null;
            return;
        }
        var locationKey = "__" + storageKey + "_DialogLocation";
        var previousOnShow = dialog.onShow;
        var previousOnMove = dialog.onMove;
        var previousOnClose = dialog.onClose;
        var windowState = {
            selectionSpan: getSelectionViewSpan(), /* 選択範囲は show() の前に測る / measured before show() */
            screenWidth: null,                     /* 最初に開いたときに推定する / estimated on the first show */
            avoidedLocation: null                  /* 避けるためにずらした位置（記録しない）/ location set to avoid the selection (not remembered) */
        };
        dialog.dialogWindowState = windowState;

        dialog.opacity = DIALOG_OPACITY;

        /* 今の位置を記録する / Remember the current location */
        function rememberDialogLocation() {
            var currentLocation = [dialog.location[0], dialog.location[1]];
            var avoidedLocation = windowState.avoidedLocation;
            if (avoidedLocation && currentLocation[0] === avoidedLocation[0] && currentLocation[1] === avoidedLocation[1]) return;
            $.global[locationKey] = currentLocation;
        }

        dialog.onShow = function () {
            /* 最初に開くときの既定の位置は画面の横中央なので、画面の幅を逆算できる。2回目からは前回の位置なので使い回す
               On the first show the default location is centered horizontally, which gives the screen width; reuse it afterwards */
            if (windowState.screenWidth === null) windowState.screenWidth = dialog.location[0] * 2 + dialog.bounds.width;
            if (previousOnShow) previousOnShow.apply(this, arguments);
            /* $.screens は実際の画面の大きさと合わない（Mac で 1280×524 など）ので、画面内かは判定しない
               $.screens does not match the real display (e.g. 1280x524 on a Mac), so no on-screen check */
            var savedLocation = $.global[locationKey];
            if (savedLocation) dialog.location = [savedLocation[0], savedLocation[1]];
            if (windowState.selectionSpan) {
                var avoidLeft = findDialogLeftAvoidingSelection(dialog.location[0], dialog.bounds.width, windowState.screenWidth, windowState.selectionSpan);
                if (avoidLeft !== null) {
                    dialog.location = [avoidLeft, dialog.location[1]];
                    /* 代入後の値で比べる（丸められることがある）/ Compare with the value after assignment, which may be rounded */
                    windowState.avoidedLocation = [dialog.location[0], dialog.location[1]];
                }
            }
        };
        dialog.onMove = function () {
            if (previousOnMove) previousOnMove.apply(this, arguments);
            rememberDialogLocation();
        };
        dialog.onClose = function () {
            rememberDialogLocation();
            /* false を返すと閉じるのを取りやめるので、戻り値は元の onClose のものを返す
               Returning false cancels the close, so pass the original onClose result through */
            if (previousOnClose) return previousOnClose.apply(this, arguments);
        };
    }

    /**
     * 選択中のオブジェクトが、ドキュメントの表示域の左端から画面上で何 px の範囲にあるかを返す。
     * @returns {{left: number, right: number, viewWidth: number}|null} 選択が無い・測れないときは null
     */
    function getSelectionViewSpan() {
        try {
            if (app.name !== "Adobe Illustrator" || !app.documents.length) return null;
            var targetDoc = app.activeDocument;
            var selectedItems = targetDoc.selection;
            if (!selectedItems || !selectedItems.length || !selectedItems[0].visibleBounds) return null;
            var itemCount = Math.min(selectedItems.length, DIALOG_AVOID_MAX_ITEMS);
            var spanLeft = Infinity;
            var spanRight = -Infinity;
            for (var i = 0; i < itemCount; i++) {
                var itemBounds = selectedItems[i].visibleBounds;
                if (itemBounds[0] < spanLeft) spanLeft = itemBounds[0];
                if (itemBounds[2] > spanRight) spanRight = itemBounds[2];
            }
            var activeView = targetDoc.activeView; /* 複数ウィンドウで開いていても今のウィンドウ / the current window even with multiple windows */
            var viewBounds = activeView.bounds;
            var zoom = activeView.zoom;
            var viewWidth = (viewBounds[2] - viewBounds[0]) * zoom;
            /* 表示域の外にはみ出した部分は数えない / Ignore the part outside the view */
            var left = Math.max(0, (spanLeft - viewBounds[0]) * zoom);
            var right = Math.min(viewWidth, (spanRight - viewBounds[0]) * zoom);
            if (right <= left) return null;
            return { left: left, right: right, viewWidth: viewWidth };
        } catch (e) {
            /* テキスト編集中など測れないときは避けない / Do not avoid when it cannot be measured, e.g. while editing text */
            return null;
        }
    }

    /**
     * ダイアログが選択範囲に重なるなら、重ならない左端の位置を返す。
     * 表示域は画面の横中央にあるとみなし、ずれは DIALOG_AVOID_MARGIN で吸収する。
     * @param {number} dialogLeft - 今のダイアログの左端
     * @param {number} dialogWidth - ダイアログの幅
     * @param {number} screenWidth - 画面の幅
     * @param {{left: number, right: number, viewWidth: number}} selectionSpan - getSelectionViewSpan() の結果
     * @returns {number|null} ずらした左端。重ならない・どちらにも収まらないときは null
     */
    function findDialogLeftAvoidingSelection(dialogLeft, dialogWidth, screenWidth, selectionSpan) {
        var viewLeft = (screenWidth - selectionSpan.viewWidth) / 2;
        var avoidLeft = viewLeft + selectionSpan.left - DIALOG_AVOID_MARGIN;
        var avoidRight = viewLeft + selectionSpan.right + DIALOG_AVOID_MARGIN;
        if (dialogLeft + dialogWidth <= avoidLeft || dialogLeft >= avoidRight) return null;

        var leftSideLeft = avoidLeft - dialogWidth;   /* 選択範囲の左に置くとき / placed left of the selection */
        var rightSideLeft = avoidRight;               /* 選択範囲の右に置くとき / placed right of the selection */
        var fitsLeft = leftSideLeft >= 0;
        var fitsRight = rightSideLeft + dialogWidth <= screenWidth;
        /* 選択範囲が画面の右寄りなら左へ、左寄りなら右へ逃がす / Move away from the side the selection leans to */
        var preferLeft = (avoidLeft + avoidRight) / 2 > screenWidth / 2;
        if (preferLeft && fitsLeft) return leftSideLeft;
        if (fitsRight) return rightSideLeft;
        if (fitsLeft) return leftSideLeft;
        return null;
    }

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // ダイアログの位置と不透明度（再利用パーツ）ここまで / End of the reusable dialog position and opacity
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

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
        var styleDialog = new Window("dialog", getLabel('dialog.title') + " " + SCRIPT_VERSION);
        setupWindow(styleDialog);

        var styleRow = styleDialog.add("group");
        setupRow(styleRow, "fill", FIELD_ROW_SPACING);
        addRowLabel(styleRow, labelText('fieldLabel.style'));

        var styleDropdown = styleRow.add("dropdownlist", undefined, libraryStyleNames);
        styleDropdown.helpTip = getLabel('tooltip.style');
        /* 名前が長くてもダイアログを広げない / A long style name must not widen the dialog */
        styleDropdown.preferredSize.width = DROPDOWN_WIDTH;
        styleDropdown.alignment = ["fill", "center"];

        /* 覚えていたスタイルがライブラリーから消えていることもあるので、無ければ先頭に戻す / A remembered style may be gone from the library, so fall back to the first one */
        var initialStyleName = loadLastStyleName() || DEFAULT_STYLE;
        styleDropdown.selection = 0;
        for (var i = 0; i < libraryStyleNames.length; i++) {
            if (libraryStyleNames[i] === initialStyleName) styleDropdown.selection = i;
        }

        /* ボタンエリア / Button row */
        var buttonRow = addButtonRow(styleDialog);
        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel('button.cancel'), { name: "cancel" });
        btnCancel.helpTip = getLabel('tooltip.cancel');
        var btnApply = buttonRow.rightGroup.add("button", undefined, getLabel('button.apply'), { name: "ok" });
        btnApply.helpTip = getLabel('tooltip.apply');

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
            if (appliedCount === 0) alert(getLabel('alert.applyFailed'));

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
            alert(getLabel('alert.noDocument'));
            return;
        }
        var doc = app.activeDocument;

        /* 取り込みの一時レイヤーを作るので、アクティブレイヤーが使えないなら始めない / A temporary layer is created for the import, so do not start when the active layer cannot take one */
        var activeLayer = doc.activeLayer;
        if (activeLayer.locked || !activeLayer.visible) {
            alert(getLabel('alert.lockedLayer'));
            return;
        }

        if (!getGraphicLibraryFile()) {
            alert(getLabel('alert.noLibrary') + "\n" + GRAPHIC_LIBRARY_NAME);
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
            alert(getLabel('alert.noSelection'));
            return;
        }

        var libraryStyleNames = getLibraryStyleNames(doc);
        if (libraryStyleNames.length === 0) {
            alert(getLabel('alert.noStyle'));
            return;
        }

        var styleDialog = createGraphicStyleDialog(doc, libraryStyleNames, targetItems);
        prepareDialogWindow(styleDialog, SCRIPT_NAME);
        styleDialog.show();
    }

    main();

})();
