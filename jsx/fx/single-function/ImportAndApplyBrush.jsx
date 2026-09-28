#target illustrator
#targetengine "ImportAndApplyBrushEngine"
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
var SCRIPT_VERSION  = "v1.0.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-09-20";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-28";                   /* 更新日 / last updated */

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

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // ダイアログの位置と不透明度（再利用パーツ） / Dialog position and opacity (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内に貼る。
    //    識別子は DIALOG_* / prepareDialogWindow / *DialogLeft* / getSelectionViewSpan の名前
    // 2. スクリプトの先頭（#target の次の行）に #targetengine "<SCRIPT_NAME>Engine" を置く。
    //    #targetengine が無いと $.global が実行ごとに消え、位置を覚えられない。すでにあればそのまま使う
    // 3. ダイアログの show() の直前で prepareDialogWindow(dialog, SCRIPT_NAME) を呼ぶ。
    //    それまでに入れた onShow / onMove / onClose はそのまま生かし、あとに位置の復元・記録をつなぐ
    //      prepareDialogWindow(mainDialog, SCRIPT_NAME);
    //      var dialogResult = mainDialog.show();
    //    同じスクリプトで複数のダイアログを開くときは、2つ目以降のキーを変える（SCRIPT_NAME + "_colorPicker" など）
    //    同じダイアログを何度も開くときも、毎回 show() の直前で呼んでよい（2回目からは選択範囲を測り直すだけ）
    // 4. 初めて開くとき（記録が無いとき）は、スクリプト側の配置（中央・オフセットなど）がそのまま効く
    // 5. 開く位置が選択中のオブジェクトに重なりそうなら左右の反対側へずらす（Illustrator のみ）。
    //    ずらした位置は記録せず、ユーザーが動かしたときだけ記録する
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

    var DIALOG_OPACITY = 0.97;       /* ダイアログの不透明度 / dialog opacity */
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

        var brushDialog = createBrushDialog(doc, libraryBrushNames, targetPaths);
        prepareDialogWindow(brushDialog, SCRIPT_NAME);
        brushDialog.show();
    }

    main();

})();
