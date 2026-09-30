#target illustrator
#targetengine "RenameAssetsEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

アセット書き出しパネルに登録されたアセットの名前を、一括で変更します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/RenameAssets.md

### Overview

Renames the assets registered in the Asset Export panel in bulk.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/RenameAssets.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "RenameAssets";                 /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.8";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-08-20";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-01";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/RenameAssets.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/RenameAssets.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // レイアウト / Layout
    // =========================================

    // UIレイアウト（再利用パーツ） / UI layout (reusable)

    /* ウィンドウ・パネルの余白と間隔 / Window & panel margins and spacing */
    var WINDOW_MARGINS = 16;                 /* ウィンドウ外周の余白 / window margin */
    var WINDOW_SPACING = 12;                 /* ウィンドウ内の要素間隔 / window spacing */
    var PANEL_MARGINS  = [16, 20, 16, 12];   /* パネル余白 [左,上,右,下] / panel margins */
    var PANEL_SPACING  = 12;                 /* パネル内の要素間隔 / panel spacing */
    var COLUMN_SPACING = 12;                 /* 2カラムの間隔 / gap between columns */
    var TAB_MARGINS    = [15, 20, 5, 10];    /* タブ余白 [左,上,右,下] / tab margins */

    /**
     * ウィンドウの共通設定
     * @param {Window} targetWindow - 対象のウィンドウ
     * @param {number} [spacing] - 要素間隔（省略時は WINDOW_SPACING）
     * @returns {void}
     */
    function setupWindow(targetWindow, spacing) {
        targetWindow.orientation = "column";
        targetWindow.alignChildren = "fill";
        targetWindow.margins = WINDOW_MARGINS;
        targetWindow.spacing = (typeof spacing === "number") ? spacing : WINDOW_SPACING;
    }

    /**
     * パネルの共通設定（子は幅いっぱい。ボタンは alignment = "left" で広げない）
     * @param {Panel} targetPanel - 対象のパネル
     * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function setupPanel(targetPanel, spacing) {
        targetPanel.orientation = "column";
        targetPanel.alignChildren = ["fill", "top"];
        targetPanel.alignment = "fill";
        targetPanel.margins = PANEL_MARGINS;
        targetPanel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /**
     * タブの共通設定
     * @param {Tab} targetTab - 対象のタブ
     * @param {number} [spacing] - 要素間隔（省略時は変えない）
     * @returns {void}
     */
    function setupTab(targetTab, spacing) {
        targetTab.orientation = "column";
        targetTab.alignChildren = "fill";
        targetTab.margins = TAB_MARGINS;
        if (typeof spacing === "number") targetTab.spacing = spacing;
    }

    /**
     * 横並びの行グループの共通設定（ボタン列など）。
     * alignment と alignChildren を対で指定し、中のボタンが横に伸びたり天地がずれたりしないようにする
     * @param {Group} rowGroup - 対象のグループ
     * @param {string|string[]} [rowAlignment] - 横方向の alignment（省略時は "left"）。配列ならそのまま使う
     * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function setupRow(rowGroup, rowAlignment, spacing) {
        rowGroup.orientation = "row";
        rowGroup.alignment = (rowAlignment instanceof Array) ? rowAlignment : [rowAlignment || "left", "center"];
        rowGroup.alignChildren = ["left", "center"];
        rowGroup.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /**
     * ボタンの高さを指定した px だけ詰める（レイアウトが決まったあとに呼ぶ）
     * @param {Button} targetButton - 対象のボタン
     * @param {number} trimPixels - 詰める量（px）
     * @returns {void}
     */
    function trimButtonHeight(targetButton, trimPixels) {
        /* レイアウト前は size が無い / size is not set until the layout runs */
        if (!targetButton.size) return;
        targetButton.size = [targetButton.size.width, targetButton.size.height - trimPixels];
    }

    // UIレイアウト（再利用パーツ）ここまで / End of the reusable UI layout

    var FIELD_LABEL_WIDTH = 120;              /* 項目名の幅（揃える） / Width of the field labels */
    var FIELD_CHARS = 30;                     /* 入力欄の幅（文字数） / Width of the text fields */
    var STATUS_CHARS = 30;                    /* ステータス表示の幅（文字数） / Width of the status text */

    // =========================================
    // ローカライズ / Localization
    // =========================================

    // ローカライズ（再利用パーツ） / Localization (reusable)

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

    // ローカライズ（再利用パーツ）ここまで / End of the reusable localization

    /* 日英ラベル定義。status.preview の %1・%2 に件数を差し込む / Japanese-English labels; status.preview takes counts in %1 and %2 */
    var LABELS = {
        dialog: {
            title: { ja: "グラフィックスタイルなどのアセット名のリネーム", en: "Rename Asset Names (Graphic Styles, etc.)" }
        },
        panel: {
            target: { ja: "対象", en: "Target" }
        },
        radio: {
            style: { ja: "グラフィックスタイル", en: "Graphic Styles" },
            brush: { ja: "ブラシ", en: "Brushes" },
            swatch: { ja: "スウォッチ", en: "Swatches" },
            symbol: { ja: "シンボル", en: "Symbols" }
        },
        fieldLabel: {
            find: { ja: "検索文字列", en: "Find" },
            replace: { ja: "置換文字列", en: "Replace" }
        },
        checkbox: {
            regex: { ja: "正規表現", en: "Regular Expression" },
            ignoreCase: { ja: "大文字／小文字を無視", en: "Ignore Case" }
        },
        status: {
            none: { ja: "—", en: "—" },
            noFind: { ja: "検索文字列なし", en: "Find string is empty" },
            preview: { ja: "プレビュー: 変更 %1／全 %2", en: "Preview: %1 changed / %2 total" }
        },
        tooltip: {
            find: { ja: "アセット名の中から探す文字列です。", en: "Text to look for in the asset names." },
            replace: {
                ja: "置き換える文字列です。空欄にすると検索文字列を削除します。",
                en: "Replacement text. Leave blank to delete the found text."
            },
            regex: {
                ja: "検索文字列を正規表現として扱います（$1 などの後方参照も使えます）。",
                en: "Treats the search text as a regular expression, including back-references such as $1."
            },
            ignoreCase: { ja: "大文字と小文字を区別せずに探します。", en: "Matches without regard to letter case." },
            target: { ja: "リネームするアセットの種類です。", en: "Which kind of asset to rename." },
            preview: {
                ja: "置換後の名前をアセットに仮に反映し、変更される件数を表示します。",
                en: "Applies the new names to the assets temporarily and shows how many change."
            }
        },
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            preview: { ja: "プレビュー", en: "Preview" },
            ok: { ja: "OK", en: "OK" }
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noAssets: {
                ja: "このドキュメントには対象となる項目（スタイル／ブラシ／スウォッチ／シンボル）がありません。",
                en: "No styles/brushes/swatches/symbols in this document."
            },
            noFind: { ja: "「検索」文字列が空です。", en: "Find string is empty." },
            noTarget: { ja: "選択した対象に名前付き項目がありません。", en: "No items found for selected target." }
        }
    };

    // =========================================
    // 名前の置換 / Name replacement
    // =========================================

    /**
     * 対象の種類に対応するアセットのコレクションを返す
     * @param {string} targetKey - "style" / "brush" / "swatch" / "symbol"
     * @returns {Object|null} コレクション（ドキュメントが無ければ null）
     */
    function getCollectionByKey(targetKey) {
        var doc = app.activeDocument;
        if (!doc) return null;
        if (targetKey === 'brush') return doc.brushes;
        if (targetKey === 'swatch') return doc.swatches;
        if (targetKey === 'symbol') return doc.symbols;
        return doc.graphicStyles;
    }

    /**
     * 置換関数を作る
     * @param {string} findStr - 検索文字列
     * @param {string} replStr - 置換文字列
     * @param {boolean} useRegex - 正規表現として扱うか
     * @param {boolean} ignoreCase - 大文字と小文字を区別しないか
     * @returns {Function} 元の名前を受け取り、置換後の名前（一致しなければ null）を返す関数
     */
    function buildReplacer(findStr, replStr, useRegex, ignoreCase) {
        if (useRegex || ignoreCase) {
            /* ［大文字／小文字を無視］だけのときも検索文字列を正規表現として組み立てる
               With only Ignore Case, the find text is still compiled as a pattern */
            var findPattern = null;
            try {
                findPattern = new RegExp(findStr, (useRegex && !ignoreCase) ? 'g' : 'gi');
            } catch (e) {
                /* 正規表現として不正な検索文字列 / invalid pattern */
                findPattern = null;
            }
            if (!findPattern) {
                return function () {
                    return null;
                };
            }
            return function (oldName) {
                return findPattern.test(oldName) ? oldName.replace(findPattern, replStr) : null;
            };
        }
        return function (oldName) {
            return oldName.indexOf(findStr) > -1 ? oldName.split(findStr).join(replStr) : null;
        };
    }

    /**
     * 角かっこで囲まれた既定の名前（[なし] など）かを返す
     * @param {string} assetName - アセット名
     * @returns {boolean} 既定の名前なら true
     */
    function isBracketedName(assetName) {
        /* 正規表現はリテラルで書く（ExtendScript の版によってはスコープの問題が出るため）
           Keep the literal inline to avoid scope issues in some ExtendScript versions */
        return /^\[.*\]$/.test(assetName);
    }

    /**
     * 名前の重複を避けるため、" (2)"、" (3)" … を付けて使われていない名前にする
     * @param {string} baseName - 希望する名前
     * @param {Object} usedNames - 使用済みの名前（名前 → true）
     * @returns {string} 使われていない名前（9999 を超えたらそこで打ち切る）
     */
    function ensureUniqueName(baseName, usedNames) {
        if (!usedNames[baseName]) return baseName;
        var i = 2;
        var candidate = baseName + " (" + i + ")";
        while (usedNames[candidate]) {
            i++;
            candidate = baseName + " (" + i + ")";
            /* 念のための上限 / safeguard */
            if (i > 9999) break;
        }
        return candidate;
    }

    /**
     * コレクションの名前を文字列として読み取る（名前が無いものは空文字）
     * @param {Object} collection - アセットのコレクション
     * @returns {string[]} 名前の配列
     */
    function readAssetNames(collection) {
        var assetNames = [];
        var totalCount = collection.length;
        for (var i = 0; i < totalCount; i++) {
            var rawName = collection[i].name;
            assetNames.push((rawName === undefined || rawName === null) ? "" : String(rawName));
        }
        return assetNames;
    }

    /**
     * 名前の配列から、使用済みの名前の表を作る
     * @param {string[]} assetNames - 名前の配列
     * @returns {Object} 名前 → true の表
     */
    function buildUsedNameMap(assetNames) {
        var usedNames = {};
        for (var i = 0; i < assetNames.length; i++) usedNames[assetNames[i]] = true;
        return usedNames;
    }

    /**
     * 変更後の名前を計算する（アセットはまだ変更しない）
     * @param {Object} collection - アセットのコレクション
     * @param {string} findStr - 検索文字列
     * @param {string} replStr - 置換文字列
     * @param {boolean} useRegex - 正規表現として扱うか
     * @param {boolean} ignoreCase - 大文字と小文字を区別しないか
     * @returns {{names: string[], changed: number}} 変更後の名前と、変更される件数
     */
    function computeProposedNames(collection, findStr, replStr, useRegex, ignoreCase) {
        var originalNames = readAssetNames(collection);
        var usedNames = buildUsedNameMap(originalNames);
        var proposedNames = originalNames.slice(0);
        var changedCount = 0;

        var replacer = buildReplacer(findStr, replStr, useRegex, ignoreCase);
        for (var j = 0; j < originalNames.length; j++) {
            var oldStr = originalNames[j];
            if (isBracketedName(oldStr)) continue;

            var replacedName = replacer(oldStr);
            if (replacedName !== null && replacedName !== oldStr) {
                /* 元の名前と、先に決めた名前の両方と重ならないようにする / avoid both original and earlier names */
                var uniqueName = ensureUniqueName(replacedName, usedNames);
                proposedNames[j] = uniqueName;
                usedNames[uniqueName] = true;
                changedCount++;
            }
        }
        return { names: proposedNames, changed: changedCount };
    }

    /**
     * 計算した名前をアセットに反映する
     * @param {Object} collection - アセットのコレクション
     * @param {string[]} newNames - 反映する名前
     * @returns {number} 変更した件数
     */
    function applyNames(collection, newNames) {
        var appliedCount = 0;
        for (var i = 0; i < collection.length && i < newNames.length; i++) {
            if (collection[i].name !== newNames[i]) {
                /* 名前を変更できないアセットがある / some assets reject a new name */
                try {
                    collection[i].name = newNames[i];
                    appliedCount++;
                } catch (e) { }
            }
        }
        return appliedCount;
    }

    /**
     * コレクション内のアセット名を置換する（確定実行）
     * @param {Object} collection - アセットのコレクション
     * @param {string} findStr - 検索文字列
     * @param {string} replStr - 置換文字列
     * @param {boolean} useRegex - 正規表現として扱うか
     * @param {boolean} ignoreCase - 大文字と小文字を区別しないか
     * @returns {number} 変更した件数
     */
    function renameItemsByCollection(collection, findStr, replStr, useRegex, ignoreCase) {
        /* 既存の名前を控えて重複を避ける / snapshot existing names for collision checks */
        var usedNames = buildUsedNameMap(readAssetNames(collection));
        var replacer = buildReplacer(findStr, replStr, useRegex, ignoreCase);
        var totalCount = collection.length;
        var changedCount = 0;

        for (var j = 0; j < totalCount; j++) {
            var asset = collection[j];
            var oldName = asset.name;
            if (oldName === undefined || oldName === null) continue;

            var oldStr = String(oldName);
            if (isBracketedName(oldStr)) continue;

            var replacedName = replacer(oldStr);
            if (replacedName === null || replacedName === oldStr) continue;

            var uniqueName = ensureUniqueName(replacedName, usedNames);
            if (uniqueName === oldStr) continue;

            /* 名前を変更できないアセットは飛ばす / skip assets that reject the new name */
            try {
                asset.name = uniqueName;
                /* 元の名前は表に残るが、新しい重複を防げれば十分 / old names stay in the map; that is fine */
                usedNames[uniqueName] = true;
                changedCount++;
            } catch (e) { }
        }
        return changedCount;
    }

    // ダイアログの位置と不透明度（再利用パーツ） / Dialog position and opacity (reusable)

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
            /* 文字ツールで文字を選択しているときは TextRange が返り、[0] が無い / Selecting characters with the Type tool returns a TextRange, which has no [0] */
            if (!selectedItems || selectedItems.typename === "TextRange" || !selectedItems.length || !selectedItems[0].visibleBounds) return null;
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

    // ダイアログの位置と不透明度（再利用パーツ）ここまで / End of the reusable dialog position and opacity

    // =========================================
    // ダイアログの部品 / Dialog helpers
    // =========================================

    /**
     * タイトル付きのパネルを追加する（共通レイアウト適用）
     * @param {Object} parentGroup - 追加先のコンテナ
     * @param {string} title - パネルのタイトル
     * @returns {Panel} 追加したパネル
     */
    function addPanel(parentGroup, title) {
        var titledPanel = parentGroup.add("panel", undefined, title);
        setupPanel(titledPanel);
        return titledPanel;
    }

    /**
     * 横一列のグループを追加する
     * @param {Object} parentGroup - 追加先のコンテナ
     * @param {string[]} [alignChildren] - 子要素の揃え（["left", "center"] など）
     * @param {string|string[]} [alignment] - グループ自身の揃え（["fill", "bottom"] など）
     * @param {number} [spacing] - 要素間隔
     * @returns {Group} 追加したグループ
     */
    function addRow(parentGroup, alignChildren, alignment, spacing) {
        var rowGroup = parentGroup.add("group");
        rowGroup.orientation = "row";
        if (alignChildren) rowGroup.alignChildren = alignChildren;
        if (alignment) rowGroup.alignment = alignment;
        if (spacing != null) rowGroup.spacing = spacing;
        return rowGroup;
    }

    /**
     * 右揃えの項目名と入力欄の行を追加する
     * @param {Object} parentGroup - 追加先のコンテナ
     * @param {string} fieldLabelText - 項目名
     * @param {number} labelWidth - 項目名の幅
     * @param {number} [editChars] - 入力欄の幅（文字数。省略時は 30）
     * @returns {EditText} 追加した入力欄
     */
    function addLabeledEdit(parentGroup, fieldLabelText, labelWidth, editChars) {
        var rowGroup = addRow(parentGroup, ["fill", "center"]);
        var fieldLabel = rowGroup.add("statictext", undefined, fieldLabelText);
        fieldLabel.preferredSize.width = labelWidth;
        fieldLabel.justify = "right";
        var inputField = rowGroup.add("edittext", undefined, "");
        inputField.characters = editChars || 30;
        return inputField;
    }

    /**
     * チェックボックスを追加する
     * @param {Object} parentGroup - 追加先のコンテナ
     * @param {string} checkboxLabel - 表示名
     * @param {boolean} defaultValue - 初期値
     * @returns {Checkbox} 追加したチェックボックス
     */
    function addCheckbox(parentGroup, checkboxLabel, defaultValue) {
        var checkboxControl = parentGroup.add("checkbox", undefined, checkboxLabel);
        checkboxControl.value = !!defaultValue;
        return checkboxControl;
    }

    /**
     * ラジオボタンのグループを追加する
     * @param {Object} parentGroup - 追加先のコンテナ
     * @param {string[]} radioLabels - 各ラジオボタンの表示名
     * @param {string} [orientation] - 並びの向き（省略時は "row"）
     * @param {number[]} [margins] - 余白
     * @param {number} [spacing] - 要素間隔
     * @returns {{group: Group, buttons: RadioButton[], select: Function, getSelectedIndex: Function, onChange: Function}} グループと操作用の関数
     */
    function addRadioGroup(parentGroup, radioLabels, orientation, margins, spacing) {
        var radioGroup = parentGroup.add("group");
        radioGroup.orientation = orientation || "row";
        if (margins) radioGroup.margins = margins;
        if (spacing != null) radioGroup.spacing = spacing;

        var radioButtons = [];
        for (var i = 0; i < radioLabels.length; i++) {
            radioButtons[i] = radioGroup.add("radiobutton", undefined, radioLabels[i]);
        }

        var changeHandlers = [];

        /**
         * 選択中のラジオボタンの番号を返す
         * @returns {number} 番号（選択なしは -1）
         */
        function getSelectedIndex() {
            for (var i = 0; i < radioButtons.length; i++) {
                if (radioButtons[i].value) return i;
            }
            return -1;
        }

        /**
         * 登録された変更時の処理を呼ぶ
         * @returns {void}
         */
        function fireChange() {
            var selectedIndex = getSelectedIndex();
            for (var h = 0; h < changeHandlers.length; h++) changeHandlers[h](selectedIndex);
        }

        for (var j = 0; j < radioButtons.length; j++) {
            radioButtons[j].onClick = fireChange;
        }

        return {
            group: radioGroup,
            buttons: radioButtons,
            select: function (selectedIndex, notify) {
                for (var i = 0; i < radioButtons.length; i++) radioButtons[i].value = (i === selectedIndex);
                if (notify === true) fireChange();
            },
            getSelectedIndex: getSelectedIndex,
            onChange: function (handler) {
                if (typeof handler === 'function') changeHandlers[changeHandlers.length] = handler;
            }
        };
    }

    /**
     * 中央寄せのテキスト（ステータス表示など）を追加する
     * @param {Object} parentGroup - 追加先のコンテナ
     * @param {number} [chars] - 幅（文字数。省略時は 30）
     * @returns {StaticText} 追加したテキスト
     */
    function addCenteredStatic(parentGroup, chars) {
        var rowGroup = addRow(parentGroup, ["center", "center"], "center");
        var centeredText = rowGroup.add("statictext", undefined, "");
        centeredText.characters = chars || 30;
        centeredText.alignment = "center";
        centeredText.justify = "center";
        return centeredText;
    }

    // ボタン行（再利用パーツ） / Button row (reusable)

    var BUTTON_ROW_TOP_MARGIN = 5; /* ボタン行の上の余白 / top margin of the button row */
    var BUTTON_ROW_BOTTOM_MARGIN = 14; /* ボタン行の下の余白。ダイアログの下余白と合わせて約30px（Illustrator 標準のダイアログに合わせる） / bottom margin; with the dialog margin about 30px, like Illustrator's own dialogs */
    var BUTTON_ROW_SPACING = 10;   /* ボタンどうしの間隔 / spacing between buttons */
    var BUTTON_ROW_CENTER_MAX_WIDTH = 200; /* 右のボタンだけの行を中央に置く、ダイアログの内側の最大幅（px、左右の余白を除く）。広いダイアログは右揃え / max inner dialog width (px, margins excluded) that centers a right-only row; wider dialogs keep it right-aligned */

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
        btnRowGroup.margins = [0, BUTTON_ROW_TOP_MARGIN, 0, BUTTON_ROW_BOTTOM_MARGIN];
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

    /**
     * 左のグループにボタンが無い（右のボタンだけの）行を、ダイアログの幅に合わせて揃える。
     * 内側の幅（左右の余白を除く）が BUTTON_ROW_CENTER_MAX_WIDTH 以下なら左右中央、それより広ければ右揃えのまま。
     * 幅はレイアウトが決まるまで分からないので、ダイアログを表示した時点（show イベント）で判定する。
     * ボタンをすべて足したあと、show() の前に呼ぶ。centered で作った行や、左にボタンがある行はそのまま
     * @param {{rowGroup: Group, leftGroup: Group|null, rightGroup: Group|null}} buttonRow - addButtonRow() の戻り値
     * @returns {void}
     */
    function alignRightOnlyButtonRow(buttonRow) {
        if (!buttonRow.leftGroup || buttonRow.leftGroup.children.length > 0) return;
        var dialogWindow = buttonRow.rowGroup.window;
        dialogWindow.addEventListener("show", function () {
            if (!buttonRow.leftGroup) return;
            var btnRowGroup = buttonRow.rowGroup;
            /* 行の幅＝ダイアログの内側の幅（左右の余白を除く）/ The row spans the dialog's inner width (margins excluded) */
            if (!btnRowGroup.size || btnRowGroup.size.width > BUTTON_ROW_CENTER_MAX_WIDTH) return;
            /* 左のグループとスペーサーを外し、右のグループだけを中央に置く / Drop the left group and the spacer so only the right group remains, centered */
            btnRowGroup.remove(buttonRow.leftGroup);
            btnRowGroup.remove(btnRowGroup.children[0]); /* 左のグループを外すと先頭はスペーサー / the spacer is first once the left group is gone */
            btnRowGroup.alignment = ["center", "bottom"];
            btnRowGroup.alignChildren = ["center", "center"];
            buttonRow.leftGroup = null;
            dialogWindow.layout.layout(true);
        });
    }

    // ボタン行（再利用パーツ）ここまで / End of the reusable button row

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * ダイアログを表示し、OK なら置換の設定を返す（プレビューは名前を仮に変更し、キャンセルで戻す）
     * @returns {{find: string, repl: string, regex: boolean, ignoreCase: boolean, target: string}|null} 設定。キャンセル時は null
     */
    function showDialog() {
        var renameDialog = new Window("dialog", getLabel('dialog.title') + " " + SCRIPT_VERSION);
        setupWindow(renameDialog);

        /* 対象（上段）/ Target radio group */
        var targetPanel = addPanel(renameDialog, getLabel('panel.target'));
        var targetRadios = addRadioGroup(targetPanel,
            [getLabel('radio.style'), getLabel('radio.brush'), getLabel('radio.swatch'), getLabel('radio.symbol')], 'row');
        var brushRadio = targetRadios.buttons[1];
        var swatchRadio = targetRadios.buttons[2];
        var symbolRadio = targetRadios.buttons[3];
        for (var i = 0; i < targetRadios.buttons.length; i++) {
            targetRadios.buttons[i].helpTip = getLabel('tooltip.target');
        }
        targetRadios.select(0);

        var findInput = addLabeledEdit(renameDialog, labelText('fieldLabel.find'), FIELD_LABEL_WIDTH, FIELD_CHARS);
        findInput.helpTip = getLabel('tooltip.find');
        var replaceInput = addLabeledEdit(renameDialog, labelText('fieldLabel.replace'), FIELD_LABEL_WIDTH, FIELD_CHARS);
        replaceInput.helpTip = getLabel('tooltip.replace');

        /* 正規表現と大文字／小文字（横並び・中央）/ Regex and ignore-case options, centered */
        var optionsRow = addRow(renameDialog, ["center", "center"], "center");
        var regexCheckbox = addCheckbox(optionsRow, getLabel('checkbox.regex'), false);
        regexCheckbox.helpTip = getLabel('tooltip.regex');
        var ignoreCaseCheckbox = addCheckbox(optionsRow, getLabel('checkbox.ignoreCase'), false);
        ignoreCaseCheckbox.helpTip = getLabel('tooltip.ignoreCase');

        /* プレビューの結果（ボタンで実行）/ Preview status, updated by the button */
        var statusText = addCenteredStatic(renameDialog, STATUS_CHARS);

        /* プレビュー前の名前と、その対象 / Names before the preview and their target */
        var snapshotNames = null;
        var snapshotTargetKey = null;
        var accepted = false;

        /**
         * 選択中の対象の種類を返す
         * @returns {string} "style" / "brush" / "swatch" / "symbol"
         */
        function getTargetKey() {
            return brushRadio.value ? 'brush' : (swatchRadio.value ? 'swatch' : (symbolRadio.value ? 'symbol' : 'style'));
        }

        /**
         * 控えた名前に戻してから処理を実行する（プレビューが重ならないように）。控えが無ければ先に取る
         * @param {string} targetKey - 対象の種類
         * @param {Function} workFn - 戻したコレクションを受け取る処理
         * @returns {Object|null} 対象のコレクション
         */
        function withSnapshot(targetKey, workFn) {
            var targetCollection = getCollectionByKey(targetKey);
            if (!targetCollection) return null;
            if (!snapshotNames || snapshotTargetKey !== targetKey) {
                snapshotNames = [];
                for (var i = 0; i < targetCollection.length; i++) snapshotNames[i] = targetCollection[i].name;
                snapshotTargetKey = targetKey;
            }
            restoreSnapshotNames(targetCollection);
            /* 名前の読み書きで例外が出てもダイアログは閉じない / keep the dialog alive on DOM errors */
            try {
                workFn && workFn(targetCollection);
            } catch (e) { }
            return targetCollection;
        }

        /**
         * 控えた名前をコレクションに書き戻す
         * @param {Object} targetCollection - 書き戻すコレクション
         * @returns {void}
         */
        function restoreSnapshotNames(targetCollection) {
            for (var i = 0; i < targetCollection.length && i < snapshotNames.length; i++) {
                /* 名前を変更できないアセットがある / some assets reject a new name */
                try {
                    targetCollection[i].name = snapshotNames[i];
                } catch (e) { }
            }
        }

        /**
         * キャンセル時にプレビュー前の名前へ戻す
         * @returns {void}
         */
        function restoreSnapshot() {
            if (!snapshotNames) return;
            var targetCollection = getCollectionByKey(snapshotTargetKey);
            if (!targetCollection) return;
            restoreSnapshotNames(targetCollection);
        }

        /**
         * 置換後の名前を仮に反映し、件数を表示する
         * @returns {void}
         */
        function updatePreview() {
            var findText = String(findInput.text || '');
            var replaceText = String(replaceInput.text || '');
            var useRegex = regexCheckbox.value === true;
            var ignoreCase = ignoreCaseCheckbox.value === true;
            if (!findText) {
                statusText.text = getLabel('status.noFind');
                return;
            }

            var previewPlan = null;
            var targetCollection = withSnapshot(getTargetKey(), function (restoredCollection) {
                previewPlan = computeProposedNames(restoredCollection, findText, replaceText, useRegex, ignoreCase);
                applyNames(restoredCollection, previewPlan.names);
            });
            if (!targetCollection) {
                statusText.text = getLabel('status.none');
                return;
            }
            statusText.text = getLabel('status.preview', [previewPlan.changed, targetCollection.length]);
            app.redraw();
        }

        /* 対象を切り替えたら控えを捨て、次のプレビューで取り直す / a new target needs a new snapshot */
        targetRadios.onChange(function () {
            snapshotNames = null;
        });

        /* ボタン行（左=プレビュー／右=キャンセル・OK）/ Button row: Preview on the left, Cancel and OK on the right */
        var buttonRow = addButtonRow(renameDialog);
        var btnPreview = buttonRow.leftGroup.add("button", undefined, getLabel('button.preview'), { name: "preview" });
        btnPreview.helpTip = getLabel('tooltip.preview');
        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel('button.cancel'), { name: "cancel" });
        var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel('button.ok'), { name: "ok" });

        btnOK.onClick = function () {
            /* プレビューで変えた名前はそのまま残す / keep the previewed names */
            accepted = true;
            snapshotNames = null;
            renameDialog.close();
        };
        btnCancel.onClick = function () {
            /* プレビューを反映していたら元の名前に戻す / revert the preview */
            try {
                restoreSnapshot();
            } catch (e) { }
            renameDialog.close();
        };
        btnPreview.onClick = function () {
            updatePreview();
        };

        findInput.active = true;

        alignRightOnlyButtonRow(buttonRow);
        prepareDialogWindow(renameDialog, SCRIPT_NAME);
        renameDialog.show();

        return accepted ? {
            find: String(findInput.text || ""),
            repl: String(replaceInput.text || ""),
            regex: regexCheckbox.value,
            ignoreCase: ignoreCaseCheckbox.value,
            target: getTargetKey()
        } : null;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * ドキュメントを確認し、ダイアログの設定でアセット名を置換する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel('alert.noDocument'));
            return;
        }
        var doc = app.activeDocument;
        /* 対象になるコレクションが1つでもあるか / at least one eligible collection */
        var hasAnyAssets = (doc.graphicStyles && doc.graphicStyles.length) ||
            (doc.brushes && doc.brushes.length) ||
            (doc.swatches && doc.swatches.length) ||
            (doc.symbols && doc.symbols.length);
        if (!hasAnyAssets) {
            alert(getLabel('alert.noAssets'));
            return;
        }

        var renameOptions = showDialog();
        if (!renameOptions) return;

        if (!renameOptions.find) {
            alert(getLabel('alert.noFind'));
            return;
        }

        var targetCollection = getCollectionByKey(renameOptions.target);
        if (!targetCollection || targetCollection.length === 0) {
            alert(getLabel('alert.noTarget'));
            return;
        }

        renameItemsByCollection(targetCollection, renameOptions.find, renameOptions.repl, renameOptions.regex, renameOptions.ignoreCase);
    }

    main();

})();
