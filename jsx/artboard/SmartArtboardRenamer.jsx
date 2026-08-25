#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

接頭辞・接尾辞と参照テキスト（最前面のテキスト、指定レイヤーのテキスト、元のアートボード名、任意の文字列）を組み合わせて、アートボード名を一括で変更します。
ダイアログでプレビューを確認しながら、アートボードの並び替えや1件ずつの手動リネームもできます。

詳細は README を参照してください。

### Overview

Renames the artboards in bulk from a prefix and suffix plus a reference text — the frontmost text, text on a chosen layer, the original name, or a string you type.
The dialog offers a preview, along with reordering and renaming artboards one at a time.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SmartArtboardRenamer";         /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.5.4";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-05-09";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-08-07";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartArtboardRenamer.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartArtboardRenamer.md
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/ne0934ee22972"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // レイアウト / Layout
    // =========================================

    /* ウィンドウ・パネルの余白と間隔 / Window & panel margins and spacing */
    var WINDOW_MARGINS = 16;                 /* ウィンドウ外周の余白 / window margin */
    var WINDOW_SPACING = 12;                 /* ウィンドウ内の要素間隔 / window spacing */
    var PANEL_MARGINS  = [16, 20, 16, 12];   /* パネル余白 [左,上,右,下] / panel margins */
    var PANEL_SPACING  = 12;                 /* パネル内の要素間隔 / panel spacing */
    var COLUMN_SPACING = 12;                 /* 2カラムの間隔 / gap between columns */

    /* 並び替えリストの寸法 / Reorder list metrics */
    var VISIBLE_ROW_COUNT       = 12;        /* 一度に表示する行数（超えるとスクロール） / rows shown at once */
    var ENTRY_ROW_HEIGHT        = 24;        /* 1行の高さの目安（スクロールバーの高さ算出用） / estimated row height, used to size the scrollbar */
    var SCROLLBAR_WIDTH         = 16;        /* スクロールバーの幅 / scrollbar width */
    var ORDER_COLUMN_WIDTH      = 24;        /* 「順」列の幅 / order column width */
    var TARGET_COLUMN_WIDTH     = 28;        /* 「対象」列の幅 / checkbox column width */
    var CURRENT_NAME_WIDTH      = 140;       /* 「元の名前」列の幅 / current name column width */
    var ARROW_COLUMN_WIDTH      = 14;        /* 矢印列の幅 / arrow column width */
    var NEW_NAME_WIDTH          = 160;       /* 「新しい名前」列の幅 / new name column width */
    var ROW_SPACING             = 6;         /* 行内の要素間隔 / spacing inside a row */

    /* ボタンの寸法 / Button metrics */
    var MOVE_BUTTON_WIDTH   = 56;            /* 並び替えボタンの幅 / move button width */
    var MOVE_BUTTON_HEIGHT  = 22;            /* 並び替えボタンの高さ / move button height */
    var TOKEN_BUTTON_WIDTH  = 28;            /* トークンボタンの既定幅 / default token button width */
    var TOKEN_BUTTON_HEIGHT = 20;            /* トークンボタンの高さ / token button height */

    /**
     * ウィンドウの共通設定を適用する
     * @param {Window} win - 対象ウィンドウ
     * @param {number} [spacing] - 要素間隔（省略時は WINDOW_SPACING）
     * @returns {void}
     */
    function setupWindow(win, spacing) {
        win.orientation = "column";
        win.alignChildren = "fill";
        win.margins = WINDOW_MARGINS;
        win.spacing = (typeof spacing === "number") ? spacing : WINDOW_SPACING;
    }

    /**
     * パネルの共通設定を適用する
     * @param {Panel} panel - 対象パネル
     * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function setupPanel(panel, spacing) {
        panel.orientation = "column";
        panel.alignChildren = ["fill", "top"];
        panel.alignment = "fill";
        panel.margins = PANEL_MARGINS;
        panel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /**
     * 行グループの共通設定を適用する（ボタン列など）
     * @param {Group} group - 対象グループ
     * @param {string} [alignment] - 配置（省略時は "left"）
     * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function setupRow(group, alignment, spacing) {
        group.orientation = "row";
        group.alignChildren = ["left", "center"];
        group.alignment = alignment || "left";
        group.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * 現在のUI言語を判定する
     * @returns {string} "ja" または "en"
     */
    function getUILanguage() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = getUILanguage();

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "アートボードのリネームとソート", en: "Smart Artboard Renamer" }
        },
        panel: {
            rename:       { ja: "リネーム", en: "Rename" },
            prefix:       { ja: "接頭辞", en: "Prefix" },
            suffix:       { ja: "接尾辞", en: "Suffix" },
            textSource:   { ja: "参照テキスト", en: "Text Source" },
            targetBoards: { ja: "対象アートボード", en: "Target Artboards" },
            reorder:      { ja: "並び替え／リネーム", en: "Reorder / Rename" }
        },
        radio: {
            originalName: { ja: "元のアートボード名", en: "Original Artboard Name" },
            customText:   { ja: "指定", en: "Custom" },
            frontmost:    { ja: "最前面のテキスト", en: "Frontmost Text" },
            layer:        { ja: "レイヤー", en: "Layer" },
            allBoards:    { ja: "すべて", en: "All" },
            rangeBoards:  { ja: "指定範囲", en: "Range" }
        },
        columnHeader: {
            order:        { ja: "順", en: "#" },
            selection:    { ja: "対象", en: "Sel" },
            currentName:  { ja: "元の名前", en: "Original Name" },
            newName:      { ja: "新しい名前", en: "New Name" }
        },
        button: {
            moveToTop:    { ja: "↑ 先頭へ", en: "↑ Top" },
            moveUp:       { ja: "↑ 上へ", en: "↑ Up" },
            moveDown:     { ja: "↓ 下へ", en: "↓ Down" },
            moveToBottom: { ja: "↓ 末尾へ", en: "↓ Bottom" },
            refresh:      { ja: "更新", en: "Refresh" },
            cancel:       { ja: "キャンセル", en: "Cancel" },
            ok:           { ja: "OK", en: "OK" }
        },
        alert: {
            noLayer: {
                ja: "指定されたレイヤーが見つからないか、非表示です。",
                en: "The specified layer was not found or is hidden."
            },
            needSettings: {
                ja: "接頭辞・接尾辞のいずれかを入力するか、参照テキストを指定してください。",
                en: "Enter a prefix or suffix, or choose a text source."
            },
            emptyName: {
                ja: "{n} 番目の新しい名前が空です。名前を入力してください。",
                en: "Artboard {n}: new name is empty. Please enter a name."
            }
        }
    };

    /**
     * ラベルを現在のUI言語で取得する
     * @param {...string} labelPath - たどるキー（例: getLabel("dialog", "title")）
     * @returns {string} ローカライズ済みラベル
     */
    function getLabel() {
        var node = LABELS;
        for (var i = 0; i < arguments.length; i++) {
            if (node == null) break;
            node = node[arguments[i]];
        }
        return (node && node[uiLang] != null) ? node[uiLang] : "";
    }

    // =========================================
    // ユーティリティ / Utilities
    // =========================================

    /**
     * 配列に値が含まれるかを判定する
     * @param {Array} list - 検索対象の配列
     * @param {*} value - 探す値
     * @returns {boolean} 含まれていれば true
     */
    function containsValue(list, value) {
        for (var i = 0; i < list.length; i++) {
            if (list[i] === value) return true;
        }
        return false;
    }

    /**
     * "1-3,5" 形式の範囲文字列を 0 始まりのインデックス配列に変換する
     * @param {string} rangeText - 範囲文字列
     * @returns {number[]} 0 始まりのインデックス配列
     */
    function parseArtboardRange(rangeText) {
        var indices = [];
        var parts = rangeText.split(",");
        for (var i = 0; i < parts.length; i++) {
            var part = parts[i].replace(/\s+/g, "");
            if (/^\d+$/.test(part)) {
                indices.push(parseInt(part, 10) - 1);
            } else if (/^\d+-\d+$/.test(part)) {
                var bounds = part.split("-");
                var rangeStart = parseInt(bounds[0], 10);
                var rangeEnd = parseInt(bounds[1], 10);
                /* "5-1" のような逆順指定も昇順として扱う / Treat a reversed range like "5-1" as ascending */
                if (rangeStart > rangeEnd) {
                    var swapped = rangeStart;
                    rangeStart = rangeEnd;
                    rangeEnd = swapped;
                }
                for (var j = rangeStart; j <= rangeEnd; j++) indices.push(j - 1);
            }
        }
        return indices;
    }

    /**
     * 0 始まりのインデックス配列から "1-3,5" 形式の文字列を生成する
     * @param {number[]} zeroBasedIndices - 0 始まりのインデックス配列
     * @returns {string} 1 始まりの範囲文字列
     */
    function buildArtboardRangeString(zeroBasedIndices) {
        if (!zeroBasedIndices || zeroBasedIndices.length === 0) return "";
        var sortedIndices = [];
        for (var i = 0; i < zeroBasedIndices.length; i++) sortedIndices.push(zeroBasedIndices[i]);
        sortedIndices.sort(function (a, b) { return a - b; });

        /* 連続する番号を "n-m" にまとめる / Collapse consecutive numbers into "n-m" */
        var parts = [];
        var rangeStart = sortedIndices[0];
        var rangeEnd = sortedIndices[0];
        for (var j = 1; j <= sortedIndices.length; j++) {
            if (j < sortedIndices.length && sortedIndices[j] === rangeEnd + 1) {
                rangeEnd = sortedIndices[j];
                continue;
            }
            parts.push(rangeStart === rangeEnd
                ? (rangeStart + 1) + ""
                : (rangeStart + 1) + "-" + (rangeEnd + 1));
            rangeStart = sortedIndices[j];
            rangeEnd = sortedIndices[j];
        }
        return parts.join(",");
    }

    /**
     * 対象アートボードのインデックス配列を返す
     * @param {number} artboardCount - アートボード総数
     * @param {string} targetType - "all" または "numbered"
     * @param {string} rangeText - "numbered" のときの範囲文字列
     * @returns {number[]} 対象インデックス配列
     */
    function getTargetArtboardIndices(artboardCount, targetType, rangeText) {
        if (targetType === "all") {
            var allIndices = [];
            for (var i = 0; i < artboardCount; i++) allIndices.push(i);
            return allIndices;
        }

        /* 存在しない番号と重複を落として、入力順のまま返す / Drop out-of-range and duplicate numbers, keeping the input order */
        var parsedIndices = parseArtboardRange(rangeText);
        var validIndices = [];
        for (var j = 0; j < parsedIndices.length; j++) {
            if (parsedIndices[j] < 0 || parsedIndices[j] >= artboardCount) continue;
            if (containsValue(validIndices, parsedIndices[j])) continue;
            validIndices.push(parsedIndices[j]);
        }
        return validIndices;
    }

    /**
     * リネーム対象のインデックスを決める（targetIndices があれば範囲文字列より優先）
     * @param {number} artboardCount - アートボード総数
     * @param {object} renameSettings - リネーム設定
     * @returns {number[]} 対象インデックス配列（連番はこの順に振られる）
     */
    function resolveTargetIndices(artboardCount, renameSettings) {
        if (renameSettings.targetIndices) return renameSettings.targetIndices;
        return getTargetArtboardIndices(artboardCount, renameSettings.artboardTarget, renameSettings.rangeText);
    }

    /**
     * 対象外アートボードの名前を衝突回避用の予約名として返す
     * @param {Artboards} artboards - ドキュメントのアートボード
     * @param {number[]} targetIndices - 対象インデックス配列
     * @returns {string[]} 予約名の配列
     */
    function getReservedArtboardNames(artboards, targetIndices) {
        var reservedNames = [];
        for (var i = 0; i < artboards.length; i++) {
            if (!containsValue(targetIndices, i)) reservedNames.push(artboards[i].name);
        }
        return reservedNames;
    }

    /**
     * 名前の配列をエントリーの並び順に合わせて並べ替える
     * @param {string[]} names - 並べ替え前の名前（元の位置順）
     * @param {object[]} artboardEntries - 並び替えリストのエントリー
     * @returns {string[]} エントリーの並び順に対応する名前
     */
    function reorderNamesByEntries(names, artboardEntries) {
        var reorderedNames = [];
        for (var i = 0; i < artboardEntries.length; i++) {
            reorderedNames.push(names[artboardEntries[i].originalIndex]);
        }
        return reorderedNames;
    }

    /**
     * 設定オブジェクトの浅いコピーを作る
     * @param {object} source - コピー元
     * @returns {object} コピーされたオブジェクト
     */
    function copySettings(source) {
        var copied = {};
        for (var settingsKey in source) {
            if (source.hasOwnProperty(settingsKey)) copied[settingsKey] = source[settingsKey];
        }
        return copied;
    }

    // =========================================
    // 参照テキストの収集 / Text source collection
    // =========================================

    /**
     * 名前でトップレベルレイヤーを検索する
     * @param {Document} doc - 対象ドキュメント
     * @param {string} layerName - レイヤー名
     * @returns {Layer|null} 見つかったレイヤー（なければ null）
     */
    function findLayerByName(doc, layerName) {
        for (var i = 0; i < doc.layers.length; i++) {
            if (doc.layers[i].name === layerName) return doc.layers[i];
        }
        return null;
    }

    /**
     * テキストフレームの可視バウンズの中心座標を取得する
     * @param {TextFrame} textFrame - 対象テキストフレーム
     * @returns {number[]} [x, y] の中心座標
     */
    function getTextCenter(textFrame) {
        var textBounds = textFrame.visibleBounds;
        return [(textBounds[0] + textBounds[2]) / 2, (textBounds[1] + textBounds[3]) / 2];
    }

    /**
     * テキストフレームの中心が指定アートボードの矩形内にあるかを判定する
     * @param {TextFrame} textFrame - 対象テキストフレーム
     * @param {number[]} artboardBounds - アートボードの矩形 [左,上,右,下]
     * @returns {boolean} 内側にあれば true
     */
    function isTextFrameOnArtboard(textFrame, artboardBounds) {
        var textCenter = getTextCenter(textFrame);
        return textCenter[0] >= artboardBounds[0] && textCenter[0] <= artboardBounds[2] &&
            textCenter[1] <= artboardBounds[1] && textCenter[1] >= artboardBounds[3];
    }

    /**
     * テキストフレームの内容を1行の参照文字列に整える
     * @param {TextFrame} textFrame - 対象テキストフレーム
     * @returns {string} 改行・タブを除いた文字列
     */
    function cleanTextContents(textFrame) {
        return textFrame.contents.replace(/[\r\n\t]/g, "");
    }

    /**
     * コンテナ内の TextFrame をグループ階層まで再帰的に収集する
     * container.pageItems はサブレイヤーやグループ内の項目も含む場合があるため、収集済みは取り込まない
     * @param {Layer|GroupItem} container - 走査対象のコンテナ
     * @param {TextFrame[]} collectedFrames - 収集先の配列
     * @returns {void}
     */
    function collectTextFramesInContainer(container, collectedFrames) {
        var containerItems = container.pageItems;
        for (var i = 0; i < containerItems.length; i++) {
            var pageItem = containerItems[i];
            if (pageItem.hidden || pageItem.locked) continue;

            if (pageItem.typename === "TextFrame") {
                if (!containsValue(collectedFrames, pageItem)) collectedFrames.push(pageItem);
            } else if (pageItem.typename === "GroupItem") {
                collectTextFramesInContainer(pageItem, collectedFrames);
            }
        }
    }

    /**
     * レイヤーとサブレイヤー配下の TextFrame を再帰的に収集する
     * @param {Layer} layer - 走査対象レイヤー
     * @returns {TextFrame[]} 収集したテキストフレーム
     */
    function getTextFramesInLayer(layer) {
        var collectedFrames = [];

        /* 非表示・ロックされたレイヤーは除外して再帰 / Skip hidden or locked layers while recursing */
        function collectInLayer(targetLayer) {
            if (!targetLayer.visible || targetLayer.locked) return;
            collectTextFramesInContainer(targetLayer, collectedFrames);
            for (var i = 0; i < targetLayer.layers.length; i++) {
                collectInLayer(targetLayer.layers[i]);
            }
        }

        collectInLayer(layer);
        return collectedFrames;
    }

    /**
     * コンテナ内を再帰的に走査し、指定アートボード上で最初に見つかった TextFrame を返す
     * @param {Layer|GroupItem} container - 走査対象のコンテナ
     * @param {number[]} artboardBounds - アートボードの矩形
     * @returns {TextFrame|null} 見つかったテキストフレーム（なければ null）
     */
    function findFrontmostTextFrameInContainer(container, artboardBounds) {
        var containerItems = container.pageItems;
        for (var i = 0; i < containerItems.length; i++) {
            var pageItem = containerItems[i];
            if (pageItem.hidden || pageItem.locked) continue;

            if (pageItem.typename === "TextFrame") {
                if (isTextFrameOnArtboard(pageItem, artboardBounds)) return pageItem;
            } else if (pageItem.typename === "GroupItem") {
                var nestedTextFrame = findFrontmostTextFrameInContainer(pageItem, artboardBounds);
                if (nestedTextFrame) return nestedTextFrame;
            }
        }
        return null;
    }

    /**
     * レイヤーとサブレイヤーを再帰的に走査し、指定アートボード上の最前面 TextFrame を返す
     * @param {Layer} layer - 走査対象レイヤー
     * @param {number[]} artboardBounds - アートボードの矩形
     * @returns {TextFrame|null} 見つかったテキストフレーム（なければ null）
     */
    function findFrontmostTextFrameInLayer(layer, artboardBounds) {
        if (!layer.visible || layer.locked) return null;

        var textFrame = findFrontmostTextFrameInContainer(layer, artboardBounds);
        if (textFrame) return textFrame;

        for (var i = 0; i < layer.layers.length; i++) {
            var nestedTextFrame = findFrontmostTextFrameInLayer(layer.layers[i], artboardBounds);
            if (nestedTextFrame) return nestedTextFrame;
        }
        return null;
    }

    /**
     * 各アートボードの最前面 TextFrame から参照文字列マップを作る
     * 判定順はレイヤー順・pageItems順に依存（Illustratorの厳密な描画Z順ではない）
     * @param {Document} doc - 対象ドキュメント
     * @returns {object} アートボードインデックスをキーにした文字列配列のマップ
     */
    function buildFrontmostTextMap(doc) {
        var textsByArtboardIndex = {};
        for (var artboardIndex = 0; artboardIndex < doc.artboards.length; artboardIndex++) {
            var artboardBounds = doc.artboards[artboardIndex].artboardRect;
            for (var layerIndex = 0; layerIndex < doc.layers.length; layerIndex++) {
                var textFrame = findFrontmostTextFrameInLayer(doc.layers[layerIndex], artboardBounds);
                if (!textFrame) continue;
                textsByArtboardIndex[artboardIndex] = [cleanTextContents(textFrame)];
                break;
            }
        }
        return textsByArtboardIndex;
    }

    /**
     * テキストフレームの中心を含むアートボードのうち、最も小さいものの位置を返す
     * アートボードが重なっている場合に、より内側のアートボードを優先する
     * @param {TextFrame} textFrame - 対象テキストフレーム
     * @param {Artboards} artboards - ドキュメントのアートボード
     * @returns {number} アートボードの位置（見つからなければ -1）
     */
    function findArtboardIndexForTextFrame(textFrame, artboards) {
        var matchedIndex = -1;
        var matchedArea = 0;
        for (var i = 0; i < artboards.length; i++) {
            var artboardBounds = artboards[i].artboardRect;
            if (!isTextFrameOnArtboard(textFrame, artboardBounds)) continue;
            var artboardArea = Math.abs(artboardBounds[2] - artboardBounds[0]) *
                Math.abs(artboardBounds[1] - artboardBounds[3]);
            if (matchedIndex === -1 || artboardArea < matchedArea) {
                matchedIndex = i;
                matchedArea = artboardArea;
            }
        }
        return matchedIndex;
    }

    /**
     * テキストフレームを、中心座標が含まれるアートボードに紐付ける
     * @param {TextFrame[]} textFrames - 対象テキストフレーム
     * @param {Artboards} artboards - ドキュメントのアートボード
     * @returns {object} アートボードインデックスをキーにした文字列配列のマップ
     */
    function mapTextFramesToArtboards(textFrames, artboards) {
        var textsByArtboardIndex = {};
        for (var i = 0; i < textFrames.length; i++) {
            var artboardIndex = findArtboardIndexForTextFrame(textFrames[i], artboards);
            if (artboardIndex === -1) continue;
            if (!textsByArtboardIndex[artboardIndex]) textsByArtboardIndex[artboardIndex] = [];
            textsByArtboardIndex[artboardIndex].push(cleanTextContents(textFrames[i]));
        }
        return textsByArtboardIndex;
    }

    /**
     * 参照テキストモードに応じて、アートボードごとの参照文字列マップを作る
     * @param {Document} doc - 対象ドキュメント
     * @param {object} renameSettings - リネーム設定
     * @returns {object} アートボードインデックスをキーにした文字列配列のマップ
     */
    function buildArtboardTextMap(doc, renameSettings) {
        var textSourceMode = renameSettings.textSourceMode;

        if (textSourceMode === "layer") {
            var sourceLayer = findLayerByName(doc, renameSettings.selectedLayerName);
            if (!sourceLayer || !sourceLayer.visible) return {};
            return mapTextFramesToArtboards(getTextFramesInLayer(sourceLayer), doc.artboards);
        }

        if (textSourceMode === "frontmost") {
            return buildFrontmostTextMap(doc);
        }

        var textsByArtboardIndex = {};
        for (var artboardIndex = 0; artboardIndex < doc.artboards.length; artboardIndex++) {
            if (textSourceMode === "originalName") {
                /* ダイアログを開いた時点の名前を参照（［更新］後の現在名ではない） / Use the name captured when the dialog opened */
                var originalNames = renameSettings.originalNames;
                textsByArtboardIndex[artboardIndex] = [(originalNames && originalNames[artboardIndex] != null)
                    ? originalNames[artboardIndex]
                    : doc.artboards[artboardIndex].name];
            } else if (textSourceMode === "custom") {
                textsByArtboardIndex[artboardIndex] = [renameSettings.customText];
            }
        }
        return textsByArtboardIndex;
    }

    // =========================================
    // 名前の生成 / Name generation
    // =========================================

    /**
     * 拡張子を除いたドキュメント名を取得する
     * @param {Document} doc - 対象ドキュメント
     * @returns {string} 拡張子なしのファイル名
     */
    function getDocumentBaseName(doc) {
        return doc.name.replace(/\.[^.]+$/, "");
    }

    /**
     * 今日の日付を yyyyMMdd 形式で取得する
     * @returns {string} yyyyMMdd 形式の日付文字列
     */
    function getTodayStamp() {
        var today = new Date();
        return today.getFullYear().toString() +
            ("0" + (today.getMonth() + 1)).slice(-2) +
            ("0" + today.getDate()).slice(-2);
    }

    /**
     * テンプレートの連番・#FN・#DT トークンを展開する
     * @param {string} template - 接頭辞または接尾辞のテンプレート
     * @param {number} sequenceNumber - 1 始まりの連番位置
     * @param {string} documentBaseName - #FN に展開するファイル名
     * @param {string} todayStamp - #DT に展開する日付文字列
     * @returns {string} 展開後の文字列
     */
    function expandTemplateTokens(template, sequenceNumber, documentBaseName, todayStamp) {
        var expanded = template;

        /* 連番置換を先に行う（#DT/#FN を先に展開すると日付やファイル名内の数字を誤って捕捉するため） / Replace the sequence first so digits from #DT/#FN are not picked up */
        var numberMatch = expanded.match(/\d+/);
        if (numberMatch) {
            var numberToken = numberMatch[0];
            var sequenceValue = parseInt(numberToken, 10) + sequenceNumber - 1;
            var isZeroPadded = numberToken.charAt(0) === "0" && numberToken.length > 1;
            expanded = expanded.replace(numberToken, isZeroPadded
                ? ("0000000000" + sequenceValue).slice(-numberToken.length)
                : sequenceValue.toString());
        }

        return expanded.replace(/#FN/g, documentBaseName).replace(/#DT/g, todayStamp);
    }

    /**
     * 重複している名前だけに "_1", "_2" を付けて最終名を返す（重複しない名前はそのまま）
     * 連番で結果がユニークになる場合は何も付かないため、連番の有無で分岐する必要はない
     * @param {string[]} baseNames - 計算済みのベース名
     * @param {string[]} reservedNames - すでに使われている名前
     * @returns {string[]} 重複を解消した名前
     */
    function resolveUniqueNames(baseNames, reservedNames) {
        var resolvedNames = [];

        /* ベース名ごとの出現回数（2以上なら重複） / Count each base name (2 or more means a collision) */
        var baseNameCount = {};
        for (var i = 0; i < baseNames.length; i++) {
            baseNameCount[baseNames[i]] = (baseNameCount[baseNames[i]] || 0) + 1;
        }

        var usedNames = [];
        for (var r = 0; r < reservedNames.length; r++) usedNames.push(reservedNames[r]);

        var suffixNumbers = {};
        for (var j = 0; j < baseNames.length; j++) {
            var baseName = baseNames[j];
            var resolvedName = baseName;
            /* 計画内で重複、予約名と衝突、または先に確定した名前と衝突する場合に連番を付ける
               / Add a suffix when the name repeats in the plan, hits a reserved name, or hits an already-resolved name */
            if (baseNameCount[baseName] > 1 || containsValue(usedNames, baseName)) {
                do {
                    suffixNumbers[baseName] = (suffixNumbers[baseName] || 0) + 1;
                    resolvedName = baseName + "_" + suffixNumbers[baseName];
                } while (containsValue(usedNames, resolvedName));
            }
            usedNames.push(resolvedName);
            resolvedNames.push(resolvedName);
        }
        return resolvedNames;
    }

    /**
     * 対象アートボードの新しい名前を計算する（canvas は変更しない）
     * @param {Artboards} artboards - ドキュメントのアートボード
     * @param {object} artboardTextMap - アートボードごとの参照文字列マップ
     * @param {object} renameSettings - リネーム設定
     * @param {number[]} targetIndices - 対象インデックス配列
     * @param {string} documentBaseName - #FN に展開するファイル名
     * @returns {string[]} 全アートボード分の名前（対象外は現在の名前）
     */
    function planArtboardNames(artboards, artboardTextMap, renameSettings, targetIndices, documentBaseName) {
        var plannedNames = [];
        for (var i = 0; i < artboards.length; i++) plannedNames.push(artboards[i].name);

        var prefixTemplate = renameSettings.prefix;
        var suffixTemplate = renameSettings.suffix;
        var todayStamp = getTodayStamp();
        var reservedNames = getReservedArtboardNames(artboards, targetIndices);

        /* 1パス目: 各ターゲットのベース名を計算（結果が空のものは予約名に逃がす） / Pass 1: build base names, keeping empty results as reserved */
        var baseNames = [];
        var baseNameIndices = [];
        var sequenceNumber = 1;
        /* targetIndices の並び順で連番を振る（並び替え中は表示順に合わせるため） / Number them in the given target order */
        for (var t = 0; t < targetIndices.length; t++) {
            var artboardIndex = targetIndices[t];
            if (artboardIndex < 0 || artboardIndex >= artboards.length) continue;
            if (containsValue(baseNameIndices, artboardIndex)) continue;
            var expandedPrefix = expandTemplateTokens(prefixTemplate, sequenceNumber, documentBaseName, todayStamp);
            var expandedSuffix = expandTemplateTokens(suffixTemplate, sequenceNumber, documentBaseName, todayStamp);
            var referenceText = artboardTextMap[artboardIndex] ? artboardTextMap[artboardIndex].join(" ") : "";
            if (!expandedPrefix && !expandedSuffix && !referenceText) {
                reservedNames.push(plannedNames[artboardIndex]);
                continue;
            }
            baseNames.push(expandedPrefix + referenceText + expandedSuffix);
            baseNameIndices.push(artboardIndex);
            sequenceNumber++;
        }

        /* 2パス目: 重複しているベース名だけ "_1", "_2" を付加 / Pass 2: append "_1", "_2" to colliding base names */
        var resolvedNames = resolveUniqueNames(baseNames, reservedNames);
        for (var k = 0; k < baseNameIndices.length; k++) {
            plannedNames[baseNameIndices[k]] = resolvedNames[k];
        }
        return plannedNames;
    }

    /**
     * リネーム設定が実行可能かを検証する
     * @param {Document} doc - 対象ドキュメント
     * @param {object} renameSettings - リネーム設定
     * @returns {string|null} 問題があれば LABELS.alert のキー、なければ null
     */
    function validateRenameSettings(doc, renameSettings) {
        if (renameSettings.textSourceMode === "custom" &&
            renameSettings.customText === "" && renameSettings.prefix === "" && renameSettings.suffix === "") {
            return "needSettings";
        }
        if (renameSettings.textSourceMode === "layer") {
            var sourceLayer = findLayerByName(doc, renameSettings.selectedLayerName);
            if (!sourceLayer || !sourceLayer.visible) return "noLayer";
        }
        return null;
    }

    /**
     * canvas を変更せずに「いま［更新］したらこうなる」名前を計算する
     * @param {Document} doc - 対象ドキュメント
     * @param {object} renameSettings - リネーム設定
     * @returns {string[]} 全アートボード分のプレビュー名
     */
    function computePreviewNames(doc, renameSettings) {
        var currentNames = [];
        for (var i = 0; i < doc.artboards.length; i++) currentNames.push(doc.artboards[i].name);
        if (validateRenameSettings(doc, renameSettings)) return currentNames;

        return planArtboardNames(
            doc.artboards, buildArtboardTextMap(doc, renameSettings), renameSettings,
            resolveTargetIndices(doc.artboards.length, renameSettings), getDocumentBaseName(doc));
    }

    /**
     * 設定に従いアートボードをリネームして canvas を更新する
     * @param {Document} doc - 対象ドキュメント
     * @param {object} renameSettings - リネーム設定
     * @param {object} [renameOptions] - 実行オプション（silent: true で警告を出さない）
     * @returns {boolean} リネームを実行したら true
     */
    function executeRename(doc, renameSettings, renameOptions) {
        var alertKey = validateRenameSettings(doc, renameSettings);
        if (alertKey) {
            if (!renameOptions || !renameOptions.silent) alert(getLabel("alert", alertKey));
            return false;
        }

        var newNames = computePreviewNames(doc, renameSettings);
        for (var i = 0; i < doc.artboards.length; i++) {
            if (doc.artboards[i].name !== newNames[i]) doc.artboards[i].name = newNames[i];
        }
        return true;
    }

    // =========================================
    // 並び替えと適用 / Reorder and apply
    // =========================================

    /**
     * 表示順が canvas の並びと変わっているかを判定する
     * @param {object[]} artboardEntries - 並び替えリストのエントリー
     * @returns {boolean} 並び替えがあれば true
     */
    function hasReorderedEntries(artboardEntries) {
        for (var i = 0; i < artboardEntries.length; i++) {
            if (artboardEntries[i].originalIndex !== i) return true;
        }
        return false;
    }

    /**
     * エントリーの並び順をアートボードの矩形と名前に反映する
     * @param {Artboards} artboards - ドキュメントのアートボード
     * @param {object[]} artboardEntries - 並び替えリストのエントリー
     * @returns {void}
     */
    function applyEntryOrderToArtboards(artboards, artboardEntries) {
        /* 並べ替え前の canvas 名を位置で記録（［更新］済みの名前を保持） / Snapshot current canvas names before reordering */
        var namesByOriginalIndex = [];
        for (var i = 0; i < artboards.length; i++) namesByOriginalIndex.push(artboards[i].name);

        /* 一時名を挟んで同名衝突を避ける / Use temporary names to avoid collisions while reordering */
        for (var t = 0; t < artboards.length; t++) {
            artboards[t].name = "__tmp_ab_" + t + "__";
        }
        for (var newPosition = 0; newPosition < artboardEntries.length; newPosition++) {
            artboards[newPosition].artboardRect = artboardEntries[newPosition].rect;
            artboards[newPosition].name = namesByOriginalIndex[artboardEntries[newPosition].originalIndex];
        }
    }

    /**
     * 並び替え後の位置を基準にリネーム設定を作り直す
     * @param {object} renameSettings - 元のリネーム設定
     * @param {object[]} artboardEntries - 並び替えリストのエントリー
     * @returns {object} 並び替え後の位置に合わせたリネーム設定
     */
    function buildReorderedSettings(renameSettings, artboardEntries) {
        var reorderedSettings = copySettings(renameSettings);

        /* 並び替え後の位置がそのまま対象インデックスになる / Reordered positions are the new target indices */
        var targetPositions = [];
        for (var i = 0; i < artboardEntries.length; i++) {
            if (artboardEntries[i].checked) targetPositions.push(i);
        }
        reorderedSettings.targetIndices = targetPositions;

        /* originalName モード用に、元の名前の参照順も並び替え後へ補正 / Reorder the original-name references as well */
        if (reorderedSettings.originalNames && reorderedSettings.originalNames.length) {
            reorderedSettings.originalNames = reorderNamesByEntries(reorderedSettings.originalNames, artboardEntries);
        }
        return reorderedSettings;
    }

    /**
     * 手動で上書きされた名前を並び替え後の位置に適用する
     * @param {Artboards} artboards - ドキュメントのアートボード
     * @param {object[]} artboardEntries - 並び替えリストのエントリー
     * @returns {void}
     */
    function applyManualNameOverrides(artboards, artboardEntries) {
        for (var position = 0; position < artboardEntries.length; position++) {
            if (position >= artboards.length) break;
            if (artboardEntries[position].userEdited && artboardEntries[position].checked) {
                artboards[position].name = artboardEntries[position].newName;
            }
        }
    }

    /**
     * リネーム・並び替え・手動上書きをまとめてアートボードに適用する
     * 並び替えがある場合は、並べ替えてから新しい位置を基準に一度だけリネームする
     * @param {Document} doc - 対象ドキュメント
     * @param {object} renameSettings - リネーム設定（artboardEntries を含む）
     * @param {object} [renameOptions] - 実行オプション（silent: true で警告を出さない）
     * @returns {boolean} 適用したら true
     */
    function applyRenameAndReorder(doc, renameSettings, renameOptions) {
        var alertKey = validateRenameSettings(doc, renameSettings);
        if (alertKey) {
            if (!renameOptions || !renameOptions.silent) alert(getLabel("alert", alertKey));
            return false;
        }

        var artboardEntries = renameSettings.artboardEntries;
        if (artboardEntries && hasReorderedEntries(artboardEntries)) {
            /* 並べ替えてから、新しい位置を基準に一度だけリネームする / Reorder first, then rename once against the new positions */
            applyEntryOrderToArtboards(doc.artboards, artboardEntries);
            executeRename(doc, buildReorderedSettings(renameSettings, artboardEntries), { silent: true });
        } else {
            executeRename(doc, renameSettings, renameOptions);
        }

        if (artboardEntries) applyManualNameOverrides(doc.artboards, artboardEntries);
        return true;
    }

    /**
     * キャンセル時にアートボードの名前と並び（矩形）を元に戻す
     * @param {Document} doc - 対象ドキュメント
     * @param {string[]} originalNames - ダイアログ表示前の名前
     * @param {number[][]} originalRects - ダイアログ表示前の矩形
     * @returns {void}
     */
    function restoreOriginalArtboards(doc, originalNames, originalRects) {
        for (var i = 0; i < doc.artboards.length; i++) {
            if (originalRects[i]) doc.artboards[i].artboardRect = originalRects[i];
            doc.artboards[i].name = originalNames[i];
        }
    }

    // =========================================
    // 並び替えリスト / Reorder list
    // =========================================

    /**
     * 並び替え／リネームリストを構築し、操作用のAPIを返す
     * @param {Panel} hostPanel - リストを追加するパネル
     * @param {Document} doc - 対象ドキュメント
     * @param {object} listCallbacks - onCheckedChange / onLayoutChange コールバック
     * @returns {object} リスト操作API
     */
    function createReorderList(hostPanel, doc, listCallbacks) {
        var artboardEntries = [];
        for (var artboardIndex = 0; artboardIndex < doc.artboards.length; artboardIndex++) {
            artboardEntries.push({
                originalIndex: artboardIndex,
                name: doc.artboards[artboardIndex].name,
                newName: doc.artboards[artboardIndex].name,
                rect: doc.artboards[artboardIndex].artboardRect,
                checked: false,
                userEdited: false
            });
        }

        var entryRowControls = [];
        /* 表示中の先頭行（スクロール位置） / Index of the first visible row */
        var scrollOffset = 0;

        var listRow = hostPanel.add("group");
        listRow.orientation = "row";
        listRow.alignChildren = ["fill", "fill"];
        listRow.spacing = 4;

        var reorderRowsHost = listRow.add("group");
        reorderRowsHost.orientation = "column";
        reorderRowsHost.alignChildren = ["fill", "top"];
        reorderRowsHost.spacing = 4;

        /* 行数が表示上限を超えるときだけスクロールバーを作る / Add a scrollbar only when the rows overflow */
        var listScrollbar = null;
        if (artboardEntries.length > VISIBLE_ROW_COUNT) {
            listScrollbar = listRow.add("scrollbar");
            listScrollbar.alignment = ["right", "fill"];
            listScrollbar.preferredSize.width = SCROLLBAR_WIDTH;
            listScrollbar.minimumSize.height = (VISIBLE_ROW_COUNT + 1) * (ENTRY_ROW_HEIGHT + 4);
            listScrollbar.minvalue = 0;
            listScrollbar.maxvalue = artboardEntries.length - VISIBLE_ROW_COUNT;
            listScrollbar.value = 0;
            listScrollbar.stepdelta = 1;
            listScrollbar.jumpdelta = VISIBLE_ROW_COUNT;
        }

        var moveButtonsRow = hostPanel.add("group");
        setupRow(moveButtonsRow, "center", 4);
        var moveToTopButton = moveButtonsRow.add("button", undefined, getLabel("button", "moveToTop"));
        var moveUpButton = moveButtonsRow.add("button", undefined, getLabel("button", "moveUp"));
        var moveDownButton = moveButtonsRow.add("button", undefined, getLabel("button", "moveDown"));
        var moveToBottomButton = moveButtonsRow.add("button", undefined, getLabel("button", "moveToBottom"));
        moveToTopButton.preferredSize = [MOVE_BUTTON_WIDTH + 4, MOVE_BUTTON_HEIGHT];
        moveUpButton.preferredSize = [MOVE_BUTTON_WIDTH, MOVE_BUTTON_HEIGHT];
        moveDownButton.preferredSize = [MOVE_BUTTON_WIDTH, MOVE_BUTTON_HEIGHT];
        moveToBottomButton.preferredSize = [MOVE_BUTTON_WIDTH + 4, MOVE_BUTTON_HEIGHT];

        /**
         * 表示中の行の値をエントリーへ書き戻す
         * @returns {void}
         */
        function syncEditingValues() {
            for (var i = 0; i < entryRowControls.length; i++) {
                var entry = entryRowControls[i].entry;
                entry.checked = entryRowControls[i].checkbox.value;
                if (entry.checked) entry.newName = entryRowControls[i].newNameField.text;
            }
        }

        /**
         * チェックを外した行の名前をプレビュー対象外へ戻す
         * @param {object} entry - 対象エントリー
         * @returns {void}
         */
        function resetEntryName(entry) {
            entry.newName = entry.name;
            entry.userEdited = false;
        }

        /**
         * エントリーをチェック済み・未チェックに分けて並べ直す
         * @param {boolean} checkedFirst - チェック済みを先頭に集めるなら true
         * @returns {void}
         */
        function partitionEntriesByChecked(checkedFirst) {
            var checkedEntries = [];
            var uncheckedEntries = [];
            for (var i = 0; i < artboardEntries.length; i++) {
                if (artboardEntries[i].checked) checkedEntries.push(artboardEntries[i]);
                else uncheckedEntries.push(artboardEntries[i]);
            }
            artboardEntries = checkedFirst
                ? checkedEntries.concat(uncheckedEntries)
                : uncheckedEntries.concat(checkedEntries);
        }

        /**
         * 指定方向へ移動できるチェック行があるかを判定する
         * @param {number} step - -1 で上方向、1 で下方向
         * @returns {boolean} 移動できる行があれば true
         */
        function hasMovableEntry(step) {
            for (var i = 0; i < artboardEntries.length; i++) {
                var neighborIndex = i + step;
                if (neighborIndex < 0 || neighborIndex >= artboardEntries.length) continue;
                if (artboardEntries[i].checked && !artboardEntries[neighborIndex].checked) return true;
            }
            return false;
        }

        /**
         * チェック行を指定方向へ1つ移動する
         * @param {number} step - -1 で上方向、1 で下方向
         * @returns {void}
         */
        function moveCheckedEntries(step) {
            syncEditingValues();
            /* 上方向は先頭から、下方向は末尾から走査する / Scan from the top when moving up, from the bottom when moving down */
            for (var i = 0; i < artboardEntries.length; i++) {
                var entryIndex = (step < 0) ? i : artboardEntries.length - 1 - i;
                var neighborIndex = entryIndex + step;
                if (neighborIndex < 0 || neighborIndex >= artboardEntries.length) continue;
                if (!artboardEntries[entryIndex].checked || artboardEntries[neighborIndex].checked) continue;
                var movedEntry = artboardEntries[entryIndex];
                artboardEntries[entryIndex] = artboardEntries[neighborIndex];
                artboardEntries[neighborIndex] = movedEntry;
            }
            scrollToFirstChecked();
            refreshRows();
        }

        /**
         * 並び替えボタンの有効・無効を更新する
         * @returns {void}
         */
        function updateMoveButtonsState() {
            var canMoveUp = hasMovableEntry(-1);
            var canMoveDown = hasMovableEntry(1);
            moveToTopButton.enabled = canMoveUp;
            moveUpButton.enabled = canMoveUp;
            moveDownButton.enabled = canMoveDown;
            moveToBottomButton.enabled = canMoveDown;
        }

        /**
         * エントリーのチェック状態を更新する（外したら名前をプレビュー対象へ戻す）
         * @param {object} entry - 対象エントリー
         * @param {boolean} isChecked - チェック状態
         * @returns {void}
         */
        function setEntryChecked(entry, isChecked) {
            entry.checked = isChecked;
            if (!isChecked) resetEntryName(entry);
        }

        /**
         * エントリーの現在の状態を表示中の行へ描画する
         * @returns {void}
         */
        function paintVisibleRows() {
            for (var i = 0; i < entryRowControls.length; i++) {
                var rowControls = entryRowControls[i];
                var entry = rowControls.entry;
                var currentName = doc.artboards[entry.originalIndex].name;

                /* 元の名前 列：現在の canvas 名（［更新］後は確定後の名前） / Current canvas name */
                rowControls.currentNameLabel.text = currentName;
                rowControls.currentNameLabel.helpTip = currentName;

                rowControls.checkbox.value = entry.checked;
                rowControls.newNameField.enabled = entry.checked;
                rowControls.newNameField.text = entry.newName;
            }
        }

        /**
         * 見出し行を追加する
         * @returns {void}
         */
        function buildHeaderRow() {
            var headerRow = reorderRowsHost.add("group");
            setupRow(headerRow, "left", ROW_SPACING);
            var headerWidths = [ORDER_COLUMN_WIDTH, TARGET_COLUMN_WIDTH, CURRENT_NAME_WIDTH, ARROW_COLUMN_WIDTH, NEW_NAME_WIDTH];
            var headerTexts = [
                getLabel("columnHeader", "order"),
                getLabel("columnHeader", "selection"),
                getLabel("columnHeader", "currentName"),
                "→",
                getLabel("columnHeader", "newName")
            ];
            for (var i = 0; i < headerTexts.length; i++) {
                headerRow.add("statictext", undefined, headerTexts[i]).preferredSize.width = headerWidths[i];
            }
        }

        /**
         * アートボード1件分の行を追加する
         * @param {number} entryIndex - エントリーの位置（0 始まり、スクロール位置を含む絶対値）
         * @returns {void}
         */
        function buildEntryRow(entryIndex) {
            var entry = artboardEntries[entryIndex];
            var entryRow = reorderRowsHost.add("group");
            setupRow(entryRow, "left", ROW_SPACING);

            entryRow.add("statictext", undefined, (entryIndex + 1) + "").preferredSize.width = ORDER_COLUMN_WIDTH;

            var rowCheckbox = entryRow.add("checkbox", undefined, "");
            rowCheckbox.preferredSize.width = TARGET_COLUMN_WIDTH;

            var currentNameLabel = entryRow.add("statictext", undefined, entry.name);
            currentNameLabel.preferredSize.width = CURRENT_NAME_WIDTH;

            entryRow.add("statictext", undefined, "→").preferredSize.width = ARROW_COLUMN_WIDTH;

            var newNameField = entryRow.add("edittext", undefined, entry.newName);
            newNameField.preferredSize.width = NEW_NAME_WIDTH;

            var rowControls = {
                entry: entry,
                checkbox: rowCheckbox,
                currentNameLabel: currentNameLabel,
                newNameField: newNameField
            };

            rowCheckbox.onClick = function () {
                var isChecked = rowCheckbox.value;
                /* Option+クリックで全行を一括切り替え / Option-click toggles every row */
                var isOptionKeyHeld = ScriptUI.environment && ScriptUI.environment.keyboardState &&
                    ScriptUI.environment.keyboardState.altKey;
                if (isOptionKeyHeld) {
                    for (var i = 0; i < artboardEntries.length; i++) setEntryChecked(artboardEntries[i], isChecked);
                } else {
                    setEntryChecked(entry, isChecked);
                }
                paintVisibleRows();
                if (listCallbacks.onCheckedChange) listCallbacks.onCheckedChange();
                updateMoveButtonsState();
            };

            newNameField.onChange = function () {
                entry.newName = newNameField.text;
                entry.userEdited = true;
            };

            entryRowControls.push(rowControls);
        }

        /**
         * 表示位置に合わせて並び替えリストの行を作り直す
         * @returns {void}
         */
        function refreshRows() {
            while (reorderRowsHost.children.length > 0) {
                reorderRowsHost.remove(reorderRowsHost.children[0]);
            }
            entryRowControls = [];

            if (listScrollbar) listScrollbar.value = scrollOffset;

            buildHeaderRow();
            var lastRowIndex = Math.min(artboardEntries.length, scrollOffset + VISIBLE_ROW_COUNT);
            for (var entryIndex = scrollOffset; entryIndex < lastRowIndex; entryIndex++) {
                buildEntryRow(entryIndex);
            }

            paintVisibleRows();
            updateMoveButtonsState();
            reorderRowsHost.layout.layout(true);
            if (listCallbacks.onLayoutChange) listCallbacks.onLayoutChange();
        }

        /**
         * 最初のチェック行が表示範囲に入るようスクロール位置を調整する
         * @returns {void}
         */
        function scrollToFirstChecked() {
            for (var i = 0; i < artboardEntries.length; i++) {
                if (!artboardEntries[i].checked) continue;
                if (i < scrollOffset) scrollOffset = i;
                else if (i >= scrollOffset + VISIBLE_ROW_COUNT) scrollOffset = i - VISIBLE_ROW_COUNT + 1;
                break;
            }
        }

        /**
         * スクロール位置を変更して行を作り直す
         * @param {number} newOffset - 新しい先頭行の位置
         * @returns {void}
         */
        function setScrollOffset(newOffset) {
            var maxScrollOffset = Math.max(0, artboardEntries.length - VISIBLE_ROW_COUNT);
            var clampedOffset = Math.max(0, Math.min(newOffset, maxScrollOffset));
            if (clampedOffset === scrollOffset) return;
            /* 行を作り直す前に、表示中の編集内容をエントリーへ退避 / Keep pending edits before rebuilding rows */
            syncEditingValues();
            scrollOffset = clampedOffset;
            refreshRows();
        }

        if (listScrollbar) {
            listScrollbar.onChanging = function () { setScrollOffset(Math.round(listScrollbar.value)); };
            listScrollbar.onChange = function () { setScrollOffset(Math.round(listScrollbar.value)); };
        }
        moveToTopButton.onClick = function () {
            syncEditingValues();
            partitionEntriesByChecked(true);
            scrollToFirstChecked();
            refreshRows();
        };
        moveToBottomButton.onClick = function () {
            syncEditingValues();
            partitionEntriesByChecked(false);
            scrollToFirstChecked();
            refreshRows();
        };
        moveUpButton.onClick = function () { moveCheckedEntries(-1); };
        moveDownButton.onClick = function () { moveCheckedEntries(1); };

        return {
            getEntries: function () { return artboardEntries; },
            syncEditingValues: syncEditingValues,
            refreshRows: refreshRows,

            /**
             * 指定したエントリーが表示範囲に入るまでスクロールする
             * @param {number} entryIndex - 表示したいエントリーの位置
             * @returns {void}
             */
            revealEntry: function (entryIndex) {
                setScrollOffset(entryIndex - Math.floor(VISIBLE_ROW_COUNT / 2));
            },

            /**
             * 確定後の canvas に合わせてエントリーの位置・名前・矩形を取り直す
             * 並び替えを適用した直後に呼び、originalIndex を新しい位置へ振り直す
             * @returns {void}
             */
            syncEntriesToCanvas: function () {
                for (var i = 0; i < artboardEntries.length && i < doc.artboards.length; i++) {
                    artboardEntries[i].originalIndex = i;
                    artboardEntries[i].name = doc.artboards[i].name;
                    artboardEntries[i].rect = doc.artboards[i].artboardRect;
                }
            },

            /**
             * プレビュー名を全エントリーに反映する（手動編集行は据え置き）
             * @param {string[]} previewNames - アートボード位置順のプレビュー名
             * @returns {void}
             */
            applyPreviewNames: function (previewNames) {
                for (var i = 0; i < artboardEntries.length; i++) {
                    var entry = artboardEntries[i];
                    if (entry.userEdited) continue;
                    entry.newName = (previewNames && previewNames[entry.originalIndex] != null)
                        ? previewNames[entry.originalIndex]
                        : doc.artboards[entry.originalIndex].name;
                }
                paintVisibleRows();
            },

            /**
             * 対象インデックスに合わせてチェック状態を一括設定する
             * @param {number[]} targetIndices - 対象行の表示位置（0 始まり）
             * @returns {void}
             */
            applyCheckedIndices: function (targetIndices) {
                /* 対象は「順」列と同じ表示位置で指定する / Target indices are display positions, matching the "#" column */
                for (var i = 0; i < artboardEntries.length; i++) {
                    setEntryChecked(artboardEntries[i], containsValue(targetIndices, i));
                }
                paintVisibleRows();
                updateMoveButtonsState();
            }
        };
    }

    // =========================================
    // ダイアログ構築 / Dialog construction
    // =========================================

    /**
     * 接頭辞・接尾辞用のトークン挿入ボタンを追加する（最終行末にクリアボタン x を配置）
     * @param {Panel} affixPanel - ボタンを追加するパネル
     * @param {EditText} targetInput - 挿入先の入力欄
     * @returns {void}
     */
    function addTokenButtons(affixPanel, targetInput) {
        var tokenButtonRows = [
            [
                { label: "1", value: "1", width: 22 },
                { label: "01", value: "01" },
                { label: "-", value: "-", width: 22 },
                { label: "_", value: "_", width: 22 }
            ],
            [
                { label: "#FN", value: "#FN", width: 40 },
                { label: "#DT", value: "#DT", width: 40 }
            ]
        ];

        /**
         * 入力欄の値を更新し、プレビューを更新させる
         * @param {string} newText - 設定する文字列
         * @returns {void}
         */
        function updateTargetInput(newText) {
            targetInput.text = newText;
            targetInput.notify("onChange");
        }

        for (var rowIndex = 0; rowIndex < tokenButtonRows.length; rowIndex++) {
            var tokenButtonRow = affixPanel.add("group");
            setupRow(tokenButtonRow, "left", 4);
            var tokenSpecs = tokenButtonRows[rowIndex];
            for (var tokenIndex = 0; tokenIndex < tokenSpecs.length; tokenIndex++) {
                (function (tokenSpec) {
                    var tokenButton = tokenButtonRow.add("button", undefined, tokenSpec.label);
                    tokenButton.preferredSize = [tokenSpec.width || TOKEN_BUTTON_WIDTH, TOKEN_BUTTON_HEIGHT];
                    tokenButton.onClick = function () { updateTargetInput(targetInput.text + tokenSpec.value); };
                })(tokenSpecs[tokenIndex]);
            }
            /* 最終行末にのみ、対象フィールドをクリアするボタンを配置 / Add the clear button on the last row only */
            if (rowIndex === tokenButtonRows.length - 1) {
                var clearFieldButton = tokenButtonRow.add("button", undefined, "x");
                clearFieldButton.preferredSize = [22, TOKEN_BUTTON_HEIGHT];
                clearFieldButton.onClick = function () { updateTargetInput(""); };
            }
        }
    }

    /**
     * 接頭辞・接尾辞パネルを構築する
     * @param {Panel} parentPanel - 追加先パネル
     * @param {string} panelLabel - パネル見出し
     * @param {number} inputCharacters - 入力欄の文字数
     * @returns {EditText} 生成した入力欄
     */
    function addAffixPanel(parentPanel, panelLabel, inputCharacters) {
        var affixPanel = parentPanel.add("panel", undefined, panelLabel);
        setupPanel(affixPanel, ROW_SPACING);
        var affixInput = affixPanel.add("edittext", undefined, "");
        affixInput.characters = inputCharacters;
        addTokenButtons(affixPanel, affixInput);
        return affixInput;
    }

    /**
     * 参照テキストパネルを構築する
     * @param {Panel} parentPanel - 追加先パネル
     * @param {Document} doc - 対象ドキュメント
     * @returns {object} 生成したラジオボタンと入力欄
     */
    function addTextSourcePanel(parentPanel, doc) {
        var textSourcePanel = parentPanel.add("panel", undefined, getLabel("panel", "textSource"));
        setupPanel(textSourcePanel, ROW_SPACING);

        var originalNameRadio = textSourcePanel.add("radiobutton", undefined, getLabel("radio", "originalName"));

        var customTextRow = textSourcePanel.add("group");
        setupRow(customTextRow, "left", ROW_SPACING);
        var customTextRadio = customTextRow.add("radiobutton", undefined, getLabel("radio", "customText"));
        var customTextInput = customTextRow.add("edittext", undefined, "");
        customTextInput.characters = 12;
        customTextInput.enabled = false;

        var frontmostRadio = textSourcePanel.add("radiobutton", undefined, getLabel("radio", "frontmost"));

        /* 「レイヤー」ラジオとレイヤードロップダウンを横並びに / Keep the layer radio and dropdown on one row */
        var layerSelectRow = textSourcePanel.add("group");
        setupRow(layerSelectRow, "left", ROW_SPACING);
        var layerRadio = layerSelectRow.add("radiobutton", undefined, getLabel("radio", "layer"));
        var layerDropdown = layerSelectRow.add("dropdownlist", undefined, []);
        layerDropdown.minimumSize.width = 100;
        for (var i = 0; i < doc.layers.length; i++) {
            if (doc.layers[i].visible) layerDropdown.add("item", doc.layers[i].name);
        }
        if (layerDropdown.items.length > 0) layerDropdown.selection = 0;
        layerDropdown.enabled = false;

        /* 全ラジオを追加した後に初期値を設定 / Set initial values after every radio exists */
        originalNameRadio.value = false;
        customTextRadio.value = false;
        frontmostRadio.value = true;
        layerRadio.value = false;

        return {
            originalNameRadio: originalNameRadio,
            customTextRadio: customTextRadio,
            customTextInput: customTextInput,
            frontmostRadio: frontmostRadio,
            layerRadio: layerRadio,
            layerDropdown: layerDropdown
        };
    }

    /**
     * 対象アートボードパネルを構築する
     * @param {Group} parentGroup - 追加先グループ
     * @returns {object} 生成したラジオボタンと入力欄
     */
    function addTargetArtboardPanel(parentGroup) {
        var targetPanel = parentGroup.add("panel", undefined, getLabel("panel", "targetBoards"));
        setupPanel(targetPanel, ROW_SPACING);

        var targetRow = targetPanel.add("group");
        setupRow(targetRow, "left", ROW_SPACING);
        var allArtboardsRadio = targetRow.add("radiobutton", undefined, getLabel("radio", "allBoards"));
        var rangeArtboardsRadio = targetRow.add("radiobutton", undefined, getLabel("radio", "rangeBoards"));
        var rangeInput = targetRow.add("edittext", undefined, "");
        rangeInput.characters = 10;
        rangeInput.enabled = false;
        allArtboardsRadio.value = true;
        rangeArtboardsRadio.value = false;

        return {
            allArtboardsRadio: allArtboardsRadio,
            rangeArtboardsRadio: rangeArtboardsRadio,
            rangeInput: rangeInput
        };
    }

    /**
     * ダイアログ下部のボタン行を構築する
     * @param {Window} renameDialog - 対象ダイアログ
     * @returns {object} 更新・キャンセル・OK ボタン
     */
    function addDialogButtonRow(renameDialog) {
        var buttonRow = renameDialog.add("group");
        setupRow(buttonRow, "fill", ROW_SPACING);

        var refreshButton = buttonRow.add("button", undefined, getLabel("button", "refresh"));
        refreshButton.alignment = ["left", "center"];

        var buttonSpacer = buttonRow.add("group");
        buttonSpacer.alignment = ["fill", "fill"];
        buttonSpacer.minimumSize.width = 0;

        var cancelButton = buttonRow.add("button", undefined, getLabel("button", "cancel"), { name: "cancel" });
        cancelButton.alignment = ["right", "center"];

        var okButton = buttonRow.add("button", undefined, getLabel("button", "ok"), { name: "ok" });
        okButton.alignment = ["right", "center"];

        return { refreshButton: refreshButton, cancelButton: cancelButton, okButton: okButton };
    }

    /**
     * リネームダイアログのUIを構築する
     * @param {Document} doc - 対象ドキュメント
     * @returns {object} ダイアログとコントロール群
     */
    function createRenameDialog(doc) {
        var renameDialog = new Window("dialog", getLabel("dialog", "title") + " " + SCRIPT_VERSION);
        setupWindow(renameDialog);

        /* コンテンツ行（左：リネーム条件、右：対象アートボード＋並び替え／リネーム） / Content row: conditions on the left, targets and reorder list on the right */
        var contentRow = renameDialog.add("group");
        setupRow(contentRow, "fill", COLUMN_SPACING);
        contentRow.alignChildren = ["fill", "top"];

        var conditionColumn = contentRow.add("group");
        conditionColumn.orientation = "column";
        conditionColumn.alignChildren = "fill";
        conditionColumn.spacing = COLUMN_SPACING;

        var renamePanel = conditionColumn.add("panel", undefined, getLabel("panel", "rename"));
        setupPanel(renamePanel, ROW_SPACING);

        var prefixInput = addAffixPanel(renamePanel, getLabel("panel", "prefix"), 14);
        prefixInput.active = true;
        var textSourceControls = addTextSourcePanel(renamePanel, doc);
        var suffixInput = addAffixPanel(renamePanel, getLabel("panel", "suffix"), 16);

        var targetColumn = contentRow.add("group");
        targetColumn.orientation = "column";
        targetColumn.alignChildren = "fill";
        targetColumn.spacing = COLUMN_SPACING;

        var targetControls = addTargetArtboardPanel(targetColumn);

        var reorderPanel = targetColumn.add("panel", undefined, getLabel("panel", "reorder"));
        setupPanel(reorderPanel, ROW_SPACING);

        /* bindDialogEvents から注入されるプレビュー更新コールバック / Preview updater injected by bindDialogEvents */
        var requestPreviewUpdate = null;

        /**
         * チェック状態から対象アートボード設定（ラジオ＋指定範囲）へ逆同期する
         * @returns {void}
         */
        function syncTargetFromCheckboxes() {
            var artboardEntries = reorderList.getEntries();
            var checkedPositions = [];
            for (var i = 0; i < artboardEntries.length; i++) {
                if (artboardEntries[i].checked) checkedPositions.push(i);
            }

            var isAllChecked = checkedPositions.length === artboardEntries.length;
            targetControls.allArtboardsRadio.value = isAllChecked;
            targetControls.rangeArtboardsRadio.value = !isAllChecked;
            targetControls.rangeInput.enabled = !isAllChecked;
            /* 表示位置を基準に指定範囲へ反映（全解除なら空にする） / Reflect display positions in the range field, clearing it when nothing is checked */
            if (!isAllChecked) {
                targetControls.rangeInput.text = buildArtboardRangeString(checkedPositions);
            }
        }

        var reorderList = createReorderList(reorderPanel, doc, {
            onCheckedChange: function () {
                syncTargetFromCheckboxes();
                if (requestPreviewUpdate) requestPreviewUpdate();
            },
            onLayoutChange: function () { renameDialog.layout.layout(true); }
        });

        var dialogButtons = addDialogButtonRow(renameDialog);

        dialogButtons.okButton.onClick = function () {
            reorderList.syncEditingValues();
            var artboardEntries = reorderList.getEntries();
            for (var i = 0; i < artboardEntries.length; i++) {
                if (artboardEntries[i].checked && artboardEntries[i].newName === "") {
                    /* 該当行が見えるところまでスクロールしてから知らせる / Scroll to the offending row before alerting */
                    reorderList.revealEntry(i);
                    alert(getLabel("alert", "emptyName").replace("{n}", (i + 1)));
                    return;
                }
            }
            renameDialog.close(1);
        };

        /* 全UI構築後に初回レンダリング / Render the list once every control exists */
        reorderList.refreshRows();

        return {
            dialog: renameDialog,
            prefixInput: prefixInput,
            suffixInput: suffixInput,
            frontmostRadio: textSourceControls.frontmostRadio,
            layerRadio: textSourceControls.layerRadio,
            originalNameRadio: textSourceControls.originalNameRadio,
            customTextRadio: textSourceControls.customTextRadio,
            customTextInput: textSourceControls.customTextInput,
            layerDropdown: textSourceControls.layerDropdown,
            allArtboardsRadio: targetControls.allArtboardsRadio,
            rangeArtboardsRadio: targetControls.rangeArtboardsRadio,
            rangeInput: targetControls.rangeInput,
            refreshButton: dialogButtons.refreshButton,
            reorderList: reorderList,
            setRequestPreviewUpdate: function (previewUpdater) { requestPreviewUpdate = previewUpdater; }
        };
    }

    // =========================================
    // ダイアログのイベント / Dialog events
    // =========================================

    /**
     * ダイアログ各コントロールから設定オブジェクトを構築する
     * @param {object} dialogUI - createRenameDialog が返したコントロール群
     * @returns {object} リネーム設定
     */
    function readDialogSettings(dialogUI) {
        var textSourceMode = dialogUI.frontmostRadio.value ? "frontmost"
            : (dialogUI.layerRadio.value ? "layer"
                : (dialogUI.originalNameRadio.value ? "originalName" : "custom"));
        return {
            textSourceMode: textSourceMode,
            prefix: dialogUI.prefixInput.text,
            suffix: dialogUI.suffixInput.text,
            customText: dialogUI.customTextInput.text,
            artboardTarget: dialogUI.allArtboardsRadio.value ? "all" : "numbered",
            rangeText: dialogUI.rangeInput.text,
            selectedLayerName: (dialogUI.layerRadio.value && dialogUI.layerDropdown.selection)
                ? dialogUI.layerDropdown.selection.text
                : null
        };
    }

    /**
     * ダイアログ各コントロールにイベントハンドラを設定する
     * @param {object} dialogUI - createRenameDialog が返したコントロール群
     * @param {Document} doc - 対象ドキュメント
     * @param {string[]} originalNames - ダイアログ表示前のアートボード名
     * @returns {function} 現在のダイアログ状態からリネーム設定を作る関数
     */
    function bindDialogEvents(dialogUI, doc, originalNames) {
        /* originalName モードの参照元。［更新］で並び替えを確定するたび canvas に合わせて並べ替える
           / Reference names for the originalName mode, kept in sync with the canvas order */
        var sourceNames = [];
        for (var nameIndex = 0; nameIndex < originalNames.length; nameIndex++) {
            sourceNames.push(originalNames[nameIndex]);
        }

        /**
         * 現在のダイアログ状態からリネーム設定を作る
         * @returns {object} リネーム設定
         */
        function buildActiveSettings() {
            var renameSettings = readDialogSettings(dialogUI);
            renameSettings.originalNames = sourceNames;

            /* チェック行を canvas 位置へ変換して対象にする（並び替え中も正しいアートボードを指す）
               / Translate the checked rows into canvas indices so reordered rows still point at the right artboards */
            var artboardEntries = dialogUI.reorderList.getEntries();
            var targetIndices = [];
            for (var i = 0; i < artboardEntries.length; i++) {
                if (artboardEntries[i].checked) targetIndices.push(artboardEntries[i].originalIndex);
            }
            renameSettings.targetIndices = targetIndices;
            return renameSettings;
        }

        /**
         * canvas は触らず、未確定プレビュー名で右カラムを更新する
         * @returns {void}
         */
        function updatePreview() {
            dialogUI.reorderList.applyPreviewNames(computePreviewNames(doc, buildActiveSettings()));
        }

        /**
         * 対象アートボード設定をチェックボックスへ反映する
         * @returns {void}
         */
        function applyTargetToCheckboxes() {
            var targetType = dialogUI.allArtboardsRadio.value ? "all" : "numbered";
            dialogUI.reorderList.applyCheckedIndices(
                getTargetArtboardIndices(doc.artboards.length, targetType, dialogUI.rangeInput.text));
        }

        /**
         * 現在の設定と右カラムの手動編集を canvas に確定する（［更新］）
         * @returns {void}
         */
        function commitCurrentSettings() {
            /* 表示中の編集内容を取り込んでから、対象アートボード設定でチェック状態を確定
               / Capture pending row edits first, then settle the checked state from the target settings */
            dialogUI.reorderList.syncEditingValues();
            applyTargetToCheckboxes();

            var renameSettings = buildActiveSettings();
            renameSettings.artboardEntries = dialogUI.reorderList.getEntries();
            if (applyRenameAndReorder(doc, renameSettings, { silent: true })) {
                /* 確定した並びを基準に参照元とエントリーの位置を取り直す（二重適用と行ずれの防止）
                   / Re-anchor the reference names and entries to the committed order */
                sourceNames = reorderNamesByEntries(sourceNames, renameSettings.artboardEntries);
                dialogUI.reorderList.syncEntriesToCanvas();
                app.redraw();
            }
            updatePreview();
        }

        /**
         * 対象アートボードのラジオ切り替えを反映する
         * @returns {void}
         */
        function syncTargetInput() {
            dialogUI.rangeInput.enabled = dialogUI.rangeArtboardsRadio.value;
            applyTargetToCheckboxes();
            updatePreview();
        }

        dialogUI.allArtboardsRadio.onClick = function () {
            dialogUI.rangeArtboardsRadio.value = false;
            syncTargetInput();
        };
        dialogUI.rangeArtboardsRadio.onClick = function () {
            dialogUI.allArtboardsRadio.value = false;
            syncTargetInput();
        };
        dialogUI.rangeInput.onChange = function () {
            if (!dialogUI.rangeArtboardsRadio.value) return;
            applyTargetToCheckboxes();
            updatePreview();
        };

        /* 参照テキストのラジオは別コンテナに分かれているため排他を手動管理 / Radios live in different containers, so exclusivity is managed by hand */
        var textSourceRadios = [
            dialogUI.originalNameRadio, dialogUI.customTextRadio,
            dialogUI.frontmostRadio, dialogUI.layerRadio
        ];

        /**
         * 参照テキストのラジオを排他選択する
         * @param {RadioButton} selectedRadio - 選択されたラジオボタン
         * @returns {void}
         */
        function selectTextSourceRadio(selectedRadio) {
            for (var i = 0; i < textSourceRadios.length; i++) {
                textSourceRadios[i].value = (textSourceRadios[i] === selectedRadio);
            }
            dialogUI.layerDropdown.enabled = dialogUI.layerRadio.value;
            dialogUI.customTextInput.enabled = dialogUI.customTextRadio.value;
            if (dialogUI.customTextRadio.value) dialogUI.customTextInput.active = true;
            updatePreview();
        }

        for (var radioIndex = 0; radioIndex < textSourceRadios.length; radioIndex++) {
            (function (textSourceRadio) {
                textSourceRadio.onClick = function () { selectTextSourceRadio(textSourceRadio); };
            })(textSourceRadios[radioIndex]);
        }

        dialogUI.layerDropdown.onChange = function () {
            if (dialogUI.layerRadio.value) updatePreview();
        };
        dialogUI.customTextInput.onChange = function () {
            if (dialogUI.customTextRadio.value) updatePreview();
        };
        dialogUI.prefixInput.onChange = updatePreview;
        dialogUI.suffixInput.onChange = updatePreview;
        dialogUI.refreshButton.onClick = commitCurrentSettings;

        /* チェックボックス操作 → 対象アートボード設定 の逆同期でプレビューを更新 / Refresh the preview when checkboxes drive the target settings */
        dialogUI.setRequestPreviewUpdate(updatePreview);

        /* 初期同期：対象アートボード設定 → チェックボックス → プレビュー / Initial sync */
        applyTargetToCheckboxes();
        updatePreview();

        return buildActiveSettings;
    }

    /**
     * ダイアログを表示し、確定された設定を返す
     * @param {Document} doc - 対象ドキュメント
     * @param {string[]} originalNames - ダイアログ表示前のアートボード名
     * @returns {object|null} 確定した設定（キャンセル時は null）
     */
    function showRenameDialog(doc, originalNames) {
        var dialogUI = createRenameDialog(doc);
        var buildActiveSettings = bindDialogEvents(dialogUI, doc, originalNames);

        if (dialogUI.dialog.show() !== 1) return null;

        var renameSettings = buildActiveSettings();
        renameSettings.artboardEntries = dialogUI.reorderList.getEntries();
        return renameSettings;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * スクリプトのエントリポイント
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) return;

        var doc = app.activeDocument;
        var originalNames = [];
        var originalRects = [];
        for (var i = 0; i < doc.artboards.length; i++) {
            originalNames.push(doc.artboards[i].name);
            originalRects.push(doc.artboards[i].artboardRect);
        }

        var dialogResult = showRenameDialog(doc, originalNames);
        if (!dialogResult) {
            /* キャンセル：ダイアログを開く前の名前と並びまで戻す / Cancel: restore the names and order captured before the dialog opened */
            restoreOriginalArtboards(doc, originalNames, originalRects);
            return;
        }

        /* OK：途中の［更新］コミットを残したまま、最後の設定を適用 / OK: apply the final settings on top of any committed refresh */
        applyRenameAndReorder(doc, dialogResult);
    }

    main();

})();
