#target illustrator
#targetengine "SelectionInspectorSession"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択中またはドキュメント全体のオブジェクトを集計し、2カラムの常駐パレットで表示します。
テキスト・配置画像・透明・グループ・パス・ガイドの内訳を確認でき、選択オブジェクトのメモの閲覧・編集にも対応します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SelectionInspector.md

note記事も参照してください。
https://note.com/dtp_tranist/n/nefcb1ce828ce

### Overview

Tallies the objects in the selection, or in the whole document, and shows them in a two-column persistent palette.
It breaks down text, placed images, transparency, groups, paths and guides, and can view and edit the note on the selected object.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SelectionInspector.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SelectionInspector";           /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.7.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-08-06";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-22";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SelectionInspector.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SelectionInspector.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/nefcb1ce828ce"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // レイアウト / Layout
    // =========================================

    var LABEL_WIDTH_LEFT = 100;               /* 左カラムの項目名の幅 / Label width, left column */
    var LABEL_WIDTH_RIGHT = 130;              /* 右カラムの項目名の幅 / Label width, right column */
    var VALUE_WIDTH = 90;                     /* 値の幅 / Value width */
    var PANEL_MARGINS = [15, 20, 0, 10];      /* パネル余白 / Panel margins */
    var PALETTE_MARGINS = [15, 10, 15, 15];   /* パレット余白 / Palette margins */
    var TAB_MARGINS = [10, 15, 10, 10];       /* タブ（情報／メモ）の余白 / Tab margins */
    var MEMO_PREVIEW_SIZE = [160, 50];        /* 情報タブのメモ表示の寸法 / Note preview size on the Info tab */
    var MEMO_FIELD_SIZE = [340, 44];          /* メモタブの入力欄の寸法 / Note field size on the Notes tab */
    var PALETTE_OPACITY = 0.97;               /* パレットの不透明度 / Palette opacity */

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * Illustrator の UI 言語から表示言語を判定する
     * @returns {string} "ja" または "en"
     */
    function detectUILanguage() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }

    var uiLang = detectUILanguage();

    /* 日英ラベル定義（カテゴリ構造） / Japanese-English labels (categorized) */
    var LABELS = {
        dialog: {
            title: { ja: "選択／全体オブジェクトのカウント", en: "Selection / All Objects Count" }
        },
        report: {
            title: { ja: "Selection Inspector Report", en: "Selection Inspector Report" },
            document: { ja: "ドキュメント", en: "Document" },
            date: { ja: "日付", en: "Date" },
            valueNote: {
                ja: "※ 値は『選択 / 全体』の形式です（アートボードのみ全体）",
                en: "Note: values are formatted as 'Selection / All' (Artboards is All only)"
            }
        },
        panel: {
            basics: { ja: "基本", en: "Basics" },
            texts: { ja: "テキスト", en: "Text Frames" },
            charPara: { ja: "文字・段落", en: "Characters & Paragraphs" },
            images: { ja: "配置画像", en: "Images" },
            memo: { ja: "メモ", en: "Notes" },
            group: { ja: "グループ", en: "Groups" },
            transparency: { ja: "透明", en: "Transparency" },
            path: { ja: "パス", en: "Paths" },
            guide: { ja: "ガイド", en: "Guides" }
        },
        radio: {
            info: { ja: "情報", en: "Info" },
            memo: { ja: "メモ", en: "Notes" }
        },
        fieldLabel: {
            artboards: { ja: "アートボード", en: "Artboards" },
            objects: { ja: "オブジェクト", en: "Objects" },
            texts: { ja: "テキスト", en: "Text Frames" },
            pointText: { ja: "ポイント文字", en: "Point Type" },
            areaText: { ja: "エリア内文字", en: "Area Type" },
            pathText: { ja: "パス上文字", en: "Path Text" },
            chars: { ja: "文字数", en: "Characters" },
            paras: { ja: "段落数", en: "Paragraphs" },
            forcedBreaks: { ja: "強制改行", en: "Line Breaks" },
            linked: { ja: "リンク", en: "Linked Images" },
            embed: { ja: "埋め込み", en: "Embedded Images" },
            broken: { ja: "リンク切れ", en: "Broken Links" },
            group: { ja: "グループ", en: "Groups" },
            clipGroup: { ja: "クリップグループ", en: "Clipping Groups" },
            opacityLt100: { ja: "不透明度<100", en: "Opacity < 100" },
            blendNotNormal: { ja: "描画モード≠通常", en: "Blend Mode != Normal" },
            pathCount: { ja: "パス", en: "Paths" },
            openPath: { ja: "オープンパス", en: "Open Paths" },
            closedPath: { ja: "クローズパス", en: "Closed Paths" },
            anchors: { ja: "アンカーポイント", en: "Anchor Points" },
            handles: { ja: "ハンドル", en: "Handles" },
            compoundPath: { ja: "複合パス", en: "Compound Paths" },
            compoundShape: { ja: "複合シェイプ", en: "Compound Shapes" },
            rulerGuides: { ja: "ルーラーガイド", en: "Ruler Guides" },
            artboardGuides: { ja: "アートボードガイド", en: "Artboard Guides" },
            otherGuides: { ja: "その他のガイド", en: "Other Guides" }
        },
        button: {
            refresh: { ja: "更新", en: "Refresh" },
            exportReport: { ja: "書き出し", en: "Export" },
            applyMemo: { ja: "適用", en: "Apply" }
        },
        memo: {
            multiple: {
                ja: "複数のメモがあります。\n「メモ」タブで確認",
                en: "Multiple notes found.\nSee the \"Notes\" tab."
            },
            none: { ja: "選択オブジェクトがありません", en: "No objects selected" }
        },
        status: {
            ready: { ja: "準備完了", en: "Ready" },
            noDoc: { ja: "ドキュメントが開かれていません", en: "No document open" },
            wholeDoc: { ja: "選択なし（全体を集計）", en: "No selection (counting all)" },
            selectedPrefix: { ja: "選択 ", en: "Selected " },
            selectedSuffix: { ja: " 件を集計", en: " object(s)" },
            timeout: { ja: "Illustrator から応答がありません", en: "No response from Illustrator" },
            busy: { ja: "処理中です", en: "Busy" },
            error: { ja: "エラー", en: "Error" },
            memoApplied: { ja: "メモを適用しました", en: "Note applied" },
            selChanged: { ja: "選択が変わりました。更新してください", en: "Selection changed. Please refresh." },
            exportedPrefix: { ja: "書き出しました: ", en: "Exported: " },
            exportFailOpen: { ja: "ファイルを開けませんでした", en: "Failed to open the file" }
        },
        tooltip: {
            shortcut: { ja: "⌥I: 情報  ⌥M: メモ  ⌘R: 更新", en: "⌥I: Info  ⌥M: Notes  ⌘R: Refresh" },
            refresh: { ja: "選択内容を再集計（⌘R）", en: "Recount selection (Cmd+R)" },
            esc: { ja: "Esc で閉じる", en: "Press Esc to close" },
            exportReport: {
                ja: "集計結果をデスクトップにテキストファイル（count-ドキュメント名-日付.txt）で書き出します",
                en: "Writes the tallies to a text file on the desktop (count-<document>-<date>.txt)"
            },
            applyMemo: {
                ja: "この欄の内容を、対応するオブジェクトのメモに設定します（上から、同じ高さなら左から順）",
                en: "Sets this text as the note of the matching object (ordered top to bottom, then left to right)"
            },
            statValue: {
                ja: "選択 / 全体 の数です（アートボードは全体のみ）",
                en: "Selection / All (Artboards shows All only)"
            },
            handles: {
                ja: "アンカーポイントから伸びている方向線の数です",
                en: "Number of direction handles that stick out of their anchor points"
            },
            rulerGuides: {
                ja: "どのアートボードの幅・高さよりも長い水平／垂直のガイドです",
                en: "Horizontal or vertical guides longer than every artboard's width and height"
            },
            artboardGuides: {
                ja: "長さがいずれかのアートボードの幅または高さと同じ（±0.5pt）水平／垂直のガイドです",
                en: "Horizontal or vertical guides as long as an artboard's width or height (±0.5 pt)"
            },
            otherGuides: {
                ja: "ルーラーガイドにもアートボードガイドにも当たらないガイドです（斜め・閉じたパスなど）",
                en: "Guides that are neither ruler guides nor artboard guides (diagonal, closed, and so on)"
            }
        }
    };

    /**
     * LABELS からドット区切りのパスで表示言語のテキストを取り出す（途中が無くても落ちない）
     * @param {string} labelPath - "panel.basics" のようなドット区切りのキー
     * @returns {string} 表示言語のテキスト（見つからない場合は labelPath をそのまま返す）
     */
    function getLabel(labelPath) {
        var labelPathKeys = String(labelPath).split(".");
        var labelNode = LABELS;
        for (var i = 0; i < labelPathKeys.length; i++) {
            if (labelNode == null) return labelPath;
            labelNode = labelNode[labelPathKeys[i]];
        }
        if (labelNode == null) return labelPath;
        if (typeof labelNode === "string") return labelNode;
        if (typeof labelNode === "object" && labelNode[uiLang] != null) return labelNode[uiLang];
        return labelPath;
    }

    /**
     * コロン付きの項目名を返す（日本語は全角、英語は半角）
     * @param {string} labelPath - ラベルのパス
     * @returns {string} コロン付きの項目名
     */
    function labelText(labelPath) {
        return getLabel(labelPath) + (uiLang === "ja" ? "：" : ":");
    }

    /**
     * 書き出し用の項目名を返す（言語にかかわらず半角コロン）
     * @param {string} labelPath - ラベルのパス
     * @returns {string} 半角コロン付きの項目名
     */
    function reportLabelText(labelPath) {
        return getLabel(labelPath) + ":";
    }

    // =========================================
    // 集計項目の構成 / Tally layout
    // =========================================

    /* パネルごとの行。label は項目名（LABELS.fieldLabel）、key は集計結果のキー（<key>Sel / <key>All）、
       tip は項目名の helpTip（LABELS.tooltip）。アートボードは全体の値だけ、オブジェクト数はキー名が別形式
       Rows per panel: label = row label, key = result key (<key>Sel / <key>All), tip = helpTip on the label */
    var STAT_ROWS = {
        basics: [
            { label: "artboards", allKey: "artboards" },
            { label: "objects", selKey: "selCount", allKey: "allCount" }
        ],
        texts: [
            { label: "texts", key: "text" },
            { label: "pointText", key: "point" },
            { label: "areaText", key: "area" },
            { label: "pathText", key: "path" }
        ],
        charPara: [
            { label: "chars", key: "char" },
            { label: "paras", key: "para" },
            { label: "forcedBreaks", key: "fb" }
        ],
        images: [
            { label: "linked", key: "linked" },
            { label: "embed", key: "embed" },
            { label: "broken", key: "broken" }
        ],
        group: [
            { label: "group", key: "group" },
            { label: "clipGroup", key: "clip" }
        ],
        transparency: [
            { label: "opacityLt100", key: "opacity" },
            { label: "blendNotNormal", key: "blend" }
        ],
        path: [
            { label: "pathCount", key: "pathCount" },
            { label: "openPath", key: "open" },
            { label: "closedPath", key: "closed" },
            { label: "anchors", key: "anchor" },
            { label: "handles", key: "handle", tip: "handles" },
            { label: "compoundPath", key: "cpath" },
            { label: "compoundShape", key: "cshape" }
        ],
        guide: [
            { label: "rulerGuides", key: "ruler", tip: "rulerGuides" },
            { label: "artboardGuides", key: "abguide", tip: "artboardGuides" },
            { label: "otherGuides", key: "otherguide", tip: "otherGuides" }
        ]
    };

    /* パレットの2カラムに並べるパネル（memo はメモの表示欄）/ Panels in the palette's two columns */
    var PALETTE_COLUMNS = [
        ["basics", "texts", "charPara", "images", "memo"],
        ["group", "transparency", "path", "guide"]
    ];

    /* 書き出すセクションの順番 / Section order in the exported report */
    var REPORT_SECTIONS = ["basics", "texts", "charPara", "images", "memo", "transparency", "group", "path", "guide"];

    /**
     * 集計結果から1行ぶんの表示値を作る（「選択 / 全体」、全体のみの行は全体の値）
     * @param {Object} statRow - STAT_ROWS の1行
     * @param {Object} statMap - parseCollect() の map
     * @returns {string} 表示値
     */
    function formatStatValue(statRow, statMap) {
        if (!statRow.key && !statRow.selKey) return statMap[statRow.allKey];
        var selKey = statRow.selKey || (statRow.key + "Sel");
        var allKey = statRow.allKey || (statRow.key + "All");
        return statMap[selKey] + " / " + statMap[allKey];
    }

    /**
     * 空でないメモだけを取り出す
     * @param {string[]} memoList - メモの一覧
     * @returns {string[]} 空でないメモ
     */
    function collectNonEmptyNotes(memoList) {
        var nonEmptyNotes = [];
        for (var i = 0; i < memoList.length; i++) {
            if (memoList[i] && memoList[i] !== "") { nonEmptyNotes.push(memoList[i]); }
        }
        return nonEmptyNotes;
    }

    // =========================================
    // worker 関数（メインエンジンで実行）/ Worker functions (run in main engine)
    // -----------------------------------------
    // toString() で文字列にして BridgeTalk で送るため、JSDoc は付けない（構文エラーになる）
    // toString() は改行を全削除するため、関数内は // 行コメント禁止・/* */ のみ・各文は必ずセミコロンで終える
    // 関数を足したら WORKER_FUNCS にも登録する
    // Serialized with toString() and sent through BridgeTalk: no JSDoc, no // comments inside,
    // and every statement ends with a semicolon. Register new functions in WORKER_FUNCS.
    // =========================================

    /* アートボードの幅・高さのうち最大の値を返す / Largest artboard width or height */
    function wkGetMaxArtboardSpan(doc) {
        var maxSpan = 0;
        try {
            var artboards = doc.artboards;
            for (var i = 0; i < artboards.length; i++) {
                var artboardRect = artboards[i].artboardRect;
                var artboardWidth = Math.abs(artboardRect[2] - artboardRect[0]);
                var artboardHeight = Math.abs(artboardRect[1] - artboardRect[3]);
                if (artboardWidth > maxSpan) { maxSpan = artboardWidth; }
                if (artboardHeight > maxSpan) { maxSpan = artboardHeight; }
            }
        } catch (e) {}
        return maxSpan;
    }

    /* 水平・垂直の2点のガイドなら長さを、それ以外は -1 を返す / Length of a straight 2-point guide, otherwise -1 */
    function wkStraightGuideLength(pathItem) {
        if (!pathItem || pathItem.typename !== "PathItem") { return -1; }
        if (!pathItem.guides) { return -1; }
        if (pathItem.closed) { return -1; }
        if (!pathItem.pathPoints || pathItem.pathPoints.length !== 2) { return -1; }
        var anchor0 = pathItem.pathPoints[0].anchor;
        var anchor1 = pathItem.pathPoints[1].anchor;
        var straightTolerance = 0.01;
        var isVertical = Math.abs(anchor0[0] - anchor1[0]) <= straightTolerance;
        var isHorizontal = Math.abs(anchor0[1] - anchor1[1]) <= straightTolerance;
        if (!(isVertical || isHorizontal)) { return -1; }
        var dx = anchor0[0] - anchor1[0];
        var dy = anchor0[1] - anchor1[1];
        return Math.sqrt(dx * dx + dy * dy);
    }

    /* 長さがいずれかのアートボードの幅または高さと同じガイドか / Whether a guide is as long as an artboard side */
    function wkIsArtboardGuide(pathItem, doc) {
        try {
            var guideLength = wkStraightGuideLength(pathItem);
            if (guideLength < 0) { return false; }
            var lengthTolerance = 0.5;
            var artboards = doc.artboards;
            for (var i = 0; i < artboards.length; i++) {
                var artboardRect = artboards[i].artboardRect;
                if (Math.abs(guideLength - Math.abs(artboardRect[2] - artboardRect[0])) <= lengthTolerance) { return true; }
                if (Math.abs(guideLength - Math.abs(artboardRect[1] - artboardRect[3])) <= lengthTolerance) { return true; }
            }
            return false;
        } catch (e) { return false; }
    }

    /* どのアートボードの幅・高さよりも長いガイド（ルーラーガイド）か / Whether a guide is longer than every artboard side */
    function wkIsRulerGuide(pathItem, doc) {
        try {
            var guideLength = wkStraightGuideLength(pathItem);
            if (guideLength < 0) { return false; }
            var maxSpan = wkGetMaxArtboardSpan(doc);
            if (!(maxSpan > 0)) { return true; }
            return (guideLength > maxSpan);
        } catch (e) { return false; }
    }

    /* ガイド1本を種類別に数える / Classify one guide */
    function wkCountGuides(pathItem, doc) {
        var guideCounts = { ruler: 0, artboard: 0, other: 0 };
        try {
            if (!pathItem || pathItem.typename !== "PathItem") { return guideCounts; }
            if (!pathItem.guides) { return guideCounts; }
            if (wkIsRulerGuide(pathItem, doc)) { guideCounts.ruler = 1; }
            else if (wkIsArtboardGuide(pathItem, doc)) { guideCounts.artboard = 1; }
            else { guideCounts.other = 1; }
        } catch (e) {}
        return guideCounts;
    }

    /* ガイドの種類別の数を集計に足す / Add guide counts into a tally */
    function wkAddGuideCounts(tally, guideCounts) {
        tally.ruler += guideCounts.ruler;
        tally.abguide += guideCounts.artboard;
        tally.otherguide += guideCounts.other;
    }

    /* 不透明度が100未満か・描画モードが通常以外かを数える / Opacity below 100 and non-normal blend mode */
    function wkCountTransparency(pageItem) {
        var transparency = { opacityLt100: 0, blendNotNormal: 0 };
        if (!pageItem) { return transparency; }
        try { if (typeof pageItem.opacity === "number" && pageItem.opacity < 100) { transparency.opacityLt100 = 1; } } catch (e) {}
        try { if (pageItem.blendingMode !== undefined && pageItem.blendingMode !== BlendModes.NORMAL) { transparency.blendNotNormal = 1; } } catch (e2) {}
        return transparency;
    }

    /* ガイドのパスか / Whether a path is a guide */
    function wkIsGuidePath(pathItem) {
        try { return (pathItem && pathItem.typename === "PathItem" && pathItem.guides === true); } catch (e) { return false; }
    }

    /* アンカーから伸びている方向線の数を数える / Count direction handles that stick out of their anchors */
    function wkCountHandles(pathItem) {
        var handleCount = 0;
        try {
            var pathPoints = pathItem.pathPoints;
            for (var i = 0; i < pathPoints.length; i++) {
                var anchor = pathPoints[i].anchor;
                var leftDirection = pathPoints[i].leftDirection;
                var rightDirection = pathPoints[i].rightDirection;
                if (leftDirection[0] !== anchor[0] || leftDirection[1] !== anchor[1]) { handleCount++; }
                if (rightDirection[0] !== anchor[0] || rightDirection[1] !== anchor[1]) { handleCount++; }
            }
        } catch (e) {}
        return handleCount;
    }

    /* ガイドでないパス1本ぶんを集計に足す / Add one non-guide path into the path stats */
    function wkAddPathStats(pathItem, pathStats) {
        if (wkIsGuidePath(pathItem)) { return; }
        pathStats.pathCount++;
        pathStats.anchorCount += pathItem.pathPoints.length;
        pathStats.handleCount += wkCountHandles(pathItem);
        if (pathItem.closed) { pathStats.closedPath++; } else { pathStats.openPath++; }
    }

    /* パスの数・アンカー・ハンドル・開閉を集計する（グループと複合パスは中まで）/ Path stats, recursing into groups and compound paths */
    function wkCountPathStats(pageItem, pathStats) {
        if (pageItem.typename === "GroupItem") {
            for (var gi = 0; gi < pageItem.pageItems.length; gi++) { wkCountPathStats(pageItem.pageItems[gi], pathStats); }
        } else if (pageItem.typename === "PathItem") {
            wkAddPathStats(pageItem, pathStats);
        } else if (pageItem.typename === "CompoundPathItem") {
            for (var ci = 0; ci < pageItem.pathItems.length; ci++) { wkAddPathStats(pageItem.pathItems[ci], pathStats); }
        }
    }

    /* 文字列に含まれる改行（\n）を数える / Count \n in a string */
    function wkCountForcedBreaks(contentsText) {
        if (!contentsText || !contentsText.length) { return 0; }
        var breakMatches = contentsText.match(/[\n]/g);
        return breakMatches ? breakMatches.length : 0;
    }

    /* テキストの数・文字数・段落数・種類を集計する（グループは中まで）/ Text stats, recursing into groups */
    function wkCountTextStats(pageItem, textStats) {
        if (pageItem.typename === "GroupItem") {
            for (var gi = 0; gi < pageItem.pageItems.length; gi++) { wkCountTextStats(pageItem.pageItems[gi], textStats); }
        } else if (pageItem.typename === "TextFrame") {
            textStats.textCount++;
            try { textStats.charCount += pageItem.characters.length; } catch (e) {}
            try { textStats.paraCount += pageItem.paragraphs.length; } catch (e2) {}
            try { textStats.forcedBreakCount += wkCountForcedBreaks(pageItem.contents); } catch (e3) {}
            if (pageItem.kind === TextType.POINTTEXT) { textStats.pointText++; }
            else if (pageItem.kind === TextType.AREATEXT) { textStats.areaText++; }
            else if (pageItem.kind === TextType.PATHTEXT) { textStats.pathText++; }
        }
    }

    /* 上端と左端を返す（取れなければ 0）/ Top and left of an item, 0 when unavailable */
    function wkTopLeft(pageItem) {
        var topLeft = [0, 0];
        try { var itemBounds = pageItem.geometricBounds; topLeft = [itemBounds[1], itemBounds[0]]; } catch (e) {}
        return topLeft;
    }

    /* 選択を上から、同じ高さなら左から並べ替える / Sort the selection top to bottom, then left to right */
    function wkSortSelection(selectedItems) {
        var sortedItems = [];
        for (var i = 0; i < selectedItems.length; i++) { sortedItems.push(selectedItems[i]); }
        sortedItems.sort(function (itemA, itemB) {
            var topLeftA = wkTopLeft(itemA);
            var topLeftB = wkTopLeft(itemB);
            if (topLeftA[0] !== topLeftB[0]) { return topLeftB[0] - topLeftA[0]; }
            return topLeftA[1] - topLeftB[1];
        });
        return sortedItems;
    }

    /* オブジェクト数を数える（グループと複合パスは中まで）/ Count items, recursing into groups and compound paths */
    function wkCountAllItems(pageItems) {
        var itemCount = 0;
        for (var i = 0; i < pageItems.length; i++) {
            itemCount++;
            if (pageItems[i].typename === "GroupItem") { itemCount += wkCountAllItems(pageItems[i].pageItems); }
            else if (pageItems[i].typename === "CompoundPathItem") { itemCount += wkCountAllItems(pageItems[i].pathItems); }
        }
        return itemCount;
    }

    /* オブジェクトの一覧を種類別に集計する（選択と全体で共通）/ Tally a list of items (shared by selection and document) */
    function wkTallyItems(pageItems, doc) {
        var tally = {
            cpath: 0, cshape: 0, opacity: 0, blend: 0, ruler: 0, abguide: 0, otherguide: 0,
            linked: 0, embed: 0, broken: 0, group: 0, clip: 0,
            path: { pathCount: 0, anchorCount: 0, handleCount: 0, openPath: 0, closedPath: 0 },
            text: { textCount: 0, charCount: 0, paraCount: 0, forcedBreakCount: 0, pointText: 0, areaText: 0, pathText: 0 }
        };
        for (var i = 0; i < pageItems.length; i++) {
            var pageItem = pageItems[i];
            if (pageItem.typename === "CompoundPathItem") { tally.cpath++; }
            if (pageItem.typename === "PluginItem") {
                try { if (pageItem.name && pageItem.name.indexOf("Compound Shape") !== -1) { tally.cshape++; } } catch (e) {}
            }
            var transparency = wkCountTransparency(pageItem);
            tally.opacity += transparency.opacityLt100;
            tally.blend += transparency.blendNotNormal;
            try {
                if (pageItem.typename === "PathItem") {
                    wkAddGuideCounts(tally, wkCountGuides(pageItem, doc));
                } else if (pageItem.typename === "CompoundPathItem") {
                    for (var j = 0; j < pageItem.pathItems.length; j++) { wkAddGuideCounts(tally, wkCountGuides(pageItem.pathItems[j], doc)); }
                }
            } catch (e2) {}
            wkCountPathStats(pageItem, tally.path);
            wkCountTextStats(pageItem, tally.text);
            if (pageItem.typename === "PlacedItem") {
                if (pageItem.embedded) { tally.embed++; }
                else {
                    tally.linked++;
                    try { var linkedFile = pageItem.file; if (!linkedFile || !linkedFile.exists) { tally.broken++; } } catch (e3) { tally.broken++; }
                }
            }
            if (pageItem.typename === "GroupItem") { tally.group++; if (pageItem.clipped) { tally.clip++; } }
        }
        return tally;
    }

    /* 「key+Sel=値」「key+All=値」の組を足す / Push a Sel/All pair */
    function wkPushPair(resultPairs, statKey, selValue, allValue) {
        resultPairs.push(statKey + "Sel=" + selValue);
        resultPairs.push(statKey + "All=" + allValue);
    }

    /* 集計して「OK|key=value|...|MEMO|件数|メモ...」の文字列で返す / Collect and return OK|key=value|...|MEMO|count|notes */
    function wkCollect() {
        if (app.documents.length === 0) { return "NODOC"; }
        var doc = app.activeDocument;
        var selectedItems = doc.selection;
        if (!selectedItems) { selectedItems = []; }
        var selCount = selectedItems.length;
        var allCount = wkCountAllItems(doc.pageItems);

        var selTally = wkTallyItems(selectedItems, doc);
        var allTally = wkTallyItems(doc.pageItems, doc);

        var sortedItems = wkSortSelection(selectedItems);
        var memoParts = [];
        for (var i = 0; i < sortedItems.length; i++) {
            var noteText = "";
            try { noteText = sortedItems[i].note || ""; } catch (e) { noteText = ""; }
            memoParts.push(encodeURIComponent(noteText));
        }

        var docName = "";
        try { docName = doc.name; } catch (e2) { docName = ""; }

        var resultPairs = [];
        resultPairs.push("selCount=" + selCount);
        resultPairs.push("allCount=" + allCount);
        resultPairs.push("artboards=" + doc.artboards.length);
        resultPairs.push("docName=" + encodeURIComponent(docName));
        wkPushPair(resultPairs, "text", selTally.text.textCount, allTally.text.textCount);
        wkPushPair(resultPairs, "point", selTally.text.pointText, allTally.text.pointText);
        wkPushPair(resultPairs, "area", selTally.text.areaText, allTally.text.areaText);
        wkPushPair(resultPairs, "path", selTally.text.pathText, allTally.text.pathText);
        wkPushPair(resultPairs, "char", selTally.text.charCount, allTally.text.charCount);
        wkPushPair(resultPairs, "para", selTally.text.paraCount, allTally.text.paraCount);
        wkPushPair(resultPairs, "fb", selTally.text.forcedBreakCount, allTally.text.forcedBreakCount);
        wkPushPair(resultPairs, "linked", selTally.linked, allTally.linked);
        wkPushPair(resultPairs, "embed", selTally.embed, allTally.embed);
        wkPushPair(resultPairs, "broken", selTally.broken, allTally.broken);
        wkPushPair(resultPairs, "group", selTally.group, allTally.group);
        wkPushPair(resultPairs, "clip", selTally.clip, allTally.clip);
        wkPushPair(resultPairs, "opacity", selTally.opacity, allTally.opacity);
        wkPushPair(resultPairs, "blend", selTally.blend, allTally.blend);
        wkPushPair(resultPairs, "pathCount", selTally.path.pathCount, allTally.path.pathCount);
        wkPushPair(resultPairs, "open", selTally.path.openPath, allTally.path.openPath);
        wkPushPair(resultPairs, "closed", selTally.path.closedPath, allTally.path.closedPath);
        wkPushPair(resultPairs, "anchor", selTally.path.anchorCount, allTally.path.anchorCount);
        wkPushPair(resultPairs, "handle", selTally.path.handleCount, allTally.path.handleCount);
        wkPushPair(resultPairs, "cpath", selTally.cpath, allTally.cpath);
        wkPushPair(resultPairs, "cshape", selTally.cshape, allTally.cshape);
        wkPushPair(resultPairs, "ruler", selTally.ruler, allTally.ruler);
        wkPushPair(resultPairs, "abguide", selTally.abguide, allTally.abguide);
        wkPushPair(resultPairs, "otherguide", selTally.otherguide, allTally.otherguide);

        return "OK|" + resultPairs.join("|") + "|MEMO|" + selCount + "|" + memoParts.join("|");
    }

    /* 並べ替えた選択の itemIndex 番目にメモを設定する / Set the note on the itemIndex-th item of the sorted selection */
    function wkApplyMemo(itemIndex, encodedNote) {
        if (app.documents.length === 0) { return "NODOC"; }
        var selectedItems = app.activeDocument.selection;
        if (!selectedItems) { selectedItems = []; }
        if (itemIndex < 0 || itemIndex >= selectedItems.length) { return "IDX"; }
        var sortedItems = wkSortSelection(selectedItems);
        try { sortedItems[itemIndex].note = decodeURIComponent(encodedNote); } catch (e) { return "ERR:" + e; }
        app.redraw();
        return "OK";
    }

    /* worker 関数は全登録（追加漏れ防止） / Register every worker function */
    var WORKER_FUNCS = [
        wkGetMaxArtboardSpan,
        wkStraightGuideLength,
        wkIsArtboardGuide,
        wkIsRulerGuide,
        wkCountGuides,
        wkAddGuideCounts,
        wkCountTransparency,
        wkIsGuidePath,
        wkCountHandles,
        wkAddPathStats,
        wkCountPathStats,
        wkCountForcedBreaks,
        wkCountTextStats,
        wkTopLeft,
        wkSortSelection,
        wkCountAllItems,
        wkTallyItems,
        wkPushPair,
        wkCollect,
        wkApplyMemo
    ];

    // =========================================
    // BridgeTalk 委譲 / Delegation to the main engine
    // =========================================

    var isBusy = false;

    /**
     * worker 関数一式をメインエンジンへ送り、指定の呼び出しを実行して結果を受け取る
     * @param {string} callExpression - メインエンジンで評価する呼び出し式（"wkCollect()" など）
     * @returns {string} 戻り値。失敗時は "ERR:" で始まる文字列
     */
    function callMainEngine(callExpression) {
        if (isBusy) { return "ERR:BUSY"; }
        isBusy = true;

        var resultHolder = { value: null };
        try {
            var workerSource = "";
            for (var i = 0; i < WORKER_FUNCS.length; i++) { workerSource += WORKER_FUNCS[i].toString(); }
            workerSource += callExpression + ";";

            var bridgeMessage = new BridgeTalk();
            bridgeMessage.target = "illustrator";
            bridgeMessage.body = "eval(decodeURIComponent(\"" + encodeURIComponent(workerSource) + "\"));";
            bridgeMessage.onResult = function (bridgeResult) {
                resultHolder.value = (bridgeResult && bridgeResult.body != null) ? String(bridgeResult.body) : "";
            };
            bridgeMessage.onError = function (bridgeError) {
                resultHolder.value = "ERR:" + ((bridgeError && bridgeError.body) ? bridgeError.body : "bridge");
            };
            bridgeMessage.send(10);
        } catch (e) {
            resultHolder.value = "ERR:" + e;
        } finally {
            isBusy = false;
        }

        if (resultHolder.value === null) { return "ERR:TIMEOUT"; }
        return resultHolder.value;
    }

    /**
     * wkCollect() の戻り値（OK|key=value|...|MEMO|count|note...）を解析する
     * @param {string} response - 戻り値
     * @returns {{map: Object, memoList: string[]}|null} 集計値とメモ。形式が違えば null
     */
    function parseCollect(response) {
        if (!response || response.indexOf("OK|") !== 0) return null;
        var responseBody = response.substring(3);
        var memoMarkerIndex = responseBody.indexOf("|MEMO|");
        if (memoMarkerIndex < 0) return null;
        var statPart = responseBody.substring(0, memoMarkerIndex);
        var memoPart = responseBody.substring(memoMarkerIndex + 6);

        var statMap = {};
        var statPairs = statPart.split("|");
        for (var i = 0; i < statPairs.length; i++) {
            var separatorIndex = statPairs[i].indexOf("=");
            if (separatorIndex > 0) { statMap[statPairs[i].substring(0, separatorIndex)] = statPairs[i].substring(separatorIndex + 1); }
        }
        if (statMap.docName != null) { try { statMap.docName = decodeURIComponent(statMap.docName); } catch (e) {} }

        var memoFieldsText = memoPart.split("|");
        var memoCount = parseInt(memoFieldsText[0], 10) || 0;
        var memoList = [];
        for (var k = 1; k <= memoCount && k < memoFieldsText.length; k++) {
            var noteText = memoFieldsText[k];
            try { noteText = decodeURIComponent(noteText); } catch (e2) {}
            memoList.push(noteText);
        }
        return { map: statMap, memoList: memoList };
    }

    // =========================================
    // 状態保持（常駐エンジン） / Session state (resident engine)
    // =========================================

    if (!$.global.__selectionInspectorState) {
        $.global.__selectionInspectorState = { location: null };
    }

    // =========================================
    // レポート書き出し / Report export
    // =========================================

    /**
     * 集計結果をデスクトップのテキストファイルに書き出す
     * @param {{map: Object, memoList: string[]}} collected - parseCollect() の結果
     * @returns {string|null} 書き出したファイルのパス。ファイルを開けなければ null
     */
    function writeReportFile(collected) {
        var statMap = collected.map;
        var fullName = statMap.docName || "";
        var baseName = fullName.replace(/\.[^\.]+$/, "");
        var today = new Date();
        var yyyy = today.getFullYear();
        var mm = ("0" + (today.getMonth() + 1)).slice(-2);
        var dd = ("0" + today.getDate()).slice(-2);

        var reportPath = Folder.desktop + "/count-" + baseName + "-" + yyyy + mm + dd + ".txt";
        var reportFile = new File(reportPath);
        if (!reportFile.open("w")) return null;

        reportFile.writeln(getLabel('report.title'));
        reportFile.writeln(labelText('report.document') + " " + fullName);
        reportFile.writeln(labelText('report.date') + " " + yyyy + "-" + mm + "-" + dd);
        reportFile.writeln("");
        reportFile.writeln(getLabel('report.valueNote'));

        for (var s = 0; s < REPORT_SECTIONS.length; s++) {
            var sectionKey = REPORT_SECTIONS[s];
            reportFile.writeln("");
            reportFile.writeln(getLabel('panel.' + sectionKey));
            if (sectionKey === "memo") {
                var nonEmptyNotes = collectNonEmptyNotes(collected.memoList);
                if (nonEmptyNotes.length > 0) { reportFile.writeln(nonEmptyNotes.join("\n")); }
                continue;
            }
            var statRows = STAT_ROWS[sectionKey];
            for (var r = 0; r < statRows.length; r++) {
                reportFile.writeln(reportLabelText('fieldLabel.' + statRows[r].label) + " " + formatStatValue(statRows[r], statMap));
            }
        }

        reportFile.close();
        return reportPath;
    }

    // =========================================
    // パレット構築 / Build palette
    // =========================================

    /**
     * 項目名と値を1行追加する
     * @param {Panel} statPanel - 追加先のパネル
     * @param {Object} statRow - STAT_ROWS の1行
     * @param {number} labelWidth - 項目名の幅
     * @returns {StaticText} 値を表示する statictext
     */
    function addStatRow(statPanel, statRow, labelWidth) {
        var rowGroup = statPanel.add("group");
        rowGroup.orientation = "row";
        var rowLabel = rowGroup.add("statictext", undefined, labelText('fieldLabel.' + statRow.label));
        rowLabel.justify = "right";
        rowLabel.preferredSize.width = labelWidth;
        if (statRow.tip) rowLabel.helpTip = getLabel('tooltip.' + statRow.tip);
        var valueText = rowGroup.add("statictext", undefined, "-");
        valueText.preferredSize.width = VALUE_WIDTH;
        valueText.helpTip = getLabel('tooltip.statValue');
        return valueText;
    }

    /**
     * 見出し付きのパネルを追加する
     * @param {Group} parentColumn - 追加先のカラム
     * @param {string} titlePath - 見出しのラベルのパス
     * @returns {Panel} 追加したパネル
     */
    function addPanel(parentColumn, titlePath) {
        var statPanel = parentColumn.add("panel", undefined, getLabel(titlePath));
        statPanel.orientation = "column";
        statPanel.alignChildren = ["fill", "top"];
        statPanel.margins = PANEL_MARGINS;
        return statPanel;
    }

    /**
     * 情報タブ（2カラムの集計パネルとメモの表示欄）を組み立てる
     * @param {Group} infoTab - 情報タブのグループ
     * @returns {{valueTexts: Object, memoPreview: StaticText}} 行ラベル名 → 値の statictext と、メモの表示欄
     */
    function buildInfoTab(infoTab) {
        var twoColGroup = infoTab.add("group");
        twoColGroup.orientation = "row";
        twoColGroup.alignChildren = ["fill", "top"];

        var valueTexts = {};
        var memoPreview = null;
        for (var c = 0; c < PALETTE_COLUMNS.length; c++) {
            var column = twoColGroup.add("group");
            column.orientation = "column";
            column.alignChildren = ["fill", "top"];
            var labelWidth = (c === 0) ? LABEL_WIDTH_LEFT : LABEL_WIDTH_RIGHT;

            for (var p = 0; p < PALETTE_COLUMNS[c].length; p++) {
                var panelKey = PALETTE_COLUMNS[c][p];
                var statPanel = addPanel(column, 'panel.' + panelKey);
                if (panelKey === "memo") {
                    memoPreview = statPanel.add("statictext", undefined, "", { multiline: true });
                    memoPreview.preferredSize = MEMO_PREVIEW_SIZE;
                    continue;
                }
                var statRows = STAT_ROWS[panelKey];
                for (var r = 0; r < statRows.length; r++) {
                    valueTexts[statRows[r].label] = addStatRow(statPanel, statRows[r], labelWidth);
                }
            }
        }
        return { valueTexts: valueTexts, memoPreview: memoPreview };
    }

    /**
     * 集計値をすべての行に反映する
     * @param {Object} valueTexts - 行ラベル名 → 値の statictext
     * @param {Object} statMap - parseCollect() の map
     * @returns {void}
     */
    function applyStatValues(valueTexts, statMap) {
        for (var panelKey in STAT_ROWS) {
            if (!STAT_ROWS.hasOwnProperty(panelKey)) continue;
            var statRows = STAT_ROWS[panelKey];
            for (var r = 0; r < statRows.length; r++) {
                valueTexts[statRows[r].label].text = formatStatValue(statRows[r], statMap);
            }
        }
    }

    /**
     * すべての行の値を「-」に戻す
     * @param {Object} valueTexts - 行ラベル名 → 値の statictext
     * @returns {void}
     */
    function clearStatValues(valueTexts) {
        for (var rowLabel in valueTexts) {
            if (valueTexts.hasOwnProperty(rowLabel)) { valueTexts[rowLabel].text = "-"; }
        }
    }

    /**
     * キーイベントの既定の動作を止める（止められないイベントは無視）
     * @param {Object} keyEvent - キーイベント
     * @returns {void}
     */
    function preventDefaultSafely(keyEvent) {
        try { if (keyEvent.preventDefault) keyEvent.preventDefault(); } catch (e) {}
    }

    /**
     * パレットのキー操作を組み込む
     * Esc: 閉じる / ⌥I・⌥M: タブ切替 / ⌘R: 更新（Enter はメモ編集と衝突するため使わない）
     * @param {Window} inspectorPalette - パレット
     * @param {{showInfo: Function, showMemo: Function, refresh: Function}} keyActions - キーに割り当てる処理
     * @returns {void}
     */
    function bindPaletteKeys(inspectorPalette, keyActions) {
        inspectorPalette.addEventListener("keydown", function (keyEvent) {
            var keyName = (keyEvent && keyEvent.keyName) ? String(keyEvent.keyName).toUpperCase() : "";

            if (keyName === "ESCAPE") {
                try { inspectorPalette.close(); } catch (e1) {}
            } else if (keyEvent && keyEvent.altKey && keyName === "I") {
                keyActions.showInfo();
                preventDefaultSafely(keyEvent);
            } else if (keyEvent && keyEvent.altKey && keyName === "M") {
                keyActions.showMemo();
                preventDefaultSafely(keyEvent);
            } else if (keyEvent && keyEvent.metaKey && keyName === "R") {
                keyActions.refresh();
                preventDefaultSafely(keyEvent);
            }
        });
    }

    /**
     * パレットを組み立てる
     * @returns {Window} 組み立てたパレット（未表示）
     */
    function buildPalette() {
        var inspectorPalette = new Window("palette", getLabel('dialog.title') + ' ' + SCRIPT_VERSION, undefined, { resizeable: false });
        inspectorPalette.orientation = "column";
        inspectorPalette.alignChildren = "center";
        inspectorPalette.margins = PALETTE_MARGINS;

        var memoFields = [];

        /* 表示切り替え（ラジオボタン＋stack）/ View switcher (radio buttons + stack) */
        var switchRow = inspectorPalette.add("group");
        switchRow.orientation = "row";
        switchRow.alignChildren = ["center", "center"];
        switchRow.helpTip = getLabel('tooltip.shortcut');

        var infoRadio = switchRow.add("radiobutton", undefined, getLabel('radio.info'));
        var memoRadio = switchRow.add("radiobutton", undefined, getLabel('radio.memo'));
        infoRadio.helpTip = getLabel('tooltip.shortcut');
        memoRadio.helpTip = getLabel('tooltip.shortcut');
        infoRadio.value = true;

        var stackWrap = inspectorPalette.add("group");
        stackWrap.orientation = "stack";
        stackWrap.alignChildren = ["fill", "fill"];

        var infoTab = stackWrap.add("group");
        infoTab.orientation = "column";
        infoTab.alignChildren = ["fill", "top"];
        infoTab.margins = TAB_MARGINS;

        var memoTab = stackWrap.add("group");
        memoTab.orientation = "column";
        memoTab.alignChildren = ["fill", "top"];
        memoTab.margins = TAB_MARGINS;
        memoTab.visible = false;

        /* 情報タブ（2カラム） / Info tab (two columns) */
        var infoParts = buildInfoTab(infoTab);
        var valueTexts = infoParts.valueTexts;
        var memoPreview = infoParts.memoPreview;

        /* ステータス / Status line */
        var statusText = inspectorPalette.add("statictext", undefined, getLabel('status.ready'));
        statusText.alignment = ["fill", "bottom"];

        /* ステータス行に表示する / Show a message on the status line */
        function setStatus(statusMessage) { statusText.text = statusMessage; }

        /* タブの中身が変わったあとに組み直す / Re-layout after the tab contents change */
        function relayout() {
            try { if (stackWrap.layout) { stackWrap.layout.layout(true); } } catch (e) {}
            try { if (inspectorPalette.layout) { inspectorPalette.layout.layout(true); } } catch (e2) {}
        }

        /* 情報タブとメモタブを切り替える / Switch between the Info and Notes tabs */
        function switchView(viewMode) {
            var showInfo = (viewMode === "info");
            infoRadio.value = showInfo;
            memoRadio.value = !showInfo;
            infoTab.visible = showInfo;
            memoTab.visible = !showInfo;
            relayout();
        }

        /* メモタブを開き、先頭の入力欄にフォーカスする / Open the Notes tab and focus its first field */
        function showMemoView() {
            switchView("memo");
            try { if (memoFields.length > 0) { memoFields[0].active = true; } } catch (e) {}
        }

        /* 情報タブのメモ欄を更新（1件ならその内容、複数なら案内）/ Update the note preview on the Info tab */
        function updateMemoPreview(memoList) {
            var nonEmptyNotes = collectNonEmptyNotes(memoList);
            memoPreview.text = (nonEmptyNotes.length === 1) ? nonEmptyNotes[0] : (nonEmptyNotes.length > 1 ? getLabel('memo.multiple') : "");
        }

        /* メモタブを再構築 / Rebuild the Notes tab */
        function rebuildMemo(memoList) {
            while (memoTab.children.length > 0) {
                try { memoTab.remove(memoTab.children[0]); } catch (e) { break; }
            }
            memoFields = [];

            if (!memoList || memoList.length === 0) {
                memoTab.add("statictext", undefined, getLabel('memo.none'));
                relayout();
                return;
            }

            for (var i = 0; i < memoList.length; i++) {
                (function (itemIndex, noteText) {
                    var memoRow = memoTab.add("group");
                    memoRow.orientation = "row";
                    memoRow.alignChildren = ["fill", "center"];
                    var memoField = memoRow.add("edittext", undefined, noteText, { multiline: true });
                    memoField.preferredSize = MEMO_FIELD_SIZE;
                    memoFields.push(memoField);
                    var btnApply = memoRow.add("button", undefined, getLabel('button.applyMemo'));
                    btnApply.helpTip = getLabel('tooltip.applyMemo');
                    btnApply.onClick = function () { applyMemo(itemIndex, memoField.text); };
                })(i, memoList[i]);
            }
            relayout();
        }

        /* 再集計 / Recount */
        function refresh() {
            setStatus(getLabel('status.busy'));
            var response = callMainEngine("wkCollect()");

            if (response === "ERR:BUSY") { setStatus(getLabel('status.busy')); return; }
            if (response === null || response === "ERR:TIMEOUT") { setStatus(getLabel('status.timeout')); return; }
            if (response === "NODOC") { setStatus(getLabel('status.noDoc')); clearStatValues(valueTexts); updateMemoPreview([]); rebuildMemo([]); return; }
            if (response.indexOf("ERR:") === 0) { setStatus(getLabel('status.error') + ": " + response.substring(4)); return; }

            var collected = parseCollect(response);
            if (!collected) { setStatus(getLabel('status.error')); return; }

            applyStatValues(valueTexts, collected.map);
            updateMemoPreview(collected.memoList);
            rebuildMemo(collected.memoList);

            var selectedCount = parseInt(collected.map.selCount, 10) || 0;
            if (selectedCount > 0) {
                setStatus(getLabel('status.selectedPrefix') + selectedCount + getLabel('status.selectedSuffix'));
            } else {
                setStatus(getLabel('status.wholeDoc'));
            }
        }

        /* メモ適用 / Apply a note */
        function applyMemo(itemIndex, noteText) {
            setStatus(getLabel('status.busy'));
            var response = callMainEngine("wkApplyMemo(" + itemIndex + ",\"" + encodeURIComponent(noteText) + "\")");
            if (response === "OK") { setStatus(getLabel('status.memoApplied')); refresh(); }
            else if (response === "NODOC") { setStatus(getLabel('status.noDoc')); }
            else if (response === "IDX") { setStatus(getLabel('status.selChanged')); }
            else { setStatus(getLabel('status.error') + ": " + response); }
        }

        /* レポート書き出し（収集データからパレット側で生成） / Export report */
        function exportReport() {
            setStatus(getLabel('status.busy'));
            var response = callMainEngine("wkCollect()");
            if (response === "NODOC") { setStatus(getLabel('status.noDoc')); return; }
            if (response === null || response === "ERR:TIMEOUT") { setStatus(getLabel('status.timeout')); return; }
            if (typeof response === "string" && response.indexOf("ERR:") === 0) { setStatus(getLabel('status.error') + ": " + response.substring(4)); return; }

            var collected = parseCollect(response);
            if (!collected) { setStatus(getLabel('status.error')); return; }

            try {
                var reportPath = writeReportFile(collected);
                if (reportPath !== null) {
                    setStatus(getLabel('status.exportedPrefix') + reportPath);
                } else {
                    setStatus(getLabel('status.exportFailOpen'));
                }
            } catch (err) {
                setStatus(getLabel('status.error') + ": " + err);
            }
        }

        /* ボタン行 / Button row */
        var btnRowGroup = inspectorPalette.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.alignChildren = ["fill", "center"];
        btnRowGroup.alignment = ["fill", "bottom"];

        var btnLeftGroup = btnRowGroup.add("group");
        btnLeftGroup.alignChildren = ["left", "center"];
        var btnExport = btnLeftGroup.add("button", undefined, getLabel('button.exportReport'));
        btnExport.helpTip = getLabel('tooltip.exportReport') + "\n" + getLabel('tooltip.esc');

        var spacer = btnRowGroup.add("statictext", undefined, "");
        spacer.alignment = ["fill", "fill"];
        spacer.minimumSize.width = 0;

        var btnRightGroup = btnRowGroup.add("group");
        btnRightGroup.alignChildren = ["right", "center"];
        var btnRefresh = btnRightGroup.add("button", undefined, getLabel('button.refresh'));
        btnRefresh.helpTip = getLabel('tooltip.refresh') + "\n" + getLabel('tooltip.esc');

        btnExport.onClick = exportReport;
        btnRefresh.onClick = refresh;

        infoRadio.onClick = function () { switchView("info"); };
        memoRadio.onClick = showMemoView;

        bindPaletteKeys(inspectorPalette, {
            showInfo: function () { switchView("info"); },
            showMemo: showMemoView,
            refresh: refresh
        });

        try { inspectorPalette.opacity = PALETTE_OPACITY; } catch (e) {}

        switchView("info");

        /* 表示直後に一度集計 / Count once right after showing */
        inspectorPalette.onShow = function () {
            refresh();
        };

        return inspectorPalette;
    }

    // =========================================
    // 位置の記憶・復元 / Remember & restore location
    // =========================================

    /**
     * 前回の位置にパレットを置く（記録が無ければ中央）
     * @param {Window} inspectorPalette - パレット
     * @returns {void}
     */
    function restoreLocation(inspectorPalette) {
        try {
            var savedLocation = $.global.__selectionInspectorState.location;
            if (savedLocation && savedLocation.length === 2) {
                inspectorPalette.location = [savedLocation[0], savedLocation[1]];
            } else {
                inspectorPalette.center();
            }
        } catch (e) {
            inspectorPalette.center();
        }
    }

    /**
     * パレットの位置を常駐エンジンに控える
     * @param {Window} inspectorPalette - パレット
     * @returns {void}
     */
    function rememberLocation(inspectorPalette) {
        try {
            if (inspectorPalette.location && inspectorPalette.location.length === 2) {
                $.global.__selectionInspectorState.location = [inspectorPalette.location[0], inspectorPalette.location[1]];
            }
        } catch (e) {}
    }

    // =========================================
    // 起動 / Entry point
    // =========================================

    /**
     * パレットを表示する（開いているものがあれば閉じてから）
     * @returns {void}
     */
    function showPalette() {
        /* 多重起動防止：既存パレットがあれば閉じる / Prevent duplicates */
        if ($.global.__SelectionInspectorPalette) {
            try { $.global.__SelectionInspectorPalette.close(); } catch (e) {}
            $.global.__SelectionInspectorPalette = null;
        }

        var inspectorPalette = buildPalette();

        /* 常駐エンジンの変数に保持して GC 回避 / Keep in resident engine to avoid GC */
        $.global.__SelectionInspectorPalette = inspectorPalette;
        inspectorPalette.onClose = function () {
            rememberLocation(inspectorPalette);
            app.redraw();
            $.global.__SelectionInspectorPalette = null;
        };

        restoreLocation(inspectorPalette);
        inspectorPalette.show();
    }

    showPalette();

})();
