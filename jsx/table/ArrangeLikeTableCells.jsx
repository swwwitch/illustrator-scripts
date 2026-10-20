#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
SCRIPTMETA-BEGIN
Script-ID=com.swwwitch.illustrator.ArrangeLikeTableCells
Name=ArrangeLikeTableCells
Author=Masahiro Takano (@swwwitch)
Target-App=Illustrator
Version=1.1.0
Release-Date=2026-10-20
Description-BEGIN
外枠の中の縦罫・横罫を均等に、またはテキストに合わせて並べ直し、罫線で区切られたセルの中央にテキストを字形の見た目で配置します。

Redistributes the rules inside the outer frame, evenly or to fit the texts, then centers each text in its cell based on its glyph outlines.
Description-END
SCRIPTMETA-END
*/

/*

### 概要

外枠の中の縦罫・横罫を均等に、またはテキストに合わせて並べ直し、罫線で区切られたセルの中央にテキストを字形の見た目で配置します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ArrangeLikeTableCells.md

### Overview

Redistributes the rules inside the outer frame, evenly or to fit the texts, then centers each text in its cell based on its glyph outlines.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ArrangeLikeTableCells.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ArrangeLikeTableCells";        /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-10-20";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-20";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ArrangeLikeTableCells.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ArrangeLikeTableCells.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* ［縦罫を並べ直す］［横罫を並べ直す］の初期値 / Initial value of "Redistribute vertical / horizontal rules" */
    var AVERAGE_RULES = true;

    /*
        ［間隔］の初期値 / Initial value of "Spacing"
        "even"    : 罫線を等間隔に並べる
                    Space the rules evenly.
        "content" : セルの大きさをテキストに合わせ、テキストの両側の余白をそろえる
                    Size each cell to its texts, with equal space around them.
    */
    var DISTRIBUTE_MODE = "even";

    /*
        ［長さを揃える］の初期値 / Initial value of "Match rule lengths"
        true  : 縦罫の上下、横罫の左右を外枠に合わせる（2点の罫線のみ）
                Snap rule endpoints to the outer frame (two-point rules only).
        false : 罫線の長さは変えず、位置だけ平均化する
                Keep rule lengths; redistribute positions only.
    */
    var MATCH_LINES_TO_RECT = true;

    /* ［セルの中央に配置］の初期値 / Initial value of "Center in cells" */
    var CENTER_IN_CELLS = true;

    /* ［中央揃えにする］の初期値 / Initial value of "Align center" */
    var JUSTIFY_CENTER = true;

    /* 縦・横の辺・罫線とみなす傾きの許容値（pt） / Tolerance for treating a segment or rule as vertical or horizontal (pt) */
    var TOLERANCE = 0.5;

    /* 位置の比較に挟む許容値（pt） / Tolerance for comparing positions (pt) */
    var EPSILON = 0.001;

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
     * 項目名の文言の末尾にコロンを付ける（日本語は半角スペース＋半角コロン「 :」、英語は「:」。Illustrator の線パネルなどの項目名に合わせる）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {Object|Array} [placeholderValues] - getLabel と同じ
     * @returns {string} コロン付きの文言
     */
    function labelText(labelRef, placeholderValues) {
        return getLabel(labelRef, placeholderValues) + (uiLang === "ja" ? " :" : ":");
    }

    /**
     * 「項目名 : 値」の1行を返す（日本語は「件数 : 5」、英語は「Count: 5」。どちらもコロンのあとに空白を入れる）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {string|number} value - コロンのあとに続ける値
     * @returns {string} 項目名と値をつないだ文字列
     */
    function labelValueText(labelRef, value) {
        return labelText(labelRef) + " " + value;
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

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "セルの中央に配置", en: "Center Text in Cells" }
        },
        panel: {
            rules: { ja: "罫線", en: "Rules" },
            text: { ja: "テキスト", en: "Text" }
        },
        checkbox: {
            averageVertical: { ja: "縦罫を並べ直す", en: "Redistribute vertical rules" },
            averageHorizontal: { ja: "横罫を並べ直す", en: "Redistribute horizontal rules" },
            matchRuleLengths: { ja: "長さを揃える", en: "Match rule lengths" },
            centerInCells: { ja: "セルの中央に配置", en: "Center in cells" },
            justifyCenter: { ja: "中央揃えにする", en: "Align center" },
            preview: { ja: "プレビュー", en: "Preview" }
        },
        fieldLabel: {
            spacing: { ja: "間隔", en: "Spacing" }
        },
        radio: {
            even: { ja: "均等", en: "Even" },
            content: { ja: "テキストに合わせる", en: "Fit to text" }
        },
        tooltip: {
            averageVertical: {
                ja: "外枠の中にある縦罫を、［間隔］の指定どおりに並べ直します。Option キーを押しながらクリックすると横罫もまとめて切り替えます。",
                en: "Redistributes the vertical rules inside the frame as set in Spacing. Option-click to toggle the horizontal rules too."
            },
            averageHorizontal: {
                ja: "外枠の中にある横罫を、［間隔］の指定どおりに並べ直します。Option キーを押しながらクリックすると縦罫もまとめて切り替えます。",
                en: "Redistributes the horizontal rules inside the frame as set in Spacing. Option-click to toggle the vertical rules too."
            },
            even: {
                ja: "罫線を等間隔に並べます。",
                en: "Spaces the rules evenly."
            },
            content: {
                ja: "セルの幅（高さ）をテキストに合わせ、テキストの両側の余白をそろえます。",
                en: "Sizes each column (row) to its texts, with equal space on both sides."
            },
            matchRuleLengths: {
                ja: "並べ直した罫線の長さを外枠の辺にそろえます。",
                en: "Matches the length of the redistributed rules to the sides of the frame."
            },
            centerInCells: {
                ja: "テキストを、枠・罫線で区切られたセルの中央へ字形の見た目で移動します。",
                en: "Moves each text to the center of its cell, based on its glyph outlines."
            },
            justifyCenter: {
                ja: "段落の行揃えを中央揃えにします。",
                en: "Sets the paragraph alignment to center."
            },
            preview: {
                ja: "結果を画面で確認します。キャンセルすると元に戻ります。",
                en: "Shows the result on the canvas. Cancel restores the original layout."
            }
        },
        button: {
            ok: { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        alert: {
            noText: {
                ja: "テキストと、セルの枠・罫線をいっしょに選択してください。",
                en: "Select the texts together with the cell frame and rules."
            },
            noRule: {
                ja: "セルの枠・罫線が選択されていません。",
                en: "No cell frame or rules are selected."
            },
            skipped: {
                ja: "{count} 個のテキストは、囲むセルが見つからなかったため移動しませんでした。",
                en: "{count} text(s) were not moved because no enclosing cell was found."
            }
        }
    };

    // =========================================
    // 選択 / Selection
    // =========================================

    /**
     * 処理対象の選択オブジェクトを返す
     * @returns {PageItem[]|null} 選択オブジェクト（前面→背面の順）。ドキュメントがない・未選択・文字の選択中は null
     */
    function getValidSelection() {
        if (!app.documents.length) return null;
        var selectedItems = app.activeDocument.selection;
        /* 文字ツールで文字を選択中は、添字で要素を取れない TextRange が返る
           A text selection returns a TextRange that cannot be indexed */
        if (!selectedItems || !selectedItems.length || selectedItems.typename === "TextRange") return null;
        return selectedItems;
    }

    /**
     * グループ・複合パスの中までたどって、テキストとパスを振り分ける
     * @param {PageItem} item - 対象のオブジェクト
     * @param {TextFrame[]} textFrames - テキストの追加先
     * @param {PathItem[]} pathItems - パスの追加先
     * @returns {void}
     */
    function collectItems(item, textFrames, pathItems) {
        var i;
        switch (item.typename) {
            case "TextFrame":
                textFrames.push(item);
                break;
            case "PathItem":
                if (!item.guides) pathItems.push(item);
                break;
            case "CompoundPathItem":
                for (i = 0; i < item.pathItems.length; i++) pathItems.push(item.pathItems[i]);
                break;
            case "GroupItem":
                for (i = 0; i < item.pageItems.length; i++) collectItems(item.pageItems[i], textFrames, pathItems);
                break;
        }
    }

    // =========================================
    // セルの辺 / Cell edges
    // =========================================

    /**
     * パスのアンカー間の線分から、垂直・水平の辺を拾う
     * @param {PathItem[]} pathItems - 枠・罫線のパス
     * @returns {{vertical: Object[], horizontal: Object[]}} 垂直は {pos: x, from: y下, to: y上}、水平は {pos: y, from: x左, to: x右}
     */
    function collectEdges(pathItems) {
        var edges = { vertical: [], horizontal: [] };
        for (var i = 0; i < pathItems.length; i++) {
            var points = pathItems[i].pathPoints;
            var pointCount = points.length;
            var segmentCount = pathItems[i].closed ? pointCount : pointCount - 1;
            for (var j = 0; j < segmentCount; j++) {
                var startPoint = points[j].anchor;
                var endPoint = points[(j + 1) % pointCount].anchor;
                var dx = Math.abs(startPoint[0] - endPoint[0]);
                var dy = Math.abs(startPoint[1] - endPoint[1]);
                if (dx <= TOLERANCE && dy > TOLERANCE) {
                    edges.vertical.push({
                        pos: (startPoint[0] + endPoint[0]) / 2,
                        from: Math.min(startPoint[1], endPoint[1]),
                        to: Math.max(startPoint[1], endPoint[1])
                    });
                } else if (dy <= TOLERANCE && dx > TOLERANCE) {
                    edges.horizontal.push({
                        pos: (startPoint[1] + endPoint[1]) / 2,
                        from: Math.min(startPoint[0], endPoint[0]),
                        to: Math.max(startPoint[0], endPoint[0])
                    });
                }
            }
        }
        return edges;
    }

    /**
     * 基準位置の両側で、いちばん近い辺の位置を返す
     * @param {Object[]} edges - collectEdges の vertical または horizontal
     * @param {number} pos - 辺と直交する軸での基準位置
     * @param {number} along - 辺に沿う軸での基準位置（この範囲を覆う辺だけを使う）
     * @returns {{lower: number, upper: number}|null} 両側の辺の位置。片側でも無ければ null
     */
    function findNearestEdges(edges, pos, along) {
        var lower = null;
        var upper = null;
        for (var i = 0; i < edges.length; i++) {
            var edge = edges[i];
            if (along < edge.from - TOLERANCE || along > edge.to + TOLERANCE) continue;
            if (edge.pos < pos - EPSILON && (lower === null || edge.pos > lower)) lower = edge.pos;
            if (edge.pos > pos + EPSILON && (upper === null || edge.pos < upper)) upper = edge.pos;
        }
        if (lower === null || upper === null) return null;
        return { lower: lower, upper: upper };
    }

    /**
     * 罫線を動かす前の位置を、動かしたあとの位置に置き換える
     * @param {Object[]} positionMap - {from: 前の位置, to: 後の位置} の配列
     * @param {number} pos - 動かす前の辺の位置
     * @returns {number} 動かしたあとの位置。対応する罫線が無ければそのまま
     */
    function mapPosition(positionMap, pos) {
        for (var i = 0; i < positionMap.length; i++) {
            if (Math.abs(positionMap[i].from - pos) <= TOLERANCE) return positionMap[i].to;
        }
        return pos;
    }

    // =========================================
    // 外枠と罫線の判定 / Detecting the frame and rules
    // =========================================

    /**
     * もっとも面積の大きい閉じたパスを外枠とみなす
     * @param {PathItem[]} candidatePathItems - 候補のパス
     * @returns {PathItem|null} 外枠のパス。見つからなければ null
     */
    function findOuterRectangle(candidatePathItems) {
        var outerRectangleCandidate = null;
        var maxArea = 0;

        for (var i = 0; i < candidatePathItems.length; i++) {
            var pathItem = candidatePathItems[i];
            if (!pathItem.closed) continue;

            var geometricBounds = pathItem.geometricBounds;
            var pathWidth = geometricBounds[2] - geometricBounds[0];
            var pathHeight = geometricBounds[1] - geometricBounds[3];
            if (pathWidth <= 0 || pathHeight <= 0) continue;

            var pathArea = pathWidth * pathHeight;
            if (pathArea > maxArea) {
                maxArea = pathArea;
                outerRectangleCandidate = pathItem;
            }
        }
        return outerRectangleCandidate;
    }

    /**
     * 外枠の座標と寸法をまとめて返す
     * @param {PathItem} outerRect - 外枠のパス
     * @returns {{left: number, top: number, right: number, bottom: number, width: number, height: number}} 外枠の範囲
     */
    function getRectBounds(outerRect) {
        var outerBounds = outerRect.geometricBounds;
        return {
            left: outerBounds[0],
            top: outerBounds[1],
            right: outerBounds[2],
            bottom: outerBounds[3],
            width: outerBounds[2] - outerBounds[0],
            height: outerBounds[1] - outerBounds[3]
        };
    }

    /**
     * 外枠の中にある縦罫・横罫を振り分け、縦罫は左から右、横罫は上から下に並べる
     * @param {PathItem[]} candidatePathItems - 候補のパス
     * @param {PathItem|null} outerRect - 外枠のパス（対象から除く）。罫線で囲んだ枠なら null
     * @param {Object} rectBounds - getRectBounds() の戻り値
     * @returns {{vertical: PathItem[], horizontal: PathItem[]}} 縦罫と横罫
     */
    function classifyRules(candidatePathItems, outerRect, rectBounds) {
        var verticalRules = [];
        var horizontalRules = [];

        for (var i = 0; i < candidatePathItems.length; i++) {
            var pathItem = candidatePathItems[i];
            if (pathItem === outerRect) continue;

            if (isRuleInRect(pathItem, rectBounds, true)) {
                verticalRules.push(pathItem);
            } else if (isRuleInRect(pathItem, rectBounds, false)) {
                horizontalRules.push(pathItem);
            }
        }

        /* 縦罫は左から右へ、横罫は上から下へソート / Sort vertical L→R, horizontal T→B */
        verticalRules.sort(function (firstRule, secondRule) {
            return getCenterX(firstRule) - getCenterX(secondRule);
        });
        horizontalRules.sort(function (firstRule, secondRule) {
            return getCenterY(secondRule) - getCenterY(firstRule);
        });

        return { vertical: verticalRules, horizontal: horizontalRules };
    }

    /**
     * 外枠の中にある縦罫（または横罫）かどうかを判定する
     * 縦罫は「長さ＝高さ、並ぶ方向＝X」、横罫は「長さ＝幅、並ぶ方向＝Y」として同じ条件で見る
     * @param {PathItem} pathItem - 判定するパス
     * @param {Object} rectBounds - getRectBounds() の戻り値
     * @param {boolean} isVertical - true なら縦罫、false なら横罫として判定する
     * @returns {boolean} 罫線とみなせるとき true
     */
    function isRuleInRect(pathItem, rectBounds, isVertical) {
        if (pathItem.closed) return false;

        var geometricBounds = pathItem.geometricBounds;
        var pathWidth = Math.abs(geometricBounds[2] - geometricBounds[0]);
        var pathHeight = Math.abs(geometricBounds[1] - geometricBounds[3]);
        var centerX = (geometricBounds[0] + geometricBounds[2]) / 2;
        var centerY = (geometricBounds[1] + geometricBounds[3]) / 2;

        if (isVertical) {
            return hasRuleShape(pathHeight, pathWidth, rectBounds.height) &&
                isInsideRange(centerX, rectBounds.left, rectBounds.right) &&
                isNearRange(centerY, rectBounds.bottom, rectBounds.top);
        }
        return hasRuleShape(pathWidth, pathHeight, rectBounds.width) &&
            isInsideRange(centerY, rectBounds.bottom, rectBounds.top) &&
            isNearRange(centerX, rectBounds.left, rectBounds.right);
    }

    /**
     * 細長い形で、外枠の辺の 20% 以上の長さがあるかどうか
     * @param {number} ruleLength - 罫線の方向の長さ
     * @param {number} ruleThickness - 罫線と直交する方向の幅
     * @param {number} outerLength - 同じ方向の外枠の長さ
     * @returns {boolean} 罫線の形とみなせるとき true
     */
    function hasRuleShape(ruleLength, ruleThickness, outerLength) {
        if (ruleLength <= 0) return false;
        if (!(ruleThickness <= TOLERANCE || ruleLength > ruleThickness * 5)) return false;
        return !(ruleLength < outerLength * 0.2);
    }

    /**
     * 値が範囲の内側に、許容値より離れて収まっているかどうか（両端ちょうどは外）
     * @param {number} value - 調べる値
     * @param {number} rangeMin - 範囲の最小値
     * @param {number} rangeMax - 範囲の最大値
     * @returns {boolean} 内側なら true
     */
    function isInsideRange(value, rangeMin, rangeMax) {
        return value > rangeMin + TOLERANCE && value < rangeMax - TOLERANCE;
    }

    /**
     * 値が範囲に、許容値の はみ出しまで含めて収まっているかどうか
     * @param {number} value - 調べる値
     * @param {number} rangeMin - 範囲の最小値
     * @param {number} rangeMax - 範囲の最大値
     * @returns {boolean} 収まっていれば true
     */
    function isNearRange(value, rangeMin, rangeMax) {
        return value >= rangeMin - TOLERANCE && value <= rangeMax + TOLERANCE;
    }

    /**
     * パスの外接矩形の中心 X を返す
     * @param {PathItem} pathItem - 対象のパス
     * @returns {number} 中心の X 座標
     */
    function getCenterX(pathItem) {
        var geometricBounds = pathItem.geometricBounds;
        return (geometricBounds[0] + geometricBounds[2]) / 2;
    }

    /**
     * パスの外接矩形の中心 Y を返す
     * @param {PathItem} pathItem - 対象のパス
     * @returns {number} 中心の Y 座標
     */
    function getCenterY(pathItem) {
        var geometricBounds = pathItem.geometricBounds;
        return (geometricBounds[1] + geometricBounds[3]) / 2;
    }

    // =========================================
    // 罫線の再配置 / Redistributing rules
    // =========================================

    /**
     * 外枠と、その中の縦罫・横罫を見つける。
     * 外枠は、いちばん大きい閉じたパスと、罫線4本で囲んだ枠のうち、広いほうを使う
     * @param {PathItem[]} pathItems - 選択中のパス
     * @returns {{rectBounds: Object, vertical: PathItem[], horizontal: PathItem[], frameLines: Object|null}|null} 外枠の範囲と罫線（罫線で囲んだ枠なら、その4辺の罫線も）。外枠が無ければ null
     */
    function findRules(pathItems) {
        var outerRect = findOuterRectangle(pathItems);
        var rectBounds = outerRect ? getRectBounds(outerRect) : null;
        var lineFrame = findLineFrame(pathItems);
        var bounds = lineFrame ? lineFrame.bounds : null;
        if (bounds && rectBounds && bounds.width * bounds.height <= rectBounds.width * rectBounds.height + TOLERANCE) lineFrame = null;
        if (lineFrame) {
            /* 外側の罫線は外枠の辺になり、classifyRules() で罫線から外れる / The outermost rules become the frame and are left out by classifyRules() */
            outerRect = null;
            rectBounds = lineFrame.bounds;
        }
        if (!rectBounds) return null;

        var rules = classifyRules(pathItems, outerRect, rectBounds);
        return {
            rectBounds: rectBounds,
            vertical: rules.vertical,
            horizontal: rules.horizontal,
            frameLines: lineFrame ? { vertical: lineFrame.vertical, horizontal: lineFrame.horizontal } : null
        };
    }

    /**
     * 長方形を使わず、罫線4本で囲んだ枠を見つける。
     * 左右は外側の縦罫の位置、上下は外側の横罫の位置で決める（罫線の端がはみ出していても枠は広がらない）
     * @param {PathItem[]} pathItems - 選択中のパス
     * @returns {{bounds: Object, vertical: PathItem[], horizontal: PathItem[]}|null} 枠の範囲と、枠になる縦罫・横罫。縦・横とも2本以上なければ null
     */
    function findLineFrame(pathItems) {
        var verticalLines = [];
        var horizontalLines = [];
        var i;
        for (i = 0; i < pathItems.length; i++) {
            if (pathItems[i].closed) continue;
            var geometricBounds = pathItems[i].geometricBounds;
            var pathWidth = geometricBounds[2] - geometricBounds[0];
            var pathHeight = geometricBounds[1] - geometricBounds[3];
            if (pathWidth <= TOLERANCE && pathHeight > TOLERANCE) verticalLines.push(pathItems[i]);
            else if (pathHeight <= TOLERANCE && pathWidth > TOLERANCE) horizontalLines.push(pathItems[i]);
        }
        if (verticalLines.length < 2 || horizontalLines.length < 2) return null;

        var left = Infinity;
        var right = -Infinity;
        var top = -Infinity;
        var bottom = Infinity;
        for (i = 0; i < verticalLines.length; i++) {
            left = Math.min(left, getCenterX(verticalLines[i]));
            right = Math.max(right, getCenterX(verticalLines[i]));
        }
        for (i = 0; i < horizontalLines.length; i++) {
            top = Math.max(top, getCenterY(horizontalLines[i]));
            bottom = Math.min(bottom, getCenterY(horizontalLines[i]));
        }
        if (right - left <= TOLERANCE || top - bottom <= TOLERANCE) return null;

        var frameVertical = [];
        var frameHorizontal = [];
        for (i = 0; i < verticalLines.length; i++) {
            var centerX = getCenterX(verticalLines[i]);
            if (Math.abs(centerX - left) <= TOLERANCE || Math.abs(centerX - right) <= TOLERANCE) frameVertical.push(verticalLines[i]);
        }
        for (i = 0; i < horizontalLines.length; i++) {
            var centerY = getCenterY(horizontalLines[i]);
            if (Math.abs(centerY - top) <= TOLERANCE || Math.abs(centerY - bottom) <= TOLERANCE) frameHorizontal.push(horizontalLines[i]);
        }

        return {
            bounds: { left: left, top: top, right: right, bottom: bottom, width: right - left, height: top - bottom },
            vertical: frameVertical,
            horizontal: frameHorizontal
        };
    }

    /**
     * 罫線で囲んだ枠の4辺を、枠の角から角までの長さにそろえる（2点の罫線のみ）
     * @param {Object} frameLines - findRules() の frameLines
     * @param {Object} rectBounds - 枠の範囲
     * @returns {void}
     */
    function matchFrameLineLengths(frameLines, rectBounds) {
        var i;
        for (i = 0; i < frameLines.vertical.length; i++) {
            if (frameLines.vertical[i].pathPoints.length !== 2) continue;
            setTwoPointVerticalLine(frameLines.vertical[i], getCenterX(frameLines.vertical[i]), rectBounds.top, rectBounds.bottom);
        }
        for (i = 0; i < frameLines.horizontal.length; i++) {
            if (frameLines.horizontal[i].pathPoints.length !== 2) continue;
            setTwoPointHorizontalLine(frameLines.horizontal[i], getCenterY(frameLines.horizontal[i]), rectBounds.left, rectBounds.right);
        }
    }

    /**
     * 縦罫・横罫を、設定に従って並べ直す
     * @param {Object|null} rules - findRules() の戻り値（columnSizes・rowSizes を足したもの）
     * @param {{averageVertical: boolean, averageHorizontal: boolean, distributeMode: string, matchRuleLengths: boolean}} settings - 設定
     * @returns {{vertical: Object[], horizontal: Object[]}} 罫線ごとの {from: 前の位置, to: 後の位置}
     */
    function averageRules(rules, settings) {
        var positionMap = { vertical: [], horizontal: [] };
        if (!rules) return positionMap;

        var fitsContent = settings.distributeMode === "content";
        if (settings.averageVertical) {
            var columnOffsets = computeRuleOffsets(rules.rectBounds.width, rules.vertical.length, fitsContent ? rules.columnSizes : null);
            positionMap.vertical = redistributeVerticalRules(rules.vertical, rules.rectBounds, settings.matchRuleLengths, columnOffsets);
        }
        if (settings.averageHorizontal) {
            var rowOffsets = computeRuleOffsets(rules.rectBounds.height, rules.horizontal.length, fitsContent ? rules.rowSizes : null);
            positionMap.horizontal = redistributeHorizontalRules(rules.horizontal, rules.rectBounds, settings.matchRuleLengths, rowOffsets);
        }
        if (rules.frameLines && settings.matchRuleLengths && (settings.averageVertical || settings.averageHorizontal)) {
            matchFrameLineLengths(rules.frameLines, rules.rectBounds);
        }
        return positionMap;
    }

    /**
     * 罫線を置く位置を、外枠の辺からの距離で返す。
     * 区画ごとに「中身の大きさ＋共通の余白」を割り当てる（中身が無ければ等間隔）
     * @param {number} totalLength - 外枠の幅（高さ）
     * @param {number} ruleCount - 罫線の数
     * @param {number[]|null} contentSizes - 区画ごとの中身の大きさ（罫線の数＋1 個）。null なら等間隔
     * @returns {number[]} 罫線ごとの、外枠の左辺（上辺）からの距離
     */
    function computeRuleOffsets(totalLength, ruleCount, contentSizes) {
        var sectionCount = ruleCount + 1;
        var sectionSizes = [];
        var contentTotal = 0;
        var i;
        for (i = 0; i < sectionCount; i++) {
            var sectionSize = contentSizes ? (contentSizes[i] || 0) : 0;
            sectionSizes.push(sectionSize);
            contentTotal += sectionSize;
        }

        var padding = (totalLength - contentTotal) / sectionCount;
        var scale = 1;
        /* 中身が外枠に収まらないときは、余白なしで中身の比率に縮める / If the content does not fit, drop the padding and scale by the content ratio */
        if (padding < 0) {
            padding = 0;
            scale = totalLength / contentTotal;
        }

        var offsets = [];
        var offset = 0;
        for (i = 0; i < ruleCount; i++) {
            offset += sectionSizes[i] * scale + padding;
            offsets.push(offset);
        }
        return offsets;
    }

    /**
     * 罫線で区切られた区画ごとに、中にあるテキストの最大の幅（高さ）を測る
     * @param {Object[]} textCells - main() で集めたテキストの情報
     * @param {Object} rules - findRules() の戻り値
     * @param {boolean} isVertical - true なら縦罫で区切った列の幅、false なら横罫で区切った行の高さ
     * @returns {number[]} 区画ごとの大きさ（左から右、上から下。罫線の数＋1 個）
     */
    function measureContentSizes(textCells, rules, isVertical) {
        var sortedRules = isVertical ? rules.vertical : rules.horizontal;
        var rulePositions = [];
        var sizes = [0];
        var i;
        for (i = 0; i < sortedRules.length; i++) {
            rulePositions.push(isVertical ? getCenterX(sortedRules[i]) : getCenterY(sortedRules[i]));
            sizes.push(0);
        }

        var rectBounds = rules.rectBounds;
        for (i = 0; i < textCells.length; i++) {
            var glyphCenter = textCells[i].glyphCenter;
            if (!glyphCenter) continue;
            /* 外枠の外にあるテキストは数えない / Ignore texts outside the frame */
            if (!isNearRange(glyphCenter[0], rectBounds.left, rectBounds.right) ||
                !isNearRange(glyphCenter[1], rectBounds.bottom, rectBounds.top)) continue;

            var center = isVertical ? glyphCenter[0] : glyphCenter[1];
            var size = isVertical ? textCells[i].glyphSize[0] : textCells[i].glyphSize[1];
            /* 縦罫は左から右、横罫は上から下（Y は下ほど小さい）に並んでいる / Vertical rules run left to right, horizontal top to bottom (Y decreases downward) */
            var sectionIndex = 0;
            for (var j = 0; j < rulePositions.length; j++) {
                if (isVertical ? rulePositions[j] < center : rulePositions[j] > center) sectionIndex++;
            }
            if (size > sizes[sectionIndex]) sizes[sectionIndex] = size;
        }
        return sizes;
    }

    /**
     * 縦罫を指定の位置に並べ直す
     * @param {PathItem[]} verticalRules - 左から右に並んだ縦罫
     * @param {Object} rectBounds - getRectBounds() の戻り値
     * @param {boolean} matchRuleLengths - 2点の罫線の上下を外枠に合わせるかどうか
     * @param {number[]} ruleOffsets - 罫線ごとの、外枠の左辺からの距離
     * @returns {Object[]} 罫線ごとの {from: 前の X, to: 後の X}
     */
    function redistributeVerticalRules(verticalRules, rectBounds, matchRuleLengths, ruleOffsets) {
        var positionMap = [];
        var ruleCount = verticalRules.length;

        for (var i = 0; i < ruleCount; i++) {
            var verticalRule = verticalRules[i];
            var currentX = getCenterX(verticalRule);
            var targetX = rectBounds.left + ruleOffsets[i];
            positionMap.push({ from: currentX, to: targetX });

            if (matchRuleLengths && verticalRule.pathPoints.length === 2) {
                setTwoPointVerticalLine(verticalRule, targetX, rectBounds.top, rectBounds.bottom);
            } else {
                translatePath(verticalRule, targetX - currentX, 0);
            }
        }
        return positionMap;
    }

    /**
     * 横罫を指定の位置に並べ直す
     * @param {PathItem[]} horizontalRules - 上から下に並んだ横罫
     * @param {Object} rectBounds - getRectBounds() の戻り値
     * @param {boolean} matchRuleLengths - 2点の罫線の左右を外枠に合わせるかどうか
     * @param {number[]} ruleOffsets - 罫線ごとの、外枠の上辺からの距離
     * @returns {Object[]} 罫線ごとの {from: 前の Y, to: 後の Y}
     */
    function redistributeHorizontalRules(horizontalRules, rectBounds, matchRuleLengths, ruleOffsets) {
        var positionMap = [];
        var ruleCount = horizontalRules.length;

        for (var i = 0; i < ruleCount; i++) {
            var horizontalRule = horizontalRules[i];
            var currentY = getCenterY(horizontalRule);
            var targetY = rectBounds.top - ruleOffsets[i];
            positionMap.push({ from: currentY, to: targetY });

            if (matchRuleLengths && horizontalRule.pathPoints.length === 2) {
                setTwoPointHorizontalLine(horizontalRule, targetY, rectBounds.left, rectBounds.right);
            } else {
                translatePath(horizontalRule, 0, targetY - currentY);
            }
        }
        return positionMap;
    }

    /**
     * パスのすべてのアンカーポイントとハンドルを平行移動する
     * @param {PathItem} pathItem - 対象のパス
     * @param {number} deltaX - X 方向の移動量
     * @param {number} deltaY - Y 方向の移動量
     * @returns {void}
     */
    function translatePath(pathItem, deltaX, deltaY) {
        for (var i = 0; i < pathItem.pathPoints.length; i++) {
            movePathPoint(pathItem.pathPoints[i], deltaX, deltaY);
        }
    }

    /**
     * 2点の縦罫を、指定の X 位置で上端から下端までの線にする
     * @param {PathItem} pathItem - 2点のパス
     * @param {number} x - 配置する X 座標
     * @param {number} top - 上端の Y 座標
     * @param {number} bottom - 下端の Y 座標
     * @returns {void}
     */
    function setTwoPointVerticalLine(pathItem, x, top, bottom) {
        var firstPoint = pathItem.pathPoints[0];
        var secondPoint = pathItem.pathPoints[1];
        var isFirstOnTop = firstPoint.anchor[1] >= secondPoint.anchor[1];

        movePointTo(isFirstOnTop ? firstPoint : secondPoint, x, top);
        movePointTo(isFirstOnTop ? secondPoint : firstPoint, x, bottom);
    }

    /**
     * 2点の横罫を、指定の Y 位置で左端から右端までの線にする
     * @param {PathItem} pathItem - 2点のパス
     * @param {number} y - 配置する Y 座標
     * @param {number} left - 左端の X 座標
     * @param {number} right - 右端の X 座標
     * @returns {void}
     */
    function setTwoPointHorizontalLine(pathItem, y, left, right) {
        var firstPoint = pathItem.pathPoints[0];
        var secondPoint = pathItem.pathPoints[1];
        var isFirstOnLeft = firstPoint.anchor[0] <= secondPoint.anchor[0];

        movePointTo(isFirstOnLeft ? firstPoint : secondPoint, left, y);
        movePointTo(isFirstOnLeft ? secondPoint : firstPoint, right, y);
    }

    /**
     * アンカーポイントを指定の座標へ移動する（ハンドルも同じ量だけ動かす）
     * @param {PathPoint} pathPoint - 対象のアンカーポイント
     * @param {number} targetX - 移動先の X 座標
     * @param {number} targetY - 移動先の Y 座標
     * @returns {void}
     */
    function movePointTo(pathPoint, targetX, targetY) {
        movePathPoint(pathPoint, targetX - pathPoint.anchor[0], targetY - pathPoint.anchor[1]);
    }

    /**
     * アンカーポイントと両ハンドルを同じ量だけ動かす
     * @param {PathPoint} pathPoint - 対象のアンカーポイント
     * @param {number} deltaX - X 方向の移動量
     * @param {number} deltaY - Y 方向の移動量
     * @returns {void}
     */
    function movePathPoint(pathPoint, deltaX, deltaY) {
        pathPoint.anchor = [pathPoint.anchor[0] + deltaX, pathPoint.anchor[1] + deltaY];
        pathPoint.leftDirection = [pathPoint.leftDirection[0] + deltaX, pathPoint.leftDirection[1] + deltaY];
        pathPoint.rightDirection = [pathPoint.rightDirection[0] + deltaX, pathPoint.rightDirection[1] + deltaY];
    }

    // =========================================
    // テキスト / Text
    // =========================================

    /**
     * 字形の外形で測ったテキストの範囲を返す（アウトライン化した複製で測る）
     * @param {TextFrame} textFrame - 対象のテキスト
     * @returns {number[]|null} [左, 上, 右, 下]。字形が無ければ null
     */
    function getGlyphBounds(textFrame) {
        var duplicateFrame = textFrame.duplicate();
        var outlineGroup = null;
        var bounds = null;
        try {
            outlineGroup = duplicateFrame.createOutline();
            bounds = outlineGroup.geometricBounds;
        } catch (e) {
            bounds = null;
        }
        /* createOutline() は複製を消費するので、残っていれば消す / createOutline() consumes the duplicate; remove it only if it remains */
        try { duplicateFrame.remove(); } catch (e) { }
        if (outlineGroup) {
            try { outlineGroup.remove(); } catch (e) { }
        }
        return bounds;
    }


    /**
     * テキストの字形の中心を返す
     * @param {TextFrame} textFrame - 対象のテキスト
     * @returns {number[]|null} [X, Y]。字形が無ければ null
     */
    function getGlyphCenter(textFrame) {
        var bounds = getGlyphBounds(textFrame);
        if (!bounds) return null;
        return [(bounds[0] + bounds[2]) / 2, (bounds[1] + bounds[3]) / 2];
    }

    /**
     * 字形の範囲から、中心と大きさを求める
     * @param {number[]|null} bounds - getGlyphBounds() の戻り値
     * @returns {{center: number[], size: number[]}|null} 中心 [X, Y] と大きさ [幅, 高さ]。範囲が無ければ null
     */
    function getBoundsGeometry(bounds) {
        if (!bounds) return null;
        return {
            center: [(bounds[0] + bounds[2]) / 2, (bounds[1] + bounds[3]) / 2],
            size: [bounds[2] - bounds[0], bounds[1] - bounds[3]]
        };
    }

    /**
     * 段落の行揃えを中央揃えにする
     * @param {TextFrame} textFrame - 対象のテキスト
     * @returns {void}
     */
    function setJustifyCenter(textFrame) {
        textFrame.textRange.paragraphAttributes.justification = Justification.CENTER;
    }

    // =========================================
    // 状態の保存と復元 / Saving and restoring state
    // =========================================

    /**
     * 編集前のパスの座標を保存する
     * @param {PathItem[]} targetPathItems - 保存するパス
     * @returns {Object[]} パスごとのアンカーポイントの状態
     */
    function savePathState(targetPathItems) {
        var pathStates = [];

        for (var pathIndex = 0; pathIndex < targetPathItems.length; pathIndex++) {
            var pathItem = targetPathItems[pathIndex];
            var pointStates = [];

            for (var pointIndex = 0; pointIndex < pathItem.pathPoints.length; pointIndex++) {
                var pathPoint = pathItem.pathPoints[pointIndex];

                pointStates.push({
                    anchor: [pathPoint.anchor[0], pathPoint.anchor[1]],
                    leftDirection: [pathPoint.leftDirection[0], pathPoint.leftDirection[1]],
                    rightDirection: [pathPoint.rightDirection[0], pathPoint.rightDirection[1]],
                    pointType: pathPoint.pointType
                });
            }

            pathStates.push({ pathItem: pathItem, pointStates: pointStates });
        }

        return pathStates;
    }

    /**
     * 保存しておいた座標へ戻す
     * @param {Object[]} pathStates - savePathState() の戻り値
     * @returns {void}
     */
    function restorePathState(pathStates) {
        for (var pathIndex = 0; pathIndex < pathStates.length; pathIndex++) {
            var pathItem = pathStates[pathIndex].pathItem;
            var pointStates = pathStates[pathIndex].pointStates;

            for (var pointIndex = 0; pointIndex < pointStates.length; pointIndex++) {
                var pathPoint = pathItem.pathPoints[pointIndex];
                var pointState = pointStates[pointIndex];

                pathPoint.anchor = [pointState.anchor[0], pointState.anchor[1]];
                pathPoint.leftDirection = [pointState.leftDirection[0], pointState.leftDirection[1]];
                pathPoint.rightDirection = [pointState.rightDirection[0], pointState.rightDirection[1]];
                pathPoint.pointType = pointState.pointType;
            }
        }
    }

    /**
     * 編集前のテキストの位置と、段落ごとの行揃えを保存する
     * @param {TextFrame[]} textFrames - 保存するテキスト
     * @returns {Object[]} テキストごとの {textFrame, position, justifications}
     */
    function saveTextState(textFrames) {
        var textStates = [];
        for (var i = 0; i < textFrames.length; i++) {
            var textFrame = textFrames[i];
            var justifications = [];
            for (var j = 0; j < textFrame.paragraphs.length; j++) {
                justifications.push(textFrame.paragraphs[j].paragraphAttributes.justification);
            }
            textStates.push({
                textFrame: textFrame,
                position: [textFrame.position[0], textFrame.position[1]],
                justifications: justifications
            });
        }
        return textStates;
    }

    /**
     * 保存しておいた行揃えと位置へ戻す（行揃えで字形が動くので、位置はそのあとで戻す）
     * @param {Object[]} textStates - saveTextState() の戻り値
     * @returns {void}
     */
    function restoreTextState(textStates) {
        for (var i = 0; i < textStates.length; i++) {
            var textState = textStates[i];
            var paragraphs = textState.textFrame.paragraphs;
            for (var j = 0; j < textState.justifications.length && j < paragraphs.length; j++) {
                paragraphs[j].paragraphAttributes.justification = textState.justifications[j];
            }
            textState.textFrame.position = [textState.position[0], textState.position[1]];
        }
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択からテキスト・枠・罫線を集め、ダイアログで均等化・配置を確定する
     * @returns {void}
     */
    function main() {
        var selectedItems = getValidSelection();
        if (!selectedItems) {
            alert(getLabel("alert.noText"));
            return;
        }

        var textFrames = [];
        var pathItems = [];
        for (var i = 0; i < selectedItems.length; i++) collectItems(selectedItems[i], textFrames, pathItems);
        if (!textFrames.length) {
            alert(getLabel("alert.noText"));
            return;
        }

        var edges = collectEdges(pathItems);
        if (!edges.vertical.length || !edges.horizontal.length) {
            alert(getLabel("alert.noRule"));
            return;
        }

        /* 罫線を動かすとセルの判定がずれるので、先に今のセルを決めておく
           Find each text's cell before moving rules, since moving them changes the result */
        var textCells = [];
        for (var j = 0; j < textFrames.length; j++) {
            var glyphGeometry = getBoundsGeometry(getGlyphBounds(textFrames[j]));
            var glyphCenter = glyphGeometry ? glyphGeometry.center : null;
            textCells.push({
                textFrame: textFrames[j],
                glyphCenter: glyphCenter,
                glyphSize: glyphGeometry ? glyphGeometry.size : null,
                column: glyphCenter ? findNearestEdges(edges.vertical, glyphCenter[0], glyphCenter[1]) : null,
                row: glyphCenter ? findNearestEdges(edges.horizontal, glyphCenter[1], glyphCenter[0]) : null
            });
        }

        var rules = findRules(pathItems);
        if (rules) {
            /* ［テキストに合わせる］で使う区画ごとの大きさは、動かす前に測っておく / Measure section sizes for "Fit to text" before anything moves */
            rules.columnSizes = measureContentSizes(textCells, rules, true);
            rules.rowSizes = measureContentSizes(textCells, rules, false);
        }
        var verticalRuleCount = rules ? rules.vertical.length : 0;
        var horizontalRuleCount = rules ? rules.horizontal.length : 0;
        var pathState = savePathState(pathItems);
        var textState = saveTextState(textFrames);

        /**
         * 元の状態に戻したうえで、設定どおりに均等化・配置する
         * @param {{averageVertical: boolean, averageHorizontal: boolean, distributeMode: string, matchRuleLengths: boolean, centerInCells: boolean, justifyCenter: boolean}} settings - 設定
         * @returns {number} セルが見つからず移動しなかったテキストの数
         */
        function applySettings(settings) {
            restorePathState(pathState);
            restoreTextState(textState);

            var positionMap = averageRules(rules, settings);
            var skippedCount = 0;

            for (var k = 0; k < textCells.length; k++) {
                var textCell = textCells[k];
                if (settings.justifyCenter) setJustifyCenter(textCell.textFrame);
                if (!settings.centerInCells) continue;
                if (!textCell.column || !textCell.row) {
                    skippedCount++;
                    continue;
                }

                /* 行揃えを変えると字形が動くので測り直す / Re-measure, since changing the alignment moves the glyphs */
                var currentCenter = settings.justifyCenter ? getGlyphCenter(textCell.textFrame) : textCell.glyphCenter;
                if (!currentCenter) {
                    skippedCount++;
                    continue;
                }
                var cellCenterX = (mapPosition(positionMap.vertical, textCell.column.lower) + mapPosition(positionMap.vertical, textCell.column.upper)) / 2;
                var cellCenterY = (mapPosition(positionMap.horizontal, textCell.row.lower) + mapPosition(positionMap.horizontal, textCell.row.upper)) / 2;
                textCell.textFrame.translate(cellCenterX - currentCenter[0], cellCenterY - currentCenter[1]);
            }
            return skippedCount;
        }

        /* 開いた時点の設定（見つかった罫線は ON） / Initial settings: rules found start ON */
        var initialSettings = {
            averageVertical: AVERAGE_RULES && verticalRuleCount > 0,
            averageHorizontal: AVERAGE_RULES && horizontalRuleCount > 0,
            distributeMode: DISTRIBUTE_MODE,
            matchRuleLengths: MATCH_LINES_TO_RECT,
            centerInCells: CENTER_IN_CELLS,
            justifyCenter: JUSTIFY_CENTER
        };

        var dialogControls = buildDialog(verticalRuleCount, horizontalRuleCount, initialSettings);
        var verticalRulesCheckbox = dialogControls.verticalRulesCheckbox;
        var horizontalRulesCheckbox = dialogControls.horizontalRulesCheckbox;
        var spacingRow = dialogControls.spacingRow;
        var evenSpacingRadio = dialogControls.evenSpacingRadio;
        var contentSpacingRadio = dialogControls.contentSpacingRadio;
        var matchRuleLengthsCheckbox = dialogControls.matchRuleLengthsCheckbox;
        var centerInCellsCheckbox = dialogControls.centerInCellsCheckbox;
        var justifyCenterCheckbox = dialogControls.justifyCenterCheckbox;
        var previewCheckbox = dialogControls.previewCheckbox;

        /**
         * チェックボックスの状態を設定にまとめる（表示前は value を読み戻せないので、ユーザー操作のあとだけ呼ぶ）
         * @returns {Object} applySettings() に渡す設定
         */
        function getCheckedSettings() {
            return {
                averageVertical: verticalRulesCheckbox.value,
                averageHorizontal: horizontalRulesCheckbox.value,
                distributeMode: contentSpacingRadio.value ? "content" : "even",
                matchRuleLengths: matchRuleLengthsCheckbox.value,
                centerInCells: centerInCellsCheckbox.value,
                justifyCenter: justifyCenterCheckbox.value
            };
        }

        /**
         * ［間隔］［長さを揃える］は、どちらかの罫線を並べ直すときだけ操作できるようにする
         * @returns {void}
         */
        function updateRuleOptionsEnabled() {
            var isAveraging = verticalRulesCheckbox.value || horizontalRulesCheckbox.value;
            spacingRow.enabled = isAveraging;
            matchRuleLengthsCheckbox.enabled = isAveraging;
        }

        /**
         * プレビューを更新する（プレビューが OFF のときは元に戻すだけ）
         * @returns {void}
         */
        function updatePreview() {
            if (previewCheckbox.value) {
                applySettings(getCheckedSettings());
            } else {
                restorePathState(pathState);
                restoreTextState(textState);
            }
            app.redraw();
        }

        /**
         * 罫線チェックボックスのクリック処理（Option／Alt キーで縦罫・横罫をまとめて切り替える）
         * @param {Checkbox} clickedCheckbox - クリックされたチェックボックス
         * @returns {void}
         */
        function handleRuleCheckboxClick(clickedCheckbox) {
            if (ScriptUI.environment.keyboardState.altKey) {
                if (verticalRulesCheckbox.enabled) verticalRulesCheckbox.value = clickedCheckbox.value;
                if (horizontalRulesCheckbox.enabled) horizontalRulesCheckbox.value = clickedCheckbox.value;
            }
            updateRuleOptionsEnabled();
            updatePreview();
        }

        verticalRulesCheckbox.onClick = function () {
            handleRuleCheckboxClick(verticalRulesCheckbox);
        };
        horizontalRulesCheckbox.onClick = function () {
            handleRuleCheckboxClick(horizontalRulesCheckbox);
        };
        evenSpacingRadio.onClick = updatePreview;
        contentSpacingRadio.onClick = updatePreview;
        matchRuleLengthsCheckbox.onClick = updatePreview;
        centerInCellsCheckbox.onClick = updatePreview;
        justifyCenterCheckbox.onClick = updatePreview;
        previewCheckbox.onClick = updatePreview;

        var skippedCount = 0;
        dialogControls.btnOK.onClick = function () {
            skippedCount = applySettings(getCheckedSettings());
            app.redraw();
            dialogControls.mainDialog.close(1);
        };
        dialogControls.btnCancel.onClick = function () {
            dialogControls.mainDialog.close(0);
        };

        /* 「キャンセル」の文字では Esc の取消ボタンと認識されない / "キャンセル" is not recognized as the Esc button */
        dialogControls.mainDialog.cancelElement = dialogControls.btnCancel;

        /* 初回のプレビュー（プレビューは ON で始まる） / Initial preview (Preview starts ON) */
        applySettings(initialSettings);
        app.redraw();

        prepareDialogWindow(dialogControls.mainDialog, SCRIPT_NAME);

        /* Esc など［OK］以外で閉じたときも元に戻す / Restore however the dialog is closed, unless by OK */
        if (dialogControls.mainDialog.show() !== 1) {
            restorePathState(pathState);
            restoreTextState(textState);
            app.redraw();
            return;
        }

        if (skippedCount) alert(getLabel("alert.skipped", { count: skippedCount }));
    }

    // ボタン行（再利用パーツ） / Button row (reusable)

    var BUTTON_ROW_TOP_MARGIN = 5; /* ボタン行の上の余白 / top margin of the button row */
    var BUTTON_ROW_BOTTOM_MARGIN = 14; /* ボタン行の下の余白。ダイアログの下余白と合わせて約30px（Illustrator 標準のダイアログに合わせる） / bottom margin; with the dialog margin about 30px, like Illustrator's own dialogs */
    var BUTTON_ROW_SPACING = 10;   /* ボタンどうしの間隔 / spacing between buttons */
    var BUTTON_ROW_CENTER_MAX_WIDTH = 240; /* 右のボタンだけの行を中央に置く、ダイアログの内側の最大幅（px、左右の余白を除く）。広いダイアログは右揃え / max inner dialog width (px, margins excluded) that centers a right-only row; wider dialogs keep it right-aligned */

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
    // ダイアログ / Dialog
    // =========================================

    /**
     * ダイアログを組み立てる
     * @param {number} verticalRuleCount - 見つかった縦罫の数
     * @param {number} horizontalRuleCount - 見つかった横罫の数
     * @param {{averageVertical: boolean, averageHorizontal: boolean, distributeMode: string, matchRuleLengths: boolean, centerInCells: boolean, justifyCenter: boolean}} initialSettings - コントロールの初期値
     * @returns {{mainDialog: Window, verticalRulesCheckbox: Checkbox, horizontalRulesCheckbox: Checkbox, spacingRow: Group, evenSpacingRadio: RadioButton, contentSpacingRadio: RadioButton, matchRuleLengthsCheckbox: Checkbox, centerInCellsCheckbox: Checkbox, justifyCenterCheckbox: Checkbox, previewCheckbox: Checkbox, btnCancel: Button, btnOK: Button}} ダイアログとコントロール
     */
    function buildDialog(verticalRuleCount, horizontalRuleCount, initialSettings) {
        var mainDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        setupWindow(mainDialog);

        var rulesPanel = mainDialog.add("panel", undefined, getLabel("panel.rules"));
        setupPanel(rulesPanel, 6);

        var verticalRulesCheckbox = addOptionCheckbox(rulesPanel,
            "checkbox.averageVertical", "tooltip.averageVertical", verticalRuleCount > 0, initialSettings.averageVertical);
        var horizontalRulesCheckbox = addOptionCheckbox(rulesPanel,
            "checkbox.averageHorizontal", "tooltip.averageHorizontal", horizontalRuleCount > 0, initialSettings.averageHorizontal);
        var isAveraging = initialSettings.averageVertical || initialSettings.averageHorizontal;

        var spacingRow = rulesPanel.add("group");
        setupRow(spacingRow, "left", 8);
        spacingRow.add("statictext", undefined, labelText("fieldLabel.spacing"));
        var evenSpacingRadio = spacingRow.add("radiobutton", undefined, getLabel("radio.even"));
        evenSpacingRadio.helpTip = getLabel("tooltip.even");
        var contentSpacingRadio = spacingRow.add("radiobutton", undefined, getLabel("radio.content"));
        contentSpacingRadio.helpTip = getLabel("tooltip.content");
        contentSpacingRadio.value = initialSettings.distributeMode === "content";
        evenSpacingRadio.value = !contentSpacingRadio.value;
        spacingRow.enabled = isAveraging;

        var matchRuleLengthsCheckbox = addOptionCheckbox(rulesPanel,
            "checkbox.matchRuleLengths", "tooltip.matchRuleLengths", isAveraging, initialSettings.matchRuleLengths);

        var textPanel = mainDialog.add("panel", undefined, getLabel("panel.text"));
        setupPanel(textPanel, 6);

        var centerInCellsCheckbox = addOptionCheckbox(textPanel,
            "checkbox.centerInCells", "tooltip.centerInCells", true, initialSettings.centerInCells);
        var justifyCenterCheckbox = addOptionCheckbox(textPanel,
            "checkbox.justifyCenter", "tooltip.justifyCenter", true, initialSettings.justifyCenter);

        var previewGroup = mainDialog.add("group");
        previewGroup.orientation = "row";
        previewGroup.alignment = "center";

        var previewCheckbox = previewGroup.add("checkbox", undefined, getLabel("checkbox.preview"));
        previewCheckbox.helpTip = getLabel("tooltip.preview");
        previewCheckbox.value = true;

        var buttonRow = addButtonRow(mainDialog);
        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"));
        var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"));
        alignRightOnlyButtonRow(buttonRow);

        return {
            mainDialog: mainDialog,
            verticalRulesCheckbox: verticalRulesCheckbox,
            horizontalRulesCheckbox: horizontalRulesCheckbox,
            spacingRow: spacingRow,
            evenSpacingRadio: evenSpacingRadio,
            contentSpacingRadio: contentSpacingRadio,
            matchRuleLengthsCheckbox: matchRuleLengthsCheckbox,
            centerInCellsCheckbox: centerInCellsCheckbox,
            justifyCenterCheckbox: justifyCenterCheckbox,
            previewCheckbox: previewCheckbox,
            btnCancel: btnCancel,
            btnOK: btnOK
        };
    }

    /**
     * チェックボックスを tooltip 付きで追加する
     * @param {Panel} parentPanel - 追加先のパネル
     * @param {string} labelPath - 表示名の LABELS パス
     * @param {string} tooltipPath - tooltip の LABELS パス
     * @param {boolean} isEnabled - 操作できるかどうか
     * @param {boolean} initialValue - 初期値
     * @returns {Checkbox} 追加したチェックボックス
     */
    function addOptionCheckbox(parentPanel, labelPath, tooltipPath, isEnabled, initialValue) {
        var optionCheckbox = parentPanel.add("checkbox", undefined, getLabel(labelPath));
        optionCheckbox.helpTip = getLabel(tooltipPath);
        optionCheckbox.value = initialValue === true;
        optionCheckbox.enabled = isEnabled;
        return optionCheckbox;
    }

    main();

})();
