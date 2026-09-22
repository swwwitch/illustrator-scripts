#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);
#targetengine "ExtendLinesEngine"

/*

### 概要

選択オブジェクト内のパス（グループ／複合パスを含む）から隣接するアンカーポイントのペアを取り、補助線として「直線を描画範囲いっぱいに延長した線」を描画します。
［円弧から円］でBezier曲線セグメントから円を推定することもできます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ExtendLines.md

### Overview

Takes pairs of adjacent anchor points from the paths in the selection — groups and compound paths included — and draws each as a construction line extended across the drawing area.
The "arc to circle" option can also estimate a circle from a Bézier segment.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ExtendLines.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ExtendLines";                  /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.1";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-02-27";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-22";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ExtendLines.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ExtendLines.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

var SCRIPT_MARKER = "__ExtendLines__";

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* 線幅の既定値（0.1mm を pt で）/ Default stroke width (0.1 mm in points) */
    var DEFAULT_STROKE_WIDTH_PT = 0.1 * 72.0 / 25.4;

    /* ［別レイヤーに］の描画先レイヤー名 / Layer used by "Separate layer" */
    var CONSTRUCTION_LAYER_NAME = "_construction";

    /* プレビュー用レイヤー名のもと（重複時は連番を付ける）/ Base name of the preview layer (numbered when taken) */
    var PREVIEW_LAYER_BASE_NAME = "__ExtendLines_Preview";

    // =========================================
    // レイアウト / Layout
    // =========================================

    var DIALOG_MARGINS = 15;                 /* ダイアログの余白 / dialog margins */
    var PANEL_MARGINS = [15, 20, 15, 10];    /* パネルの余白 [左,上,右,下] / panel margins */
    var COLUMN_SPACING = 10;                 /* 2カラムの間隔 / gap between the two columns */
    var STROKE_ROW_SPACING = 6;              /* 線幅の行の間隔 / spacing in the stroke-width row */
    var STROKE_INPUT_CHARS = 6;              /* 線幅の入力欄の幅（文字数）/ width of the stroke-width field */
    var BUTTON_SPACING = 10;                 /* ボタンの間隔 / spacing between buttons */

    /**
     * パネルを縦並び・左揃えにして共通の余白を付ける
     * @param {Panel} targetPanel - 対象のパネル
     * @returns {void}
     */
    function applyPanelLayout(targetPanel) {
        targetPanel.orientation = "column";
        targetPanel.alignChildren = ["left", "top"];
        targetPanel.margins = PANEL_MARGINS;
    }

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * UI の表示言語を判定する
     * @returns {string} "ja" または "en"
     */
    function detectUILanguage() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = detectUILanguage();

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "補助線の描画", en: "Extend Lines" }
        },
        panel: {
            auxLines: { ja: "補助線を描画", en: "Draw guide lines" },
            arcOptions: { ja: "円弧オプション", en: "Arc options" },
            options: { ja: "オプション", en: "Options" }
        },
        checkbox: {
            straightOnly: { ja: "直線", en: "Straight" },
            arcToCircle: { ja: "円弧から円", en: "Create circle from arc" },
            group: { ja: "グループ化", en: "Group" },
            separateLayer: { ja: "別レイヤーに", en: "Separate layer" },
            guide: { ja: "ガイド化", en: "Convert to guides" },
            dedup: { ja: "線のダブりを削除", en: "Remove duplicates" },
            preview: { ja: "プレビュー", en: "Preview" }
        },
        radio: {
            arcIgnore: { ja: "無視", en: "Ignore" },
            arcStraight: { ja: "直線", en: "Straight" },
            arcExtend: { ja: "直線（延長）", en: "Straight (extend)" }
        },
        fieldLabel: {
            strokeWidth: { ja: "線幅", en: "Stroke width" }
        },
        tooltip: {
            arcOptions: { ja: "完全な円弧以外の場合", en: "If not a perfect circular arc" },
            straightOnly: {
                ja: "オンのときは直線のセグメントだけ、オフのときは曲線のセグメントだけを延長します。",
                en: "On: extends straight segments only. Off: extends curved segments only."
            },
            arcToCircle: {
                ja: "円弧になっている曲線のセグメントから、その円を描きます。",
                en: "Draws the full circle of each curved segment that is a circular arc."
            },
            arcIgnore: { ja: "円弧でない曲線には何も描きません。", en: "Draws nothing for curves that are not circular arcs." },
            arcStraight: {
                ja: "円弧でない曲線には、両端のアンカーポイントを結ぶ線分を描きます。",
                en: "For curves that are not circular arcs, draws the segment between their end anchor points."
            },
            arcExtend: {
                ja: "円弧でない曲線には、両端のアンカーポイントを結ぶ直線を描画範囲いっぱいに延長して描きます。",
                en: "For curves that are not circular arcs, draws the line through their end anchor points across the drawing area."
            },
            separateLayer: {
                ja: "「_construction」レイヤーに描きます。再実行時は、このスクリプトで描いた線を消してから描き直します。",
                en: "Draws on the \"_construction\" layer. On a rerun, lines this script drew earlier are cleared first."
            },
            dedup: { ja: "重なる延長線を1本にまとめます。", en: "Draws only one line where extended lines overlap." }
        },
        itemName: {
            lineGroup: { ja: "補助線", en: "Aux Lines" }
        },
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" }
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noSelection: { ja: "オブジェクトを選択してください。", en: "Please select objects." },
            noValidPath: { ja: "有効なパスが見つかりません。", en: "No valid paths were found." },
            error: { ja: "エラー: ", en: "Error: " }
        }
    };

    /**
     * ドット区切りのパスで表示言語のラベルを取り出す（見つからなければパスをそのまま返す）
     * @param {string} labelPath - LABELS 内のパス（例: "checkbox.group"）
     * @returns {string} 表示言語のラベル
     */
    function getLabel(labelPath) {
        var pathKeys = labelPath.split(".");
        var labelNode = LABELS;
        for (var i = 0; i < pathKeys.length && labelNode; i++) {
            labelNode = labelNode[pathKeys[i]];
        }
        if (!labelNode) return labelPath;
        return labelNode[uiLang] || labelNode.en || labelNode.ja || labelPath;
    }

    /**
     * コロン付きの項目名を返す（日本語は全角、英語は半角）
     * @param {string} labelPath - ラベルのパス
     * @returns {string} コロン付きの項目名
     */
    function labelText(labelPath) {
        return getLabel(labelPath) + (uiLang === "ja" ? "：" : ":");
    }

    // =========================================
    // セッション記憶 / Session memory
    // =========================================

    /* セッション保持（Illustrator終了で破棄） / Session-only state (forgotten when Illustrator quits) */
    var __EXTENDLINES_SESSION__ = (typeof __EXTENDLINES_SESSION__ !== "undefined") ? __EXTENDLINES_SESSION__ : {};

    // =========================================
    // 単位 / Units
    // =========================================

    /* 単位テーブル（配列の添字が rulerType コードと一致：0=in, 1=mm, 2=pt …）/ Unit table; the array index equals the rulerType code */
    var UNITS = [
        { label: "in",    pointsPerUnit: 72 },                /* 0 */
        { label: "mm",    pointsPerUnit: 72 / 25.4 },         /* 1 */
        { label: "pt",    pointsPerUnit: 1 },                 /* 2 */
        { label: "pica",  pointsPerUnit: 12 },                /* 3 */
        { label: "cm",    pointsPerUnit: 72 / 2.54 },         /* 4 */
        { label: "Q",     pointsPerUnit: 72 / 25.4 * 0.25 },  /* 5 */
        { label: "px",    pointsPerUnit: 1 },                 /* 6 */
        { label: "ft/in", pointsPerUnit: 72 * 12 },           /* 7 */
        { label: "m",     pointsPerUnit: 72 / 25.4 * 1000 },  /* 8 */
        { label: "yd",    pointsPerUnit: 72 * 36 },           /* 9 */
        { label: "ft",    pointsPerUnit: 72 * 12 }            /* 10 */
    ];

    /* 単位コード5を「歯（H）」と表示する環境設定キー。文字サイズ（text/units）だけ「級（Q）」
       Preference keys that show unit code 5 as H; only the type size (text/units) shows Q */
    var HA_UNIT_PREF_KEYS = { "rulerType": true, "strokeUnits": true, "text/asianunits": true };

    /**
     * 設定キーごとの単位情報を取得する
     * @param {string} prefKey - 環境設定キー（省略時は "rulerType"）
     * @returns {{code: number, label: string, pointsPerUnit: number}} 単位情報
     */
    function getUnitInfo(prefKey) {
        var unitKey = prefKey || "rulerType";
        var unitCode = app.preferences.getIntegerPreference(unitKey);
        var unit = UNITS[unitCode] || UNITS[2];
        var label = (unitCode === 5 && HA_UNIT_PREF_KEYS[unitKey]) ? "H" : unit.label;
        return { code: unitCode, label: label, pointsPerUnit: unit.pointsPerUnit };
    }

    // =========================================
    // 入力の補助 / Input helpers
    // =========================================

    /**
     * テキストフィールドを↑↓キーで増減できるようにする（±1、Shift で10単位、Option で±0.1）
     * @param {EditText} editText - 対象のテキストフィールド
     * @param {boolean} allowNegative - 負の値を許すか
     * @param {Function} [onChanged] - 値を変えたあとに呼ぶ関数
     * @returns {void}
     */
    function changeValueByArrowKey(editText, allowNegative, onChanged) {
        editText.addEventListener("keydown", function (event) {
            var value = Number(editText.text);
            if (isNaN(value)) return;

            var keyboard = ScriptUI.environment.keyboardState;
            var delta = 1;

            if (keyboard.shiftKey) {
                delta = 10;
                // Shiftキー押下時は10の倍数にスナップ
                if (event.keyName == "Up") {
                    value = Math.ceil((value + 1) / delta) * delta;
                    event.preventDefault();
                } else if (event.keyName == "Down") {
                    value = Math.floor((value - 1) / delta) * delta;
                    event.preventDefault();
                }
            } else if (keyboard.altKey) {
                delta = 0.1;
                // Optionキー押下時は0.1単位で増減
                if (event.keyName == "Up") {
                    value += delta;
                    event.preventDefault();
                } else if (event.keyName == "Down") {
                    value -= delta;
                    event.preventDefault();
                }
            } else {
                if (event.keyName == "Up") {
                    value += delta;
                    event.preventDefault();
                } else if (event.keyName == "Down") {
                    value -= delta;
                    event.preventDefault();
                }
            }

            if (keyboard.altKey) {
                // 小数第1位までに丸め
                value = Math.round(value * 10) / 10;
            } else {
                // 整数に丸め
                value = Math.round(value);
            }

            if (!allowNegative && value < 0) value = 0;

            editText.text = String(value);
            if (typeof onChanged === "function") {
                /* プレビューの描き直しで DOM が失敗してもキー操作は続ける / keep the key handler alive if the preview redraw fails */
                try { onChanged(); } catch (e) { }
            }
        });
    }

    /**
     * 数値の入力文字列を読む。空白を除き、全角のカンマ・小数点を半角にし、
     * 「1,5」のような小数カンマは小数点に、「1,000」のような桁区切りは取り除く
     * @param {string} inputText - 入力された文字列
     * @returns {number} 読み取った数値（読めなければ NaN）
     */
    function parseLocaleNumber(inputText) {
        var normalizedText = String(inputText);
        normalizedText = normalizedText.replace(/\s+/g, "").replace(/，/g, ",").replace(/．/g, ".");
        if (normalizedText.indexOf(",") >= 0 && normalizedText.indexOf(".") < 0) {
            normalizedText = normalizedText.replace(/,/g, ".");
        } else {
            normalizedText = normalizedText.replace(/,/g, "");
        }
        return Number(normalizedText);
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * ドキュメントを確認して補助線の描画を実行する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }

        try {
            runExtendLines(app.activeDocument);
        } catch (e) {
            alert(getLabel("alert.error") + e);
        }
    }

    /**
     * 選択からパスを集め、ダイアログの設定で補助線を描く
     * @param {Document} doc - 対象のドキュメント
     * @returns {void}
     */
    function runExtendLines(doc) {
        var currentSelection = doc.selection;

        if (!currentSelection || currentSelection.length === 0) {
            alert(getLabel("alert.noSelection"));
            return;
        }

        /* 選択からパスを抽出し、テキストは一時的に複製・アウトライン化してから追加
           / Collect paths from the selection; text is duplicated and outlined temporarily, then added */
        var targetPaths = [];
        extractPathItems(currentSelection, targetPaths);
        var tempOutlineRoots = outlineTextFromSelection(currentSelection);
        if (tempOutlineRoots.length > 0) {
            extractPathItems(tempOutlineRoots, targetPaths);
        }

        if (targetPaths.length === 0) {
            alert(getLabel("alert.noValidPath"));
            cleanupTempOutlines(tempOutlineRoots);
            return;
        }

        var drawSettings = showDialog(doc, currentSelection, targetPaths);
        if (drawSettings === null) {
            cleanupTempOutlines(tempOutlineRoots);
            return; /* キャンセル / cancelled */
        }

        var drawBounds = getDrawBounds(doc, currentSelection);
        var targetLayer = prepareTargetLayer(doc, currentSelection, drawSettings.separateLayer);
        var lineContainer = createLineContainer(targetLayer, drawSettings.group);

        try {
            drawConstructionLines(lineContainer, targetPaths, drawSettings, drawBounds);
        } finally {
            /* 一時アウトラインの後始末（失敗時も確実に削除）/ Always remove the temporary outlines */
            cleanupTempOutlines(tempOutlineRoots);
        }
    }

    /**
     * 描画範囲を求める。通常はアクティブなアートボードで、選択がアートボードと
     * まったく交差しないときだけ、選択の中心を中心に幅・高さ4倍の矩形を使う
     * @param {Document} doc - 対象のドキュメント
     * @param {PageItem[]} currentSelection - 選択オブジェクト
     * @returns {{left: number, top: number, right: number, bottom: number}} 描画範囲
     */
    function getDrawBounds(doc, currentSelection) {
        var artboardRect = doc.artboards[doc.artboards.getActiveArtboardIndex()].artboardRect;
        var drawBounds = { left: artboardRect[0], top: artboardRect[1], right: artboardRect[2], bottom: artboardRect[3] };

        var selectionBounds = getUnionBounds(currentSelection);
        if (!selectionBounds) return drawBounds;

        var selLeft = selectionBounds[0], selTop = selectionBounds[1], selRight = selectionBounds[2], selBottom = selectionBounds[3];

        /* Illustrator の座標は top > bottom（Yが上に行くほど増える）/ Illustrator's Y axis points up (top > bottom) */
        var intersects = !(selRight < drawBounds.left || selLeft > drawBounds.right || selTop < drawBounds.bottom || selBottom > drawBounds.top);
        if (intersects) return drawBounds;

        var selWidth = selRight - selLeft;
        var selHeight = selTop - selBottom;
        /* 幅・高さが極端に小さい場合の安全策 / Guard against a degenerate selection */
        if (selWidth < 1) selWidth = 1;
        if (selHeight < 1) selHeight = 1;

        var centerX = (selLeft + selRight) / 2;
        var centerY = (selTop + selBottom) / 2;
        var drawWidth = selWidth * 4;
        var drawHeight = selHeight * 4;

        return {
            left: centerX - drawWidth / 2,
            top: centerY + drawHeight / 2,
            right: centerX + drawWidth / 2,
            bottom: centerY - drawHeight / 2
        };
    }

    /**
     * 描画先のレイヤーを用意する。［別レイヤーに］がオフならアクティブレイヤー。
     * オンなら _construction レイヤーを再利用し（このスクリプトの生成物は消す）、
     * 選択が _construction 上にあるときは既存を _backup に改名して新しく作る
     * @param {Document} doc - 対象のドキュメント
     * @param {PageItem[]} currentSelection - 選択オブジェクト
     * @param {boolean} shouldSeparateLayer - ［別レイヤーに］の値
     * @returns {Layer} 描画先のレイヤー
     */
    function prepareTargetLayer(doc, currentSelection, shouldSeparateLayer) {
        if (!shouldSeparateLayer) return doc.activeLayer;

        var targetLayer;
        if (isSelectionOnLayer(currentSelection, CONSTRUCTION_LAYER_NAME)) {
            /* 元のオブジェクトを消さないため、既存をバックアップ名へ改名して新しく作る / Keep the originals: rename the existing layer and start a new one */
            var existingLayer = findLayerByName(doc, CONSTRUCTION_LAYER_NAME);
            if (existingLayer) {
                var backupName = createUniqueLayerName(doc, CONSTRUCTION_LAYER_NAME + "_backup");
                try { existingLayer.name = backupName; } catch (e) { }
            }
            targetLayer = doc.layers.add();
            targetLayer.name = CONSTRUCTION_LAYER_NAME;
        } else {
            targetLayer = findLayerByName(doc, CONSTRUCTION_LAYER_NAME);
            if (!targetLayer) {
                targetLayer = doc.layers.add();
                targetLayer.name = CONSTRUCTION_LAYER_NAME;
            }
            /* 再実行時は、このスクリプトが生成したものだけクリア / On a rerun, clear only what this script generated */
            clearGeneratedItemsInLayer(targetLayer);
        }

        /* 最前面へ移動 / Bring to front */
        try {
            targetLayer.zOrder(ZOrderMethod.BRINGTOFRONT);
        } catch (e) { }
        return targetLayer;
    }

    /**
     * 補助線の格納先を用意する（［グループ化］がオンなら目印付きのグループ、オフならレイヤーそのもの）
     * @param {Layer} targetLayer - 描画先のレイヤー
     * @param {boolean} shouldGroup - ［グループ化］の値
     * @returns {Layer|GroupItem} 補助線の格納先
     */
    function createLineContainer(targetLayer, shouldGroup) {
        if (!shouldGroup) return targetLayer;
        var lineGroup = targetLayer.groupItems.add();
        lineGroup.name = SCRIPT_MARKER + "_" + getLabel("itemName.lineGroup");
        try { lineGroup.note = SCRIPT_MARKER; } catch (e) { }
        return lineGroup;
    }

    /**
     * 対象パスから補助線（と［円弧から円］の円）を描く。プレビューと確定の両方で使う
     * @param {Layer|GroupItem} targetContainer - 描画先
     * @param {PathItem[]} targetPaths - 対象のパス
     * @param {Object} drawSettings - readDialogSettings() の戻り値と同じ形の設定
     * @param {{left: number, top: number, right: number, bottom: number}} drawBounds - 描画範囲
     * @returns {void}
     */
    function drawConstructionLines(targetContainer, targetPaths, drawSettings, drawBounds) {
        var dedupMap = {}; /* 線のダブり検出用 / for detecting duplicate lines */

        /* 円弧から円（先に）/ Circles from arcs first */
        if (drawSettings.arcToCircle) {
            for (var c = 0; c < targetPaths.length; c++) {
                createCirclesFromArcPath(targetPaths[c], targetContainer, drawSettings, drawBounds, dedupMap);
            }
        }

        for (var p = 0; p < targetPaths.length; p++) {
            var pathItem = targetPaths[p];
            var pathPoints = pathItem.pathPoints;
            if (!pathPoints || pathPoints.length < 2) continue;

            /* 隣接するアンカーポイントのペア（閉じたパスは最後と最初もつなぐ）/ Adjacent anchor pairs (closed paths also join the last and first) */
            var pointPairs = [];
            for (var i = 0; i < pathPoints.length - 1; i++) {
                pointPairs.push([i, i + 1]);
            }
            if (pathItem.closed && pathPoints.length >= 3) {
                pointPairs.push([pathPoints.length - 1, 0]);
            }

            for (var j = 0; j < pointPairs.length; j++) {
                var startPoint = pathPoints[pointPairs[j][0]];
                var endPoint = pathPoints[pointPairs[j][1]];
                var isStraight = isStraightSegment(startPoint, endPoint);

                /* STRAIGHT は直線セグメントだけ、CURVE は曲線セグメントだけ / STRAIGHT keeps straight segments, CURVE keeps curved ones */
                if (drawSettings.mode === "STRAIGHT" && !isStraight) continue;
                if (drawSettings.mode === "CURVE" && isStraight) continue;

                /* 「円弧から円」ON のときは曲線を円の処理に任せる（ここで結ぶと円弧が直線になる）
                   / With "arc to circle" on, curves are left to the circle routine (joining them here would flatten the arc) */
                if (drawSettings.arcToCircle && !isStraight) continue;

                var startAnchor = startPoint.anchor;
                var endAnchor = endPoint.anchor;

                /* 2点が同じ座標（ゴミパスなど）はスキップ / Skip coincident points (stray paths and the like) */
                if (Math.abs(startAnchor[0] - endAnchor[0]) < 0.001 && Math.abs(startAnchor[1] - endAnchor[1]) < 0.001) continue;

                drawLineAcrossArtboard(targetContainer, startAnchor, endAnchor, drawBounds, drawSettings, dedupMap);
            }
        }
    }

    /**
     * 選択オブジェクト全体の外接バウンディング（geometricBounds）を取得する
     * @param {PageItem[]} items - 対象のオブジェクト
     * @returns {number[]|null} [左, 上, 右, 下]（取れなければ null）
     */
    function getUnionBounds(items) {
        var unionBounds = null;
        for (var i = 0; i < items.length; i++) {
            try {
                if (!items[i] || !items[i].geometricBounds) continue;
                var itemBounds = items[i].geometricBounds; /* [L, T, R, B] */
                if (!unionBounds) {
                    unionBounds = [itemBounds[0], itemBounds[1], itemBounds[2], itemBounds[3]];
                } else {
                    if (itemBounds[0] < unionBounds[0]) unionBounds[0] = itemBounds[0];
                    if (itemBounds[1] > unionBounds[1]) unionBounds[1] = itemBounds[1];
                    if (itemBounds[2] > unionBounds[2]) unionBounds[2] = itemBounds[2];
                    if (itemBounds[3] < unionBounds[3]) unionBounds[3] = itemBounds[3];
                }
            } catch (e) { }
        }
        return unionBounds;
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * ダイアログと各コントロールを作る
     * @returns {Object} ダイアログ（dialog）と各コントロールの参照
     */
    function buildDialog() {
        var extendDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        extendDialog.orientation = "column";
        extendDialog.alignChildren = ["left", "top"];
        extendDialog.margins = DIALOG_MARGINS;

        /* 前回の位置を復元（セッション中のみ）/ Restore the last dialog position (session only) */
        try {
            if (__EXTENDLINES_SESSION__.dlgLoc && __EXTENDLINES_SESSION__.dlgLoc.length === 2) {
                extendDialog.location = __EXTENDLINES_SESSION__.dlgLoc;
            }
        } catch (e) { }

        /* 2カラムレイアウト / Two-column layout */
        var columnsGroup = extendDialog.add("group");
        columnsGroup.orientation = "row";
        columnsGroup.alignChildren = ["fill", "top"];
        columnsGroup.alignment = ["fill", "top"];
        columnsGroup.spacing = COLUMN_SPACING;

        /* 左カラム：補助線を描画 / Left column: draw construction lines */
        var auxLinesPanel = columnsGroup.add("panel", undefined, getLabel("panel.auxLines"));
        applyPanelLayout(auxLinesPanel);
        auxLinesPanel.alignment = ["fill", "top"];

        var auxLinesGroup = auxLinesPanel.add("group");
        auxLinesGroup.orientation = "column";
        auxLinesGroup.alignChildren = ["left", "top"];
        auxLinesGroup.margins = [0, 0, 0, 0];

        var straightOnlyRow = auxLinesGroup.add("group");
        straightOnlyRow.orientation = "row";
        straightOnlyRow.alignChildren = ["left", "center"];
        straightOnlyRow.spacing = 0;

        var straightOnlyCheckbox = straightOnlyRow.add("checkbox", undefined, getLabel("checkbox.straightOnly"));
        straightOnlyCheckbox.helpTip = getLabel("tooltip.straightOnly");
        straightOnlyCheckbox.value = true; /* デフォルトON / on by default */

        var arcToCircleCheckbox = auxLinesGroup.add("checkbox", undefined, getLabel("checkbox.arcToCircle"));
        arcToCircleCheckbox.helpTip = getLabel("tooltip.arcToCircle");
        arcToCircleCheckbox.value = true; /* デフォルトON / on by default */

        /* 円弧オプション（［円弧から円］がオフのときはディム）/ Arc options (dimmed while "arc to circle" is off) */
        var arcOptionsPanel = auxLinesGroup.add("panel", undefined, getLabel("panel.arcOptions"));
        applyPanelLayout(arcOptionsPanel);
        arcOptionsPanel.helpTip = getLabel("tooltip.arcOptions");

        var arcFallbackGroup = arcOptionsPanel.add("group");
        arcFallbackGroup.orientation = "column";
        arcFallbackGroup.alignChildren = ["left", "top"];

        var arcIgnoreRadio = arcFallbackGroup.add("radiobutton", undefined, getLabel("radio.arcIgnore"));
        arcIgnoreRadio.helpTip = getLabel("tooltip.arcIgnore");
        var arcStraightRadio = arcFallbackGroup.add("radiobutton", undefined, getLabel("radio.arcStraight"));
        arcStraightRadio.helpTip = getLabel("tooltip.arcStraight");
        var arcExtendRadio = arcFallbackGroup.add("radiobutton", undefined, getLabel("radio.arcExtend"));
        arcExtendRadio.helpTip = getLabel("tooltip.arcExtend");
        arcIgnoreRadio.value = true; /* デフォルト：無視 / default: ignore */

        arcOptionsPanel.enabled = arcToCircleCheckbox.value;

        /* 線幅（線の単位で入力。既定値は 0.1mm を換算）/ Stroke width (entered in the stroke unit; defaults to 0.1 mm converted) */
        var strokeRow = auxLinesGroup.add("group");
        strokeRow.orientation = "row";
        strokeRow.alignChildren = ["left", "center"];
        strokeRow.spacing = STROKE_ROW_SPACING;

        strokeRow.add("statictext", undefined, labelText("fieldLabel.strokeWidth"));

        var strokeUnit = getUnitInfo("strokeUnits");
        var defaultUnitValue = DEFAULT_STROKE_WIDTH_PT / strokeUnit.pointsPerUnit;
        var strokeWidthInput = strokeRow.add("edittext", undefined, defaultUnitValue.toFixed(3));
        strokeWidthInput.characters = STROKE_INPUT_CHARS;

        strokeRow.add("statictext", undefined, strokeUnit.label);

        /* 右カラム：オプション / Right column: options */
        var optionsPanel = columnsGroup.add("panel", undefined, getLabel("panel.options"));
        applyPanelLayout(optionsPanel);
        optionsPanel.alignment = ["fill", "top"];

        var groupCheckbox = optionsPanel.add("checkbox", undefined, getLabel("checkbox.group"));
        groupCheckbox.value = true; /* デフォルトON / on by default */

        var separateLayerCheckbox = optionsPanel.add("checkbox", undefined, getLabel("checkbox.separateLayer"));
        separateLayerCheckbox.helpTip = getLabel("tooltip.separateLayer");
        separateLayerCheckbox.value = true; /* デフォルトON / on by default */

        var guideCheckbox = optionsPanel.add("checkbox", undefined, getLabel("checkbox.guide"));
        guideCheckbox.value = false; /* デフォルトOFF / off by default */

        var dedupCheckbox = optionsPanel.add("checkbox", undefined, getLabel("checkbox.dedup"));
        dedupCheckbox.helpTip = getLabel("tooltip.dedup");
        dedupCheckbox.value = true; /* デフォルトON / on by default */

        /* 2カラムの下に余白 / Spacer below the two columns */
        extendDialog.add("panel", undefined, undefined);

        /* ボタンエリア（左：プレビュー／中央：スペーサー／右：キャンセル・OK）/ Button row: preview, spacer, Cancel/OK */
        var btnRowGroup = extendDialog.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.alignChildren = ["fill", "center"];
        btnRowGroup.alignment = ["fill", "top"];

        var previewCheckbox = btnRowGroup.add("checkbox", undefined, getLabel("checkbox.preview"));
        previewCheckbox.value = false;

        var spacer = btnRowGroup.add("group");
        spacer.alignment = ["fill", "fill"];
        spacer.minimumSize.width = 0;

        var btnRightGroup = btnRowGroup.add("group");
        btnRightGroup.orientation = "row";
        btnRightGroup.alignChildren = ["right", "center"];
        btnRightGroup.spacing = BUTTON_SPACING;

        var btnCancel = btnRightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnOK = btnRightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });

        return {
            dialog: extendDialog,
            straightOnlyCheckbox: straightOnlyCheckbox,
            arcToCircleCheckbox: arcToCircleCheckbox,
            arcOptionsPanel: arcOptionsPanel,
            arcIgnoreRadio: arcIgnoreRadio,
            arcStraightRadio: arcStraightRadio,
            arcExtendRadio: arcExtendRadio,
            strokeWidthInput: strokeWidthInput,
            groupCheckbox: groupCheckbox,
            separateLayerCheckbox: separateLayerCheckbox,
            guideCheckbox: guideCheckbox,
            dedupCheckbox: dedupCheckbox,
            previewCheckbox: previewCheckbox
        };
    }

    /**
     * ダイアログを表示して設定を取得する（プレビューはダイアログ中だけ描き、閉じたら消す）
     * @param {Document} doc - 対象のドキュメント
     * @param {PageItem[]} currentSelection - 選択オブジェクト
     * @param {PathItem[]} targetPaths - 対象のパス
     * @returns {Object|null} 設定（キャンセル時は null）
     */
    function showDialog(doc, currentSelection, targetPaths) {
        var dialogUI = buildDialog();

        /* 線幅の最後の有効値（入力途中で NaN や 0 になったときに使う）/ Last valid stroke width, used while the field is mid-edit */
        var lastValidStrokeWidthPt = DEFAULT_STROKE_WIDTH_PT;

        /* プレビュー用レイヤー名（ダイアログごとに一意）/ Preview layer name (unique per dialog) */
        var previewLayerName = createUniqueLayerName(doc, PREVIEW_LAYER_BASE_NAME);

        /**
         * 線幅の入力を pt に換算する（読めないときは最後の有効値）
         * @returns {number} 線幅（pt）
         */
        function readStrokeWidthPt() {
            var pointsPerUnit = getUnitInfo("strokeUnits").pointsPerUnit;
            var enteredValue = parseLocaleNumber(dialogUI.strokeWidthInput.text);
            if (!(enteredValue > 0)) return lastValidStrokeWidthPt;

            var strokeWidthPt = enteredValue * pointsPerUnit;
            if (!(strokeWidthPt > 0)) return lastValidStrokeWidthPt;

            lastValidStrokeWidthPt = strokeWidthPt;
            return strokeWidthPt;
        }

        /**
         * ダイアログの状態を設定オブジェクトにまとめる
         * @returns {Object} mode / group / separateLayer / guide / dedup / arcToCircle / strokeWidthPt / arcFallback
         */
        function readDialogSettings() {
            return {
                mode: dialogUI.straightOnlyCheckbox.value ? "STRAIGHT" : "CURVE",
                group: dialogUI.groupCheckbox.value,
                separateLayer: dialogUI.separateLayerCheckbox.value,
                guide: dialogUI.guideCheckbox.value,
                dedup: dialogUI.dedupCheckbox.value,
                arcToCircle: dialogUI.arcToCircleCheckbox.value,
                strokeWidthPt: readStrokeWidthPt(),
                arcFallback: dialogUI.arcIgnoreRadio.value ? "IGNORE" : (dialogUI.arcExtendRadio.value ? "EXTEND" : "STRAIGHT")
            };
        }

        /**
         * プレビューを消す
         * @returns {void}
         */
        function clearPreview() {
            removeLayerIfExists(doc, previewLayerName);
            app.redraw();
        }

        /**
         * プレビューを描き直す（既存のプレビューは消して作り直す。ユーザーの元オブジェクトは触らない）
         * @returns {void}
         */
        function updatePreviewFromUI() {
            if (!dialogUI.previewCheckbox.value) return;

            removeLayerIfExists(doc, previewLayerName);

            var previewLayer = doc.layers.add();
            previewLayer.name = previewLayerName;
            try { previewLayer.zOrder(ZOrderMethod.BRINGTOFRONT); } catch (e) { }

            var previewGroup = previewLayer.groupItems.add();
            previewGroup.name = SCRIPT_MARKER + "_PREVIEW";
            try { previewGroup.note = SCRIPT_MARKER; } catch (e) { }

            drawConstructionLines(previewGroup, targetPaths, readDialogSettings(), getDrawBounds(doc, currentSelection));
            app.redraw();
        }

        /**
         * プレビューがオンなら描き直す
         * @returns {void}
         */
        function refreshPreviewIfOn() {
            if (dialogUI.previewCheckbox.value) updatePreviewFromUI();
        }

        /**
         * 線幅が変わったとき、最後の有効値を更新してプレビューへ反映する
         * @returns {void}
         */
        function onStrokeWidthChanged() {
            readStrokeWidthPt();
            refreshPreviewIfOn();
        }

        /**
         * 円弧オプションのラジオを排他で選ぶ（環境差で同時にオンになる事故を防ぐ）
         * @param {string} fallbackMode - "IGNORE" / "STRAIGHT" / "EXTEND"
         * @returns {void}
         */
        function setArcFallback(fallbackMode) {
            dialogUI.arcIgnoreRadio.value = (fallbackMode === "IGNORE");
            dialogUI.arcStraightRadio.value = (fallbackMode === "STRAIGHT");
            dialogUI.arcExtendRadio.value = (fallbackMode === "EXTEND");
        }

        /**
         * 円弧オプションのラジオにクリック処理を付ける
         * @param {RadioButton} fallbackRadio - 対象のラジオ
         * @param {string} fallbackMode - そのラジオが表すモード
         * @returns {void}
         */
        function bindArcFallbackRadio(fallbackRadio, fallbackMode) {
            fallbackRadio.onClick = function () {
                setArcFallback(fallbackMode);
                refreshPreviewIfOn();
            };
        }

        dialogUI.previewCheckbox.onClick = function () {
            if (dialogUI.previewCheckbox.value) updatePreviewFromUI();
            else clearPreview();
        };

        dialogUI.arcToCircleCheckbox.onClick = function () {
            dialogUI.arcOptionsPanel.enabled = dialogUI.arcToCircleCheckbox.value;
            refreshPreviewIfOn();
        };

        var previewTriggerCheckboxes = [dialogUI.straightOnlyCheckbox, dialogUI.groupCheckbox, dialogUI.separateLayerCheckbox, dialogUI.guideCheckbox, dialogUI.dedupCheckbox];
        for (var i = 0; i < previewTriggerCheckboxes.length; i++) {
            previewTriggerCheckboxes[i].onClick = refreshPreviewIfOn;
        }

        bindArcFallbackRadio(dialogUI.arcIgnoreRadio, "IGNORE");
        bindArcFallbackRadio(dialogUI.arcStraightRadio, "STRAIGHT");
        bindArcFallbackRadio(dialogUI.arcExtendRadio, "EXTEND");

        /* 線幅の変更をプレビューに反映し、常に最後の有効値も更新 / Reflect stroke-width edits in the preview and keep the last valid value */
        dialogUI.strokeWidthInput.onChanging = onStrokeWidthChanged;

        /* ↑↓ / Shift+↑↓ / Option+↑↓ での増減 / Arrow-key stepping */
        changeValueByArrowKey(dialogUI.strokeWidthInput, false, onStrokeWidthChanged);

        var dialogResult = dialogUI.dialog.show();

        /* ダイアログ終了時は必ずプレビューを消す / Always remove the preview when the dialog closes */
        clearPreview();

        /* 位置を控える（セッション中のみ）/ Save the dialog position (session only) */
        try {
            __EXTENDLINES_SESSION__.dlgLoc = [dialogUI.dialog.location[0], dialogUI.dialog.location[1]];
        } catch (e) { }

        return (dialogResult === 1) ? readDialogSettings() : null;
    }

    // =========================================
    // 対象の収集 / Target collection
    // =========================================

    /**
     * 選択内のテキストを一時的に複製してアウトライン化し、アウトライン（グループ等）を返す。
     * 元のテキストは変更しない
     * @param {PageItem[]} items - 選択オブジェクト
     * @returns {PageItem[]} 一時アウトライン（テキストが無ければ空配列）
     */
    function outlineTextFromSelection(items) {
        var outlineRoots = [];

        /**
         * テキストを探してアウトライン化する（グループ・複合パスの中も辿る）
         * @param {PageItem} pageItem - 対象のオブジェクト
         * @returns {void}
         */
        function walk(pageItem) {
            if (!pageItem) return;

            if (pageItem.typename === "TextFrame") {
                try {
                    /* 同一レイヤー末尾に複製し、複製だけアウトライン化 / Duplicate to the end of the same layer and outline only the copy */
                    var duplicatedText = pageItem.duplicate(pageItem.layer, ElementPlacement.PLACEATEND);
                    var outlined = duplicatedText.createOutline();
                    /* createOutline は複製を消費するので remove は例外になりうる / createOutline consumes the copy, so remove() may throw */
                    try { duplicatedText.remove(); } catch (e) { }

                    if (outlined) {
                        outlineRoots.push(outlined);
                    }
                } catch (e) {
                    /* 失敗しても全体は止めない / Keep going on failure */
                }
                return;
            }

            if (pageItem.typename === "GroupItem") {
                for (var i = 0; i < pageItem.pageItems.length; i++) {
                    walk(pageItem.pageItems[i]);
                }
            } else if (pageItem.typename === "CompoundPathItem") {
                for (var j = 0; j < pageItem.pathItems.length; j++) {
                    walk(pageItem.pathItems[j]);
                }
            }
        }

        for (var i = 0; i < items.length; i++) {
            walk(items[i]);
        }

        return outlineRoots;
    }

    /**
     * 一時アウトライン（グループ等）を削除する
     * @param {PageItem[]} outlineRoots - outlineTextFromSelection() の戻り値
     * @returns {void}
     */
    function cleanupTempOutlines(outlineRoots) {
        if (!outlineRoots || outlineRoots.length === 0) return;
        for (var i = outlineRoots.length - 1; i >= 0; i--) {
            try {
                if (outlineRoots[i] && outlineRoots[i].remove) outlineRoots[i].remove();
            } catch (e) { }
        }
    }

    /**
     * グループや複合パスの中から再帰的にパスアイテムを抽出する
     * @param {PageItem[]} items - 対象のオブジェクト
     * @param {PathItem[]} targetArray - 見つけたパスを追加する配列
     * @returns {void}
     */
    function extractPathItems(items, targetArray) {
        for (var i = 0; i < items.length; i++) {
            var pageItem = items[i];
            if (pageItem.typename === "PathItem") {
                targetArray.push(pageItem);
            } else if (pageItem.typename === "CompoundPathItem") {
                extractPathItems(pageItem.pathItems, targetArray);
            } else if (pageItem.typename === "GroupItem") {
                extractPathItems(pageItem.pageItems, targetArray);
            }
        }
    }

    // =========================================
    // レイヤー / Layers
    // =========================================

    /**
     * 選択がすべて指定レイヤー上にあるかを返す
     * @param {PageItem[]} items - 選択オブジェクト
     * @param {string} layerName - レイヤー名
     * @returns {boolean} すべてそのレイヤーに属していれば true
     */
    function isSelectionOnLayer(items, layerName) {
        try {
            if (!items || items.length === 0) return false;
            for (var i = 0; i < items.length; i++) {
                if (!items[i]) return false;
                var itemLayer = items[i].layer;
                if (!itemLayer || itemLayer.name !== layerName) return false;
            }
            return true;
        } catch (e) { }
        return false;
    }

    /**
     * 名前でレイヤーを探す
     * @param {Document} doc - 対象のドキュメント
     * @param {string} layerName - レイヤー名
     * @returns {Layer|null} 見つかったレイヤー（無ければ null）
     */
    function findLayerByName(doc, layerName) {
        /* getByName は見つからないと例外 / getByName throws when missing */
        try { return doc.layers.getByName(layerName); } catch (e) { }
        return null;
    }

    /**
     * 既存と重ならないレイヤー名を作る（重なれば _2, _3 … を付ける）
     * @param {Document} doc - 対象のドキュメント
     * @param {string} baseName - もとの名前
     * @returns {string} 使われていないレイヤー名
     */
    function createUniqueLayerName(doc, baseName) {
        var layerName = baseName;
        var suffixNumber = 2;
        while (findLayerByName(doc, layerName)) {
            layerName = baseName + "_" + suffixNumber;
            suffixNumber++;
        }
        return layerName;
    }

    /**
     * 指定した名前のレイヤーがあれば削除する
     * @param {Document} doc - 対象のドキュメント
     * @param {string} layerName - レイヤー名
     * @returns {void}
     */
    function removeLayerIfExists(doc, layerName) {
        try {
            var targetLayer = findLayerByName(doc, layerName);
            if (targetLayer) targetLayer.remove();
        } catch (e) { }
    }

    /**
     * このスクリプトが生成したオブジェクトだけを削除する（ユーザーの既存オブジェクトは残す）
     * @param {Layer} targetLayer - 対象のレイヤー
     * @returns {void}
     */
    function clearGeneratedItemsInLayer(targetLayer) {
        if (!targetLayer) return;
        try {
            /* groupItems を後ろから走査 / Walk the groups from the back */
            for (var i = targetLayer.groupItems.length - 1; i >= 0; i--) {
                var groupItem = targetLayer.groupItems[i];
                if (!groupItem) continue;
                var groupNote = "";
                try { groupNote = groupItem.note || ""; } catch (e) { }
                if (groupNote === SCRIPT_MARKER) {
                    try { groupItem.remove(); } catch (e) { }
                    continue;
                }
                var groupName = "";
                try { groupName = groupItem.name || ""; } catch (e) { }
                if (groupName.indexOf(SCRIPT_MARKER) === 0) {
                    try { groupItem.remove(); } catch (e) { }
                }
            }

            /* レイヤー直下に描いた線もあるので pathItems も後ろから走査 / Lines may sit directly on the layer, so walk the paths too */
            for (var j = targetLayer.pathItems.length - 1; j >= 0; j--) {
                var pathItem = targetLayer.pathItems[j];
                if (!pathItem) continue;
                var pathNote = "";
                try { pathNote = pathItem.note || ""; } catch (e) { }
                if (pathNote === SCRIPT_MARKER) {
                    try { pathItem.remove(); } catch (e) { }
                }
            }
        } catch (e) { }
    }

    // =========================================
    // セグメントの判定 / Segment helpers
    // =========================================

    /**
     * 2つのアンカーポイント間が直線（ハンドルが出ていない）かどうかを判定する
     * @param {PathPoint} startPoint - 始点
     * @param {PathPoint} endPoint - 終点
     * @returns {boolean} 直線なら true
     */
    function isStraightSegment(startPoint, endPoint) {
        var epsilon = 0.001;

        /* 始点の出力ハンドル・終点の入力ハンドルがアンカーと同じ位置か / Is each handle on its anchor? */
        var rightDx = Math.abs(startPoint.rightDirection[0] - startPoint.anchor[0]);
        var rightDy = Math.abs(startPoint.rightDirection[1] - startPoint.anchor[1]);
        var leftDx = Math.abs(endPoint.leftDirection[0] - endPoint.anchor[0]);
        var leftDy = Math.abs(endPoint.leftDirection[1] - endPoint.anchor[1]);

        return (rightDx < epsilon && rightDy < epsilon && leftDx < epsilon && leftDy < epsilon);
    }

    /**
     * 2点から線の重複判定キーを作る（向きは無視）
     * @param {number[]} pointA - 端点
     * @param {number[]} pointB - もう一方の端点
     * @returns {string} 重複判定キー
     */
    function makeLineKey(pointA, pointB) {
        var keyA = pointKey(pointA);
        var keyB = pointKey(pointB);
        return (keyA < keyB) ? (keyA + "|" + keyB) : (keyB + "|" + keyA);
    }

    /**
     * 点を 0.001pt 単位に丸めたキーにする（浮動小数の揺れを抑える）
     * @param {number[]} point - [x, y]
     * @returns {string} 点のキー
     */
    function pointKey(point) {
        return (Number(point[0]).toFixed(3) + "," + Number(point[1]).toFixed(3));
    }

    /**
     * 曲線になっている隣接セグメント（どちらかのハンドルが出ているもの）を集める
     * @param {PathItem} pathItem - 対象のパス
     * @returns {Object[]} セグメントの配列（startPoint / endPoint）
     */
    function getCurvedAdjacentSegments(pathItem) {
        var pathPoints = pathItem.pathPoints;
        var pointCount = pathPoints.length;
        var tolerance = 1e-6;
        var isClosed = pathItem.closed;
        var curvedSegments = [];

        for (var i = 0; i < pointCount; i++) {
            if (!isClosed && i === pointCount - 1) break;

            var startPoint = pathPoints[i];
            var endPoint = pathPoints[(i + 1) % pointCount];

            var startRight = startPoint.rightDirection;
            var startAnchor = startPoint.anchor;
            var endLeft = endPoint.leftDirection;
            var endAnchor = endPoint.anchor;

            var startHandleOut = Math.abs(startRight[0] - startAnchor[0]) > tolerance || Math.abs(startRight[1] - startAnchor[1]) > tolerance;
            var endHandleOut = Math.abs(endLeft[0] - endAnchor[0]) > tolerance || Math.abs(endLeft[1] - endAnchor[1]) > tolerance;

            if (startHandleOut || endHandleOut) {
                curvedSegments.push({ startPoint: startPoint, endPoint: endPoint });
            }
        }

        return curvedSegments;
    }

    // =========================================
    // 円弧・円の計算 / Arc and circle math
    // =========================================

    /**
     * 2アンカーの円弧から円の中心を求める（両端の接線に垂直な直線の交点）
     * @param {number[]} startAnchor - 始点のアンカー
     * @param {number[]} startHandle - 始点の出力ハンドル
     * @param {number[]} endAnchor - 終点のアンカー
     * @param {number[]} endHandle - 終点の入力ハンドル
     * @returns {number[]|null} 中心 [x, y]（平行で求まらなければ null）
     */
    function getCircleCenterFrom2AnchorArc(startAnchor, startHandle, endAnchor, endHandle) {
        var startTangent = [startHandle[0] - startAnchor[0], startHandle[1] - startAnchor[1]];
        var endTangent = [endAnchor[0] - endHandle[0], endAnchor[1] - endHandle[1]];

        /* 法線（半径方向）/ Normals (radius directions) */
        return intersectLines(startAnchor, rotate90(startTangent), endAnchor, rotate90(endTangent));
    }

    /**
     * ベクトルを90°回す
     * @param {number[]} vector - [x, y]
     * @returns {number[]} 回したベクトル
     */
    function rotate90(vector) {
        return [-vector[1], vector[0]];
    }

    /**
     * 点と方向で表した2直線の交点を求める
     * @param {number[]} pointA - 直線Aの通る点
     * @param {number[]} directionA - 直線Aの方向
     * @param {number[]} pointB - 直線Bの通る点
     * @param {number[]} directionB - 直線Bの方向
     * @returns {number[]|null} 交点（平行なら null）
     */
    function intersectLines(pointA, directionA, pointB, directionB) {
        var denominator = directionA[0] * directionB[1] - directionA[1] * directionB[0];
        if (Math.abs(denominator) < 1e-9) return null;

        var dx = pointB[0] - pointA[0];
        var dy = pointB[1] - pointA[1];

        var distanceA = (dx * directionB[1] - dy * directionB[0]) / denominator;
        return [pointA[0] + directionA[0] * distanceA, pointA[1] + directionA[1] * distanceA];
    }

    /**
     * 2次元ベクトルの内積
     * @param {number[]} vectorA - [x, y]
     * @param {number[]} vectorB - [x, y]
     * @returns {number} 内積
     */
    function dotProduct(vectorA, vectorB) {
        return vectorA[0] * vectorB[0] + vectorA[1] * vectorB[1];
    }

    /**
     * 2次元ベクトルの長さ
     * @param {number[]} vector - [x, y]
     * @returns {number} 長さ
     */
    function vectorLength(vector) {
        return Math.sqrt(vector[0] * vector[0] + vector[1] * vector[1]);
    }

    /**
     * 2次元ベクトルの差（a − b）
     * @param {number[]} vectorA - [x, y]
     * @param {number[]} vectorB - [x, y]
     * @returns {number[]} 差のベクトル
     */
    function subtractVectors(vectorA, vectorB) {
        return [vectorA[0] - vectorB[0], vectorA[1] - vectorB[1]];
    }

    /**
     * 3次ベジェ曲線上の点を求める
     * @param {number[]} p0 - 始点
     * @param {number[]} p1 - 始点側の制御点
     * @param {number[]} p2 - 終点側の制御点
     * @param {number[]} p3 - 終点
     * @param {number} t - パラメーター（0〜1）
     * @returns {number[]} 曲線上の点
     */
    function cubicBezierPoint(p0, p1, p2, p3, t) {
        var u = 1 - t;
        var uu = u * u;
        var uuu = uu * u;
        var tt = t * t;
        var ttt = tt * t;

        return [
            uuu * p0[0] + 3 * uu * t * p1[0] + 3 * u * tt * p2[0] + ttt * p3[0],
            uuu * p0[1] + 3 * uu * t * p1[1] + 3 * u * tt * p2[1] + ttt * p3[1]
        ];
    }

    /**
     * 円弧とみなせるかを判定する（両端と中点が同じ円上にあり、両端の接線が半径に垂直）
     * @param {number[]} startAnchor - 始点のアンカー
     * @param {number[]} startHandle - 始点の出力ハンドル
     * @param {number[]} endHandle - 終点の入力ハンドル
     * @param {number[]} endAnchor - 終点のアンカー
     * @param {number[]} center - 推定した中心
     * @param {number} radius - 推定した半径
     * @returns {boolean} 円弧とみなせれば true
     */
    function isApproxCircularArc(startAnchor, startHandle, endHandle, endAnchor, center, radius) {
        if (!center || !(radius > 0)) return false;

        /* 半径の1%（最低 0.2pt）を許容 / Tolerance: 1% of the radius, at least 0.2 pt */
        var tolerance = Math.max(0.2, radius * 0.01);

        var startDistance = vectorLength(subtractVectors(startAnchor, center));
        var endDistance = vectorLength(subtractVectors(endAnchor, center));
        if (Math.abs(startDistance - radius) > tolerance) return false;
        if (Math.abs(endDistance - radius) > tolerance) return false;

        var midPoint = cubicBezierPoint(startAnchor, startHandle, endHandle, endAnchor, 0.5);
        var midDistance = vectorLength(subtractVectors(midPoint, center));
        if (Math.abs(midDistance - radius) > tolerance) return false;

        /* 両端で「半径・接線 ≈ 0」/ radius · tangent ≈ 0 at both ends */
        var startTangent = subtractVectors(startHandle, startAnchor);
        var endTangent = subtractVectors(endAnchor, endHandle);
        var startRadius = subtractVectors(startAnchor, center);
        var endRadius = subtractVectors(endAnchor, center);

        /* 接線が短すぎるものは円弧ではないとみなす / Treat vanishing tangents as not an arc */
        if (vectorLength(startTangent) < 1e-6 || vectorLength(endTangent) < 1e-6) return false;

        if (Math.abs(dotProduct(startRadius, startTangent)) > tolerance * vectorLength(startTangent)) return false;
        if (Math.abs(dotProduct(endRadius, endTangent)) > tolerance * vectorLength(endTangent)) return false;

        return true;
    }

    // =========================================
    // 描画 / Drawing
    // =========================================

    /**
     * 補助線（延長線）と同じスタイルを適用する
     * @param {PathItem} pathItem - 対象のパス
     * @param {boolean} shouldGuide - ガイドにするか
     * @param {number} strokeWidthPt - 線幅（pt）
     * @returns {void}
     */
    function applyAuxStyle(pathItem, shouldGuide, strokeWidthPt) {
        try { pathItem.note = SCRIPT_MARKER; } catch (e) { }
        pathItem.filled = false;
        try { pathItem.fillColor = new NoColor(); } catch (e) { }
        if (shouldGuide) {
            pathItem.stroked = false;
            try { pathItem.guides = true; } catch (e) { }
        } else {
            pathItem.stroked = true;

            var blackColor = new CMYKColor();
            blackColor.cyan = 0;
            blackColor.magenta = 0;
            blackColor.yellow = 0;
            blackColor.black = 100;
            pathItem.strokeColor = blackColor;

            pathItem.strokeWidth = (strokeWidthPt && strokeWidthPt > 0) ? strokeWidthPt : DEFAULT_STROKE_WIDTH_PT;
        }
    }

    /**
     * アンカー同士を結ぶ線分（弦）を描く
     * @param {Layer|GroupItem} targetContainer - 描画先
     * @param {number[]} startAnchor - 始点
     * @param {number[]} endAnchor - 終点
     * @param {boolean} shouldGuide - ガイドにするか
     * @param {number} strokeWidthPt - 線幅（pt）
     * @returns {PathItem|null} 描いた線（失敗時は null）
     */
    function createChordLine(targetContainer, startAnchor, endAnchor, shouldGuide, strokeWidthPt) {
        try {
            var chordLine = targetContainer.pathItems.add();
            chordLine.setEntirePath([startAnchor, endAnchor]);
            chordLine.closed = false;
            applyAuxStyle(chordLine, shouldGuide, strokeWidthPt);
            return chordLine;
        } catch (e) { }
        return null;
    }

    /**
     * 曲線セグメントごとに円弧から円を推定して描く。円弧とみなせないものは円弧オプションに従う
     * @param {PathItem} arcPath - 対象のパス
     * @param {Layer|GroupItem} targetContainer - 描画先
     * @param {Object} drawSettings - 設定（guide / arcFallback / dedup / strokeWidthPt）
     * @param {{left: number, top: number, right: number, bottom: number}} drawBounds - 描画範囲
     * @param {Object} dedupMap - 線のダブり検出用の表
     * @returns {PathItem[]} 描いた円
     */
    function createCirclesFromArcPath(arcPath, targetContainer, drawSettings, drawBounds, dedupMap) {
        var circles = [];
        try {
            var curvedSegments = getCurvedAdjacentSegments(arcPath);
            if (!curvedSegments || curvedSegments.length === 0) return circles;

            for (var i = 0; i < curvedSegments.length; i++) {
                var startAnchor = curvedSegments[i].startPoint.anchor;
                var endAnchor = curvedSegments[i].endPoint.anchor;
                var startHandle = curvedSegments[i].startPoint.rightDirection;
                var endHandle = curvedSegments[i].endPoint.leftDirection;

                var center = getCircleCenterFrom2AnchorArc(startAnchor, startHandle, endAnchor, endHandle);
                if (!center) continue;

                var dx = startAnchor[0] - center[0];
                var dy = startAnchor[1] - center[1];
                var radius = Math.sqrt(dx * dx + dy * dy);
                if (!(radius > 0)) continue;

                /* 正確な円弧の一部でない場合の扱い / Segments that are not true arcs */
                if (!isApproxCircularArc(startAnchor, startHandle, endHandle, endAnchor, center, radius)) {
                    if (drawSettings.arcFallback === "STRAIGHT") {
                        /* 直線：アンカー同士を結ぶ弦 / Straight: the chord between the anchors */
                        createChordLine(targetContainer, startAnchor, endAnchor, drawSettings.guide, drawSettings.strokeWidthPt);
                    } else if (drawSettings.arcFallback === "EXTEND") {
                        /* 直線（延長）：弦を描画範囲いっぱいに延長 / Straight (extend): the chord extended across the drawing area */
                        drawLineAcrossArtboard(targetContainer, startAnchor, endAnchor, drawBounds, drawSettings, dedupMap);
                    }
                    /* IGNORE は何もしない / IGNORE draws nothing */
                    continue;
                }

                var circle = targetContainer.pathItems.ellipse(
                    center[1] + radius,
                    center[0] - radius,
                    radius * 2,
                    radius * 2
                );
                /* 塗りなしは applyAuxStyle が設定する / applyAuxStyle clears the fill */
                applyAuxStyle(circle, drawSettings.guide, drawSettings.strokeWidthPt);
                circles.push(circle);
            }
        } catch (e) { }

        return circles;
    }

    /**
     * 2点を通る直線を描画範囲の端まで延長して描く（［線のダブりを削除］がオンなら同じ線は1本だけ）
     * @param {Layer|GroupItem} targetContainer - 描画先
     * @param {number[]} startAnchor - 通る点
     * @param {number[]} endAnchor - もう一つの通る点
     * @param {{left: number, top: number, right: number, bottom: number}} drawBounds - 描画範囲
     * @param {Object} drawSettings - 設定（guide / dedup / strokeWidthPt）
     * @param {Object} dedupMap - 線のダブり検出用の表
     * @returns {void}
     */
    function drawLineAcrossArtboard(targetContainer, startAnchor, endAnchor, drawBounds, drawSettings, dedupMap) {
        var left = drawBounds.left, top = drawBounds.top, right = drawBounds.right, bottom = drawBounds.bottom;
        var x1 = startAnchor[0], y1 = startAnchor[1];
        var x2 = endAnchor[0], y2 = endAnchor[1];

        var intersections = [];
        var epsilon = 0.001;

        if (Math.abs(x1 - x2) < epsilon) {
            /* 垂直線 / vertical */
            intersections.push([x1, top]);
            intersections.push([x1, bottom]);
        } else if (Math.abs(y1 - y2) < epsilon) {
            /* 水平線 / horizontal */
            intersections.push([left, y1]);
            intersections.push([right, y1]);
        } else {
            /* 傾きがある場合は4辺との交点を調べる / Sloped: test all four edges */
            var slope = (y2 - y1) / (x2 - x1);
            var intercept = y1 - slope * x1;

            var yAtLeft = slope * left + intercept;
            if (yAtLeft <= top + epsilon && yAtLeft >= bottom - epsilon) intersections.push([left, yAtLeft]);

            var yAtRight = slope * right + intercept;
            if (yAtRight <= top + epsilon && yAtRight >= bottom - epsilon) intersections.push([right, yAtRight]);

            var xAtTop = (top - intercept) / slope;
            if (xAtTop >= left - epsilon && xAtTop <= right + epsilon) intersections.push([xAtTop, top]);

            var xAtBottom = (bottom - intercept) / slope;
            if (xAtBottom >= left - epsilon && xAtBottom <= right + epsilon) intersections.push([xAtBottom, bottom]);
        }

        if (intersections.length < 2) return;

        /* 最初の交点と重ならない交点を1つ選ぶ / Pick the first intersection that differs from the first one */
        var uniquePoints = [intersections[0]];
        for (var j = 1; j < intersections.length; j++) {
            var dx = intersections[j][0] - uniquePoints[0][0];
            var dy = intersections[j][1] - uniquePoints[0][1];
            if (Math.sqrt(dx * dx + dy * dy) > epsilon) {
                uniquePoints.push(intersections[j]);
                break;
            }
        }
        if (uniquePoints.length !== 2) return;

        /* 線のダブりを削除（同じ2点で構成される延長線は1本だけ描く）/ Draw each extended line only once */
        if (drawSettings.dedup && dedupMap) {
            var lineKey = makeLineKey(uniquePoints[0], uniquePoints[1]);
            if (dedupMap[lineKey]) return;
            dedupMap[lineKey] = true;
        }
        var newLine = targetContainer.pathItems.add();
        newLine.setEntirePath(uniquePoints);
        newLine.closed = false;
        applyAuxStyle(newLine, drawSettings.guide, drawSettings.strokeWidthPt);
    }

    main();

})();
