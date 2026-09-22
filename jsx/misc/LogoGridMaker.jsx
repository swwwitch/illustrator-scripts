#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択オブジェクトの visibleBounds を基準に、ロゴ用の補助線およびクリアスペースを生成します。
「文字形状からグリッド構造を抽出する」ことを目的としたスクリプトです。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/LogoGridMaker.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n95a285784495

### Overview

Generates construction lines and clear space for a logo, based on the visibleBounds of the selection.
The script is aimed at extracting a grid structure from letterforms.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/LogoGridMaker.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "LogoGridMaker";                /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.4.3";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-04-10";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-22";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/LogoGridMaker.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/LogoGridMaker.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n95a285784495"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* 補助線の外観 / Appearance of the guide lines */
    var LINE_STYLE = {
        grayTint: 100,          /* 線のグレー濃度（K%） / gray tint of the lines */
        boundsMultiplier: 4     /* ［境界線を強調］時の線幅倍率 / stroke multiplier for emphasized bounds */
    };

    /* クリアスペースの外観 / Appearance of the clear space */
    var CLEAR_SPACE_STYLE = {
        cyan: 60,               /* 帯のシアン濃度（C%） / cyan tint of the bands */
        fillOpacity: 20,        /* 塗りの不透明度（%） / fill opacity */
        strokeWidth: 0.25       /* 枠線の線幅（pt） / stroke width of the bands */
    };

    /* ライン検出のしきい値 / Thresholds of the line detection */
    var DETECTION = {
        coordTolerance: 0.001,       /* 同じ座標とみなす許容値（pt） / tolerance for identical coordinates */
        clusterRatio: 0.03,          /* クラスタ許容値＝対象サイズ×この比率 / cluster tolerance ratio */
        clusterMin: 1.0,             /* クラスタ許容値の下限（pt） / minimum cluster tolerance */
        clusterMax: 5.0,             /* クラスタ許容値の上限（pt） / maximum cluster tolerance */
        shortSegmentRatio: 0.03,     /* 短いセグメントを捨てる比率（平均長比） / ratio used to drop short segments */
        shortSegmentFloor: 2,        /* 短いセグメント判定の下限（pt） / floor of the short segment threshold */
        horizontalFlatness: 0.2,     /* 水平とみなす上下の差（pt） / vertical gap that still counts as horizontal */
        horizontalMinRun: 0.5,       /* 水平とみなす最小の長さ（pt） / minimum run that counts as horizontal */
        fallbackLineRatio: 0.3,      /* 内部の線を推定できないときの位置比率 / ratio used when inner lines are unknown */
        descenderRatio: 0.15,        /* ディセンダーとみなす下方向の比率 / ratio that still counts as a descender */
        diagonalAngleTolerance: 2.0, /* 同じ傾きとみなす角度差（度） / angle gap that counts as the same slope */
        diagonalMinAngle: 45,        /* 斜線として扱う角度の下限（度） / lower angle bound of diagonal elements */
        diagonalMaxAngle: 90,        /* 斜線として扱う角度の上限（度） / upper angle bound of diagonal elements */
        diagonalMergeRatio: 0.6      /* 斜線をまとめる間隔の比率 / ratio used to merge diagonal lines */
    };

    // =========================================
    // レイアウト / Layout
    // =========================================

    var PANEL_MARGINS = [15, 20, 15, 10];   /* パネル余白 [左,上,右,下] / panel margins */

    /* 入力欄の幅（文字数） / Width of the input fields, in characters */
    var FIELD_WIDTH = {
        count: 2,        /* 本数・分割数 / counts and divisions */
        strokeWidth: 4,  /* 線幅 / stroke width */
        scale: 6,        /* 伸張率 / scale */
        layerName: 16    /* レイヤー名 / layer name */
    };

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * 実行環境のロケールからUI言語を判定します。
     *
     * @returns {string} "ja" または "en"。
     */
    function detectUILanguage() {
        return ($.locale && $.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = detectUILanguage();

    /* UI文言の定義 / UI string definitions */
    var LABELS = {
        dialog: {
            title: { ja: "ロゴグリッドの作成", en: "Logo Grid Maker" }
        },
        panel: {
            common: { ja: "共通設定とクリアスペース", en: "Common Settings & Clear Space" },
            horizontal: { ja: "横線", en: "Horizontal Lines" },
            vertical: { ja: "縦線", en: "Vertical Lines" },
            lineMethod: { ja: "ラインの決め方", en: "Line Generation Method" },
            options: { ja: "オプション", en: "Options" }
        },
        fieldLabel: {
            preset: { ja: "プリセット", en: "Preset" },
            divisions: { ja: "分割数", en: "Divisions" },
            strokeWidth: { ja: "線幅", en: "Stroke Width" },
            targetLayer: { ja: "作成レイヤー", en: "Target Layer" },
            widthScale: { ja: "横伸張率", en: "Width Scale" },
            heightScale: { ja: "縦伸張率", en: "Height Scale" }
        },
        radio: {
            lineNone: { ja: "なし", en: "None" },
            lineAuto: { ja: "自動判定", en: "Auto Detect" },
            lineSegment: { ja: "水平セグメント", en: "Horizontal Segments" },
            lineEven: { ja: "均等分割", en: "Even Divisions" }
        },
        checkbox: {
            extraHorizontal: { ja: "上下に線を追加", en: "Add Top & Bottom Lines" },
            extendLeft: { ja: "左方向に延長", en: "Extend Leftward" },
            outerVertical: { ja: "左右に線を追加", en: "Add Left & Right Lines" },
            columnDiv: { ja: "均等分割", en: "Even Divisions" },
            extendUp: { ja: "上方向に延長", en: "Extend Upward" },
            verticalElements: { ja: "垂直エレメント", en: "Vertical Elements" },
            diagonalElements: { ja: "斜線エレメント", en: "Diagonal Elements" },
            convertToGuides: { ja: "ガイドに変換する", en: "Convert to Guides" },
            emphasizeBounds: { ja: "境界線を強調", en: "Highlight Bounds" },
            clearSpace: { ja: "クリアスペース", en: "Clear Space" },
            groupItems: { ja: "グループ化", en: "Group Items" },
            preview: { ja: "プレビュー", en: "Preview" }
        },
        unit: {
            lines: { ja: "本", en: "lines" }
        },
        button: {
            exportPreset: { ja: "書き出し", en: "Export" },
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" }
        },
        itemName: {
            layer: { ja: "// logo_guide", en: "// logo_guide" },
            guideGroup: { ja: "補助線", en: "Guides" },
            previewGroup: { ja: "補助線_プレビュー", en: "Guides (Preview)" },
            clearSpaceGroup: { ja: "クリアスペース", en: "Clear Space" }
        },
        tooltip: {
            preset: {
                ja: "よく使う設定の組み合わせを一括で反映します。",
                en: "Applies a ready-made combination of settings."
            },
            exportPreset: {
                ja: "現在の設定を、プリセット定義に貼り付けられる形でデスクトップに書き出します。",
                en: "Writes the current settings to the desktop, ready to paste into the preset definitions."
            },
            clearSpace: {
                ja: "選択範囲の外側に、ユニット幅のクリアスペースを作成します。ONのあいだ横線・縦線は作成されません。",
                en: "Draws a clear space one unit wide around the selection. Horizontal and vertical lines are skipped while this is on."
            },
            divisions: {
                ja: "選択範囲の高さを何分割した長さをユニット（1単位）とするかを指定します。",
                en: "Sets the unit size, as the number of parts the height of the selection is divided into."
            },
            strokeWidth: {
                ja: "補助線の線幅。単位は環境設定の［線］に従います。",
                en: "Stroke width of the guide lines, in the unit set under Preferences > Stroke."
            },
            emphasizeBounds: {
                ja: "選択範囲の境界に重なる線だけ、指定した線幅の4倍で描きます。",
                en: "Draws the lines that sit on the bounds of the selection four times heavier."
            },
            convertToGuides: {
                ja: "作成した線をガイドに変換します。ONのあいだグループ化は使えません。",
                en: "Converts the generated lines into guides. Grouping is unavailable while this is on."
            },
            groupItems: {
                ja: "作成した線をひとつのグループにまとめます。OFFならレイヤー直下に置きます。",
                en: "Puts the generated lines into a single group. When off, they are placed directly on the layer."
            },
            targetLayer: {
                ja: "補助線を作成するレイヤー名。同名のレイヤーがなければ新規作成します。",
                en: "Name of the layer the lines are created on. A layer is added when no layer has this name."
            },
            widthScale: {
                ja: "横線の長さ。選択範囲の幅に対する割合で指定します。",
                en: "Length of the horizontal lines, as a percentage of the width of the selection."
            },
            extraHorizontal: {
                ja: "選択範囲の上端・下端の線と、その外側にユニット間隔で指定した本数の線を追加します。",
                en: "Adds lines on the top and bottom bounds, plus the given number of lines outside them, one unit apart."
            },
            extendLeft: {
                ja: "横線の右端を選択範囲の右外に固定し、余った長さを左方向に伸ばします。",
                en: "Pins the right end of the horizontal lines and extends the remaining length leftward."
            },
            lineMethod: {
                ja: "選択範囲の内側に引く横線の決め方。",
                en: "How the horizontal lines inside the selection are placed."
            },
            lineNone: {
                ja: "内側の横線を作成しません。",
                en: "Draws no lines inside the selection."
            },
            lineAuto: {
                ja: "アセンダー・ミーンライン・ベースライン・ディセンダーを推定して引きます。",
                en: "Estimates the ascender, mean line, baseline and descender."
            },
            lineSegment: {
                ja: "パスに含まれる水平セグメントの位置に引きます。",
                en: "Places the lines on the horizontal segments found in the paths."
            },
            lineEven: {
                ja: "上端から下端までを、指定した本数で均等に分割して引きます。",
                en: "Divides the height evenly into the given number of lines."
            },
            heightScale: {
                ja: "縦線の長さ。選択範囲の高さに対する割合で指定します。",
                en: "Length of the vertical lines, as a percentage of the height of the selection."
            },
            outerVertical: {
                ja: "選択範囲の左端・右端の線と、その外側にユニット間隔で指定した本数の線を追加します。",
                en: "Adds lines on the left and right bounds, plus the given number of lines outside them, one unit apart."
            },
            extendUp: {
                ja: "縦線の下端を選択範囲の下（2ユニット）に固定し、余った長さを上方向に伸ばします。",
                en: "Pins the bottom of the vertical lines two units below the selection and extends the rest upward."
            },
            columnDiv: {
                ja: "選択範囲の幅を、指定した数で均等分割する縦線を引きます。",
                en: "Divides the width of the selection evenly with vertical lines."
            },
            verticalElements: {
                ja: "パスに含まれる垂直セグメントの位置に縦線を引きます。",
                en: "Places vertical lines on the vertical segments found in the paths."
            },
            diagonalElements: {
                ja: "斜体などの傾きを検出し、その角度に沿った線を引きます。",
                en: "Detects the dominant slant of the letterforms and draws lines along it."
            },
            preview: {
                ja: "設定の結果をドキュメント上で確認します。",
                en: "Shows the result on the document while you adjust the settings."
            }
        },
        prompt: {
            presetName: { ja: "プリセット名を入力してください：", en: "Enter preset name:" }
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noSelection: { ja: "オブジェクトを選択してください。", en: "Please select an object." },
            noBounds: { ja: "選択範囲の境界を取得できませんでした。", en: "Could not get the bounds of the selection." },
            invalidSize: { ja: "選択オブジェクトのサイズを取得できませんでした。", en: "Could not get the size of the selected object." },
            invalidInput: { ja: "入力値が不正です。数値を正しく入力してください。", en: "Invalid input. Please enter valid numeric values." },
            exported: {
                ja: "書き出しました（配列に貼り付け用）:\n",
                en: "Exported (for array paste):\n"
            },
            exportFailed: {
                ja: "プリセットを書き出せませんでした。",
                en: "Could not export the preset."
            }
        }
    };

    /**
     * ラベル（ja/en の組）を現在の言語の文言にします。
     *
     * @param {Object} labelSet - { ja, en } 形式のラベル。
     * @returns {string} 現在の言語の文言。
     */
    function getLabel(labelSet) {
        return (labelSet && labelSet[uiLang]) || "";
    }

    /**
     * 入力欄の前に置く項目名に、言語別のコロンを付けます（日本語は全角、英語は半角）。
     *
     * @param {Object} labelSet - { ja, en } 形式のラベル。
     * @returns {string} コロンを付けた文言。
     */
    function labelText(labelSet) {
        return getLabel(labelSet) + ((uiLang === "ja") ? "：" : ":");
    }

    // =========================================
    // 単位 / Units
    // =========================================

    /* 単位コードに対応する表示ラベルと、1単位あたりのポイント数
       Unit code -> display label and points per unit */
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
     * 環境設定キーの単位を返す
     * @param {string} [prefKey] - "rulerType"（既定）/ "strokeUnits" / "text/units" / "text/asianunits"
     * @returns {{code: number, label: string, pointsPerUnit: number}} 単位の情報
     */
    function getUnitInfo(prefKey) {
        var unitKey = prefKey || "rulerType";
        var unitCode = app.preferences.getIntegerPreference(unitKey);
        /* 未知のコードは pt に寄せる / unknown codes fall back to points */
        var unit = UNITS[unitCode] || UNITS[2];
        /* 級（Q）と歯（H）は同じ長さだが、文字サイズは「Q」、距離は「H」と呼び分ける */
        var label = (unitCode === 5 && HA_UNIT_PREF_KEYS[unitKey]) ? "H" : unit.label;
        return { code: unitCode, label: label, pointsPerUnit: unit.pointsPerUnit };
    }

    /**
     * 環境設定の［線］の単位で入力された値をptに換算します。
     *
     * @param {number} value - 線の単位での値。
     * @returns {number} pt換算した値。
     */
    function strokeUnitsToPoints(value) {
        return value * getUnitInfo("strokeUnits").pointsPerUnit;
    }

    // =========================================
    // 共通ユーティリティ / Shared utilities
    // =========================================

    /**
     * 2つの座標を同じ位置とみなせるか判定します（吸着結果の微小なズレを吸収）。
     *
     * @param {number} value1 - 比較する座標。
     * @param {number} value2 - 比較する座標。
     * @returns {boolean} 同じ位置とみなせるとき true。
     */
    function isSameCoord(value1, value2) {
        return Math.abs(value1 - value2) < DETECTION.coordTolerance;
    }

    /**
     * 対象サイズからクラスタリングの許容値を求めます。
     *
     * @param {number} size - 対象の幅または高さ（pt）。
     * @returns {number} 同じ線とみなす許容値（pt）。
     */
    function getClusterTolerance(size) {
        return Math.max(DETECTION.clusterMin, Math.min(DETECTION.clusterMax, size * DETECTION.clusterRatio));
    }

    /**
     * 短すぎるセグメントを捨てるためのしきい値を求めます。
     *
     * @param {number} totalLength - セグメント長の合計（pt）。
     * @param {number} count - セグメント数。
     * @returns {number} 採用する最小の長さ（pt）。
     */
    function getMinSegmentLength(totalLength, count) {
        return Math.max(DETECTION.shortSegmentFloor, totalLength * DETECTION.shortSegmentRatio / count);
    }

    /**
     * 並べ替え済みの配列を、値の近さでクラスタに分けます。
     *
     * @param {Array<Object>} sortedItems - 値の順に並べた配列。
     * @param {function(Object):number} getValue - 要素から比較する値を取り出す関数。
     * @param {number} tolerance - 同じクラスタとみなす差。
     * @returns {Array<Array<Object>>} クラスタの配列。
     */
    function clusterByValue(sortedItems, getValue, tolerance) {
        var clusters = [];
        var currentCluster = null;
        var baseValue = 0;
        for (var i = 0; i < sortedItems.length; i++) {
            var value = getValue(sortedItems[i]);
            if (!currentCluster || Math.abs(baseValue - value) > tolerance) {
                currentCluster = [];
                baseValue = value;
                clusters.push(currentCluster);
            }
            currentCluster.push(sortedItems[i]);
        }
        return clusters;
    }

    /**
     * 重み付き平均を求めます。
     *
     * @param {Array<Object>} weightedItems - 対象の配列。
     * @param {function(Object):number} getValue - 平均する値を取り出す関数。
     * @param {function(Object):number} getWeight - 重みを取り出す関数。
     * @returns {number} 重み付き平均（重みの合計が0のときは0）。
     */
    function weightedAverage(weightedItems, getValue, getWeight) {
        var total = 0;
        var sum = 0;
        for (var i = 0; i < weightedItems.length; i++) {
            var weight = getWeight(weightedItems[i]);
            total += weight;
            sum += getValue(weightedItems[i]) * weight;
        }
        return total > 0 ? (sum / total) : 0;
    }

    /**
     * グレー（K）のカラーを作成します。
     *
     * @param {number} tint - 濃度（0〜100）。
     * @returns {GrayColor} 作成したカラー。
     */
    function makeGrayColor(tint) {
        var grayColor = new GrayColor();
        grayColor.gray = tint;
        return grayColor;
    }

    /**
     * シアンのみのCMYKカラーを作成します。
     *
     * @param {number} tint - シアンの濃度（0〜100）。
     * @returns {CMYKColor} 作成したカラー。
     */
    function makeCyanColor(tint) {
        var cyanColor = new CMYKColor();
        cyanColor.cyan = tint;
        cyanColor.magenta = 0;
        cyanColor.yellow = 0;
        cyanColor.black = 0;
        return cyanColor;
    }

    /**
     * 補助線を作成するレイヤーを用意します。新しく作ったレイヤーは控えておきます。
     *
     * @param {Object} gridContext - buildGridContext() が返すコンテキスト。
     * @param {string} layerName - レイヤー名。
     * @returns {Layer} 取得または作成したレイヤー。
     */
    function ensureGuideLayer(gridContext, layerName) {
        for (var i = 0; i < gridContext.doc.layers.length; i++) {
            if (gridContext.doc.layers[i].name === layerName) {
                return gridContext.doc.layers[i];
            }
        }
        var newLayer = gridContext.doc.layers.add();
        newLayer.name = layerName;
        gridContext.createdLayers.push(newLayer);
        return newLayer;
    }

    /**
     * このスクリプトが作ったレイヤーのうち、空のまま残ったものを削除します。
     *
     * @param {Object} gridContext - buildGridContext() が返すコンテキスト。
     * @returns {void}
     */
    function removeEmptyCreatedLayers(gridContext) {
        for (var i = gridContext.createdLayers.length - 1; i >= 0; i--) {
            /* ユーザーが先に削除したレイヤーは参照が無効になる / The layer may already be gone */
            try {
                if (gridContext.createdLayers[i].pageItems.length === 0) {
                    gridContext.createdLayers[i].remove();
                }
            } catch (e) { }
        }
        gridContext.createdLayers = [];
    }

    /**
     * レイヤー名の入力値から改行・タブと前後の空白を取り除きます。
     *
     * @param {string} inputText - 入力された文字列。
     * @returns {string} 整えた文字列。
     */
    function normalizeLayerNameText(inputText) {
        return String(inputText).replace(/[\r\n\t]+/g, " ").replace(/^\s+|\s+$/g, "");
    }

    // =========================================
    // 選択範囲とジオメトリ / Selection and geometry
    // =========================================

    /**
     * オブジェクトの visibleBounds を取得します（取得できないときは null）。
     *
     * @param {PageItem} pageItem - 対象のオブジェクト。
     * @returns {Array<number>|null} [左, 上, 右, 下] または null。
     */
    function getVisibleBounds(pageItem) {
        /* visibleBounds を持たない・取得できないオブジェクトがある / Some items throw on visibleBounds */
        try {
            return pageItem.visibleBounds;
        } catch (e) {
            return null;
        }
    }

    /**
     * 選択範囲全体の visibleBounds を求めます。
     *
     * @param {Array<PageItem>} pageItems - 選択中のオブジェクト。
     * @returns {Array<number>|null} [左, 上, 右, 下] または null。
     */
    function getSelectionBounds(pageItems) {
        var left, top, right, bottom;
        var hasBounds = false;

        for (var i = 0; i < pageItems.length; i++) {
            var itemBounds = getVisibleBounds(pageItems[i]);
            if (!itemBounds) continue;

            if (!hasBounds) {
                left = itemBounds[0];
                top = itemBounds[1];
                right = itemBounds[2];
                bottom = itemBounds[3];
                hasBounds = true;
                continue;
            }
            if (itemBounds[0] < left) left = itemBounds[0];
            if (itemBounds[1] > top) top = itemBounds[1];
            if (itemBounds[2] > right) right = itemBounds[2];
            if (itemBounds[3] < bottom) bottom = itemBounds[3];
        }

        return hasBounds ? [left, top, right, bottom] : null;
    }

    /**
     * グループ・複合パスをたどって、すべてのパスに処理を適用します。
     *
     * @param {Array<PageItem>} pageItems - 対象のオブジェクト。
     * @param {function(PathItem):void} handler - パスごとに呼ぶ処理。
     * @returns {void}
     */
    function forEachPathItem(pageItems, handler) {
        for (var i = 0; i < pageItems.length; i++) {
            var pageItem = pageItems[i];
            if (pageItem.typename === "PathItem") {
                handler(pageItem);
            } else if (pageItem.typename === "GroupItem") {
                forEachPathItem(pageItem.pageItems, handler);
            } else if (pageItem.typename === "CompoundPathItem") {
                forEachPathItem(pageItem.pathItems, handler);
            }
        }
    }

    /**
     * 2つのアンカーポイントのあいだが直線（ハンドルなし）か判定します。
     *
     * @param {PathPoint} startPoint - 始点のアンカーポイント。
     * @param {PathPoint} endPoint - 終点のアンカーポイント。
     * @returns {boolean} 直線のとき true。
     */
    function isStraightSegment(startPoint, endPoint) {
        var tolerance = DETECTION.coordTolerance;
        return Math.abs(startPoint.rightDirection[0] - startPoint.anchor[0]) < tolerance &&
            Math.abs(startPoint.rightDirection[1] - startPoint.anchor[1]) < tolerance &&
            Math.abs(endPoint.leftDirection[0] - endPoint.anchor[0]) < tolerance &&
            Math.abs(endPoint.leftDirection[1] - endPoint.anchor[1]) < tolerance;
    }

    /**
     * 直線セグメントの向きを判定します。
     *
     * @param {Array<number>} startAnchor - 始点の座標 [x, y]。
     * @param {Array<number>} endAnchor - 終点の座標 [x, y]。
     * @returns {string} "HORIZONTAL" / "VERTICAL" / "DIAGONAL"。
     */
    function classifySegmentDirection(startAnchor, endAnchor) {
        if (isSameCoord(startAnchor[1], endAnchor[1])) return "HORIZONTAL";
        if (isSameCoord(startAnchor[0], endAnchor[0])) return "VERTICAL";
        return "DIAGONAL";
    }

    /**
     * 斜線セグメントの情報（角度は0〜180度に正規化）を作成します。
     *
     * @param {Array<number>} startAnchor - 始点の座標 [x, y]。
     * @param {Array<number>} endAnchor - 終点の座標 [x, y]。
     * @returns {{startAnchor: Array<number>, endAnchor: Array<number>, angle: number, length: number}} 斜線セグメント。
     */
    function makeDiagonalSegment(startAnchor, endAnchor) {
        var dx = endAnchor[0] - startAnchor[0];
        var dy = endAnchor[1] - startAnchor[1];
        var angle = Math.atan2(dy, dx) * 180 / Math.PI;
        if (angle < 0) angle += 180;
        if (angle >= 180) angle -= 180;
        return {
            startAnchor: startAnchor,
            endAnchor: endAnchor,
            angle: angle,
            length: Math.sqrt(dx * dx + dy * dy)
        };
    }

    /**
     * 選択範囲のアンカーポイントと直線セグメントを、向きごとに集めます。
     *
     * @param {Array<PageItem>} pageItems - 選択中のオブジェクト。
     * @returns {{points: Array<Object>, horizontal: Array<Object>, vertical: Array<Object>, diagonal: Array<Object>, top: number, bottom: number}} 解析用のジオメトリ。
     */
    function collectGeometry(pageItems) {
        var geometry = {
            points: [],
            horizontal: [],
            vertical: [],
            diagonal: [],
            top: -Infinity,
            bottom: Infinity
        };

        forEachPathItem(pageItems, function (pathItem) {
            var pathPoints = pathItem.pathPoints;
            var count = pathPoints.length;
            for (var i = 0; i < count; i++) {
                var nextIndex = (i + 1) % count;
                var startAnchor = pathPoints[i].anchor;
                var endAnchor = pathPoints[nextIndex].anchor;

                if (startAnchor[1] > geometry.top) geometry.top = startAnchor[1];
                if (startAnchor[1] < geometry.bottom) geometry.bottom = startAnchor[1];

                /* 水平に近い辺の端点は、自動判定で横線の手がかりにする */
                var onHorizontal = Math.abs(startAnchor[1] - endAnchor[1]) < DETECTION.horizontalFlatness &&
                    Math.abs(startAnchor[0] - endAnchor[0]) > DETECTION.horizontalMinRun;
                geometry.points.push({ x: startAnchor[0], y: startAnchor[1], onHorizontal: onHorizontal });

                /* 開いたパスの終端は次の点とつながらない / The last point of an open path has no next segment */
                if (i >= count - 1 && !pathItem.closed) continue;
                if (!isStraightSegment(pathPoints[i], pathPoints[nextIndex])) continue;

                var direction = classifySegmentDirection(startAnchor, endAnchor);
                if (direction === "HORIZONTAL") {
                    geometry.horizontal.push({ position: startAnchor[1], length: Math.abs(startAnchor[0] - endAnchor[0]) });
                } else if (direction === "VERTICAL") {
                    geometry.vertical.push({ position: startAnchor[0], length: Math.abs(startAnchor[1] - endAnchor[1]) });
                } else {
                    geometry.diagonal.push(makeDiagonalSegment(startAnchor, endAnchor));
                }
            }
        });

        return geometry;
    }

    /**
     * 選択範囲の寸法とジオメトリをまとめた、処理用のコンテキストを作成します。
     *
     * @param {Document} targetDoc - 対象のドキュメント。
     * @param {Array<PageItem>} selectedItems - 選択中のオブジェクト。
     * @returns {Object|null} コンテキスト。作成できないときは null。
     */
    function buildGridContext(targetDoc, selectedItems) {
        var selectionBounds = getSelectionBounds(selectedItems);
        if (!selectionBounds) {
            alert(getLabel(LABELS.alert.noBounds));
            return null;
        }

        var selLeft = selectionBounds[0];
        var selTop = selectionBounds[1];
        var selRight = selectionBounds[2];
        var selBottom = selectionBounds[3];
        var selWidth = selRight - selLeft;
        var selHeight = selTop - selBottom;

        if (selWidth <= 0 || selHeight <= 0) {
            alert(getLabel(LABELS.alert.invalidSize));
            return null;
        }

        return {
            doc: targetDoc,
            items: selectedItems,
            selLeft: selLeft,
            selTop: selTop,
            selRight: selRight,
            selBottom: selBottom,
            selWidth: selWidth,
            selHeight: selHeight,
            centerX: (selLeft + selRight) / 2,
            centerY: (selTop + selBottom) / 2,
            geometry: collectGeometry(selectedItems),
            guideLayer: null,
            createdLayers: []
        };
    }

    // =========================================
    // ライン解析 / Line analysis
    // =========================================

    /**
     * アンカーポイントのクラスタから、代表となる横線の位置を求めます。
     *
     * @param {Array<Object>} points - 同じ高さに集まったアンカーポイント。
     * @returns {{y: number, count: number, onHorizontal: boolean}} 代表位置と手がかり。
     */
    function summarizeAnchorCluster(points) {
        var frequency = {};
        var onHorizontal = false;

        for (var i = 0; i < points.length; i++) {
            if (points[i].onHorizontal) onHorizontal = true;
            var yKey = points[i].y.toFixed(3);
            frequency[yKey] = (frequency[yKey] || 0) + 1;
        }

        var bestY = points[0].y;
        var bestCount = -1;
        for (var frequencyKey in frequency) {
            if (frequency.hasOwnProperty(frequencyKey) && frequency[frequencyKey] > bestCount) {
                bestCount = frequency[frequencyKey];
                bestY = parseFloat(frequencyKey);
            }
        }

        return { y: bestY, count: points.length, onHorizontal: onHorizontal };
    }

    /**
     * 候補のうち、もっとも横線らしい1本を選びます（水平な辺を持つものを優先）。
     *
     * @param {Array<Object>} candidateLines - 候補の横線。
     * @returns {Object} 選ばれた横線。
     */
    function pickDominantLine(candidateLines) {
        var dominantLine = candidateLines[0];
        var bestScore = -Infinity;
        for (var i = 0; i < candidateLines.length; i++) {
            var score = candidateLines[i].count + (candidateLines[i].onHorizontal ? 1000 : 0);
            if (score > bestScore) {
                bestScore = score;
                dominantLine = candidateLines[i];
            }
        }
        return dominantLine;
    }

    /**
     * アンカーポイントの分布から、書体の基準線（アセンダー〜ディセンダー）を推定します。
     *
     * @param {Object} geometry - collectGeometry() の結果。
     * @returns {{ascenderY: number, meanY: number, baseY: number, descenderY: number, hasDescender: boolean}} 推定した基準線。
     */
    function detectTypographicLines(geometry) {
        var tolerance = getClusterTolerance(geometry.top - geometry.bottom);
        var sortedPoints = geometry.points.slice().sort(function (a, b) { return b.y - a.y; });
        var clusters = clusterByValue(sortedPoints, function (point) { return point.y; }, tolerance);

        var anchorLines = [];
        for (var i = 0; i < clusters.length; i++) {
            anchorLines.push(summarizeAnchorCluster(clusters[i]));
        }

        var ascenderY = anchorLines[0].y;
        var descenderY = anchorLines[anchorLines.length - 1].y;
        var height = ascenderY - descenderY;
        var middleY = (ascenderY + descenderY) / 2;

        /* 上下端から離れた候補だけを、ミーンラインとベースラインの候補にする */
        var upperLines = [];
        var lowerLines = [];
        for (var j = 1; j < anchorLines.length - 1; j++) {
            if (Math.abs(ascenderY - anchorLines[j].y) <= tolerance || Math.abs(anchorLines[j].y - descenderY) <= tolerance) continue;
            if (anchorLines[j].y > middleY) {
                upperLines.push(anchorLines[j]);
            } else {
                lowerLines.push(anchorLines[j]);
            }
        }

        var meanY, baseY;
        if (upperLines.length > 0 && lowerLines.length > 0) {
            meanY = pickDominantLine(upperLines).y;
            baseY = pickDominantLine(lowerLines).y;
        } else {
            meanY = ascenderY - height * DETECTION.fallbackLineRatio;
            baseY = descenderY + height * DETECTION.fallbackLineRatio;
        }

        var hasDescender = !(height > 0 && (baseY - descenderY) / height < DETECTION.descenderRatio);

        return {
            ascenderY: ascenderY,
            meanY: meanY,
            baseY: baseY,
            descenderY: descenderY,
            hasDescender: hasDescender
        };
    }

    /**
     * 直線セグメントをクラスタリングし、長さで加重平均した代表位置を返します。
     *
     * @param {Array<Object>} segments - position と length を持つセグメント。
     * @param {{descending: boolean, dropShort: boolean, span: number}} clusterOptions - 並び順・短いセグメントの扱い・許容値の基準サイズ。
     * @returns {Array<number>} 線の位置。
     */
    function getSegmentLinePositions(segments, clusterOptions) {
        if (segments.length === 0) return [];

        var minPosition = Infinity;
        var maxPosition = -Infinity;
        var totalLength = 0;
        for (var i = 0; i < segments.length; i++) {
            if (segments[i].position < minPosition) minPosition = segments[i].position;
            if (segments[i].position > maxPosition) maxPosition = segments[i].position;
            totalLength += segments[i].length;
        }

        var keptSegments = segments;
        if (clusterOptions.dropShort) {
            var minLength = getMinSegmentLength(totalLength, segments.length);
            keptSegments = [];
            for (var j = 0; j < segments.length; j++) {
                if (segments[j].length >= minLength) keptSegments.push(segments[j]);
            }
            if (keptSegments.length === 0) return [];
        }

        var span = (typeof clusterOptions.span === "number") ? clusterOptions.span : (maxPosition - minPosition);
        var descending = !!clusterOptions.descending;
        var sortedSegments = keptSegments.slice().sort(function (a, b) {
            return descending ? (b.position - a.position) : (a.position - b.position);
        });
        var clusters = clusterByValue(sortedSegments, function (segment) { return segment.position; }, getClusterTolerance(span));

        var positions = [];
        for (var k = 0; k < clusters.length; k++) {
            positions.push(weightedAverage(clusters[k],
                function (segment) { return segment.position; },
                function (segment) { return segment.length; }));
        }

        positions.sort(function (a, b) { return descending ? (b - a) : (a - b); });
        return positions;
    }

    /**
     * ［ラインの決め方］に従って、横線を引くY座標を求めます。
     *
     * @param {Object} geometry - collectGeometry() の結果。
     * @param {string} lineMethod - "none" / "segment" / "even" / "auto"。
     * @param {number} evenLineCount - 均等分割の本数。
     * @returns {Array<number>} 横線のY座標。
     */
    function getHorizontalLineYs(geometry, lineMethod, evenLineCount) {
        if (geometry.points.length === 0 || lineMethod === "none") return [];

        if (lineMethod === "segment") {
            return getSegmentLinePositions(geometry.horizontal, {
                descending: true,
                dropShort: false,
                span: geometry.top - geometry.bottom
            });
        }

        if (lineMethod === "even") {
            var evenLineYs = [geometry.top];
            var step = (geometry.top - geometry.bottom) / (evenLineCount + 1);
            for (var i = 1; i <= evenLineCount; i++) {
                evenLineYs.push(geometry.top - step * i);
            }
            evenLineYs.push(geometry.bottom);
            return evenLineYs;
        }

        var zones = detectTypographicLines(geometry);
        var autoLines = [zones.ascenderY, zones.meanY];
        if (zones.hasDescender) autoLines.push(zones.baseY);
        autoLines.push(zones.descenderY);
        return autoLines;
    }

    /**
     * 垂直エレメント（縦のステム）のX座標を求めます。
     *
     * @param {Object} geometry - collectGeometry() の結果。
     * @returns {Array<number>} 縦線のX座標。
     */
    function getVerticalElementXs(geometry) {
        return getSegmentLinePositions(geometry.vertical, { descending: false, dropShort: true });
    }

    /**
     * 2点を通る直線を、指定した矩形の境界まで伸ばします。
     *
     * @param {Array<number>} startAnchor - 始点の座標 [x, y]。
     * @param {Array<number>} endAnchor - 終点の座標 [x, y]。
     * @param {number} left - 左端。
     * @param {number} top - 上端。
     * @param {number} right - 右端。
     * @param {number} bottom - 下端。
     * @returns {Array<Array<number>>|null} 交点2つ。求まらないときは null。
     */
    function extendLineToBounds(startAnchor, endAnchor, left, top, right, bottom) {
        var x1 = startAnchor[0], y1 = startAnchor[1];
        var x2 = endAnchor[0], y2 = endAnchor[1];
        var tolerance = DETECTION.coordTolerance;
        var intersections = [];

        if (isSameCoord(x1, x2)) {
            intersections.push([x1, top]);
            intersections.push([x1, bottom]);
        } else if (isSameCoord(y1, y2)) {
            intersections.push([left, y1]);
            intersections.push([right, y1]);
        } else {
            var slope = (y2 - y1) / (x2 - x1);
            var intercept = y1 - slope * x1;

            var yAtLeft = slope * left + intercept;
            if (yAtLeft <= top + tolerance && yAtLeft >= bottom - tolerance) intersections.push([left, yAtLeft]);

            var yAtRight = slope * right + intercept;
            if (yAtRight <= top + tolerance && yAtRight >= bottom - tolerance) intersections.push([right, yAtRight]);

            var xAtTop = (top - intercept) / slope;
            if (xAtTop >= left - tolerance && xAtTop <= right + tolerance) intersections.push([xAtTop, top]);

            var xAtBottom = (bottom - intercept) / slope;
            if (xAtBottom >= left - tolerance && xAtBottom <= right + tolerance) intersections.push([xAtBottom, bottom]);
        }

        if (intersections.length < 2) return null;

        /* 同じ点が並ぶことがあるため、離れた2点を選ぶ / Pick two points that are actually apart */
        for (var i = 1; i < intersections.length; i++) {
            var dx = intersections[i][0] - intersections[0][0];
            var dy = intersections[i][1] - intersections[0][1];
            if (Math.sqrt(dx * dx + dy * dy) > tolerance) {
                return [intersections[0], intersections[i]];
            }
        }
        return null;
    }

    /**
     * 1点と角度から、指定した矩形の境界まで伸びる直線を求めます。
     *
     * @param {Array<number>} point - 通過する座標 [x, y]。
     * @param {number} angleRad - 角度（ラジアン）。
     * @param {number} left - 左端。
     * @param {number} top - 上端。
     * @param {number} right - 右端。
     * @param {number} bottom - 下端。
     * @returns {Array<Array<number>>|null} 交点2つ。求まらないときは null。
     */
    function extendPointAtAngle(point, angleRad, left, top, right, bottom) {
        var farDistance = 100000;
        var farPoint = [point[0] + Math.cos(angleRad) * farDistance, point[1] + Math.sin(angleRad) * farDistance];
        return extendLineToBounds(point, farPoint, left, top, right, bottom);
    }

    /**
     * 斜線の重複を判定するためのキーを作ります。
     *
     * @param {Array<Array<number>>} endpoints - 線の両端 [[x, y], [x, y]]。
     * @returns {string} 両端の座標から作ったキー。
     */
    function makeDiagonalLineKey(endpoints) {
        var keyStart = endpoints[0][0].toFixed(1) + "," + endpoints[0][1].toFixed(1);
        var keyEnd = endpoints[1][0].toFixed(1) + "," + endpoints[1][1].toFixed(1);
        return (keyStart < keyEnd) ? (keyStart + "|" + keyEnd) : (keyEnd + "|" + keyStart);
    }

    /**
     * 斜体の傾きとして扱える角度のセグメントだけを取り出します。
     *
     * @param {Array<Object>} segments - 斜線セグメント。
     * @returns {Array<Object>} 角度が範囲内のセグメント。
     */
    function filterDiagonalSegments(segments) {
        var filtered = [];
        for (var i = 0; i < segments.length; i++) {
            if (segments[i].angle > DETECTION.diagonalMinAngle && segments[i].angle < DETECTION.diagonalMaxAngle) {
                filtered.push(segments[i]);
            }
        }
        return filtered;
    }

    /**
     * もっとも多く使われている傾きのセグメント群を選びます。
     *
     * @param {Array<Object>} segments - 斜線セグメント。
     * @returns {Array<Object>|null} 選ばれたセグメント群。見つからないときは null。
     */
    function pickDominantAngleCluster(segments) {
        var sortedSegments = segments.slice().sort(function (a, b) { return a.angle - b.angle; });
        var clusters = clusterByValue(sortedSegments, function (segment) { return segment.angle; }, DETECTION.diagonalAngleTolerance);

        var dominantCluster = null;
        var bestLength = -1;
        for (var i = 0; i < clusters.length; i++) {
            var totalLength = 0;
            for (var j = 0; j < clusters[i].length; j++) {
                totalLength += clusters[i][j].length;
            }
            if (totalLength > bestLength) {
                bestLength = totalLength;
                dominantCluster = clusters[i];
            }
        }
        return dominantCluster;
    }

    /**
     * 斜線セグメントの中点を、傾きに直交する方向でまとめます。
     *
     * @param {Array<Object>} segments - 同じ傾きのセグメント。
     * @param {number} angleRad - 傾き（ラジアン）。
     * @param {number} minLength - 採用する最小の長さ。
     * @returns {Array<Array<number>>} まとめた中点の座標。
     */
    function mergeSegmentMidpoints(segments, angleRad, minLength) {
        var normalX = -Math.sin(angleRad);
        var normalY = Math.cos(angleRad);

        var midpoints = [];
        for (var i = 0; i < segments.length; i++) {
            if (segments[i].length < minLength) continue;
            var x = (segments[i].startAnchor[0] + segments[i].endAnchor[0]) / 2;
            var y = (segments[i].startAnchor[1] + segments[i].endAnchor[1]) / 2;
            midpoints.push({ x: x, y: y, length: segments[i].length, offset: x * normalX + y * normalY });
        }
        if (midpoints.length === 0) return [];

        midpoints.sort(function (a, b) { return a.offset - b.offset; });
        var mergeTolerance = Math.max(DETECTION.shortSegmentFloor, minLength * DETECTION.diagonalMergeRatio);
        var clusters = clusterByValue(midpoints, function (midpoint) { return midpoint.offset; }, mergeTolerance);

        var merged = [];
        for (var j = 0; j < clusters.length; j++) {
            merged.push([
                weightedAverage(clusters[j], function (midpoint) { return midpoint.x; }, function (midpoint) { return midpoint.length; }),
                weightedAverage(clusters[j], function (midpoint) { return midpoint.y; }, function (midpoint) { return midpoint.length; })
            ]);
        }
        return merged;
    }

    /**
     * 斜線エレメント（斜体の傾きに沿った線）を求めます。
     *
     * @param {Object} geometry - collectGeometry() の結果。
     * @param {number} left - 線を伸ばす範囲の左端。
     * @param {number} top - 線を伸ばす範囲の上端。
     * @param {number} right - 線を伸ばす範囲の右端。
     * @param {number} bottom - 線を伸ばす範囲の下端。
     * @returns {Array<Array<Array<number>>>} 線の両端の座標の配列。
     */
    function getDiagonalElementLines(geometry, left, top, right, bottom) {
        var segments = filterDiagonalSegments(geometry.diagonal);
        if (segments.length === 0) return [];

        var dominantCluster = pickDominantAngleCluster(segments);
        if (!dominantCluster) return [];

        var totalLength = 0;
        for (var i = 0; i < dominantCluster.length; i++) {
            totalLength += dominantCluster[i].length;
        }
        var angleRad = weightedAverage(dominantCluster,
            function (segment) { return segment.angle; },
            function (segment) { return segment.length; }) * Math.PI / 180;

        var midpoints = mergeSegmentMidpoints(dominantCluster, angleRad, getMinSegmentLength(totalLength, dominantCluster.length));

        var diagonalLines = [];
        var usedKeys = {};
        for (var j = 0; j < midpoints.length; j++) {
            var diagonalLine = extendPointAtAngle(midpoints[j], angleRad, left, top, right, bottom);
            if (!diagonalLine) continue;
            var key = makeDiagonalLineKey(diagonalLine);
            if (usedKeys[key]) continue;
            usedKeys[key] = true;
            diagonalLines.push(diagonalLine);
        }
        return diagonalLines;
    }

    // =========================================
    // プリセット / Presets
    // =========================================

    /* プリセットの定義（伸張率と本数は入力欄の文字列） / Preset definitions */
    var PRESET_DEFINITIONS = {
        "1x1": {
            widthScale: "140.0", extraHLines: true, extraHLineCount: "1", extendLeft: false,
            lineMethod: "none", evenLineCount: "4",
            heightScale: "200.0", outerVLines: true, outerVLineCount: "1", extendUp: false,
            columnDiv: false, columnDivisions: "2", verticalElements: false, diagonalElements: false,
            emphasizeBounds: false,
            clearSpace: false, clearSpaceDivisions: "4"
        },
        "auto": {
            widthScale: "120.0", extraHLines: true, extraHLineCount: "2", extendLeft: false,
            lineMethod: "auto", evenLineCount: "4",
            heightScale: "200.0", outerVLines: true, outerVLineCount: "2", extendUp: false,
            columnDiv: false, columnDivisions: "2", verticalElements: false, diagonalElements: false,
            emphasizeBounds: true,
            clearSpace: false, clearSpaceDivisions: "4"
        },
        "element": {
            widthScale: "120.0", extraHLines: false, extraHLineCount: "2", extendLeft: false,
            lineMethod: "segment", evenLineCount: "4",
            heightScale: "300.0", outerVLines: false, outerVLineCount: "2", extendUp: false,
            columnDiv: false, columnDivisions: "3", verticalElements: true, diagonalElements: true,
            emphasizeBounds: false,
            clearSpace: false, clearSpaceDivisions: "5"
        },
        "element+": {
            widthScale: "120.0", extraHLines: true, extraHLineCount: "2", extendLeft: false,
            lineMethod: "segment", evenLineCount: "4",
            heightScale: "300.0", outerVLines: true, outerVLineCount: "2", extendUp: false,
            columnDiv: false, columnDivisions: "3", verticalElements: true, diagonalElements: true,
            emphasizeBounds: true,
            clearSpace: false, clearSpaceDivisions: "5"
        },
        "left": {
            widthScale: "150.0", extraHLines: true, extraHLineCount: "2", extendLeft: true,
            lineMethod: "even", evenLineCount: "4",
            heightScale: "200.0", outerVLines: true, outerVLineCount: "2", extendUp: false,
            columnDiv: false, columnDivisions: "2", verticalElements: false, diagonalElements: false,
            emphasizeBounds: true,
            clearSpace: false, clearSpaceDivisions: "5"
        },
        "up-3": {
            widthScale: "120.0", extraHLines: true, extraHLineCount: "2", extendLeft: false,
            lineMethod: "even", evenLineCount: "4",
            heightScale: "300.0", outerVLines: true, outerVLineCount: "2", extendUp: true,
            columnDiv: true, columnDivisions: "3", verticalElements: false, diagonalElements: false,
            emphasizeBounds: true,
            clearSpace: false, clearSpaceDivisions: "5"
        },
        "clear space": {
            widthScale: "140.0", extraHLines: true, extraHLineCount: "1", extendLeft: false,
            lineMethod: "none", evenLineCount: "4",
            heightScale: "200.0", outerVLines: true, outerVLineCount: "1", extendUp: false,
            columnDiv: false, columnDivisions: "2", verticalElements: false, diagonalElements: false,
            emphasizeBounds: false,
            clearSpace: true, clearSpaceDivisions: "4"
        }
    };

    /* プリセットのキーとUIコントロールの対応 / Preset keys and the controls they map to */
    var PRESET_FIELDS = {
        widthScale: { control: "widthScaleInput", kind: "scale" },
        extraHLines: { control: "extraHLinesCheck", kind: "check" },
        extraHLineCount: { control: "extraHLinesInput", kind: "text" },
        extendLeft: { control: "extendLeftCheck", kind: "check" },
        lineMethod: { control: "", kind: "method" },
        evenLineCount: { control: "evenLineCountInput", kind: "text" },
        heightScale: { control: "heightScaleInput", kind: "scale" },
        outerVLines: { control: "outerVLinesCheck", kind: "check" },
        outerVLineCount: { control: "outerVLinesInput", kind: "text" },
        extendUp: { control: "extendUpCheck", kind: "check" },
        columnDiv: { control: "columnDivCheck", kind: "check" },
        columnDivisions: { control: "columnDivInput", kind: "text" },
        verticalElements: { control: "verticalElementsCheck", kind: "check" },
        diagonalElements: { control: "diagonalElementsCheck", kind: "check" },
        emphasizeBounds: { control: "emphasizeBoundsCheck", kind: "check" },
        clearSpace: { control: "clearSpaceCheck", kind: "check" },
        clearSpaceDivisions: { control: "clearSpaceDivInput", kind: "text" }
    };

    /* 書き出し時の行の並び / Key order used when a preset is exported */
    var PRESET_FIELD_ROWS = [
        ["widthScale", "extraHLines", "extraHLineCount", "extendLeft"],
        ["lineMethod", "evenLineCount"],
        ["heightScale", "outerVLines", "outerVLineCount", "extendUp"],
        ["columnDiv", "columnDivisions", "verticalElements", "diagonalElements"],
        ["emphasizeBounds"],
        ["clearSpace", "clearSpaceDivisions"]
    ];

    /**
     * ［ラインの決め方］の選択状態を取得します。
     *
     * @param {Object} dialogUI - buildDialog() が返すUI参照。
     * @returns {string} "none" / "even" / "segment" / "auto"。
     */
    function getLineMethod(dialogUI) {
        if (dialogUI.methodNoneRadio.value) return "none";
        if (dialogUI.methodEvenRadio.value) return "even";
        if (dialogUI.methodSegmentRadio.value) return "segment";
        return "auto";
    }

    /**
     * プリセットの1項目を、現在のUIから読み取ります。
     *
     * @param {Object} dialogUI - buildDialog() が返すUI参照。
     * @param {string} key - プリセットのキー。
     * @returns {string|boolean} 読み取った値。
     */
    function readPresetValue(dialogUI, key) {
        var presetField = PRESET_FIELDS[key];
        if (presetField.kind === "method") return getLineMethod(dialogUI);
        if (presetField.kind === "check") return dialogUI[presetField.control].value;
        return dialogUI[presetField.control].text;
    }

    /**
     * プリセットの1項目を、UIに反映します。
     *
     * @param {Object} dialogUI - buildDialog() が返すUI参照。
     * @param {string} key - プリセットのキー。
     * @param {string|boolean} value - 反映する値。
     * @returns {void}
     */
    function writePresetValue(dialogUI, key, value) {
        var presetField = PRESET_FIELDS[key];
        if (presetField.kind === "method") {
            dialogUI.methodNoneRadio.value = (value === "none");
            dialogUI.methodAutoRadio.value = (value === "auto");
            dialogUI.methodSegmentRadio.value = (value === "segment");
            dialogUI.methodEvenRadio.value = (value === "even");
            return;
        }
        if (presetField.kind === "check") {
            dialogUI[presetField.control].value = !!value;
            return;
        }
        if (presetField.kind === "scale") {
            dialogUI[presetField.control].text = parseFloat(value).toFixed(1);
            return;
        }
        dialogUI[presetField.control].text = String(value);
    }

    /**
     * 選択中のプリセットをUIに反映します。
     *
     * @param {Object} dialogUI - buildDialog() が返すUI参照。
     * @param {Object} gridContext - buildGridContext() が返すコンテキスト。
     * @returns {void}
     */
    function applyPreset(dialogUI, gridContext) {
        if (!dialogUI.presetDropdown.selection) return;
        var presetValues = PRESET_DEFINITIONS[dialogUI.presetDropdown.selection._key];
        if (!presetValues) return;

        for (var key in PRESET_FIELDS) {
            if (PRESET_FIELDS.hasOwnProperty(key)) {
                writePresetValue(dialogUI, key, presetValues[key]);
            }
        }

        updateClearSpaceState(dialogUI);
        refreshPreview(dialogUI, gridContext);
    }

    /**
     * プリセットの値を、定義に貼り付けられる表記にします。
     *
     * @param {string|boolean} value - プリセットの値。
     * @returns {string} 表記した値。
     */
    function formatPresetValue(value) {
        return (typeof value === "string") ? ('"' + value + '"') : String(value);
    }

    /**
     * 現在の設定を、プリセット定義に貼り付けられる形の文字列にします。
     *
     * @param {Object} dialogUI - buildDialog() が返すUI参照。
     * @param {string} presetName - プリセット名。
     * @returns {string} 書き出す文字列。
     */
    function buildPresetSnippet(dialogUI, presetName) {
        var snippetLines = ['                "' + presetName + '": {'];

        for (var i = 0; i < PRESET_FIELD_ROWS.length; i++) {
            var fieldRow = PRESET_FIELD_ROWS[i];
            var rowParts = [];
            for (var j = 0; j < fieldRow.length; j++) {
                rowParts.push(fieldRow[j] + ": " + formatPresetValue(readPresetValue(dialogUI, fieldRow[j])));
            }
            var isLastRow = (i === PRESET_FIELD_ROWS.length - 1);
            snippetLines.push("                    " + rowParts.join(", ") + (isLastRow ? "" : ","));
        }

        snippetLines.push("                },");
        return snippetLines.join("\n");
    }

    /**
     * 現在の設定をプリセットとしてデスクトップに書き出します。
     *
     * @param {Object} dialogUI - buildDialog() が返すUI参照。
     * @returns {void}
     */
    function exportPreset(dialogUI) {
        var presetName = prompt(getLabel(LABELS.prompt.presetName), "");
        if (!presetName) return;

        var fileName = presetName.replace(/[\\\/:*?"<>|]/g, "_") + ".json";
        var presetFile = new File(Folder.desktop + "/" + fileName);
        presetFile.encoding = "UTF-8";
        if (!presetFile.open("w")) {
            alert(getLabel(LABELS.alert.exportFailed));
            return;
        }
        presetFile.write(buildPresetSnippet(dialogUI, presetName));
        presetFile.close();
        alert(getLabel(LABELS.alert.exported) + presetFile.fsName);
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * パネルの余白と並びを設定します。
     *
     * @param {Panel} targetPanel - 対象のパネル。
     * @param {string} horizontalAlign - 子要素の横方向の揃え。
     * @returns {Panel} 設定したパネル。
     */
    function setupPanel(targetPanel, horizontalAlign) {
        targetPanel.margins = PANEL_MARGINS;
        targetPanel.alignChildren = [horizontalAlign || "left", "top"];
        return targetPanel;
    }

    /**
     * 行グループを作成します（横並び・左揃え・天地中央）。
     *
     * @param {Object} parentGroup - 追加先のコンテナ。
     * @returns {Group} 作成した行グループ。
     */
    function addRow(parentGroup) {
        var rowGroup = parentGroup.add("group");
        rowGroup.orientation = "row";
        rowGroup.alignment = ["left", "top"];
        rowGroup.alignChildren = ["left", "center"];
        return rowGroup;
    }

    /**
     * 縦並びのカラム用グループを作成します。
     *
     * @param {Group} parentGroup - 追加先のグループ。
     * @param {string} horizontalAlign - 子要素の横方向の揃え。
     * @returns {Group} 作成したカラム。
     */
    function addColumn(parentGroup, horizontalAlign) {
        var columnGroup = parentGroup.add("group");
        columnGroup.orientation = "column";
        columnGroup.alignChildren = [horizontalAlign, "top"];
        return columnGroup;
    }

    /**
     * カラムを横に並べるグループを作成します。
     *
     * @param {Object} parentGroup - 追加先のコンテナ。
     * @returns {Group} 作成したグループ。
     */
    function addColumnsGroup(parentGroup) {
        var columnsGroup = parentGroup.add("group");
        columnsGroup.orientation = "row";
        columnsGroup.alignChildren = ["fill", "top"];
        return columnsGroup;
    }

    /**
     * tooltip 付きのチェックボックスを追加します。
     *
     * @param {Object} parentGroup - 追加先のコンテナ。
     * @param {Object} labelSet - 表示名のラベル。
     * @param {Object} tooltipSet - tooltip のラベル。
     * @param {boolean} initialValue - 初期値。
     * @returns {Checkbox} 追加したチェックボックス。
     */
    function addCheckbox(parentGroup, labelSet, tooltipSet, initialValue) {
        var checkbox = parentGroup.add("checkbox", undefined, getLabel(labelSet));
        checkbox.value = initialValue;
        checkbox.helpTip = getLabel(tooltipSet);
        return checkbox;
    }

    /**
     * tooltip 付きのラジオボタンを追加します。
     *
     * @param {Object} parentGroup - 追加先のコンテナ。
     * @param {Object} labelSet - 表示名のラベル。
     * @param {Object} tooltipSet - tooltip のラベル。
     * @returns {RadioButton} 追加したラジオボタン。
     */
    function addRadio(parentGroup, labelSet, tooltipSet) {
        var radio = parentGroup.add("radiobutton", undefined, getLabel(labelSet));
        radio.helpTip = getLabel(tooltipSet);
        return radio;
    }

    /**
     * tooltip 付きの入力欄を追加します。
     *
     * @param {Object} parentGroup - 追加先のコンテナ。
     * @param {string} initialText - 初期値。
     * @param {number} characters - 入力欄の幅（文字数）。
     * @param {Object} tooltipSet - tooltip のラベル。
     * @returns {EditText} 追加した入力欄。
     */
    function addInput(parentGroup, initialText, characters, tooltipSet) {
        var inputField = parentGroup.add("edittext", undefined, initialText);
        inputField.characters = characters;
        inputField.helpTip = getLabel(tooltipSet);
        return inputField;
    }

    /**
     * 「項目名：入力欄［単位］」の行を追加します。
     *
     * @param {Object} parentGroup - 追加先のコンテナ。
     * @param {Object} labelSet - 項目名のラベル。
     * @param {string} initialText - 初期値。
     * @param {number} characters - 入力欄の幅（文字数）。
     * @param {Object} tooltipSet - tooltip のラベル。
     * @param {string} [unitText] - 入力欄の後ろに置く単位。
     * @returns {EditText} 追加した入力欄。
     */
    function addLabeledInput(parentGroup, labelSet, initialText, characters, tooltipSet, unitText) {
        var rowGroup = addRow(parentGroup);
        rowGroup.add("statictext", undefined, labelText(labelSet));
        var inputField = addInput(rowGroup, initialText, characters, tooltipSet);
        if (unitText) rowGroup.add("statictext", undefined, unitText);
        return inputField;
    }

    /**
     * プリセットの選択と書き出しの行を作成します。
     *
     * @param {Window} gridDialog - ダイアログ。
     * @param {Object} dialogUI - コントロールへの参照を書き込む先。
     * @returns {void}
     */
    function addPresetRow(gridDialog, dialogUI) {
        var presetRow = gridDialog.add("group");
        presetRow.orientation = "row";
        presetRow.alignment = ["center", "top"];
        presetRow.alignChildren = ["left", "center"];
        presetRow.add("statictext", undefined, labelText(LABELS.fieldLabel.preset));

        var presetDropdown = presetRow.add("dropdownlist");
        presetDropdown.helpTip = getLabel(LABELS.tooltip.preset);
        var autoIndex = 0;
        for (var presetKey in PRESET_DEFINITIONS) {
            if (!PRESET_DEFINITIONS.hasOwnProperty(presetKey)) continue;
            if (presetKey === "auto") autoIndex = presetDropdown.items.length;
            presetDropdown.add("item", presetKey)._key = presetKey;
        }
        presetDropdown.selection = autoIndex;

        var btnExport = presetRow.add("button", undefined, getLabel(LABELS.button.exportPreset));
        btnExport.helpTip = getLabel(LABELS.tooltip.exportPreset);

        dialogUI.presetDropdown = presetDropdown;
        dialogUI.btnExport = btnExport;
    }

    /**
     * ［共通設定とクリアスペース］パネルを作成します。
     *
     * @param {Window} gridDialog - ダイアログ。
     * @param {Object} dialogUI - コントロールへの参照を書き込む先。
     * @returns {void}
     */
    function addCommonPanel(gridDialog, dialogUI) {
        var commonPanel = setupPanel(gridDialog.add("panel", undefined, getLabel(LABELS.panel.common)), "fill");
        var commonColumns = addColumnsGroup(commonPanel);
        var commonLeftColumn = addColumn(commonColumns, "left");
        var commonRightColumn = addColumn(commonColumns, "left");

        /* 左：分割数・クリアスペース・ガイド化 / Left: divisions, clear space, guides */
        dialogUI.clearSpaceDivInput = addLabeledInput(commonLeftColumn, LABELS.fieldLabel.divisions, "4",
            FIELD_WIDTH.count, LABELS.tooltip.divisions);
        dialogUI.clearSpaceCheck = addCheckbox(addRow(commonLeftColumn),
            LABELS.checkbox.clearSpace, LABELS.tooltip.clearSpace, false);
        dialogUI.convertToGuidesCheck = addCheckbox(addRow(commonLeftColumn),
            LABELS.checkbox.convertToGuides, LABELS.tooltip.convertToGuides, false);

        /* 右：線幅・境界線の強調・グループ化 / Right: stroke width, emphasized bounds, grouping */
        dialogUI.strokeWidthInput = addLabeledInput(commonRightColumn, LABELS.fieldLabel.strokeWidth, "0.25",
            FIELD_WIDTH.strokeWidth, LABELS.tooltip.strokeWidth, getUnitInfo("strokeUnits").label);
        dialogUI.emphasizeBoundsCheck = addCheckbox(addRow(commonRightColumn),
            LABELS.checkbox.emphasizeBounds, LABELS.tooltip.emphasizeBounds, false);
        dialogUI.groupItemsCheck = addCheckbox(addRow(commonRightColumn),
            LABELS.checkbox.groupItems, LABELS.tooltip.groupItems, true);

        dialogUI.layerNameInput = addLabeledInput(commonPanel, LABELS.fieldLabel.targetLayer, getLabel(LABELS.itemName.layer),
            FIELD_WIDTH.layerName, LABELS.tooltip.targetLayer);
    }

    /**
     * ［横線］パネル（［ラインの決め方］を含む）を作成します。
     *
     * @param {Group} parentColumn - 追加先のカラム。
     * @param {Object} dialogUI - コントロールへの参照を書き込む先。
     * @returns {void}
     */
    function addHorizontalPanel(parentColumn, dialogUI) {
        var horizontalPanel = setupPanel(parentColumn.add("panel", undefined, getLabel(LABELS.panel.horizontal)));
        dialogUI.horizontalPanel = horizontalPanel;

        dialogUI.widthScaleInput = addLabeledInput(horizontalPanel, LABELS.fieldLabel.widthScale, "140",
            FIELD_WIDTH.scale, LABELS.tooltip.widthScale, "%");

        var extraHLinesRow = addRow(horizontalPanel);
        dialogUI.extraHLinesCheck = addCheckbox(extraHLinesRow,
            LABELS.checkbox.extraHorizontal, LABELS.tooltip.extraHorizontal, false);
        dialogUI.extraHLinesInput = addInput(extraHLinesRow, "1", FIELD_WIDTH.count, LABELS.tooltip.extraHorizontal);

        dialogUI.extendLeftCheck = addCheckbox(addRow(horizontalPanel),
            LABELS.checkbox.extendLeft, LABELS.tooltip.extendLeft, false);

        /* ラインの決め方 / Line generation method */
        var lineMethodPanel = setupPanel(horizontalPanel.add("panel", undefined, getLabel(LABELS.panel.lineMethod)));
        lineMethodPanel.helpTip = getLabel(LABELS.tooltip.lineMethod);

        dialogUI.methodNoneRadio = addRadio(lineMethodPanel, LABELS.radio.lineNone, LABELS.tooltip.lineNone);
        dialogUI.methodAutoRadio = addRadio(lineMethodPanel, LABELS.radio.lineAuto, LABELS.tooltip.lineAuto);
        dialogUI.methodAutoRadio.value = true;
        dialogUI.methodSegmentRadio = addRadio(lineMethodPanel, LABELS.radio.lineSegment, LABELS.tooltip.lineSegment);

        var evenLineCountRow = addRow(lineMethodPanel);
        dialogUI.methodEvenRadio = addRadio(evenLineCountRow, LABELS.radio.lineEven, LABELS.tooltip.lineEven);
        dialogUI.evenLineCountInput = addInput(evenLineCountRow, "4", FIELD_WIDTH.count, LABELS.tooltip.lineEven);
        evenLineCountRow.add("statictext", undefined, getLabel(LABELS.unit.lines));
    }

    /**
     * ［縦線］パネル（［オプション］を含む）を作成します。
     *
     * @param {Group} parentColumn - 追加先のカラム。
     * @param {Object} dialogUI - コントロールへの参照を書き込む先。
     * @returns {void}
     */
    function addVerticalPanel(parentColumn, dialogUI) {
        var verticalPanel = setupPanel(parentColumn.add("panel", undefined, getLabel(LABELS.panel.vertical)));
        dialogUI.verticalPanel = verticalPanel;

        dialogUI.heightScaleInput = addLabeledInput(verticalPanel, LABELS.fieldLabel.heightScale, "200",
            FIELD_WIDTH.scale, LABELS.tooltip.heightScale, "%");

        var outerVLinesRow = addRow(verticalPanel);
        dialogUI.outerVLinesCheck = addCheckbox(outerVLinesRow,
            LABELS.checkbox.outerVertical, LABELS.tooltip.outerVertical, false);
        dialogUI.outerVLinesInput = addInput(outerVLinesRow, "1", FIELD_WIDTH.count, LABELS.tooltip.outerVertical);

        dialogUI.extendUpCheck = addCheckbox(addRow(verticalPanel),
            LABELS.checkbox.extendUp, LABELS.tooltip.extendUp, false);

        /* オプション / Options */
        var verticalOptionsPanel = setupPanel(verticalPanel.add("panel", undefined, getLabel(LABELS.panel.options)));

        var columnDivRow = addRow(verticalOptionsPanel);
        dialogUI.columnDivCheck = addCheckbox(columnDivRow, LABELS.checkbox.columnDiv, LABELS.tooltip.columnDiv, false);
        dialogUI.columnDivInput = addInput(columnDivRow, "2", FIELD_WIDTH.count, LABELS.tooltip.columnDiv);
        dialogUI.columnDivInput.enabled = false;

        dialogUI.verticalElementsCheck = addCheckbox(addRow(verticalOptionsPanel),
            LABELS.checkbox.verticalElements, LABELS.tooltip.verticalElements, false);
        dialogUI.diagonalElementsCheck = addCheckbox(addRow(verticalOptionsPanel),
            LABELS.checkbox.diagonalElements, LABELS.tooltip.diagonalElements, false);
    }

    /**
     * ［プレビュー］とOK／キャンセルのボタン行を作成します。
     *
     * @param {Window} gridDialog - ダイアログ。
     * @param {Object} dialogUI - コントロールへの参照を書き込む先。
     * @returns {void}
     */
    function addButtonRow(gridDialog, dialogUI) {
        var btnRowGroup = gridDialog.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.alignment = ["fill", "bottom"];

        var btnLeftGroup = btnRowGroup.add("group");
        btnLeftGroup.orientation = "row";
        btnLeftGroup.alignChildren = ["left", "center"];
        dialogUI.previewCheck = addCheckbox(btnLeftGroup, LABELS.checkbox.preview, LABELS.tooltip.preview, true);

        var spacer = btnRowGroup.add("group");
        spacer.alignment = ["fill", "fill"];
        spacer.minimumSize.width = 0;

        var btnRightGroup = btnRowGroup.add("group");
        btnRightGroup.orientation = "row";
        btnRightGroup.alignChildren = ["right", "center"];
        dialogUI.btnCancel = btnRightGroup.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });
        dialogUI.btnOK = btnRightGroup.add("button", undefined, getLabel(LABELS.button.ok), { name: "ok" });
    }

    /**
     * 設定用のダイアログを作成します。
     *
     * @returns {Object} ダイアログ（dialog）と各コントロールへの参照。
     */
    function buildDialog() {
        var gridDialog = new Window("dialog", getLabel(LABELS.dialog.title) + " " + SCRIPT_VERSION);
        gridDialog.orientation = "column";
        gridDialog.alignChildren = ["fill", "top"];

        var dialogUI = { dialog: gridDialog };
        addPresetRow(gridDialog, dialogUI);
        addCommonPanel(gridDialog, dialogUI);

        /* 横線・縦線の2カラム / Two columns for the horizontal and vertical lines */
        var panelColumns = addColumnsGroup(gridDialog);
        addHorizontalPanel(addColumn(panelColumns, "fill"), dialogUI);
        addVerticalPanel(addColumn(panelColumns, "fill"), dialogUI);

        addButtonRow(gridDialog, dialogUI);
        return dialogUI;
    }

    // =========================================
    // UIの状態 / UI state
    // =========================================

    /**
     * ［グループ化］の使用可否を、［ガイドに変換する］に合わせます。
     *
     * @param {Object} dialogUI - buildDialog() が返すUI参照。
     * @returns {void}
     */
    function updateGroupState(dialogUI) {
        dialogUI.groupItemsCheck.enabled = !dialogUI.convertToGuidesCheck.value;
        if (dialogUI.convertToGuidesCheck.value) dialogUI.groupItemsCheck.value = false;
    }

    /**
     * ［均等分割］の本数入力の使用可否を更新します。
     *
     * @param {Object} dialogUI - buildDialog() が返すUI参照。
     * @returns {void}
     */
    function updateColumnDivField(dialogUI) {
        dialogUI.columnDivInput.enabled = dialogUI.columnDivCheck.enabled && dialogUI.columnDivCheck.value;
    }

    /**
     * ［分割数］の値と使用可否を、横線の均等分割に合わせます。
     *
     * @param {Object} dialogUI - buildDialog() が返すUI参照。
     * @returns {void}
     */
    function updateClearSpaceDivField(dialogUI) {
        var evenMode = dialogUI.methodEvenRadio.value;
        var evenCount = parseInt(dialogUI.evenLineCountInput.text, 10);

        /* 均等分割の本数＋1がユニット数になる / The unit count follows the even division count */
        if (evenMode && !isNaN(evenCount) && evenCount >= 1) {
            dialogUI.clearSpaceDivInput.text = String(evenCount + 1);
        }
        dialogUI.clearSpaceDivInput.enabled = dialogUI.clearSpaceCheck.value || !evenMode;
    }

    /**
     * 分割にかかわる入力欄の状態をまとめて更新します。
     *
     * @param {Object} dialogUI - buildDialog() が返すUI参照。
     * @returns {void}
     */
    function syncDivisionFields(dialogUI) {
        updateColumnDivField(dialogUI);
        updateClearSpaceDivField(dialogUI);
    }

    /**
     * ［クリアスペース］の状態に合わせて、横線・縦線パネルなどの使用可否を更新します。
     *
     * @param {Object} dialogUI - buildDialog() が返すUI参照。
     * @returns {void}
     */
    function updateClearSpaceState(dialogUI) {
        var clearSpace = dialogUI.clearSpaceCheck.value;

        /* 横線・縦線はパネルごとディム / Dim the whole horizontal and vertical panels */
        dialogUI.horizontalPanel.enabled = !clearSpace;
        dialogUI.verticalPanel.enabled = !clearSpace;
        dialogUI.extraHLinesInput.enabled = !clearSpace && dialogUI.extraHLinesCheck.value;
        dialogUI.evenLineCountInput.enabled = !clearSpace && dialogUI.methodEvenRadio.value;
        dialogUI.outerVLinesInput.enabled = !clearSpace && dialogUI.outerVLinesCheck.value;
        dialogUI.columnDivInput.enabled = !clearSpace && dialogUI.columnDivCheck.value;

        /* クリアスペース中は線幅とガイド化を使わず、グループ化はON固定 */
        dialogUI.strokeWidthInput.enabled = !clearSpace;
        dialogUI.convertToGuidesCheck.enabled = !clearSpace;
        if (clearSpace) {
            dialogUI.convertToGuidesCheck.value = false;
            dialogUI.groupItemsCheck.value = true;
            dialogUI.groupItemsCheck.enabled = false;
        } else {
            updateGroupState(dialogUI);
        }

        syncDivisionFields(dialogUI);
    }

    // =========================================
    // 入力値 / Parameters
    // =========================================

    /**
     * 伸張率の入力欄か判定します（小数第1位で表示する欄）。
     *
     * @param {Object} dialogUI - buildDialog() が返すUI参照。
     * @param {EditText} inputField - 対象の入力欄。
     * @returns {boolean} 伸張率の入力欄のとき true。
     */
    function isScaleField(dialogUI, inputField) {
        return inputField === dialogUI.widthScaleInput || inputField === dialogUI.heightScaleInput;
    }

    /**
     * 正の数として扱える入力か判定します。
     *
     * @param {string} inputText - 入力された文字列。
     * @returns {boolean} 正の数のとき true。
     */
    function isPositiveNumberText(inputText) {
        return /^\d+(?:\.\d+)?$/.test(inputText) && parseFloat(inputText) > 0;
    }

    /**
     * 1以上の整数として扱える入力か判定します。
     *
     * @param {string} inputText - 入力された文字列。
     * @returns {boolean} 1以上の整数のとき true。
     */
    function isCountText(inputText) {
        return /^\d+$/.test(inputText) && parseInt(inputText, 10) >= 1;
    }

    /**
     * ダイアログの入力内容をそのまま読み取ります。
     *
     * @param {Object} dialogUI - buildDialog() が返すUI参照。
     * @returns {Object} 入力内容。
     */
    function readDialogValues(dialogUI) {
        return {
            layerName: normalizeLayerNameText(dialogUI.layerNameInput.text),
            widthScaleText: dialogUI.widthScaleInput.text,
            heightScaleText: dialogUI.heightScaleInput.text,
            strokeWidthText: dialogUI.strokeWidthInput.text,
            lineMethod: getLineMethod(dialogUI),
            evenLineCountText: dialogUI.evenLineCountInput.text,
            extraHLines: dialogUI.extraHLinesCheck.value,
            extraHLineCountText: dialogUI.extraHLinesInput.text,
            extendLeft: dialogUI.extendLeftCheck.value,
            outerVLines: dialogUI.outerVLinesCheck.value,
            outerVLineCountText: dialogUI.outerVLinesInput.text,
            extendUp: dialogUI.extendUpCheck.value,
            columnDiv: dialogUI.columnDivCheck.value,
            columnDivisionsText: dialogUI.columnDivInput.text,
            verticalElements: dialogUI.verticalElementsCheck.value,
            diagonalElements: dialogUI.diagonalElementsCheck.value,
            emphasizeBounds: dialogUI.emphasizeBoundsCheck.value,
            convertToGuides: dialogUI.convertToGuidesCheck.value,
            groupItems: dialogUI.groupItemsCheck.value,
            clearSpace: dialogUI.clearSpaceCheck.value,
            clearSpaceDivisionsText: dialogUI.clearSpaceDivInput.text
        };
    }

    /**
     * 入力内容が処理に使える値か検証します。
     *
     * @param {Object} dialogValues - readDialogValues() の結果。
     * @returns {boolean} すべて正しいとき true。
     */
    function validateDialogValues(dialogValues) {
        if (!dialogValues.layerName) return false;
        if (!isPositiveNumberText(dialogValues.widthScaleText)) return false;
        if (!isPositiveNumberText(dialogValues.heightScaleText)) return false;
        if (!isPositiveNumberText(dialogValues.strokeWidthText)) return false;
        if (!isPositiveNumberText(dialogValues.clearSpaceDivisionsText)) return false;
        if (dialogValues.lineMethod === "even" && !isCountText(dialogValues.evenLineCountText)) return false;
        if (dialogValues.extraHLines && !isCountText(dialogValues.extraHLineCountText)) return false;
        if (dialogValues.outerVLines && !isCountText(dialogValues.outerVLineCountText)) return false;
        if (dialogValues.columnDiv && !isCountText(dialogValues.columnDivisionsText)) return false;
        return true;
    }

    /**
     * 入力内容を、描画で使う単位・型に変換します。
     *
     * @param {Object} dialogValues - readDialogValues() の結果。
     * @returns {Object|null} 変換した設定。線幅を換算できないときは null。
     */
    function toGridParams(dialogValues) {
        var strokeWidth = strokeUnitsToPoints(parseFloat(dialogValues.strokeWidthText));
        if (!(strokeWidth > 0)) return null;

        return {
            layerName: dialogValues.layerName,
            widthScale: parseFloat(dialogValues.widthScaleText) / 100,
            heightScale: parseFloat(dialogValues.heightScaleText) / 100,
            strokeWidth: strokeWidth,
            lineMethod: dialogValues.lineMethod,
            evenLineCount: (dialogValues.lineMethod === "even") ? parseInt(dialogValues.evenLineCountText, 10) : 0,
            extraHLines: dialogValues.extraHLines,
            extraHLineCount: dialogValues.extraHLines ? parseInt(dialogValues.extraHLineCountText, 10) : 0,
            extendLeft: dialogValues.extendLeft,
            outerVLines: dialogValues.outerVLines,
            outerVLineCount: dialogValues.outerVLines ? parseInt(dialogValues.outerVLineCountText, 10) : 0,
            extendUp: dialogValues.extendUp,
            columnDiv: dialogValues.columnDiv,
            columnDivisions: dialogValues.columnDiv ? parseInt(dialogValues.columnDivisionsText, 10) : 0,
            verticalElements: dialogValues.verticalElements,
            diagonalElements: dialogValues.diagonalElements,
            emphasizeBounds: dialogValues.emphasizeBounds,
            convertToGuides: dialogValues.convertToGuides,
            groupItems: dialogValues.groupItems,
            clearSpace: dialogValues.clearSpace,
            clearSpaceDivisions: parseFloat(dialogValues.clearSpaceDivisionsText)
        };
    }

    /**
     * ダイアログの入力内容を、検証したうえで設定にまとめます。
     *
     * @param {Object} dialogUI - buildDialog() が返すUI参照。
     * @returns {Object|null} 設定。入力値が不正なときは null。
     */
    function getGridParams(dialogUI) {
        var dialogValues = readDialogValues(dialogUI);
        return validateDialogValues(dialogValues) ? toGridParams(dialogValues) : null;
    }

    // =========================================
    // 描画 / Drawing
    // =========================================

    /**
     * 補助線を描く範囲（線の長さとユニットの大きさ）を求めます。
     *
     * @param {Object} gridContext - buildGridContext() が返すコンテキスト。
     * @param {Object} gridParams - getGridParams() が返す設定。
     * @returns {{left: number, right: number, top: number, bottom: number, unitSize: number}} 描画範囲。
     */
    function computeFrame(gridContext, gridParams) {
        var unitSize = gridContext.selHeight / gridParams.clearSpaceDivisions;

        var halfWidth = gridContext.selWidth * gridParams.widthScale / 2;
        var left = gridContext.centerX - halfWidth;
        var right = gridContext.centerX + halfWidth;
        if (gridParams.extendLeft) {
            /* 右端を選択範囲の右外に固定し、余りを左へ伸ばす / Pin the right end and extend leftward */
            var shift = right - (gridContext.selRight + gridContext.selHeight / 2);
            left -= shift;
            right -= shift;
        }

        var lineHeight = gridContext.selHeight * gridParams.heightScale;
        var top, bottom;
        if (gridParams.extendUp) {
            bottom = gridContext.selBottom - unitSize * 2;
            top = bottom + lineHeight;
        } else {
            top = gridContext.centerY + lineHeight / 2;
            bottom = gridContext.centerY - lineHeight / 2;
        }

        if (!gridParams.clearSpace) {
            /* 外側に追加した線とも必ず交差するよう、線の長さを自動で補正する */
            var outerOffset = gridParams.outerVLines ? unitSize * (gridParams.outerVLineCount - 1) : 0;
            var extraOffset = gridParams.extraHLines ? unitSize * (gridParams.extraHLineCount - 1) : 0;
            left = Math.min(left, gridContext.selLeft - outerOffset);
            right = Math.max(right, gridContext.selRight + outerOffset);
            top = Math.max(top, gridContext.selTop + extraOffset);
            bottom = Math.min(bottom, gridContext.selBottom - extraOffset);
        }

        return { left: left, right: right, top: top, bottom: bottom, unitSize: unitSize };
    }

    /**
     * 直線を1本描きます。
     *
     * @param {GroupItem} targetGroup - 追加先のグループ。
     * @param {number} x1 - 始点のX座標。
     * @param {number} y1 - 始点のY座標。
     * @param {number} x2 - 終点のX座標。
     * @param {number} y2 - 終点のY座標。
     * @param {Object} strokeColor - 線のカラー。
     * @param {number} strokeWidth - 線幅（pt）。
     * @returns {PathItem} 作成したパス。
     */
    function drawLine(targetGroup, x1, y1, x2, y2, strokeColor, strokeWidth) {
        var linePath = targetGroup.pathItems.add();
        linePath.setEntirePath([[x1, y1], [x2, y2]]);
        linePath.stroked = true;
        linePath.filled = false;
        linePath.strokeWidth = strokeWidth;
        linePath.strokeColor = strokeColor;
        return linePath;
    }

    /**
     * 横線を1本描きます。
     *
     * @param {GroupItem} targetGroup - 追加先のグループ。
     * @param {Object} frame - computeFrame() が返す描画範囲。
     * @param {number} y - 描くY座標。
     * @param {Object} strokeStyle - 線のカラーと線幅。
     * @param {boolean} onBounds - 選択範囲の境界に重なるか。
     * @returns {void}
     */
    function drawHorizontalLine(targetGroup, frame, y, strokeStyle, onBounds) {
        drawLine(targetGroup, frame.left, y, frame.right, y, strokeStyle.color, onBounds ? strokeStyle.bounds : strokeStyle.normal);
    }

    /**
     * 縦線を1本描きます。
     *
     * @param {GroupItem} targetGroup - 追加先のグループ。
     * @param {Object} frame - computeFrame() が返す描画範囲。
     * @param {number} x - 描くX座標。
     * @param {Object} strokeStyle - 線のカラーと線幅。
     * @param {boolean} onBounds - 選択範囲の境界に重なるか。
     * @returns {void}
     */
    function drawVerticalLine(targetGroup, frame, x, strokeStyle, onBounds) {
        drawLine(targetGroup, x, frame.top, x, frame.bottom, strokeStyle.color, onBounds ? strokeStyle.bounds : strokeStyle.normal);
    }

    /**
     * 横線（境界線・外側の線・ラインの決め方による線）を描きます。
     *
     * @param {GroupItem} linesGroup - 追加先のグループ。
     * @param {Object} gridContext - buildGridContext() が返すコンテキスト。
     * @param {Object} gridParams - getGridParams() が返す設定。
     * @param {Object} frame - computeFrame() が返す描画範囲。
     * @param {Object} strokeStyle - 線のカラーと線幅。
     * @returns {void}
     */
    function drawHorizontalGuides(linesGroup, gridContext, gridParams, frame, strokeStyle) {
        if (gridParams.extraHLines) {
            drawHorizontalLine(linesGroup, frame, gridContext.selTop, strokeStyle, true);
            drawHorizontalLine(linesGroup, frame, gridContext.selBottom, strokeStyle, true);
            for (var i = 1; i < gridParams.extraHLineCount; i++) {
                drawHorizontalLine(linesGroup, frame, gridContext.selTop + frame.unitSize * i, strokeStyle, false);
                drawHorizontalLine(linesGroup, frame, gridContext.selBottom - frame.unitSize * i, strokeStyle, false);
            }
        }

        var lineYs = getHorizontalLineYs(gridContext.geometry, gridParams.lineMethod, gridParams.evenLineCount);
        for (var j = 0; j < lineYs.length; j++) {
            var onBounds = gridParams.emphasizeBounds &&
                (isSameCoord(lineYs[j], gridContext.selTop) || isSameCoord(lineYs[j], gridContext.selBottom));
            drawHorizontalLine(linesGroup, frame, lineYs[j], strokeStyle, onBounds);
        }
    }

    /**
     * 縦線（境界線・外側の線・均等分割・垂直／斜線エレメント）を描きます。
     *
     * @param {GroupItem} linesGroup - 追加先のグループ。
     * @param {Object} gridContext - buildGridContext() が返すコンテキスト。
     * @param {Object} gridParams - getGridParams() が返す設定。
     * @param {Object} frame - computeFrame() が返す描画範囲。
     * @param {Object} strokeStyle - 線のカラーと線幅。
     * @returns {void}
     */
    function drawVerticalGuides(linesGroup, gridContext, gridParams, frame, strokeStyle) {
        if (gridParams.outerVLines) {
            drawVerticalLine(linesGroup, frame, gridContext.selLeft, strokeStyle, true);
            drawVerticalLine(linesGroup, frame, gridContext.selRight, strokeStyle, true);
            for (var i = 1; i < gridParams.outerVLineCount; i++) {
                drawVerticalLine(linesGroup, frame, gridContext.selLeft - frame.unitSize * i, strokeStyle, false);
                drawVerticalLine(linesGroup, frame, gridContext.selRight + frame.unitSize * i, strokeStyle, false);
            }
        }

        if (gridParams.columnDiv && gridParams.columnDivisions > 1) {
            var step = gridContext.selWidth / gridParams.columnDivisions;
            for (var j = 1; j < gridParams.columnDivisions; j++) {
                drawVerticalLine(linesGroup, frame, gridContext.selLeft + step * j, strokeStyle, false);
            }
        }

        if (gridParams.verticalElements) {
            var elementXs = getVerticalElementXs(gridContext.geometry);
            for (var k = 0; k < elementXs.length; k++) {
                var onBounds = gridParams.emphasizeBounds &&
                    (isSameCoord(elementXs[k], gridContext.selLeft) || isSameCoord(elementXs[k], gridContext.selRight));
                drawVerticalLine(linesGroup, frame, elementXs[k], strokeStyle, onBounds);
            }
        }

        if (gridParams.diagonalElements) {
            var diagonalLines = getDiagonalElementLines(gridContext.geometry, frame.left, frame.top, frame.right, frame.bottom);
            for (var lineIndex = 0; lineIndex < diagonalLines.length; lineIndex++) {
                drawLine(linesGroup,
                    diagonalLines[lineIndex][0][0], diagonalLines[lineIndex][0][1],
                    diagonalLines[lineIndex][1][0], diagonalLines[lineIndex][1][1],
                    strokeStyle.color, strokeStyle.normal);
            }
        }
    }

    /**
     * クリアスペースの帯を1つ描きます（塗りと枠線の2枚重ね）。
     *
     * @param {GroupItem} targetGroup - 追加先のグループ。
     * @param {number} left - 左端。
     * @param {number} top - 上端。
     * @param {number} width - 幅。
     * @param {number} height - 高さ。
     * @param {Object} bandColor - 帯のカラー。
     * @returns {void}
     */
    function drawClearSpaceBand(targetGroup, left, top, width, height, bandColor) {
        var fillRect = targetGroup.pathItems.rectangle(top, left, width, height);
        fillRect.filled = true;
        fillRect.fillColor = bandColor;
        fillRect.stroked = false;
        fillRect.opacity = CLEAR_SPACE_STYLE.fillOpacity;

        var outlineRect = targetGroup.pathItems.rectangle(top, left, width, height);
        outlineRect.filled = false;
        outlineRect.stroked = true;
        outlineRect.strokeColor = bandColor;
        outlineRect.strokeWidth = CLEAR_SPACE_STYLE.strokeWidth;
    }

    /**
     * 選択範囲の四方にクリアスペースを描きます。
     *
     * @param {GroupItem} linesGroup - 追加先のグループ。
     * @param {Object} gridContext - buildGridContext() が返すコンテキスト。
     * @param {number} unitSize - ユニット（1単位）の大きさ。
     * @returns {void}
     */
    function drawClearSpace(linesGroup, gridContext, unitSize) {
        var left = gridContext.selLeft - unitSize;
        var right = gridContext.selRight + unitSize;
        var top = gridContext.selTop + unitSize;
        var outerWidth = right - left;
        var outerHeight = top - (gridContext.selBottom - unitSize);
        var bandColor = makeCyanColor(CLEAR_SPACE_STYLE.cyan);

        var clearSpaceGroup = linesGroup.groupItems.add();
        clearSpaceGroup.name = getLabel(LABELS.itemName.clearSpaceGroup);

        drawClearSpaceBand(clearSpaceGroup, left, top, outerWidth, unitSize, bandColor);
        drawClearSpaceBand(clearSpaceGroup, left, gridContext.selBottom, outerWidth, unitSize, bandColor);
        drawClearSpaceBand(clearSpaceGroup, left, top, unitSize, outerHeight, bandColor);
        drawClearSpaceBand(clearSpaceGroup, gridContext.selRight, top, unitSize, outerHeight, bandColor);
    }

    /**
     * 作成した線をガイドに変換し、グループ化しない設定ならレイヤー直下に移します。
     *
     * @param {GroupItem} linesGroup - 作成した線のグループ。
     * @param {Object} gridContext - buildGridContext() が返すコンテキスト。
     * @param {Object} gridParams - getGridParams() が返す設定。
     * @param {boolean} isPreview - プレビューとして作成したか。
     * @returns {GroupItem|null} 残ったグループ。解除したときは null。
     */
    function finalizeLinesGroup(linesGroup, gridContext, gridParams, isPreview) {
        /* プレビューは削除できるようグループのまま残す / Keep previews grouped so they can be removed */
        if (isPreview) return linesGroup;

        if (gridParams.convertToGuides) {
            for (var i = 0; i < linesGroup.pathItems.length; i++) {
                linesGroup.pathItems[i].guides = true;
            }
        }

        if (!gridParams.groupItems && !gridParams.convertToGuides) {
            while (linesGroup.pageItems.length > 0) {
                linesGroup.pageItems[0].move(gridContext.guideLayer, ElementPlacement.PLACEATBEGINNING);
            }
            linesGroup.remove();
            return null;
        }

        return linesGroup;
    }

    /**
     * 設定に従って補助線またはクリアスペースを作成します。
     *
     * @param {Object} gridContext - buildGridContext() が返すコンテキスト。
     * @param {Object} gridParams - getGridParams() が返す設定。
     * @param {boolean} isPreview - プレビューとして作成するか。
     * @returns {GroupItem|null} 作成したグループ。グループ化しないときは null。
     */
    function createGridItems(gridContext, gridParams, isPreview) {
        var frame = computeFrame(gridContext, gridParams);
        var strokeStyle = {
            color: makeGrayColor(LINE_STYLE.grayTint),
            normal: gridParams.strokeWidth,
            bounds: gridParams.emphasizeBounds ? (gridParams.strokeWidth * LINE_STYLE.boundsMultiplier) : gridParams.strokeWidth
        };

        var linesGroup = gridContext.guideLayer.groupItems.add();
        linesGroup.name = getLabel(isPreview ? LABELS.itemName.previewGroup : LABELS.itemName.guideGroup);

        if (gridParams.clearSpace) {
            drawClearSpace(linesGroup, gridContext, frame.unitSize);
        } else {
            drawHorizontalGuides(linesGroup, gridContext, gridParams, frame, strokeStyle);
            drawVerticalGuides(linesGroup, gridContext, gridParams, frame, strokeStyle);
        }

        return finalizeLinesGroup(linesGroup, gridContext, gridParams, isPreview);
    }

    // =========================================
    // プレビュー / Preview
    // =========================================

    var previewGroup = null;

    /**
     * プレビューで作成したオブジェクトを削除します。
     *
     * @returns {void}
     */
    function removePreview() {
        if (!previewGroup) return;
        /* 削除済みのときは参照が無効 / The group may already be gone */
        try {
            previewGroup.remove();
        } catch (e) { }
        previewGroup = null;
        app.redraw();
    }

    /**
     * 現在の設定でプレビューを作り直します。
     *
     * @param {Object} dialogUI - buildDialog() が返すUI参照。
     * @param {Object} gridContext - buildGridContext() が返すコンテキスト。
     * @returns {void}
     */
    function updatePreview(dialogUI, gridContext) {
        removePreview();
        var gridParams = getGridParams(dialogUI);
        if (!gridParams) return;
        gridContext.guideLayer = ensureGuideLayer(gridContext, gridParams.layerName);
        previewGroup = createGridItems(gridContext, gridParams, true);
        app.redraw();
    }

    /**
     * ［プレビュー］がONのときだけ、プレビューを作り直します。
     *
     * @param {Object} dialogUI - buildDialog() が返すUI参照。
     * @param {Object} gridContext - buildGridContext() が返すコンテキスト。
     * @returns {void}
     */
    function refreshPreview(dialogUI, gridContext) {
        if (dialogUI.previewCheck.value) updatePreview(dialogUI, gridContext);
    }

    // =========================================
    // イベント / Events
    // =========================================

    /**
     * 入力欄で↑↓キーによる増減を有効にします（shiftで10、optionで0.1刻み）。
     *
     * @param {EditText} editText - 対象の入力欄。
     * @param {boolean} integerOnly - 整数だけを扱うか。
     * @param {Object} dialogUI - buildDialog() が返すUI参照。
     * @param {Object} gridContext - buildGridContext() が返すコンテキスト。
     * @returns {void}
     */
    function changeValueByArrowKey(editText, integerOnly, dialogUI, gridContext) {
        editText.addEventListener("keydown", function (event) {
            if (event.keyName !== "Up" && event.keyName !== "Down") return;

            var value = Number(editText.text);
            if (isNaN(value)) return;

            var isUp = (event.keyName === "Up");
            var keyboardState = ScriptUI.environment.keyboardState;
            if (keyboardState.shiftKey) {
                value = isUp ? (Math.ceil((value + 1) / 10) * 10) : (Math.floor((value - 1) / 10) * 10);
            } else if (keyboardState.altKey && !integerOnly) {
                value += isUp ? 0.1 : -0.1;
            } else {
                value += isUp ? 1 : -1;
            }

            if (value < 0) value = 0;
            value = (keyboardState.altKey && !integerOnly) ? (Math.round(value * 10) / 10) : Math.round(value);

            event.preventDefault();
            editText.text = isScaleField(dialogUI, editText) ? value.toFixed(1) : String(value);

            if (editText === dialogUI.evenLineCountInput || editText === dialogUI.columnDivInput) {
                syncDivisionFields(dialogUI);
            }
            refreshPreview(dialogUI, gridContext);
        });
    }

    /**
     * ↑↓キーで増減できる入力欄をまとめて設定します。
     *
     * @param {Object} dialogUI - buildDialog() が返すUI参照。
     * @param {Object} gridContext - buildGridContext() が返すコンテキスト。
     * @returns {void}
     */
    function bindArrowKeyHandlers(dialogUI, gridContext) {
        changeValueByArrowKey(dialogUI.widthScaleInput, false, dialogUI, gridContext);
        changeValueByArrowKey(dialogUI.heightScaleInput, false, dialogUI, gridContext);
        changeValueByArrowKey(dialogUI.strokeWidthInput, false, dialogUI, gridContext);
        changeValueByArrowKey(dialogUI.clearSpaceDivInput, false, dialogUI, gridContext);
        changeValueByArrowKey(dialogUI.columnDivInput, true, dialogUI, gridContext);
        changeValueByArrowKey(dialogUI.extraHLinesInput, true, dialogUI, gridContext);
        changeValueByArrowKey(dialogUI.outerVLinesInput, true, dialogUI, gridContext);
        changeValueByArrowKey(dialogUI.evenLineCountInput, true, dialogUI, gridContext);
    }

    /**
     * 伸張率の入力欄を、フォーカスが外れたときに小数第1位の表記にそろえます。
     *
     * @param {EditText} scaleField - 対象の入力欄。
     * @returns {void}
     */
    function bindScaleFieldBlur(scaleField) {
        scaleField.onBlur = function () {
            var value = parseFloat(scaleField.text);
            if (!isNaN(value)) scaleField.text = value.toFixed(1);
        };
    }

    /**
     * ダイアログのイベントハンドラーを設定します。
     *
     * @param {Object} dialogUI - buildDialog() が返すUI参照。
     * @param {Object} gridContext - buildGridContext() が返すコンテキスト。
     * @returns {void}
     */
    function bindEvents(dialogUI, gridContext) {
        dialogUI.presetDropdown.onChange = function () {
            applyPreset(dialogUI, gridContext);
        };

        dialogUI.btnExport.onClick = function () {
            exportPreset(dialogUI);
        };

        dialogUI.previewCheck.onClick = function () {
            if (dialogUI.previewCheck.value) {
                updatePreview(dialogUI, gridContext);
            } else {
                removePreview();
            }
        };

        dialogUI.widthScaleInput.onChanging = dialogUI.heightScaleInput.onChanging =
            dialogUI.strokeWidthInput.onChanging = dialogUI.columnDivInput.onChanging =
            dialogUI.extraHLinesInput.onChanging = dialogUI.outerVLinesInput.onChanging =
            dialogUI.clearSpaceDivInput.onChanging = function () {
                refreshPreview(dialogUI, gridContext);
            };

        dialogUI.evenLineCountInput.onChanging = function () {
            syncDivisionFields(dialogUI);
            refreshPreview(dialogUI, gridContext);
        };

        bindScaleFieldBlur(dialogUI.widthScaleInput);
        bindScaleFieldBlur(dialogUI.heightScaleInput);

        dialogUI.layerNameInput.onBlur = function () {
            dialogUI.layerNameInput.text = normalizeLayerNameText(dialogUI.layerNameInput.text);
        };

        dialogUI.methodNoneRadio.onClick = dialogUI.methodAutoRadio.onClick =
            dialogUI.methodSegmentRadio.onClick = dialogUI.methodEvenRadio.onClick = function () {
                dialogUI.evenLineCountInput.enabled = dialogUI.methodEvenRadio.value;
                syncDivisionFields(dialogUI);
                refreshPreview(dialogUI, gridContext);
            };

        dialogUI.extraHLinesCheck.onClick = dialogUI.extendLeftCheck.onClick =
            dialogUI.outerVLinesCheck.onClick = dialogUI.columnDivCheck.onClick =
            dialogUI.extendUpCheck.onClick = dialogUI.verticalElementsCheck.onClick =
            dialogUI.diagonalElementsCheck.onClick = dialogUI.groupItemsCheck.onClick = function () {
                dialogUI.extraHLinesInput.enabled = dialogUI.extraHLinesCheck.enabled && dialogUI.extraHLinesCheck.value;
                dialogUI.outerVLinesInput.enabled = dialogUI.outerVLinesCheck.enabled && dialogUI.outerVLinesCheck.value;
                syncDivisionFields(dialogUI);
                refreshPreview(dialogUI, gridContext);
            };

        dialogUI.clearSpaceCheck.onClick = function () {
            updateClearSpaceState(dialogUI);
            refreshPreview(dialogUI, gridContext);
        };

        dialogUI.emphasizeBoundsCheck.onClick = function () {
            refreshPreview(dialogUI, gridContext);
        };

        dialogUI.convertToGuidesCheck.onClick = function () {
            updateGroupState(dialogUI);
            refreshPreview(dialogUI, gridContext);
        };

        dialogUI.btnOK.onClick = function () {
            removePreview();
            var gridParams = getGridParams(dialogUI);
            if (!gridParams) {
                alert(getLabel(LABELS.alert.invalidInput));
                return;
            }
            gridContext.guideLayer = ensureGuideLayer(gridContext, gridParams.layerName);
            createGridItems(gridContext, gridParams, false);
            dialogUI.dialog.close(1);
        };

        dialogUI.btnCancel.onClick = function () {
            removePreview();
            dialogUI.dialog.close(2);
        };

        dialogUI.dialog.onClose = function () {
            removePreview();
            removeEmptyCreatedLayers(gridContext);
        };
    }

    // =========================================
    // メイン / Main
    // =========================================

    /**
     * コーナーウィジェットの表示を切り替えます（ドラッグ中の誤操作と表示の乱れを避けるため）。
     *
     * @returns {void}
     */
    function toggleLiveCornerAnnotator() {
        try {
            app.executeMenuCommand('Live Corner Annotator');
        } catch (e) { }
    }

    /**
     * ドキュメントと選択を確認し、ダイアログを表示します。
     *
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel(LABELS.alert.noDocument));
            return;
        }

        var doc = app.activeDocument;
        if (doc.selection.length === 0) {
            alert(getLabel(LABELS.alert.noSelection));
            return;
        }

        var gridContext = buildGridContext(doc, doc.selection);
        if (!gridContext) return;

        previewGroup = null;
        toggleLiveCornerAnnotator();
        try {
            var dialogUI = buildDialog();
            bindArrowKeyHandlers(dialogUI, gridContext);
            bindEvents(dialogUI, gridContext);
            applyPreset(dialogUI, gridContext);
            dialogUI.dialog.show();
        } finally {
            toggleLiveCornerAnnotator();
        }
    }

    main();

})();
