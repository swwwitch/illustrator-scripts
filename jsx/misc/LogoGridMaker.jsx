#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択オブジェクトの visibleBounds を基準に、ロゴ用の補助線およびクリアスペースを生成します。
「文字形状からグリッド構造を抽出する」ことを目的としたスクリプトです。

詳細は README を参照してください。

### Overview

Generates construction lines and clear space for a logo, based on the visibleBounds of the selection.
The script is aimed at extracting a grid structure from letterforms.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "LogoGridMaker";                /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.4.3";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-04-10";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-19";                   /* 更新日 / last updated */

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
    function getCurrentLang() {
        return ($.locale && $.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = getCurrentLang();

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

    /* 線の単位コード→ラベルとpt換算係数（コード5はH=0.25mm） / Stroke unit code to label and points factor */
    var STROKE_UNITS = {
        0: { label: "in", factor: 72.0 },
        1: { label: "mm", factor: 72.0 / 25.4 },
        2: { label: "pt", factor: 1.0 },
        3: { label: "pica", factor: 12.0 },
        4: { label: "cm", factor: 72.0 / 2.54 },
        5: { label: "H", factor: 72.0 / 25.4 * 0.25 },
        6: { label: "px", factor: 1.0 },
        7: { label: "ft/in", factor: 72.0 * 12.0 },
        8: { label: "m", factor: 72.0 / 25.4 * 1000.0 },
        9: { label: "yd", factor: 72.0 * 36.0 },
        10: { label: "ft", factor: 72.0 * 12.0 }
    };

    /**
     * 環境設定の［線］の単位を取得します（不明な単位は pt 扱い）。
     *
     * @returns {{label: string, factor: number}} 単位ラベルとpt換算係数。
     */
    function getStrokeUnit() {
        var code = app.preferences.getIntegerPreference("strokeUnits");
        return STROKE_UNITS[code] || STROKE_UNITS[2];
    }

    /**
     * 線の単位で入力された値をptに換算します。
     *
     * @param {number} value - 線の単位での値。
     * @returns {number} pt換算した値。
     */
    function toPoints(value) {
        return value * getStrokeUnit().factor;
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
        var current = null;
        var baseValue = 0;
        for (var i = 0; i < sortedItems.length; i++) {
            var value = getValue(sortedItems[i]);
            if (!current || Math.abs(baseValue - value) > tolerance) {
                current = [];
                baseValue = value;
                clusters.push(current);
            }
            current.push(sortedItems[i]);
        }
        return clusters;
    }

    /**
     * 重み付き平均を求めます。
     *
     * @param {Array<Object>} items - 対象の配列。
     * @param {function(Object):number} getValue - 平均する値を取り出す関数。
     * @param {function(Object):number} getWeight - 重みを取り出す関数。
     * @returns {number} 重み付き平均（重みの合計が0のときは0）。
     */
    function weightedAverage(items, getValue, getWeight) {
        var total = 0;
        var sum = 0;
        for (var i = 0; i < items.length; i++) {
            var weight = getWeight(items[i]);
            total += weight;
            sum += getValue(items[i]) * weight;
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
        var color = new GrayColor();
        color.gray = tint;
        return color;
    }

    /**
     * シアンのみのCMYKカラーを作成します。
     *
     * @param {number} tint - シアンの濃度（0〜100）。
     * @returns {CMYKColor} 作成したカラー。
     */
    function makeCyanColor(tint) {
        var color = new CMYKColor();
        color.cyan = tint;
        color.magenta = 0;
        color.yellow = 0;
        color.black = 0;
        return color;
    }

    /**
     * 補助線を作成するレイヤーを用意します。新しく作ったレイヤーは控えておきます。
     *
     * @param {Object} context - buildContext() が返すコンテキスト。
     * @param {string} name - レイヤー名。
     * @returns {Layer} 取得または作成したレイヤー。
     */
    function ensureGuideLayer(context, name) {
        for (var i = 0; i < context.doc.layers.length; i++) {
            if (context.doc.layers[i].name === name) {
                return context.doc.layers[i];
            }
        }
        var layer = context.doc.layers.add();
        layer.name = name;
        context.createdLayers.push(layer);
        return layer;
    }

    /**
     * このスクリプトが作ったレイヤーのうち、空のまま残ったものを削除します。
     *
     * @param {Object} context - buildContext() が返すコンテキスト。
     * @returns {void}
     */
    function removeEmptyCreatedLayers(context) {
        for (var i = context.createdLayers.length - 1; i >= 0; i--) {
            try {
                if (context.createdLayers[i].pageItems.length === 0) {
                    context.createdLayers[i].remove();
                }
            } catch (e) { }
        }
        context.createdLayers = [];
    }

    /**
     * レイヤー名の入力値から改行・タブと前後の空白を取り除きます。
     *
     * @param {string} text - 入力された文字列。
     * @returns {string} 整えた文字列。
     */
    function normalizeLayerNameText(text) {
        return String(text).replace(/[\r\n\t]+/g, " ").replace(/^\s+|\s+$/g, "");
    }

    // =========================================
    // 選択範囲とジオメトリ / Selection and geometry
    // =========================================

    /**
     * オブジェクトの visibleBounds を取得します（取得できないときは null）。
     *
     * @param {PageItem} item - 対象のオブジェクト。
     * @returns {Array<number>|null} [左, 上, 右, 下] または null。
     */
    function getItemBounds(item) {
        try {
            return item.visibleBounds;
        } catch (e) {
            return null;
        }
    }

    /**
     * 選択範囲全体の visibleBounds を求めます。
     *
     * @param {Array<PageItem>} items - 選択中のオブジェクト。
     * @returns {Array<number>|null} [左, 上, 右, 下] または null。
     */
    function getSelectionBounds(items) {
        var left, top, right, bottom;
        var found = false;

        for (var i = 0; i < items.length; i++) {
            var bounds = getItemBounds(items[i]);
            if (!bounds) continue;

            if (!found) {
                left = bounds[0];
                top = bounds[1];
                right = bounds[2];
                bottom = bounds[3];
                found = true;
                continue;
            }
            if (bounds[0] < left) left = bounds[0];
            if (bounds[1] > top) top = bounds[1];
            if (bounds[2] > right) right = bounds[2];
            if (bounds[3] < bottom) bottom = bounds[3];
        }

        return found ? [left, top, right, bottom] : null;
    }

    /**
     * グループ・複合パスをたどって、すべてのパスに処理を適用します。
     *
     * @param {Array<PageItem>} items - 対象のオブジェクト。
     * @param {function(PathItem):void} handler - パスごとに呼ぶ処理。
     * @returns {void}
     */
    function forEachPathItem(items, handler) {
        for (var i = 0; i < items.length; i++) {
            var item = items[i];
            if (item.typename === "PathItem") {
                handler(item);
            } else if (item.typename === "GroupItem") {
                forEachPathItem(item.pageItems, handler);
            } else if (item.typename === "CompoundPathItem") {
                forEachPathItem(item.pathItems, handler);
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
     * @param {Array<PageItem>} items - 選択中のオブジェクト。
     * @returns {{points: Array<Object>, horizontal: Array<Object>, vertical: Array<Object>, diagonal: Array<Object>, top: number, bottom: number}} 解析用のジオメトリ。
     */
    function collectGeometry(items) {
        var geometry = {
            points: [],
            horizontal: [],
            vertical: [],
            diagonal: [],
            top: -Infinity,
            bottom: Infinity
        };

        forEachPathItem(items, function (pathItem) {
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
     * @param {Array<PageItem>} items - 選択中のオブジェクト。
     * @returns {Object|null} コンテキスト。作成できないときは null。
     */
    function buildContext(targetDoc, items) {
        var bounds = getSelectionBounds(items);
        if (!bounds) {
            alert(getLabel(LABELS.alert.noBounds));
            return null;
        }

        var selLeft = bounds[0];
        var selTop = bounds[1];
        var selRight = bounds[2];
        var selBottom = bounds[3];
        var selWidth = selRight - selLeft;
        var selHeight = selTop - selBottom;

        if (selWidth <= 0 || selHeight <= 0) {
            alert(getLabel(LABELS.alert.invalidSize));
            return null;
        }

        return {
            doc: targetDoc,
            items: items,
            selLeft: selLeft,
            selTop: selTop,
            selRight: selRight,
            selBottom: selBottom,
            selWidth: selWidth,
            selHeight: selHeight,
            centerX: (selLeft + selRight) / 2,
            centerY: (selTop + selBottom) / 2,
            geometry: collectGeometry(items),
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
            var key = points[i].y.toFixed(3);
            frequency[key] = (frequency[key] || 0) + 1;
        }

        var bestY = points[0].y;
        var bestCount = -1;
        for (var key in frequency) {
            if (frequency.hasOwnProperty(key) && frequency[key] > bestCount) {
                bestCount = frequency[key];
                bestY = parseFloat(key);
            }
        }

        return { y: bestY, count: points.length, onHorizontal: onHorizontal };
    }

    /**
     * 候補のうち、もっとも横線らしい1本を選びます（水平な辺を持つものを優先）。
     *
     * @param {Array<Object>} lines - 候補の横線。
     * @returns {Object} 選ばれた横線。
     */
    function pickDominantLine(lines) {
        var best = lines[0];
        var bestScore = -Infinity;
        for (var i = 0; i < lines.length; i++) {
            var score = lines[i].count + (lines[i].onHorizontal ? 1000 : 0);
            if (score > bestScore) {
                bestScore = score;
                best = lines[i];
            }
        }
        return best;
    }

    /**
     * アンカーポイントの分布から、書体の基準線（アセンダー〜ディセンダー）を推定します。
     *
     * @param {Object} geometry - collectGeometry() の結果。
     * @returns {{ascenderY: number, meanY: number, baseY: number, descenderY: number, hasDescender: boolean}} 推定した基準線。
     */
    function detectTypographicLines(geometry) {
        var tolerance = getClusterTolerance(geometry.top - geometry.bottom);
        var sorted = geometry.points.slice().sort(function (a, b) { return b.y - a.y; });
        var clusters = clusterByValue(sorted, function (point) { return point.y; }, tolerance);

        var lines = [];
        for (var i = 0; i < clusters.length; i++) {
            lines.push(summarizeAnchorCluster(clusters[i]));
        }

        var ascenderY = lines[0].y;
        var descenderY = lines[lines.length - 1].y;
        var height = ascenderY - descenderY;
        var middleY = (ascenderY + descenderY) / 2;

        /* 上下端から離れた候補だけを、ミーンラインとベースラインの候補にする */
        var upperLines = [];
        var lowerLines = [];
        for (var j = 1; j < lines.length - 1; j++) {
            if (Math.abs(ascenderY - lines[j].y) <= tolerance || Math.abs(lines[j].y - descenderY) <= tolerance) continue;
            if (lines[j].y > middleY) {
                upperLines.push(lines[j]);
            } else {
                lowerLines.push(lines[j]);
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
     * @param {{descending: boolean, dropShort: boolean, span: number}} options - 並び順・短いセグメントの扱い・許容値の基準サイズ。
     * @returns {Array<number>} 線の位置。
     */
    function getSegmentLinePositions(segments, options) {
        if (segments.length === 0) return [];

        var minPosition = Infinity;
        var maxPosition = -Infinity;
        var totalLength = 0;
        for (var i = 0; i < segments.length; i++) {
            if (segments[i].position < minPosition) minPosition = segments[i].position;
            if (segments[i].position > maxPosition) maxPosition = segments[i].position;
            totalLength += segments[i].length;
        }

        var targets = segments;
        if (options.dropShort) {
            var minLength = getMinSegmentLength(totalLength, segments.length);
            targets = [];
            for (var j = 0; j < segments.length; j++) {
                if (segments[j].length >= minLength) targets.push(segments[j]);
            }
            if (targets.length === 0) return [];
        }

        var span = (typeof options.span === "number") ? options.span : (maxPosition - minPosition);
        var descending = !!options.descending;
        var sorted = targets.slice().sort(function (a, b) {
            return descending ? (b.position - a.position) : (a.position - b.position);
        });
        var clusters = clusterByValue(sorted, function (segment) { return segment.position; }, getClusterTolerance(span));

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
            var lines = [geometry.top];
            var step = (geometry.top - geometry.bottom) / (evenLineCount + 1);
            for (var i = 1; i <= evenLineCount; i++) {
                lines.push(geometry.top - step * i);
            }
            lines.push(geometry.bottom);
            return lines;
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
        var sorted = segments.slice().sort(function (a, b) { return a.angle - b.angle; });
        var clusters = clusterByValue(sorted, function (segment) { return segment.angle; }, DETECTION.diagonalAngleTolerance);

        var best = null;
        var bestLength = -1;
        for (var i = 0; i < clusters.length; i++) {
            var totalLength = 0;
            for (var j = 0; j < clusters[i].length; j++) {
                totalLength += clusters[i][j].length;
            }
            if (totalLength > bestLength) {
                bestLength = totalLength;
                best = clusters[i];
            }
        }
        return best;
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

        var cluster = pickDominantAngleCluster(segments);
        if (!cluster) return [];

        var totalLength = 0;
        for (var i = 0; i < cluster.length; i++) {
            totalLength += cluster[i].length;
        }
        var angleRad = weightedAverage(cluster,
            function (segment) { return segment.angle; },
            function (segment) { return segment.length; }) * Math.PI / 180;

        var midpoints = mergeSegmentMidpoints(cluster, angleRad, getMinSegmentLength(totalLength, cluster.length));

        var lines = [];
        var usedKeys = {};
        for (var j = 0; j < midpoints.length; j++) {
            var line = extendPointAtAngle(midpoints[j], angleRad, left, top, right, bottom);
            if (!line) continue;
            var key = makeDiagonalLineKey(line);
            if (usedKeys[key]) continue;
            usedKeys[key] = true;
            lines.push(line);
        }
        return lines;
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
     * @param {Object} ui - createDialog() が返すUI参照。
     * @returns {string} "none" / "even" / "segment" / "auto"。
     */
    function getLineMethod(ui) {
        if (ui.methodNoneRadio.value) return "none";
        if (ui.methodEvenRadio.value) return "even";
        if (ui.methodSegmentRadio.value) return "segment";
        return "auto";
    }

    /**
     * プリセットの1項目を、現在のUIから読み取ります。
     *
     * @param {Object} ui - createDialog() が返すUI参照。
     * @param {string} key - プリセットのキー。
     * @returns {string|boolean} 読み取った値。
     */
    function readPresetValue(ui, key) {
        var field = PRESET_FIELDS[key];
        if (field.kind === "method") return getLineMethod(ui);
        if (field.kind === "check") return ui[field.control].value;
        return ui[field.control].text;
    }

    /**
     * プリセットの1項目を、UIに反映します。
     *
     * @param {Object} ui - createDialog() が返すUI参照。
     * @param {string} key - プリセットのキー。
     * @param {string|boolean} value - 反映する値。
     * @returns {void}
     */
    function writePresetValue(ui, key, value) {
        var field = PRESET_FIELDS[key];
        if (field.kind === "method") {
            ui.methodNoneRadio.value = (value === "none");
            ui.methodAutoRadio.value = (value === "auto");
            ui.methodSegmentRadio.value = (value === "segment");
            ui.methodEvenRadio.value = (value === "even");
            return;
        }
        if (field.kind === "check") {
            ui[field.control].value = !!value;
            return;
        }
        if (field.kind === "scale") {
            ui[field.control].text = parseFloat(value).toFixed(1);
            return;
        }
        ui[field.control].text = String(value);
    }

    /**
     * 選択中のプリセットをUIに反映します。
     *
     * @param {Object} ui - createDialog() が返すUI参照。
     * @param {Object} context - buildContext() が返すコンテキスト。
     * @returns {void}
     */
    function applyPreset(ui, context) {
        if (!ui.presetDropdown.selection) return;
        var preset = PRESET_DEFINITIONS[ui.presetDropdown.selection._key];
        if (!preset) return;

        for (var key in PRESET_FIELDS) {
            if (PRESET_FIELDS.hasOwnProperty(key)) {
                writePresetValue(ui, key, preset[key]);
            }
        }

        updateClearSpaceState(ui);
        refreshPreview(ui, context);
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
     * @param {Object} ui - createDialog() が返すUI参照。
     * @param {string} presetName - プリセット名。
     * @returns {string} 書き出す文字列。
     */
    function buildPresetSnippet(ui, presetName) {
        var lines = ['                "' + presetName + '": {'];

        for (var i = 0; i < PRESET_FIELD_ROWS.length; i++) {
            var row = PRESET_FIELD_ROWS[i];
            var parts = [];
            for (var j = 0; j < row.length; j++) {
                parts.push(row[j] + ": " + formatPresetValue(readPresetValue(ui, row[j])));
            }
            var isLastRow = (i === PRESET_FIELD_ROWS.length - 1);
            lines.push("                    " + parts.join(", ") + (isLastRow ? "" : ","));
        }

        lines.push("                },");
        return lines.join("\n");
    }

    /**
     * 現在の設定をプリセットとしてデスクトップに書き出します。
     *
     * @param {Object} ui - createDialog() が返すUI参照。
     * @returns {void}
     */
    function exportPreset(ui) {
        var presetName = prompt(getLabel(LABELS.prompt.presetName), "");
        if (!presetName) return;

        var fileName = presetName.replace(/[\\\/:*?"<>|]/g, "_") + ".json";
        var file = new File(Folder.desktop + "/" + fileName);
        file.encoding = "UTF-8";
        if (!file.open("w")) {
            alert(getLabel(LABELS.alert.exportFailed));
            return;
        }
        file.write(buildPresetSnippet(ui, presetName));
        file.close();
        alert(getLabel(LABELS.alert.exported) + file.fsName);
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * パネルの余白と並びを設定します。
     *
     * @param {Panel} panel - 対象のパネル。
     * @param {string} horizontalAlign - 子要素の横方向の揃え。
     * @returns {Panel} 設定したパネル。
     */
    function setupPanel(panel, horizontalAlign) {
        panel.margins = PANEL_MARGINS;
        panel.alignChildren = [horizontalAlign || "left", "top"];
        return panel;
    }

    /**
     * 行グループを作成します（横並び・左揃え・天地中央）。
     *
     * @param {Object} parent - 追加先のコンテナ。
     * @returns {Group} 作成した行グループ。
     */
    function addRow(parent) {
        var row = parent.add("group");
        row.orientation = "row";
        row.alignment = ["left", "top"];
        row.alignChildren = ["left", "center"];
        return row;
    }

    /**
     * 設定用のダイアログを作成します。
     *
     * @returns {Object} ダイアログと各コントロールへの参照。
     */
    function createDialog() {
        var dialog = new Window("dialog", getLabel(LABELS.dialog.title) + " " + SCRIPT_VERSION);
        dialog.orientation = "column";
        dialog.alignChildren = ["fill", "top"];

        /* プリセット / Presets */
        var presetRow = dialog.add("group");
        presetRow.orientation = "row";
        presetRow.alignment = ["center", "top"];
        presetRow.alignChildren = ["left", "center"];
        presetRow.add("statictext", undefined, labelText(LABELS.fieldLabel.preset));

        var presetDropdown = presetRow.add("dropdownlist");
        presetDropdown.helpTip = getLabel(LABELS.tooltip.preset);
        var autoIndex = 0;
        var presetIndex = 0;
        for (var presetKey in PRESET_DEFINITIONS) {
            if (!PRESET_DEFINITIONS.hasOwnProperty(presetKey)) continue;
            presetDropdown.add("item", presetKey)._key = presetKey;
            if (presetKey === "auto") autoIndex = presetIndex;
            presetIndex++;
        }
        if (presetDropdown.items.length > 0) presetDropdown.selection = autoIndex;

        var btnExport = presetRow.add("button", undefined, getLabel(LABELS.button.exportPreset));
        btnExport.helpTip = getLabel(LABELS.tooltip.exportPreset);

        /* 共通設定とクリアスペース / Common settings and clear space */
        var commonPanel = setupPanel(dialog.add("panel", undefined, getLabel(LABELS.panel.common)), "fill");

        var commonColumns = commonPanel.add("group");
        commonColumns.orientation = "row";
        commonColumns.alignChildren = ["fill", "top"];

        var commonLeftColumn = commonColumns.add("group");
        commonLeftColumn.orientation = "column";
        commonLeftColumn.alignChildren = ["left", "top"];

        var commonRightColumn = commonColumns.add("group");
        commonRightColumn.orientation = "column";
        commonRightColumn.alignChildren = ["left", "top"];

        var clearSpaceDivRow = addRow(commonLeftColumn);
        clearSpaceDivRow.add("statictext", undefined, labelText(LABELS.fieldLabel.divisions));
        var clearSpaceDivInput = clearSpaceDivRow.add("edittext", undefined, "4");
        clearSpaceDivInput.characters = FIELD_WIDTH.count;
        clearSpaceDivInput.helpTip = getLabel(LABELS.tooltip.divisions);

        var strokeWidthRow = addRow(commonRightColumn);
        strokeWidthRow.add("statictext", undefined, labelText(LABELS.fieldLabel.strokeWidth));
        var strokeWidthInput = strokeWidthRow.add("edittext", undefined, "0.25");
        strokeWidthInput.characters = FIELD_WIDTH.strokeWidth;
        strokeWidthInput.helpTip = getLabel(LABELS.tooltip.strokeWidth);
        strokeWidthRow.add("statictext", undefined, getStrokeUnit().label);

        var emphasizeBoundsRow = addRow(commonRightColumn);
        var emphasizeBoundsCheck = emphasizeBoundsRow.add("checkbox", undefined, getLabel(LABELS.checkbox.emphasizeBounds));
        emphasizeBoundsCheck.value = false;
        emphasizeBoundsCheck.helpTip = getLabel(LABELS.tooltip.emphasizeBounds);

        var clearSpaceRow = addRow(commonLeftColumn);
        var clearSpaceCheck = clearSpaceRow.add("checkbox", undefined, getLabel(LABELS.checkbox.clearSpace));
        clearSpaceCheck.value = false;
        clearSpaceCheck.helpTip = getLabel(LABELS.tooltip.clearSpace);

        var convertToGuidesRow = addRow(commonLeftColumn);
        var convertToGuidesCheck = convertToGuidesRow.add("checkbox", undefined, getLabel(LABELS.checkbox.convertToGuides));
        convertToGuidesCheck.value = false;
        convertToGuidesCheck.helpTip = getLabel(LABELS.tooltip.convertToGuides);

        var groupItemsRow = addRow(commonRightColumn);
        var groupItemsCheck = groupItemsRow.add("checkbox", undefined, getLabel(LABELS.checkbox.groupItems));
        groupItemsCheck.value = true;
        groupItemsCheck.helpTip = getLabel(LABELS.tooltip.groupItems);

        var layerNameRow = addRow(commonPanel);
        layerNameRow.add("statictext", undefined, labelText(LABELS.fieldLabel.targetLayer));
        var layerNameInput = layerNameRow.add("edittext", undefined, getLabel(LABELS.itemName.layer));
        layerNameInput.characters = FIELD_WIDTH.layerName;
        layerNameInput.helpTip = getLabel(LABELS.tooltip.targetLayer);

        /* 横線・縦線の2カラム / Two columns for the horizontal and vertical lines */
        var panelColumns = dialog.add("group");
        panelColumns.orientation = "row";
        panelColumns.alignChildren = ["fill", "top"];

        var leftColumn = panelColumns.add("group");
        leftColumn.orientation = "column";
        leftColumn.alignChildren = ["fill", "top"];

        var rightColumn = panelColumns.add("group");
        rightColumn.orientation = "column";
        rightColumn.alignChildren = ["fill", "top"];

        /* 横線 / Horizontal lines */
        var horizontalPanel = setupPanel(leftColumn.add("panel", undefined, getLabel(LABELS.panel.horizontal)));

        var widthScaleRow = addRow(horizontalPanel);
        widthScaleRow.add("statictext", undefined, labelText(LABELS.fieldLabel.widthScale));
        var widthScaleInput = widthScaleRow.add("edittext", undefined, "140");
        widthScaleInput.characters = FIELD_WIDTH.scale;
        widthScaleInput.helpTip = getLabel(LABELS.tooltip.widthScale);
        widthScaleRow.add("statictext", undefined, "%");

        var extraHLinesRow = addRow(horizontalPanel);
        var extraHLinesCheck = extraHLinesRow.add("checkbox", undefined, getLabel(LABELS.checkbox.extraHorizontal));
        extraHLinesCheck.value = false;
        extraHLinesCheck.helpTip = getLabel(LABELS.tooltip.extraHorizontal);
        var extraHLinesInput = extraHLinesRow.add("edittext", undefined, "1");
        extraHLinesInput.characters = FIELD_WIDTH.count;
        extraHLinesInput.helpTip = getLabel(LABELS.tooltip.extraHorizontal);

        var extendLeftRow = addRow(horizontalPanel);
        var extendLeftCheck = extendLeftRow.add("checkbox", undefined, getLabel(LABELS.checkbox.extendLeft));
        extendLeftCheck.value = false;
        extendLeftCheck.helpTip = getLabel(LABELS.tooltip.extendLeft);

        /* ラインの決め方 / Line generation method */
        var lineMethodPanel = setupPanel(horizontalPanel.add("panel", undefined, getLabel(LABELS.panel.lineMethod)));
        lineMethodPanel.helpTip = getLabel(LABELS.tooltip.lineMethod);

        var methodNoneRadio = lineMethodPanel.add("radiobutton", undefined, getLabel(LABELS.radio.lineNone));
        methodNoneRadio.helpTip = getLabel(LABELS.tooltip.lineNone);
        var methodAutoRadio = lineMethodPanel.add("radiobutton", undefined, getLabel(LABELS.radio.lineAuto));
        methodAutoRadio.value = true;
        methodAutoRadio.helpTip = getLabel(LABELS.tooltip.lineAuto);
        var methodSegmentRadio = lineMethodPanel.add("radiobutton", undefined, getLabel(LABELS.radio.lineSegment));
        methodSegmentRadio.helpTip = getLabel(LABELS.tooltip.lineSegment);

        var evenLineCountRow = addRow(lineMethodPanel);
        var methodEvenRadio = evenLineCountRow.add("radiobutton", undefined, getLabel(LABELS.radio.lineEven));
        methodEvenRadio.helpTip = getLabel(LABELS.tooltip.lineEven);
        var evenLineCountInput = evenLineCountRow.add("edittext", undefined, "4");
        evenLineCountInput.characters = FIELD_WIDTH.count;
        evenLineCountInput.helpTip = getLabel(LABELS.tooltip.lineEven);
        evenLineCountRow.add("statictext", undefined, getLabel(LABELS.unit.lines));

        /* 縦線 / Vertical lines */
        var verticalPanel = setupPanel(rightColumn.add("panel", undefined, getLabel(LABELS.panel.vertical)));

        var heightScaleRow = addRow(verticalPanel);
        heightScaleRow.add("statictext", undefined, labelText(LABELS.fieldLabel.heightScale));
        var heightScaleInput = heightScaleRow.add("edittext", undefined, "200");
        heightScaleInput.characters = FIELD_WIDTH.scale;
        heightScaleInput.helpTip = getLabel(LABELS.tooltip.heightScale);
        heightScaleRow.add("statictext", undefined, "%");

        var outerVLinesRow = addRow(verticalPanel);
        var outerVLinesCheck = outerVLinesRow.add("checkbox", undefined, getLabel(LABELS.checkbox.outerVertical));
        outerVLinesCheck.value = false;
        outerVLinesCheck.helpTip = getLabel(LABELS.tooltip.outerVertical);
        var outerVLinesInput = outerVLinesRow.add("edittext", undefined, "1");
        outerVLinesInput.characters = FIELD_WIDTH.count;
        outerVLinesInput.helpTip = getLabel(LABELS.tooltip.outerVertical);

        var extendUpRow = addRow(verticalPanel);
        var extendUpCheck = extendUpRow.add("checkbox", undefined, getLabel(LABELS.checkbox.extendUp));
        extendUpCheck.value = false;
        extendUpCheck.helpTip = getLabel(LABELS.tooltip.extendUp);

        /* オプション / Options */
        var verticalOptionsPanel = setupPanel(verticalPanel.add("panel", undefined, getLabel(LABELS.panel.options)));

        var columnDivRow = addRow(verticalOptionsPanel);
        var columnDivCheck = columnDivRow.add("checkbox", undefined, getLabel(LABELS.checkbox.columnDiv));
        columnDivCheck.value = false;
        columnDivCheck.helpTip = getLabel(LABELS.tooltip.columnDiv);
        var columnDivInput = columnDivRow.add("edittext", undefined, "2");
        columnDivInput.characters = FIELD_WIDTH.count;
        columnDivInput.enabled = false;
        columnDivInput.helpTip = getLabel(LABELS.tooltip.columnDiv);

        var verticalElementsRow = addRow(verticalOptionsPanel);
        var verticalElementsCheck = verticalElementsRow.add("checkbox", undefined, getLabel(LABELS.checkbox.verticalElements));
        verticalElementsCheck.value = false;
        verticalElementsCheck.helpTip = getLabel(LABELS.tooltip.verticalElements);

        var diagonalElementsRow = addRow(verticalOptionsPanel);
        var diagonalElementsCheck = diagonalElementsRow.add("checkbox", undefined, getLabel(LABELS.checkbox.diagonalElements));
        diagonalElementsCheck.value = false;
        diagonalElementsCheck.helpTip = getLabel(LABELS.tooltip.diagonalElements);

        /* ボタンエリア / Button area */
        var btnRowGroup = dialog.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.alignment = ["fill", "bottom"];

        var btnLeftGroup = btnRowGroup.add("group");
        btnLeftGroup.orientation = "row";
        btnLeftGroup.alignChildren = ["left", "center"];
        var previewCheck = btnLeftGroup.add("checkbox", undefined, getLabel(LABELS.checkbox.preview));
        previewCheck.value = true;
        previewCheck.helpTip = getLabel(LABELS.tooltip.preview);

        var spacer = btnRowGroup.add("group");
        spacer.alignment = ["fill", "fill"];
        spacer.minimumSize.width = 0;

        var btnRightGroup = btnRowGroup.add("group");
        btnRightGroup.orientation = "row";
        btnRightGroup.alignChildren = ["right", "center"];
        var btnCancel = btnRightGroup.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });
        var btnOK = btnRightGroup.add("button", undefined, getLabel(LABELS.button.ok), { name: "ok" });

        return {
            dialog: dialog,
            presetDropdown: presetDropdown,
            btnExport: btnExport,
            horizontalPanel: horizontalPanel,
            verticalPanel: verticalPanel,
            widthScaleInput: widthScaleInput,
            extraHLinesCheck: extraHLinesCheck,
            extraHLinesInput: extraHLinesInput,
            extendLeftCheck: extendLeftCheck,
            methodNoneRadio: methodNoneRadio,
            methodAutoRadio: methodAutoRadio,
            methodSegmentRadio: methodSegmentRadio,
            methodEvenRadio: methodEvenRadio,
            evenLineCountInput: evenLineCountInput,
            heightScaleInput: heightScaleInput,
            outerVLinesCheck: outerVLinesCheck,
            outerVLinesInput: outerVLinesInput,
            columnDivCheck: columnDivCheck,
            columnDivInput: columnDivInput,
            extendUpCheck: extendUpCheck,
            verticalElementsCheck: verticalElementsCheck,
            diagonalElementsCheck: diagonalElementsCheck,
            clearSpaceCheck: clearSpaceCheck,
            clearSpaceDivInput: clearSpaceDivInput,
            layerNameInput: layerNameInput,
            strokeWidthInput: strokeWidthInput,
            convertToGuidesCheck: convertToGuidesCheck,
            groupItemsCheck: groupItemsCheck,
            emphasizeBoundsCheck: emphasizeBoundsCheck,
            previewCheck: previewCheck,
            btnCancel: btnCancel,
            btnOK: btnOK
        };
    }

    // =========================================
    // UIの状態 / UI state
    // =========================================

    /**
     * ［グループ化］の使用可否を、［ガイドに変換する］に合わせます。
     *
     * @param {Object} ui - createDialog() が返すUI参照。
     * @returns {void}
     */
    function updateGroupState(ui) {
        ui.groupItemsCheck.enabled = !ui.convertToGuidesCheck.value;
        if (ui.convertToGuidesCheck.value) ui.groupItemsCheck.value = false;
    }

    /**
     * ［均等分割］の本数入力の使用可否を更新します。
     *
     * @param {Object} ui - createDialog() が返すUI参照。
     * @returns {void}
     */
    function updateColumnDivField(ui) {
        ui.columnDivInput.enabled = ui.columnDivCheck.enabled && ui.columnDivCheck.value;
    }

    /**
     * ［分割数］の値と使用可否を、横線の均等分割に合わせます。
     *
     * @param {Object} ui - createDialog() が返すUI参照。
     * @returns {void}
     */
    function updateClearSpaceDivField(ui) {
        var evenMode = ui.methodEvenRadio.value;
        var evenCount = parseInt(ui.evenLineCountInput.text, 10);

        /* 均等分割の本数＋1がユニット数になる / The unit count follows the even division count */
        if (evenMode && !isNaN(evenCount) && evenCount >= 1) {
            ui.clearSpaceDivInput.text = String(evenCount + 1);
        }
        ui.clearSpaceDivInput.enabled = ui.clearSpaceCheck.value || !evenMode;
    }

    /**
     * 分割にかかわる入力欄の状態をまとめて更新します。
     *
     * @param {Object} ui - createDialog() が返すUI参照。
     * @returns {void}
     */
    function syncDivisionFields(ui) {
        updateColumnDivField(ui);
        updateClearSpaceDivField(ui);
    }

    /**
     * ［クリアスペース］の状態に合わせて、横線・縦線パネルなどの使用可否を更新します。
     *
     * @param {Object} ui - createDialog() が返すUI参照。
     * @returns {void}
     */
    function updateClearSpaceState(ui) {
        var clearSpace = ui.clearSpaceCheck.value;

        /* 横線・縦線はパネルごとディム / Dim the whole horizontal and vertical panels */
        ui.horizontalPanel.enabled = !clearSpace;
        ui.verticalPanel.enabled = !clearSpace;
        ui.extraHLinesInput.enabled = !clearSpace && ui.extraHLinesCheck.value;
        ui.evenLineCountInput.enabled = !clearSpace && ui.methodEvenRadio.value;
        ui.outerVLinesInput.enabled = !clearSpace && ui.outerVLinesCheck.value;
        ui.columnDivInput.enabled = !clearSpace && ui.columnDivCheck.value;

        /* クリアスペース中は線幅とガイド化を使わず、グループ化はON固定 */
        ui.strokeWidthInput.enabled = !clearSpace;
        ui.convertToGuidesCheck.enabled = !clearSpace;
        if (clearSpace) {
            ui.convertToGuidesCheck.value = false;
            ui.groupItemsCheck.value = true;
            ui.groupItemsCheck.enabled = false;
        } else {
            updateGroupState(ui);
        }

        syncDivisionFields(ui);
    }

    // =========================================
    // 入力値 / Parameters
    // =========================================

    /**
     * 伸張率の入力欄か判定します（小数第1位で表示する欄）。
     *
     * @param {Object} ui - createDialog() が返すUI参照。
     * @param {EditText} field - 対象の入力欄。
     * @returns {boolean} 伸張率の入力欄のとき true。
     */
    function isScaleField(ui, field) {
        return field === ui.widthScaleInput || field === ui.heightScaleInput;
    }

    /**
     * 正の数として扱える入力か判定します。
     *
     * @param {string} text - 入力された文字列。
     * @returns {boolean} 正の数のとき true。
     */
    function isPositiveNumberText(text) {
        return /^\d+(?:\.\d+)?$/.test(text) && parseFloat(text) > 0;
    }

    /**
     * 1以上の整数として扱える入力か判定します。
     *
     * @param {string} text - 入力された文字列。
     * @returns {boolean} 1以上の整数のとき true。
     */
    function isCountText(text) {
        return /^\d+$/.test(text) && parseInt(text, 10) >= 1;
    }

    /**
     * ダイアログの入力内容をそのまま読み取ります。
     *
     * @param {Object} ui - createDialog() が返すUI参照。
     * @returns {Object} 入力内容。
     */
    function readParams(ui) {
        return {
            layerName: normalizeLayerNameText(ui.layerNameInput.text),
            widthScaleText: ui.widthScaleInput.text,
            heightScaleText: ui.heightScaleInput.text,
            strokeWidthText: ui.strokeWidthInput.text,
            lineMethod: getLineMethod(ui),
            evenLineCountText: ui.evenLineCountInput.text,
            extraHLines: ui.extraHLinesCheck.value,
            extraHLineCountText: ui.extraHLinesInput.text,
            extendLeft: ui.extendLeftCheck.value,
            outerVLines: ui.outerVLinesCheck.value,
            outerVLineCountText: ui.outerVLinesInput.text,
            extendUp: ui.extendUpCheck.value,
            columnDiv: ui.columnDivCheck.value,
            columnDivisionsText: ui.columnDivInput.text,
            verticalElements: ui.verticalElementsCheck.value,
            diagonalElements: ui.diagonalElementsCheck.value,
            emphasizeBounds: ui.emphasizeBoundsCheck.value,
            convertToGuides: ui.convertToGuidesCheck.value,
            groupItems: ui.groupItemsCheck.value,
            clearSpace: ui.clearSpaceCheck.value,
            clearSpaceDivisionsText: ui.clearSpaceDivInput.text
        };
    }

    /**
     * 入力内容が処理に使える値か検証します。
     *
     * @param {Object} raw - readParams() の結果。
     * @returns {boolean} すべて正しいとき true。
     */
    function validateParams(raw) {
        if (!raw.layerName) return false;
        if (!isPositiveNumberText(raw.widthScaleText)) return false;
        if (!isPositiveNumberText(raw.heightScaleText)) return false;
        if (!isPositiveNumberText(raw.strokeWidthText)) return false;
        if (!isPositiveNumberText(raw.clearSpaceDivisionsText)) return false;
        if (raw.lineMethod === "even" && !isCountText(raw.evenLineCountText)) return false;
        if (raw.extraHLines && !isCountText(raw.extraHLineCountText)) return false;
        if (raw.outerVLines && !isCountText(raw.outerVLineCountText)) return false;
        if (raw.columnDiv && !isCountText(raw.columnDivisionsText)) return false;
        return true;
    }

    /**
     * 入力内容を、描画で使う単位・型に変換します。
     *
     * @param {Object} raw - readParams() の結果。
     * @returns {Object|null} 変換した設定。線幅を換算できないときは null。
     */
    function normalizeParams(raw) {
        var strokeWidth = toPoints(parseFloat(raw.strokeWidthText));
        if (!(strokeWidth > 0)) return null;

        return {
            layerName: raw.layerName,
            widthScale: parseFloat(raw.widthScaleText) / 100,
            heightScale: parseFloat(raw.heightScaleText) / 100,
            strokeWidth: strokeWidth,
            lineMethod: raw.lineMethod,
            evenLineCount: (raw.lineMethod === "even") ? parseInt(raw.evenLineCountText, 10) : 0,
            extraHLines: raw.extraHLines,
            extraHLineCount: raw.extraHLines ? parseInt(raw.extraHLineCountText, 10) : 0,
            extendLeft: raw.extendLeft,
            outerVLines: raw.outerVLines,
            outerVLineCount: raw.outerVLines ? parseInt(raw.outerVLineCountText, 10) : 0,
            extendUp: raw.extendUp,
            columnDiv: raw.columnDiv,
            columnDivisions: raw.columnDiv ? parseInt(raw.columnDivisionsText, 10) : 0,
            verticalElements: raw.verticalElements,
            diagonalElements: raw.diagonalElements,
            emphasizeBounds: raw.emphasizeBounds,
            convertToGuides: raw.convertToGuides,
            groupItems: raw.groupItems,
            clearSpace: raw.clearSpace,
            clearSpaceDivisions: parseFloat(raw.clearSpaceDivisionsText)
        };
    }

    /**
     * ダイアログの入力内容を、検証したうえで設定にまとめます。
     *
     * @param {Object} ui - createDialog() が返すUI参照。
     * @returns {Object|null} 設定。入力値が不正なときは null。
     */
    function getParams(ui) {
        var raw = readParams(ui);
        return validateParams(raw) ? normalizeParams(raw) : null;
    }

    // =========================================
    // 描画 / Drawing
    // =========================================

    /**
     * 補助線を描く範囲（線の長さとユニットの大きさ）を求めます。
     *
     * @param {Object} context - buildContext() が返すコンテキスト。
     * @param {Object} params - getParams() が返す設定。
     * @returns {{left: number, right: number, top: number, bottom: number, unitSize: number}} 描画範囲。
     */
    function computeFrame(context, params) {
        var unitSize = context.selHeight / params.clearSpaceDivisions;

        var halfWidth = context.selWidth * params.widthScale / 2;
        var left = context.centerX - halfWidth;
        var right = context.centerX + halfWidth;
        if (params.extendLeft) {
            /* 右端を選択範囲の右外に固定し、余りを左へ伸ばす / Pin the right end and extend leftward */
            var shift = right - (context.selRight + context.selHeight / 2);
            left -= shift;
            right -= shift;
        }

        var lineHeight = context.selHeight * params.heightScale;
        var top, bottom;
        if (params.extendUp) {
            bottom = context.selBottom - unitSize * 2;
            top = bottom + lineHeight;
        } else {
            top = context.centerY + lineHeight / 2;
            bottom = context.centerY - lineHeight / 2;
        }

        if (!params.clearSpace) {
            /* 外側に追加した線とも必ず交差するよう、線の長さを自動で補正する */
            var outerOffset = params.outerVLines ? unitSize * (params.outerVLineCount - 1) : 0;
            var extraOffset = params.extraHLines ? unitSize * (params.extraHLineCount - 1) : 0;
            left = Math.min(left, context.selLeft - outerOffset);
            right = Math.max(right, context.selRight + outerOffset);
            top = Math.max(top, context.selTop + extraOffset);
            bottom = Math.min(bottom, context.selBottom - extraOffset);
        }

        return { left: left, right: right, top: top, bottom: bottom, unitSize: unitSize };
    }

    /**
     * 直線を1本描きます。
     *
     * @param {GroupItem} parent - 追加先のグループ。
     * @param {number} x1 - 始点のX座標。
     * @param {number} y1 - 始点のY座標。
     * @param {number} x2 - 終点のX座標。
     * @param {number} y2 - 終点のY座標。
     * @param {Object} color - 線のカラー。
     * @param {number} strokeWidth - 線幅（pt）。
     * @returns {PathItem} 作成したパス。
     */
    function drawLine(parent, x1, y1, x2, y2, color, strokeWidth) {
        var line = parent.pathItems.add();
        line.setEntirePath([[x1, y1], [x2, y2]]);
        line.stroked = true;
        line.filled = false;
        line.strokeWidth = strokeWidth;
        line.strokeColor = color;
        return line;
    }

    /**
     * 横線を1本描きます。
     *
     * @param {GroupItem} parent - 追加先のグループ。
     * @param {Object} frame - computeFrame() が返す描画範囲。
     * @param {number} y - 描くY座標。
     * @param {Object} stroke - 線のカラーと線幅。
     * @param {boolean} onBounds - 選択範囲の境界に重なるか。
     * @returns {void}
     */
    function drawHorizontalLine(parent, frame, y, stroke, onBounds) {
        drawLine(parent, frame.left, y, frame.right, y, stroke.color, onBounds ? stroke.bounds : stroke.normal);
    }

    /**
     * 縦線を1本描きます。
     *
     * @param {GroupItem} parent - 追加先のグループ。
     * @param {Object} frame - computeFrame() が返す描画範囲。
     * @param {number} x - 描くX座標。
     * @param {Object} stroke - 線のカラーと線幅。
     * @param {boolean} onBounds - 選択範囲の境界に重なるか。
     * @returns {void}
     */
    function drawVerticalLine(parent, frame, x, stroke, onBounds) {
        drawLine(parent, x, frame.top, x, frame.bottom, stroke.color, onBounds ? stroke.bounds : stroke.normal);
    }

    /**
     * 横線（境界線・外側の線・ラインの決め方による線）を描きます。
     *
     * @param {GroupItem} group - 追加先のグループ。
     * @param {Object} context - buildContext() が返すコンテキスト。
     * @param {Object} params - getParams() が返す設定。
     * @param {Object} frame - computeFrame() が返す描画範囲。
     * @param {Object} stroke - 線のカラーと線幅。
     * @returns {void}
     */
    function drawHorizontalGuides(group, context, params, frame, stroke) {
        if (params.extraHLines) {
            drawHorizontalLine(group, frame, context.selTop, stroke, true);
            drawHorizontalLine(group, frame, context.selBottom, stroke, true);
            for (var i = 1; i < params.extraHLineCount; i++) {
                drawHorizontalLine(group, frame, context.selTop + frame.unitSize * i, stroke, false);
                drawHorizontalLine(group, frame, context.selBottom - frame.unitSize * i, stroke, false);
            }
        }

        var lineYs = getHorizontalLineYs(context.geometry, params.lineMethod, params.evenLineCount);
        for (var j = 0; j < lineYs.length; j++) {
            var onBounds = params.emphasizeBounds &&
                (isSameCoord(lineYs[j], context.selTop) || isSameCoord(lineYs[j], context.selBottom));
            drawHorizontalLine(group, frame, lineYs[j], stroke, onBounds);
        }
    }

    /**
     * 縦線（境界線・外側の線・均等分割・垂直／斜線エレメント）を描きます。
     *
     * @param {GroupItem} group - 追加先のグループ。
     * @param {Object} context - buildContext() が返すコンテキスト。
     * @param {Object} params - getParams() が返す設定。
     * @param {Object} frame - computeFrame() が返す描画範囲。
     * @param {Object} stroke - 線のカラーと線幅。
     * @returns {void}
     */
    function drawVerticalGuides(group, context, params, frame, stroke) {
        if (params.outerVLines) {
            drawVerticalLine(group, frame, context.selLeft, stroke, true);
            drawVerticalLine(group, frame, context.selRight, stroke, true);
            for (var i = 1; i < params.outerVLineCount; i++) {
                drawVerticalLine(group, frame, context.selLeft - frame.unitSize * i, stroke, false);
                drawVerticalLine(group, frame, context.selRight + frame.unitSize * i, stroke, false);
            }
        }

        if (params.columnDiv && params.columnDivisions > 1) {
            var step = context.selWidth / params.columnDivisions;
            for (var j = 1; j < params.columnDivisions; j++) {
                drawVerticalLine(group, frame, context.selLeft + step * j, stroke, false);
            }
        }

        if (params.verticalElements) {
            var elementXs = getVerticalElementXs(context.geometry);
            for (var k = 0; k < elementXs.length; k++) {
                var onBounds = params.emphasizeBounds &&
                    (isSameCoord(elementXs[k], context.selLeft) || isSameCoord(elementXs[k], context.selRight));
                drawVerticalLine(group, frame, elementXs[k], stroke, onBounds);
            }
        }

        if (params.diagonalElements) {
            var diagonalLines = getDiagonalElementLines(context.geometry, frame.left, frame.top, frame.right, frame.bottom);
            for (var m = 0; m < diagonalLines.length; m++) {
                drawLine(group,
                    diagonalLines[m][0][0], diagonalLines[m][0][1],
                    diagonalLines[m][1][0], diagonalLines[m][1][1],
                    stroke.color, stroke.normal);
            }
        }
    }

    /**
     * クリアスペースの帯を1つ描きます（塗りと枠線の2枚重ね）。
     *
     * @param {GroupItem} parent - 追加先のグループ。
     * @param {number} left - 左端。
     * @param {number} top - 上端。
     * @param {number} width - 幅。
     * @param {number} height - 高さ。
     * @param {Object} color - 帯のカラー。
     * @returns {void}
     */
    function drawClearSpaceBand(parent, left, top, width, height, color) {
        var fill = parent.pathItems.rectangle(top, left, width, height);
        fill.filled = true;
        fill.fillColor = color;
        fill.stroked = false;
        fill.opacity = CLEAR_SPACE_STYLE.fillOpacity;

        var outline = parent.pathItems.rectangle(top, left, width, height);
        outline.filled = false;
        outline.stroked = true;
        outline.strokeColor = color;
        outline.strokeWidth = CLEAR_SPACE_STYLE.strokeWidth;
    }

    /**
     * 選択範囲の四方にクリアスペースを描きます。
     *
     * @param {GroupItem} group - 追加先のグループ。
     * @param {Object} context - buildContext() が返すコンテキスト。
     * @param {number} unitSize - ユニット（1単位）の大きさ。
     * @returns {void}
     */
    function drawClearSpace(group, context, unitSize) {
        var left = context.selLeft - unitSize;
        var right = context.selRight + unitSize;
        var top = context.selTop + unitSize;
        var outerWidth = right - left;
        var outerHeight = top - (context.selBottom - unitSize);
        var color = makeCyanColor(CLEAR_SPACE_STYLE.cyan);

        var clearSpaceGroup = group.groupItems.add();
        clearSpaceGroup.name = getLabel(LABELS.itemName.clearSpaceGroup);

        drawClearSpaceBand(clearSpaceGroup, left, top, outerWidth, unitSize, color);
        drawClearSpaceBand(clearSpaceGroup, left, context.selBottom, outerWidth, unitSize, color);
        drawClearSpaceBand(clearSpaceGroup, left, top, unitSize, outerHeight, color);
        drawClearSpaceBand(clearSpaceGroup, context.selRight, top, unitSize, outerHeight, color);
    }

    /**
     * 作成した線をガイドに変換し、グループ化しない設定ならレイヤー直下に移します。
     *
     * @param {GroupItem} group - 作成した線のグループ。
     * @param {Object} context - buildContext() が返すコンテキスト。
     * @param {Object} params - getParams() が返す設定。
     * @param {boolean} isPreview - プレビューとして作成したか。
     * @returns {GroupItem|null} 残ったグループ。解除したときは null。
     */
    function finishGroup(group, context, params, isPreview) {
        /* プレビューは削除できるようグループのまま残す / Keep previews grouped so they can be removed */
        if (isPreview) return group;

        if (params.convertToGuides) {
            for (var i = 0; i < group.pathItems.length; i++) {
                group.pathItems[i].guides = true;
            }
        }

        if (!params.groupItems && !params.convertToGuides) {
            while (group.pageItems.length > 0) {
                group.pageItems[0].move(context.guideLayer, ElementPlacement.PLACEATBEGINNING);
            }
            group.remove();
            return null;
        }

        return group;
    }

    /**
     * 設定に従って補助線またはクリアスペースを作成します。
     *
     * @param {Object} context - buildContext() が返すコンテキスト。
     * @param {Object} params - getParams() が返す設定。
     * @param {boolean} isPreview - プレビューとして作成するか。
     * @returns {GroupItem|null} 作成したグループ。グループ化しないときは null。
     */
    function createLines(context, params, isPreview) {
        var frame = computeFrame(context, params);
        var stroke = {
            color: makeGrayColor(LINE_STYLE.grayTint),
            normal: params.strokeWidth,
            bounds: params.emphasizeBounds ? (params.strokeWidth * LINE_STYLE.boundsMultiplier) : params.strokeWidth
        };

        var group = context.guideLayer.groupItems.add();
        group.name = getLabel(isPreview ? LABELS.itemName.previewGroup : LABELS.itemName.guideGroup);

        if (params.clearSpace) {
            drawClearSpace(group, context, frame.unitSize);
        } else {
            drawHorizontalGuides(group, context, params, frame, stroke);
            drawVerticalGuides(group, context, params, frame, stroke);
        }

        return finishGroup(group, context, params, isPreview);
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
        try {
            previewGroup.remove();
        } catch (e) { }
        previewGroup = null;
        app.redraw();
    }

    /**
     * 現在の設定でプレビューを作り直します。
     *
     * @param {Object} ui - createDialog() が返すUI参照。
     * @param {Object} context - buildContext() が返すコンテキスト。
     * @returns {void}
     */
    function updatePreview(ui, context) {
        removePreview();
        var params = getParams(ui);
        if (!params) return;
        context.guideLayer = ensureGuideLayer(context, params.layerName);
        previewGroup = createLines(context, params, true);
        app.redraw();
    }

    /**
     * ［プレビュー］がONのときだけ、プレビューを作り直します。
     *
     * @param {Object} ui - createDialog() が返すUI参照。
     * @param {Object} context - buildContext() が返すコンテキスト。
     * @returns {void}
     */
    function refreshPreview(ui, context) {
        if (ui.previewCheck.value) updatePreview(ui, context);
    }

    // =========================================
    // イベント / Events
    // =========================================

    /**
     * 入力欄で↑↓キーによる増減を有効にします（shiftで10、optionで0.1刻み）。
     *
     * @param {EditText} editText - 対象の入力欄。
     * @param {boolean} integerOnly - 整数だけを扱うか。
     * @param {Object} ui - createDialog() が返すUI参照。
     * @param {Object} context - buildContext() が返すコンテキスト。
     * @returns {void}
     */
    function changeValueByArrowKey(editText, integerOnly, ui, context) {
        editText.addEventListener("keydown", function (event) {
            if (event.keyName !== "Up" && event.keyName !== "Down") return;

            var value = Number(editText.text);
            if (isNaN(value)) return;

            var isUp = (event.keyName === "Up");
            var keyboard = ScriptUI.environment.keyboardState;
            if (keyboard.shiftKey) {
                value = isUp ? (Math.ceil((value + 1) / 10) * 10) : (Math.floor((value - 1) / 10) * 10);
            } else if (keyboard.altKey && !integerOnly) {
                value += isUp ? 0.1 : -0.1;
            } else {
                value += isUp ? 1 : -1;
            }

            if (value < 0) value = 0;
            value = (keyboard.altKey && !integerOnly) ? (Math.round(value * 10) / 10) : Math.round(value);

            event.preventDefault();
            editText.text = isScaleField(ui, editText) ? value.toFixed(1) : String(value);

            if (editText === ui.evenLineCountInput || editText === ui.columnDivInput) {
                syncDivisionFields(ui);
            }
            refreshPreview(ui, context);
        });
    }

    /**
     * ↑↓キーで増減できる入力欄をまとめて設定します。
     *
     * @param {Object} ui - createDialog() が返すUI参照。
     * @param {Object} context - buildContext() が返すコンテキスト。
     * @returns {void}
     */
    function bindArrowKeyHandlers(ui, context) {
        changeValueByArrowKey(ui.widthScaleInput, false, ui, context);
        changeValueByArrowKey(ui.heightScaleInput, false, ui, context);
        changeValueByArrowKey(ui.strokeWidthInput, false, ui, context);
        changeValueByArrowKey(ui.clearSpaceDivInput, false, ui, context);
        changeValueByArrowKey(ui.columnDivInput, true, ui, context);
        changeValueByArrowKey(ui.extraHLinesInput, true, ui, context);
        changeValueByArrowKey(ui.outerVLinesInput, true, ui, context);
        changeValueByArrowKey(ui.evenLineCountInput, true, ui, context);
    }

    /**
     * 伸張率の入力欄を、フォーカスが外れたときに小数第1位の表記にそろえます。
     *
     * @param {EditText} field - 対象の入力欄。
     * @returns {void}
     */
    function bindScaleFieldBlur(field) {
        field.onBlur = function () {
            var value = parseFloat(field.text);
            if (!isNaN(value)) field.text = value.toFixed(1);
        };
    }

    /**
     * ダイアログのイベントハンドラーを設定します。
     *
     * @param {Object} ui - createDialog() が返すUI参照。
     * @param {Object} context - buildContext() が返すコンテキスト。
     * @returns {void}
     */
    function bindEvents(ui, context) {
        ui.presetDropdown.onChange = function () {
            applyPreset(ui, context);
        };

        ui.btnExport.onClick = function () {
            exportPreset(ui);
        };

        ui.previewCheck.onClick = function () {
            if (ui.previewCheck.value) {
                updatePreview(ui, context);
            } else {
                removePreview();
            }
        };

        ui.widthScaleInput.onChanging = ui.heightScaleInput.onChanging =
            ui.strokeWidthInput.onChanging = ui.columnDivInput.onChanging =
            ui.extraHLinesInput.onChanging = ui.outerVLinesInput.onChanging =
            ui.clearSpaceDivInput.onChanging = function () {
                refreshPreview(ui, context);
            };

        ui.evenLineCountInput.onChanging = function () {
            syncDivisionFields(ui);
            refreshPreview(ui, context);
        };

        bindScaleFieldBlur(ui.widthScaleInput);
        bindScaleFieldBlur(ui.heightScaleInput);

        ui.layerNameInput.onBlur = function () {
            ui.layerNameInput.text = normalizeLayerNameText(ui.layerNameInput.text);
        };

        ui.methodNoneRadio.onClick = ui.methodAutoRadio.onClick =
            ui.methodSegmentRadio.onClick = ui.methodEvenRadio.onClick = function () {
                ui.evenLineCountInput.enabled = ui.methodEvenRadio.value;
                syncDivisionFields(ui);
                refreshPreview(ui, context);
            };

        ui.extraHLinesCheck.onClick = ui.extendLeftCheck.onClick =
            ui.outerVLinesCheck.onClick = ui.columnDivCheck.onClick =
            ui.extendUpCheck.onClick = ui.verticalElementsCheck.onClick =
            ui.diagonalElementsCheck.onClick = ui.groupItemsCheck.onClick = function () {
                ui.extraHLinesInput.enabled = ui.extraHLinesCheck.enabled && ui.extraHLinesCheck.value;
                ui.outerVLinesInput.enabled = ui.outerVLinesCheck.enabled && ui.outerVLinesCheck.value;
                syncDivisionFields(ui);
                refreshPreview(ui, context);
            };

        ui.clearSpaceCheck.onClick = function () {
            updateClearSpaceState(ui);
            refreshPreview(ui, context);
        };

        ui.emphasizeBoundsCheck.onClick = function () {
            refreshPreview(ui, context);
        };

        ui.convertToGuidesCheck.onClick = function () {
            updateGroupState(ui);
            refreshPreview(ui, context);
        };

        ui.btnOK.onClick = function () {
            removePreview();
            var params = getParams(ui);
            if (!params) {
                alert(getLabel(LABELS.alert.invalidInput));
                return;
            }
            context.guideLayer = ensureGuideLayer(context, params.layerName);
            createLines(context, params, false);
            ui.dialog.close(1);
        };

        ui.btnCancel.onClick = function () {
            removePreview();
            ui.dialog.close(2);
        };

        ui.dialog.onClose = function () {
            removePreview();
            removeEmptyCreatedLayers(context);
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

        var context = buildContext(doc, doc.selection);
        if (!context) return;

        previewGroup = null;
        toggleLiveCornerAnnotator();
        try {
            var ui = createDialog();
            bindArrowKeyHandlers(ui, context);
            bindEvents(ui, context);
            applyPreset(ui, context);
            ui.dialog.show();
        } finally {
            toggleLiveCornerAnnotator();
        }
    }

    main();

})();
