#target illustrator
#targetengine "MasonryLayoutEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
SCRIPTMETA-BEGIN
Script-ID=com.swwwitch.illustrator.MasonryLayout
Name=オブジェクトをメイソンリーレイアウトで並べる
Author=David DeSandro
Target-App=Illustrator
Version=1.1.0
Release-Date=2026-11-07
Description-BEGIN
選択したオブジェクトを、いちばん短い列へ順に詰めていくメイソンリーレイアウトで並べ直します。
基準は選択オブジェクトか現在のアートボードで、アートボードのときは余白の内側に収まるように列数と大きさを決めます。幅を列にそろえたり、順序をランダムにしたりもできます。

Rearranges the selected objects in a masonry layout, placing each one in the shortest column.
The reference is the selected objects or the current artboard; for the artboard, the column count and size are chosen to fit inside the margin. You can also match widths to the column and randomize the order.
Description-END
SCRIPTMETA-END
*/

/*

### 概要

選択したオブジェクトを、いちばん短い列へ順に詰めていくメイソンリーレイアウトで並べ直します。
基準は選択オブジェクトか現在のアートボードで、アートボードのときは余白の内側に収まるように列数と大きさを決めます。幅を列にそろえたり、順序をランダムにしたりもできます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/MasonryLayout.md

### Overview

Rearranges the selected objects in a masonry layout, placing each one in the shortest column.
The reference is the selected objects or the current artboard; for the artboard, the column count and size are chosen to fit inside the margin. You can also match widths to the column and randomize the order.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/MasonryLayout.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "MasonryLayout";                /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "David DeSandro";               /* 作者 / author */
var SCRIPT_MODIFIED = "Masahiro Takano (@swwwitch)";  /* 改変 / modified by */
var SCRIPT_RELEASED = "2026-11-07";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-11-07";                   /* 更新日 / last updated */

var SCRIPT_DESCRIPTION_JA = "オブジェクトをメイソンリーレイアウトで並べる"; /* README の見出し / README title (Japanese) */

var SCRIPT_README_JA  = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/MasonryLayout.md"; /* README（日本語） */
var SCRIPT_README_EN  = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/MasonryLayout.md"; /* README (English) */
var SCRIPT_SOURCE_URL = "https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/alignment/MasonryLayout.jsx"; /* 最新版のソース / latest source */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

/**
 * @author David DeSandro
 * @discussion Masonry v3.3.2（MIT License）の配置計算を Illustrator のオブジェクト向けに移植した。
 * Ported from the layout calculation of Masonry v3.3.2 (MIT License).
 * http://masonry.desandro.com
 */

/*
Masonry
Copyright 2015 David DeSandro

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in
all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
THE SOFTWARE.
*/

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================
    var MIN_ITEM_COUNT = 2;      /* 必要な数 / items needed */
    var MAX_COLUMN_COUNT = 50;   /* 列数の上限 / highest column count */
    var DEFAULT_GAP = 0;         /* 間隔の初期値（定規の単位）。Masonry の gutter の既定値と同じ / initial gap in ruler units, as Masonry's default gutter */
    var DEFAULT_LINK_GAPS = true; /* 左右と上下の間隔を連動して開く / open with the two gaps linked */
    var DEFAULT_MARGIN = 0;      /* アートボードの余白の初期値（定規の単位） / initial artboard margin in ruler units */
    var MAX_MARGIN_RATIO = 0.45; /* 余白の上限（アートボードの短辺に対する比率） / highest margin, as a ratio of the artboard's short side */
    var GAP_STEP = 1;            /* 間隔・列の幅の∧∨・↑↓の増減量 / stepper increment for gaps and column width */
    var MIN_COLUMN_WIDTH = 0.01; /* 列の幅の下限（定規の単位） / lowest column width in ruler units */

    // =========================================
    // レイアウト / Layout
    // =========================================
    var FIELD_LABEL_WIDTHS = { ja: 56, en: 92 }; /* 項目名の幅（px） / label width (px) */

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
     * 単位の値を pt に換算する
     * @param {number} value - 単位の値
     * @param {{pointsPerUnit: number}} unit - getUnitInfo() の戻り値か UNITS[i]
     * @returns {number} pt の値
     */
    function unitToPt(value, unit) {
        return value * unit.pointsPerUnit;
    }

    /**
     * pt の値を単位の値に換算する
     * @param {number} ptValue - pt の値
     * @param {{pointsPerUnit: number}} unit - getUnitInfo() の戻り値か UNITS[i]
     * @returns {number} 単位の値
     */
    function ptToUnit(ptValue, unit) {
        return ptValue / unit.pointsPerUnit;
    }

    // =========================================
    // 共通部品 / Shared parts
    // =========================================

    // UI の明暗（再利用パーツ） / UI theme (reusable)

    /**
     * UI がダークテーマかどうかを判定する（Illustrator は uiBrightness、InDesign は uiBrightnessPreference）
     * @returns {boolean} ダークなら true。取得できない環境では false（明るいUI扱い）
     */
    function isDarkUI() {
        try {
            if (app.preferences && app.preferences.getRealPreference) {
                return app.preferences.getRealPreference("uiBrightness") <= 0.5; /* Illustrator */
            }
            return app.generalPreferences.uiBrightnessPreference <= 0.5; /* InDesign */
        } catch (e) {
            return false;
        }
    }

    // UI の明暗（再利用パーツ）ここまで / End of the reusable UI theme

    // リンクアイコン（再利用パーツ） / Link toggle (reusable)

    // -----------------------------------------
    // リンクアイコンの寸法 / Link toggle metrics
    // -----------------------------------------
    var LINK_ICON_SIZE          = [22, 22]; /* アイコンの大きさ / icon size */
    var LINK_ICON_STROKE        = 1.5;      /* 線幅 / stroke width */
    var LINK_CUT_DIRECTION      = [1, 0];   /* 連動中の左辺の切れ目の向き（水平）/ direction of the left-leg cut when linked (horizontal) */
    var LINK_HOOK_CUT_DIRECTION = [0, 1];   /* 連動中の巻き込みの切れ目の向き（垂直）/ direction of the hook cut when linked (vertical) */
    var LINK_STRAND_COUNT       = 4;        /* 切れ目の向きをそろえるための細い線の本数 / strands used to shape the cuts */
    var LINK_SLASH_CLEARANCE    = 2.2;      /* 連動OFFの斜線とフックの間（22px 基準）/ gap between the slash and the hooks when unlinked */

    // -----------------------------------------
    // リンクアイコンの配色 / Link toggle colors
    // -----------------------------------------
    var LINK_UI_DARK = isDarkUI();
    /* ダイアログの地に重ねる半透明の黒・白（UIの明るさの段階に追従する）。値はステップボタンの配色と同じ
       Translucent overlays that follow the dialog background; same values as the stepper buttons */
    var LINK_PRESSED_COLOR  = LINK_UI_DARK ? [1, 1, 1, 0.12] : [0, 0, 0, 0.13]; /* 連動中の地 / background while linked */
    var LINK_FRAME_COLOR    = LINK_UI_DARK ? [1, 1, 1, 0.07] : [0, 0, 0, 0.10]; /* 連動中の枠 / frame while linked */
    var LINK_ICON_COLOR     = LINK_UI_DARK ? [1, 1, 1, 1]    : [0, 0, 0, 0.70]; /* アイコンの線 / icon strokes */
    var LINK_DIM_ICON_COLOR = LINK_UI_DARK ? [1, 1, 1, 0.20] : [0, 0, 0, 0.25]; /* 無効時の線 / strokes when disabled */

    // -----------------------------------------
    // アイコンを作る・切り替える（外から呼ぶ関数） / Public API
    // -----------------------------------------
    /**
     * 連動の ON／OFF を切り替えるリンクアイコンを追加する（onDraw で自作描画）。
     * クリックで切り替わる。連動中は押し込んだボタンのように地と枠を描く。
     * @param {Group} parent - 追加先
     * @param {boolean} initialValue - 連動の初期値
     * @param {Function} onToggle - 切り替えたあとに呼ぶ関数
     * @param {number[]} [iconSize] - アイコンの [幅, 高さ]（省略時は LINK_ICON_SIZE。絵は 22px 基準から拡大縮小する）
     * @param {number} [chainRatio] - 鎖の絵の大きさの比率（省略時は 1。枠・地の大きさは変えず、鎖だけ縮める）
     * @returns {Group} アイコン（.value で連動中かを読む）
     */
    function addLinkToggle(parent, initialValue, onToggle, iconSize, chainRatio) {
        var toggleSize = iconSize || LINK_ICON_SIZE;
        var linkToggle = parent.add("group");
        linkToggle.preferredSize = toggleSize;
        linkToggle.minimumSize = toggleSize;
        linkToggle.maximumSize = toggleSize;
        linkToggle.value = initialValue;

        linkToggle.onDraw = function () {
            var iconGraphics = linkToggle.graphics;
            var iconWidth = toggleSize[0];
            var iconHeight = toggleSize[1];
            /* 自作描画は自動でディムにならないため、親もたどって判定する / Custom drawing is not dimmed automatically */
            var isDimmed = !isLinkToggleEnabledInTree(linkToggle);
            /* 連動中は押し込んだボタンのように地と枠を描く / While linked, draw it like a pressed button */
            if (linkToggle.value && !isDimmed) {
                iconGraphics.newPath();
                iconGraphics.rectPath(0, 0, iconWidth, iconHeight);
                iconGraphics.fillPath(iconGraphics.newBrush(iconGraphics.BrushType.SOLID_COLOR, LINK_PRESSED_COLOR));
                iconGraphics.newPath();
                iconGraphics.rectPath(0.5, 0.5, iconWidth - 1, iconHeight - 1);
                iconGraphics.strokePath(iconGraphics.newPen(iconGraphics.PenType.SOLID_COLOR, LINK_FRAME_COLOR, 1));
            }
            drawLinkIcon(iconGraphics, iconWidth, iconHeight, linkToggle.value, isDimmed ? LINK_DIM_ICON_COLOR : LINK_ICON_COLOR, chainRatio);
        };

        linkToggle.addEventListener("mousedown", function () {
            if (!isLinkToggleEnabledInTree(linkToggle)) return;
            linkToggle.value = !linkToggle.value;
            redrawLinkToggle(linkToggle);
            if (onToggle) onToggle();
        });
        return linkToggle;
    }

    /**
     * 連動の状態をコードから変えて描き直す（onToggle は呼ばない）
     * @param {Group} linkToggle - addLinkToggle() で作ったアイコン
     * @param {boolean} isLinked - 連動にするなら true
     * @returns {void}
     */
    function setLinkToggleValue(linkToggle, isLinked) {
        if (linkToggle.value === isLinked) return;
        linkToggle.value = isLinked;
        redrawLinkToggle(linkToggle);
    }

    /**
     * アイコンの有効／無効を切り替えて描き直す（変わらないときは描き直さない）
     * @param {Group} linkToggle - addLinkToggle() で作ったアイコン
     * @param {boolean} isEnabled - 有効にするなら true
     * @returns {void}
     */
    function setLinkToggleEnabled(linkToggle, isEnabled) {
        if (linkToggle.enabled === isEnabled) return;
        linkToggle.enabled = isEnabled;
        redrawLinkToggle(linkToggle);
    }

    /**
     * コントロールと親がすべて有効かを判定する（親の無効化は子の enabled に出ないため、親もたどる）
     * @param {Object} control - 判定するコントロール
     * @returns {boolean} すべて有効なら true
     */
    function isLinkToggleEnabledInTree(control) {
        for (var node = control; node; node = node.parent) {
            if (!node.enabled) return false;
        }
        return true;
    }

    /**
     * group の onDraw を呼び直す。group には notify() が無いため、隠して再表示して描き直させる
     * @param {Group} linkToggle - 描き直すアイコン
     * @returns {void}
     */
    function redrawLinkToggle(linkToggle) {
        linkToggle.hide();
        linkToggle.show();
    }

    // -----------------------------------------
    // アイコンの形 / Icon geometry
    // -----------------------------------------
    /**
     * 連動アイコンを描く。Illustrator の［縦横比を固定］に合わせ、連動中は縦につながったチェーン、
     * 連動していないときは上下に分かれたチェーンに斜線を重ねる。座標は 22px 四方を基準に拡大縮小する。
     * @param {ScriptUIGraphics} iconGraphics - 描画先
     * @param {number} iconWidth - 描画範囲の幅
     * @param {number} iconHeight - 描画範囲の高さ
     * @param {boolean} isLinked - 連動中なら true
     * @param {number[]} iconColor - [r, g, b, a]
     * @param {number} [chainRatio] - 鎖の大きさの比率（省略時は 1）。中央に置いたまま縮める
     * @returns {void}
     */
    function drawLinkIcon(iconGraphics, iconWidth, iconHeight, isLinked, iconColor, chainRatio) {
        var iconScale = Math.min(iconWidth, iconHeight) / 22 * (chainRatio || 1);
        var offsetX = (iconWidth - 22 * iconScale) / 2;
        var offsetY = (iconHeight - 22 * iconScale) / 2;
        var strokes = isLinked ? buildLinkedChainStrokes() : buildUnlinkedChainStrokes();
        for (var i = 0; i < strokes.length; i++) {
            var strokePoints = strokes[i].points;
            /* newPath() を呼ばないとパスが前の描画に積み重なる / Without newPath() the paths accumulate */
            iconGraphics.newPath();
            for (var j = 0; j < strokePoints.length; j++) {
                var pointX = offsetX + strokePoints[j][0] * iconScale;
                var pointY = offsetY + strokePoints[j][1] * iconScale;
                if (j === 0) iconGraphics.moveTo(pointX, pointY);
                else iconGraphics.lineTo(pointX, pointY);
            }
            iconGraphics.strokePath(iconGraphics.newPen(iconGraphics.PenType.SOLID_COLOR, iconColor, strokes[i].width * iconScale));
        }
    }

    /**
     * 連動中のチェーン（縦に組み合った2つの輪）の線を返す。
     * 上の輪は左辺の途中から上端を回って右辺を下り、下端で内側へ巻き込む。下の輪はそれを180度回したもの。
     * 切れ目の向きをそろえるため、輪を細い線の束にし、両端を延ばしてから直線で切る（左辺は水平、巻き込みは垂直）
     * @returns {Array<{points: Array<number[]>, width: number}>} 線ごとの点列と線幅（22px 四方の座標）
     */
    function buildLinkedChainStrokes() {
        /* 左辺は上端の丸みだけ残して短く切り、下の輪の巻き込みとの間を空ける
           Keep only a stub on the left so it stays clear of the lower ring's hook */
        var upperRing = densifyPoints(buildArcPoints(11, 7, 3.5, 3.5, 180, 360)
            .concat([[14.5, 11.2]])
            .concat(buildArcPoints(11, 11.2, 3.5, 2.3, 0, 115)));
        var ringStart = upperRing[0];
        var ringEnd = upperRing[upperRing.length - 1];
        var extendedRing = extendPolylineEnds(upperRing, LINK_ICON_STROKE);
        /* 延ばした先がどちら側かで、切り捨てる側を決める / The extended tips tell which side to cut away */
        var startOutsideSign = sideOfLine(extendedRing[0], ringStart, LINK_CUT_DIRECTION);
        var endOutsideSign = sideOfLine(extendedRing[extendedRing.length - 1], ringEnd, LINK_HOOK_CUT_DIRECTION);

        var upperStrands = buildStrandStrokes(extendedRing, function (strandPoints) {
            var trimmed = trimPolylineTail(strandPoints, ringEnd, LINK_HOOK_CUT_DIRECTION, endOutsideSign);
            trimmed = trimPolylineTail(trimmed.reverse(), ringStart, LINK_CUT_DIRECTION, startOutsideSign).reverse();
            return [trimmed];
        });
        var strokes = [];
        for (var i = 0; i < upperStrands.length; i++) {
            strokes.push(upperStrands[i]);
            strokes.push({ points: rotatePointsHalfTurn(upperStrands[i].points), width: upperStrands[i].width });
        }
        return strokes;
    }

    /**
     * 中心線を線幅の中で等分した細い線に分け、clipStrand で切った結果を線として返す。
     * @param {Array<number[]>} centerline - 中心線の点列
     * @param {Function} clipStrand - 細い線の点列を受け取り、残す点列の配列を返す関数
     * @returns {Array<{points: Array<number[]>, width: number}>} 細い線ごとの点列と線幅
     */
    function buildStrandStrokes(centerline, clipStrand) {
        var strandWidth = LINK_ICON_STROKE / LINK_STRAND_COUNT;
        var strokes = [];
        for (var k = 0; k < LINK_STRAND_COUNT; k++) {
            /* 線幅の中を等分した位置に細い線を並べる / Lay the strands evenly across the stroke width */
            var strandOffset = -LINK_ICON_STROKE / 2 + strandWidth * (k + 0.5);
            var strandPieces = clipStrand(offsetPolyline(centerline, strandOffset));
            for (var j = 0; j < strandPieces.length; j++) {
                /* 隣の線と少し重ねて隙間を埋める / Overlap neighbours slightly so no seams show */
                if (strandPieces[j].length > 1) strokes.push({ points: strandPieces[j], width: strandWidth * 1.4 });
            }
        }
        return strokes;
    }

    /**
     * 点列の両端を、端の向きのまま length だけ延ばす。
     * @param {Array<number[]>} points - 点列
     * @param {number} length - 延ばす長さ
     * @returns {Array<number[]>} 延ばした点列
     */
    function extendPolylineEnds(points, length) {
        /* from から to の向きへ、to から length 先の点 / point length beyond to, heading from from to to */
        function extendBeyond(from, to) {
            var dx = to[0] - from[0];
            var dy = to[1] - from[1];
            var segmentLength = Math.sqrt(dx * dx + dy * dy) || 1;
            return [to[0] + dx / segmentLength * length, to[1] + dy / segmentLength * length];
        }
        var lastIndex = points.length - 1;
        return [extendBeyond(points[1], points[0])].concat(points, [extendBeyond(points[lastIndex - 1], points[lastIndex])]);
    }

    /**
     * 点が直線のどちら側にあるかを符号で返す。
     * @param {number[]} point - 点
     * @param {number[]} linePoint - 直線上の1点
     * @param {number[]} direction - 直線の向き
     * @returns {number} 正・負で側を表す値
     */
    function sideOfLine(point, linePoint, direction) {
        return direction[0] * (point[1] - linePoint[1]) - direction[1] * (point[0] - linePoint[0]);
    }

    /**
     * 点列の終わり側で、直線より outsideSign の側にはみ出した部分を切り、直線との交点で止める。
     * 輪の別の場所が同じ直線をまたいでも切らないよう、終わりから数点の範囲だけを見る。
     * @param {Array<number[]>} points - 点列
     * @param {number[]} cutPoint - 切る直線上の1点
     * @param {number[]} direction - 切る直線の向き
     * @param {number} outsideSign - 切り捨てる側の符号
     * @returns {Array<number[]>} 切った点列
     */
    function trimPolylineTail(points, cutPoint, direction, outsideSign) {
        var lastIndex = points.length - 1;
        var searchLimit = Math.max(0, lastIndex - 12);
        var index = lastIndex;
        while (index > searchLimit && sideOfLine(points[index], cutPoint, direction) * outsideSign > 0) index--;
        if (index === lastIndex) return points.slice(0);
        var inside = points[index];
        var outside = points[index + 1];
        var insideSide = sideOfLine(inside, cutPoint, direction);
        var ratio = insideSide / (insideSide - sideOfLine(outside, cutPoint, direction));
        return points.slice(0, index + 1).concat([[inside[0] + (outside[0] - inside[0]) * ratio, inside[1] + (outside[1] - inside[1]) * ratio]]);
    }

    /**
     * 連動していないときのチェーン（上下に分かれた輪と斜線）の線を返す。
     * フックは斜線の近くで切る。線の端は進む向きに直角にしか切れないため、フックを細い線の束にして
     * 1本ずつ斜線と平行な境界で切り、切り口が斜線に沿って見えるようにする。
     * @returns {Array<{points: Array<number[]>, width: number}>} 線ごとの点列と線幅（22px 四方の座標）
     */
    function buildUnlinkedChainStrokes() {
        var slashStart = [3.5, 3.5];
        var slashEnd = [18.5, 18.5];
        var upperHook = densifyPoints(buildArcPoints(11, 7, 3.5, 3.5, 180, 360).concat([[14.5, 11.5]]));
        var hooks = [upperHook, rotatePointsHalfTurn(upperHook)];

        /* 斜線の近くの帯を切り取る / Cut away the band around the slash */
        function clipAroundSlash(strandPoints) {
            return clipOutsideBand(strandPoints, slashStart, slashEnd, LINK_SLASH_CLEARANCE);
        }
        var strokes = buildStrandStrokes(hooks[0], clipAroundSlash).concat(buildStrandStrokes(hooks[1], clipAroundSlash));
        strokes.push({ points: [slashStart, slashEnd], width: LINK_ICON_STROKE });
        return strokes;
    }

    /**
     * 点の間隔が 0.5 以下になるよう、線分の間に点を足す。
     * @param {Array<number[]>} points - 点列
     * @returns {Array<number[]>} 細かくした点列
     */
    function densifyPoints(points) {
        var densePoints = [points[0]];
        for (var i = 1; i < points.length; i++) {
            var from = points[i - 1];
            var to = points[i];
            var steps = Math.max(1, Math.ceil(Math.sqrt(Math.pow(to[0] - from[0], 2) + Math.pow(to[1] - from[1], 2)) / 0.5));
            for (var j = 1; j <= steps; j++) {
                densePoints.push([from[0] + (to[0] - from[0]) * j / steps, from[1] + (to[1] - from[1]) * j / steps]);
            }
        }
        return densePoints;
    }

    /**
     * 点列を、進む向きの左側へ offset だけずらした点列を返す（負の値なら右側）。
     * @param {Array<number[]>} points - 点列
     * @param {number} offset - ずらす距離
     * @returns {Array<number[]>} ずらした点列
     */
    function offsetPolyline(points, offset) {
        var shifted = [];
        for (var i = 0; i < points.length; i++) {
            var before = points[Math.max(0, i - 1)];
            var after = points[Math.min(points.length - 1, i + 1)];
            var tangentX = after[0] - before[0];
            var tangentY = after[1] - before[1];
            var tangentLength = Math.sqrt(tangentX * tangentX + tangentY * tangentY) || 1;
            shifted.push([points[i][0] - tangentY / tangentLength * offset, points[i][1] + tangentX / tangentLength * offset]);
        }
        return shifted;
    }

    /**
     * 直線（線分を延長したもの）から clearance 未満の帯に入る部分を切り取り、残りを点列に分けて返す。
     * 帯の境界で線分を補間して切るので、切り口は直線と平行にそろう。
     * @param {Array<number[]>} points - 点列
     * @param {number[]} lineStart - 直線上の1点
     * @param {number[]} lineEnd - 直線上のもう1点
     * @param {number} clearance - 空ける距離
     * @returns {Array<Array<number[]>>} 帯の外側に残った点列（2点未満のものは除く）
     */
    function clipOutsideBand(points, lineStart, lineEnd, clearance) {
        var directionX = lineEnd[0] - lineStart[0];
        var directionY = lineEnd[1] - lineStart[1];
        var directionLength = Math.sqrt(directionX * directionX + directionY * directionY);

        /* 直線からの符号付き距離 / signed distance from the line */
        function signedDistance(point) {
            return (directionX * (point[1] - lineStart[1]) - directionY * (point[0] - lineStart[0])) / directionLength;
        }
        /* 2点の間で、距離が boundary になる点 / point between two points where the distance equals boundary */
        function interpolateAt(from, to, fromDistance, toDistance, boundary) {
            var ratio = (boundary - fromDistance) / (toDistance - fromDistance);
            return [from[0] + (to[0] - from[0]) * ratio, from[1] + (to[1] - from[1]) * ratio];
        }

        var pieces = [];
        var currentPiece = [];
        for (var i = 0; i < points.length; i++) {
            var distance = signedDistance(points[i]);
            var isOutside = Math.abs(distance) >= clearance;
            if (i > 0) {
                var previousDistance = signedDistance(points[i - 1]);
                var wasOutside = Math.abs(previousDistance) >= clearance;
                if (wasOutside && !isOutside) {
                    /* 帯に入る: 境界で止める / entering the band: stop at the boundary */
                    currentPiece.push(interpolateAt(points[i - 1], points[i], previousDistance, distance, previousDistance > 0 ? clearance : -clearance));
                    if (currentPiece.length > 1) pieces.push(currentPiece);
                    currentPiece = [];
                } else if (!wasOutside && isOutside) {
                    /* 帯から出る: 境界から始める / leaving the band: start at the boundary */
                    currentPiece = [interpolateAt(points[i - 1], points[i], previousDistance, distance, distance > 0 ? clearance : -clearance)];
                }
            }
            if (isOutside) currentPiece.push(points[i]);
        }
        if (currentPiece.length > 1) pieces.push(currentPiece);
        return pieces;
    }

    /**
     * 楕円弧の点列を返す（角度は右が0度、下が90度の画面座標）。
     * @param {number} centerX - 中心X
     * @param {number} centerY - 中心Y
     * @param {number} radiusX - 横の半径
     * @param {number} radiusY - 縦の半径
     * @param {number} startDegrees - 開始角度
     * @param {number} endDegrees - 終了角度
     * @returns {Array<number[]>} 点列
     */
    function buildArcPoints(centerX, centerY, radiusX, radiusY, startDegrees, endDegrees) {
        var arcSteps = 12;
        var arcPoints = [];
        for (var i = 0; i <= arcSteps; i++) {
            var angle = (startDegrees + (endDegrees - startDegrees) * i / arcSteps) * Math.PI / 180;
            arcPoints.push([centerX + radiusX * Math.cos(angle), centerY + radiusY * Math.sin(angle)]);
        }
        return arcPoints;
    }

    /**
     * 点列を 22px 四方の中心で180度回す。
     * @param {Array<number[]>} points - 点列
     * @returns {Array<number[]>} 回した点列
     */
    function rotatePointsHalfTurn(points) {
        var rotated = [];
        for (var i = 0; i < points.length; i++) {
            rotated.push([22 - points[i][0], 22 - points[i][1]]);
        }
        return rotated;
    }

    // リンクアイコン（再利用パーツ）ここまで / End of the reusable link toggle

    // ステップボタン（再利用パーツ） / Stepper buttons (reusable)

    // -----------------------------------------
    // ステップボタンの寸法・増減量 / Stepper metrics and steps
    // -----------------------------------------
    var STEPPER_BUTTON_WIDTH   = 20;  /* ∧∨ボタンの幅 / button width */
    var STEPPER_BUTTON_HEIGHT  = 11;  /* ∧∨ボタン1つの高さ（2つ重ねた全体の高さは22） / button height (22 for the pair) */
    var STEPPER_CORNER_RADIUS  = 2;   /* 枠の角丸の半径（ScriptUIは円弧を描けないため短い線分で近似） / corner radius, approximated with segments */
    var STEPPER_FIELD_SPACING  = 3;   /* 項目名と∧∨の間隔 / spacing between the label and the stepper */
    var STEPPER_SIDE_MARGIN    = 3;   /* ∧∨の左に足す余白（右は入力欄に突き合わせる） / extra space left of the stepper */
    var STEPPER_SHIFT_MULTIPLE = 10;  /* shift＋クリックでそろえる倍数 / Shift-click snaps to multiples of this */
    var STEPPER_OPTION_STEP    = 0.1; /* option＋クリックの増減量 / Option-click step */

    /* 単位の換算は UnitValue に任せる（in / ft / yd / mm / cm / m / pt / pc / px ほか、単数形・複数形も可）。
       UnitValue に無い単位だけ、ここで UnitValue の単位に読み替える（値は「1単位＝何 unit か」）。
       「p」は「1p6」（1パイカ6ポイント）の形にも使う
       Units UnitValue lacks, mapped onto UnitValue units (how many of `unit` make one) */
    var STEPPER_UNIT_ALIASES = {
        "q": { unit: "mm", amount: 0.25 },    /* 級 / Q */
        "h": { unit: "mm", amount: 0.25 },    /* 歯 / H */
        "p": { unit: "pc", amount: 1 },       /* パイカ / pica */
        "ft/in": { unit: "ft", amount: 1 },   /* Illustrator の単位コード7の表示 / Illustrator unit code 7 */
        "c": { unit: "ci", amount: 1 },       /* シセロ（InDesign の表示） / ciceros as InDesign shows them */
        "ag": { unit: "in", amount: 1 / 14 }, /* アゲート / agates */
        "ap": { unit: "tpt", amount: 1 }      /* アメリカンポイント / American points */
    };

    // -----------------------------------------
    // ステップボタンの配色 / Stepper colors
    // -----------------------------------------
    var STEPPER_UI_DARK           = isDarkUI();
    /* UIの明るさは4段階あり、段階ごとに背景色が違う。どの段階でも背景に対する差で見せるよう、黒・白の半透明を重ねる。
       ダーク側は Illustrator 標準のスピナー（［グリッドに分割］）で実測、明るい側は最も明るい段階（背景 約0.94）から逆算
       UI brightness has four levels with different backgrounds, so colors are translucent overlays that follow the
       dialog background. Dark values are measured from Illustrator's own spinner; light values derived for the lightest level */
    var STEPPER_FILL_COLOR        = STEPPER_UI_DARK ? [0, 0, 0, 0.10]  : [1, 1, 1, 0.50];  /* 地 / background */
    var STEPPER_FRAME_COLOR       = STEPPER_UI_DARK ? [1, 1, 1, 0.07]  : [0, 0, 0, 0.10];  /* 枠線 / frame */
    var STEPPER_PRESSED_COLOR     = STEPPER_UI_DARK ? [1, 1, 1, 0.12]  : [0, 0, 0, 0.13];  /* 押下中 / pressed */
    var STEPPER_CHEVRON_COLOR     = STEPPER_UI_DARK ? [1, 1, 1, 1]     : [0, 0, 0, 0.70];  /* 山形の線 / chevron */
    var STEPPER_DIM_FILL_COLOR    = STEPPER_UI_DARK ? [1, 1, 1, 0.035] : [1, 1, 1, 0.30];  /* 無効時の地 / background when disabled */
    var STEPPER_DIM_FRAME_COLOR   = STEPPER_UI_DARK ? [1, 1, 1, 0.035] : [0, 0, 0, 0.05];  /* 無効時の枠線（ダークは地と同じで見せない） / frame when disabled */
    var STEPPER_DIM_CHEVRON_COLOR = STEPPER_UI_DARK ? [1, 1, 1, 0.20]  : [0, 0, 0, 0.25];  /* 無効時の山形 / chevron when disabled */

    // -----------------------------------------
    // 数値欄を作る（外から呼ぶ関数） / Public API
    // -----------------------------------------
    /**
     * 「項目名・∧∨・入力欄」をひと組にした数値欄を追加する。
     * ↑↓キーでも∧∨と同じように増減する。直接入力した値も、フォーカスが外れたときに
     * 整数化・下限・上限・単位（「20 mm」の形）へそろえ、数値でなければ直前の値に戻す。
     * 四則演算（+ - * / と括弧）を入れると、確定時に計算した値にする。欄と違う単位で入れた値は欄の単位へ換算する（mm の欄に「1 in」→「25.4 mm」）
     * @param {Group|Panel} parent - 追加先
     * @param {Object} fieldOptions - label（コロン込みの項目名）/ labelWidth / text / characters /
     *     step / min / max / integer（true で整数のみ）/ unit / onStep
     * @returns {EditText} 入力欄（項目名は .fieldLabel、∧∨は .stepperGroup で参照できる）
     */
    function addSteppedField(parent, fieldOptions) {
        var fieldRowGroup = parent.add("group");
        fieldRowGroup.orientation = "row";
        fieldRowGroup.alignChildren = ["left", "center"];
        fieldRowGroup.spacing = STEPPER_FIELD_SPACING;

        var fieldLabel = fieldRowGroup.add("statictext", undefined, fieldOptions.label || "");
        if (fieldOptions.labelWidth) {
            fieldLabel.preferredSize.width = fieldOptions.labelWidth;
            fieldLabel.justify = "right";
        }

        /* ∧∨と入力欄は隙間0で突き合わせる / butt the stepper against the field */
        var stepperInputGroup = fieldRowGroup.add("group");
        stepperInputGroup.orientation = "row";
        stepperInputGroup.alignChildren = ["left", "center"];
        stepperInputGroup.spacing = 0;
        stepperInputGroup.margins = 0;

        var numberInput;
        var stepperGroup = addStepper(stepperInputGroup, function () { return numberInput; }, fieldOptions);
        numberInput = stepperInputGroup.add("edittext", undefined, fieldOptions.text || "");
        numberInput.characters = fieldOptions.characters || 6;
        numberInput.fieldLabel = fieldLabel;
        numberInput.stepperGroup = stepperGroup;

        /* ↑↓キーも∧∨と同じ処理で増減する（増減量・下限・上限・単位・修飾キーをそろえる） / arrow keys share the stepper's logic */
        bindSteppedArrowKeys(numberInput, stepperGroup);

        /* 項目名のクリックで入力欄にフォーカスを移す / clicking the label focuses the field */
        fieldLabel.addEventListener("click", function () { focusNumberInput(numberInput); });

        /* 直接入力をそろえる。計算式は計算し、数値でなければ直前の値に戻す / normalize typed values; evaluate arithmetic, revert non-numbers */
        numberInput.lastValidText = numberInput.text;
        numberInput.onChange = function () {
            var value = evaluateArithmetic(numberInput.text, fieldOptions.unit);
            if (isNaN(value)) {
                numberInput.text = numberInput.lastValidText;
                return;
            }
            writeSteppedValue(numberInput, value, fieldOptions);
        };
        return numberInput;
    }

    /**
     * 数値欄の有効／無効を、項目名・∧∨ごとまとめて切り替える
     * @param {EditText} numberInput - addSteppedField() で作った入力欄
     * @param {boolean} isEnabled - 有効にするなら true
     * @returns {void}
     */
    function setSteppedFieldEnabled(numberInput, isEnabled) {
        numberInput.enabled = isEnabled;
        numberInput.fieldLabel.enabled = isEnabled;
        numberInput.stepperGroup.enabled = isEnabled;
        /* ∧∨は自作描画なので、描き直してディム表示を切り替える / redraw the custom-drawn buttons to update the dimming */
        for (var i = 0; i < numberInput.stepperGroup.children.length; i++) {
            redrawStepperGroup(numberInput.stepperGroup.children[i]);
        }
    }

    /**
     * 数値欄の単位を差し替える（単位の設定やドロップダウンを切り替えたとき用）。
     * shouldConvert が true なら値を新しい単位へ換算し（10 mm → 28.35 pt）、false なら数値はそのままで単位だけ付け替える
     * @param {EditText} numberInput - addSteppedField() で作った入力欄、または bindSteppedArrowKeys() を呼んだ入力欄
     * @param {string} unit - 新しい単位（例 " pt"。単位なしは ""）
     * @param {boolean} [shouldConvert] - 値も換算するなら true
     * @returns {void}
     */
    function setSteppedFieldUnit(numberInput, unit, shouldConvert) {
        var stepOptions = numberInput.stepperGroup.stepOptions;
        var oldUnit = stepOptions.unit || "";
        var value = parseFloat(numberInput.text);
        stepOptions.unit = unit;
        if (isNaN(value)) return;
        if (shouldConvert) {
            var converted = evaluateArithmetic(String(value) + oldUnit, unit);
            if (!isNaN(converted)) value = converted;
        }
        numberInput.text = formatStepperNumber(value) + unit;
        numberInput.lastValidText = numberInput.text;
    }

    /**
     * 入力欄の値を増減する∧∨ボタンを、隙間なく縦に積んで追加する
     * @param {Group|Panel} parent - 追加先
     * @param {Function} getNumberInput - 対象の入力欄を返す関数（入力欄を∧∨より後に作れるよう、クリック時に引く）
     * @param {Object} stepOptions - step（増減量）/ min / max / integer / unit（例 " mm"）/ onStep(numberInput)
     * @returns {Group} ∧∨をまとめた group（.stepBy(direction) で同じ増減を呼べる）
     */
    function addStepper(parent, getNumberInput, stepOptions) {
        var stepperGroup = parent.add("group");
        stepperGroup.orientation = "column";
        stepperGroup.spacing = 0; /* 2つのボタンをつなげて1つの枠に見せる / join the buttons into one frame */
        stepperGroup.margins = [STEPPER_SIDE_MARGIN, 0, 0, 0]; /* 右は入力欄に突き合わせる / butt against the field on the right */
        stepperGroup.alignment = ["left", "center"];

        /**
         * 入力欄の値を増減する（shift を押しながらなら STEPPER_SHIFT_MULTIPLE の倍数へ、option なら STEPPER_OPTION_STEP ずつ。下限・上限で止める）
         * @param {number} direction - 増やすなら 1、減らすなら -1
         * @returns {void}
         */
        function stepBy(direction) {
            var numberInput = getNumberInput();
            if (!isStepperEnabledInTree(numberInput)) return; /* 入力欄か親が無効の間は動かさない */
            var value = evaluateArithmetic(numberInput.text, stepOptions.unit); /* 確定前の計算式も計算してから増減 / evaluate an uncommitted expression first */
            if (isNaN(value)) value = parseFloat(numberInput.text); /* 計算できなければ従来どおり先頭の数値 / fall back to the leading number */
            if (isNaN(value)) value = 0;
            writeSteppedValue(numberInput, computeSteppedValue(value, direction, stepOptions), stepOptions);
            if (stepOptions.onStep) stepOptions.onStep(numberInput);
        }

        /**
         * ∧∨を離したときに入力欄へフォーカスを移す（mousedown で移しても、離したときに外れる）
         * @param {Group} chevronButton - makeStepperChevronButton() で作ったボタン
         * @returns {Group} 渡したボタン
         */
        function focusInputOnRelease(chevronButton) {
            chevronButton.addEventListener("mouseup", function () {
                var numberInput = getNumberInput();
                if (isStepperEnabledInTree(numberInput)) focusNumberInput(numberInput);
            });
            return chevronButton;
        }

        /* 整数の欄では option＋クリックの0.1刻みが効かないので、説明から外す / integer fields have no 0.1 step */
        var upTooltip = stepOptions.integer ? LABELS.tooltip.stepUpInteger : LABELS.tooltip.stepUp;
        var downTooltip = stepOptions.integer ? LABELS.tooltip.stepDownInteger : LABELS.tooltip.stepDown;
        focusInputOnRelease(makeStepperChevronButton(stepperGroup, "up", function () { stepBy(1); })).helpTip = getLabel(upTooltip);
        focusInputOnRelease(makeStepperChevronButton(stepperGroup, "down", function () { stepBy(-1); })).helpTip = getLabel(downTooltip);
        stepperGroup.stepBy = stepBy; /* ↑↓キーからも同じ処理で増減できるよう公開 / shared with the arrow keys */
        stepperGroup.stepOptions = stepOptions; /* 確定時の計算で欄の単位を引けるよう公開 / lets the commit-time evaluation find the unit */
        return stepperGroup;
    }

    /**
     * 入力欄の↑↓キーを、∧∨と同じ処理で増減させる。ほかのキーは素通し。
     * あわせて、確定時に計算式・単位付きの値を計算して書き戻す（各スクリプトの onChange より先に呼ばれるので、onChange は計算後の値を読む）
     * @param {EditText} numberInput - 対象の入力欄
     * @param {Group} stepperGroup - addStepper() で作った∧∨
     * @returns {void}
     */
    function bindSteppedArrowKeys(numberInput, stepperGroup) {
        numberInput.addEventListener("keydown", function (event) {
            if (event.keyName !== "Up" && event.keyName !== "Down") return;
            stepperGroup.stepBy(event.keyName === "Up" ? 1 : -1);
            event.preventDefault(); /* カーソル移動を止める / keep the caret from moving */
        });
        numberInput.addEventListener("change", function () {
            var fieldUnit = stepperGroup.stepOptions ? stepperGroup.stepOptions.unit : undefined;
            var value = evaluateArithmetic(numberInput.text, fieldUnit);
            if (isNaN(value)) return; /* 計算できなければ各スクリプトの処理に任せる / leave it to the script's own handler */
            /* 式か、換算で値が変わったときだけ書き戻す（ただの数値は書式を崩さない） / rewrite only expressions and converted values */
            var hasOperator = /[*\/()\u00D7\u00F7\uFF0A\uFF0F\uFF08\uFF09]|[\d.\uFF10-\uFF19][^\d.\uFF10-\uFF19]*[+\-\u2212\uFF0B\uFF0D]/.test(numberInput.text);
            if (!hasOperator && value === parseFloat(numberInput.text)) {
                /* 単位を省いて入れた数値には、欄の単位だけ付け足す（桁は丸めない） / append the field unit to a bare number */
                var trimmedText = numberInput.text.replace(/^\s+|\s+$/g, "");
                if (fieldUnit && /[\d.]$/.test(trimmedText)) numberInput.text = trimmedText + fieldUnit;
                return;
            }
            numberInput.text = formatStepperNumber(value) + (fieldUnit || "");
        });
        numberInput.stepperGroup = stepperGroup; /* setSteppedFieldUnit() から∧∨の設定を引けるようにする / lets setSteppedFieldUnit() find the options */
    }

    // -----------------------------------------
    // 値の計算 / Value helpers
    // -----------------------------------------
    /**
     * 押された修飾キーに応じて、1回分増減した値を返す
     * （shift なら STEPPER_SHIFT_MULTIPLE の倍数へ、option なら STEPPER_OPTION_STEP ずつ、それ以外は step の倍数へ（1.5→2、1.5→1）。
     * 整数の欄では option を無視して step の倍数へ）
     * @param {number} value - 元の値
     * @param {number} direction - 増やすなら 1、減らすなら -1
     * @param {Object} stepOptions - step（通常の増減量。省略時は 1）/ integer
     * @returns {number} 増減した値（下限・上限は未適用）
     */
    function computeSteppedValue(value, direction, stepOptions) {
        var keyState = ScriptUI.environment.keyboardState;
        if (keyState.shiftKey) return snapStepperToNextMultiple(value, STEPPER_SHIFT_MULTIPLE, direction);
        if (keyState.altKey && !stepOptions.integer) return value + direction * STEPPER_OPTION_STEP;
        return snapStepperToNextMultiple(value, stepOptions.step || 1, direction);
    }

    /**
     * 値を、指定した方向にある次の倍数へ移す（230→240、232→240、下げるときは 232→230、230→220）
     * @param {number} value - 元の値
     * @param {number} multiple - 倍数の単位（例 10）
     * @param {number} direction - 上げるなら 1、下げるなら -1
     * @returns {number} 移した値
     */
    function snapStepperToNextMultiple(value, multiple, direction) {
        /* 0.29 / 0.01 = 28.999… のような浮動小数の誤差で同じ値に戻らないよう、商を丸めてから切り捨て・切り上げる
           round the quotient first so float error (0.29 / 0.01 = 28.999…) does not step back to the same value */
        var quotient = Math.round(value / multiple * 1e6) / 1e6;
        if (direction > 0) return Math.round((Math.floor(quotient) + 1) * multiple * 1e6) / 1e6;
        return Math.round((Math.ceil(quotient) - 1) * multiple * 1e6) / 1e6;
    }

    /**
     * 値を下限・上限の範囲に収める
     * @param {number} value - 数値
     * @param {Object} rangeOptions - min / max（どちらも省略可）
     * @returns {number} 範囲に収めた値
     */
    function clampSteppedValue(value, rangeOptions) {
        if (rangeOptions.min !== undefined && value < rangeOptions.min) return rangeOptions.min;
        if (rangeOptions.max !== undefined && value > rangeOptions.max) return rangeOptions.max;
        return value;
    }

    /**
     * 入力欄の文字列を四則演算（+ - * / と括弧）として計算する。eval は使わない。
     * 数値の後ろの単位は UnitValue で欄の単位へ換算する（mm の欄に「1in」→ 25.4、「1p6」は1パイカ6ポイント）。単位のない数値は欄の単位とみなす。
     * 全角の数字・記号と × ÷ は半角に直す
     * @param {string} text - 入力欄の文字列
     * @param {string} [fieldUnit] - 欄の単位（例 " mm"。前後の空白は無視）
     * @returns {number} 欄の単位での計算結果（式として読めない・換算できない単位・0で割ったときは NaN）
     */
    function evaluateArithmetic(text, fieldUnit) {
        var source = String(text)
            .replace(/[！-～]/g, function (ch) { return String.fromCharCode(ch.charCodeAt(0) - 0xFEE0); })
            .replace(/×/g, "*")
            .replace(/÷/g, "/")
            .replace(/[−–—]/g, "-")
            .replace(/\s/g, "");
        if (source === "") return NaN;
        var fieldUnitKey = String(fieldUnit || "").replace(/^\s+|\s+$/g, "").toLowerCase();
        var fieldUnitValue = createStepperUnitValue(1, fieldUnitKey); /* 欄の単位の1単位（換算できない欄は null） / one field unit */
        var position = 0;

        /**
         * 加減算の並び（項 ± 項 …）を読む
         * @returns {number} 値（読めなければ NaN）
         */
        function readSum() {
            var total = readProduct();
            while (position < source.length && (source.charAt(position) === "+" || source.charAt(position) === "-")) {
                var operator = source.charAt(position++);
                var operand = readProduct();
                total = (operator === "+") ? total + operand : total - operand;
            }
            return total;
        }

        /**
         * 乗除算の並び（因子 × 因子 …）を読む
         * @returns {number} 値（読めなければ NaN）
         */
        function readProduct() {
            var total = readFactor();
            while (position < source.length && (source.charAt(position) === "*" || source.charAt(position) === "/")) {
                var operator = source.charAt(position++);
                var operand = readFactor();
                if (operator === "/" && operand === 0) return NaN;
                total = (operator === "*") ? total * operand : total / operand;
            }
            return total;
        }

        /**
         * 符号付きの数値（単位付きなら欄の単位へ換算）か、括弧で囲んだ式を読む
         * @returns {number} 値（読めなければ NaN）
         */
        function readFactor() {
            var ch = source.charAt(position);
            if (ch === "+" || ch === "-") {
                position++;
                var signedValue = readFactor();
                return (ch === "-") ? -signedValue : signedValue;
            }
            if (ch === "(") {
                position++;
                var innerValue = readSum();
                if (source.charAt(position) !== ")") return NaN;
                position++;
                return innerValue;
            }
            var numberMatch = /^(\d+\.?\d*|\.\d+)/.exec(source.substring(position));
            if (!numberMatch) return NaN;
            position += numberMatch[0].length;
            return readUnitSuffix(parseFloat(numberMatch[0]));
        }

        /**
         * 数値の直後の単位を読み、欄の単位へ換算する
         * @param {number} value - 単位の前の数値
         * @returns {number} 欄の単位での値（換算できない単位なら NaN）
         */
        function readUnitSuffix(value) {
            var unitMatch = /^(ft\/in|[A-Za-z]+|%|°)/i.exec(source.substring(position));
            if (!unitMatch) return value; /* 単位なしは欄の単位 / no unit means the field's unit */
            position += unitMatch[0].length;
            var unitKey = unitMatch[0].toLowerCase();
            if (unitKey === fieldUnitKey) return value;
            var typedValue = createStepperUnitValue(value, unitKey);
            if (!typedValue || !fieldUnitValue) return NaN; /* 知らない単位・単位のない欄 / unknown unit or unitless field */
            var points = typedValue.as("pt");
            /* 「1p6」＝1パイカ6ポイント / pica-point notation */
            if (unitKey === "p") {
                var pointMatch = /^(\d+\.?\d*|\.\d+)/.exec(source.substring(position));
                if (pointMatch) {
                    position += pointMatch[0].length;
                    points += parseFloat(pointMatch[0]);
                }
            }
            return points / fieldUnitValue.as("pt");
        }

        var result = readSum();
        if (position !== source.length || !isFinite(result)) return NaN; /* 読み残しがあれば式として不正 / leftovers mean a malformed expression */
        return result;
    }

    /**
     * 数値と単位から UnitValue を作る。Q・H・p は STEPPER_UNIT_ALIASES で UnitValue の単位に読み替える。
     * %（percent）は基準の長さが無いと換算できないので扱わない
     * @param {number} value - 数値
     * @param {string} unitKey - 単位（小文字。例 "mm"、"inches"、"q"）
     * @returns {UnitValue|null} UnitValue（UnitValue が知らない単位・空・% なら null）
     */
    function createStepperUnitValue(value, unitKey) {
        if (unitKey === "" || unitKey === "%") return null;
        var alias = STEPPER_UNIT_ALIASES[unitKey];
        var unitValue = alias ? new UnitValue(value * alias.amount, alias.unit) : new UnitValue(value, unitKey);
        if (unitValue.type === "?" || unitValue.type === "%") return null; /* 知らない単位は例外にならず "?" になる。"percent" も除く / unknown units become "?" */
        return unitValue;
    }

    /**
     * 値を整数化・下限・上限でそろえ、単位を付けて入力欄に書き込む（直前の正しい値としても控える）
     * @param {EditText} numberInput - 書き込む入力欄
     * @param {number} value - 数値
     * @param {Object} valueOptions - integer / min / max / unit（どれも省略可）
     * @returns {void}
     */
    function writeSteppedValue(numberInput, value, valueOptions) {
        numberInput.text = formatSteppedValue(value, valueOptions);
        numberInput.lastValidText = numberInput.text;
    }

    /**
     * 値を整数化・下限・上限でそろえ、丸めて単位を付けた表示用の文字列にする。
     * 整数化してから下限で止めるので、「整数・下限1」の欄に 0.4 が入っても 1 になる
     * @param {number} value - 数値
     * @param {Object} valueOptions - integer / min / max / unit（どれも省略可）
     * @returns {string} 入力欄に入れる文字列（例 "20 mm"）
     */
    function formatSteppedValue(value, valueOptions) {
        if (valueOptions.integer) value = Math.round(value);
        return formatStepperNumber(clampSteppedValue(value, valueOptions)) + (valueOptions.unit || "");
    }

    /**
     * 小数第2位で丸めた数値を文字列で返す
     * @param {number} value - 数値
     * @returns {string} 表示用の数値文字列
     */
    function formatStepperNumber(value) {
        return String(Math.round(value * 100) / 100);
    }

    // -----------------------------------------
    // ∧∨ボタンの描画 / Drawing
    // -----------------------------------------
    /**
     * 山形（∧／∨）の極小ボタンを作成する。
     * 上下2つを隙間なく積んで1つの枠に見えるよう、枠線は外側の辺だけ描き（上ボタンは上側、下ボタンは下側）、
     * 継ぎ目に線は引かない
     * @param {Group|Panel} parent - 追加先
     * @param {string} direction - "up" または "down"
     * @param {Function} onClickFn - クリック時の処理
     * @returns {Group} ボタンとして使う group
     */
    function makeStepperChevronButton(parent, direction, onClickFn) {
        var buttonWidth = STEPPER_BUTTON_WIDTH;
        var buttonHeight = STEPPER_BUTTON_HEIGHT;
        var isUp = (direction === "up");
        var chevronBox = parent.add("group");
        chevronBox.margins = 0;
        chevronBox.spacing = 0;
        chevronBox.preferredSize = [buttonWidth, buttonHeight];
        chevronBox.minimumSize = [buttonWidth, buttonHeight];
        chevronBox.maximumSize = [buttonWidth, buttonHeight];
        chevronBox.isPressed = false;
        chevronBox.isStepperButton = true; /* redrawSteppersIn() の目印 / marker for redrawSteppersIn() */

        chevronBox.onDraw = function () {
            var boxGraphics = chevronBox.graphics;
            /* 自作描画は自動でディムにならないため、無効なら薄い色で描く。親の無効化は子の enabled に出ないので親も見る
               Custom drawing is not dimmed automatically; the parent's state does not reach the child's enabled */
            var isDimmed = !isStepperEnabledInTree(chevronBox);

            /* 枠線の内側の地（押下中は押下色） / background inside the frame, pressed color while pressed */
            var fillColor = isDimmed ? STEPPER_DIM_FILL_COLOR : (chevronBox.isPressed ? STEPPER_PRESSED_COLOR : STEPPER_FILL_COLOR);
            boxGraphics.newPath();
            boxGraphics.rectPath(1, isUp ? 1 : 0, buttonWidth - 2, buttonHeight - 1);
            boxGraphics.fillPath(boxGraphics.newBrush(boxGraphics.BrushType.SOLID_COLOR, fillColor));

            drawStepperFrame(boxGraphics, buttonWidth, buttonHeight, isUp, isDimmed ? STEPPER_DIM_FRAME_COLOR : STEPPER_FRAME_COLOR);
            drawStepperChevron(boxGraphics, buttonWidth, buttonHeight, isUp, isDimmed ? STEPPER_DIM_CHEVRON_COLOR : STEPPER_CHEVRON_COLOR);
        };

        /**
         * 押下状態を変えて描き直す
         * @param {boolean} isPressed - 押下中なら true
         * @returns {void}
         */
        function repaint(isPressed) {
            if (chevronBox.isPressed === isPressed) return;
            chevronBox.isPressed = isPressed;
            redrawStepperGroup(chevronBox);
        }
        chevronBox.addEventListener("mousedown", function () {
            if (!isStepperEnabledInTree(chevronBox)) return;
            repaint(true);
            if (onClickFn) onClickFn();
        });
        chevronBox.addEventListener("mouseup", function () { repaint(false); });
        /* 押したまま外へ出たときも押下色を残さない / reset when the pointer leaves while pressed */
        chevronBox.addEventListener("mouseout", function () { repaint(false); });
        return chevronBox;
    }

    /**
     * 外側の辺だけの枠を描く（角は丸める）。継ぎ目側は開けておき、上下2つで1つの枠に見せる。
     * ScriptUI は円弧を描けないため、角丸は短い線分で近似する
     * @param {ScriptUIGraphics} boxGraphics - 描画先
     * @param {number} boxWidth - ボタンの幅
     * @param {number} boxHeight - ボタンの高さ
     * @param {boolean} isUp - 上のボタンなら true（上側に枠を描く）
     * @param {number[]} frameColor - [r, g, b, a]
     * @returns {void}
     */
    function drawStepperFrame(boxGraphics, boxWidth, boxHeight, isUp, frameColor) {
        var frameLeft = 0.5;
        var frameRight = boxWidth - 0.5;
        var outerY = isUp ? 0.5 : boxHeight - 0.5;
        var seamY = isUp ? boxHeight : 0;
        var towardSeam = isUp ? 1 : -1; /* 外側の辺から継ぎ目へ向かう向き / direction from the outer edge to the seam */
        var radius = STEPPER_CORNER_RADIUS;
        var arcSteps = 4; /* 角丸1つを何本の線分で近似するか / segments per corner */
        var angle, k;

        boxGraphics.newPath();
        boxGraphics.moveTo(frameLeft, seamY);
        /* 左の角丸 / left corner */
        for (k = 0; k <= arcSteps; k++) {
            angle = (Math.PI / 2) * k / arcSteps;
            boxGraphics.lineTo(frameLeft + radius - radius * Math.cos(angle), outerY + towardSeam * (radius - radius * Math.sin(angle)));
        }
        /* 右の角丸 / right corner */
        for (k = 0; k <= arcSteps; k++) {
            angle = (Math.PI / 2) * k / arcSteps;
            boxGraphics.lineTo(frameRight - radius + radius * Math.sin(angle), outerY + towardSeam * (radius - radius * Math.cos(angle)));
        }
        boxGraphics.lineTo(frameRight, seamY);
        boxGraphics.strokePath(boxGraphics.newPen(boxGraphics.PenType.SOLID_COLOR, frameColor, 1));
    }

    /**
     * 山形（∧／∨）を描く。文字グリフの▲▼は上下で大きさやベースラインが揃わないため、線で描く
     * @param {ScriptUIGraphics} boxGraphics - 描画先
     * @param {number} boxWidth - ボタンの幅
     * @param {number} boxHeight - ボタンの高さ
     * @param {boolean} isUp - ∧なら true、∨なら false
     * @param {number[]} chevronColor - [r, g, b, a]
     * @returns {void}
     */
    function drawStepperChevron(boxGraphics, boxWidth, boxHeight, isUp, chevronColor) {
        var centerX = boxWidth / 2;
        var centerY = isUp ? boxHeight / 2 + 0.5 : boxHeight / 2 - 0.5; /* 継ぎ目から少し離す / nudged away from the seam */
        var halfWidth = 3.6; /* 山形の半幅（高さ1.8に対して開き約127°） / half width of the chevron */
        var tipOffsetY = isUp ? -1.8 : 1.8; /* 頂点の中心からのずれ（上向きは上、下向きは下） */
        boxGraphics.newPath();
        boxGraphics.moveTo(centerX - halfWidth, centerY - tipOffsetY);
        boxGraphics.lineTo(centerX, centerY + tipOffsetY);
        boxGraphics.lineTo(centerX + halfWidth, centerY - tipOffsetY);
        boxGraphics.strokePath(boxGraphics.newPen(boxGraphics.PenType.SOLID_COLOR, chevronColor, 1.2));
    }

    /**
     * 入力欄にフォーカスを移す
     * @param {EditText} numberInput - 対象の入力欄
     * @returns {void}
     */
    function focusNumberInput(numberInput) {
        numberInput.active = false; /* 一度外さないとフォーカスが移らないことがある / reset first or focus may not move */
        numberInput.active = true;
    }

    /**
     * コントロールと、その親をたどってすべて有効かを返す（親の無効化は子の enabled に出ない）
     * @param {Object} control - 対象のコントロール
     * @returns {boolean} すべて有効なら true
     */
    function isStepperEnabledInTree(control) {
        for (var node = control; node; node = node.parent) {
            if (!node.enabled) return false;
        }
        return true;
    }

    /**
     * コンテナ以下にある∧∨ボタンをすべて描き直す。行やパネルの enabled を切り替えたあとに呼ぶ
     * @param {Object} container - 行・グループ・パネルなど
     * @returns {void}
     */
    function redrawSteppersIn(container) {
        if (!container.children) return;
        for (var i = 0; i < container.children.length; i++) {
            var child = container.children[i];
            if (child.isStepperButton) redrawStepperGroup(child);
            else redrawSteppersIn(child);
        }
    }

    /**
     * group の onDraw を呼び直す。group には notify() が無いため、隠して再表示して描き直させる
     * @param {Group} targetGroup - 描き直す group
     * @returns {void}
     */
    function redrawStepperGroup(targetGroup) {
        targetGroup.hide();
        targetGroup.show();
    }

    // ステップボタン（再利用パーツ）ここまで / End of the reusable stepper

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

    // バージョン表示（再利用パーツ） / Version footer (reusable)

    var VERSION_FOOTER_TOP_MARGIN = 4;    /* バージョンの行の上に足す余白（px。ウィンドウの間隔に足す）/ extra space above the version row, added to the window spacing */
    var VERSION_FOOTER_BOTTOM_MARGIN = 8; /* バージョンの行の下の余白（px。ウィンドウの下余白をこの値にする）/ space below the version row (replaces the window's bottom margin) */
    var VERSION_FOOTER_PIXEL_SIZE = 1;    /* ドット1つの大きさ（px）/ size of one dot (px) */
    var VERSION_FOOTER_LETTER_GAP = 1;    /* 字の間（ドット）/ gap between glyphs (dots) */
    var VERSION_FOOTER_BRAND = "DTP Transit"; /* 右下に出す名前（字形にある文字だけ）/ name shown at the bottom right (glyph characters only) */
    var VERSION_FOOTER_COLOR = isDarkUI() ? [0.45, 0.45, 0.45, 1] : [0.62, 0.62, 0.62, 1]; /* 文字の色（ディム表示に近い灰色）/ text color, a gray close to the dimmed look */

    /* 字形：Sigi 7px Regular（Rasmus Andersson、CC0。https://github.com/rsms/sigi-pixel-font）の大文字・数字・「. - _ +」・空白を
       9 行のドット（# が点。上7行が大文字の高さ、下2行は小文字の下に出る部分）に写したもの。Sigi に小文字が無いので、a〜z は同じ大きさで描き足した。I は l と見分けるよう上下に横棒を足した
       / Glyphs: uppercase, digits, ". - _ +" and space from Sigi 7px Regular by Rasmus Andersson (CC0, https://github.com/rsms/sigi-pixel-font),
         nine rows each ("#" is a dot; seven rows of cap height plus two for descenders). Sigi has no lowercase, so a-z were drawn to match. I has top and bottom bars to tell it from l */
    var VERSION_FOOTER_GLYPH_HEIGHT = 9;  /* 字の高さ（ドット。下2行は g・j・p・q・y の下に出る部分）/ glyph height (dots), including two descender rows */
    var VERSION_FOOTER_GLYPHS = {
        "A": [".###.", "#...#", "#...#", "#####", "#...#", "#...#", "#...#", ".....", "....."],
        "B": ["####.", "#...#", "#...#", "####.", "#...#", "#...#", "####.", ".....", "....."],
        "C": [".###.", "#...#", "#....", "#....", "#....", "#...#", ".###.", ".....", "....."],
        "D": ["####.", "#...#", "#...#", "#...#", "#...#", "#...#", "####.", ".....", "....."],
        "E": ["#####", "#....", "#....", "####.", "#....", "#....", "#####", ".....", "....."],
        "F": ["#####", "#....", "#....", "####.", "#....", "#....", "#....", ".....", "....."],
        "G": [".###.", "#...#", "#....", "#..##", "#...#", "#...#", ".###.", ".....", "....."],
        "H": ["#...#", "#...#", "#...#", "#####", "#...#", "#...#", "#...#", ".....", "....."],
        "I": ["###", ".#.", ".#.", ".#.", ".#.", ".#.", "###", "...", "..."],
        "J": ["....#", "....#", "....#", "....#", "....#", "#...#", ".###.", ".....", "....."],
        "K": ["#...#", "#..#.", "#.#..", "##...", "#.#..", "#..#.", "#...#", ".....", "....."],
        "L": ["#...", "#...", "#...", "#...", "#...", "#...", "####", "....", "...."],
        "M": ["#...#", "##.##", "#.#.#", "#...#", "#...#", "#...#", "#...#", ".....", "....."],
        "N": ["#...#", "#...#", "##..#", "#.#.#", "#..##", "#...#", "#...#", ".....", "....."],
        "O": [".###.", "#...#", "#...#", "#...#", "#...#", "#...#", ".###.", ".....", "....."],
        "P": ["####.", "#...#", "#...#", "####.", "#....", "#....", "#....", ".....", "....."],
        "Q": [".###.", "#...#", "#...#", "#...#", "#...#", "#..##", ".####", ".....", "....."],
        "R": ["####.", "#...#", "#...#", "####.", "#.#..", "#..#.", "#...#", ".....", "....."],
        "S": [".###.", "#...#", "#....", ".###.", "....#", "#...#", ".###.", ".....", "....."],
        "T": ["#####", "..#..", "..#..", "..#..", "..#..", "..#..", "..#..", ".....", "....."],
        "U": ["#...#", "#...#", "#...#", "#...#", "#...#", "#...#", ".###.", ".....", "....."],
        "V": ["#...#", "#...#", "#...#", ".#.#.", ".#.#.", "..#..", "..#..", ".....", "....."],
        "W": ["#..#..#", "#..#..#", "#..#..#", "#..#..#", "#..#..#", "#..#..#", ".##.##.", ".......", "......."],
        "X": ["#...#", "#...#", ".#.#.", "..#..", ".#.#.", "#...#", "#...#", ".....", "....."],
        "Y": ["#...#", "#...#", ".#.#.", "..#..", "..#..", "..#..", "..#..", ".....", "....."],
        "Z": ["#####", "....#", "...#.", "..#..", ".#...", "#....", "#####", ".....", "....."],
        "a": ["....", "....", ".##.", "...#", ".###", "#..#", ".###", "....", "...."],
        "b": ["#...", "#...", "###.", "#..#", "#..#", "#..#", "###.", "....", "...."],
        "c": ["....", "....", ".###", "#...", "#...", "#...", ".###", "....", "...."],
        "d": ["...#", "...#", ".###", "#..#", "#..#", "#..#", ".###", "....", "...."],
        "e": ["....", "....", ".##.", "#..#", "####", "#...", ".###", "....", "...."],
        "f": ["..#", ".#.", "###", ".#.", ".#.", ".#.", ".#.", "...", "..."],
        "g": ["....", "....", ".###", "#..#", "#..#", "#..#", ".###", "...#", "###."],
        "h": ["#...", "#...", "###.", "#..#", "#..#", "#..#", "#..#", "....", "...."],
        "i": ["#", ".", "#", "#", "#", "#", "#", ".", "."],
        "j": [".#", "..", ".#", ".#", ".#", ".#", ".#", ".#", "#."],
        "k": ["#...", "#...", "#..#", "#.#.", "##..", "#.#.", "#..#", "....", "...."],
        "l": ["#.", "#.", "#.", "#.", "#.", "#.", ".#", "..", ".."],
        "m": [".....", ".....", "####.", "#.#.#", "#.#.#", "#.#.#", "#.#.#", ".....", "....."],
        "n": ["....", "....", "###.", "#..#", "#..#", "#..#", "#..#", "....", "...."],
        "o": ["....", "....", ".##.", "#..#", "#..#", "#..#", ".##.", "....", "...."],
        "p": ["....", "....", "###.", "#..#", "#..#", "#..#", "###.", "#...", "#..."],
        "q": ["....", "....", ".###", "#..#", "#..#", "#..#", ".###", "...#", "...#"],
        "r": ["...", "...", "#.#", "##.", "#..", "#..", "#..", "...", "..."],
        "s": ["....", "....", ".###", "#...", ".##.", "...#", "###.", "....", "...."],
        "t": ["...", ".#.", "###", ".#.", ".#.", ".#.", ".##", "...", "..."],
        "u": ["....", "....", "#..#", "#..#", "#..#", "#..#", ".###", "....", "...."],
        "v": [".....", ".....", "#...#", "#...#", ".#.#.", ".#.#.", "..#..", ".....", "....."],
        "w": [".....", ".....", "#...#", "#...#", "#.#.#", "#.#.#", ".#.#.", ".....", "....."],
        "x": ["....", "....", "#..#", "#..#", ".##.", "#..#", "#..#", "....", "...."],
        "y": ["....", "....", "#..#", "#..#", "#..#", "#..#", ".###", "...#", "###."],
        "z": ["....", "....", "####", "...#", ".##.", "#...", "####", "....", "...."],
        "0": [".###.", "#...#", "#...#", "#...#", "#...#", "#...#", ".###.", ".....", "....."],
        "1": [".#.", "##.", ".#.", ".#.", ".#.", ".#.", "###", "...", "..."],
        "2": [".###.", "#...#", "....#", "...#.", "..#..", ".#...", "#####", ".....", "....."],
        "3": [".###.", "#...#", "....#", "..##.", "....#", "#...#", ".###.", ".....", "....."],
        "4": ["...#.", "..##.", ".#.#.", "#..#.", "#####", "...#.", "...#.", ".....", "....."],
        "5": ["#####", "#....", "####.", "....#", "....#", "#...#", ".###.", ".....", "....."],
        "6": [".###.", "#...#", "#....", "####.", "#...#", "#...#", ".###.", ".....", "....."],
        "7": ["#####", "....#", "....#", "...#.", "...#.", "..#..", "..#..", ".....", "....."],
        "8": [".###.", "#...#", "#...#", ".###.", "#...#", "#...#", ".###.", ".....", "....."],
        "9": [".###.", "#...#", "#...#", ".####", "....#", "#...#", ".###.", ".....", "....."],
        ".": [".", ".", ".", ".", ".", ".", "#", ".", "."],
        "-": ["...", "...", "...", "###", "...", "...", "...", "...", "..."],
        "_": ["....", "....", "....", "....", "....", "....", "####", "....", "...."],
        "+": [".....", "..#..", "..#..", "#####", "..#..", "..#..", ".....", ".....", "....."],
        " ": ["...", "...", "...", "...", "...", "...", "...", "...", "..."]
    };

    /**
     * ウィンドウの下端の左にスクリプト名とバージョン（SCRIPT_NAME・SCRIPT_VERSION）、右に VERSION_FOOTER_BRAND を小さく薄く入れる。部品をすべて足したあと、show() の直前に呼ぶ。
     * 直前の行（ボタン行など）の下余白を 0 にし、ウィンドウの下余白を VERSION_FOOTER_BOTTOM_MARGIN にする。
     * InDesign の ScriptUI は文字サイズの指定を無視するので、onDraw でドットを並べて描く（UITheme の isDarkUI() を使う）
     * @param {Window} targetWindow - ダイアログまたはパレット（縦並び）
     * @returns {Group|null} バージョンを描くグループ。すでに足してあれば null
     */
    function addVersionFooter(targetWindow) {
        if (targetWindow.versionFooter) return null;

        /* ボタン行の下余白（BUTTON_ROW_BOTTOM_MARGIN）は、この行が引き受ける / this row takes over the button row's bottom margin */
        var lastChild = targetWindow.children[targetWindow.children.length - 1];
        if (lastChild && lastChild.type === "group") {
            var rowMargins = lastChild.margins;
            lastChild.margins = [rowMargins[0], rowMargins[1], rowMargins[2], 0];
        }
        var windowMargins = targetWindow.margins;
        targetWindow.margins = [windowMargins[0], windowMargins[1], windowMargins[2], VERSION_FOOTER_BOTTOM_MARGIN];

        var footerGroup = targetWindow.add("group");
        footerGroup.orientation = "row";
        footerGroup.alignment = ["fill", "bottom"];
        footerGroup.alignChildren = ["left", "center"];
        footerGroup.margins = [0, VERSION_FOOTER_TOP_MARGIN, 0, 0];

        /* 左は「スクリプト名 バージョン」。字形に無い文字（日本語など）を含む名前や、幅が足りないときはバージョンだけ。
           左の欄は余りの幅まで伸ばし、ダイアログの幅はバージョンの分しか求めない（名前でダイアログを広げない）
           / Left: "script name version", or just the version when the name has characters without glyphs (e.g. Japanese)
             or does not fit. The left area takes the extra width and only asks for the version's width, so the name never widens the dialog */
        var versionText = String(SCRIPT_VERSION);
        var leftCandidates = [versionText];
        if (typeof SCRIPT_NAME === "string" && hasVersionFooterGlyphs(SCRIPT_NAME)) leftCandidates.unshift(SCRIPT_NAME + " " + versionText);
        var versionCanvas = addVersionFooterCanvas(footerGroup, leftCandidates);
        versionCanvas.alignment = ["fill", "center"];
        var brandCanvas = addVersionFooterCanvas(footerGroup, [VERSION_FOOTER_BRAND]);
        brandCanvas.alignment = ["right", "center"];
        brandCanvas.maximumSize = brandCanvas.preferredSize;
        targetWindow.versionFooter = footerGroup;
        return versionCanvas;
    }

    /**
     * 文字列のすべての文字に字形があるかを返す
     * @param {string} dotText - 文字列
     * @returns {boolean} すべてあれば true
     */
    function hasVersionFooterGlyphs(dotText) {
        for (var i = 0; i < dotText.length; i++) {
            if (!VERSION_FOOTER_GLYPHS.hasOwnProperty(dotText.charAt(i))) return false;
        }
        return dotText.length > 0;
    }

    /**
     * ドットで文字列を描くグループを足す。候補のうち、欄の幅に収まる最初のものを描く
     * @param {Group} parentGroup - 足す先
     * @param {string[]} dotTexts - 描く文字列の候補（長い順。いちばん短いものの幅を欄の幅として求める）
     * @returns {Group} 描くグループ
     */
    function addVersionFooterCanvas(parentGroup, dotTexts) {
        var dotCanvas = parentGroup.add("group");
        var shortestText = dotTexts[dotTexts.length - 1];
        dotCanvas.preferredSize = [measureVersionFooterDots(shortestText) * VERSION_FOOTER_PIXEL_SIZE, VERSION_FOOTER_GLYPH_HEIGHT * VERSION_FOOTER_PIXEL_SIZE];
        dotCanvas.helpTip = dotTexts[0];
        dotCanvas.onDraw = function () {
            var availableWidth = this.size ? this.size.width : 0;
            var drawText = shortestText;
            for (var i = 0; i < dotTexts.length; i++) {
                if (measureVersionFooterDots(dotTexts[i]) * VERSION_FOOTER_PIXEL_SIZE <= availableWidth) {
                    drawText = dotTexts[i];
                    break;
                }
            }
            drawVersionFooterDots(this.graphics, drawText);
        };
        return dotCanvas;
    }

    /**
     * 文字列の幅をドットの数で返す（字形に無い文字は数えない）
     * @param {string} versionText - 文字列
     * @returns {number} 幅（ドット）
     */
    function measureVersionFooterDots(versionText) {
        var dotCount = 0;
        for (var i = 0; i < versionText.length; i++) {
            var glyphRows = VERSION_FOOTER_GLYPHS[versionText.charAt(i)];
            if (!glyphRows) continue;
            if (dotCount > 0) dotCount += VERSION_FOOTER_LETTER_GAP;
            dotCount += glyphRows[0].length;
        }
        return dotCount;
    }

    /**
     * 文字列をドットで描く。横に続くドットは1つの四角にまとめて塗る
     * @param {ScriptUIGraphics} canvasGraphics - 描く先の graphics
     * @param {string} versionText - 文字列
     * @returns {void}
     */
    function drawVersionFooterDots(canvasGraphics, versionText) {
        var footerBrush = canvasGraphics.newBrush(canvasGraphics.BrushType.SOLID_COLOR, VERSION_FOOTER_COLOR);
        var dotSize = VERSION_FOOTER_PIXEL_SIZE;
        var cursorDots = 0;
        canvasGraphics.newPath();
        for (var i = 0; i < versionText.length; i++) {
            var glyphRows = VERSION_FOOTER_GLYPHS[versionText.charAt(i)];
            if (!glyphRows) continue;
            if (cursorDots > 0) cursorDots += VERSION_FOOTER_LETTER_GAP;
            for (var row = 0; row < glyphRows.length; row++) {
                var rowText = glyphRows[row];
                var runStart = -1;
                for (var col = 0; col <= rowText.length; col++) {
                    var isDot = col < rowText.length && rowText.charAt(col) === "#";
                    if (isDot && runStart < 0) runStart = col;
                    if (!isDot && runStart >= 0) {
                        canvasGraphics.rectPath((cursorDots + runStart) * dotSize, row * dotSize, (col - runStart) * dotSize, dotSize);
                        runStart = -1;
                    }
                }
            }
            cursorDots += glyphRows[0].length;
        }
        canvasGraphics.fillPath(footerBrush);
    }

    // バージョン表示（再利用パーツ）ここまで / End of the reusable version footer

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

    // 選択の収集と境界（再利用パーツ） / Selection items and bounds (reusable)

    /* 座標を同じと見なす許容値（pt） / Tolerance for treating coordinates as equal, in points */
    var SELECTION_ITEMS_TOLERANCE = 0.001;

    /**
     * 選択やコレクションを、オブジェクトの配列にそろえる
     * TextRange・PathItem は length を持つので、typename で1個か集まりかを見分ける
     * @param {*} source - doc.selection、配列、DOM のコレクション、または単独のオブジェクト
     * @returns {Array} オブジェクトの配列（空なら []）
     */
    function normalizeSelectionItems(source) {
        var items = [];
        if (!source) return items;
        var typeName = "";
        try { typeName = source.typename || ""; } catch (e) { /* 読めない種類 / unreadable kind */ }
        /* 単数形の typename は1個（PageItems などのコレクションは s で終わる）
           A singular typename is one object (collections such as PageItems end in s) */
        if (typeName && !/s$/.test(typeName)) return [source];
        if (typeof source.length !== "number") return items;
        for (var i = 0; i < source.length; i++) items.push(source[i]);
        return items;
    }

    /**
     * 文字カーソルの選択（TextRange）を、それを含むテキストフレームに読み替える
     * @param {TextRange} textRange - 文字の範囲
     * @returns {TextFrame|null} テキストフレーム（たどれなければ null）
     */
    function resolveTextRangeFrame(textRange) {
        var current = textRange;
        /* parent をたどる（深さは念のため制限） / Walk up the parents, with a safety limit */
        for (var depth = 0; depth < 10 && current; depth++) {
            try {
                if (current.typename === "TextFrame") return current;
                current = current.parent;
            } catch (e) {
                break;
            }
        }
        /* ストーリーの先頭フレームで代用する / Fall back to the first frame of the story */
        try {
            var storyFrames = textRange.story.textFrames;
            if (storyFrames.length > 0) return storyFrames[0];
        } catch (e2) { /* ストーリーを持たない / no story */ }
        return null;
    }

    /**
     * 選択から条件に合うオブジェクトを集める（グループ・レイヤーを再帰でたどり、重複は除く）
     * 条件に合ったオブジェクトの中へは進まない
     * @param {*} source - doc.selection、配列、コレクション、または単独のオブジェクト
     * @param {Object} [options] - 収集の設定
     * @param {function(PageItem): boolean} [options.accept] - 集める条件（既定はグループ・レイヤー以外すべて）
     * @param {boolean} [options.enterGroups] - グループの中をたどる（既定 true）
     * @param {boolean} [options.enterClipGroups] - クリップグループの中をたどる（既定は enterGroups と同じ）
     * @param {boolean} [options.enterCompoundPaths] - 複合パスの中のパスをたどる（既定 false）
     * @param {boolean} [options.textRangeToFrame] - 文字の選択をテキストフレームに読み替える（既定 true）
     * @param {boolean} [options.skipLocked] - ロックされたものを中ごと外す（既定 false）
     * @param {boolean} [options.skipHidden] - 非表示のものを中ごと外す（既定 false）
     * @param {boolean} [options.skipClipMasks] - クリッピングマスクを外す（既定 false）
     * @param {boolean} [options.skipGuides] - ガイドを外す（既定 false）
     * @param {boolean} [options.unique] - 同じ参照を1回だけにする（既定 true。数千件で遅ければ false）
     * @returns {Array} 集めたオブジェクト（前面→背面の順）
     */
    function collectSelectionItems(source, options) {
        var opts = options || {};
        var enterGroups = (opts.enterGroups !== false);
        var enterClipGroups = (opts.enterClipGroups === undefined) ? enterGroups : (opts.enterClipGroups === true);
        var accept = opts.accept || function (item) {
            return item.typename !== "GroupItem" && item.typename !== "Layer";
        };
        var collected = [];

        /**
         * 集めた配列に加える（unique のときは同じ参照を足さない）
         * @param {PageItem} item - 加えるオブジェクト
         * @returns {void}
         */
        function pushItem(item) {
            if (opts.unique !== false) {
                for (var k = 0; k < collected.length; k++) {
                    if (collected[k] === item) return;
                }
            }
            collected.push(item);
        }

        /**
         * 設定に従って外すオブジェクトか判定する
         * @param {PageItem} item - 判定するオブジェクト
         * @returns {boolean} 外すなら true
         */
        function isSkipped(item) {
            try {
                if (item.typename === "Layer") {
                    if (opts.skipLocked && item.locked) return true;
                    if (opts.skipHidden && !item.visible) return true;
                    return false;
                }
                if (opts.skipLocked && item.locked) return true;
                if (opts.skipHidden && item.hidden) return true;
                if (opts.skipGuides && item.guides === true) return true;
                if (opts.skipClipMasks && isClipMaskItem(item)) return true;
            } catch (e) {
                /* 読めないプロパティは「外さない」に倒す / Unreadable properties do not exclude */
            }
            return false;
        }

        /**
         * 1件をたどって集める
         * @param {PageItem} item - 対象のオブジェクト
         * @returns {void}
         */
        function visit(item) {
            if (!item) return;
            var typeName = "";
            try { typeName = item.typename; } catch (e) { return; }

            if (typeName === "TextRange" || typeName === "InsertionPoint") {
                if (opts.textRangeToFrame === false) {
                    if (accept(item)) pushItem(item);
                    return;
                }
                visit(resolveTextRangeFrame(item));
                return;
            }
            if (isSkipped(item)) return;
            if (accept(item)) {
                pushItem(item);
                return;
            }

            var children = null;
            if (typeName === "GroupItem") {
                var isClipped = false;
                try { isClipped = (item.clipped === true); } catch (e2) { }
                if (isClipped ? enterClipGroups : enterGroups) children = item.pageItems;
            } else if (typeName === "CompoundPathItem") {
                if (opts.enterCompoundPaths) children = item.pathItems;
            } else if (typeName === "Layer") {
                /* 重なり順はサブレイヤーとページアイテムで別々なので、ページアイテム→サブレイヤーの順にする
                   Page items and sublayers stack separately; visit page items first, then sublayers */
                walk(item.pageItems);
                walk(item.layers);
                return;
            }
            if (children) walk(children);
        }

        /**
         * 集まりの各要素をたどる
         * @param {*} list - 配列またはコレクション
         * @returns {void}
         */
        function walk(list) {
            var listItems = normalizeSelectionItems(list);
            for (var i = 0; i < listItems.length; i++) visit(listItems[i]);
        }

        walk(source);
        return collected;
    }

    /**
     * テキストフレームの種類を "point" / "area" / "path" で返す
     * @param {TextFrame} textFrame - テキストフレーム
     * @returns {string} 種類のキー（判定できなければ ""）
     */
    function getTextFrameKindKey(textFrame) {
        try {
            if (textFrame.kind === TextType.POINTTEXT) return "point";
            if (textFrame.kind === TextType.AREATEXT) return "area";
            if (textFrame.kind === TextType.PATHTEXT) return "path";
        } catch (e) { /* kind を読めない / kind is unreadable */ }
        return "";
    }

    /**
     * 選択からテキストフレームを集める（グループの中・文字カーソルの選択を含む）
     * @param {*} source - doc.selection など
     * @param {Object} [options] - collectSelectionItems と同じ設定に加えて次を受ける
     * @param {string[]} [options.kinds] - 集める種類（"point" / "area" / "path"。既定はすべて）
     * @returns {TextFrame[]} テキストフレーム（前面→背面の順）
     */
    function collectSelectionTextFrames(source, options) {
        var opts = {};
        var sourceOptions = options || {};
        for (var key in sourceOptions) {
            if (sourceOptions.hasOwnProperty(key)) opts[key] = sourceOptions[key];
        }
        var kindFilter = null;
        if (opts.kinds && opts.kinds.length) {
            kindFilter = {};
            for (var i = 0; i < opts.kinds.length; i++) kindFilter[opts.kinds[i]] = true;
        }
        opts.accept = function (item) {
            if (item.typename !== "TextFrame") return false;
            return !kindFilter || kindFilter[getTextFrameKindKey(item)] === true;
        };
        /* 種類で外したテキストは中をたどらない（accept が false でも子は無い） / Text frames have no children to walk */
        return collectSelectionItems(source, opts);
    }

    /**
     * 選択からパスを集める（グループの中を含む）
     * @param {*} source - doc.selection など
     * @param {Object} [options] - collectSelectionItems と同じ設定に加えて次を受ける
     * @param {string} [options.compoundPaths] - 複合パスの扱い。"children"（中のパス、既定）/ "whole"（複合パスごと）/ "skip"（外す）
     * @returns {Array} PathItem（"whole" のときは CompoundPathItem も）の配列
     */
    function collectSelectionPathItems(source, options) {
        var opts = {};
        var sourceOptions = options || {};
        for (var key in sourceOptions) {
            if (sourceOptions.hasOwnProperty(key)) opts[key] = sourceOptions[key];
        }
        var compoundMode = opts.compoundPaths || "children";
        opts.enterCompoundPaths = (compoundMode === "children");
        opts.accept = function (item) {
            if (item.typename === "PathItem") return true;
            return compoundMode === "whole" && item.typename === "CompoundPathItem";
        };
        return collectSelectionItems(source, opts);
    }

    /**
     * クリッピングマスク（クリップグループの型）か判定する
     * パスは clipping、複合パスは中の先頭パスの clipping、テキストは clipping が無いので「クリップグループの先頭」で見る
     * @param {PageItem} item - 判定するオブジェクト
     * @returns {boolean} マスクなら true
     */
    function isClipMaskItem(item) {
        try {
            if (item.typename === "PathItem") return item.clipping === true;
            if (item.typename === "CompoundPathItem") {
                return item.pathItems.length > 0 && item.pathItems[0].clipping === true;
            }
            if (item.typename === "TextFrame") {
                var parentGroup = item.parent;
                return parentGroup.typename === "GroupItem" && parentGroup.clipped === true &&
                    parentGroup.pageItems.length > 0 && parentGroup.pageItems[0] === item;
            }
        } catch (e) { /* 読めない種類はマスクではない / unreadable kinds are not masks */ }
        return false;
    }

    /**
     * クリップグループの型（マスク）を返す
     * フラグで探し、見つからなければ先頭（pageItems[0]）を返す（型は常に最前面。テキストの型はフラグを持たない）
     * @param {GroupItem} groupItem - 対象のグループ
     * @returns {PageItem|null} マスク（クリップグループでなければ null）
     */
    function getClipMaskItem(groupItem) {
        try {
            if (!groupItem || groupItem.typename !== "GroupItem" || groupItem.clipped !== true) return null;
            var groupChildren = groupItem.pageItems;
            if (groupChildren.length === 0) return null;
            for (var i = 0; i < groupChildren.length; i++) {
                var childType = groupChildren[i].typename;
                if ((childType === "PathItem" || childType === "CompoundPathItem") && isClipMaskItem(groupChildren[i])) {
                    return groupChildren[i];
                }
            }
            return groupChildren[0];
        } catch (e) {
            return null;
        }
    }

    /**
     * グループの中（入れ子を含む）にクリップグループがあるか判定する
     * @param {GroupItem} groupItem - 対象のグループ
     * @returns {boolean} あれば true
     */
    function hasClippedDescendant(groupItem) {
        try {
            var groupChildren = groupItem.pageItems;
            for (var i = 0; i < groupChildren.length; i++) {
                if (groupChildren[i].typename !== "GroupItem") continue;
                if (groupChildren[i].clipped === true || hasClippedDescendant(groupChildren[i])) return true;
            }
        } catch (e) { /* 中を読めない / cannot read the children */ }
        return false;
    }

    /**
     * 環境設定の［プレビュー境界を使用］を読む
     * @returns {boolean} オンなら true（読めなければ false）
     */
    function readUsePreviewBoundsPreference() {
        try {
            return app.preferences.getBooleanPreference("includeStrokeInBounds");
        } catch (e) {
            return false;
        }
    }

    /**
     * 見た目どおりの境界を返す。クリップグループはマスクの境界、
     * 中にクリップグループを含むグループは子の境界を合わせたもの（隠れた部分を含めない）
     * @param {PageItem} item - 対象のオブジェクト
     * @param {boolean} [usePreviewBounds] - true で visibleBounds、false で geometricBounds（省略時は環境設定に従う）
     * @returns {number[]|null} [左, 上, 右, 下] の新しい配列（測れなければ null）
     */
    function getClipAwareBounds(item, usePreviewBounds) {
        var usePreview = (usePreviewBounds === undefined || usePreviewBounds === null) ?
            readUsePreviewBoundsPreference() : (usePreviewBounds === true);
        try {
            var measuredItem = item;
            if (item.typename === "GroupItem") {
                var maskItem = getClipMaskItem(item);
                if (maskItem) {
                    measuredItem = maskItem;
                } else if (hasClippedDescendant(item)) {
                    /* グループ自体の効果（影など）の広がりは含まれなくなる
                       This leaves out the reach of effects applied to the group itself (drop shadows etc.) */
                    var childBounds = getClipAwareUnionBounds(filterMeasurableChildren(item.pageItems), usePreview);
                    if (childBounds) return childBounds;
                }
            }
            var bounds = usePreview ? measuredItem.visibleBounds : measuredItem.geometricBounds;
            return [bounds[0], bounds[1], bounds[2], bounds[3]];
        } catch (e) {
            return null;
        }
    }

    /**
     * 境界の計算に入れる子だけを残す（非表示とガイドを外す）
     * @param {*} childList - 子のコレクション
     * @returns {Array} 残した子
     */
    function filterMeasurableChildren(childList) {
        var childItems = normalizeSelectionItems(childList);
        var measurable = [];
        for (var i = 0; i < childItems.length; i++) {
            try {
                if (childItems[i].hidden === true || childItems[i].guides === true) continue;
            } catch (e) { /* 読めなければ残す / keep when unreadable */ }
            measurable.push(childItems[i]);
        }
        return measurable;
    }

    /**
     * 複数のオブジェクトを囲む外接範囲を返す（クリップグループはマスクで測る）
     * @param {*} items - オブジェクトの配列・コレクション・選択
     * @param {boolean} [usePreviewBounds] - true で visibleBounds、false で geometricBounds（省略時は環境設定に従う）
     * @returns {number[]|null} [左, 上, 右, 下]（測れるものが無ければ null）
     */
    function getClipAwareUnionBounds(items, usePreviewBounds) {
        var usePreview = (usePreviewBounds === undefined || usePreviewBounds === null) ?
            readUsePreviewBoundsPreference() : (usePreviewBounds === true);
        var itemList = normalizeSelectionItems(items);
        var unionBounds = null;
        for (var i = 0; i < itemList.length; i++) {
            var itemBounds = getClipAwareBounds(itemList[i], usePreview);
            if (!itemBounds) continue;
            if (!unionBounds) {
                unionBounds = itemBounds;
                continue;
            }
            if (itemBounds[0] < unionBounds[0]) unionBounds[0] = itemBounds[0];
            if (itemBounds[1] > unionBounds[1]) unionBounds[1] = itemBounds[1];
            if (itemBounds[2] > unionBounds[2]) unionBounds[2] = itemBounds[2];
            if (itemBounds[3] < unionBounds[3]) unionBounds[3] = itemBounds[3];
        }
        return unionBounds;
    }

    /**
     * 2つの座標を許容値つきで比べる
     * @param {number} valueA - 座標A（pt）
     * @param {number} valueB - 座標B（pt）
     * @param {number} [tolerance] - 許容値（pt、既定は SELECTION_ITEMS_TOLERANCE）
     * @returns {boolean} 差が許容値以下なら true
     */
    function isNearlySameCoordinate(valueA, valueB, tolerance) {
        var limit = (typeof tolerance === "number") ? tolerance : SELECTION_ITEMS_TOLERANCE;
        return Math.abs(valueA - valueB) <= limit;
    }

    /**
     * 2つの境界を許容値つきで比べる
     * @param {number[]} boundsA - [左, 上, 右, 下]
     * @param {number[]} boundsB - [左, 上, 右, 下]
     * @param {number} [tolerance] - 許容値（pt、既定は SELECTION_ITEMS_TOLERANCE）
     * @returns {boolean} 4辺とも許容値以内なら true
     */
    function areBoundsNearlyEqual(boundsA, boundsB, tolerance) {
        if (!boundsA || !boundsB) return false;
        for (var i = 0; i < 4; i++) {
            if (!isNearlySameCoordinate(boundsA[i], boundsB[i], tolerance)) return false;
        }
        return true;
    }

    // 選択の収集と境界（再利用パーツ）ここまで / End of the reusable selection items and bounds

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

    var LABELS = {
        dialog: {
            title: { ja: "列に詰めて並べる", en: "Pack into Columns" }
        },
        panel: {
            reference: { ja: "基準", en: "Reference" },
            columns: { ja: "列", en: "Columns" },
            gap: { ja: "間隔", en: "Gap" },
            arrangement: { ja: "並べ方", en: "Arrangement" }
        },
        radio: {
            selection: { ja: "選択オブジェクト", en: "Selected Objects" },
            artboard: { ja: "現在のアートボード", en: "Current Artboard" },
            left: { ja: "左から", en: "From Left" },
            centerX: { ja: "中央から", en: "From Center" },
            right: { ja: "右から", en: "From Right" },
            top: { ja: "上から", en: "From Top" },
            centerY: { ja: "中央から", en: "From Center" },
            bottom: { ja: "下から", en: "From Bottom" },
            position: { ja: "位置", en: "Position" },
            stacking: { ja: "重ね順", en: "Stacking" },
            random: { ja: "ランダム", en: "Random" }
        },
        checkbox: {
            fitArtboard: { ja: "アートボードにフィット", en: "Fit to Artboard" },
            matchWidth: { ja: "幅を列にそろえる", en: "Match Width to Column" }
        },
        fieldLabel: {
            margin: { ja: "余白", en: "Margin" },
            columnCount: { ja: "列数", en: "Columns" },
            columnWidth: { ja: "列の幅", en: "Column Width" },
            columnGap: { ja: "左右", en: "Horizontal" },
            rowGap: { ja: "上下", en: "Vertical" },
            horizontalStart: { ja: "横", en: "Horizontal" },
            verticalStart: { ja: "縦", en: "Vertical" },
            order: { ja: "順序", en: "Order" }
        },
        tooltip: {
            columnCount: { ja: "列の数", en: "Number of columns" },
            columnWidth: {
                ja: "1列の幅。これより幅の広いオブジェクトは複数の列にまたがる",
                en: "Width of one column; wider objects span several columns"
            },
            columnGap: { ja: "列と列のあいだの間隔", en: "Gap between columns" },
            rowGap: { ja: "列の中のオブジェクトどうしの間隔", en: "Gap between objects in a column" },
            selection: { ja: "選択オブジェクト全体の範囲を基準に並べる", en: "Use the area of all the selected objects" },
            artboard: {
                ja: "現在のアートボードに収まるように列数を決める。収まらなければ縮小する",
                en: "Choose the column count to fit the current artboard, scaling down if needed"
            },
            fitArtboard: {
                ja: "拡大もして、アートボードいっぱいに並べる",
                en: "Scale up as well, to fill the artboard"
            },
            margin: { ja: "アートボードの端から内側に取る余白", en: "Space kept inside the artboard edges" },
            matchWidth: {
                ja: "縦横比を保ったまま、すべてのオブジェクトの幅を列の幅にする",
                en: "Scale every object to the column width, keeping its proportions"
            },
            linkGaps: { ja: "左右と上下の間隔を連動する", en: "Link the horizontal and vertical gaps" },
            left: { ja: "基準の左端から右へ列を並べる", en: "Lay out columns rightward from the left edge of the reference" },
            centerX: { ja: "全体をアートボードの左右中央に置く", en: "Center the whole layout horizontally on the artboard" },
            right: { ja: "基準の右端から左へ列を並べる", en: "Lay out columns leftward from the right edge of the reference" },
            top: { ja: "基準の上端から下へ積む", en: "Stack downward from the top edge of the reference" },
            centerY: { ja: "全体をアートボードの上下中央に置く", en: "Center the whole layout vertically on the artboard" },
            bottom: { ja: "基準の下端から上へ積む", en: "Stack upward from the bottom edge of the reference" },
            position: { ja: "今の位置の上の行から順に、行の中は左から置く", en: "Place by current position: rows from the top, left to right in each row" },
            stacking: { ja: "前面のオブジェクトから順に置く", en: "Place from the frontmost object backward" },
            random: { ja: "ランダムな順に置く", en: "Place in random order" },
            shuffle: { ja: "ランダムな順を作り直す", en: "Make a new random order" },
            stepUp: {
                ja: "値を増やす（shift＋クリックで10の倍数へ、option＋クリックで0.1ずつ）",
                en: "Increase (Shift-click to snap to 10s, Option-click by 0.1)"
            },
            stepDown: {
                ja: "値を減らす（shift＋クリックで10の倍数へ、option＋クリックで0.1ずつ）",
                en: "Decrease (Shift-click to snap to 10s, Option-click by 0.1)"
            },
            stepUpInteger: { ja: "値を増やす（shift＋クリックで10の倍数へ）", en: "Increase (Shift-click to snap to 10s)" },
            stepDownInteger: { ja: "値を減らす（shift＋クリックで10の倍数へ）", en: "Decrease (Shift-click to snap to 10s)" }
        },
        button: {
            shuffle: { ja: "シャッフル", en: "Shuffle" },
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" }
        },
        alert: {
            noDocument: { ja: "ドキュメントを開いてください。", en: "Open a document first." },
            needItems: { ja: "オブジェクトを{count}つ以上選択してください。", en: "Select at least {count} objects." }
        }
    };

    // =========================================
    // メイソンリーの計算 / Masonry calculation
    // =========================================
    /* Masonry v3.3.2 の measureColumns / _getItemLayoutPosition / _getColGroup の移植。
       座標はどれも左上を原点に、右と下を正にした pt。Illustrator の座標への読み替えは computeLayoutPlan() で行う
       Ported from Masonry v3.3.2. Coordinates here are points from the top left, growing right and down;
       computeLayoutPlan() maps them onto Illustrator's coordinates */

    /**
     * 配列の最小値を返す
     * @param {number[]} values - 数値の配列
     * @returns {number} 最小値
     */
    function getMinValue(values) {
        var minValue = values[0];
        for (var i = 1; i < values.length; i++) {
            if (values[i] < minValue) minValue = values[i];
        }
        return minValue;
    }

    /**
     * 配列の最大値を返す
     * @param {number[]} values - 数値の配列
     * @returns {number} 最大値
     */
    function getMaxValue(values) {
        var maxValue = values[0];
        for (var i = 1; i < values.length; i++) {
            if (values[i] > maxValue) maxValue = values[i];
        }
        return maxValue;
    }

    /**
     * 指定の値が最初に現れる位置を返す（ES3 の配列には indexOf が無い）
     * @param {number[]} values - 数値の配列
     * @param {number} targetValue - 探す値
     * @returns {number} 位置（無ければ -1）
     */
    function findValueIndex(values, targetValue) {
        for (var i = 0; i < values.length; i++) {
            if (values[i] === targetValue) return i;
        }
        return -1;
    }

    /**
     * 全体の幅に入る列数を求める（measureColumns）。端数が1pt未満なら誤差とみなして四捨五入する
     * @param {number} containerWidth - 全体の幅（pt）
     * @param {number} columnWidth - 列の幅（pt）
     * @param {number} columnGap - 列の間隔（pt）
     * @returns {number} 列数（1以上）
     */
    function measureColumnCount(containerWidth, columnWidth, columnGap) {
        var columnPitch = columnWidth + columnGap;
        var paddedWidth = containerWidth + columnGap;
        var shortfall = columnPitch - paddedWidth % columnPitch;
        var columnRatio = paddedWidth / columnPitch;
        var columnCount = (shortfall && shortfall < 1) ? Math.round(columnRatio) : Math.floor(columnRatio);
        return Math.max(columnCount, 1);
    }

    /**
     * 基準の幅に入る列数を、オブジェクトの数と列数の上限の範囲で求める
     * @param {number} referenceWidth - 基準の幅（pt）
     * @param {number} columnWidth - 列の幅（pt）
     * @param {number} columnGap - 列の間隔（pt）
     * @param {number} itemCount - オブジェクトの数
     * @returns {number} 列数
     */
    function fitColumnCount(referenceWidth, columnWidth, columnGap, itemCount) {
        return Math.min(measureColumnCount(referenceWidth, columnWidth, columnGap), itemCount, MAX_COLUMN_COUNT);
    }

    /**
     * オブジェクトがまたがる列の数を求める。余りが1pt未満なら誤差とみなして四捨五入する
     * @param {number} itemWidth - オブジェクトの幅（pt）
     * @param {number} columnPitch - 列の幅＋列の間隔（pt）
     * @param {number} columnCount - 列数
     * @returns {number} またがる列の数（列数が上限）
     */
    function getColumnSpan(itemWidth, columnPitch, columnCount) {
        var remainder = itemWidth % columnPitch;
        var spanRatio = itemWidth / columnPitch;
        var columnSpan = (remainder && remainder < 1) ? Math.round(spanRatio) : Math.ceil(spanRatio);
        return Math.min(columnSpan, columnCount);
    }

    /**
     * 複数の列にまたがるオブジェクトを置ける位置ごとに、またがる列のうち最も低い地点を返す（_getColGroup）
     * @param {number[]} columnBottoms - 列ごとの積み上がった高さ
     * @param {number} columnSpan - またがる列の数
     * @returns {number[]} 置ける位置ごとの高さ
     */
    function getColumnGroupBottoms(columnBottoms, columnSpan) {
        if (columnSpan < 2) return columnBottoms;
        var groupBottoms = [];
        var groupCount = columnBottoms.length + 1 - columnSpan;
        for (var i = 0; i < groupCount; i++) {
            groupBottoms.push(getMaxValue(columnBottoms.slice(i, i + columnSpan)));
        }
        return groupBottoms;
    }

    /**
     * 順に、いちばん短い列へオブジェクトを置いていく（_getItemLayoutPosition）
     * @param {Array<{width: number, height: number}>} itemSizes - 置く順に並べた大きさ（pt）
     * @param {{columnCount: number, columnWidth: number, columnGap: number, rowGap: number}} layoutSettings - 列と間隔（pt）
     * @returns {Array<{x: number, y: number}>} itemSizes と同じ順の位置
     */
    function computeMasonryPositions(itemSizes, layoutSettings) {
        var columnPitch = layoutSettings.columnWidth + layoutSettings.columnGap;
        var columnBottoms = [];
        for (var column = 0; column < layoutSettings.columnCount; column++) {
            columnBottoms.push(0);
        }

        var positions = [];
        for (var i = 0; i < itemSizes.length; i++) {
            var columnSpan = getColumnSpan(itemSizes[i].width, columnPitch, layoutSettings.columnCount);
            var groupBottoms = getColumnGroupBottoms(columnBottoms, columnSpan);
            var shortestBottom = getMinValue(groupBottoms);
            var shortestIndex = findValueIndex(groupBottoms, shortestBottom);
            positions.push({ x: columnPitch * shortestIndex, y: shortestBottom });

            var nextBottom = shortestBottom + (itemSizes[i].height + layoutSettings.rowGap);
            var coveredColumns = layoutSettings.columnCount + 1 - groupBottoms.length;
            for (var j = 0; j < coveredColumns; j++) {
                columnBottoms[shortestIndex + j] = nextBottom;
            }
        }
        return positions;
    }

    // =========================================
    // 基準への収まり / Fitting to the reference
    // =========================================

    var MIN_FIT_SCALE = 0.01;  /* 縮小の下限（1 が原寸） / lowest scale (1 is actual size) */
    var FIT_SEARCH_STEPS = 16; /* 倍率を二分探索する上限の回数 / most bisection steps for the scale */
    var FIT_SCALE_PRECISION = 0.0005; /* 倍率の探索を止める幅 / the scale search stops at this range */
    var SCALE_TOLERANCE = 0.00001;    /* 原寸と見なす倍率の差 / scale differences treated as actual size */

    /**
     * 倍率を掛けた大きさで配置を求める。間隔は倍率を掛けない
     * @param {Array<{width: number, height: number}>} itemSizes - 置く順に並べた原寸の大きさ（pt）
     * @param {{columnWidth: number, columnGap: number, rowGap: number}} layoutSettings - 列と間隔（pt）
     * @param {number} columnCount - 列数
     * @param {number} scale - 倍率（1 が原寸）
     * @returns {{columnCount: number, scale: number, sizes: Array, positions: Array}} 配置
     */
    function buildScaledLayout(itemSizes, layoutSettings, columnCount, scale) {
        var scaledSizes = [];
        for (var i = 0; i < itemSizes.length; i++) {
            scaledSizes.push({ width: itemSizes[i].width * scale, height: itemSizes[i].height * scale });
        }
        var positions = computeMasonryPositions(scaledSizes, {
            columnCount: columnCount,
            columnWidth: layoutSettings.columnWidth * scale,
            columnGap: layoutSettings.columnGap,
            rowGap: layoutSettings.rowGap
        });
        return { columnCount: columnCount, scale: scale, sizes: scaledSizes, positions: positions };
    }

    /**
     * 配置全体の大きさを測る
     * @param {{sizes: Array, positions: Array}} scaledLayout - buildScaledLayout() の戻り値
     * @returns {{width: number, height: number}} 全体の幅と高さ（pt）
     */
    function measureLayoutExtent(scaledLayout) {
        var layoutWidth = 0;
        var layoutHeight = 0;
        for (var i = 0; i < scaledLayout.sizes.length; i++) {
            layoutWidth = Math.max(layoutWidth, scaledLayout.positions[i].x + scaledLayout.sizes[i].width);
            layoutHeight = Math.max(layoutHeight, scaledLayout.positions[i].y + scaledLayout.sizes[i].height);
        }
        return { width: layoutWidth, height: layoutHeight };
    }

    /**
     * 配置が基準の大きさに収まるか調べる
     * @param {{sizes: Array, positions: Array}} scaledLayout - buildScaledLayout() の戻り値
     * @param {{width: number, height: number}} referenceSize - 基準の大きさ（pt）
     * @returns {boolean} 収まれば true
     */
    function doesLayoutFit(scaledLayout, referenceSize) {
        var layoutExtent = measureLayoutExtent(scaledLayout);
        return layoutExtent.width <= referenceSize.width + SELECTION_ITEMS_TOLERANCE &&
            layoutExtent.height <= referenceSize.height + SELECTION_ITEMS_TOLERANCE;
    }

    /**
     * 指定の列数で基準に収まる最大の倍率を探す（scalesUp でなければ原寸より大きくしない）
     * @param {Array<{width: number, height: number}>} itemSizes - 置く順に並べた原寸の大きさ（pt）
     * @param {Object} layoutSettings - columnWidth / columnGap / rowGap（pt）/ scalesUp
     * @param {number} columnCount - 列数
     * @param {{width: number, height: number}} referenceSize - 基準の大きさ（pt）
     * @param {number} lowestScale - これまでの最良の倍率。これを超えられなければ探さない
     * @returns {number} 倍率（超えられなければ 0）
     */
    function findFitScale(itemSizes, layoutSettings, columnCount, referenceSize, lowestScale) {
        /* 列の並びが基準の幅に入る倍率が上限 / The columns must fit the reference width */
        var columnsWidth = referenceSize.width - (columnCount - 1) * layoutSettings.columnGap;
        var highScale = columnsWidth / (columnCount * layoutSettings.columnWidth);
        if (!layoutSettings.scalesUp) highScale = Math.min(highScale, 1);
        var lowScale = Math.max(lowestScale, MIN_FIT_SCALE);
        if (highScale <= lowScale) return 0;

        /**
         * 倍率が収まるか調べる
         * @param {number} scale - 倍率
         * @returns {boolean} 収まれば true
         */
        function fitsAt(scale) {
            return doesLayoutFit(buildScaledLayout(itemSizes, layoutSettings, columnCount, scale), referenceSize);
        }

        if (fitsAt(highScale)) return highScale;
        if (!fitsAt(lowScale)) return 0;
        for (var i = 0; i < FIT_SEARCH_STEPS && highScale - lowScale > FIT_SCALE_PRECISION; i++) {
            var middleScale = (lowScale + highScale) / 2;
            if (fitsAt(middleScale)) {
                lowScale = middleScale;
            } else {
                highScale = middleScale;
            }
        }
        return lowScale;
    }

    /**
     * 基準に収まる列数と倍率を決める。倍率が最も大きくなる列数を選び、同じなら列の多いほうを取る
     * @param {Array<{width: number, height: number}>} itemSizes - 置く順に並べた原寸の大きさ（pt）
     * @param {Object} layoutSettings - columnWidth / columnGap / rowGap（pt）/ scalesUp
     * @param {{width: number, height: number}} referenceSize - 基準の大きさ（pt）
     * @returns {{columnCount: number, scale: number}} 列数と倍率
     */
    function findReferenceFit(itemSizes, layoutSettings, referenceSize) {
        var bestFit = { columnCount: 1, scale: 0 };
        for (var columnCount = Math.min(itemSizes.length, MAX_COLUMN_COUNT); columnCount >= 1; columnCount--) {
            var fitScale = findFitScale(itemSizes, layoutSettings, columnCount, referenceSize, bestFit.scale);
            if (fitScale > bestFit.scale) bestFit = { columnCount: columnCount, scale: fitScale };
            /* 原寸が上限なら、原寸で収まった時点で探し終わる / With actual size as the limit, stop once it fits */
            if (!layoutSettings.scalesUp && bestFit.scale >= 1) break;
        }
        /* 間隔だけで基準を超えるときは下限まで縮める / When the gaps alone overflow, shrink to the lower limit */
        if (bestFit.scale === 0) bestFit.scale = MIN_FIT_SCALE;
        return bestFit;
    }

    // =========================================
    // 並べる順 / Placement order
    // =========================================

    /* 並べ替えのキーに足す値（pt）。千倍したキーがどの座標でも10桁にそろい、文字列のまま比べられる
       Added to sort keys (pt) so the key times 1000 has 10 digits at any coordinate and compares as a string */
    var SORT_KEY_OFFSET = 2000000;
    var SORT_KEY_SEPARATOR = "\u0001"; /* キーの区切り。数字より前に並ぶ / key separator, sorts before digits */
    var ROW_KEY_OFFSET = 1000000;      /* 行番号の桁をそろえるために足す値 / pads row numbers to a fixed width */

    /**
     * 座標を、文字列のまま大小を比べられるキーにする
     * @param {number} coordinate - 座標（pt）
     * @returns {string} 桁のそろったキー
     */
    function buildCoordinateKey(coordinate) {
        return String(Math.round((coordinate + SORT_KEY_OFFSET) * 1000));
    }

    /**
     * 末尾に番号を付けたキーを並べ替え、番号だけを取り出す。
     * 比較関数つきの sort() は並びが狂うので、文字列キーで並べる
     * @param {string[]} sortKeys - 「キー＋区切り＋番号」の配列
     * @returns {number[]} 並べ替えた番号
     */
    function sortIndexesByKeys(sortKeys) {
        sortKeys.sort();
        var sortedIndexes = [];
        for (var i = 0; i < sortKeys.length; i++) {
            var keyParts = sortKeys[i].split(SORT_KEY_SEPARATOR);
            sortedIndexes.push(parseInt(keyParts[keyParts.length - 1], 10));
        }
        return sortedIndexes;
    }

    /**
     * 今の位置を読む順（上の行から、行の中は左から）に並べる。
     * 行の先頭のオブジェクトの中心より上辺が上にあるものを同じ行とみなす
     * @param {Array<{bounds: number[]}>} itemStates - captureItemStates() の戻り値
     * @returns {number[]} itemStates の番号
     */
    function buildPositionOrder(itemStates) {
        var topKeys = [];
        for (var i = 0; i < itemStates.length; i++) {
            topKeys.push(buildCoordinateKey(-itemStates[i].bounds[1]) + SORT_KEY_SEPARATOR + i);
        }
        var topOrder = sortIndexesByKeys(topKeys);

        var readingKeys = [];
        var rowIndex = -1;
        var rowCenter = Infinity;
        for (var k = 0; k < topOrder.length; k++) {
            var itemBounds = itemStates[topOrder[k]].bounds;
            if (itemBounds[1] < rowCenter) {
                rowIndex++;
                rowCenter = (itemBounds[1] + itemBounds[3]) / 2;
            }
            readingKeys.push((ROW_KEY_OFFSET + rowIndex) + SORT_KEY_SEPARATOR +
                buildCoordinateKey(itemBounds[0]) + SORT_KEY_SEPARATOR + topOrder[k]);
        }
        return sortIndexesByKeys(readingKeys);
    }

    /**
     * 重ね順（前面から背面へ）に並べる。選択は前面から順に入っている
     * @param {Array} itemStates - captureItemStates() の戻り値
     * @returns {number[]} itemStates の番号
     */
    function buildStackingOrder(itemStates) {
        var stackingOrder = [];
        for (var i = 0; i < itemStates.length; i++) {
            stackingOrder.push(i);
        }
        return stackingOrder;
    }

    /**
     * 2つの順が同じか調べる
     * @param {number[]} orderA - 番号の配列
     * @param {number[]} orderB - 番号の配列（orderA と同じ長さ）
     * @returns {boolean} 同じなら true
     */
    function isSameOrder(orderA, orderB) {
        for (var i = 0; i < orderA.length; i++) {
            if (orderA[i] !== orderB[i]) return false;
        }
        return true;
    }

    /**
     * ランダムな順を作る（Fisher-Yates）。
     * 数が少ないと同じ順が出やすいので、直前と同じなら作り直す（［シャッフル］が空振りしないように）
     * @param {number} itemCount - オブジェクトの数（2以上）
     * @param {number[]} [previousOrder] - 直前の順
     * @returns {number[]} itemStates の番号
     */
    function buildRandomOrder(itemCount, previousOrder) {
        var randomOrder;
        do {
            randomOrder = [];
            for (var i = 0; i < itemCount; i++) {
                randomOrder.push(i);
            }
            for (var j = itemCount - 1; j > 0; j--) {
                var swapIndex = Math.floor(Math.random() * (j + 1));
                var swappedValue = randomOrder[j];
                randomOrder[j] = randomOrder[swapIndex];
                randomOrder[swapIndex] = swappedValue;
            }
        } while (previousOrder && isSameOrder(randomOrder, previousOrder));
        return randomOrder;
    }

    /**
     * 設定に合わせて並べる順を返す
     * @param {Array<{bounds: number[]}>} itemStates - captureItemStates() の戻り値
     * @param {{orderKey: string, randomOrder: number[]}} layoutSettings - 順序の設定
     * @returns {number[]} itemStates の番号
     */
    function getPlacementOrder(itemStates, layoutSettings) {
        if (layoutSettings.orderKey === "random") return layoutSettings.randomOrder;
        if (layoutSettings.orderKey === "stacking") return buildStackingOrder(itemStates);
        return buildPositionOrder(itemStates);
    }

    // =========================================
    // オブジェクトへの当てはめ / Applying to objects
    // =========================================

    /**
     * 選択したオブジェクトと、動かす前の境界を控える
     * @param {PageItem[]} items - 対象のオブジェクト
     * @param {boolean} usePreviewBounds - true で visibleBounds、false で geometricBounds
     * @returns {Array<{item: PageItem, bounds: number[]}>} 測れたオブジェクトと境界 [左, 上, 右, 下]
     */
    function captureItemStates(items, usePreviewBounds) {
        var itemStates = [];
        for (var i = 0; i < items.length; i++) {
            var itemBounds = getClipAwareBounds(items[i], usePreviewBounds);
            if (itemBounds) itemStates.push({ item: items[i], bounds: itemBounds });
        }
        return itemStates;
    }

    /**
     * 境界の幅を返す
     * @param {number[]} itemBounds - 境界 [左, 上, 右, 下]
     * @returns {number} 幅（pt）
     */
    function getBoundsWidth(itemBounds) {
        return itemBounds[2] - itemBounds[0];
    }

    /**
     * 境界の高さを返す
     * @param {number[]} itemBounds - 境界 [左, 上, 右, 下]
     * @returns {number} 高さ（pt）
     */
    function getBoundsHeight(itemBounds) {
        return itemBounds[1] - itemBounds[3];
    }

    /**
     * 境界を内側へ縮める
     * @param {number[]} outerBounds - 境界 [左, 上, 右, 下]
     * @param {number} insetAmount - 縮める量（pt）
     * @returns {number[]} 縮めた境界 [左, 上, 右, 下]
     */
    function getInsetBounds(outerBounds, insetAmount) {
        return [outerBounds[0] + insetAmount, outerBounds[1] - insetAmount, outerBounds[2] - insetAmount, outerBounds[3] + insetAmount];
    }

    /**
     * 並べた結果にするための動かし方を求める。
     * 起点は基準の角に固定し、右・下から並べるときは左右・上下を折り返す（Masonry の isOriginLeft / isOriginTop）。
     * 中央からのときは左上から並べた全体を、余りの半分だけずらす。
     * fitsReference のときは列数の設定を使わず、基準に収まる列数と倍率を決める。
     * matchesColumnWidth のときは先に各オブジェクトの幅を列の幅にしてから配置を求める（幅の無い線はそのまま）
     * @param {Array<{bounds: number[]}>} itemStates - captureItemStates() の戻り値
     * @param {Object} layoutSettings - referenceBounds（基準の [左, 上, 右, 下]）/ fitsReference / scalesUp / matchesColumnWidth / columnCount /
     *     columnWidth / columnGap / rowGap（pt）/ originX / originY / orderKey / randomOrder
     * @returns {{columnCount: number, moves: Array}} 列数と動かし方。moves の各要素は stateIndex / scale /
     *     targetX / targetY（合わせる座標）/ edgeX / edgeY（座標を合わせる辺の、境界の中の番号）
     */
    function computeLayoutPlan(itemStates, layoutSettings) {
        var placementOrder = getPlacementOrder(itemStates, layoutSettings);

        var itemSizes = [];
        var itemScales = [];
        for (var i = 0; i < placementOrder.length; i++) {
            var orderedBounds = itemStates[placementOrder[i]].bounds;
            var itemWidth = getBoundsWidth(orderedBounds);
            var itemScale = (layoutSettings.matchesColumnWidth && itemWidth > 0) ? layoutSettings.columnWidth / itemWidth : 1;
            itemScales.push(itemScale);
            itemSizes.push({ width: itemWidth * itemScale, height: getBoundsHeight(orderedBounds) * itemScale });
        }

        var referenceBounds = layoutSettings.referenceBounds;
        var layoutFit = { columnCount: layoutSettings.columnCount, scale: 1 };
        if (layoutSettings.fitsReference) {
            layoutFit = findReferenceFit(itemSizes, layoutSettings, {
                width: getBoundsWidth(referenceBounds),
                height: getBoundsHeight(referenceBounds)
            });
        }
        var scaledLayout = buildScaledLayout(itemSizes, layoutSettings, layoutFit.columnCount, layoutFit.scale);
        var positions = scaledLayout.positions;
        var layoutExtent = measureLayoutExtent(scaledLayout);

        var isFromRight = (layoutSettings.originX === "right");
        var isFromBottom = (layoutSettings.originY === "bottom");
        var centerOffsetX = (layoutSettings.originX === "centerX") ? (getBoundsWidth(referenceBounds) - layoutExtent.width) / 2 : 0;
        var centerOffsetY = (layoutSettings.originY === "centerY") ? (getBoundsHeight(referenceBounds) - layoutExtent.height) / 2 : 0;
        var layoutMoves = [];
        for (var j = 0; j < placementOrder.length; j++) {
            layoutMoves.push({
                stateIndex: placementOrder[j],
                scale: itemScales[j] * layoutFit.scale,
                /* Illustrator の Y は上が正 / Illustrator's Y grows upward */
                targetX: isFromRight ? referenceBounds[2] - positions[j].x : referenceBounds[0] + centerOffsetX + positions[j].x,
                targetY: isFromBottom ? referenceBounds[3] + positions[j].y : referenceBounds[1] - centerOffsetY - positions[j].y,
                edgeX: isFromRight ? 2 : 0,
                edgeY: isFromBottom ? 3 : 1
            });
        }
        return { columnCount: layoutFit.columnCount, moves: layoutMoves };
    }

    /**
     * オブジェクトを線幅ごと、左上を基準に拡大・縮小する
     * @param {PageItem} item - 対象のオブジェクト
     * @param {number} scale - 倍率（1 が原寸）
     * @returns {void}
     */
    function resizeItem(item, scale) {
        var scalePercent = scale * 100;
        item.resize(scalePercent, scalePercent, true, true, true, true, scalePercent, Transformation.TOPLEFT);
    }

    /**
     * 求めたとおりにオブジェクトを拡大・縮小して動かす。線や効果で大きさが倍率どおりにならないので、変えたあとに測り直して合わせる
     * @param {Array<{item: PageItem, bounds: number[]}>} itemStates - captureItemStates() の戻り値
     * @param {Array} layoutMoves - computeLayoutPlan() の moves
     * @param {boolean} usePreviewBounds - 控えたときと同じ境界の種類
     * @returns {boolean} 1つでも変えたら true
     */
    function applyLayoutMoves(itemStates, layoutMoves, usePreviewBounds) {
        var hasMoved = false;
        for (var i = 0; i < layoutMoves.length; i++) {
            var layoutMove = layoutMoves[i];
            var itemState = itemStates[layoutMove.stateIndex];
            var itemBounds = itemState.bounds;
            if (!isNearlySameCoordinate(layoutMove.scale, 1, SCALE_TOLERANCE)) {
                resizeItem(itemState.item, layoutMove.scale);
                itemBounds = getClipAwareBounds(itemState.item, usePreviewBounds) || itemBounds;
                hasMoved = true;
            }
            var dx = layoutMove.targetX - itemBounds[layoutMove.edgeX];
            var dy = layoutMove.targetY - itemBounds[layoutMove.edgeY];
            if (isNearlySameCoordinate(dx, 0) && isNearlySameCoordinate(dy, 0)) continue;
            itemState.item.translate(dx, dy);
            hasMoved = true;
        }
        return hasMoved;
    }

    /**
     * 控えた大きさと位置からずれているオブジェクトだけを戻す（app.undo() で戻りきらなかった分の保険）
     * @param {Array<{item: PageItem, bounds: number[]}>} itemStates - captureItemStates() の戻り値
     * @param {boolean} usePreviewBounds - 控えたときと同じ境界の種類
     * @returns {void}
     */
    function restoreItemStates(itemStates, usePreviewBounds) {
        for (var i = 0; i < itemStates.length; i++) {
            var itemState = itemStates[i];
            var currentBounds = getClipAwareBounds(itemState.item, usePreviewBounds);
            if (!currentBounds) continue;

            /* 線のように幅か高さが0のものがあるので、長いほうで比べる / Compare the longer side, as lines can have no width or height */
            var originalSpan = Math.max(getBoundsWidth(itemState.bounds), getBoundsHeight(itemState.bounds));
            var currentSpan = Math.max(getBoundsWidth(currentBounds), getBoundsHeight(currentBounds));
            if (currentSpan > 0 && !isNearlySameCoordinate(originalSpan, currentSpan)) {
                resizeItem(itemState.item, originalSpan / currentSpan);
                currentBounds = getClipAwareBounds(itemState.item, usePreviewBounds) || currentBounds;
            }

            var dx = itemState.bounds[0] - currentBounds[0];
            var dy = itemState.bounds[1] - currentBounds[1];
            /* undo が効いていれば差は0。ここで動かすと履歴が伸びる / Zero after a full undo; moving would add history */
            if (isNearlySameCoordinate(dx, 0) && isNearlySameCoordinate(dy, 0)) continue;
            itemState.item.translate(dx, dy);
        }
    }

    /**
     * プレビューの描き直しと取り消しをまとめる
     * @param {Array<{item: PageItem, bounds: number[]}>} itemStates - captureItemStates() の戻り値
     * @param {boolean} usePreviewBounds - 控えたときと同じ境界の種類
     * @returns {{update: function(Object): Object, clear: function(): void}} 描き直しと取り消し
     */
    function createLayoutPreview(itemStates, usePreviewBounds) {
        /* 1つも動かなかったときは false のまま。app.undo() がユーザーの操作を取り消さないようにする
           Stays false when nothing moved, so app.undo() never reverts the user's own step */
        var hasUncommittedPreview = false;

        /**
         * プレビューを取り除き、元の位置へ戻す。app.undo() で履歴ごと取り除き、戻りきらなかった分は座標で戻す
         * @returns {void}
         */
        function clearPreview() {
            if (hasUncommittedPreview) app.undo();
            hasUncommittedPreview = false;
            restoreItemStates(itemStates, usePreviewBounds);
        }

        /**
         * 指定の設定でプレビューを描き直す
         * @param {Object} layoutSettings - computeLayoutPlan() の設定
         * @returns {{columnCount: number, moves: Array}} computeLayoutPlan() の戻り値
         */
        function updatePreview(layoutSettings) {
            clearPreview();
            var layoutPlan = computeLayoutPlan(itemStates, layoutSettings);
            hasUncommittedPreview = applyLayoutMoves(itemStates, layoutPlan.moves, usePreviewBounds);
            app.redraw();
            return layoutPlan;
        }

        return { update: updatePreview, clear: clearPreview };
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 数値欄の値を読む。確定前の入力もそのまま読み、読めなければ直前の値を使う
     * @param {EditText} numberInput - addSteppedField() で作った入力欄
     * @param {Object} fieldOptions - addSteppedField() に渡した設定
     * @returns {number} 値（欄の単位）
     */
    function readFieldValue(numberInput, fieldOptions) {
        var fieldValue = evaluateArithmetic(numberInput.text, fieldOptions.unit);
        if (isNaN(fieldValue)) fieldValue = parseFloat(numberInput.lastValidText);
        if (fieldOptions.integer) fieldValue = Math.round(fieldValue);
        return clampSteppedValue(fieldValue, fieldOptions);
    }

    /**
     * 変更でプレビューを描き直す数値欄を足す
     * @param {Group|Panel} parentPanel - 追加先
     * @param {Object} fieldOptions - addSteppedField() の設定（onStep に描き直しの処理を入れる）
     * @param {string} helpTipText - ツールチップ
     * @returns {EditText} 入力欄
     */
    function addPreviewField(parentPanel, fieldOptions, helpTipText) {
        var numberInput = addSteppedField(parentPanel, fieldOptions);
        numberInput.helpTip = helpTipText;
        var normalizeField = numberInput.onChange;
        numberInput.onChange = function () {
            normalizeField();
            fieldOptions.onStep();
        };
        return numberInput;
    }

    /**
     * ラジオボタンを1つ足す
     * @param {Group|Panel} parentGroup - 追加先
     * @param {string} choiceKey - 選択肢のキー（LABELS.radio と LABELS.tooltip の名前）
     * @param {boolean} isInitial - 最初に選んでおくなら true
     * @param {function(string): void} onChoose - 選んだときの処理（キーを受け取る）
     * @returns {RadioButton} 足したラジオボタン
     */
    function addChoiceRadio(parentGroup, choiceKey, isInitial, onChoose) {
        var choiceRadio = parentGroup.add("radiobutton", undefined, getLabel(LABELS.radio[choiceKey]));
        choiceRadio.helpTip = getLabel(LABELS.tooltip[choiceKey]);
        choiceRadio.value = isInitial;
        choiceRadio.onClick = function () {
            onChoose(choiceKey);
        };
        return choiceRadio;
    }

    /**
     * 項目名とラジオボタンの行を足す
     * @param {Panel} parentPanel - 追加先
     * @param {Object} fieldLabelRef - 項目名の LABELS
     * @param {string[]} choiceKeys - 選択肢のキー（LABELS.radio と LABELS.tooltip の名前）
     * @param {function(string): void} onChoose - 選んだときの処理（キーを受け取る）
     * @returns {Object} キーごとのラジオボタン
     */
    function addChoiceRow(parentPanel, fieldLabelRef, choiceKeys, onChoose) {
        var choiceRow = parentPanel.add("group");
        setupRow(choiceRow, "left");
        var choiceLabel = choiceRow.add("statictext", undefined, labelText(fieldLabelRef));
        choiceLabel.preferredSize.width = FIELD_LABEL_WIDTHS[uiLang];
        choiceLabel.justify = "right";

        var choiceRadios = {};
        for (var i = 0; i < choiceKeys.length; i++) {
            choiceRadios[choiceKeys[i]] = addChoiceRadio(choiceRow, choiceKeys[i], i === 0, onChoose);
        }
        return choiceRadios;
    }

    /**
     * ダイアログを開き、OK なら配置を確定する。キャンセルなら元の位置に戻す
     * @param {Array<{item: PageItem, bounds: number[]}>} itemStates - captureItemStates() の戻り値
     * @param {{selection: number[], artboard: number[]}} referenceBoundsByKey - 基準ごとの境界 [左, 上, 右, 下]
     * @param {boolean} usePreviewBounds - 境界の種類
     * @returns {void}
     */
    function showLayoutDialog(itemStates, referenceBoundsByKey, usePreviewBounds) {
        var layoutPreview = createLayoutPreview(itemStates, usePreviewBounds);
        var rulerUnit = getUnitInfo();
        var unitSuffix = " " + rulerUnit.label;
        var labelWidth = FIELD_LABEL_WIDTHS[uiLang];

        /* 列の幅の既定は Masonry と同じく最初のオブジェクトの幅。幅が無ければ全体の幅にする
           The default column width is the first object's width, as in Masonry, or the whole width when it has none */
        var selectionWidth = getBoundsWidth(referenceBoundsByKey.selection);
        var firstItemBounds = itemStates[buildPositionOrder(itemStates)[0]].bounds;
        var defaultColumnWidth = getBoundsWidth(firstItemBounds) || selectionWidth || 1;
        var defaultColumnCount = fitColumnCount(selectionWidth, defaultColumnWidth, 0, itemStates.length);

        /* 表示前のラジオの value は読み戻せないので、選んでいる値は変数で持つ
           A radio's value cannot be read back before show(), so the choices live in variables */
        var selectedReferenceKey = "selection";
        var isFitArtboardOn = false;
        var isMatchWidthOn = false;
        var selectedOriginX = "left";
        var selectedOriginY = "top";
        var selectedOrderKey = "position";
        /* ほかの値を変えても並びが変わらないよう、順は［シャッフル］まで持ち続ける
           Kept until Shuffle is clicked, so changing other values does not reorder the objects */
        var randomOrder = buildRandomOrder(itemStates.length);
        /** @type {Object|null} OK を押したときの設定（キャンセルなら null） */
        var committedSettings = null;

        var layoutDialog = new Window("dialog", getLabel(LABELS.dialog.title));
        setupWindow(layoutDialog);

        var referencePanel = layoutDialog.add("panel", undefined, getLabel(LABELS.panel.reference));
        setupPanel(referencePanel, 6);
        addChoiceRadio(referencePanel, "selection", true, applyReference);
        addChoiceRadio(referencePanel, "artboard", false, applyReference);
        var fitArtboardCheckbox = referencePanel.add("checkbox", undefined, getLabel(LABELS.checkbox.fitArtboard));
        fitArtboardCheckbox.helpTip = getLabel(LABELS.tooltip.fitArtboard);
        fitArtboardCheckbox.enabled = false;
        fitArtboardCheckbox.onClick = function () {
            isFitArtboardOn = fitArtboardCheckbox.value;
            refreshPreview();
        };
        var artboardBounds = referenceBoundsByKey.artboard;
        var marginOptions = {
            label: labelText(LABELS.fieldLabel.margin),
            labelWidth: labelWidth,
            text: DEFAULT_MARGIN + unitSuffix,
            characters: 8,
            step: GAP_STEP,
            min: 0,
            max: ptToUnit(Math.min(getBoundsWidth(artboardBounds), getBoundsHeight(artboardBounds)) * MAX_MARGIN_RATIO, rulerUnit),
            unit: unitSuffix,
            onStep: refreshPreview
        };
        var marginField = addPreviewField(referencePanel, marginOptions, getLabel(LABELS.tooltip.margin));
        setSteppedFieldEnabled(marginField, false);

        var columnsPanel = layoutDialog.add("panel", undefined, getLabel(LABELS.panel.columns));
        setupPanel(columnsPanel);
        var columnCountOptions = {
            label: labelText(LABELS.fieldLabel.columnCount),
            labelWidth: labelWidth,
            text: String(defaultColumnCount),
            characters: 8,
            step: 1,
            min: 1,
            max: MAX_COLUMN_COUNT,
            integer: true,
            onStep: refreshPreview
        };
        var columnCountField = addPreviewField(columnsPanel, columnCountOptions, getLabel(LABELS.tooltip.columnCount));
        var columnWidthOptions = {
            label: labelText(LABELS.fieldLabel.columnWidth),
            labelWidth: labelWidth,
            characters: 8,
            step: GAP_STEP,
            min: MIN_COLUMN_WIDTH,
            unit: unitSuffix,
            onStep: refreshPreview
        };
        columnWidthOptions.text = formatSteppedValue(ptToUnit(defaultColumnWidth, rulerUnit), columnWidthOptions);
        var columnWidthField = addPreviewField(columnsPanel, columnWidthOptions, getLabel(LABELS.tooltip.columnWidth));
        var defaultColumnWidthText = columnWidthField.text;
        var matchWidthCheckbox = columnsPanel.add("checkbox", undefined, getLabel(LABELS.checkbox.matchWidth));
        matchWidthCheckbox.helpTip = getLabel(LABELS.tooltip.matchWidth);
        matchWidthCheckbox.onClick = function () {
            isMatchWidthOn = matchWidthCheckbox.value;
            refreshPreview();
        };

        var gapPanel = layoutDialog.add("panel", undefined, getLabel(LABELS.panel.gap));
        setupPanel(gapPanel);

        /* 欄を縦に積んだ group の右にリンクアイコンを置く / The link icon sits right of the stacked fields */
        var gapRow = gapPanel.add("group");
        setupRow(gapRow, "left");
        var gapFieldsGroup = gapRow.add("group");
        gapFieldsGroup.orientation = "column";
        gapFieldsGroup.alignChildren = ["left", "center"];
        gapFieldsGroup.spacing = PANEL_SPACING;

        /**
         * 間隔の数値欄の設定を作る
         * @param {Object} fieldLabelRef - 項目名の LABELS
         * @param {function(): void} onGapChange - 値を変えたときの処理
         * @returns {Object} addSteppedField() の設定
         */
        function buildGapOptions(fieldLabelRef, onGapChange) {
            return {
                label: labelText(fieldLabelRef),
                labelWidth: labelWidth,
                text: DEFAULT_GAP + unitSuffix,
                characters: 8,
                step: GAP_STEP,
                min: 0,
                unit: unitSuffix,
                onStep: onGapChange
            };
        }

        /**
         * 連動中なら、片方の間隔をもう片方へ写す
         * @param {EditText} sourceField - 写す元の欄
         * @param {Object} sourceOptions - 元の欄の設定
         * @param {EditText} targetField - 写す先の欄
         * @param {Object} targetOptions - 先の欄の設定
         * @returns {void}
         */
        function syncLinkedGap(sourceField, sourceOptions, targetField, targetOptions) {
            if (!gapLinkToggle.value) return;
            writeSteppedValue(targetField, readFieldValue(sourceField, sourceOptions), targetOptions);
        }

        var columnGapOptions = buildGapOptions(LABELS.fieldLabel.columnGap, function () {
            syncLinkedGap(columnGapField, columnGapOptions, rowGapField, rowGapOptions);
            refreshPreview();
        });
        var columnGapField = addPreviewField(gapFieldsGroup, columnGapOptions, getLabel(LABELS.tooltip.columnGap));
        var rowGapOptions = buildGapOptions(LABELS.fieldLabel.rowGap, function () {
            syncLinkedGap(rowGapField, rowGapOptions, columnGapField, columnGapOptions);
            refreshPreview();
        });
        var rowGapField = addPreviewField(gapFieldsGroup, rowGapOptions, getLabel(LABELS.tooltip.rowGap));

        var gapLinkToggle = addLinkToggle(gapRow, DEFAULT_LINK_GAPS, function () {
            syncLinkedGap(columnGapField, columnGapOptions, rowGapField, rowGapOptions);
            refreshPreview();
        });
        gapLinkToggle.helpTip = getLabel(LABELS.tooltip.linkGaps);

        /* 確定前の入力もそろえておく（OK を押すとフォーカスが外れないまま読むことがある）
           Mirror text as it is typed, since OK can read a field that still has focus */
        columnGapField.onChanging = function () {
            if (gapLinkToggle.value) rowGapField.text = columnGapField.text;
        };
        rowGapField.onChanging = function () {
            if (gapLinkToggle.value) columnGapField.text = rowGapField.text;
        };

        var arrangementPanel = layoutDialog.add("panel", undefined, getLabel(LABELS.panel.arrangement));
        setupPanel(arrangementPanel, 6);
        var originXRadios = addChoiceRow(arrangementPanel, LABELS.fieldLabel.horizontalStart, ["left", "centerX", "right"], function (choiceKey) {
            selectedOriginX = choiceKey;
            refreshPreview();
        });
        var originYRadios = addChoiceRow(arrangementPanel, LABELS.fieldLabel.verticalStart, ["top", "centerY", "bottom"], function (choiceKey) {
            selectedOriginY = choiceKey;
            refreshPreview();
        });
        /* ［中央から］はアートボードが基準のときだけ / From Center is only for the artboard */
        originXRadios.centerX.enabled = false;
        originYRadios.centerY.enabled = false;
        addChoiceRow(arrangementPanel, LABELS.fieldLabel.order, ["position", "stacking", "random"], function (choiceKey) {
            selectedOrderKey = choiceKey;
            btnShuffle.enabled = (choiceKey === "random");
            refreshPreview();
        });

        var buttonRow = addButtonRow(layoutDialog);
        var btnShuffle = buttonRow.leftGroup.add("button", undefined, getLabel(LABELS.button.shuffle));
        btnShuffle.helpTip = getLabel(LABELS.tooltip.shuffle);
        btnShuffle.enabled = false;
        buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });
        var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.ok), { name: "ok" });
        alignRightOnlyButtonRow(buttonRow);

        /**
         * 列の幅を pt で読む。既定のままなら表示の丸めを通さず、測った幅をそのまま使う
         * @returns {number} 列の幅（pt）
         */
        function readColumnWidth() {
            if (columnWidthField.text === defaultColumnWidthText) return defaultColumnWidth;
            return unitToPt(readFieldValue(columnWidthField, columnWidthOptions), rulerUnit);
        }

        /**
         * 基準の境界を読む。アートボードなら余白のぶん内側にする
         * @returns {number[]} 基準の境界 [左, 上, 右, 下]
         */
        function readReferenceBounds() {
            if (selectedReferenceKey === "selection") return referenceBoundsByKey.selection;
            return getInsetBounds(artboardBounds, unitToPt(readFieldValue(marginField, marginOptions), rulerUnit));
        }

        /**
         * 今の設定を集める
         * @returns {Object} computeLayoutPlan() の設定
         */
        function readLayoutSettings() {
            return {
                referenceBounds: readReferenceBounds(),
                fitsReference: (selectedReferenceKey === "artboard"),
                scalesUp: isFitArtboardOn,
                matchesColumnWidth: isMatchWidthOn,
                columnCount: readFieldValue(columnCountField, columnCountOptions),
                columnWidth: readColumnWidth(),
                columnGap: unitToPt(readFieldValue(columnGapField, columnGapOptions), rulerUnit),
                rowGap: unitToPt(readFieldValue(rowGapField, rowGapOptions), rulerUnit),
                originX: selectedOriginX,
                originY: selectedOriginY,
                orderKey: selectedOrderKey,
                randomOrder: randomOrder
            };
        }

        /**
         * 今の設定でプレビューを描き直す。アートボードが基準なら、決まった列数を欄に出す
         * @returns {void}
         */
        function refreshPreview() {
            var layoutPlan = layoutPreview.update(readLayoutSettings());
            if (selectedReferenceKey === "artboard") writeSteppedValue(columnCountField, layoutPlan.columnCount, columnCountOptions);
        }

        /**
         * 基準を切り替える。アートボードでは収まりを優先して列の設定を止め、選択オブジェクトでは幅に入る列数に戻す
         * @param {string} referenceKey - "selection" または "artboard"
         * @returns {void}
         */
        function applyReference(referenceKey) {
            selectedReferenceKey = referenceKey;
            var usesColumnSettings = (referenceKey === "selection");
            fitArtboardCheckbox.enabled = !usesColumnSettings;
            setSteppedFieldEnabled(marginField, !usesColumnSettings);
            originXRadios.centerX.enabled = !usesColumnSettings;
            originYRadios.centerY.enabled = !usesColumnSettings;
            setSteppedFieldEnabled(columnCountField, usesColumnSettings);
            setSteppedFieldEnabled(columnWidthField, usesColumnSettings);
            if (usesColumnSettings) {
                /* 使えなくなる［中央から］は、左・上に戻す / From Center becomes unavailable, so fall back to left and top */
                if (selectedOriginX === "centerX") {
                    selectedOriginX = "left";
                    originXRadios.centerX.value = false;
                    originXRadios.left.value = true;
                }
                if (selectedOriginY === "centerY") {
                    selectedOriginY = "top";
                    originYRadios.centerY.value = false;
                    originYRadios.top.value = true;
                }
                var columnGap = unitToPt(readFieldValue(columnGapField, columnGapOptions), rulerUnit);
                var fittedCount = fitColumnCount(selectionWidth, readColumnWidth(), columnGap, itemStates.length);
                writeSteppedValue(columnCountField, fittedCount, columnCountOptions);
            }
            refreshPreview();
        }

        btnShuffle.onClick = function () {
            randomOrder = buildRandomOrder(itemStates.length, randomOrder);
            refreshPreview();
        };

        btnOK.onClick = function () {
            committedSettings = readLayoutSettings();
            layoutDialog.close(1);
        };

        addVersionFooter(layoutDialog);
        prepareDialogWindow(layoutDialog, SCRIPT_NAME);
        refreshPreview();
        layoutDialog.show();

        /* プレビューを消してから本番を1回だけ流す（取り消しが1回で済む）
           Drop the preview, then run once for real so a single undo reverts it */
        layoutPreview.clear();
        if (committedSettings !== null) {
            applyLayoutMoves(itemStates, computeLayoutPlan(itemStates, committedSettings).moves, usePreviewBounds);
        }
        app.redraw();
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択を確かめてダイアログを開く
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel(LABELS.alert.noDocument));
            return;
        }

        var activeDoc = app.activeDocument;
        var selectedItems = collectSelectionItems(activeDoc.selection, {
            accept: function () { return true; }, /* グループは中へ入らず1つとして扱う / keep groups whole */
            skipLocked: true,
            skipHidden: true
        });
        var usePreviewBounds = readUsePreviewBoundsPreference();
        var itemStates = captureItemStates(selectedItems, usePreviewBounds);
        if (itemStates.length < MIN_ITEM_COUNT) {
            alert(getLabel(LABELS.alert.needItems, { count: MIN_ITEM_COUNT }));
            return;
        }
        var referenceBoundsByKey = {
            selection: getClipAwareUnionBounds(selectedItems, usePreviewBounds),
            artboard: activeDoc.artboards[activeDoc.artboards.getActiveArtboardIndex()].artboardRect
        };
        showLayoutDialog(itemStates, referenceBoundsByKey, usePreviewBounds);
    }

    main();

})();
