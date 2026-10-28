#target illustrator
#targetengine "PatternSwatchSquaresEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
SCRIPTMETA-BEGIN
Script-ID=com.swwwitch.illustrator.PatternSwatchSquares
Name=PatternSwatchSquares
Author=Masahiro Takano (@swwwitch)
Target-App=Illustrator
Version=1.1.2
Release-Date=2026-10-27
Description-BEGIN
Illustrator 付属のスウォッチライブラリ（パターン・グラデーション）やドキュメントのスウォッチを、長方形に1つずつ適用して並べた見本を作ります。
ライブラリは複数選べ、ライブラリごとのブロックを縦に並べて、作業中のアートボードの中央に置きます。

Creates samples by filling rectangles, one per swatch, with Illustrator's bundled swatch libraries (patterns and gradients) or the document's swatches.
Several libraries can be chosen; one block per library is stacked vertically and centered on the active artboard.
Description-END
SCRIPTMETA-END
*/

/*

### 概要

Illustrator 付属のスウォッチライブラリ（パターン・グラデーション）やドキュメントのスウォッチを、長方形に1つずつ適用して並べた見本を作ります。
ライブラリは複数選べ、ライブラリごとのブロックを縦に並べて、作業中のアートボードの中央に置きます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/PatternSwatchSquares.md

note記事も参照してください。
https://note.com/dtp_tranist/n/nbf88198f1581

### Overview

Creates samples by filling rectangles, one per swatch, with Illustrator's bundled swatch libraries (patterns and gradients) or the document's swatches.
Several libraries can be chosen; one block per library is stacked vertically and centered on the active artboard.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/PatternSwatchSquares.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "PatternSwatchSquares";         /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-10-27";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-27";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/PatternSwatchSquares.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/PatternSwatchSquares.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/nbf88198f1581"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================
    var DEFAULT_SQUARE_WIDTH = 60;    /* 長方形の幅の初期値（pt = px） / Default square width */
    var DEFAULT_SQUARE_HEIGHT = 60;   /* 長方形の高さの初期値 / Default square height */
    var DEFAULT_SIZE_LINKED = true;   /* 幅と高さの連動の初期値 / Width and height linked by default */
    var DEFAULT_SQUARE_MARGIN = 40;   /* 長方形どうしの間隔の初期値 / Default gap between squares */
    var DEFAULT_STROKE_WIDTH = 1;     /* 線幅の初期値（pt）。線の色の初期値は黒 / Default stroke width; the stroke color defaults to black */
    var DEFAULT_ADD_TITLE = false;    /* ［ライブラリ名］の初期値 / Library Name is off by default */
    var DEFAULT_SHOW_SWATCH_NAMES = false; /* ［スウォッチ名］の初期値 / Swatch Name is off by default */
    var DEFAULT_PREVIEW = true;       /* ［プレビュー（長方形のみ）］の初期値 / Preview is on by default */
    var SQUARES_PER_ROW = 10;         /* 1行に並べる数 / Squares per row */
    var TITLE_FONT_SIZE = 14;         /* タイトルの文字サイズ（pt） / Title font size */
    var TITLE_GAP = 10;               /* 長方形（スウォッチ名）の上端とタイトルの下端の間隔（pt） / Gap below the title */
    var SWATCH_NAME_FONT_SIZE = 5;    /* スウォッチ名の文字サイズ（pt） / Swatch name font size */
    var SWATCH_NAME_GAP = 2;          /* 長方形の上端とスウォッチ名の下端の間隔（pt） / Gap between a square and its name */
    var APPLY_TEXT_MOJIKUMI = true;   /* 日本語環境でテキストに文字組みアキ量設定を適用するか / apply the mojikumi set in Japanese UIs */
    var TEXT_MOJIKUMI_INDEX = 5;      /* 日本語環境で使う文字組みアキ量設定（mojikumiSet の番号。5 = 詰め組み） / mojikumi set for Japanese UIs (5 = Tsume-gumi) */
    var TEXT_FONT_NAME = "HiraginoSans-W3"; /* タイトル・スウォッチ名のフォント（PostScript 名）。無ければ既定のまま / Font for the title and names; the default is kept if missing */

    // =========================================
    // レイアウト / Layout
    // =========================================
    var FIELD_LABEL_WIDTH = 60;       /* 数値欄の項目名の幅 / width of the field labels */
    var FIELD_CHARACTERS = 7;         /* 数値欄の文字数（単位込み） / characters of the number fields, unit included */
    var PREVIEW_HAIRLINE_WIDTH = 0.5; /* 線幅 0 のときのプレビューの線幅（pt） / preview stroke when the weight is 0 */
    var COLOR_SWATCH_SIZE = [40, 20]; /* 線のカラーの色見本の大きさ / size of the stroke color swatch */
    var LIBRARY_LIST_WIDTH = 200;     /* ライブラリ一覧のチェックボックスの幅 / width of the checkboxes in the library list */
    var DIVIDER_HEIGHT = 2;           /* 一覧の区切り線の高さ / height of the list dividers */

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
            title: { ja: "スウォッチの見本を作成", en: "Swatch Samples" }
        },
        panel: {
            swatch: { ja: "スウォッチ", en: "Swatches" },
            libraryList: { ja: "ライブラリ", en: "Libraries" },
            square: { ja: "長方形", en: "Rectangle" },
            stroke: { ja: "線", en: "Stroke" },
            text: { ja: "テキストを追加", en: "Add Text" }
        },
        radio: {
            pattern: { ja: "パターン", en: "Patterns" },
            patternClassic: { ja: "パターン（クラシック）", en: "Patterns (Classic)" },
            gradient: { ja: "グラデーション", en: "Gradients" },
            gradientClassic: { ja: "グラデーション（クラシック）", en: "Gradients (Classic)" },
            library: { ja: "ライブラリ", en: "Library" },
            document: { ja: "ドキュメント", en: "Document" },
            selectedSwatches: { ja: "選択したスウォッチ", en: "Selected Swatches" },
            allSwatches: { ja: "すべてのスウォッチ", en: "All Swatches" }
        },
        checkbox: {
            libraryName: { ja: "ライブラリ名（ドキュメント名）", en: "Library Name (Document Name)" },
            swatchName: { ja: "スウォッチ名", en: "Swatch Name" },
            preview: { ja: "プレビュー（長方形のみ）", en: "Preview (Rectangles Only)" }
        },
        fieldLabel: {
            width: { ja: "幅", en: "Width" },
            height: { ja: "高さ", en: "Height" },
            margin: { ja: "間隔", en: "Spacing" },
            strokeWidth: { ja: "線幅", en: "Weight" },
            strokeColor: { ja: "カラー", en: "Color" },
            count: { ja: "件数", en: "Count" }
        },
        tooltip: {
            library: { ja: "右の一覧でチェックを入れた Illustrator 付属のスウォッチライブラリを使います", en: "Use the bundled swatch libraries checked in the list on the right" },
            documentSource: { ja: "このドキュメントのスウォッチを、種類を問わず（カラー・パターン・グラデーション）使います", en: "Use this document's swatches of any kind (colors, patterns, gradients)" },
            selectedSwatches: { ja: "スウォッチパネルで選択しているスウォッチだけを使います", en: "Use only the swatches selected in the Swatches panel" },
            libraryList: { ja: "チェックを入れたライブラリごとに見本のブロックを作り、縦に並べます", en: "One block of samples is created for each checked library, stacked vertically" },
            libraryCheckbox: { ja: "option＋クリックで、これだけにチェック（もう一度ですべてにチェック）", en: "Option-click to check only this one (again to check all)" },
            linkToggle: { ja: "幅と高さを連動", en: "Link width and height" },
            margin: { ja: "長方形どうしの間隔。ライブラリのブロックの間はこの2倍", en: "Spacing between the rectangles; blocks of libraries are spaced twice this" },
            strokeWidth: { ja: "0 にすると線なし", en: "0 for no stroke" },
            strokeColor: { ja: "クリックしてカラーピッカーを開く", en: "Click to open the Color Picker" },
            libraryName: {
                ja: "各ブロックの左上の長方形の上に、ライブラリ名（ドキュメントのときはドキュメント名）を入れます",
                en: "Adds the library name (the document name for document swatches) above the top-left rectangle of each block"
            },
            swatchName: { ja: "各長方形の上に、スウォッチ名を入れます", en: "Adds the swatch name above each rectangle" },
            preview: { ja: "設定を変えるたびに、長方形の並び（パターン・グラデーションは塗らない）をドキュメントに描いて確かめます", en: "Draws the layout of the rectangles in the document as you change the settings (without the patterns or gradients)" },
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
            selectAll: { ja: "すべてを選択", en: "Select All" },
            deselectAll: { ja: "選択を解除", en: "Deselect All" },
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" }
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません", en: "No document is open." },
            noPattern: { ja: "パターンスウォッチがありません", en: "There are no pattern swatches." },
            noGradient: { ja: "グラデーションスウォッチがありません", en: "There are no gradient swatches." },
            noSelectedSwatch: { ja: "スウォッチパネルでスウォッチが選択されていません", en: "No swatch is selected in the Swatches panel." },
            noDocumentSwatch: { ja: "このドキュメントには使えるスウォッチがありません", en: "This document has no usable swatches." },
            noLibraryFolder: { ja: "Illustrator 付属のスウォッチライブラリが見つかりません", en: "Illustrator's bundled swatch libraries were not found." },
            noCheckedLibrary: { ja: "一覧でライブラリにチェックを入れてください", en: "Check at least one library in the list." }
        }
    };

    // =========================================
    // スウォッチの種類 / Swatch kinds
    // =========================================

    /* 種類ごとの色の typename・ライブラリの場所（Illustrator のプリセットからの相対パス。言語ごとに候補を並べる）・
       サブフォルダーも探すか・ラベル。クラシックでない種類はフォルダー直下のファイルだけ（クラシックのフォルダーは含めない）
       Per kind: color typename, library folders relative to the presets (candidates per language), whether to search subfolders, labels.
       Non-classic kinds take only the files directly in the folder (not the classic folder) */
    var SWATCH_KINDS = {
        pattern: {
            colorTypename: "PatternColor",
            librarySubpaths: ["スウォッチ/パターン", "Swatches/Patterns"],
            includesSubfolders: false,
            radioLabel: "radio.pattern",
            emptyAlert: "alert.noPattern"
        },
        patternClassic: {
            colorTypename: "PatternColor",
            librarySubpaths: ["スウォッチ/パターン/パターン (クラシック)", "Swatches/Patterns/Patterns (Classic)"],
            includesSubfolders: true,
            radioLabel: "radio.patternClassic",
            emptyAlert: "alert.noPattern"
        },
        gradient: {
            colorTypename: "GradientColor",
            librarySubpaths: ["スウォッチ/グラデーション", "Swatches/Gradients"],
            includesSubfolders: false,
            radioLabel: "radio.gradient",
            emptyAlert: "alert.noGradient"
        },
        gradientClassic: {
            colorTypename: "GradientColor",
            librarySubpaths: ["スウォッチ/グラデーション/グラデーション (クラシック)", "Swatches/Gradients/Gradients (Classic)"],
            includesSubfolders: true,
            radioLabel: "radio.gradientClassic",
            emptyAlert: "alert.noGradient"
        }
    };
    var SWATCH_KIND_KEYS = ["pattern", "patternClassic", "gradient", "gradientClassic"]; /* ラジオボタンの並び / radio button order */

    // =========================================
    // スウォッチライブラリ / Swatch libraries
    // =========================================

    /**
     * 種類のスウォッチライブラリのファイルを返す（Illustrator 付属のフォルダーが見つからなければ空）
     * @param {string} kindKey - SWATCH_KINDS のキー
     * @returns {Array<File|null>} ライブラリのファイル。サブフォルダーの切れ目は null（一覧の区切り線）
     */
    function getKindLibraryFiles(kindKey) {
        var illustratorPath = app.path.fsName;
        var presetsRoots = [illustratorPath + "/Presets.localized/" + app.locale, illustratorPath + "/Presets/" + app.locale];
        var swatchKind = SWATCH_KINDS[kindKey];
        for (var i = 0; i < presetsRoots.length; i++) {
            for (var j = 0; j < swatchKind.librarySubpaths.length; j++) {
                var libraryFolder = new Folder(presetsRoots[i] + "/" + swatchKind.librarySubpaths[j]);
                if (libraryFolder.exists) return collectLibraryFiles(libraryFolder, swatchKind.includesSubfolders);
            }
        }
        return [];
    }

    /**
     * フォルダー内の .ai ファイルを名前順に集める。
     * サブフォルダーも探すときは、ファイルとサブフォルダーを名前順に混ぜ、サブフォルダーはその位置に中身を展開し、
     * グループの切れ目に null を挟む（Illustrator の［スウォッチライブラリを開く］のメニューと同じ順・区切り）
     * @param {Folder} libraryFolder - 探すフォルダー
     * @param {boolean} includesSubfolders - サブフォルダーも探すなら true
     * @returns {Array<File|null>} ライブラリのファイル（切れ目は null）
     */
    function collectLibraryFiles(libraryFolder, includesSubfolders) {
        var fileGroups = [];
        var ownFiles = [];
        var folderEntries = libraryFolder.getFiles();
        folderEntries.sort(compareByComposedName);
        for (var i = 0; i < folderEntries.length; i++) {
            var folderEntry = folderEntries[i];
            if (folderEntry instanceof Folder) {
                if (!includesSubfolders) continue;
                if (ownFiles.length > 0) {
                    fileGroups.push(ownFiles);
                    ownFiles = [];
                }
                var subfolderFiles = collectLibraryFiles(folderEntry, true);
                if (subfolderFiles.length > 0) fileGroups.push(subfolderFiles);
            } else if (/\.ai$/i.test(folderEntry.name)) {
                ownFiles.push(folderEntry);
            }
        }
        if (ownFiles.length > 0) fileGroups.push(ownFiles);

        var libraryFiles = [];
        for (var j = 0; j < fileGroups.length; j++) {
            if (j > 0) libraryFiles.push(null);
            libraryFiles = libraryFiles.concat(fileGroups[j]);
        }
        return libraryFiles;
    }

    /**
     * ファイル・フォルダーの名前を、濁点・半濁点を合成して返す
     * @param {File|Folder} fileEntry - ファイルまたはフォルダー
     * @returns {string} 合成した名前
     */
    function getComposedName(fileEntry) {
        return composeKanaMarks(decodeURI(fileEntry.name));
    }

    /**
     * ファイル・フォルダーを、濁点・半濁点を合成した名前で比べる
     * （NFD のままだと「パ」が「ハ＋゜」になり、「ハッチング」より前に来てしまう）
     * @param {File|Folder} entryA - 比べるもの
     * @param {File|Folder} entryB - 比べるもの
     * @returns {number} 並び順（負・0・正）
     */
    function compareByComposedName(entryA, entryB) {
        var nameA = getComposedName(entryA);
        var nameB = getComposedName(entryB);
        return (nameA < nameB) ? -1 : ((nameA > nameB) ? 1 : 0);
    }

    /**
     * ライブラリのファイル名から拡張子を外した表示名を返す
     * @param {File} libraryFile - ライブラリのファイル
     * @returns {string} 表示名
     */
    function getLibraryDisplayName(libraryFile) {
        return getComposedName(libraryFile).replace(/\.ai$/i, "");
    }

    /* 濁点・半濁点を合成する前と後の仮名 / Kana before and after composing the voiced marks */
    var DAKUTEN_BASE_KANA    = "かきくけこさしすせそたちつてとはひふへほうカキクケコサシスセソタチツテトハヒフヘホウ";
    var DAKUTEN_KANA         = "がぎぐげござじずぜぞだぢづでどばびぶべぼゔガギグゲゴザジズゼゾダヂヅデドバビブベボヴ";
    var HANDAKUTEN_BASE_KANA = "はひふへほハヒフヘホ";
    var HANDAKUTEN_KANA      = "ぱぴぷぺぽパピプペポ";

    /**
     * 結合文字の濁点（U+3099）・半濁点（U+309A）を前の仮名と合成する。
     * Illustrator 付属のファイル名は NFD で、そのままでは一覧の文字が空白になる
     * @param {string} sourceText - 元の文字列
     * @returns {string} 合成した文字列
     */
    function composeKanaMarks(sourceText) {
        return sourceText.replace(/([\u3040-\u30ff])([\u3099\u309a])/g, function (matched, baseKana, combiningMark) {
            var isDakuten = (combiningMark === "\u3099");
            var markIndex = (isDakuten ? DAKUTEN_BASE_KANA : HANDAKUTEN_BASE_KANA).indexOf(baseKana);
            if (markIndex < 0) return matched;
            return (isDakuten ? DAKUTEN_KANA : HANDAKUTEN_KANA).charAt(markIndex);
        });
    }

    /**
     * ライブラリのファイルを開く（すでに開いていればそのドキュメント）。開くときの警告は出さない
     * @param {File} libraryFile - スウォッチライブラリのファイル
     * @returns {{doc: Document, wasAlreadyOpen: boolean}} 開いたドキュメントと、もともと開いていたか
     */
    function openLibraryDocument(libraryFile) {
        var documentCountBefore = app.documents.length;
        var previousInteractionLevel = app.userInteractionLevel;
        app.userInteractionLevel = UserInteractionLevel.DONTDISPLAYALERTS;
        var libraryDoc;
        /* 開けなくても警告の設定は戻す / Restore the alert setting even when the file cannot be opened */
        try {
            libraryDoc = app.open(libraryFile);
        } finally {
            app.userInteractionLevel = previousInteractionLevel;
        }
        app.activeDocument = libraryDoc;
        /* 数が増えなければ、すでに開いていたドキュメント / If the count did not grow, it was already open */
        return { doc: libraryDoc, wasAlreadyOpen: (app.documents.length === documentCountBefore) };
    }

    /**
     * 開いたライブラリを閉じ（もともと開いていたものは残し）、元のドキュメントに戻る
     * @param {{doc: Document, wasAlreadyOpen: boolean}} openedLibrary - openLibraryDocument() の戻り値
     * @param {Document} targetDoc - 戻るドキュメント
     * @returns {void}
     */
    function closeLibraryDocument(openedLibrary, targetDoc) {
        if (!openedLibrary.wasAlreadyOpen) openedLibrary.doc.close(SaveOptions.DONOTSAVECHANGES);
        app.activeDocument = targetDoc;
    }

    // =========================================
    // スウォッチの取得 / Collecting swatches
    // =========================================

    /**
     * 塗りに使えるスウォッチか（［なし］と［レジストレーション］は除く）
     * @param {Swatch} swatch - スウォッチ
     * @returns {boolean} 使えるなら true
     */
    function isFillableSwatch(swatch) {
        var swatchColor;
        /* 色を読むと MRAP エラーになるスウォッチがあるので、読めないものは使わない
           Some swatches throw MRAP when their color is read; skip those */
        try {
            swatchColor = swatch.color;
        } catch (e) {
            return false;
        }
        if (swatchColor.typename === "NoColor") return false;
        return !(swatchColor.typename === "SpotColor" && swatchColor.spot.colorType === ColorModel.REGISTRATION);
    }

    /**
     * スウォッチの配列から、指定した種類の色のスウォッチだけを抜き出す
     * @param {Swatches|Swatch[]} swatchList - スウォッチの一覧
     * @param {string|null} colorTypename - 色の typename（"PatternColor" / "GradientColor"）。null なら種類を問わない
     * @returns {Swatch[]} 抜き出したスウォッチ
     */
    function filterSwatchesByColorType(swatchList, colorTypename) {
        var matchedSwatches = [];
        for (var i = 0; i < swatchList.length; i++) {
            if (!isFillableSwatch(swatchList[i])) continue;
            if (colorTypename === null || swatchList[i].color.typename === colorTypename) matchedSwatches.push(swatchList[i]);
        }
        return matchedSwatches;
    }

    /**
     * ドキュメントの対象のスウォッチを返す
     * @param {Document} doc - 対象のドキュメント
     * @param {string|null} colorTypename - 色の typename。null なら種類を問わない
     * @param {boolean} selectedOnly - スウォッチパネルで選択しているものだけなら true
     * @returns {Swatch[]} スウォッチ
     */
    function getDocumentSwatches(doc, colorTypename, selectedOnly) {
        return filterSwatchesByColorType(selectedOnly ? doc.swatches.getSelected() : doc.swatches, colorTypename);
    }

    // =========================================
    // 長方形とタイトルの作成 / Drawing squares and the title
    // =========================================

    /**
     * @typedef {Object} SquareLayout
     * @property {number} width - 長方形の幅（pt）
     * @property {number} height - 長方形の高さ（pt）
     * @property {number} margin - 長方形どうしの間隔（pt）
     * @property {number} strokeWidth - 線幅（pt）。0 なら線なし
     * @property {Color} strokeColor - 線の色
     * @property {boolean} showSwatchNames - 長方形の上にスウォッチ名を入れるなら true
     */

    /**
     * レイヤーとその親がすべて編集できる（ロック・非表示でない）か
     * @param {Layer} layer - 調べるレイヤー
     * @returns {boolean} 編集できるなら true
     */
    function isLayerEditable(layer) {
        for (var currentLayer = layer; currentLayer && currentLayer.typename === "Layer"; currentLayer = currentLayer.parent) {
            if (currentLayer.locked || !currentLayer.visible) return false;
        }
        return true;
    }

    /**
     * 描く先のレイヤーを返す。作業中のレイヤー（またはその親）がロック・非表示なら、新しいレイヤーを足す
     * @param {Document} doc - 対象のドキュメント
     * @returns {Layer} 描く先のレイヤー
     */
    function getDrawableLayer(doc) {
        if (isLayerEditable(doc.activeLayer)) return doc.activeLayer;
        return doc.layers.add();
    }

    /**
     * スウォッチ名を入れるとき、長方形の上に空ける高さを返す
     * @param {SquareLayout} squareLayout - 長方形の設定
     * @returns {number} 空ける高さ（pt）。名前を入れないなら 0
     */
    function getSwatchNameSpace(squareLayout) {
        return squareLayout.showSwatchNames ? SWATCH_NAME_FONT_SIZE + SWATCH_NAME_GAP : 0;
    }

    /**
     * TEXT_FONT_NAME のフォントを返す
     * @returns {TextFont|null} フォント。入っていなければ null
     */
    function findTextFont() {
        /* getByName は見つからないと例外になる / getByName throws when the font is missing */
        try {
            return app.textFonts.getByName(TEXT_FONT_NAME);
        } catch (e) {
            return null;
        }
    }

    var textFont; /* 初めて文字を置くときに引く / looked up when the first text is placed */

    /**
     * テキストの段落に文字組みアキ量設定（TEXT_MOJIKUMI_INDEX、詰め組み）を適用する。
     * textRange にまとめて代入すると MRAP エラーの元になったので、段落ごとに代入する（mojikumiSet の要素は内部 ID の文字列）
     * @param {TextFrame} pointText - 対象のテキスト
     * @param {Document} doc - テキストのあるドキュメント
     * @returns {void}
     */
    function applyTextMojikumi(pointText, doc) {
        var mojikumiValue = doc.mojikumiSet[TEXT_MOJIKUMI_INDEX];
        var textParagraphs = pointText.paragraphs;
        for (var i = 0; i < textParagraphs.length; i++) textParagraphs[i].paragraphAttributes.mojikumi = mojikumiValue;
    }

    /**
     * ポイント文字を作り、左端と下端をそろえて置く（フォントは TEXT_FONT_NAME）
     * @param {Layer} targetLayer - 置く先のレイヤー
     * @param {string} textContents - 文字列
     * @param {number} fontSize - 文字サイズ（pt）
     * @param {number} left - 左端の X 座標
     * @param {number} bottom - 下端の Y 座標
     * @returns {TextFrame} 置いたポイント文字
     */
    function addPointTextAt(targetLayer, textContents, fontSize, left, bottom) {
        if (textFont === undefined) textFont = findTextFont();
        var pointText = targetLayer.textFrames.add();
        pointText.contents = textContents;
        var textAttributes = pointText.textRange.characterAttributes;
        if (textFont) textAttributes.textFont = textFont;
        textAttributes.size = fontSize;
        textAttributes.kerningMethod = AutoKernType.AUTO; /* カーニング：メトリクス / Kerning: Metrics */
        textAttributes.proportionalMetrics = true;        /* プロポーショナルメトリクス / Proportional metrics */
        if (APPLY_TEXT_MOJIKUMI && uiLang === "ja") applyTextMojikumi(pointText, getParentDocument(targetLayer));
        var textBounds = pointText.geometricBounds;
        pointText.translate(left - textBounds[0], bottom - textBounds[3]);
        return pointText;
    }

    /**
     * レイヤーの属するドキュメントを返す（サブレイヤーも親をたどる）
     * @param {Layer} layer - レイヤー
     * @returns {Document} ドキュメント
     */
    function getParentDocument(layer) {
        var parentItem = layer.parent;
        while (parentItem.typename !== "Document") parentItem = parentItem.parent;
        return parentItem;
    }

    /**
     * ドキュメントのカラーモードに合わせた黒を返す
     * @param {Document} doc - 対象のドキュメント
     * @returns {CMYKColor|RGBColor} 黒
     */
    function createBlackColor(doc) {
        if (doc.documentColorSpace === DocumentColorSpace.CMYK) {
            var cmykBlack = new CMYKColor();
            cmykBlack.black = 100;
            return cmykBlack;
        }
        return new RGBColor(); /* 既定値が 0,0,0 / Defaults to 0,0,0 */
    }

    /**
     * スウォッチを塗りにした長方形を描く
     * @param {Layer} targetLayer - 描く先のレイヤー
     * @param {number} left - 左端の X 座標
     * @param {number} top - 上端の Y 座標
     * @param {SquareLayout} squareLayout - 長方形の大きさと線
     * @param {Swatch|SwatchEntry} fillSwatch - 適用するスウォッチ（color を読む。null ならプレビュー用に塗らない）
     * @returns {PathItem} 描いた長方形
     */
    function drawSwatchSquare(targetLayer, left, top, squareLayout, fillSwatch) {
        var square = targetLayer.pathItems.rectangle(top, left, squareLayout.width, squareLayout.height);
        var isOutlineOnly = (fillSwatch.color === null);
        square.filled = !isOutlineOnly;
        if (!isOutlineOnly) square.fillColor = fillSwatch.color;
        /* 塗らないプレビューで線も無いと見えないので、細い線を付ける / A fill-less preview needs a hairline when there is no stroke */
        var strokeWidth = (isOutlineOnly && squareLayout.strokeWidth <= 0) ? PREVIEW_HAIRLINE_WIDTH : squareLayout.strokeWidth;
        square.stroked = strokeWidth > 0;
        if (square.stroked) {
            square.strokeColor = squareLayout.strokeColor;
            square.strokeWidth = strokeWidth;
        }
        return square;
    }

    /**
     * 塗らずに描くための見本を作る（名前だけを持ち、色は null）
     * @param {string[]} swatchNames - スウォッチ名
     * @returns {SwatchEntry[]} 色が null の見本
     */
    function toOutlineEntries(swatchNames) {
        var outlineEntries = [];
        for (var i = 0; i < swatchNames.length; i++) outlineEntries.push({ name: swatchNames[i], color: null });
        return outlineEntries;
    }

    /**
     * スウォッチの名前を並べる
     * @param {Array<Swatch|SwatchEntry>} swatchList - スウォッチ
     * @returns {string[]} 名前
     */
    function getSwatchNames(swatchList) {
        var swatchNames = [];
        for (var i = 0; i < swatchList.length; i++) swatchNames.push(swatchList[i].name);
        return swatchNames;
    }

    /**
     * ライブラリを開いてスウォッチ名だけを読み、閉じる（取り込まない）
     * @param {Document} targetDoc - 読み終えたあとに戻るドキュメント
     * @param {File} libraryFile - スウォッチライブラリのファイル
     * @param {string} colorTypename - 色の typename
     * @returns {string[]} スウォッチ名
     */
    function readLibrarySwatchNames(targetDoc, libraryFile, colorTypename) {
        var openedLibrary = openLibraryDocument(libraryFile);
        /* 読めなくても開いたライブラリは閉じる / Close the library even when reading fails */
        try {
            return getSwatchNames(filterSwatchesByColorType(openedLibrary.doc.swatches, colorTypename));
        } finally {
            closeLibraryDocument(openedLibrary, targetDoc);
        }
    }

    /**
     * スウォッチごとに長方形を描き、SQUARES_PER_ROW 個ずつ折り返して並べる（左上を原点に置く。位置はあとで動かす）。
     * スウォッチ名を入れるときは、各長方形の上に左端をそろえて置き、その分だけ行の間を広げる
     * @param {Layer} targetLayer - 描く先のレイヤー（スウォッチを持つドキュメントのもの）
     * @param {Array<Swatch|SwatchEntry>} fillSwatches - 適用するスウォッチ（name と color を読む）
     * @param {SquareLayout} squareLayout - 長方形の大きさと間隔
     * @returns {PageItem[]} 描いた長方形（先頭が左上）と、そのあとにスウォッチ名
     */
    function drawSwatchSquareGrid(targetLayer, fillSwatches, squareLayout) {
        var nameSpace = getSwatchNameSpace(squareLayout);
        var pitchX = squareLayout.width + squareLayout.margin;
        var pitchY = nameSpace + squareLayout.height + squareLayout.margin;

        var squares = [];
        var swatchNames = [];
        for (var i = 0; i < fillSwatches.length; i++) {
            var squareLeft = (i % SQUARES_PER_ROW) * pitchX;
            var squareTop = -nameSpace - Math.floor(i / SQUARES_PER_ROW) * pitchY;
            squares.push(drawSwatchSquare(targetLayer, squareLeft, squareTop, squareLayout, fillSwatches[i]));
            if (squareLayout.showSwatchNames) {
                swatchNames.push(addPointTextAt(targetLayer, fillSwatches[i].name, SWATCH_NAME_FONT_SIZE, squareLeft, squareTop + SWATCH_NAME_GAP));
            }
        }
        return squares.concat(swatchNames);
    }

    /**
     * ドキュメントの指定した種類のスウォッチを、名前で引ける表にする
     * @param {Document} doc - 対象のドキュメント
     * @param {string} colorTypename - 色の typename
     * @returns {Object} { スウォッチ名: Swatch }
     */
    function mapSwatchesByName(doc, colorTypename) {
        var swatchesByName = {};
        var matchedSwatches = filterSwatchesByColorType(doc.swatches, colorTypename);
        for (var i = 0; i < matchedSwatches.length; i++) swatchesByName[matchedSwatches[i].name] = matchedSwatches[i];
        return swatchesByName;
    }

    /**
     * @typedef {Object} SwatchEntry
     * @property {string} name - スウォッチ名
     * @property {Color} color - 描く先のドキュメントで使える色
     */

    /**
     * ライブラリのスウォッチを、描く先のドキュメントに取り込む。
     * パターン・グラデーションは別のドキュメントの色として代入できないので、ライブラリ側で塗った長方形を複製して
     * スウォッチを持ち込み、複製の塗りの色を控えてから複製を消す。ライブラリは開いていなかったときだけ保存せずに閉じる。
     * 複製先は、作業中のレイヤーの状態（ロック・テンプレートなど）に左右されないよう、一時的に足したレイヤーにする
     * @param {Document} targetDoc - 描く先のドキュメント
     * @param {File} libraryFile - スウォッチライブラリのファイル
     * @param {string} colorTypename - 使うスウォッチの色の typename
     * @returns {SwatchEntry[]} 取り込んだスウォッチ
     */
    function importLibrarySwatchEntries(targetDoc, libraryFile, colorTypename) {
        var importLayer = targetDoc.layers.add();
        var openedLibrary = openLibraryDocument(libraryFile);
        var libraryDoc = openedLibrary.doc;
        var swatchEntries = [];
        /* 途中で失敗しても、開いたライブラリを閉じて元のドキュメントに戻す / Close the library and return to the document even on failure */
        try {
            var librarySwatches = filterSwatchesByColorType(libraryDoc.swatches, colorTypename);
            if (libraryDoc === targetDoc) {
                /* 作業中のドキュメントがそのライブラリなら、そのまま使う / Use the swatches as is when the document is the library itself */
                swatchEntries = librarySwatches;
            } else {
                /* 同じ名前・種類のスウォッチが作業中のドキュメントにあれば、複製せずにそれを使う。
                   名前が重なるものを複製すると、スウォッチの競合でレイヤーを変更できないエラーになる
                   Reuse a swatch of the same name and kind in the document; duplicating one that conflicts fails with "Target layer cannot be modified" */
                var existingSwatches = mapSwatchesByName(targetDoc, colorTypename);
                var workLayer = libraryDoc.layers.add();
                for (var i = 0; i < librarySwatches.length; i++) {
                    var swatchName = librarySwatches[i].name;
                    if (existingSwatches[swatchName]) {
                        swatchEntries.push({ name: swatchName, color: existingSwatches[swatchName].color });
                        continue;
                    }
                    var carrierSquare = workLayer.pathItems.rectangle(0, 0, 10, 10);
                    carrierSquare.fillColor = librarySwatches[i].color;
                    var importedSquare = carrierSquare.duplicate(importLayer, ElementPlacement.PLACEATEND);
                    swatchEntries.push({ name: swatchName, color: importedSquare.fillColor });
                    importedSquare.remove();
                }
                if (openedLibrary.wasAlreadyOpen) workLayer.remove();
            }
        } finally {
            closeLibraryDocument(openedLibrary, targetDoc);
            importLayer.remove();
        }
        return swatchEntries;
    }

    /**
     * 左上の長方形（スウォッチ名を入れるときはその名前）の上に、左端をそろえてタイトルを置く
     * @param {Layer} targetLayer - 置く先のレイヤー
     * @param {PathItem} topLeftSquare - 左上の長方形
     * @param {string} titleText - タイトルの文字列
     * @param {SquareLayout} squareLayout - 長方形の設定
     * @returns {TextFrame} 置いたタイトル
     */
    function addGridTitle(targetLayer, topLeftSquare, titleText, squareLayout) {
        var squareBounds = topLeftSquare.geometricBounds;
        var titleBottom = squareBounds[1] + getSwatchNameSpace(squareLayout) + TITLE_GAP;
        return addPointTextAt(targetLayer, titleText, TITLE_FONT_SIZE, squareBounds[0], titleBottom);
    }

    /**
     * オブジェクトをまとめて囲む範囲を返す
     * @param {PageItem[]} pageItems - 対象のオブジェクト
     * @returns {number[]} [左, 上, 右, 下]
     */
    function getItemsBounds(pageItems) {
        var unionBounds = pageItems[0].geometricBounds.slice(0);
        for (var i = 1; i < pageItems.length; i++) {
            var itemBounds = pageItems[i].geometricBounds;
            if (itemBounds[0] < unionBounds[0]) unionBounds[0] = itemBounds[0];
            if (itemBounds[1] > unionBounds[1]) unionBounds[1] = itemBounds[1];
            if (itemBounds[2] > unionBounds[2]) unionBounds[2] = itemBounds[2];
            if (itemBounds[3] < unionBounds[3]) unionBounds[3] = itemBounds[3];
        }
        return unionBounds;
    }

    /**
     * オブジェクトをまとめて動かす
     * @param {PageItem[]} pageItems - 対象のオブジェクト
     * @param {number} dx - 横の移動量
     * @param {number} dy - 縦の移動量
     * @returns {void}
     */
    function translateItems(pageItems, dx, dy) {
        for (var i = 0; i < pageItems.length; i++) pageItems[i].translate(dx, dy);
    }

    /**
     * 1ブロック（長方形・スウォッチ名・タイトル）の左上を指定の位置へ動かし、下端を返す
     * @param {PageItem[]} blockItems - ブロックのオブジェクト
     * @param {number} blockTop - 置く上端
     * @returns {number} 置いたあとの下端
     */
    function placeBlockAt(blockItems, blockTop) {
        var blockBounds = getItemsBounds(blockItems);
        translateItems(blockItems, -blockBounds[0], blockTop - blockBounds[1]);
        return blockTop - (blockBounds[1] - blockBounds[3]);
    }

    // =========================================
    // ダイアログの設定 / Dialog settings
    // =========================================

    /* 前回の設定（Illustrator 終了まで） / Last settings (kept until Illustrator quits) */
    var DIALOG_SETTINGS_KEY = "__" + SCRIPT_NAME + "_DialogSettings";
    /* プレビュー用に読んだライブラリのスウォッチ名（Illustrator 終了まで） / Library swatch names read for the preview (kept until Illustrator quits) */
    var LIBRARY_NAMES_CACHE_KEY = "__" + SCRIPT_NAME + "_LibrarySwatchNames";

    /**
     * @typedef {Object} DialogSettings
     * @property {string} kindKey - SWATCH_KINDS のキー
     * @property {boolean} useDocument - ドキュメントのスウォッチを使うなら true
     * @property {boolean} selectedSwatchesOnly - ドキュメントのとき、スウォッチパネルで選択しているものだけなら true
     * @property {Object} checkedLibraries - 種類ごとに、チェックを入れたライブラリの fsName（{ fsName: true }）
     * @property {number} width - 長方形の幅（pt）
     * @property {number} height - 長方形の高さ（pt）
     * @property {boolean} isSizeLinked - 幅と高さを連動するなら true
     * @property {number} margin - 長方形どうしの間隔（pt）
     * @property {number} strokeWidth - 線幅（pt）
     * @property {Color} strokeColor - 線の色
     * @property {boolean} addTitle - 各ブロックの左上の長方形の上にライブラリ名を入れるなら true
     * @property {boolean} showSwatchNames - 長方形の上にスウォッチ名を入れるなら true
     */

    /**
     * 前回の設定を返す。無い項目（古い版で控えたもの）は既定値で補う
     * @param {Document} doc - 線の色の既定値（黒）のカラーモードを決めるドキュメント
     * @returns {DialogSettings} 設定
     */
    function loadDialogSettings(doc) {
        var defaultSettings = {
            kindKey: SWATCH_KIND_KEYS[0],
            useDocument: false,
            selectedSwatchesOnly: true,
            checkedLibraries: {},
            width: DEFAULT_SQUARE_WIDTH,
            height: DEFAULT_SQUARE_HEIGHT,
            isSizeLinked: DEFAULT_SIZE_LINKED,
            margin: DEFAULT_SQUARE_MARGIN,
            strokeWidth: DEFAULT_STROKE_WIDTH,
            strokeColor: createBlackColor(doc),
            addTitle: DEFAULT_ADD_TITLE,
            showSwatchNames: DEFAULT_SHOW_SWATCH_NAMES
        };
        var settings = $.global[DIALOG_SETTINGS_KEY] || {};
        for (var settingKey in defaultSettings) {
            if (!defaultSettings.hasOwnProperty(settingKey)) continue;
            if (settings[settingKey] === undefined) settings[settingKey] = defaultSettings[settingKey];
        }
        return settings;
    }

    // =========================================
    // 色見本 / Color swatch
    // =========================================

    /**
     * 色を、色見本に塗る RGB（0〜1）に簡易換算する
     * @param {Color} sourceColor - RGBColor / CMYKColor / GrayColor
     * @returns {number[]} [r, g, b]（0〜1）
     */
    function colorToPreviewRgb(sourceColor) {
        switch (sourceColor.typename) {
            case "RGBColor":
                return [sourceColor.red / 255, sourceColor.green / 255, sourceColor.blue / 255];
            case "CMYKColor":
                var whiteRatio = 1 - sourceColor.black / 100;
                return [(1 - sourceColor.cyan / 100) * whiteRatio, (1 - sourceColor.magenta / 100) * whiteRatio, (1 - sourceColor.yellow / 100) * whiteRatio];
            case "GrayColor":
                var grayLevel = 1 - sourceColor.gray / 100;
                return [grayLevel, grayLevel, grayLevel];
        }
        return [0, 0, 0];
    }

    /**
     * クリックで標準のカラーピッカーを開く色見本を足す（group はクリックを受けないので image で作る）
     * @param {Group|Panel} parent - 追加先
     * @param {Color} initialColor - 初めの色
     * @param {Function} onColorChange - 色を選び直したあとに呼ぶ関数
     * @returns {Image} 色見本（.swatchColor で今の色を読む）
     */
    function addColorSwatch(parent, initialColor, onColorChange) {
        var colorSwatch = parent.add("image", undefined, undefined);
        colorSwatch.preferredSize = COLOR_SWATCH_SIZE;
        colorSwatch.swatchColor = initialColor;

        colorSwatch.onDraw = function () {
            var swatchGraphics = colorSwatch.graphics;
            var swatchWidth = colorSwatch.size[0];
            var swatchHeight = colorSwatch.size[1];
            /* newPath() を呼ばないとパスが積み重なる / Without newPath() the paths accumulate */
            swatchGraphics.newPath();
            swatchGraphics.rectPath(0, 0, swatchWidth, swatchHeight);
            swatchGraphics.fillPath(swatchGraphics.newBrush(swatchGraphics.BrushType.SOLID_COLOR, colorToPreviewRgb(colorSwatch.swatchColor).concat([1])));
            swatchGraphics.newPath();
            swatchGraphics.rectPath(0.5, 0.5, swatchWidth - 1, swatchHeight - 1);
            swatchGraphics.strokePath(swatchGraphics.newPen(swatchGraphics.PenType.SOLID_COLOR, [0.5, 0.5, 0.5, 1], 1));
        };

        colorSwatch.addEventListener("click", function () {
            /* キャンセルでも渡した色が返るので、戻り値をそのまま使う / Cancel returns the passed color, so use the result as is */
            colorSwatch.swatchColor = app.showColorPicker(colorSwatch.swatchColor);
            /* 隠して再表示して描き直させる / Hide and show to redraw */
            colorSwatch.hide();
            colorSwatch.show();
            onColorChange();
        });
        return colorSwatch;
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 単位付きの数値欄（項目名・∧∨・入力欄）を足す
     * @param {Group|Panel} parent - 追加先
     * @param {string} labelPath - 項目名の LABELS のパス
     * @param {number} valuePt - 初期値（pt）
     * @param {{label: string, pointsPerUnit: number}} fieldUnit - 欄の単位
     * @param {number} minValue - 下限（欄の単位）
     * @returns {EditText} 入力欄
     */
    function addUnitField(parent, labelPath, valuePt, fieldUnit, minValue) {
        return addSteppedField(parent, {
            label: labelText(labelPath),
            labelWidth: FIELD_LABEL_WIDTH,
            text: formatStepperNumber(Math.round(ptToUnit(valuePt, fieldUnit) * 1000) / 1000) + " " + fieldUnit.label,
            characters: FIELD_CHARACTERS,
            step: 1,
            min: minValue,
            unit: " " + fieldUnit.label
        });
    }

    /**
     * 入力欄の値を pt で読む
     * @param {EditText} unitField - addUnitField() で作った入力欄
     * @param {{pointsPerUnit: number}} fieldUnit - 欄の単位
     * @returns {number} 値（pt）
     */
    function readUnitFieldPt(unitField, fieldUnit) {
        return unitToPt(parseFloat(unitField.text), fieldUnit);
    }

    /**
     * 区切り線（高さ DIVIDER_HEIGHT のパネル、幅いっぱい）を足す
     * @param {Group|Panel} parent - 追加先
     * @returns {Panel} 区切り線
     */
    function addDivider(parent) {
        var dividerPanel = parent.add("panel");
        dividerPanel.alignment = ["fill", "top"];
        dividerPanel.preferredSize.height = DIVIDER_HEIGHT;
        return dividerPanel;
    }

    /**
     * ラジオボタンの並びを足す（同じ group に入れて排他にする）
     * @param {Group|Panel} parent - 追加先
     * @param {string[]} labelPaths - 各ラジオボタンの LABELS のパス
     * @param {boolean} [isVertical] - 縦に並べるなら true（省略時は横）
     * @returns {RadioButton[]} 作ったラジオボタン
     */
    function addRadioRow(parent, labelPaths, isVertical) {
        var radioRow = parent.add("group");
        setupRow(radioRow);
        if (isVertical) {
            radioRow.orientation = "column";
            radioRow.alignChildren = ["left", "top"];
            radioRow.spacing = 6;
        }
        var radioButtons = [];
        for (var i = 0; i < labelPaths.length; i++) {
            radioButtons.push(radioRow.add("radiobutton", undefined, getLabel(labelPaths[i])));
        }
        return radioButtons;
    }

    /**
     * 種類のライブラリの一覧（チェックボックスと区切り線）を作る。前回チェックしたものにチェックを入れ、無ければ先頭だけにする
     * @param {Group} host - 重ねる先
     * @param {Array<File|null>} libraryFiles - ライブラリのファイル（切れ目は null）
     * @param {Object} [checkedNames] - 前回チェックした fsName（{ fsName: true }）
     * @param {Function} onCheckboxClick - クリックしたチェックボックスを受け取って呼ぶ関数
     * @returns {{group: Group, entries: Array<{file: File, checkbox: Checkbox}>}} 一覧
     */
    function addLibraryCheckboxList(host, libraryFiles, checkedNames, onCheckboxClick) {
        var kindGroup = host.add("group");
        kindGroup.orientation = "column";
        kindGroup.alignChildren = ["left", "top"];
        kindGroup.spacing = 4;
        var listEntries = [];
        var hasChecked = false;
        for (var i = 0; i < libraryFiles.length; i++) {
            if (libraryFiles[i] === null) {
                addDivider(kindGroup).preferredSize.width = LIBRARY_LIST_WIDTH;
                continue;
            }
            var entryCheckbox = kindGroup.add("checkbox", undefined, getLibraryDisplayName(libraryFiles[i]));
            entryCheckbox.preferredSize.width = LIBRARY_LIST_WIDTH;
            entryCheckbox.helpTip = getLabel("tooltip.libraryCheckbox");
            entryCheckbox.value = !!(checkedNames && checkedNames[libraryFiles[i].fsName]);
            entryCheckbox.onClick = function () { onCheckboxClick(this); };
            hasChecked = hasChecked || entryCheckbox.value;
            listEntries.push({ file: libraryFiles[i], checkbox: entryCheckbox });
        }
        if (!hasChecked && listEntries.length > 0) listEntries[0].checkbox.value = true;
        return { group: kindGroup, entries: listEntries };
    }

    /**
     * ［ライブラリ］パネル（右カラム）を作る。種類ごとのチェックボックスの一覧を最初にすべて作って重ねておき、
     * 種類を切り替えたら表示する一覧だけを入れ替える（表示後に足したコントロールは描かれないことがあるため）
     * @param {Group} parent - 追加先
     * @param {DialogSettings} settings - 前回の設定（種類ごとにチェックしたライブラリ）
     * @param {Function} onSettingsChange - チェックを変えたあとに呼ぶ関数
     * @returns {{showKind: Function, setEnabled: Function, getCheckedFiles: Function, getCheckedNamesByKind: Function, getEntryCount: Function}} 一覧を操作する関数
     */
    function buildLibraryListPanel(parent, settings, onSettingsChange) {
        var libraryPanel = parent.add("panel", undefined, getLabel("panel.libraryList"));
        setupPanel(libraryPanel, 6);
        libraryPanel.helpTip = getLabel("tooltip.libraryList");

        var kindListsHost = libraryPanel.add("group");
        kindListsHost.orientation = "stack";
        kindListsHost.alignChildren = ["left", "top"];

        /* 種類ごとの { group, entries: [{ file, checkbox }] } / per kind: the list group and its rows */
        var kindLists = {};
        for (var i = 0; i < SWATCH_KIND_KEYS.length; i++) {
            var kindKey = SWATCH_KIND_KEYS[i];
            kindLists[kindKey] = addLibraryCheckboxList(kindListsHost, getKindLibraryFiles(kindKey), settings.checkedLibraries[kindKey], handleCheckboxClick);
        }

        var countText = libraryPanel.add("statictext", undefined, "");
        countText.preferredSize.width = LIBRARY_LIST_WIDTH; /* あとで文言を入れるので幅を確保 / reserve the width for later text */

        var selectButtonRow = libraryPanel.add("group");
        setupRow(selectButtonRow, "left", BUTTON_ROW_SPACING);
        var selectAllButton = selectButtonRow.add("button", undefined, getLabel("button.selectAll"));
        var deselectAllButton = selectButtonRow.add("button", undefined, getLabel("button.deselectAll"));

        var currentKindKey = SWATCH_KIND_KEYS[0];

        /**
         * チェックボックスをクリックしたあとの処理。option＋クリックなら、クリックしたものだけにするか、すべてにチェックを入れる
         * （クリックしたものだけがオンだったときは、すべてオン）
         * @param {Checkbox} clickedCheckbox - クリックしたチェックボックス
         * @returns {void}
         */
        function handleCheckboxClick(clickedCheckbox) {
            if (ScriptUI.environment.keyboardState.altKey) {
                var listEntries = kindLists[currentKindKey].entries;
                var hasOtherChecked = false;
                for (var j = 0; j < listEntries.length; j++) {
                    if (listEntries[j].checkbox !== clickedCheckbox && listEntries[j].checkbox.value) hasOtherChecked = true;
                }
                for (var k = 0; k < listEntries.length; k++) {
                    listEntries[k].checkbox.value = !hasOtherChecked || listEntries[k].checkbox === clickedCheckbox;
                }
            }
            updateCountText();
            onSettingsChange();
        }

        /**
         * 表示中の種類で、チェック済みのライブラリを返す
         * @returns {File[]} チェック済みのファイル
         */
        function getCheckedFiles() {
            var checkedFiles = [];
            var listEntries = kindLists[currentKindKey].entries;
            for (var j = 0; j < listEntries.length; j++) {
                if (listEntries[j].checkbox.value) checkedFiles.push(listEntries[j].file);
            }
            return checkedFiles;
        }

        /**
         * 件数の表示を更新する
         * @returns {void}
         */
        function updateCountText() {
            countText.text = labelValueText("fieldLabel.count", getCheckedFiles().length + " / " + kindLists[currentKindKey].entries.length);
        }

        /**
         * 種類の一覧だけを表示する
         * @param {string} kindKey - SWATCH_KINDS のキー
         * @returns {void}
         */
        function showKind(kindKey) {
            currentKindKey = kindKey;
            for (var listKey in kindLists) {
                if (kindLists.hasOwnProperty(listKey)) kindLists[listKey].group.visible = (listKey === kindKey);
            }
            updateCountText();
        }

        /**
         * 表示中の種類の、すべての行のチェックをそろえる
         * @param {boolean} isChecked - チェックを入れるなら true
         * @returns {void}
         */
        function setAllChecked(isChecked) {
            var listEntries = kindLists[currentKindKey].entries;
            for (var j = 0; j < listEntries.length; j++) listEntries[j].checkbox.value = isChecked;
            updateCountText();
            onSettingsChange();
        }

        selectAllButton.onClick = function () { setAllChecked(true); };
        deselectAllButton.onClick = function () { setAllChecked(false); };

        return {
            showKind: showKind,
            setEnabled: function (isEnabled) { libraryPanel.enabled = isEnabled; },
            getCheckedFiles: getCheckedFiles,
            /* 種類ごとのチェック（表示していない種類も） / Checked libraries per kind, including hidden kinds */
            getCheckedNamesByKind: function () {
                var checkedNamesByKind = {};
                for (var listKey in kindLists) {
                    if (!kindLists.hasOwnProperty(listKey)) continue;
                    var checkedNames = {};
                    var listEntries = kindLists[listKey].entries;
                    for (var j = 0; j < listEntries.length; j++) {
                        if (listEntries[j].checkbox.value) checkedNames[listEntries[j].file.fsName] = true;
                    }
                    checkedNamesByKind[listKey] = checkedNames;
                }
                return checkedNamesByKind;
            },
            getEntryCount: function () { return kindLists[currentKindKey].entries.length; }
        };
    }

    /**
     * ［スウォッチ］パネル（種類・出どころ・ドキュメントの対象）を作る。種類を切り替えると右の一覧を入れ替える
     * @param {Group} parent - 追加先
     * @param {DialogSettings} settings - 前回の設定
     * @param {Object} libraryList - buildLibraryListPanel() の戻り値
     * @param {Function} onSettingsChange - 切り替えたあとに呼ぶ関数
     * @returns {{getKindKey: Function, isLibrarySource: Function, isSelectedOnly: Function}} 結果を読むための関数
     */
    function buildSwatchPanel(parent, settings, libraryList, onSettingsChange) {
        var swatchPanel = parent.add("panel", undefined, getLabel("panel.swatch"));
        setupPanel(swatchPanel, 6);

        /* 出どころ・ドキュメントの対象を上に、種類をその下に / Source and document scope on top, kinds below */
        var sourceRadios = addRadioRow(swatchPanel, ["radio.library", "radio.document"]);
        var libraryRadio = sourceRadios[0];
        var documentRadio = sourceRadios[1];
        libraryRadio.helpTip = getLabel("tooltip.library");
        documentRadio.helpTip = getLabel("tooltip.documentSource");
        libraryRadio.value = !settings.useDocument;
        documentRadio.value = settings.useDocument;

        addDivider(swatchPanel);
        var scopeRadios = addRadioRow(swatchPanel, ["radio.selectedSwatches", "radio.allSwatches"], true);
        var selectedSwatchesRadio = scopeRadios[0];
        selectedSwatchesRadio.helpTip = getLabel("tooltip.selectedSwatches");
        selectedSwatchesRadio.value = settings.selectedSwatchesOnly;
        scopeRadios[1].value = !settings.selectedSwatchesOnly;

        addDivider(swatchPanel);
        var kindLabelPaths = [];
        for (var i = 0; i < SWATCH_KIND_KEYS.length; i++) kindLabelPaths.push(SWATCH_KINDS[SWATCH_KIND_KEYS[i]].radioLabel);
        var kindRadios = addRadioRow(swatchPanel, kindLabelPaths, true);

        var currentKindIndex = Math.max(0, indexOfKindKey(settings.kindKey));

        /**
         * 今の種類のキーを返す
         * @returns {string} SWATCH_KINDS のキー
         */
        function getKindKey() {
            return SWATCH_KIND_KEYS[currentKindIndex];
        }

        /**
         * ［ライブラリ］のときは種類と右の一覧、［ドキュメント］のときは対象のラジオボタンを有効にする
         * （ドキュメントは種類を問わないので、種類はディムにする）
         * @returns {void}
         */
        function updateSourceControls() {
            libraryList.setEnabled(libraryRadio.value);
            kindRadios[0].parent.enabled = libraryRadio.value;
            selectedSwatchesRadio.parent.enabled = documentRadio.value;
        }

        /**
         * 種類を切り替え、右の一覧を入れ替える
         * @param {number} kindIndex - SWATCH_KIND_KEYS の位置
         * @returns {void}
         */
        function selectKind(kindIndex) {
            currentKindIndex = kindIndex;
            kindRadios[kindIndex].value = true;
            var kindKey = getKindKey();
            libraryList.showKind(kindKey);
        }

        for (var j = 0; j < kindRadios.length; j++) {
            kindRadios[j].onClick = (function (kindIndex) {
                return function () {
                    selectKind(kindIndex);
                    onSettingsChange();
                };
            })(j);
        }
        libraryRadio.onClick = documentRadio.onClick = function () {
            updateSourceControls();
            onSettingsChange();
        };
        scopeRadios[0].onClick = scopeRadios[1].onClick = onSettingsChange;
        selectKind(currentKindIndex);
        updateSourceControls();

        return {
            getKindKey: getKindKey,
            isLibrarySource: function () { return libraryRadio.value; },
            isSelectedOnly: function () { return selectedSwatchesRadio.value; }
        };
    }

    /**
     * SWATCH_KIND_KEYS の中での種類の位置を返す
     * @param {string} kindKey - SWATCH_KINDS のキー
     * @returns {number} 位置（無ければ -1）
     */
    function indexOfKindKey(kindKey) {
        for (var i = 0; i < SWATCH_KIND_KEYS.length; i++) {
            if (SWATCH_KIND_KEYS[i] === kindKey) return i;
        }
        return -1;
    }

    /**
     * ［長方形］パネル（幅・高さ・連動・間隔）を作る
     * @param {Group} parent - 追加先
     * @param {DialogSettings} settings - 前回の設定
     * @param {{label: string, pointsPerUnit: number}} rulerUnit - 定規の単位
     * @param {Function} onSettingsChange - 値を変えたあとに呼ぶ関数
     * @returns {{widthField: EditText, heightField: EditText, marginField: EditText, linkToggle: Group}} 入力欄と連動アイコン
     */
    function buildSquarePanel(parent, settings, rulerUnit, onSettingsChange) {
        var squarePanel = parent.add("panel", undefined, getLabel("panel.square"));
        setupPanel(squarePanel, 6);

        /* 幅・高さを縦に積み、右に連動アイコンを置く / Stack width and height with the link toggle on the right */
        var sizeRow = squarePanel.add("group");
        setupRow(sizeRow, "left", 6);
        var sizeFieldsColumn = sizeRow.add("group");
        sizeFieldsColumn.orientation = "column";
        sizeFieldsColumn.alignChildren = ["left", "center"];
        sizeFieldsColumn.spacing = 6;
        var widthField = addUnitField(sizeFieldsColumn, "fieldLabel.width", settings.width, rulerUnit, 0.01);
        var heightField = addUnitField(sizeFieldsColumn, "fieldLabel.height", settings.height, rulerUnit, 0.01);

        /**
         * 連動中は、変えた側の値をもう一方へ写す
         * @param {EditText} sourceField - 変えた側の入力欄
         * @param {EditText} targetField - 写す先の入力欄
         * @returns {void}
         */
        function copyLinkedSize(sourceField, targetField) {
            if (!linkToggle.value) return;
            targetField.text = sourceField.text;
            targetField.lastValidText = sourceField.text;
        }
        var linkToggle = addLinkToggle(sizeRow, settings.isSizeLinked, function () {
            copyLinkedSize(widthField, heightField);
            onSettingsChange();
        });
        linkToggle.helpTip = getLabel("tooltip.linkToggle");
        linkSizeFields(widthField, heightField, copyLinkedSize);
        linkSizeFields(heightField, widthField, copyLinkedSize);

        var marginField = addUnitField(squarePanel, "fieldLabel.margin", settings.margin, rulerUnit, 0);
        marginField.helpTip = getLabel("tooltip.margin");
        watchUnitField(widthField, onSettingsChange);
        watchUnitField(heightField, onSettingsChange);
        watchUnitField(marginField, onSettingsChange);

        return { widthField: widthField, heightField: heightField, marginField: marginField, linkToggle: linkToggle };
    }

    /**
     * 入力欄の確定（直接入力）と∧∨・↑↓キーの増減のあとに、もう一方の欄へ写す処理を足す
     * @param {EditText} sourceField - 変える側の入力欄
     * @param {EditText} targetField - 写す先の入力欄
     * @param {Function} copyLinkedSize - (sourceField, targetField) を受け取って写す関数
     * @returns {void}
     */
    function linkSizeFields(sourceField, targetField, copyLinkedSize) {
        var normalizeOnChange = sourceField.onChange;
        sourceField.onChange = function () {
            normalizeOnChange();
            copyLinkedSize(sourceField, targetField);
        };
        sourceField.stepperGroup.stepOptions.onStep = function () { copyLinkedSize(sourceField, targetField); };
    }

    /**
     * 数値欄の確定（直接入力）と∧∨・↑↓キーの増減のあとに呼ぶ処理を足す（既存の処理のあとに呼ぶ）
     * @param {EditText} unitField - addUnitField() で作った入力欄
     * @param {Function} onValueChange - 値が変わったあとに呼ぶ関数
     * @returns {void}
     */
    function watchUnitField(unitField, onValueChange) {
        var previousOnChange = unitField.onChange;
        unitField.onChange = function () {
            previousOnChange();
            onValueChange();
        };
        var stepOptions = unitField.stepperGroup.stepOptions;
        var previousOnStep = stepOptions.onStep;
        stepOptions.onStep = function (numberInput) {
            if (previousOnStep) previousOnStep(numberInput);
            onValueChange();
        };
    }

    /**
     * ［線］パネル（線幅・カラー）を作る
     * @param {Group} parent - 追加先
     * @param {DialogSettings} settings - 前回の設定
     * @param {{label: string, pointsPerUnit: number}} strokeUnit - 線の単位
     * @param {Function} onSettingsChange - 値を変えたあとに呼ぶ関数
     * @returns {{strokeWidthField: EditText, strokeColorSwatch: Image}} 入力欄と色見本
     */
    function buildStrokePanel(parent, settings, strokeUnit, onSettingsChange) {
        var strokePanel = parent.add("panel", undefined, getLabel("panel.stroke"));
        setupPanel(strokePanel, 6);
        var strokeWidthField = addUnitField(strokePanel, "fieldLabel.strokeWidth", settings.strokeWidth, strokeUnit, 0);
        strokeWidthField.helpTip = getLabel("tooltip.strokeWidth");
        watchUnitField(strokeWidthField, onSettingsChange);

        var strokeColorRow = strokePanel.add("group");
        setupRow(strokeColorRow, "left", STEPPER_FIELD_SPACING);
        var strokeColorLabel = strokeColorRow.add("statictext", undefined, labelText("fieldLabel.strokeColor"));
        strokeColorLabel.preferredSize.width = FIELD_LABEL_WIDTH;
        strokeColorLabel.justify = "right";
        var strokeColorSwatch = addColorSwatch(strokeColorRow, settings.strokeColor, onSettingsChange);
        strokeColorSwatch.helpTip = getLabel("tooltip.strokeColor");

        return { strokeWidthField: strokeWidthField, strokeColorSwatch: strokeColorSwatch };
    }

    /**
     * ［テキストを追加］パネル（ライブラリ名・スウォッチ名）を作る
     * @param {Group} parent - 追加先
     * @param {DialogSettings} settings - 前回の設定
     * @param {Function} onSettingsChange - チェックを変えたあとに呼ぶ関数
     * @returns {{libraryNameCheckbox: Checkbox, swatchNameCheckbox: Checkbox}} チェックボックス
     */
    function buildTextPanel(parent, settings, onSettingsChange) {
        var textPanel = parent.add("panel", undefined, getLabel("panel.text"));
        setupPanel(textPanel, 6);
        var libraryNameCheckbox = textPanel.add("checkbox", undefined, getLabel("checkbox.libraryName"));
        libraryNameCheckbox.value = settings.addTitle;
        libraryNameCheckbox.helpTip = getLabel("tooltip.libraryName");
        libraryNameCheckbox.onClick = onSettingsChange;
        var swatchNameCheckbox = textPanel.add("checkbox", undefined, getLabel("checkbox.swatchName"));
        swatchNameCheckbox.value = settings.showSwatchNames;
        swatchNameCheckbox.helpTip = getLabel("tooltip.swatchName");
        swatchNameCheckbox.onClick = onSettingsChange;
        return { libraryNameCheckbox: libraryNameCheckbox, swatchNameCheckbox: swatchNameCheckbox };
    }

    /**
     * ダイアログ下部のボタン行（左に［プレビュー（長方形のみ）］、右に［キャンセル］［OK］）を作る
     * @param {Window} dialog - ダイアログ
     * @param {Object} sampleSession - createSampleSession() の戻り値
     * @param {Function} updatePreview - プレビューを描き直す関数
     * @param {Function} canClose - ［OK］で閉じてよいかを返す関数
     * @returns {Checkbox} ［プレビュー（長方形のみ）］のチェックボックス
     */
    function buildDialogButtonRow(dialog, sampleSession, updatePreview, canClose) {
        var buttonRow = addButtonRow(dialog);
        var previewCheckbox = buttonRow.leftGroup.add("checkbox", undefined, getLabel("checkbox.preview"));
        previewCheckbox.helpTip = getLabel("tooltip.preview");
        previewCheckbox.value = DEFAULT_PREVIEW;
        previewCheckbox.onClick = function () {
            if (previewCheckbox.value) {
                updatePreview();
                return;
            }
            sampleSession.clear();
            app.redraw();
        };
        buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var okButton = buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
        okButton.onClick = function () {
            if (canClose()) dialog.close(1);
        };
        alignRightOnlyButtonRow(buttonRow);
        return previewCheckbox;
    }

    /**
     * 2カラムの片側の group を作る
     * @param {Group} columnsGroup - 左右を並べる group
     * @returns {Group} 縦に積む group
     */
    function addColumn(columnsGroup) {
        var columnGroup = columnsGroup.add("group");
        columnGroup.orientation = "column";
        columnGroup.alignChildren = ["fill", "top"];
        columnGroup.spacing = WINDOW_SPACING;
        return columnGroup;
    }

    // =========================================
    // 見本の作成とプレビュー / Drawing samples and the preview
    // =========================================

    /**
     * @typedef {Object} SampleRequest
     * @property {DialogSettings} settings - 設定
     * @property {File[]} libraryFiles - チェックを入れたライブラリ（ドキュメントなら空）
     */

    /**
     * 見本のブロック（長方形・スウォッチ名・ライブラリ名）を描き、ブロックを縦に並べてアートボードの中央に置く。
     * アートボードと描く位置を同じ座標系で扱うため、描く間だけドキュメント座標にする
     * @param {Document} doc - 対象のドキュメント
     * @param {Layer} targetLayer - 描く先のレイヤー
     * @param {SampleRequest} sampleRequest - 描く内容
     * @param {Function} getLibraryEntries - (libraryFile, colorTypename) を受け取り SwatchEntry[] を返す関数
     * @param {boolean} isOutlineOnly - プレビュー用に塗らずに描くなら true
     * @param {number} artboardIndex - 中央に置くアートボードの番号
     * @returns {PageItem[]} 描いたオブジェクト（スウォッチが無ければ空）
     */
    function drawSampleBlocks(doc, targetLayer, sampleRequest, getLibraryEntries, isOutlineOnly, artboardIndex) {
        var settings = sampleRequest.settings;
        var squareLayout = {
            width: settings.width,
            height: settings.height,
            margin: settings.margin,
            strokeWidth: settings.strokeWidth,
            strokeColor: settings.strokeColor,
            showSwatchNames: settings.showSwatchNames
        };

        var createdItems = [];
        var blockTop = 0;

        /**
         * 1ブロックを描いて足し、次のブロックの上端を進める
         * @param {Array<Swatch|SwatchEntry>} fillSwatches - 塗るスウォッチ
         * @param {string} titleText - ライブラリ名（ドキュメント名）
         * @returns {void}
         */
        function appendBlock(fillSwatches, titleText) {
            if (fillSwatches.length === 0) return;
            var blockItems = drawSwatchSquareGrid(targetLayer, fillSwatches, squareLayout);
            if (settings.addTitle) blockItems.push(addGridTitle(targetLayer, blockItems[0], titleText, squareLayout));
            /* ブロックの間は長方形の間隔の2倍 / Blocks are spaced twice the spacing */
            blockTop = placeBlockAt(blockItems, blockTop) - squareLayout.margin * 2;
            createdItems = createdItems.concat(blockItems);
        }

        var previousCoordinateSystem = app.coordinateSystem;
        app.coordinateSystem = CoordinateSystem.DOCUMENTCOORDINATESYSTEM;
        /* 途中で失敗しても座標系を戻す / Restore the coordinate system even on failure */
        try {
            if (settings.useDocument) {
                var documentSwatches = getDocumentSwatches(doc, null, settings.selectedSwatchesOnly);
                if (isOutlineOnly) documentSwatches = toOutlineEntries(getSwatchNames(documentSwatches));
                appendBlock(documentSwatches, doc.name.replace(/\.[^.]+$/, ""));
            } else {
                var colorTypename = SWATCH_KINDS[settings.kindKey].colorTypename;
                for (var i = 0; i < sampleRequest.libraryFiles.length; i++) {
                    var libraryFile = sampleRequest.libraryFiles[i];
                    appendBlock(getLibraryEntries(libraryFile, colorTypename), getLibraryDisplayName(libraryFile));
                }
            }
            if (createdItems.length === 0) return createdItems;

            /* 全体をアートボードの中央へ。ライブラリを開くと表示が変わるので、画面の中心は使わない
               Center everything on the artboard; the view changes when a library is opened, so the view center is not used */
            var artboardRect = doc.artboards[artboardIndex].artboardRect;
            var allBounds = getItemsBounds(createdItems);
            translateItems(createdItems, (artboardRect[0] + artboardRect[2]) / 2 - (allBounds[0] + allBounds[2]) / 2, (artboardRect[1] + artboardRect[3]) / 2 - (allBounds[1] + allBounds[3]) / 2);
            return createdItems;
        } finally {
            app.coordinateSystem = previousCoordinateSystem;
        }
    }

    /**
     * 見本を描く作業をまとめる（プレビューの描き直しと、取り消したときの後始末）。
     * プレビューはライブラリのスウォッチ名だけを読んで長方形を描き、確定したときにだけスウォッチを取り込んで塗る
     * @param {Document} doc - 対象のドキュメント
     * @returns {{render: Function, clear: Function, cancel: Function}} 描く・消す・取り消す関数
     */
    function createSampleSession(doc) {
        var layerCountBefore = doc.layers.length;
        var artboardIndex = doc.artboards.getActiveArtboardIndex(); /* 実行した時点の作業中のアートボード / the active artboard when run */
        var targetLayer = null;
        var drawnItems = [];
        var importsByKey = {};   /* ライブラリごとに取り込んだスウォッチ / swatches imported per library */

        /**
         * 描く先のレイヤーを返す（初めて描くときに決める）
         * @returns {Layer} 描く先のレイヤー
         */
        function getTargetLayer() {
            if (!targetLayer) targetLayer = getDrawableLayer(doc);
            return targetLayer;
        }

        /**
         * ライブラリのスウォッチを返す（初めてのときだけ取り込む）
         * @param {File} libraryFile - スウォッチライブラリのファイル
         * @param {string} colorTypename - 色の typename
         * @returns {SwatchEntry[]} スウォッチ
         */
        function getLibraryEntries(libraryFile, colorTypename) {
            var importKey = libraryFile.fsName + "|" + colorTypename;
            if (!importsByKey[importKey]) importsByKey[importKey] = importLibrarySwatchEntries(doc, libraryFile, colorTypename);
            return importsByKey[importKey];
        }

        /**
         * プレビュー用に、ライブラリのスウォッチ名だけを返す（取り込まない）。
         * 読んだ名前は Illustrator の終了まで $.global に控え、次の実行からはライブラリを開かない
         * @param {File} libraryFile - スウォッチライブラリのファイル
         * @param {string} colorTypename - 色の typename
         * @returns {SwatchEntry[]} 色が null の見本
         */
        function getLibraryOutlineEntries(libraryFile, colorTypename) {
            var namesCache = $.global[LIBRARY_NAMES_CACHE_KEY];
            if (!namesCache) namesCache = $.global[LIBRARY_NAMES_CACHE_KEY] = {};
            /* 更新日時を含めて、ライブラリが更新されたら読み直す / Include the modified time so an updated library is read again */
            var modifiedTime = libraryFile.modified ? libraryFile.modified.getTime() : 0;
            var namesKey = libraryFile.fsName + "|" + modifiedTime + "|" + colorTypename;
            if (!namesCache[namesKey]) namesCache[namesKey] = readLibrarySwatchNames(doc, libraryFile, colorTypename);
            return toOutlineEntries(namesCache[namesKey]);
        }

        /**
         * 描いたものを消す
         * @returns {void}
         */
        function clear() {
            for (var i = 0; i < drawnItems.length; i++) drawnItems[i].remove();
            drawnItems = [];
        }

        /**
         * 描いたものを消してから描き直す。プレビューは長方形だけ（スウォッチを取り込まず、塗らない）
         * @param {SampleRequest} sampleRequest - 描く内容
         * @param {boolean} isPreview - プレビューなら true
         * @returns {PageItem[]} 描いたオブジェクト
         */
        function render(sampleRequest, isPreview) {
            clear();
            var getEntries = isPreview ? getLibraryOutlineEntries : getLibraryEntries;
            drawnItems = drawSampleBlocks(doc, getTargetLayer(), sampleRequest, getEntries, isPreview, artboardIndex);
            return drawnItems;
        }

        /**
         * ［キャンセル］で閉じたときの後始末。プレビューを消し、ロック中のために足したレイヤーが空なら消す
         * （プレビューはスウォッチを取り込まないので、スウォッチパネルは変わっていない）
         * @returns {void}
         */
        function cancel() {
            clear();
            if (targetLayer && doc.layers.length > layerCountBefore && targetLayer.pageItems.length === 0) targetLayer.remove();
        }

        return { render: render, clear: clear, cancel: cancel };
    }

    // =========================================
    // ダイアログ（本体） / Dialog (main)
    // =========================================

    /**
     * 設定のダイアログを表示する（左に設定、右にライブラリの一覧）。［プレビュー］がオンなら、変えるたびに描き直す
     * @param {Document} doc - 対象のドキュメント
     * @param {Object} sampleSession - createSampleSession() の戻り値
     * @returns {SampleRequest|null} 選んだ内容。キャンセルなら null
     */
    function showSettingsDialog(doc, sampleSession) {
        var settings = loadDialogSettings(doc);
        var rulerUnit = getUnitInfo();
        var strokeUnit = getUnitInfo("strokeUnits");
        var previewCheckbox = null; /* 作り終えるまでは描き直さない / no redraw until the dialog is built */

        /**
         * ［プレビュー］がオンなら描き直す
         * @returns {void}
         */
        function updatePreview() {
            if (!previewCheckbox || !previewCheckbox.value) return;
            sampleSession.render(collectSampleRequest(), true);
            app.redraw();
        }

        var dialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        setupWindow(dialog);

        var columnsGroup = dialog.add("group");
        columnsGroup.orientation = "row";
        columnsGroup.alignChildren = ["fill", "top"];
        columnsGroup.spacing = COLUMN_SPACING;
        var leftColumn = addColumn(columnsGroup);
        var rightColumn = addColumn(columnsGroup);

        /* 一覧は先に作る（左の種類の切り替えで入れ替えるため） / Build the list first so the kind radios can switch it */
        var libraryList = buildLibraryListPanel(rightColumn, settings, updatePreview);
        var swatchControls = buildSwatchPanel(leftColumn, settings, libraryList, updatePreview);
        var squareControls = buildSquarePanel(leftColumn, settings, rulerUnit, updatePreview);
        var strokeControls = buildStrokePanel(leftColumn, settings, strokeUnit, updatePreview);
        var textControls = buildTextPanel(leftColumn, settings, updatePreview);

        /**
         * 今のダイアログの内容を読む（控えには書かない）
         * @returns {SampleRequest} 描く内容
         */
        function collectSampleRequest() {
            var kindKey = swatchControls.getKindKey();
            var currentSettings = {
                kindKey: kindKey,
                useDocument: !swatchControls.isLibrarySource(),
                selectedSwatchesOnly: swatchControls.isSelectedOnly(),
                checkedLibraries: settings.checkedLibraries,
                width: readUnitFieldPt(squareControls.widthField, rulerUnit),
                height: readUnitFieldPt(squareControls.heightField, rulerUnit),
                isSizeLinked: squareControls.linkToggle.value,
                margin: readUnitFieldPt(squareControls.marginField, rulerUnit),
                strokeWidth: readUnitFieldPt(strokeControls.strokeWidthField, strokeUnit),
                strokeColor: strokeControls.strokeColorSwatch.swatchColor,
                addTitle: textControls.libraryNameCheckbox.value,
                showSwatchNames: textControls.swatchNameCheckbox.value
            };
            return {
                settings: currentSettings,
                libraryFiles: currentSettings.useDocument ? [] : libraryList.getCheckedFiles()
            };
        }

        previewCheckbox = buildDialogButtonRow(dialog, sampleSession, updatePreview, function () {
            /* ライブラリのとき、一覧が空かチェックが無ければ閉じない / For libraries, stay open when the list is empty or nothing is checked */
            if (!swatchControls.isLibrarySource()) return true;
            if (libraryList.getEntryCount() === 0) {
                alert(getLabel("alert.noLibraryFolder"));
                return false;
            }
            if (libraryList.getCheckedFiles().length === 0) {
                alert(getLabel("alert.noCheckedLibrary"));
                return false;
            }
            return true;
        });

        /* 開いたときに最初のプレビューを描く（onShow は prepareDialogWindow の前に入れ、その中から呼ばせる）
           Draw the first preview when the dialog opens (set before prepareDialogWindow, which chains it) */
        dialog.onShow = updatePreview;
        prepareDialogWindow(dialog, SCRIPT_NAME);
        if (dialog.show() !== 1) return null;

        var sampleRequest = collectSampleRequest();
        sampleRequest.settings.checkedLibraries = libraryList.getCheckedNamesByKind();
        $.global[DIALOG_SETTINGS_KEY] = sampleRequest.settings;
        return sampleRequest;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 何も描けなかったときの警告の LABELS のパスを返す
     * @param {DialogSettings} settings - 設定
     * @returns {string} 警告のパス
     */
    function getEmptyAlertPath(settings) {
        if (!settings.useDocument) return SWATCH_KINDS[settings.kindKey].emptyAlert;
        return settings.selectedSwatchesOnly ? "alert.noSelectedSwatch" : "alert.noDocumentSwatch";
    }

    /**
     * エントリポイント
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }

        var doc = app.activeDocument;
        var sampleSession = createSampleSession(doc);
        var sampleRequest = showSettingsDialog(doc, sampleSession);
        if (sampleRequest === null) {
            sampleSession.cancel();
            app.redraw();
            return;
        }

        /* プレビューの有無にかかわらず、確定した内容で描き直す / Redraw with the final settings whether or not previewing */
        var createdItems = sampleSession.render(sampleRequest, false);
        if (createdItems.length === 0) {
            alert(getLabel(getEmptyAlertPath(sampleRequest.settings)));
            return;
        }
        doc.selection = null;
        doc.selection = createdItems;
    }

    main();
})();
