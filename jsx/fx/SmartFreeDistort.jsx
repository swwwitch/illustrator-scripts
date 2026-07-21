#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択オブジェクトに Illustrator の「自由変形」ライブ効果を適用するスクリプトです。
台形・平行四辺形・三角形・対角線の全18プリセットをアイコンから選択し、
調整可能なプリセットでは変形量と強度を設定できます。Undoベースのプレビューで
結果を確認してから、効果を各オブジェクトへ個別に適用します。

### 主な機能

- 調整変形（プリセットはすべてアイコンで選択）
  - 台形（上辺を狭く / 広く、下辺を狭く / 広く）
  - 平行四辺形（4基準点 × 左右 / 上下 の8パターン）
  - 変形量スライダー（0.00〜0.49、ドラッグと↑↓キー操作に対応）
  - 強度切替（マイルド / ノーマル / ブースト）
- 固定変形
  - 三角形（左上 / 右上 / 左下 / 右下）
  - 対角線（＼ / ／）
- プレビュー機能（Undoベースで一時適用）
- 日英ローカライズ対応UI

### 仕様・注意

- 三角形および対角線プリセットでは、変形量と強度は使用されません。
- 台形 / 平行四辺形の強度倍率は、マイルド=0.25、ノーマル=0.5、ブースト=1.0 です。
- 選択対象にテキストが含まれる場合、強度の初期値はマイルドになります。
- プレビューは Undo を利用して制御しているため、他の操作と混在すると不整合が起こる可能性があります。
- 複数オブジェクト選択時は、すべてに同一のライブ効果を適用します。

### 紹介記事（note）

https://note.com/dtp_tranist/n/n15a7ae196a23

### GitHub

https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/fx/SmartFreeDistort.jsx

### 更新履歴

- v1.5.4 (2026-07-21)
  - 日本語・英語READMEを追加し、基本情報からリンク

- v1.5.3 (2026-07-21)
  - 日本語・英語の概要を、現在のプリセット数と適用フローに合わせて更新

- v1.5.2 (2026-07-21)
  - プレビュー解除のUndoを、目印レイヤーを使って過不足なく巻き戻すよう修正
    （複数オブジェクト選択時に効果が残留し、二重適用になる問題への対応）
  - 効果を適用できるオブジェクトがない場合、ダイアログを出す前に警告するよう変更
  - プレビューのUndoで選択が失われた場合に備え、初期選択をフォールバックとして保持
  - 本適用が途中で失敗した場合、何件まで適用したかを表示するよう変更
  - アイコンの平行四辺形サンプル値を、実際の適用値（変形量×2）と一致するよう修正

- v1.5.1 (2026-07-21)
  - 平行四辺形を、4基準点 × 左右 / 上下 の8パターンに拡張
  - 台形 / 平行四辺形 / 三角形 / 対角線のプリセットをすべてアイコンボタン化
  - 台形 / 平行四辺形 / 三角形・対角線 の3カラムレイアウトに変更
  - 変形量パネルをプリセットの下へ移動し、強度ラジオを左右中央に配置
  - ボタンエリアを「左：プレビュー／中央：スペーサー／右：キャンセル・OK」に整理

- v1.5.0 (2026-07-21)
  - プレビュー解除のUndoを、適用済みのときだけ1回実行するよう整理（戻りすぎの防止）
  - 変形量スライダーを整数（%）に丸め、表示値と実適用値のずれを解消
  - 効果を適用できない選択（テキスト範囲など）を除外
  - 適用中の例外でプレビューが残留しないよう巻き戻し処理を追加
  - プリセットをテーブル定義に集約し、インデックス依存の分岐を廃止
  - UIレイアウトの余白・間隔を共通設定に統一

- v1.1.1 (2026-04-24)
  - 台形 / 平行四辺形 / 三角形 / 対角線のプリセットUIを実装
  - 変形量スライダーとプレビューを実装
  - 強度切替（マイルド / ノーマル / ブースト）を追加
  - 選択対象にテキストが含まれる場合、強度初期値をマイルドに設定
  - プレビュー更新の間引きと同一条件スキップを追加
  - 日英ローカライズ対応UIを実装

*/

/*

### Overview

Applies Illustrator's "Free Distort" live effect to selected objects. Choose from
18 icon-based trapezoid, parallelogram, triangle and diagonal presets; adjustable
presets also support amount and strength controls. An Undo-based preview lets you
check the result before the effect is applied to each object individually.

### Key features

- Adjustable transform
  - Trapezoid (narrow / wide top, narrow / wide bottom)
  - Parallelogram (four anchor corners by two axes, eight presets)
  - Amount slider (0.00-0.49, supports dragging and arrow keys)
  - Strength switch (Mild / Normal / Boost)
- Fixed transform
  - Triangle (top left / top right / bottom left / bottom right)
  - Diagonal (backslash / slash)
- Preview based on Undo
- Localized UI (Japanese / English)

### Notes

- Amount and strength are ignored for the triangle and diagonal presets.
- Strength factors are Mild=0.25, Normal=0.5, Boost=1.0.
- Strength defaults to Mild when the selection contains text.
- Preview relies on Undo, so mixing it with other operations may cause inconsistencies.
- The same live effect is applied to every selected object.

### Article (note, in Japanese)

https://note.com/dtp_tranist/n/n15a7ae196a23

### GitHub

https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/fx/SmartFreeDistort.jsx

### Changelog

- v1.5.4 (2026-07-21)
  - Added Japanese and English READMEs and linked them from Basic info

- v1.5.3 (2026-07-21)
  - Updated the Japanese and English overviews to reflect the current preset count
    and application flow

- v1.5.2 (2026-07-21)
  - Preview removal now undoes exactly the applied steps, using a marker layer as the
    undo boundary (previously a leftover preview could be applied twice)
  - Warn about a selection with no effect target before the dialog opens
  - Keep the initial selection as a fallback in case an undo clears it
  - Report how many objects were done when the final apply stops partway
  - Icon sample shear now matches the amount actually applied (amount times two)

- v1.5.1 (2026-07-21)
  - Expanded the parallelogram presets to eight: four anchors by two axes
  - Turned every preset into an icon button drawn from its own corner data
  - Reorganized the dialog into three columns: trapezoid, parallelogram, triangle / diagonal
  - Moved the amount panel below the presets and centered the strength radios
  - Rearranged the footer: preview on the left, Cancel / OK on the right

- v1.5.0 (2026-07-21)
  - Preview removal now runs a single undo, and only when one is actually pending
  - Rounded the amount slider to integer percent so the readout matches what is applied
  - Skipped selected items that cannot receive an effect (text ranges and similar)
  - Rolled back partially applied previews when an error occurs
  - Moved presets into a single table and removed index-based branching
  - Unified panel margins and spacing through shared layout helpers

- v1.1.1 (2026-04-24)
  - Initial preset UI, amount slider, preview and strength switch

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SmartFreeDistort";             /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.5.4";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-04-24";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-07-21";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartFreeDistort.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartFreeDistort.md
var SCRIPT_URL      = "https://note.com/dtp_tranist/n/n15a7ae196a23"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

// =========================================
// ユーザー設定 / User settings
// =========================================
var DISTORT_CONFIG = {
    amountMinPercent: 0,       /* 変形量スライダーの最小値（%） / slider minimum (percent) */
    amountMaxPercent: 49,      /* 変形量スライダーの最大値（%） / slider maximum (percent) */
    amountDefaultPercent: 20,  /* 変形量スライダーの初期値（%） / slider default (percent) */
    amountSliderWidth: 220,    /* 変形量スライダーの幅（px） / amount slider width in pixels */
    previewThrottleMs: 150,    /* プレビュー更新の最小間隔（ミリ秒） / minimum preview interval in ms */
    shearBaseFactor: 2         /* シアー量の基準倍率 / base multiplier for the shear amount */
};

/* 強度倍率。台形・平行四辺形で共通（平行四辺形はさらに shearBaseFactor を掛ける）
   Strength factors shared by trapezoid and parallelogram (the latter also uses shearBaseFactor) */
var STRENGTH_FACTORS = {
    mild: 0.25,
    normal: 0.5,
    boost: 1.0
};

// =========================================
// ローカライズ / Localization
// =========================================

/* 実行環境のロケールから UI 言語を判定 / Detect UI language from the host locale */
function getCurrentLang() {
    return ($.locale && $.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

/* UI 文言の定義 / UI string definitions */
var LABELS = {
    dialog: {
        title: { ja: "スマート自由変形", en: "Smart Free Distort" }
    },
    alert: {
        noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
        noSelection: {
            ja: "オブジェクトを選択してください。",
            en: "Please select one or more objects."
        },
        noTarget: {
            ja: "効果を適用できるオブジェクトが選択されていません。",
            en: "The selection contains no object that can receive an effect."
        },
        partialApply: {
            ja: "{count}件に適用した時点で中断しました。取り消すには、その回数ぶん取り消してください。",
            en: "Stopped after applying to {count} object(s). Undo that many times to revert."
        }
    },
    panel: {
        trapezoid: { ja: "台形", en: "Trapezoid" },
        parallelogram: { ja: "平行四辺形", en: "Parallelogram" },
        triangle: { ja: "三角形", en: "Triangle" },
        diagonal: { ja: "対角線", en: "Diagonal" },
        amount: { ja: "変形量", en: "Amount" },
        strength: { ja: "強度", en: "Strength" }
    },
    trapezoid: {
        narrowTop: { ja: "上辺を狭く", en: "Narrow Top" },
        wideTop: { ja: "上辺を広く", en: "Wide Top" },
        narrowBottom: { ja: "下辺を狭く", en: "Narrow Bottom" },
        wideBottom: { ja: "下辺を広く", en: "Wide Bottom" }
    },
    shear: {
        axisHorizontal: { ja: "左右", en: "Horizontal" },
        axisVertical: { ja: "上下", en: "Vertical" },
        anchorTopLeft: { ja: "左上", en: "Top Left" },
        anchorTopRight: { ja: "右上", en: "Top Right" },
        anchorBottomLeft: { ja: "左下", en: "Bottom Left" },
        anchorBottomRight: { ja: "右下", en: "Bottom Right" },
        tipFormat: {
            ja: "{axis}（基準点：{anchor}）",
            en: "{axis} (anchor: {anchor})"
        }
    },
    triangle: {
        bottomLeft: { ja: "左下", en: "Bottom Left" },
        bottomRight: { ja: "右下", en: "Bottom Right" },
        topLeft: { ja: "左上", en: "Top Left" },
        topRight: { ja: "右上", en: "Top Right" }
    },
    diagonal: {
        backslash: { ja: "＼", en: "\\" },
        slash: { ja: "／", en: "/" }
    },
    strength: {
        mild: { ja: "マイルド", en: "Mild" },
        normal: { ja: "ノーマル", en: "Normal" },
        boost: { ja: "ブースト", en: "Boost" }
    },
    checkbox: {
        preview: { ja: "プレビュー", en: "Preview" }
    },
    button: {
        cancel: { ja: "キャンセル", en: "Cancel" },
        ok: { ja: "OK", en: "OK" }
    },
    tip: {
        amount: { ja: "0.00-0.49", en: "0.00-0.49" }
    }
};

/* ドット区切りのパスで LABELS から文言を取得 / Look up a label by dot-separated path */
function L(labelPath) {
    var pathSegments = String(labelPath).split(".");
    var labelNode = LABELS;

    for (var i = 0; i < pathSegments.length; i++) {
        if (!labelNode || labelNode[pathSegments[i]] == null) return labelPath;
        labelNode = labelNode[pathSegments[i]];
    }

    if (labelNode[lang] != null) return labelNode[lang];
    if (labelNode.ja != null) return labelNode.ja;
    if (labelNode.en != null) return labelNode.en;
    return labelPath;
}

/* {name} 形式のプレースホルダを埋める / Fill in {name} placeholders */
function formatLabel(template, values) {
    return String(template).replace(/\{(\w+)\}/g, function (token, name) {
        return (values[name] == null) ? token : values[name];
    });
}

// =========================================
// 単位 / Units
// =========================================
/* 自由変形の座標は 0〜1 の正規化値で扱うため、定規単位には依存しません。
   Free Distort coordinates are normalized to 0-1, so the ruler unit is irrelevant here. */

// =========================================
// UIレイアウトの共通設定 / Shared UI layout
// =========================================

/* ウィンドウ・パネルの余白と間隔 / Window & panel margins and spacing */
var WINDOW_MARGINS = 16;                 /* ウィンドウ外周の余白 / window margin */
var WINDOW_SPACING = 12;                 /* ウィンドウ内の要素間隔 / window spacing */
var PANEL_MARGINS  = [16, 20, 16, 12];   /* パネル余白 [左,上,右,下] / panel margins */
var PANEL_SPACING  = 8;                  /* パネル内の要素間隔 / panel spacing */
var COLUMN_SPACING = 12;                 /* 2カラムの間隔 / gap between columns */

/* ウィンドウの共通設定 / Apply shared window layout */
function setupWindow(win, spacing) {
    win.orientation = "column";
    win.alignChildren = "fill";
    win.margins = WINDOW_MARGINS;
    win.spacing = (typeof spacing === "number") ? spacing : WINDOW_SPACING;
}

/* パネルの共通設定 / Apply shared panel layout */
function setupPanel(panel, spacing) {
    panel.orientation = "column";
    panel.alignChildren = ["fill", "top"];
    panel.alignment = "fill";
    panel.margins = PANEL_MARGINS;
    panel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
}

/* 行グループの共通設定（ボタン列など） / Apply a horizontal row group */
function setupRow(group, alignment, spacing) {
    group.orientation = "row";
    group.alignment = alignment || "left";
    group.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
}

/* 固定変形アイコンボタンの外観 / Appearance of the fixed-transform icon buttons */
var ICON_BUTTON_SIZE = 40;          /* ボタンの一辺（px） / button side in pixels */
var ICON_BUTTON_PADDING = 5;        /* ボタン枠とタイルの余白（px） / gap between the button edge and the tile */
var ICON_GRID_COLUMNS = 2;          /* アイコンを並べる既定の列数 / default icons per row */
var TRAPEZOID_ICON_COLUMNS = 1;     /* 台形は1列に縦積み / the trapezoid presets stack in one column */
var ICON_GRID_SPACING = 4;          /* アイコン同士の間隔（px） / gap between icons */

var ICON_TILE_LINE_WIDTH = 2;       /* タイルの枠線の太さ（px） / width of the tile frame */
var ICON_SHAPE_SHRINK = 0.94;       /* 図形の縮小率（枠に触れさせない） / shrink so the shape clears the frame */
var ICON_SHAPE_LINE_WIDTH = 1;      /* 図形の輪郭線の太さ（px） / outline width of a filled shape */

/* アイコンに描く変形量のサンプル値。シアーは実際の適用と同じ比率（変形量×基準倍率）
   にして、アイコンの形と適用結果がずれないようにする。
   Sample amounts used when drawing the icons. The shear keeps the same ratio as the real
   application (amount times the base factor) so the icon matches what gets applied. */
var ICON_SAMPLE_TRAPEZOID_AMOUNT = 0.20;
var ICON_SAMPLE_SHEAR_AMOUNT =
    ICON_SAMPLE_TRAPEZOID_AMOUNT * DISTORT_CONFIG.shearBaseFactor;
var ICON_DIAGONAL_LINE_WIDTH = 3.5; /* 対角線の太さ（px） / stroke width of the diagonal presets */
var ICON_SELECTED_LINE_WIDTH = 2;   /* 選択枠の太さ（px） / width of the selection frame */

/* この面積を下回る図形は線として描く（対角線プリセット用）
   Shapes below this area are drawn as a line; this covers the diagonal presets */
var DEGENERATE_AREA_THRESHOLD = 0.0001;

/* アイコンの配色（RGBA、各0〜1） / Icon colors, RGBA in the 0-1 range */
var ICON_COLORS = {
    tileFill:   [0.91, 0.91, 0.91, 1],  /* 変形前の領域 / the area before distortion */
    tileBorder: [0.76, 0.76, 0.76, 1],  /* タイルの枠線 / the tile frame */
    distorted:  [0.09, 0.07, 0.06, 1],  /* 変形後の形 / the distorted shape */
    selected:   [0.20, 0.55, 0.95, 1]   /* 選択枠 / the selection frame */
};

// =========================================
// プリセット定義 / Preset definitions
// =========================================
/* makeCorners は変形後の4隅を [TL, TR, BL, BR] の順で返す。
   引数 trapezoidAmount / shearAmount は強度倍率を掛けたあとの変形量。
   makeCorners returns the destination corners in [TL, TR, BL, BR] order.
   Its arguments are the amounts after the strength factor has been applied.
   UIの並び順・ラベル・可変/固定・座標をこのテーブル1か所で管理する。
   Order, label, adjustable flag and geometry all live in this single table. */
var DISTORT_PRESETS = [
    { presetKey: "trapNarrowTop", labelText: L("trapezoid.narrowTop"), groupKey: "trapezoid", adjustable: true,
      makeCorners: function (trapezoidAmount) {
          return [[trapezoidAmount, 0], [1 - trapezoidAmount, 0], [0, 1], [1, 1]];
      } },

    { presetKey: "trapWideTop", labelText: L("trapezoid.wideTop"), groupKey: "trapezoid", adjustable: true,
      makeCorners: function (trapezoidAmount) {
          return [[-trapezoidAmount, 0], [1 + trapezoidAmount, 0], [0, 1], [1, 1]];
      } },

    { presetKey: "trapNarrowBottom", labelText: L("trapezoid.narrowBottom"), groupKey: "trapezoid", adjustable: true,
      makeCorners: function (trapezoidAmount) {
          return [[0, 0], [1, 0], [trapezoidAmount, 1], [1 - trapezoidAmount, 1]];
      } },

    { presetKey: "trapWideBottom", labelText: L("trapezoid.wideBottom"), groupKey: "trapezoid", adjustable: true,
      makeCorners: function (trapezoidAmount) {
          return [[0, 0], [1, 0], [-trapezoidAmount, 1], [1 + trapezoidAmount, 1]];
      } },

    /* 平行四辺形は 4基準点 × 2軸 の8パターンを後段で生成して追加する
       The eight parallelogram presets are generated further below */

    /* 三角形：名前の角と対角の頂点をたたむ。並び順はアイコングリッドの
       左上→右上→左下→右下に対応させる。
       Triangle: fold the corner opposite to the name. The order matches the icon
       grid, filled top-left, top-right, bottom-left, bottom-right. */
    { presetKey: "triTopLeft", labelText: L("triangle.topLeft"), groupKey: "triangle", adjustable: false,
      makeCorners: function () { return [[0, 0], [1, 0], [0, 1], [0, 1]]; } },

    { presetKey: "triTopRight", labelText: L("triangle.topRight"), groupKey: "triangle", adjustable: false,
      makeCorners: function () { return [[0, 0], [1, 0], [1, 1], [1, 1]]; } },

    { presetKey: "triBottomLeft", labelText: L("triangle.bottomLeft"), groupKey: "triangle", adjustable: false,
      makeCorners: function () { return [[0, 0], [0, 0], [0, 1], [1, 1]]; } },

    { presetKey: "triBottomRight", labelText: L("triangle.bottomRight"), groupKey: "triangle", adjustable: false,
      makeCorners: function () { return [[1, 0], [1, 0], [0, 1], [1, 1]]; } },

    /* 対角線：上辺と下辺をそれぞれ1点につぶす / Diagonal: collapse the top and bottom edges */
    { presetKey: "diagonalBackslash", labelText: L("diagonal.backslash"), groupKey: "diagonal", adjustable: false,
      makeCorners: function () { return [[0, 0], [0, 0], [1, 1], [1, 1]]; } },

    { presetKey: "diagonalSlash", labelText: L("diagonal.slash"), groupKey: "diagonal", adjustable: false,
      makeCorners: function () { return [[1, 0], [1, 0], [0, 1], [0, 1]]; } }
];

/* 平行四辺形の基準点。x=0 が左、y=0 が上。
   Anchor corners for the parallelogram presets; x=0 is left, y=0 is top. */
var SHEAR_ANCHORS = [
    { anchorKey: "TopLeft",     x: 0, y: 0 },
    { anchorKey: "TopRight",    x: 1, y: 0 },
    { anchorKey: "BottomLeft",  x: 0, y: 1 },
    { anchorKey: "BottomRight", x: 1, y: 1 }
];

/* シアーの方向 / Shear axes */
var SHEAR_AXES = [
    { axisKey: "Horizontal", labelPath: "shear.axisHorizontal" },
    { axisKey: "Vertical",   labelPath: "shear.axisVertical" }
];

/* 基準点の角を固定し、その角を含む辺を残したまま、反対側の辺を基準点から
   遠ざかる向きにずらす。これで 4基準点 × 2軸 の8パターンが一意に決まる。
   Pin the anchor corner: the edge through it stays put and the opposite edge slides
   away from the anchor. Four anchors times two axes give eight distinct shears. */
function makeShearCorners(anchor, axisKey, shearAmount) {
    var corners = [[0, 0], [1, 0], [0, 1], [1, 1]]; /* [TL, TR, BL, BR] */
    var i;

    if (axisKey === "Horizontal") {
        /* 基準点の反対側の横辺が、左右にずれる / the opposite horizontal edge slides sideways */
        var movingRow = 1 - anchor.y;
        var horizontalShift = (anchor.x === 0) ? shearAmount : -shearAmount;

        for (i = 0; i < corners.length; i++) {
            if (corners[i][1] === movingRow) corners[i][0] += horizontalShift;
        }
    } else {
        /* 基準点の反対側の縦辺が、上下にずれる / the opposite vertical edge slides up or down */
        var movingColumn = 1 - anchor.x;
        var verticalShift = (anchor.y === 0) ? shearAmount : -shearAmount;

        for (i = 0; i < corners.length; i++) {
            if (corners[i][0] === movingColumn) corners[i][1] += verticalShift;
        }
    }
    return corners;
}

/* 8パターンを生成してテーブルに追加する。軸ごとに4基準点を並べるので、
   2列のアイコングリッドでは上2行が左右、下2行が上下になる。
   Generate the eight presets; anchors are grouped by axis, so in the two-column icon
   grid the top two rows are the horizontal shears and the bottom two the vertical ones. */
(function () {
    for (var axisIndex = 0; axisIndex < SHEAR_AXES.length; axisIndex++) {
        for (var anchorIndex = 0; anchorIndex < SHEAR_ANCHORS.length; anchorIndex++) {
            var axis = SHEAR_AXES[axisIndex];
            var anchor = SHEAR_ANCHORS[anchorIndex];

            DISTORT_PRESETS.push({
                presetKey: "shear" + axis.axisKey + anchor.anchorKey,
                labelText: formatLabel(L("shear.tipFormat"), {
                    axis: L(axis.labelPath),
                    anchor: L("shear.anchor" + anchor.anchorKey)
                }),
                groupKey: "parallelogram",
                adjustable: true,
                /* クロージャで軸と基準点を閉じ込める / capture the axis and anchor in a closure */
                makeCorners: (function (capturedAnchor, capturedAxisKey) {
                    return function (trapezoidAmount, shearAmount) {
                        return makeShearCorners(capturedAnchor, capturedAxisKey, shearAmount);
                    };
                })(anchor, axis.axisKey)
            });
        }
    }
})();

var DEFAULT_PRESET_INDEX = 0; /* 初期選択プリセット / initially selected preset */

/* パネルに割り当てるプリセットグループ。値は DISTORT_PRESETS の groupKey と
   LABELS.panel のキーを兼ねる。
   Preset groups assigned to panels; each value doubles as a groupKey in
   DISTORT_PRESETS and as a key in LABELS.panel. */
var TRAPEZOID_GROUP_KEY = "trapezoid";
var SHEAR_GROUP_KEY = "parallelogram";
var FIXED_GROUP_KEYS = ["triangle", "diagonal"];


(function () {

    // =========================================
    // 事前チェック / Preflight checks
    // =========================================
    if (app.documents.length === 0) {
        alert(L("alert.noDocument"));
        return;
    }

    var doc = app.activeDocument;
    if (!doc.selection || doc.selection.length === 0) {
        alert(L("alert.noSelection"));
        return;
    }

    // =========================================
    // ライブ効果の適用 / Live effect application
    // =========================================

    /* 選択のうち、ライブ効果を適用できるオブジェクトだけを抽出。
       TextRange などは applyEffect を持たないため除外する。
       Collect the selected objects that can take a live effect; TextRange and
       similar selections have no applyEffect and are skipped. */
    function getEffectTargets() {
        var selectedItems = doc.selection;
        var targetItems = [];

        if (!selectedItems || selectedItems.length == null) return targetItems;

        for (var i = 0; i < selectedItems.length; i++) {
            if (selectedItems[i] && selectedItems[i].applyEffect != undefined) {
                targetItems.push(selectedItems[i]);
            }
        }
        return targetItems;
    }

    /* 保持しておいた参照のうち、まだ有効なものだけを返す。
       Undo をまたぐと参照が無効になっていることがあるため、プロパティ参照で確かめる。
       Keep only the references that are still alive: an undo can invalidate them, so
       probe each one by touching a property. */
    function filterLiveItems(items) {
        var liveItems = [];

        for (var i = 0; i < items.length; i++) {
            try {
                if (items[i].typename) liveItems.push(items[i]);
            } catch (error) {
                /* すでに無効な参照は捨てる / drop references that are already dead */
            }
        }
        return liveItems;
    }

    /* 変形後の4隅から自由変形の XML を組み立てる。
       プレースホルダは位置ではなく名前で解決する（src0h と src0v の取り違えを防ぐ）。
       Build the Free Distort XML; placeholders resolve by name, not by position. */
    function buildFreeDistortXML(destCorners) {
        var sourceCorners = [[0, 0], [1, 0], [0, 1], [1, 1]];
        var cornerValues = {};

        for (var i = 0; i < 4; i++) {
            /* Y軸は Illustrator の座標系に合わせて符号を反転 / flip Y for the Illustrator axis */
            cornerValues["src" + i + "h"] = sourceCorners[i][0];
            cornerValues["src" + i + "v"] = -sourceCorners[i][1];
            cornerValues["dst" + i + "h"] = destCorners[i][0];
            cornerValues["dst" + i + "v"] = -destCorners[i][1];
        }

        var xmlTemplate = '<LiveEffect name="Adobe Free Distort"><Dict data="' +
            'R src0h {src0h} R src0v {src0v} R src1h {src1h} R src1v {src1v} ' +
            'R src2h {src2h} R src2v {src2v} R src3h {src3h} R src3v {src3v} ' +
            'R dst0h {dst0h} R dst0v {dst0v} R dst1h {dst1h} R dst1v {dst1v} ' +
            'R dst2h {dst2h} R dst2v {dst2v} R dst3h {dst3h} R dst3v {dst3v} ' +
            '"/></LiveEffect>';

        return xmlTemplate.replace(/\{(\w+)\}/g, function (token, tokenName) {
            return (cornerValues[tokenName] == null) ? token : String(cornerValues[tokenName]);
        });
    }

    /* 各オブジェクトへ効果を適用する。
       onApplied は1件適用するたびに、それまでの適用件数を引数として呼ばれる。
       例外で中断しても、呼び出し側は何件まで適用済みかを把握できる。
       Apply the effect to each object. onApplied receives the running count after every
       object, so the caller still knows how far it got when an exception interrupts it. */
    function applyFreeDistort(targetItems, destCorners, onApplied) {
        var xml = buildFreeDistortXML(destCorners);

        for (var i = 0; i < targetItems.length; i++) {
            targetItems[i].applyEffect(xml);
            if (onApplied) onApplied(i + 1);
        }
    }

    // =========================================
    // Undo境界の目印 / Undo boundary marker
    // =========================================

    /* applyEffect がオブジェクトごとに別々のアンドゥステップになるか、まとめて1ステップに
       なるかは環境によって変わる。件数ぶん undo すると戻りすぎ、1回だけだと戻し足りない。
       そこで適用前に空レイヤーを1枚足しておき、それが消えるまで undo する。
       上限は適用件数なので、ユーザー自身の操作まで巻き戻すことは決してない。
       Whether applyEffect produces one undo step per object or a single coalesced step
       varies by environment: undoing once can leave effects behind, and undoing once per
       object can eat the user's own edits. So add an empty layer before applying and undo
       until it disappears, capped at the number of applications - never further back. */
    var UNDO_MARK_LAYER_NAME = "__" + SCRIPT_NAME + "_undo_mark__";
    var undoMarkBaseLayerCount = -1;

    /* 目印レイヤーを追加する。追加できなければ false
       Add the marker layer; returns false when it cannot be created */
    function beginUndoMark() {
        try {
            undoMarkBaseLayerCount = doc.layers.length;
            doc.layers.add().name = UNDO_MARK_LAYER_NAME;
            return true;
        } catch (error) {
            undoMarkBaseLayerCount = -1;
            return false;
        }
    }

    /* 目印がまだ残っているか。名前ではなくレイヤー数で見るのは、レイヤー追加と
       リネームが別のアンドゥステップに分かれる場合があるため。
       Whether the marker is still there. This counts layers rather than matching the name,
       because adding and renaming a layer can land in two separate undo steps. */
    function hasUndoMark() {
        return undoMarkBaseLayerCount >= 0 && doc.layers.length > undoMarkBaseLayerCount;
    }

    /* 目印を直接取り除く。undo で消えなかった場合の後始末に使う。
       名前で引き当てるので、取り違えて他のレイヤーを消すことはない。
       Delete the marker outright when undo did not remove it. The lookup is by name, so
       this can never take out one of the user's own layers. */
    function removeUndoMark() {
        try {
            doc.layers.getByName(UNDO_MARK_LAYER_NAME).remove();
        } catch (error) {
            /* すでに無ければ何もしない / nothing to do when it is already gone */
        }
        undoMarkBaseLayerCount = -1;
    }

    /* 目印が消えるまで undo する。stepLimit（＝適用件数）を超えて戻すことはない
       Undo until the marker is gone, never exceeding stepLimit (the applied count) */
    function undoToMark(stepLimit) {
        var steps = 0;

        while (steps < stepLimit && hasUndoMark()) {
            app.undo();
            steps++;
        }
        if (hasUndoMark()) removeUndoMark();
        undoMarkBaseLayerCount = -1;
    }

    /* 例外内容をユーザーに通知 / Report an error to the user */
    function showError(error) {
        alert(SCRIPT_NAME + ": " + (error && error.message ? error.message : String(error)));
    }

    // =========================================
    // 変形量の計算 / Amount calculation
    // =========================================

    /* 強度キーから倍率を取得（未知のキーはノーマル扱い）
       Resolve the strength factor; unknown keys fall back to normal */
    function getStrengthFactor(strengthKey) {
        var factor = STRENGTH_FACTORS[strengthKey];
        return (factor == null) ? STRENGTH_FACTORS.normal : factor;
    }

    /* プリセットから変形後の4隅座標を生成。固定変形では変形量・強度を使わない。
       Build the destination corners for a preset; fixed presets ignore amount and strength. */
    function makeDestCorners(preset, amountRatio, strengthKey) {
        if (!preset.adjustable) return preset.makeCorners();

        var trapezoidAmount = amountRatio * getStrengthFactor(strengthKey);
        var shearAmount = trapezoidAmount * DISTORT_CONFIG.shearBaseFactor;

        return preset.makeCorners(trapezoidAmount, shearAmount);
    }

    // =========================================
    // 選択の判定 / Selection inspection
    // =========================================

    /* オブジェクト内にテキストが含まれるかを再帰的に判定（グループの入れ子に対応）
       Recursively check whether an object contains text, walking nested groups */
    function itemContainsText(item) {
        if (!item || !item.typename) return false;
        if (item.typename === "TextFrame") return true;
        if (!item.pageItems) return false;

        for (var i = 0; i < item.pageItems.length; i++) {
            if (itemContainsText(item.pageItems[i])) return true;
        }
        return false;
    }

    /* 選択全体にテキストが含まれるか / Whether the whole selection contains text */
    function selectionContainsText(selectedItems) {
        if (!selectedItems || selectedItems.length == null) return false;

        for (var i = 0; i < selectedItems.length; i++) {
            if (itemContainsText(selectedItems[i])) return true;
        }
        return false;
    }

    // =========================================
    // ダイアログの構築 / Dialog construction
    // =========================================

    /* [TL, TR, BL, BR] を外周をたどる順に並べ替える
       Reorder [TL, TR, BL, BR] so the points walk the perimeter */
    function toPerimeter(corners) {
        return [corners[0], corners[1], corners[3], corners[2]];
    }

    /* 点群を囲む矩形 / The rectangle enclosing the points */
    function getBounds(points) {
        var bounds = { minX: points[0][0], minY: points[0][1],
                       maxX: points[0][0], maxY: points[0][1] };

        for (var i = 1; i < points.length; i++) {
            bounds.minX = Math.min(bounds.minX, points[i][0]);
            bounds.minY = Math.min(bounds.minY, points[i][1]);
            bounds.maxX = Math.max(bounds.maxX, points[i][0]);
            bounds.maxY = Math.max(bounds.maxY, points[i][1]);
        }
        return bounds;
    }

    /* 点群を (0.5, 0.5) を中心に縮小する。図形がタイル枠に触れないようにするため。
       Shrink the points about (0.5, 0.5) so the shape never touches the tile frame. */
    function shrinkAboutCenter(points, factor) {
        var shrunk = [];

        for (var i = 0; i < points.length; i++) {
            shrunk.push([
                0.5 + (points[i][0] - 0.5) * factor,
                0.5 + (points[i][1] - 0.5) * factor
            ]);
        }
        return shrunk;
    }

    /* 変形前の正方形と変形後の図形の両方が収まるよう、描画領域に合わせて
       拡大率と位置を決め、正規化座標をピクセル座標へ移す関数を返す。
       平行四辺形や台形は 0〜1 の外へはみ出すため、この収め込みが必要になる。
       Return a mapper from normalized to pixel coordinates, scaled so that both the
       source square and the distorted shape fit the drawing area. The parallelogram
       and trapezoid presets reach outside 0-1, so the fit is required. */
    function makeIconScaler(points, areaLeft, areaTop, areaSide) {
        var bounds = getBounds(points);
        var width = bounds.maxX - bounds.minX;
        var height = bounds.maxY - bounds.minY;
        var scale = areaSide / Math.max(width, height);
        var offsetX = areaLeft + (areaSide - width * scale) / 2 - bounds.minX * scale;
        var offsetY = areaTop + (areaSide - height * scale) / 2 - bounds.minY * scale;

        return function (point) {
            return [offsetX + point[0] * scale, offsetY + point[1] * scale];
        };
    }

    /* 多角形のパスを引く / Trace a polygon path */
    function tracePolygon(graphics, points, toPixel) {
        graphics.newPath();

        for (var i = 0; i < points.length; i++) {
            var pixel = toPixel(points[i]);
            if (i === 0) graphics.moveTo(pixel[0], pixel[1]);
            else graphics.lineTo(pixel[0], pixel[1]);
        }
        graphics.closePath();
    }

    /* アイコンボタンの描画。変形後の4隅を makeCorners から直接得るため、
       アイコンの形と実際に適用される効果が必ず一致する。
       Draw an icon button. The corners come straight from makeCorners, so the icon
       always matches the effect that will actually be applied.
       this は描画対象のボタン / `this` is the button being drawn. */
    function drawPresetIcon() {
        var graphics = this.graphics;
        var preset = this.presetRef;

        var areaSide = Math.min(this.size.width, this.size.height) - ICON_BUTTON_PADDING * 2;
        var areaLeft = (this.size.width - areaSide) / 2;
        var areaTop = (this.size.height - areaSide) / 2;

        var sourceSquare = toPerimeter([[0, 0], [1, 0], [0, 1], [1, 1]]);
        var distorted = shrinkAboutCenter(
            toPerimeter(preset.adjustable
                ? preset.makeCorners(ICON_SAMPLE_TRAPEZOID_AMOUNT, ICON_SAMPLE_SHEAR_AMOUNT)
                : preset.makeCorners()),
            ICON_SHAPE_SHRINK);

        var toPixel = makeIconScaler(
            sourceSquare.concat(distorted), areaLeft, areaTop, areaSide);

        /* 変形前の領域をタイルとして塗り、枠線を回す
           Fill the source area as a tile and outline it */
        tracePolygon(graphics, sourceSquare, toPixel);
        graphics.fillPath(graphics.newBrush(graphics.BrushType.SOLID_COLOR, ICON_COLORS.tileFill));
        graphics.strokePath(graphics.newPen(
            graphics.PenType.SOLID_COLOR, ICON_COLORS.tileBorder, ICON_TILE_LINE_WIDTH));

        tracePolygon(graphics, distorted, toPixel);

        if (getPolygonArea(distorted) < DEGENERATE_AREA_THRESHOLD) {
            /* 対角線は面積が0で塗りが出ないため、太い線として描く
               The diagonals have no area to fill, so draw them as a thick stroke */
            graphics.strokePath(graphics.newPen(
                graphics.PenType.SOLID_COLOR, ICON_COLORS.distorted, ICON_DIAGONAL_LINE_WIDTH));
        } else {
            /* 塗りだけだと輪郭が粗く出るため、同色の細線を重ねる
               Fill alone leaves ragged edges, so overlay a thin outline in the same color */
            graphics.fillPath(graphics.newBrush(
                graphics.BrushType.SOLID_COLOR, ICON_COLORS.distorted));
            graphics.strokePath(graphics.newPen(
                graphics.PenType.SOLID_COLOR, ICON_COLORS.distorted, ICON_SHAPE_LINE_WIDTH));
        }

        /* 選択中はボタン全体を枠で囲んで示す
           Mark the selected button with a frame around the whole button */
        if (this.value) {
            graphics.newPath();
            graphics.rectPath(1, 1, this.size.width - 2, this.size.height - 2);
            graphics.strokePath(graphics.newPen(
                graphics.PenType.SOLID_COLOR, ICON_COLORS.selected, ICON_SELECTED_LINE_WIDTH));
        }
    }

    /* 多角形の面積（靴ひも公式）。対角線プリセットのように面積が0のとき、
       塗りではなく線で描くべきかを判定するために使う。
       Polygon area via the shoelace formula, used to decide whether a preset has to be
       stroked as a line instead of filled, as with the diagonal presets. */
    function getPolygonArea(points) {
        var doubledArea = 0;

        for (var i = 0; i < points.length; i++) {
            var next = points[(i + 1) % points.length];
            doubledArea += points[i][0] * next[1] - next[0] * points[i][1];
        }
        return Math.abs(doubledArea) / 2;
    }

    /* プリセット1つぶんのアイコンボタンを追加 / Add one icon button for a preset */
    function addPresetIconButton(parentGroup, preset) {
        var iconButton = parentGroup.add("iconbutton", undefined, undefined,
            { style: "toolbutton", toggle: true });

        iconButton.preferredSize = [ICON_BUTTON_SIZE, ICON_BUTTON_SIZE];
        iconButton.helpTip = preset.labelText;
        iconButton.presetRef = preset;
        iconButton.onDraw = drawPresetIcon;

        return iconButton;
    }

    /* 指定グループのプリセットを、専用パネル内のアイコングリッドとして並べる。
       columns 個ごとに行を折り返す（省略時は ICON_GRID_COLUMNS）。
       Lay out the presets of a group as an icon grid inside its own panel, wrapping
       to a new row every `columns` icons, defaulting to ICON_GRID_COLUMNS. */
    function addPresetIconPanel(parentGroup, groupKey, presetControls, columns) {
        var groupPanel = parentGroup.add("panel", undefined, L("panel." + groupKey));
        setupPanel(groupPanel);
        groupPanel.alignChildren = ["center", "top"];
        groupPanel.spacing = ICON_GRID_SPACING;

        var iconsPerRow = (typeof columns === "number") ? columns : ICON_GRID_COLUMNS;
        var currentRow = null;
        var iconsInRow = 0;

        for (var i = 0; i < DISTORT_PRESETS.length; i++) {
            if (DISTORT_PRESETS[i].groupKey !== groupKey) continue;

            if (!currentRow || iconsInRow >= iconsPerRow) {
                currentRow = groupPanel.add("group");
                setupRow(currentRow, "center", ICON_GRID_SPACING);
                iconsInRow = 0;
            }

            presetControls[i] = addPresetIconButton(currentRow, DISTORT_PRESETS[i]);
            iconsInRow++;
        }
        return groupPanel;
    }

    /* 複数グループのアイコンパネルをまとめて追加
       Add the icon panels for several groups at once */
    function addPresetIconPanels(parentGroup, groupKeys, presetControls) {
        for (var i = 0; i < groupKeys.length; i++) {
            addPresetIconPanel(parentGroup, groupKeys[i], presetControls);
        }
    }

    /* 変形量スライダーと強度ラジオを組み立てる / Build the amount slider and strength radios */
    function buildAmountPanel(parentPanel) {
        var amountPanel = parentPanel.add("panel", undefined, L("panel.amount"));
        setupPanel(amountPanel);

        var amountSlider = amountPanel.add("slider", undefined,
            DISTORT_CONFIG.amountDefaultPercent,
            DISTORT_CONFIG.amountMinPercent,
            DISTORT_CONFIG.amountMaxPercent);
        amountSlider.helpTip = L("tip.amount");
        amountSlider.preferredSize.width = DISTORT_CONFIG.amountSliderWidth;

        var amountReadout = amountPanel.add("statictext", undefined, "");
        amountReadout.alignment = ["center", "center"];

        var strengthPanel = amountPanel.add("panel", undefined, L("panel.strength"));
        setupPanel(strengthPanel);

        /* ラジオを内側のグループにまとめ、パネルの左右中央に置く
           Group the radios so the row can be centered inside the panel */
        var strengthRow = strengthPanel.add("group");
        setupRow(strengthRow, "center");

        var mildRadio = strengthRow.add("radiobutton", undefined, L("strength.mild"));
        var normalRadio = strengthRow.add("radiobutton", undefined, L("strength.normal"));
        var boostRadio = strengthRow.add("radiobutton", undefined, L("strength.boost"));

        return {
            amountPanel: amountPanel,
            amountSlider: amountSlider,
            amountReadout: amountReadout,
            mildRadio: mildRadio,
            normalRadio: normalRadio,
            boostRadio: boostRadio
        };
    }

    /* ボタンエリアを左右分割で組み立てる。
       左：プレビュー、中央：伸縮スペーサー、右：キャンセル / OK。
       Build the footer split left and right: Preview on the left, a stretchable spacer
       in the middle, and Cancel / OK on the right. */
    function buildFooterRow(parentWindow) {
        /* メイングループ（横並び） / Main group (horizontal layout) */
        var btnRowGroup = parentWindow.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.margins = [10, 10, 10, 0];
        btnRowGroup.alignment = ["fill", "bottom"];

        /* 左側グループ / Left-side button group
           プレビューは押しっぱなしの切り替えなので、ボタンではなくチェックボックスにしている
           Preview is a sticky toggle, so it stays a checkbox rather than a button */
        var btnLeftGroup = btnRowGroup.add("group");
        btnLeftGroup.alignChildren = ["left", "center"];
        var previewCheckbox = btnLeftGroup.add("checkbox", undefined, L("checkbox.preview"));

        /* スペーサー（伸縮） / Spacer (stretchable) */
        var spacer = btnRowGroup.add("group");
        spacer.alignment = ["fill", "fill"];
        spacer.minimumSize.width = 0;

        /* 右側グループ / Right-side button group */
        var btnRightGroup = btnRowGroup.add("group");
        btnRightGroup.alignChildren = ["right", "center"];
        btnRightGroup.add("button", undefined, L("button.cancel"), { name: "cancel" });
        btnRightGroup.add("button", undefined, L("button.ok"), { name: "ok" });

        return previewCheckbox;
    }

    /* ダイアログ本体を組み立てて各コントロールを返す
       Build the dialog and return references to its controls */
    function buildDialog(preferMildStrength) {
        var dialog = new Window("dialog", L("dialog.title") + " " + SCRIPT_VERSION);
        setupWindow(dialog);

        var columnsGroup = dialog.add("group");
        setupRow(columnsGroup, "fill", COLUMN_SPACING);
        columnsGroup.alignChildren = ["fill", "top"];

        /* 添字は DISTORT_PRESETS と対応させる / indices match DISTORT_PRESETS */
        var presetControls = [];

        /* 名前では変形の向きが伝わらないため、すべてアイコンで形を見せる
           Names cannot convey the direction of a distortion, so every preset is an icon */

        /* --- 1カラム目：台形（1列×4行） / Column 1: trapezoid, one column of four --- */
        addPresetIconPanel(columnsGroup, TRAPEZOID_GROUP_KEY, presetControls,
            TRAPEZOID_ICON_COLUMNS);

        /* --- 2カラム目：平行四辺形（2列×4行、上2行が左右・下2行が上下）
               Column 2: parallelogram, 2 by 4 with horizontal on top and vertical below --- */
        addPresetIconPanel(columnsGroup, SHEAR_GROUP_KEY, presetControls);

        /* --- 3カラム目：三角形と対角線を縦に積む
               Column 3: the triangle and diagonal panels stacked --- */
        var fixedColumn = columnsGroup.add("group");
        fixedColumn.orientation = "column";
        fixedColumn.alignChildren = ["fill", "top"];
        fixedColumn.spacing = PANEL_SPACING;
        addPresetIconPanels(fixedColumn, FIXED_GROUP_KEYS, presetControls);

        /* --- 変形量：プリセットの2カラムの下に、全幅で置く
               Amount: placed full width below the two preset columns --- */
        var amountControls = buildAmountPanel(dialog);

        var previewCheckbox = buildFooterRow(dialog);

        /* --- 初期状態 / Initial state --- */
        presetControls[DEFAULT_PRESET_INDEX].value = true;
        previewCheckbox.value = false;
        /* テキストを含む選択では、崩れを抑えるためマイルドを初期値にする
           Default to Mild when text is selected, to keep the shapes readable */
        amountControls.mildRadio.value = !!preferMildStrength;
        amountControls.normalRadio.value = !preferMildStrength;
        amountControls.boostRadio.value = false;

        return {
            dialog: dialog,
            presetControls: presetControls,
            amountControls: amountControls,
            previewCheckbox: previewCheckbox
        };
    }

    // =========================================
    // ダイアログの制御 / Dialog interaction
    // =========================================

    /* ダイアログを表示し、確定した設定を返す（キャンセル時は null）
       Show the dialog and return the confirmed settings, or null when cancelled */
    function showDialog() {
        var dialogUI = buildDialog(selectionContainsText(doc.selection));
        var presetControls = dialogUI.presetControls;
        var amountControls = dialogUI.amountControls;
        var previewCheckbox = dialogUI.previewCheckbox;

        /* 取り消すべきプレビューが何件ぶん残っているか。undoToMark の上限に使う。
           How many applications of the pending preview remain; used as the undo cap. */
        var previewAppliedCount = 0;
        var lastPreviewUpdateTime = 0;
        var lastPreviewSignature = null;

        /* 選択中のプリセットを取得 / Get the selected preset */
        function getSelectedPreset() {
            for (var i = 0; i < presetControls.length; i++) {
                if (presetControls[i].value) return DISTORT_PRESETS[i];
            }
            return DISTORT_PRESETS[DEFAULT_PRESET_INDEX];
        }

        /* スライダー値を整数（%）に丸める。表示・適用・差分判定で同じ値を使うため
           Round the slider to an integer percent so display, apply and diffing agree */
        function getAmountPercent() {
            return Math.round(amountControls.amountSlider.value);
        }

        /* 実際の変形量（0.00〜0.49） / The actual amount, in the 0.00-0.49 range */
        function getAmountRatio() {
            return getAmountPercent() / 100;
        }

        /* 選択中の強度キー / The selected strength key */
        function getStrengthKey() {
            if (amountControls.mildRadio.value) return "mild";
            if (amountControls.boostRadio.value) return "boost";
            return "normal";
        }

        /* 変形量の数値表示を更新 / Refresh the numeric amount readout */
        function updateAmountReadout() {
            amountControls.amountReadout.text = getAmountRatio().toFixed(2);
        }

        /* 固定変形では変形量・強度をディム表示にする（子コントロールもまとめて無効化）
           Dim the amount panel for fixed presets; its children follow automatically */
        function updateAmountPanelEnabled() {
            amountControls.amountPanel.enabled = getSelectedPreset().adjustable;
        }

        /* 同一条件の再適用を避けるための署名 / Signature used to skip redundant work */
        function buildPreviewSignature() {
            return [
                getSelectedPreset().presetKey,
                getAmountPercent(),
                getStrengthKey(),
                previewCheckbox.value ? 1 : 0
            ].join("|");
        }

        /* プレビューを取り消す。目印レイヤーが消えるまで、適用件数を上限に undo する
           Remove the preview: undo until the marker layer is gone, capped at the count applied */
        function clearPreview() {
            if (previewAppliedCount > 0 || hasUndoMark()) {
                undoToMark(previewAppliedCount);
                app.redraw();
            }
            previewAppliedCount = 0;
            lastPreviewSignature = null;
        }

        /* プレビューを適用（ignoreThrottle=true で間引きを無視）
           Apply the preview; ignoreThrottle=true bypasses the update throttling */
        function updatePreview(ignoreThrottle) {
            if (!previewCheckbox.value) return;

            var signature = buildPreviewSignature();
            if (signature === lastPreviewSignature) return;

            var currentTime = new Date().getTime();
            if (!ignoreThrottle &&
                (currentTime - lastPreviewUpdateTime) < DISTORT_CONFIG.previewThrottleMs) {
                return;
            }
            lastPreviewUpdateTime = currentTime;

            clearPreview();

            var targetItems = getEffectTargets();
            if (targetItems.length === 0) return;

            /* 目印を置けないと安全に巻き戻せないため、プレビュー自体を諦める
               Without a marker the preview cannot be reverted safely, so skip it entirely */
            if (!beginUndoMark()) {
                previewCheckbox.value = false;
                return;
            }

            var destCorners = makeDestCorners(getSelectedPreset(), getAmountRatio(), getStrengthKey());

            try {
                applyFreeDistort(targetItems, destCorners,
                    function (appliedCount) { previewAppliedCount = appliedCount; });
            } catch (error) {
                /* 適用済みのぶんだけ巻き戻す / roll back exactly what was applied */
                clearPreview();
                previewCheckbox.value = false;
                showError(error);
                return;
            }

            app.redraw();
            lastPreviewSignature = signature;
        }

        // --- イベントハンドラ / Event handlers ---

        amountControls.amountSlider.onChanging = function () {
            updateAmountReadout();
            updatePreview(false);
        };

        amountControls.amountSlider.onChange = function () {
            updateAmountReadout();
            updatePreview(true);
        };

        previewCheckbox.onClick = function () {
            if (previewCheckbox.value) updatePreview(true);
            else clearPreview();
        };

        amountControls.mildRadio.onClick =
        amountControls.normalRadio.onClick =
        amountControls.boostRadio.onClick = function () {
            updatePreview(true);
        };

        /* プリセットは複数パネルに分かれており、さらにラジオとトグルボタンが混在するため、
           ScriptUI の排他選択は使えない。押されたものを選択状態に固定し、他を手動で解除する。
           The presets span several panels and mix radio buttons with toggle buttons, so
           ScriptUI cannot group them: pin the clicked one and clear the others by hand. */
        function selectPresetControl(clickedControl) {
            for (var k = 0; k < presetControls.length; k++) {
                presetControls[k].value = (presetControls[k] === clickedControl);
            }
            /* アイコンボタンは onDraw で選択枠を描くので、再描画させる
               The icon buttons draw their selection frame in onDraw, so force a repaint */
            if (dialogUI.dialog.update) dialogUI.dialog.update();
        }

        for (var i = 0; i < presetControls.length; i++) {
            presetControls[i].onClick = function () {
                /* トグルボタンは再クリックで解除されてしまうため、選択状態を維持する
                   A toggle button would clear itself on a second click, so keep it selected */
                selectPresetControl(this);
                updateAmountPanelEnabled();
                updatePreview(true);
            };
        }

        updateAmountReadout();
        updateAmountPanelEnabled();

        var isAccepted = (dialogUI.dialog.show() == 1);

        /* キャンセル時も OK 時も、いったんプレビューを完全に戻してから抜ける
           Either way, revert the preview completely before leaving */
        clearPreview();
        if (!isAccepted) return null;

        return {
            preset: getSelectedPreset(),
            amountRatio: getAmountRatio(),
            strengthKey: getStrengthKey()
        };
    }

    // =========================================
    // 実行 / Run
    // =========================================
    /* 効果を適用できる対象が無いなら、ダイアログを出す前に知らせる
       Tell the user before opening the dialog when nothing can take an effect */
    var initialTargets = getEffectTargets();
    if (initialTargets.length === 0) {
        alert(L("alert.noTarget"));
        return;
    }

    var distortSettings = showDialog();
    if (!distortSettings) return;

    /* プレビュー解除の Undo で選択が失われることがあるため、初期選択を保険にする
       An undo from removing the preview can clear the selection, so fall back to the
       objects captured before the dialog */
    var targetItems = getEffectTargets();
    if (targetItems.length === 0) targetItems = filterLiveItems(initialTargets);
    if (targetItems.length === 0) {
        alert(L("alert.noTarget"));
        return;
    }

    var destCorners = makeDestCorners(
        distortSettings.preset,
        distortSettings.amountRatio,
        distortSettings.strengthKey);

    /* ここでは目印レイヤーを使わない。成功時に目印を消す操作がアンドゥ履歴の最後に残り、
       ユーザーが Undo したときに空レイヤーが復活してしまうため。
       代わりに、途中で失敗したら何件まで適用したかを伝える。
       No marker layer here: removing it on success would sit at the top of the undo stack,
       so the user's first undo would resurrect an empty layer. Report how far the run got
       instead, and leave the undo to the user. */
    var appliedCount = 0;

    try {
        applyFreeDistort(targetItems, destCorners,
            function (count) { appliedCount = count; });
    } catch (error) {
        showError(error);
        if (appliedCount > 0) {
            alert(formatLabel(L("alert.partialApply"), { count: appliedCount }));
        }
    }

})();
