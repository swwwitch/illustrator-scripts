#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

「自由変形」効果を、プリセット・変形量・強度の組み合わせで適用するスクリプトです。
選択オブジェクトに対して、ダイアログUIから変形プリセットを選び、プレビューで確認しながら適用できます。

### 主な機能

- 調整変形
  - 台形（上辺を狭く / 広く、下辺を狭く / 広く）
  - 平行四辺形（右シアー / 左シアー / 上シアー / 下シアー）
  - 変形量スライダー（0.00〜0.49、ドラッグと↑↓キー操作に対応）
  - 強度切替（マイルド / ノーマル / ブースト）
- 固定変形
  - 三角形（左下 / 右下 / 左上 / 右上）
  - 対角線（＼ / ／）
- プレビュー機能（Undoベースで一時適用）
- 日英ローカライズ対応UI

### 仕様・注意

- 三角形および対角線プリセットでは、変形量と強度は使用されません。
- 台形 / 平行四辺形の強度倍率は、マイルド=0.25、ノーマル=0.5、ブースト=1.0 です。
- 選択対象にテキストが含まれる場合、強度の初期値はマイルドになります。
- プレビューは Undo を利用して制御しているため、他の操作と混在すると不整合が起こる可能性があります。
- 複数オブジェクト選択時は、すべてに同一のライブ効果を適用します。
- シアー系の名称は、座標の増減方向そのものではなく見た目基準です。

### GitHub

https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/fx/SmartFreeDistort.jsx

### 更新履歴

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

Applies the "Free Distort" live effect using a combination of preset, amount and strength.
Pick a distortion preset from the dialog UI and apply it to the selection while checking the preview.

### Key features

- Adjustable transform
  - Trapezoid (narrow / wide top, narrow / wide bottom)
  - Parallelogram (right / left / up / down shear)
  - Amount slider (0.00-0.49, supports dragging and arrow keys)
  - Strength switch (Mild / Normal / Boost)
- Fixed transform
  - Triangle (bottom left / bottom right / top left / top right)
  - Diagonal (backslash / slash)
- Preview based on Undo
- Localized UI (Japanese / English)

### Notes

- Amount and strength are ignored for the triangle and diagonal presets.
- Strength factors are Mild=0.25, Normal=0.5, Boost=1.0.
- Strength defaults to Mild when the selection contains text.
- Preview relies on Undo, so mixing it with other operations may cause inconsistencies.
- The same live effect is applied to every selected object.
- Shear names follow the visual direction, not the raw coordinate delta.

### GitHub

https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/fx/SmartFreeDistort.jsx

### Changelog

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
var SCRIPT_VERSION  = "v1.5.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-04-24";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-07-21";                   /* 更新日 / last updated */

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
        }
    },
    panel: {
        adjustable: { ja: "調整変形", en: "Adjustable Transform" },
        fixed: { ja: "固定変形", en: "Fixed Transform" },
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
        right: { ja: "右シアー", en: "Right Shear" },
        left: { ja: "左シアー", en: "Left Shear" },
        up: { ja: "上シアー", en: "Up Shear" },
        down: { ja: "下シアー", en: "Down Shear" }
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

    return labelNode[lang] || labelNode.ja || labelNode.en || labelPath;
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
    { presetKey: "trapNarrowTop", labelPath: "trapezoid.narrowTop", groupKey: "trapezoid", adjustable: true,
      makeCorners: function (trapezoidAmount) {
          return [[trapezoidAmount, 0], [1 - trapezoidAmount, 0], [0, 1], [1, 1]];
      } },

    { presetKey: "trapWideTop", labelPath: "trapezoid.wideTop", groupKey: "trapezoid", adjustable: true,
      makeCorners: function (trapezoidAmount) {
          return [[-trapezoidAmount, 0], [1 + trapezoidAmount, 0], [0, 1], [1, 1]];
      } },

    { presetKey: "trapNarrowBottom", labelPath: "trapezoid.narrowBottom", groupKey: "trapezoid", adjustable: true,
      makeCorners: function (trapezoidAmount) {
          return [[0, 0], [1, 0], [trapezoidAmount, 1], [1 - trapezoidAmount, 1]];
      } },

    { presetKey: "trapWideBottom", labelPath: "trapezoid.wideBottom", groupKey: "trapezoid", adjustable: true,
      makeCorners: function (trapezoidAmount) {
          return [[0, 0], [1, 0], [-trapezoidAmount, 1], [1 + trapezoidAmount, 1]];
      } },

    { presetKey: "shearRight", labelPath: "shear.right", groupKey: "parallelogram", adjustable: true,
      makeCorners: function (trapezoidAmount, shearAmount) {
          return [[shearAmount, 0], [1 + shearAmount, 0], [0, 1], [1, 1]];
      } },

    { presetKey: "shearLeft", labelPath: "shear.left", groupKey: "parallelogram", adjustable: true,
      makeCorners: function (trapezoidAmount, shearAmount) {
          return [[0, 0], [1, 0], [shearAmount, 1], [1 + shearAmount, 1]];
      } },

    { presetKey: "shearUp", labelPath: "shear.up", groupKey: "parallelogram", adjustable: true,
      makeCorners: function (trapezoidAmount, shearAmount) {
          return [[0, 0], [1, shearAmount], [0, 1], [1, 1 + shearAmount]];
      } },

    { presetKey: "shearDown", labelPath: "shear.down", groupKey: "parallelogram", adjustable: true,
      makeCorners: function (trapezoidAmount, shearAmount) {
          return [[0, shearAmount], [1, 0], [0, 1 + shearAmount], [1, 1]];
      } },

    /* 三角形：名前の角と対角の頂点をたたむ / Triangle: fold the corner opposite to the name */
    { presetKey: "triBottomLeft", labelPath: "triangle.bottomLeft", groupKey: "triangle", adjustable: false,
      makeCorners: function () { return [[0, 0], [0, 0], [0, 1], [1, 1]]; } },

    { presetKey: "triBottomRight", labelPath: "triangle.bottomRight", groupKey: "triangle", adjustable: false,
      makeCorners: function () { return [[1, 0], [1, 0], [0, 1], [1, 1]]; } },

    { presetKey: "triTopLeft", labelPath: "triangle.topLeft", groupKey: "triangle", adjustable: false,
      makeCorners: function () { return [[0, 0], [1, 0], [0, 1], [0, 1]]; } },

    { presetKey: "triTopRight", labelPath: "triangle.topRight", groupKey: "triangle", adjustable: false,
      makeCorners: function () { return [[0, 0], [1, 0], [1, 1], [1, 1]]; } },

    /* 対角線：上辺と下辺をそれぞれ1点につぶす / Diagonal: collapse the top and bottom edges */
    { presetKey: "diagonalBackslash", labelPath: "diagonal.backslash", groupKey: "diagonal", adjustable: false,
      makeCorners: function () { return [[0, 0], [0, 0], [1, 1], [1, 1]]; } },

    { presetKey: "diagonalSlash", labelPath: "diagonal.slash", groupKey: "diagonal", adjustable: false,
      makeCorners: function () { return [[1, 0], [1, 0], [0, 1], [0, 1]]; } }
];

var DEFAULT_PRESET_INDEX = 0; /* 初期選択プリセット / initially selected preset */

/* 左右カラムに並べるプリセットグループ。値は DISTORT_PRESETS の groupKey と
   LABELS.panel のキーを兼ねる。
   Preset groups laid out in the left and right columns; each value doubles as a
   groupKey in DISTORT_PRESETS and as a key in LABELS.panel. */
var ADJUSTABLE_GROUP_KEYS = ["trapezoid", "parallelogram"];
var FIXED_GROUP_KEYS = ["triangle", "diagonal"];

/* シアー系の名称は座標の増減方向ではなく、ダイアログ上での見た目を基準にしています。
   Shear names follow the visual direction seen in the dialog, not the raw coordinate delta. */

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
       redraw を挟まない連続した applyEffect は Illustrator 側で1つのアンドゥステップに
       まとめられるため、複数選択でも取り消しは undo 1回で足りる。
       onFirstApplied は最初の1件が適用された時点で呼ばれ、巻き戻しの要否判定に使う。
       Apply the effect to each object. Consecutive applyEffect calls without a redraw
       are coalesced into a single undo step by Illustrator, so one undo reverts them all.
       onFirstApplied fires once the first object has been touched, so the caller knows a
       rollback is needed. */
    function applyFreeDistort(targetItems, destCorners, onFirstApplied) {
        var xml = buildFreeDistortXML(destCorners);

        for (var i = 0; i < targetItems.length; i++) {
            targetItems[i].applyEffect(xml);
            if (onFirstApplied) {
                onFirstApplied();
                onFirstApplied = null;
            }
        }
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

    /* 指定グループのプリセットを1枚のパネルにラジオボタンとして並べる。
       presetRadios の添字は DISTORT_PRESETS の添字と一致させる。
       Lay out one panel of radio buttons for the given group; the presetRadios
       indices match those of DISTORT_PRESETS. */
    function addPresetGroupPanel(parentGroup, groupKey, presetRadios) {
        var groupPanel = parentGroup.add("panel", undefined, L("panel." + groupKey));
        setupPanel(groupPanel);
        groupPanel.alignChildren = ["left", "top"];

        for (var i = 0; i < DISTORT_PRESETS.length; i++) {
            if (DISTORT_PRESETS[i].groupKey !== groupKey) continue;
            presetRadios[i] = groupPanel.add("radiobutton", undefined, L(DISTORT_PRESETS[i].labelPath));
        }
        return groupPanel;
    }

    /* 複数グループのパネルをまとめて追加 / Add the panels for several groups at once */
    function addPresetGroupPanels(parentGroup, groupKeys, presetRadios) {
        for (var i = 0; i < groupKeys.length; i++) {
            addPresetGroupPanel(parentGroup, groupKeys[i], presetRadios);
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
        strengthPanel.orientation = "row";
        strengthPanel.alignChildren = ["left", "center"];

        var mildRadio = strengthPanel.add("radiobutton", undefined, L("strength.mild"));
        var normalRadio = strengthPanel.add("radiobutton", undefined, L("strength.normal"));
        var boostRadio = strengthPanel.add("radiobutton", undefined, L("strength.boost"));

        return {
            amountPanel: amountPanel,
            amountSlider: amountSlider,
            amountReadout: amountReadout,
            mildRadio: mildRadio,
            normalRadio: normalRadio,
            boostRadio: boostRadio
        };
    }

    /* プレビューチェックボックスと OK / キャンセルの行を組み立てる
       Build the footer row holding the preview checkbox and the OK / Cancel buttons */
    function buildFooterRow(parentWindow) {
        var footerRow = parentWindow.add("group");
        setupRow(footerRow, "fill");

        var footerLeftGroup = footerRow.add("group");
        setupRow(footerLeftGroup, "left");
        var previewCheckbox = footerLeftGroup.add("checkbox", undefined, L("checkbox.preview"));

        /* 左右を押し広げるためのスペーサー / spacer that pushes the two sides apart */
        var footerSpacer = footerRow.add("group");
        footerSpacer.alignment = ["fill", "fill"];

        var footerRightGroup = footerRow.add("group");
        setupRow(footerRightGroup, "right");
        footerRightGroup.add("button", undefined, L("button.cancel"), { name: "cancel" });
        footerRightGroup.add("button", undefined, L("button.ok"), { name: "ok" });

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

        var presetRadios = [];

        /* --- 調整変形：変形量・強度が有効 / Adjustable transform --- */
        var adjustablePanel = columnsGroup.add("panel", undefined, L("panel.adjustable"));
        setupPanel(adjustablePanel);

        var adjustablePresetsRow = adjustablePanel.add("group");
        setupRow(adjustablePresetsRow, "fill");
        adjustablePresetsRow.alignChildren = ["fill", "top"];
        addPresetGroupPanels(adjustablePresetsRow, ADJUSTABLE_GROUP_KEYS, presetRadios);

        var amountControls = buildAmountPanel(adjustablePanel);

        /* --- 固定変形：変形量・強度は使用しない / Fixed transform --- */
        var fixedPanel = columnsGroup.add("panel", undefined, L("panel.fixed"));
        setupPanel(fixedPanel);
        addPresetGroupPanels(fixedPanel, FIXED_GROUP_KEYS, presetRadios);

        var previewCheckbox = buildFooterRow(dialog);

        /* --- 初期状態 / Initial state --- */
        presetRadios[DEFAULT_PRESET_INDEX].value = true;
        previewCheckbox.value = false;
        /* テキストを含む選択では、崩れを抑えるためマイルドを初期値にする
           Default to Mild when text is selected, to keep the shapes readable */
        amountControls.mildRadio.value = !!preferMildStrength;
        amountControls.normalRadio.value = !preferMildStrength;
        amountControls.boostRadio.value = false;

        return {
            dialog: dialog,
            presetRadios: presetRadios,
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
        var presetRadios = dialogUI.presetRadios;
        var amountControls = dialogUI.amountControls;
        var previewCheckbox = dialogUI.previewCheckbox;

        /* previewApplied は「取り消すべきプレビューが残っているか」を表す。
           複数オブジェクトでも Undo は1回で足りるため、件数ではなく真偽値で管理する。
           previewApplied tracks whether a preview is still pending removal. One undo
           covers every object, so this is a flag rather than a counter. */
        var previewApplied = false;
        var lastPreviewUpdateTime = 0;
        var lastPreviewSignature = null;

        /* 選択中のプリセットを取得 / Get the selected preset */
        function getSelectedPreset() {
            for (var i = 0; i < presetRadios.length; i++) {
                if (presetRadios[i].value) return DISTORT_PRESETS[i];
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

        /* プレビューを取り消す。適用済みのときだけ Undo を1回だけ実行する
           （余分に undo するとユーザーの直前の操作まで巻き戻ってしまう）
           Remove the preview with a single undo, and only when one is pending:
           any extra undo would roll back the user's own previous edit. */
        function clearPreview() {
            if (previewApplied) {
                app.undo();
                app.redraw();
                previewApplied = false;
            }
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

            var destCorners = makeDestCorners(getSelectedPreset(), getAmountRatio(), getStrengthKey());

            try {
                applyFreeDistort(targetItems, destCorners, function () { previewApplied = true; });
            } catch (error) {
                /* 1件でも適用済みなら巻き戻す。1件目で失敗した場合は何もしない
                   Roll back if at least one object was touched; do nothing if the first failed */
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

        /* プリセットは複数パネルに分かれており ScriptUI の排他が効かないため、手動で解除する
           The presets span several panels, so ScriptUI cannot group them: clear the others by hand */
        for (var i = 0; i < presetRadios.length; i++) {
            presetRadios[i].onClick = function () {
                for (var k = 0; k < presetRadios.length; k++) {
                    if (presetRadios[k] !== this) presetRadios[k].value = false;
                }
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
    var distortSettings = showDialog();
    if (!distortSettings) return;

    var targetItems = getEffectTargets();
    if (targetItems.length === 0) {
        alert(L("alert.noTarget"));
        return;
    }

    var destCorners = makeDestCorners(
        distortSettings.preset,
        distortSettings.amountRatio,
        distortSettings.strengthKey);

    try {
        applyFreeDistort(targetItems, destCorners);
    } catch (error) {
        showError(error);
    }

})();
