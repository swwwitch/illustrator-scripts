#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

# ペア間隔調整スクリプト

オブジェクトの間隔を指定値にそろえます。2つのモードがあります。

- 自動ペア認識：選択オブジェクトを最も近いもの同士でペアにし、各ペアの間隔をそろえます
- グループ：選択した各グループの中身を並べ、隣り合う間隔をすべて指定値にそろえます（2個に限らず、3個以上もカスケードで分配）

選択がすべてグループのときは起動時に自動で「グループ」、それ以外は「自動ペア認識」が選ばれます（ダイアログで切り替え可能）。

固定オブジェクト（上 / 左 / 右 / 下）で基準にする側を選びます。左右を固定すると水平方向、上下を固定すると垂直方向の間隔をそろえます。固定した側は動かさず、残りを移動します。

さらに「整列（移動側）」で、動かす側を固定オブジェクトに対して直交方向にそろえられます。上下を固定（縦並び）なら左右方向（左 / 中央 / 右）、左右を固定（横並び）なら上下方向（上 / 中央 / 下）。間隔の調整と直交する方向だけが有効になります。

## 使い方

- 間隔をそろえたいオブジェクト、またはグループを選択して実行
- ダイアログで「モード」「固定オブジェクト（上 / 左 / 右 / 下）」「間隔」を指定（間隔は↑↓キーで±1、Shiftで±10、Optionで±0.1。マイナス値で重ねられます）
- 「固定オブジェクト」で基準にする側を選びます。左右なら水平、上下なら垂直の間隔をそろえます。グループモードでは、固定した端を起点に残りをカスケード配置します
- 「整列（移動側）」で、動かす側を固定オブジェクトの端／中央にそろえます（不要なら「なし」）
- 「クリア」ボタンで間隔を 0 にできます
- 「プレビュー境界」をオンにすると、線幅や効果を含む見た目の境界で間隔を測ります
- 設定はダイアログを閉じずにライブプレビューで確認できます

## 注意

- 自動ペア認識：選択数が奇数のときは末尾の1つがペアにならず、そのまま残ります
- グループ：子が2個未満のグループ、およびグループ以外のオブジェクトは対象外です

### 紹介記事

https://note.com/dtp_tranist/n/nc8fab19d8164

*/

// =========================================
// バージョン / Version
// =========================================
var SCRIPT_VERSION = "v1.2.2";

// =========================================
// ユーザー設定 / User settings
// =========================================
var DEFAULT_GAP = 30; // 平均間隔を測れないときのフォールバック（pt）/ Fallback gap when no average can be measured (pt)

// =========================================
// ローカライズ / Localization
// =========================================
var LABELS = {
    dialogTitle: { ja: "ペア間隔の調整", en: "Adjust Pair Gap" },
    mode: { ja: "モード", en: "Mode" },
    modeGroup: { ja: "グループ", en: "Group" },
    modeAutoPair: { ja: "自動ペア認識", en: "Auto Pair Detection" },
    fixedSide: { ja: "固定オブジェクト", en: "Fixed Object" },
    sideTop: { ja: "上", en: "Top" },
    sideLeft: { ja: "左", en: "Left" },
    sideRight: { ja: "右", en: "Right" },
    sideBottom: { ja: "下", en: "Bottom" },
    spacing: { ja: "間隔", en: "Gap" },
    clear: { ja: "クリア", en: "Clear" },
    align: { ja: "整列（移動側）", en: "Alignment (moved side)" },
    alignH: { ja: "左右方向", en: "Horizontal" },
    alignV: { ja: "上下方向", en: "Vertical" },
    alignNone: { ja: "なし", en: "None" },
    alignCenter: { ja: "中央", en: "Center" },
    options: { ja: "オプション", en: "Options" },
    previewBounds: { ja: "プレビュー境界", en: "Preview Bounds" },
    cancel: { ja: "キャンセル", en: "Cancel" },
    // ツールチップ / Tooltips
    tipModeGroup: {
        ja: "選択した各グループの中身を、隣り合う間隔がすべて同じになるように並べます（3個以上も対応）。",
        en: "Distributes the contents of each selected group so every adjacent gap is equal (3+ objects supported)."
    },
    tipModeAutoPair: {
        ja: "選択オブジェクトを最も近いもの同士でペアにし、各ペアの間隔をそろえます。",
        en: "Pairs the selected objects by nearest neighbor and sets the gap of each pair."
    },
    tipFixedSide: {
        ja: "固定した側のオブジェクトは動かさず、残りを移動して間隔をそろえます。左右で水平、上下で垂直方向にそろえます。",
        en: "Keeps the fixed side in place and moves the rest to set the gap. Left/Right adjusts horizontally, Top/Bottom vertically."
    },
    tipSpacing: {
        ja: "オブジェクト間の間隔。マイナスにすると重なります。↑↓キーで±1、Shiftで±10、Optionで±0.1。",
        en: "Gap between objects; a negative value overlaps them. Arrow keys ±1, Shift ±10, Option ±0.1."
    },
    tipClear: {
        ja: "間隔を0にします。",
        en: "Sets the gap to 0."
    },
    tipAlignH: {
        ja: "縦に並べたとき（上下を固定）、動く側を固定オブジェクトの左端／中央／右端にそろえます。",
        en: "When stacking vertically (top/bottom fixed), aligns moved objects to the fixed object's left/center/right."
    },
    tipAlignV: {
        ja: "横に並べたとき（左右を固定）、動く側を固定オブジェクトの上端／中央／下端にそろえます。",
        en: "When laying out horizontally (left/right fixed), aligns moved objects to the fixed object's top/center/bottom."
    },
    tipPreviewBounds: {
        ja: "オンで線幅や効果を含む見た目の境界、オフで幾何境界を基準に間隔を測ります。",
        en: "On: measure by visible bounds (incl. stroke/effects). Off: geometric bounds."
    },
    // アラート / Alerts
    alertNoDocument: {
        ja: "ドキュメントが開かれていません。",
        en: "No document is open."
    },
    alertSelectTwo: {
        ja: "オブジェクトを2つ以上選択してください。",
        en: "Select at least two objects."
    }
    // OK はローカライズしない / "OK" is not localized
};

var currentLanguage = ($.locale && $.locale.indexOf("ja") === 0) ? "ja" : "en";

/* キーからロケールに応じたラベルを返す / Return a localized label by key */
function getLocalizedText(key) {
    var entry = LABELS[key];
    return entry ? (entry[currentLanguage] || entry.en) : key;
}

/* コロン付きラベル（日本語は全角、英語は半角）/ Label with colon (full-width JA, half-width EN) */
function labelText(key) {
    return getLocalizedText(key) + (currentLanguage === 'ja' ? '：' : ':');
}

// =========================================
// 単位 / Units
// =========================================
// ルーラー単位を取得し、1単位あたりの pt 係数を求める / Resolve ruler unit and pt factor
var UNIT_LABELS = ["in", "mm", "pt", "pc", "cm", "Q", "px"];
var UNIT_FACTORS = [72, 72 / 25.4, 1, 12, 72 / 2.54, 72 / 25.4 * 0.25, 1];
var rulerType = app.preferences.getIntegerPreference("rulerType");
if (rulerType < 0 || rulerType > 6) { rulerType = 2; } // 不明時は pt / Fallback to pt
var rulerUnitLabel = UNIT_LABELS[rulerType];
var pointsPerUnit = UNIT_FACTORS[rulerType]; // 1単位 = pointsPerUnit pt

// =========================================
// ユーティリティ / Utility
// =========================================
/* テキストフィールドで↑↓キーによる増減を有効化する / Enable arrow-key value stepping
   onUpdate を渡すと値変更後に呼ぶ（プレビュー更新用）/ onUpdate runs after each change */
function changeValueByArrowKey(editText, onUpdate) {
    editText.addEventListener("keydown", function (event) {
        // ↑↓キー以外（Enter・数字入力など）は通常処理に任せる。
        // ここで毎回 onUpdate を呼ぶと、Enter 等でもプレビューが走り不安定になる。
        // Only handle Up/Down; let other keys (Enter, digits...) fall through.
        if (event.keyName !== "Up" && event.keyName !== "Down") return;

        var value = Number(editText.text);
        if (isNaN(value)) return;

        var keyboard = ScriptUI.environment.keyboardState;
        var isUp = (event.keyName === "Up");

        if (keyboard.shiftKey) {
            // Shift：±10、10の倍数にスナップ / Shift: ±10 snapped to multiples of 10
            var step = 10;
            value = isUp ? Math.ceil((value + 1) / step) * step
                : Math.floor((value - 1) / step) * step;
        } else if (keyboard.altKey) {
            // Option：±0.1 / Option: ±0.1
            value = isUp ? value + 0.1 : value - 0.1;
        } else {
            // 通常：±1 / Default: ±1
            value = isUp ? value + 1 : value - 1;
        }

        // 負の値も許容（オブジェクトを重ねる）/ Allow negative values (overlap objects)

        // 丸め（Option は小数第1位、その他は整数）/ Round (Option: 1 decimal, else integer)
        value = keyboard.altKey ? Math.round(value * 10) / 10 : Math.round(value);

        event.preventDefault();
        editText.text = value;

        // プレビューを更新 / Refresh live preview
        if (typeof onUpdate === "function") {
            onUpdate();
        } else {
            editText.notify("onChanging");
        }
    });
}

/* パネルの余白と間隔 / Panel margins and spacing */
var PANEL_MARGINS = [16, 20, 16, 12];
var PANEL_SPACING = 8;

/* パネルの共通設定 / Apply shared panel layout */
function setupPanel(panel, spacing) {
    panel.orientation = "column";
    panel.alignChildren = ["fill", "top"];
    panel.alignment = "fill";
    panel.margins = PANEL_MARGINS;
    panel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
}

/* グループの共通設定（orientation は呼び出し側で指定）/ Apply shared group layout (orientation passed in) */
function setupGroup(group, orientation, spacing) {
    group.orientation = orientation || "column";
    group.alignChildren = ["left", "center"];
    group.alignment = "fill";
    group.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
}

/* 固定オブジェクト（上 / 左 / 右 / 下）パネルを生成する。
   上下左右を厳密な3×3グリッドに配置する（中央・四隅は空セル）。
   セルを固定幅にして列を確実にそろえる。4つのラジオは親が別なので自動グループ化されない。
   既定は「右」を固定（左が動く＝水平）。イベント結線は呼び出し側で行う。
   Build the Fixed Object panel (top/left/right/bottom on a strict 3×3 grid; center &
   corners are empty cells, fixed-width cells keep columns aligned). The four radios live
   in different parents, so ScriptUI won't group them — exclusivity is enforced by hand.
   Default fixes the right side (horizontal). Event wiring is left to the caller.
   返り値 / Returns: { panel, fixedRadios, selectFixedRadio, getFixedSide } */
function buildFixedSidePanel(parentGroup) {
    var panel = parentGroup.add("panel", undefined, getLocalizedText('fixedSide'));
    setupPanel(panel, 2);
    // 十字レイアウトの左右に余白を足す（+8）/ Add left/right margin to the cross layout (+8)
    panel.margins = [PANEL_MARGINS[0] + 8, PANEL_MARGINS[1], PANEL_MARGINS[2] + 8, PANEL_MARGINS[3]];
    panel.alignChildren = ["center", "top"]; // 各行を中央そろえで十字に / Center each row

    //   ・  上  ・
    //   左  ・  右
    //   ・  下  ・
    var GRID_CELL = [22, 20]; // セルの幅・高さ(px) / cell width, height (px)

    /* グリッドの1セルを作る。withRadio が真ならラジオを入れて返す / Build one grid cell */
    function addGridCell(rowGroup, withRadio) {
        var cell = rowGroup.add("group");
        cell.margins = 0;
        cell.alignChildren = ["center", "center"];
        cell.preferredSize = GRID_CELL;
        return withRadio ? cell.add("radiobutton", undefined, "") : null;
    }

    /* グリッドの1行（3セル）を作る / Build one grid row (3 cells) */
    function addGridRow() {
        var row = panel.add("group");
        row.orientation = "row";
        row.alignment = "center";
        row.spacing = 0;
        row.margins = 0;
        return row;
    }

    var gridRowTop = addGridRow();
    addGridCell(gridRowTop, false);
    var topRadio = addGridCell(gridRowTop, true);       // 上 / Top
    addGridCell(gridRowTop, false);

    var gridRowMid = addGridRow();
    var leftRadio = addGridCell(gridRowMid, true);      // 左 / Left
    addGridCell(gridRowMid, false);                     // 中央は空 / center empty
    var rightRadio = addGridCell(gridRowMid, true);     // 右 / Right

    var gridRowBottom = addGridRow();
    addGridCell(gridRowBottom, false);
    var bottomRadio = addGridCell(gridRowBottom, true); // 下 / Bottom
    addGridCell(gridRowBottom, false);

    var fixedRadios = [topRadio, leftRadio, rightRadio, bottomRadio];

    /* 指定したラジオだけ ON にして排他制御する / Turn on only the chosen radio (manual exclusivity) */
    function selectFixedRadio(chosen) {
        for (var fi = 0; fi < fixedRadios.length; fi++) {
            fixedRadios[fi].value = (fixedRadios[fi] === chosen);
        }
    }
    for (var ti = 0; ti < fixedRadios.length; ti++) {
        fixedRadios[ti].helpTip = getLocalizedText('tipFixedSide');
    }
    rightRadio.value = true; // 既定：右を固定（左が動く＝水平）/ Default: fix right (horizontal)

    /* 現在固定する側を取得する / Get the currently fixed side */
    function getFixedSide() {
        if (topRadio.value) return "top";
        if (leftRadio.value) return "left";
        if (bottomRadio.value) return "bottom";
        return "right";
    }

    /* 固定する側を設定する（"top"/"left"/"right"/"bottom"）。未知の値は無視 / Set the fixed side */
    function setFixedSide(side) {
        var radioBySide = { top: topRadio, left: leftRadio, right: rightRadio, bottom: bottomRadio };
        if (radioBySide[side]) selectFixedRadio(radioBySide[side]);
    }

    return {
        panel: panel,
        fixedRadios: fixedRadios,
        selectFixedRadio: selectFixedRadio,
        getFixedSide: getFixedSide,
        setFixedSide: setFixedSide
    };
}

/* オプションパネル（間隔の入力・クリア・プレビュー境界）を生成する。
   間隔は現在のルーラー単位で表示し、getSpacingInPoints() が pt に換算して返す。
   initialGapPoints は初期値・入力が空のときのフォールバック（pt）。
   イベント結線（↑↓キー・クリア・プレビュー更新）は呼び出し側で行う。
   Build the Options panel (gap input, clear, preview-bounds). The gap is shown in the
   current ruler unit; getSpacingInPoints() converts back to pt. initialGapPoints is the
   default and the fallback when the field is empty. Event wiring is left to the caller.
   返り値 / Returns: { panel, spacingInput, clearButton, previewBoundsCheckbox,
                       getSpacingInPoints, getBoundsType } */
function buildSpacingOptionsPanel(parentGroup, initialGapPoints) {
    var panel = parentGroup.add("panel", undefined, getLocalizedText('options'));
    setupPanel(panel, 6);

    // 行：間隔（ラベル＋入力＋単位）/ Row: gap (label + input + unit)
    var spacingRow = panel.add("group");
    setupGroup(spacingRow, "row");
    spacingRow.alignment = "left"; // 広げず左寄せ / Keep at natural width, packed left
    spacingRow.add("statictext", undefined, labelText('spacing'));
    var defaultSpacingDisplay = Math.round((initialGapPoints / pointsPerUnit) * 100) / 100;
    var spacingInput = spacingRow.add("edittext", undefined, String(defaultSpacingDisplay));
    spacingInput.characters = 4;
    spacingInput.helpTip = getLocalizedText('tipSpacing');
    spacingRow.add("statictext", undefined, rulerUnitLabel);

    // クリアボタン：間隔を 0 にする / Clear button: set the gap to 0
    var clearButton = spacingRow.add("button", undefined, getLocalizedText('clear'));
    clearButton.helpTip = getLocalizedText('tipClear');

    // チェックボックス：プレビュー境界（groupに入れて左添え）/ Checkbox in a left-aligned group
    var previewBoundsGroup = panel.add("group");
    previewBoundsGroup.orientation = "row";
    previewBoundsGroup.alignment = "left";   // パネル内で左添え / Left within panel
    previewBoundsGroup.margins = [0, 5, 0, 0]; // 上マージン5 / Top margin 5
    var previewBoundsCheckbox = previewBoundsGroup.add("checkbox", undefined, getLocalizedText('previewBounds'));
    previewBoundsCheckbox.value = false; // OFF=幾何境界 / ON=プレビュー境界
    previewBoundsCheckbox.helpTip = getLocalizedText('tipPreviewBounds');

    /* 入力値を pt に換算して取得する / Get the gap in pt from the input */
    function getSpacingInPoints() {
        var value = parseFloat(spacingInput.text);
        if (isNaN(value)) { value = initialGapPoints / pointsPerUnit; }
        return value * pointsPerUnit;
    }

    /* チェックに応じた境界プロパティ名を返す / Bounds property per checkbox */
    function getBoundsType() {
        return previewBoundsCheckbox.value ? "visibleBounds" : "geometricBounds";
    }

    return {
        panel: panel,
        spacingInput: spacingInput,
        clearButton: clearButton,
        previewBoundsCheckbox: previewBoundsCheckbox,
        getSpacingInPoints: getSpacingInPoints,
        getBoundsType: getBoundsType
    };
}

/* 整列パネル（移動側の整列）を生成する。左右方向と上下方向の2グループを1つのパネルにまとめる。
   行ごとに独立したラジオグループ（なし / 先頭 / 中央 / 後ろ → "none"/"start"/"center"/"end"）。
   整列はギャップ軸に直交する方向：上下固定→左右整列、左右固定→上下整列。該当しない側の行は無効化する。
   現在の固定側に依存するので getFixedSide を受け取り、updateAlignPanels()/getAlignMode() がそれを参照する。
   イベント結線（onClick→プレビュー更新）は呼び出し側で行う。
   Build the Alignment panel (one panel with two independent radio groups). Alignment is
   perpendicular to the gap axis; the inapplicable row is disabled. Takes getFixedSide since
   the active row depends on the fixed side. Event wiring is left to the caller.
   返り値 / Returns: { panel, hAlign, vAlign, updateAlignPanels, getAlignMode } */
function buildAlignmentPanel(parentGroup, getFixedSide) {
    var panel = parentGroup.add("panel", undefined, getLocalizedText('align'));
    setupPanel(panel, 6);

    /* 整列1行（ラベル＋なし/先頭/中央/後ろ）を作る / Build one alignment row (label + radios) */
    function buildAlignRow(labelKey, startKey, endKey, tipKey) {
        var row = panel.add("group");
        setupGroup(row, "row");
        row.alignment = "left";              // 広げず左寄せ / Keep at natural width, packed left
        row.alignChildren = ["left", "center"];
        var label = row.add("statictext", undefined, labelText(labelKey));
        var radios = {
            none: row.add("radiobutton", undefined, getLocalizedText('alignNone')),
            start: row.add("radiobutton", undefined, getLocalizedText(startKey)),
            center: row.add("radiobutton", undefined, getLocalizedText('alignCenter')),
            end: row.add("radiobutton", undefined, getLocalizedText(endKey))
        };
        radios.none.value = true; // 既定：整列なし / Default: no alignment
        var tip = getLocalizedText(tipKey);
        for (var key in radios) { radios[key].helpTip = tip; }
        return { row: row, label: label, radios: radios };
    }

    // 左右方向の整列（上下固定で縦に並べたとき有効）/ Horizontal align (active when stacking vertically)
    var hAlign = buildAlignRow('alignH', 'sideLeft', 'sideRight', 'tipAlignH');
    // 上下方向の整列（左右固定で横に並べたとき有効）/ Vertical align (active when laying out horizontally)
    var vAlign = buildAlignRow('alignV', 'sideTop', 'sideBottom', 'tipAlignV');

    // ラベル幅をそろえて2行のラジオ開始位置を合わせる / Match label widths so radios line up
    var alignLabelWidth = Math.max(hAlign.label.preferredSize.width, vAlign.label.preferredSize.width);
    hAlign.label.preferredSize.width = alignLabelWidth;
    vAlign.label.preferredSize.width = alignLabelWidth;

    /* 上下固定ならギャップは垂直 / Gap is vertical when the top/bottom side is fixed */
    function isVerticalGap() {
        var side = getFixedSide();
        return side === "top" || side === "bottom";
    }
    /* 固定側に合わせて該当する整列行だけ有効化する / Enable only the applicable alignment row */
    function updateAlignPanels() {
        var vertical = isVerticalGap();
        hAlign.row.enabled = vertical;   // 左右整列は縦並びのとき / Horizontal align when stacking vertically
        vAlign.row.enabled = !vertical;  // 上下整列は横並びのとき / Vertical align when laying out horizontally
    }
    /* 現在有効な整列パネルの値を "none"/"start"/"center"/"end" で返す / Active alignment value */
    function getAlignMode() {
        var radios = isVerticalGap() ? hAlign.radios : vAlign.radios;
        if (radios.start.value) return "start";
        if (radios.center.value) return "center";
        if (radios.end.value) return "end";
        return "none";
    }

    return {
        panel: panel,
        hAlign: hAlign,
        vAlign: vAlign,
        updateAlignPanels: updateAlignPanels,
        getAlignMode: getAlignMode
    };
}

// =========================================
// 軸 / Axis
// =========================================
/* 水平 / 垂直を共通の「進行方向の座標 p」に抽象化する。p は増えるほど右（水平）または下（垂直）。
   これで間隔調整のロジックを 1 本で両軸に使える。geometricBounds = [左, 上, 右, 下]。
   Abstracts horizontal/vertical into a single "travel coordinate p" that increases
   rightward (horizontal) or downward (vertical), so the spacing logic works on both axes.
   start = 進行方向の先頭側の辺、end = 後ろ側の辺 / start = leading edge, end = trailing edge.
   垂直は Illustrator の y が上ほど大きいので符号を反転して「下方向で増加」にそろえる。
   Vertical negates y (Illustrator y grows upward) so p still increases in the travel direction. */
function makeAxis(isVertical) {
    if (isVertical) {
        return {
            vertical: true,
            start: function (bounds) { return -bounds[1]; }, // 上端 / top edge
            end: function (bounds) { return -bounds[3]; }  // 下端 / bottom edge
        };
    }
    return {
        vertical: false,
        start: function (bounds) { return bounds[0]; }, // 左端 / left edge
        end: function (bounds) { return bounds[2]; }  // 右端 / right edge
    };
}

/* 固定する側から軸を求める（上下→垂直、左右→水平）/ Axis from the fixed side (top/bottom → vertical) */
function axisForSide(fixedSide) {
    return makeAxis(fixedSide === "top" || fixedSide === "bottom");
}

/* 固定側が進行方向の「後ろ側」(右 or 下) かどうか / Whether the fixed side is the trailing end (right/bottom) */
function isAnchorEnd(fixedSide) {
    return fixedSide === "right" || fixedSide === "bottom";
}

// =========================================
// ペアリング / Pairing
// =========================================
/* 選択オブジェクトを最も近いもの同士でペアにする / Pair nearest objects.
   ペアリングは常に geometricBounds（中心）で判定する。
   「プレビュー境界」を ON にしても間隔計算の基準が変わるだけで、
   ペアの組み合わせ自体は変わらない。
   Pairing always uses geometricBounds (centers); the Preview Bounds option
   only changes how the gap is measured, never which objects are paired. */
function createNearestPairs(selectedItems) {
    /* オブジェクトの中心座標を取得する（常に幾何境界。クリップグループはクリッピングパス基準）
       Object center (always geometric bounds; clip groups use their clipping path) */
    function getObjectCenter(item) {
        var bounds = geometricBoundsOf(item);
        return {
            x: (bounds[0] + bounds[2]) / 2,
            y: (bounds[1] + bounds[3]) / 2
        };
    }

    /* 2点間の距離を計算する / Distance between two points */
    function getPointDistance(point1, point2) {
        var dx = point1.x - point2.x;
        var dy = point1.y - point2.y;
        return Math.sqrt(dx * dx + dy * dy);
    }

    // 中心座標付きの作業リスト / Working list with centers
    var remainingItems = [];
    for (var i = 0; i < selectedItems.length; i++) {
        remainingItems.push({
            obj: selectedItems[i],
            center: getObjectCenter(selectedItems[i])
        });
    }

    var pairs = [];
    while (remainingItems.length > 0) {
        var currentItem = remainingItems.shift();
        var nearestIndex = -1;
        var nearestDistance = Infinity;
        for (var j = 0; j < remainingItems.length; j++) {
            var distance = getPointDistance(currentItem.center, remainingItems[j].center);
            if (distance < nearestDistance) {
                nearestDistance = distance;
                nearestIndex = j;
            }
        }
        if (nearestIndex !== -1) {
            var partnerItem = remainingItems.splice(nearestIndex, 1)[0];
            pairs.push({ a: currentItem.obj, b: partnerItem.obj });
        }
    }
    return pairs;
}

/* GroupItem の直下オブジェクトを配列で返す。グループ以外は null。
   Return a GroupItem's direct children as an array; null for non-groups. */
function getGroupChildren(item) {
    if (item.typename !== "GroupItem") return null;
    var children = [];
    for (var i = 0; i < item.pageItems.length; i++) {
        children.push(item.pageItems[i]);
    }
    return children;
}

/* クリップグループならクリッピングパス（マスク）を返す。クリップグループでなければ null。
   クリップグループの geometricBounds／visibleBounds はマスクで隠れた部分まで含むことがあり、
   間隔計算がずれる。基準をマスクの形（クリッピングパス）にそろえるために使う。
   Return the clipping path (mask) if the item is a clip group, else null. A clip group's
   bounds can include artwork hidden by the mask, throwing off gap math, so we measure
   against the mask shape instead. */
function getClippingPath(item) {
    if (item.typename !== "GroupItem" || !item.clipped) return null;
    var children = item.pageItems;
    for (var i = 0; i < children.length; i++) {
        var child = children[i];
        if (child.typename === "PathItem" && child.clipping) return child;
        // 複合パスのマスクは先頭サブパスの clipping で判定 / Compound-path mask: check the first subpath
        if (child.typename === "CompoundPathItem" &&
            child.pathItems.length > 0 && child.pathItems[0].clipping) return child;
    }
    return null;
}

/* 間隔計算に使う幾何境界。クリップグループはクリッピングパスを基準にする。
   Geometric bounds for gap math; clip groups measure against their clipping path. */
function geometricBoundsOf(item) {
    var clip = getClippingPath(item);
    return clip ? clip.geometricBounds : item.geometricBounds;
}

/* 間隔計算に使うプレビュー境界。クリップグループはクリッピングパスを基準にする。
   Visible bounds for gap math; clip groups measure against their clipping path. */
function visibleBoundsOf(item) {
    var clip = getClippingPath(item);
    return clip ? clip.visibleBounds : item.visibleBounds;
}

/* 選択した各グループの一番左・一番右を取り出してペアにする / Pair the leftmost and
   rightmost object inside each selected group.
   ペア認識の代わりにグループ単位で組む。左右は geometricBounds の左端 x で判定し、
   どちらを固定するかは固定オブジェクト（左/右）ラジオに従う（applySpacing 側で処理）。
   Built per group instead of nearest-neighbor pairing; which side stays fixed is
   decided by the Fixed Object (left/right) radio, handled in applySpacing. */
function createGroupPairs(selectedItems) {
    var pairs = [];
    for (var i = 0; i < selectedItems.length; i++) {
        var children = getGroupChildren(selectedItems[i]);
        // グループでない、または子が2個未満なら間隔を調整できない / Need a group with 2+ children
        if (!children || children.length < 2) continue;

        // グループは2つに限らない。含まれる全オブジェクトを members として持ち、
        // applySpacing 側で左端順に等間隔へ分配する（固定側を基準に配置）。
        // A group may hold any number of objects; keep them all as members and let
        // applySpacing distribute them with a uniform gap, anchored at the fixed side.
        pairs.push({ members: children });
    }
    return pairs;
}

/* 選択オブジェクトの現在の間隔の平均（pt）を求める。測れない場合は null。axis で測る軸を指定。
   モードに合わせて測る：自動ペア認識は各ペアの間隔、グループは各グループ内の隣接間隔。
   常に geometricBounds（幾何境界）基準。Average current gap (pt) along the given axis,
   matching the mode; null if nothing measurable. Always geometric bounds. */
function computeAverageGap(selectedItems, mode, axis) {
    var gaps = [];

    if (mode === "group") {
        // 各グループ内：進行方向順に並べて隣り合う間隔を測る / Adjacent gaps inside each group
        for (var i = 0; i < selectedItems.length; i++) {
            var children = getGroupChildren(selectedItems[i]);
            if (!children || children.length < 2) continue;
            var sorted = children.slice(0);
            sorted.sort(function (a, b) { return axis.start(geometricBoundsOf(a)) - axis.start(geometricBoundsOf(b)); });
            for (var j = 1; j < sorted.length; j++) {
                // 次の先頭辺 - 前の後ろ辺 / next leading edge - previous trailing edge
                gaps.push(axis.start(geometricBoundsOf(sorted[j])) - axis.end(geometricBoundsOf(sorted[j - 1])));
            }
        }
    } else {
        // 自動ペア認識：各ペアの間隔を測る / Gap of each nearest pair
        var pairs = createNearestPairs(selectedItems);
        for (var p = 0; p < pairs.length; p++) {
            var boundsA = geometricBoundsOf(pairs[p].a);
            var boundsB = geometricBoundsOf(pairs[p].b);
            var leadBounds = (axis.start(boundsA) < axis.start(boundsB)) ? boundsA : boundsB;
            var trailBounds = (axis.start(boundsA) < axis.start(boundsB)) ? boundsB : boundsA;
            gaps.push(axis.start(trailBounds) - axis.end(leadBounds));
        }
    }

    if (gaps.length === 0) return null;
    var sum = 0;
    for (var g = 0; g < gaps.length; g++) sum += gaps[g];
    return sum / gaps.length;
}

// =========================================
// 設定の保存 / Persistence
// =========================================
// ダイアログを閉じるときに「固定側・プレビュー境界・整列」だけを覚え、次回開いたときに復元する。
// モードと間隔は選択内容から毎回決めるので保存しない。Folder.userData に key=value 形式で1ファイル。
// Remember only the fixed side, preview-bounds, and alignment on close; restore them next time.
// Mode and gap are re-derived from the selection each run, so they are not saved.
var SETTINGS_FILE = Folder.userData + "/AdjustPairGapSettings.txt";

/* 設定ファイルを key=value 形式で読み、オブジェクトで返す。無ければ空オブジェクト。
   Read the settings file (key=value lines) into an object; empty object if missing. */
function loadSettings() {
    var settings = {};
    var file = new File(SETTINGS_FILE);
    if (!file.exists || !file.open("r")) return settings;
    var content = file.read();
    file.close();
    var lines = content.split("\n");
    for (var i = 0; i < lines.length; i++) {
        var eq = lines[i].indexOf("=");
        if (eq <= 0) continue;
        var key = lines[i].substring(0, eq).replace(/^\s+|\s+$/g, "");
        var value = lines[i].substring(eq + 1).replace(/^\s+|\s+$/g, "");
        settings[key] = value;
    }
    return settings;
}

/* 設定を key=value 形式で書き出す / Persist settings as key=value lines */
function saveSettings(settings) {
    var file = new File(SETTINGS_FILE);
    if (!file.open("w")) return;
    var lines = [];
    for (var key in settings) {
        if (settings.hasOwnProperty(key)) lines.push(key + "=" + settings[key]);
    }
    file.write(lines.join("\n"));
    file.close();
}

/* 整列行で選択中の値（none/start/center/end）を返す / Selected value of an alignment row */
function getAlignRowValue(alignRow) {
    var radios = alignRow.radios;
    if (radios.start.value) return "start";
    if (radios.center.value) return "center";
    if (radios.end.value) return "end";
    return "none";
}

/* 保存済みの整列値を行のラジオへ反映する。未知の値は無視 / Apply a saved alignment value (ignore unknown) */
function applySavedAlign(alignRow, value) {
    if (!alignRow || !value) return;
    var radio = alignRow.radios[value];
    if (radio) radio.value = true;
}

(function () {
    if (app.documents.length === 0) {
        alert(getLocalizedText('alertNoDocument'));
        return;
    }

    var activeDocument = app.activeDocument;

    // doc.selection を配列にコピーしておく（後の選択変更や undo の影響を受けないように）
    // Copy doc.selection into a plain array (immune to later selection changes / undo)
    var liveSelection = activeDocument.selection;
    // selection が無効・未選択のケースを先に弾く（Illustrator では稀に null になる）
    // Guard against an invalid / empty selection (selection can be null in rare cases)
    if (!liveSelection || liveSelection.length === 0) {
        alert(getLocalizedText('alertSelectTwo'));
        return;
    }
    var selectedItems = [];
    for (var i = 0; i < liveSelection.length; i++) {
        selectedItems.push(liveSelection[i]);
    }

    if (selectedItems.length < 2) {
        alert(getLocalizedText('alertSelectTwo'));
        return;
    }

    // モード切り替え時に組み直すペアを保持する / Holds pairs rebuilt when the mode changes
    var objectPairs = [];

    // 選択がすべてグループなら既定でグループモードにする（それ以外は自動ペア認識）。
    // Default to group mode when every selected object is a group; otherwise auto pair.
    var selectionIsGroupsOnly = true;
    for (var gi = 0; gi < selectedItems.length; gi++) {
        if (selectedItems[gi].typename !== "GroupItem") { selectionIsGroupsOnly = false; break; }
    }
    var defaultMode = selectionIsGroupsOnly ? "group" : "auto";

    // 間隔の初期値は選択オブジェクトの現在の平均間隔。測れなければ DEFAULT_GAP を使う。
    // 負（重なり）の場合は 0 にクランプ。Initial gap = current average gap of the
    // selection (clamped to >= 0); fall back to DEFAULT_GAP when nothing measurable.
    // 既定の固定側は「右」＝水平軸。初期間隔は水平で測る / Default fixed side is right → horizontal axis
    var measuredGap = computeAverageGap(selectedItems, defaultMode, makeAxis(false));
    var initialGapPoints = (measuredGap !== null) ? Math.max(0, measuredGap) : DEFAULT_GAP;

    // =========================================
    // プレビュー / Live preview
    //   巻き戻しは移動量の逆適用で行う（app.undo() は使わない）。
    //   app.undo() をキーボードイベント内で同期実行すると Illustrator が
    //   不安定になり得るため、ここでは記録した移動を translate で打ち消す。
    //   Revert by reversing recorded moves — never app.undo(), which is
    //   unstable when called synchronously inside keyboard event handlers.
    // =========================================
    var appliedMoves = []; // 適用済みの移動 / Applied moves: { obj, dx, dy }

    /* 記録した移動を逆向きに適用して元に戻す / Reverse recorded moves */
    function undoPreview() {
        for (var m = appliedMoves.length - 1; m >= 0; m--) {
            appliedMoves[m].obj.translate(-appliedMoves[m].dx, -appliedMoves[m].dy);
        }
        appliedMoves = [];
    }

    /* 軸に沿って p だけ移動し、巻き戻し用に記録する。p は右（水平）/ 下（垂直）で増加。
       Move an object by p along the axis and record it for undo (p increases right/down). */
    function moveByAxis(axis, obj, p) {
        if (p === 0) return;
        // 垂直は Illustrator の y が上ほど大きいので、下方向(+p)へは translate(0, -p)
        // Vertical: Illustrator y grows upward, so moving down (+p) means translate(0, -p)
        var dx = axis.vertical ? 0 : p;
        var dy = axis.vertical ? -p : 0;
        obj.translate(dx, dy);
        appliedMoves.push({ obj: obj, dx: dx, dy: dy });
    }

    /* グループ内の全オブジェクトを進行方向順に並べ、隣り合う間隔を gap にそろえる。
       固定側のオブジェクトは動かさず、そこを起点にカスケードで再配置する。
       axis で水平/垂直を切り替える（左右→水平、上下→垂直）。
       Distribute all objects in a group along the axis with a uniform gap, anchored at
       the fixed side (the fixed-side object stays; the rest cascade from it). */
    function distributeGroup(pair, fixedSide, gapInPoints, useVisible, axis) {
        var members = pair.members;
        var cachedBounds = useVisible ? pair.memberVis : pair.memberGeo;
        var count = members.length;
        var anchorEnd = isAnchorEnd(fixedSide);

        // 先頭辺の昇順に並べたインデックス / Indices ordered by leading edge ascending
        var order = [];
        for (var i = 0; i < count; i++) order.push(i);
        order.sort(function (x, y) { return axis.start(cachedBounds[x]) - axis.start(cachedBounds[y]); });

        if (anchorEnd) {
            // 後ろ側（右 or 下）を固定し、後ろから前へ配置 / Anchor the trailing end; walk backward
            var nextStart = axis.start(cachedBounds[order[count - 1]]); // 後ろ端オブジェクトの先頭辺（不動）
            for (var r = count - 2; r >= 0; r--) {
                var idxR = order[r];
                var desiredEnd = nextStart - gapInPoints;
                var pR = desiredEnd - axis.end(cachedBounds[idxR]);
                moveByAxis(axis, members[idxR], pR);
                nextStart = axis.start(cachedBounds[idxR]) + pR; // この要素の新しい先頭辺 / its new leading edge
            }
        } else {
            // 先頭側（左 or 上）を固定し、前から後ろへ配置 / Anchor the leading end; walk forward
            var prevEnd = axis.end(cachedBounds[order[0]]); // 先頭端オブジェクトの後ろ辺（不動）
            for (var f = 1; f < count; f++) {
                var idxF = order[f];
                var desiredStart = prevEnd + gapInPoints;
                var pF = desiredStart - axis.start(cachedBounds[idxF]);
                moveByAxis(axis, members[idxF], pF);
                prevEnd = axis.end(cachedBounds[idxF]) + pF; // この要素の新しい後ろ辺 / its new trailing edge
            }
        }
    }

    /* オブジェクトを alignAxis 方向で anchor にそろえる。alignMode は "none"/"start"/"center"/"end"。
       start = 先頭辺（左 or 上）、end = 後ろ辺（右 or 下）、center = 中央。
       ギャップ調整は gap 軸方向のみ動かすので、整列軸の位置はキャッシュ境界のまま使える。
       Align obj to anchor along alignAxis. Gap moves only along the gap axis, so the
       align-axis position is unchanged and cached bounds stay valid. */
    function alignToAnchor(alignAxis, obj, objBounds, anchorBounds, alignMode) {
        if (alignMode === "none") return;
        var delta;
        if (alignMode === "start") {
            delta = alignAxis.start(anchorBounds) - alignAxis.start(objBounds);
        } else if (alignMode === "end") {
            delta = alignAxis.end(anchorBounds) - alignAxis.end(objBounds);
        } else { // center
            var anchorCenter = (alignAxis.start(anchorBounds) + alignAxis.end(anchorBounds)) / 2;
            var objCenter = (alignAxis.start(objBounds) + alignAxis.end(objBounds)) / 2;
            delta = anchorCenter - objCenter;
        }
        moveByAxis(alignAxis, obj, delta);
    }

    /* グループの各メンバーを固定端のメンバー（アンカー）に整列軸方向でそろえる。
       Align every group member to the anchor (the member at the fixed end) along alignAxis. */
    function alignGroup(pair, gapAxis, anchorEnd, useVisible, alignAxis, alignMode) {
        if (alignMode === "none") return;
        var members = pair.members;
        var cachedBounds = useVisible ? pair.memberVis : pair.memberGeo;
        var count = members.length;

        var order = [];
        for (var i = 0; i < count; i++) order.push(i);
        order.sort(function (x, y) { return gapAxis.start(cachedBounds[x]) - gapAxis.start(cachedBounds[y]); });

        var anchorIdx = anchorEnd ? order[count - 1] : order[0]; // 固定端のメンバー / member at the fixed end
        var anchorBounds = cachedBounds[anchorIdx];
        for (var k = 0; k < count; k++) {
            if (k === anchorIdx) continue;
            alignToAnchor(alignAxis, members[k], cachedBounds[k], anchorBounds, alignMode);
        }
    }

    /* 設定値で各ペアの間隔を調整する / Apply the gap to every pair.
       固定側から軸（水平/垂直）と固定端（先頭/後ろ）を決める / Axis and anchor end follow the fixed side.
       alignMode は整列（ギャップ軸に直交する方向、固定オブジェクト基準）/ alignMode aligns on the
       axis perpendicular to the gap, anchored at the fixed object. */
    function applySpacing(fixedSide, gapInPoints, boundsType, alignMode) {
        var useVisible = (boundsType === "visibleBounds");
        var axis = axisForSide(fixedSide);
        var anchorEnd = isAnchorEnd(fixedSide);
        var alignAxis = makeAxis(!axis.vertical); // 整列はギャップ軸に直交 / perpendicular to the gap axis

        for (var k = 0; k < objectPairs.length; k++) {
            var pair = objectPairs[k];

            if (pair.members) {
                // グループモード：全オブジェクトを等間隔に分配（固定側を基準）、その後に整列
                // Group mode: distribute with a uniform gap (anchored at the fixed side), then align
                distributeGroup(pair, fixedSide, gapInPoints, useVisible, axis);
                alignGroup(pair, axis, anchorEnd, useVisible, alignAxis, alignMode);
                continue;
            }

            // 自動ペア認識：2オブジェクトの間隔を調整 / Auto mode: adjust the gap of a pair
            // キャッシュ済みの元の境界を使用 / Use cached original bounds
            var boundsA = useVisible ? pair.visibleA : pair.geometricA;
            var boundsB = useVisible ? pair.visibleB : pair.geometricB;

            // 進行方向の先頭辺で前後を判定 / Decide leading/trailing by the axis start edge
            var leadObject, trailObject, leadBounds, trailBounds;
            if (axis.start(boundsA) < axis.start(boundsB)) {
                leadObject = pair.a; leadBounds = boundsA;
                trailObject = pair.b; trailBounds = boundsB;
            } else {
                leadObject = pair.b; leadBounds = boundsB;
                trailObject = pair.a; trailBounds = boundsA;
            }

            // 固定側を基準に動く側だけ移動し、同じ動く側を整列軸でも固定側へそろえる。
            // Move only the non-fixed object for the gap, then align that same object to the fixed one.
            var p, movedObject, movedBounds, anchorBounds;
            if (anchorEnd) {
                // 後ろ側（右 or 下）を固定：先頭側オブジェクトを動かす / Fix trailing end, move the leading object
                p = (axis.start(trailBounds) - gapInPoints) - axis.end(leadBounds);
                moveByAxis(axis, leadObject, p);
                movedObject = leadObject; movedBounds = leadBounds; anchorBounds = trailBounds;
            } else {
                // 先頭側（左 or 上）を固定：後ろ側オブジェクトを動かす / Fix leading end, move the trailing object
                p = (axis.end(leadBounds) + gapInPoints) - axis.start(trailBounds);
                moveByAxis(axis, trailObject, p);
                movedObject = trailObject; movedBounds = trailBounds; anchorBounds = leadBounds;
            }
            alignToAnchor(alignAxis, movedObject, movedBounds, anchorBounds, alignMode);
        }
    }

    /* 直前のプレビューを巻き戻してから再適用する / Re-run preview from scratch */
    function runPreview(fixedSide, gapInPoints, boundsType, alignMode) {
        undoPreview();
        applySpacing(fixedSide, gapInPoints, boundsType, alignMode);
        app.redraw();
    }

    /* 各ペアの元の境界をキャッシュする / Cache original bounds of every pair.
       適用時は常に元位置へ巻き戻してから計算するので、ここで一度取れば使い回せる。 */
    function cachePairBounds(pairs) {
        for (var p = 0; p < pairs.length; p++) {
            var pair = pairs[p];
            if (pair.members) {
                // グループ：全メンバーの境界をキャッシュ / Group: cache every member's bounds
                pair.memberGeo = [];
                pair.memberVis = [];
                for (var m = 0; m < pair.members.length; m++) {
                    pair.memberGeo.push(geometricBoundsOf(pair.members[m]));
                    pair.memberVis.push(visibleBoundsOf(pair.members[m]));
                }
            } else {
                pair.geometricA = geometricBoundsOf(pair.a);
                pair.visibleA = visibleBoundsOf(pair.a);
                pair.geometricB = geometricBoundsOf(pair.b);
                pair.visibleB = visibleBoundsOf(pair.b);
            }
        }
    }

    /* モードに応じてペアを組み直す / Rebuild pairs for the given mode ("auto" | "group").
       境界キャッシュは必ず元位置で取るため、先にプレビューを巻き戻してから組む。
       Always revert the preview first so cached bounds reflect original positions. */
    function buildPairs(mode) {
        undoPreview();
        var pairs = (mode === "group")
            ? createGroupPairs(selectedItems)
            : createNearestPairs(selectedItems);
        cachePairBounds(pairs);
        objectPairs = pairs;
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================
    /* ダイアログを生成して表示し、OK されたら true を返す / Build, show dialog; return true on OK */
    function showSettingsDialog() {
        var dialog = new Window("dialog", getLocalizedText('dialogTitle') + ' ' + SCRIPT_VERSION);
        dialog.orientation = "column";
        dialog.alignChildren = "fill";

        // モードと固定オブジェクトを横並び（2カラム）に置く / Mode and Fixed object side by side (two columns)
        var topColumns = dialog.add("group");
        topColumns.orientation = "row";
        topColumns.alignChildren = ["fill", "fill"]; // 2パネルの高さをそろえる / Match panel heights
        topColumns.spacing = PANEL_SPACING;

        // モード / Mode（ラジオ縦並び）
        var modePanel = topColumns.add("panel", undefined, getLocalizedText('mode'));
        setupPanel(modePanel, 6);
        modePanel.alignChildren = ["left", "top"]; // ラジオは左寄せで縦並び / Stack radios, left-aligned
        var modeGroupRadio = modePanel.add("radiobutton", undefined, getLocalizedText('modeGroup'));       // グループ
        var modeAutoPairRadio = modePanel.add("radiobutton", undefined, getLocalizedText('modeAutoPair')); // 自動ペア認識
        modeGroupRadio.helpTip = getLocalizedText('tipModeGroup');
        modeAutoPairRadio.helpTip = getLocalizedText('tipModeAutoPair');
        // 既定モードは選択内容で決める：グループのみ→グループ、それ以外→自動ペア認識。
        // Default mode follows the selection: groups only → group, otherwise → auto pair.
        modeGroupRadio.value = selectionIsGroupsOnly;
        modeAutoPairRadio.value = !selectionIsGroupsOnly;

        // 固定オブジェクト / Fixed object（上・左・右・下を十字に配置 / arranged as a cross）
        var fixedSideRefs = buildFixedSidePanel(topColumns);
        var fixedSidePanel = fixedSideRefs.panel;
        var fixedRadios = fixedSideRefs.fixedRadios;
        var selectFixedRadio = fixedSideRefs.selectFixedRadio;
        var getFixedSide = fixedSideRefs.getFixedSide;
        var setFixedSide = fixedSideRefs.setFixedSide;

        // 上段2パネルを全幅・等幅にそろえる：preferred 幅を広い方に合わせ、fill で均等に広げる。
        // Make the two top panels full-width and equal: match preferred widths, fill expands them evenly.
        modePanel.alignment = ["fill", "fill"];
        fixedSidePanel.alignment = ["fill", "fill"];
        var topPanelWidth = Math.max(modePanel.preferredSize.width, fixedSidePanel.preferredSize.width);
        modePanel.preferredSize.width = topPanelWidth;
        fixedSidePanel.preferredSize.width = topPanelWidth;

        // オプション / Options（間隔の入力・クリア・プレビュー境界）
        var optionsRefs = buildSpacingOptionsPanel(dialog, initialGapPoints);
        var spacingInput = optionsRefs.spacingInput;
        var clearButton = optionsRefs.clearButton;
        var previewBoundsCheckbox = optionsRefs.previewBoundsCheckbox;
        var getSpacingInPoints = optionsRefs.getSpacingInPoints;
        var getBoundsType = optionsRefs.getBoundsType;

        // 整列パネル / Alignment（移動側の整列。固定側に依存するので getFixedSide を渡す）
        var alignRefs = buildAlignmentPanel(dialog, getFixedSide);
        var hAlign = alignRefs.hAlign;
        var vAlign = alignRefs.vAlign;
        var updateAlignPanels = alignRefs.updateAlignPanels;
        var getAlignMode = alignRefs.getAlignMode;

        // 前回終了時の設定を復元（固定側・プレビュー境界・整列の3つだけ。モード／間隔は選択から毎回決める）。
        // Restore last-used settings: fixed side, preview bounds, alignment only (mode/gap stay auto).
        var savedSettings = loadSettings();
        if (savedSettings.fixedSide) setFixedSide(savedSettings.fixedSide);
        if (savedSettings.previewBounds === "true") previewBoundsCheckbox.value = true;
        applySavedAlign(hAlign, savedSettings.alignH);
        applySavedAlign(vAlign, savedSettings.alignV);

        /* 現在のモードを取得する / Get the current mode ("group" | "auto") */
        function getMode() {
            return modeGroupRadio.value ? "group" : "auto";
        }

        /* モードを切り替えてペアを組み直し、プレビューを更新する / Switch mode, rebuild pairs, refresh */
        function onModeChange() {
            buildPairs(getMode());
            refreshPreview();
        }

        /* 現在の設定でプレビューを更新する / Refresh preview with current settings */
        function refreshPreview() {
            runPreview(getFixedSide(), getSpacingInPoints(), getBoundsType(), getAlignMode());
        }

        // 設定変更でライブプレビュー / Update preview on change
        modeGroupRadio.onClick = onModeChange;
        modeAutoPairRadio.onClick = onModeChange;
        // 固定側のラジオ：手動で排他にしてからプレビュー更新（軸の切替もここで反映）
        // Fixed-side radios: enforce exclusivity by hand, then refresh (axis switch applies here too)
        for (var ri = 0; ri < fixedRadios.length; ri++) {
            fixedRadios[ri].onClick = (function (radio) {
                return function () { selectFixedRadio(radio); updateAlignPanels(); refreshPreview(); };
            })(fixedRadios[ri]);
        }
        changeValueByArrowKey(spacingInput, refreshPreview); // ↑↓キーで増減＋プレビュー更新 / Arrow keys + preview
        spacingInput.onChanging = refreshPreview;
        previewBoundsCheckbox.onClick = refreshPreview;
        // 整列ラジオ（左右/上下の各行）：選択でプレビュー更新 / Alignment radios refresh on click
        var alignRows = [hAlign, vAlign];
        for (var ai = 0; ai < alignRows.length; ai++) {
            var alignRadios = alignRows[ai].radios;
            for (var key in alignRadios) { alignRadios[key].onClick = refreshPreview; }
        }
        // クリアボタン：間隔を 0 にしてプレビュー更新 / Clear button: set gap to 0, refresh
        clearButton.onClick = function () {
            spacingInput.text = "0";
            refreshPreview();
        };

        // ボタン（Mac 規約：Cancel → OK）/ Buttons (Mac order: Cancel → OK)
        var buttonRow = dialog.add("group");
        buttonRow.alignment = "right";
        buttonRow.add("button", undefined, getLocalizedText('cancel'), { name: "cancel" });
        buttonRow.add("button", undefined, "OK", { name: "ok" });

        // ダイアログ表示時に既定モードでペアを組んで初回プレビュー（同期側 undo を避けて onShow から起動）
        // Build pairs for the default mode, then run the first preview (from onShow to avoid sync undo)
        dialog.onShow = function () {
            updateAlignPanels(); // 既定の固定側に合わせて整列パネルの有効/無効を初期化 / Init enabled state
            buildPairs(getMode());
            refreshPreview();
        };

        var accepted = (dialog.show() === 1);
        if (accepted) {
            // OK時に現在の設定を保存（固定側・プレビュー境界・整列）/ Save settings on OK
            saveSettings({
                fixedSide: getFixedSide(),
                previewBounds: previewBoundsCheckbox.value ? "true" : "false",
                alignH: getAlignRowValue(hAlign),
                alignV: getAlignRowValue(vAlign)
            });
        }
        return accepted;
    }

    if (showSettingsDialog()) {
        // OK：プレビューをそのまま確定 / Keep the applied preview
    } else {
        // キャンセル：プレビューを巻き戻す / Revert the preview
        undoPreview();
        app.redraw();
    }
})();
