#target illustrator
#targetengine "AdjustPairGap"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

# ペア配置調整スクリプト

オブジェクトの間隔（および位置）を指定値にそろえます。「固定」で選んだ側を基準に、残りを動かします。ライブプレビュー対応。

## モード

- 自動ペア認識：選択オブジェクトを最も近いもの同士でペアにし、各ペアの間隔をそろえます
- グループ：選択した各グループの中身を並べ、隣り合う間隔をすべて指定値にそろえます（3個以上もカスケードで分配）
- アートボード：各選択オブジェクトと、「固定」で選んだアートボード端（上 / 左 / 右 / 下）との間隔（マージン）をそろえます

選択がすべてグループなら起動時に「グループ」、それ以外は「自動ペア認識」が初期選択です（「アートボード」は手動。ダイアログで切替可）。

## 固定（基準）

上 / 左 / 右 / 下 で基準にする側を選びます。左右なら水平方向、上下なら垂直方向の間隔をそろえ、選んだ側は動かしません。グループモードでは選んだ端を起点に残りをカスケード配置します。

## 間隔（位置調整パネル）

- 「間隔」で対象間の距離を指定します（↑↓キーで±1、Shiftで±10、Optionで±0.1。マイナス値で重ねられます）
- 「プレビュー境界」オンで線幅や効果を含む見た目の境界を基準に測ります

## 水平 / 垂直パネル（移動側）

「固定」が上下のときは「水平」パネル、左右のときは「垂直」パネルが有効になります（直交する側だけ操作可）。

- 整列：動かす側を固定側の端／中央にそろえます（不要なら「移動しない」）
- 位置：移動側（＝固定でない側）を直交方向へ微調整します（正＝右／下）。整列が「中央」のときは 0 にして無効化します
- ショートカット：水平＝L/C/R（左/中央/右）、垂直＝T/M/B（上/中央/下）

## テキストの行揃え

テキストオブジェクトの段落の行揃えを指定します（自動 / 左 / 中央 / 右 / 均等配置（最終行左））。

- 自動（既定）：整列・固定に連動します。エリア内文字＝均等配置（最終行左）、ポイント文字＝縦並びは水平整列に連動（左/中央/右）、横並びは固定側に連動（固定が左→左／固定が右→右）
- 左 / 中央 / 右 / 均等配置：ポイント文字・エリア内文字とも一律にその行揃え

行揃えは間隔・整列より先に適用し、境界を測り直してから配置します（ポイント文字は行揃えで字幅が変わるため）。

## 設定の記憶

ダイアログを閉じる（OK）と設定を記憶し、次回はその状態で開きます。

- 同一セッション中：モード・固定・整列・間隔・位置・プレビュー境界をすべて記憶（#targetengine）
- Illustrator 再起動後：固定・整列・位置・プレビュー境界を復元（モード・間隔は選択内容に依存するため毎回決め直し）

## 注意

- 自動ペア認識：選択数が奇数のときは末尾の1つがペアにならず残ります
- グループ：子が2個未満のグループ、およびグループ以外のオブジェクトは対象外です
- アートボード：各オブジェクトは中心が含まれるアートボードを基準にします（該当がなければアクティブアートボード）

### 紹介記事

https://note.com/dtp_tranist/n/nc8fab19d8164

---

# Adjust Pair Gap

Aligns the gap (and position) between objects to a specified value. Uses the side chosen by the key object as the anchor and moves the rest. Supports live preview.

## Modes

- Auto pair detection: pairs the selected objects by nearest neighbor and aligns the gap of each pair
- Group: lays out the contents of each selected group and aligns every adjacent gap to the specified value (3+ items are distributed in a cascade)
- Artboard: aligns the gap (margin) between each selected object and the artboard edge (top / left / right / bottom) chosen by the key object

If the whole selection is groups, "Group" is selected at launch; otherwise "Auto pair detection" is selected ("Artboard" is manual; switchable in the dialog).

## Key object (anchor)

Choose the anchor side with top / left / right / bottom. Left/right aligns the horizontal gap, top/bottom the vertical gap, and the chosen side does not move. In Group mode, the chosen edge is the starting point for cascading the rest.

## Gap (position panel)

- "Gap" specifies the distance between targets (↑↓ keys ±1, Shift ±10, Option ±0.1; negative values overlap them)
- With "Preview bounds" on, the visual bounds including stroke width and effects are used for measurement

## Horizontal / Vertical panel (moving side)

When the key is top/bottom, the "Horizontal" panel is enabled; when left/right, the "Vertical" panel is enabled (only the orthogonal side can be operated).

- Align: aligns the moving side to the edge/center of the key object (use "Don't move" if not needed)
- Position: fine-tunes the moving side (= the non-key side) in the orthogonal direction (positive = right / down). Set to 0 to disable when Align is "Center"
- Shortcuts: horizontal = L/C/R (left/center/right), vertical = T/M/B (top/center/bottom)

## Text justification

Specifies the paragraph justification of text objects (auto / left / center / right / justify (last line left)).

- Auto (default): follows alignment and key. Area text = justify (last line left); point text = vertical layout follows horizontal alignment (left/center/right), horizontal layout follows the key side (key left → left / key right → right)
- Left / Center / Right / Justify: applies that justification uniformly to both point text and area text

Justification is applied before the gap and alignment, then the bounds are re-measured before placement (because point text changes its width with justification).

## Remembering settings

Closing the dialog (OK) remembers the settings, and they are restored on the next open.

- Within the same session: mode, key, alignment, gap, position, and preview bounds are all remembered (#targetengine)
- After an Illustrator restart: key, alignment, position, and preview bounds are restored (mode and gap depend on the selection, so they are decided each time)

## Notes

- Auto pair detection: when the selection count is odd, the last one is left unpaired
- Group: groups with fewer than 2 children, and non-group objects, are excluded
- Artboard: each object uses the artboard that contains its center as the anchor (or the active artboard if none applies)

### Article

https://note.com/dtp_tranist/n/nc8fab19d8164

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AdjustPairGap";                /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.3.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "";                             /* 更新日 / last updated */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

// =========================================
// ユーザー設定 / User settings
// =========================================
var DEFAULT_GAP = 30; // 平均間隔を測れないときのフォールバック（pt）/ Fallback gap when no average can be measured (pt)

// =========================================
// ローカライズ / Localization
// =========================================
var LABELS = {
    // ダイアログ / Dialog
    dialog: {
        title: { ja: "ペア配置の調整", en: "Adjust Pair Layout" }
    },
    // モード / Mode
    mode: {
        label: { ja: "モード", en: "Mode" },
        group: { ja: "グループ", en: "Group" },
        auto: { ja: "自動ペア認識", en: "Auto Pair Detection" },
        artboard: { ja: "アートボード", en: "Artboard" }
    },
    // キーオブジェクト / Key object
    fixedSide: {
        label: { ja: "固定", en: "Key Object" },
        top: { ja: "上", en: "Top" },
        left: { ja: "左", en: "Left" },
        right: { ja: "右", en: "Right" },
        bottom: { ja: "下", en: "Bottom" }
    },
    // 間隔 / Spacing
    spacing: {
        label: { ja: "間隔", en: "Gap" }
    },
    // 位置オフセット（移動側をキーオブジェクトに対し直交方向へずらす）/ Position offset (perpendicular to the gap)
    offset: {
        horizontal: { ja: "左右", en: "Horizontal" },
        vertical: { ja: "上下", en: "Vertical" }
    },
    // 水平／垂直パネル内の行ラベル / Row labels inside the Horizontal/Vertical panels
    panel: {
        align: { ja: "整列", en: "Align" },
        position: { ja: "位置", en: "Position" }
    },
    // テキストの行揃え / Text alignment (justification)
    justify: {
        label: { ja: "テキストの行揃え", en: "Text alignment" },
        auto: { ja: "自動", en: "Auto" },
        left: { ja: "左", en: "Left" },
        center: { ja: "中央", en: "Center" },
        right: { ja: "右", en: "Right" },
        full: { ja: "均等配置", en: "Justify" }
    },
    // 整列 / Alignment
    align: {
        label: { ja: "整列（移動側）", en: "Alignment (moved side)" },
        h: { ja: "水平", en: "Horizontal" },
        v: { ja: "垂直", en: "Vertical" },
        none: { ja: "移動しない", en: "Don't move" },
        center: { ja: "中央", en: "Center" }
    },
    // 位置調整 / Position
    options: {
        label: { ja: "位置調整", en: "Position" },
        previewBounds: { ja: "プレビュー境界", en: "Preview Bounds" }
    },
    // ボタン / Buttons（OK はローカライズしない / "OK" is not localized）
    button: {
        cancel: { ja: "キャンセル", en: "Cancel" }
    },
    // ツールチップ / Tooltips
    tip: {
        modeGroup: {
            ja: "選択した各グループの中身を、隣り合う間隔がすべて同じになるように並べます（3個以上も対応）。",
            en: "Distributes the contents of each selected group so every adjacent gap is equal (3+ objects supported)."
        },
        modeAutoPair: {
            ja: "選択オブジェクトを最も近いもの同士でペアにし、各ペアの間隔をそろえます。",
            en: "Pairs the selected objects by nearest neighbor and sets the gap of each pair."
        },
        modeArtboard: {
            ja: "アートボードの端（キーオブジェクトで選んだ上／左／右／下）を基準に、各選択オブジェクトとの間隔（マージン）を指定値にそろえます。",
            en: "Sets each selected object's gap (margin) to the artboard edge chosen in Key Object (top/left/right/bottom)."
        },
        fixedSide: {
            ja: "基準にする側。選んだ側は動かさず残りを移動します（左右＝水平、上下＝垂直）。",
            en: "The anchor side. The chosen side stays put while the rest move (Left/Right = horizontal, Top/Bottom = vertical)."
        },
        spacing: {
            ja: "オブジェクト間の間隔。マイナスにすると重なります。↑↓キーで±1、Shiftで±10、Optionで±0.1。",
            en: "Gap between objects; a negative value overlaps them. Arrow keys ±1, Shift ±10, Option ±0.1."
        },
        offsetHorizontal: {
            ja: "上下をキーにしたとき有効。整列後の移動側を左右へ追加でずらします（正＝右）。↑↓キーで±1、Shiftで±10、Optionで±0.1。",
            en: "Active when the key is top/bottom. Nudges the moved side horizontally after alignment (positive = right). Arrow keys ±1, Shift ±10, Option ±0.1."
        },
        offsetVertical: {
            ja: "左右をキーにしたとき有効。整列後の移動側を上下へ追加でずらします（正＝下）。↑↓キーで±1、Shiftで±10、Optionで±0.1。",
            en: "Active when the key is left/right. Nudges the moved side vertically after alignment (positive = down). Arrow keys ±1, Shift ±10, Option ±0.1."
        },
        alignH: {
            ja: "縦に並べたとき（上下をキーに）、動く側をキーオブジェクトの左端／中央／右端にそろえます。",
            en: "When stacking vertically (top/bottom key), aligns moved objects to the key object's left/center/right."
        },
        alignV: {
            ja: "横に並べたとき（左右をキーに）、動く側をキーオブジェクトの上端／中央／下端にそろえます。",
            en: "When laying out horizontally (left/right key), aligns moved objects to the key object's top/center/bottom."
        },
        previewBounds: {
            ja: "オンで線幅や効果を含む見た目の境界、オフで幾何境界を基準に間隔を測ります。",
            en: "On: measure by visible bounds (incl. stroke/effects). Off: geometric bounds."
        },
        justifyAuto: {
            ja: "整列・キーに連動。エリア内文字は均等配置、ポイント文字は縦並びなら水平整列・横並びならキー側に合わせます。",
            en: "Linked to align/key: area text is justified; point text follows the horizontal align (vertical stack) or the key side (horizontal row)."
        },
        justifyLeft: {
            ja: "テキストを左揃えにします。",
            en: "Left-aligns the text."
        },
        justifyCenter: {
            ja: "テキストを中央揃えにします。",
            en: "Center-aligns the text."
        },
        justifyRight: {
            ja: "テキストを右揃えにします。",
            en: "Right-aligns the text."
        },
        justifyFull: {
            ja: "テキストを均等配置（最終行左）にします。",
            en: "Justifies the text (last line left-aligned)."
        }
    },
    // アラート / Alerts
    alert: {
        noDocument: {
            ja: "ドキュメントが開かれていません。",
            en: "No document is open."
        },
        selectTwo: {
            ja: "オブジェクトを2つ以上選択してください。",
            en: "Select at least two objects."
        }
    }
};

var currentLanguage = ($.locale && $.locale.indexOf("ja") === 0) ? "ja" : "en";

/* ドット区切りキー（"mode.group" など）でロケールに応じたラベルを返す / Return a localized label by dotted key */
function getLocalizedText(key) {
    var parts = key.split(".");
    var entry = LABELS;
    for (var i = 0; i < parts.length; i++) {
        if (!entry) break;
        entry = entry[parts[i]];
    }
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

/* 整列のキーボードショートカットをダイアログに付ける。今アクティブな向きのキーだけ反応する。
   水平がアクティブ（上下キー時）：L=左 / C=中央 / R=右、垂直がアクティブ（左右キー時）：T=上 / M=中央 / B=下。
   isHorizontalActive() で現在の有効な向きを判定する（パネルの enabled では子の enabled が追従しないため）。
   Attach alignment keyboard shortcuts; only the currently active orientation responds.
   isHorizontalActive() decides which orientation is active (don't rely on child .enabled, which does
   not follow a disabled parent panel). Calls onChange on change. */
function addAlignmentKeyHandler(dialog, hAlign, vAlign, isHorizontalActive, onChange) {
    dialog.addEventListener("keydown", function (event) {
        var target = null;
        var key = event.keyName;
        if (isHorizontalActive()) {
            if (key === "L") target = hAlign.radios.start;
            else if (key === "C") target = hAlign.radios.center;
            else if (key === "R") target = hAlign.radios.end;
        } else {
            if (key === "T") target = vAlign.radios.start;
            else if (key === "M") target = vAlign.radios.center;
            else if (key === "B") target = vAlign.radios.end;
        }
        if (!target) return;
        event.preventDefault();
        if (target.value) return;     // 既に選択済みなら何もしない / already selected
        target.value = true;          // 同じ行内のラジオは自動排他 / radios in a row auto-exclude
        if (typeof onChange === "function") onChange();
    });
}

/* パネルの余白と間隔 / Panel margins and spacing */
var PANEL_MARGINS = [16, 20, 16, 12];
var PANEL_SPACING = 8;
var COLUMN_SPACING = 12; /* 2カラムの間隔 / Gap between the two columns */

/* パネルの共通設定 / Apply shared panel layout */
function setupPanel(panel, spacing) {
    panel.orientation = "column";
    panel.alignChildren = ["fill", "top"];
    panel.alignment = "fill";
    panel.margins = PANEL_MARGINS;
    panel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
}

/* グループの共通設定（row/column で整列を切り替え）/ Apply shared group layout (alignChildren switches by orientation) */
function setupGroup(group, orientation, spacing) {
    var groupOrientation = orientation || "column";
    group.orientation = groupOrientation;
    /* row は横並びなので縦中央、column は縦並びなので左揃え / row: vertically centered, column: left-aligned */
    group.alignChildren = (groupOrientation === "row") ? ["left", "center"] : ["left", "top"];
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
    var panel = parentGroup.add("panel", undefined, getLocalizedText('fixedSide.label'));
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
        for (var i = 0; i < fixedRadios.length; i++) {
            fixedRadios[i].value = (fixedRadios[i] === chosen);
        }
    }
    for (var i = 0; i < fixedRadios.length; i++) {
        fixedRadios[i].helpTip = getLocalizedText('tip.fixedSide');
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

/* 位置調整パネル（間隔の入力・プレビュー境界）を生成する。
   左右/上下オフセットは「水平」「垂直」パネルへ移動したので、ここは間隔だけを扱う。
   間隔は現在のルーラー単位で表示し、getSpacingInPoints() が pt に換算して返す。
   initialGapPoints は間隔の初期値・空欄時のフォールバック（pt）。イベント結線は呼び出し側で行う。
   Build the Position panel (gap input, preview-bounds). Offsets moved to the Horizontal/
   Vertical panels, so this panel handles only the gap. Event wiring is left to the caller.
   返り値 / Returns: { panel, spacingInput, previewBoundsCheckbox,
                       getSpacingInPoints, getBoundsType } */
function buildGapPanel(parentGroup, initialGapPoints) {
    var panel = parentGroup.add("panel", undefined, getLocalizedText('options.label'));
    setupPanel(panel, 6);

    // 間隔行（ラベル＋入力＋単位）/ Gap row (label + input + unit)
    var spacingRow = panel.add("group");
    setupGroup(spacingRow, "row");
    spacingRow.alignment = "left"; // 広げず左寄せ / Keep at natural width, packed left
    spacingRow.add("statictext", undefined, labelText('spacing.label'));
    var defaultSpacingDisplay = Math.round((initialGapPoints / pointsPerUnit) * 100) / 100;
    var spacingInput = spacingRow.add("edittext", undefined, String(defaultSpacingDisplay));
    spacingInput.characters = 4;
    spacingInput.helpTip = getLocalizedText('tip.spacing');
    spacingRow.add("statictext", undefined, rulerUnitLabel);

    // チェックボックス：プレビュー境界（左添え）/ Preview-bounds checkbox (left)
    var previewBoundsGroup = panel.add("group");
    previewBoundsGroup.orientation = "row";
    previewBoundsGroup.alignment = "left";
    previewBoundsGroup.margins = [0, 5, 0, 0]; // 上マージン5 / Top margin 5
    var previewBoundsCheckbox = previewBoundsGroup.add("checkbox", undefined, getLocalizedText('options.previewBounds'));
    previewBoundsCheckbox.value = false; // OFF=幾何境界 / ON=プレビュー境界
    previewBoundsCheckbox.helpTip = getLocalizedText('tip.previewBounds');

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
        previewBoundsCheckbox: previewBoundsCheckbox,
        getSpacingInPoints: getSpacingInPoints,
        getBoundsType: getBoundsType
    };
}

/* 「水平」または「垂直」パネルを生成する。整列（移動しない/開始/中央/終端）と位置（オフセット）の
   2行を1枚にまとめる。which="h" は水平（整列＝左/中央/右・位置＝左右）、"v" は垂直（上/中央/下・上下）。
   有効/無効・中央時のオフセット無効化・getAlignMode などは呼び出し側で行う。
   Build the "Horizontal" or "Vertical" panel: an alignment row (none/start/center/end) and a
   position (offset) row. which="h" → horizontal (left/center/right, left-right offset);
   "v" → vertical (top/center/bottom, up-down offset). Enable/disable and event wiring are the caller's.
   返り値 / Returns: { panel, align:{row,label,radios}, offsetRow, offsetLabel, offsetInput, getOffsetInPoints } */
function buildOrientationPanel(parentGroup, which) {
    var isH = (which === "h");
    var titleKey = isH ? 'align.h' : 'align.v';                  // 水平 / 垂直
    var startKey = isH ? 'fixedSide.left' : 'fixedSide.top';     // 左 / 上
    var endKey = isH ? 'fixedSide.right' : 'fixedSide.bottom';   // 右 / 下
    var tipAlignKey = isH ? 'tip.alignH' : 'tip.alignV';
    var tipOffsetKey = isH ? 'tip.offsetHorizontal' : 'tip.offsetVertical';

    var panel = parentGroup.add("panel", undefined, getLocalizedText(titleKey));
    setupPanel(panel, 6);

    // 整列行（ラベル＋移動しない/開始/中央/終端）/ Alignment row (label + none/start/center/end)
    var alignRow = panel.add("group");
    setupGroup(alignRow, "row");
    alignRow.alignment = "left";
    alignRow.alignChildren = ["left", "center"];
    var alignLabel = alignRow.add("statictext", undefined, labelText('panel.align'));
    var radios = {
        none: alignRow.add("radiobutton", undefined, getLocalizedText('align.none')),
        start: alignRow.add("radiobutton", undefined, getLocalizedText(startKey)),
        center: alignRow.add("radiobutton", undefined, getLocalizedText('align.center')),
        end: alignRow.add("radiobutton", undefined, getLocalizedText(endKey))
    };
    radios.none.value = true; // 既定：整列なし / Default: no alignment
    var alignTip = getLocalizedText(tipAlignKey);
    for (var key in radios) { radios[key].helpTip = alignTip; }

    // 位置行（ラベル＋入力＋単位）/ Position row (label + input + unit)
    var offsetRow = panel.add("group");
    setupGroup(offsetRow, "row");
    offsetRow.alignment = "left";
    offsetRow.alignChildren = ["left", "center"];
    var offsetLabel = offsetRow.add("statictext", undefined, labelText('panel.position'));
    var offsetInput = offsetRow.add("edittext", undefined, "0");
    offsetInput.characters = 4;
    offsetInput.helpTip = getLocalizedText(tipOffsetKey);
    offsetRow.add("statictext", undefined, rulerUnitLabel);

    // ラベル幅をそろえて整列ラジオと入力の開始位置を合わせる / Match label widths
    var labelWidth = Math.max(alignLabel.preferredSize.width, offsetLabel.preferredSize.width);
    alignLabel.preferredSize.width = labelWidth;
    offsetLabel.preferredSize.width = labelWidth;

    /* オフセット入力を pt に換算（空・不正は 0）/ Offset input in pt (empty/invalid → 0) */
    function getOffsetInPoints() {
        var value = parseFloat(offsetInput.text);
        if (isNaN(value)) value = 0;
        return value * pointsPerUnit;
    }

    return {
        panel: panel,
        align: { row: alignRow, label: alignLabel, radios: radios },
        offsetRow: offsetRow,
        offsetLabel: offsetLabel,
        offsetInput: offsetInput,
        getOffsetInPoints: getOffsetInPoints
    };
}

/* テキストの行揃えパネルを生成する。ラジオ横並び：自動 / 左 / 中央 / 右 / 均等配置（最終行左）。
   getJustifyMode() で選択を文字列（"auto"/"left"/"center"/"right"/"full"）で返す。
   「自動」は整列・キーに連動（エリア内文字は均等配置）。解決は applySpacing 側で行う。イベント結線は呼び出し側。
   Build the "Text alignment" panel: horizontal radios (auto/left/center/right/justify).
   getJustifyMode() returns the selection as a string; "auto" links to align/key (resolved in
   applySpacing). Event wiring is left to the caller.
   返り値 / Returns: { panel, radios, getJustifyMode } */
function buildJustifyPanel(parentGroup) {
    var panel = parentGroup.add("panel", undefined, getLocalizedText('justify.label'));
    setupPanel(panel, 6);
    panel.orientation = "row"; // ラジオ横並び / radios in a row
    panel.alignChildren = ["left", "center"];

    var radios = {
        auto: panel.add("radiobutton", undefined, getLocalizedText('justify.auto')),
        left: panel.add("radiobutton", undefined, getLocalizedText('justify.left')),
        center: panel.add("radiobutton", undefined, getLocalizedText('justify.center')),
        right: panel.add("radiobutton", undefined, getLocalizedText('justify.right')),
        full: panel.add("radiobutton", undefined, getLocalizedText('justify.full'))
    };
    radios.auto.value = true; // 既定：自動（整列・キーに連動）/ Default: auto (linked to align/key)
    // 各ラジオにツールチップ / Tooltip per radio
    radios.auto.helpTip = getLocalizedText('tip.justifyAuto');
    radios.left.helpTip = getLocalizedText('tip.justifyLeft');
    radios.center.helpTip = getLocalizedText('tip.justifyCenter');
    radios.right.helpTip = getLocalizedText('tip.justifyRight');
    radios.full.helpTip = getLocalizedText('tip.justifyFull');

    /* 選択中の行揃えモードを文字列で返す / Selected justify mode as a string */
    function getJustifyMode() {
        if (radios.left.value) return "left";
        if (radios.center.value) return "center";
        if (radios.right.value) return "right";
        if (radios.full.value) return "full";
        return "auto";
    }

    return { panel: panel, radios: radios, getJustifyMode: getJustifyMode };
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
// 行揃え / Text justification
// =========================================
/* 左右方向の整列値（start/center/end）を段落の行揃え（Justification）へ対応づける。
   none・不明は null（行揃えを変更しない）/ Map a horizontal align value to a Justification; null = leave as-is.
   start=左揃え, center=中央揃え, end=右揃え / start=left, center=center, end=right */
function justificationForAlign(alignMode) {
    if (alignMode === "start") return Justification.LEFT;
    if (alignMode === "center") return Justification.CENTER;
    if (alignMode === "end") return Justification.RIGHT;
    return null;
}

/* テキストの段落行揃えを設定する。Justification.LEFT は代入が無視される Illustrator のバグが
   あるので、一時 resize（200%→50%）で段落属性をリフレッシュしてから代入し、位置を保存して戻す。
   RIGHT/CENTER はそのまま代入できる。
   Set paragraph justification. Assigning Justification.LEFT is ignored by Illustrator, so a
   temporary resize (200% then 50%) refreshes the paragraph attributes before assignment; the
   position is saved and restored. RIGHT/CENTER assign directly. */
function setParagraphJustification(textFrame, justification) {
    if (justification === Justification.LEFT) {
        var savedPosition = [textFrame.position[0], textFrame.position[1]];
        textFrame.resize(200, 200);
        textFrame.textRange.paragraphAttributes.justification = Justification.LEFT;
        textFrame.resize(50, 50);
        try { textFrame.position = savedPosition; } catch (ePos) {}
        return;
    }
    textFrame.textRange.paragraphAttributes.justification = justification;
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

/* オブジェクトの中心を含むアートボードの矩形 [左, 上, 右, 下] を返す。該当が無ければアクティブ
   アートボードを使う。geometricBounds と同じ並び（上 > 下、y は上方向で増加）なので軸関数で扱える。
   Return the artboardRect [left, top, right, bottom] of the artboard containing the item's center
   (falls back to the active artboard). Same layout as geometricBounds, so the axis helpers apply. */
function artboardRectFor(item) {
    var artboards = app.activeDocument.artboards;
    var bounds = geometricBoundsOf(item);
    var centerX = (bounds[0] + bounds[2]) / 2;
    var centerY = (bounds[1] + bounds[3]) / 2;
    for (var i = 0; i < artboards.length; i++) {
        var rect = artboards[i].artboardRect; // [左, 上, 右, 下] / [left, top, right, bottom]
        if (centerX >= rect[0] && centerX <= rect[2] && centerY <= rect[1] && centerY >= rect[3]) {
            return rect;
        }
    }
    return artboards[artboards.getActiveArtboardIndex()].artboardRect;
}

/* アートボードモードの作業単位を作る。選択オブジェクトを1つずつ独立した単位にし、
   applySpacing 側で各オブジェクトとアートボード端の間隔（マージン）を指定値にそろえる。
   Build the units for Artboard mode: each selected object becomes an independent unit, and
   applySpacing sets each object's gap (margin) to the chosen artboard edge. */
function createArtboardUnits(selectedItems) {
    var units = [];
    for (var i = 0; i < selectedItems.length; i++) {
        units.push({ single: selectedItems[i] });
    }
    return units;
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
        for (var i = 0; i < pairs.length; i++) {
            var boundsA = geometricBoundsOf(pairs[i].a);
            var boundsB = geometricBoundsOf(pairs[i].b);
            var leadBounds = (axis.start(boundsA) < axis.start(boundsB)) ? boundsA : boundsB;
            var trailBounds = (axis.start(boundsA) < axis.start(boundsB)) ? boundsB : boundsA;
            gaps.push(axis.start(trailBounds) - axis.end(leadBounds));
        }
    }

    if (gaps.length === 0) return null;
    var gapSum = 0;
    for (var i = 0; i < gaps.length; i++) gapSum += gaps[i];
    return gapSum / gaps.length;
}

// =========================================
// 設定の保存 / Persistence
// =========================================
// ダイアログを閉じるときに設定を覚え、次回開いたときに復元する
// （モード・キー・プレビュー境界・整列・間隔・左右/上下オフセット。数値は pt で保存）。
// モードと間隔は選択内容に依存するため、同一セッション中だけ覚える（ファイルには残さない）。
// 2段構えで記憶する：
//   - #targetengine の常駐グローバル（$.global）… 同一セッション内は即座に前回状態を復元（ファイル I/O なし）
//   - Folder.userData のファイル … Illustrator を再起動してもまたいで永続
// Remember only the fixed side, preview-bounds, and alignment on close; restore them next time.
// Mode and gap are re-derived from the selection each run, so they are not saved.
// Two layers: the #targetengine persistent global ($.global) restores instantly within a session,
// and a Folder.userData file persists across Illustrator restarts.
var SETTINGS_FILE = Folder.userData + "/AdjustPairGapSettings.txt";
var SETTINGS_GLOBAL_KEY = "adjustPairGapSettings"; // $.global 上のキー / key on $.global

/* 設定オブジェクトを浅くコピーする（常駐グローバルを直接書き換えないため）/ Shallow copy (avoid mutating the persistent global) */
function cloneSettings(src) {
    var copy = {};
    for (var key in src) {
        if (src.hasOwnProperty(key)) copy[key] = src[key];
    }
    return copy;
}

/* 設定を読み込む。まず #targetengine の常駐グローバル（同一セッションの最新）を優先し、
   無ければ設定ファイル（セッションまたぎ）を読む。どちらも無ければ空オブジェクト。
   Load settings: prefer the #targetengine persistent global (latest in-session); otherwise read
   the file (cross-session). Empty object if neither exists. */
function loadSettings() {
    if ($.global[SETTINGS_GLOBAL_KEY]) {
        return cloneSettings($.global[SETTINGS_GLOBAL_KEY]);
    }
    var settings = {};
    var file = new File(SETTINGS_FILE);
    if (!file.exists || !file.open("r")) return settings;
    var content = file.read();
    file.close();
    var lines = content.split("\n");
    for (var i = 0; i < lines.length; i++) {
        var eqIndex = lines[i].indexOf("=");
        if (eqIndex <= 0) continue;
        var key = lines[i].substring(0, eqIndex).replace(/^\s+|\s+$/g, "");
        var value = lines[i].substring(eqIndex + 1).replace(/^\s+|\s+$/g, "");
        settings[key] = value;
    }
    return settings;
}

// ファイル（セッションまたぎ）には保存しないキー。モード・間隔は選択内容に依存するので
// 同一セッション内（常駐グローバル）だけ覚え、再起動後は毎回選択から決め直す。
// Keys NOT written to the file (cross-session): mode and gap depend on the selection, so they are
// remembered only within the session (persistent global) and re-derived from the selection after a restart.
var SETTINGS_FILE_SKIP = { mode: true, gap: true };

/* 設定を保存する。常駐グローバル（同一セッション用）には全項目、ファイル（セッションまたぎ用）には
   モード・間隔を除いた項目を書く。
   Save settings: everything to the persistent global (in-session); everything except mode/gap to the file. */
function saveSettings(settings) {
    $.global[SETTINGS_GLOBAL_KEY] = cloneSettings(settings);
    var file = new File(SETTINGS_FILE);
    if (!file.open("w")) return;
    var lines = [];
    for (var key in settings) {
        if (settings.hasOwnProperty(key) && !SETTINGS_FILE_SKIP[key]) lines.push(key + "=" + settings[key]);
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
        alert(getLocalizedText('alert.noDocument'));
        return;
    }

    var activeDocument = app.activeDocument;

    // doc.selection を配列にコピーしておく（後の選択変更や undo の影響を受けないように）
    // Copy doc.selection into a plain array (immune to later selection changes / undo)
    var liveSelection = activeDocument.selection;
    // selection が無効・未選択のケースを先に弾く（Illustrator では稀に null になる）
    // Guard against an invalid / empty selection (selection can be null in rare cases)
    if (!liveSelection || liveSelection.length === 0) {
        alert(getLocalizedText('alert.selectTwo'));
        return;
    }
    var selectedItems = [];
    for (var i = 0; i < liveSelection.length; i++) {
        selectedItems.push(liveSelection[i]);
    }

    if (selectedItems.length < 2) {
        alert(getLocalizedText('alert.selectTwo'));
        return;
    }

    // モード切り替え時に組み直すペアを保持する / Holds pairs rebuilt when the mode changes
    var objectPairs = [];

    // 選択がすべてグループなら既定でグループモードにする（それ以外は自動ペア認識）。
    // Default to group mode when every selected object is a group; otherwise auto pair.
    var selectionIsGroupsOnly = true;
    for (var i = 0; i < selectedItems.length; i++) {
        if (selectedItems[i].typename !== "GroupItem") { selectionIsGroupsOnly = false; break; }
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
    var appliedJustifications = []; // 適用済みの行揃え変更 / Applied justification changes: { obj, original }

    /* 記録した移動を逆向きに適用して元に戻す（位置のみ。行揃えは別管理）/ Reverse recorded moves (position only).
       行揃えはリフレッシュ毎に巻き戻すと resize が連発して重いので、ここでは触らない。
       行揃えの巻き戻しは整列系ラジオの変更時とキャンセル時にだけ undoJustifications() で行う。
       Justification is NOT reverted here (reverting every refresh would thrash resize); it is only
       reverted on alignment-related radio changes and on cancel via undoJustifications(). */
    function undoPreview() {
        for (var i = appliedMoves.length - 1; i >= 0; i--) {
            appliedMoves[i].obj.translate(-appliedMoves[i].dx, -appliedMoves[i].dy);
        }
        appliedMoves = [];
    }

    /* テキストの行揃えを justification にそろえ、元の値を記録する（巻き戻し用）。
       既に同じなら何もしない（無駄な undo を作らない）。テキスト以外は無視。
       Set a TextFrame's justification, recording the original for undo. Skips no-ops and non-text. */
    function applyJustification(obj, justification) {
        if (justification === null) return;
        if (obj.constructor.name !== "TextFrame") return;
        var current = obj.textRange.paragraphAttributes.justification;
        if (current === justification) return;
        appliedJustifications.push({ obj: obj, original: current });
        setParagraphJustification(obj, justification);
    }

    /* 記録した行揃え変更を元の値へ戻す / Restore recorded justification changes */
    function undoJustifications() {
        for (var j = appliedJustifications.length - 1; j >= 0; j--) {
            setParagraphJustification(appliedJustifications[j].obj, appliedJustifications[j].original);
        }
        appliedJustifications = [];
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
            var nextLeadingEdge = axis.start(cachedBounds[order[count - 1]]); // 後ろ端オブジェクトの先頭辺（不動）
            for (var i = count - 2; i >= 0; i--) {
                var memberIndex = order[i];
                var desiredEnd = nextLeadingEdge - gapInPoints;
                var shiftAmount = desiredEnd - axis.end(cachedBounds[memberIndex]);
                moveByAxis(axis, members[memberIndex], shiftAmount);
                nextLeadingEdge = axis.start(cachedBounds[memberIndex]) + shiftAmount; // この要素の新しい先頭辺 / its new leading edge
            }
        } else {
            // 先頭側（左 or 上）を固定し、前から後ろへ配置 / Anchor the leading end; walk forward
            var prevTrailingEdge = axis.end(cachedBounds[order[0]]); // 先頭端オブジェクトの後ろ辺（不動）
            for (var i = 1; i < count; i++) {
                var memberIndex = order[i];
                var desiredStart = prevTrailingEdge + gapInPoints;
                var shiftAmount = desiredStart - axis.start(cachedBounds[memberIndex]);
                moveByAxis(axis, members[memberIndex], shiftAmount);
                prevTrailingEdge = axis.end(cachedBounds[memberIndex]) + shiftAmount; // この要素の新しい後ろ辺 / its new trailing edge
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

    /* グループの固定端メンバー（キーオブジェクト）のインデックスを返す。並び順は不変なので
       キャッシュ境界で判定して十分（ライブ境界を読まない）/ Index of the group's key member
       (at the fixed end). Order is stable, so cached bounds suffice (no live reads). */
    function groupAnchorIndex(pair, gapAxis, anchorEnd, useVisible) {
        var cached = useVisible ? pair.memberVis : pair.memberGeo;
        var count = pair.members.length;
        var order = [];
        for (var i = 0; i < count; i++) order.push(i);
        order.sort(function (x, y) { return gapAxis.start(cached[x]) - gapAxis.start(cached[y]); });
        return anchorEnd ? order[count - 1] : order[0];
    }

    /* グループの各メンバーを固定端のメンバー（アンカー）に整列軸方向でそろえる。
       行揃えは間隔調整より前に適用して境界を取り直す（cachePairBounds）ので、ここはキャッシュ境界で十分。
       Justification is applied and bounds re-cached before gap adjustment, so cached bounds suffice here. */
    function alignGroup(pair, gapAxis, anchorEnd, useVisible, alignAxis, alignMode) {
        if (alignMode === "none") return;
        var members = pair.members;
        var count = members.length;
        var bounds = useVisible ? pair.memberVis : pair.memberGeo;

        var order = [];
        for (var i = 0; i < count; i++) order.push(i);
        order.sort(function (x, y) { return gapAxis.start(bounds[x]) - gapAxis.start(bounds[y]); });

        var anchorIndex = anchorEnd ? order[count - 1] : order[0]; // 固定端のメンバー / member at the fixed end
        var anchorBounds = bounds[anchorIndex];
        for (var i = 0; i < count; i++) {
            if (i === anchorIndex) continue;
            alignToAnchor(alignAxis, members[i], bounds[i], anchorBounds, alignMode);
        }
    }

    /* 設定値で各ペアの間隔を調整する / Apply the gap to every pair.
       固定側から軸（水平/垂直）と固定端（先頭/後ろ）を決める / Axis and anchor end follow the fixed side.
       alignMode は整列（ギャップ軸に直交する方向、固定オブジェクト基準）/ alignMode aligns on the
       axis perpendicular to the gap, anchored at the fixed object. */
    function applySpacing(fixedSide, gapInPoints, boundsType, alignMode, offsetHorizontalPt, offsetVerticalPt, justifyMode) {
        var useVisible = (boundsType === "visibleBounds");
        var axis = axisForSide(fixedSide);
        var anchorEnd = isAnchorEnd(fixedSide);
        var alignAxis = makeAxis(!axis.vertical); // 整列はギャップ軸に直交 / perpendicular to the gap axis

        /* 「テキストの行揃え」パネルの選択 justifyMode を、オブジェクトごとに Justification へ解決する。
           "left"/"center"/"right"/"full" は一律。"auto" は連動：エリア内文字＝均等配置、
           ポイント文字＝縦並びなら水平整列に連動（左/中央/右）、横並びならキー側に連動（左→左/右→右）。
           Resolve justifyMode to a Justification per object. Explicit modes apply uniformly; "auto"
           links: area text → justify; point text → horizontal align (vertical stack) or key side
           (horizontal row). null = leave unchanged. */
        function resolveJustifyForObject(obj) {
            if (justifyMode === "left") return Justification.LEFT;
            if (justifyMode === "center") return Justification.CENTER;
            if (justifyMode === "right") return Justification.RIGHT;
            if (justifyMode === "full") return Justification.FULLJUSTIFYLASTLINELEFT;
            // auto（連動）/ auto (linked)
            if (obj.constructor.name === "TextFrame" && obj.kind === TextType.AREATEXT) {
                return Justification.FULLJUSTIFYLASTLINELEFT; // エリア内文字は均等配置 / area text → justify
            }
            if (!alignAxis.vertical) return justificationForAlign(alignMode); // 縦並び：水平整列に連動
            return (fixedSide === "left") ? Justification.LEFT
                : (fixedSide === "right") ? Justification.RIGHT : null;        // 横並び：キー側に連動
        }

        // 行揃えを先に全テキストへ適用し、境界を取り直す（ポイント文字は行揃えで字幅が変わるため、
        // 間隔・整列を行揃え後の実際の形で計算する）。applyJustification は冪等。
        // Apply justification to all text FIRST, then re-measure bounds, so gap/align use the
        // post-justification shape (point-text width changes with justification). Idempotent.
        for (var i = 0; i < objectPairs.length; i++) {
            var justifyPair = objectPairs[i];
            if (justifyPair.single) {
                applyJustification(justifyPair.single, resolveJustifyForObject(justifyPair.single));
            } else if (justifyPair.members) {
                for (var j = 0; j < justifyPair.members.length; j++) {
                    applyJustification(justifyPair.members[j], resolveJustifyForObject(justifyPair.members[j]));
                }
            } else {
                applyJustification(justifyPair.a, resolveJustifyForObject(justifyPair.a));
                applyJustification(justifyPair.b, resolveJustifyForObject(justifyPair.b));
            }
        }
        cachePairBounds(objectPairs); // 行揃え後の境界で取り直す / re-cache post-justification bounds

        // 位置オフセット：ギャップ軸に直交する側だけ使う（上下キー→左右、左右キー→上下）。
        // キーオブジェクトでない側（移動側）だけを alignAxis 方向へずらす（右＝正／下＝正）。
        // Position offset along the perpendicular (align) axis only, applied to the non-key (moved) object(s).
        var offsetAlong = (!alignAxis.vertical) ? offsetHorizontalPt : offsetVerticalPt;

        for (var i = 0; i < objectPairs.length; i++) {
            var pair = objectPairs[i];

            if (pair.single) {
                // アートボードモード：オブジェクトとアートボード端の間隔（マージン）を指定値にそろえる。
                // Artboard mode: set the object's gap (margin) to the chosen artboard edge; align to it.
                var singleBounds = useVisible ? pair.visibleS : pair.geometricS;
                var artboardRect = pair.artboardRect;
                var marginShift = anchorEnd
                    ? (axis.end(artboardRect) - gapInPoints) - axis.end(singleBounds)    // 右/下端から内側へ / inset from trailing edge
                    : (axis.start(artboardRect) + gapInPoints) - axis.start(singleBounds); // 左/上端から内側へ / inset from leading edge
                moveByAxis(axis, pair.single, marginShift);
                if (alignMode !== "none") {
                    alignToAnchor(alignAxis, pair.single, singleBounds, artboardRect, alignMode);
                }
                moveByAxis(alignAxis, pair.single, offsetAlong); // 位置オフセット（アートボード基準なので常に対象）/ offset
                continue;
            }

            if (pair.members) {
                // グループモード：全オブジェクトを等間隔に分配（固定側を基準）、その後に整列
                // Group mode: distribute (anchored at the fixed side), then align
                distributeGroup(pair, fixedSide, gapInPoints, useVisible, axis);
                alignGroup(pair, axis, anchorEnd, useVisible, alignAxis, alignMode);
                // 位置オフセット：キーオブジェクト（固定端メンバー）以外を直交方向へずらす / Offset non-key members only
                if (offsetAlong !== 0) {
                    var keyMemberIndex = groupAnchorIndex(pair, axis, anchorEnd, useVisible);
                    for (var j = 0; j < pair.members.length; j++) {
                        if (j === keyMemberIndex) continue;
                        moveByAxis(alignAxis, pair.members[j], offsetAlong);
                    }
                }
                continue;
            }

            // 自動ペア認識：2オブジェクトの間隔を調整（行揃え後の境界）/ Auto mode (post-justification bounds)
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
            var gapShift, movedObject, movedBounds, anchorBounds;
            if (anchorEnd) {
                // 後ろ側（右 or 下）を固定：先頭側オブジェクトを動かす / Fix trailing end, move the leading object
                gapShift = (axis.start(trailBounds) - gapInPoints) - axis.end(leadBounds);
                moveByAxis(axis, leadObject, gapShift);
                movedObject = leadObject; movedBounds = leadBounds; anchorBounds = trailBounds;
            } else {
                // 先頭側（左 or 上）を固定：後ろ側オブジェクトを動かす / Fix leading end, move the trailing object
                gapShift = (axis.end(leadBounds) + gapInPoints) - axis.start(trailBounds);
                moveByAxis(axis, trailObject, gapShift);
                movedObject = trailObject; movedBounds = trailBounds; anchorBounds = leadBounds;
            }
            // 整列はギャップ軸と直交方向：ギャップ移動で整列軸の値は変わらないのでキャッシュ境界で可。
            // Alignment is perpendicular to the gap; the gap move doesn't change the align-axis value, so cache is fine.
            if (alignMode !== "none") {
                alignToAnchor(alignAxis, movedObject, movedBounds, anchorBounds, alignMode);
            }
            moveByAxis(alignAxis, movedObject, offsetAlong); // 位置オフセット（移動側のみ）/ position offset (moved object only)
        }
    }

    /* 直前のプレビューを巻き戻してから再適用する / Re-run preview from scratch.
       行揃え・整列ともプレビュー時点で最終結果と一致するので、OK では別処理は不要。
       Preview already matches the final result (justification + alignment), so OK needs no extra pass. */
    function runPreview(fixedSide, gapInPoints, boundsType, alignMode, offsetHorizontalPt, offsetVerticalPt, justifyMode) {
        undoPreview();
        applySpacing(fixedSide, gapInPoints, boundsType, alignMode, offsetHorizontalPt, offsetVerticalPt, justifyMode);
        app.redraw();
    }

    /* 各ペアの元の境界をキャッシュする / Cache original bounds of every pair.
       適用時は常に元位置へ巻き戻してから計算するので、ここで一度取れば使い回せる。 */
    function cachePairBounds(pairs) {
        for (var i = 0; i < pairs.length; i++) {
            var pair = pairs[i];
            if (pair.single) {
                // アートボード：オブジェクト境界と所属アートボードの矩形をキャッシュ
                // Artboard: cache the object's bounds and its artboard rect
                pair.geometricS = geometricBoundsOf(pair.single);
                pair.visibleS = visibleBoundsOf(pair.single);
                pair.artboardRect = artboardRectFor(pair.single);
            } else if (pair.members) {
                // グループ：全メンバーの境界をキャッシュ / Group: cache every member's bounds
                pair.memberGeo = [];
                pair.memberVis = [];
                for (var j = 0; j < pair.members.length; j++) {
                    pair.memberGeo.push(geometricBoundsOf(pair.members[j]));
                    pair.memberVis.push(visibleBoundsOf(pair.members[j]));
                }
            } else {
                pair.geometricA = geometricBoundsOf(pair.a);
                pair.visibleA = visibleBoundsOf(pair.a);
                pair.geometricB = geometricBoundsOf(pair.b);
                pair.visibleB = visibleBoundsOf(pair.b);
            }
        }
    }

    /* モードに応じてペアを組み直す / Rebuild pairs for the given mode ("auto" | "group" | "artboard").
       境界キャッシュは必ず元位置で取るため、先にプレビューを巻き戻してから組む。
       Always revert the preview first so cached bounds reflect original positions. */
    function buildPairs(mode) {
        undoPreview();
        var pairs;
        if (mode === "group") {
            pairs = createGroupPairs(selectedItems);
        } else if (mode === "artboard") {
            pairs = createArtboardUnits(selectedItems);
        } else {
            pairs = createNearestPairs(selectedItems);
        }
        cachePairBounds(pairs);
        objectPairs = pairs;
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================
    /* ダイアログを生成して表示し、OK されたら true を返す / Build, show dialog; return true on OK */
    function showSettingsDialog() {
        var dialog = new Window("dialog", getLocalizedText('dialog.title') + ' ' + SCRIPT_VERSION);
        dialog.orientation = "column";
        dialog.alignChildren = "fill";

        // モード / Mode（1カラム・ラジオ横並び）/ Mode (single column, radios in a row)
        var modePanel = dialog.add("panel", undefined, getLocalizedText('mode.label'));
        setupPanel(modePanel, 6);
        modePanel.orientation = "row";
        modePanel.alignChildren = ["left", "center"];
        var modeGroupRadio = modePanel.add("radiobutton", undefined, getLocalizedText('mode.group'));        // グループ
        var modeAutoPairRadio = modePanel.add("radiobutton", undefined, getLocalizedText('mode.auto'));      // 自動ペア認識
        var modeArtboardRadio = modePanel.add("radiobutton", undefined, getLocalizedText('mode.artboard'));  // アートボード
        modeGroupRadio.helpTip = getLocalizedText('tip.modeGroup');
        modeAutoPairRadio.helpTip = getLocalizedText('tip.modeAutoPair');
        modeArtboardRadio.helpTip = getLocalizedText('tip.modeArtboard');
        // 既定モードは選択内容で決める：グループのみ→グループ、それ以外→自動ペア認識（アートボードは手動選択）。
        // Default mode follows the selection: groups only → group, otherwise → auto pair (artboard is manual).
        modeGroupRadio.value = selectionIsGroupsOnly;
        modeAutoPairRadio.value = !selectionIsGroupsOnly;

        // キーオブジェクト と 位置調整 を2カラムで左右に並べる / Key object + Position side by side (two columns)
        var keyPositionColumns = dialog.add("group");
        keyPositionColumns.orientation = "row";
        keyPositionColumns.alignChildren = ["fill", "fill"]; // 2パネルの高さをそろえる / Match panel heights
        keyPositionColumns.spacing = PANEL_SPACING;

        // キーオブジェクト / Key object（上・左・右・下を十字に配置 / arranged as a cross）
        var fixedSideRefs = buildFixedSidePanel(keyPositionColumns);
        var fixedSidePanel = fixedSideRefs.panel;
        var fixedRadios = fixedSideRefs.fixedRadios;
        var selectFixedRadio = fixedSideRefs.selectFixedRadio;
        var getFixedSide = fixedSideRefs.getFixedSide;
        var setFixedSide = fixedSideRefs.setFixedSide;

        // 位置調整 / Position（間隔・プレビュー境界）
        var gapPanelRefs = buildGapPanel(keyPositionColumns, initialGapPoints);
        var spacingInput = gapPanelRefs.spacingInput;
        var previewBoundsCheckbox = gapPanelRefs.previewBoundsCheckbox;
        var getSpacingInPoints = gapPanelRefs.getSpacingInPoints;
        var getBoundsType = gapPanelRefs.getBoundsType;

        // キー／位置の2パネルの高さをそろえる / Match the two panel heights
        fixedSidePanel.alignment = ["fill", "fill"];
        gapPanelRefs.panel.alignment = ["fill", "fill"];

        // 水平 / 垂直パネル（整列＋位置オフセット）/ Horizontal & Vertical panels (alignment + offset)
        var horizontalRefs = buildOrientationPanel(dialog, "h");
        var verticalRefs = buildOrientationPanel(dialog, "v");
        var hAlign = horizontalRefs.align;
        var vAlign = verticalRefs.align;
        var offsetHorizontalInput = horizontalRefs.offsetInput;
        var offsetVerticalInput = verticalRefs.offsetInput;
        var offsetHorizontalRow = horizontalRefs.offsetRow;
        var offsetVerticalRow = verticalRefs.offsetRow;
        var getOffsetHorizontalPoints = horizontalRefs.getOffsetInPoints;
        var getOffsetVerticalPoints = verticalRefs.getOffsetInPoints;

        // テキストの行揃え / Text alignment（自動 / 左 / 中央 / 右 / 均等配置）
        var justifyRefs = buildJustifyPanel(dialog);
        var justifyRadios = justifyRefs.radios;
        var getJustifyMode = justifyRefs.getJustifyMode;

        /* キーオブジェクトの側からギャップが垂直か（上下キー）を判定 / Gap is vertical when key is top/bottom */
        function isVerticalGap() {
            var side = getFixedSide();
            return side === "top" || side === "bottom";
        }
        /* キー側に応じて有効なパネル（水平/垂直）を切り替え、整列「中央」なら対応オフセットを 0＋無効にする。
           Switch the active panel (horizontal/vertical) by the key side; zero & disable the offset
           when that panel's alignment is center. */
        function updateActivePanels() {
            var vertical = isVerticalGap();
            // 上下キー（縦並び）→ 水平パネル有効、左右キー（横並び）→ 垂直パネル有効
            horizontalRefs.panel.enabled = vertical;
            verticalRefs.panel.enabled = !vertical;
            var hCenter = hAlign.radios.center.value;
            var vCenter = vAlign.radios.center.value;
            offsetHorizontalRow.enabled = vertical && !hCenter;
            if (vertical && hCenter) offsetHorizontalInput.text = "0";
            offsetVerticalRow.enabled = !vertical && !vCenter;
            if (!vertical && vCenter) offsetVerticalInput.text = "0";
        }
        /* 現在有効な整列パネルの値を "none"/"start"/"center"/"end" で返す / Active alignment value */
        function getAlignMode() {
            var radios = isVerticalGap() ? hAlign.radios : vAlign.radios;
            if (radios.start.value) return "start";
            if (radios.center.value) return "center";
            if (radios.end.value) return "end";
            return "none";
        }

        // 前回終了時の設定をすべて復元（モード・キー・プレビュー境界・整列・間隔・左右/上下オフセット）。
        // 数値は pt で保存しているので、現在のルーラー単位の表示へ戻す。
        // Restore all last-used settings (mode, key, preview-bounds, alignment, gap, offsets).
        // Numeric values are stored in pt, so convert back to the current ruler unit for display.
        var savedSettings = loadSettings();
        /* pt 文字列を現在の単位の表示文字列へ。無効なら null / pt string → display string in the current unit (null if invalid) */
        function savedPtToDisplay(ptString) {
            if (ptString === undefined || ptString === null || ptString === "") return null;
            var pt = parseFloat(ptString);
            if (isNaN(pt)) return null;
            return String(Math.round((pt / pointsPerUnit) * 100) / 100);
        }
        // モード（選択に依存せず保存値を優先）/ Mode (prefer the saved value)
        if (savedSettings.mode === "group") modeGroupRadio.value = true;
        else if (savedSettings.mode === "artboard") modeArtboardRadio.value = true;
        else if (savedSettings.mode === "auto") modeAutoPairRadio.value = true;
        if (savedSettings.fixedSide) setFixedSide(savedSettings.fixedSide);
        if (savedSettings.previewBounds === "true") previewBoundsCheckbox.value = true;
        applySavedAlign(hAlign, savedSettings.alignH);
        applySavedAlign(vAlign, savedSettings.alignV);
        // テキストの行揃え / Text alignment
        if (savedSettings.justify && justifyRadios[savedSettings.justify]) {
            justifyRadios[savedSettings.justify].value = true;
        }
        // 数値（間隔・左右・上下）/ Numeric values (gap, horizontal/vertical offsets)
        var savedGapDisplay = savedPtToDisplay(savedSettings.gap);
        if (savedGapDisplay !== null) spacingInput.text = savedGapDisplay;
        var savedOffsetHDisplay = savedPtToDisplay(savedSettings.offsetH);
        if (savedOffsetHDisplay !== null) offsetHorizontalInput.text = savedOffsetHDisplay;
        var savedOffsetVDisplay = savedPtToDisplay(savedSettings.offsetV);
        if (savedOffsetVDisplay !== null) offsetVerticalInput.text = savedOffsetVDisplay;

        /* 現在のモードを取得する / Get the current mode ("group" | "artboard" | "auto") */
        function getMode() {
            if (modeGroupRadio.value) return "group";
            if (modeArtboardRadio.value) return "artboard";
            return "auto";
        }

        /* モードを切り替えてペアを組み直し、プレビューを更新する / Switch mode, rebuild pairs, refresh.
           モードで対象オブジェクトが変わるので行揃えを元へ戻してから組み直す / object set changes, so revert justify first */
        function onModeChange() {
            undoJustifications();
            buildPairs(getMode());
            refreshPreview();
        }

        /* 現在の設定でプレビューを更新する（行揃え・整列とも最終結果と一致）/ Refresh preview (matches final result) */
        function refreshPreview() {
            runPreview(getFixedSide(), getSpacingInPoints(), getBoundsType(), getAlignMode(),
                getOffsetHorizontalPoints(), getOffsetVerticalPoints(), getJustifyMode());
        }

        /* 行揃えの対象・向きが変わりうる操作（モード／キー／整列の切替）用：先に行揃えを元へ戻してから
           プレビューを更新する。これで「なし」へ戻したときや別の向きへ変えたときに正しく反映される。
           For changes that can alter the justification target (mode/key/alignment): revert justification
           first, then refresh — so switching to "none" or another direction is reflected correctly. */
        function refreshPreviewResetJustify() {
            undoJustifications();
            refreshPreview();
        }

        /* 整列の変更時：中央↔それ以外でオフセットの有効/無効が変わるので更新してからプレビュー。
           On alignment change: refresh offset enabled state (center toggles it), then preview. */
        function onAlignChange() {
            updateActivePanels();
            refreshPreviewResetJustify();
        }

        // 設定変更でライブプレビュー / Update preview on change
        modeGroupRadio.onClick = onModeChange;
        modeAutoPairRadio.onClick = onModeChange;
        modeArtboardRadio.onClick = onModeChange;
        // 固定側のラジオ：手動で排他にしてからプレビュー更新（軸の切替もここで反映）
        // Fixed-side radios: enforce exclusivity by hand, then refresh (axis switch applies here too)
        for (var i = 0; i < fixedRadios.length; i++) {
            fixedRadios[i].onClick = (function (radio) {
                return function () { selectFixedRadio(radio); updateActivePanels(); refreshPreviewResetJustify(); };
            })(fixedRadios[i]);
        }
        changeValueByArrowKey(spacingInput, refreshPreview); // ↑↓キーで増減＋プレビュー更新 / Arrow keys + preview
        spacingInput.onChanging = refreshPreview;
        // 左右/上下オフセット：↑↓キーで増減＋入力でプレビュー更新 / Offsets: arrow-key step + refresh on change
        changeValueByArrowKey(offsetHorizontalInput, refreshPreview);
        changeValueByArrowKey(offsetVerticalInput, refreshPreview);
        offsetHorizontalInput.onChanging = refreshPreview;
        offsetVerticalInput.onChanging = refreshPreview;
        previewBoundsCheckbox.onClick = refreshPreview;
        // 整列ラジオ（左右/上下の各行）：オフセットの有効/無効を更新し、行揃えを戻してから更新 / Alignment radios
        var alignRows = [hAlign, vAlign];
        for (var i = 0; i < alignRows.length; i++) {
            var alignRadios = alignRows[i].radios;
            for (var key in alignRadios) { alignRadios[key].onClick = onAlignChange; }
        }
        // 整列のキーボードショートカット（水平 L/C/R・垂直 T/M/B）。有効な向きだけ反応 / Alignment keyboard shortcuts (active orientation only)
        addAlignmentKeyHandler(dialog, hAlign, vAlign, isVerticalGap, onAlignChange);
        // テキストの行揃えラジオ：行揃えを戻してから再適用 / Justification radios: revert justify, then refresh
        for (var key in justifyRadios) { justifyRadios[key].onClick = refreshPreviewResetJustify; }

        // ボタン（Mac 規約：Cancel → OK）/ Buttons (Mac order: Cancel → OK)
        var buttonRow = dialog.add("group");
        buttonRow.alignment = "right";
        buttonRow.add("button", undefined, getLocalizedText('button.cancel'), { name: "cancel" });
        buttonRow.add("button", undefined, "OK", { name: "ok" });

        // ダイアログ表示時に既定モードでペアを組んで初回プレビュー（同期側 undo を避けて onShow から起動）
        // Build pairs for the default mode, then run the first preview (from onShow to avoid sync undo)
        dialog.onShow = function () {
            updateActivePanels(); // 既定のキー側に合わせて水平/垂直パネルとオフセットの有効/無効を初期化 / Init enabled state
            buildPairs(getMode());
            refreshPreview();
        };

        var accepted = (dialog.show() === 1);
        if (accepted) {
            // プレビュー状態がそのまま最終結果（行揃え・整列とも反映済み）なので、確定処理は保存のみ。
            // The preview already is the final result (justification + alignment), so OK just saves.
            // OK時に現在の設定をすべて保存（数値は pt で保存）/ Save all settings on OK (numeric values in pt)
            saveSettings({
                mode: getMode(),
                fixedSide: getFixedSide(),
                previewBounds: previewBoundsCheckbox.value ? "true" : "false",
                alignH: getAlignRowValue(hAlign),
                alignV: getAlignRowValue(vAlign),
                justify: getJustifyMode(),
                gap: String(getSpacingInPoints()),
                offsetH: String(getOffsetHorizontalPoints()),
                offsetV: String(getOffsetVerticalPoints())
            });
        }
        return accepted;
    }

    if (showSettingsDialog()) {
        // OK：プレビューをそのまま確定 / Keep the applied preview
    } else {
        // キャンセル：位置と行揃えを巻き戻す / Revert the preview (position + justification)
        undoPreview();
        undoJustifications();
        app.redraw();
    }
})();
