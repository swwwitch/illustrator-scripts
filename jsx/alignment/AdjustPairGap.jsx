#target illustrator
#targetengine "AdjustPairGap"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したオブジェクトの間隔と位置を、指定した値にそろえます。
グループ内の等間隔配置・最も近いもの同士のペア・アートボード端からのマージンの3モードがあり、
［固定］で選んだ側は動かさず、残りをライブプレビューで確認しながら動かします。

詳細は README を参照してください。

### Overview

Sets the gap and the position of the selected objects to a value you specify.
Three modes — even spacing inside a group, nearest-neighbour pairs, and margins from an
artboard edge — hold the side picked as the key object and move the rest, with a live preview.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AdjustPairGap";                /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.3.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-06-08";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-06";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AdjustPairGap.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AdjustPairGap.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/nc8fab19d8164"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

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
        // 位置調整パネル内の行ラベル / Row labels inside the Position panel
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
        // 位置調整（整列・オフセット）/ Position (alignment and offset)
        align: {
            label: { ja: "位置調整", en: "Position" },
            none: { ja: "なし", en: "None" },
            center: { ja: "中央", en: "Center" }
        },
        // オフセット（間隔・プレビュー境界）/ Offset (gap and preview bounds)
        options: {
            label: { ja: "オフセット", en: "Offset" },
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

    /* 単位を括弧で添えたタイトル（日本語は全角括弧、英語は半角）。各行に単位を並べる代わりに
       パネル名へまとめる / Title with the unit in parentheses (full-width in JA), so the rows
       themselves don't need to repeat it */
    function labelWithUnit(key, unit) {
        return getLocalizedText(key) + (currentLanguage === 'ja' ? '（' + unit + '）' : ' (' + unit + ')');
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

    /* 整列のキーボードショートカットをダイアログに付ける。整列パネルは1枚なので、同じラジオを
       今の向きで読み替える。水平（上下キー時）：L=左 / C=中央 / R=右、垂直（左右キー時）：T=上 / M=中央 / B=下。
       Attach the alignment keyboard shortcuts. There is a single alignment panel, so the same radios
       are read according to the current orientation, which isHorizontalActive() reports.
       Calls onChange on change. */
    function addAlignmentKeyHandler(dialog, radios, isHorizontalActive, onChange) {
        dialog.addEventListener("keydown", function (event) {
            // Cmd+C などの修飾キー付きの入力は横取りしない（コピー等を潰さないため）
            // Do not swallow modified keystrokes such as Cmd+C
            var keyboard = ScriptUI.environment.keyboardState;
            if (keyboard.metaKey || keyboard.ctrlKey || keyboard.altKey || keyboard.shiftKey) return;

            var target = null;
            var key = event.keyName;
            if (isHorizontalActive()) {
                if (key === "L") target = radios.start;
                else if (key === "C") target = radios.center;
                else if (key === "R") target = radios.end;
            } else {
                if (key === "T") target = radios.start;
                else if (key === "M") target = radios.center;
                else if (key === "B") target = radios.end;
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

    /* オフセットパネル（間隔の入力・プレビュー境界）を生成する。整列後のずらし量は［位置調整］パネルが
       持つので、ここは間隔だけを扱う。間隔は現在のルーラー単位で表示し（単位はパネル名に出す）、
       getSpacingInPoints() が pt に換算して返す。
       initialGapPoints は間隔の初期値・空欄時のフォールバック（pt）。イベント結線は呼び出し側で行う。
       Build the Offset panel (gap input, preview-bounds). The post-alignment nudge lives in the
       Position panel, so this one handles only the gap; its unit is shown in the panel title.
       Event wiring is left to the caller.
       返り値 / Returns: { panel, spacingInput, previewBoundsCheckbox,
                           getSpacingInPoints, getBoundsType } */
    function buildGapPanel(parentGroup, initialGapPoints) {
        var panel = parentGroup.add("panel", undefined, labelWithUnit('options.label', rulerUnitLabel));
        setupPanel(panel, 6);

        // 間隔行（ラベル＋入力）。単位はパネル名に出しているので行には並べない
        // Gap row (label + input); the unit lives in the panel title instead
        var spacingRow = panel.add("group");
        setupGroup(spacingRow, "row");
        spacingRow.alignment = "left"; // 広げず左寄せ / Keep at natural width, packed left
        spacingRow.add("statictext", undefined, labelText('spacing.label'));
        var defaultSpacingDisplay = Math.round((initialGapPoints / pointsPerUnit) * 100) / 100;
        var spacingInput = spacingRow.add("edittext", undefined, String(defaultSpacingDisplay));
        spacingInput.characters = 4;
        spacingInput.helpTip = getLocalizedText('tip.spacing');

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

    /* 位置調整パネルを生成する。整列（なし/開始/中央/終端）と位置の2行を1枚にまとめ、［固定］で選んだ側に
       応じて setOrientation() で水平／垂直に切り替える（開始/終端のラベルとツールチップが入れ替わる）。
       どちらの向きかは 左・右／上・下 のラベルで示すので、パネル名は「位置調整」で固定
       （オフセットパネルと同じく、単位はパネル名に出す）。
       向きごとの値の保持と中央時のオフセット無効化は呼び出し側で行う。
       開始/終端のラベルはレイアウト確定後に差し替えるので、文字数の多いほうで組み立てて幅を確保しておく。
       Build the Position panel: an alignment row (none/start/center/end) and a position row.
       setOrientation() switches it between horizontal and vertical (start/end labels and tooltips);
       the left/right vs top/bottom labels show which orientation is live, so the panel title stays put.
       Keeping the per-orientation values and disabling the offset on center are the caller's job.
       The start/end radios are built with the longer of the two candidate labels, so swapping the text
       after layout never clips it.
       返り値 / Returns: { alignRadios, offsetRow, offsetInput, setOrientation } */
    function buildAlignmentPanel(parentGroup) {
        /* 文字数の多いほうを返す（ラベル領域の確保用）/ The longer of the two (to reserve label width) */
        function widerText(a, b) {
            return (a.length >= b.length) ? a : b;
        }

        var panel = parentGroup.add("panel", undefined, labelWithUnit('align.label', rulerUnitLabel));
        setupPanel(panel, 6);

        // 整列行（ラベル＋なし/開始/中央/終端）/ Alignment row (label + none/start/center/end)
        var alignRow = panel.add("group");
        setupGroup(alignRow, "row");
        alignRow.alignment = "left";
        alignRow.alignChildren = ["left", "center"];
        var alignLabel = alignRow.add("statictext", undefined, labelText('panel.align'));
        var radios = {
            none: alignRow.add("radiobutton", undefined, getLocalizedText('align.none')),
            start: alignRow.add("radiobutton", undefined,
                widerText(getLocalizedText('fixedSide.left'), getLocalizedText('fixedSide.top'))),
            center: alignRow.add("radiobutton", undefined, getLocalizedText('align.center')),
            end: alignRow.add("radiobutton", undefined,
                widerText(getLocalizedText('fixedSide.right'), getLocalizedText('fixedSide.bottom')))
        };
        radios.none.value = true; // 既定：整列なし / Default: no alignment

        // 位置行（ラベル＋入力）。単位はパネル名に出しているので行には並べない
        // Position row (label + input); the unit lives in the panel title instead
        var offsetRow = panel.add("group");
        setupGroup(offsetRow, "row");
        offsetRow.alignment = "left";
        offsetRow.alignChildren = ["left", "center"];
        var offsetLabel = offsetRow.add("statictext", undefined, labelText('panel.position'));
        var offsetInput = offsetRow.add("edittext", undefined, "0");
        offsetInput.characters = 4;

        // ラベル幅をそろえて整列ラジオと入力の開始位置を合わせる / Match label widths
        var labelWidth = Math.max(alignLabel.preferredSize.width, offsetLabel.preferredSize.width);
        alignLabel.preferredSize.width = labelWidth;
        offsetLabel.preferredSize.width = labelWidth;

        /* 水平／垂直を切り替える（開始/終端のラベルとツールチップ）。ラベルの差し替えで
           レイアウトは組み直さないので、レイアウト確定後（onShow 以降）に呼ぶこと。
           Switch the orientation (start/end labels and tooltips). Changing the text does not
           re-run layout, so call this once the layout is settled (from onShow onward). */
        function setOrientation(isHorizontal) {
            radios.start.text = getLocalizedText(isHorizontal ? 'fixedSide.left' : 'fixedSide.top');
            radios.end.text = getLocalizedText(isHorizontal ? 'fixedSide.right' : 'fixedSide.bottom');
            var alignTip = getLocalizedText(isHorizontal ? 'tip.alignH' : 'tip.alignV');
            for (var key in radios) { radios[key].helpTip = alignTip; }
            offsetInput.helpTip = getLocalizedText(isHorizontal ? 'tip.offsetHorizontal' : 'tip.offsetVertical');
        }

        return {
            alignRadios: radios,
            offsetRow: offsetRow,
            offsetInput: offsetInput,
            setOrientation: setOrientation
        };
    }

    /* 表示文字列（現在のルーラー単位）を pt に換算する。空・不正は 0
       Convert a display string in the current ruler unit to pt (empty/invalid → 0) */
    function offsetToPoints(text) {
        var value = parseFloat(text);
        if (isNaN(value)) value = 0;
        return value * pointsPerUnit;
    }

    // =========================================
    // 行揃えボタン / Justification buttons
    // ScriptUI の button では選択状態を表示できないので、背景とアイコンを onDraw で自前描画する。
    // 描画方式は UnifiedTypePanel.jsx にそろえる。
    // ScriptUI buttons cannot show a selected state, so the background and icon are drawn in onDraw;
    // the drawing follows UnifiedTypePanel.jsx.
    // =========================================
    var JUSTIFY_BUTTON_SIZE = [26, 26]; /* ボタンの幅・高さ(px) / button width and height (px) */

    /* 環境設定のUI明るさが明るい側か / Whether the UI brightness preference is on the light side
       uiBrightness は 0（最暗）〜1（最明）。0.5（やや暗め）は暗い側に含めるため 0.5 超で判定 */
    function isLightUI() {
        try {
            return app.preferences.getRealPreference("uiBrightness") > 0.5;
        } catch (e) {
            return false;
        }
    }

    /* テーマ＋アクティブ状態に応じたボタン配色 / Button colors per theme and active state */
    function getJustifyColors(isLight, isActive) {
        if (isLight) {
            return {
                bg: isActive ? [0.40, 0.40, 0.40, 1] : [1, 1, 1, 1],
                border: isActive ? [0.30, 0.30, 0.30, 1] : [0.62, 0.62, 0.62, 1],
                line: isActive ? [1, 1, 1, 1] : [0.25, 0.25, 0.25, 1]
            };
        }
        return {
            bg: isActive ? [0.92, 0.92, 0.92, 1] : [0.30, 0.30, 0.30, 1],
            border: null,
            line: isActive ? [0.16, 0.16, 0.16, 1] : [0.82, 0.82, 0.82, 1]
        };
    }

    /* 行ごとの線幅。均等配置だけ最終行以外を長くする / Per-line widths (justify keeps all but the last line long) */
    function getJustifyLineWidths(iconType, longWidth, shortWidth) {
        if (iconType === "full") return [longWidth, longWidth, longWidth, shortWidth];
        return [longWidth, shortWidth, longWidth, shortWidth];
    }

    /* 行の開始 X（左／中央／右）/ Line start X (left/center/right) */
    function getJustifyLineX(iconType, buttonWidth, lineWidth) {
        var margin = 5;
        if (iconType === "right") return buttonWidth - margin - lineWidth;
        if (iconType === "center") return Math.round((buttonWidth - lineWidth) / 2);
        return margin;
    }

    /* アイコンの罫線を描く / Draw the icon's lines */
    function drawJustifyIconLines(graphics, iconType, buttonWidth, lineColor) {
        var pen = graphics.newPen(graphics.PenType.SOLID_COLOR, lineColor, 1.2);
        var rowYs = [7, 11, 15, 19];
        var lineWidths = getJustifyLineWidths(iconType, 15, 10);
        for (var i = 0; i < rowYs.length; i++) {
            var lineWidth = lineWidths[i];
            var lineStartX = getJustifyLineX(iconType, buttonWidth, lineWidth);
            graphics.newPath();
            graphics.moveTo(lineStartX, rowYs[i]);
            graphics.lineTo(lineStartX + lineWidth, rowYs[i]);
            graphics.strokePath(pen);
        }
    }

    /* 「自動」のアイコンを描く。他の行揃えと同じ高さに収まる「A」を線で描く
       （drawString はフォントの解決に環境差があり空欄になることがあるため使わない）。
       Draw the "auto" icon: an "A" built from strokes, spanning the same rows as the other icons
       (drawString is avoided — resolving a font is environment-dependent and can render nothing). */
    function drawAutoIcon(graphics, buttonWidth, lineColor) {
        var pen = graphics.newPen(graphics.PenType.SOLID_COLOR, lineColor, 1.2);
        var centerX = Math.round(buttonWidth / 2);
        var topY = 7;
        var bottomY = 19;
        var halfWidth = 5;
        var crossY = 15; // 横棒。斜線上の位置に合わせて幅を決める / crossbar, width taken from the diagonals
        var crossHalf = Math.round(halfWidth * (crossY - topY) / (bottomY - topY));
        graphics.newPath();
        graphics.moveTo(centerX - halfWidth, bottomY);
        graphics.lineTo(centerX, topY);
        graphics.lineTo(centerX + halfWidth, bottomY);
        graphics.strokePath(pen);
        graphics.newPath();
        graphics.moveTo(centerX - crossHalf, crossY);
        graphics.lineTo(centerX + crossHalf, crossY);
        graphics.strokePath(pen);
    }

    /* ボタン背景＋アイコン（または「自動」のラベル）を描く / Draw the button background + icon (or the "auto" label) */
    function drawJustifyButton(button, isActive, isLight) {
        var graphics = button.graphics;
        var colors = getJustifyColors(isLight, isActive);
        try {
            graphics.rectPath(0, 0, button.size[0], button.size[1]);
            graphics.fillPath(graphics.newBrush(graphics.BrushType.SOLID_COLOR, colors.bg));
            if (colors.border) {
                graphics.rectPath(0, 0, button.size[0], button.size[1]);
                graphics.strokePath(graphics.newPen(graphics.PenType.SOLID_COLOR, colors.border, 1));
            }
            if (button.iconType === "auto") {
                drawAutoIcon(graphics, button.size[0], colors.line);
            } else {
                drawJustifyIconLines(graphics, button.iconType, button.size[0], colors.line);
            }
        } catch (eDraw) {
            // 描画できない環境ではOS標準のボタン（ラベル付き）に任せる / Fall back to the OS control
            try { graphics.drawOSControl(); } catch (eOs) {}
        }
    }

    /* テキストの行揃えパネルを生成する。自動 / 左 / 中央 / 右 / 均等配置（最終行左）をボタンで並べる。
       getJustifyMode() で選択を文字列（"auto"/"left"/"center"/"right"/"full"）で返す。
       「自動」は整列・キーに連動（エリア内文字は均等配置）。解決は applySpacing 側で行う。イベント結線は呼び出し側。
       Build the "Text alignment" panel as a row of buttons (auto/left/center/right/justify).
       getJustifyMode() returns the selection as a string; "auto" links to align/key (resolved in
       applySpacing). Event wiring is left to the caller.
       返り値 / Returns: { panel, buttons, getJustifyMode, setJustifyMode } */
    function buildJustifyPanel(parentGroup) {
        var panel = parentGroup.add("panel", undefined, getLocalizedText('justify.label'));
        setupPanel(panel, 4);
        panel.orientation = "row"; // ボタン横並び / buttons in a row
        panel.alignChildren = ["center", "center"]; // ボタン列をパネルの左右中央に / Center the button row in the panel

        // アクティブな行揃えと UI 明暗を共有する（onDraw のクロージャから参照）
        // Shared active id + theme, read by the onDraw closures
        var state = { activeId: "auto", isLight: isLightUI() }; // 既定：自動（整列・キーに連動）/ Default: auto
        var options = [
            { id: "auto", tip: 'tip.justifyAuto' },
            { id: "left", tip: 'tip.justifyLeft' },
            { id: "center", tip: 'tip.justifyCenter' },
            { id: "right", tip: 'tip.justifyRight' },
            { id: "full", tip: 'tip.justifyFull' }
        ];
        var buttons = [];
        for (var i = 0; i < options.length; i++) {
            // ラベルは描画に失敗したときのフォールバック（drawOSControl）でも使う / text is also the fallback label
            var button = panel.add("button", undefined, getLocalizedText('justify.' + options[i].id));
            button.helpTip = getLocalizedText(options[i].tip);
            button.preferredSize = JUSTIFY_BUTTON_SIZE;
            button.minimumSize = JUSTIFY_BUTTON_SIZE;
            button.maximumSize = JUSTIFY_BUTTON_SIZE; // 伸ばさない / keep the fixed size
            button.justifyId = options[i].id;
            button.iconType = options[i].id;
            button.onDraw = function () { drawJustifyButton(this, this.justifyId === state.activeId, state.isLight); };
            buttons.push(button);
        }

        /* 選択中の行揃えモードを文字列で返す / Selected justify mode as a string */
        function getJustifyMode() {
            return state.activeId;
        }

        /* 行揃えモードを設定してボタンを描き直す。未知の値は無視 / Set the mode and redraw (unknown values ignored) */
        function setJustifyMode(mode) {
            var found = false;
            for (var i = 0; i < buttons.length; i++) {
                if (buttons[i].justifyId === mode) { found = true; break; }
            }
            if (!found) return;
            state.activeId = mode;
            for (var j = 0; j < buttons.length; j++) {
                try { buttons[j].notify("onDraw"); } catch (eDraw) {}
            }
            // notify だけでは画面に反映されないことがあるので、ウィンドウの再描画も要求する
            // notify alone may not reach the screen, so ask the window to repaint as well
            try { panel.window.update(); } catch (eUpdate) {}
        }

        return { panel: panel, buttons: buttons, getJustifyMode: getJustifyMode, setJustifyMode: setJustifyMode };
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

    /* テキストの段落行揃えを設定する。ポイント文字は行揃えを変えるとアンカー基準で組み直されて
       フレームが動く。見た目の位置は間隔・整列の計算で決めるので、どの値でも元の位置へ戻す
       （戻さないと行揃えを往復するたびに字幅の半分ずつずれ、キャンセルしても戻らない）。
       Justification.LEFT は代入が無視される Illustrator のバグがあるので、一時 resize（200%→50%）で
       段落属性をリフレッシュしてから代入する。
       Set paragraph justification. Changing it re-lays point text around its anchor and moves the
       frame, so the position is saved and restored for every value (otherwise toggling justification
       drifts the text by half its width each time and cancel cannot undo it). Assigning
       Justification.LEFT is ignored by Illustrator, so a temporary resize (200% then 50%) refreshes
       the paragraph attributes first. */
    function setParagraphJustification(textFrame, justification) {
        var savedPosition = [textFrame.position[0], textFrame.position[1]];
        if (justification === Justification.LEFT) {
            textFrame.resize(200, 200);
            textFrame.textRange.paragraphAttributes.justification = Justification.LEFT;
            textFrame.resize(50, 50);
        } else {
            textFrame.textRange.paragraphAttributes.justification = justification;
        }
        try { textFrame.position = savedPosition; } catch (ePos) {}
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

    /* 選択オブジェクトの現在の間隔の平均（pt）を求める。測れない場合は null。
       モードに合わせて測る：グループは各グループ内の隣接間隔、アートボードは固定側の端との距離
       （マージン）、自動ペア認識は各ペアの間隔。測る軸と固定端は fixedSide から決める。
       常に geometricBounds（幾何境界）基準。Average current gap (pt) for the given mode; the axis
       and the anchor end follow fixedSide. Null if nothing measurable. Always geometric bounds. */
    function computeAverageGap(selectedItems, mode, fixedSide) {
        var axis = axisForSide(fixedSide);
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
        } else if (mode === "artboard") {
            // アートボード：各オブジェクトと固定側のアートボード端との距離（マージン）を測る
            // Artboard: distance (margin) from each object to the fixed artboard edge
            var anchorEnd = isAnchorEnd(fixedSide);
            for (var i = 0; i < selectedItems.length; i++) {
                var itemBounds = geometricBoundsOf(selectedItems[i]);
                var rect = artboardRectFor(selectedItems[i]);
                gaps.push(anchorEnd ? axis.end(rect) - axis.end(itemBounds)
                    : axis.start(itemBounds) - axis.start(rect));
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
    // OK で設定を覚え、次回開いたときに復元する
    // （モード・キー・プレビュー境界・整列・行揃え・間隔・左右/上下オフセット。数値は pt で保存）。
    // モードと間隔は選択内容に依存するため、同一セッション中だけ覚える（ファイルには残さない）。
    // 2段構えで記憶する：
    //   - #targetengine の常駐グローバル（$.global）… 同一セッション内は即座に前回状態を復元（ファイル I/O なし）
    //   - Folder.userData のファイル … Illustrator を再起動してもまたいで永続
    // Remember the settings on OK and restore them next time (mode, key side, preview-bounds,
    // alignment, justification, gap, and the horizontal/vertical offsets; numbers stored in pt).
    // Mode and gap depend on the selection, so they are kept for the session only (never written to
    // the file). Two layers: the #targetengine persistent global ($.global) restores instantly within
    // a session, and a Folder.userData file persists across Illustrator restarts.
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

    /* 整列ラジオで選択中の値（none/start/center/end）を返す / Selected value of the alignment radios */
    function getAlignValue(radios) {
        if (radios.start.value) return "start";
        if (radios.center.value) return "center";
        if (radios.end.value) return "end";
        return "none";
    }

    /* 保存済みの整列値をラジオへ反映する。未知の値は無視 / Apply a saved alignment value (ignore unknown) */
    function applySavedAlign(radios, value) {
        if (!radios || !value) return;
        var radio = radios[value];
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

        // 前回終了時の設定。初期間隔を「復元後のモード・キー側」で測るため、ダイアログ生成より先に読む。
        // Last-used settings, loaded before building the dialog so the initial gap can be measured
        // with the mode and key side that will actually be restored.
        var savedSettings = loadSettings();

        // 選択がすべてグループなら既定でグループモードにする（それ以外は自動ペア認識）。
        // Default to group mode when every selected object is a group; otherwise auto pair.
        var selectionIsGroupsOnly = true;
        for (var i = 0; i < selectedItems.length; i++) {
            if (selectedItems[i].typename !== "GroupItem") { selectionIsGroupsOnly = false; break; }
        }
        var defaultMode = selectionIsGroupsOnly ? "group" : "auto";

        // 実際に開くときのモードとキー側（保存値が有効ならそちらを優先）。
        // 固定側はファイルに永続する一方で間隔は永続しないので、再起動直後は必ずここで測り直される。
        // Mode and key side the dialog actually opens with (a valid saved value wins). The key side
        // persists across restarts while the gap does not, so the gap is always re-measured here.
        var initialMode = (savedSettings.mode === "group" || savedSettings.mode === "auto" ||
            savedSettings.mode === "artboard") ? savedSettings.mode : defaultMode;
        var initialFixedSide = (savedSettings.fixedSide === "top" || savedSettings.fixedSide === "left" ||
            savedSettings.fixedSide === "right" || savedSettings.fixedSide === "bottom")
            ? savedSettings.fixedSide : "right";

        // 間隔の初期値は選択オブジェクトの現在の平均間隔。測れなければ DEFAULT_GAP を使う。
        // 負（重なり）の場合は 0 にクランプ。Initial gap = current average gap of the selection
        // (clamped to >= 0); falls back to DEFAULT_GAP when nothing measurable.
        var measuredGap = computeAverageGap(selectedItems, initialMode, initialFixedSide);
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
        // 行揃えの変更で境界が変わったか。プレビューは毎回元位置へ巻き戻すので、境界が変わるのは
        // 行揃えを当てた／戻したときだけ。真のときだけ測り直す（全オブジェクトの visibleBounds
        // 再取得は重く、間隔欄の1打鍵ごとに走ると効いてくる）。
        // Whether justification changed the bounds. The preview always reverts to the original
        // positions, so only applying/reverting justification invalidates them; re-measure only then
        // (re-reading visibleBounds for every object on each keystroke is expensive).
        var boundsCacheStale = false;

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
            boundsCacheStale = true;
        }

        /* 記録した行揃え変更を元の値へ戻す / Restore recorded justification changes */
        function undoJustifications() {
            if (appliedJustifications.length === 0) return;
            for (var j = appliedJustifications.length - 1; j >= 0; j--) {
                setParagraphJustification(appliedJustifications[j].obj, appliedJustifications[j].original);
            }
            appliedJustifications = [];
            boundsCacheStale = true;
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

        /* グループのメンバー境界キャッシュを選ぶ（プレビュー境界 / 幾何境界）
           Pick the cached member bounds (visible or geometric) */
        function memberBounds(pair, useVisible) {
            return useVisible ? pair.memberVis : pair.memberGeo;
        }

        /* 境界を進行方向（先頭辺の昇順）に並べたインデックス配列を返す
           Indices into bounds, ordered by leading edge ascending */
        function memberOrder(bounds, axis) {
            var order = [];
            for (var i = 0; i < bounds.length; i++) order.push(i);
            order.sort(function (x, y) { return axis.start(bounds[x]) - axis.start(bounds[y]); });
            return order;
        }

        /* グループ内の全オブジェクトを進行方向順に並べ、隣り合う間隔を gap にそろえる。
           固定側のオブジェクトは動かさず、そこを起点にカスケードで再配置する。
           axis で水平/垂直を切り替える（左右→水平、上下→垂直）。
           Distribute all objects in a group along the axis with a uniform gap, anchored at
           the fixed side (the fixed-side object stays; the rest cascade from it). */
        function distributeGroup(pair, fixedSide, gapInPoints, useVisible, axis) {
            var members = pair.members;
            var cachedBounds = memberBounds(pair, useVisible);
            var count = members.length;
            var anchorEnd = isAnchorEnd(fixedSide);
            var order = memberOrder(cachedBounds, axis); // 先頭辺の昇順 / ordered by leading edge

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
            var order = memberOrder(memberBounds(pair, useVisible), gapAxis);
            return anchorEnd ? order[order.length - 1] : order[0];
        }

        /* グループの各メンバーを固定端のメンバー（アンカー）に整列軸方向でそろえる。
           行揃えは間隔調整より前に適用して境界を取り直す（cachePairBounds）ので、ここはキャッシュ境界で十分。
           Justification is applied and bounds re-cached before gap adjustment, so cached bounds suffice here. */
        function alignGroup(pair, gapAxis, anchorEnd, useVisible, alignAxis, alignMode) {
            if (alignMode === "none") return;
            var members = pair.members;
            var bounds = memberBounds(pair, useVisible);
            var anchorIndex = groupAnchorIndex(pair, gapAxis, anchorEnd, useVisible); // 固定端のメンバー / member at the fixed end
            var anchorBounds = bounds[anchorIndex];
            for (var i = 0; i < members.length; i++) {
                if (i === anchorIndex) continue;
                alignToAnchor(alignAxis, members[i], bounds[i], anchorBounds, alignMode);
            }
        }

        /* 設定値で各ペアの間隔を調整する / Apply the gap to every pair.
           固定側から軸（水平/垂直）と固定端（先頭/後ろ）を決める / Axis and anchor end follow the fixed side.
           alignMode は整列（ギャップ軸に直交する方向、固定オブジェクト基準）/ alignMode aligns on the
           axis perpendicular to the gap, anchored at the fixed object. */
        function applySpacing(fixedSide, gapInPoints, boundsType, alignMode, offsetAlong, justifyMode) {
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
            // 行揃えで境界が変わったときだけ取り直す / Re-cache only when justification changed the bounds
            if (boundsCacheStale) {
                cachePairBounds(objectPairs);
                boundsCacheStale = false;
            }

            // 位置オフセット（offsetAlong）はギャップ軸に直交する alignAxis 方向。
            // キーオブジェクトでない側（移動側）だけをずらす（右＝正／下＝正）。
            // The offset runs along alignAxis (perpendicular to the gap) and moves only the non-key object(s).

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
        function runPreview(fixedSide, gapInPoints, boundsType, alignMode, offsetAlong, justifyMode) {
            undoPreview();
            applySpacing(fixedSide, gapInPoints, boundsType, alignMode, offsetAlong, justifyMode);
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
            boundsCacheStale = false;
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
            modePanel.orientation = "column"; // ラジオ縦並び / radios stacked
            modePanel.alignChildren = ["left", "top"];
            var modeGroupRadio = modePanel.add("radiobutton", undefined, getLocalizedText('mode.group'));        // グループ
            var modeAutoPairRadio = modePanel.add("radiobutton", undefined, getLocalizedText('mode.auto'));      // 自動ペア認識
            var modeArtboardRadio = modePanel.add("radiobutton", undefined, getLocalizedText('mode.artboard'));  // アートボード
            modeGroupRadio.helpTip = getLocalizedText('tip.modeGroup');
            modeAutoPairRadio.helpTip = getLocalizedText('tip.modeAutoPair');
            modeArtboardRadio.helpTip = getLocalizedText('tip.modeArtboard');
            // 初期選択は下の復元ブロックで initialMode に従って入れる / Initial selection is set from initialMode below

            // キーオブジェクト と オフセット を2カラムで左右に並べる / Key object + Offset side by side (two columns)
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

            // オフセット / Offset（間隔・プレビュー境界）
            var gapPanelRefs = buildGapPanel(keyPositionColumns, initialGapPoints);
            var spacingInput = gapPanelRefs.spacingInput;
            var previewBoundsCheckbox = gapPanelRefs.previewBoundsCheckbox;
            var getSpacingInPoints = gapPanelRefs.getSpacingInPoints;
            var getBoundsType = gapPanelRefs.getBoundsType;

            // キー／位置の2パネルの高さをそろえる / Match the two panel heights
            fixedSidePanel.alignment = ["fill", "fill"];
            gapPanelRefs.panel.alignment = ["fill", "fill"];

            // 位置調整パネル（整列＋位置）。1枚で、［固定］に応じて水平／垂直に切り替わる
            // One Position panel (alignment + offset); it switches with the Key Object side
            var alignmentRefs = buildAlignmentPanel(dialog);
            var alignRadios = alignmentRefs.alignRadios;
            var offsetRow = alignmentRefs.offsetRow;
            var offsetInput = alignmentRefs.offsetInput;

            // 水平／垂直それぞれの入力内容。パネルを切り替えるときに退避・復元する
            // Per-orientation values, stashed and restored as the panel switches
            var orientationValues = {
                h: { align: "none", offset: "0" },
                v: { align: "none", offset: "0" }
            };
            var shownOrientation = null; // 今パネルに出ている向き / the orientation currently shown

            // テキストの行揃え / Text alignment（自動 / 左 / 中央 / 右 / 均等配置）
            var justifyRefs = buildJustifyPanel(dialog);
            var justifyButtons = justifyRefs.buttons;
            var getJustifyMode = justifyRefs.getJustifyMode;
            var setJustifyMode = justifyRefs.setJustifyMode;

            /* キーオブジェクトの側からギャップが垂直か（上下キー）を判定 / Gap is vertical when key is top/bottom */
            function isVerticalGap() {
                var side = getFixedSide();
                return side === "top" || side === "bottom";
            }
            /* パネルに出ている内容を、その向きの控えへ退避する / Stash the shown values into their orientation */
            function stashOrientationValues() {
                if (!shownOrientation) return;
                orientationValues[shownOrientation].align = getAlignValue(alignRadios);
                orientationValues[shownOrientation].offset = offsetInput.text;
            }
            /* キー側に合わせて整列パネルの向きと中身を入れ替え、整列「中央」ならオフセットを 0＋無効にする。
               上下キー（縦並び）→ 水平の整列、左右キー（横並び）→ 垂直の整列。
               Switch the alignment panel to match the key side, and zero & disable the offset on center.
               Key top/bottom (vertical stack) → horizontal alignment; key left/right → vertical. */
            function updateActivePanels() {
                var orientation = isVerticalGap() ? "h" : "v";
                if (orientation !== shownOrientation) {
                    stashOrientationValues();
                    shownOrientation = orientation;
                    alignmentRefs.setOrientation(orientation === "h");
                    applySavedAlign(alignRadios, orientationValues[orientation].align);
                    offsetInput.text = orientationValues[orientation].offset;
                }
                var isCenter = alignRadios.center.value;
                offsetRow.enabled = !isCenter;
                if (isCenter) offsetInput.text = "0";
            }
            /* 整列パネルの値を "none"/"start"/"center"/"end" で返す / Alignment value of the panel */
            function getAlignMode() {
                return getAlignValue(alignRadios);
            }

            // 前回終了時の設定をすべて復元（モード・キー・プレビュー境界・整列・行揃え・間隔・左右/上下オフセット）。
            // 数値は pt で保存しているので、現在のルーラー単位の表示へ戻す。
            // Restore all last-used settings (mode, key, preview-bounds, alignment, justification,
            // gap, offsets). Numeric values are stored in pt, so convert back to the current ruler unit.
            /* pt 文字列を現在の単位の表示文字列へ。無効なら null / pt string → display string in the current unit (null if invalid) */
            function savedPtToDisplay(ptString) {
                if (ptString === undefined || ptString === null || ptString === "") return null;
                var pt = parseFloat(ptString);
                if (isNaN(pt)) return null;
                return String(Math.round((pt / pointsPerUnit) * 100) / 100);
            }
            // モードとキー側（初期間隔を測ったときと同じ値を使う）/ Mode and key side (same values the gap was measured with)
            if (initialMode === "group") modeGroupRadio.value = true;
            else if (initialMode === "artboard") modeArtboardRadio.value = true;
            else modeAutoPairRadio.value = true;
            setFixedSide(initialFixedSide);
            if (savedSettings.previewBounds === "true") previewBoundsCheckbox.value = true;
            // 整列は向きごとの控えに入れておき、パネルを切り替えたときに反映する（キーは none/start/center/end）
            // Alignment goes into the per-orientation stash and lands when the panel switches
            if (alignRadios[savedSettings.alignH]) orientationValues.h.align = savedSettings.alignH;
            if (alignRadios[savedSettings.alignV]) orientationValues.v.align = savedSettings.alignV;
            // テキストの行揃え / Text alignment
            if (savedSettings.justify) setJustifyMode(savedSettings.justify);
            // 数値（間隔・左右・上下）/ Numeric values (gap, horizontal/vertical offsets)
            var savedGapDisplay = savedPtToDisplay(savedSettings.gap);
            if (savedGapDisplay !== null) spacingInput.text = savedGapDisplay;
            var savedOffsetHDisplay = savedPtToDisplay(savedSettings.offsetH);
            if (savedOffsetHDisplay !== null) orientationValues.h.offset = savedOffsetHDisplay;
            var savedOffsetVDisplay = savedPtToDisplay(savedSettings.offsetV);
            if (savedOffsetVDisplay !== null) orientationValues.v.offset = savedOffsetVDisplay;

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
                    offsetToPoints(offsetInput.text), getJustifyMode());
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
            // 位置オフセット：↑↓キーで増減＋入力でプレビュー更新 / Offset: arrow-key step + refresh on change
            changeValueByArrowKey(offsetInput, refreshPreview);
            offsetInput.onChanging = refreshPreview;
            previewBoundsCheckbox.onClick = refreshPreview;
            // 整列ラジオ：オフセットの有効/無効を更新し、行揃えを戻してから更新 / Alignment radios
            for (var key in alignRadios) { alignRadios[key].onClick = onAlignChange; }
            // 整列のキーボードショートカット（水平 L/C/R・垂直 T/M/B）。今の向きで読み替える / Alignment keyboard shortcuts
            addAlignmentKeyHandler(dialog, alignRadios, isVerticalGap, onAlignChange);
            // テキストの行揃えボタン：押した値をアクティブにし、行揃えを戻してから再適用
            // Justification buttons: activate the clicked value, revert justification, then refresh
            for (var i = 0; i < justifyButtons.length; i++) {
                justifyButtons[i].onClick = function () {
                    setJustifyMode(this.justifyId);
                    refreshPreviewResetJustify();
                };
            }

            // ボタン（Mac 規約：Cancel → OK）/ Buttons (Mac order: Cancel → OK)
            var btnRowGroup = dialog.add("group");
            btnRowGroup.orientation = "row";
            btnRowGroup.alignment = ["center", "bottom"]; // ボタンをダイアログの左右中央に / Center the buttons in the dialog
            btnRowGroup.alignChildren = ["center", "center"];
            var btnCancel = btnRowGroup.add("button", undefined, getLocalizedText('button.cancel'), { name: "cancel" });
            var btnOK = btnRowGroup.add("button", undefined, "OK", { name: "ok" });
            // 行揃えのボタンが増えたので Enter / ESC の行き先を明示する
            // Spell out where Enter / ESC go, now that the justification buttons are pushbuttons too
            dialog.defaultElement = btnOK;
            dialog.cancelElement = btnCancel;

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
                stashOrientationValues(); // 出ている向きの値を控えへ入れてから保存 / stash the shown values first
                saveSettings({
                    mode: getMode(),
                    fixedSide: getFixedSide(),
                    previewBounds: previewBoundsCheckbox.value ? "true" : "false",
                    alignH: orientationValues.h.align,
                    alignV: orientationValues.v.align,
                    justify: getJustifyMode(),
                    gap: String(getSpacingInPoints()),
                    offsetH: String(offsetToPoints(orientationValues.h.offset)),
                    offsetV: String(offsetToPoints(orientationValues.v.offset))
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

})();
