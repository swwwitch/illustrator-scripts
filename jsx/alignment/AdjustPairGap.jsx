#target illustrator
#targetengine "AdjustPairGap"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したオブジェクトの間隔と位置を、指定した値にそろえます。
グループ内の等間隔配置・最も近いもの同士のペア・アートボード端からのマージンの3モードがあり、
［固定］で選んだ側は動かさず、残りをライブプレビューで確認しながら動かします。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AdjustPairGap.md

note記事も参照してください。
https://note.com/dtp_tranist/n/nc8fab19d8164

### Overview

Sets the gap and the position of the selected objects to a value you specify.
Three modes — even spacing inside a group, nearest-neighbour pairs, and margins from an
artboard edge — hold the side picked as the key object and move the rest, with a live preview.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AdjustPairGap.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AdjustPairGap";                /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.3.3";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-06-08";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-22";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AdjustPairGap.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AdjustPairGap.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/nc8fab19d8164"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================
    var DEFAULT_GAP = 30; /* 平均間隔を測れないときのフォールバック（pt）/ Fallback gap when no average can be measured (pt) */

    // =========================================
    // レイアウト / Layout
    // =========================================
    var PANEL_MARGINS       = [16, 20, 16, 12]; /* パネル余白 [左,上,右,下] / Panel margins [left, top, right, bottom] */
    var PANEL_SPACING       = 8;                /* パネル内の要素間隔 / Spacing inside panels */
    var GRID_CELL_SIZE      = [22, 20];         /* ［固定］の十字のセルの幅・高さ (px) / Cross-grid cell width, height (px) */
    var JUSTIFY_BUTTON_SIZE = [26, 26];         /* 行揃えボタンの幅・高さ (px) / Justification button width, height (px) */

    /**
     * パネルの共通設定を適用する
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
     * グループの共通設定を適用する（row は横並びなので縦中央、column は縦並びなので左揃え）
     * @param {Group} targetGroup - 対象のグループ
     * @param {string} [orientation] - "row" または "column"（既定）
     * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function setupGroup(targetGroup, orientation, spacing) {
        var groupOrientation = orientation || "column";
        targetGroup.orientation = groupOrientation;
        targetGroup.alignChildren = (groupOrientation === "row") ? ["left", "center"] : ["left", "top"];
        targetGroup.alignment = "fill";
        targetGroup.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
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

    /* ダイアログの数値は定規の単位で表示し、pt に換算して使う / Dialog values are shown in the ruler unit and converted to pt */
    var rulerUnit = getUnitInfo();
    var rulerUnitLabel = rulerUnit.label;
    var pointsPerUnit = rulerUnit.pointsPerUnit; /* 1単位 = pointsPerUnit pt */

    /**
     * pt の値を定規の単位の表示文字列にする（小数第2位で丸める）
     * @param {number} points - pt の値
     * @returns {string} 表示文字列
     */
    function pointsToDisplayText(points) {
        return String(Math.round((points / pointsPerUnit) * 100) / 100);
    }

    /**
     * 表示文字列（定規の単位）を pt に換算する。空・不正は 0
     * @param {string} displayText - 入力欄の文字列
     * @returns {number} pt の値
     */
    function displayTextToPoints(displayText) {
        var value = parseFloat(displayText);
        if (isNaN(value)) value = 0;
        return value * pointsPerUnit;
    }

    /**
     * 保存した pt の文字列を定規の単位の表示文字列へ戻す。無効なら null
     * @param {string} pointsText - 保存済みの pt の文字列
     * @returns {string|null} 表示文字列
     */
    function savedPointsToDisplayText(pointsText) {
        if (pointsText === undefined || pointsText === null || pointsText === "") return null;
        var points = parseFloat(pointsText);
        if (isNaN(points)) return null;
        return pointsToDisplayText(points);
    }

    // =========================================
    // ローカライズ / Localization
    // =========================================
    var uiLang = ($.locale && $.locale.indexOf("ja") === 0) ? "ja" : "en";

    var LABELS = {
        dialog: {
            title: { ja: "ペア配置の調整", en: "Adjust Pair Layout" }
        },
        panel: {
            mode: { ja: "モード", en: "Mode" },
            fixedSide: { ja: "固定", en: "Key Object" },
            offset: { ja: "オフセット", en: "Offset" },
            position: { ja: "位置調整", en: "Position" },
            justify: { ja: "テキストの行揃え", en: "Text alignment" }
        },
        radio: {
            modeGroup: { ja: "グループ", en: "Group" },
            modeAutoPair: { ja: "自動ペア認識", en: "Auto Pair Detection" },
            modeArtboard: { ja: "アートボード", en: "Artboard" },
            alignNone: { ja: "なし", en: "None" },
            alignLeft: { ja: "左", en: "Left" },
            alignTop: { ja: "上", en: "Top" },
            alignCenter: { ja: "中央", en: "Center" },
            alignRight: { ja: "右", en: "Right" },
            alignBottom: { ja: "下", en: "Bottom" }
        },
        checkbox: {
            previewBounds: { ja: "プレビュー境界", en: "Preview Bounds" }
        },
        fieldLabel: {
            spacing: { ja: "間隔", en: "Gap" },
            align: { ja: "整列", en: "Align" },
            position: { ja: "位置", en: "Position" }
        },
        /* OK はローカライズしない / "OK" is not localized */
        button: {
            justifyAuto: { ja: "自動", en: "Auto" },
            justifyLeft: { ja: "左", en: "Left" },
            justifyCenter: { ja: "中央", en: "Center" },
            justifyRight: { ja: "右", en: "Right" },
            justifyFull: { ja: "均等配置", en: "Justify" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        tooltip: {
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
            justifyLeft: { ja: "テキストを左揃えにします。", en: "Left-aligns the text." },
            justifyCenter: { ja: "テキストを中央揃えにします。", en: "Center-aligns the text." },
            justifyRight: { ja: "テキストを右揃えにします。", en: "Right-aligns the text." },
            justifyFull: { ja: "テキストを均等配置（最終行左）にします。", en: "Justifies the text (last line left-aligned)." }
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            selectTwo: { ja: "オブジェクトを2つ以上選択してください。", en: "Select at least two objects." }
        }
    };

    /**
     * LABELS からドット区切りのパスで表示言語のテキストを取り出す
     * @param {string} labelPath - "panel.mode" のようなドット区切りのキー
     * @returns {string} 表示言語のテキスト（見つからない場合は labelPath をそのまま返す）
     */
    function getLabel(labelPath) {
        var pathKeys = labelPath.split(".");
        var labelNode = LABELS;
        for (var i = 0; i < pathKeys.length; i++) {
            labelNode = labelNode[pathKeys[i]];
            if (!labelNode) return labelPath;
        }
        return labelNode[uiLang] || labelNode.en;
    }

    /**
     * コロン付きの項目名を返す（日本語は全角、英語は半角）
     * @param {Object|string} labelSet - ラベル、またはラベルのパス
     * @returns {string} コロン付きの項目名
     */
    function labelText(labelSet) {
        return getLabel(labelSet) + (uiLang === "ja" ? "：" : ":");
    }

    /**
     * 単位を括弧で添えたパネル名を返す（日本語は全角括弧、英語は半角）。
     * 各行に単位を並べる代わりにパネル名へまとめる
     * @param {string} labelPath - パネル名のラベルのパス
     * @param {string} unitLabel - 単位の表示ラベル
     * @returns {string} 単位付きのパネル名
     */
    function labelWithUnit(labelPath, unitLabel) {
        return getLabel(labelPath) + (uiLang === "ja" ? "（" + unitLabel + "）" : " (" + unitLabel + ")");
    }

    // =========================================
    // キー操作 / Keyboard
    // =========================================

    /**
     * テキストフィールドで↑↓キーによる増減を有効化する。onUpdate を渡すと値変更後に呼ぶ（プレビュー更新用）
     * @param {EditText} editText - 対象の入力欄
     * @param {Function} [onUpdate] - 値を変えたあとに呼ぶ関数
     * @returns {void}
     */
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

    /**
     * 整列のキーボードショートカットをダイアログに付ける。整列パネルは1枚なので、同じラジオを
     * 今の向きで読み替える。水平（上下キー時）：L=左 / C=中央 / R=右、垂直（左右キー時）：T=上 / M=中央 / B=下。
     * Cmd+C などの修飾キー付きの入力は横取りしない
     * @param {Window} targetDialog - キー入力を受けるダイアログ
     * @param {Object} alignRadios - 整列ラジオ（none / start / center / end）
     * @param {Function} isHorizontalAlignActive - 水平の整列が有効なら true を返す関数
     * @param {Function} [onChange] - 選択を変えたあとに呼ぶ関数
     * @returns {void}
     */
    function addAlignmentKeyHandler(targetDialog, alignRadios, isHorizontalAlignActive, onChange) {
        targetDialog.addEventListener("keydown", function (event) {
            // Cmd+C などの修飾キー付きの入力は横取りしない（コピー等を潰さないため）
            // Do not swallow modified keystrokes such as Cmd+C
            var keyboardState = ScriptUI.environment.keyboardState;
            if (keyboardState.metaKey || keyboardState.ctrlKey || keyboardState.altKey || keyboardState.shiftKey) return;

            var targetRadio = null;
            var keyName = event.keyName;
            if (isHorizontalAlignActive()) {
                if (keyName === "L") targetRadio = alignRadios.start;
                else if (keyName === "C") targetRadio = alignRadios.center;
                else if (keyName === "R") targetRadio = alignRadios.end;
            } else {
                if (keyName === "T") targetRadio = alignRadios.start;
                else if (keyName === "M") targetRadio = alignRadios.center;
                else if (keyName === "B") targetRadio = alignRadios.end;
            }
            if (!targetRadio) return;
            event.preventDefault();
            if (targetRadio.value) return;     // 既に選択済みなら何もしない / already selected
            targetRadio.value = true;          // 同じ行内のラジオは自動排他 / radios in a row auto-exclude
            if (typeof onChange === "function") onChange();
        });
    }

    // =========================================
    // ダイアログの部品 / Dialog parts
    // =========================================

    /**
     * モードパネル（グループ／自動ペア認識／アートボード）を生成する。イベント結線は呼び出し側で行う
     * @param {Window} parentGroup - 追加先
     * @returns {{modeRadios: RadioButton[], getMode: Function, setMode: Function}} モードのラジオと読み書きの関数
     */
    function buildModePanel(parentGroup) {
        var modePanel = parentGroup.add("panel", undefined, getLabel("panel.mode"));
        setupPanel(modePanel, 6);
        modePanel.alignChildren = ["left", "top"]; // ラジオ縦並び / radios stacked
        var modeGroupRadio = modePanel.add("radiobutton", undefined, getLabel("radio.modeGroup"));
        var modeAutoPairRadio = modePanel.add("radiobutton", undefined, getLabel("radio.modeAutoPair"));
        var modeArtboardRadio = modePanel.add("radiobutton", undefined, getLabel("radio.modeArtboard"));
        modeGroupRadio.helpTip = getLabel("tooltip.modeGroup");
        modeAutoPairRadio.helpTip = getLabel("tooltip.modeAutoPair");
        modeArtboardRadio.helpTip = getLabel("tooltip.modeArtboard");

        /**
         * 現在のモードを返す
         * @returns {string} "group" / "artboard" / "auto"
         */
        function getMode() {
            if (modeGroupRadio.value) return "group";
            if (modeArtboardRadio.value) return "artboard";
            return "auto";
        }

        /**
         * モードを選ぶ。"group" / "artboard" 以外は自動ペア認識
         * @param {string} mode - モード
         * @returns {void}
         */
        function setMode(mode) {
            if (mode === "group") modeGroupRadio.value = true;
            else if (mode === "artboard") modeArtboardRadio.value = true;
            else modeAutoPairRadio.value = true;
        }

        return {
            modeRadios: [modeGroupRadio, modeAutoPairRadio, modeArtboardRadio],
            getMode: getMode,
            setMode: setMode
        };
    }

    /**
     * 固定オブジェクト（上 / 左 / 右 / 下）パネルを生成する。
     * 上下左右を厳密な3×3グリッドに配置する（中央・四隅は空セル）。セルを固定幅にして列を確実にそろえる。
     * 4つのラジオは親が別なので自動グループ化されない（排他は手動）。
     * 既定は「右」を固定（左が動く＝水平）。イベント結線は呼び出し側で行う
     * @param {Group} parentGroup - 追加先
     * @returns {{panel: Panel, fixedRadios: RadioButton[], selectFixedRadio: Function, getFixedSide: Function, setFixedSide: Function}} パネルと操作用の関数
     */
    function buildFixedSidePanel(parentGroup) {
        var fixedSidePanel = parentGroup.add("panel", undefined, getLabel("panel.fixedSide"));
        setupPanel(fixedSidePanel, 2);
        // 十字レイアウトの左右に余白を足す（+8）/ Add left/right margin to the cross layout (+8)
        fixedSidePanel.margins = [PANEL_MARGINS[0] + 8, PANEL_MARGINS[1], PANEL_MARGINS[2] + 8, PANEL_MARGINS[3]];
        fixedSidePanel.alignChildren = ["center", "top"]; // 各行を中央そろえで十字に / Center each row

        /**
         * グリッドの1セルを作る。withRadio が真ならラジオを入れて返す
         * @param {Group} gridRow - 追加先の行
         * @param {boolean} withRadio - ラジオを入れるか
         * @returns {RadioButton|null} 入れたラジオ（無ければ null）
         */
        function addGridCell(gridRow, withRadio) {
            var gridCell = gridRow.add("group");
            gridCell.margins = 0;
            gridCell.alignChildren = ["center", "center"];
            gridCell.preferredSize = GRID_CELL_SIZE;
            return withRadio ? gridCell.add("radiobutton", undefined, "") : null;
        }

        /**
         * グリッドの1行（3セル分の入れ物）を作る
         * @returns {Group} 行のグループ
         */
        function addGridRow() {
            var gridRow = fixedSidePanel.add("group");
            gridRow.orientation = "row";
            gridRow.alignment = "center";
            gridRow.spacing = 0;
            gridRow.margins = 0;
            return gridRow;
        }

        //   ・  上  ・
        //   左  ・  右
        //   ・  下  ・
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
        var radioBySide = { top: topRadio, left: leftRadio, right: rightRadio, bottom: bottomRadio };

        /**
         * 指定したラジオだけ ON にして排他制御する
         * @param {RadioButton} chosenRadio - ON にするラジオ
         * @returns {void}
         */
        function selectFixedRadio(chosenRadio) {
            for (var i = 0; i < fixedRadios.length; i++) {
                fixedRadios[i].value = (fixedRadios[i] === chosenRadio);
            }
        }
        for (var i = 0; i < fixedRadios.length; i++) {
            fixedRadios[i].helpTip = getLabel("tooltip.fixedSide");
        }
        rightRadio.value = true; // 既定：右を固定（左が動く＝水平）/ Default: fix right (horizontal)

        /**
         * 現在固定する側を返す
         * @returns {string} "top" / "left" / "right" / "bottom"
         */
        function getFixedSide() {
            if (topRadio.value) return "top";
            if (leftRadio.value) return "left";
            if (bottomRadio.value) return "bottom";
            return "right";
        }

        /**
         * 固定する側を設定する。未知の値は無視
         * @param {string} side - "top" / "left" / "right" / "bottom"
         * @returns {void}
         */
        function setFixedSide(side) {
            if (radioBySide[side]) selectFixedRadio(radioBySide[side]);
        }

        return {
            panel: fixedSidePanel,
            fixedRadios: fixedRadios,
            selectFixedRadio: selectFixedRadio,
            getFixedSide: getFixedSide,
            setFixedSide: setFixedSide
        };
    }

    /**
     * オフセットパネル（間隔の入力・プレビュー境界）を生成する。整列後のずらし量は［位置調整］パネルが
     * 持つので、ここは間隔だけを扱う。間隔は定規の単位で表示し（単位はパネル名に出す）、
     * getSpacingInPoints() が pt に換算して返す。イベント結線は呼び出し側で行う
     * @param {Group} parentGroup - 追加先
     * @param {number} initialGapPoints - 間隔の初期値・空欄時のフォールバック（pt）
     * @returns {{panel: Panel, spacingInput: EditText, previewBoundsCheckbox: Checkbox, getSpacingInPoints: Function, getBoundsType: Function}} パネルと操作用の関数
     */
    function buildGapPanel(parentGroup, initialGapPoints) {
        var gapPanel = parentGroup.add("panel", undefined, labelWithUnit("panel.offset", rulerUnitLabel));
        setupPanel(gapPanel, 6);

        // 間隔行（ラベル＋入力）。単位はパネル名に出しているので行には並べない
        // Gap row (label + input); the unit lives in the panel title instead
        var spacingRow = gapPanel.add("group");
        setupGroup(spacingRow, "row");
        spacingRow.alignment = "left"; // 広げず左寄せ / Keep at natural width, packed left
        spacingRow.add("statictext", undefined, labelText("fieldLabel.spacing"));
        var spacingInput = spacingRow.add("edittext", undefined, pointsToDisplayText(initialGapPoints));
        spacingInput.characters = 4;
        spacingInput.helpTip = getLabel("tooltip.spacing");

        // チェックボックス：プレビュー境界（左添え）/ Preview-bounds checkbox (left)
        var previewBoundsGroup = gapPanel.add("group");
        previewBoundsGroup.orientation = "row";
        previewBoundsGroup.alignment = "left";
        previewBoundsGroup.margins = [0, 5, 0, 0]; // 上マージン5 / Top margin 5
        var previewBoundsCheckbox = previewBoundsGroup.add("checkbox", undefined, getLabel("checkbox.previewBounds"));
        previewBoundsCheckbox.value = false; // OFF=幾何境界 / ON=プレビュー境界
        previewBoundsCheckbox.helpTip = getLabel("tooltip.previewBounds");

        /**
         * 入力値を pt に換算して返す。数値でなければ初期値
         * @returns {number} 間隔（pt）
         */
        function getSpacingInPoints() {
            var value = parseFloat(spacingInput.text);
            if (isNaN(value)) { value = initialGapPoints / pointsPerUnit; }
            return value * pointsPerUnit;
        }

        /**
         * チェックに応じた境界のプロパティ名を返す
         * @returns {string} "visibleBounds" または "geometricBounds"
         */
        function getBoundsType() {
            return previewBoundsCheckbox.value ? "visibleBounds" : "geometricBounds";
        }

        return {
            panel: gapPanel,
            spacingInput: spacingInput,
            previewBoundsCheckbox: previewBoundsCheckbox,
            getSpacingInPoints: getSpacingInPoints,
            getBoundsType: getBoundsType
        };
    }

    /**
     * 位置調整パネルを生成する。整列（なし/開始/中央/終端）と位置の2行を1枚にまとめ、［固定］で選んだ側に
     * 応じて setOrientation() で水平／垂直に切り替える（開始/終端のラベルとツールチップが入れ替わる）。
     * どちらの向きかは 左・右／上・下 のラベルで示すので、パネル名は「位置調整」で固定
     * （オフセットパネルと同じく、単位はパネル名に出す）。
     * 向きごとの値の保持と中央時のオフセット無効化は createOrientationSwitcher() が行う。
     * 開始/終端のラベルはレイアウト確定後に差し替えるので、文字数の多いほうで組み立てて幅を確保しておく
     * @param {Window} parentGroup - 追加先
     * @returns {{alignRadios: Object, offsetRow: Group, offsetInput: EditText, setOrientation: Function}} 整列ラジオ・位置の入力欄と切り替え関数
     */
    function buildAlignmentPanel(parentGroup) {
        /**
         * 文字数の多いほうを返す（ラベル領域の確保用）
         * @param {string} firstText - 候補1
         * @param {string} secondText - 候補2
         * @returns {string} 長いほうの文字列
         */
        function longerText(firstText, secondText) {
            return (firstText.length >= secondText.length) ? firstText : secondText;
        }

        var alignmentPanel = parentGroup.add("panel", undefined, labelWithUnit("panel.position", rulerUnitLabel));
        setupPanel(alignmentPanel, 6);

        // 整列行（ラベル＋なし/開始/中央/終端）/ Alignment row (label + none/start/center/end)
        var alignRow = alignmentPanel.add("group");
        setupGroup(alignRow, "row");
        alignRow.alignment = "left";
        var alignLabel = alignRow.add("statictext", undefined, labelText("fieldLabel.align"));
        var alignRadios = {
            none: alignRow.add("radiobutton", undefined, getLabel("radio.alignNone")),
            start: alignRow.add("radiobutton", undefined,
                longerText(getLabel("radio.alignLeft"), getLabel("radio.alignTop"))),
            center: alignRow.add("radiobutton", undefined, getLabel("radio.alignCenter")),
            end: alignRow.add("radiobutton", undefined,
                longerText(getLabel("radio.alignRight"), getLabel("radio.alignBottom")))
        };
        alignRadios.none.value = true; // 既定：整列なし / Default: no alignment

        // 位置行（ラベル＋入力）。単位はパネル名に出しているので行には並べない
        // Position row (label + input); the unit lives in the panel title instead
        var offsetRow = alignmentPanel.add("group");
        setupGroup(offsetRow, "row");
        offsetRow.alignment = "left";
        var offsetLabel = offsetRow.add("statictext", undefined, labelText("fieldLabel.position"));
        var offsetInput = offsetRow.add("edittext", undefined, "0");
        offsetInput.characters = 4;

        // ラベル幅をそろえて整列ラジオと入力の開始位置を合わせる / Match label widths
        var labelWidth = Math.max(alignLabel.preferredSize.width, offsetLabel.preferredSize.width);
        alignLabel.preferredSize.width = labelWidth;
        offsetLabel.preferredSize.width = labelWidth;

        /**
         * 水平／垂直を切り替える（開始/終端のラベルとツールチップ）。ラベルの差し替えで
         * レイアウトは組み直さないので、レイアウト確定後（onShow 以降）に呼ぶこと
         * @param {boolean} isHorizontal - 水平の整列にするか
         * @returns {void}
         */
        function setOrientation(isHorizontal) {
            alignRadios.start.text = getLabel(isHorizontal ? "radio.alignLeft" : "radio.alignTop");
            alignRadios.end.text = getLabel(isHorizontal ? "radio.alignRight" : "radio.alignBottom");
            var alignTip = getLabel(isHorizontal ? "tooltip.alignH" : "tooltip.alignV");
            for (var alignKey in alignRadios) { alignRadios[alignKey].helpTip = alignTip; }
            offsetInput.helpTip = getLabel(isHorizontal ? "tooltip.offsetHorizontal" : "tooltip.offsetVertical");
        }

        return {
            alignRadios: alignRadios,
            offsetRow: offsetRow,
            offsetInput: offsetInput,
            setOrientation: setOrientation
        };
    }

    /**
     * 整列ラジオで選択中の値を返す
     * @param {Object} alignRadios - 整列ラジオ（none / start / center / end）
     * @returns {string} "none" / "start" / "center" / "end"
     */
    function getAlignValue(alignRadios) {
        if (alignRadios.start.value) return "start";
        if (alignRadios.center.value) return "center";
        if (alignRadios.end.value) return "end";
        return "none";
    }

    /**
     * 保存済みの整列値をラジオへ反映する。未知の値は無視
     * @param {Object} alignRadios - 整列ラジオ（none / start / center / end）
     * @param {string} alignValue - 整列値
     * @returns {void}
     */
    function applySavedAlign(alignRadios, alignValue) {
        if (!alignRadios || !alignValue) return;
        var targetRadio = alignRadios[alignValue];
        if (targetRadio) targetRadio.value = true;
    }

    /**
     * ［位置調整］パネルの向き（水平／垂直）の切り替えを受け持つ。水平／垂直それぞれの入力内容を控え、
     * ［固定］の側が変わったら退避・復元する。整列「中央」のときはオフセットを 0 にして無効にする。
     * 上下キー（縦並び）→ 水平の整列、左右キー（横並び）→ 垂直の整列
     * @param {Object} fixedSideRefs - buildFixedSidePanel() の戻り値
     * @param {Object} alignmentRefs - buildAlignmentPanel() の戻り値
     * @returns {{orientationValues: Object, isVerticalGap: Function, stashOrientationValues: Function, updateActivePanels: Function}} 向きごとの控えと操作用の関数
     */
    function createOrientationSwitcher(fixedSideRefs, alignmentRefs) {
        var alignRadios = alignmentRefs.alignRadios;
        var offsetInput = alignmentRefs.offsetInput;
        // 水平／垂直それぞれの入力内容。パネルを切り替えるときに退避・復元する
        // Per-orientation values, stashed and restored as the panel switches
        var orientationValues = {
            h: { align: "none", offset: "0" },
            v: { align: "none", offset: "0" }
        };
        var shownOrientation = null; // 今パネルに出ている向き / the orientation currently shown

        /**
         * キーオブジェクトの側からギャップが垂直か（上下キー）を判定する
         * @returns {boolean} 上下をキーにしているなら true
         */
        function isVerticalGap() {
            var side = fixedSideRefs.getFixedSide();
            return side === "top" || side === "bottom";
        }

        /**
         * パネルに出ている内容を、その向きの控えへ退避する
         * @returns {void}
         */
        function stashOrientationValues() {
            if (!shownOrientation) return;
            orientationValues[shownOrientation].align = getAlignValue(alignRadios);
            orientationValues[shownOrientation].offset = offsetInput.text;
        }

        /**
         * キー側に合わせて整列パネルの向きと中身を入れ替え、整列「中央」ならオフセットを 0＋無効にする
         * @returns {void}
         */
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
            alignmentRefs.offsetRow.enabled = !isCenter;
            if (isCenter) offsetInput.text = "0";
        }

        return {
            orientationValues: orientationValues,
            isVerticalGap: isVerticalGap,
            stashOrientationValues: stashOrientationValues,
            updateActivePanels: updateActivePanels
        };
    }

    // =========================================
    // 行揃えボタン / Justification buttons
    // ScriptUI の button では選択状態を表示できないので、背景とアイコンを onDraw で自前描画する。
    // 描画方式は UnifiedTypePalette.jsx にそろえる。
    // ScriptUI buttons cannot show a selected state, so the background and icon are drawn in onDraw;
    // the drawing follows UnifiedTypePalette.jsx.
    // =========================================

    /**
     * 環境設定のUI明るさが明るい側かを返す。
     * uiBrightness は 0（最暗）〜1（最明）。0.5（やや暗め）は暗い側に含めるため 0.5 超で判定
     * @returns {boolean} 明るい側なら true
     */
    function isLightUI() {
        try {
            return app.preferences.getRealPreference("uiBrightness") > 0.5;
        } catch (e) {
            return false;
        }
    }

    /**
     * テーマとアクティブ状態に応じたボタンの配色を返す
     * @param {boolean} isLight - 明るいテーマか
     * @param {boolean} isActive - 選択中か
     * @returns {{bg: number[], border: (number[]|null), line: number[]}} 背景・枠・線の色
     */
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

    /**
     * アイコンの行ごとの線幅を返す。均等配置だけ最終行以外を長くする
     * @param {string} iconType - "left" / "center" / "right" / "full"
     * @param {number} longWidth - 長い行の幅
     * @param {number} shortWidth - 短い行の幅
     * @returns {number[]} 4行分の線幅
     */
    function getJustifyLineWidths(iconType, longWidth, shortWidth) {
        if (iconType === "full") return [longWidth, longWidth, longWidth, shortWidth];
        return [longWidth, shortWidth, longWidth, shortWidth];
    }

    /**
     * 行の開始 X（左／中央／右）を返す
     * @param {string} iconType - "left" / "center" / "right" / "full"
     * @param {number} buttonWidth - ボタンの幅
     * @param {number} lineWidth - 線の幅
     * @returns {number} 線の開始 X
     */
    function getJustifyLineX(iconType, buttonWidth, lineWidth) {
        var margin = 5;
        if (iconType === "right") return buttonWidth - margin - lineWidth;
        if (iconType === "center") return Math.round((buttonWidth - lineWidth) / 2);
        return margin;
    }

    /**
     * 行揃えアイコンの罫線を描く
     * @param {ScriptUIGraphics} graphics - 描画先
     * @param {string} iconType - "left" / "center" / "right" / "full"
     * @param {number} buttonWidth - ボタンの幅
     * @param {number[]} lineColor - 線の色
     * @returns {void}
     */
    function drawJustifyIconLines(graphics, iconType, buttonWidth, lineColor) {
        var linePen = graphics.newPen(graphics.PenType.SOLID_COLOR, lineColor, 1.2);
        var lineYs = [7, 11, 15, 19];
        var lineWidths = getJustifyLineWidths(iconType, 15, 10);
        for (var i = 0; i < lineYs.length; i++) {
            var lineWidth = lineWidths[i];
            var lineStartX = getJustifyLineX(iconType, buttonWidth, lineWidth);
            graphics.newPath();
            graphics.moveTo(lineStartX, lineYs[i]);
            graphics.lineTo(lineStartX + lineWidth, lineYs[i]);
            graphics.strokePath(linePen);
        }
    }

    /**
     * 「自動」のアイコンを描く。他の行揃えと同じ高さに収まる「A」を線で描く
     * （drawString はフォントの解決に環境差があり空欄になることがあるため使わない）
     * @param {ScriptUIGraphics} graphics - 描画先
     * @param {number} buttonWidth - ボタンの幅
     * @param {number[]} lineColor - 線の色
     * @returns {void}
     */
    function drawAutoIcon(graphics, buttonWidth, lineColor) {
        var linePen = graphics.newPen(graphics.PenType.SOLID_COLOR, lineColor, 1.2);
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
        graphics.strokePath(linePen);
        graphics.newPath();
        graphics.moveTo(centerX - crossHalf, crossY);
        graphics.lineTo(centerX + crossHalf, crossY);
        graphics.strokePath(linePen);
    }

    /**
     * ボタンの背景とアイコンを描く。描画できない環境では OS 標準のボタン（ラベル付き）に任せる
     * @param {Button} justifyButton - 描くボタン（iconType を持つ）
     * @param {boolean} isActive - 選択中か
     * @param {boolean} isLight - 明るいテーマか
     * @returns {void}
     */
    function drawJustifyButton(justifyButton, isActive, isLight) {
        var graphics = justifyButton.graphics;
        var buttonColors = getJustifyColors(isLight, isActive);
        try {
            graphics.rectPath(0, 0, justifyButton.size[0], justifyButton.size[1]);
            graphics.fillPath(graphics.newBrush(graphics.BrushType.SOLID_COLOR, buttonColors.bg));
            if (buttonColors.border) {
                graphics.rectPath(0, 0, justifyButton.size[0], justifyButton.size[1]);
                graphics.strokePath(graphics.newPen(graphics.PenType.SOLID_COLOR, buttonColors.border, 1));
            }
            if (justifyButton.iconType === "auto") {
                drawAutoIcon(graphics, justifyButton.size[0], buttonColors.line);
            } else {
                drawJustifyIconLines(graphics, justifyButton.iconType, justifyButton.size[0], buttonColors.line);
            }
        } catch (eDraw) {
            // 描画できない環境ではOS標準のボタン（ラベル付き）に任せる / Fall back to the OS control
            try { graphics.drawOSControl(); } catch (eOs) {}
        }
    }

    /**
     * テキストの行揃えパネルを生成する。自動 / 左 / 中央 / 右 / 均等配置（最終行左）をボタンで並べる。
     * 「自動」は整列・キーに連動（エリア内文字は均等配置）。解決は resolveJustification() で行う。
     * イベント結線は呼び出し側で行う
     * @param {Window} parentGroup - 追加先
     * @returns {{panel: Panel, buttons: Button[], getJustifyMode: Function, setJustifyMode: Function}} パネルとボタン、読み書きの関数
     */
    function buildJustifyPanel(parentGroup) {
        var justifyPanel = parentGroup.add("panel", undefined, getLabel("panel.justify"));
        setupPanel(justifyPanel, 4);
        justifyPanel.orientation = "row"; // ボタン横並び / buttons in a row
        justifyPanel.alignChildren = ["center", "center"]; // ボタン列をパネルの左右中央に / Center the button row in the panel

        // アクティブな行揃えと UI 明暗を共有する（onDraw のクロージャから参照）
        // Shared active id + theme, read by the onDraw closures
        var justifyState = { activeId: "auto", isLight: isLightUI() }; // 既定：自動（整列・キーに連動）/ Default: auto
        var justifyOptions = [
            { id: "auto", label: "button.justifyAuto", tip: "tooltip.justifyAuto" },
            { id: "left", label: "button.justifyLeft", tip: "tooltip.justifyLeft" },
            { id: "center", label: "button.justifyCenter", tip: "tooltip.justifyCenter" },
            { id: "right", label: "button.justifyRight", tip: "tooltip.justifyRight" },
            { id: "full", label: "button.justifyFull", tip: "tooltip.justifyFull" }
        ];
        var justifyButtons = [];
        for (var i = 0; i < justifyOptions.length; i++) {
            // ラベルは描画に失敗したときのフォールバック（drawOSControl）でも使う / text is also the fallback label
            var justifyButton = justifyPanel.add("button", undefined, getLabel(justifyOptions[i].label));
            justifyButton.helpTip = getLabel(justifyOptions[i].tip);
            justifyButton.preferredSize = JUSTIFY_BUTTON_SIZE;
            justifyButton.minimumSize = JUSTIFY_BUTTON_SIZE;
            justifyButton.maximumSize = JUSTIFY_BUTTON_SIZE; // 伸ばさない / keep the fixed size
            justifyButton.justifyId = justifyOptions[i].id;
            justifyButton.iconType = justifyOptions[i].id;
            justifyButton.onDraw = function () { drawJustifyButton(this, this.justifyId === justifyState.activeId, justifyState.isLight); };
            justifyButtons.push(justifyButton);
        }

        /**
         * 選択中の行揃えモードを返す
         * @returns {string} "auto" / "left" / "center" / "right" / "full"
         */
        function getJustifyMode() {
            return justifyState.activeId;
        }

        /**
         * 行揃えモードを設定してボタンを描き直す。未知の値は無視
         * @param {string} mode - "auto" / "left" / "center" / "right" / "full"
         * @returns {void}
         */
        function setJustifyMode(mode) {
            var found = false;
            for (var i = 0; i < justifyButtons.length; i++) {
                if (justifyButtons[i].justifyId === mode) { found = true; break; }
            }
            if (!found) return;
            justifyState.activeId = mode;
            for (var j = 0; j < justifyButtons.length; j++) {
                try { justifyButtons[j].notify("onDraw"); } catch (eDraw) {}
            }
            // notify だけでは画面に反映されないことがあるので、ウィンドウの再描画も要求する
            // notify alone may not reach the screen, so ask the window to repaint as well
            try { justifyPanel.window.update(); } catch (eUpdate) {}
        }

        return { panel: justifyPanel, buttons: justifyButtons, getJustifyMode: getJustifyMode, setJustifyMode: setJustifyMode };
    }

    /**
     * 設定ダイアログを組み立てる（イベント結線と値の復元は呼び出し側で行う）
     * @param {number} initialGapPoints - 間隔の初期値（pt）
     * @returns {{settingsDialog: Window, modeRefs: Object, fixedSideRefs: Object, gapRefs: Object, alignmentRefs: Object, justifyRefs: Object, btnCancel: Button, btnOK: Button}} ダイアログと各パネルの参照
     */
    function buildSettingsDialog(initialGapPoints) {
        var settingsDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        settingsDialog.orientation = "column";
        settingsDialog.alignChildren = "fill";

        // モード（1カラム・ラジオ縦並び）/ Mode (single column, radios stacked)
        var modeRefs = buildModePanel(settingsDialog);

        // キーオブジェクト と オフセット を2カラムで左右に並べる / Key object + Offset side by side (two columns)
        var keyPositionColumns = settingsDialog.add("group");
        keyPositionColumns.orientation = "row";
        keyPositionColumns.alignChildren = ["fill", "fill"]; // 2パネルの高さをそろえる / Match panel heights
        keyPositionColumns.spacing = PANEL_SPACING;

        // キーオブジェクト（上・左・右・下を十字に配置）/ Key object (arranged as a cross)
        var fixedSideRefs = buildFixedSidePanel(keyPositionColumns);
        // オフセット（間隔・プレビュー境界）/ Offset (gap, preview bounds)
        var gapRefs = buildGapPanel(keyPositionColumns, initialGapPoints);

        // キー／位置の2パネルの高さをそろえる / Match the two panel heights
        fixedSideRefs.panel.alignment = ["fill", "fill"];
        gapRefs.panel.alignment = ["fill", "fill"];

        // 位置調整パネル（整列＋位置）。1枚で、［固定］に応じて水平／垂直に切り替わる
        // One Position panel (alignment + offset); it switches with the Key Object side
        var alignmentRefs = buildAlignmentPanel(settingsDialog);

        // テキストの行揃え（自動 / 左 / 中央 / 右 / 均等配置）/ Text alignment
        var justifyRefs = buildJustifyPanel(settingsDialog);

        // ボタン（Mac 規約：Cancel → OK）/ Buttons (Mac order: Cancel → OK)
        var btnRowGroup = settingsDialog.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.alignment = ["center", "bottom"]; // ボタンをダイアログの左右中央に / Center the buttons in the dialog
        btnRowGroup.alignChildren = ["center", "center"];
        var btnCancel = btnRowGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnOK = btnRowGroup.add("button", undefined, "OK", { name: "ok" });
        // 行揃えのボタンが増えたので Enter / ESC の行き先を明示する
        // Spell out where Enter / ESC go, now that the justification buttons are pushbuttons too
        settingsDialog.defaultElement = btnOK;
        settingsDialog.cancelElement = btnCancel;

        return {
            settingsDialog: settingsDialog,
            modeRefs: modeRefs,
            fixedSideRefs: fixedSideRefs,
            gapRefs: gapRefs,
            alignmentRefs: alignmentRefs,
            justifyRefs: justifyRefs,
            btnCancel: btnCancel,
            btnOK: btnOK
        };
    }

    // =========================================
    // 軸 / Axis
    // =========================================

    /**
     * 水平 / 垂直を共通の「進行方向の座標 p」に抽象化する。p は増えるほど右（水平）または下（垂直）。
     * これで間隔調整のロジックを 1 本で両軸に使える。geometricBounds = [左, 上, 右, 下]。
     * start = 進行方向の先頭側の辺、end = 後ろ側の辺。
     * 垂直は Illustrator の y が上ほど大きいので符号を反転して「下方向で増加」にそろえる
     * @param {boolean} isVertical - 垂直の軸にするか
     * @returns {{vertical: boolean, start: Function, end: Function}} 軸（境界から先頭辺・後ろ辺を取り出す関数）
     */
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

    /**
     * 固定する側から軸を求める（上下→垂直、左右→水平）
     * @param {string} fixedSide - "top" / "left" / "right" / "bottom"
     * @returns {Object} makeAxis() の軸
     */
    function axisForSide(fixedSide) {
        return makeAxis(fixedSide === "top" || fixedSide === "bottom");
    }

    /**
     * 固定側が進行方向の「後ろ側」(右 or 下) かどうかを返す
     * @param {string} fixedSide - "top" / "left" / "right" / "bottom"
     * @returns {boolean} 右・下なら true
     */
    function isAnchorEnd(fixedSide) {
        return fixedSide === "right" || fixedSide === "bottom";
    }

    // =========================================
    // 行揃え / Text justification
    // =========================================

    /**
     * 左右方向の整列値を段落の行揃えへ対応づける（start=左揃え、center=中央揃え、end=右揃え）
     * @param {string} alignMode - "none" / "start" / "center" / "end"
     * @returns {Justification|null} 行揃え。none・不明は null（行揃えを変更しない）
     */
    function justificationForAlign(alignMode) {
        if (alignMode === "start") return Justification.LEFT;
        if (alignMode === "center") return Justification.CENTER;
        if (alignMode === "end") return Justification.RIGHT;
        return null;
    }

    /**
     * 「テキストの行揃え」パネルの選択を、オブジェクトごとの行揃えへ解決する。
     * "left"/"center"/"right"/"full" は一律。"auto" は連動：エリア内文字＝均等配置、
     * ポイント文字＝縦並びなら水平整列に連動（左/中央/右）、横並びならキー側に連動（左→左/右→右）
     * @param {PageItem} targetItem - 対象のオブジェクト
     * @param {Object} spacingContext - applySpacing() の設定（justifyMode / alignAxis / alignMode / fixedSide）
     * @returns {Justification|null} 行揃え。null は変更しない
     */
    function resolveJustification(targetItem, spacingContext) {
        var justifyMode = spacingContext.justifyMode;
        if (justifyMode === "left") return Justification.LEFT;
        if (justifyMode === "center") return Justification.CENTER;
        if (justifyMode === "right") return Justification.RIGHT;
        if (justifyMode === "full") return Justification.FULLJUSTIFYLASTLINELEFT;
        // auto（連動）/ auto (linked)
        if (targetItem.constructor.name === "TextFrame" && targetItem.kind === TextType.AREATEXT) {
            return Justification.FULLJUSTIFYLASTLINELEFT; // エリア内文字は均等配置 / area text → justify
        }
        if (!spacingContext.alignAxis.vertical) return justificationForAlign(spacingContext.alignMode); // 縦並び：水平整列に連動
        var fixedSide = spacingContext.fixedSide;
        return (fixedSide === "left") ? Justification.LEFT
            : (fixedSide === "right") ? Justification.RIGHT : null;        // 横並び：キー側に連動
    }

    /**
     * テキストの段落行揃えを設定する。ポイント文字は行揃えを変えるとアンカー基準で組み直されて
     * フレームが動く。見た目の位置は間隔・整列の計算で決めるので、どの値でも元の位置へ戻す
     * （戻さないと行揃えを往復するたびに字幅の半分ずつずれ、キャンセルしても戻らない）。
     * Justification.LEFT は代入が無視される Illustrator のバグがあるので、一時 resize（200%→50%）で
     * 段落属性をリフレッシュしてから代入する
     * @param {TextFrame} textFrame - 対象のテキスト
     * @param {Justification} justification - 行揃え
     * @returns {void}
     */
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

    /**
     * 選択オブジェクトを最も近いもの同士でペアにする。
     * ペアリングは常に geometricBounds（中心）で判定する。「プレビュー境界」を ON にしても
     * 間隔計算の基準が変わるだけで、ペアの組み合わせ自体は変わらない
     * @param {PageItem[]} selectedItems - 選択オブジェクト
     * @returns {Object[]} ペア（{ a, b }）の配列
     */
    function createNearestPairs(selectedItems) {
        /**
         * オブジェクトの中心座標を返す（常に幾何境界。クリップグループはクリッピングパス基準）
         * @param {PageItem} item - 対象のオブジェクト
         * @returns {{x: number, y: number}} 中心座標
         */
        function getObjectCenter(item) {
            var bounds = geometricBoundsOf(item);
            return {
                x: (bounds[0] + bounds[2]) / 2,
                y: (bounds[1] + bounds[3]) / 2
            };
        }

        /**
         * 2点間の距離を返す
         * @param {{x: number, y: number}} point1 - 点1
         * @param {{x: number, y: number}} point2 - 点2
         * @returns {number} 距離
         */
        function getPointDistance(point1, point2) {
            var dx = point1.x - point2.x;
            var dy = point1.y - point2.y;
            return Math.sqrt(dx * dx + dy * dy);
        }

        // 中心座標付きの作業リスト / Working list with centers
        var remainingItems = [];
        for (var i = 0; i < selectedItems.length; i++) {
            remainingItems.push({
                item: selectedItems[i],
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
                pairs.push({ a: currentItem.item, b: partnerItem.item });
            }
        }
        return pairs;
    }

    /**
     * GroupItem の直下のオブジェクトを配列で返す
     * @param {PageItem} item - 対象のオブジェクト
     * @returns {PageItem[]|null} 直下のオブジェクト。グループ以外は null
     */
    function getGroupChildren(item) {
        if (item.typename !== "GroupItem") return null;
        var groupChildren = [];
        for (var i = 0; i < item.pageItems.length; i++) {
            groupChildren.push(item.pageItems[i]);
        }
        return groupChildren;
    }

    /**
     * クリップグループならクリッピングパス（マスク）を返す。
     * クリップグループの geometricBounds／visibleBounds はマスクで隠れた部分まで含むことがあり、
     * 間隔計算がずれる。基準をマスクの形（クリッピングパス）にそろえるために使う
     * @param {PageItem} item - 対象のオブジェクト
     * @returns {PageItem|null} クリッピングパス。クリップグループでなければ null
     */
    function getClippingPath(item) {
        if (item.typename !== "GroupItem" || !item.clipped) return null;
        var groupItems = item.pageItems;
        for (var i = 0; i < groupItems.length; i++) {
            var childItem = groupItems[i];
            if (childItem.typename === "PathItem" && childItem.clipping) return childItem;
            // 複合パスのマスクは先頭サブパスの clipping で判定 / Compound-path mask: check the first subpath
            if (childItem.typename === "CompoundPathItem" &&
                childItem.pathItems.length > 0 && childItem.pathItems[0].clipping) return childItem;
        }
        return null;
    }

    /**
     * 間隔計算に使う幾何境界を返す。クリップグループはクリッピングパスを基準にする
     * @param {PageItem} item - 対象のオブジェクト
     * @returns {number[]} [左, 上, 右, 下]
     */
    function geometricBoundsOf(item) {
        var clippingPath = getClippingPath(item);
        return clippingPath ? clippingPath.geometricBounds : item.geometricBounds;
    }

    /**
     * 間隔計算に使うプレビュー境界を返す。クリップグループはクリッピングパスを基準にする
     * @param {PageItem} item - 対象のオブジェクト
     * @returns {number[]} [左, 上, 右, 下]
     */
    function visibleBoundsOf(item) {
        var clippingPath = getClippingPath(item);
        return clippingPath ? clippingPath.visibleBounds : item.visibleBounds;
    }

    /**
     * 選択した各グループの中身を1単位にする。グループは2つに限らない。含まれる全オブジェクトを
     * members として持ち、applySpacing 側で先頭辺の順に等間隔へ分配する（固定側を基準に配置）
     * @param {PageItem[]} selectedItems - 選択オブジェクト
     * @returns {Object[]} 単位（{ members }）の配列。子が2個未満のグループとグループ以外は含まない
     */
    function createGroupPairs(selectedItems) {
        var pairs = [];
        for (var i = 0; i < selectedItems.length; i++) {
            var groupChildren = getGroupChildren(selectedItems[i]);
            // グループでない、または子が2個未満なら間隔を調整できない / Need a group with 2+ children
            if (!groupChildren || groupChildren.length < 2) continue;
            pairs.push({ members: groupChildren });
        }
        return pairs;
    }

    /**
     * オブジェクトの中心を含むアートボードの矩形を返す。該当が無ければアクティブアートボードを使う。
     * geometricBounds と同じ並び（上 > 下、y は上方向で増加）なので軸関数で扱える
     * @param {PageItem} item - 対象のオブジェクト
     * @returns {number[]} artboardRect [左, 上, 右, 下]
     */
    function artboardRectFor(item) {
        var artboards = app.activeDocument.artboards;
        var bounds = geometricBoundsOf(item);
        var centerX = (bounds[0] + bounds[2]) / 2;
        var centerY = (bounds[1] + bounds[3]) / 2;
        for (var i = 0; i < artboards.length; i++) {
            var artboardRect = artboards[i].artboardRect; // [左, 上, 右, 下] / [left, top, right, bottom]
            if (centerX >= artboardRect[0] && centerX <= artboardRect[2] &&
                centerY <= artboardRect[1] && centerY >= artboardRect[3]) {
                return artboardRect;
            }
        }
        return artboards[artboards.getActiveArtboardIndex()].artboardRect;
    }

    /**
     * アートボードモードの作業単位を作る。選択オブジェクトを1つずつ独立した単位にし、
     * applySpacing 側で各オブジェクトとアートボード端の間隔（マージン）を指定値にそろえる
     * @param {PageItem[]} selectedItems - 選択オブジェクト
     * @returns {Object[]} 単位（{ single }）の配列
     */
    function createArtboardUnits(selectedItems) {
        var artboardUnits = [];
        for (var i = 0; i < selectedItems.length; i++) {
            artboardUnits.push({ single: selectedItems[i] });
        }
        return artboardUnits;
    }

    /**
     * 作業単位に含まれるオブジェクトを配列で返す（アートボード：1個、グループ：全メンバー、ペア：2個）
     * @param {Object} pair - 作業単位（{ single } / { members } / { a, b }）
     * @returns {PageItem[]} オブジェクトの配列
     */
    function getPairItems(pair) {
        if (pair.single) return [pair.single];
        if (pair.members) return pair.members;
        return [pair.a, pair.b];
    }

    /**
     * 選択オブジェクトの現在の間隔の平均（pt）を求める。
     * モードに合わせて測る：グループは各グループ内の隣接間隔、アートボードは固定側の端との距離
     * （マージン）、自動ペア認識は各ペアの間隔。測る軸と固定端は fixedSide から決める。
     * 常に geometricBounds（幾何境界）基準
     * @param {PageItem[]} selectedItems - 選択オブジェクト
     * @param {string} mode - "group" / "artboard" / "auto"
     * @param {string} fixedSide - "top" / "left" / "right" / "bottom"
     * @returns {number|null} 平均間隔（pt）。測れない場合は null
     */
    function computeAverageGap(selectedItems, mode, fixedSide) {
        var axis = axisForSide(fixedSide);
        var gaps = [];

        if (mode === "group") {
            // 各グループ内：進行方向順に並べて隣り合う間隔を測る / Adjacent gaps inside each group
            for (var i = 0; i < selectedItems.length; i++) {
                var groupChildren = getGroupChildren(selectedItems[i]);
                if (!groupChildren || groupChildren.length < 2) continue;
                var sortedChildren = groupChildren.slice(0);
                sortedChildren.sort(function (itemA, itemB) {
                    return axis.start(geometricBoundsOf(itemA)) - axis.start(geometricBoundsOf(itemB));
                });
                for (var j = 1; j < sortedChildren.length; j++) {
                    // 次の先頭辺 - 前の後ろ辺 / next leading edge - previous trailing edge
                    gaps.push(axis.start(geometricBoundsOf(sortedChildren[j])) - axis.end(geometricBoundsOf(sortedChildren[j - 1])));
                }
            }
        } else if (mode === "artboard") {
            // アートボード：各オブジェクトと固定側のアートボード端との距離（マージン）を測る
            // Artboard: distance (margin) from each object to the fixed artboard edge
            var anchorEnd = isAnchorEnd(fixedSide);
            for (var i = 0; i < selectedItems.length; i++) {
                var itemBounds = geometricBoundsOf(selectedItems[i]);
                var artboardRect = artboardRectFor(selectedItems[i]);
                gaps.push(anchorEnd ? axis.end(artboardRect) - axis.end(itemBounds)
                    : axis.start(itemBounds) - axis.start(artboardRect));
            }
        } else {
            // 自動ペア認識：各ペアの間隔を測る / Gap of each nearest pair
            var nearestPairs = createNearestPairs(selectedItems);
            for (var i = 0; i < nearestPairs.length; i++) {
                var boundsA = geometricBoundsOf(nearestPairs[i].a);
                var boundsB = geometricBoundsOf(nearestPairs[i].b);
                var isALeading = axis.start(boundsA) < axis.start(boundsB);
                var leadBounds = isALeading ? boundsA : boundsB;
                var trailBounds = isALeading ? boundsB : boundsA;
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

    // ファイル（セッションまたぎ）には保存しないキー。モード・間隔は選択内容に依存するので
    // 同一セッション内（常駐グローバル）だけ覚え、再起動後は毎回選択から決め直す。
    // Keys NOT written to the file (cross-session): mode and gap depend on the selection, so they are
    // remembered only within the session (persistent global) and re-derived from the selection after a restart.
    var SETTINGS_FILE_SKIP = { mode: true, gap: true };

    /**
     * 設定オブジェクトを浅くコピーする（常駐グローバルを直接書き換えないため）
     * @param {Object} sourceSettings - コピー元
     * @returns {Object} コピー
     */
    function cloneSettings(sourceSettings) {
        var settingsCopy = {};
        for (var key in sourceSettings) {
            if (sourceSettings.hasOwnProperty(key)) settingsCopy[key] = sourceSettings[key];
        }
        return settingsCopy;
    }

    /**
     * 設定を読み込む。まず #targetengine の常駐グローバル（同一セッションの最新）を優先し、
     * 無ければ設定ファイル（セッションまたぎ）を読む
     * @returns {Object} 設定（どちらも無ければ空オブジェクト）
     */
    function loadSettings() {
        if ($.global[SETTINGS_GLOBAL_KEY]) {
            return cloneSettings($.global[SETTINGS_GLOBAL_KEY]);
        }
        var settings = {};
        var settingsFile = new File(SETTINGS_FILE);
        if (!settingsFile.exists || !settingsFile.open("r")) return settings;
        var fileContent = settingsFile.read();
        settingsFile.close();
        var settingLines = fileContent.split("\n");
        for (var i = 0; i < settingLines.length; i++) {
            var eqIndex = settingLines[i].indexOf("=");
            if (eqIndex <= 0) continue;
            var key = settingLines[i].substring(0, eqIndex).replace(/^\s+|\s+$/g, "");
            var value = settingLines[i].substring(eqIndex + 1).replace(/^\s+|\s+$/g, "");
            settings[key] = value;
        }
        return settings;
    }

    /**
     * 設定を保存する。常駐グローバル（同一セッション用）には全項目、ファイル（セッションまたぎ用）には
     * モード・間隔を除いた項目を書く
     * @param {Object} settings - 保存する設定
     * @returns {void}
     */
    function saveSettings(settings) {
        $.global[SETTINGS_GLOBAL_KEY] = cloneSettings(settings);
        var settingsFile = new File(SETTINGS_FILE);
        if (!settingsFile.open("w")) return;
        var settingLines = [];
        for (var key in settings) {
            if (settings.hasOwnProperty(key) && !SETTINGS_FILE_SKIP[key]) settingLines.push(key + "=" + settings[key]);
        }
        settingsFile.write(settingLines.join("\n"));
        settingsFile.close();
    }

    // =========================================
    // メイン処理 / Main
    // =========================================
    (function () {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }

        var doc = app.activeDocument;

        // doc.selection を配列にコピーしておく（後の選択変更や undo の影響を受けないように）
        // Copy doc.selection into a plain array (immune to later selection changes / undo)
        var liveSelection = doc.selection;
        // selection が無効・未選択のケースを先に弾く（Illustrator では稀に null になる）
        // Guard against an invalid / empty selection (selection can be null in rare cases)
        if (!liveSelection || liveSelection.length === 0) {
            alert(getLabel("alert.selectTwo"));
            return;
        }
        var selectedItems = [];
        for (var i = 0; i < liveSelection.length; i++) {
            selectedItems.push(liveSelection[i]);
        }

        if (selectedItems.length < 2) {
            alert(getLabel("alert.selectTwo"));
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
        var appliedMoves = []; // 適用済みの移動 / Applied moves: { targetItem, dx, dy }
        var appliedJustifications = []; // 適用済みの行揃え変更 / Applied justification changes: { targetItem, original }
        // 行揃えの変更で境界が変わったか。プレビューは毎回元位置へ巻き戻すので、境界が変わるのは
        // 行揃えを当てた／戻したときだけ。真のときだけ測り直す（全オブジェクトの visibleBounds
        // 再取得は重く、間隔欄の1打鍵ごとに走ると効いてくる）。
        // Whether justification changed the bounds. The preview always reverts to the original
        // positions, so only applying/reverting justification invalidates them; re-measure only then
        // (re-reading visibleBounds for every object on each keystroke is expensive).
        var boundsCacheStale = false;

        /**
         * 記録した移動を逆向きに適用して元に戻す（位置のみ）。
         * 行揃えはリフレッシュ毎に巻き戻すと resize が連発して重いので、ここでは触らない。
         * 行揃えの巻き戻しは整列系ラジオの変更時とキャンセル時にだけ undoJustifications() で行う
         * @returns {void}
         */
        function undoPreview() {
            for (var i = appliedMoves.length - 1; i >= 0; i--) {
                appliedMoves[i].targetItem.translate(-appliedMoves[i].dx, -appliedMoves[i].dy);
            }
            appliedMoves = [];
        }

        /**
         * テキストの行揃えを justification にそろえ、元の値を記録する（巻き戻し用）。
         * 既に同じなら何もしない（無駄な undo を作らない）。テキスト以外は無視
         * @param {PageItem} targetItem - 対象のオブジェクト
         * @param {Justification|null} justification - 行揃え（null は変更しない）
         * @returns {void}
         */
        function applyJustification(targetItem, justification) {
            if (justification === null) return;
            if (targetItem.constructor.name !== "TextFrame") return;
            var currentJustification = targetItem.textRange.paragraphAttributes.justification;
            if (currentJustification === justification) return;
            appliedJustifications.push({ targetItem: targetItem, original: currentJustification });
            setParagraphJustification(targetItem, justification);
            boundsCacheStale = true;
        }

        /**
         * 記録した行揃え変更を元の値へ戻す
         * @returns {void}
         */
        function undoJustifications() {
            if (appliedJustifications.length === 0) return;
            for (var j = appliedJustifications.length - 1; j >= 0; j--) {
                setParagraphJustification(appliedJustifications[j].targetItem, appliedJustifications[j].original);
            }
            appliedJustifications = [];
            boundsCacheStale = true;
        }

        /**
         * 軸に沿って移動し、巻き戻し用に記録する。移動量は右（水平）/ 下（垂直）で正
         * @param {Object} axis - makeAxis() の軸
         * @param {PageItem} targetItem - 動かすオブジェクト
         * @param {number} travelAmount - 進行方向の移動量（pt）
         * @returns {void}
         */
        function moveByAxis(axis, targetItem, travelAmount) {
            if (travelAmount === 0) return;
            // 垂直は Illustrator の y が上ほど大きいので、下方向(+)へは translate(0, -移動量)
            // Vertical: Illustrator y grows upward, so moving down (+) means translate(0, -amount)
            var dx = axis.vertical ? 0 : travelAmount;
            var dy = axis.vertical ? -travelAmount : 0;
            targetItem.translate(dx, dy);
            appliedMoves.push({ targetItem: targetItem, dx: dx, dy: dy });
        }

        /**
         * グループのメンバー境界のキャッシュを選ぶ（プレビュー境界 / 幾何境界）
         * @param {Object} pair - グループの作業単位
         * @param {boolean} useVisible - プレビュー境界を使うか
         * @returns {number[][]} メンバーごとの境界
         */
        function getMemberBounds(pair, useVisible) {
            return useVisible ? pair.memberVis : pair.memberGeo;
        }

        /**
         * 境界を進行方向（先頭辺の昇順）に並べたインデックス配列を返す
         * @param {number[][]} boundsList - メンバーごとの境界
         * @param {Object} axis - makeAxis() の軸
         * @returns {number[]} インデックスの配列
         */
        function getMemberOrder(boundsList, axis) {
            var memberOrder = [];
            for (var i = 0; i < boundsList.length; i++) memberOrder.push(i);
            memberOrder.sort(function (indexA, indexB) { return axis.start(boundsList[indexA]) - axis.start(boundsList[indexB]); });
            return memberOrder;
        }

        /**
         * グループ内の全オブジェクトを進行方向順に並べ、隣り合う間隔を gap にそろえる。
         * 固定側のオブジェクトは動かさず、そこを起点にカスケードで再配置する。
         * axis で水平/垂直を切り替える（左右→水平、上下→垂直）
         * @param {Object} pair - グループの作業単位
         * @param {string} fixedSide - 固定する側
         * @param {number} gapInPoints - 間隔（pt）
         * @param {boolean} useVisible - プレビュー境界を使うか
         * @param {Object} axis - 間隔の軸
         * @returns {void}
         */
        function distributeGroup(pair, fixedSide, gapInPoints, useVisible, axis) {
            var members = pair.members;
            var cachedBounds = getMemberBounds(pair, useVisible);
            var count = members.length;
            var anchorEnd = isAnchorEnd(fixedSide);
            var memberOrder = getMemberOrder(cachedBounds, axis); // 先頭辺の昇順 / ordered by leading edge

            if (anchorEnd) {
                // 後ろ側（右 or 下）を固定し、後ろから前へ配置 / Anchor the trailing end; walk backward
                var nextLeadingEdge = axis.start(cachedBounds[memberOrder[count - 1]]); // 後ろ端オブジェクトの先頭辺（不動）
                for (var i = count - 2; i >= 0; i--) {
                    var memberIndex = memberOrder[i];
                    var desiredEnd = nextLeadingEdge - gapInPoints;
                    var shiftAmount = desiredEnd - axis.end(cachedBounds[memberIndex]);
                    moveByAxis(axis, members[memberIndex], shiftAmount);
                    nextLeadingEdge = axis.start(cachedBounds[memberIndex]) + shiftAmount; // この要素の新しい先頭辺 / its new leading edge
                }
            } else {
                // 先頭側（左 or 上）を固定し、前から後ろへ配置 / Anchor the leading end; walk forward
                var prevTrailingEdge = axis.end(cachedBounds[memberOrder[0]]); // 先頭端オブジェクトの後ろ辺（不動）
                for (var i = 1; i < count; i++) {
                    var memberIndex = memberOrder[i];
                    var desiredStart = prevTrailingEdge + gapInPoints;
                    var shiftAmount = desiredStart - axis.start(cachedBounds[memberIndex]);
                    moveByAxis(axis, members[memberIndex], shiftAmount);
                    prevTrailingEdge = axis.end(cachedBounds[memberIndex]) + shiftAmount; // この要素の新しい後ろ辺 / its new trailing edge
                }
            }
        }

        /**
         * オブジェクトを alignAxis 方向で基準にそろえる。start = 先頭辺（左 or 上）、end = 後ろ辺（右 or 下）、
         * center = 中央。ギャップ調整は gap 軸方向のみ動かすので、整列軸の位置はキャッシュ境界のまま使える
         * @param {Object} alignAxis - 整列の軸
         * @param {PageItem} targetItem - 動かすオブジェクト
         * @param {number[]} itemBounds - 動かすオブジェクトの境界
         * @param {number[]} anchorBounds - 基準の境界
         * @param {string} alignMode - "none" / "start" / "center" / "end"
         * @returns {void}
         */
        function alignToAnchor(alignAxis, targetItem, itemBounds, anchorBounds, alignMode) {
            if (alignMode === "none") return;
            var alignShift;
            if (alignMode === "start") {
                alignShift = alignAxis.start(anchorBounds) - alignAxis.start(itemBounds);
            } else if (alignMode === "end") {
                alignShift = alignAxis.end(anchorBounds) - alignAxis.end(itemBounds);
            } else { // center
                var anchorCenter = (alignAxis.start(anchorBounds) + alignAxis.end(anchorBounds)) / 2;
                var itemCenter = (alignAxis.start(itemBounds) + alignAxis.end(itemBounds)) / 2;
                alignShift = anchorCenter - itemCenter;
            }
            moveByAxis(alignAxis, targetItem, alignShift);
        }

        /**
         * グループの固定端メンバー（キーオブジェクト）のインデックスを返す。並び順は不変なので
         * キャッシュ境界で判定して十分（ライブ境界を読まない）
         * @param {Object} pair - グループの作業単位
         * @param {Object} gapAxis - 間隔の軸
         * @param {boolean} anchorEnd - 後ろ側を固定するか
         * @param {boolean} useVisible - プレビュー境界を使うか
         * @returns {number} メンバーのインデックス
         */
        function groupAnchorIndex(pair, gapAxis, anchorEnd, useVisible) {
            var memberOrder = getMemberOrder(getMemberBounds(pair, useVisible), gapAxis);
            return anchorEnd ? memberOrder[memberOrder.length - 1] : memberOrder[0];
        }

        /**
         * グループの各メンバーを固定端のメンバー（アンカー）に整列軸方向でそろえる。
         * 行揃えは間隔調整より前に適用して境界を取り直す（cachePairBounds）ので、ここはキャッシュ境界で十分
         * @param {Object} pair - グループの作業単位
         * @param {Object} gapAxis - 間隔の軸
         * @param {boolean} anchorEnd - 後ろ側を固定するか
         * @param {boolean} useVisible - プレビュー境界を使うか
         * @param {Object} alignAxis - 整列の軸
         * @param {string} alignMode - "none" / "start" / "center" / "end"
         * @returns {void}
         */
        function alignGroup(pair, gapAxis, anchorEnd, useVisible, alignAxis, alignMode) {
            if (alignMode === "none") return;
            var members = pair.members;
            var cachedBounds = getMemberBounds(pair, useVisible);
            var anchorIndex = groupAnchorIndex(pair, gapAxis, anchorEnd, useVisible); // 固定端のメンバー / member at the fixed end
            var anchorBounds = cachedBounds[anchorIndex];
            for (var i = 0; i < members.length; i++) {
                if (i === anchorIndex) continue;
                alignToAnchor(alignAxis, members[i], cachedBounds[i], anchorBounds, alignMode);
            }
        }

        /**
         * アートボードモード：オブジェクトとアートボード端の間隔（マージン）を指定値にそろえ、
         * アートボードを基準に整列し、位置オフセットをかける（アートボード基準なので常に対象）
         * @param {Object} pair - アートボードの作業単位（{ single }）
         * @param {Object} spacingContext - applySpacing() の設定
         * @returns {void}
         */
        function applyArtboardMargin(pair, spacingContext) {
            var axis = spacingContext.axis;
            var gapInPoints = spacingContext.gapInPoints;
            var singleBounds = spacingContext.useVisible ? pair.visibleS : pair.geometricS;
            var artboardRect = pair.artboardRect;
            var marginShift = spacingContext.anchorEnd
                ? (axis.end(artboardRect) - gapInPoints) - axis.end(singleBounds)    // 右/下端から内側へ / inset from trailing edge
                : (axis.start(artboardRect) + gapInPoints) - axis.start(singleBounds); // 左/上端から内側へ / inset from leading edge
            moveByAxis(axis, pair.single, marginShift);
            if (spacingContext.alignMode !== "none") {
                alignToAnchor(spacingContext.alignAxis, pair.single, singleBounds, artboardRect, spacingContext.alignMode);
            }
            moveByAxis(spacingContext.alignAxis, pair.single, spacingContext.offsetAlong); // 位置オフセット / offset
        }

        /**
         * グループモード：全オブジェクトを等間隔に分配（固定側を基準）し、整列してから、
         * キーオブジェクト（固定端メンバー）以外を直交方向へずらす
         * @param {Object} pair - グループの作業単位（{ members }）
         * @param {Object} spacingContext - applySpacing() の設定
         * @returns {void}
         */
        function applyGroupSpacing(pair, spacingContext) {
            var axis = spacingContext.axis;
            var anchorEnd = spacingContext.anchorEnd;
            var useVisible = spacingContext.useVisible;
            distributeGroup(pair, spacingContext.fixedSide, spacingContext.gapInPoints, useVisible, axis);
            alignGroup(pair, axis, anchorEnd, useVisible, spacingContext.alignAxis, spacingContext.alignMode);
            // 位置オフセット：キーオブジェクト（固定端メンバー）以外を直交方向へずらす / Offset non-key members only
            if (spacingContext.offsetAlong !== 0) {
                var keyMemberIndex = groupAnchorIndex(pair, axis, anchorEnd, useVisible);
                for (var j = 0; j < pair.members.length; j++) {
                    if (j === keyMemberIndex) continue;
                    moveByAxis(spacingContext.alignAxis, pair.members[j], spacingContext.offsetAlong);
                }
            }
        }

        /**
         * 自動ペア認識：2オブジェクトの間隔を調整する（行揃え後の境界）。固定側を基準に動く側だけ移動し、
         * 同じ動く側を整列軸でも固定側へそろえ、位置オフセットをかける
         * @param {Object} pair - ペアの作業単位（{ a, b }）
         * @param {Object} spacingContext - applySpacing() の設定
         * @returns {void}
         */
        function applyPairGap(pair, spacingContext) {
            var axis = spacingContext.axis;
            var gapInPoints = spacingContext.gapInPoints;
            var boundsA = spacingContext.useVisible ? pair.visibleA : pair.geometricA;
            var boundsB = spacingContext.useVisible ? pair.visibleB : pair.geometricB;

            // 進行方向の先頭辺で前後を判定 / Decide leading/trailing by the axis start edge
            var leadObject, trailObject, leadBounds, trailBounds;
            if (axis.start(boundsA) < axis.start(boundsB)) {
                leadObject = pair.a; leadBounds = boundsA;
                trailObject = pair.b; trailBounds = boundsB;
            } else {
                leadObject = pair.b; leadBounds = boundsB;
                trailObject = pair.a; trailBounds = boundsA;
            }

            var gapShift, movedObject, movedBounds, anchorBounds;
            if (spacingContext.anchorEnd) {
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
            if (spacingContext.alignMode !== "none") {
                alignToAnchor(spacingContext.alignAxis, movedObject, movedBounds, anchorBounds, spacingContext.alignMode);
            }
            moveByAxis(spacingContext.alignAxis, movedObject, spacingContext.offsetAlong); // 位置オフセット（移動側のみ）/ position offset (moved object only)
        }

        /**
         * 設定値で各ペアの間隔を調整する。固定側から軸（水平/垂直）と固定端（先頭/後ろ）を決める。
         * alignMode は整列（ギャップ軸に直交する方向、固定オブジェクト基準）。
         * 位置オフセット（offsetAlong）も直交方向で、キーオブジェクトでない側（移動側）だけをずらす（右＝正／下＝正）
         * @param {string} fixedSide - 固定する側
         * @param {number} gapInPoints - 間隔（pt）
         * @param {string} boundsType - "visibleBounds" または "geometricBounds"
         * @param {string} alignMode - "none" / "start" / "center" / "end"
         * @param {number} offsetAlong - 位置オフセット（pt）
         * @param {string} justifyMode - 行揃えモード
         * @returns {void}
         */
        function applySpacing(fixedSide, gapInPoints, boundsType, alignMode, offsetAlong, justifyMode) {
            var axis = axisForSide(fixedSide);
            var spacingContext = {
                fixedSide: fixedSide,
                gapInPoints: gapInPoints,
                useVisible: (boundsType === "visibleBounds"),
                axis: axis,
                anchorEnd: isAnchorEnd(fixedSide),
                alignAxis: makeAxis(!axis.vertical), // 整列はギャップ軸に直交 / perpendicular to the gap axis
                alignMode: alignMode,
                offsetAlong: offsetAlong,
                justifyMode: justifyMode
            };

            // 行揃えを先に全テキストへ適用し、境界を取り直す（ポイント文字は行揃えで字幅が変わるため、
            // 間隔・整列を行揃え後の実際の形で計算する）。applyJustification は冪等。
            // Apply justification to all text FIRST, then re-measure bounds, so gap/align use the
            // post-justification shape (point-text width changes with justification). Idempotent.
            for (var i = 0; i < objectPairs.length; i++) {
                var pairItems = getPairItems(objectPairs[i]);
                for (var j = 0; j < pairItems.length; j++) {
                    applyJustification(pairItems[j], resolveJustification(pairItems[j], spacingContext));
                }
            }
            // 行揃えで境界が変わったときだけ取り直す / Re-cache only when justification changed the bounds
            if (boundsCacheStale) {
                cachePairBounds(objectPairs);
                boundsCacheStale = false;
            }

            for (var k = 0; k < objectPairs.length; k++) {
                var pair = objectPairs[k];
                if (pair.single) {
                    applyArtboardMargin(pair, spacingContext);
                } else if (pair.members) {
                    applyGroupSpacing(pair, spacingContext);
                } else {
                    applyPairGap(pair, spacingContext);
                }
            }
        }

        /**
         * 直前のプレビューを巻き戻してから再適用する。
         * 行揃え・整列ともプレビュー時点で最終結果と一致するので、OK では別処理は不要
         * @param {string} fixedSide - 固定する側
         * @param {number} gapInPoints - 間隔（pt）
         * @param {string} boundsType - "visibleBounds" または "geometricBounds"
         * @param {string} alignMode - "none" / "start" / "center" / "end"
         * @param {number} offsetAlong - 位置オフセット（pt）
         * @param {string} justifyMode - 行揃えモード
         * @returns {void}
         */
        function runPreview(fixedSide, gapInPoints, boundsType, alignMode, offsetAlong, justifyMode) {
            undoPreview();
            applySpacing(fixedSide, gapInPoints, boundsType, alignMode, offsetAlong, justifyMode);
            app.redraw();
        }

        /**
         * 各ペアの元の境界をキャッシュする。
         * 適用時は常に元位置へ巻き戻してから計算するので、ここで一度取れば使い回せる
         * @param {Object[]} pairs - 作業単位の配列
         * @returns {void}
         */
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

        /**
         * モードに応じてペアを組み直す。境界キャッシュは必ず元位置で取るため、先にプレビューを巻き戻してから組む
         * @param {string} mode - "auto" / "group" / "artboard"
         * @returns {void}
         */
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

        /**
         * 前回終了時の設定をダイアログへ戻す（モード・キー・プレビュー境界・整列・行揃え・間隔・左右/上下オフセット）。
         * モードとキー側は初期間隔を測ったときと同じ値を使う。数値は pt で保存しているので定規の単位の表示へ戻す。
         * 整列とオフセットは向きごとの控えに入れておき、パネルを切り替えたときに反映する
         * @param {Object} dialogControls - buildSettingsDialog() の戻り値
         * @param {Object} orientationValues - 向きごとの控え（h / v）
         * @returns {void}
         */
        function restoreDialogSettings(dialogControls, orientationValues) {
            var alignRadios = dialogControls.alignmentRefs.alignRadios;
            dialogControls.modeRefs.setMode(initialMode);
            dialogControls.fixedSideRefs.setFixedSide(initialFixedSide);
            if (savedSettings.previewBounds === "true") dialogControls.gapRefs.previewBoundsCheckbox.value = true;
            // 整列（キーは none/start/center/end）/ Alignment
            if (alignRadios[savedSettings.alignH]) orientationValues.h.align = savedSettings.alignH;
            if (alignRadios[savedSettings.alignV]) orientationValues.v.align = savedSettings.alignV;
            // テキストの行揃え / Text alignment
            if (savedSettings.justify) dialogControls.justifyRefs.setJustifyMode(savedSettings.justify);
            // 数値（間隔・左右・上下）/ Numeric values (gap, horizontal/vertical offsets)
            var savedGapDisplay = savedPointsToDisplayText(savedSettings.gap);
            if (savedGapDisplay !== null) dialogControls.gapRefs.spacingInput.text = savedGapDisplay;
            var savedOffsetHDisplay = savedPointsToDisplayText(savedSettings.offsetH);
            if (savedOffsetHDisplay !== null) orientationValues.h.offset = savedOffsetHDisplay;
            var savedOffsetVDisplay = savedPointsToDisplayText(savedSettings.offsetV);
            if (savedOffsetVDisplay !== null) orientationValues.v.offset = savedOffsetVDisplay;
        }

        /**
         * ダイアログを生成して表示する。OK では現在の設定をすべて保存する（数値は pt で保存）
         * @returns {boolean} OK なら true
         */
        function showSettingsDialog() {
            var dialogControls = buildSettingsDialog(initialGapPoints);
            var settingsDialog = dialogControls.settingsDialog;
            var modeRefs = dialogControls.modeRefs;
            var fixedSideRefs = dialogControls.fixedSideRefs;
            var gapRefs = dialogControls.gapRefs;
            var alignmentRefs = dialogControls.alignmentRefs;
            var justifyRefs = dialogControls.justifyRefs;
            var alignRadios = alignmentRefs.alignRadios;
            var offsetInput = alignmentRefs.offsetInput;

            var orientationSwitcher = createOrientationSwitcher(fixedSideRefs, alignmentRefs);
            var orientationValues = orientationSwitcher.orientationValues;
            restoreDialogSettings(dialogControls, orientationValues);

            /**
             * 現在の設定でプレビューを更新する（行揃え・整列とも最終結果と一致）
             * @returns {void}
             */
            function refreshPreview() {
                runPreview(fixedSideRefs.getFixedSide(), gapRefs.getSpacingInPoints(), gapRefs.getBoundsType(),
                    getAlignValue(alignRadios), displayTextToPoints(offsetInput.text), justifyRefs.getJustifyMode());
            }

            /**
             * 行揃えの対象・向きが変わりうる操作（モード／キー／整列の切替）用：先に行揃えを元へ戻してから
             * プレビューを更新する。これで「なし」へ戻したときや別の向きへ変えたときに正しく反映される
             * @returns {void}
             */
            function refreshPreviewResetJustify() {
                undoJustifications();
                refreshPreview();
            }

            /**
             * モードを切り替えてペアを組み直し、プレビューを更新する。
             * モードで対象オブジェクトが変わるので行揃えを元へ戻してから組み直す
             * @returns {void}
             */
            function onModeChange() {
                undoJustifications();
                buildPairs(modeRefs.getMode());
                refreshPreview();
            }

            /**
             * 整列の変更時：中央↔それ以外でオフセットの有効/無効が変わるので更新してからプレビュー
             * @returns {void}
             */
            function onAlignChange() {
                orientationSwitcher.updateActivePanels();
                refreshPreviewResetJustify();
            }

            // 設定変更でライブプレビュー / Update preview on change
            for (var i = 0; i < modeRefs.modeRadios.length; i++) {
                modeRefs.modeRadios[i].onClick = onModeChange;
            }
            // 固定側のラジオ：手動で排他にしてからプレビュー更新（軸の切替もここで反映）
            // Fixed-side radios: enforce exclusivity by hand, then refresh (axis switch applies here too)
            var fixedRadios = fixedSideRefs.fixedRadios;
            for (var j = 0; j < fixedRadios.length; j++) {
                fixedRadios[j].onClick = (function (fixedRadio) {
                    return function () {
                        fixedSideRefs.selectFixedRadio(fixedRadio);
                        orientationSwitcher.updateActivePanels();
                        refreshPreviewResetJustify();
                    };
                })(fixedRadios[j]);
            }
            changeValueByArrowKey(gapRefs.spacingInput, refreshPreview); // ↑↓キーで増減＋プレビュー更新 / Arrow keys + preview
            gapRefs.spacingInput.onChanging = refreshPreview;
            // 位置オフセット：↑↓キーで増減＋入力でプレビュー更新 / Offset: arrow-key step + refresh on change
            changeValueByArrowKey(offsetInput, refreshPreview);
            offsetInput.onChanging = refreshPreview;
            gapRefs.previewBoundsCheckbox.onClick = refreshPreview;
            // 整列ラジオ：オフセットの有効/無効を更新し、行揃えを戻してから更新 / Alignment radios
            for (var alignKey in alignRadios) { alignRadios[alignKey].onClick = onAlignChange; }
            // 整列のキーボードショートカット（水平 L/C/R・垂直 T/M/B）。今の向きで読み替える / Alignment keyboard shortcuts
            addAlignmentKeyHandler(settingsDialog, alignRadios, orientationSwitcher.isVerticalGap, onAlignChange);
            // テキストの行揃えボタン：押した値をアクティブにし、行揃えを戻してから再適用
            // Justification buttons: activate the clicked value, revert justification, then refresh
            var justifyButtons = justifyRefs.buttons;
            for (var k = 0; k < justifyButtons.length; k++) {
                justifyButtons[k].onClick = function () {
                    justifyRefs.setJustifyMode(this.justifyId);
                    refreshPreviewResetJustify();
                };
            }

            // ダイアログ表示時に既定モードでペアを組んで初回プレビュー（同期側 undo を避けて onShow から起動）
            // Build pairs for the default mode, then run the first preview (from onShow to avoid sync undo)
            settingsDialog.onShow = function () {
                orientationSwitcher.updateActivePanels(); // 既定のキー側に合わせて水平/垂直パネルとオフセットの有効/無効を初期化 / Init enabled state
                buildPairs(modeRefs.getMode());
                refreshPreview();
            };

            var isAccepted = (settingsDialog.show() === 1);
            if (isAccepted) {
                // プレビュー状態がそのまま最終結果（行揃え・整列とも反映済み）なので、確定処理は保存のみ。
                // The preview already is the final result (justification + alignment), so OK just saves.
                orientationSwitcher.stashOrientationValues(); // 出ている向きの値を控えへ入れてから保存 / stash the shown values first
                saveSettings({
                    mode: modeRefs.getMode(),
                    fixedSide: fixedSideRefs.getFixedSide(),
                    previewBounds: gapRefs.previewBoundsCheckbox.value ? "true" : "false",
                    alignH: orientationValues.h.align,
                    alignV: orientationValues.v.align,
                    justify: justifyRefs.getJustifyMode(),
                    gap: String(gapRefs.getSpacingInPoints()),
                    offsetH: String(displayTextToPoints(orientationValues.h.offset)),
                    offsetV: String(displayTextToPoints(orientationValues.v.offset))
                });
            }
            return isAccepted;
        }

        // OK はプレビューをそのまま確定。キャンセルは位置と行揃えを巻き戻す
        // OK keeps the applied preview; cancel reverts it (position + justification)
        if (!showSettingsDialog()) {
            undoPreview();
            undoJustifications();
            app.redraw();
        }
    })();

})();
