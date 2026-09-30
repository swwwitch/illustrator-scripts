#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

UIPartsPlain.jsx からさらにステップボタンを外し、数値欄を入力欄だけにした比較用のテンプレートです。
基準点はラジオボタン9個、リンクアイコンはチェックボックスです。

### Overview

A comparison template that goes one step further than UIPartsPlain.jsx and drops the stepper buttons, leaving plain input fields.
The reference point uses nine radio buttons and the link toggle a checkbox.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "UIPartsUltraPlain";            /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-09-30";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-30";                   /* 更新日 / last updated */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // 【比較のためのメモ / Notes for the comparison】
    // UIPartsPlain.jsx からステップボタンを外したもの。↑↓キーでの増減と、直接入力した値の整え方は残す
    //   ステップボタン … なし（入力欄だけ）
    //   基準点         … radiobutton 9個。行ごとの group に分かれるので、排他は手で管理する
    //   リンクアイコン … checkbox
    //   UI の明暗      … 配色を持たないので不要

    // =========================================
    // レイアウト / Layout
    // =========================================
    var DEMO_LABEL_WIDTH = 48;             /* デモの項目名の幅 / label width in the demo */
    var DEMO_PANEL_MARGINS = [15, 20, 15, 10]; /* デモのパネルの余白 / panel margins in the demo */

    // =========================================
    // 数値欄 / Number fields
    // =========================================

    // -----------------------------------------
    // 数値欄の寸法・増減量 / Field metrics and steps
    // -----------------------------------------
    var STEPPER_FIELD_SPACING  = 3;   /* 項目名と入力欄の間隔 / spacing between the label and the field */
    var STEPPER_SHIFT_MULTIPLE = 10;  /* shift＋↑↓でそろえる倍数 / Shift+arrow snaps to multiples of this */
    var STEPPER_OPTION_STEP    = 0.1; /* option＋↑↓の増減量 / Option+arrow step */

    // -----------------------------------------
    // 数値欄を作る（外から呼ぶ関数） / Public API
    // -----------------------------------------
    /**
     * 「項目名・入力欄」をひと組にした数値欄を追加する。
     * ↑↓キーで増減する。直接入力した値も、フォーカスが外れたときに
     * 整数化・下限・上限・単位（「20 mm」の形）へそろえ、数値でなければ直前の値に戻す
     * @param {Group|Panel} parent - 追加先
     * @param {Object} fieldOptions - label（コロン込みの項目名）/ labelWidth / text / characters /
     *     step / min / max / integer（true で整数のみ）/ unit / onStep
     * @returns {EditText} 入力欄（項目名は .fieldLabel で参照できる）
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

        var numberInput = fieldRowGroup.add("edittext", undefined, fieldOptions.text || "");
        numberInput.characters = fieldOptions.characters || 6;
        numberInput.fieldLabel = fieldLabel;

        bindSteppedArrowKeys(numberInput, fieldOptions);

        /* 直接入力をそろえる。数値でなければ直前の値に戻す / normalize typed values; revert non-numbers */
        numberInput.lastValidText = numberInput.text;
        numberInput.onChange = function () {
            var value = parseFloat(numberInput.text);
            if (isNaN(value)) {
                numberInput.text = numberInput.lastValidText;
                return;
            }
            writeSteppedValue(numberInput, value, fieldOptions);
        };
        return numberInput;
    }

    /**
     * 数値欄の有効／無効を、項目名ごとまとめて切り替える（標準コントロールなのでディム表示は自動）
     * @param {EditText} numberInput - addSteppedField() で作った入力欄
     * @param {boolean} isEnabled - 有効にするなら true
     * @returns {void}
     */
    function setSteppedFieldEnabled(numberInput, isEnabled) {
        numberInput.enabled = isEnabled;
        numberInput.fieldLabel.enabled = isEnabled;
    }

    /**
     * 入力欄の↑↓キーで値を増減させる（shift なら STEPPER_SHIFT_MULTIPLE の倍数へ、option なら STEPPER_OPTION_STEP ずつ。下限・上限で止める）。
     * ほかのキーは素通し
     * @param {EditText} numberInput - 対象の入力欄
     * @param {Object} stepOptions - step（増減量）/ min / max / integer / unit（例 " mm"）/ onStep(numberInput)
     * @returns {void}
     */
    function bindSteppedArrowKeys(numberInput, stepOptions) {
        numberInput.addEventListener("keydown", function (event) {
            if (event.keyName !== "Up" && event.keyName !== "Down") return;
            var direction = (event.keyName === "Up") ? 1 : -1;
            var value = parseFloat(numberInput.text);
            if (isNaN(value)) value = 0;
            writeSteppedValue(numberInput, computeSteppedValue(value, direction, stepOptions), stepOptions);
            if (stepOptions.onStep) stepOptions.onStep(numberInput);
            event.preventDefault(); /* カーソル移動を止める / keep the caret from moving */
        });
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

    // =========================================
    // 基準点ウィジェット / Anchor widget
    // =========================================

    // -----------------------------------------
    // 基準点ウィジェットの定数 / Anchor widget constants
    // -----------------------------------------
    var ANCHOR_WIDGET_NONE = -1; /* 未選択のインデックス / index while nothing is selected */

    /* セルの名前（行優先：上 → 中 → 下、列：左 → 中 → 右）。Transformation の列挙名にそろえる
       Cell names in row-major order, matching the Transformation enumeration */
    var ANCHOR_WIDGET_NAMES = ["topLeft", "top", "topRight", "left", "center", "right", "bottomLeft", "bottom", "bottomRight"];

    // -----------------------------------------
    // ウィジェットを作る・読み書きする（外から呼ぶ関数） / Public API
    // -----------------------------------------
    /**
     * 基準点（3×3）を選ぶラジオボタンを追加する。クリックしたセルを選び、onChange を呼ぶ
     * @param {Group|Panel} parent - 追加先
     * @param {number|string} initialValue - 最初に選ぶセル（0〜8 か "topLeft" などの名前。allowNone なら -1 も可）
     * @param {Function} [onChange] - クリックで選んだときに呼ぶ関数（引数はセルのインデックスとウィジェット）
     * @param {Object} [widgetOptions] - allowNone（true で未選択 -1 を許す）/ disabledCells（選べないセルの配列）
     * @returns {Group} ウィジェット（ラジオは .anchorRadios、値は getAnchorWidgetIndex() / getAnchorWidgetName() で読む）
     */
    function addAnchorWidget(parent, initialValue, onChange, widgetOptions) {
        var anchorOptions = widgetOptions || {};
        var anchorWidget = parent.add("group");
        anchorWidget.orientation = "column";
        anchorWidget.alignChildren = ["center", "center"];
        anchorWidget.spacing = 0;
        anchorWidget.anchorAllowNone = !!anchorOptions.allowNone;
        anchorWidget.anchorRadios = [];

        /* 3行の group にラジオを3個ずつ並べる / three rows of three radios */
        for (var row = 0; row < 3; row++) {
            var anchorRowGroup = anchorWidget.add("group");
            anchorRowGroup.orientation = "row";
            anchorRowGroup.spacing = 0;
            for (var column = 0; column < 3; column++) {
                var anchorRadio = anchorRowGroup.add("radiobutton", undefined, "");
                anchorRadio.anchorIndex = row * 3 + column;
                anchorRadio.onClick = function () {
                    /* ラジオは同じ親の中だけ排他なので、ほかの行の選択は手で外す / radios are exclusive only within one parent */
                    selectAnchorRadio(anchorWidget, this.anchorIndex);
                    if (onChange) onChange(anchorWidget.anchorWidgetIndex, anchorWidget);
                };
                anchorWidget.anchorRadios.push(anchorRadio);
            }
        }
        setAnchorWidgetCellsDisabled(anchorWidget, anchorOptions.disabledCells || []);
        setAnchorWidgetValue(anchorWidget, initialValue);
        return anchorWidget;
    }

    /**
     * 選択中のセルのインデックスを返す
     * @param {Group} anchorWidget - addAnchorWidget() で作ったウィジェット
     * @returns {number} 0〜8（行優先）。未選択なら -1
     */
    function getAnchorWidgetIndex(anchorWidget) {
        return anchorWidget.anchorWidgetIndex;
    }

    /**
     * 選択中のセルの名前を返す
     * @param {Group} anchorWidget - addAnchorWidget() で作ったウィジェット
     * @returns {string} "topLeft" など。未選択なら ""
     */
    function getAnchorWidgetName(anchorWidget) {
        return ANCHOR_WIDGET_NAMES[anchorWidget.anchorWidgetIndex] || "";
    }

    /**
     * 選択するセルを変える（onChange は呼ばない）
     * @param {Group} anchorWidget - addAnchorWidget() で作ったウィジェット
     * @param {number|string} anchorValue - 0〜8 か名前（allowNone なら -1 も可）
     * @returns {void}
     */
    function setAnchorWidgetValue(anchorWidget, anchorValue) {
        selectAnchorRadio(anchorWidget, resolveAnchorWidgetIndex(anchorValue, anchorWidget.anchorAllowNone));
    }

    /**
     * ウィジェットの有効／無効を切り替える（標準コントロールなのでディム表示は自動）
     * @param {Group} anchorWidget - addAnchorWidget() で作ったウィジェット
     * @param {boolean} isEnabled - 有効にするなら true
     * @returns {void}
     */
    function setAnchorWidgetEnabled(anchorWidget, isEnabled) {
        anchorWidget.enabled = isEnabled;
    }

    /**
     * 選べないセルを指定し直す（選択中のセルは変えない）
     * @param {Group} anchorWidget - addAnchorWidget() で作ったウィジェット
     * @param {number[]} disabledCells - 選べないセルのインデックス（空配列ですべて選べる）
     * @returns {void}
     */
    function setAnchorWidgetCellsDisabled(anchorWidget, disabledCells) {
        var cellFlags = toAnchorCellFlags(disabledCells);
        for (var i = 0; i < anchorWidget.anchorRadios.length; i++) {
            anchorWidget.anchorRadios[i].enabled = !cellFlags[i];
        }
    }

    /**
     * 指定したセルのラジオだけを選び、ほかは外す（-1 ならすべて外す）
     * @param {Group} anchorWidget - addAnchorWidget() で作ったウィジェット
     * @param {number} anchorIndex - 0〜8 か -1
     * @returns {void}
     */
    function selectAnchorRadio(anchorWidget, anchorIndex) {
        anchorWidget.anchorWidgetIndex = anchorIndex;
        for (var i = 0; i < anchorWidget.anchorRadios.length; i++) {
            anchorWidget.anchorRadios[i].value = (i === anchorIndex);
        }
    }

    // -----------------------------------------
    // 値の変換 / Value helpers
    // -----------------------------------------
    /**
     * セルのインデックスか名前を 0〜8 のインデックスにする。解釈できない値は中央（4）
     * @param {number|string} anchorValue - 0〜8 / -1 / "topLeft" などの名前
     * @param {boolean} [allowNone] - true なら -1（未選択）をそのまま返す
     * @returns {number} 0〜8。allowNone で -1 を渡したときだけ -1
     */
    function resolveAnchorWidgetIndex(anchorValue, allowNone) {
        if (typeof anchorValue === "string") {
            for (var i = 0; i < ANCHOR_WIDGET_NAMES.length; i++) {
                if (ANCHOR_WIDGET_NAMES[i] === anchorValue) return i;
            }
            return 4;
        }
        if (anchorValue === ANCHOR_WIDGET_NONE && allowNone) return ANCHOR_WIDGET_NONE;
        if (typeof anchorValue === "number" && anchorValue >= 0 && anchorValue <= 8 && anchorValue === Math.floor(anchorValue)) {
            return anchorValue;
        }
        return 4;
    }

    /**
     * セルの位置を割合で返す（左・上が 0、中央が 0.5、右・下が 1）
     * @param {number|string} anchorValue - 0〜8 か名前
     * @returns {number[]} [横の割合, 縦の割合]
     */
    function getAnchorRatio(anchorValue) {
        var anchorIndex = resolveAnchorWidgetIndex(anchorValue);
        return [(anchorIndex % 3) / 2, Math.floor(anchorIndex / 3) / 2];
    }

    /**
     * 境界ボックス上の基準点の座標を返す（Illustrator の [左, 上, 右, 下] でも、y 下向きの座標でもそのまま使える）
     * @param {number[]} bounds - [左, 上, 右, 下]（geometricBounds・visibleBounds・artboardRect など）
     * @param {number|string} anchorValue - 0〜8 か名前
     * @returns {number[]} [x, y]
     */
    function getAnchorPointOnBounds(bounds, anchorValue) {
        var anchorRatio = getAnchorRatio(anchorValue);
        return [
            bounds[0] + (bounds[2] - bounds[0]) * anchorRatio[0],
            bounds[1] + (bounds[3] - bounds[1]) * anchorRatio[1]
        ];
    }

    /**
     * resize()・rotate()・transform() に渡す基準点を返す（Illustrator 専用）。
     * 基準は効果を含まない境界（geometricBounds）
     * @param {number|string} anchorValue - 0〜8 か名前
     * @returns {Transformation} Transformation.TOPLEFT など
     */
    function getAnchorTransformation(anchorValue) {
        var transformations = [
            Transformation.TOPLEFT, Transformation.TOP, Transformation.TOPRIGHT,
            Transformation.LEFT, Transformation.CENTER, Transformation.RIGHT,
            Transformation.BOTTOMLEFT, Transformation.BOTTOM, Transformation.BOTTOMRIGHT
        ];
        return transformations[resolveAnchorWidgetIndex(anchorValue)];
    }

    /**
     * symbols.add() に渡す登録点を返す（Illustrator 専用）
     * @param {number|string} anchorValue - 0〜8 か名前
     * @returns {SymbolRegistrationPoint} SymbolRegistrationPoint.SYMBOLTOPLEFTPOINT など
     */
    function getAnchorSymbolRegistrationPoint(anchorValue) {
        var registrationPoints = [
            SymbolRegistrationPoint.SYMBOLTOPLEFTPOINT, SymbolRegistrationPoint.SYMBOLTOPMIDDLEPOINT, SymbolRegistrationPoint.SYMBOLTOPRIGHTPOINT,
            SymbolRegistrationPoint.SYMBOLMIDDLELEFTPOINT, SymbolRegistrationPoint.SYMBOLCENTERPOINT, SymbolRegistrationPoint.SYMBOLMIDDLERIGHTPOINT,
            SymbolRegistrationPoint.SYMBOLBOTTOMLEFTPOINT, SymbolRegistrationPoint.SYMBOLBOTTOMMIDDLEPOINT, SymbolRegistrationPoint.SYMBOLBOTTOMRIGHTPOINT
        ];
        return registrationPoints[resolveAnchorWidgetIndex(anchorValue)];
    }

    /**
     * セルのインデックスの配列を、9個の真偽値に直す
     * @param {number[]} [cellIndexes] - セルのインデックスの配列
     * @returns {boolean[]} 含まれるセルだけ true
     */
    function toAnchorCellFlags(cellIndexes) {
        var cellFlags = [false, false, false, false, false, false, false, false, false];
        if (!cellIndexes) return cellFlags;
        for (var i = 0; i < cellIndexes.length; i++) {
            if (cellIndexes[i] >= 0 && cellIndexes[i] <= 8) cellFlags[cellIndexes[i]] = true;
        }
        return cellFlags;
    }

    // =========================================
    // リンクアイコン / Link toggle
    // =========================================

    /**
     * 連動の ON／OFF を切り替えるチェックボックスを追加する
     * @param {Group} parent - 追加先
     * @param {boolean} initialValue - 連動の初期値
     * @param {Function} onToggle - 切り替えたあとに呼ぶ関数
     * @returns {Checkbox} チェックボックス（.value で連動中かを読む）
     */
    function addLinkToggle(parent, initialValue, onToggle) {
        var linkToggle = parent.add("checkbox", undefined, getLabel("checkbox.link"));
        linkToggle.value = initialValue;
        linkToggle.onClick = function () {
            if (onToggle) onToggle();
        };
        return linkToggle;
    }

    /**
     * 連動の状態をコードから変える（onToggle は呼ばない）
     * @param {Checkbox} linkToggle - addLinkToggle() で作ったチェックボックス
     * @param {boolean} isLinked - 連動にするなら true
     * @returns {void}
     */
    function setLinkToggleValue(linkToggle, isLinked) {
        linkToggle.value = isLinked;
    }

    /**
     * チェックボックスの有効／無効を切り替える
     * @param {Checkbox} linkToggle - addLinkToggle() で作ったチェックボックス
     * @param {boolean} isEnabled - 有効にするなら true
     * @returns {void}
     */
    function setLinkToggleEnabled(linkToggle, isEnabled) {
        linkToggle.enabled = isEnabled;
    }

    // =========================================
    // ボタン行 / Button row
    // =========================================

    // ボタン行（再利用パーツ） / Button row (reusable)

    var BUTTON_ROW_TOP_MARGIN = 5; /* ボタン行の上の余白 / top margin of the button row */
    var BUTTON_ROW_SPACING = 10;   /* ボタンどうしの間隔 / spacing between buttons */
    var BUTTON_ROW_CENTER_MAX_WIDTH = 200; /* 右のボタンだけの行を中央に置く、ダイアログの内側の最大幅（px、左右の余白を除く）。広いダイアログは右揃え / max inner dialog width (px, margins excluded) that centers a right-only row; wider dialogs keep it right-aligned */

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
        btnRowGroup.margins = [0, BUTTON_ROW_TOP_MARGIN, 0, 0];
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
     * 項目名の文言の末尾にコロンを付ける（日本語は全角「：」、英語は半角「:」）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {Object|Array} [placeholderValues] - getLabel と同じ
     * @returns {string} コロン付きの文言
     */
    function labelText(labelRef, placeholderValues) {
        return getLabel(labelRef, placeholderValues) + (uiLang === "ja" ? "：" : ":");
    }

    /**
     * 「項目名：値」の1行を返す（日本語は「件数：5」、英語は「Count: 5」とコロンのあとに空白を入れる）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {string|number} value - コロンのあとに続ける値
     * @returns {string} 項目名と値をつないだ文字列
     */
    function labelValueText(labelRef, value) {
        return labelText(labelRef) + (uiLang === "ja" ? "" : " ") + value;
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
            title: { ja: "UI パーツ", en: "UI Parts" }
        },
        panel: {
            size: { ja: "サイズ", en: "Size" },
            anchor: { ja: "基準点", en: "Reference Point" }
        },
        fieldLabel: {
            width: { ja: "幅", en: "Width" },
            height: { ja: "高さ", en: "Height" },
            selected: { ja: "選択", en: "Selected" }
        },
        checkbox: {
            link: { ja: "連動", en: "Link" },
            enableControls: { ja: "有効", en: "Enabled" }
        },
        tooltip: {
            linkToggle: { ja: "幅と高さを同じ値にそろえる（クリックで切り替え）", en: "Keep the width and height equal (click to toggle)" },
            anchor: { ja: "基準点（拡大・縮小、回転の基点）", en: "Reference point (origin for scale / rotate)" },
            enableControls: { ja: "OFF にすると、ステップボタン・基準点・リンクアイコンをディム表示にします", en: "When off, dims the steppers, the reference point and the link toggle" }
        },
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" }
        }
    };

    // =========================================
    // メイン処理（デモ） / Main (demo)
    // =========================================

    /**
     * 「サイズ」パネル（幅・高さの欄とリンクのチェックボックス）を追加する
     * @param {Group} parent - 追加先
     * @returns {{widthInput: EditText, heightInput: EditText, linkToggle: Checkbox}} 作ったコントロール
     */
    function buildDemoSizePanel(parent) {
        var sizePanel = parent.add("panel", undefined, getLabel("panel.size"));
        sizePanel.orientation = "column";
        sizePanel.alignChildren = ["left", "top"];
        sizePanel.margins = DEMO_PANEL_MARGINS;

        /* 入力欄を縦に積んだ group とチェックボックスを横に並べ、上下中央に置く / fields stacked, checkbox centered beside them */
        var sizeRowGroup = sizePanel.add("group");
        sizeRowGroup.orientation = "row";
        sizeRowGroup.alignChildren = ["left", "center"];
        var sizeFieldsColumn = sizeRowGroup.add("group");
        sizeFieldsColumn.orientation = "column";
        sizeFieldsColumn.alignChildren = ["left", "top"];

        var linkToggle;
        var widthInput;
        var heightInput;

        /**
         * 連動中なら、変えた欄の値をもう一方へ写す
         * @param {EditText} editedInput - 変えた欄
         * @returns {void}
         */
        function copyWhenLinked(editedInput) {
            if (!linkToggle.value) return;
            var otherInput = (editedInput === widthInput) ? heightInput : widthInput;
            otherInput.text = editedInput.text;
        }

        widthInput = addSteppedField(sizeFieldsColumn, {
            label: labelText("fieldLabel.width"), labelWidth: DEMO_LABEL_WIDTH,
            text: "100 mm", characters: 8, step: 1, min: 1, unit: " mm",
            onStep: copyWhenLinked
        });
        heightInput = addSteppedField(sizeFieldsColumn, {
            label: labelText("fieldLabel.height"), labelWidth: DEMO_LABEL_WIDTH,
            text: "100 mm", characters: 8, step: 1, min: 1, unit: " mm",
            onStep: copyWhenLinked
        });
        widthInput.onChanging = function () { copyWhenLinked(widthInput); };
        heightInput.onChanging = function () { copyWhenLinked(heightInput); };

        linkToggle = addLinkToggle(sizeRowGroup, true, function () { copyWhenLinked(widthInput); });
        linkToggle.helpTip = getLabel("tooltip.linkToggle");
        return { widthInput: widthInput, heightInput: heightInput, linkToggle: linkToggle };
    }

    /**
     * 「基準点」パネル（ラジオボタン9個と選択中のセルの表示）を追加する
     * @param {Group} parent - 追加先
     * @returns {Group} 基準点ウィジェット
     */
    function buildDemoAnchorPanel(parent) {
        var anchorPanel = parent.add("panel", undefined, getLabel("panel.anchor"));
        anchorPanel.orientation = "column";
        anchorPanel.alignChildren = ["center", "top"];
        anchorPanel.margins = DEMO_PANEL_MARGINS;
        var anchorWidget = addAnchorWidget(anchorPanel, "center", function (anchorIndex, clickedWidget) {
            selectedText.text = labelValueText("fieldLabel.selected", getAnchorWidgetName(clickedWidget));
        });
        /* group には helpTip が効かないので、ラジオ1個ずつに付ける / set the tip on each radio */
        for (var i = 0; i < anchorWidget.anchorRadios.length; i++) {
            anchorWidget.anchorRadios[i].helpTip = getLabel("tooltip.anchor");
        }
        var selectedText = anchorPanel.add("statictext", undefined, labelValueText("fieldLabel.selected", "center"));
        selectedText.preferredSize.width = 140; /* 実行時に文字を入れるので幅を確保 / reserve the width for runtime text */
        selectedText.justify = "center";
        return anchorWidget;
    }

    /**
     * すべての部品を並べたデモのダイアログを表示する
     * @returns {void}
     */
    function showDemoDialog() {
        var demoDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        demoDialog.orientation = "column";
        demoDialog.alignChildren = ["fill", "top"];

        var panelColumnsGroup = demoDialog.add("group");
        panelColumnsGroup.orientation = "row";
        panelColumnsGroup.alignChildren = ["fill", "fill"];
        var sizeControls = buildDemoSizePanel(panelColumnsGroup);
        var anchorWidget = buildDemoAnchorPanel(panelColumnsGroup);

        /* 有効／無効の切り替え / toggle enabled state */
        var enableCheckbox = demoDialog.add("checkbox", undefined, getLabel("checkbox.enableControls"));
        enableCheckbox.helpTip = getLabel("tooltip.enableControls");
        enableCheckbox.value = true;
        enableCheckbox.onClick = function () {
            var isEnabled = enableCheckbox.value;
            setSteppedFieldEnabled(sizeControls.widthInput, isEnabled);
            setSteppedFieldEnabled(sizeControls.heightInput, isEnabled);
            setLinkToggleEnabled(sizeControls.linkToggle, isEnabled);
            setAnchorWidgetEnabled(anchorWidget, isEnabled);
        };

        var buttonRow = addButtonRow(demoDialog);
        buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
        alignRightOnlyButtonRow(buttonRow); /* 右のボタンだけなので中央に並ぶ / right-only, so the row is centered */

        demoDialog.show();
    }

    showDemoDialog();

})();
