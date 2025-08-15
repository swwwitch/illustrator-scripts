#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
### スクリプト名 / Script Name：
PreferenceManagerForTransformAndAlign.jsx

### 概要 / Overview：
- Illustrator の主要な環境設定（整列・変形、リアルタイム編集、カーソル移動量 ほか）を1つのダイアログで安全に切り替え・保存します。
- カーソル移動量（cursorKeyLength）は **現在の定規単位（rulerType）** に合わせて表示・入力でき、保存時に pt へ変換されます。
- Provides a single dialog to toggle/save common Illustrator preferences (Transform & Align, Real-time edit, cursor step, etc.).
- `cursorKeyLength` is displayed/entered in the current ruler unit and converted to pt on save.

### 主な機能 / Main Features：
- 字形の境界に整列（ポイント文字／エリア内文字）
- 変形と整列：プレビュー境界、パターンを変形、角／線幅と効果の拡大縮小
- カーソル移動量（`cursorKeyLength`）：単位連動表示、↑↓/Shift/Option での値調整、起動時にフォーカス
- Align to glyph bounds (Point/Area Type)
- Transform & Align options: Preview Bounds, Transform Patterns, Scale Corners/Strokes & Effects
- `cursorKeyLength` with unit-aware display, arrow-key increments, and initial focus

### 更新履歴 / Change Log：
- v1.5 (20250815): `cursorKeyLength` を単位連動表示に変更、矢印キーでの増減・小数1桁表示・起動時フォーカス・UI整列を追加。
- v1.4 (20250815): キー入力フィールドのロジックを修正 / Fixed logic for key input field
- v1.3 (20250805): 単位と増減値のUIとロジックを削除 / Removed unit and increment UI and logic
- v1.2 (20250804): 角の拡大のロジックを修正 / Fixed logic for corner scaling
- v1.1 (20250804): ダイアログを2カラムに改修、単位とフォント設定を追加 / Dialog changed to two columns; added units and font settings
- v1.0 (20250804): 初期バージョン / Initial version
*/

/* スクリプトバージョン / Script Version */
var SCRIPT_VERSION = "v1.5";

/* 現在のUI言語を取得 / Get the current UI language */
function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

/* UIラベル定義 / UI Label Definitions */
var LABELS = {
    dialogTitle: {
        ja: "まとめて環境設定 " + SCRIPT_VERSION,
        en: "Preferences " + SCRIPT_VERSION
    },
    glyphBounds: {
        ja: "字形の境界に整列",
        en: "Align to Glyph Bounds"
    },
    pointText: {
        ja: "ポイント文字",
        en: "Point Type"
    },
    areaText: {
        ja: "エリア内文字",
        en: "Area Type"
    },
    transformTitle: {
        ja: "変形と整列",
        en: "Transform & Align"
    },
    keyInputLabel: {
        ja: "キー入力：",
        en: "Key:"
    },
    previewBounds: {
        ja: "プレビュー境界",
        en: "Preview Bounds"
    },
    transformPattern: {
        ja: "パターンを変形",
        en: "Transform Pattern Tiles"
    },
    scaleCorners: {
        ja: "角を拡大・縮小",
        en: "Scale Corners"
    },
    scaleStroke: {
        ja: "線幅と効果も拡大・縮小",
        en: "Scale Strokes & Effects"
    },
    realtimeDrawing: {
        ja: "リアルタイムの描画と編集",
        en: "Real-time Drawing & Editing"
    },
    ok: {
        ja: "OK",
        en: "OK"
    },
    cancel: {
        ja: "キャンセル",
        en: "Cancel"
    }
};

// 単位コードとラベルのマップ / Map of unit codes to labels
var unitLabelMap = {
    0: "in",
    1: "mm",
    2: "pt",
    3: "pica",
    4: "cm",
    5: "Q/H",
    6: "px",
    7: "ft/in",
    8: "m",
    9: "yd",
    10: "ft"
};

// 現在の単位ラベルを取得 / Get current unit label from rulerType
function getCurrentUnitLabel() {
    var unitCode = app.preferences.getIntegerPreference("rulerType");
    return unitLabelMap[unitCode] || "pt";
}

// 単位コードから pt への換算係数を取得 / Get pt factor from unit code
function getPtFactorFromUnitCode(code) {
    switch (code) {
        case 0:
            return 72.0; // in
        case 1:
            return 72.0 / 25.4; // mm
        case 2:
            return 1.0; // pt
        case 3:
            return 12.0; // pica
        case 4:
            return 72.0 / 2.54; // cm
        case 5:
            return 72.0 / 25.4 * 0.25; // Q or H
        case 6:
            return 1.0; // px
        case 7:
            return 72.0 * 12.0; // ft/in
        case 8:
            return 72.0 / 25.4 * 1000.0; // m
        case 9:
            return 72.0 * 36.0; // yd
        case 10:
            return 72.0 * 12.0; // ft
        default:
            return 1.0;
    }
}

function saveCursorKeyLengthFromField(editText) {
    var v = parseFloat(editText.text);
    if (isNaN(v) || v < 0) return;
    var unitCode = app.preferences.getIntegerPreference("rulerType");
    var ptPerUnit = getPtFactorFromUnitCode(unitCode);
    var ptVal = v * ptPerUnit;
    app.preferences.setRealPreference("cursorKeyLength", ptVal);
}

// ===== 保存のデバウンス / Debounced saving =====
var __cursorKeyDebounceTaskId = null;
var __cursorKeyPendingText = null;

function __runSaveCursorKeyLength() {
    try {
        if (__cursorKeyPendingText !== null) {
            var v = parseFloat(__cursorKeyPendingText);
            if (!isNaN(v) && v >= 0) {
                var unitCode = app.preferences.getIntegerPreference("rulerType");
                var ptPerUnit = getPtFactorFromUnitCode(unitCode);
                app.preferences.setRealPreference("cursorKeyLength", v * ptPerUnit);
            }
        }
    } catch (e) {
        // no-op on failure
    }
    __cursorKeyDebounceTaskId = null;
}

function scheduleSaveCursorKeyLengthDebounced(editText, delayMs) {
    try {
        __cursorKeyPendingText = String(editText.text);
        if (__cursorKeyDebounceTaskId) {
            app.cancelTask(__cursorKeyDebounceTaskId);
            __cursorKeyDebounceTaskId = null;
        }
        __cursorKeyDebounceTaskId = app.scheduleTask("__runSaveCursorKeyLength()", delayMs, false);
    } catch (e) {
        // Fallback: immediate save if scheduleTask unavailable
        saveCursorKeyLengthFromField(editText);
    }
}

function changeValueByArrowKey(editText) {
    editText.addEventListener("keydown", function(event) {
        var k = event.keyName;
        if (k !== "Up" && k !== "Down") {
            return; // 非矢印キーでは処理しない（Tabで0に丸められるのを防止）
        }
        var value = Number(editText.text);
        if (isNaN(value)) return;

        var keyboard = ScriptUI.environment.keyboardState;
        var delta = 1;

        if (keyboard.shiftKey) {
            delta = 10;
            // Shiftキー押下時は10の倍数にスナップ
            if (k == "Up") {
                value = Math.ceil((value + 1) / delta) * delta;
                event.preventDefault();
            } else if (k == "Down") {
                value = Math.floor((value - 1) / delta) * delta;
                if (value < 0) value = 0;
                event.preventDefault();
            }
        } else if (keyboard.altKey) {
            delta = 0.1;
            // Optionキー押下時は0.1単位で増減
            if (k == "Up") {
                value += delta;
                event.preventDefault();
            } else if (k == "Down") {
                value -= delta;
                event.preventDefault();
            }
        } else {
            delta = 1;
            if (k == "Up") {
                value += delta;
                event.preventDefault();
            } else if (k == "Down") {
                value -= delta;
                if (value < 0) value = 0;
                event.preventDefault();
            }
        }

        if (keyboard.altKey) {
            // 小数第1位までに丸め
            value = Math.round(value * 10) / 10;
        } else {
            // 整数に丸め
            value = Math.round(value);
        }

        // 常に小数第1位で表示し、即時保存
        editText.text = (Math.round(value * 10) / 10).toFixed(1);
        // 30–50ms デバウンスで保存（ここでは40ms）
        scheduleSaveCursorKeyLengthDebounced(editText, 40);
    });
}

/* 複数チェックボックスを環境設定キーにバインド / Bind multiple checkboxes to preference keys */
function bindCheckboxes(pairs) {
    for (var i = 0; i < pairs.length; i++) {
        (function(pair) {
            pair.checkbox.onClick = function() {
                app.preferences.setBooleanPreference(pair.prefKey, pair.checkbox.value === true);
            };
        })(pairs[i]);
    }
}

/* ダイアログ位置を調整 / Adjust dialog position */
function shiftDialogPosition(dlg, offsetX, offsetY) {
    dlg.onShow = function() {
        var currentX = dlg.location[0];
        var currentY = dlg.location[1];
        dlg.location = [currentX + offsetX, currentY + offsetY];
    };
}

/* ダイアログの透明度を設定 / Set dialog opacity */
function setDialogOpacity(dlg, opacityValue) {
    dlg.opacity = opacityValue;
}

/* メイン処理 / Main entry point */
function main() {

    var dialog = new Window('dialog');
    dialog.text = LABELS.dialogTitle[lang];
    dialog.orientation = 'column';
    dialog.alignChildren = ['fill', 'top'];


    /* 位置と不透明度 / Position & Opacity */
    var offsetX = 300;
    var dialogOpacity = 0.98;
    setDialogOpacity(dialog, dialogOpacity);
    shiftDialogPosition(dialog, offsetX, 0);


    var mainGroup = dialog.add('group');
    mainGroup.orientation = 'column';
    mainGroup.alignChildren = ['fill', 'top'];



    /* 字形の境界に整列パネル / Align to glyph bounds panel */
    var glyphPanel = mainGroup.add('panel', undefined, LABELS.glyphBounds[lang]);
    glyphPanel.orientation = 'column';
    glyphPanel.alignChildren = ['left', 'top'];
    glyphPanel.margins = [8, 20, 8, 15];

    var checkboxPoint = glyphPanel.add('checkbox', undefined, LABELS.pointText[lang]);
    checkboxPoint.value = app.preferences.getBooleanPreference('EnableActualPointTextSpaceAlign');

    var checkboxArea = glyphPanel.add('checkbox', undefined, LABELS.areaText[lang]);
    checkboxArea.value = app.preferences.getBooleanPreference('EnableActualAreaTextSpaceAlign');

    bindCheckboxes([{
            checkbox: checkboxPoint,
            prefKey: 'EnableActualPointTextSpaceAlign'
        },
        {
            checkbox: checkboxArea,
            prefKey: 'EnableActualAreaTextSpaceAlign'
        }
    ]);



    /* 変形と整列 / Transform & Align */
    var otherPanel = mainGroup.add('panel', undefined, LABELS.transformTitle[lang]);
    otherPanel.orientation = 'column';
    otherPanel.alignChildren = ['left', 'top'];
    otherPanel.margins = [8, 20, 8, 15];

    /* キー入力パネル / Key input panel */
    var keyInputPanel = otherPanel.add('group');
    keyInputPanel.orientation = 'row';
    keyInputPanel.alignChildren = ['left', 'center'];
    keyInputPanel.alignment = ['fill', 'center'];
    keyInputPanel.margins = 0;

    var keyHint = keyInputPanel.add('statictext', undefined, LABELS.keyInputLabel[lang]);
    var keyField = keyInputPanel.add('edittext', undefined, "");
    keyField.characters = 4;
    var unitLabel = keyInputPanel.add('statictext', undefined, getCurrentUnitLabel());

    // Initialize with current real preference for cursorKeyLength (convert pt -> current unit)
    try {
        var _pt = app.preferences.getRealPreference("cursorKeyLength");
        var _unitCode = app.preferences.getIntegerPreference("rulerType");
        var _ptPerUnit = getPtFactorFromUnitCode(_unitCode);
        keyField.text = (_pt / _ptPerUnit).toFixed(1);
    } catch (e) {
        // Fallback to 1.0 (in current unit) if preference is missing
        keyField.text = "1.0";
    }

    changeValueByArrowKey(keyField);

    dialog.onShow = function() {
        keyField.active = true;
    };

    dialog.onActivate = function() {
        try {
            // Update unit label from current rulerType
            unitLabel.text = getCurrentUnitLabel();
            // Recompute keyField from stored pt value to current unit (1 decimal place)
            var ptVal = app.preferences.getRealPreference("cursorKeyLength");
            var code = app.preferences.getIntegerPreference("rulerType");
            var factor = getPtFactorFromUnitCode(code);
            keyField.text = (ptVal / factor).toFixed(1);
        } catch (e) {
            // Keep existing values on failure
        }
    };

    // Update preference on change (real number only) — convert current unit -> pt before saving
    keyField.onChange = function() {
        var v = parseFloat(keyField.text);
        if (isNaN(v) || v < 0) {
            // Restore previous or default value on invalid input
            try {
                var _ptRestore = app.preferences.getRealPreference("cursorKeyLength");
                var _unitCodeRestore = app.preferences.getIntegerPreference("rulerType");
                var _ptPerUnitRestore = getPtFactorFromUnitCode(_unitCodeRestore);
                keyField.text = (_ptRestore / _ptPerUnitRestore).toFixed(1);
            } catch (e) {
                keyField.text = "1.0";
            }
            return;
        }
        var unitCode = app.preferences.getIntegerPreference("rulerType");
        var ptPerUnit = getPtFactorFromUnitCode(unitCode);
        var ptVal = v * ptPerUnit;
        app.preferences.setRealPreference("cursorKeyLength", ptVal);
    };

    /* プレビュー境界 / Preview bounds */
    var checkboxPreview = otherPanel.add('checkbox', undefined, LABELS.previewBounds[lang]);
    checkboxPreview.value = app.preferences.getBooleanPreference("includeStrokeInBounds");
    checkboxPreview.onClick = function() {
        app.preferences.setBooleanPreference("includeStrokeInBounds", checkboxPreview.value === true);
    };


    /* パターンを変形 / Transform patterns */
    var checkboxPattern = otherPanel.add('checkbox', undefined, LABELS.transformPattern[lang]);
    checkboxPattern.value = app.preferences.getBooleanPreference("transformPatterns");
    checkboxPattern.onClick = function() {
        app.preferences.setBooleanPreference("transformPatterns", checkboxPattern.value === true);
    };

    /* 角を拡大・縮小 / Scale corners */
    var checkboxCorner = otherPanel.add('checkbox', undefined, LABELS.scaleCorners[lang]);
    /* 初期値を取得（1=ON, 2=OFF） / Get initial value (1=ON, 2=OFF) */
    checkboxCorner.value = (app.preferences.getIntegerPreference("policyForPreservingCorners") === 1);
    checkboxCorner.onClick = function() {
        app.preferences.setIntegerPreference(
            "policyForPreservingCorners",
            checkboxCorner.value ? 1 : 2
        );
    };

    /* 線幅と効果も拡大・縮小 / Scale strokes and effects */
    var checkboxStroke = otherPanel.add('checkbox', undefined, LABELS.scaleStroke[lang]);
    checkboxStroke.value = app.preferences.getBooleanPreference("scaleLineWeight");
    checkboxStroke.onClick = function() {
        app.preferences.setBooleanPreference("scaleLineWeight", checkboxStroke.value === true);
    };

    var realtimeGroup = mainGroup.add('group', undefined, "");
    realtimeGroup.orientation = 'column';
    realtimeGroup.alignChildren = ['center', 'top'];
    realtimeGroup.margins = [3, 0, 10, 0];

    var checkboxRealtime = realtimeGroup.add('checkbox', undefined, LABELS.realtimeDrawing[lang]);
    checkboxRealtime.value = app.preferences.getBooleanPreference("LiveEdit_State_Machine");
    checkboxRealtime.onClick = function() {
        app.preferences.setBooleanPreference("LiveEdit_State_Machine", checkboxRealtime.value === true);
    };


    var group2 = dialog.add('group', undefined, {
        name: 'group2'
    });
    group2.orientation = 'row';
    group2.alignChildren = ['center', 'center']; /* ボタン行（中央） / Buttons (center) */
    group2.alignment = ['center', 'bottom']; /* ボタン行（中央） / Buttons (center) */

    var cancelBtn = group2.add('button', undefined, LABELS.cancel[lang], {
        name: 'cancel'
    });

    var okBtn = group2.add('button', undefined, LABELS.ok[lang], {
        name: 'ok'
    });

    dialog.show();
}

main();