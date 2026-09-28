#target illustrator
#targetengine "AiSmartRotateView"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

ビューの回転角度と環境設定の「角度の制限」をパレットから確認・変更し、選択オブジェクトの角度と合わせられます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AiSmartRotateViewPalette.md

### Overview

A palette for checking and changing the view rotation and the "constrain angle" preference, and for aligning them with the selected object's angle.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AiSmartRotateViewPalette.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AiSmartRotateViewPalette";            /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-06-05";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-27";                  /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AiSmartRotateViewPalette.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AiSmartRotateViewPalette.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* 同一リポジトリ内の transform フォルダのパス（外部スクリプト実行用）/ Path to the sibling transform folder (for running external scripts) */
    var TRANSFORM_FOLDER = (function () {
        try {
            return File($.fileName).parent.parent.fsName + "/transform";
        } catch (e) {
            return "";
        }
    })();

    /* 「角度の制限」をワンクリックで適用するプリセットの角度（0°は従来のリセットを兼ねる）
       / Preset angles applied to the constrain angle in one click (0° doubles as the former Reset) */
    var CONSTRAIN_PRESET_ANGLES = [-30, 0, 30];

    // =========================================
    // レイアウト / Layout
    // =========================================

    /* パネルの余白と間隔 / Panel margins and spacing */
    var PANEL_MARGINS = [16, 20, 16, 12];
    var PANEL_SPACING = 8;

    /**
     * パネルの共通設定をまとめて適用する
     * @param {Panel} targetPanel - 対象のパネル
     * @param {number} [spacing] - 子どうしの間隔（省略時は PANEL_SPACING）
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
     * グループの共通設定をまとめて適用する（row は縦中央、column は左揃え）
     * @param {Group} targetGroup - 対象のグループ
     * @param {string} [orientation] - "row" / "column"（省略時は "column"）
     * @param {number} [spacing] - 子どうしの間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function setupGroup(targetGroup, orientation, spacing) {
        var groupOrientation = orientation || "column";
        targetGroup.orientation = groupOrientation;
        /* row は横並びなので縦中央、column は縦並びなので左揃え / row: vertically centered, column: left-aligned */
        targetGroup.alignChildren = (groupOrientation === "row") ? ["left", "center"] : ["left", "top"];
        targetGroup.alignment = "fill";
        targetGroup.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    // =========================================
    // セッション記憶 / Session memory
    // #targetengine で確保した永続エンジンの $.global に置くため、パレットを閉じて開き直しても
    // Illustrator を終了するまで保持される（環境設定には書き込まない）
    // Kept on $.global of the persistent engine declared by #targetengine, so it survives closing and
    // reopening the palette until Illustrator quits. Nothing is written to the application preferences.
    // =========================================

    /* 回転角度と制限角度はドキュメントと環境設定から取り直せるので、再現できない項目だけを残す
       / The rotation and constrain angles are re-read from the document and preferences, so only the values that cannot be recovered are stored */
    var SESSION_STATE_KEY = "__aiSmartRotateViewSession";
    if (typeof $.global[SESSION_STATE_KEY] === "undefined") {
        $.global[SESSION_STATE_KEY] = {
            linkRotation: false, /* 「ビューの回転に連動」の状態 / the "Link to view rotation" state */
            location: null       /* パレットの表示位置 / the palette position */
        };
    }
    var sessionState = $.global[SESSION_STATE_KEY];

    /* 開いているパレット自体の置き場所（多重起動の防止とGC回避）/ Where the open palette itself is kept (prevents duplicates and GC) */
    var PALETTE_KEY = "__aiSmartRotateViewPalette";

    // =========================================
    // ローカライズ / Localization
    // =========================================

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // ローカライズ（再利用パーツ） / Localization (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内のローカライズ節（LABELS の直前）に貼る。
    //    uiLang を使うコード（StepperButtons・LinkToggle の部品など）より前に置く
    // 2. 識別子は uiLang / getCurrentLang / getLabel / labelText / labelValueText / fillLabelPlaceholders。
    //    同じ役割の既存の関数・変数（getCurrentLanguage、currentLanguage、formatLabel など）は消して、これに寄せる
    // 3. 呼び出しはどちらの形でもよい（混ぜてもよい）
    //      getLabel("dialog.title")        … パス
    //      getLabel(LABELS.dialog.title)   … { ja, en } を直接
    //      getLabel("alert.count", { count: 3 })  … "{count} 個" の {count} を差し込む
    //      getLabel("alert.range", [1, 10])       … "%1〜%2" の %1・%2 を差し込む
    //      labelText("fieldLabel.width")   … 末尾にコロン（日本語は全角「：」、英語は半角「:」）
    //      labelValueText("message.count", 5) … 「件数：5」／「Count: 5」（値が続く1行。英語はコロンのあとに空白）
    // 4. 見つからないパスはパスの文字列をそのまま返す（表示で気づけるように）。{ ja, en } が無いときは空文字
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

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

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // ローカライズ（再利用パーツ）ここまで / End of the reusable localization
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    /* ラベル定義 / Label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "ビューの回転角度", en: "View Rotation Angle" }
        },
        panel: {
            viewRotation: { ja: "アクティブビューの回転角度", en: "Active View Rotation Angle" },
            selection: { ja: "選択したオブジェクト", en: "Selected Object" },
            reset: { ja: "リセット", en: "Reset" },
            constrain: { ja: "角度の制限", en: "Constrain Angle" }
        },
        fieldLabel: {
            selectionAngle: { ja: "選択したオブジェクトの角度", en: "Selected object angle" },
            constrain: { ja: "角度の制限", en: "Constrain angle" }
        },
        checkbox: {
            linkRotation: { ja: "ビューの回転に連動", en: "Link to view rotation" }
        },
        button: {
            rotateToSelection: { ja: "選択したオブジェクトに合わせてビューを回転", en: "Rotate view to match selection" },
            rotateSelectionToView: { ja: "選択したオブジェクトをビューの回転に合わせて回転", en: "Rotate selection to match view" },
            resetSelectionRotation: { ja: "選択したオブジェクトの回転をリセット", en: "Reset selection rotation" },
            resetTextTilt: { ja: "テキストの傾きをリセット", en: "Reset text tilt" },
            resetImageTilt: { ja: "画像の傾きをリセット", en: "Reset image tilt" },
            resetRotation: { ja: "リセット", en: "Reset" },
            refresh: { ja: "更新", en: "Refresh" }
        },
        tooltip: {
            rotationSlider: {
                ja: "アクティブビューの回転角度です。Shiftを押しながらドラッグすると15°単位になります。",
                en: "Rotation angle of the active view. Hold Shift while dragging to snap to 15°."
            },
            constrainInput: {
                ja: "環境設定の「角度の制限」です。入力を確定すると、その場で適用します。",
                en: "The Constrain Angle preference. Committing the field applies it right away."
            },
            constrainSlider: {
                ja: "角度の制限をドラッグで変えます。Shiftを押しながらで15°単位になります。",
                en: "Drags the constrain angle. Hold Shift to snap to 15°."
            },
            constrainPreset: { ja: "この角度を「角度の制限」に適用します。", en: "Applies this angle to the Constrain Angle preference." },
            linkRotation: { ja: "ビューを回転したとき、角度の制限も同じ値に合わせます。", en: "Keeps the constrain angle in step with the view rotation." },
            resetRotation: { ja: "ビューの回転角度を0°に戻します。", en: "Resets the view rotation to 0°." },
            rotateToSelection: {
                ja: "選択したパスの最初の2つのアンカーポイントを結ぶ線の角度に合わせて、ビューを回転します。",
                en: "Rotates the view to the angle of the line through the first two anchor points of the selected path."
            },
            rotateSelectionToView: {
                ja: "選択したオブジェクトをビューの回転角度だけまとめて回転し、バウンディングボックスをリセットします。",
                en: "Rotates the selected objects together by the view rotation angle, then resets the bounding box."
            },
            resetSelectionRotation: {
                ja: "最前面の選択オブジェクトに記録された回転（BBAccumRotation）を打ち消して、元の向きに戻します。",
                en: "Undoes the rotation recorded on the frontmost selected object (BBAccumRotation), restoring its original orientation."
            },
            resetTextTilt: { ja: "transform フォルダーの ResetTransform.jsx を実行します。", en: "Runs ResetTransform.jsx in the transform folder." },
            resetImageTilt: { ja: "transform フォルダーの ResetRotation.jsx を実行します。", en: "Runs ResetRotation.jsx in the transform folder." },
            refresh: {
                ja: "ビューの回転角度・角度の制限・選択したオブジェクトの角度を読み直します。",
                en: "Re-reads the view rotation, the constrain angle, and the selected object's angle."
            },
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
        status: {
            applied: { ja: "制限角度に適用しました。", en: "Applied to the constrain angle." },
            resetRotation: { ja: "ビューの回転を0°にリセットしました。", en: "Reset the view rotation to 0°." },
            rotatedToSelection: { ja: "選択したオブジェクトに合わせてビューを回転しました。", en: "Rotated the view to match the selection." },
            rotatedSelectionToView: { ja: "選択したオブジェクトをビューの回転に合わせて回転しました。", en: "Rotated the selection to match the view." },
            resetSelectionRotation: { ja: "選択したオブジェクトの回転をリセットしました。", en: "Reset the selection's rotation." }
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noSelection: { ja: "オブジェクトが選択されていません。", en: "No object is selected." },
            noAngle: { ja: "角度を取得できるパスを選択してください。", en: "Select a path whose angle can be measured." },
            noRotation: { ja: "ビューが回転していません。", en: "The view is not rotated." },
            noTag: { ja: "回転情報（BBAccumRotation）が見つかりません。", en: "Rotation data (BBAccumRotation) not found." },
            noFile: { ja: "スクリプトファイルが見つかりません。", en: "Script file not found." },
            invalidAngle: { ja: "角度には数値を入力してください。", en: "Please enter a numeric angle." },
            error: { ja: "エラーが発生しました：", en: "An error occurred:" }
        }
    };

    // =========================================
    // 結果ハンドリング / Result handling
    // =========================================

    /* ワーカーの 'ERR:XXX' コードを対応するアラート文へ変換 / Map a worker 'ERR:XXX' code to its alert text */
    var RESULT_ERROR_LABELS = {
        NODOC: "alert.noDocument",
        NOSEL: "alert.noSelection",
        NOANGLE: "alert.noAngle",
        NOROTATION: "alert.noRotation",
        NOTAG: "alert.noTag",
        NOFILE: "alert.noFile"
    };

    /**
     * ワーカーの結果文字列を表示用のステータス文にする（既知のコードは専用文、未知のものは汎用エラー）
     * @param {string} result - ワーカーから返った結果文字列
     * @returns {string} ステータス文
     */
    function statusFromResult(result) {
        for (var errorCode in RESULT_ERROR_LABELS) {
            if (RESULT_ERROR_LABELS.hasOwnProperty(errorCode) && result.indexOf(errorCode) !== -1) {
                return getLabel(RESULT_ERROR_LABELS[errorCode]);
            }
        }
        return getLabel("alert.error") + " " + result;
    }

    // =========================================
    // 入力の補助 / Input helpers
    // =========================================

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // UI の明暗（再利用パーツ） / UI theme (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内に貼る（StepperButtons・LinkToggle の部品より前）。識別子は isDarkUI
    // 2. 配色を明暗で切り替えるときは isDarkUI() を1回だけ呼んで定数に控える
    //      var MY_UI_DARK = isDarkUI();
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

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

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // UI の明暗（再利用パーツ）ここまで / End of the reusable UI theme
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // ステップボタン（再利用パーツ） / Stepper buttons (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内（ローカライズより前）に貼る。
    //    識別子はすべて STEPPER_* / *Stepper* / *Stepped* の名前なので、既存の名前とはぶつからない
    //    UI の明暗は UITheme 部品の isDarkUI() を使う（先に UITheme の ▼〜▲ も貼っておく）
    // 2. コピー先の LABELS.tooltip に stepUp / stepDown / stepUpInteger / stepDownInteger を足す（このファイルの LABELS から写す）。
    //    getLabel() と uiLang はコピー先のものをそのまま使う
    // 3. 数値欄を addSteppedField() で作る。項目名・∧∨・入力欄がひと組で入り、↑↓キーも∧∨と同じ処理で増減する
    //      var widthInput = addSteppedField(parentPanel, {
    //          label: labelText(LABELS.fieldLabel.width), labelWidth: 60,
    //          text: "210 mm", characters: 8, step: 1, min: 1, unit: " mm",
    //          onStep: function (numberInput) { updatePreview(); }
    //      });
    //    値の種類は options で切り分ける:
    //      小数あり（幅・位置など）   … 指定なし（option＋クリックで0.1ずつ）
    //      整数・1以上（段数・個数など）… integer: true, min: 1（0・小数・負数は受け付けず、option＋クリックも1ずつ）
    //      整数・0以上（間隔の数など）  … integer: true, min: 0
    //      範囲つき（％など）           … min: 0, max: 100, unit: "%"
    // 4. 有効／無効は setSteppedFieldEnabled(widthInput, isEnabled)（∧∨のディム表示も切り替わる）。
    //    行・パネルなど親の enabled を切り替えたときは、そのあとで redrawSteppersIn(親) を呼んで∧∨を描き直す
    //    （∧∨は親をたどって無効を判定し、無効の間はクリックも↑↓キーも効かない）
    // 5. 値は parseFloat(widthInput.text) で読む（unit 付きの欄は「210 mm」の形で入っている）
    // 6. この欄に別の↑↓キー処理を付けない（↑↓キーが二重に効く）
    // 既存の edittext をそのまま使うときは、同じ行の group（spacing 0）に addStepper() → edittext の順で置き、
    // bindSteppedArrowKeys(edittext, stepperGroup) を呼ぶ
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

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
     * 整数化・下限・上限・単位（「20 mm」の形）へそろえ、数値でなければ直前の値に戻す
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
            var value = parseFloat(numberInput.text);
            if (isNaN(value)) value = 0;
            writeSteppedValue(numberInput, computeSteppedValue(value, direction, stepOptions), stepOptions);
            if (stepOptions.onStep) stepOptions.onStep(numberInput);
        }

        /* 整数の欄では option＋クリックの0.1刻みが効かないので、説明から外す / integer fields have no 0.1 step */
        var upTooltip = stepOptions.integer ? LABELS.tooltip.stepUpInteger : LABELS.tooltip.stepUp;
        var downTooltip = stepOptions.integer ? LABELS.tooltip.stepDownInteger : LABELS.tooltip.stepDown;
        makeStepperChevronButton(stepperGroup, "up", function () { stepBy(1); }).helpTip = getLabel(upTooltip);
        makeStepperChevronButton(stepperGroup, "down", function () { stepBy(-1); }).helpTip = getLabel(downTooltip);
        stepperGroup.stepBy = stepBy; /* ↑↓キーからも同じ処理で増減できるよう公開 / shared with the arrow keys */
        return stepperGroup;
    }

    /**
     * 入力欄の↑↓キーを、∧∨と同じ処理で増減させる。ほかのキーは素通し
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

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // ステップボタン（再利用パーツ）ここまで / End of the reusable stepper
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    // =========================================
    // 角度の計算 / Angle helpers
    // =========================================

    /**
     * 角度を -180〜180 に正規化する
     * @param {number} angle - 角度（度）
     * @returns {number} 正規化した角度
     */
    function normalizeAngle(angle) {
        var normalized = angle % 360;
        if (normalized > 180) { normalized -= 360; }
        if (normalized < -180) { normalized += 360; }
        return normalized;
    }

    /**
     * スライダー値を整数に丸める。Shift 併用時は15°単位にする
     * @param {number} angle - スライダーの値
     * @param {boolean} useShift - Shift を押しているか
     * @returns {number} 丸めた角度
     */
    function snapAngle(angle, useShift) {
        if (useShift) {
            return Math.round(angle / 15) * 15;
        }
        return Math.round(angle);
    }

    /**
     * 確定したスライダーの角度を読む（Shift を押していれば15°単位）
     * @param {Slider} targetSlider - 対象のスライダー
     * @returns {number} 丸めた角度
     */
    function readSnappedSliderAngle(targetSlider) {
        return snapAngle(targetSlider.value, ScriptUI.environment.keyboardState.shiftKey);
    }

    /**
     * ドラッグ中のスライダーの角度を読む。Shift を押しているときはつまみも15°単位の位置へ寄せる
     * @param {Slider} targetSlider - 対象のスライダー
     * @returns {number} 丸めた角度
     */
    function snapSliderWhileDragging(targetSlider) {
        var useShift = ScriptUI.environment.keyboardState.shiftKey;
        var angle = snapAngle(targetSlider.value, useShift);
        if (useShift) { targetSlider.value = angle; }
        return angle;
    }

    /**
     * 表示用に小数2桁へ丸める（atan2 由来の 30.00000001 のような桁あふれを抑える）
     * @param {number} angle - 角度（度）
     * @returns {number} 丸めた角度
     */
    function roundAngle(angle) {
        return Math.round(angle * 100) / 100;
    }

    /**
     * 角度を表示用の文字列にする（丸めて°を付ける）
     * @param {number} angle - 角度（度）
     * @returns {string} 表示用の文字列
     */
    function formatAngle(angle) {
        return roundAngle(angle) + "°";
    }

    // =========================================
    // メインエンジンへの委譲 / Delegation to the main engine
    // =========================================

    /**
     * メインエンジンでコードを実行する（常駐パレットの app は DOM 接続を失うため）。
     * 本文は encodeURIComponent + eval で送り、バックスラッシュ・多バイト文字を無傷で渡す
     * @param {string} code - 実行するコード
     * @param {Function} onResult - 結果文字列を受け取るコールバック（失敗時は "ERR:" で始まる）
     * @returns {void}
     */
    function runInMainEngine(code, onResult) {
        var bridgeTalk = new BridgeTalk();
        bridgeTalk.target = "illustrator";
        bridgeTalk.body = 'eval(decodeURIComponent("' + encodeURIComponent(code) + '"));';
        bridgeTalk.onResult = function (response) {
            onResult(String(response.body));
        };
        bridgeTalk.onError = function (response) {
            onResult("ERR:" + String(response.body));
        };
        bridgeTalk.send();
    }

    /**
     * ワーカーの本文を IIFE で包む
     * @param {string} workerCode - ワーカーの本文
     * @returns {string} IIFE で包んだコード
     */
    function wrapWorkerBody(workerCode) {
        return "(function(){" + workerCode + "})()";
    }

    /* ワーカー断片：ドキュメント有無を確認し doc を確保 / Worker fragment: guard the open document, bind `doc` */
    var WORKER_REQUIRE_DOCUMENT = "if(app.documents.length===0){return 'ERR:NODOC';}var doc=app.activeDocument;";

    /* ワーカー断片：選択を確認し selectedItems を確保（WORKER_REQUIRE_DOCUMENT の後に置く）
       / Worker fragment: guard the selection, bind `selectedItems` (place after WORKER_REQUIRE_DOCUMENT) */
    var WORKER_REQUIRE_SELECTION = "var selectedItems=doc.selection;if(!selectedItems||selectedItems.length===0){return 'ERR:NOSEL';}";

    /* ワーカー断片：パスの最初の2点を結ぶ線の傾きを返す measurePathAngle() を定義（2点以上のパスでなければ null）
       / Worker fragment: define measurePathAngle(), the slope of the line through a path's first two anchors (null unless it is a path with >=2 points) */
    var WORKER_DEFINE_PATH_ANGLE =
        "function measurePathAngle(pathItem){" +
        "if(pathItem.typename!=='PathItem'||pathItem.pathPoints.length<2){return null;}" +
        "var startAnchor=pathItem.pathPoints[0].anchor,endAnchor=pathItem.pathPoints[1].anchor;" +
        "return Math.atan2(endAnchor[1]-startAnchor[1],endAnchor[0]-startAnchor[0])*180/Math.PI;}";

    /* ワーカー断片：selectedItems[0] の角度を pathAngle に入れる。測れなければエラーで返す（WORKER_REQUIRE_SELECTION の後）
       / Worker fragment: put the angle of selectedItems[0] in `pathAngle`, returning an error when it cannot be measured (place after WORKER_REQUIRE_SELECTION) */
    var WORKER_SELECTION_PATH_ANGLE =
        WORKER_DEFINE_PATH_ANGLE +
        "var pathAngle=measurePathAngle(selectedItems[0]);" +
        "if(pathAngle===null){return 'ERR:NOANGLE';}";

    /* ワーカー断片：「角度の制限」を読む。実際の拘束方向は constrain/sin・constrain/cos が持っているため、
       constrain/angle ではなくこの2つから角度を復元する（angle は書いても拘束に反映されない）
       / Worker fragment: read the constrain angle. The real constraint direction lives in constrain/sin and
       constrain/cos, so recover the angle from those instead of constrain/angle (writing `angle` alone has no effect) */
    var WORKER_READ_CONSTRAIN =
        "var constrainAngle=Math.atan2(app.preferences.getRealPreference('constrain/sin')," +
        "app.preferences.getRealPreference('constrain/cos'))*180/Math.PI;";

    /**
     * 「角度の制限」を書き込むワーカー断片を作る。angle は環境設定ダイアログの表示用で度、
     * sin・cos は実際の拘束方向でラジアン由来。angle だけでは拘束に効かないので3つとも書く
     * @param {number} angle - 書き込む角度（度）
     * @returns {string} ワーカー断片
     */
    function buildConstrainWriteCode(angle) {
        return "var radians=(" + angle + ")*Math.PI/180;" +
            "app.preferences.setRealPreference('constrain/angle'," + angle + ");" +
            "app.preferences.setRealPreference('constrain/sin',Math.sin(radians));" +
            "app.preferences.setRealPreference('constrain/cos',Math.cos(radians));";
    }

    /**
     * ビューの回転角度・制限角度・選択角度を1回の委譲でまとめて取得する（"OK:回転,制限,選択"。選択は測れなければ空）。
     * 状態取得を1本化することで、別々に投げていた頃の表示ズレを防ぐ
     * @param {Function} onResult - 結果文字列を受け取るコールバック
     * @returns {void}
     */
    function fetchState(onResult) {
        runInMainEngine(wrapWorkerBody(
            WORKER_REQUIRE_DOCUMENT +
            "var viewRotation=doc.activeView.rotateAngle;" +
            WORKER_READ_CONSTRAIN +
            WORKER_DEFINE_PATH_ANGLE +
            "var selectionAngle='';" +
            "var selectedItems=doc.selection;" +
            "if(selectedItems&&selectedItems.length>0){var pathAngle=measurePathAngle(selectedItems[0]);" +
            "if(pathAngle!==null){selectionAngle=pathAngle;}}" +
            "return 'OK:'+viewRotation+','+constrainAngle+','+selectionAngle;"
        ), onResult);
    }

    /**
     * ビューの回転角度を設定する（制限角度は自動では変えず、適用の操作をしたときだけ反映）
     * @param {number} angle - 回転角度（度）
     * @param {Function} onResult - 結果文字列を受け取るコールバック
     * @returns {void}
     */
    function applyViewRotation(angle, onResult) {
        runInMainEngine(wrapWorkerBody(
            WORKER_REQUIRE_DOCUMENT +
            "doc.activeView.rotateAngle=" + angle + ";" +
            "return 'OK';"
        ), onResult);
    }

    /**
     * 制限角度をメインエンジンで環境設定に適用する
     * @param {number} angle - 制限角度（度）
     * @param {Function} onResult - 結果文字列を受け取るコールバック
     * @returns {void}
     */
    function applyConstrainAngle(angle, onResult) {
        runInMainEngine(wrapWorkerBody(
            buildConstrainWriteCode(angle) +
            "return 'OK';"
        ), onResult);
    }

    /**
     * ビューの回転角度だけ0°に戻す
     * @param {Function} onResult - 結果文字列を受け取るコールバック
     * @returns {void}
     */
    function resetViewRotation(onResult) {
        runInMainEngine(wrapWorkerBody(
            WORKER_REQUIRE_DOCUMENT +
            "doc.activeView.rotateAngle=0;" +
            "return 'OK';"
        ), onResult);
    }

    /**
     * 選択パスの角度（最初の2アンカー点を結ぶ線の傾き）を測り、ビューをその角度に回転する（"OK:角度"）
     * @param {Function} onResult - 結果文字列を受け取るコールバック
     * @returns {void}
     */
    function rotateViewToSelection(onResult) {
        runInMainEngine(wrapWorkerBody(
            WORKER_REQUIRE_DOCUMENT + WORKER_REQUIRE_SELECTION + WORKER_SELECTION_PATH_ANGLE +
            "doc.activeView.rotateAngle=pathAngle;" +
            "return 'OK:'+pathAngle;"
        ), onResult);
    }

    /**
     * 選択オブジェクトをビューの回転角度だけ回転し、バウンディングボックスをリセットする。
     * 複数選択のときは選択全体の外接矩形の中心を軸にまとめて回す（1つずつその場で回すと配置が崩れるため）。
     * 基準ボックスは環境設定「プレビュー境界を使用」に合わせ、回転角度は UI の表示ではなく実行時のビューから読む
     * @param {Function} onResult - 結果文字列を受け取るコールバック
     * @returns {void}
     */
    function rotateSelectionToView(onResult) {
        runInMainEngine(wrapWorkerBody(
            WORKER_REQUIRE_DOCUMENT + WORKER_REQUIRE_SELECTION +
            "var viewAngle=doc.activeView.rotateAngle;" +
            "if(viewAngle===0){return 'ERR:NOROTATION';}" +
            "var usePreviewBounds=app.preferences.getBooleanPreference('includeStrokeInBounds');" +
            "function getItemBounds(targetItem){return usePreviewBounds?targetItem.visibleBounds:targetItem.geometricBounds;}" +
            /* 選択全体の外接矩形（[左,上,右,下]、上下はY軸が上向き）/ Combined bounds of the selection ([l,t,r,b] with Y pointing up) */
            "var unionBounds=getItemBounds(selectedItems[0]).slice(0);" +
            "for(var i=1;i<selectedItems.length;i++){var itemBounds=getItemBounds(selectedItems[i]);" +
            "if(itemBounds[0]<unionBounds[0]){unionBounds[0]=itemBounds[0];}if(itemBounds[1]>unionBounds[1]){unionBounds[1]=itemBounds[1];}" +
            "if(itemBounds[2]>unionBounds[2]){unionBounds[2]=itemBounds[2];}if(itemBounds[3]<unionBounds[3]){unionBounds[3]=itemBounds[3];}}" +
            "var pivotX=(unionBounds[0]+unionBounds[2])/2,pivotY=(unionBounds[1]+unionBounds[3])/2;" +
            "var radians=viewAngle*Math.PI/180,cosAngle=Math.cos(radians),sinAngle=Math.sin(radians);" +
            /* 各オブジェクトを自身の中心で回し、その中心が共通の軸まわりに動くぶんだけ平行移動して補正
               / Rotate each object about its own center, then translate it by however far that center moves around the shared pivot */
            "for(var j=0;j<selectedItems.length;j++){var rotatedItem=selectedItems[j];var rotatedBounds=getItemBounds(rotatedItem);" +
            "var centerX=(rotatedBounds[0]+rotatedBounds[2])/2,centerY=(rotatedBounds[1]+rotatedBounds[3])/2;" +
            "rotatedItem.rotate(viewAngle,true,true,true,true,Transformation.CENTER);" +
            "var dx=centerX-pivotX,dy=centerY-pivotY;" +
            "rotatedItem.translate(pivotX+dx*cosAngle-dy*sinAngle-centerX,pivotY+dx*sinAngle+dy*cosAngle-centerY);}" +
            "app.executeMenuCommand('AI Reset Bounding Box');" +
            "return 'OK';"
        ), onResult);
    }

    /**
     * 指定したスクリプトファイルをメインエンジンで実行する（各スクリプトが自前で選択チェック・アラートを行う）
     * @param {string} scriptPath - スクリプトファイルのパス
     * @param {Function} onResult - 結果文字列を受け取るコールバック
     * @returns {void}
     */
    function runScriptFile(scriptPath, onResult) {
        runInMainEngine(wrapWorkerBody(
            "var scriptFile=new File('" + scriptPath + "');" +
            "if(!scriptFile.exists){return 'ERR:NOFILE';}" +
            "$.evalFile(scriptFile);" +
            "return 'OK';"
        ), onResult);
    }

    /**
     * 選択オブジェクトの回転をリセットする（BBAccumRotation タグの蓄積回転を戻す）
     * @param {Function} onResult - 結果文字列を受け取るコールバック
     * @returns {void}
     */
    function resetSelectionRotation(onResult) {
        runInMainEngine(wrapWorkerBody(
            WORKER_REQUIRE_DOCUMENT + WORKER_REQUIRE_SELECTION +
            "var targetItem=selectedItems[0];" +
            "if(targetItem.tags.length>0&&targetItem.tags[0].name==='BBAccumRotation'){" +
            "var accumDegrees=180*parseFloat(targetItem.tags[0].value)/Math.PI;" +
            "targetItem.rotate(accumDegrees);" +
            "return 'OK';" +
            "}" +
            "return 'ERR:NOTAG';"
        ), onResult);
    }

    // =========================================
    // パレット / Palette
    // =========================================

    /**
     * パレットと各コントロールを作る
     * @returns {Object} パレット（palette）と各コントロールの参照
     */
    function buildPalette() {
        var rotateViewPalette = new Window("palette", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        rotateViewPalette.orientation = "column";
        rotateViewPalette.alignChildren = "fill";
        rotateViewPalette.margins = 16;
        rotateViewPalette.spacing = 12;

        /* ビューの回転角度を表示するパネル / Panel showing the view rotation angle */
        var viewRotationPanel = rotateViewPalette.add("panel", undefined, getLabel("panel.viewRotation"));
        setupPanel(viewRotationPanel, 6);

        /* アクティブビューの回転角度（パネルタイトルと重複するため内側ラベルは省略）/ Active view rotation angle (inner label omitted; the panel title already states it) */
        var rotationValueGroup = viewRotationPanel.add("group");
        setupGroup(rotationValueGroup, "row");
        var rotationValue = rotationValueGroup.add("statictext", undefined, "—°");
        rotationValue.preferredSize.width = 50;

        /* 回転角度スライダー（-180〜180、Shiftで15°単位にクランプ）/ Rotation slider (-180..180, snaps to 15° with Shift) */
        var rotationSlider = viewRotationPanel.add("slider", undefined, 0, -180, 180);
        rotationSlider.helpTip = getLabel("tooltip.rotationSlider");
        rotationSlider.alignment = "fill";

        /* ビューの回転だけ0°に戻す / Reset only the view rotation to 0° */
        var resetRotationButton = viewRotationPanel.add("button", undefined, getLabel("button.resetRotation"));
        resetRotationButton.helpTip = getLabel("tooltip.resetRotation");
        resetRotationButton.alignment = "right";

        /* 角度の制限を変更するパネル / Panel for changing the constrain angle */
        var constrainPanel = rotateViewPalette.add("panel", undefined, getLabel("panel.constrain"));
        setupPanel(constrainPanel, 6);

        /* 角度の制限（編集可。入力の確定・↑↓キー・スライダー・プリセットのいずれでもその場で環境設定へ適用）
           / Constrain angle (editable; committing the field, the arrow keys, the slider, and the presets all apply it to the preference right away) */
        var constrainInputGroup = constrainPanel.add("group");
        setupGroup(constrainInputGroup, "row");
        constrainInputGroup.add("statictext", undefined, labelText("fieldLabel.constrain"));
        /* ∧∨と入力欄は隙間0で突き合わせる / butt the stepper against the field */
        var constrainStepperGroup = constrainInputGroup.add("group");
        constrainStepperGroup.orientation = "row";
        constrainStepperGroup.alignChildren = ["left", "center"];
        constrainStepperGroup.spacing = 0;
        constrainStepperGroup.margins = 0;
        var constrainInput;
        var constrainStepper = addStepper(constrainStepperGroup, function () { return constrainInput; }, {
            min: -180, max: 180,
            /* text の代入では onChange が起きないので、押すたびに適用されるよう自前で通知
               / Assigning to text does not fire onChange, so notify manually to apply on every step */
            onStep: function (numberInput) { numberInput.notify("onChange"); }
        });
        constrainInput = constrainStepperGroup.add("edittext", undefined, "");
        constrainInput.characters = 6;
        constrainInput.helpTip = getLabel("tooltip.constrainInput");
        bindSteppedArrowKeys(constrainInput, constrainStepper);
        constrainInputGroup.add("statictext", undefined, "°");

        /* 角度の制限スライダー（-180〜180、Shiftで15°単位にクランプ。離した時点で環境設定へ適用）
           / Constrain angle slider (-180..180, snaps to 15° with Shift; applied to the preference on release) */
        var constrainSlider = constrainPanel.add("slider", undefined, 0, -180, 180);
        constrainSlider.helpTip = getLabel("tooltip.constrainSlider");
        constrainSlider.alignment = "fill";

        /* よく使う角度をワンクリックで適用するプリセットボタン / Preset buttons that apply a common angle in one click */
        var constrainPresetGroup = constrainPanel.add("group");
        setupGroup(constrainPresetGroup, "row");
        constrainPresetGroup.alignment = "left";

        /* プリセットの角度とボタンの対（現在値と同じものをディムするために保持）/ Preset angle-button pairs (kept so the one matching the current value can be dimmed) */
        var constrainPresets = [];
        for (var i = 0; i < CONSTRAIN_PRESET_ANGLES.length; i++) {
            var presetButton = constrainPresetGroup.add("button", undefined, CONSTRAIN_PRESET_ANGLES[i] + "°");
            presetButton.preferredSize.width = 56;
            presetButton.helpTip = getLabel("tooltip.constrainPreset");
            constrainPresets.push({ angle: CONSTRAIN_PRESET_ANGLES[i], button: presetButton });
        }

        /* ビューの回転に連動するかどうか（ONで回転角度と同じ値、OFFでは現在の制限角度のまま）
           / Whether to follow the view rotation (on: the same value as the rotation; off: leaves the current constrain angle alone) */
        var linkRotationCheckbox = constrainPanel.add("checkbox", undefined, getLabel("checkbox.linkRotation"));
        linkRotationCheckbox.helpTip = getLabel("tooltip.linkRotation");
        linkRotationCheckbox.alignment = "left";
        /* 前回の状態を復元（Illustrator のセッション中のみ）/ Restore the previous state (only within the Illustrator session) */
        linkRotationCheckbox.value = sessionState.linkRotation;

        /* 選択したオブジェクトのパネル / Selected-object panel */
        var selectionPanel = rotateViewPalette.add("panel", undefined, getLabel("panel.selection"));
        setupPanel(selectionPanel, 6);

        /* 選択したオブジェクトの角度（表示）/ Selected object angle (display) */
        var selectionAngleGroup = selectionPanel.add("group");
        setupGroup(selectionAngleGroup, "row");
        selectionAngleGroup.add("statictext", undefined, labelText("fieldLabel.selectionAngle"));
        var selectionAngleValue = selectionAngleGroup.add("statictext", undefined, "—°");
        selectionAngleValue.preferredSize.width = 50;

        /* 選択に合わせてビューを回転 / Rotate the view to match the selection */
        var rotateViewButton = selectionPanel.add("button", undefined, getLabel("button.rotateToSelection"));
        rotateViewButton.helpTip = getLabel("tooltip.rotateToSelection");
        rotateViewButton.alignment = "left";

        /* 選択をビューの回転に合わせて回転 / Rotate the selection to match the view rotation */
        var rotateSelectionButton = selectionPanel.add("button", undefined, getLabel("button.rotateSelectionToView"));
        rotateSelectionButton.helpTip = getLabel("tooltip.rotateSelectionToView");
        rotateSelectionButton.alignment = "left";

        /* 選択の回転をリセット / Reset the selection's rotation */
        var resetSelectionButton = selectionPanel.add("button", undefined, getLabel("button.resetSelectionRotation"));
        resetSelectionButton.helpTip = getLabel("tooltip.resetSelectionRotation");
        resetSelectionButton.alignment = "left";

        /* リセットパネル（外部スクリプトを実行）/ Reset panel (runs external scripts) */
        var resetPanel = rotateViewPalette.add("panel", undefined, getLabel("panel.reset"));
        setupPanel(resetPanel, 6);

        /* テキストの傾き（ResetTransform.jsx）/ Text tilt (ResetTransform.jsx) */
        var resetTextButton = resetPanel.add("button", undefined, getLabel("button.resetTextTilt"));
        resetTextButton.helpTip = getLabel("tooltip.resetTextTilt");
        resetTextButton.alignment = "left";

        /* 画像の傾き（ResetRotation.jsx）/ Image tilt (ResetRotation.jsx) */
        var resetImageButton = resetPanel.add("button", undefined, getLabel("button.resetImageTilt"));
        resetImageButton.helpTip = getLabel("tooltip.resetImageTilt");
        resetImageButton.alignment = "left";

        /* 最新の状態を手動で取得し直す / Manually re-fetch the latest state */
        var refreshButton = rotateViewPalette.add("button", undefined, getLabel("button.refresh"));
        refreshButton.helpTip = getLabel("tooltip.refresh");
        refreshButton.alignment = "right";

        /* ステータス表示 / Status line */
        var statusText = rotateViewPalette.add("statictext", undefined, "");
        statusText.alignment = "fill";

        return {
            palette: rotateViewPalette,
            rotationValue: rotationValue,
            rotationSlider: rotationSlider,
            resetRotationButton: resetRotationButton,
            constrainInput: constrainInput,
            constrainSlider: constrainSlider,
            constrainPresets: constrainPresets,
            linkRotationCheckbox: linkRotationCheckbox,
            selectionAngleValue: selectionAngleValue,
            rotateViewButton: rotateViewButton,
            rotateSelectionButton: rotateSelectionButton,
            resetSelectionButton: resetSelectionButton,
            resetTextButton: resetTextButton,
            resetImageButton: resetImageButton,
            refreshButton: refreshButton,
            statusText: statusText
        };
    }

    /**
     * パレットの表示更新と各コントロールの操作を結び付ける
     * @param {Object} paletteUI - buildPalette() が返したパレットとコントロールの参照
     * @returns {void}
     */
    function bindPaletteHandlers(paletteUI) {
        /* 現在の状態（各リセットボタンのディム判定に使用）/ Current state (used to dim each Reset button) */
        var currentRotation = 0;
        var currentConstrain = 0;

        /**
         * 押しても値が変わらないボタンをディムする。ビューの回転が0°ならリセットと「ビューの回転に合わせて回転」、
         * 制限角度と同じ値のプリセットもディムして、いま効いている角度が分かるようにする
         * @returns {void}
         */
        function updateButtonState() {
            paletteUI.resetRotationButton.enabled = (currentRotation !== 0);
            paletteUI.rotateSelectionButton.enabled = (currentRotation !== 0);
            for (var i = 0; i < paletteUI.constrainPresets.length; i++) {
                paletteUI.constrainPresets[i].button.enabled = (paletteUI.constrainPresets[i].angle !== currentConstrain);
            }
        }

        /**
         * 制限角度を入力欄とスライダーの両方に表示する
         * @param {number} angle - 制限角度（度）
         * @returns {void}
         */
        function showConstrain(angle) {
            paletteUI.constrainInput.text = roundAngle(angle);
            paletteUI.constrainSlider.value = normalizeAngle(angle);
        }

        /**
         * 制限の入力欄の表示を更新する。連動 ON なら回転角度に追従し、OFF なら現在の制限角度をそのまま残す
         * @returns {void}
         */
        function updateConstrainSuggestion() {
            showConstrain(paletteUI.linkRotationCheckbox.value ? currentRotation : currentConstrain);
        }

        /**
         * 回転角度の表示・スライダー・状態を更新し、制限の入力欄の候補値も更新する
         * @param {number} angle - 回転角度（度）
         * @returns {void}
         */
        function setRotation(angle) {
            currentRotation = normalizeAngle(angle);
            paletteUI.rotationValue.text = formatAngle(currentRotation);
            paletteUI.rotationSlider.value = currentRotation;
            updateConstrainSuggestion();
            updateButtonState();
        }

        /**
         * 制限角度を状態に記録する。連動 OFF のときは入力欄にも実際の値を出し、ON のときは回転追従の候補値を残す
         * @param {number} angle - 制限角度（度）
         * @returns {void}
         */
        function setConstrain(angle) {
            currentConstrain = roundAngle(angle);
            if (!paletteUI.linkRotationCheckbox.value) {
                showConstrain(currentConstrain);
            }
            updateButtonState();
        }

        /**
         * ワーカーの結果を受けるコールバックを作る。失敗ならステータスにエラーを出し、成功なら onSuccess を呼ぶ
         * @param {Function} onSuccess - 成功時の処理（結果文字列を受け取る）
         * @returns {Function} runInMainEngine に渡すコールバック
         */
        function handleWorkerResult(onSuccess) {
            return function (result) {
                if (result.indexOf("OK") === 0) {
                    onSuccess(result);
                } else {
                    paletteUI.statusText.text = statusFromResult(result);
                }
            };
        }

        /**
         * 制限角度を環境設定へ適用し、成功したら状態と表示を更新する
         * @param {number} angle - 制限角度（度）
         * @returns {void}
         */
        function commitConstrain(angle) {
            applyConstrainAngle(angle, handleWorkerResult(function () {
                setConstrain(angle);
                showConstrain(angle);
                paletteUI.statusText.text = getLabel("status.applied");
            }));
        }

        /**
         * 連動 ON のとき、入力欄の候補値を環境設定へ自動で適用する（値が変わっていないときは何もしない）
         * @returns {void}
         */
        function applyLinkedConstrain() {
            if (!paletteUI.linkRotationCheckbox.value) { return; }
            var targetAngle = roundAngle(currentRotation);
            if (targetAngle === currentConstrain) { return; }
            commitConstrain(targetAngle);
        }

        /**
         * ビューの回転角度・制限角度・選択角度を1回の委譲で取得して表示へ反映する
         * @returns {void}
         */
        function refreshState() {
            fetchState(function (result) {
                if (result.indexOf("OK:") === 0) {
                    var stateValues = result.substring(3).split(",");
                    setRotation(parseFloat(stateValues[0]));
                    setConstrain(parseFloat(stateValues[1]));
                    paletteUI.selectionAngleValue.text = (stateValues[2] === "") ? "—°" : formatAngle(normalizeAngle(parseFloat(stateValues[2])));
                    paletteUI.statusText.text = "";
                    applyLinkedConstrain();
                } else {
                    if (result.indexOf("NODOC") !== -1) {
                        paletteUI.rotationValue.text = "—°";
                        paletteUI.selectionAngleValue.text = "—°";
                    }
                    paletteUI.statusText.text = statusFromResult(result);
                }
            });
        }

        /**
         * transform フォルダーの外部スクリプトを実行する
         * @param {string} fileName - スクリプトのファイル名
         * @returns {void}
         */
        function runExternalScript(fileName) {
            runScriptFile(TRANSFORM_FOLDER + "/" + fileName, handleWorkerResult(function () {
                paletteUI.statusText.text = "";
            }));
        }

        /**
         * プリセットボタンに適用処理を付ける（ループ変数を閉じ込めるため関数に切り出す）
         * @param {Object} constrainPreset - プリセットの角度（angle）とボタン（button）の対
         * @returns {void}
         */
        function bindConstrainPreset(constrainPreset) {
            constrainPreset.button.onClick = function () {
                commitConstrain(constrainPreset.angle);
            };
        }
        for (var i = 0; i < paletteUI.constrainPresets.length; i++) {
            bindConstrainPreset(paletteUI.constrainPresets[i]);
        }

        /* スライダー操作中：表示だけ更新（Shiftで15°クランプ）/ While dragging: update the display only (snap to 15° with Shift) */
        paletteUI.rotationSlider.onChanging = function () {
            paletteUI.rotationValue.text = snapSliderWhileDragging(paletteUI.rotationSlider) + "°";
        };

        /* スライダー確定：ビューの回転を適用 / On release: apply the view rotation */
        paletteUI.rotationSlider.onChange = function () {
            var angle = readSnappedSliderAngle(paletteUI.rotationSlider);
            setRotation(angle);
            applyViewRotation(angle, handleWorkerResult(function () {
                paletteUI.statusText.text = "";
                applyLinkedConstrain();
            }));
        };

        /* スライダー操作中：入力欄だけ更新（Shiftで15°クランプ）/ While dragging: update the field only (snap to 15° with Shift) */
        paletteUI.constrainSlider.onChanging = function () {
            paletteUI.constrainInput.text = snapSliderWhileDragging(paletteUI.constrainSlider);
        };

        /* スライダー確定：「角度の制限」へ適用 / On release: apply to the constrain angle */
        paletteUI.constrainSlider.onChange = function () {
            commitConstrain(readSnappedSliderAngle(paletteUI.constrainSlider));
        };

        /* 入力欄の確定（Enter・フォーカス移動・↑↓キー）でそのまま適用 / Apply as soon as the field is committed (Enter, focus change, or the arrow keys) */
        paletteUI.constrainInput.onChange = function () {
            var angle = parseFloat(paletteUI.constrainInput.text);
            if (isNaN(angle)) {
                paletteUI.statusText.text = getLabel("alert.invalidAngle");
                return;
            }
            commitConstrain(angle);
        };

        /* 選択に合わせてビューを回転：選択パスの角度を取得し、表示とビュー回転へ反映 / Rotate view to match selection: measure the path angle and reflect it in the display and view */
        paletteUI.rotateViewButton.onClick = function () {
            rotateViewToSelection(handleWorkerResult(function (result) {
                var angle = parseFloat(result.substring(3));
                paletteUI.selectionAngleValue.text = formatAngle(normalizeAngle(angle));
                setRotation(angle);
                paletteUI.statusText.text = getLabel("status.rotatedToSelection");
                applyLinkedConstrain();
            }));
        };

        /* 選択をビューの回転に合わせて回転 / Rotate the selection to match the view rotation */
        paletteUI.rotateSelectionButton.onClick = function () {
            rotateSelectionToView(handleWorkerResult(function () {
                paletteUI.statusText.text = getLabel("status.rotatedSelectionToView");
            }));
        };

        /* 選択の回転をリセット / Reset the selection's rotation */
        paletteUI.resetSelectionButton.onClick = function () {
            resetSelectionRotation(handleWorkerResult(function () {
                paletteUI.statusText.text = getLabel("status.resetSelectionRotation");
            }));
        };

        /* テキストの傾きをリセット（ResetTransform.jsx）/ Reset text tilt (ResetTransform.jsx) */
        paletteUI.resetTextButton.onClick = function () {
            runExternalScript("ResetTransform.jsx");
        };

        /* 画像の傾きをリセット（ResetRotation.jsx）/ Reset image tilt (ResetRotation.jsx) */
        paletteUI.resetImageButton.onClick = function () {
            runExternalScript("ResetRotation.jsx");
        };

        /* リセット（ビューの回転だけ）：0°に戻して表示を更新 / Reset (view rotation only): set it to 0° and refresh the display */
        paletteUI.resetRotationButton.onClick = function () {
            resetViewRotation(handleWorkerResult(function () {
                setRotation(0);
                paletteUI.statusText.text = getLabel("status.resetRotation");
                applyLinkedConstrain();
            }));
        };

        /* 連動の切り替え：ONなら現在の回転角度を環境設定へ適用、OFFなら現在の制限角度を表示し直すだけ。状態はセッションに残す
           / Toggle linking: on, push the current rotation to the preference; off, just redisplay the current constrain angle. The state is kept in the session */
        paletteUI.linkRotationCheckbox.onClick = function () {
            sessionState.linkRotation = paletteUI.linkRotationCheckbox.value;
            updateConstrainSuggestion();
            applyLinkedConstrain();
        };

        /* 更新ボタン：最新の状態を取得し直す / Refresh button: re-fetch the latest state */
        paletteUI.refreshButton.onClick = refreshState;

        /* 初期表示時、およびパレットがアクティブになるたびに最新のビュー回転角度を取得
           / Fetch the latest view rotation on first show and whenever the palette becomes active */
        paletteUI.palette.onShow = refreshState;
        paletteUI.palette.onActivate = refreshState;
    }

    /**
     * パレットを前回の位置に置く（記憶がなければ中央）
     * @param {Window} rotateViewPalette - 対象のパレット
     * @returns {void}
     */
    function restorePaletteLocation(rotateViewPalette) {
        try {
            if (sessionState.location && sessionState.location.length === 2) {
                rotateViewPalette.location = [sessionState.location[0], sessionState.location[1]];
            } else {
                rotateViewPalette.center();
            }
        } catch (e) {
            rotateViewPalette.center();
        }
    }

    /**
     * パレットを作って表示する。すでに開いているパレットがあれば閉じてから作り直す
     * @returns {void}
     */
    function showRotateViewPalette() {
        /* 多重起動の防止：すでに開いているパレットがあれば閉じてから作り直す（閉じる際に位置がセッションへ残る）
           / Prevent duplicates: close an already open palette before rebuilding (closing stores its position in the session) */
        if ($.global[PALETTE_KEY]) {
            try { $.global[PALETTE_KEY].close(); } catch (e) {}
            $.global[PALETTE_KEY] = null;
        }

        var paletteUI = buildPalette();
        bindPaletteHandlers(paletteUI);
        var rotateViewPalette = paletteUI.palette;

        /* 閉じるときに表示位置をセッションへ控え、常駐エンジンの参照も解放
           / Store the position in the session on close, and release the reference held by the resident engine */
        rotateViewPalette.onClose = function () {
            try {
                if (rotateViewPalette.location) {
                    sessionState.location = [rotateViewPalette.location[0], rotateViewPalette.location[1]];
                }
            } catch (e) {}
            $.global[PALETTE_KEY] = null;
        };

        /* 常駐エンジンに保持してGCを避ける / Keep it in the resident engine so it is not garbage-collected */
        $.global[PALETTE_KEY] = rotateViewPalette;

        restorePaletteLocation(rotateViewPalette);
        rotateViewPalette.show();
    }

    showRotateViewPalette();

})();
