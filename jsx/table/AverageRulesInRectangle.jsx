#target illustrator
#targetengine "AverageRulesInRectangleEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
SCRIPTMETA-BEGIN
Script-ID=com.swwwitch.illustrator.AverageRulesInRectangle
Name=外枠の中で罫線を均等に再配置
Author=Masahiro Takano (@swwwitch)
Target-App=Illustrator
Version=1.0.11
Release-Date=2026-04-29
Description-BEGIN
選択した外枠の長方形と、その内側にある縦罫・横罫を、外枠の中で均等に再配置します。
横罫・縦罫はそれぞれON/OFFでき、罫線の長さを外枠の辺に揃えることもできます。

Evenly redistributes the vertical and horizontal rules inside the selected outer rectangle.
Horizontal and vertical rules can be toggled independently, and rules can be stretched to the sides of the frame.
Description-END
SCRIPTMETA-END
*/

/*

### 概要

選択した外枠の長方形と、その内側にある縦罫・横罫を、外枠の中で均等に再配置します。
横罫・縦罫はそれぞれON/OFFでき、罫線の長さを外枠の辺に揃えることもできます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AverageRulesInRectangle.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n7e8c9ce0173e

### Overview

Evenly redistributes the vertical and horizontal rules inside the selected outer rectangle.
Horizontal and vertical rules can be toggled independently, and rules can be stretched to the sides of the frame.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AverageRulesInRectangle.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AverageRulesInRectangle";      /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.11";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-04-29";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-11-20";                   /* 更新日 / last updated */

var SCRIPT_DESCRIPTION_JA = "外枠の中で罫線を均等に再配置"; /* README の見出し / README title (Japanese) */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AverageRulesInRectangle.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AverageRulesInRectangle.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n7e8c9ce0173e"; /* 紹介記事 / article URL */
var SCRIPT_SOURCE_URL  = "https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/table/AverageRulesInRectangle.jsx"; /* 最新版のソース / latest source */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /*
        ［長さを揃える］の初期値 / Initial value of "Match rule lengths"
        true  : 縦罫の上下、横罫の左右を外枠に合わせる
                Snap rule endpoints to the outer rectangle.
        false : 罫線の長さは変えず、位置だけ平均化する
                Keep rule lengths; redistribute positions only.
    */
    var MATCH_LINES_TO_RECT = true;

    /* 罫線の判定に使う許容値（pt） / Tolerance for detecting rules (pt) */
    var TOLERANCE = 0.5;

    // =========================================
    // レイアウト / Layout
    // =========================================

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
            title: { ja: "罫線を平均化", en: "Distribute Rules" }
        },
        panel: {
            target: { ja: "対象", en: "Target" }
        },
        checkbox: {
            averageHorizontal: { ja: "横罫を平均化", en: "Distribute horizontal rules" },
            averageVertical: { ja: "縦罫を平均化", en: "Distribute vertical rules" },
            matchRuleLengths: { ja: "長さを揃える", en: "Match rule lengths" },
            preview: { ja: "プレビュー", en: "Preview" }
        },
        tooltip: {
            averageHorizontal: {
                ja: "長方形の中にある横罫の間隔を均等にします。",
                en: "Evens out the spacing of the horizontal rules inside the rectangle."
            },
            averageVertical: {
                ja: "長方形の中にある縦罫の間隔を均等にします。",
                en: "Evens out the spacing of the vertical rules inside the rectangle."
            },
            matchRuleLengths: {
                ja: "罫の長さを長方形の辺にそろえます。",
                en: "Matches the length of the rules to the sides of the rectangle."
            },
            preview: {
                ja: "結果を画面で確認します。キャンセルすると元に戻ります。",
                en: "Shows the result on the canvas. Cancel restores the original layout."
            }
        },
        button: {
            ok: { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noSelection: { ja: "外枠の長方形と罫線を選択してください。", en: "Please select the outer rectangle and the rules." },
            noPathInSelection: { ja: "選択内にパスがありません。", en: "No paths found in the selection." },
            outerRectNotFound: { ja: "外枠の長方形が見つかりませんでした。", en: "Outer rectangle not found." },
            noRulesFound: { ja: "縦罫または横罫が見つかりませんでした。", en: "No vertical or horizontal rules were found." }
        }
    };

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択から外枠と罫線を見つけ、ダイアログで平均化を確定する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }

        var doc = app.activeDocument;

        if (doc.selection.length === 0) {
            alert(getLabel("alert.noSelection"));
            return;
        }

        /* ロック中・非表示のものは除き、グループ・複合パスの中のパスも集める / Collect paths inside groups and compound paths, skipping locked and hidden ones */
        var selectedPathItems = collectSelectionPathItems(doc.selection, { skipLocked: true, skipHidden: true });

        if (selectedPathItems.length === 0) {
            alert(getLabel("alert.noPathInSelection"));
            return;
        }

        var outerRect = findOuterRectangle(selectedPathItems);

        if (!outerRect) {
            alert(getLabel("alert.outerRectNotFound"));
            return;
        }

        var rectBounds = getRectBounds(outerRect);
        var rules = classifyRules(selectedPathItems, outerRect, rectBounds);

        if (rules.vertical.length === 0 && rules.horizontal.length === 0) {
            alert(getLabel("alert.noRulesFound"));
            return;
        }

        var originalState = savePathState(selectedPathItems);

        /* 開いた時点の設定（見つかった罫線は ON、長さはユーザー設定どおり） / Initial settings: rules found start ON, lengths follow the user setting */
        var initialSettings = {
            averageHorizontal: rules.horizontal.length > 0,
            averageVertical: rules.vertical.length > 0,
            matchRuleLengths: MATCH_LINES_TO_RECT
        };

        var dialogControls = buildDialog(rules.horizontal.length, rules.vertical.length, initialSettings);
        var horizontalRulesCheckbox = dialogControls.horizontalRulesCheckbox;
        var verticalRulesCheckbox = dialogControls.verticalRulesCheckbox;
        var matchRuleLengthsCheckbox = dialogControls.matchRuleLengthsCheckbox;
        var previewCheckbox = dialogControls.previewCheckbox;

        /**
         * 罫線チェックボックスのクリック処理（Option／Alt キーで横罫・縦罫をまとめて切り替える）
         * @param {Checkbox} clickedCheckbox - クリックされたチェックボックス
         * @returns {void}
         */
        function handleRuleCheckboxClick(clickedCheckbox) {
            if (ScriptUI.environment.keyboardState.altKey) {
                toggleBothRuleCheckboxes(clickedCheckbox.value);
            }
            updatePreview();
        }

        /**
         * 有効な横罫・縦罫のチェックボックスをまとめて同じ状態にする
         * @param {boolean} checkboxValue - 設定する値
         * @returns {void}
         */
        function toggleBothRuleCheckboxes(checkboxValue) {
            if (horizontalRulesCheckbox.enabled) {
                horizontalRulesCheckbox.value = checkboxValue;
            }
            if (verticalRulesCheckbox.enabled) {
                verticalRulesCheckbox.value = checkboxValue;
            }
        }

        /**
         * 元の座標に戻したうえで、設定どおりに平均化する
         * @param {{averageHorizontal: boolean, averageVertical: boolean, matchRuleLengths: boolean}} ruleSettings - 平均化の設定
         * @returns {void}
         */
        function applyRuleSettings(ruleSettings) {
            restorePathState(originalState);
            if (ruleSettings.averageVertical) {
                redistributeVerticalRules(rules.vertical, rectBounds, ruleSettings.matchRuleLengths);
            }
            if (ruleSettings.averageHorizontal) {
                redistributeHorizontalRules(rules.horizontal, rectBounds, ruleSettings.matchRuleLengths);
            }
        }

        /**
         * チェックボックスの状態どおりに平均化する（表示前は value を読み戻せないので、ユーザー操作のあとだけ呼ぶ）
         * @returns {void}
         */
        function applyCheckedRules() {
            applyRuleSettings({
                averageHorizontal: horizontalRulesCheckbox.value,
                averageVertical: verticalRulesCheckbox.value,
                matchRuleLengths: matchRuleLengthsCheckbox.value
            });
        }

        /**
         * プレビューを更新する（プレビューが OFF のときは元の座標に戻すだけ）
         * @returns {void}
         */
        function updatePreview() {
            if (previewCheckbox.value) {
                applyCheckedRules();
            } else {
                restorePathState(originalState);
            }
            app.redraw();
        }

        horizontalRulesCheckbox.onClick = function () {
            handleRuleCheckboxClick(horizontalRulesCheckbox);
        };

        verticalRulesCheckbox.onClick = function () {
            handleRuleCheckboxClick(verticalRulesCheckbox);
        };

        matchRuleLengthsCheckbox.onClick = function () {
            updatePreview();
        };

        previewCheckbox.onClick = function () {
            updatePreview();
        };

        dialogControls.btnOK.onClick = function () {
            applyCheckedRules();
            app.redraw();
            dialogControls.rulesDialog.close(1);
        };

        dialogControls.btnCancel.onClick = function () {
            dialogControls.rulesDialog.close(0);
        };

        /* 「キャンセル」の文字では Esc の取消ボタンと認識されない / "キャンセル" is not recognized as the Esc button */
        dialogControls.rulesDialog.cancelElement = dialogControls.btnCancel;

        /* 初回のプレビュー（プレビューは ON で始まる） / Initial preview (Preview starts ON) */
        applyRuleSettings(initialSettings);
        app.redraw();

        addVersionFooter(dialogControls.rulesDialog);
        prepareDialogWindow(dialogControls.rulesDialog, SCRIPT_NAME);

        /* Esc など［OK］以外で閉じたときも元に戻す / Restore however the dialog is closed, unless by OK */
        if (dialogControls.rulesDialog.show() !== 1) {
            restorePathState(originalState);
            app.redraw();
        }
    }

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

// バージョン表示（再利用パーツ） / Version footer (reusable)

var VERSION_FOOTER_TOP_MARGIN = 4;    /* バージョンの行の上に足す余白（px。ウィンドウの間隔に足す）/ extra space above the version row, added to the window spacing */
var VERSION_FOOTER_BOTTOM_MARGIN = 8; /* バージョンの行の下の余白（px。ウィンドウの下余白をこの値にする）/ space below the version row (replaces the window's bottom margin) */
var VERSION_FOOTER_PIXEL_SIZE = 1;    /* ドット1つの大きさ（px）/ size of one dot (px) */
var VERSION_FOOTER_LETTER_GAP = 1;    /* 字の間（ドット）/ gap between glyphs (dots) */
var VERSION_FOOTER_BRAND = "DTP Transit"; /* 右下に出す名前（字形にある文字だけ）/ name shown at the bottom right (glyph characters only) */
var VERSION_FOOTER_COLOR = isDarkUI() ? [0.45, 0.45, 0.45, 1] : [0.62, 0.62, 0.62, 1]; /* 文字の色（ディム表示に近い灰色）/ text color, a gray close to the dimmed look */

/* 字形：Sigi 7px Regular（Rasmus Andersson、CC0。https://github.com/rsms/sigi-pixel-font）の大文字・数字・「. - _ +」・空白を
   9 行のドット（# が点。上7行が大文字の高さ、下2行は小文字の下に出る部分）に写したもの。Sigi に小文字が無いので、a〜z は同じ大きさで描き足した。I は l と見分けるよう上下に横棒を足した
   / Glyphs: uppercase, digits, ". - _ +" and space from Sigi 7px Regular by Rasmus Andersson (CC0, https://github.com/rsms/sigi-pixel-font),
     nine rows each ("#" is a dot; seven rows of cap height plus two for descenders). Sigi has no lowercase, so a-z were drawn to match. I has top and bottom bars to tell it from l */
var VERSION_FOOTER_GLYPH_HEIGHT = 9;  /* 字の高さ（ドット。下2行は g・j・p・q・y の下に出る部分）/ glyph height (dots), including two descender rows */
var VERSION_FOOTER_GLYPHS = {
    "A": [".###.", "#...#", "#...#", "#####", "#...#", "#...#", "#...#", ".....", "....."],
    "B": ["####.", "#...#", "#...#", "####.", "#...#", "#...#", "####.", ".....", "....."],
    "C": [".###.", "#...#", "#....", "#....", "#....", "#...#", ".###.", ".....", "....."],
    "D": ["####.", "#...#", "#...#", "#...#", "#...#", "#...#", "####.", ".....", "....."],
    "E": ["#####", "#....", "#....", "####.", "#....", "#....", "#####", ".....", "....."],
    "F": ["#####", "#....", "#....", "####.", "#....", "#....", "#....", ".....", "....."],
    "G": [".###.", "#...#", "#....", "#..##", "#...#", "#...#", ".###.", ".....", "....."],
    "H": ["#...#", "#...#", "#...#", "#####", "#...#", "#...#", "#...#", ".....", "....."],
    "I": ["###", ".#.", ".#.", ".#.", ".#.", ".#.", "###", "...", "..."],
    "J": ["....#", "....#", "....#", "....#", "....#", "#...#", ".###.", ".....", "....."],
    "K": ["#...#", "#..#.", "#.#..", "##...", "#.#..", "#..#.", "#...#", ".....", "....."],
    "L": ["#...", "#...", "#...", "#...", "#...", "#...", "####", "....", "...."],
    "M": ["#...#", "##.##", "#.#.#", "#...#", "#...#", "#...#", "#...#", ".....", "....."],
    "N": ["#...#", "#...#", "##..#", "#.#.#", "#..##", "#...#", "#...#", ".....", "....."],
    "O": [".###.", "#...#", "#...#", "#...#", "#...#", "#...#", ".###.", ".....", "....."],
    "P": ["####.", "#...#", "#...#", "####.", "#....", "#....", "#....", ".....", "....."],
    "Q": [".###.", "#...#", "#...#", "#...#", "#...#", "#..##", ".####", ".....", "....."],
    "R": ["####.", "#...#", "#...#", "####.", "#.#..", "#..#.", "#...#", ".....", "....."],
    "S": [".###.", "#...#", "#....", ".###.", "....#", "#...#", ".###.", ".....", "....."],
    "T": ["#####", "..#..", "..#..", "..#..", "..#..", "..#..", "..#..", ".....", "....."],
    "U": ["#...#", "#...#", "#...#", "#...#", "#...#", "#...#", ".###.", ".....", "....."],
    "V": ["#...#", "#...#", "#...#", ".#.#.", ".#.#.", "..#..", "..#..", ".....", "....."],
    "W": ["#..#..#", "#..#..#", "#..#..#", "#..#..#", "#..#..#", "#..#..#", ".##.##.", ".......", "......."],
    "X": ["#...#", "#...#", ".#.#.", "..#..", ".#.#.", "#...#", "#...#", ".....", "....."],
    "Y": ["#...#", "#...#", ".#.#.", "..#..", "..#..", "..#..", "..#..", ".....", "....."],
    "Z": ["#####", "....#", "...#.", "..#..", ".#...", "#....", "#####", ".....", "....."],
    "a": ["....", "....", ".##.", "...#", ".###", "#..#", ".###", "....", "...."],
    "b": ["#...", "#...", "###.", "#..#", "#..#", "#..#", "###.", "....", "...."],
    "c": ["....", "....", ".###", "#...", "#...", "#...", ".###", "....", "...."],
    "d": ["...#", "...#", ".###", "#..#", "#..#", "#..#", ".###", "....", "...."],
    "e": ["....", "....", ".##.", "#..#", "####", "#...", ".###", "....", "...."],
    "f": ["..#", ".#.", "###", ".#.", ".#.", ".#.", ".#.", "...", "..."],
    "g": ["....", "....", ".###", "#..#", "#..#", "#..#", ".###", "...#", "###."],
    "h": ["#...", "#...", "###.", "#..#", "#..#", "#..#", "#..#", "....", "...."],
    "i": ["#", ".", "#", "#", "#", "#", "#", ".", "."],
    "j": [".#", "..", ".#", ".#", ".#", ".#", ".#", ".#", "#."],
    "k": ["#...", "#...", "#..#", "#.#.", "##..", "#.#.", "#..#", "....", "...."],
    "l": ["#.", "#.", "#.", "#.", "#.", "#.", ".#", "..", ".."],
    "m": [".....", ".....", "####.", "#.#.#", "#.#.#", "#.#.#", "#.#.#", ".....", "....."],
    "n": ["....", "....", "###.", "#..#", "#..#", "#..#", "#..#", "....", "...."],
    "o": ["....", "....", ".##.", "#..#", "#..#", "#..#", ".##.", "....", "...."],
    "p": ["....", "....", "###.", "#..#", "#..#", "#..#", "###.", "#...", "#..."],
    "q": ["....", "....", ".###", "#..#", "#..#", "#..#", ".###", "...#", "...#"],
    "r": ["...", "...", "#.#", "##.", "#..", "#..", "#..", "...", "..."],
    "s": ["....", "....", ".###", "#...", ".##.", "...#", "###.", "....", "...."],
    "t": ["...", ".#.", "###", ".#.", ".#.", ".#.", ".##", "...", "..."],
    "u": ["....", "....", "#..#", "#..#", "#..#", "#..#", ".###", "....", "...."],
    "v": [".....", ".....", "#...#", "#...#", ".#.#.", ".#.#.", "..#..", ".....", "....."],
    "w": [".....", ".....", "#...#", "#...#", "#.#.#", "#.#.#", ".#.#.", ".....", "....."],
    "x": ["....", "....", "#..#", "#..#", ".##.", "#..#", "#..#", "....", "...."],
    "y": ["....", "....", "#..#", "#..#", "#..#", "#..#", ".###", "...#", "###."],
    "z": ["....", "....", "####", "...#", ".##.", "#...", "####", "....", "...."],
    "0": [".###.", "#...#", "#...#", "#...#", "#...#", "#...#", ".###.", ".....", "....."],
    "1": [".#.", "##.", ".#.", ".#.", ".#.", ".#.", "###", "...", "..."],
    "2": [".###.", "#...#", "....#", "...#.", "..#..", ".#...", "#####", ".....", "....."],
    "3": [".###.", "#...#", "....#", "..##.", "....#", "#...#", ".###.", ".....", "....."],
    "4": ["...#.", "..##.", ".#.#.", "#..#.", "#####", "...#.", "...#.", ".....", "....."],
    "5": ["#####", "#....", "####.", "....#", "....#", "#...#", ".###.", ".....", "....."],
    "6": [".###.", "#...#", "#....", "####.", "#...#", "#...#", ".###.", ".....", "....."],
    "7": ["#####", "....#", "....#", "...#.", "...#.", "..#..", "..#..", ".....", "....."],
    "8": [".###.", "#...#", "#...#", ".###.", "#...#", "#...#", ".###.", ".....", "....."],
    "9": [".###.", "#...#", "#...#", ".####", "....#", "#...#", ".###.", ".....", "....."],
    ".": [".", ".", ".", ".", ".", ".", "#", ".", "."],
    "-": ["...", "...", "...", "###", "...", "...", "...", "...", "..."],
    "_": ["....", "....", "....", "....", "....", "....", "####", "....", "...."],
    "+": [".....", "..#..", "..#..", "#####", "..#..", "..#..", ".....", ".....", "....."],
    " ": ["...", "...", "...", "...", "...", "...", "...", "...", "..."]
};

/**
 * ウィンドウの下端の左にスクリプト名とバージョン（SCRIPT_NAME・SCRIPT_VERSION）、右に VERSION_FOOTER_BRAND を小さく薄く入れる。部品をすべて足したあと、show() の直前に呼ぶ。
 * 直前の行（ボタン行など）の下余白を 0 にし、ウィンドウの下余白を VERSION_FOOTER_BOTTOM_MARGIN にする。
 * InDesign の ScriptUI は文字サイズの指定を無視するので、onDraw でドットを並べて描く（UITheme の isDarkUI() を使う）
 * @param {Window} targetWindow - ダイアログまたはパレット（縦並び）
 * @returns {Group|null} バージョンを描くグループ。すでに足してあれば null
 */
function addVersionFooter(targetWindow) {
    if (targetWindow.versionFooter) return null;

    /* ボタン行の下余白（BUTTON_ROW_BOTTOM_MARGIN）は、この行が引き受ける / this row takes over the button row's bottom margin */
    var lastChild = targetWindow.children[targetWindow.children.length - 1];
    if (lastChild && lastChild.type === "group") {
        var rowMargins = lastChild.margins;
        lastChild.margins = [rowMargins[0], rowMargins[1], rowMargins[2], 0];
    }
    var windowMargins = targetWindow.margins;
    targetWindow.margins = [windowMargins[0], windowMargins[1], windowMargins[2], VERSION_FOOTER_BOTTOM_MARGIN];

    var footerGroup = targetWindow.add("group");
    footerGroup.orientation = "row";
    footerGroup.alignment = ["fill", "bottom"];
    footerGroup.alignChildren = ["left", "center"];
    footerGroup.margins = [0, VERSION_FOOTER_TOP_MARGIN, 0, 0];

    /* 左は「スクリプト名 バージョン」。字形に無い文字（日本語など）を含む名前や、幅が足りないときはバージョンだけ。
       左の欄は余りの幅まで伸ばし、ダイアログの幅はバージョンの分しか求めない（名前でダイアログを広げない）
       / Left: "script name version", or just the version when the name has characters without glyphs (e.g. Japanese)
         or does not fit. The left area takes the extra width and only asks for the version's width, so the name never widens the dialog */
    var versionText = String(SCRIPT_VERSION);
    var leftCandidates = [versionText];
    if (typeof SCRIPT_NAME === "string" && hasVersionFooterGlyphs(SCRIPT_NAME)) leftCandidates.unshift(SCRIPT_NAME + " " + versionText);
    var versionCanvas = addVersionFooterCanvas(footerGroup, leftCandidates);
    versionCanvas.alignment = ["fill", "center"];
    bindVersionInfoClick(versionCanvas); /* クリックで「このスクリプトについて」を開く / click opens About This Script */
    var brandCanvas = addVersionFooterCanvas(footerGroup, [VERSION_FOOTER_BRAND]);
    brandCanvas.alignment = ["right", "center"];
    brandCanvas.maximumSize = brandCanvas.preferredSize;
    targetWindow.versionFooter = footerGroup;
    return versionCanvas;
}

/**
 * 文字列のすべての文字に字形があるかを返す
 * @param {string} dotText - 文字列
 * @returns {boolean} すべてあれば true
 */
function hasVersionFooterGlyphs(dotText) {
    for (var i = 0; i < dotText.length; i++) {
        if (!VERSION_FOOTER_GLYPHS.hasOwnProperty(dotText.charAt(i))) return false;
    }
    return dotText.length > 0;
}

/**
 * ドットで文字列を描くグループを足す。候補のうち、欄の幅に収まる最初のものを描く
 * @param {Group} parentGroup - 足す先
 * @param {string[]} dotTexts - 描く文字列の候補（長い順。いちばん短いものの幅を欄の幅として求める）
 * @returns {Group} 描くグループ
 */
function addVersionFooterCanvas(parentGroup, dotTexts) {
    var dotCanvas = parentGroup.add("group");
    var shortestText = dotTexts[dotTexts.length - 1];
    dotCanvas.preferredSize = [measureVersionFooterDots(shortestText) * VERSION_FOOTER_PIXEL_SIZE, VERSION_FOOTER_GLYPH_HEIGHT * VERSION_FOOTER_PIXEL_SIZE];
    dotCanvas.helpTip = dotTexts[0];
    dotCanvas.onDraw = function () {
        var availableWidth = this.size ? this.size.width : 0;
        var drawText = shortestText;
        for (var i = 0; i < dotTexts.length; i++) {
            if (measureVersionFooterDots(dotTexts[i]) * VERSION_FOOTER_PIXEL_SIZE <= availableWidth) {
                drawText = dotTexts[i];
                break;
            }
        }
        drawVersionFooterDots(this.graphics, drawText);
    };
    return dotCanvas;
}

/**
 * 文字列の幅をドットの数で返す（字形に無い文字は数えない）
 * @param {string} versionText - 文字列
 * @returns {number} 幅（ドット）
 */
function measureVersionFooterDots(versionText) {
    var dotCount = 0;
    for (var i = 0; i < versionText.length; i++) {
        var glyphRows = VERSION_FOOTER_GLYPHS[versionText.charAt(i)];
        if (!glyphRows) continue;
        if (dotCount > 0) dotCount += VERSION_FOOTER_LETTER_GAP;
        dotCount += glyphRows[0].length;
    }
    return dotCount;
}

/**
 * 文字列をドットで描く。横に続くドットは1つの四角にまとめて塗る
 * @param {ScriptUIGraphics} canvasGraphics - 描く先の graphics
 * @param {string} versionText - 文字列
 * @returns {void}
 */
function drawVersionFooterDots(canvasGraphics, versionText) {
    var footerBrush = canvasGraphics.newBrush(canvasGraphics.BrushType.SOLID_COLOR, VERSION_FOOTER_COLOR);
    var dotSize = VERSION_FOOTER_PIXEL_SIZE;
    var cursorDots = 0;
    canvasGraphics.newPath();
    for (var i = 0; i < versionText.length; i++) {
        var glyphRows = VERSION_FOOTER_GLYPHS[versionText.charAt(i)];
        if (!glyphRows) continue;
        if (cursorDots > 0) cursorDots += VERSION_FOOTER_LETTER_GAP;
        for (var row = 0; row < glyphRows.length; row++) {
            var rowText = glyphRows[row];
            var runStart = -1;
            for (var col = 0; col <= rowText.length; col++) {
                var isDot = col < rowText.length && rowText.charAt(col) === "#";
                if (isDot && runStart < 0) runStart = col;
                if (!isDot && runStart >= 0) {
                    canvasGraphics.rectPath((cursorDots + runStart) * dotSize, row * dotSize, (col - runStart) * dotSize, dotSize);
                    runStart = -1;
                }
            }
        }
        cursorDots += glyphRows[0].length;
    }
    canvasGraphics.fillPath(footerBrush);
}

// -----------------------------------------
// このスクリプトについて / About this script
// -----------------------------------------
/* 左下のスクリプト名・バージョンをクリックすると開く。基本情報ブロックの SCRIPT_* を読み、
   この部品だけで動くよう、文言（VERSION_INFO_LABELS）とレイアウトの値も部品の中に持つ
   Opened by clicking the name and version at the bottom left. Reads the SCRIPT_* basic info,
   and keeps its own labels and layout values so the part works on its own */

var VERSION_INFO_LANG = (String($.locale || "").indexOf("ja") === 0) ? "ja" : "en"; /* 文言の言語 / label language */
var VERSION_INFO_IS_INDESIGN = (String(app.name).indexOf("InDesign") >= 0);         /* InDesign で動いているか / running in InDesign */
/* 実行中のスクリプトのファイル。ボタンを押すころには $.fileName が別のものを指すことがあるので、読み込んだ時点で控える
   The running script file, captured at load time; $.fileName may point elsewhere by the time the button is clicked */
var VERSION_INFO_SCRIPT_FILE = new File($.fileName);
var VERSION_INFO_REVEAL_APP_PATH = "/Applications/RevealInFinder.app";   /* Finder で選択表示する補助アプリ / helper app that reveals a file in the Finder */
var VERSION_INFO_REVEAL_PATH_FILE = "/tmp/illustrator_reveal_path.txt";  /* 補助アプリが読むパス（アプリの旧名のまま。変えると読めない）/ path the helper reads (keeps the app's old name) */

var VERSION_INFO_WINDOW_MARGINS = 16;               /* ウィンドウの余白 / window margin */
var VERSION_INFO_WINDOW_SPACING = 12;               /* ウィンドウ内の間隔 / window spacing */
var VERSION_INFO_PANEL_MARGINS = [16, 20, 16, 12];  /* パネルの余白 [左,上,右,下] / panel margins */
var VERSION_INFO_ROW_SPACING = 6;                   /* 項目名と値の間隔 / gap between a label and its value */
var VERSION_INFO_BUTTON_SPACING = 10;               /* ボタンどうしの間隔 / gap between buttons */
var VERSION_INFO_BUTTON_ROW_MARGINS = [0, 5, 0, 14]; /* ボタン行の余白（Illustrator 標準のダイアログに合わせる）/ button row margins */
var VERSION_INFO_COPY_BUTTON_SIZE = 22;             /* コピーボタンの一辺 / copy button size */
/* コピーボタンの配色。∧∨（ステップボタン）と同じく、背景に重ねる半透明 / Copy button colors, translucent overlays like the stepper */
var VERSION_INFO_COPY_FILL_COLOR    = isDarkUI() ? [0, 0, 0, 0.10] : [1, 1, 1, 0.50]; /* 地 / background */
var VERSION_INFO_COPY_FRAME_COLOR   = isDarkUI() ? [1, 1, 1, 0.12] : [0, 0, 0, 0.15]; /* 枠線 / frame */
var VERSION_INFO_COPY_PRESSED_COLOR = isDarkUI() ? [1, 1, 1, 0.12] : [0, 0, 0, 0.13]; /* 押下中 / pressed */
var VERSION_INFO_COPY_ICON_COLOR    = isDarkUI() ? [1, 1, 1, 0.85] : [0, 0, 0, 0.70]; /* アイコン / icon */

var VERSION_INFO_LABELS = {
    dialog: {
        title: { ja: "このスクリプトについて", en: "About This Script" }
    },
    panel: {
        generalInfo: { ja: "基本情報", en: "General" },
        links: { ja: "リンク", en: "Links" }
    },
    fieldLabel: {
        scriptName: { ja: "スクリプト名", en: "Script Name" },
        scriptFeature: { ja: "機能", en: "Feature" },
        scriptVersion: { ja: "バージョン", en: "Version" }
    },
    button: {
        openReadme: { ja: "README", en: "README" },
        openSourceCode: { ja: "ソースコード", en: "Source Code" },
        openArticle: { ja: "記事（note）", en: "Article (note)" },
        revealInFinder: { ja: "Finder に表示", en: "Show in Finder" },
        revealInExplorer: { ja: "エクスプローラーに表示", en: "Show in Explorer" },
        close: { ja: "閉じる", en: "Close" }
    },
    tooltip: {
        versionInfo: { ja: "クリックで、このスクリプトについての情報を表示", en: "Click for info about this script" },
        copyScriptName: { ja: "スクリプト名をコピー", en: "Copy Script Name" }
    },
    alert: {
        copyFailed: { ja: "スクリプト名をコピーできませんでした。", en: "Could not copy the script name." }
    }
};

/**
 * VERSION_INFO_LABELS の文言を今の言語で返す
 * @param {Object} labelEntry - { ja, en }
 * @returns {string} 文言
 */
function getVersionInfoLabel(labelEntry) {
    return labelEntry[VERSION_INFO_LANG] || labelEntry.en;
}

/**
 * 基本情報ブロックの値を返す（変数が無い・文字列でなければ空文字）。typeof で読むので、未定義でも例外にならない
 * @param {string} variableKey - "name" / "version" / "feature" / "readme" / "source" / "article"
 * @returns {string} 値
 */
function getVersionInfoValue(variableKey) {
    var infoValue;
    if (variableKey === "name") infoValue = (typeof SCRIPT_NAME !== "undefined") ? SCRIPT_NAME : "";
    else if (variableKey === "version") infoValue = (typeof SCRIPT_VERSION !== "undefined") ? SCRIPT_VERSION : "";
    else if (variableKey === "feature") infoValue = (typeof SCRIPT_DESCRIPTION_JA !== "undefined") ? SCRIPT_DESCRIPTION_JA : "";
    else if (variableKey === "readme") {
        if (VERSION_INFO_LANG === "ja") infoValue = (typeof SCRIPT_README_JA !== "undefined") ? SCRIPT_README_JA : "";
        else infoValue = (typeof SCRIPT_README_EN !== "undefined") ? SCRIPT_README_EN : "";
    }
    else if (variableKey === "source") infoValue = (typeof SCRIPT_SOURCE_URL !== "undefined") ? SCRIPT_SOURCE_URL : "";
    else if (variableKey === "article") infoValue = (typeof SCRIPT_ARTICLE_URL !== "undefined") ? SCRIPT_ARTICLE_URL : "";
    return (infoValue === undefined || infoValue === null) ? "" : String(infoValue);
}

/**
 * 左下のスクリプト名・バージョンをクリックしたら、「このスクリプトについて」のダイアログを開くようにする
 * @param {Group} versionCanvas - addVersionFooterCanvas() で作った左の欄
 * @returns {void}
 */
function bindVersionInfoClick(versionCanvas) {
    versionCanvas.helpTip = getVersionInfoLabel(VERSION_INFO_LABELS.tooltip.versionInfo);
    /* mousedown で開くと、離したときのイベントが閉じたあとの欄に届くので mouseup で開く
       Open on mouseup so the release does not land on the footer after the dialog closes */
    versionCanvas.addEventListener("mouseup", function () {
        showVersionInfoDialog(versionCanvas.window.type === "palette");
    });
}

/**
 * 「このスクリプトについて」のダイアログを開く。スクリプト名（コピーボタン付き）・機能・バージョンの［基本情報］パネル、
 * README・ソースコード・記事を開く［リンク］パネル、［Finder に表示］［閉じる］を並べる
 * @param {boolean} isFromPalette - 常駐パレットから開いたなら true（Illustrator のコピーを BridgeTalk 経由にする）
 * @returns {void}
 */
function showVersionInfoDialog(isFromPalette) {
    var infoDialog = new Window("dialog", getVersionInfoLabel(VERSION_INFO_LABELS.dialog.title));
    infoDialog.orientation = "column";
    infoDialog.alignChildren = "fill";
    infoDialog.margins = VERSION_INFO_WINDOW_MARGINS;
    infoDialog.spacing = VERSION_INFO_WINDOW_SPACING;

    var infoPanel = addVersionInfoPanel(infoDialog, VERSION_INFO_LABELS.panel.generalInfo);
    /* コピーボタンはスクリプト名だけ / Only the script name gets a copy button */
    addVersionInfoRow(infoPanel, VERSION_INFO_LABELS.fieldLabel.scriptName, getVersionInfoValue("name"), isFromPalette);
    /* 機能は README の見出しで日本語だけなので、日本語の画面だけに出す / The feature is the Japanese README title, shown only in the Japanese UI */
    if (VERSION_INFO_LANG === "ja" && getVersionInfoValue("feature") !== "") {
        addVersionInfoRow(infoPanel, VERSION_INFO_LABELS.fieldLabel.scriptFeature, getVersionInfoValue("feature"), null);
    }
    addVersionInfoRow(infoPanel, VERSION_INFO_LABELS.fieldLabel.scriptVersion, getVersionInfoValue("version"), null);
    alignVersionInfoLabels(infoPanel);

    /* URL が無いボタンはディム表示にして、並びはスクリプトごとに変えない / Dim buttons without a URL; keep the same buttons everywhere */
    var linksPanel = addVersionInfoPanel(infoDialog, VERSION_INFO_LABELS.panel.links);
    var linkButtonRow = linksPanel.add("group");
    linkButtonRow.orientation = "row";
    linkButtonRow.alignment = ["left", "center"];
    linkButtonRow.alignChildren = ["left", "center"];
    linkButtonRow.spacing = VERSION_INFO_BUTTON_SPACING;
    addVersionInfoLinkButton(infoDialog, linkButtonRow, VERSION_INFO_LABELS.button.openReadme, getVersionInfoValue("readme"));
    addVersionInfoLinkButton(infoDialog, linkButtonRow, VERSION_INFO_LABELS.button.openSourceCode, getVersionInfoValue("source"));
    addVersionInfoLinkButton(infoDialog, linkButtonRow, VERSION_INFO_LABELS.button.openArticle, getVersionInfoValue("article"));

    /* ボタン行：左に［Finder に表示］、右に［閉じる］ / Button row: Show in Finder at the left, Close at the right */
    var buttonRow = infoDialog.add("group");
    buttonRow.orientation = "row";
    buttonRow.alignment = ["fill", "bottom"];
    buttonRow.alignChildren = ["left", "center"];
    buttonRow.margins = VERSION_INFO_BUTTON_ROW_MARGINS;
    buttonRow.spacing = VERSION_INFO_BUTTON_SPACING;
    var revealLabel = (File.fs === "Windows") ? VERSION_INFO_LABELS.button.revealInExplorer : VERSION_INFO_LABELS.button.revealInFinder;
    var revealButton = buttonRow.add("button", undefined, getVersionInfoLabel(revealLabel));
    if (VERSION_INFO_SCRIPT_FILE.exists) {
        revealButton.helpTip = VERSION_INFO_SCRIPT_FILE.fsName;
        revealButton.onClick = function () {
            revealVersionInfoScriptFile(VERSION_INFO_SCRIPT_FILE);
            infoDialog.close();
        };
    } else {
        revealButton.enabled = false;
    }
    /* 余りの幅を吸って、［閉じる］を右端に寄せる / Absorbs the extra width so Close sits at the right edge */
    var buttonSpacer = buttonRow.add("group");
    buttonSpacer.alignment = ["fill", "fill"];
    buttonSpacer.minimumSize.width = 0;
    var closeButton = buttonRow.add("button", undefined, getVersionInfoLabel(VERSION_INFO_LABELS.button.close), { name: "ok" });
    closeButton.alignment = ["right", "center"];
    /* return でも esc でも閉じる / Close with both Return and Esc */
    infoDialog.defaultElement = closeButton;
    infoDialog.cancelElement = closeButton;
    infoDialog.show();
}

/**
 * 「このスクリプトについて」のダイアログにパネルを足す
 * @param {Window} infoDialog - 足す先のダイアログ
 * @param {Object} titleEntry - パネル名の文言 { ja, en }
 * @returns {Panel} パネル
 */
function addVersionInfoPanel(infoDialog, titleEntry) {
    var infoPanel = infoDialog.add("panel", undefined, getVersionInfoLabel(titleEntry));
    infoPanel.orientation = "column";
    infoPanel.alignChildren = ["fill", "top"];
    infoPanel.alignment = "fill";
    infoPanel.margins = VERSION_INFO_PANEL_MARGINS;
    infoPanel.spacing = VERSION_INFO_WINDOW_SPACING;
    return infoPanel;
}

/**
 * ［基本情報］パネルに「項目名 : 値」の1行を足す（値は見せるだけなので statictext）
 * @param {Panel} infoPanel - 足す先のパネル
 * @param {Object} labelEntry - 項目名の文言 { ja, en }
 * @param {string} valueText - 表示する値
 * @param {boolean|null} copyFromPalette - 値の右にコピーボタンを置くなら true / false（常駐パレットから開いたか）、置かないなら null
 * @returns {StaticText} 値の statictext
 */
function addVersionInfoRow(infoPanel, labelEntry, valueText, copyFromPalette) {
    var infoRow = infoPanel.add("group");
    infoRow.orientation = "row";
    infoRow.alignment = ["fill", "center"];
    infoRow.alignChildren = ["left", "center"];
    infoRow.spacing = VERSION_INFO_ROW_SPACING;
    /* 項目名のコロンは日本語が「 :」、英語が「:」 / Colon: " :" in Japanese, ":" in English */
    var rowLabel = infoRow.add("statictext", undefined, getVersionInfoLabel(labelEntry) + (VERSION_INFO_LANG === "ja" ? " :" : ":"));
    rowLabel.justify = "right";
    var valueLabel = infoRow.add("statictext", undefined, valueText);
    /* 余りの幅を取り、コピーボタンを右端に置く / Take the extra width so the copy button sits at the right */
    valueLabel.alignment = ["fill", "center"];
    if (copyFromPalette !== null) addVersionInfoCopyButton(infoRow, valueLabel, copyFromPalette);
    return valueLabel;
}

/**
 * ［基本情報］パネルの項目名を、いちばん長いものの幅にそろえる
 * @param {Panel} infoPanel - 対象のパネル
 * @returns {void}
 */
function alignVersionInfoLabels(infoPanel) {
    var rowLabels = [];
    var labelWidth = 0;
    for (var i = 0; i < infoPanel.children.length; i++) {
        var rowLabel = infoPanel.children[i].children[0]; /* 行の先頭は項目名 / each row starts with its label */
        rowLabels.push(rowLabel);
        labelWidth = Math.max(labelWidth, rowLabel.graphics.measureString(rowLabel.text)[0]);
    }
    for (var j = 0; j < rowLabels.length; j++) {
        rowLabels[j].preferredSize.width = Math.ceil(labelWidth) + 2;
    }
}

/**
 * URL を開くボタンを足す。URL が無ければディム表示にする。開いたらダイアログは閉じる
 * @param {Window} infoDialog - ボタンを置くダイアログ
 * @param {Group} buttonGroup - 足す先のグループ
 * @param {Object} labelEntry - ボタンの文言 { ja, en }
 * @param {string} targetUrl - 開く URL（空ならディム表示）
 * @returns {void}
 */
function addVersionInfoLinkButton(infoDialog, buttonGroup, labelEntry, targetUrl) {
    var linkButton = buttonGroup.add("button", undefined, getVersionInfoLabel(labelEntry));
    if (!targetUrl) {
        linkButton.enabled = false;
        return;
    }
    linkButton.helpTip = targetUrl;
    linkButton.onClick = function () {
        openVersionInfoUrl(targetUrl);
        infoDialog.close();
    };
}

/**
 * 値をクリップボードへコピーするアイコンのボタンを足す。コピーしたらアイコンをチェックマークに変え、ポインタが外れたら戻す
 * @param {Group} infoRow - 足す先の行
 * @param {StaticText} valueLabel - コピーする値の statictext
 * @param {boolean} isFromPalette - 常駐パレットから開いたなら true
 * @returns {Group} ボタンとして使う group
 */
function addVersionInfoCopyButton(infoRow, valueLabel, isFromPalette) {
    var buttonSize = VERSION_INFO_COPY_BUTTON_SIZE;
    var copyButton = infoRow.add("group");
    copyButton.preferredSize = [buttonSize, buttonSize];
    copyButton.minimumSize = [buttonSize, buttonSize];
    copyButton.maximumSize = [buttonSize, buttonSize];
    copyButton.helpTip = getVersionInfoLabel(VERSION_INFO_LABELS.tooltip.copyScriptName);
    copyButton.isPressed = false;
    copyButton.isCopied = false;

    copyButton.onDraw = function () {
        var buttonGraphics = copyButton.graphics;
        var fillColor = copyButton.isPressed ? VERSION_INFO_COPY_PRESSED_COLOR : VERSION_INFO_COPY_FILL_COLOR;
        buttonGraphics.newPath();
        buttonGraphics.rectPath(0, 0, buttonSize, buttonSize);
        buttonGraphics.fillPath(buttonGraphics.newBrush(buttonGraphics.BrushType.SOLID_COLOR, fillColor));
        buttonGraphics.newPath();
        buttonGraphics.rectPath(0.5, 0.5, buttonSize - 1, buttonSize - 1);
        buttonGraphics.strokePath(buttonGraphics.newPen(buttonGraphics.PenType.SOLID_COLOR, VERSION_INFO_COPY_FRAME_COLOR, 1));
        if (copyButton.isCopied) drawVersionInfoCheckIcon(buttonGraphics, buttonSize);
        else drawVersionInfoCopyIcon(buttonGraphics, buttonSize);
    };

    /**
     * 状態を変えて描き直す（group には notify() が無いので隠して再表示する）
     * @param {boolean} isPressed - 押下中なら true
     * @param {boolean} isCopied - コピー済みなら true
     * @returns {void}
     */
    function repaint(isPressed, isCopied) {
        if (copyButton.isPressed === isPressed && copyButton.isCopied === isCopied) return;
        copyButton.isPressed = isPressed;
        copyButton.isCopied = isCopied;
        copyButton.hide();
        copyButton.show();
    }
    copyButton.addEventListener("mousedown", function () { repaint(true, false); });
    copyButton.addEventListener("mouseup", function () {
        if (!copyButton.isPressed) return;
        var hasCopied = copyVersionInfoText(valueLabel.text, isFromPalette);
        repaint(false, hasCopied);
        if (!hasCopied) alert(getVersionInfoLabel(VERSION_INFO_LABELS.alert.copyFailed));
    });
    /* 押したまま外へ出たとき・コピーのあと離れたときは元の見た目に戻す / Reset when the pointer leaves */
    copyButton.addEventListener("mouseout", function () { repaint(false, false); });
    return copyButton;
}

/**
 * コピーのアイコン（後ろの四角に前の四角を重ねた形）を描く
 * @param {ScriptUIGraphics} buttonGraphics - 描画先
 * @param {number} buttonSize - ボタンの一辺
 * @returns {void}
 */
function drawVersionInfoCopyIcon(buttonGraphics, buttonSize) {
    var iconPen = buttonGraphics.newPen(buttonGraphics.PenType.SOLID_COLOR, VERSION_INFO_COPY_ICON_COLOR, 1);
    var sheetSize = 9;
    var offset = 3;
    var frontLeft = Math.round((buttonSize - sheetSize - offset) / 2) + offset + 0.5;
    var frontTop = Math.round((buttonSize - sheetSize - offset) / 2) + offset + 0.5;
    var backLeft = frontLeft - offset;
    var backTop = frontTop - offset;
    /* 後ろの四角は、前の四角に隠れる部分を描かない（上と左の辺だけ）/ Draw only the visible top and left of the back sheet */
    buttonGraphics.newPath();
    buttonGraphics.moveTo(frontLeft, backTop + sheetSize);
    buttonGraphics.lineTo(backLeft, backTop + sheetSize);
    buttonGraphics.lineTo(backLeft, backTop);
    buttonGraphics.lineTo(backLeft + sheetSize, backTop);
    buttonGraphics.lineTo(backLeft + sheetSize, frontTop);
    buttonGraphics.strokePath(iconPen);
    buttonGraphics.newPath();
    buttonGraphics.rectPath(frontLeft, frontTop, sheetSize, sheetSize);
    buttonGraphics.strokePath(iconPen);
}

/**
 * コピーしたあとのチェックマークを描く
 * @param {ScriptUIGraphics} buttonGraphics - 描画先
 * @param {number} buttonSize - ボタンの一辺
 * @returns {void}
 */
function drawVersionInfoCheckIcon(buttonGraphics, buttonSize) {
    var center = buttonSize / 2;
    buttonGraphics.newPath();
    buttonGraphics.moveTo(center - 4.5, center);
    buttonGraphics.lineTo(center - 1.5, center + 3);
    buttonGraphics.lineTo(center + 4.5, center - 3.5);
    buttonGraphics.strokePath(buttonGraphics.newPen(buttonGraphics.PenType.SOLID_COLOR, VERSION_INFO_COPY_ICON_COLOR, 1.5));
}

/**
 * 文字列をクリップボードにコピーする。
 * InDesign は AppleScript（Windows は VBScript）の doScript で書き込む。
 * Illustrator は文字列を直接送る API が無いので、一時レイヤーのテキストフレームを選んでコピーする（プレーンテキストも入る）。
 * Illustrator の常駐パレットからは DOM に触れないので、同じ処理を BridgeTalk でメインエンジンに渡す
 * @param {string} copyText - コピーする文字列
 * @param {boolean} isFromPalette - 常駐パレットから開いたなら true
 * @returns {boolean} コピーの操作まで進めたら true
 */
function copyVersionInfoText(copyText, isFromPalette) {
    try {
        if (VERSION_INFO_IS_INDESIGN) {
            if (File.fs === "Windows") {
                app.doScript('CreateObject("htmlfile").parentWindow.clipboardData.setData "text", "' + copyText.replace(/"/g, '""') + '"', ScriptLanguage.VISUAL_BASIC);
            } else {
                app.doScript('set the clipboard to "' + copyText.replace(/\\/g, "\\\\").replace(/"/g, '\\"') + '"', ScriptLanguage.APPLESCRIPT_LANGUAGE);
            }
            return true;
        }
        var copyScript = buildVersionInfoCopyScript(copyText);
        if (!isFromPalette) return eval(copyScript) === true;
        var copyMessage = new BridgeTalk();
        copyMessage.target = BridgeTalk.appSpecifier; /* 複数バージョンを起動していても、今のバージョンに送る / the running version, even with several installed */
        copyMessage.body = copyScript;
        copyMessage.send();
        return true;
    } catch (e) {
        return false;
    }
}

/**
 * Illustrator でコピーする処理を文字列で組み立てる（BridgeTalk でも送れるよう、関数の toString() は使わない）。
 * 一時レイヤーにテキストフレームを作り、フレームを選んでメニューコマンドでコピーして、レイヤーごと消す。
 * 文字として選択（textRange.select()）してコピーすると「オブジェクトをコピーできません」の警告が出るので使わない。
 * 選択と作業レイヤーは元に戻す。ドキュメントが無ければ一時ドキュメントを作って閉じる
 * @param {string} copyText - コピーする文字列
 * @returns {string} 実行すると true / false を返すスクリプト
 */
function buildVersionInfoCopyScript(copyText) {
    return [
        "(function (copyText) {",
        "    var usingTempDoc = (app.documents.length === 0);",
        "    var targetDoc = usingTempDoc ? app.documents.add() : app.activeDocument;",
        "    var savedSelection = (!usingTempDoc && targetDoc.selection instanceof Array) ? targetDoc.selection : [];",
        "    var savedActiveLayer = usingTempDoc ? null : targetDoc.activeLayer;",
        "    var tempLayer = null;",
        "    var hasCopied = false;",
        "    try {",
        "        tempLayer = targetDoc.layers.add();",
        "        var tempFrame = tempLayer.textFrames.add();",
        "        tempFrame.contents = copyText;",
        "        app.redraw();",
        "        app.executeMenuCommand('deselectall');",
        "        tempFrame.selected = true;",
        "        app.redraw();",
        "        app.executeMenuCommand('copy');",
        "        app.redraw();",
        "        hasCopied = true;",
        "    } catch (e) { }",
        "    if (tempLayer) { try { tempLayer.remove(); } catch (e) { } }",
        "    if (usingTempDoc) { targetDoc.close(SaveOptions.DONOTSAVECHANGES); return hasCopied; }",
        "    try { targetDoc.activeLayer = savedActiveLayer; } catch (e) { }",
        "    targetDoc.selection = null;",
        "    for (var i = 0; i < savedSelection.length; i++) { try { savedSelection[i].selected = true; } catch (e) { } }",
        "    app.redraw();",
        "    return hasCopied;",
        "})(" + quoteVersionInfoString(copyText) + ");"
    ].join(" "); /* 行コメントを含まないので1行につなぐ（BridgeTalk の転送で改行が化けない）/ joined on one line; there are no line comments */
}

/**
 * 文字列を、スクリプトに埋め込める二重引用符のリテラルにする
 * @param {string} sourceText - 文字列
 * @returns {string} "…" の形のリテラル
 */
function quoteVersionInfoString(sourceText) {
    return '"' + String(sourceText).replace(/\\/g, "\\\\").replace(/"/g, '\\"').replace(/\r/g, "\\r").replace(/\n/g, "\\n") + '"';
}

/**
 * ファイルを Finder で表示する。RevealInFinder.app があればファイルを選択した状態で表示し、無ければ入っているフォルダーを開く
 * @param {File} scriptFile - 表示するファイル
 * @returns {void}
 */
function revealVersionInfoScriptFile(scriptFile) {
    /* .app は実体がフォルダーなので File.exists が false になる環境がある。Folder でも確かめる / An .app is a folder, so check with Folder too */
    var hasRevealApp = (File.fs === "Macintosh")
        && (new Folder(VERSION_INFO_REVEAL_APP_PATH).exists || new File(VERSION_INFO_REVEAL_APP_PATH).exists);
    /* アプリは一時ファイルからパスを読むので、書き出してから起動する / The app reads the path from a temp file */
    if (hasRevealApp && writeVersionInfoRevealPath(scriptFile.fsName) && new File(VERSION_INFO_REVEAL_APP_PATH).execute()) return;
    /* アプリが無い・起動できないときは、何も起きないままにせずフォルダーを開く / Otherwise open the enclosing folder */
    scriptFile.parent.execute();
}

/**
 * RevealInFinder.app に渡すパスを一時ファイルへ書き出す（UTF-8、改行なし）
 * @param {string} targetPath - 絶対パス
 * @returns {boolean} 書き出せたら true
 */
function writeVersionInfoRevealPath(targetPath) {
    var pathFile = new File(VERSION_INFO_REVEAL_PATH_FILE);
    pathFile.encoding = "UTF-8";
    pathFile.lineFeed = "Unix";
    if (!pathFile.open("w")) return false;
    var hasWritten = pathFile.write(targetPath);
    pathFile.close();
    return hasWritten;
}

/**
 * URL を既定のブラウザーで開く（一時的なショートカットファイルを作って開く）
 * @param {string} targetUrl - 開く URL
 * @returns {void}
 */
function openVersionInfoUrl(targetUrl) {
    var isMac = (File.fs === "Macintosh");
    var shortcutFile = new File(Folder.temp.fsName + "/VersionInfoLink" + (isMac ? ".webloc" : ".url"));
    shortcutFile.encoding = "UTF-8";
    shortcutFile.open("w");
    if (isMac) {
        var escapedUrl = targetUrl.replace(/&/g, "&amp;").replace(/</g, "&lt;");
        shortcutFile.write('<?xml version="1.0" encoding="UTF-8"?>\n'
            + '<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">\n'
            + '<plist version="1.0"><dict><key>URL</key><string>' + escapedUrl + '</string></dict></plist>\n');
    } else {
        shortcutFile.write("[InternetShortcut]\r\nURL=" + targetUrl + "\r\n");
    }
    shortcutFile.close();
    shortcutFile.execute();
}

// バージョン表示（再利用パーツ）ここまで / End of the reusable version footer
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
    // ダイアログ / Dialog
    // =========================================

    /**
     * ダイアログを組み立てる
     * @param {number} horizontalRuleCount - 見つかった横罫の数
     * @param {number} verticalRuleCount - 見つかった縦罫の数
     * @param {{averageHorizontal: boolean, averageVertical: boolean, matchRuleLengths: boolean}} initialSettings - チェックボックスの初期値
     * @returns {{rulesDialog: Window, horizontalRulesCheckbox: Checkbox, verticalRulesCheckbox: Checkbox, matchRuleLengthsCheckbox: Checkbox, previewCheckbox: Checkbox, btnCancel: Button, btnOK: Button}} ダイアログとコントロール
     */
    function buildDialog(horizontalRuleCount, verticalRuleCount, initialSettings) {
        var rulesDialog = new Window("dialog", getLabel("dialog.title"));
        setupWindow(rulesDialog);

        var targetPanel = rulesDialog.add("panel", undefined, getLabel("panel.target"));
        setupPanel(targetPanel, 6);

        var horizontalRulesCheckbox = addOptionCheckbox(targetPanel,
            "checkbox.averageHorizontal", "tooltip.averageHorizontal", horizontalRuleCount > 0, initialSettings.averageHorizontal);
        var verticalRulesCheckbox = addOptionCheckbox(targetPanel,
            "checkbox.averageVertical", "tooltip.averageVertical", verticalRuleCount > 0, initialSettings.averageVertical);
        var matchRuleLengthsCheckbox = addOptionCheckbox(targetPanel,
            "checkbox.matchRuleLengths", "tooltip.matchRuleLengths", horizontalRuleCount > 0 || verticalRuleCount > 0, initialSettings.matchRuleLengths);

        var previewGroup = rulesDialog.add("group");
        previewGroup.orientation = "row";
        previewGroup.alignment = "center";

        var previewCheckbox = previewGroup.add("checkbox", undefined, getLabel("checkbox.preview"));
        previewCheckbox.helpTip = getLabel("tooltip.preview");
        previewCheckbox.value = true;

        var buttonRow = addButtonRow(rulesDialog);
        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"));
        var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"));
        alignRightOnlyButtonRow(buttonRow);

        return {
            rulesDialog: rulesDialog,
            horizontalRulesCheckbox: horizontalRulesCheckbox,
            verticalRulesCheckbox: verticalRulesCheckbox,
            matchRuleLengthsCheckbox: matchRuleLengthsCheckbox,
            previewCheckbox: previewCheckbox,
            btnCancel: btnCancel,
            btnOK: btnOK
        };
    }

    /**
     * チェックボックスを tooltip 付きで追加する
     * @param {Panel} parentPanel - 追加先のパネル
     * @param {string} labelPath - 表示名の LABELS パス
     * @param {string} tooltipPath - tooltip の LABELS パス
     * @param {boolean} isEnabled - 操作できるかどうか
     * @param {boolean} initialValue - 初期値
     * @returns {Checkbox} 追加したチェックボックス
     */
    function addOptionCheckbox(parentPanel, labelPath, tooltipPath, isEnabled, initialValue) {
        var optionCheckbox = parentPanel.add("checkbox", undefined, getLabel(labelPath));
        optionCheckbox.helpTip = getLabel(tooltipPath);
        optionCheckbox.value = initialValue === true;
        optionCheckbox.enabled = isEnabled;
        return optionCheckbox;
    }

    // 選択の収集と境界（再利用パーツ） / Selection items and bounds (reusable)

    /* 座標を同じと見なす許容値（pt） / Tolerance for treating coordinates as equal, in points */
    var SELECTION_ITEMS_TOLERANCE = 0.001;

    /**
     * 選択やコレクションを、オブジェクトの配列にそろえる
     * TextRange・PathItem は length を持つので、typename で1個か集まりかを見分ける
     * @param {*} source - doc.selection、配列、DOM のコレクション、または単独のオブジェクト
     * @returns {Array} オブジェクトの配列（空なら []）
     */
    function normalizeSelectionItems(source) {
        var items = [];
        if (!source) return items;
        var typeName = "";
        try { typeName = source.typename || ""; } catch (e) { /* 読めない種類 / unreadable kind */ }
        /* 単数形の typename は1個（PageItems などのコレクションは s で終わる）
           A singular typename is one object (collections such as PageItems end in s) */
        if (typeName && !/s$/.test(typeName)) return [source];
        if (typeof source.length !== "number") return items;
        for (var i = 0; i < source.length; i++) items.push(source[i]);
        return items;
    }

    /**
     * 文字カーソルの選択（TextRange）を、それを含むテキストフレームに読み替える
     * @param {TextRange} textRange - 文字の範囲
     * @returns {TextFrame|null} テキストフレーム（たどれなければ null）
     */
    function resolveTextRangeFrame(textRange) {
        var current = textRange;
        /* parent をたどる（深さは念のため制限） / Walk up the parents, with a safety limit */
        for (var depth = 0; depth < 10 && current; depth++) {
            try {
                if (current.typename === "TextFrame") return current;
                current = current.parent;
            } catch (e) {
                break;
            }
        }
        /* ストーリーの先頭フレームで代用する / Fall back to the first frame of the story */
        try {
            var storyFrames = textRange.story.textFrames;
            if (storyFrames.length > 0) return storyFrames[0];
        } catch (e2) { /* ストーリーを持たない / no story */ }
        return null;
    }

    /**
     * 選択から条件に合うオブジェクトを集める（グループ・レイヤーを再帰でたどり、重複は除く）
     * 条件に合ったオブジェクトの中へは進まない
     * @param {*} source - doc.selection、配列、コレクション、または単独のオブジェクト
     * @param {Object} [options] - 収集の設定
     * @param {function(PageItem): boolean} [options.accept] - 集める条件（既定はグループ・レイヤー以外すべて）
     * @param {boolean} [options.enterGroups] - グループの中をたどる（既定 true）
     * @param {boolean} [options.enterClipGroups] - クリップグループの中をたどる（既定は enterGroups と同じ）
     * @param {boolean} [options.enterCompoundPaths] - 複合パスの中のパスをたどる（既定 false）
     * @param {boolean} [options.textRangeToFrame] - 文字の選択をテキストフレームに読み替える（既定 true）
     * @param {boolean} [options.skipLocked] - ロックされたものを中ごと外す（既定 false）
     * @param {boolean} [options.skipHidden] - 非表示のものを中ごと外す（既定 false）
     * @param {boolean} [options.skipClipMasks] - クリッピングマスクを外す（既定 false）
     * @param {boolean} [options.skipGuides] - ガイドを外す（既定 false）
     * @param {boolean} [options.unique] - 同じ参照を1回だけにする（既定 true。数千件で遅ければ false）
     * @returns {Array} 集めたオブジェクト（前面→背面の順）
     */
    function collectSelectionItems(source, options) {
        var opts = options || {};
        var enterGroups = (opts.enterGroups !== false);
        var enterClipGroups = (opts.enterClipGroups === undefined) ? enterGroups : (opts.enterClipGroups === true);
        var accept = opts.accept || function (item) {
            return item.typename !== "GroupItem" && item.typename !== "Layer";
        };
        var collected = [];

        /**
         * 集めた配列に加える（unique のときは同じ参照を足さない）
         * @param {PageItem} item - 加えるオブジェクト
         * @returns {void}
         */
        function pushItem(item) {
            if (opts.unique !== false) {
                for (var k = 0; k < collected.length; k++) {
                    if (collected[k] === item) return;
                }
            }
            collected.push(item);
        }

        /**
         * 設定に従って外すオブジェクトか判定する
         * @param {PageItem} item - 判定するオブジェクト
         * @returns {boolean} 外すなら true
         */
        function isSkipped(item) {
            try {
                if (item.typename === "Layer") {
                    if (opts.skipLocked && item.locked) return true;
                    if (opts.skipHidden && !item.visible) return true;
                    return false;
                }
                if (opts.skipLocked && item.locked) return true;
                if (opts.skipHidden && item.hidden) return true;
                if (opts.skipGuides && item.guides === true) return true;
                if (opts.skipClipMasks && isClipMaskItem(item)) return true;
            } catch (e) {
                /* 読めないプロパティは「外さない」に倒す / Unreadable properties do not exclude */
            }
            return false;
        }

        /**
         * 1件をたどって集める
         * @param {PageItem} item - 対象のオブジェクト
         * @returns {void}
         */
        function visit(item) {
            if (!item) return;
            var typeName = "";
            try { typeName = item.typename; } catch (e) { return; }

            if (typeName === "TextRange" || typeName === "InsertionPoint") {
                if (opts.textRangeToFrame === false) {
                    if (accept(item)) pushItem(item);
                    return;
                }
                visit(resolveTextRangeFrame(item));
                return;
            }
            if (isSkipped(item)) return;
            if (accept(item)) {
                pushItem(item);
                return;
            }

            var children = null;
            if (typeName === "GroupItem") {
                var isClipped = false;
                try { isClipped = (item.clipped === true); } catch (e2) { }
                if (isClipped ? enterClipGroups : enterGroups) children = item.pageItems;
            } else if (typeName === "CompoundPathItem") {
                if (opts.enterCompoundPaths) children = item.pathItems;
            } else if (typeName === "Layer") {
                /* 重なり順はサブレイヤーとページアイテムで別々なので、ページアイテム→サブレイヤーの順にする
                   Page items and sublayers stack separately; visit page items first, then sublayers */
                walk(item.pageItems);
                walk(item.layers);
                return;
            }
            if (children) walk(children);
        }

        /**
         * 集まりの各要素をたどる
         * @param {*} list - 配列またはコレクション
         * @returns {void}
         */
        function walk(list) {
            var listItems = normalizeSelectionItems(list);
            for (var i = 0; i < listItems.length; i++) visit(listItems[i]);
        }

        walk(source);
        return collected;
    }

    /**
     * テキストフレームの種類を "point" / "area" / "path" で返す
     * @param {TextFrame} textFrame - テキストフレーム
     * @returns {string} 種類のキー（判定できなければ ""）
     */
    function getTextFrameKindKey(textFrame) {
        try {
            if (textFrame.kind === TextType.POINTTEXT) return "point";
            if (textFrame.kind === TextType.AREATEXT) return "area";
            if (textFrame.kind === TextType.PATHTEXT) return "path";
        } catch (e) { /* kind を読めない / kind is unreadable */ }
        return "";
    }

    /**
     * 選択からテキストフレームを集める（グループの中・文字カーソルの選択を含む）
     * @param {*} source - doc.selection など
     * @param {Object} [options] - collectSelectionItems と同じ設定に加えて次を受ける
     * @param {string[]} [options.kinds] - 集める種類（"point" / "area" / "path"。既定はすべて）
     * @returns {TextFrame[]} テキストフレーム（前面→背面の順）
     */
    function collectSelectionTextFrames(source, options) {
        var opts = {};
        var sourceOptions = options || {};
        for (var key in sourceOptions) {
            if (sourceOptions.hasOwnProperty(key)) opts[key] = sourceOptions[key];
        }
        var kindFilter = null;
        if (opts.kinds && opts.kinds.length) {
            kindFilter = {};
            for (var i = 0; i < opts.kinds.length; i++) kindFilter[opts.kinds[i]] = true;
        }
        opts.accept = function (item) {
            if (item.typename !== "TextFrame") return false;
            return !kindFilter || kindFilter[getTextFrameKindKey(item)] === true;
        };
        /* 種類で外したテキストは中をたどらない（accept が false でも子は無い） / Text frames have no children to walk */
        return collectSelectionItems(source, opts);
    }

    /**
     * 選択からパスを集める（グループの中を含む）
     * @param {*} source - doc.selection など
     * @param {Object} [options] - collectSelectionItems と同じ設定に加えて次を受ける
     * @param {string} [options.compoundPaths] - 複合パスの扱い。"children"（中のパス、既定）/ "whole"（複合パスごと）/ "skip"（外す）
     * @returns {Array} PathItem（"whole" のときは CompoundPathItem も）の配列
     */
    function collectSelectionPathItems(source, options) {
        var opts = {};
        var sourceOptions = options || {};
        for (var key in sourceOptions) {
            if (sourceOptions.hasOwnProperty(key)) opts[key] = sourceOptions[key];
        }
        var compoundMode = opts.compoundPaths || "children";
        opts.enterCompoundPaths = (compoundMode === "children");
        opts.accept = function (item) {
            if (item.typename === "PathItem") return true;
            return compoundMode === "whole" && item.typename === "CompoundPathItem";
        };
        return collectSelectionItems(source, opts);
    }

    /**
     * クリッピングマスク（クリップグループの型）か判定する
     * パスは clipping、複合パスは中の先頭パスの clipping、テキストは clipping が無いので「クリップグループの先頭」で見る
     * @param {PageItem} item - 判定するオブジェクト
     * @returns {boolean} マスクなら true
     */
    function isClipMaskItem(item) {
        try {
            if (item.typename === "PathItem") return item.clipping === true;
            if (item.typename === "CompoundPathItem") {
                return item.pathItems.length > 0 && item.pathItems[0].clipping === true;
            }
            if (item.typename === "TextFrame") {
                var parentGroup = item.parent;
                return parentGroup.typename === "GroupItem" && parentGroup.clipped === true &&
                    parentGroup.pageItems.length > 0 && parentGroup.pageItems[0] === item;
            }
        } catch (e) { /* 読めない種類はマスクではない / unreadable kinds are not masks */ }
        return false;
    }

    /**
     * クリップグループの型（マスク）を返す
     * フラグで探し、見つからなければ先頭（pageItems[0]）を返す（型は常に最前面。テキストの型はフラグを持たない）
     * @param {GroupItem} groupItem - 対象のグループ
     * @returns {PageItem|null} マスク（クリップグループでなければ null）
     */
    function getClipMaskItem(groupItem) {
        try {
            if (!groupItem || groupItem.typename !== "GroupItem" || groupItem.clipped !== true) return null;
            var groupChildren = groupItem.pageItems;
            if (groupChildren.length === 0) return null;
            for (var i = 0; i < groupChildren.length; i++) {
                var childType = groupChildren[i].typename;
                if ((childType === "PathItem" || childType === "CompoundPathItem") && isClipMaskItem(groupChildren[i])) {
                    return groupChildren[i];
                }
            }
            return groupChildren[0];
        } catch (e) {
            return null;
        }
    }

    /**
     * グループの中（入れ子を含む）にクリップグループがあるか判定する
     * @param {GroupItem} groupItem - 対象のグループ
     * @returns {boolean} あれば true
     */
    function hasClippedDescendant(groupItem) {
        try {
            var groupChildren = groupItem.pageItems;
            for (var i = 0; i < groupChildren.length; i++) {
                if (groupChildren[i].typename !== "GroupItem") continue;
                if (groupChildren[i].clipped === true || hasClippedDescendant(groupChildren[i])) return true;
            }
        } catch (e) { /* 中を読めない / cannot read the children */ }
        return false;
    }

    /**
     * 環境設定の［プレビュー境界を使用］を読む
     * @returns {boolean} オンなら true（読めなければ false）
     */
    function readUsePreviewBoundsPreference() {
        try {
            return app.preferences.getBooleanPreference("includeStrokeInBounds");
        } catch (e) {
            return false;
        }
    }

    /**
     * 見た目どおりの境界を返す。クリップグループはマスクの境界、
     * 中にクリップグループを含むグループは子の境界を合わせたもの（隠れた部分を含めない）
     * @param {PageItem} item - 対象のオブジェクト
     * @param {boolean} [usePreviewBounds] - true で visibleBounds、false で geometricBounds（省略時は環境設定に従う）
     * @returns {number[]|null} [左, 上, 右, 下] の新しい配列（測れなければ null）
     */
    function getClipAwareBounds(item, usePreviewBounds) {
        var usePreview = (usePreviewBounds === undefined || usePreviewBounds === null) ?
            readUsePreviewBoundsPreference() : (usePreviewBounds === true);
        try {
            var measuredItem = item;
            if (item.typename === "GroupItem") {
                var maskItem = getClipMaskItem(item);
                if (maskItem) {
                    measuredItem = maskItem;
                } else if (hasClippedDescendant(item)) {
                    /* グループ自体の効果（影など）の広がりは含まれなくなる
                       This leaves out the reach of effects applied to the group itself (drop shadows etc.) */
                    var childBounds = getClipAwareUnionBounds(filterMeasurableChildren(item.pageItems), usePreview);
                    if (childBounds) return childBounds;
                }
            }
            var bounds = usePreview ? measuredItem.visibleBounds : measuredItem.geometricBounds;
            return [bounds[0], bounds[1], bounds[2], bounds[3]];
        } catch (e) {
            return null;
        }
    }

    /**
     * 境界の計算に入れる子だけを残す（非表示とガイドを外す）
     * @param {*} childList - 子のコレクション
     * @returns {Array} 残した子
     */
    function filterMeasurableChildren(childList) {
        var childItems = normalizeSelectionItems(childList);
        var measurable = [];
        for (var i = 0; i < childItems.length; i++) {
            try {
                if (childItems[i].hidden === true || childItems[i].guides === true) continue;
            } catch (e) { /* 読めなければ残す / keep when unreadable */ }
            measurable.push(childItems[i]);
        }
        return measurable;
    }

    /**
     * 複数のオブジェクトを囲む外接範囲を返す（クリップグループはマスクで測る）
     * @param {*} items - オブジェクトの配列・コレクション・選択
     * @param {boolean} [usePreviewBounds] - true で visibleBounds、false で geometricBounds（省略時は環境設定に従う）
     * @returns {number[]|null} [左, 上, 右, 下]（測れるものが無ければ null）
     */
    function getClipAwareUnionBounds(items, usePreviewBounds) {
        var usePreview = (usePreviewBounds === undefined || usePreviewBounds === null) ?
            readUsePreviewBoundsPreference() : (usePreviewBounds === true);
        var itemList = normalizeSelectionItems(items);
        var unionBounds = null;
        for (var i = 0; i < itemList.length; i++) {
            var itemBounds = getClipAwareBounds(itemList[i], usePreview);
            if (!itemBounds) continue;
            if (!unionBounds) {
                unionBounds = itemBounds;
                continue;
            }
            if (itemBounds[0] < unionBounds[0]) unionBounds[0] = itemBounds[0];
            if (itemBounds[1] > unionBounds[1]) unionBounds[1] = itemBounds[1];
            if (itemBounds[2] > unionBounds[2]) unionBounds[2] = itemBounds[2];
            if (itemBounds[3] < unionBounds[3]) unionBounds[3] = itemBounds[3];
        }
        return unionBounds;
    }

    /**
     * 2つの座標を許容値つきで比べる
     * @param {number} valueA - 座標A（pt）
     * @param {number} valueB - 座標B（pt）
     * @param {number} [tolerance] - 許容値（pt、既定は SELECTION_ITEMS_TOLERANCE）
     * @returns {boolean} 差が許容値以下なら true
     */
    function isNearlySameCoordinate(valueA, valueB, tolerance) {
        var limit = (typeof tolerance === "number") ? tolerance : SELECTION_ITEMS_TOLERANCE;
        return Math.abs(valueA - valueB) <= limit;
    }

    /**
     * 2つの境界を許容値つきで比べる
     * @param {number[]} boundsA - [左, 上, 右, 下]
     * @param {number[]} boundsB - [左, 上, 右, 下]
     * @param {number} [tolerance] - 許容値（pt、既定は SELECTION_ITEMS_TOLERANCE）
     * @returns {boolean} 4辺とも許容値以内なら true
     */
    function areBoundsNearlyEqual(boundsA, boundsB, tolerance) {
        if (!boundsA || !boundsB) return false;
        for (var i = 0; i < 4; i++) {
            if (!isNearlySameCoordinate(boundsA[i], boundsB[i], tolerance)) return false;
        }
        return true;
    }

    // 選択の収集と境界（再利用パーツ）ここまで / End of the reusable selection items and bounds

    // =========================================
    // 外枠と罫線の判定 / Detecting the frame and rules
    // =========================================

    /**
     * もっとも面積の大きい閉じたパスを外枠とみなす
     * @param {PathItem[]} candidatePathItems - 候補のパス
     * @returns {PathItem|null} 外枠のパス。見つからなければ null
     */
    function findOuterRectangle(candidatePathItems) {
        var outerRectangleCandidate = null;
        var maxArea = 0;

        for (var i = 0; i < candidatePathItems.length; i++) {
            var pathItem = candidatePathItems[i];

            if (!pathItem.closed) {
                continue;
            }

            var geometricBounds = pathItem.geometricBounds;
            var pathWidth = geometricBounds[2] - geometricBounds[0];
            var pathHeight = geometricBounds[1] - geometricBounds[3];

            if (pathWidth <= 0 || pathHeight <= 0) {
                continue;
            }

            var pathArea = pathWidth * pathHeight;

            if (pathArea > maxArea) {
                maxArea = pathArea;
                outerRectangleCandidate = pathItem;
            }
        }

        return outerRectangleCandidate;
    }

    /**
     * 外枠の座標と寸法をまとめて返す
     * @param {PathItem} outerRect - 外枠のパス
     * @returns {{left: number, top: number, right: number, bottom: number, width: number, height: number}} 外枠の範囲
     */
    function getRectBounds(outerRect) {
        var outerBounds = outerRect.geometricBounds;
        return {
            left: outerBounds[0],
            top: outerBounds[1],
            right: outerBounds[2],
            bottom: outerBounds[3],
            width: outerBounds[2] - outerBounds[0],
            height: outerBounds[1] - outerBounds[3]
        };
    }

    /**
     * 外枠の中にある縦罫・横罫を振り分け、縦罫は左から右、横罫は上から下に並べる
     * @param {PathItem[]} candidatePathItems - 候補のパス
     * @param {PathItem} outerRect - 外枠のパス（対象から除く）
     * @param {Object} rectBounds - getRectBounds() の戻り値
     * @returns {{vertical: PathItem[], horizontal: PathItem[]}} 縦罫と横罫
     */
    function classifyRules(candidatePathItems, outerRect, rectBounds) {
        var verticalRules = [];
        var horizontalRules = [];

        for (var i = 0; i < candidatePathItems.length; i++) {
            var pathItem = candidatePathItems[i];

            if (pathItem === outerRect) {
                continue;
            }

            if (isRuleInRect(pathItem, rectBounds, true)) {
                verticalRules.push(pathItem);
            } else if (isRuleInRect(pathItem, rectBounds, false)) {
                horizontalRules.push(pathItem);
            }
        }

        /* 縦罫は左から右へ、横罫は上から下へソート / Sort vertical L→R, horizontal T→B */
        verticalRules.sort(function (firstRule, secondRule) {
            return getCenterX(firstRule) - getCenterX(secondRule);
        });

        horizontalRules.sort(function (firstRule, secondRule) {
            return getCenterY(secondRule) - getCenterY(firstRule);
        });

        return { vertical: verticalRules, horizontal: horizontalRules };
    }

    /**
     * 外枠の中にある縦罫（または横罫）かどうかを判定する
     * 縦罫は「長さ＝高さ、並ぶ方向＝X」、横罫は「長さ＝幅、並ぶ方向＝Y」として同じ条件で見る
     * @param {PathItem} pathItem - 判定するパス
     * @param {Object} rectBounds - getRectBounds() の戻り値
     * @param {boolean} isVertical - true なら縦罫、false なら横罫として判定する
     * @returns {boolean} 罫線とみなせるとき true
     */
    function isRuleInRect(pathItem, rectBounds, isVertical) {
        if (pathItem.closed) {
            return false;
        }

        var geometricBounds = pathItem.geometricBounds;
        var pathWidth = Math.abs(geometricBounds[2] - geometricBounds[0]);
        var pathHeight = Math.abs(geometricBounds[1] - geometricBounds[3]);
        var centerX = (geometricBounds[0] + geometricBounds[2]) / 2;
        var centerY = (geometricBounds[1] + geometricBounds[3]) / 2;

        if (isVertical) {
            return hasRuleShape(pathHeight, pathWidth, rectBounds.height) &&
                isInsideRange(centerX, rectBounds.left, rectBounds.right) &&
                isNearRange(centerY, rectBounds.bottom, rectBounds.top);
        }
        return hasRuleShape(pathWidth, pathHeight, rectBounds.width) &&
            isInsideRange(centerY, rectBounds.bottom, rectBounds.top) &&
            isNearRange(centerX, rectBounds.left, rectBounds.right);
    }

    /**
     * 細長い形で、外枠の辺の 20% 以上の長さがあるかどうか
     * @param {number} ruleLength - 罫線の方向の長さ
     * @param {number} ruleThickness - 罫線と直交する方向の幅
     * @param {number} outerLength - 同じ方向の外枠の長さ
     * @returns {boolean} 罫線の形とみなせるとき true
     */
    function hasRuleShape(ruleLength, ruleThickness, outerLength) {
        if (ruleLength <= 0) {
            return false;
        }
        if (!(ruleThickness <= TOLERANCE || ruleLength > ruleThickness * 5)) {
            return false;
        }
        return !(ruleLength < outerLength * 0.2);
    }

    /**
     * 値が範囲の内側に、許容値より離れて収まっているかどうか（両端ちょうどは外）
     * @param {number} value - 調べる値
     * @param {number} rangeMin - 範囲の最小値
     * @param {number} rangeMax - 範囲の最大値
     * @returns {boolean} 内側なら true
     */
    function isInsideRange(value, rangeMin, rangeMax) {
        return value > rangeMin + TOLERANCE && value < rangeMax - TOLERANCE;
    }

    /**
     * 値が範囲に、許容値の はみ出しまで含めて収まっているかどうか
     * @param {number} value - 調べる値
     * @param {number} rangeMin - 範囲の最小値
     * @param {number} rangeMax - 範囲の最大値
     * @returns {boolean} 収まっていれば true
     */
    function isNearRange(value, rangeMin, rangeMax) {
        return value >= rangeMin - TOLERANCE && value <= rangeMax + TOLERANCE;
    }

    /**
     * パスの外接矩形の中心 X を返す
     * @param {PathItem} pathItem - 対象のパス
     * @returns {number} 中心の X 座標
     */
    function getCenterX(pathItem) {
        var geometricBounds = pathItem.geometricBounds;
        return (geometricBounds[0] + geometricBounds[2]) / 2;
    }

    /**
     * パスの外接矩形の中心 Y を返す
     * @param {PathItem} pathItem - 対象のパス
     * @returns {number} 中心の Y 座標
     */
    function getCenterY(pathItem) {
        var geometricBounds = pathItem.geometricBounds;
        return (geometricBounds[1] + geometricBounds[3]) / 2;
    }

    // =========================================
    // 罫線の再配置 / Redistributing rules
    // =========================================

    /**
     * 縦罫を外枠の幅で等間隔に並べ直す
     * @param {PathItem[]} verticalRules - 左から右に並んだ縦罫
     * @param {Object} rectBounds - getRectBounds() の戻り値
     * @param {boolean} matchRuleLengths - 2点の罫線の上下を外枠に合わせるかどうか
     * @returns {void}
     */
    function redistributeVerticalRules(verticalRules, rectBounds, matchRuleLengths) {
        var ruleCount = verticalRules.length;

        for (var i = 0; i < ruleCount; i++) {
            var verticalRule = verticalRules[i];
            var targetX = rectBounds.left + rectBounds.width * (i + 1) / (ruleCount + 1);

            if (matchRuleLengths && verticalRule.pathPoints.length === 2) {
                setTwoPointVerticalLine(verticalRule, targetX, rectBounds.top, rectBounds.bottom);
            } else {
                translatePath(verticalRule, targetX - getCenterX(verticalRule), 0);
            }
        }
    }

    /**
     * 横罫を外枠の高さで等間隔に並べ直す
     * @param {PathItem[]} horizontalRules - 上から下に並んだ横罫
     * @param {Object} rectBounds - getRectBounds() の戻り値
     * @param {boolean} matchRuleLengths - 2点の罫線の左右を外枠に合わせるかどうか
     * @returns {void}
     */
    function redistributeHorizontalRules(horizontalRules, rectBounds, matchRuleLengths) {
        var ruleCount = horizontalRules.length;

        for (var i = 0; i < ruleCount; i++) {
            var horizontalRule = horizontalRules[i];
            var targetY = rectBounds.top - rectBounds.height * (i + 1) / (ruleCount + 1);

            if (matchRuleLengths && horizontalRule.pathPoints.length === 2) {
                setTwoPointHorizontalLine(horizontalRule, targetY, rectBounds.left, rectBounds.right);
            } else {
                translatePath(horizontalRule, 0, targetY - getCenterY(horizontalRule));
            }
        }
    }

    /**
     * パスのすべてのアンカーポイントとハンドルを平行移動する
     * @param {PathItem} pathItem - 対象のパス
     * @param {number} deltaX - X 方向の移動量
     * @param {number} deltaY - Y 方向の移動量
     * @returns {void}
     */
    function translatePath(pathItem, deltaX, deltaY) {
        for (var i = 0; i < pathItem.pathPoints.length; i++) {
            movePathPoint(pathItem.pathPoints[i], deltaX, deltaY);
        }
    }

    /**
     * 2点の縦罫を、指定の X 位置で上端から下端までの線にする
     * @param {PathItem} pathItem - 2点のパス
     * @param {number} x - 配置する X 座標
     * @param {number} top - 上端の Y 座標
     * @param {number} bottom - 下端の Y 座標
     * @returns {void}
     */
    function setTwoPointVerticalLine(pathItem, x, top, bottom) {
        var firstPoint = pathItem.pathPoints[0];
        var secondPoint = pathItem.pathPoints[1];
        var isFirstOnTop = firstPoint.anchor[1] >= secondPoint.anchor[1];

        movePointTo(isFirstOnTop ? firstPoint : secondPoint, x, top);
        movePointTo(isFirstOnTop ? secondPoint : firstPoint, x, bottom);
    }

    /**
     * 2点の横罫を、指定の Y 位置で左端から右端までの線にする
     * @param {PathItem} pathItem - 2点のパス
     * @param {number} y - 配置する Y 座標
     * @param {number} left - 左端の X 座標
     * @param {number} right - 右端の X 座標
     * @returns {void}
     */
    function setTwoPointHorizontalLine(pathItem, y, left, right) {
        var firstPoint = pathItem.pathPoints[0];
        var secondPoint = pathItem.pathPoints[1];
        var isFirstOnLeft = firstPoint.anchor[0] <= secondPoint.anchor[0];

        movePointTo(isFirstOnLeft ? firstPoint : secondPoint, left, y);
        movePointTo(isFirstOnLeft ? secondPoint : firstPoint, right, y);
    }

    /**
     * アンカーポイントを指定の座標へ移動する（ハンドルも同じ量だけ動かす）
     * @param {PathPoint} pathPoint - 対象のアンカーポイント
     * @param {number} targetX - 移動先の X 座標
     * @param {number} targetY - 移動先の Y 座標
     * @returns {void}
     */
    function movePointTo(pathPoint, targetX, targetY) {
        movePathPoint(pathPoint, targetX - pathPoint.anchor[0], targetY - pathPoint.anchor[1]);
    }

    /**
     * アンカーポイントと両ハンドルを同じ量だけ動かす
     * @param {PathPoint} pathPoint - 対象のアンカーポイント
     * @param {number} deltaX - X 方向の移動量
     * @param {number} deltaY - Y 方向の移動量
     * @returns {void}
     */
    function movePathPoint(pathPoint, deltaX, deltaY) {
        pathPoint.anchor = [pathPoint.anchor[0] + deltaX, pathPoint.anchor[1] + deltaY];
        pathPoint.leftDirection = [pathPoint.leftDirection[0] + deltaX, pathPoint.leftDirection[1] + deltaY];
        pathPoint.rightDirection = [pathPoint.rightDirection[0] + deltaX, pathPoint.rightDirection[1] + deltaY];
    }

    // =========================================
    // 座標の保存と復元 / Saving and restoring coordinates
    // =========================================

    /**
     * 編集前のパスの座標を保存する
     * @param {PathItem[]} targetPathItems - 保存するパス
     * @returns {Object[]} パスごとのアンカーポイントの状態
     */
    function savePathState(targetPathItems) {
        var pathStates = [];

        for (var pathIndex = 0; pathIndex < targetPathItems.length; pathIndex++) {
            var pathItem = targetPathItems[pathIndex];
            var pointStates = [];

            for (var pointIndex = 0; pointIndex < pathItem.pathPoints.length; pointIndex++) {
                var pathPoint = pathItem.pathPoints[pointIndex];

                pointStates.push({
                    anchor: [pathPoint.anchor[0], pathPoint.anchor[1]],
                    leftDirection: [pathPoint.leftDirection[0], pathPoint.leftDirection[1]],
                    rightDirection: [pathPoint.rightDirection[0], pathPoint.rightDirection[1]],
                    pointType: pathPoint.pointType
                });
            }

            pathStates.push({ pathItem: pathItem, pointStates: pointStates });
        }

        return pathStates;
    }

    /**
     * 保存しておいた座標へ戻す
     * @param {Object[]} pathStates - savePathState() の戻り値
     * @returns {void}
     */
    function restorePathState(pathStates) {
        for (var pathIndex = 0; pathIndex < pathStates.length; pathIndex++) {
            var pathItem = pathStates[pathIndex].pathItem;
            var pointStates = pathStates[pathIndex].pointStates;

            for (var pointIndex = 0; pointIndex < pointStates.length; pointIndex++) {
                var pathPoint = pathItem.pathPoints[pointIndex];
                var pointState = pointStates[pointIndex];

                pathPoint.anchor = [pointState.anchor[0], pointState.anchor[1]];
                pathPoint.leftDirection = [pointState.leftDirection[0], pointState.leftDirection[1]];
                pathPoint.rightDirection = [pointState.rightDirection[0], pointState.rightDirection[1]];
                pathPoint.pointType = pointState.pointType;
            }
        }
    }

    main();

})();
