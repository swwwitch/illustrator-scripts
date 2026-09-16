#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したオブジェクトの明るさを、スライダーで調整します。
プラスで明るく、マイナスで暗くなり、プレビューを確認しながら決めた結果は、取り消し1ステップにまとまります。

詳細は README を参照してください。

### Overview

Adjusts the brightness of the selected objects with a slider.
Positive values lighten and negative values darken, and the previewed result commits as a single undo step.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ColorToneSlider";              /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-12-28";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-17";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ColorToneSlider.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ColorToneSlider.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n88e33648b19a"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function() {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* スライダーで指定できる範囲（%）/ Range available on the slider (%) */
    var SLIDER_MIN_PERCENT = -50;
    var SLIDER_MAX_PERCENT = 50;

    /* 入力欄で指定できる範囲（%）/ Range available in the text field (%) */
    var INPUT_MIN_PERCENT = -100;
    var INPUT_MAX_PERCENT = 100;

    /* 調整量の初期値（%）/ Initial adjustment amount (%) */
    var DEFAULT_PERCENT = 0;

    /* shift 併用時のスライダーの刻み（%）/ Slider step while shift is held (%) */
    var SLIDER_SHIFT_STEP = 10;

    /* リセットのショートカットキー / Shortcut key for reset */
    var RESET_SHORTCUT_KEY = "R";

    // =========================================
    // レイアウト / Layout
    // =========================================

    var WINDOW_MARGINS           = 16;   /* ウィンドウ外周の余白 / window margin */
    var WINDOW_SPACING           = 12;   /* ウィンドウ内の要素間隔 / window spacing */
    var ROW_SPACING              = 8;    /* 行内の要素間隔 / spacing inside a row */
    var UNIT_SPACING             = 4;    /* 入力欄と単位の間隔 / spacing before a unit label */
    var BUTTON_ROW_TOP_MARGIN    = 8;    /* ボタン行の上余白 / top margin of the button row */
    var SLIDER_WIDTH             = 300;  /* スライダーの幅 / slider width */
    var PERCENT_INPUT_CHARACTERS = 4;    /* 調整量の入力欄の文字数 / width of the percent field */

    /**
     * ウィンドウの共通レイアウトを設定する
     * @param {Window} targetWindow - 対象のウィンドウ
     * @returns {void}
     */
    function setupWindow(targetWindow) {
        targetWindow.orientation = "column";
        targetWindow.alignChildren = ["fill", "top"];
        targetWindow.margins = WINDOW_MARGINS;
        targetWindow.spacing = WINDOW_SPACING;
    }

    /**
     * 行グループの共通レイアウトを設定する（横位置と天地を対で指定し、子は伸ばさない）
     * @param {Group} rowGroup - 対象のグループ
     * @param {string} [alignment] - 行の横位置（省略時は "left"）
     * @param {number} [spacing] - 要素間隔（省略時は ROW_SPACING）
     * @returns {void}
     */
    function setupRow(rowGroup, alignment, spacing) {
        rowGroup.orientation = "row";
        rowGroup.alignment = [alignment || "left", "center"];
        rowGroup.alignChildren = ["left", "center"];
        rowGroup.spacing = (typeof spacing === "number") ? spacing : ROW_SPACING;
    }

    /**
     * 数値入力欄に↑↓キーでの増減を付ける（shift 併用で ±10）
     * @param {EditText} editText - 対象の入力欄
     * @param {boolean} allowNegative - マイナスの値を許可するか
     * @param {function} [onChange] - 値を変えたあとに呼ぶ処理
     * @param {number} [minValue] - 下限値
     * @returns {void}
     */
    function changeValueByArrowKey(editText, allowNegative, onChange, minValue) {
        editText.addEventListener("keydown", function(event) {
            if (event.keyName !== "Up" && event.keyName !== "Down") return;
            var current = Number(editText.text);
            if (isNaN(current)) return;

            var sign = (event.keyName === "Up") ? 1 : -1;
            var step = ScriptUI.environment.keyboardState.shiftKey ? 10 : 1;
            var next = (step === 10) ?
                (sign > 0 ? Math.ceil((current + 1) / step) * step : Math.floor((current - 1) / step) * step) :
                current + sign * step;

            next = Math.round(next);
            if (!allowNegative && next < 0) next = 0;
            if (typeof minValue === "number" && next < minValue) next = minValue;

            editText.text = next;
            event.preventDefault();
            if (typeof onChange === "function") onChange();
        });
    }

    // =========================================
    // 定数 / Constants
    // =========================================

    /**
     * @typedef {object} ColorModel
     * @property {Array<string>} channels - 明るさの調整対象になるカラー値のプロパティ名
     * @property {Array<string>} references - 値ではなく参照として引き継ぐプロパティ名
     * @property {number} maxValue - カラー値の最大値
     * @property {boolean} additive - 加法混色（値が大きいほど明るい）かどうか
     * @property {function(): object} createColor - 同じカラーモデルの空のカラーを作る
     */

    /* 調整に対応するカラーモデル（グラデーション・パターンは対象外）/ Supported color models */
    var COLOR_MODELS = {
        CMYKColor: {
            channels: ["cyan", "magenta", "yellow", "black"],
            references: [],
            maxValue: 100,
            additive: false,
            createColor: function() { return new CMYKColor(); }
        },
        RGBColor: {
            channels: ["red", "green", "blue"],
            references: [],
            maxValue: 255,
            additive: true,
            createColor: function() { return new RGBColor(); }
        },
        GrayColor: {
            channels: ["gray"],
            references: [],
            maxValue: 100,
            additive: false,
            createColor: function() { return new GrayColor(); }
        },
        SpotColor: {
            channels: ["tint"],
            references: ["spot"],
            maxValue: 100,
            additive: false,
            createColor: function() { return new SpotColor(); }
        }
    };

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /* 現在の UI 言語 / Current UI language */
    var uiLang = ($.locale.indexOf("ja") === 0) ? "ja" : "en";

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "明るさ調整", en: "Adjust Brightness" }
        },
        fieldLabel: {
            amount: { ja: "調整量", en: "Amount" }
        },
        sliderEnd: {
            darker: { ja: "暗く", en: "Darker" },
            lighter: { ja: "明るく", en: "Lighter" }
        },
        button: {
            reset: { ja: "リセット", en: "Reset" },
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" }
        },
        tooltip: {
            slider: {
                ja: "ドラッグして明るさを調整します。右へ動かすと明るく、左へ動かすと暗くなります。shift 併用で10%刻み。",
                en: "Drag to adjust the brightness. Drag right to lighten, left to darken. Hold shift for 10% steps."
            },
            percentInput: {
                ja: "-100〜100% の範囲で直接入力できます。↑↓キーで ±1、shift 併用で ±10。",
                en: "Enter a value from -100 to 100%. Arrow keys step by 1, and shift steps by 10."
            },
            reset: {
                ja: "調整量を 0% に戻します（R キー）。",
                en: "Resets the adjustment to 0% (R key)."
            }
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noSelection: { ja: "調整したいオブジェクトを選択してください。", en: "Select the objects you want to adjust." },
            noTarget: { ja: "色を調整できるオブジェクトが見つかりません。", en: "No objects with adjustable colors were found." }
        }
    };

    /**
     * ラベル（ja/en）を現在の UI 言語の文字列にする
     * @param {object} labelSet - ja/en を持つラベル
     * @returns {string} 現在の言語の文字列
     */
    function getLabel(labelSet) {
        return (labelSet && labelSet[uiLang]) || "";
    }

    /**
     * 項目名にコロンを付ける（日本語は全角、英語は半角）
     * @param {object} labelSet - ja/en を持つラベル
     * @returns {string} コロン付きの項目名
     */
    function labelText(labelSet) {
        return getLabel(labelSet) + (uiLang === "ja" ? "：" : ":");
    }

    // =========================================
    // ユーティリティ / Utilities
    // =========================================

    /**
     * 値を範囲内に収める
     * @param {number} value - 対象の値
     * @param {number} minValue - 下限値
     * @param {number} maxValue - 上限値
     * @returns {number} 範囲内に収めた値
     */
    function clampValue(value, minValue, maxValue) {
        if (value < minValue) return minValue;
        if (value > maxValue) return maxValue;
        return value;
    }

    // =========================================
    // カラー処理 / Color handling
    // =========================================

    /**
     * カラーを複製する（対応していないカラーモデルは null）
     * @param {object} sourceColor - 複製元のカラー
     * @returns {object|null} 複製したカラー
     */
    function cloneColor(sourceColor) {
        if (!sourceColor) return null;
        var colorModel = COLOR_MODELS[sourceColor.typename];
        if (!colorModel) return null;

        var clonedColor = colorModel.createColor();
        for (var i = 0; i < colorModel.channels.length; i++) {
            clonedColor[colorModel.channels[i]] = sourceColor[colorModel.channels[i]];
        }
        for (var j = 0; j < colorModel.references.length; j++) {
            clonedColor[colorModel.references[j]] = sourceColor[colorModel.references[j]];
        }
        return clonedColor;
    }

    /**
     * カラー値1つを明るく／暗くする
     * @param {number} channelValue - 元のカラー値
     * @param {number} ratio - 調整量の比率（正で明るく、負で暗く）
     * @param {number} maxValue - カラー値の最大値
     * @param {boolean} additive - 加法混色かどうか
     * @returns {number} 調整後のカラー値
     */
    function shiftChannelValue(channelValue, ratio, maxValue, additive) {
        /* 加法混色は明るくするほど最大値へ、減法混色は明るくするほど 0 へ寄せる / Lightening moves toward max for additive models, toward 0 for subtractive ones */
        var movesToMax = ((ratio > 0) === additive);
        var amount = Math.abs(ratio);
        return movesToMax ? channelValue + (maxValue - channelValue) * amount : channelValue * (1 - amount);
    }

    /**
     * 明るさを調整したカラーを作る
     * @param {object} baseColor - 元のカラー
     * @param {number} ratio - 調整量の比率（-1〜1）
     * @returns {object} 調整後のカラー（未対応のカラーモデルは元のカラー）
     */
    function buildAdjustedColor(baseColor, ratio) {
        var colorModel = COLOR_MODELS[baseColor.typename];
        if (!colorModel) return baseColor;

        var adjustedColor = cloneColor(baseColor);
        for (var i = 0; i < colorModel.channels.length; i++) {
            var channelName = colorModel.channels[i];
            adjustedColor[channelName] = shiftChannelValue(baseColor[channelName], ratio, colorModel.maxValue, colorModel.additive);
        }
        return adjustedColor;
    }

    // =========================================
    // 調整対象 / Adjustment targets
    // =========================================

    /**
     * @typedef {object} AdjustTarget
     * @property {PageItem} item - 対象のオブジェクト
     * @property {object|null} fillColor - 元の塗りのカラー
     * @property {object|null} strokeColor - 元の線のカラー
     */

    /**
     * 調整対象を配列に加える（塗りも線も調整できない場合は加えない）
     * @param {Array<AdjustTarget>} adjustTargets - 収集先の配列
     * @param {PageItem} pageItem - 対象のオブジェクト
     * @param {object} fillColor - 元の塗りのカラー
     * @param {object} strokeColor - 元の線のカラー
     * @returns {void}
     */
    function addAdjustTarget(adjustTargets, pageItem, fillColor, strokeColor) {
        var adjustTarget = {
            item: pageItem,
            fillColor: cloneColor(fillColor),
            strokeColor: cloneColor(strokeColor)
        };
        if (adjustTarget.fillColor || adjustTarget.strokeColor) adjustTargets.push(adjustTarget);
    }

    /**
     * 選択範囲を走査して、色を調整できるオブジェクトと元のカラーを集める
     * @param {Array} pageItems - 走査するオブジェクトの配列
     * @param {Array<AdjustTarget>} adjustTargets - 収集先の配列
     * @returns {void}
     */
    function collectAdjustTargets(pageItems, adjustTargets) {
        for (var i = 0; i < pageItems.length; i++) {
            var pageItem = pageItems[i];
            if (pageItem.typename === "GroupItem") {
                collectAdjustTargets(pageItem.pageItems, adjustTargets);
            } else if (pageItem.typename === "CompoundPathItem") {
                collectAdjustTargets(pageItem.pathItems, adjustTargets);
            } else if (pageItem.typename === "PathItem") {
                addAdjustTarget(adjustTargets, pageItem,
                    pageItem.filled ? pageItem.fillColor : null,
                    pageItem.stroked ? pageItem.strokeColor : null);
            } else if (pageItem.typename === "TextFrame") {
                /* 文字ごとに書式が異なるとカラーを取得できないことがある / Character attributes can fail on mixed formatting */
                try {
                    var textAttributes = pageItem.textRange.characterAttributes;
                    addAdjustTarget(adjustTargets, pageItem, textAttributes.fillColor, textAttributes.strokeColor);
                } catch (e) {}
            }
        }
    }

    /**
     * オブジェクトに塗りまたは線のカラーを設定する
     * @param {PageItem} pageItem - 対象のオブジェクト
     * @param {string} colorProperty - "fillColor" または "strokeColor"
     * @param {object} color - 設定するカラー
     * @returns {void}
     */
    function applyColor(pageItem, colorProperty, color) {
        if (pageItem.typename === "TextFrame") pageItem.textRange.characterAttributes[colorProperty] = color;
        else pageItem[colorProperty] = color;
    }

    /**
     * すべての調整対象に、指定した調整量を適用する
     * @param {Array<AdjustTarget>} adjustTargets - 調整対象
     * @param {number} percent - 調整量（%）
     * @returns {void}
     */
    function applyAdjustment(adjustTargets, percent) {
        var ratio = percent / 100;
        for (var i = 0; i < adjustTargets.length; i++) {
            var adjustTarget = adjustTargets[i];
            if (adjustTarget.fillColor) applyColor(adjustTarget.item, "fillColor", buildAdjustedColor(adjustTarget.fillColor, ratio));
            if (adjustTarget.strokeColor) applyColor(adjustTarget.item, "strokeColor", buildAdjustedColor(adjustTarget.strokeColor, ratio));
        }
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * スライダーの行を作る
     * @param {Window} adjustDialog - 追加先のダイアログ
     * @returns {Slider} 追加したスライダー
     */
    function buildSliderRow(adjustDialog) {
        var sliderRow = adjustDialog.add("group");
        setupRow(sliderRow, "fill");

        sliderRow.add("statictext", undefined, getLabel(LABELS.sliderEnd.darker));
        var brightnessSlider = sliderRow.add("slider", undefined, DEFAULT_PERCENT, SLIDER_MIN_PERCENT, SLIDER_MAX_PERCENT);
        brightnessSlider.preferredSize.width = SLIDER_WIDTH;
        brightnessSlider.alignment = ["fill", "center"];
        brightnessSlider.helpTip = getLabel(LABELS.tooltip.slider);
        sliderRow.add("statictext", undefined, getLabel(LABELS.sliderEnd.lighter));

        return brightnessSlider;
    }

    /**
     * 調整量の入力欄の行を作る
     * @param {Window} adjustDialog - 追加先のダイアログ
     * @returns {EditText} 追加した入力欄
     */
    function buildPercentRow(adjustDialog) {
        var percentRow = adjustDialog.add("group");
        setupRow(percentRow, "center", UNIT_SPACING);

        var amountLabel = percentRow.add("statictext", undefined, labelText(LABELS.fieldLabel.amount));
        amountLabel.helpTip = getLabel(LABELS.tooltip.percentInput);

        var percentInput = percentRow.add("edittext", undefined, String(DEFAULT_PERCENT));
        percentInput.characters = PERCENT_INPUT_CHARACTERS;
        percentInput.helpTip = getLabel(LABELS.tooltip.percentInput);
        percentRow.add("statictext", undefined, "%");

        return percentInput;
    }

    /**
     * ボタンの行を作る（左：リセット／右：キャンセル・OK）
     * @param {Window} adjustDialog - 追加先のダイアログ
     * @returns {Button} 追加したリセットボタン
     */
    function buildButtonRow(adjustDialog) {
        var btnRowGroup = adjustDialog.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.margins = [0, BUTTON_ROW_TOP_MARGIN, 0, 0];
        btnRowGroup.alignment = ["fill", "bottom"];

        var btnLeftGroup = btnRowGroup.add("group");
        btnLeftGroup.alignChildren = ["left", "center"];
        var btnReset = btnLeftGroup.add("button", undefined, getLabel(LABELS.button.reset));
        btnReset.helpTip = getLabel(LABELS.tooltip.reset);

        var spacer = btnRowGroup.add("group");
        spacer.alignment = ["fill", "fill"];
        spacer.minimumSize.width = 0;

        var btnRightGroup = btnRowGroup.add("group");
        btnRightGroup.alignChildren = ["right", "center"];
        btnRightGroup.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });
        btnRightGroup.add("button", undefined, getLabel(LABELS.button.ok), { name: "ok" });

        return btnReset;
    }

    /**
     * 調整用ダイアログを表示し、プレビューしながら調整量を決める
     * @param {Array<AdjustTarget>} adjustTargets - 調整対象
     * @returns {void}
     */
    function showAdjustDialog(adjustTargets) {
        var adjustDialog = new Window("dialog", getLabel(LABELS.dialog.title) + " " + SCRIPT_VERSION);
        setupWindow(adjustDialog);

        var brightnessSlider = buildSliderRow(adjustDialog);
        var percentInput = buildPercentRow(adjustDialog);
        var btnReset = buildButtonRow(adjustDialog);

        /* プレビューを適用済みかどうか / Whether a preview is currently applied */
        var hasPreview = false;

        /* 現在プレビュー中の調整量（%）/ Adjustment amount currently previewed (%) */
        var appliedPercent = null;

        /**
         * 入力値を丸めてプレビューを更新する
         * @param {number} percent - 調整量（%）
         * @returns {void}
         */
        function updatePreview(percent) {
            var nextPercent = Math.round(clampValue(percent, INPUT_MIN_PERCENT, INPUT_MAX_PERCENT));
            percentInput.text = nextPercent;
            brightnessSlider.value = clampValue(nextPercent, SLIDER_MIN_PERCENT, SLIDER_MAX_PERCENT);

            /* 同じ値なら適用し直さない（shift 併用のドラッグで取り消しを繰り返さない）/ Skip re-applying the same amount */
            if (hasPreview && nextPercent === appliedPercent) return;

            /* 履歴を汚さないよう、2回目以降は直前のプレビューを取り消してから適用する / Undo the previous preview so the history keeps a single step */
            if (hasPreview) app.undo();
            applyAdjustment(adjustTargets, nextPercent);
            appliedPercent = nextPercent;
            hasPreview = true;
            app.redraw();
        }

        brightnessSlider.onChanging = function() {
            var sliderPercent = brightnessSlider.value;
            /* shift 併用時は10%刻みにスナップ / Snap to 10% steps while shift is held */
            if (ScriptUI.environment.keyboardState.shiftKey) {
                sliderPercent = Math.round(sliderPercent / SLIDER_SHIFT_STEP) * SLIDER_SHIFT_STEP;
            }
            updatePreview(sliderPercent);
        };

        percentInput.onChange = function() {
            var inputPercent = Number(percentInput.text);
            updatePreview(isNaN(inputPercent) ? DEFAULT_PERCENT : inputPercent);
        };
        changeValueByArrowKey(percentInput, true, function() {
            updatePreview(Number(percentInput.text));
        }, INPUT_MIN_PERCENT);

        btnReset.onClick = function() {
            updatePreview(DEFAULT_PERCENT);
        };

        adjustDialog.addEventListener("keydown", function(event) {
            if (event.keyName !== RESET_SHORTCUT_KEY) return;
            updatePreview(DEFAULT_PERCENT);
            event.preventDefault();
        });

        percentInput.active = true;

        /* キャンセル（Esc を含む）ではプレビューを破棄して元に戻す / Cancel discards the preview */
        if (adjustDialog.show() !== 1 && hasPreview) {
            app.undo();
            app.redraw();
        }
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択範囲を集めて調整用ダイアログを開く
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel(LABELS.alert.noDocument));
            return;
        }

        var selectedItems = app.activeDocument.selection;
        if (!selectedItems || selectedItems.length === 0) {
            alert(getLabel(LABELS.alert.noSelection));
            return;
        }

        var adjustTargets = [];
        collectAdjustTargets(selectedItems, adjustTargets);
        if (adjustTargets.length === 0) {
            alert(getLabel(LABELS.alert.noTarget));
            return;
        }

        showAdjustDialog(adjustTargets);
    }

    main();

})();
