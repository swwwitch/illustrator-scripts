#target illustrator
#targetengine "FillStrokeSwitcherEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したオブジェクトの塗りと線を入れ替えたり、一方をもう一方へ移したりします。
塗り↔線、塗り→線、線→塗りの3モードに対応します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/FillStrokeSwitcher.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n81ee3a9e09b4

参照（しぶやみゃむ さんの記事）：
https://note.com/shibumi/n/n5229b4357dd3

### Overview

Swaps the fill and stroke of the selected objects, or moves one into the other.
Three modes are available: fill ↔ stroke, fill → stroke, and stroke → fill.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/FillStrokeSwitcher.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "FillStrokeSwitcher";           /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.4";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-03-25";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-29";                   /* 更新日 / last updated */

var SCRIPT_README_JA     = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/FillStrokeSwitcher.md"; /* README（日本語） */
var SCRIPT_README_EN     = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/FillStrokeSwitcher.md"; /* README (English) */
var SCRIPT_ARTICLE_URL   = "https://note.com/dtp_tranist/n/n81ee3a9e09b4"; /* 紹介記事 / article URL */
var SCRIPT_REFERENCE_URL = "https://note.com/shibumi/n/n5229b4357dd3";     /* 参照記事（しぶやみゃむ） / reference article */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // レイアウト / Layout
    // =========================================

    var PANEL_MARGINS = [15, 20, 15, 10];
    var PANEL_SPACING = 8;

    /**
     * パネルに共通のレイアウトを適用する
     * @param {Panel} targetPanel - 対象のパネル
     * @param {number} [spacing] - 子の間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function setupPanel(targetPanel, spacing) {
        targetPanel.orientation = "column";
        targetPanel.alignChildren = ['fill', 'top'];
        targetPanel.alignment = "fill";
        targetPanel.margins = PANEL_MARGINS;
        targetPanel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    // =========================================
    // 処理モード / Processing modes
    // =========================================

    var MODE_SWAP = "swap";
    var MODE_FILL_TO_STROKE = "fillToStroke";
    var MODE_STROKE_TO_FILL = "strokeToFill";
    var MODE_SWAP_BETWEEN = "swapBetween";
    var MODE_FILL_NONE = "fillNone";
    var MODE_STROKE_NONE = "strokeNone";
    var MODE_FILL_AND_STROKE_NONE = "fillAndStrokeNone";

    // =========================================
    // ローカライズ / Localization
    // =========================================

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // ローカライズ（再利用パーツ） / Localization (reusable)
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

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "塗りと線の調整", en: "Fill and Stroke Adjustments" }
        },
        panel: {
            convert: { ja: "変換", en: "Convert" },
            erase: { ja: "消去", en: "Erase" }
        },
        radio: {
            swap: { ja: "塗り↔線", en: "Fill ↔ Stroke" },
            fillToStroke: { ja: "塗り→線", en: "Fill → Stroke" },
            strokeToFill: { ja: "線→塗り", en: "Stroke → Fill" },
            swapBetween: { ja: "2つのオブジェクト間で交換", en: "Swap Between 2 Objects" },
            fillNone: { ja: "塗りを消去", en: "Erase Fill" },
            strokeNone: { ja: "線を消去", en: "Erase Stroke" },
            fillStrokeNone: { ja: "塗りと線を消去", en: "Erase Fill and Stroke" }
        },
        checkbox: {
            preview: { ja: "プレビュー", en: "Preview" }
        },
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" }
        },
        tooltip: {
            swap: { ja: "選択オブジェクトごとに、塗りと線を入れ替えます。", en: "Swaps fill and stroke within each selected object." },
            fillToStroke: {
                ja: "塗りの色を線に適用します。必要に応じて線幅を補完します。",
                en: "Applies the fill color to the stroke. Adds a stroke width when needed."
            },
            strokeToFill: { ja: "線の色を塗りに適用します。", en: "Applies the stroke color to the fill." },
            swapBetween: {
                ja: "選択した2つのオブジェクトの塗りと線の見た目を交換します。",
                en: "Swaps the fill and stroke appearance between the two selected objects."
            },
            fillNone: { ja: "選択オブジェクトの塗りをなしにします。", en: "Removes the fill from selected objects." },
            strokeNone: { ja: "選択オブジェクトの線をなしにします。", en: "Removes the stroke from selected objects." },
            fillStrokeNone: {
                ja: "選択オブジェクトの塗りと線をどちらもなしにします。",
                en: "Removes both fill and stroke from selected objects."
            },
            preview: {
                ja: "結果を一時的に表示します。キャンセル時は元に戻します。",
                en: "Temporarily shows the result. Restores the original appearance when canceled."
            }
        },
        alert: {
            noSelection: { ja: "オブジェクトを選択してください", en: "Please select at least one object." },
            tooManyObjects: { ja: "オブジェクトは 2 つまでにしてください", en: "Please select up to two objects." },
            noDocument: { ja: "ドキュメントが開かれていません", en: "No document is open." },
            pathFailures: { ja: "パス処理の失敗", en: "Path processing failures" },
            textFailures: { ja: "テキスト処理の失敗", en: "Text processing failures" },
            selectionRestoreFailures: { ja: "選択の復元失敗", en: "Selection restoration failures" },
            details: { ja: "詳細", en: "Details" }
        }
    };

    // =========================================
    // 失敗の集計 / Failure stats
    // =========================================

    /**
     * 処理結果の集計オブジェクトを作る
     * @returns {{pathFailureCount: number, textFailureCount: number, selectionRestoreFailureCount: number, failureDetails: string[]}} 空の集計
     */
    function createProcessStats() {
        return {
            pathFailureCount: 0,
            textFailureCount: 0,
            selectionRestoreFailureCount: 0,
            failureDetails: []
        };
    }

    /**
     * 失敗の詳細を集計に追加する（最大 8 件）
     * @param {Object|null} processStats - 集計（null なら何もしない）
     * @param {string} category - 失敗の分類（"Path" など）
     * @param {PageItem|null} failedItem - 失敗したオブジェクト
     * @param {Error} error - 発生した例外
     * @returns {void}
     */
    function addFailureDetail(processStats, category, failedItem, error) {
        if (!processStats || !processStats.failureDetails) return;
        if (processStats.failureDetails.length >= 8) return;

        var itemTypeName = 'Unknown';
        var itemName = '';
        var errorMessage = 'Unknown error';
        try {
            if (failedItem && failedItem.typename) itemTypeName = failedItem.typename;
            if (failedItem && failedItem.name) itemName = String(failedItem.name);
            if (error && error.message) {
                errorMessage = String(error.message);
            } else if (error) {
                errorMessage = String(error);
            }
        } catch (e) { /* 削除済みのオブジェクトはプロパティを読めない / removed items cannot be read */ }

        var detail = category + ': ' + itemTypeName;
        if (itemName !== '') detail += ' [' + itemName + ']';
        detail += ' - ' + errorMessage;

        processStats.failureDetails.push(detail);
    }

    /**
     * 失敗を数えて詳細を記録する
     * @param {Object|null} processStats - 集計（null なら何もしない）
     * @param {string} counterKey - 増やすカウンター（"pathFailureCount" など）
     * @param {string} category - 失敗の分類（"Path" など）
     * @param {PageItem|null} failedItem - 失敗したオブジェクト
     * @param {Error} error - 発生した例外
     * @returns {void}
     */
    function recordFailure(processStats, counterKey, category, failedItem, error) {
        if (processStats) processStats[counterKey]++;
        addFailureDetail(processStats, category, failedItem, error);
    }

    /**
     * 集計から、失敗を知らせるメッセージを作る
     * @param {Object} processStats - 集計
     * @returns {string} メッセージ（失敗が無ければ空文字）
     */
    function buildFailureMessage(processStats) {
        var messageLines = [];
        if (processStats.pathFailureCount > 0) {
            messageLines.push(labelValueText('alert.pathFailures', processStats.pathFailureCount));
        }
        if (processStats.textFailureCount > 0) {
            messageLines.push(labelValueText('alert.textFailures', processStats.textFailureCount));
        }
        if (processStats.selectionRestoreFailureCount > 0) {
            messageLines.push(labelValueText('alert.selectionRestoreFailures', processStats.selectionRestoreFailureCount));
        }
        if (processStats.failureDetails.length > 0) {
            messageLines.push('');
            messageLines.push(labelText('alert.details'));
            for (var i = 0; i < processStats.failureDetails.length; i++) {
                messageLines.push('- ' + processStats.failureDetails[i]);
            }
        }
        return messageLines.join('\n');
    }

    // =========================================
    // 色処理 / Color handling
    // =========================================

    /**
     * 色を複製する（未対応の種類は null）
     * @param {Color} color - 元の色
     * @returns {Color|null} 複製した色
     */
    function cloneColor(color) {
        if (!color) return null;

        switch (color.typename) {

            case "RGBColor":
                var rgbColor = new RGBColor();
                rgbColor.red = color.red;
                rgbColor.green = color.green;
                rgbColor.blue = color.blue;
                return rgbColor;

            case "CMYKColor":
                var cmykColor = new CMYKColor();
                cmykColor.cyan = color.cyan;
                cmykColor.magenta = color.magenta;
                cmykColor.yellow = color.yellow;
                cmykColor.black = color.black;
                return cmykColor;

            case "GrayColor":
                var grayColor = new GrayColor();
                grayColor.gray = color.gray;
                return grayColor;

            case "SpotColor":
                var spotColor = new SpotColor();
                spotColor.spot = color.spot;
                spotColor.tint = color.tint;
                return spotColor;

            case "GradientColor":
                var gradientColor = new GradientColor();
                gradientColor.gradient = color.gradient;
                gradientColor.angle = color.angle;
                gradientColor.length = color.length;
                gradientColor.origin = color.origin;
                gradientColor.matrix = color.matrix;
                return gradientColor;

            case "NoColor":
                return new NoColor();

            default:
                return null;
        }
    }

    /**
     * 「なし」の色を作る
     * @returns {NoColor} 「なし」の色
     */
    function createNoColor() {
        return new NoColor();
    }

    /**
     * テキストに使える色（「なし」以外）かを判定する
     * @param {Color|null} color - 判定する色
     * @returns {boolean} 使える色なら true
     */
    function isUsableTextColor(color) {
        return color && color.typename && color.typename !== "NoColor";
    }

    /**
     * TextRange の塗りの色を読む
     * @param {TextRange} textRange - 対象の文字範囲
     * @returns {Color|null} 塗りの色（読めないときは null）
     */
    function getTextRangeFillColor(textRange) {
        try {
            return textRange.characterAttributes.fillColor;
        } catch (e) {
            return null;
        }
    }

    /**
     * TextRange の線の色を読む
     * @param {TextRange} textRange - 対象の文字範囲
     * @returns {Color|null} 線の色（読めないときは null）
     */
    function getTextRangeStrokeColor(textRange) {
        try {
            return textRange.characterAttributes.strokeColor;
        } catch (e) {
            return null;
        }
    }

    /**
     * TextRange が有効な塗りを持つか
     * @param {TextRange} textRange - 対象の文字範囲
     * @returns {boolean} 持っていれば true
     */
    function hasTextRangeFill(textRange) {
        return isUsableTextColor(getTextRangeFillColor(textRange));
    }

    /**
     * TextRange が有効な線を持つか
     * @param {TextRange} textRange - 対象の文字範囲
     * @returns {boolean} 持っていれば true
     */
    function hasTextRangeStroke(textRange) {
        return isUsableTextColor(getTextRangeStrokeColor(textRange));
    }

    /**
     * テキストの1文字ずつ（文字が取れなければ全体）に処理を適用する
     * @param {TextFrame} textFrame - 対象のテキスト
     * @param {Function} applyToRange - TextRange を受け取る関数
     * @returns {void}
     */
    function forEachCharacterRange(textFrame, applyToRange) {
        var textRange = textFrame.textRange;
        var textCharacters = null;
        try { textCharacters = textRange.characters; } catch (eC) { textCharacters = null; }

        if (textCharacters && textCharacters.length > 0) {
            for (var i = 0; i < textCharacters.length; i++) {
                applyToRange(textCharacters[i]);
            }
        } else {
            applyToRange(textRange);
        }
    }

    /**
     * TextRange に処理モードに応じた塗り・線の変更を適用する
     * @param {TextRange} textRange - 対象の文字範囲
     * @param {string} mode - 処理モード（MODE_*）
     * @returns {void}
     */
    function applyTextFillAndStrokeToRange(textRange, mode) {
        var charAttributes = textRange.characterAttributes;
        var hasFill = hasTextRangeFill(textRange);
        var hasStroke = hasTextRangeStroke(textRange);

        var fillCopy = hasFill ? cloneColor(getTextRangeFillColor(textRange)) : null;
        var strokeCopy = hasStroke ? cloneColor(getTextRangeStrokeColor(textRange)) : null;

        switch (mode) {
            case MODE_FILL_TO_STROKE:
                if (hasFill && fillCopy) {
                    charAttributes.strokeColor = fillCopy;
                    if (!hasStroke || !charAttributes.strokeWeight || charAttributes.strokeWeight <= 0) {
                        charAttributes.strokeWeight = 1;
                    }
                }
                break;

            case MODE_STROKE_TO_FILL:
                if (hasStroke && strokeCopy) {
                    charAttributes.fillColor = strokeCopy;
                }
                break;

            case MODE_FILL_NONE:
                charAttributes.fillColor = createNoColor();
                break;

            case MODE_STROKE_NONE:
                charAttributes.strokeColor = createNoColor();
                break;

            case MODE_FILL_AND_STROKE_NONE:
                charAttributes.fillColor = createNoColor();
                charAttributes.strokeColor = createNoColor();
                break;

            case MODE_SWAP:
            default:
                if (!hasFill && !hasStroke) {
                    return;
                }

                if (hasStroke && strokeCopy) {
                    charAttributes.fillColor = strokeCopy;
                } else {
                    charAttributes.fillColor = createNoColor();
                }

                if (hasFill && fillCopy) {
                    charAttributes.strokeColor = fillCopy;
                    if (!hasStroke || !charAttributes.strokeWeight || charAttributes.strokeWeight <= 0) {
                        charAttributes.strokeWeight = 1;
                    }
                } else {
                    charAttributes.strokeColor = createNoColor();
                }
                break;
        }
    }

    // =========================================
    // 処理モードの適用 / Applying the mode
    // =========================================

    /**
     * パスに処理モードを適用する（失敗は集計に記録）
     * @param {PathItem} pathItem - 対象のパス
     * @param {string} mode - 処理モード（MODE_*）
     * @param {Object|null} processStats - 集計（プレビュー時は null）
     * @returns {void}
     */
    function applyPathFillStrokeMode(pathItem, mode, processStats) {
        try {
            var hasFill = pathItem.filled;
            var hasStroke = pathItem.stroked;

            var fillCopy = hasFill ? cloneColor(pathItem.fillColor) : null;
            var strokeCopy = hasStroke ? cloneColor(pathItem.strokeColor) : null;

            switch (mode) {
                case MODE_FILL_TO_STROKE:
                    if (hasFill && fillCopy) {
                        pathItem.stroked = true;
                        pathItem.strokeColor = fillCopy;
                        if (!hasStroke) {
                            pathItem.strokeWidth = 1;
                        }
                    }
                    break;

                case MODE_STROKE_TO_FILL:
                    if (hasStroke && strokeCopy) {
                        pathItem.filled = true;
                        pathItem.fillColor = strokeCopy;
                    }
                    break;

                case MODE_FILL_NONE:
                    pathItem.filled = false;
                    pathItem.fillColor = createNoColor();
                    break;

                case MODE_STROKE_NONE:
                    pathItem.stroked = false;
                    pathItem.strokeColor = createNoColor();
                    break;

                case MODE_FILL_AND_STROKE_NONE:
                    pathItem.filled = false;
                    pathItem.stroked = false;
                    pathItem.fillColor = createNoColor();
                    pathItem.strokeColor = createNoColor();
                    break;

                case MODE_SWAP:
                default:
                    if (!hasFill && !hasStroke) break;
                    pathItem.filled = hasStroke;
                    pathItem.fillColor = (hasStroke && strokeCopy) ? strokeCopy : createNoColor();
                    pathItem.stroked = hasFill;
                    if (hasFill && fillCopy) {
                        pathItem.strokeColor = fillCopy;
                        if (!hasStroke || !pathItem.strokeWidth || pathItem.strokeWidth <= 0) {
                            pathItem.strokeWidth = 1;
                        }
                    } else {
                        pathItem.strokeColor = createNoColor();
                    }
                    break;
            }
        } catch (e) {
            recordFailure(processStats, 'pathFailureCount', 'Path', pathItem, e);
        }
    }

    /**
     * テキストの各文字に処理モードを適用する（失敗は集計に記録）
     * @param {TextFrame} textFrame - 対象のテキスト
     * @param {string} mode - 処理モード（MODE_*）
     * @param {Object|null} processStats - 集計（プレビュー時は null）
     * @returns {void}
     */
    function applyTextFillStrokeMode(textFrame, mode, processStats) {
        try {
            forEachCharacterRange(textFrame, function (textRange) {
                applyTextFillAndStrokeToRange(textRange, mode);
            });
        } catch (e) {
            recordFailure(processStats, 'textFailureCount', 'Text', textFrame, e);
        }
    }

    /**
     * 選択オブジェクトに処理モードを適用する（グループ・複合パスは再帰）
     * @param {PageItem[]} targetItems - 対象のオブジェクト
     * @param {string} mode - 処理モード（MODE_*）
     * @param {Object|null} processStats - 集計（プレビュー時は null）
     * @returns {void}
     */
    function processItems(targetItems, mode, processStats) {
        for (var i = 0; i < targetItems.length; i++) {
            var targetItem = targetItems[i];

            switch (targetItem.typename) {

                case "GroupItem":
                    processItems(targetItem.pageItems, mode, processStats);
                    break;

                case "PathItem":
                    applyPathFillStrokeMode(targetItem, mode, processStats);
                    break;

                case "CompoundPathItem":
                    processItems(targetItem.pathItems, mode, processStats);
                    break;

                case "TextFrame":
                    applyTextFillStrokeMode(targetItem, mode, processStats);
                    break;
            }
        }
    }

    // =========================================
    // 見た目の取得と適用 / Appearance capture and apply
    // =========================================

    /**
     * パスの塗り・線の見た目を控える
     * @param {PathItem} pathItem - 対象のパス
     * @returns {{filled: boolean, stroked: boolean, fill: Color, stroke: Color, strokeWidth: number}} 見た目
     */
    function capturePathAppearance(pathItem) {
        return {
            filled: pathItem.filled,
            stroked: pathItem.stroked,
            fill: pathItem.filled ? cloneColor(pathItem.fillColor) : null,
            stroke: pathItem.stroked ? cloneColor(pathItem.strokeColor) : null,
            strokeWidth: pathItem.strokeWidth
        };
    }

    /**
     * TextRange の塗り・線の見た目を控える
     * @param {TextRange} textRange - 対象の文字範囲
     * @returns {{filled: boolean, stroked: boolean, fill: Color, stroke: Color, strokeWidth: number}} 見た目
     */
    function captureTextRangeAppearance(textRange) {
        var hasFill = hasTextRangeFill(textRange);
        var hasStroke = hasTextRangeStroke(textRange);
        return {
            filled: hasFill,
            stroked: hasStroke,
            fill: hasFill ? cloneColor(getTextRangeFillColor(textRange)) : null,
            stroke: hasStroke ? cloneColor(getTextRangeStrokeColor(textRange)) : null,
            strokeWidth: textRange.characterAttributes.strokeWeight
        };
    }

    /**
     * パスに見た目を適用する
     * @param {PathItem} pathItem - 対象のパス
     * @param {Object} appearance - capturePathAppearance() などで控えた見た目
     * @returns {void}
     */
    function applyPathAppearance(pathItem, appearance) {
        pathItem.filled = appearance.filled;
        pathItem.fillColor = (appearance.filled && appearance.fill) ? appearance.fill : createNoColor();
        pathItem.stroked = appearance.stroked;
        if (appearance.stroked && appearance.stroke) {
            pathItem.strokeColor = appearance.stroke;
            if (appearance.strokeWidth && appearance.strokeWidth > 0) {
                pathItem.strokeWidth = appearance.strokeWidth;
            }
        } else {
            pathItem.strokeColor = createNoColor();
        }
    }

    /**
     * TextRange に見た目を適用する
     * @param {TextRange} textRange - 対象の文字範囲
     * @param {Object} appearance - captureTextRangeAppearance() などで控えた見た目
     * @returns {void}
     */
    function applyTextRangeAppearance(textRange, appearance) {
        var charAttributes = textRange.characterAttributes;
        charAttributes.fillColor = (appearance.filled && appearance.fill) ? appearance.fill : createNoColor();
        if (appearance.stroked && appearance.stroke) {
            charAttributes.strokeColor = appearance.stroke;
            if (appearance.strokeWidth && appearance.strokeWidth > 0) {
                charAttributes.strokeWeight = appearance.strokeWidth;
            }
        } else {
            charAttributes.strokeColor = createNoColor();
        }
    }

    /**
     * オブジェクトの代表の見た目を1つ控える（複合パスは最初のパス、テキストは全体）
     * @param {PageItem} sourceItem - 対象のオブジェクト
     * @returns {Object} 見た目
     */
    function snapshotAppearance(sourceItem) {
        var typeName = sourceItem.typename;
        if (typeName === "PathItem") return capturePathAppearance(sourceItem);
        if (typeName === "CompoundPathItem") {
            if (!sourceItem.pathItems || sourceItem.pathItems.length === 0) {
                throw new Error("CompoundPathItem has no pathItems");
            }
            return capturePathAppearance(sourceItem.pathItems[0]);
        }
        if (typeName === "TextFrame") return captureTextRangeAppearance(sourceItem.textRange);
        throw new Error("Unsupported type for swap: " + typeName);
    }

    /**
     * 1つの見た目をオブジェクト全体に均一に適用する
     * @param {PageItem} targetItem - 対象のオブジェクト
     * @param {Object} appearance - 適用する見た目
     * @returns {void}
     */
    function applyAppearance(targetItem, appearance) {
        var typeName = targetItem.typename;
        if (typeName === "PathItem") {
            applyPathAppearance(targetItem, appearance);
            return;
        }
        if (typeName === "CompoundPathItem") {
            for (var i = 0; i < targetItem.pathItems.length; i++) {
                applyPathAppearance(targetItem.pathItems[i], appearance);
            }
            return;
        }
        if (typeName === "TextFrame") {
            forEachCharacterRange(targetItem, function (textRange) {
                applyTextRangeAppearance(textRange, appearance);
            });
            return;
        }
        throw new Error("Unsupported type for swap: " + typeName);
    }

    /**
     * 2つのオブジェクトの塗りと線の見た目を交換する（失敗は集計に記録）
     * @param {PageItem} itemA - 1つ目のオブジェクト
     * @param {PageItem} itemB - 2つ目のオブジェクト
     * @param {Object|null} processStats - 集計（プレビュー時は null）
     * @returns {void}
     */
    function swapAppearanceBetween(itemA, itemB, processStats) {
        try {
            var appearanceA = snapshotAppearance(itemA);
            var appearanceB = snapshotAppearance(itemB);
            applyAppearance(itemA, appearanceB);
            applyAppearance(itemB, appearanceA);
        } catch (e) {
            recordFailure(processStats, 'pathFailureCount', 'SwapBetween', itemA, e);
        }
    }

    // =========================================
    // プレビュー用スナップショット / Preview snapshot
    // =========================================

    /**
     * プレビューの取り消し用に、パス・テキスト1つの状態を控える
     * @param {PageItem} leafItem - パスまたはテキスト
     * @returns {Object|null} 控えた状態（対象外の種類は null）
     */
    function snapshotLeafForPreview(leafItem) {
        var typeName = leafItem.typename;
        if (typeName === "PathItem") {
            var pathSnapshot = capturePathAppearance(leafItem);
            pathSnapshot.kind = "path";
            pathSnapshot.item = leafItem;
            return pathSnapshot;
        }
        if (typeName === "TextFrame") {
            var textSnapshot = { kind: "text", item: leafItem, characters: [] };
            try {
                var textCharacters = leafItem.textRange.characters;
                for (var i = 0; i < textCharacters.length; i++) {
                    textSnapshot.characters.push(captureTextRangeAppearance(textCharacters[i]));
                }
            } catch (e) { /* 読めた文字までで止める / keep the characters read so far */ }
            return textSnapshot;
        }
        return null;
    }

    /**
     * 選択の中のパス・テキストを再帰的にたどって状態を控える
     * @param {PageItem[]} sourceItems - 対象のオブジェクト
     * @param {Object[]} leafSnapshots - 控えた状態を追加する配列
     * @returns {void}
     */
    function captureLeafSnapshots(sourceItems, leafSnapshots) {
        for (var i = 0; i < sourceItems.length; i++) {
            var sourceItem = sourceItems[i];
            switch (sourceItem.typename) {
                case "GroupItem":
                    captureLeafSnapshots(sourceItem.pageItems, leafSnapshots);
                    break;
                case "PathItem":
                case "TextFrame":
                    var leafSnapshot = snapshotLeafForPreview(sourceItem);
                    if (leafSnapshot) leafSnapshots.push(leafSnapshot);
                    break;
                case "CompoundPathItem":
                    captureLeafSnapshots(sourceItem.pathItems, leafSnapshots);
                    break;
            }
        }
    }

    /**
     * 選択全体の状態を控える
     * @param {PageItem[]} sourceItems - 対象のオブジェクト
     * @returns {Object[]} 控えた状態
     */
    function captureSelectionState(sourceItems) {
        var leafSnapshots = [];
        captureLeafSnapshots(sourceItems, leafSnapshots);
        return leafSnapshots;
    }

    /**
     * 控えた状態からパス・テキスト1つを元に戻す
     * @param {Object} leafSnapshot - snapshotLeafForPreview() の結果
     * @returns {void}
     */
    function restoreLeafSnapshot(leafSnapshot) {
        if (leafSnapshot.kind === "path") {
            applyPathAppearance(leafSnapshot.item, leafSnapshot);
            return;
        }
        if (leafSnapshot.kind === "text") {
            var textCharacters = null;
            try { textCharacters = leafSnapshot.item.textRange.characters; } catch (eC) { return; }
            if (!textCharacters) return;
            var pairCount = (textCharacters.length < leafSnapshot.characters.length) ? textCharacters.length : leafSnapshot.characters.length;
            for (var i = 0; i < pairCount; i++) {
                try { applyTextRangeAppearance(textCharacters[i], leafSnapshot.characters[i]); } catch (eR) { }
            }
        }
    }

    /**
     * 控えた状態から選択全体を元に戻す
     * @param {Object[]} leafSnapshots - captureSelectionState() の結果
     * @returns {void}
     */
    function restoreSelectionState(leafSnapshots) {
        for (var i = 0; i < leafSnapshots.length; i++) {
            if (!leafSnapshots[i]) continue;
            try { restoreLeafSnapshot(leafSnapshots[i]); } catch (e) { }
        }
    }

    /**
     * 選択を元のオブジェクトに戻す（失敗は集計に記録）
     * @param {Document} doc - 対象ドキュメント
     * @param {PageItem[]} originalItems - 元の選択
     * @param {Object} processStats - 集計
     * @returns {void}
     */
    function restoreSelectedItems(doc, originalItems, processStats) {
        var restoredItems = [];
        for (var i = 0; i < originalItems.length; i++) {
            if (originalItems[i]) restoredItems.push(originalItems[i]);
        }

        try {
            doc.selection = restoredItems;
        } catch (eR) {
            recordFailure(processStats, 'selectionRestoreFailureCount', 'Selection', null, eR);
        }
    }

    /**
     * プレビューとして選択中のオブジェクトに処理モードを適用する
     * @param {PageItem[]} targetItems - 対象のオブジェクト
     * @param {string} mode - 処理モード（MODE_*）
     * @returns {void}
     */
    function applyPreview(targetItems, mode) {
        if (mode === MODE_SWAP_BETWEEN) {
            if (targetItems.length === 2) {
                swapAppearanceBetween(targetItems[0], targetItems[1], null);
            }
            return;
        }
        processItems(targetItems, mode, null);
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // ダイアログの位置と不透明度（再利用パーツ） / Dialog position and opacity (reusable)
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

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
            if (!selectedItems || !selectedItems.length || !selectedItems[0].visibleBounds) return null;
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

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // ダイアログの位置と不透明度（再利用パーツ）ここまで / End of the reusable dialog position and opacity
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // ボタン行（再利用パーツ） / Button row (reusable)
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

    var BUTTON_ROW_TOP_MARGIN = 5; /* ボタン行の上の余白 / top margin of the button row */
    var BUTTON_ROW_SPACING = 10;   /* ボタンどうしの間隔 / spacing between buttons */

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

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // ボタン行（再利用パーツ）ここまで / End of the reusable button row
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    /**
     * ラジオボタンを1つだけ選択状態にする（パネルをまたぐので手動で排他にする）
     * @param {RadioButton} selectedRadio - 選択するラジオボタン
     * @param {Object[]} modeRadioEntries - {radio, mode} の一覧
     * @returns {void}
     */
    function setExclusiveRadio(selectedRadio, modeRadioEntries) {
        for (var i = 0; i < modeRadioEntries.length; i++) {
            modeRadioEntries[i].radio.value = (modeRadioEntries[i].radio === selectedRadio);
        }
    }

    /**
     * 選択中のラジオボタンから処理モードを返す
     * @param {Object[]} modeRadioEntries - {radio, mode} の一覧
     * @returns {string} 処理モード（どれも選ばれていなければ MODE_SWAP）
     */
    function getModeFromRadios(modeRadioEntries) {
        for (var i = 0; i < modeRadioEntries.length; i++) {
            if (modeRadioEntries[i].radio.value) return modeRadioEntries[i].mode;
        }
        return MODE_SWAP;
    }

    /**
     * ラジオボタンのクリックに排他選択とプレビュー更新をつなぐ
     * @param {Object[]} modeRadioEntries - {radio, mode} の一覧
     * @param {Function} refreshPreview - プレビューを更新する関数
     * @returns {void}
     */
    function bindModeRadioEvents(modeRadioEntries, refreshPreview) {
        /**
         * 1つのラジオボタンにクリック時の処理をつなぐ
         * @param {RadioButton} modeRadio - 対象のラジオボタン
         * @returns {void}
         */
        function bindExclusiveRadio(modeRadio) {
            modeRadio.onClick = function () {
                setExclusiveRadio(modeRadio, modeRadioEntries);
                refreshPreview();
            };
        }

        for (var i = 0; i < modeRadioEntries.length; i++) {
            bindExclusiveRadio(modeRadioEntries[i].radio);
        }
    }

    /**
     * 処理モードのラジオボタンを追加する（ラベルと tooltip は radio.<key> / tooltip.<key>）
     * @param {Panel} parentPanel - 追加先のパネル
     * @param {string} labelKey - LABELS のキー
     * @param {string} mode - このラジオボタンの処理モード（MODE_*）
     * @param {Object[]} modeRadioEntries - {radio, mode} を追加する一覧
     * @returns {RadioButton} 追加したラジオボタン
     */
    function addModeRadio(parentPanel, labelKey, mode, modeRadioEntries) {
        var modeRadio = parentPanel.add('radiobutton', undefined, getLabel('radio.' + labelKey));
        modeRadio.helpTip = getLabel('tooltip.' + labelKey);
        modeRadioEntries.push({ radio: modeRadio, mode: mode });
        return modeRadio;
    }

    /**
     * 処理モード選択ダイアログを表示する。キャンセル時はプレビューを取り消す
     * @param {PageItem[]} originalSelection - 元の選択
     * @returns {{mode: string, previewApplied: boolean}|null} 選んだモードとプレビュー適用済みか（キャンセル時は null）
     */
    function showModeDialog(originalSelection) {
        var leafSnapshots = captureSelectionState(originalSelection);
        var isPairSelected = (originalSelection && originalSelection.length === 2);

        var modeDialog = new Window('dialog', getLabel('dialog.title') + ' ' + SCRIPT_VERSION);
        modeDialog.orientation = 'column';
        modeDialog.alignChildren = ['fill', 'top'];

        var modePanelsGroup = modeDialog.add('group');
        modePanelsGroup.orientation = 'row';
        modePanelsGroup.alignChildren = ['fill', 'fill'];

        var convertPanel = modePanelsGroup.add('panel', undefined, getLabel('panel.convert'));
        setupPanel(convertPanel);

        var erasePanel = modePanelsGroup.add('panel', undefined, getLabel('panel.erase'));
        setupPanel(erasePanel);

        /* 2つのパネルにまたがるラジオボタン / Radio buttons spread across two panels */
        var modeRadioEntries = [];
        var swapRadio = addModeRadio(convertPanel, 'swap', MODE_SWAP, modeRadioEntries);
        addModeRadio(convertPanel, 'fillToStroke', MODE_FILL_TO_STROKE, modeRadioEntries);
        addModeRadio(convertPanel, 'strokeToFill', MODE_STROKE_TO_FILL, modeRadioEntries);
        var swapBetweenRadio = addModeRadio(convertPanel, 'swapBetween', MODE_SWAP_BETWEEN, modeRadioEntries);
        swapBetweenRadio.enabled = isPairSelected;
        addModeRadio(erasePanel, 'fillNone', MODE_FILL_NONE, modeRadioEntries);
        addModeRadio(erasePanel, 'strokeNone', MODE_STROKE_NONE, modeRadioEntries);
        addModeRadio(erasePanel, 'fillStrokeNone', MODE_FILL_AND_STROKE_NONE, modeRadioEntries);

        /**
         * プレビューをいったん元に戻し、ON なら選んだモードで掛け直す
         * @returns {void}
         */
        function refreshPreview() {
            restoreSelectionState(leafSnapshots);
            if (previewCheckbox.value) {
                applyPreview(originalSelection, getModeFromRadios(modeRadioEntries));
            }
            app.redraw();
        }

        bindModeRadioEvents(modeRadioEntries, refreshPreview);
        setExclusiveRadio(isPairSelected ? swapBetweenRadio : swapRadio, modeRadioEntries);

        /* ボタン行（左にプレビュー、右に OK／キャンセル） / Button row: preview on the left, OK/Cancel on the right */
        var buttonRow = addButtonRow(modeDialog);
        var previewCheckbox = buttonRow.leftGroup.add('checkbox', undefined, getLabel('checkbox.preview'));
        previewCheckbox.helpTip = getLabel('tooltip.preview');
        previewCheckbox.onClick = refreshPreview;
        var btnCancel = buttonRow.rightGroup.add('button', undefined, getLabel('button.cancel'), { name: 'cancel' });
        var btnOK = buttonRow.rightGroup.add('button', undefined, getLabel('button.ok'), { name: 'ok' });

        prepareDialogWindow(modeDialog, SCRIPT_NAME);
        if (modeDialog.show() !== 1) {
            restoreSelectionState(leafSnapshots);
            app.redraw();
            return null;
        }

        return {
            mode: getModeFromRadios(modeRadioEntries),
            previewApplied: previewCheckbox.value
        };
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * エントリポイント
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel('alert.noDocument'));
            return;
        }

        var doc = app.activeDocument;
        if (doc.selection.length === 0) {
            alert(getLabel('alert.noSelection'));
            return;
        }

        if (doc.selection.length > 2) {
            alert(getLabel('alert.tooManyObjects'));
            return;
        }

        var processStats = createProcessStats();
        var originalSelection = [];
        for (var i = 0; i < doc.selection.length; i++) {
            originalSelection.push(doc.selection[i]);
        }

        var dialogResult = showModeDialog(originalSelection);
        if (dialogResult === null) {
            return;
        }

        try {
            /* プレビューで適用済みなら掛け直さない / Skip when the preview already applied it */
            if (!dialogResult.previewApplied) {
                if (dialogResult.mode === MODE_SWAP_BETWEEN) {
                    swapAppearanceBetween(originalSelection[0], originalSelection[1], processStats);
                } else {
                    processItems(originalSelection, dialogResult.mode, processStats);
                }
            }
        } finally {
            restoreSelectedItems(doc, originalSelection, processStats);
        }

        var failureMessage = buildFailureMessage(processStats);
        if (failureMessage !== '') {
            alert(failureMessage);
        }
    }

    main();
})();
