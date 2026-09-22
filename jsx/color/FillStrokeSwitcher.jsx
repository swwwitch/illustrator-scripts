#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したオブジェクトの塗りと線を入れ替えたり、一方をもう一方へ移したりします。
塗り↔線、塗り→線、線→塗りの3モードに対応します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/FillStrokeSwitcher.md

note記事も参照してください。
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
var SCRIPT_VERSION  = "v1.1.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-23";                             /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/FillStrokeSwitcher.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/FillStrokeSwitcher.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/shibumi/n/n5229b4357dd3"; /* 紹介記事 / article URL */

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

    /**
     * 表示言語を判定する
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = getCurrentLang();

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

    /**
     * LABELS からドット区切りのパスで表示言語のテキストを取り出す
     * @param {string} labelPath - "radio.swap" のようなドット区切りのキー
     * @returns {string} 表示言語のテキスト（見つからない場合は labelPath をそのまま返す）
     */
    function getLabel(labelPath) {
        var labelPathKeys = labelPath.split(".");
        var labelNode = LABELS;
        for (var i = 0; i < labelPathKeys.length; i++) {
            labelNode = labelNode[labelPathKeys[i]];
            if (!labelNode) return labelPath;
        }
        return labelNode[uiLang] || labelNode.en || labelPath;
    }

    /**
     * コロン付きの項目名を返す（日本語は全角、英語は半角）
     * @param {string} labelPath - ラベルのパス
     * @returns {string} コロン付きの項目名
     */
    function labelText(labelPath) {
        return getLabel(labelPath) + (uiLang === "ja" ? "：" : ":");
    }

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
            messageLines.push(getLabel('alert.pathFailures') + ': ' + processStats.pathFailureCount);
        }
        if (processStats.textFailureCount > 0) {
            messageLines.push(getLabel('alert.textFailures') + ': ' + processStats.textFailureCount);
        }
        if (processStats.selectionRestoreFailureCount > 0) {
            messageLines.push(getLabel('alert.selectionRestoreFailures') + ': ' + processStats.selectionRestoreFailureCount);
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
        var btnRowGroup = modeDialog.add('group');
        btnRowGroup.orientation = 'row';
        btnRowGroup.alignment = ['fill', 'center'];
        btnRowGroup.alignChildren = ['fill', 'center'];

        var btnLeftGroup = btnRowGroup.add('group');
        btnLeftGroup.orientation = 'row';
        btnLeftGroup.alignment = ['left', 'center'];
        var previewCheckbox = btnLeftGroup.add('checkbox', undefined, getLabel('checkbox.preview'));
        previewCheckbox.helpTip = getLabel('tooltip.preview');
        previewCheckbox.onClick = refreshPreview;

        var spacer = btnRowGroup.add('group');
        spacer.alignment = ['fill', 'center'];

        var btnRightGroup = btnRowGroup.add('group');
        btnRightGroup.orientation = 'row';
        btnRightGroup.alignment = ['right', 'center'];
        btnRightGroup.add('button', undefined, getLabel('button.cancel'), { name: 'cancel' });
        btnRightGroup.add('button', undefined, getLabel('button.ok'), { name: 'ok' });

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
