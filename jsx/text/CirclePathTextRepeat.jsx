#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

円（パス）とテキストを1つずつ選択して実行すると、テキストを指定回数繰り返して、円を複製したパス上文字に変換します。
区切り文字はスペースまたは任意の文字（初期値は「•」）から選べます。

詳細は README を参照してください。

### Overview

With one circle and one text frame selected, repeats the text a given number of times and converts it into text on a copy of the circle.
The separator can be a space or any character you type ("•" by default).

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "CirclePathTextRepeat";         /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-06-12";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-07";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/CirclePathTextRepeat.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/CirclePathTextRepeat.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    var DEFAULT_REPEAT_COUNT    = 3;     /* 繰り返し数 / repeat count */
    var DEFAULT_SEPARATOR_CHAR  = "•";   /* 区切り文字（欧文 bullet）/ separator character (bullet) */
    var DEFAULT_SPACE_COUNT     = 1;     /* 区切り文字の前後に入れる半角スペース数 / half-width spaces on each side */
    var DEFAULT_SEPARATOR_SCALE = 100;   /* 区切り文字の比率（%）/ separator scale (%) */
    var DEFAULT_BASELINE_SHIFT  = 0;     /* 区切り文字のベースライン（環境設定のテキスト単位）/ separator baseline (preferences text unit) */
    var DEFAULT_CORRECTION      = 103;   /* 円周の補正率（%）/ circumference correction (%) */
    var DEFAULT_ROTATION        = 0;     /* 回転角度（度）/ rotation angle (degrees) */
    var FALLBACK_FONT_SIZE      = 12;    /* 元テキストのサイズを読めないときの代替（pt）/ fallback font size (pt) */

    // =========================================
    // Illustrator の制限 / Illustrator limits
    // =========================================

    var MIN_FONT_SIZE        = 0.1;      /* 文字サイズの下限（pt）/ minimum font size (pt) */
    var MAX_FONT_SIZE        = 1296;     /* 文字サイズの上限（pt）/ maximum font size (pt) */
    var MEASURE_FRAME_OFFSET = -100000;  /* 計測用フレームを置く画面外の座標 / off-canvas position of the measurement frame */

    // =========================================
    // レイアウト / Layout
    // =========================================

    var WINDOW_MARGINS      = 16;               /* ウィンドウ外周の余白 / window margin */
    var WINDOW_SPACING      = 12;               /* ウィンドウ内の要素間隔 / window spacing */
    var PANEL_MARGINS       = [16, 20, 16, 12]; /* パネル余白 [左,上,右,下] / panel margins */
    var PANEL_SPACING       = 8;                /* パネル内の要素間隔 / panel spacing */
    var BUTTON_ROW_MARGINS  = [0, 10, 0, 0];    /* ボタン行の余白 / margins of the button row */
    var COUNT_FIELD_CHARS   = 3;                /* 個数を入れる欄の文字数 / width of count fields */
    var VALUE_FIELD_CHARS   = 6;                /* 数値を入れる欄の文字数 / width of value fields */

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * 現在の表示言語を取得する
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = getCurrentLang();

    /* カテゴリ分けした日英ラベル定義 / Categorized Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "円周にテキストを繰り返す", en: "Repeat Text Around a Circle" }
        },
        panel: {
            repeat:    { ja: "繰り返しと回転", en: "Repeat & Rotation" },
            separator: { ja: "区切り文字", en: "Separator" },
            fontSize:  { ja: "文字サイズ", en: "Font Size" }
        },
        label: {
            repeatCount: { ja: "繰り返し数", en: "Repeat count" },
            separator:   { ja: "種類", en: "Type" },
            spaceCount:  { ja: "スペース数", en: "Spaces" },
            scale:       { ja: "区切り文字の比率", en: "Separator scale" },
            baseline:    { ja: "区切り文字のベースライン", en: "Separator baseline" },
            correction:  { ja: "補正率", en: "Correction" },
            rotation:    { ja: "回転角度", en: "Rotation angle" }
        },
        radio: {
            space:     { ja: "スペース", en: "Space" },
            character: { ja: "任意の文字", en: "Custom character" }
        },
        checkbox: {
            preview: { ja: "プレビュー", en: "Preview" },
            fitSize: { ja: "円周に合わせて文字サイズを調整", en: "Adjust font size to circumference" }
        },
        button: {
            ok:     { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        tooltip: {
            repeatCount: {
                ja: "元のテキストを円周上に繰り返す回数です。",
                en: "Number of times to repeat the source text around the path."
            },
            rotation: {
                ja: "生成したパス上文字を、円の中心を基準に回転します。",
                en: "Rotates the generated path text around the circle center."
            },
            separatorSpace: {
                ja: "テキスト同士をスペースだけで区切ります。",
                en: "Separates repeated text with spaces only."
            },
            separatorCharacter: {
                ja: "入力した文字を区切り文字として使用します。初期値は欧文 bullet です。",
                en: "Uses the typed character as the separator. The default is a bullet."
            },
            spaceCount: {
                ja: "区切り文字の前後に入れる半角スペース数です。スペースのみの場合は区切りのスペース数になります。",
                en: "Number of half-width spaces on each side of the separator. For Space mode, this is the separator spacing."
            },
            scale: {
                ja: "スペース以外の区切り文字にだけ適用する水平・垂直比率です。",
                en: "Horizontal and vertical scale applied only to non-space separator characters."
            },
            baseline: {
                ja: "スペース以外の区切り文字にだけ適用するベースラインシフトです。単位はIllustratorの文字単位設定に従います。",
                en: "Baseline shift applied only to non-space separator characters. The unit follows Illustrator's text unit preference."
            },
            fitSize: {
                ja: "円周に収まるように、元テキストの文字サイズを基準に自動調整します。",
                en: "Automatically adjusts font size based on the source text so the repeated text fits the circumference."
            },
            correction: {
                ja: "文字サイズ自動調整時の隙間を微調整します。100%が基準で、大きくすると文字が詰まり、小さくすると隙間が空きます。",
                en: "Fine-tunes the gap when auto-adjusting font size. 100% is the baseline; larger values tighten the text, smaller values open the gap."
            },
            preview: {
                ja: "設定変更を一時的に反映して確認します。確定するまでは元のオブジェクトを保持します。",
                en: "Temporarily previews the result while keeping the original objects until you click OK."
            },
            arrowKeys: {
                ja: "↑↓キーで増減できます（Shift+↑↓で10単位、Option+↑↓で0.1単位）。",
                en: "Step with the Up/Down keys (Shift for 10, Option for 0.1)."
            }
        },
        alert: {
            noDocument: {
                ja: "ドキュメントが開かれていません。",
                en: "No document is open."
            },
            needCircleAndText: {
                ja: "円のパスとテキストを1つずつ選択してください。",
                en: "Please select one circle path and one text frame."
            },
            emptyText: {
                ja: "テキストが空です。",
                en: "The text is empty."
            },
            invalidCount: {
                ja: "繰り返し数には1以上の整数を入力してください。",
                en: "Enter an integer of 1 or more for the repeat count."
            },
            invalidCorrection: {
                ja: "補正率には0より大きい数値を入力してください。",
                en: "Enter a number greater than 0 for the correction."
            }
        }
    };

    /**
     * ネストキーからローカライズ文字列を取得する（例: getLabel("panel.fontSize")）
     * @param {string} labelKey - ドットでつないだ LABELS のキー
     * @returns {string} 現在の言語のラベル（見つからない場合はキーそのもの）
     */
    function getLabel(labelKey) {
        var keyPath = labelKey.split(".");
        var labelNode = LABELS;
        for (var i = 0; i < keyPath.length; i++) {
            if (labelNode == null) return labelKey;
            labelNode = labelNode[keyPath[i]];
        }
        if (labelNode == null) return labelKey;
        return (labelNode[uiLang] != null) ? labelNode[uiLang] : labelNode.en;
    }

    /**
     * コロン付きの項目名を返す（日本語は全角、英語は半角）
     * @param {string} labelKey - ドットでつないだ LABELS のキー
     * @returns {string} コロンを付けた項目名
     */
    function labelText(labelKey) {
        return getLabel(labelKey) + (uiLang === "ja" ? "：" : ":");
    }

    /**
     * コントロールにツールチップを設定する
     * @param {object} control - ScriptUI コントロール
     * @param {string} tooltipKey - ドットでつないだ LABELS のキー
     * @returns {void}
     */
    function setTooltip(control, tooltipKey) {
        control.helpTip = getLabel(tooltipKey);
    }

    // =========================================
    // 単位 / Units
    // =========================================

    /**
     * 単位コードからラベルと1単位あたりの pt 数を取得する（環境設定の text/units 用）
     * @param {number} unitCode - 0=inch / 1=mm / 2=pt / 3=pica / 4=cm / 5=Q / 6=px
     * @returns {{label: string, ptPerUnit: number}} 単位表記と換算係数
     */
    function getUnitInfo(unitCode) {
        switch (unitCode) {
            case 0: return { label: "in", ptPerUnit: 72 };
            case 1: return { label: "mm", ptPerUnit: 72 / 25.4 };
            case 3: return { label: "pica", ptPerUnit: 12 };
            case 4: return { label: "cm", ptPerUnit: 72 / 2.54 };
            case 5: return { label: "Q", ptPerUnit: (72 / 25.4) * 0.25 };  /* 1Q = 0.25mm */
            case 6: return { label: "px", ptPerUnit: 1 };                  /* 72ppi 前提 / assumes 72ppi */
            default: return { label: "pt", ptPerUnit: 1 };                 /* 2=pt ほか / 2=pt and fallback */
        }
    }

    // =========================================
    // レイアウトの共通処理 / Shared layout helpers
    // =========================================

    /**
     * ウィンドウの共通設定
     * @param {Window} targetWindow - 対象のウィンドウ
     * @returns {void}
     */
    function setupWindow(targetWindow) {
        targetWindow.orientation = "column";
        targetWindow.alignChildren = "fill";
        targetWindow.margins = WINDOW_MARGINS;
        targetWindow.spacing = WINDOW_SPACING;
    }

    /**
     * パネルの共通設定
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
     * 行グループの共通設定（親パネルの fill 継承を打ち消す）
     * @param {Group} targetGroup - 対象のグループ
     * @param {string} [verticalAlign] - 天地の揃え（省略時は "center"）
     * @returns {void}
     */
    function setupRow(targetGroup, verticalAlign) {
        targetGroup.orientation = "row";
        targetGroup.alignment = ["left", "center"];
        targetGroup.alignChildren = ["left", verticalAlign || "center"];
    }

    /**
     * パネルを追加して共通設定を適用する
     * @param {object} parentContainer - 追加先のコンテナー
     * @param {string} panelTitleKey - ドットでつないだ LABELS のキー
     * @returns {Panel} 追加したパネル
     */
    function addPanel(parentContainer, panelTitleKey) {
        var createdPanel = parentContainer.add("panel", undefined, getLabel(panelTitleKey));
        setupPanel(createdPanel);
        return createdPanel;
    }

    // =========================================
    // 文字列 / Strings
    // =========================================

    /**
     * 指定個数の半角スペース文字列を作る
     * @param {number} spaceCount - スペースの個数
     * @returns {string} 半角スペースだけの文字列
     */
    function buildSpaces(spaceCount) {
        return new Array(spaceCount + 1).join(" ");
    }

    /**
     * 繰り返し文字列を作る
     * @param {string} sourceText - 元のテキスト
     * @param {number} repeatCount - 繰り返し数
     * @param {string} separatorText - 区切り文字
     * @param {boolean} appendTrailing - 末尾にも区切り文字を付けるか
     * @returns {string} 連結後の文字列
     */
    function buildRepeatedText(sourceText, repeatCount, separatorText, appendTrailing) {
        var segments = [];
        for (var i = 0; i < repeatCount; i++) {
            segments.push(sourceText);
        }
        var joinedText = segments.join(separatorText);
        return appendTrailing ? (joinedText + separatorText) : joinedText;
    }

    // =========================================
    // 図形 / Geometry
    // =========================================

    /**
     * 円・楕円のおおよその周長を返す（Ramanujan 近似）
     * @param {PathItem} ellipsePath - 対象のパス
     * @returns {number} 周長（pt）
     */
    function getEllipsePerimeter(ellipsePath) {
        var pathBounds = ellipsePath.geometricBounds;
        var radiusX = Math.abs(pathBounds[2] - pathBounds[0]) / 2;
        var radiusY = Math.abs(pathBounds[1] - pathBounds[3]) / 2;

        /* ramanujanH は (a-b)^2/(a+b)^2 / ramanujanH is (a-b)^2/(a+b)^2 */
        var ramanujanH = Math.pow(radiusX - radiusY, 2) / Math.pow(radiusX + radiusY, 2);
        return Math.PI * (radiusX + radiusY) *
            (1 + (3 * ramanujanH) / (10 + Math.sqrt(4 - 3 * ramanujanH)));
    }

    /**
     * 指定点を中心にアイテムを回転する
     * @param {PageItem} targetItem - 対象のアイテム
     * @param {number} rotationAngle - 回転角度（度）
     * @param {number} centerX - 回転中心のX座標
     * @param {number} centerY - 回転中心のY座標
     * @returns {void}
     */
    function rotateAroundCenter(targetItem, rotationAngle, centerX, centerY) {
        var rotationMatrix = app.getRotationMatrix(rotationAngle);
        targetItem.translate(-centerX, -centerY);
        targetItem.transform(rotationMatrix);
        targetItem.translate(centerX, centerY);
    }

    // =========================================
    // 入力値の解析 / Field parsing
    // =========================================

    /**
     * 1以上の整数として読み取る
     * @param {string} fieldText - 入力欄の文字列
     * @returns {number|null} 整数、条件を満たさない場合は null
     */
    function parsePositiveInteger(fieldText) {
        var parsedValue = parseInt(fieldText, 10);
        return (isNaN(parsedValue) || parsedValue < 1) ? null : parsedValue;
    }

    /**
     * 0 より大きい数値として読み取る
     * @param {string} fieldText - 入力欄の文字列
     * @returns {number|null} 数値、条件を満たさない場合は null
     */
    function parsePositiveNumber(fieldText) {
        var parsedValue = parseFloat(fieldText);
        return (isNaN(parsedValue) || parsedValue <= 0) ? null : parsedValue;
    }

    /**
     * 数値として読み取る（負値・小数可）
     * @param {string} fieldText - 入力欄の文字列
     * @param {number} fallbackValue - 読み取れないときの値
     * @returns {number} 読み取った数値
     */
    function parseNumberOr(fieldText, fallbackValue) {
        var parsedValue = parseFloat(fieldText);
        return isNaN(parsedValue) ? fallbackValue : parsedValue;
    }

    /**
     * 修飾キーに応じて次の値を計算する（通常 ±1 / Shift ±10・10 の倍数にスナップ / Option ±0.1）
     * @param {number} currentValue - 現在の値
     * @param {number} direction - 1 なら増加、-1 なら減少
     * @param {object} keyboardState - ScriptUI.environment.keyboardState
     * @returns {number} 変更後の値
     */
    function stepValueByModifier(currentValue, direction, keyboardState) {
        if (keyboardState.shiftKey) {
            /* Shift：±10 で 10 の倍数にスナップ / Shift: step ±10 snapped to a multiple of 10 */
            var snappedValue = (direction > 0)
                ? Math.ceil((currentValue + 1) / 10) * 10
                : Math.floor((currentValue - 1) / 10) * 10;
            return Math.round(snappedValue);
        }
        if (keyboardState.altKey) {
            /* Option：±0.1（小数第1位に丸め）/ Option: step ±0.1 rounded to one decimal */
            return Math.round((currentValue + direction * 0.1) * 10) / 10;
        }
        /* 通常：±1（整数に丸め）/ Default: step ±1 rounded to an integer */
        return Math.round(currentValue + direction);
    }

    /**
     * ↑↓キーで数値欄の値を増減できるようにする
     * @param {EditText} numberField - 対象の入力欄
     * @param {boolean} [allowNegative] - 負値を許可するか
     * @returns {void}
     */
    function changeValueByArrowKey(numberField, allowNegative) {
        numberField.addEventListener("keydown", function (event) {
            var direction = 0;
            if (event.keyName === "Up") direction = 1;
            else if (event.keyName === "Down") direction = -1;
            else return;

            var currentValue = Number(numberField.text);
            if (isNaN(currentValue)) return;

            event.preventDefault();

            var keyboardState = ScriptUI.environment.keyboardState;
            var nextValue = stepValueByModifier(currentValue, direction, keyboardState);

            /* 減算時のみ 0 で下げ止め（Option と allowNegative は除外）/ Floor at 0 only when decreasing (Option and allowNegative excluded) */
            if (!allowNegative && direction < 0 && !keyboardState.altKey && nextValue < 0) {
                nextValue = 0;
            }

            numberField.text = nextValue;

            /* .text への代入では onChanging が発火しないため明示的に呼ぶ / Setting .text does not fire onChanging, so call it manually */
            if (typeof numberField.onChanging === "function") {
                numberField.onChanging();
            }
        });
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択内容を検証し、ダイアログを開いてパス上文字を作成する
     * @returns {void}
     */
    function main() {

        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }

        var activeDoc = app.activeDocument;

        /* 選択数を確認（円とテキストの2つ）/ Ensure exactly two objects are selected */
        if (activeDoc.selection.length !== 2) {
            alert(getLabel("alert.needCircleAndText"));
            return;
        }

        var selectedItems = activeDoc.selection;
        var sourceTextFrame = null;
        var circlePath = null;

        /* 選択からテキストフレームと円のパスを振り分け / Pick the text frame and the circle path from the selection */
        for (var i = 0; i < selectedItems.length; i++) {
            if (selectedItems[i].typename === "TextFrame") {
                sourceTextFrame = selectedItems[i];
            } else if (selectedItems[i].typename === "PathItem") {
                circlePath = selectedItems[i];
            }
        }

        if (sourceTextFrame === null || circlePath === null) {
            alert(getLabel("alert.needCircleAndText"));
            return;
        }

        if (sourceTextFrame.contents === "") {
            alert(getLabel("alert.emptyText"));
            return;
        }

        /* パス上文字は改行を保持できないため、改行はスペースに置き換える / Path text cannot keep line breaks, so replace them with spaces */
        var originalText = sourceTextFrame.contents.replace(/[\r\n]/g, " ");

        var measureFrame = null;      /* 画面外に常駐させる計測用フレーム（使い回し）/ Reusable off-canvas measurement frame */
        var isUpdatingPreview = false;
        var isPreviewApplied = false; /* 適用済みで undo が必要か / Whether an applied preview still needs undoing */
        var hasCommitted = false;     /* OK で確定したか / Whether OK has committed the result */

        // -----------------------------------------
        // 文字属性のコピー / Copying character attributes
        // -----------------------------------------

        /**
         * 属性を1つだけ安全にコピーする（1つ失敗しても他へ波及させない）
         * @param {object} sourceAttributes - コピー元の属性
         * @param {object} targetAttributes - コピー先の属性
         * @param {string} attributeName - 属性名
         * @returns {void}
         */
        function safeCopyAttribute(sourceAttributes, targetAttributes, attributeName) {
            try {
                targetAttributes[attributeName] = sourceAttributes[attributeName];
            } catch (e) { }
        }

        /**
         * 元テキストの文字属性・段落属性をコピーする
         * @param {TextFrame} targetFrame - コピー先のテキストフレーム
         * @returns {void}
         */
        function copySourceAttributes(targetFrame) {
            var sourceAttributes = sourceTextFrame.textRange.characterAttributes;
            var targetAttributes = targetFrame.textRange.characterAttributes;
            var attributeNames = ["size", "textFont", "fillColor", "tracking",
                "horizontalScale", "verticalScale", "baselineShift"];
            for (var i = 0; i < attributeNames.length; i++) {
                safeCopyAttribute(sourceAttributes, targetAttributes, attributeNames[i]);
            }
            safeCopyAttribute(
                sourceTextFrame.textRange.paragraphAttributes,
                targetFrame.textRange.paragraphAttributes,
                "justification"
            );
        }

        // -----------------------------------------
        // 文字幅の計測 / Measuring text width
        // -----------------------------------------

        /**
         * 計測用フレームがまだ document 上に存在するか調べる
         * @returns {boolean} 参照できれば true
         */
        function isMeasureFrameAlive() {
            if (measureFrame === null) return false;
            try {
                measureFrame.contents;   /* 参照できれば有効 / accessible means still valid */
                return true;
            } catch (e) {
                return false;
            }
        }

        /**
         * 計測用フレームを1枚だけ用意する（画面外に配置して使い回す）
         * @returns {TextFrame} 計測用のテキストフレーム
         */
        function ensureMeasureFrame() {
            if (!isMeasureFrameAlive()) {
                measureFrame = activeDoc.textFrames.add();
                measureFrame.top = MEASURE_FRAME_OFFSET;
                measureFrame.left = MEASURE_FRAME_OFFSET;
                /* 生成を確定し、プレビューの app.undo() で消えないようにする / Commit the creation so the preview's app.undo() cannot remove it */
                app.redraw();
            }
            return measureFrame;
        }

        /**
         * 計測用フレームを削除する
         * @returns {void}
         */
        function removeMeasureFrame() {
            if (isMeasureFrameAlive()) {
                measureFrame.remove();
            }
            measureFrame = null;
        }

        /**
         * 元テキストの属性を写した使い回しフレームで文字幅を測る
         * @param {string} textToMeasure - 計測する文字列
         * @returns {{fontSize: number, width: number}} 計測時の文字サイズと実測幅（pt）
         */
        function measureText(textToMeasure) {
            var measureTarget = ensureMeasureFrame();
            measureTarget.contents = textToMeasure;
            copySourceAttributes(measureTarget);

            /* 元テキストのサイズをそのまま基準にする（読めなければ代替値を書き込む）/ Use the source size as the base (write a fallback when unreadable) */
            var measureAttributes = measureTarget.textRange.characterAttributes;
            var baseFontSize = measureAttributes.size;
            if (isNaN(baseFontSize) || baseFontSize <= 0) {
                baseFontSize = FALLBACK_FONT_SIZE;
                measureAttributes.size = baseFontSize;
            }

            /* geometricBounds を正しく更新するため redraw が必要。常駐フレームなので毎回の add/remove は無く、蓄積しない。
               A redraw is required for geometricBounds to update correctly. The frame is reused (no per-call add/remove), so nothing accumulates. */
            app.redraw();

            var frameBounds = measureTarget.geometricBounds;
            return { fontSize: baseFontSize, width: Math.abs(frameBounds[2] - frameBounds[0]) };
        }

        /**
         * 文字サイズを Illustrator が受け付ける範囲に収める
         * @param {number} fontSize - 調整前の文字サイズ（pt）
         * @returns {number} 範囲内に収めた文字サイズ（pt）
         */
        function clampFontSize(fontSize) {
            if (isNaN(fontSize)) return FALLBACK_FONT_SIZE;
            if (fontSize < MIN_FONT_SIZE) return MIN_FONT_SIZE;
            if (fontSize > MAX_FONT_SIZE) return MAX_FONT_SIZE;
            return fontSize;
        }

        // -----------------------------------------
        // 区切り文字 / Separator
        // -----------------------------------------

        /**
         * @typedef {object} SeparatorInfo
         * @property {string} text - 実際に挟む区切り文字列
         * @property {string} styledChars - スケール・ベースラインを適用する文字（スペースのみなら ""）
         * @property {number} leadingSpaces - 区切り文字列の先頭にあるスペース数
         */

        /**
         * 現在のUI状態から区切り文字の情報を返す
         * @param {number} spaceCount - 区切り文字の前後に入れる半角スペース数
         * @returns {SeparatorInfo} 区切り文字の情報
         */
        function getSeparatorInfo(spaceCount) {
            var sideSpaces = buildSpaces(spaceCount);
            if (!separatorCharRadio.value) {
                /* スペースのみ（末尾にも付与し、円の折り返し位置の間隔も揃える）/ Spaces only (also trailing, so the wrap-around gap matches) */
                return { text: sideSpaces, styledChars: "", leadingSpaces: 0 };
            }
            /* 入力文字の左右にスペースを付与 / Pad the typed character with spaces on both sides */
            var separatorChar = separatorCharInput.text;
            return {
                text: sideSpaces + separatorChar + sideSpaces,
                styledChars: separatorChar,
                leadingSpaces: sideSpaces.length
            };
        }

        /**
         * 区切り文字（スペース以外）だけに水平・垂直比率とベースラインを適用する
         * 文字の内容ではなく位置で特定するため、元テキストに同じ文字があっても影響しない
         * @param {TextFrame} pathTypeFrame - 対象のパス上文字
         * @param {SeparatorInfo} separatorInfo - 区切り文字の情報
         * @param {number} repeatCount - 繰り返し数（＝区切り文字の個数）
         * @param {number} scalePercent - 水平・垂直比率（%）
         * @param {number} baselineShiftPt - ベースラインシフト（pt）
         * @returns {void}
         */
        function applySeparatorStyle(pathTypeFrame, separatorInfo, repeatCount, scalePercent, baselineShiftPt) {
            var styledLength = separatorInfo.styledChars.length;
            if (styledLength === 0) return;

            var characters = pathTypeFrame.textRange.characters;
            var sourceLength = originalText.length;
            var separatorLength = separatorInfo.text.length;

            for (var i = 0; i < repeatCount; i++) {
                /* i 番目の区切り文字が始まる位置 / Start index of the i-th separator */
                var styleStart = sourceLength * (i + 1) + separatorLength * i + separatorInfo.leadingSpaces;
                for (var j = 0; j < styledLength; j++) {
                    if (styleStart + j >= characters.length) return;
                    var attributes = characters[styleStart + j].characterAttributes;
                    attributes.horizontalScale = scalePercent;
                    attributes.verticalScale = scalePercent;
                    attributes.baselineShift = baselineShiftPt;
                }
            }
        }

        // -----------------------------------------
        // 生成 / Generating
        // -----------------------------------------

        /**
         * @typedef {object} RepeatSettings
         * @property {number|null} repeatCount - 繰り返し数（不正なら null）
         * @property {number} spaceCount - 区切り文字の前後のスペース数
         * @property {number} separatorScale - 区切り文字のスケール（%）
         * @property {number} baselineShiftPt - 区切り文字のベースラインシフト（pt）
         * @property {number|null} correctionPercent - 円周の補正率（%、不正なら null）
         * @property {number} rotationAngle - 回転角度（度）
         * @property {boolean} shouldFit - 文字サイズを円周に合わせるか
         */

        /**
         * ダイアログの入力値をまとめて読み取る
         * @returns {RepeatSettings} 現在の設定
         */
        function readSettings() {
            return {
                repeatCount: parsePositiveInteger(repeatCountInput.text),
                spaceCount: parsePositiveInteger(spaceCountInput.text) || DEFAULT_SPACE_COUNT,
                separatorScale: parsePositiveNumber(scaleInput.text) || DEFAULT_SEPARATOR_SCALE,
                baselineShiftPt: parseNumberOr(baselineInput.text, 0) * textUnitInfo.ptPerUnit,
                correctionPercent: parsePositiveNumber(correctionInput.text),
                rotationAngle: parseNumberOr(rotationInput.text, 0),
                shouldFit: fitSizeCheckbox.value
            };
        }

        /**
         * 必須項目を検証する
         * @param {RepeatSettings} settings - 検証する設定
         * @returns {string|null} 不正なら alert のキー、問題なければ null
         */
        function validateSettings(settings) {
            if (settings.repeatCount === null) return "alert.invalidCount";
            /* 補正率はフィット ON のときだけ必須 / The correction is required only when fitting is on */
            if (settings.shouldFit && settings.correctionPercent === null) return "alert.invalidCorrection";
            return null;
        }

        /**
         * 円周に合う文字サイズを計算する（計測のため redraw を伴う）
         * @param {RepeatSettings} settings - 現在の設定
         * @returns {number|null} 文字サイズ（pt）、フィット OFF なら null
         */
        function computeFittedFontSize(settings) {
            if (!settings.shouldFit) return null;

            var separatorInfo = getSeparatorInfo(settings.spaceCount);
            var repeatedText = buildRepeatedText(originalText, settings.repeatCount, separatorInfo.text, true);
            var measured = measureText(repeatedText);
            if (measured.width <= 0) return measured.fontSize;

            var perimeter = getEllipsePerimeter(circlePath);
            var correctionRatio = settings.correctionPercent / 100;
            return clampFontSize(measured.fontSize * ((perimeter * correctionRatio) / measured.width));
        }

        /**
         * パス上文字を作成する（計測・redraw は含めない＝1 undo グループにするため）
         * @param {RepeatSettings} settings - 現在の設定
         * @param {number|null} fontSize - 適用する文字サイズ（pt）、null ならフィット OFF
         * @param {boolean} isPreview - プレビューとして作成するか
         * @returns {TextFrame} 作成したパス上文字
         */
        function createPathTypeText(settings, fontSize, isPreview) {
            var separatorInfo = getSeparatorInfo(settings.spaceCount);
            var repeatedText = buildRepeatedText(originalText, settings.repeatCount, separatorInfo.text, true);

            var pathTypeFrame = activeDoc.textFrames.pathText(circlePath.duplicate());
            pathTypeFrame.contents = repeatedText;
            copySourceAttributes(pathTypeFrame);

            /* 事前計算したフィットサイズを適用（null はフィット OFF）/ Apply the precomputed fit size (null means fitting is off) */
            if (fontSize !== null) {
                pathTypeFrame.textRange.characterAttributes.size = fontSize;
            }

            /* 区切り文字（スペース以外）のスケール・ベースラインを適用 / Apply scale and baseline to the separator's non-space characters */
            if (settings.separatorScale !== 100 || settings.baselineShiftPt !== 0) {
                applySeparatorStyle(pathTypeFrame, separatorInfo, settings.repeatCount,
                    settings.separatorScale, settings.baselineShiftPt);
            }

            /* 円の中心を基準に回転 / Rotate around the circle center */
            if (settings.rotationAngle !== 0) {
                var pathBounds = circlePath.geometricBounds;
                rotateAroundCenter(pathTypeFrame, settings.rotationAngle,
                    (pathBounds[0] + pathBounds[2]) / 2, (pathBounds[1] + pathBounds[3]) / 2);
            }

            /* プレビュー時は元のテキスト・円を一時的に隠す（undo で復帰）/ Hide the originals during preview (restored by undo) */
            if (isPreview) {
                sourceTextFrame.hidden = true;
                circlePath.hidden = true;
            }

            return pathTypeFrame;
        }

        // -----------------------------------------
        // ダイアログ / Dialog
        // -----------------------------------------

        /* 環境設定のテキスト単位（ベースライン入力に使用）/ Preferences text unit (used by the baseline field) */
        var textUnitInfo = getUnitInfo(app.preferences.getIntegerPreference("text/units"));

        /* 幅を揃える行ラベル / Row labels to align to one width */
        var rowLabels = [];

        /**
         * 行ラベルを追加し、幅揃えの対象に登録する
         * @param {Group} parentGroup - 追加先のグループ
         * @param {string} labelKey - ドットでつないだ LABELS のキー
         * @returns {StaticText} 追加したラベル
         */
        function addRowLabel(parentGroup, labelKey) {
            var rowLabel = parentGroup.add("statictext", undefined, labelText(labelKey));
            rowLabel.justify = "right";
            rowLabels.push(rowLabel);
            return rowLabel;
        }

        /**
         * ラベル＋数値入力欄（必要なら単位表記）の行を追加する
         * @param {Panel} parentPanel - 追加先のパネル
         * @param {string} labelKey - 項目名の LABELS キー
         * @param {number} defaultValue - 入力欄の初期値
         * @param {number} fieldChars - 入力欄の文字数
         * @param {string} tooltipKey - ツールチップの LABELS キー
         * @param {string} [suffixText] - 入力欄の右に添える単位表記
         * @returns {EditText} 追加した入力欄
         */
        function addNumberField(parentPanel, labelKey, defaultValue, fieldChars, tooltipKey, suffixText) {
            var fieldRow = parentPanel.add("group");
            setupRow(fieldRow);
            addRowLabel(fieldRow, labelKey);

            var numberField = fieldRow.add("edittext", undefined, String(defaultValue));
            numberField.characters = fieldChars;
            /* 数値欄は共通で ↑↓ 操作を案内する / Every number field documents the arrow-key stepping */
            numberField.helpTip = getLabel(tooltipKey) + "\n" + getLabel("tooltip.arrowKeys");

            if (suffixText) {
                fieldRow.add("statictext", undefined, suffixText);
            }
            return numberField;
        }

        /**
         * 全ラベルの幅を最も広いものに揃える
         * @returns {void}
         */
        function alignLabelWidths() {
            var maxWidth = 0;
            for (var i = 0; i < rowLabels.length; i++) {
                if (rowLabels[i].preferredSize.width > maxWidth) {
                    maxWidth = rowLabels[i].preferredSize.width;
                }
            }
            for (var j = 0; j < rowLabels.length; j++) {
                rowLabels[j].preferredSize.width = maxWidth;
            }
        }

        /* タイトルバーにバージョンを表示 / Show the version in the title bar */
        var dialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        setupWindow(dialog);

        var repeatPanel = addPanel(dialog, "panel.repeat");
        var repeatCountInput = addNumberField(repeatPanel, "label.repeatCount", DEFAULT_REPEAT_COUNT, COUNT_FIELD_CHARS, "tooltip.repeatCount");
        var rotationInput = addNumberField(repeatPanel, "label.rotation", DEFAULT_ROTATION, COUNT_FIELD_CHARS, "tooltip.rotation", "°");

        var separatorPanel = addPanel(dialog, "panel.separator");

        var separatorRow = separatorPanel.add("group");
        setupRow(separatorRow, "top");
        addRowLabel(separatorRow, "label.separator");

        var separatorRadioGroup = separatorRow.add("group");
        separatorRadioGroup.orientation = "column";
        separatorRadioGroup.alignChildren = "left";

        var separatorSpaceRadio = separatorRadioGroup.add("radiobutton", undefined, getLabel("radio.space"));
        setTooltip(separatorSpaceRadio, "tooltip.separatorSpace");

        /* 「文字」ラジオ＋自由入力フィールド（初期値は欧文 bullet）/ "Character" radio plus a free-input field (defaults to the bullet) */
        var separatorCharRow = separatorRadioGroup.add("group");
        setupRow(separatorCharRow);
        var separatorCharRadio = separatorCharRow.add("radiobutton", undefined, getLabel("radio.character"));
        setTooltip(separatorCharRadio, "tooltip.separatorCharacter");
        var separatorCharInput = separatorCharRow.add("edittext", undefined, DEFAULT_SEPARATOR_CHAR);
        separatorCharInput.characters = COUNT_FIELD_CHARS;
        setTooltip(separatorCharInput, "tooltip.separatorCharacter");

        /* スペースと文字は親グループが異なるため排他制御は手動 / Space and character radios live in different parents, so exclusivity is handled manually */
        separatorSpaceRadio.value = false;
        separatorCharRadio.value = true;

        var spaceCountInput = addNumberField(separatorPanel, "label.spaceCount", DEFAULT_SPACE_COUNT, VALUE_FIELD_CHARS, "tooltip.spaceCount");
        var scaleInput = addNumberField(separatorPanel, "label.scale", DEFAULT_SEPARATOR_SCALE, VALUE_FIELD_CHARS, "tooltip.scale", "%");
        /* 単位表記は環境設定のテキスト単位に従う / The unit label follows the preferences text unit */
        var baselineInput = addNumberField(separatorPanel, "label.baseline", DEFAULT_BASELINE_SHIFT, VALUE_FIELD_CHARS, "tooltip.baseline", textUnitInfo.label);

        var fontSizePanel = addPanel(dialog, "panel.fontSize");

        var fitSizeCheckbox = fontSizePanel.add("checkbox", undefined, getLabel("checkbox.fitSize"));
        fitSizeCheckbox.value = true;
        setTooltip(fitSizeCheckbox, "tooltip.fitSize");

        var correctionInput = addNumberField(fontSizePanel, "label.correction", DEFAULT_CORRECTION, VALUE_FIELD_CHARS, "tooltip.correction", "%");

        alignLabelWidths();

        /* メイングループ（横並び）/ Main group (horizontal layout) */
        var btnRowGroup = dialog.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.margins = BUTTON_ROW_MARGINS;
        btnRowGroup.alignment = ["fill", "bottom"];

        /* 左側グループ：プレビュー / Left-side group: preview */
        var btnLeftGroup = btnRowGroup.add("group");
        btnLeftGroup.alignChildren = ["left", "center"];
        var previewCheckbox = btnLeftGroup.add("checkbox", undefined, getLabel("checkbox.preview"));
        previewCheckbox.value = true;
        setTooltip(previewCheckbox, "tooltip.preview");

        /* スペーサー（伸縮）/ Spacer (stretchable) */
        var spacer = btnRowGroup.add("group");
        spacer.alignment = ["fill", "fill"];
        spacer.minimumSize.width = 0;

        /* 右側グループ：ボタン2つ / Right-side group: two buttons */
        var btnRightGroup = btnRowGroup.add("group");
        btnRightGroup.alignChildren = ["right", "center"];
        var btnCancel = btnRightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnOK = btnRightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });

        // -----------------------------------------
        // プレビュー / Preview
        // -----------------------------------------

        /**
         * プレビューを適用して表示する（undo は呼び出し元が行う）
         * @returns {void}
         */
        function applyPreview() {
            if (!previewCheckbox.value) return;

            var settings = readSettings();
            if (validateSettings(settings) !== null) return;

            /* 計測は redraw を含むため適用バッチの前に実行 / Measure before the apply batch (it involves a redraw) */
            var fontSize = computeFittedFontSize(settings);

            /* ここから先は必ず undo が要る / Everything past this point must be undone */
            isPreviewApplied = true;

            /* 仮アイテムで強制的に変化を起こし、undo の空振りを防ぐ。画面外に作り、変数にも保持しない（undo で消える）/ Force a change with an off-canvas dummy so undo cannot misfire (not kept in a variable since undo removes it) */
            activeDoc.pathItems.rectangle(MEASURE_FRAME_OFFSET, MEASURE_FRAME_OFFSET, 1, 1);

            createPathTypeText(settings, fontSize, true);
            app.redraw();   /* 見せる / show the applied result */
        }

        /**
         * プレビューを更新する（適用 → redraw → undo。再入と失敗を吸収する）
         * @returns {void}
         */
        function runPreview() {
            if (isUpdatingPreview) return;
            isUpdatingPreview = true;
            isPreviewApplied = false;

            try {
                applyPreview();
            } catch (e) {
                /* プレビューは best-effort。失敗しても操作を続けられるようにする / Preview is best-effort; keep the dialog usable on failure */
            }

            if (isPreviewApplied) {
                app.undo();   /* 内部を未適用へ戻す（画面は適用後のまま）/ revert the model (screen keeps showing it) */
            } else {
                app.redraw(); /* 未適用状態を表示 / show the clean state */
            }
            isUpdatingPreview = false;
        }

        // -----------------------------------------
        // イベント / Events
        // -----------------------------------------

        repeatCountInput.onChanging = runPreview;
        separatorCharInput.onChanging = runPreview;
        spaceCountInput.onChanging = runPreview;
        scaleInput.onChanging = runPreview;
        baselineInput.onChanging = runPreview;
        correctionInput.onChanging = runPreview;
        rotationInput.onChanging = runPreview;

        previewCheckbox.onClick = runPreview;

        fitSizeCheckbox.onClick = function () {
            correctionInput.enabled = fitSizeCheckbox.value;
            runPreview();
        };

        /**
         * 区切り文字の選択に合わせて入力欄の有効・無効を切り替える
         * スケールとベースラインは「文字」のときだけ効く
         * @returns {void}
         */
        function updateSeparatorFields() {
            var isCharSelected = separatorCharRadio.value;
            separatorCharInput.enabled = isCharSelected;
            scaleInput.enabled = isCharSelected;
            baselineInput.enabled = isCharSelected;
        }

        /* クリックした側を選択し、もう一方を解除（手動排他）/ Select the clicked radio and clear the other (manual exclusivity) */
        separatorSpaceRadio.onClick = function () {
            separatorSpaceRadio.value = true;
            separatorCharRadio.value = false;
            updateSeparatorFields();
            runPreview();
        };
        separatorCharRadio.onClick = function () {
            separatorCharRadio.value = true;
            separatorSpaceRadio.value = false;
            updateSeparatorFields();
            runPreview();
        };

        btnOK.onClick = function () {
            var settings = readSettings();
            var invalidKey = validateSettings(settings);
            if (invalidKey !== null) {
                alert(getLabel(invalidKey));
                return;
            }

            /* 確定：未適用のクリーンな状態から1回だけ適用 / Commit: apply once from the clean state */
            var fontSize = computeFittedFontSize(settings);
            removeMeasureFrame();   /* 計測フレームを片付けてから確定 / Clean up the measurement frame before committing */

            var resultTextFrame = createPathTypeText(settings, fontSize, false);

            sourceTextFrame.remove();
            circlePath.remove();

            activeDoc.selection = null;
            resultTextFrame.selected = true;

            hasCommitted = true;
            app.redraw();
            dialog.close();
        };

        btnCancel.onClick = function () {
            /* 内部は undo 済み。計測フレームを片付け、元の選択へ戻して閉じる / Model already reverted; clean up the measurement frame, restore the selection, and close */
            removeMeasureFrame();
            activeDoc.selection = null;
            sourceTextFrame.selected = true;
            circlePath.selected = true;
            app.redraw();
            dialog.close();
        };

        dialog.onClose = function () {
            /* 計測フレームを片付け、未確定なら未適用状態を画面に反映 / Clean up the measurement frame; if not committed, refresh the screen to the reverted state */
            removeMeasureFrame();
            if (!hasCommitted) {
                app.redraw();
            }
        };

        /* 初期プレビューはウィンドウ表示後に起動（同期実行を避ける）/ Start the initial preview after the window is shown (avoid running it synchronously) */
        dialog.onShow = function () {
            runPreview();
        };

        /* 初期状態を反映 / Apply the initial state */
        correctionInput.enabled = fitSizeCheckbox.value;
        updateSeparatorFields();

        changeValueByArrowKey(repeatCountInput);
        changeValueByArrowKey(spaceCountInput);
        changeValueByArrowKey(scaleInput);
        changeValueByArrowKey(baselineInput, true);
        changeValueByArrowKey(correctionInput);
        changeValueByArrowKey(rotationInput, true);

        dialog.show();
    }

    main();

})();
