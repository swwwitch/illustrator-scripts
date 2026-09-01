#target illustrator
app.preferences.setBooleanPreference("ShowExternalJSXWarning", false);

/*

### 概要

選択した複数行のテキストを1段落ごとのテキストオブジェクトに分け、指定したアートボードから順番に配置し、アートボードの幅に合わせて中央にそろえます。

詳細は README を参照してください。

### Overview

Splits the selected multi-line text into one text object per paragraph, places them on the artboards in order starting from the one you choose, and fits each to the artboard width, centered.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SplitTextToArtboards";         /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-09-01";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-01";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SplitTextToArtboards.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SplitTextToArtboards.md
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/xxxxxxxx"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function() {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* ダイアログの初期値 / Dialog defaults */
    var DEFAULT_START_ARTBOARD = 1;     /* 配置を始めるアートボード番号 / artboard to start from */
    var DEFAULT_WIDTH_PERCENT  = 90;    /* アートボード幅に対する仕上がり幅（%） / target width as a percentage of the artboard */
    var DEFAULT_COLUMN_COUNT   = 6;     /* アートボード追加時の列数（既存の並びから読めないとき） / columns used when the layout gives none */
    var DEFAULT_ARTBOARD_GAP   = 20;    /* アートボード追加時の間隔（pt、既存の並びから読めないとき） / gap when the layout gives none */
    var DEFAULT_KEEP_SOURCE    = false; /* 元のテキストを残すか / keep the original text */

    /* 計測に使う境界 / Bounds used for measuring
       true : プレビュー境界（線幅・効果込みの見た目の端） / preview bounds (incl. strokes & effects)
       false: 幾何境界（パスの端） / geometric bounds (path edges) */
    var USE_PREVIEW_BOUNDS = true;

    /* リサイズ後に上下中央へもそろえるか（false なら縦位置はそのまま）
       Also center vertically after resizing (false keeps the vertical position) */
    var CENTER_VERTICALLY = true;

    /* 空段落を読み飛ばすか（false なら空段落にもアートボードを1枚使う）
       Skip empty paragraphs (false spends one artboard on each empty line) */
    var SKIP_EMPTY_PARAGRAPHS = true;

    /* 同じ列・同じ行とみなす座標の許容差（pt） / Tolerance for treating positions as the same row or column (pt) */
    var POSITION_TOLERANCE = 0.5;

    // =========================================
    // レイアウト / Layout
    // =========================================

    var WINDOW_MARGINS     = 16;               /* ウィンドウ外周の余白 / window margin */
    var WINDOW_SPACING     = 12;               /* ウィンドウ内の要素間隔 / window spacing */
    var PANEL_MARGINS      = [16, 20, 16, 12]; /* パネル余白 [左,上,右,下] / panel margins */
    var PANEL_SPACING      = 6;                /* パネル内の要素間隔 / panel spacing */
    var FIELD_ROW_SPACING  = 6;                /* ラベル・入力欄・単位表記の間隔 / gap inside a labeled field row */
    var LABEL_WIDTH        = 140;              /* 行ラベルの共通幅 / shared width of row labels */
    var FIELD_CHARACTERS   = 5;                /* 数値入力欄の文字数 / width of a number field in characters */
    var BUTTON_BAR_MARGINS = [0, 10, 0, 0];    /* ボタンバーの余白 / margins of the bottom button bar */
    var BUTTON_BAR_SPACING = 10;               /* ボタンバー内の要素間隔 / spacing inside the button bar */

    // =========================================
    // 日英ラベル定義 / Japanese-English label definitions
    // =========================================

    /**
     * 現在のUI言語を判定する
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        var localeText = ($.locale || "") + ""; /* 文字列化して扱う / Ensure a string */
        return (localeText.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = getCurrentLang();

    /* カテゴリ分けした日英ラベル定義（fieldLabel と tooltip はキーを共有する）
       Categorized labels; fieldLabel and tooltip share their keys */
    var LABELS = {
        dialog: {
            title: { ja: "段落をアートボードへ分配", en: "Split Text to Artboards" }
        },
        panel: {
            settings: { ja: "設定", en: "Settings" }
        },
        fieldLabel: {
            startArtboard: { ja: "どのアートボードから", en: "Start from artboard" },
            widthPercent:  { ja: "アートボードの幅", en: "Artboard width" },
            columnCount:   { ja: "追加時の列数", en: "Columns when adding" },
            artboardGap:   { ja: "アートボードの間隔", en: "Artboard gap" }
        },
        checkbox: {
            keepSource: { ja: "元のテキストを残す", en: "Keep the original text" }
        },
        tooltip: {
            startArtboard: { ja: "1段落目を置くアートボード番号。↑↓で増減、Shift+↑↓で10単位スナップ", en: "Artboard for the first paragraph. Up/Down to step, Shift+Up/Down snaps to 10" },
            widthPercent:  { ja: "アートボード幅に対する仕上がり幅。100 でアートボード幅ぴったり", en: "Target width as a percentage of the artboard width (100 = full width)" },
            columnCount:   { ja: "アートボードが足りないときに追加する並びの列数。既存の並びから読み取れた場合は初期値に入ります", en: "Columns used when adding artboards; prefilled from the existing layout when it can be read" },
            artboardGap:   { ja: "追加するアートボードどうしの間隔。既存の並びから読み取れた場合は初期値に入ります", en: "Gap between the added artboards; prefilled from the existing layout when it can be read" }
        },
        info: {
            summary: { ja: "対象の段落 %1 ／ 現在のアートボード %2", en: "%1 paragraph(s) / %2 artboard(s)" }
        },
        button: {
            ok:     { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        alert: {
            noDocument:    { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noSelection:   { ja: "オブジェクトが選択されていません。", en: "No object is selected." },
            noTextFrame:   { ja: "テキストオブジェクトを選択してください。", en: "Please select a text object." },
            noParagraph:   { ja: "分配できる段落がありません。", en: "There is no paragraph to distribute." },
            artboardLimit: {
                ja: "アートボードをこれ以上追加できないため、先頭の %1 段落だけを配置しました。",
                en: "No more artboards could be added, so only the first %1 paragraph(s) were placed."
            }
        }
    };

    /**
     * LABELS からカテゴリを辿って現在の言語のラベルを取得する（例: getLabel('alert','noDocument')）
     * @param {...string} keys - LABELS を辿るキー列
     * @returns {string} 該当するラベル（見つからない場合は空文字）
     */
    function getLabel() {
        var labelNode = LABELS;
        for (var i = 0; i < arguments.length; i++) {
            if (labelNode == null) break;
            labelNode = labelNode[arguments[i]];
        }
        return (labelNode && labelNode[uiLang] != null) ? labelNode[uiLang] : "";
    }

    /**
     * 項目名にコロンを付けて返す（日本語は全角、英語は半角）
     * @param {...string} keys - LABELS を辿るキー列
     * @returns {string} コロン付きの項目名
     */
    function labelText() {
        return getLabel.apply(null, arguments) + ((uiLang === "ja") ? "：" : ":");
    }

    // =========================================
    // UIレイアウト補助 / UI layout helpers
    // =========================================

    /**
     * ダイアログウィンドウを生成する（共通レイアウト適用）
     * @param {string} windowTitle - ウィンドウタイトル
     * @returns {Window} 生成したダイアログ
     */
    function createDialogWindow(windowTitle) {
        var dialogWindow = new Window("dialog", windowTitle);
        dialogWindow.orientation = "column";
        dialogWindow.alignChildren = ["fill", "top"];
        dialogWindow.spacing = WINDOW_SPACING;
        dialogWindow.margins = WINDOW_MARGINS;
        return dialogWindow;
    }

    /**
     * パネルに共通レイアウトを適用する
     * @param {Panel} targetPanel - 対象パネル
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
     * グループを横並びの行として設定する
     * @param {Group} targetGroup - 対象グループ
     * @param {string} [horizontalAlign] - 横方向の揃え（省略時は "left"）
     * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function setupRow(targetGroup, horizontalAlign, spacing) {
        targetGroup.orientation = "row";
        /* 揃えは横と天地を対で指定し、親の fill 継承を打ち消す / Pair both axes to cancel the parent's fill */
        targetGroup.alignment = [horizontalAlign || "left", "center"];
        targetGroup.alignChildren = ["left", "center"];
        targetGroup.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /**
     * ラベル付きパネルを生成する（共通レイアウト適用）
     * @param {Window|Group} parentContainer - 追加先
     * @param {string} panelTitle - パネルの見出し
     * @returns {Panel} 生成したパネル
     */
    function addPanel(parentContainer, panelTitle) {
        var createdPanel = parentContainer.add("panel");
        createdPanel.text = panelTitle;
        setupPanel(createdPanel);
        return createdPanel;
    }

    /**
     * 共通幅で右揃えの行ラベルを追加する
     * @param {Group} parentContainer - 追加先の行グループ
     * @param {string} text - 表示する文字列
     * @returns {StaticText} 生成したラベル
     */
    function addRowLabel(parentContainer, text) {
        var rowLabel = parentContainer.add("statictext", undefined, text);
        rowLabel.preferredSize.width = LABEL_WIDTH;
        rowLabel.justify = "right";
        return rowLabel;
    }

    /**
     * 入力欄に↑↓キーでの値増減を追加する（Shift併用で10単位スナップ）
     * @param {EditText} editText - 対象の入力欄
     * @returns {void}
     */
    function changeValueByArrowKey(editText) {
        editText.addEventListener("keydown", function (event) {
            var currentValue = Number(editText.text);
            if (isNaN(currentValue)) return;
            var keyboardState = ScriptUI.environment.keyboardState;

            if (event.keyName == "Up" || event.keyName == "Down") {
                if (keyboardState.shiftKey) {
                    /* Shift押下時は10の倍数スナップ / Snap to tens if Shift is pressed */
                    currentValue = Math.round(currentValue / 10) * 10 + (event.keyName == "Up" ? 10 : -10);
                } else {
                    currentValue += (event.keyName == "Up" ? 1 : -1);
                }

                event.preventDefault();
                editText.text = currentValue;
                /* プログラム変更は onChanging を発火しないため明示的に呼ぶ / fire onChanging manually */
                if (typeof editText.onChanging === "function") editText.onChanging();
            }
        });
    }

    /**
     * ラベル・数値入力欄・単位表記をひと組にした行を追加する
     * ラベルとツールチップは LABELS の同じキーから引く
     * @param {Panel|Group} parentContainer - 追加先
     * @param {string} fieldKey - LABELS.fieldLabel / LABELS.tooltip のキー
     * @param {number} defaultValue - 入力欄の初期値
     * @param {string} [unitText] - 入力欄の右に添える単位表記
     * @returns {EditText} 生成した入力欄
     */
    function addNumberFieldRow(parentContainer, fieldKey, defaultValue, unitText) {
        var fieldRow = parentContainer.add("group");
        setupRow(fieldRow, "left", FIELD_ROW_SPACING);
        var fieldLabel = addRowLabel(fieldRow, labelText("fieldLabel", fieldKey));
        var numberField = fieldRow.add("edittext", undefined, String(defaultValue));
        numberField.characters = FIELD_CHARACTERS;
        if (unitText) {
            fieldRow.add("statictext", undefined, unitText);
        }
        fieldLabel.helpTip = getLabel("tooltip", fieldKey);
        numberField.helpTip = getLabel("tooltip", fieldKey);
        changeValueByArrowKey(numberField);
        return numberField;
    }

    /**
     * 入力欄の値を数値として読み、範囲外なら既定値に戻す
     * @param {EditText} inputField - 対象の入力欄
     * @param {number} defaultValue - 既定値
     * @param {number} minimumValue - 下限
     * @param {number} maximumValue - 上限
     * @returns {number} 読み取った数値
     */
    function readNumberField(inputField, defaultValue, minimumValue, maximumValue) {
        var value = Number(inputField.text);
        if (isNaN(value) || value < minimumValue || value > maximumValue) {
            return defaultValue;
        }
        return value;
    }

    // =========================================
    // 矩形の計測 / Bounds
    // =========================================

    /**
     * オブジェクト1つの境界を返す（USE_PREVIEW_BOUNDS に従う）
     * @param {PageItem} item - 対象オブジェクト
     * @returns {number[]} [左, 上, 右, 下] の座標
     */
    function getItemBounds(item) {
        return USE_PREVIEW_BOUNDS ? item.visibleBounds : item.geometricBounds;
    }

    /**
     * 矩形の中心座標を求める
     * @param {number[]} rect - [左, 上, 右, 下] の座標
     * @returns {number[]} [中心X, 中心Y] の座標
     */
    function getRectCenter(rect) {
        return [(rect[0] + rect[2]) / 2, (rect[1] + rect[3]) / 2];
    }

    // =========================================
    // アートボード / Artboards
    // =========================================

    /**
     * @typedef {object} ArtboardLayout
     * @property {number} originLeft - 1列目の左端X
     * @property {number} lastLeft - 最後のアートボードの左端X
     * @property {number} lastTop - 最後のアートボードの上端Y
     * @property {number} width - アートボードの幅
     * @property {number} height - アートボードの高さ
     * @property {number} columnCount - 読み取れた列数（読めなければ既定値）
     * @property {number} artboardGap - 読み取れた間隔（読めなければ既定値）
     */

    /**
     * 座標の配列に、まだ無い値だけを足す（微小なずれは同じ位置とみなす）
     * @param {number[]} values - 追加先の配列
     * @param {number} value - 追加する座標
     * @returns {void}
     */
    function addUniquePosition(values, value) {
        for (var i = 0; i < values.length; i++) {
            if (Math.abs(values[i] - value) <= POSITION_TOLERANCE) {
                return;
            }
        }
        values.push(value);
    }

    /**
     * 並べ替え済みの座標列から、隣り合う値の最小間隔を求める
     * @param {number[]} sortedValues - 並べ替え済みの座標
     * @returns {number} 最小間隔（値が1つだけなら 0）
     */
    function getSmallestStep(sortedValues) {
        var smallestStep = 0;
        for (var i = 1; i < sortedValues.length; i++) {
            var step = Math.abs(sortedValues[i] - sortedValues[i - 1]);
            if (smallestStep === 0 || step < smallestStep) {
                smallestStep = step;
            }
        }
        return smallestStep;
    }

    /**
     * 既存アートボードの並びを1回の走査で読み取る（追加位置とダイアログ初期値の両方に使う）
     * @param {Document} doc - 対象ドキュメント
     * @returns {ArtboardLayout} 並びの情報
     */
    function readArtboardLayout(doc) {
        var lefts = [];
        var tops = [];
        for (var i = 0; i < doc.artboards.length; i++) {
            var rect = doc.artboards[i].artboardRect;
            addUniquePosition(lefts, rect[0]);
            addUniquePosition(tops, rect[1]);
        }
        lefts.sort(function(a, b) { return a - b; });
        tops.sort(function(a, b) { return b - a; });

        var lastRect = doc.artboards[doc.artboards.length - 1].artboardRect;
        var width = lastRect[2] - lastRect[0];
        var height = lastRect[1] - lastRect[3];
        /* 列ピッチ・行ピッチからアートボードのサイズを引いた残りが間隔 / Pitch minus size leaves the gap */
        var gapX = getSmallestStep(lefts) - width;
        var gapY = getSmallestStep(tops) - height;
        var artboardGap = DEFAULT_ARTBOARD_GAP;
        if (gapX > 0 && gapY > 0) {
            /* 縦横で違うときは狭いほうに合わせる / Take the tighter of the two when they differ */
            artboardGap = Math.min(gapX, gapY);
        } else if (gapX > 0 || gapY > 0) {
            artboardGap = Math.max(gapX, gapY);
        }
        return {
            originLeft:  lefts[0],
            lastLeft:    lastRect[0],
            lastTop:     lastRect[1],
            width:       width,
            height:      height,
            /* アートボードが1枚だと列数は読み取れない / A single artboard tells us nothing about columns */
            columnCount: (lefts.length > 1) ? lefts.length : DEFAULT_COLUMN_COUNT,
            artboardGap: artboardGap
        };
    }

    /**
     * グリッドの指定マスにアートボードを追加する
     * @param {Document} doc - 対象ドキュメント
     * @param {ArtboardLayout} layout - 並びの情報
     * @param {number} stepX - 列の送り
     * @param {number} stepY - 行の送り
     * @param {number} column - 列番号（0始まり）
     * @param {number} rowOffset - 追加を始める行からの行送り
     * @returns {boolean} 追加できたら true
     */
    function addArtboardAt(doc, layout, stepX, stepY, column, rowOffset) {
        var left = layout.originLeft + column * stepX;
        var top = layout.lastTop - rowOffset * stepY;
        /* カンバス（227×227inch）の外は Illustrator が受け付けず Error 1200 になる。
           座標の上限を返すAPIが無いため、実際に追加して可否を見るしかない
           Illustrator rejects rects outside the canvas; there is no API for the limit, so we probe */
        try {
            doc.artboards.add([left, top, left + layout.width, top - layout.height]);
            return true;
        } catch (err) {
            return false;
        }
    }

    /**
     * 必要な枚数に足りるまで、既存の並びを引き継いでアートボードを追加する
     * @param {Document} doc - 対象ドキュメント
     * @param {ArtboardLayout} layout - 並びの情報
     * @param {number} requiredCount - 必要なアートボードの枚数
     * @param {SplitSettings} settings - ダイアログで決めた設定
     * @returns {number} 実際に用意できたアートボードの枚数
     */
    function ensureArtboardCount(doc, layout, requiredCount, settings) {
        if (doc.artboards.length >= requiredCount) {
            return doc.artboards.length;
        }
        var stepX = layout.width + settings.artboardGap;
        var stepY = layout.height + settings.artboardGap;
        var addCount = requiredCount - doc.artboards.length;
        var column = Math.round((layout.lastLeft - layout.originLeft) / stepX);
        var rowOffset = 0;
        for (var i = 0; i < addCount; i++) {
            column++;
            if (column >= settings.columnCount) {
                column = 0;
                rowOffset++;
            }
            if (addArtboardAt(doc, layout, stepX, stepY, column, rowOffset)) {
                continue;
            }
            /* 行の先頭で失敗したら下にも伸ばせないので打ち切り、
               途中なら右端に達しただけなので次の行の先頭で試し直す
               Failing at column 0 means we are out of canvas; otherwise just wrap to the next row */
            if (column === 0) {
                break;
            }
            column = 0;
            rowOffset++;
            if (!addArtboardAt(doc, layout, stepX, stepY, column, rowOffset)) {
                break;
            }
        }
        return doc.artboards.length;
    }

    // =========================================
    // リサイズと配置 / Resize & placement
    // =========================================

    /**
     * オブジェクトをアートボードの幅に合わせて等倍スケールし、中央へ移動する
     * @param {PageItem} item - 対象オブジェクト
     * @param {number[]} artboardRect - [左, 上, 右, 下] の座標
     * @param {number} widthPercent - アートボード幅に対する仕上がり幅（%）
     * @returns {void}
     */
    function placeItemOnArtboard(item, artboardRect, widthPercent) {
        var itemBounds = getItemBounds(item);
        var currentWidth = itemBounds[2] - itemBounds[0];
        var targetWidth = (artboardRect[2] - artboardRect[0]) * widthPercent / 100;
        if (currentWidth > 0 && targetWidth > 0) {
            var scalePercent = targetWidth / currentWidth * 100;
            /* 線幅・パターン・グラデーションも同じ倍率で変形する / Scale strokes, patterns and gradients alike */
            item.resize(scalePercent, scalePercent, true, true, true, true, scalePercent, Transformation.TOPLEFT);
        }

        var artboardCenter = getRectCenter(artboardRect);
        var itemCenter = getRectCenter(getItemBounds(item));
        item.left += artboardCenter[0] - itemCenter[0];
        if (CENTER_VERTICALLY) {
            item.top += artboardCenter[1] - itemCenter[1];
        }
    }

    // =========================================
    // 段落の切り出し / Paragraphs
    // =========================================

    /**
     * 分配対象になる段落の番号を集める（SKIP_EMPTY_PARAGRAPHS に従う）
     * @param {TextFrame} textFrame - 対象テキストオブジェクト
     * @returns {number[]} 段落番号
     */
    function collectParagraphIndexes(textFrame) {
        var indexes = [];
        var paragraphs = textFrame.paragraphs;
        for (var i = 0; i < paragraphs.length; i++) {
            /* 空白と改行だけの段落は空とみなす / A paragraph of whitespace only counts as blank */
            if (SKIP_EMPTY_PARAGRAPHS && /^[\s　]*$/.test(paragraphs[i].contents)) {
                continue;
            }
            indexes.push(i);
        }
        return indexes;
    }

    /**
     * 末尾に残った改行を削除する（空行の分だけ高さが増えるのを防ぐ）
     * @param {TextFrame} textFrame - 対象テキストオブジェクト
     * @returns {void}
     */
    function removeTrailingReturns(textFrame) {
        var characters = textFrame.characters;
        /* 消せない文字に当たっても止まるよう、回数は文字数で頭打ちにする / Cap the loop at the character count */
        for (var i = characters.length; i > 0; i--) {
            var lastCharacter = characters[characters.length - 1];
            if (lastCharacter.contents !== "\r" && lastCharacter.contents !== "\n") {
                return;
            }
            lastCharacter.remove();
        }
    }

    /**
     * テキストオブジェクトから指定した段落だけを残す（書式はそのまま）
     * @param {TextFrame} textFrame - 対象テキストオブジェクト
     * @param {number} paragraphIndex - 残す段落の番号
     * @returns {void}
     */
    function keepOnlyParagraph(textFrame, paragraphIndex) {
        var paragraphs = textFrame.paragraphs;
        /* 後ろから消していけば、残す段落より前の番号がずれない / Deleting from the end keeps the earlier indexes valid */
        for (var i = paragraphs.length - 1; i >= 0; i--) {
            if (i !== paragraphIndex) {
                paragraphs[i].remove();
            }
        }
        removeTrailingReturns(textFrame);
    }

    // =========================================
    // 選択の取得 / Selection
    // =========================================

    /**
     * 文字を部分選択している場合に、その文字を含むテキストオブジェクトを選択し直す
     * @param {Document} doc - 対象ドキュメント
     * @returns {Array} 選択し直したテキストオブジェクト
     */
    function selectTextFramesFromTextRange(doc) {
        var storyFrames = doc.selection.story.textFrames;
        var targetFrames = [];
        for (var i = 0; i < storyFrames.length; i++) {
            targetFrames.push(storyFrames[i]);
        }
        /* 文字編集を抜けてからテキストオブジェクトを選択 / Leave text editing, then select the frames */
        app.executeMenuCommand("deselectall");
        selectItems(doc, targetFrames);
        return targetFrames;
    }

    /**
     * 指定したオブジェクトだけを選択状態にする
     * @param {Document} doc - 対象ドキュメント
     * @param {Array} items - 選択するオブジェクト
     * @returns {void}
     */
    function selectItems(doc, items) {
        doc.selection = null;
        for (var i = 0; i < items.length; i++) {
            items[i].selected = true;
        }
    }

    /**
     * 選択の中から最初のテキストオブジェクトを返す
     * doc.selection はライブ参照になりうるため、変形前に配列へ写して固定する
     * @param {Document} doc - 対象ドキュメント
     * @returns {TextFrame} テキストオブジェクト（見つからない場合は null）
     */
    function getSelectedTextFrame(doc) {
        var selection = doc.selection;
        if (!selection) {
            return null;
        }
        /* 文字を部分選択しているときは selection が TextRange になるため、テキストオブジェクトに置き換える
           A partial text selection comes back as a TextRange; promote it to the text object */
        var items = (selection instanceof Array) ? selection : selectTextFramesFromTextRange(doc);
        for (var i = 0; i < items.length; i++) {
            if (items[i].typename === "TextFrame") {
                return items[i];
            }
        }
        return null;
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * @typedef {object} SplitSettings
     * @property {number} startIndex - 配置を始めるアートボード番号（0始まり）
     * @property {number} widthPercent - アートボード幅に対する仕上がり幅（%）
     * @property {number} columnCount - アートボード追加時の列数
     * @property {number} artboardGap - アートボード追加時の間隔（pt）
     * @property {boolean} keepSource - 元のテキストを残すなら true
     */

    /**
     * 設定ダイアログを表示する
     * @param {number} paragraphCount - 分配対象の段落数
     * @param {number} artboardCount - 現在のアートボード枚数
     * @param {ArtboardLayout} layout - 既存の並び（列数と間隔の初期値に使う）
     * @returns {SplitSettings} 設定（キャンセル時は null）
     */
    function showSettingsDialog(paragraphCount, artboardCount, layout) {
        var dialog = createDialogWindow(getLabel("dialog", "title") + " " + SCRIPT_VERSION);

        var settingsPanel = addPanel(dialog, getLabel("panel", "settings"));
        var startInput = addNumberFieldRow(settingsPanel, "startArtboard", DEFAULT_START_ARTBOARD);
        var widthInput = addNumberFieldRow(settingsPanel, "widthPercent", DEFAULT_WIDTH_PERCENT, "%");
        var columnInput = addNumberFieldRow(settingsPanel, "columnCount", layout.columnCount);
        var gapInput = addNumberFieldRow(settingsPanel, "artboardGap", layout.artboardGap, "pt");

        /* 元のテキストを残す（ラベル幅ぶん字下げして入力欄と頭をそろえる）
           Keep the original text; indented by the label width to line up with the fields */
        var keepSourceRow = settingsPanel.add("group");
        setupRow(keepSourceRow, "left", FIELD_ROW_SPACING);
        addRowLabel(keepSourceRow, "");
        var keepSourceCheckbox = keepSourceRow.add("checkbox", undefined, getLabel("checkbox", "keepSource"));
        keepSourceCheckbox.value = DEFAULT_KEEP_SOURCE;

        /* 段落数とアートボード枚数の確認表示 / Show what was detected */
        var summaryText = dialog.add("statictext", undefined,
            getLabel("info", "summary").replace("%1", paragraphCount).replace("%2", artboardCount));
        summaryText.alignment = ["fill", "center"];

        var buttonBarGroup = dialog.add("group");
        setupRow(buttonBarGroup, "right", BUTTON_BAR_SPACING);
        buttonBarGroup.margins = BUTTON_BAR_MARGINS;
        /* Mac 規約でキャンセル → OK の順 / Cancel before OK per macOS */
        buttonBarGroup.add("button", undefined, getLabel("button", "cancel"), { name: "cancel" });
        buttonBarGroup.add("button", undefined, getLabel("button", "ok"), { name: "ok" });

        if (dialog.show() !== 1) {
            return null;
        }
        return {
            startIndex:   readNumberField(startInput, DEFAULT_START_ARTBOARD, 1, artboardCount) - 1,
            widthPercent: readNumberField(widthInput, DEFAULT_WIDTH_PERCENT, 1, 1000),
            columnCount:  Math.round(readNumberField(columnInput, layout.columnCount, 1, 100)),
            artboardGap:  readNumberField(gapInput, layout.artboardGap, 0, 10000),
            keepSource:   keepSourceCheckbox.value
        };
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 段落ごとにテキストを複製し、開始アートボードから順番に配置する
     * @param {Document} doc - 対象ドキュメント
     * @param {TextFrame} sourceFrame - 分配元のテキストオブジェクト
     * @param {number[]} paragraphIndexes - 分配する段落番号
     * @param {number} placeCount - 実際に配置する段落数
     * @param {SplitSettings} settings - ダイアログで決めた設定
     * @returns {Array} 生成したテキストオブジェクト
     */
    function distributeParagraphs(doc, sourceFrame, paragraphIndexes, placeCount, settings) {
        var createdFrames = [];
        for (var i = 0; i < placeCount; i++) {
            var duplicatedFrame = sourceFrame.duplicate();
            keepOnlyParagraph(duplicatedFrame, paragraphIndexes[i]);
            placeItemOnArtboard(duplicatedFrame, doc.artboards[settings.startIndex + i].artboardRect, settings.widthPercent);
            createdFrames.push(duplicatedFrame);
        }
        return createdFrames;
    }

    /**
     * 選択したテキストを段落ごとに分け、指定したアートボードから順番に配置する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert", "noDocument"));
            return;
        }
        var doc = app.activeDocument;

        var sourceFrame = getSelectedTextFrame(doc);
        if (sourceFrame === null) {
            alert(getLabel("alert", (doc.selection && doc.selection.length) ? "noTextFrame" : "noSelection"));
            return;
        }

        var paragraphIndexes = collectParagraphIndexes(sourceFrame);
        if (paragraphIndexes.length === 0) {
            alert(getLabel("alert", "noParagraph"));
            return;
        }

        var layout = readArtboardLayout(doc);
        var settings = showSettingsDialog(paragraphIndexes.length, doc.artboards.length, layout);
        if (settings === null) {
            return;
        }

        var artboardCount = ensureArtboardCount(doc, layout, settings.startIndex + paragraphIndexes.length, settings);
        /* アートボードを増やしきれなかったときは、置ける分だけ処理する / Place only what fits */
        var placeCount = Math.min(paragraphIndexes.length, artboardCount - settings.startIndex);
        var createdFrames = distributeParagraphs(doc, sourceFrame, paragraphIndexes, placeCount, settings);

        if (!settings.keepSource) {
            sourceFrame.remove();
        }
        selectItems(doc, createdFrames);
        app.redraw();

        if (placeCount < paragraphIndexes.length) {
            alert(getLabel("alert", "artboardLimit").replace("%1", placeCount));
        }
    }

    main();

})();
