#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したオブジェクトを、グリッド／行／列／ランダムのいずれかの方式で複製・配置します。
繰り返し数・間隔・方向・アートボードへの敷き詰めを2カラムのダイアログで指定でき、結果はライブプレビューで確認できます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/DuplicateInGridPlus.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n228720785a71

### Overview

Duplicates and lays out the selected objects as a grid, a row, a column, or at random.
Repeat count, spacing, direction and filling the artboard are set in a two-column dialog, with a live preview of the result.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/DuplicateInGridPlus.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "DuplicateInGridPlus";          /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v2.0.3";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-10-23";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-22";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/DuplicateInGridPlus.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/DuplicateInGridPlus.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n228720785a71"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* プレビュー用レイヤーと一時オブジェクトの識別タグ / Preview layer and temporary-item tag */
    var PREVIEW_LAYER_NAME = "_preview";
    var PREVIEW_ITEM_TAG = "__grid_preview__";

    /* 繰り返し数の下限・上限（スライダーの範囲）/ Repeat count range (slider bounds) */
    var REPEAT_COUNT_MIN = 1;
    var REPEAT_COUNT_MAX = 20;

    /* プレビュー更新の最小間隔（ミリ秒）/ Minimum interval between preview updates (ms) */
    var PREVIEW_THROTTLE_MS = 120;

    /* 画面ズームの範囲 / Zoom range */
    var VIEW_ZOOM_MIN = 0.1;
    var VIEW_ZOOM_MAX = 16;

    // =========================================
    // レイアウト / Layout
    // =========================================
    var WINDOW_MARGINS    = 16;                 /* ウィンドウ外周の余白 / window margin */
    var WINDOW_SPACING    = 12;                 /* ウィンドウ内の要素間隔 / window spacing */
    var PANEL_MARGINS     = [16, 20, 16, 12];   /* パネル余白 [左,上,右,下] / panel margins */
    var PANEL_SPACING     = 8;                  /* パネル内の要素間隔 / panel spacing */
    var COLUMN_SPACING    = 12;                 /* 2カラムの間隔 / gap between columns */
    var FIELD_ROW_SPACING = 20;                 /* 入力欄と［連動］の間隔 / gap between fields and the link checkbox */
    var FIELD_CHARS       = 4;                  /* 数値入力欄の文字数 / width of numeric fields */
    var ZOOM_SLIDER_WIDTH = 240;                /* ズームスライダーの幅 / zoom slider width */
    var ZOOM_GROUP_MARGINS = [0, 0, 0, 10];     /* ズームの行の余白 / zoom row margins */
    var DIALOG_OFFSET_X   = 300;                /* ダイアログを右へずらす量 / horizontal dialog offset */
    var DIALOG_OPACITY    = 0.98;               /* ダイアログの不透明度 / dialog opacity */

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * 現在のUI言語を判定する
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        return ($.locale && $.locale.toLowerCase().indexOf("ja") === 0) ? "ja" : "en";
    }

    var uiLang = getCurrentLang();

    /* ラベル定義 / Label definitions (JA/EN) */
    var LABELS = {
        dialog: {
            title: { ja: "複製配置（グリッド／行／列／ランダム）", en: "Duplicate & Arrange" }
        },
        panel: {
            repeatCount: { ja: "繰り返し数", en: "Count" },
            repeatMethod: { ja: "繰り返し方式", en: "Repeat Method" },
            gap: { ja: "間隔（{unit}）", en: "Gap ({unit})" },
            direction: { ja: "方向", en: "Direction" },
            fill: { ja: "敷き詰め", en: "Fill" }
        },
        fieldLabel: {
            countHorizontal: { ja: "横", en: "Horizontal" },
            countVertical: { ja: "縦", en: "Vertical" },
            gapHorizontal: { ja: "左右", en: "Horizontal" },
            gapVertical: { ja: "上下", en: "Vertical" },
            directionHorizontal: { ja: "横方向", en: "Horizontal" },
            directionVertical: { ja: "縦方向", en: "Vertical" },
            zoom: { ja: "画面ズーム", en: "Zoom" }
        },
        radio: {
            methodGrid: { ja: "グリッド", en: "Grid" },
            methodRow: { ja: "行", en: "Row" },
            methodColumn: { ja: "列", en: "Column" },
            methodRandom: { ja: "ランダム配置", en: "Random" },
            directionRight: { ja: "右", en: "Right" },
            directionLeft: { ja: "左", en: "Left" },
            directionUp: { ja: "上", en: "Up" },
            directionDown: { ja: "下", en: "Down" }
        },
        checkbox: {
            link: { ja: "連動", en: "Link Horizontal & Vertical" },
            fillToEdge: { ja: "アートボードの端まで", en: "Fill to Artboard Edge" },
            fillFull: { ja: "アートボードいっぱいに", en: "Fill Full Artboard" },
            lightMode: { ja: "軽量モード", en: "Light mode" }
        },
        tooltip: {
            countHorizontal: { ja: "横方向に並べる数です（元のオブジェクトを含む）。", en: "How many to place horizontally, including the original." },
            countVertical: { ja: "縦方向に並べる数です（元のオブジェクトを含む）。", en: "How many to place vertically, including the original." },
            countLink: { ja: "横と縦の数を同じにします。", en: "Keeps the horizontal and vertical counts the same." },
            countSlider: { ja: "繰り返し数をまとめて変更します。", en: "Changes the counts together." },
            methodGrid: { ja: "横と縦の両方に並べます。", en: "Places copies both horizontally and vertically." },
            methodRow: { ja: "横一列に並べます。", en: "Places copies in a single row." },
            methodColumn: { ja: "縦一列に並べます。", en: "Places copies in a single column." },
            methodRandom: { ja: "グリッドの枠内でランダムにずらして配置します。", en: "Scatters the copies randomly within the grid area." },
            gapHorizontal: { ja: "隣り合うオブジェクトの左右のアキです。", en: "Space between neighbouring objects horizontally." },
            gapVertical: { ja: "隣り合うオブジェクトの上下のアキです。", en: "Space between neighbouring objects vertically." },
            gapLink: { ja: "左右と上下の間隔を同じにします。", en: "Keeps the horizontal and vertical gaps the same." },
            directionRight: { ja: "元のオブジェクトの右へ複製します。", en: "Duplicates to the right of the original." },
            directionLeft: { ja: "元のオブジェクトの左へ複製します。", en: "Duplicates to the left of the original." },
            directionUp: { ja: "元のオブジェクトの上へ複製します。", en: "Duplicates above the original." },
            directionDown: { ja: "元のオブジェクトの下へ複製します。", en: "Duplicates below the original." },
            fillToEdge: { ja: "アートボードの端に届くまで数を自動で増やします。", en: "Increases the count automatically until the copies reach the artboard edge." },
            fillFull: { ja: "アートボード全面を埋めるように、元の位置に関係なく敷き詰めます。", en: "Tiles the whole artboard, ignoring the original position." },
            lightMode: { ja: "プレビューを簡易表示にして、重いオブジェクトでも操作を軽くします。", en: "Simplifies the preview so heavy objects stay responsive." },
            zoom: { ja: "作業中の画面表示倍率を変えます。結果には影響しません。", en: "Changes the view zoom while you work. It does not affect the result." }
        },
        button: {
            ok: { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noSelection: { ja: "オブジェクトを選択してください。", en: "Please select an object." },
            invalidCount: { ja: "繰り返し数は1以上の整数を入力してください。", en: "Enter an integer count of 1 or more." },
            invalidGap: { ja: "間隔は数値で入力してください。", en: "Enter a numeric gap value." }
        }
    };

    /**
     * ドット区切りのキーからUI言語のラベルを取得する
     * @param {string} labelPath - "panel.gap" のようなドット区切りのキー
     * @param {object} [params] - {unit: "mm"} のような差し込み値
     * @returns {string} 表示言語のテキスト（見つからない場合は labelPath をそのまま返す）
     */
    function getLabel(labelPath, params) {
        var pathKeys = labelPath.split(".");
        var labelNode = LABELS;
        for (var i = 0; i < pathKeys.length; i++) {
            labelNode = labelNode[pathKeys[i]];
            if (!labelNode) return labelPath;
        }
        var localizedText = labelNode[uiLang] || labelNode.en;
        if (!localizedText) return labelPath;
        if (params) {
            for (var paramKey in params) {
                if (!params.hasOwnProperty(paramKey)) continue;
                localizedText = localizedText.replace(new RegExp("\\{" + paramKey + "\\}", "g"), params[paramKey]);
            }
        }
        return localizedText;
    }

    /**
     * コロン付きの項目名を返す（日本語は全角、英語は半角）
     * @param {string} labelPath - "fieldLabel.countHorizontal" のようなドット区切りのキー
     * @returns {string} コロンを添えたラベル
     */
    function labelText(labelPath) {
        return getLabel(labelPath) + (uiLang === "ja" ? "：" : ":");
    }

    // =========================================
    // UIレイアウト補助 / UI layout helpers
    // =========================================

    /**
     * パネルの共通設定を適用する
     * @param {Panel} targetPanel - 対象のパネル
     * @param {number} [spacing] - パネル内の要素間隔
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
     * 行グループの共通設定を適用する
     * @param {Group} targetGroup - 対象のグループ
     * @param {string} [alignment] - グループ自身の配置（既定は "left"）
     * @param {number} [spacing] - グループ内の要素間隔
     * @returns {void}
     */
    function setupRow(targetGroup, alignment, spacing) {
        targetGroup.orientation = "row";
        targetGroup.alignChildren = ["left", "center"];
        targetGroup.alignment = alignment || "left";
        targetGroup.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /**
     * ラベル付きパネルを生成する（共通レイアウト適用）
     * @param {Group|Window} parentContainer - 追加先
     * @param {string} panelTitle - パネルのタイトル
     * @returns {Panel} 生成したパネル
     */
    function addPanel(parentContainer, panelTitle) {
        var createdPanel = parentContainer.add("panel");
        createdPanel.text = panelTitle;
        setupPanel(createdPanel);
        return createdPanel;
    }

    /**
     * 左寄せの縦並びグループを生成する（入力欄の列・チェックボックス列など）
     * @param {Group|Panel} parentContainer - 追加先
     * @returns {Group} 生成したグループ
     */
    function addColumnGroup(parentContainer) {
        var columnGroup = parentContainer.add("group");
        columnGroup.orientation = "column";
        columnGroup.alignChildren = ["left", "center"];
        return columnGroup;
    }

    /**
     * ダイアログウィンドウを生成する
     * @param {string} title - タイトルバーの文字列
     * @returns {Window} 生成したダイアログ
     */
    function createDialogWindow(title) {
        var dialogWindow = new Window("dialog", title);
        dialogWindow.orientation = "column";
        dialogWindow.alignChildren = ["fill", "top"];
        dialogWindow.spacing = WINDOW_SPACING;
        dialogWindow.margins = WINDOW_MARGINS;
        dialogWindow.opacity = DIALOG_OPACITY;
        return dialogWindow;
    }

    /**
     * 数値入力欄を上下キーで増減できるようにする（shift:10単位 / option:0.1単位）
     * @param {EditText} numericField - 対象のテキストフィールド
     * @returns {void}
     */
    function changeValueByArrowKey(numericField) {
        numericField.addEventListener("keydown", function (event) {
            var fieldValue = Number(numericField.text);
            if (isNaN(fieldValue)) return;
            var keyboardState = ScriptUI.environment.keyboardState;

            if (keyboardState.shiftKey) {
                if (event.keyName == "Up") { fieldValue = Math.ceil((fieldValue + 1) / 10) * 10; event.preventDefault(); }
                else if (event.keyName == "Down") { fieldValue = Math.floor((fieldValue - 1) / 10) * 10; event.preventDefault(); }
            } else if (keyboardState.altKey) {
                if (event.keyName == "Up") { fieldValue += 0.1; event.preventDefault(); }
                else if (event.keyName == "Down") { fieldValue -= 0.1; event.preventDefault(); }
            } else {
                if (event.keyName == "Up") { fieldValue += 1; event.preventDefault(); }
                else if (event.keyName == "Down") { fieldValue -= 1; event.preventDefault(); }
            }

            fieldValue = keyboardState.altKey ? Math.round(fieldValue * 10) / 10 : Math.round(fieldValue);
            if (numericField.isInteger) fieldValue = Math.max(0, Math.round(fieldValue));
            numericField.text = fieldValue;
            if (typeof numericField.onChanging === "function") numericField.onChanging();
        });
    }

    /**
     * ラジオボタンを選択し、クリック時と同じ処理を走らせる
     * @param {RadioButton} radioButton - 選択するラジオボタン
     * @returns {void}
     */
    function selectRadioButton(radioButton) {
        if (!radioButton || radioButton.enabled === false) return;
        radioButton.value = true;
        if (typeof radioButton.onClick === "function") radioButton.onClick();
    }

    /**
     * ラベルと tooltip が同じキーのラジオボタン／チェックボックスを追加する
     * @param {Group|Panel} parentContainer - 追加先
     * @param {string} controlType - "radiobutton" または "checkbox"
     * @param {string} labelKey - LABELS.radio（または LABELS.checkbox）と LABELS.tooltip に共通のキー
     * @returns {RadioButton|Checkbox} 追加したコントロール
     */
    function addLabeledControl(parentContainer, controlType, labelKey) {
        var labelCategory = (controlType === "radiobutton") ? "radio" : "checkbox";
        var labeledControl = parentContainer.add(controlType, undefined, getLabel(labelCategory + "." + labelKey));
        labeledControl.helpTip = getLabel("tooltip." + labelKey);
        return labeledControl;
    }

    /**
     * 項目名付きの数値入力欄を追加する（上下キーで増減できる）
     * @param {Group} parentGroup - 追加先
     * @param {string} fieldKey - LABELS.fieldLabel と LABELS.tooltip に共通のキー
     * @param {string} initialText - 初期値
     * @param {boolean} isInteger - 整数だけにするなら true
     * @returns {EditText} 追加した入力欄
     */
    function addNumericField(parentGroup, fieldKey, initialText, isInteger) {
        var fieldGroup = parentGroup.add("group");
        fieldGroup.add("statictext", undefined, labelText("fieldLabel." + fieldKey));
        var numericInput = fieldGroup.add("edittext", undefined, initialText);
        numericInput.helpTip = getLabel("tooltip." + fieldKey);
        numericInput.characters = FIELD_CHARS;
        if (isInteger) numericInput.isInteger = true;
        changeValueByArrowKey(numericInput);
        return numericInput;
    }

    /**
     * 横／縦の数値入力欄と［連動］チェックボックスの組を追加する
     * @param {Panel} parentPanel - 追加先
     * @param {string} horizontalKey - 横の入力欄のキー（LABELS.fieldLabel / LABELS.tooltip）
     * @param {string} verticalKey - 縦の入力欄のキー（LABELS.fieldLabel / LABELS.tooltip）
     * @param {string} linkTooltipKey - ［連動］の tooltip のキー
     * @param {string} initialText - 入力欄の初期値
     * @param {boolean} isInteger - 整数だけにするなら true
     * @returns {{horizontalInput: EditText, verticalInput: EditText, linkCheck: Checkbox}} 追加したコントロール
     */
    function addLinkedFieldPair(parentPanel, horizontalKey, verticalKey, linkTooltipKey, initialText, isInteger) {
        var pairRow = parentPanel.add("group");
        setupRow(pairRow, "left", FIELD_ROW_SPACING);
        pairRow.alignChildren = ["left", "top"];

        var fieldsColumn = addColumnGroup(pairRow);
        var horizontalInput = addNumericField(fieldsColumn, horizontalKey, initialText, isInteger);
        var verticalInput = addNumericField(fieldsColumn, verticalKey, initialText, isInteger);

        var linkGroup = addColumnGroup(pairRow);
        var linkCheck = linkGroup.add("checkbox", undefined, getLabel("checkbox.link"));
        linkCheck.helpTip = getLabel("tooltip." + linkTooltipKey);
        linkCheck.value = true;

        return { horizontalInput: horizontalInput, verticalInput: verticalInput, linkCheck: linkCheck };
    }

    /**
     * 項目名と2つのラジオボタンの行を追加する（方向の指定用）
     * @param {Panel} parentPanel - 追加先
     * @param {string} labelKey - LABELS.fieldLabel のキー
     * @param {string} firstKey - 1つ目のラジオのキー（LABELS.radio / LABELS.tooltip）
     * @param {string} secondKey - 2つ目のラジオのキー（LABELS.radio / LABELS.tooltip）
     * @returns {RadioButton[]} 追加したラジオボタン [1つ目, 2つ目]
     */
    function addDirectionRow(parentPanel, labelKey, firstKey, secondKey) {
        var directionRow = parentPanel.add("group");
        setupRow(directionRow);
        directionRow.add("statictext", undefined, labelText("fieldLabel." + labelKey));
        var firstRadio = addLabeledControl(directionRow, "radiobutton", firstKey);
        var secondRadio = addLabeledControl(directionRow, "radiobutton", secondKey);
        return [firstRadio, secondRadio];
    }

    /**
     * 2カラムレイアウトの列を追加する
     * @param {Group} parentGroup - 追加先
     * @returns {Group} 追加した列
     */
    function addLayoutColumn(parentGroup) {
        var layoutColumn = parentGroup.add("group");
        layoutColumn.orientation = "column";
        layoutColumn.alignChildren = "fill";
        layoutColumn.spacing = WINDOW_SPACING;
        return layoutColumn;
    }

    // =========================================
    // 単位 / Units
    // =========================================

    /* 単位テーブル（配列の添字が rulerType コードと一致：0=in, 1=mm, 2=pt …）/ Unit table; the array index equals the rulerType code */
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
     * 設定キーごとの単位情報を取得する
     * @param {string} prefKey - 環境設定キー（省略時は "rulerType"）
     * @returns {{code: number, label: string, pointsPerUnit: number}} 単位情報
     */
    function getUnitInfo(prefKey) {
        var unitKey = prefKey || "rulerType";
        var unitCode = app.preferences.getIntegerPreference(unitKey);
        var unit = UNITS[unitCode] || UNITS[2];
        var label = (unitCode === 5 && HA_UNIT_PREF_KEYS[unitKey]) ? "H" : unit.label;
        return { code: unitCode, label: label, pointsPerUnit: unit.pointsPerUnit };
    }

    // =========================================
    // 境界と複製 / Bounds and duplication
    // =========================================

    /**
     * クリッピングマスクを優先して境界を取得する（なければ可視境界）
     * @param {PageItem} targetItem - 対象オブジェクト
     * @returns {Array<number>} [左, 上, 右, 下]
     */
    function getMaskedBounds(targetItem) {
        try {
            if (targetItem.typename === "GroupItem" && targetItem.clipped) {
                for (var i = 0; i < targetItem.pageItems.length; i++) {
                    var childItem = targetItem.pageItems[i];
                    if (childItem.typename === "PathItem" && childItem.clipping) return childItem.geometricBounds;
                    if (childItem.typename === "CompoundPathItem" && childItem.pathItems.length > 0 && childItem.pathItems[0].clipping)
                        return childItem.pathItems[0].geometricBounds;
                }
            }
            if (targetItem.typename === "PathItem" && targetItem.clipping) return targetItem.geometricBounds;
            if (targetItem.typename === "CompoundPathItem" && targetItem.pathItems.length > 0 && targetItem.pathItems[0].clipping)
                return targetItem.pathItems[0].geometricBounds;
        } catch (e) { }
        return targetItem.visibleBounds;
    }

    /**
     * アクティブなアートボードの矩形を取得する
     * @param {Document} doc - 対象ドキュメント
     * @returns {Array<number>} [左, 上, 右, 下]
     */
    function getActiveArtboardRect(doc) {
        return doc.artboards[doc.artboards.getActiveArtboardIndex()].artboardRect;
    }

    /**
     * 複数オブジェクトを囲む境界を求める
     * @param {Array<PageItem>} targetItems - 対象オブジェクトの配列
     * @returns {Array<number>} [左, 上, 右, 下]
     */
    function getUnionBounds(targetItems) {
        var unionLeft = Infinity, unionTop = -Infinity, unionRight = -Infinity, unionBottom = Infinity;
        for (var i = 0; i < targetItems.length; i++) {
            var itemBounds = getMaskedBounds(targetItems[i]);
            if (itemBounds[0] < unionLeft) unionLeft = itemBounds[0];
            if (itemBounds[1] > unionTop) unionTop = itemBounds[1];
            if (itemBounds[2] > unionRight) unionRight = itemBounds[2];
            if (itemBounds[3] < unionBottom) unionBottom = itemBounds[3];
        }
        return [unionLeft, unionTop, unionRight, unionBottom];
    }

    /**
     * グリッド配置のオフセット一覧を作る（元オブジェクトのぶんは含まない）
     * @param {number} rowCount - 縦の数
     * @param {number} columnCount - 横の数
     * @param {number} sourceWidth - 元オブジェクトの幅（pt）
     * @param {number} sourceHeight - 元オブジェクトの高さ（pt）
     * @param {number} gapX - 左右の間隔（pt）
     * @param {number} gapY - 上下の間隔（pt）
     * @param {string} horizontalDirection - "right" または "left"
     * @param {string} verticalDirection - "up" または "down"
     * @returns {Array<Array<number>>} [dx, dy] の配列（pt）
     */
    function computeGridOffsets(rowCount, columnCount, sourceWidth, sourceHeight, gapX, gapY, horizontalDirection, verticalDirection) {
        var gridOffsets = [];
        for (var rowIndex = 0; rowIndex < rowCount; rowIndex++) {
            for (var columnIndex = 0; columnIndex < columnCount; columnIndex++) {
                if (rowIndex === 0 && columnIndex === 0) continue;
                var dx = (sourceWidth + gapX) * columnIndex;
                if (horizontalDirection === "left") dx = -dx;
                var dy = (sourceHeight + gapY) * rowIndex;
                if (verticalDirection !== "up") dy = -dy;
                gridOffsets.push([dx, dy]);
            }
        }
        return gridOffsets;
    }

    /**
     * 元オブジェクトを指定オフセットぶん複製する
     * @param {Array<PageItem>} sourceItems - 複製元（複数可。まとめて同じ量だけずらす）
     * @param {Array<Array<number>>} placementOffsets - [dx, dy] の配列（pt）
     * @param {Layer} [targetLayer] - 複製先レイヤー（省略時は元と同じ場所に複製）
     * @returns {Array<PageItem>} 生成した複製の配列
     */
    function duplicateWithOffsets(sourceItems, placementOffsets, targetLayer) {
        var duplicatedItems = [];

        for (var i = 0; i < placementOffsets.length; i++) {
            var dx = placementOffsets[i][0];
            var dy = placementOffsets[i][1];
            for (var j = 0; j < sourceItems.length; j++) {
                /* duplicate() は同じ座標に作られるので、オフセットぶん動かすだけでよい
                   duplicate() keeps the original coordinates, so shifting by the offset is enough */
                var duplicatedItem = targetLayer
                    ? sourceItems[j].duplicate(targetLayer, ElementPlacement.PLACEATBEGINNING)
                    : sourceItems[j].duplicate();
                if (targetLayer) duplicatedItem.note = PREVIEW_ITEM_TAG;

                duplicatedItem.left += dx;
                duplicatedItem.top += dy;
                duplicatedItems.push(duplicatedItem);
            }
        }
        return duplicatedItems;
    }

    // =========================================
    // プレビュー / Preview
    // =========================================

    /**
     * プレビュー用レイヤーを取得する（なければ作成し、最前面へ）
     * @param {Document} doc - 対象ドキュメント
     * @returns {Layer} プレビュー用レイヤー
     */
    function getPreviewLayer(doc) {
        var previewLayer;
        try {
            previewLayer = doc.layers.getByName(PREVIEW_LAYER_NAME);
        } catch (e) {
            previewLayer = doc.layers.add();
            previewLayer.name = PREVIEW_LAYER_NAME;
        }
        previewLayer.visible = true;
        previewLayer.locked = false;
        previewLayer.zOrder(ZOrderMethod.BRINGTOFRONT);
        return previewLayer;
    }

    /**
     * プレビューで作った一時オブジェクトだけを削除する（noteタグで判別）
     * @param {Document} doc - 対象ドキュメント
     * @param {boolean} [skipRedraw] - true なら再描画を呼ばない（直後に描き直す場合）
     * @returns {void}
     */
    function clearPreview(doc, skipRedraw) {
        var previewLayer;
        try {
            previewLayer = doc.layers.getByName(PREVIEW_LAYER_NAME);
        } catch (e) {
            return;
        }
        /* pageItems はグループの子まで含むため、削除で添字がずれても止まらないよう1件ずつ受ける
           pageItems includes group children, so guard each removal against the shifting index */
        for (var i = previewLayer.pageItems.length - 1; i >= 0; i--) {
            try {
                var previewItem = previewLayer.pageItems[i];
                if (previewItem.note === PREVIEW_ITEM_TAG) previewItem.remove();
            } catch (e) { }
        }
        if (!skipRedraw) app.redraw();
    }

    /**
     * プレビューを描き直す
     * @param {Document} doc - 対象ドキュメント
     * @param {Array<PageItem>} sourceItems - 複製元
     * @param {Array<Array<number>>} placementOffsets - [dx, dy] の配列（pt）
     * @returns {void}
     */
    function renderPreview(doc, sourceItems, placementOffsets) {
        var previewLayer = getPreviewLayer(doc);
        /* 空の状態を挟むとちらつくので、消去時は再描画しない / Skip the intermediate repaint so the canvas never flashes empty */
        clearPreview(doc, true);
        duplicateWithOffsets(sourceItems, placementOffsets, previewLayer);
        app.redraw();
    }

    // =========================================
    // 画面ズーム / View zoom
    // 軽量モードではスライダーを離したときだけ適用 / Light mode applies the zoom only on release
    // =========================================

    /**
     * 現在のビュー状態（ズーム倍率と中心）を控える
     * @param {Document} doc - 対象ドキュメント
     * @returns {object} {view, zoom, center} の状態オブジェクト
     */
    function captureViewState(doc) {
        var viewState = { view: null, zoom: null, center: null };
        /* ビューが取れないことがある / The view may be unavailable */
        try {
            viewState.view = doc.activeView;
            viewState.zoom = viewState.view.zoom;
            viewState.center = viewState.view.centerPoint;
        } catch (e) { }
        return viewState;
    }

    /**
     * 控えておいたビュー状態を復元する
     * @param {Document} doc - 対象ドキュメント
     * @param {object} viewState - captureViewState() の戻り値
     * @returns {void}
     */
    function restoreViewState(doc, viewState) {
        if (!viewState) return;
        /* ビューが閉じられていることがある / The view may be gone */
        try {
            var targetView = viewState.view || doc.activeView;
            if (targetView && viewState.zoom != null) targetView.zoom = viewState.zoom;
            if (targetView && viewState.center != null) targetView.centerPoint = viewState.center;
        } catch (e) { }
    }

    /**
     * 画面ズーム用のスライダーと軽量モードのチェックボックスを追加する
     * @param {Group|Window} parentContainer - 追加先
     * @param {Document} doc - 対象ドキュメント
     * @param {object} initialViewState - captureViewState() の戻り値
     * @returns {{restoreInitial: Function}} 開いたときのビュー状態に戻す関数
     */
    function addZoomControls(parentContainer, doc, initialViewState) {
        var zoomGroup = parentContainer.add("group");
        zoomGroup.orientation = "row";
        zoomGroup.alignChildren = ["center", "center"];
        zoomGroup.alignment = "center";
        zoomGroup.margins = ZOOM_GROUP_MARGINS;

        zoomGroup.add("statictext", undefined, labelText("fieldLabel.zoom"));

        var initialZoom = 1;
        /* ビューが取れないことがある / The view may be unavailable */
        try {
            initialZoom = Number((initialViewState.zoom != null) ? initialViewState.zoom : doc.activeView.zoom);
        } catch (e) { }
        if (!initialZoom || isNaN(initialZoom)) initialZoom = 1;

        var zoomSlider = zoomGroup.add("slider", undefined, initialZoom, VIEW_ZOOM_MIN, VIEW_ZOOM_MAX);
        zoomSlider.helpTip = getLabel("tooltip.zoom");
        zoomSlider.preferredSize.width = ZOOM_SLIDER_WIDTH;

        var lightModeCheck = zoomGroup.add("checkbox", undefined, getLabel("checkbox.lightMode"));
        lightModeCheck.helpTip = getLabel("tooltip.lightMode");
        lightModeCheck.value = false;

        /**
         * 指定倍率をビューへ適用する
         * @param {number} zoomLevel - ズーム倍率
         * @returns {void}
         */
        function applyZoom(zoomLevel) {
            /* ビューが閉じられている・範囲外の倍率のときは何もしない / Ignore a missing view or an out-of-range zoom */
            try {
                var targetView = initialViewState.view ? initialViewState.view : doc.activeView;
                if (!targetView) return;
                targetView.zoom = zoomLevel;
                app.redraw();
            } catch (e) { }
        }

        /* ドラッグ中の追従（軽量モードでは無効）/ Live drag (disabled in light mode) */
        zoomSlider.onChanging = function () {
            if (lightModeCheck.value) return;
            applyZoom(Number(zoomSlider.value));
        };

        /* 離したときは必ず1回適用 / Always apply once on release */
        zoomSlider.onChange = function () {
            applyZoom(Number(zoomSlider.value));
        };

        lightModeCheck.onClick = function () {
            applyZoom(Number(zoomSlider.value));
        };

        return {
            restoreInitial: function () { restoreViewState(doc, initialViewState); }
        };
    }

    // =========================================
    // 繰り返し数とランダム配置 / Repeat counts and random placement
    // =========================================

    /**
     * ［アートボードの端まで］：選択オブジェクトを起点に、アートボードの端まで並ぶ行列数を求める
     * @param {number[]} artboardRect - アートボードの矩形 [左, 上, 右, 下]
     * @param {number[]} sourceBounds - 複製元の境界 [左, 上, 右, 下]
     * @param {number} sourceWidth - 複製元の幅（pt）
     * @param {number} sourceHeight - 複製元の高さ（pt）
     * @param {{x: number, y: number}} gapPoints - 左右・上下の間隔（pt）
     * @param {string} repeatMethod - "grid" / "row" / "column" / "random"
     * @param {boolean} isRightward - 右方向に並べるなら true（false なら左）
     * @param {boolean} isUpward - 上方向に並べるなら true（false なら下）
     * @returns {{columns: number, rows: number}} 横と縦の数
     */
    function computeCountsToArtboardEdge(artboardRect, sourceBounds, sourceWidth, sourceHeight, gapPoints, repeatMethod, isRightward, isUpward) {
        var sourceLeft = sourceBounds[0], sourceTop = sourceBounds[1];
        var stepWidth = sourceWidth + gapPoints.x, stepHeight = sourceHeight + gapPoints.y;
        var columnCount = 1, rowCount = 1;

        if (repeatMethod !== "column" && stepWidth > 0) {
            var availableWidth = isRightward ? (artboardRect[2] - sourceLeft) : (sourceLeft - artboardRect[0]);
            columnCount = Math.max(1, Math.floor((availableWidth + gapPoints.x) / stepWidth));
        }
        if (repeatMethod !== "row" && stepHeight > 0) {
            var availableHeight = isUpward ? (artboardRect[1] - sourceTop) : (sourceTop - artboardRect[3]);
            rowCount = Math.max(1, Math.floor((availableHeight + gapPoints.y) / stepHeight));
        }
        return { columns: columnCount, rows: rowCount };
    }

    /**
     * ［アートボードいっぱいに］：アートボードに収まる最大の行列数を求める（方向は無視）
     * @param {number[]} artboardRect - アートボードの矩形 [左, 上, 右, 下]
     * @param {number} sourceWidth - 複製元の幅（pt）
     * @param {number} sourceHeight - 複製元の高さ（pt）
     * @param {{x: number, y: number}} gapPoints - 左右・上下の間隔（pt）
     * @param {string} repeatMethod - "grid" / "row" / "column" / "random"
     * @returns {{columns: number, rows: number}} 横と縦の数
     */
    function computeCountsToFullArtboard(artboardRect, sourceWidth, sourceHeight, gapPoints, repeatMethod) {
        var artboardWidth = Math.abs(artboardRect[2] - artboardRect[0]);
        var artboardHeight = Math.abs(artboardRect[1] - artboardRect[3]);
        var stepWidth = sourceWidth + gapPoints.x, stepHeight = sourceHeight + gapPoints.y;

        return {
            columns: (repeatMethod === "column" || stepWidth <= 0)
                ? 1 : Math.max(1, Math.floor((artboardWidth + gapPoints.x) / stepWidth)),
            rows: (repeatMethod === "row" || stepHeight <= 0)
                ? 1 : Math.max(1, Math.floor((artboardHeight + gapPoints.y) / stepHeight))
        };
    }

    /**
     * 数値を繰り返し数の範囲に収める
     * @param {string|number} countValue - 入力値
     * @returns {number} REPEAT_COUNT_MIN〜REPEAT_COUNT_MAX に収めた整数
     */
    function clampRepeatCount(countValue) {
        /* スライダーは実数を返すので、切り捨てず四捨五入する / Sliders report real numbers, so round instead of truncating */
        var repeatCount = Math.round(Number(countValue));
        if (isNaN(repeatCount) || repeatCount < REPEAT_COUNT_MIN) return REPEAT_COUNT_MIN;
        if (repeatCount > REPEAT_COUNT_MAX) return REPEAT_COUNT_MAX;
        return repeatCount;
    }

    /**
     * 線形合同法で擬似乱数を返す（同じ種から同じ並びを再現するため）
     * @param {object} seedHolder - {v: number} 形式の内部状態
     * @returns {number} 0以上1未満の擬似乱数
     */
    function nextRandomValue(seedHolder) {
        seedHolder.v = (seedHolder.v * 1664525 + 1013904223) % 4294967296;
        return seedHolder.v / 4294967296;
    }

    /**
     * ランダム配置のオフセット一覧を作る
     * @param {number} duplicateCount - 複製する数
     * @param {number} rangeX - 左右の散らばり範囲（pt）
     * @param {number} rangeY - 上下の散らばり範囲（pt）
     * @param {number} randomSeed - 乱数の種
     * @returns {Array<Array<number>>} [dx, dy] の配列（pt）
     */
    function createRandomOffsets(duplicateCount, rangeX, rangeY, randomSeed) {
        var seedHolder = { v: randomSeed >>> 0 };
        var randomOffsets = [];
        for (var i = 0; i < duplicateCount; i++) {
            var dx = (rangeX > 0) ? (-rangeX + 2 * rangeX * nextRandomValue(seedHolder)) : 0;
            var dy = (rangeY > 0) ? (-rangeY + 2 * rangeY * nextRandomValue(seedHolder)) : 0;
            randomOffsets.push([dx, dy]);
        }
        return randomOffsets;
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 設定ダイアログを組み立てる（イベントは showDuplicateDialog() で結線する）
     * @param {Document} doc - 対象ドキュメント
     * @param {string} rulerUnitLabel - 定規の単位の表示名
     * @returns {object} ダイアログ・各コントロール・ズームの操作をまとめたオブジェクト
     */
    function buildDuplicateDialog(doc, rulerUnitLabel) {
        var duplicateDialog = createDialogWindow(getLabel("dialog.title") + " " + SCRIPT_VERSION);

        /* 2カラムレイアウト：左（繰り返し数／方式）、右（間隔／方向／敷き詰め）
           Two-column layout: counts & method on the left, gap / direction / fill on the right */
        var columnsGroup = duplicateDialog.add("group");
        columnsGroup.orientation = "row";
        columnsGroup.alignChildren = ["fill", "top"];
        columnsGroup.spacing = COLUMN_SPACING;

        var leftColumnGroup = addLayoutColumn(columnsGroup);
        var rightColumnGroup = addLayoutColumn(columnsGroup);

        /* 繰り返し数 / Repeat count */
        var repeatCountPanel = addPanel(leftColumnGroup, getLabel("panel.repeatCount"));
        var countFields = addLinkedFieldPair(repeatCountPanel, "countHorizontal", "countVertical", "countLink", "2", true);

        var countSliderGroup = repeatCountPanel.add("group");
        countSliderGroup.orientation = "row";
        countSliderGroup.alignChildren = ["fill", "center"];
        var countSlider = countSliderGroup.add("slider", undefined, 2, REPEAT_COUNT_MIN, REPEAT_COUNT_MAX);
        countSlider.helpTip = getLabel("tooltip.countSlider");
        countSlider.alignment = ["fill", "center"];

        /* 繰り返し方式 / Repeat method */
        var repeatMethodPanel = addPanel(leftColumnGroup, getLabel("panel.repeatMethod"));
        repeatMethodPanel.alignChildren = ["left", "top"];
        var methodGridRadio = addLabeledControl(repeatMethodPanel, "radiobutton", "methodGrid");
        var methodRowRadio = addLabeledControl(repeatMethodPanel, "radiobutton", "methodRow");
        var methodColumnRadio = addLabeledControl(repeatMethodPanel, "radiobutton", "methodColumn");
        var methodRandomRadio = addLabeledControl(repeatMethodPanel, "radiobutton", "methodRandom");
        methodGridRadio.value = true;

        /* 間隔（現在の定規単位で入力し、内部ではptへ変換）
           Gap (entered in the current ruler unit, converted to points internally) */
        var gapPanel = addPanel(rightColumnGroup, getLabel("panel.gap", { unit: rulerUnitLabel }));
        var gapFields = addLinkedFieldPair(gapPanel, "gapHorizontal", "gapVertical", "gapLink", "10", false);

        /* 方向 / Direction */
        var directionPanel = addPanel(rightColumnGroup, getLabel("panel.direction"));
        directionPanel.alignChildren = ["left", "top"];
        var horizontalDirectionRadios = addDirectionRow(directionPanel, "directionHorizontal", "directionRight", "directionLeft");
        horizontalDirectionRadios[0].value = true;
        var verticalDirectionRadios = addDirectionRow(directionPanel, "directionVertical", "directionUp", "directionDown");
        verticalDirectionRadios[1].value = true;

        /* 敷き詰め / Fill */
        var fillPanel = addPanel(rightColumnGroup, getLabel("panel.fill"));
        fillPanel.alignChildren = ["left", "top"];
        var fillToEdgeCheck = addLabeledControl(fillPanel, "checkbox", "fillToEdge");
        fillToEdgeCheck.value = false;
        var fillFullCheck = addLabeledControl(fillPanel, "checkbox", "fillFull");
        fillFullCheck.value = false;

        /* 画面ズーム / Zoom */
        var zoomControls = addZoomControls(duplicateDialog, doc, captureViewState(doc));

        var btnRowGroup = duplicateDialog.add("group");
        setupRow(btnRowGroup, "center");
        var btnCancel = btnRowGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnOK = btnRowGroup.add("button", undefined, getLabel("button.ok"));

        return {
            dialog: duplicateDialog,
            countHorizontalInput: countFields.horizontalInput,
            countVerticalInput: countFields.verticalInput,
            countLinkCheck: countFields.linkCheck,
            countSlider: countSlider,
            methodGridRadio: methodGridRadio,
            methodRowRadio: methodRowRadio,
            methodColumnRadio: methodColumnRadio,
            methodRandomRadio: methodRandomRadio,
            gapHorizontalInput: gapFields.horizontalInput,
            gapVerticalInput: gapFields.verticalInput,
            gapLinkCheck: gapFields.linkCheck,
            directionRightRadio: horizontalDirectionRadios[0],
            directionLeftRadio: horizontalDirectionRadios[1],
            directionUpRadio: verticalDirectionRadios[0],
            directionDownRadio: verticalDirectionRadios[1],
            fillToEdgeCheck: fillToEdgeCheck,
            fillFullCheck: fillFullCheck,
            zoomControls: zoomControls,
            btnCancel: btnCancel,
            btnOK: btnOK
        };
    }

    /**
     * 設定ダイアログを表示し、プレビューを結線する
     * @param {Document} doc - 対象ドキュメント
     * @param {Array<PageItem>} sourceItems - 複製元オブジェクト（複数選択のまま扱う）
     * @param {number} sourceWidth - 複製元全体の幅（pt）
     * @param {number} sourceHeight - 複製元全体の高さ（pt）
     * @returns {object} OKなら {placementOffsets, fillFullArtboard}、キャンセルなら null
     */
    function showDuplicateDialog(doc, sourceItems, sourceWidth, sourceHeight) {
        var rulerUnit = getUnitInfo("rulerType");
        var dialogControls = buildDuplicateDialog(doc, rulerUnit.label);

        var duplicateDialog = dialogControls.dialog;
        var countHorizontalInput = dialogControls.countHorizontalInput;
        var countVerticalInput = dialogControls.countVerticalInput;
        var countLinkCheck = dialogControls.countLinkCheck;
        var countSlider = dialogControls.countSlider;
        var methodGridRadio = dialogControls.methodGridRadio;
        var methodRowRadio = dialogControls.methodRowRadio;
        var methodColumnRadio = dialogControls.methodColumnRadio;
        var methodRandomRadio = dialogControls.methodRandomRadio;
        var gapHorizontalInput = dialogControls.gapHorizontalInput;
        var gapVerticalInput = dialogControls.gapVerticalInput;
        var gapLinkCheck = dialogControls.gapLinkCheck;
        var directionRightRadio = dialogControls.directionRightRadio;
        var directionLeftRadio = dialogControls.directionLeftRadio;
        var directionUpRadio = dialogControls.directionUpRadio;
        var directionDownRadio = dialogControls.directionDownRadio;
        var fillToEdgeCheck = dialogControls.fillToEdgeCheck;
        var fillFullCheck = dialogControls.fillFullCheck;

        // -----------------------------------------
        // 入力値の読み取り / Reading the input values
        // -----------------------------------------

        /**
         * 選択中の繰り返し方式を返す
         * @returns {string} "grid" / "row" / "column" / "random"
         */
        function getRepeatMethod() {
            if (methodRowRadio.value) return "row";
            if (methodColumnRadio.value) return "column";
            if (methodRandomRadio.value) return "random";
            return "grid";
        }

        /**
         * 繰り返し数を読み取る（方式に応じて不要側を1に固定）
         * @returns {object} {columns, rows}。数値として読めない場合は null
         */
        function readRepeatCounts() {
            var columnCount = parseInt(countHorizontalInput.text, 10);
            var rowCount = parseInt(countVerticalInput.text, 10);
            var repeatMethod = getRepeatMethod();
            if (repeatMethod === "row" || repeatMethod === "random") rowCount = 1;
            if (repeatMethod === "column") columnCount = 1;
            if (isNaN(columnCount) || columnCount < 1 || isNaN(rowCount) || rowCount < 1) return null;
            return { columns: columnCount, rows: rowCount };
        }

        /**
         * 間隔を読み取り、ポイントへ変換する
         * @returns {object} {x, y}（pt）。数値として読めない場合は null
         */
        function readGapPoints() {
            var horizontalGap = parseFloat(gapHorizontalInput.text);
            var verticalGap = parseFloat(gapVerticalInput.text);
            if (isNaN(horizontalGap) || isNaN(verticalGap)) return null;
            return {
                x: horizontalGap * rulerUnit.pointsPerUnit,
                y: verticalGap * rulerUnit.pointsPerUnit
            };
        }

        // -----------------------------------------
        // ランダム配置 / Random placement
        // -----------------------------------------

        /* OK後もプレビューと同じ配置にするためのキャッシュ / Cache so OK keeps the previewed layout */
        var randomOffsetCache = { key: null, offsets: [] };

        /**
         * ランダム配置のオフセットを取得する（同じ条件ならキャッシュを返す）
         * @param {number} repeatCount - 繰り返し数
         * @param {number} gapX - 左右の間隔（pt）
         * @param {number} gapY - 上下の間隔（pt）
         * @returns {Array<Array<number>>} [dx, dy] の配列（pt）
         */
        function getRandomOffsets(repeatCount, gapX, gapY) {
            var GAP_EPSILON = 1e-9;
            var duplicateCount = Math.max(0, repeatCount - 1);

            /* 間隔が0の軸はランダムを完全にOFF / A gap of 0 disables randomness on that axis */
            var rangeX = (Math.abs(gapX) <= GAP_EPSILON) ? 0 : Math.max(0, (sourceWidth + gapX) * (repeatCount - 1) / 2);
            var rangeY = (Math.abs(gapY) <= GAP_EPSILON) ? 0 : Math.max(0, (sourceHeight + gapY) * (repeatCount - 1) / 2);

            var cacheKey = [repeatCount, gapX.toFixed(4), gapY.toFixed(4), sourceWidth.toFixed(4), sourceHeight.toFixed(4)].join("|");
            if (randomOffsetCache.key === cacheKey && randomOffsetCache.offsets.length === duplicateCount) {
                return randomOffsetCache.offsets;
            }

            var randomSeed = (new Date().getTime() & 0xFFFFFFFF) ^ (repeatCount << 16);
            randomOffsetCache.key = cacheKey;
            randomOffsetCache.offsets = createRandomOffsets(duplicateCount, rangeX, rangeY, randomSeed);
            return randomOffsetCache.offsets;
        }

        /**
         * 現在の設定から配置オフセットを組み立てる
         * @returns {Array<Array<number>>} [dx, dy] の配列（pt）。入力値が読めない場合は null
         */
        function buildPlacementOffsets() {
            var repeatCounts = readRepeatCounts();
            var gapPoints = readGapPoints();
            if (!repeatCounts || !gapPoints) return null;

            if (getRepeatMethod() === "random") {
                return getRandomOffsets(repeatCounts.columns, gapPoints.x, gapPoints.y);
            }
            var horizontalDirection = directionRightRadio.value ? "right" : "left";
            var verticalDirection = directionUpRadio.value ? "up" : "down";
            return computeGridOffsets(
                repeatCounts.rows, repeatCounts.columns,
                sourceWidth, sourceHeight,
                gapPoints.x, gapPoints.y,
                horizontalDirection, verticalDirection
            );
        }

        // -----------------------------------------
        // プレビュー更新 / Preview updates
        // -----------------------------------------

        var lastPreviewTime = 0;

        /**
         * 現在の設定でプレビューを描き直す
         * @returns {void}
         */
        function applyPreview() {
            var placementOffsets = buildPlacementOffsets();
            if (!placementOffsets) return;
            renderPreview(doc, sourceItems, placementOffsets);
        }

        /**
         * プレビュー更新を間引く（スライダードラッグ中のチラつき軽減）
         * @param {boolean} force - true なら間引かずに更新する
         * @returns {void}
         */
        function applyPreviewThrottled(force) {
            var currentTime = new Date().getTime();
            if (!force && (currentTime - lastPreviewTime) < PREVIEW_THROTTLE_MS) return;
            lastPreviewTime = currentTime;
            applyPreview();
        }

        // -----------------------------------------
        // UIの同期 / Keeping the UI in sync
        // -----------------------------------------

        /**
         * ［連動］に合わせて縦の繰り返し数を横へそろえる
         * @returns {void}
         */
        function syncCountFields() {
            if (countLinkCheck.value) {
                countVerticalInput.enabled = false;
                countVerticalInput.text = countHorizontalInput.text;
            } else {
                countVerticalInput.enabled = true;
            }
            updateCountSliderFromFields();
        }

        /**
         * ［連動］に合わせて上下の間隔を左右へそろえる
         * @returns {void}
         */
        function syncGapFields() {
            if (gapLinkCheck.value) {
                gapVerticalInput.enabled = false;
                gapVerticalInput.text = gapHorizontalInput.text;
            } else {
                gapVerticalInput.enabled = true;
            }
        }

        /**
         * 方向ラジオの有効／無効をまとめて切り替える
         * @param {boolean} horizontalEnabled - 横方向を有効にするか
         * @param {boolean} verticalEnabled - 縦方向を有効にするか
         * @returns {void}
         */
        function setDirectionEnabled(horizontalEnabled, verticalEnabled) {
            directionRightRadio.enabled = horizontalEnabled;
            directionLeftRadio.enabled = horizontalEnabled;
            directionUpRadio.enabled = verticalEnabled;
            directionDownRadio.enabled = verticalEnabled;
        }

        /**
         * 繰り返し方式と敷き詰めの状態から、方向ラジオの有効／無効を決める
         * @returns {void}
         */
        function applyDirectionStateForMethod() {
            /* ［アートボードいっぱいに］のあいだは方向を固定 / Fill Full pins the direction */
            if (fillFullCheck.value) { setDirectionEnabled(false, false); return; }

            var repeatMethod = getRepeatMethod();
            if (repeatMethod === "row") setDirectionEnabled(true, false);
            else if (repeatMethod === "column") setDirectionEnabled(false, true);
            else if (repeatMethod === "random") setDirectionEnabled(false, false);
            else setDirectionEnabled(true, true);
        }

        /**
         * 方向を既定（右・下）に戻す
         * @returns {void}
         */
        function resetDirectionToDefault() {
            directionRightRadio.value = true;
            directionLeftRadio.value = false;
            directionDownRadio.value = true;
            directionUpRadio.value = false;
        }

        /**
         * 入力欄の値をスライダーへ反映し、スライダーの有効／無効を決める
         * @returns {void}
         */
        function updateCountSliderFromFields() {
            var repeatMethod = getRepeatMethod();
            var sourceField = (repeatMethod === "column") ? countVerticalInput : countHorizontalInput;
            countSlider.value = clampRepeatCount(sourceField.text);

            /* 敷き詰め中は行列数を自動計算するので操作させない（上限20で潰れるため）
               While a Fill option drives the counts, block the slider so it cannot clamp them to 20 */
            if (fillToEdgeCheck.value || fillFullCheck.value) {
                countSlider.enabled = false;
                return;
            }
            /* 行／列／ランダムでは常に有効、グリッドは［連動］ONのときだけ有効
               Enabled in Row / Column / Random, and in Grid only while Link is on */
            countSlider.enabled = (repeatMethod !== "grid") || (countLinkCheck.enabled && countLinkCheck.value);
        }

        /**
         * スライダーの値を繰り返し数の入力欄へ反映する
         * @returns {void}
         */
        function applyCountFieldsFromSlider() {
            var repeatCount = clampRepeatCount(countSlider.value);
            countSlider.value = repeatCount;

            var repeatMethod = getRepeatMethod();
            if (repeatMethod === "column") {
                /* 列：横は常に1で、スライダーは縦を操作 / Column: horizontal stays 1, the slider drives vertical */
                countHorizontalInput.text = "1";
                countVerticalInput.text = String(repeatCount);
                return;
            }
            countHorizontalInput.text = String(repeatCount);
            if (repeatMethod === "row" || repeatMethod === "random") {
                countVerticalInput.text = "1";
            } else if (countLinkCheck.value) {
                countVerticalInput.text = String(repeatCount);
            }
        }

        /**
         * 繰り返し数の片方を1に固定し、もう片方だけを入力できるようにする（［連動］は OFF で無効）
         * @param {EditText} fixedInput - 1に固定する入力欄
         * @param {EditText} activeInput - 入力できるようにする入力欄
         * @returns {void}
         */
        function fixCountFieldToOne(fixedInput, activeInput) {
            fixedInput.text = "1";
            fixedInput.enabled = false;
            activeInput.enabled = true;
            countLinkCheck.value = false;
            countLinkCheck.enabled = false;
        }

        /**
         * 間隔の［連動］の値と有効／無効をまとめて設定する
         * @param {boolean} isLinked - ON かつ有効にするなら true、OFF かつ無効にするなら false
         * @returns {void}
         */
        function setGapLinkState(isLinked) {
            gapLinkCheck.value = isLinked;
            gapLinkCheck.enabled = isLinked;
        }

        /**
         * 繰り返し方式に合わせて各コントロールの状態を更新する
         * @returns {void}
         */
        function updateRepeatMethodUI() {
            var repeatMethod = getRepeatMethod();

            if (repeatMethod === "column") {
                /* 列：横は常に1 / Column: horizontal fixed to 1 */
                fixCountFieldToOne(countHorizontalInput, countVerticalInput);
                setGapLinkState(false);
                gapHorizontalInput.enabled = false;
                gapVerticalInput.enabled = true;

            } else if (repeatMethod === "row") {
                /* 行：縦は常に1 / Row: vertical fixed to 1 */
                fixCountFieldToOne(countVerticalInput, countHorizontalInput);
                setGapLinkState(false);
                gapHorizontalInput.enabled = true;
                gapVerticalInput.enabled = false;

            } else if (repeatMethod === "random") {
                /* ランダム：繰り返し数は横だけ、間隔は連動ON、方向と敷き詰めは無効
                   Random: a single count, gaps linked, direction & fill turned off */
                fixCountFieldToOne(countVerticalInput, countHorizontalInput);
                setGapLinkState(true);
                gapHorizontalInput.enabled = true;
                syncGapFields();
                fillToEdgeCheck.value = false;
                fillFullCheck.value = false;

            } else {
                /* グリッド：繰り返し数・間隔とも［連動］ONに戻す / Grid: restore both link checkboxes */
                countHorizontalInput.enabled = true;
                countLinkCheck.enabled = true;
                countLinkCheck.value = true;
                syncCountFields();

                setGapLinkState(true);
                gapHorizontalInput.enabled = true;
                syncGapFields();
            }

            applyDirectionStateForMethod();
            updateCountSliderFromFields();
        }

        // -----------------------------------------
        // 敷き詰めの自動計算 / Automatic fill counts
        // -----------------------------------------

        /**
         * 求めた行列数を入力欄とスライダーへ反映する
         * @param {{columns: number, rows: number}} fillCounts - 横と縦の数
         * @returns {void}
         */
        function setCountFields(fillCounts) {
            countHorizontalInput.text = String(fillCounts.columns);
            countVerticalInput.text = String(fillCounts.rows);
            updateCountSliderFromFields();
        }

        /**
         * ［アートボードの端まで］：選択オブジェクトを起点に行列数を求めて反映する
         * @returns {void}
         */
        function recalcCountsToArtboardEdge() {
            var gapPoints = readGapPoints();
            if (!gapPoints) return;
            setCountFields(computeCountsToArtboardEdge(
                getActiveArtboardRect(doc), getUnionBounds(sourceItems),
                sourceWidth, sourceHeight, gapPoints,
                getRepeatMethod(), directionRightRadio.value, directionUpRadio.value
            ));
        }

        /**
         * ［アートボードいっぱいに］：アートボードに収まる最大の行列数を求めて反映する（方向は無視）
         * @returns {void}
         */
        function recalcCountsToFullArtboard() {
            var gapPoints = readGapPoints();
            if (!gapPoints) return;
            setCountFields(computeCountsToFullArtboard(
                getActiveArtboardRect(doc), sourceWidth, sourceHeight, gapPoints, getRepeatMethod()
            ));
        }

        /**
         * ONになっている敷き詰めオプションに応じて行列数を再計算する
         * @returns {void}
         */
        function recalcFillCounts() {
            if (fillToEdgeCheck.value) recalcCountsToArtboardEdge();
            else if (fillFullCheck.value) recalcCountsToFullArtboard();
        }

        // -----------------------------------------
        // イベント結線 / Event wiring
        // -----------------------------------------

        /**
         * 横の繰り返し数が変わったときの処理
         * @returns {void}
         */
        function onCountHorizontalChanged() {
            if (countLinkCheck.value) countVerticalInput.text = countHorizontalInput.text;
            updateCountSliderFromFields();
            applyPreview();
        }

        /**
         * 縦の繰り返し数が変わったときの処理
         * @returns {void}
         */
        function onCountVerticalChanged() {
            updateCountSliderFromFields();
            applyPreview();
        }

        /**
         * 左右の間隔が変わったときの処理
         * @returns {void}
         */
        function onGapHorizontalChanged() {
            if (gapLinkCheck.value) gapVerticalInput.text = gapHorizontalInput.text;
            recalcFillCounts();
            applyPreview();
        }

        /**
         * 上下の間隔が変わったときの処理
         * @returns {void}
         */
        function onGapVerticalChanged() {
            recalcFillCounts();
            applyPreview();
        }

        /**
         * 繰り返し数のスライダーが動いたときの処理
         * @param {boolean} isReleased - 離したとき（間引かずにプレビューする）なら true
         * @returns {void}
         */
        function onCountSliderMoved(isReleased) {
            if (!countSlider.enabled) return;
            applyCountFieldsFromSlider();
            applyPreviewThrottled(isReleased);
        }

        /**
         * 左／上の方向が選ばれたときの処理（［アートボードの端まで］は右・下が起点なので自動で OFF）
         * @returns {void}
         */
        function onReverseDirectionChosen() {
            fillToEdgeCheck.value = false;
            updateCountSliderFromFields();
            applyPreview();
        }

        /**
         * 繰り返し方式のラジオが選ばれたときの処理
         * @returns {void}
         */
        function onRepeatMethodChanged() {
            updateRepeatMethodUI();
            recalcFillCounts();
            applyPreview();
        }

        countHorizontalInput.onChanging = onCountHorizontalChanged;
        countHorizontalInput.onChange = onCountHorizontalChanged;
        countVerticalInput.onChanging = onCountVerticalChanged;
        countVerticalInput.onChange = onCountVerticalChanged;
        gapHorizontalInput.onChanging = onGapHorizontalChanged;
        gapHorizontalInput.onChange = onGapHorizontalChanged;
        gapVerticalInput.onChanging = onGapVerticalChanged;
        gapVerticalInput.onChange = onGapVerticalChanged;

        countLinkCheck.onClick = function () {
            syncCountFields();
            applyPreview();
        };
        gapLinkCheck.onClick = function () {
            syncGapFields();
            /* 上下の間隔が左右にそろうと敷き詰めの行列数も変わる / Linking the gaps changes the fill counts too */
            recalcFillCounts();
            applyPreview();
        };

        countSlider.onChanging = function () {
            onCountSliderMoved(false);
        };
        countSlider.onChange = function () {
            onCountSliderMoved(true);
        };

        methodGridRadio.onClick = onRepeatMethodChanged;
        methodRowRadio.onClick = function () {
            /* 列→行では縦の数を横の数として引き継ぐ / Carry the vertical count over when switching Column -> Row */
            countHorizontalInput.text = String(clampRepeatCount(countVerticalInput.text));
            onRepeatMethodChanged();
        };
        methodColumnRadio.onClick = function () {
            /* 行→列では横の数を縦の数として引き継ぐ / Carry the horizontal count over when switching Row -> Column */
            countVerticalInput.text = String(clampRepeatCount(countHorizontalInput.text));
            onRepeatMethodChanged();
        };
        methodRandomRadio.onClick = function () {
            /* 列→ランダムでは縦の数を横の数として引き継ぐ / Carry the vertical count over when switching Column -> Random */
            if (getRepeatMethod() === "random" && countHorizontalInput.text === "1") {
                countHorizontalInput.text = String(clampRepeatCount(countVerticalInput.text));
            }
            onRepeatMethodChanged();
        };

        directionRightRadio.onClick = applyPreview;
        directionDownRadio.onClick = applyPreview;
        /* 左・上を選んだら［アートボードの端まで］を自動OFF / Choosing Left or Up turns Fill to Edge off */
        directionLeftRadio.onClick = onReverseDirectionChosen;
        directionUpRadio.onClick = onReverseDirectionChosen;

        fillToEdgeCheck.onClick = function () {
            if (fillToEdgeCheck.value) {
                /* 端までは右・下を起点にする（方向は変更可）/ Fill to Edge starts from the right & bottom, direction still editable */
                fillFullCheck.value = false;
                resetDirectionToDefault();
                applyDirectionStateForMethod();
                countLinkCheck.value = false;
                syncCountFields();
                recalcCountsToArtboardEdge();
            }
            updateCountSliderFromFields();
            applyPreview();
        };

        fillFullCheck.onClick = function () {
            if (fillFullCheck.value) {
                /* いっぱいには方向を固定してディム / Fill Full fixes the direction and dims the radios */
                fillToEdgeCheck.value = false;
                resetDirectionToDefault();
                recalcCountsToFullArtboard();
            }
            applyDirectionStateForMethod();
            updateCountSliderFromFields();
            applyPreview();
        };

        /* キー操作でラジオを選択 / Keyboard shortcuts for radios */
        duplicateDialog.addEventListener("keydown", function (event) {
            var keyName = event.keyName;
            if (!keyName) return;

            var targetRadio = null;
            if (keyName === "G") targetRadio = methodGridRadio;
            else if (keyName === "C") targetRadio = methodColumnRadio;
            else if (keyName === "A") targetRadio = methodRandomRadio;
            else if (keyName === "R") targetRadio = event.shiftKey ? directionRightRadio : methodRowRadio;
            else if (keyName === "L") targetRadio = directionLeftRadio;
            else if (keyName === "T") targetRadio = directionUpRadio;
            else if (keyName === "B") targetRadio = directionDownRadio;
            if (!targetRadio) return;

            selectRadioButton(targetRadio);
            event.preventDefault();
        });

        var dialogResult = null;

        dialogControls.btnOK.onClick = function () {
            if (!readRepeatCounts()) { alert(getLabel("alert.invalidCount")); return; }
            if (!readGapPoints()) { alert(getLabel("alert.invalidGap")); return; }

            dialogResult = {
                placementOffsets: buildPlacementOffsets(),
                fillFullArtboard: fillFullCheck.value
            };
            clearPreview(doc);
            duplicateDialog.close();
        };

        dialogControls.btnCancel.onClick = function () {
            dialogControls.zoomControls.restoreInitial();
            clearPreview(doc);
            duplicateDialog.close();
        };

        duplicateDialog.onClose = function () {
            /* OK以外で閉じたときはキャンセル扱い / Closing without OK is treated as Cancel */
            if (dialogResult) return;
            dialogControls.zoomControls.restoreInitial();
            clearPreview(doc);
        };

        duplicateDialog.onShow = function () {
            var dialogLocation = duplicateDialog.location;
            duplicateDialog.location = [dialogLocation[0] + DIALOG_OFFSET_X, dialogLocation[1]];
            applyPreview();
            countHorizontalInput.active = true;
        };

        updateRepeatMethodUI();
        duplicateDialog.show();
        return dialogResult;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 複数選択をひとつのグループにまとめる
     * @param {Document} doc - 対象ドキュメント
     * @param {Array<PageItem>} selectedItems - 選択中のオブジェクト
     * @returns {PageItem} まとめたグループ（失敗時は先頭のオブジェクト）
     */
    function groupSelectedItems(doc, selectedItems) {
        try {
            var wrapperGroup = doc.groupItems.add();
            /* 選択配列はライブで変化するのでコピーを回す / The live selection array can change, so iterate a copy */
            var itemsToMove = [];
            for (var i = 0; i < selectedItems.length; i++) itemsToMove.push(selectedItems[i]);
            /* 移動できないオブジェクトが混じっても残りは処理する / Keep going even if one item refuses to move */
            for (var j = 0; j < itemsToMove.length; j++) {
                try { itemsToMove[j].move(wrapperGroup, ElementPlacement.PLACEATEND); } catch (err) { }
            }

            doc.selection = null;
            wrapperGroup.selected = true;
            return wrapperGroup;
        } catch (e) {
            return selectedItems[0];
        }
    }

    /**
     * 元オブジェクトと複製をまとめてアートボード中央へ移動する
     * @param {Document} doc - 対象ドキュメント
     * @param {Array<PageItem>} sourceItems - 複製元
     * @param {Array<PageItem>} duplicatedItems - 生成した複製
     * @returns {void}
     */
    function centerOnArtboard(doc, sourceItems, duplicatedItems) {
        var artboardRect = getActiveArtboardRect(doc);
        var itemsToMove = sourceItems.concat(duplicatedItems);
        var unionBounds = getUnionBounds(itemsToMove);

        var dx = (artboardRect[0] + artboardRect[2]) / 2 - (unionBounds[0] + unionBounds[2]) / 2;
        var dy = (artboardRect[1] + artboardRect[3]) / 2 - (unionBounds[1] + unionBounds[3]) / 2;

        for (var i = 0; i < itemsToMove.length; i++) {
            itemsToMove[i].left += dx;
            itemsToMove[i].top += dy;
        }
    }

    /**
     * 検証 → ダイアログ → 複製
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) { alert(getLabel("alert.noDocument")); return; }
        var doc = app.activeDocument;
        if (doc.selection.length === 0) { alert(getLabel("alert.noSelection")); return; }

        /* 選択配列はライブで変化するのでコピーを保持 / The live selection array can change, so keep a copy */
        var selectedItems = [];
        for (var i = 0; i < doc.selection.length; i++) selectedItems.push(doc.selection[i]);

        var sourceBounds = getUnionBounds(selectedItems);
        var sourceWidth = sourceBounds[2] - sourceBounds[0];
        var sourceHeight = sourceBounds[1] - sourceBounds[3];

        /* グループ化はOK後まで遅らせる（キャンセル時にドキュメントを変更しないため）
           Grouping is deferred until after OK so that cancelling leaves the document untouched */
        var duplicateSettings = showDuplicateDialog(doc, selectedItems, sourceWidth, sourceHeight);
        if (!duplicateSettings) return;

        var sourceItems = (selectedItems.length > 1) ? [groupSelectedItems(doc, selectedItems)] : selectedItems;
        var duplicatedItems = duplicateWithOffsets(sourceItems, duplicateSettings.placementOffsets, null);
        if (duplicateSettings.fillFullArtboard) centerOnArtboard(doc, sourceItems, duplicatedItems);
    }

    main();

})();
