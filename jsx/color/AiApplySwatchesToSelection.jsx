#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したオブジェクトやテキストに、スウォッチや定義済みカラーを適用するモーダルダイアログです。
配色は選択中のスウォッチかスウォッチグループから取り込め、適用単位（オブジェクト／1文字／単語／行／段落）と適用順（そのまま／逆順／ランダム／完全ランダム）を変えるたびにライブプレビューします。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AiApplySwatchesToSelection.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n5602f3084d2b

### Overview

A modal dialog that applies swatches, or predefined colors, to the selected objects and text.
Colors are captured from the selected swatches or from a swatch group, and every change to the application unit (object, character, word, line or paragraph) or order (as-is, reversed, random or fully random) is previewed live.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AiApplySwatchesToSelection.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AiApplySwatchesToSelection";   /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.8.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2024-11-03";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-23";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AiApplySwatchesToSelection.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AiApplySwatchesToSelection.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n5602f3084d2b"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* 自動カラー（CMYK ドキュメント）の生成設定 / Auto colors for CMYK documents */
    var CMYK_FALLBACK_MAX_TOTAL = 200;        /* 合計上限 C+M / C+Y / M+Y / total channel limit */
    var CMYK_FALLBACK_MIN_DISTANCE = 35;      /* 近似色回避の最小距離 / min distance to avoid similar colors */

    /* 1文字単位プレビューで着色する最大文字数（超過分は確定時に着色）/ max chars colored in per-character preview */
    var PREVIEW_CHAR_CAP = 500;

    /* 初回プレビューでプログレスバーを出す目安（文字数・オブジェクト数）/ Show a progress bar above these counts */
    var HEAVY_SELECTION_CHAR_COUNT = 1000;
    var HEAVY_SELECTION_ITEM_COUNT = 300;

    // =========================================
    // レイアウト / Layout
    // =========================================

    var WINDOW_MARGINS = 16;                 /* ウィンドウ外周の余白 / window margin */
    var WINDOW_SPACING = 12;                 /* ウィンドウ内の要素間隔 / window spacing */
    var PANEL_MARGINS  = [16, 20, 16, 12];   /* パネル余白 [左,上,右,下] / panel margins */
    var PANEL_SPACING  = 8;                  /* パネル内の要素間隔 / panel spacing */
    var COLUMN_SPACING = 12;                 /* 2カラムの間隔 / gap between columns */

    var COLOR_CHIPS_PER_ROW = 12;            /* 1行あたりのチップ数 / chips per row */
    var CHIP_SIZE = 18;                      /* チップの一辺(px) / chip size */
    var CHIP_GAP = 4;                        /* チップ間の間隔(px) / gap between chips */

    var PROGRESS_WIDTH = 320;                /* プログレスバーの幅 / progress bar width */

    /**
     * ウィンドウの共通設定
     * @param {Window} targetWindow - 対象ウィンドウ
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
     * パネルの共通設定
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
     * 行グループの共通設定（ボタン列など）
     * @param {Group} rowGroup - 対象グループ
     * @param {string} [alignment] - 配置（省略時は "left"）
     * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function setupRow(rowGroup, alignment, spacing) {
        rowGroup.orientation = "row";
        rowGroup.alignment = alignment || "left";
        rowGroup.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * 現在の UI 言語を判定する
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = getCurrentLang();

    /* ラベル定義（カテゴリ分け）/ Label definitions (categorized)
       radio.unit / radio.order と tooltip.unit / tooltip.order はキーで引く（addRadioPanel）
       radio.unit / radio.order and tooltip.unit / tooltip.order are looked up by key in addRadioPanel */
    var LABELS = {
        dialog: {
            title: { ja: "カラーを配色", en: "Distribute Colors" }
        },
        panel: {
            unit: { ja: "配色単位", en: "Distribute By" },
            colors: { ja: "使用するカラー", en: "Colors to Use" },
            order: { ja: "配色順", en: "Order" }
        },
        radio: {
            unit: {
                object:    { ja: "オブジェクト", en: "Per object" },
                character: { ja: "1文字", en: "Per character" },
                word:      { ja: "単語", en: "Per word" },
                line:      { ja: "行", en: "Per line" },
                paragraph: { ja: "段落", en: "Per paragraph" }
            },
            order: {
                asis:       { ja: "取り込み順", en: "As captured" },
                reverse:    { ja: "逆順", en: "Reverse" },
                random:     { ja: "ランダム", en: "Random" },
                fullrandom: { ja: "完全ランダム", en: "Fully random" }
            },
            source: {
                selected: { ja: "選択しているスウォッチ", en: "Selected swatches" },
                group:    { ja: "スウォッチグループ", en: "Swatch group" }
            }
        },
        dropdown: {
            noGroup: { ja: "（グループなし）", en: "(no groups)" }
        },
        button: {
            ok:     { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        tooltip: {
            colors: {
                ja: "開いた時点で選択していたスウォッチを適用に使います。",
                en: "The swatches selected when the dialog opened are used for applying."
            },
            source: {
                selected: {
                    ja: "開いた時点で選択していたスウォッチを使います。選択していない（または白だけの）ときは自動カラーを使います。",
                    en: "Uses the swatches selected when the dialog opened. With none (or only white) selected, auto colors are used."
                },
                group: {
                    ja: "下で選んだスウォッチグループのカラーを使います（［なし］は除きます）。",
                    en: "Uses the colors of the swatch group chosen below ([None] is skipped)."
                },
                groupDropdown: {
                    ja: "使うスウォッチグループを選びます。選ぶと［スウォッチグループ］に切り替わります。",
                    en: "Chooses the swatch group to use. Choosing one switches to Swatch group."
                }
            },
            unit: {
                object: {
                    ja: "選択オブジェクト単位でカラーを順番に適用します。",
                    en: "Applies colors in order, one per selected object."
                },
                character: { ja: "テキストを1文字ずつ分けてカラーを適用します。", en: "Applies a color to each character of the text." },
                word: {
                    ja: "英文テキストを単語ごとに分けてカラーを適用します。英文以外では単語が正しく分割されないことがあります。",
                    en: "Applies a color to each word of English text. Words may not split correctly for non-English text."
                },
                line:      { ja: "テキストの行ごとにカラーを適用します。", en: "Applies a color to each line of the text." },
                paragraph: { ja: "テキストの段落ごとにカラーを適用します。", en: "Applies a color to each paragraph of the text." }
            },
            order: {
                asis: {
                    ja: "取り込んだカラーの並び順で適用します（適用先は位置順・文字順）。",
                    en: "Applies colors in the captured order (targets follow position / reading order)."
                },
                reverse: { ja: "取り込んだカラーの並びを逆にして適用します。", en: "Applies the captured colors in reverse order." },
                random: {
                    ja: "取り込んだカラーの並びをランダムにして適用します（並びは繰り返します）。",
                    en: "Applies the captured colors in a random order (the sequence repeats)."
                },
                fullrandom: {
                    ja: "適用先ごとにカラーを毎回ランダムに選びます（並びは繰り返しません）。",
                    en: "Picks a color at random for each target (no repeating sequence)."
                }
            }
        },
        alert: {
            noDoc: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noSel: { ja: "オブジェクトを選択してください。", en: "Please select objects." }
        },
        note: {
            autoColor: { ja: "カラー未選択：自動カラーを使用します", en: "No colors selected: using auto colors" }
        },
        progress: {
            title:    { ja: "準備しています…", en: "Preparing…" },
            analyze:  { ja: "対象を解析しています…", en: "Analyzing selection…" },
            snapshot: { ja: "元の状態を保存しています…", en: "Saving original state…" },
            apply:    { ja: "カラーを適用しています…", en: "Applying colors…" }
        }
    };

    /* 配色単位・配色順の選択肢（LABELS の radio.unit / radio.order のキー）/ Option keys */
    var UNIT_KEYS = ["object", "character", "word", "line", "paragraph"];
    var ORDER_KEYS = ["asis", "reverse", "random", "fullrandom"];

    /**
     * ドット区切りキーでローカライズ文字列を取得する
     * @param {string} labelPath - "panel.unit" のようなドット区切りのキー
     * @returns {string} 表示言語のテキスト（見つからない場合は labelPath をそのまま返す）
     */
    function getLabel(labelPath) {
        var pathKeys = labelPath.split(".");
        var labelNode = LABELS;
        for (var i = 0; i < pathKeys.length; i++) {
            labelNode = labelNode[pathKeys[i]];
            if (labelNode == null) return labelPath;
        }
        return labelNode[uiLang];
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
    // プログレスバー / Progress bar
    // =========================================

    /**
     * 進捗表示用のパレットウィンドウを生成して表示する（非モーダルなので処理中に更新できる）
     *
     * 戻り値の set(fraction 0..1, message) で進捗を更新、close() で閉じる。
     * @param {string} title - ウィンドウのタイトル
     * @returns {{set: function, close: function}} 進捗の更新と終了の操作
     */
    function createProgressWindow(title) {
        var progressWindow = new Window("palette", title);
        progressWindow.orientation = "column";
        progressWindow.alignChildren = ["fill", "top"];
        progressWindow.margins = 16;
        progressWindow.spacing = 8;
        var progressBar = progressWindow.add("progressbar", undefined, 0, 100);
        progressBar.preferredSize = [PROGRESS_WIDTH, 8];
        var progressLabel = progressWindow.add("statictext", undefined, "");
        progressLabel.preferredSize.width = PROGRESS_WIDTH;
        progressWindow.show();
        progressWindow.update();
        return {
            set: function (fraction, message) {
                var percent = Math.round(fraction * 100);
                progressBar.value = (percent < 0) ? 0 : (percent > 100 ? 100 : percent);
                if (message != null) { progressLabel.text = message; }
                progressWindow.update(); /* 同期ループ中でも再描画させる / force a repaint even inside a synchronous loop */
            },
            close: function () { progressWindow.close(); }
        };
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * モーダルダイアログを表示し、OK でプレビューを確定、それ以外は元へ戻す
     *
     * メインエンジンで動くので DOM 操作は直接呼ぶ。ライブプレビューは applyColorsToSelection が
     * ベースラインを基準に上書きする。OK でプレビューを確定、キャンセル／Esc／閉じるで元へ戻す。
     * @returns {void}
     */
    function showDialog() {
        if (app.documents.length === 0) { alert(getLabel("alert.noDoc")); return; }

        /* 選択情報を取得（初期単位・チップ・スウォッチ名）/ Read selection info (default unit, chips, swatch names) */
        var selectionInfo = readSelectionInfo();
        if (!selectionInfo.ok || (!selectionInfo.isText && selectionInfo.itemCount === 0)) { alert(getLabel("alert.noSel")); return; }

        var colorDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        setupWindow(colorDialog);

        /* ドキュメントのスウォッチグループ（未分類は除外）/ Document swatch groups (uncategorized excluded) */
        var swatchGroups = readSwatchGroupsInfo();

        /* 適用に使うスウォッチ名（スウォッチ名で参照）。カラーソースに応じて差し替える
           Swatch names used for applying (referenced by name); swapped by the chosen color source */
        var loadedSwatchNames = selectionInfo.swatchNames;

        var colorSource = buildColorPanel(colorDialog, selectionInfo.chips, swatchGroups);

        /**
         * カラーソースを適用してチップとプレビューを更新する
         * @returns {void}
         */
        function applyColorSource() {
            var chips;
            if (colorSource.groupRadio.value && swatchGroups.length > 0) {
                var swatchGroup = swatchGroups[colorSource.groupDropdown.selection.index];
                loadedSwatchNames = swatchGroup.swatchNames;
                chips = swatchGroup.chips;
            } else {
                loadedSwatchNames = selectionInfo.swatchNames;
                chips = selectionInfo.chips;
            }
            /* チップを描き直してレイアウトを取り直す / Rebuild chips and relayout */
            clearChildren(colorSource.chipContainer);
            buildColorChips(colorSource.chipContainer, chips);
            colorDialog.layout.layout(true);
            runPreview();
        }

        /**
         * カラーソースのラジオを手動で排他的に切り替える（別のグループに入っているため）
         * @param {boolean} useGroup - スウォッチグループを使うなら true
         * @returns {void}
         */
        function selectColorSource(useGroup) {
            colorSource.groupRadio.value = useGroup;
            colorSource.selectedRadio.value = !useGroup;
        }

        colorSource.selectedRadio.onClick = function () {
            selectColorSource(false);
            applyColorSource();
        };
        colorSource.groupRadio.onClick = function () {
            selectColorSource(true);
            applyColorSource();
        };
        colorSource.groupDropdown.onChange = function () {
            if (!colorSource.groupRadio.value) { selectColorSource(true); }
            applyColorSource();
        };

        /* 配色単位・配色順を2カラムで配置 / Coloring unit and order in two columns */
        var unitsRow = colorDialog.add("group");
        setupRow(unitsRow, "fill", COLUMN_SPACING);
        unitsRow.alignChildren = ["fill", "top"];

        var unitRadioButtons = addRadioPanel(unitsRow, "unit", UNIT_KEYS, selectionInfo.defaultUnit, runPreview);
        updateUnitAvailability(unitRadioButtons, selectionInfo);
        var orderRadioButtons = addRadioPanel(unitsRow, "order", ORDER_KEYS, "asis", runPreview);

        var dialogButtons = buildButtonRow(colorDialog);
        dialogButtons.btnOK.onClick = function () { colorDialog.close(1); };
        dialogButtons.btnCancel.onClick = function () { colorDialog.close(2); };

        /**
         * 現在の配色単位と配色順を取得する
         * @returns {{unit: string, order: string}} 配色の設定
         */
        function readColoringOptions() {
            return {
                unit: getSelectedOption(unitRadioButtons, "object"),
                order: getSelectedOption(orderRadioButtons, "asis")
            };
        }

        /**
         * 現在の設定でプレビューを適用する（メインエンジンで直接実行）
         * @returns {void}
         */
        function runPreview() {
            applyColorsToSelection(readColoringOptions(), loadedSwatchNames);
        }

        /* 初回プレビューはダイアログ表示前に済ませる。元状態の保存（全文字読み取り）が重いので、
           重い選択のときはプログレスバーを出して進捗を見せ、終わってから塗り済みのダイアログを開く
           Do the first preview before showing the dialog. Saving originals (reading every character) is
           heavy, so show a progress bar for heavy selections, then open the already-painted dialog */
        commitPreview(); /* 旧ベースラインを破棄（現在の状態を基準にする）/ discard any stale baseline */
        runFirstPreview(readColoringOptions(), loadedSwatchNames);

        /* モーダル実行。OK 以外（キャンセル・Esc・クローズボックス）は元へ戻す
           Run modally; anything other than OK (Cancel, Esc, close box) reverts the preview */
        if (colorDialog.show() === 1) {
            /* 間引きプレビューだったときだけ、保存した配色で全文字をフル着色（プレビューと同じ結果になる）。
               間引きでなければプレビューが最終結果なので再適用しない（再適用するとランダムが引き直されて変わる）
               Only when the preview was decimated, apply fully using the saved colors (matches the preview);
               otherwise the preview is already final, so skip re-apply (it would re-randomize) */
            if ($.global.__aiApplyWasDecimated) {
                applyColorsToSelection(readColoringOptions(), loadedSwatchNames, true);
            }
            commitPreview();
        } else {
            restoreBaseline();
            app.redraw();
        }
    }

    /**
     * 「使用するカラー」パネル（配色チップとカラーソースの選択）を組み立てる
     * @param {Window} colorDialog - 追加先のダイアログ
     * @param {number[][]} chips - 最初に表示するチップの RGB（0..255）
     * @param {Object[]} swatchGroups - readSwatchGroupsInfo() の結果
     * @returns {{chipContainer: Group, selectedRadio: RadioButton, groupRadio: RadioButton, groupDropdown: DropDownList}} 各コントロール
     */
    function buildColorPanel(colorDialog, chips, swatchGroups) {
        var colorPanel = colorDialog.add("panel", undefined, getLabel("panel.colors"));
        setupPanel(colorPanel);
        colorPanel.helpTip = getLabel("tooltip.colors");
        var chipContainer = colorPanel.add("group");
        chipContainer.orientation = "column";
        chipContainer.alignChildren = ["left", "top"];
        buildColorChips(chipContainer, chips);

        /* カラーソース：選択スウォッチ or スウォッチグループ（ラジオは手動で排他制御）。
           ラジオを group にまとめ、上マージンでチップとの間隔をとる
           Color source radios wrapped in a group; top margin adds space above them */
        var colorSourceGroup = colorPanel.add("group");
        colorSourceGroup.orientation = "column";
        colorSourceGroup.alignChildren = ["left", "top"];
        colorSourceGroup.margins = [0, 10, 0, 0];
        var selectedRadio = colorSourceGroup.add("radiobutton", undefined, getLabel("radio.source.selected"));
        selectedRadio.value = true;
        selectedRadio.helpTip = getLabel("tooltip.source.selected");
        var groupRadio = colorSourceGroup.add("radiobutton", undefined, labelText("radio.source.group"));
        groupRadio.helpTip = getLabel("tooltip.source.group");
        var groupNames = [];
        for (var i = 0; i < swatchGroups.length; i++) { groupNames.push(swatchGroups[i].name); }
        /* ポップアップは次の行に置き、グループラジオの下へ少しインデントする
           Put the dropdown on the next line, slightly indented under the group radio */
        var groupDropdownRow = colorSourceGroup.add("group");
        setupRow(groupDropdownRow, "left");
        groupDropdownRow.margins = [16, 0, 0, 0];
        var groupDropdown = groupDropdownRow.add("dropdownlist", undefined, groupNames.length > 0 ? groupNames : [getLabel("dropdown.noGroup")]);
        groupDropdown.selection = 0;
        groupDropdown.helpTip = getLabel("tooltip.source.groupDropdown");
        if (swatchGroups.length === 0) {
            /* グループが無ければグループ選択は無効 / disable the group option when there are none */
            groupRadio.enabled = false;
            groupDropdown.enabled = false;
        }
        return { chipContainer: chipContainer, selectedRadio: selectedRadio, groupRadio: groupRadio, groupDropdown: groupDropdown };
    }

    /**
     * ボタン行（キャンセル／OK）を組み立てる
     * @param {Window} colorDialog - 追加先のダイアログ
     * @returns {{btnOK: Button, btnCancel: Button}} OK・キャンセルボタン
     */
    function buildButtonRow(colorDialog) {
        var btnRowGroup = colorDialog.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.alignment = "right";
        btnRowGroup.spacing = PANEL_SPACING;
        var btnCancel = btnRowGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnOK = btnRowGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
        return { btnOK: btnOK, btnCancel: btnCancel };
    }

    /**
     * ラジオボタンのパネルを追加して配列を返す
     *
     * パネル名は panel.<category>、ラベルは radio.<category>.<key>、tooltip は tooltip.<category>.<key> から引く。
     * @param {Group} parentGroup - 追加先
     * @param {string} category - "unit" / "order"
     * @param {string[]} optionKeys - 選択肢のキー
     * @param {string} defaultKey - 最初に選ぶキー
     * @param {function} onChange - クリック時に呼ぶ関数
     * @returns {RadioButton[]} 追加したラジオボタン（optionKey にキーを持つ）
     */
    function addRadioPanel(parentGroup, category, optionKeys, defaultKey, onChange) {
        var radioPanel = parentGroup.add("panel", undefined, getLabel("panel." + category));
        setupPanel(radioPanel);
        var radioButtons = [];
        for (var i = 0; i < optionKeys.length; i++) {
            var radioButton = radioPanel.add("radiobutton", undefined, getLabel("radio." + category + "." + optionKeys[i]));
            radioButton.optionKey = optionKeys[i];
            if (optionKeys[i] === defaultKey) radioButton.value = true;
            radioButton.onClick = onChange;
            radioButtons.push(radioButton);
        }
        for (var j = 0; j < radioButtons.length; j++) {
            radioButtons[j].helpTip = getLabel("tooltip." + category + "." + radioButtons[j].optionKey);
        }
        return radioButtons;
    }

    /**
     * 選択中のラジオボタンのキーを取得する
     * @param {RadioButton[]} radioButtons - 対象のラジオボタン
     * @param {string} fallbackKey - どれも選ばれていないときのキー
     * @returns {string} 選択中のキー
     */
    function getSelectedOption(radioButtons, fallbackKey) {
        for (var i = 0; i < radioButtons.length; i++) {
            if (radioButtons[i].value) return radioButtons[i].optionKey;
        }
        return fallbackKey;
    }

    /**
     * 選択内容に応じて配色単位ラジオの有効／無効を更新する
     *
     * - オブジェクトが1個ならオブジェクトを無効
     * - テキストでない場合は文字/単語/行/段落を無効
     * - 段落が1つだけなら段落も無効
     * - すべて無効になる場合（単一図形など）はオブジェクトを残す
     * @param {RadioButton[]} unitRadioButtons - 配色単位のラジオボタン
     * @param {Object} selectionInfo - readSelectionInfo() の結果
     * @returns {void}
     */
    function updateUnitAvailability(unitRadioButtons, selectionInfo) {
        var enabledByUnit = {
            object: (selectionInfo.itemCount !== 1),
            character: selectionInfo.isText,
            word: selectionInfo.isText,
            line: selectionInfo.isText,
            /* 段落は、テキストが複数 or 単一テキストの段落が2つ以上のとき有効
               （＝単一テキストで段落1つのときだけディム）
               Paragraph enabled when multiple text objects, or a single text object with 2+ paragraphs */
            paragraph: (selectionInfo.isText && (selectionInfo.textObjectCount !== 1 || selectionInfo.paragraphCount > 1))
        };

        if (!(enabledByUnit.object || enabledByUnit.character || enabledByUnit.word || enabledByUnit.line || enabledByUnit.paragraph)) {
            enabledByUnit.object = true;
        }

        for (var i = 0; i < unitRadioButtons.length; i++) {
            unitRadioButtons[i].enabled = enabledByUnit[unitRadioButtons[i].optionKey];
        }

        /* 選択中が無効になったら有効な先頭の単位へ寄せる / Move selection to the first enabled unit if the current one is disabled */
        for (var j = 0; j < unitRadioButtons.length; j++) {
            if (unitRadioButtons[j].value && unitRadioButtons[j].enabled) return;
        }
        for (var m = 0; m < unitRadioButtons.length; m++) {
            if (unitRadioButtons[m].enabled) {
                for (var n = 0; n < unitRadioButtons.length; n++) {
                    unitRadioButtons[n].value = (n === m);
                }
                return;
            }
        }
    }

    /**
     * コンテナの子要素をすべて削除する
     * @param {Group} container - 対象のコンテナ
     * @returns {void}
     */
    function clearChildren(container) {
        while (container.children.length > 0) {
            container.remove(container.children[container.children.length - 1]);
        }
    }

    // =========================================
    // カラーチップ表示 / Color chips
    // =========================================

    /**
     * スウォッチのカラーを■で並べる（無いときは注記を表示）
     * @param {Group} container - 追加先
     * @param {number[][]} chips - チップの RGB（0..255）
     * @returns {void}
     */
    function buildColorChips(container, chips) {
        if (!chips || chips.length === 0) {
            container.add("statictext", undefined, getLabel("note.autoColor"));
            return;
        }
        /* 1行を1つの描画キャンバスにまとめて描く（要素ごとのレイアウトずれを避ける）
           Draw each row on a single canvas to avoid per-element vertical misalignment */
        for (var start = 0; start < chips.length; start += COLOR_CHIPS_PER_ROW) {
            addChipRowCanvas(container, chips.slice(start, start + COLOR_CHIPS_PER_ROW));
        }
    }

    /**
     * 1行ぶんのチップを1つの group に手描きする
     * @param {Group} container - 追加先
     * @param {number[][]} rowChips - この行のチップの RGB（0..255）
     * @returns {void}
     */
    function addChipRowCanvas(container, rowChips) {
        var rowWidth = rowChips.length * CHIP_SIZE + (rowChips.length - 1) * CHIP_GAP;
        var chipRowCanvas = container.add("group");
        chipRowCanvas.preferredSize = [rowWidth, CHIP_SIZE];
        chipRowCanvas.minimumSize = [rowWidth, CHIP_SIZE];
        chipRowCanvas.maximumSize = [rowWidth, CHIP_SIZE];
        chipRowCanvas.chipColors = rowChips;
        chipRowCanvas.onDraw = function () {
            var graphics = this.graphics;
            for (var i = 0; i < this.chipColors.length; i++) {
                var chipRgb = this.chipColors[i];
                var brush = graphics.newBrush(graphics.BrushType.SOLID_COLOR, [chipRgb[0] / 255, chipRgb[1] / 255, chipRgb[2] / 255, 1]);
                graphics.newPath();
                graphics.rectPath(i * (CHIP_SIZE + CHIP_GAP), 0, CHIP_SIZE, CHIP_SIZE);
                graphics.fillPath(brush);
            }
        };
    }

    // =========================================
    // 選択情報の取得 / Reading the selection
    // =========================================

    /**
     * 選択情報（初期の配色単位・スウォッチのチップと名前・テキストの統計）を取得する
     * @returns {Object} ok / defaultUnit / chips / swatchNames / itemCount / isText / paragraphCount / textObjectCount
     */
    function readSelectionInfo() {
        var selectionInfo = { ok: false, defaultUnit: "object", chips: [], swatchNames: [], itemCount: 0, isText: false, paragraphCount: 0, textObjectCount: 0 };
        if (app.documents.length === 0) return selectionInfo;
        var doc = app.activeDocument;
        var items = flattenSelection(app.selection);
        var textRange = getSelectedTextRange(app.selection);
        if (textRange) { selectionInfo.defaultUnit = "character"; }
        else if (items.length === 1 && items[0].typename === "TextFrame") { selectionInfo.defaultUnit = "character"; }

        /* 選択スウォッチ（白だけのときは使わない）/ Selected swatches, ignored when all are white */
        var swatches = doc.swatches.getSelected();
        if (swatches && swatches.length >= 1 && !allWhiteSwatches(swatches)) {
            for (var i = 0; i < swatches.length; i++) {
                selectionInfo.chips.push(colorToRGB255(swatches[i].color));
                selectionInfo.swatchNames.push(swatches[i].name);
            }
        }

        var textStats = getSelectionTextStats(items, textRange);
        selectionInfo.itemCount = items.length;
        selectionInfo.isText = textStats.isText;
        selectionInfo.paragraphCount = textStats.paragraphCount;
        selectionInfo.textObjectCount = textStats.textObjectCount;
        selectionInfo.ok = true;
        return selectionInfo;
    }

    /**
     * ドキュメントのスウォッチグループを取得する（未分類＝名前なしと、[なし] のスウォッチは除外）
     * @returns {Object[]} 各グループの { name, swatchNames, chips }
     */
    function readSwatchGroupsInfo() {
        var swatchGroupsInfo = [];
        if (app.documents.length === 0) return swatchGroupsInfo;
        var doc = app.activeDocument;
        for (var i = 0; i < doc.swatchGroups.length; i++) {
            var swatchGroup = doc.swatchGroups[i];
            var groupName, groupSwatches;
            /* 未分類グループの name は ""、getAllSwatches はまれに投げる。1つの try でまとめてスキップ判定
               The uncategorized group's name is ""; getAllSwatches occasionally throws — one try covers both */
            try {
                groupName = swatchGroup.name;
                if (groupName === "") { continue; } /* 未分類グループは除外 / skip the uncategorized group */
                groupSwatches = swatchGroup.getAllSwatches();
            } catch (e) { continue; }
            var swatchNames = [];
            var chips = [];
            for (var j = 0; j < groupSwatches.length; j++) {
                var swatchColor = groupSwatches[j].color;
                if (swatchColor.typename === "NoColor") { continue; } /* [なし] は除外 / skip the None swatch */
                swatchNames.push(groupSwatches[j].name);
                chips.push(colorToRGB255(swatchColor));
            }
            if (swatchNames.length === 0) { continue; }
            swatchGroupsInfo.push({ name: groupName, swatchNames: swatchNames, chips: chips });
        }
        return swatchGroupsInfo;
    }

    /**
     * 選択のテキスト統計（テキスト有無・段落数・テキストオブジェクト数）を求める
     *
     * 段落数はテキストが1つのときだけ数える（ディム判定がその場合しか参照せず、長文では重いため）。
     * @param {PageItem[]} items - 色付け対象のアイテム
     * @param {TextRange|null} textRange - テキスト編集中の選択範囲
     * @returns {{isText: boolean, paragraphCount: number, textObjectCount: number}} テキストの統計
     */
    function getSelectionTextStats(items, textRange) {
        var isText = false;
        var textObjectCount = 0;
        var singleTextRange = null;
        if (textRange) { isText = true; textObjectCount++; singleTextRange = textRange; }
        for (var i = 0; i < items.length; i++) {
            if (items[i].typename === "TextFrame") { isText = true; textObjectCount++; singleTextRange = items[i].textRange; }
        }
        /* 段落数も \r 基準で数える（paragraphs.length は強制改行でも増えるので、
           強制改行入りの1段落が「2段落」に見えて段落単位が有効になってしまう）
           Count paragraphs by \r as well: paragraphs.length also counts forced line breaks,
           which would make a single paragraph look like two and wrongly enable the paragraph unit */
        var paragraphCount = (textObjectCount === 1) ? countParagraphs(singleTextRange) : 0;
        return { isText: isText, paragraphCount: paragraphCount, textObjectCount: textObjectCount };
    }

    /**
     * 選択をフラット化して色付け対象のみ集める
     * @param {PageItem[]} selection - 選択
     * @returns {PageItem[]} パス・複合パス・テキストフレーム
     */
    function flattenSelection(selection) {
        var collected = [];
        if (!selection || selection.length === 0) return collected;
        for (var i = 0; i < selection.length; i++) { collectColorableItems(selection[i], collected); }
        return collected;
    }

    /**
     * 色付け可能な要素を再帰的に集める（グループ内も）
     * @param {PageItem} pageItem - 対象のアイテム
     * @param {PageItem[]} collected - 集めた結果（追加する）
     * @returns {void}
     */
    function collectColorableItems(pageItem, collected) {
        if (!pageItem) return;
        if (pageItem.typename === "PathItem" || pageItem.typename === "CompoundPathItem" || pageItem.typename === "TextFrame") { collected.push(pageItem); return; }
        if (pageItem.typename === "GroupItem") {
            var children = pageItem.pageItems;
            for (var i = 0; i < children.length; i++) { collectColorableItems(children[i], collected); }
        }
    }

    /**
     * テキスト編集中の選択範囲を取得する
     * @param {Object[]} selection - 選択
     * @returns {TextRange|null} 文字を選択中なら TextRange、そうでなければ null
     */
    function getSelectedTextRange(selection) {
        if (!selection || selection.length !== 1) return null;
        if (selection[0] && selection[0].typename === "TextRange") return selection[0];
        return null;
    }

    // =========================================
    // 着色とプレビュー / Coloring and preview
    // =========================================

    /**
     * 初回プレビューを実行する。重い選択（長文・多数オブジェクト）のときだけプログレスバーを出す
     *
     * 軽い選択では一瞬で終わるためバーは出さない（点滅を避ける）。
     * @param {{unit: string, order: string}} coloringOptions - 配色の設定
     * @param {string[]} swatchNames - 適用するスウォッチ名
     * @returns {void}
     */
    function runFirstPreview(coloringOptions, swatchNames) {
        var selection = (app.documents.length > 0) ? app.activeDocument.selection : null;
        var items = selection ? flattenSelection(selection) : [];
        var textRange = selection ? getSelectedTextRange(selection) : null;
        var isHeavy = (countBaselineCharacters(items, textRange) > HEAVY_SELECTION_CHAR_COUNT) || (items.length > HEAVY_SELECTION_ITEM_COUNT);
        if (!isHeavy) { applyColorsToSelection(coloringOptions, swatchNames); return; }
        var progress = createProgressWindow(getLabel("progress.title"));
        /* 失敗してもプログレスバーは必ず閉じる / Always close the progress bar */
        try {
            applyColorsToSelection(coloringOptions, swatchNames, false, progress);
        } finally {
            progress.close();
        }
    }

    /**
     * 選択にカラーを適用する
     *
     * 元の塗り/線/不透明度は開いた最初の1回だけ全選択ぶんスナップショット（ベースライン）し、
     * 以降のプレビューは上書きのみ＝復元も再スナップショットもしない。
     * どの配色単位でも可視グリフ・全オブジェクトを必ず塗り直すので見た目は常に正しく、
     * キャンセル時はベースラインから完全復元する（app.undo はグローバル履歴を巻き戻すため使わない）。
     * @param {{unit: string, order: string}} coloringOptions - 配色の設定
     * @param {string[]} swatchNames - 適用するスウォッチ名
     * @param {boolean} [finalApply] - 確定時の全文字着色なら true（保存した配色を再利用する）
     * @param {Object} [progress] - createProgressWindow() の戻り値（進捗を表示するとき）
     * @returns {void}
     */
    function applyColorsToSelection(coloringOptions, swatchNames, finalApply, progress) {
        if (app.documents.length === 0) return;
        var doc = app.activeDocument;
        var selection = doc.selection;
        var items = flattenSelection(selection);
        var textRange = getSelectedTextRange(selection);
        /* オブジェクト選択は適用後に選択フォーカスを戻すため控える（テキスト編集中は戻さない）
           Save the object selection to restore focus after applying (skip while editing text) */
        var savedSelection = textRange ? null : selection;
        if (progress) { progress.set(0.1, getLabel("progress.analyze")); }
        /* 同一単位のプレビューでは対象集合をキャッシュ再利用（順序変更だけなら作り直さない）
           Reuse the cached target set across previews of the same unit (order-only changes skip the rebuild) */
        var targets = collectTargetsCached(items, textRange, coloringOptions.unit);
        if (targets.length === 0) { restoreBaseline(); app.redraw(); return; }
        /* 元状態は初回だけ取得（重い文字読み取りは1回きり）。進捗はここが主コスト
           Snapshot originals only once (the heavy per-char read happens once); this is the main progress cost */
        if (progress) { progress.set(0.2, getLabel("progress.snapshot")); }
        ensureBaseline(items, textRange, progress ? function (fraction) { progress.set(0.2 + fraction * 0.6, getLabel("progress.snapshot")); } : null);

        /* 着色前の準備（間引きの要否と着色件数を決める）/ Prepare before coloring (decide decimation and count) */
        var coloringPlan = prepareColoring(items, textRange, coloringOptions.unit, targets.length, finalApply);
        if (!finalApply) { $.global.__aiApplyWasDecimated = coloringPlan.decimated; }

        if (progress) { progress.set(0.85, getLabel("progress.apply")); }
        /* ランダム系（並びシャッフル・完全ランダム抽選・自動CMYK生成）はプレビューと確定で結果が変わらないよう、
           プレビュー時に生成した配色と抽選結果を保存し、確定（finalApply）時はそれを再利用する
           Keep random results stable between preview and commit: generate & store the colors and the
           fully-random draws on preview, then reuse them on commit */
        var colors, groupIndexCache;
        if (finalApply && $.global.__aiApplyColors) {
            colors = $.global.__aiApplyColors;
            groupIndexCache = $.global.__aiApplyGroupIndex || {};
        } else {
            colors = orderColors(resolveAppliedColors(doc, targets.length, swatchNames), coloringOptions.order);
            groupIndexCache = {};
            $.global.__aiApplyColors = colors;
            $.global.__aiApplyGroupIndex = groupIndexCache;
        }
        applyColorsToTargets(targets, coloringPlan.limit, colors, coloringOptions.order === "fullrandom", coloringPlan.decimated, groupIndexCache);
        if (progress) { progress.set(1, getLabel("progress.apply")); }

        /* 適用で変わった選択（フォーカス）を元に戻す / Restore the selection changed by applying */
        if (savedSelection) { try { app.selection = savedSelection; } catch (e) {} }
        app.redraw();
    }

    /**
     * 着色前の準備をして、着色ループの制御値を返す
     *
     * 1文字単位で対象が多いプレビューは間引き（先頭 PREVIEW_CHAR_CAP だけ着色し残りは元色）にし、
     * それ以外は線なし・不透明度100 を範囲へ一括設定する。
     * @param {PageItem[]} items - 色付け対象のアイテム
     * @param {TextRange|null} textRange - テキスト編集中の選択範囲
     * @param {string} unitMode - 配色単位
     * @param {number} targetCount - 対象の数
     * @param {boolean} [finalApply] - 確定時の全文字着色なら true
     * @returns {{decimated: boolean, limit: number}} 間引くかどうかと着色する件数
     */
    function prepareColoring(items, textRange, unitMode, targetCount, finalApply) {
        var decimated = (!finalApply && unitMode === "character" && targetCount > PREVIEW_CHAR_CAP);
        if (decimated) {
            /* 前回プレビューの着色を元へ戻してから先頭だけ塗る（残りが元の色で見える）
               Repaint originals first, then color just the head (so the tail shows the original color) */
            paintBaseline();
            return { decimated: true, limit: PREVIEW_CHAR_CAP };
        }
        /* 線なし・不透明度100 は全テキスト対象で共通。文字ごとに書かず範囲へ1回だけ書く
           No stroke / 100% opacity is common to all text targets; write it once per range, not per character */
        resetTextStrokeOpacity(items, textRange);
        return { decimated: false, limit: targetCount };
    }

    /**
     * 対象の先頭 limit 件へ色を適用する
     *
     * 完全ランダムはグループごとに1回抽選（groupIndexCache に保持）。間引き時は文字ごとに線・不透明度も設定する。
     * @param {Object[]} targets - 着色対象
     * @param {number} limit - 着色する件数
     * @param {Color[]} colors - 使うカラー
     * @param {boolean} isFullRandom - 完全ランダムなら true
     * @param {boolean} decimated - 間引きプレビューなら true
     * @param {Object} groupIndexCache - 完全ランダムの抽選結果（グループキー → カラー番号）
     * @returns {void}
     */
    function applyColorsToTargets(targets, limit, colors, isFullRandom, decimated, groupIndexCache) {
        for (var i = 0; i < limit; i++) {
            var colorIndex;
            if (isFullRandom) {
                var groupKey = (typeof targets[i].groupKey === "string") ? targets[i].groupKey : String(i);
                colorIndex = fullRandomColorIndex(groupKey, groupIndexCache, colors.length);
            } else {
                colorIndex = (typeof targets[i].colorIndex === "number") ? targets[i].colorIndex : i;
            }
            var color = pickColorByIndex(colorIndex, colors);
            if (decimated) { applyColorFullToCharacter(targets[i], color); }
            else { applyColorToTarget(targets[i], color); }
        }
    }

    /**
     * 着色対象を配色単位ごとにキャッシュして返す（単位が同じなら作り直さない）
     *
     * モーダル中は選択が変わらないので、Character や TextRange などの参照も次回まで有効。
     * @param {PageItem[]} items - 色付け対象のアイテム
     * @param {TextRange|null} textRange - テキスト編集中の選択範囲
     * @param {string} unitMode - 配色単位
     * @returns {Object[]} 着色対象
     */
    function collectTargetsCached(items, textRange, unitMode) {
        if ($.global.__aiApplyTargetsUnit === unitMode && $.global.__aiApplyTargets) {
            return $.global.__aiApplyTargets;
        }
        var targets = (items.length > 0 || textRange) ? collectColorTargetsByUnit(items, textRange, unitMode) : [];
        $.global.__aiApplyTargets = targets;
        $.global.__aiApplyTargetsUnit = unitMode;
        return targets;
    }

    /**
     * 元状態のベースラインを初回だけ取得する（全選択ぶん・テキストは文字単位で忠実に保存）
     *
     * 単位に依存せず選択全体を控えるので、以後どの単位に切り替えても復元・再取得が不要。
     * @param {PageItem[]} items - 色付け対象のアイテム
     * @param {TextRange|null} textRange - テキスト編集中の選択範囲
     * @param {function|null} onProgress - 進捗（0..1）を受け取る関数
     * @returns {void}
     */
    function ensureBaseline(items, textRange, onProgress) {
        if ($.global.__aiApplyBaseline) { return; }
        var baseline = [];
        /* 進捗コンテキスト：処理済み文字数を総数で割って onProgress へ通知（テキストが無ければ null）
           Progress context: report processed chars / total to onProgress (null when there is no text) */
        var progressContext = onProgress ? { done: 0, total: countBaselineCharacters(items, textRange), report: onProgress } : null;
        if (textRange) { snapshotTarget({ kind: "textrange", node: textRange }, baseline, progressContext); }
        for (var i = 0; i < items.length; i++) {
            var pageItem = items[i];
            if (pageItem.typename === "TextFrame") { snapshotTarget({ kind: "textframe", node: pageItem }, baseline, progressContext); }
            else if (pageItem.typename === "PathItem") { snapshotTarget({ kind: "path", node: pageItem }, baseline, progressContext); }
            else if (pageItem.typename === "CompoundPathItem") { snapshotTarget({ kind: "compound", node: pageItem }, baseline, progressContext); }
        }
        $.global.__aiApplyBaseline = baseline;
    }

    /**
     * ベースライン取得で読む総文字数を見積もる（進捗の分母）
     * @param {PageItem[]} items - 色付け対象のアイテム
     * @param {TextRange|null} textRange - テキスト編集中の選択範囲
     * @returns {number} 総文字数（0 のときは 1）
     */
    function countBaselineCharacters(items, textRange) {
        var total = 0;
        if (textRange) { total += textRange.characters.length; }
        for (var i = 0; i < items.length; i++) {
            if (items[i].typename === "TextFrame") { total += items[i].textRange.characters.length; }
        }
        return total > 0 ? total : 1;
    }

    /**
     * 現在のプレビューを確定する（ベースラインとキャッシュを破棄＝以後 restore しない）
     *
     * 閉じる時・初回開始前に呼ぶ。
     * @returns {void}
     */
    function commitPreview() {
        $.global.__aiApplyBaseline = null;
        $.global.__aiApplySwatchMap = null;
        $.global.__aiApplyTargets = null;
        $.global.__aiApplyTargetsUnit = null;
        $.global.__aiApplyColors = null;
        $.global.__aiApplyGroupIndex = null;
        $.global.__aiApplyWasDecimated = null;
    }

    /**
     * ベースライン（元の塗り/線/不透明度）を画面へ書き戻す（破棄はしない）
     *
     * 間引きプレビューで着色しなかった残りを元の色に戻すのに使う。
     * @returns {void}
     */
    function paintBaseline() {
        var baseline = $.global.__aiApplyBaseline;
        if (!baseline) { return; }
        for (var i = 0; i < baseline.length; i++) {
            /* 書き戻せないものは飛ばす / Skip entries that cannot be restored */
            try { restoreSnapshotEntry(baseline[i]); } catch (e) { }
        }
    }

    /**
     * ベースラインを書き戻して破棄する（キャンセル・Esc・閉じる時。以後 restore しない）
     * @returns {void}
     */
    function restoreBaseline() {
        paintBaseline();
        $.global.__aiApplyBaseline = null;
    }

    /**
     * スナップショット1件を復元する
     * @param {Object} snapshotEntry - snapshotTarget() が記録した1件
     * @returns {void}
     */
    function restoreSnapshotEntry(snapshotEntry) {
        var node = snapshotEntry.node;
        if (snapshotEntry.kind === "textrun") {
            /* 同色ランを範囲1回で書き戻す（1文字ずつより桁違いに速い）/ Restore a run in one span write (far faster than per character) */
            var runRange = snapshotEntry.story.textRange;
            runRange.start = snapshotEntry.start; runRange.end = snapshotEntry.end;
            runRange.fillColor = snapshotEntry.fill; runRange.strokeColor = snapshotEntry.stroke; runRange.opacity = snapshotEntry.opacity;
        }
        else if (snapshotEntry.kind === "path") { node.fillColor = snapshotEntry.fill; node.stroked = snapshotEntry.stroked; node.opacity = snapshotEntry.opacity; }
        else if (snapshotEntry.kind === "compound") {
            for (var i = 0; i < snapshotEntry.subs.length; i++) {
                var subSnapshot = snapshotEntry.subs[i];
                /* 書き戻せないサブパスは飛ばす / Skip sub-paths that cannot be restored */
                try { subSnapshot.sub.fillColor = subSnapshot.fill; subSnapshot.sub.stroked = subSnapshot.stroked; subSnapshot.sub.opacity = subSnapshot.opacity; } catch (e) { }
            }
        }
    }

    /**
     * 着色前の対象の元の状態をスナップショット配列へ保存する（テキストは文字単位で忠実に保存）
     * @param {{kind: string, node: Object}} target - 対象（textrange / textframe / path / compound）
     * @param {Object[]} snapshotEntries - 保存先（追加する）
     * @param {Object|null} progressContext - 進捗の通知先
     * @returns {void}
     */
    function snapshotTarget(target, snapshotEntries, progressContext) {
        var node = target.node;
        if (target.kind === "textrange") { snapshotCharacters(node, snapshotEntries, progressContext); }
        else if (target.kind === "textframe") { snapshotCharacters(node.textRange, snapshotEntries, progressContext); }
        else if (target.kind === "path") { snapshotEntries.push({ kind: "path", node: node, fill: node.fillColor, stroked: node.stroked, opacity: node.opacity }); }
        else if (target.kind === "compound") {
            var subPaths = node.pathItems;
            var subSnapshots = [];
            for (var i = 0; i < subPaths.length; i++) { subSnapshots.push({ sub: subPaths[i], fill: subPaths[i].fillColor, stroked: subPaths[i].stroked, opacity: subPaths[i].opacity }); }
            snapshotEntries.push({ kind: "compound", node: node, subs: subSnapshots });
        }
    }

    /**
     * テキスト範囲を文字単位でスナップショットする（元が混色でも忠実に戻せる）
     *
     * 見た目が同じ連続文字は1つのランにまとめて保存する。
     * @param {TextRange} range - 対象のテキスト範囲
     * @param {Object[]} snapshotEntries - 保存先（追加する）
     * @param {Object|null} progressContext - 進捗の通知先
     * @returns {void}
     */
    function snapshotCharacters(range, snapshotEntries, progressContext) {
        var story = range.story;
        var characters = range.characters;
        var count = characters.length;
        /* 連続範囲なので i 文字目の story オフセットは range.start + i（ch.start の DOM 読み取りを省く）
           The range is contiguous, so character i's story offset is range.start + i (avoids reading ch.start) */
        var rangeStart = range.start;
        var started = false;
        var runFill = null, runStroke = null, runOpacity = 0, runKey = null, runStart = 0, runEnd = 0;
        for (var i = 0; i < count; i++) {
            var character = characters[i];
            var fill = character.fillColor;
            var stroke = character.strokeColor;
            var opacity = character.opacity;
            var charStart = rangeStart + i;
            /* 一定文字ごとに進捗を通知（毎回だと更新自体が重い）/ Report progress every N chars (updating each time is itself costly) */
            if (progressContext) { progressContext.done++; if ((progressContext.done % 200) === 0) { progressContext.report(progressContext.done / progressContext.total); } }
            var runCandidateKey = colorKey(fill) + "|" + colorKey(stroke) + "|" + opacity;
            if (started && runCandidateKey === runKey) {
                /* 見た目が同じ連続文字は run を伸ばす / extend the run for contiguous identical characters */
                runEnd = charStart + 1;
            } else {
                if (started) { snapshotEntries.push({ kind: "textrun", story: story, start: runStart, end: runEnd, fill: runFill, stroke: runStroke, opacity: runOpacity }); }
                started = true;
                runKey = runCandidateKey; runFill = fill; runStroke = stroke; runOpacity = opacity;
                runStart = charStart; runEnd = charStart + 1;
            }
        }
        if (started) { snapshotEntries.push({ kind: "textrun", story: story, start: runStart, end: runEnd, fill: runFill, stroke: runStroke, opacity: runOpacity }); }
    }

    /**
     * 色の同一性キーを返す（NoColor 対応。塗り・線のラン判定に使う）
     * @param {Color} color - 対象の色
     * @returns {string} 同一性キー
     */
    function colorKey(color) {
        if (color.typename === "NoColor") { return "N"; }
        return serializeColor(color);
    }

    /**
     * 配色単位に応じた着色対象を集める
     * @param {PageItem[]} items - 色付け対象のアイテム
     * @param {TextRange|null} textRange - テキスト編集中の選択範囲
     * @param {string} unitMode - 配色単位
     * @returns {Object[]} 着色対象
     */
    function collectColorTargetsByUnit(items, textRange, unitMode) {
        var targets = [];
        if (textRange) { pushTextRangeTargets(textRange, unitMode, targets); return targets; }
        if (items.length === 1 && items[0].typename === "TextFrame") {
            if (unitMode === "object") { targets.push({ kind: "textframe", node: items[0] }); }
            else { pushTextRangeTargets(items[0].textRange, unitMode, targets); }
            return targets;
        }
        var sortedItems = items.slice();
        sortByPosition(sortedItems);
        for (var i = 0; i < sortedItems.length; i++) {
            var pageItem = sortedItems[i];
            if (pageItem.typename === "TextFrame" && unitMode !== "object") { pushTextRangeTargets(pageItem.textRange, unitMode, targets); }
            else if (pageItem.typename === "PathItem") { targets.push({ kind: "path", node: pageItem }); }
            else if (pageItem.typename === "CompoundPathItem") { targets.push({ kind: "compound", node: pageItem }); }
            else if (pageItem.typename === "TextFrame") { targets.push({ kind: "textframe", node: pageItem }); }
        }
        return targets;
    }

    /**
     * テキスト範囲を配色単位ごとに分割して着色対象にする
     * @param {TextRange} textRange - 対象のテキスト範囲
     * @param {string} unitMode - 配色単位
     * @param {Object[]} targets - 着色対象（追加する）
     * @returns {void}
     */
    function pushTextRangeTargets(textRange, unitMode, targets) {
        if (unitMode === "word") { pushStaggeredWordTargets(textRange, targets); return; }
        if (unitMode === "paragraph") { pushParagraphTargets(textRange, targets); return; }
        var ranges = getTextUnitRanges(textRange, unitMode);
        if (!ranges) { targets.push({ kind: "textrange", node: textRange }); return; }
        for (var i = 0; i < ranges.length; i++) { targets.push({ kind: "textrange", node: ranges[i] }); }
    }

    /**
     * 単語ごとの着色対象を作る（Illustrator の単語境界で単語スパンを作り、各行先頭は互い違い）
     *
     * 1文字ずつ塗ると文字数の二乗で重くなるため、単語スパン1回の書き込みにまとめている。
     * スパンは次の単語の開始まで伸ばし、句読点・スペースを直前の単語色にする
     * （行頭は先頭の単語色、行末の残りは最後の単語色。1文字ずつ塗っていた頃と同じ結果）。
     * @param {TextRange} textRange - 対象のテキスト範囲
     * @param {Object[]} targets - 着色対象（追加する）
     * @returns {void}
     */
    function pushStaggeredWordTargets(textRange, targets) {
        var story = textRange.story;
        var lines = textRange.lines;
        var lineCount = lines.length;
        for (var lineIndex = 0; lineIndex < lineCount; lineIndex++) {
            var line = lines[lineIndex];
            var words = line.words;
            var wordCount = words.length;
            for (var wordIndex = 0; wordIndex < wordCount; wordIndex++) {
                /* 行頭の記号類は先頭の単語に、行末の残りは最後の単語に含める
                   Leading symbols join the first word; the tail of the line joins the last word */
                var spanStart = (wordIndex === 0) ? line.start : words[wordIndex].start;
                var spanEnd = (wordIndex === wordCount - 1) ? line.end : words[wordIndex + 1].start;
                if (spanEnd <= spanStart) { continue; }
                /* groupKey は行内で一意（完全ランダムで単語ごとに別色を抽選するため）
                   groupKey is unique per line+word so fully random draws a distinct color per word */
                targets.push({ kind: "span", story: story, start: spanStart, end: spanEnd, colorIndex: wordIndex + lineIndex, groupKey: lineIndex + ":" + wordIndex });
            }
        }
    }

    /**
     * 段落ごとの着色対象を作る（本文を \r で走査してスパンを作る）
     *
     * textRange.paragraphs は強制改行でも区切られてしまい、1段落が複数に割れるため使わない。
     * @param {TextRange} textRange - 対象のテキスト範囲
     * @param {Object[]} targets - 着色対象（追加する）
     * @returns {void}
     */
    function pushParagraphTargets(textRange, targets) {
        var story = textRange.story;
        var storyText = story.textRange.contents;
        var rangeEnd = textRange.end;
        var head = textRange.start;
        while (head < rangeEnd) {
            var returnAt = storyText.indexOf("\r", head);
            /* 改行が無い、または選択範囲の外なら、そこが最後の段落 / No return, or one past the selection, ends the last paragraph */
            var isLastParagraph = (returnAt === -1 || returnAt >= rangeEnd);
            var tail = isLastParagraph ? rangeEnd : returnAt;
            /* 空段落（改行だけ）は塗る文字が無いので飛ばす / An empty paragraph has nothing to color */
            if (tail > head) { targets.push({ kind: "span", story: story, start: head, end: tail }); }
            if (isLastParagraph) break;
            head = returnAt + 1;
        }
    }

    /**
     * 範囲内の段落数を \r 基準で数える（塗る文字が無い空段落は数えない）
     * @param {TextRange} textRange - 対象のテキスト範囲
     * @returns {number} 段落数
     */
    function countParagraphs(textRange) {
        var paragraphSpans = [];
        pushParagraphTargets(textRange, paragraphSpans);
        return paragraphSpans.length;
    }

    /**
     * 配色単位に対応するテキスト範囲のコレクションを返す（単語・段落は pushTextRangeTargets で先に処理）
     * @param {TextRange} textRange - 対象のテキスト範囲
     * @param {string} unitMode - 配色単位
     * @returns {Object|null} characters / lines。該当しなければ null
     */
    function getTextUnitRanges(textRange, unitMode) {
        if (unitMode === "character") return textRange.characters;
        if (unitMode === "line") return textRange.lines;
        return null;
    }

    /**
     * スパン（story 内の start〜end）に対応する TextRange を作る
     *
     * story.textRange は呼ぶたび独立したオブジェクトを返すので、start/end を書き換えて使える。
     * @param {{story: Story, start: number, end: number}} target - スパンの着色対象
     * @returns {TextRange} スパンの範囲
     */
    function getSpanRange(target) {
        var spanRange = target.story.textRange;
        spanRange.start = target.start;
        spanRange.end = target.end;
        return spanRange;
    }

    /**
     * テキストの線・不透明度をまとめて初期化する（線なし・不透明度100）
     *
     * 全テキスト対象で同じ値なので、対象範囲へ1回だけ書き込む（1文字ずつ書くと重い）。
     * @param {PageItem[]} items - 色付け対象のアイテム
     * @param {TextRange|null} textRange - テキスト編集中の選択範囲
     * @returns {void}
     */
    function resetTextStrokeOpacity(items, textRange) {
        if (textRange) { applyNoStrokeFullOpacity(textRange); }
        for (var i = 0; i < items.length; i++) {
            if (items[i].typename === "TextFrame") { applyNoStrokeFullOpacity(items[i].textRange); }
        }
    }

    /**
     * テキスト範囲に「線なし・不透明度100」を1回で設定する
     * @param {TextRange} range - 対象のテキスト範囲（Character も可）
     * @returns {void}
     */
    function applyNoStrokeFullOpacity(range) {
        range.strokeColor = new NoColor();
        range.opacity = 100;
    }

    /**
     * 間引きプレビュー用に、1文字へ塗り・線なし・不透明度100 をまとめて設定する
     *
     * 範囲全体の resetTextStrokeOpacity を使わず文字単位で設定するのは、着色範囲外の元の線・不透明度を保つため。
     * @param {{node: Object}} target - 1文字の着色対象
     * @param {Color} color - 塗りの色
     * @returns {void}
     */
    function applyColorFullToCharacter(target, color) {
        var node = target.node;
        node.fillColor = color;
        applyNoStrokeFullOpacity(node); /* Character も strokeColor/opacity を持つので流用 / a Character also has strokeColor/opacity */
    }

    /**
     * 着色対象に塗りを設定する（テキストの線・不透明度は resetTextStrokeOpacity でまとめて処理済み）
     * @param {Object} target - 着色対象
     * @param {Color} color - 塗りの色
     * @returns {void}
     */
    function applyColorToTarget(target, color) {
        var node = target.node;
        if (target.kind === "span") { getSpanRange(target).fillColor = color; }
        else if (target.kind === "textrange") { node.fillColor = color; }
        else if (target.kind === "textframe") { node.textRange.fillColor = color; }
        else if (target.kind === "path") { node.fillColor = color; node.stroked = false; node.opacity = 100; }
        else if (target.kind === "compound") {
            var subPaths = node.pathItems;
            for (var i = 0; i < subPaths.length; i++) { subPaths[i].fillColor = color; subPaths[i].stroked = false; subPaths[i].opacity = 100; }
        }
    }

    /**
     * アイテムを位置順に並べ替える（横に広がっていれば左から、縦なら上から）
     * @param {PageItem[]} items - 対象（その場で並べ替える）
     * @returns {void}
     */
    function sortByPosition(items) {
        var minLeft = Infinity, maxLeft = -Infinity, minTop = Infinity, maxTop = -Infinity;
        for (var i = 0; i < items.length; i++) {
            var left = items[i].left;
            var top = items[i].top;
            if (left < minLeft) minLeft = left;
            if (left > maxLeft) maxLeft = left;
            if (top < minTop) minTop = top;
            if (top > maxTop) maxTop = top;
        }
        if (maxLeft - minLeft > maxTop - minTop) { items.sort(function (itemA, itemB) { return comparePositionKeys(itemA.left, itemB.left, itemB.top, itemA.top); }); }
        else { items.sort(function (itemA, itemB) { return comparePositionKeys(itemB.top, itemA.top, itemA.left, itemB.left); }); }
    }

    /**
     * ソート用の比較（主キー → 副キー）
     * @param {number} primaryA - 主キー（A）
     * @param {number} primaryB - 主キー（B）
     * @param {number} secondaryA - 副キー（A）
     * @param {number} secondaryB - 副キー（B）
     * @returns {number} 比較結果
     */
    function comparePositionKeys(primaryA, primaryB, secondaryA, secondaryB) {
        return primaryA == primaryB ? secondaryA - secondaryB : primaryA - primaryB;
    }

    // =========================================
    // カラーの決定と生成 / Resolving and generating colors
    // =========================================

    /**
     * 配列をランダムに並べ替えた複製を返す
     * @param {Array} source - 元の配列
     * @returns {Array} シャッフルした複製
     */
    function shuffleArray(source) {
        var shuffled = source.slice();
        for (var i = shuffled.length - 1; i > 0; i--) {
            var j = Math.floor(Math.random() * (i + 1));
            var swappedValue = shuffled[i];
            shuffled[i] = shuffled[j];
            shuffled[j] = swappedValue;
        }
        return shuffled;
    }

    /**
     * 整数の乱数を返す
     * @param {number} minValue - 下限（含む）
     * @param {number} maxValue - 上限（含む）
     * @returns {number} minValue〜maxValue の整数
     */
    function randomInt(minValue, maxValue) {
        return Math.floor(Math.random() * (maxValue - minValue + 1)) + minValue;
    }

    /**
     * 使用するカラーを決める（取り込んだスウォッチ名を最優先、無ければ自動カラー）
     * @param {Document} doc - 対象ドキュメント
     * @param {number} targetCount - 着色対象の数（自動 CMYK の生成数）
     * @param {string[]} swatchNames - 取り込んだスウォッチ名
     * @returns {Color[]} 使うカラー
     */
    function resolveAppliedColors(doc, targetCount, swatchNames) {
        var colors = [];
        if (swatchNames && swatchNames.length > 0) {
            /* 名前→色マップはダイアログ中1回だけ構築してキャッシュ（モーダル中スウォッチは変わらない）。
               プレビューは頻繁に走るので毎回の全スウォッチ走査を避ける
               Build the name->color map once per dialog and cache it (swatches never change during the
               modal); the preview fires often, so avoid rescanning all swatches every time */
            var swatchColorsByName = $.global.__aiApplySwatchMap || ($.global.__aiApplySwatchMap = buildSwatchColorMap(doc));
            for (var i = 0; i < swatchNames.length; i++) {
                var swatchColor = swatchColorsByName["$" + swatchNames[i]];
                if (swatchColor) { colors.push(swatchColor); }
            }
        }
        if (colors.length >= 1) { return colors; }
        if (doc.documentColorSpace === DocumentColorSpace.CMYK) { return generateUniqueCMYPalette(targetCount, CMYK_FALLBACK_MAX_TOTAL); }
        return getDefaultRGBColors();
    }

    /**
     * スウォッチ名 → 色のマップを作る（選択状態に依存しない）。キーは "$"+名前でプロトタイプ汚染を避ける
     * @param {Document} doc - 対象ドキュメント
     * @returns {Object} "$"+スウォッチ名 → 色
     */
    function buildSwatchColorMap(doc) {
        var swatchColorsByName = {};
        var swatches = doc.swatches;
        for (var i = 0; i < swatches.length; i++) {
            var mapKey = "$" + swatches[i].name;
            if (swatchColorsByName[mapKey] === undefined) { swatchColorsByName[mapKey] = swatches[i].color; }
        }
        return swatchColorsByName;
    }

    /**
     * カラーを同一性キー文字列にする（colorKey のラン判定に使う）
     *
     * SpotColor は tint も含める（tint 違いを別色として区別するため）。
     * @param {Color} color - 対象の色
     * @returns {string} 同一性キー
     */
    function serializeColor(color) {
        var colorType = color.typename;
        if (colorType === "CMYKColor") { return "C," + color.cyan + "," + color.magenta + "," + color.yellow + "," + color.black; }
        if (colorType === "RGBColor") { return "R," + color.red + "," + color.green + "," + color.blue; }
        if (colorType === "GrayColor") { return "G," + color.gray; }
        if (colorType === "SpotColor") {
            var tint = (typeof color.tint === "number") ? color.tint : 100;
            return "S," + tint + "," + serializeColor(color.spot.color);
        }
        return "R,128,128,128";
    }

    /**
     * 配色順に並べ替える（完全ランダムは適用側でインデックスを抽選するので並びは変えない）
     * @param {Color[]} colors - 取り込んだカラー
     * @param {string} orderMode - 配色順
     * @returns {Color[]} 並べ替えたカラー
     */
    function orderColors(colors, orderMode) {
        if (orderMode === "reverse") return colors.slice().reverse();
        if (orderMode === "random") return shuffleArray(colors);
        return colors;
    }

    /**
     * 完全ランダム用に、グループキーごとのカラー番号を抽選してキャッシュする
     * @param {string} groupKey - グループキー
     * @param {Object} groupIndexCache - 抽選結果のキャッシュ
     * @param {number} colorCount - カラーの数
     * @returns {number} カラー番号
     */
    function fullRandomColorIndex(groupKey, groupIndexCache, colorCount) {
        if (groupIndexCache[groupKey] === undefined) { groupIndexCache[groupKey] = randomInt(0, colorCount - 1); }
        return groupIndexCache[groupKey];
    }

    /**
     * 番号に応じて色を取得する（末尾を超えたらループ）。不透明度は適用側で設定する
     * @param {number} index - カラー番号
     * @param {Color[]} colors - 使うカラー
     * @returns {Color} 色
     */
    function pickColorByIndex(index, colors) {
        return colors[index % colors.length];
    }

    /**
     * CMYK カラーを作る
     * @param {number} cyan - シアン
     * @param {number} magenta - マゼンタ
     * @param {number} yellow - イエロー
     * @param {number} black - ブラック
     * @returns {CMYKColor} 色
     */
    function buildCMYKColor(cyan, magenta, yellow, black) {
        var color = new CMYKColor();
        color.cyan = cyan; color.magenta = magenta; color.yellow = yellow; color.black = black;
        return color;
    }

    /**
     * RGB カラーを作る
     * @param {number} red - 赤
     * @param {number} green - 緑
     * @param {number} blue - 青
     * @returns {RGBColor} 色
     */
    function buildRGBColor(red, green, blue) {
        var color = new RGBColor();
        color.red = red; color.green = green; color.blue = blue;
        return color;
    }

    /**
     * RGB ドキュメント用の定義済みカラーを返す
     * @returns {RGBColor[]} 定義済みカラー
     */
    function getDefaultRGBColors() {
        var rgbDefinitions = [[222, 84, 25], [245, 233, 40], [41, 163, 57], [53, 157, 209], [173, 127, 71], [238, 176, 51]];
        var colors = [];
        for (var i = 0; i < rgbDefinitions.length; i++) { colors.push(buildRGBColor(rgbDefinitions[i][0], rgbDefinitions[i][1], rgbDefinitions[i][2])); }
        return colors;
    }

    /**
     * CM/CY/MY の2チャンネル（K=0）をランダムに1色ぶん作る
     * @param {string} channelPair - "CM" / "CY" / "MY"
     * @param {number} maxTotal - 2チャンネルの合計上限
     * @returns {{cyan: number, magenta: number, yellow: number}|null} 各チャンネルの値。作れなければ null
     */
    function pickTwoChannelCMY(channelPair, maxTotal) {
        var amountA = randomInt(1, Math.min(100, maxTotal - 1));
        var amountBMax = Math.min(100, maxTotal - amountA);
        if (amountBMax < 1) return null;
        var amountB = randomInt(1, amountBMax);
        if (channelPair === "CM") return { cyan: amountA, magenta: amountB, yellow: 0 };
        if (channelPair === "CY") return { cyan: amountA, magenta: 0, yellow: amountB };
        return { cyan: 0, magenta: amountA, yellow: amountB };
    }

    /**
     * CMYK 用に、CM/CY/MY だけでできるだけ重ならない色を作る
     * @param {number} count - 作る色の数
     * @param {number} maxTotal - 2チャンネルの合計上限
     * @returns {CMYKColor[]} 作った色
     */
    function generateUniqueCMYPalette(count, maxTotal) {
        var palette = [];
        var seen = {};
        var channelPairs = ["CM", "CY", "MY"];
        var minDistance = CMYK_FALLBACK_MIN_DISTANCE;
        var maxAttempts = Math.max(3000, count * 120);
        var attempts = 0;
        while (palette.length < count) {
            attempts++;
            /* 見つかりにくくなったら距離の条件を少しずつ緩める / Relax the distance as attempts grow */
            if (minDistance > 0 && (attempts % 500) === 0) { minDistance = Math.max(0, minDistance - 5); }
            var allowDuplicate = attempts > maxAttempts;
            if (allowDuplicate) { minDistance = 0; }
            if ((attempts % 37) === 0) { channelPairs = shuffleArray(channelPairs); }
            var channelPair = channelPairs[palette.length % channelPairs.length];
            var channels = pickTwoChannelCMY(channelPair, maxTotal);
            if (!channels) continue;
            var channelKey = channels.cyan + "," + channels.magenta + "," + channels.yellow;
            if (!allowDuplicate && seen[channelKey]) continue;
            if (minDistance > 0 && !isFarEnoughCMY(channels.cyan, channels.magenta, channels.yellow, palette, minDistance)) continue;
            seen[channelKey] = true;
            palette.push(buildCMYKColor(channels.cyan, channels.magenta, channels.yellow, 0));
        }
        return palette;
    }

    /**
     * 採用済みの色から十分離れているかを判定する
     * @param {number} cyan - シアン
     * @param {number} magenta - マゼンタ
     * @param {number} yellow - イエロー
     * @param {CMYKColor[]} acceptedColors - 採用済みの色
     * @param {number} minDistance - 必要な最小距離
     * @returns {boolean} 十分離れていれば true
     */
    function isFarEnoughCMY(cyan, magenta, yellow, acceptedColors, minDistance) {
        for (var i = 0; i < acceptedColors.length; i++) {
            if (cmyDistance(cyan, magenta, yellow, acceptedColors[i].cyan, acceptedColors[i].magenta, acceptedColors[i].yellow) < minDistance) return false;
        }
        return true;
    }

    /**
     * CMY のマンハッタン距離を求める
     * @param {number} cyan1 - 色1のシアン
     * @param {number} magenta1 - 色1のマゼンタ
     * @param {number} yellow1 - 色1のイエロー
     * @param {number} cyan2 - 色2のシアン
     * @param {number} magenta2 - 色2のマゼンタ
     * @param {number} yellow2 - 色2のイエロー
     * @returns {number} 距離
     */
    function cmyDistance(cyan1, magenta1, yellow1, cyan2, magenta2, yellow2) {
        return Math.abs(cyan1 - cyan2) + Math.abs(magenta1 - magenta2) + Math.abs(yellow1 - yellow2);
    }

    /**
     * 色が白かを判定する
     * @param {Color} color - 対象の色
     * @returns {boolean} 白なら true
     */
    function isWhiteColor(color) {
        if (color.typename === "CMYKColor") return color.cyan === 0 && color.magenta === 0 && color.yellow === 0 && color.black === 0;
        if (color.typename === "RGBColor") return color.red === 255 && color.green === 255 && color.blue === 255;
        return false;
    }

    /**
     * スウォッチが白だけかを判定する
     * @param {Swatch[]} swatches - 対象のスウォッチ
     * @returns {boolean} すべて白なら true
     */
    function allWhiteSwatches(swatches) {
        for (var i = 0; i < swatches.length; i++) {
            if (!isWhiteColor(swatches[i].color)) return false;
        }
        return true;
    }

    /**
     * Illustrator のカラーを RGB（0..255）に換算する（チップ表示用）
     * @param {Color} color - 対象の色
     * @returns {number[]} [r, g, b]
     */
    function colorToRGB255(color) {
        var colorType = color.typename;
        if (colorType === "RGBColor") return [Math.round(color.red), Math.round(color.green), Math.round(color.blue)];
        if (colorType === "CMYKColor") {
            var blackFactor = 1 - color.black / 100;
            return [
                Math.round(255 * (1 - color.cyan / 100) * blackFactor),
                Math.round(255 * (1 - color.magenta / 100) * blackFactor),
                Math.round(255 * (1 - color.yellow / 100) * blackFactor)
            ];
        }
        if (colorType === "GrayColor") {
            var grayLevel = Math.round(255 * (1 - color.gray / 100));
            return [grayLevel, grayLevel, grayLevel];
        }
        if (colorType === "SpotColor") {
            var baseRgb = colorToRGB255(color.spot.color);
            var tint = (typeof color.tint === "number") ? color.tint / 100 : 1;
            return [
                Math.round(255 - (255 - baseRgb[0]) * tint),
                Math.round(255 - (255 - baseRgb[1]) * tint),
                Math.round(255 - (255 - baseRgb[2]) * tint)
            ];
        }
        return [128, 128, 128];
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    showDialog();

})();
