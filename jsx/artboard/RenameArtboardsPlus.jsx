#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

接頭辞・接尾辞・元のアートボード名・連番を組み合わせて、アートボード名を一括で変更するツール。プレビューで結果を確認してから適用でき、命名ルールはプリセットとして書き出せる。

詳細はREADMEを参照。

*/

/*

### Overview

Batch renames artboards by combining prefixes, suffixes, the original artboard name, and sequential numbering. The result can be checked in the preview before applying, and naming rules can be exported as presets.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "RenameArtboardsPlus";          /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.3.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-04-20";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-08-06";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/RenameArtboardsPlus.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/RenameArtboardsPlus.md
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n80f9534bc6fb"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * 現在のUI言語を判定する
     * @returns {string} "ja" または "en"
     */
    function getUILanguage() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = getUILanguage();

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "アートボード名の一括設定", en: "Batch Rename Artboards" }
        },
        panel: {
            prefix:  { ja: "接頭辞", en: "Prefix" },
            name:    { ja: "アートボード名と番号", en: "Artboard Name & Number" },
            suffix:  { ja: "接尾辞", en: "Suffix" },
            preview: { ja: "プレビュー", en: "Preview" }
        },
        fieldLabel: {
            fileName:    { ja: "ファイル名", en: "File Name" },
            separator:   { ja: "区切り文字", en: "Separator" },
            string:      { ja: "文字列", en: "String" },
            format:      { ja: "連番形式", en: "Numbering Format" },
            startNumber: { ja: "開始番号", en: "Start Number" },
            increment:   { ja: "増分", en: "Increment" }
        },
        radio: {
            useFileNo:     { ja: "参照しない", en: "Do not use" },
            useFileYes:    { ja: "参照する", en: "Use" },
            separatorNone: { ja: "なし", en: "None" }
        },
        nameStyle: {
            none:                 { ja: "なし", en: "None" },
            number:               { ja: "番号", en: "Number" },
            name:                 { ja: "名称", en: "Name" },
            numberDashName:       { ja: "番号-名称", en: "Number-Name" },
            numberUnderscoreName: { ja: "番号_名称", en: "Number_Name" }
        },
        numberingFormat: {
            none:       { ja: "なし", en: "None" },
            numeric:    { ja: "数字", en: "Number" },
            alphaUpper: { ja: "アルファベット（大文字）", en: "Alphabet (Upper)" },
            alphaLower: { ja: "アルファベット（小文字）", en: "Alphabet (Lower)" }
        },
        preset: {
            none: { ja: "(未選択)", en: "(None)" },
            savePrompt: {
                ja: "プリセットを書き出す場所と名前を指定してください",
                en: "Specify the location and name to export the preset"
            }
        },
        preview: {
            invalid:   { ja: "※数値が正しくありません", en: "※ Invalid number" },
            truncated: { ja: "...（以下省略）", en: "... (more)" }
        },
        button: {
            ok:           { ja: "OK", en: "OK" },
            cancel:       { ja: "キャンセル", en: "Cancel" },
            exportPreset: { ja: "プリセット書き出し", en: "Export Preset" }
        },
        alert: {
            title:        { ja: "エラー", en: "Error" },
            exportFailed: { ja: "ファイルを書き込めませんでした。", en: "Could not write the file." },
            exportSuccess: {
                ja: "プリセットを書き出しました：\n",
                en: "Preset exported:\n"
            },
            exportError: {
                ja: "プリセットの保存に失敗しました：\n",
                en: "Failed to save preset:\n"
            },
            generalError: {
                ja: "エラーが発生しました：\n",
                en: "An error occurred:\n"
            },
            invalidInput: {
                ja: "入力内容に誤りがあります。\n正しい数値を指定してください。",
                en: "There is an error in your input.\nPlease enter valid numbers."
            },
            emptyName: {
                ja: "アートボード名が空になります。\n接頭辞・名称・接尾辞・連番のいずれかを指定してください。",
                en: "Artboard name would be empty.\nSpecify at least one of prefix, name, suffix, or numbering."
            }
        }
    };

    /**
     * LABELS をカテゴリ・キーの順にたどってラベルを取得する
     * @param {...string} labelPath - たどるキー（例: getLabel("dialog", "title")）
     * @returns {string} ローカライズされた文字列（見つからなければ空文字）
     */
    function getLabel() {
        var node = LABELS;
        for (var i = 0; i < arguments.length; i++) {
            if (node == null) break;
            node = node[arguments[i]];
        }
        return (node && node[uiLang] != null) ? node[uiLang] : "";
    }

    /**
     * 入力欄の見出しをコロン付きで取得する（日本語は全角、英語は半角）
     * @param {string} labelKey - LABELS.fieldLabel のキー
     * @returns {string} コロンを付けたラベル
     */
    function getFieldLabel(labelKey) {
        return getLabel("fieldLabel", labelKey) + (uiLang === "ja" ? "：" : ":");
    }

    /**
     * キー配列をローカライズ済みのドロップダウン項目に変換する
     * @param {string} categoryName - LABELS のカテゴリ名
     * @param {string[]} labelKeys - カテゴリ内のキーを表示順に並べた配列
     * @returns {string[]} 表示用ラベルの配列
     */
    function toLabelList(categoryName, labelKeys) {
        var labels = [];
        for (var i = 0; i < labelKeys.length; i++) labels.push(getLabel(categoryName, labelKeys[i]));
        return labels;
    }

    // =========================================
    // UIレイアウトの共通設定 / Shared UI layout
    // =========================================

    /* ウィンドウ・パネルの余白と間隔 / Window & panel margins and spacing */
    var WINDOW_MARGINS = 16;                 /* ウィンドウ外周の余白 / window margin */
    var WINDOW_SPACING = 12;                 /* ウィンドウ内の要素間隔 / window spacing */
    var PANEL_MARGINS  = [16, 20, 16, 12];   /* パネル余白 [左,上,右,下] / panel margins */
    var PANEL_SPACING  = 12;                 /* パネル内の要素間隔 / panel spacing */
    var COLUMN_SPACING = 12;                 /* 2カラムの間隔 / gap between columns */
    var FIELD_SPACING  = 6;                  /* 入力欄が並ぶパネル内の間隔 / spacing inside input-heavy panels */

    /**
     * ウィンドウの共通設定
     * @param {Window} win - 対象のウィンドウ
     * @param {number} [spacing] - 要素間隔（省略時は WINDOW_SPACING）
     * @returns {void}
     */
    function setupWindow(win, spacing) {
        win.orientation = "column";
        win.alignChildren = "fill";
        win.margins = WINDOW_MARGINS;
        win.spacing = (typeof spacing === "number") ? spacing : WINDOW_SPACING;
    }

    /**
     * パネルの共通設定
     * @param {Object} panel - 対象のパネル
     * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function setupPanel(panel, spacing) {
        panel.orientation = "column";
        panel.alignChildren = ["fill", "top"];
        panel.alignment = "fill";
        panel.margins = PANEL_MARGINS;
        panel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /**
     * 行グループの共通設定（ボタン列など）
     * @param {Object} group - 対象のグループ
     * @param {string} [alignment] - グループ自体の配置（省略時は "left"）
     * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function setupRow(group, alignment, spacing) {
        group.orientation = "row";
        group.alignment = alignment || "left";
        group.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /**
     * ラベル付きパネルを生成する（共通レイアウト適用）
     * @param {Object} parentGroup - 追加先のグループ
     * @param {string} panelTitle - パネル見出し
     * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
     * @returns {Object} 追加したパネル
     */
    function addPanel(parentGroup, panelTitle, spacing) {
        var panel = parentGroup.add("panel");
        panel.text = panelTitle;
        setupPanel(panel, spacing);
        return panel;
    }

    // =========================================
    // 定数 / Constants
    // =========================================

    /* 連番形式キー（LABELS.numberingFormat のキーを兼ね、dropdown の表示順になる。"none" は連番なし）
       Numbering format keys (also LABELS.numberingFormat keys; the array order is the dropdown order) */
    var NUMBERING_FORMAT_KEYS = ["none", "numeric", "alphaUpper", "alphaLower"];

    /* アートボード名スタイルキー（LABELS.nameStyle のキーを兼ね、dropdown の表示順になる）
       Artboard name style keys (also LABELS.nameStyle keys; the array order is the dropdown order) */
    var ARTBOARD_NAME_STYLE_KEYS = ["none", "number", "name", "numberDashName", "numberUnderscoreName"];

    /* 区切り文字の候補（ラジオボタンの並び順になる）/ Separator choices (the array order is the radio button order) */
    var SEPARATOR_VALUES = ["", "-", "_"];

    /* 連番形式ごとの開始値デフォルト / Default start value per numbering format */
    var DEFAULT_START_VALUES = { numeric: "1", alphaUpper: "A", alphaLower: "a" };

    /* プレビューに表示する最大件数 / Maximum number of preview rows */
    var PREVIEW_MAX_ROWS = 15;

    /* プリセットに保存する項目（label 以外。書き出し時の並び順になる）
       Preset fields other than label (the array order is the export order) */
    var PRESET_KEYS = [
        "useFilename", "prefixSeparator", "prefix", "nameStyleKey",
        "separator", "formatKey", "start", "increment", "suffix"
    ];

    /* 内蔵プリセット定義（label はローカライズしない固定文字列）/ Built-in presets (label is a fixed, non-localized string) */
    var BUILTIN_NAMING_PRESETS = [
        {
            label: "ファイル名+連番3",
            useFilename: true,
            prefixSeparator: "-",
            prefix: "",
            nameStyleKey: "none",
            separator: "",
            formatKey: "numeric",
            start: "001",
            increment: "1",
            suffix: ""
        },
        {
            label: "アートボード名と連番",
            useFilename: false,
            prefixSeparator: "-",
            prefix: "",
            nameStyleKey: "name",
            separator: "-",
            formatKey: "numeric",
            start: "1",
            increment: "1",
            suffix: ""
        }
    ];

    // =========================================
    // 補助関数群 / Helper functions
    // =========================================

    /**
     * 数値をゼロ埋めした文字列にする
     * @param {number} value - 対象の数値
     * @param {number} digits - 最低桁数
     * @returns {string} ゼロ埋めした文字列
     */
    function padNumberWithZeros(value, digits) {
        var text = String(value);
        while (text.length < digits) text = "0" + text;
        return text;
    }

    /**
     * 0始まりインデックスからアルファベットラベルを生成する（A, B, ... Z, AA, AB, ...）
     * @param {number} index - 0始まりのインデックス
     * @param {boolean} useLowercase - 小文字で返すかどうか
     * @returns {string} アルファベットラベル
     */
    function buildAlphaLabel(index, useLowercase) {
        var label = "";
        while (index >= 0) {
            label = String.fromCharCode((index % 26) + 65) + label;
            index = Math.floor(index / 26) - 1;
        }
        return useLowercase ? label.toLowerCase() : label;
    }

    /**
     * アルファベット文字列を1始まりのインデックスに変換する
     * @param {string} alphaText - "A" や "ab" などのアルファベット文字列
     * @returns {number} 1始まりのインデックス（アルファベット以外を含む場合は NaN）
     */
    function getIndexFromAlphaLabel(alphaText) {
        var upperText = alphaText.toUpperCase();
        var total = 0;
        for (var i = 0; i < upperText.length; i++) {
            var charCode = upperText.charCodeAt(i);
            if (charCode < 65 || charCode > 90) return NaN;
            total = total * 26 + (charCode - 64);
        }
        return total;
    }

    /**
     * 前後の空白を取り除く（ES3 には String#trim がないため）
     * @param {string} text - 対象の文字列
     * @returns {string} 前後の空白を除いた文字列
     */
    function trimText(text) {
        return String(text).replace(/^\s+|\s+$/g, "");
    }

    /**
     * 数字だけで構成された文字列を整数に変換する（"1あ" のような入力を弾く）
     * @param {string} text - 入力文字列（前後の空白は除去済みであること）
     * @returns {number} 変換した整数、数字以外を含む場合は NaN
     */
    function parseDigitsOnly(text) {
        if (!/^\d+$/.test(text)) return NaN;
        return parseInt(text, 10);
    }

    /**
     * 開始番号の入力文字列からゼロ埋め桁数を求める
     * @param {string} startText - 開始番号の入力文字列（前後の空白は除去済みであること）
     * @returns {number} 数字のみなら文字数、それ以外は 0
     */
    function getPadDigitsFromStartText(startText) {
        return isNaN(parseDigitsOnly(startText)) ? 0 : startText.length;
    }

    /**
     * キー配列から一致するインデックスを返す（ES3 には Array#indexOf がないため）
     * @param {string[]} keys - 検索対象のキー配列
     * @param {string} targetKey - 探すキー
     * @returns {number} 見つかったインデックス、なければ -1
     */
    function findKeyIndex(keys, targetKey) {
        for (var i = 0; i < keys.length; i++) if (keys[i] === targetKey) return i;
        return -1;
    }

    /**
     * ドロップダウンの選択位置に対応するキーを取得する
     * @param {DropDownList} dropdown - 対象のドロップダウン
     * @param {string[]} keys - 表示順に並んだキー配列
     * @returns {string} 選択中のキー
     */
    function getSelectedKey(dropdown, keys) {
        return keys[dropdown.selection.index];
    }

    /**
     * 区切り文字ラジオ群から選択中の区切り文字を取得する
     * @param {Array<RadioButton>} separatorRadios - 区切り文字のラジオボタン群
     * @returns {string} 選択中の区切り文字（未選択なら空文字）
     */
    function getSelectedSeparator(separatorRadios) {
        for (var i = 0; i < separatorRadios.length; i++) {
            if (separatorRadios[i].value) return separatorRadios[i]._separator;
        }
        return "";
    }

    /**
     * 区切り文字の値からラジオボタンの選択状態を設定する
     * @param {Array<RadioButton>} separatorRadios - 区切り文字のラジオボタン群
     * @param {string} separatorValue - 選択したい区切り文字
     * @returns {void}
     */
    function selectSeparatorRadio(separatorRadios, separatorValue) {
        for (var i = 0; i < separatorRadios.length; i++) {
            separatorRadios[i].value = (separatorRadios[i]._separator === separatorValue);
        }
    }

    /**
     * 複数のコントロールに同じイベントハンドラを割り当てる
     * @param {Array<Object>} controls - ScriptUI コントロールの配列
     * @param {string} eventName - "onClick" などのイベント名
     * @param {function} handler - 割り当てるハンドラ
     * @returns {void}
     */
    function bindEventToAll(controls, eventName, handler) {
        for (var i = 0; i < controls.length; i++) controls[i][eventName] = handler;
    }

    /**
     * 複数のコントロールの活性状態をまとめて切り替える
     * @param {Array<Object>} controls - ScriptUI コントロールの配列
     * @param {boolean} isEnabled - 有効にするかどうか
     * @returns {void}
     */
    function setControlsEnabled(controls, isEnabled) {
        for (var i = 0; i < controls.length; i++) controls[i].enabled = isEnabled;
    }

    // =========================================
    // 名前の組み立て / Name building
    // =========================================

    /**
     * アートボード名スタイルに応じた名称部分を組み立てる
     * @param {number} artboardIndex - 0始まりのアートボード番号
     * @param {string} nameStyleKey - ARTBOARD_NAME_STYLE_KEYS のいずれか
     * @param {string[]} originalNames - 変更前のアートボード名
     * @returns {string} 名称部分（スタイルが "none" なら空文字）
     */
    function buildNameSegment(artboardIndex, nameStyleKey, originalNames) {
        var positionNumberText = (artboardIndex + 1).toString();
        var originalName = originalNames[artboardIndex];
        switch (nameStyleKey) {
            case "number": return positionNumberText;
            case "name": return originalName;
            case "numberDashName": return positionNumberText + "-" + originalName;
            case "numberUnderscoreName": return positionNumberText + "_" + originalName;
            default: return "";
        }
    }

    /**
     * リネーム設定一式
     * @typedef {Object} RenameContext
     * @property {boolean} hasNumber - 連番を付けるかどうか
     * @property {boolean} isNumeric - 連番が数字かどうか
     * @property {boolean} isLowerAlpha - 連番が小文字アルファベットかどうか
     * @property {number} startValue - 連番の開始値（アルファベットは1始まりインデックス）
     * @property {number} incrementValue - 連番の増分（数字のときのみ使用）
     * @property {number} padDigits - ゼロ埋め桁数
     * @property {string} nameStyleKey - アートボード名スタイルキー
     * @property {string} composedPrefix - 組み立て済みの接頭辞
     * @property {string} numberSeparator - 名称と連番の間に入れる区切り文字
     * @property {string} suffixText - 接尾辞の文字列
     * @property {boolean} isValid - 入力値が妥当かどうか
     */

    /**
     * 指定インデックスのアートボード名を組み立てる
     * @param {number} artboardIndex - 0始まりのアートボード番号
     * @param {RenameContext} renameContext - リネーム設定
     * @param {string[]} originalNames - 変更前のアートボード名
     * @returns {string} 新しいアートボード名
     */
    function buildArtboardName(artboardIndex, renameContext, originalNames) {
        var nameSegment = buildNameSegment(artboardIndex, renameContext.nameStyleKey, originalNames);
        if (!renameContext.hasNumber) {
            return renameContext.composedPrefix + nameSegment + renameContext.suffixText;
        }
        var offset = renameContext.isNumeric ? artboardIndex * renameContext.incrementValue : artboardIndex;
        var numberValue = renameContext.startValue + offset;
        var numberLabel = renameContext.isNumeric ?
            padNumberWithZeros(numberValue, renameContext.padDigits) :
            buildAlphaLabel(numberValue - 1, renameContext.isLowerAlpha);
        return renameContext.composedPrefix + nameSegment + renameContext.numberSeparator + numberLabel + renameContext.suffixText;
    }

    // =========================================
    // プリセット書き出し / Preset export
    // =========================================

    /**
     * JSリテラルに埋め込めるよう文字列をエスケープする（\ と " のみ。改行はテキストフィールドに入らない想定）
     * @param {string} text - エスケープ対象の文字列
     * @returns {string} エスケープ済みの文字列
     */
    function escapeForJSLiteral(text) {
        return String(text).replace(/\\/g, "\\\\").replace(/"/g, '\\"');
    }

    /**
     * プリセット設定を BUILTIN_NAMING_PRESETS へ直接貼り付けられるテキストに変換する
     * @param {string} presetLabel - プリセット名
     * @param {Object} settings - collectCurrentSettings() が返す設定
     * @returns {string} オブジェクトリテラル形式の文字列
     */
    function serializePreset(presetLabel, settings) {
        var fields = ['label: "' + escapeForJSLiteral(presetLabel) + '"'];
        for (var i = 0; i < PRESET_KEYS.length; i++) {
            var presetKey = PRESET_KEYS[i];
            var value = settings[presetKey];
            var literal = (typeof value === "boolean") ? String(value) : '"' + escapeForJSLiteral(value) + '"';
            fields.push(presetKey + ": " + literal);
        }
        return "{ " + fields.join(", ") + " }";
    }

    /**
     * プリセットをテキストファイルに書き出す
     * @param {Object} settings - collectCurrentSettings() が返す設定
     * @returns {void}
     */
    function exportPresetToFile(settings) {
        var presetFile = File.saveDialog(getLabel("preset", "savePrompt"), "*.txt");
        if (!presetFile) return;
        if (presetFile.name.indexOf(".") === -1) presetFile = new File(presetFile.fsName + ".txt");

        var presetLabel = decodeURIComponent(presetFile.name.replace(/\.txt$/i, "")); // 日本語ファイル名もOK / Handles JA filenames
        presetFile.encoding = "UTF-8"; /* 日本語ラベルの文字化けを防ぐ / Keep JA labels readable */
        if (!presetFile.open("w")) {
            alert(getLabel("alert", "exportFailed"));
            return;
        }

        var didWrite = presetFile.write(serializePreset(presetLabel, settings));
        presetFile.close();
        if (!didWrite) {
            alert(getLabel("alert", "exportError") + presetFile.error);
            return;
        }
        alert(getLabel("alert", "exportSuccess") + presetFile.fsName);
    }

    // =========================================
    // UI構築 / UI builders
    // =========================================

    /**
     * 「区切り文字：なし / - / _」のラジオ行を構築する
     * @param {Object} parentPanel - 追加先のパネル
     * @returns {Array<RadioButton>} 区切り文字のラジオボタン群（各ラジオが _separator に値を持つ）
     */
    function buildSeparatorRadioRow(parentPanel) {
        var separatorGroup = parentPanel.add("group");
        setupRow(separatorGroup);
        separatorGroup.add("statictext", undefined, getFieldLabel("separator"));

        var separatorRadios = [];
        for (var i = 0; i < SEPARATOR_VALUES.length; i++) {
            var separatorValue = SEPARATOR_VALUES[i];
            var radioLabel = (separatorValue === "") ? getLabel("radio", "separatorNone") : separatorValue;
            var radio = separatorGroup.add("radiobutton", undefined, radioLabel);
            /* ラジオ配列の順序に依存せず値を引けるよう、ラジオ自身に持たせる / Keep the value on the radio so callers don't depend on order */
            radio._separator = separatorValue;
            separatorRadios.push(radio);
        }
        separatorRadios[0].value = true;
        return separatorRadios;
    }

    /**
     * 「ラベル：入力欄」の1行を構築する
     * @param {Object} parentPanel - 追加先のパネル
     * @param {string} labelKey - LABELS.fieldLabel のキー
     * @param {string} initialText - 入力欄の初期値
     * @param {number} charWidth - 入力欄の幅（文字数）
     * @returns {EditText} 追加した入力欄
     */
    function buildLabeledInput(parentPanel, labelKey, initialText, charWidth) {
        var inputRowGroup = parentPanel.add("group");
        setupRow(inputRowGroup);
        inputRowGroup.add("statictext", undefined, getFieldLabel(labelKey));
        var textInput = inputRowGroup.add("edittext", undefined, initialText);
        textInput.characters = charWidth;
        return textInput;
    }

    /**
     * プリセット選択行を構築する
     * @param {Object} parentGroup - 追加先のグループ
     * @returns {Object} presetDropdown / exportPresetButton を持つオブジェクト
     */
    function buildPresetRow(parentGroup) {
        var presetRowGroup = parentGroup.add("group");
        setupRow(presetRowGroup, "left");

        var presetItemLabels = [getLabel("preset", "none")];
        for (var i = 0; i < BUILTIN_NAMING_PRESETS.length; i++) {
            presetItemLabels.push(BUILTIN_NAMING_PRESETS[i].label);
        }

        var presetDropdown = presetRowGroup.add("dropdownlist", undefined, presetItemLabels);
        presetDropdown.selection = 0;

        /* ボタンは行幅いっぱいに広げない / Keep the button at its natural width */
        var exportPresetButton = presetRowGroup.add("button", undefined, getLabel("button", "exportPreset"));
        exportPresetButton.alignment = "left";

        return {
            presetDropdown: presetDropdown,
            exportPresetButton: exportPresetButton
        };
    }

    /**
     * 接頭辞パネルを構築する
     * @param {Object} parentGroup - 追加先のグループ
     * @returns {Object} useFilenameRadios / prefixSeparatorRadios / prefixTextInput を持つオブジェクト
     */
    function buildPrefixPanel(parentGroup) {
        var prefixPanel = addPanel(parentGroup, getLabel("panel", "prefix"), FIELD_SPACING);

        var useFilenameGroup = prefixPanel.add("group");
        setupRow(useFilenameGroup);
        useFilenameGroup.add("statictext", undefined, getFieldLabel("fileName"));
        var useFilenameRadios = [
            useFilenameGroup.add("radiobutton", undefined, getLabel("radio", "useFileNo")),
            useFilenameGroup.add("radiobutton", undefined, getLabel("radio", "useFileYes"))
        ];
        useFilenameRadios[0].value = true;

        var prefixSeparatorRadios = buildSeparatorRadioRow(prefixPanel);
        var prefixTextInput = buildLabeledInput(prefixPanel, "string", "", 16);

        return {
            useFilenameRadios: useFilenameRadios,
            prefixSeparatorRadios: prefixSeparatorRadios,
            prefixTextInput: prefixTextInput
        };
    }

    /**
     * アートボード名スタイルのパネルを構築する
     * @param {Object} parentGroup - 追加先のグループ
     * @returns {DropDownList} アートボード名スタイルのドロップダウン
     */
    function buildNameStylePanel(parentGroup) {
        var nameStylePanel = addPanel(parentGroup, getLabel("panel", "name"), FIELD_SPACING);
        var nameStyleDropdown = nameStylePanel.add("dropdownlist", undefined, toLabelList("nameStyle", ARTBOARD_NAME_STYLE_KEYS));
        nameStyleDropdown.selection = findKeyIndex(ARTBOARD_NAME_STYLE_KEYS, "none");
        /* ドロップダウンはパネル幅いっぱいに広げない / Keep the dropdown at its natural width */
        nameStyleDropdown.alignment = "left";
        return nameStyleDropdown;
    }

    /**
     * 接尾辞パネル（区切り文字・連番・接尾辞文字列）を構築する
     * @param {Object} parentGroup - 追加先のグループ
     * @returns {Object} numberSeparatorRadios / numberingFormatDropdown / startValueInput / incrementInput / suffixTextInput を持つオブジェクト
     */
    function buildSuffixPanel(parentGroup) {
        var suffixPanel = addPanel(parentGroup, getLabel("panel", "suffix"), FIELD_SPACING);

        var numberSeparatorRadios = buildSeparatorRadioRow(suffixPanel);

        var numberingFormatGroup = suffixPanel.add("group");
        setupRow(numberingFormatGroup);
        numberingFormatGroup.add("statictext", undefined, getFieldLabel("format"));
        var numberingFormatDropdown = numberingFormatGroup.add("dropdownlist", undefined, toLabelList("numberingFormat", NUMBERING_FORMAT_KEYS));
        numberingFormatDropdown.selection = findKeyIndex(NUMBERING_FORMAT_KEYS, "numeric");

        var startValueInput = buildLabeledInput(suffixPanel, "startNumber", DEFAULT_START_VALUES.numeric, 5);
        var incrementInput = buildLabeledInput(suffixPanel, "increment", "1", 5);
        var suffixTextInput = buildLabeledInput(suffixPanel, "string", "", 16);

        return {
            numberSeparatorRadios: numberSeparatorRadios,
            numberingFormatDropdown: numberingFormatDropdown,
            startValueInput: startValueInput,
            incrementInput: incrementInput,
            suffixTextInput: suffixTextInput
        };
    }

    /**
     * プレビューパネルを構築する
     * @param {Object} parentGroup - 追加先のグループ
     * @returns {ListBox} プレビュー用のリストボックス
     */
    function buildPreviewPanel(parentGroup) {
        var previewPanel = addPanel(parentGroup, getLabel("panel", "preview"));
        previewPanel.preferredSize.width = 250;
        previewPanel.preferredSize.height = 380;

        var previewListBox = previewPanel.add("listbox", undefined, [], {
            multiselect: false,
            numberOfColumns: 1,
            showHeaders: false
        });
        /* リストだけはパネルいっぱいに広げる / The list is the one control that should fill the panel */
        previewListBox.alignment = ["fill", "fill"];
        return previewListBox;
    }

    /**
     * ダイアログ下部のボタン行を構築する（左にキャンセル、右にOK）
     * @param {Window} parentWindow - 追加先のダイアログ
     * @returns {Button} OKボタン
     */
    function buildDialogButtonRow(parentWindow) {
        var dialogButtonRow = parentWindow.add("group");
        dialogButtonRow.orientation = "row";
        dialogButtonRow.alignChildren = ["fill", "center"];
        dialogButtonRow.spacing = 0;

        /* alignChildren の "fill" がボタン自体を伸ばさないよう、それぞれグループで包む
           Wrap each button in a group so the row's "fill" does not stretch the button */
        var cancelButtonGroup = dialogButtonRow.add("group");
        setupRow(cancelButtonGroup, "left");
        cancelButtonGroup.add("button", undefined, getLabel("button", "cancel"), { name: "cancel" });

        var buttonSpacer = dialogButtonRow.add("group");
        buttonSpacer.alignment = ["fill", "fill"];
        buttonSpacer.minimumSize.width = 50;

        var okButtonGroup = dialogButtonRow.add("group");
        setupRow(okButtonGroup, "right");
        return okButtonGroup.add("button", undefined, getLabel("button", "ok"), { name: "ok" });
    }

    /**
     * ダイアログ本体とすべてのコントロールを構築する
     * @typedef {Object} DialogUI
     * @property {Window} dialog - ダイアログ本体
     * @property {Object} preset - プリセット行のコントロール
     * @property {Object} prefix - 接頭辞パネルのコントロール
     * @property {DropDownList} nameStyleDropdown - アートボード名スタイルのドロップダウン
     * @property {Object} suffix - 接尾辞パネルのコントロール
     * @property {ListBox} previewListBox - プレビュー用のリストボックス
     * @property {Button} okButton - OKボタン
     *
     * @returns {DialogUI} ダイアログとコントロール一式
     */
    function buildRenameDialog() {
        var renameDialog = new Window("dialog", getLabel("dialog", "title") + " " + SCRIPT_VERSION);
        setupWindow(renameDialog);

        var dialogBodyRow = renameDialog.add("group");
        dialogBodyRow.orientation = "row";
        dialogBodyRow.alignChildren = ["top", "fill"];
        dialogBodyRow.spacing = COLUMN_SPACING;

        var settingsColumnGroup = dialogBodyRow.add("group");
        settingsColumnGroup.orientation = "column";
        settingsColumnGroup.alignChildren = ["fill", "top"];
        settingsColumnGroup.spacing = PANEL_SPACING;

        return {
            dialog: renameDialog,
            preset: buildPresetRow(settingsColumnGroup),
            prefix: buildPrefixPanel(settingsColumnGroup),
            nameStyleDropdown: buildNameStylePanel(settingsColumnGroup),
            suffix: buildSuffixPanel(settingsColumnGroup),
            previewListBox: buildPreviewPanel(dialogBodyRow),
            okButton: buildDialogButtonRow(renameDialog)
        };
    }

    // =========================================
    // ダイアログ制御 / Dialog controller
    // =========================================

    /**
     * ダイアログの状態を読み書きする操作一式を作る
     * @param {DialogUI} dialogUI - 構築済みのダイアログとコントロール
     * @param {Document} activeDoc - 対象のドキュメント
     * @returns {Object} プレビュー更新・プリセット反映・リネーム実行などの関数を持つオブジェクト
     */
    function createRenameController(dialogUI, activeDoc) {
        var prefix = dialogUI.prefix;
        var suffix = dialogUI.suffix;
        var artboards = activeDoc.artboards;
        var documentBaseName = activeDoc.name.replace(/\.[^\.]+$/, "");

        /* 変更前のアートボード名を控えておく（何度リネームしても元の名前から組み立てられる）
           Snapshot the original names so every rename starts from the same baseline */
        var originalArtboardNames = [];
        for (var i = 0; i < artboards.length; i++) originalArtboardNames.push(artboards[i].name);
        var artboardCount = originalArtboardNames.length;

        /**
         * 現在のUI状態から接頭辞文字列を組み立てる
         * @returns {string} 接頭辞（ファイル名参照時はファイル名＋区切り文字を前置）
         */
        function composePrefixText() {
            var prefixText = prefix.prefixTextInput.text;
            if (prefix.useFilenameRadios[1].value) {
                prefixText = documentBaseName + getSelectedSeparator(prefix.prefixSeparatorRadios) + prefixText;
            }
            return prefixText;
        }

        /**
         * 現在のUI状態からリネーム設定を組み立てる
         * @returns {RenameContext} リネーム設定
         */
        function buildRenameContext() {
            var formatKey = getSelectedKey(suffix.numberingFormatDropdown, NUMBERING_FORMAT_KEYS);
            var hasNumber = (formatKey !== "none");
            var isNumeric = (formatKey === "numeric");
            var startText = trimText(suffix.startValueInput.text);
            var startValue = !hasNumber ? 0 : (isNumeric ? parseDigitsOnly(startText) : getIndexFromAlphaLabel(startText));
            var incrementValue = parseDigitsOnly(trimText(suffix.incrementInput.text));

            var isStartValid = !isNaN(startValue) && startValue > 0;
            var isIncrementValid = !isNumeric || (!isNaN(incrementValue) && incrementValue > 0);

            return {
                hasNumber: hasNumber,
                isNumeric: isNumeric,
                isLowerAlpha: (formatKey === "alphaLower"),
                startValue: startValue,
                incrementValue: incrementValue,
                padDigits: getPadDigitsFromStartText(startText),
                nameStyleKey: getSelectedKey(dialogUI.nameStyleDropdown, ARTBOARD_NAME_STYLE_KEYS),
                composedPrefix: composePrefixText(),
                numberSeparator: getSelectedSeparator(suffix.numberSeparatorRadios),
                suffixText: suffix.suffixTextInput.text,
                isValid: !hasNumber || (isStartValid && isIncrementValid)
            };
        }

        /**
         * 入力を検証したうえでリネーム設定を返す（不正ならアラートを表示）
         * @returns {RenameContext|null} 妥当なリネーム設定、不正な場合は null
         */
        function validateAndBuildContext() {
            var renameContext = buildRenameContext();
            if (!renameContext.isValid) {
                alert(getLabel("alert", "invalidInput"), getLabel("alert", "title"));
                return null;
            }
            /* 連番なしの場合のみ空名になり得る（連番ありなら連番ラベルが必ず入る）
               Empty name only possible when there is no numbering (the number label is always non-empty otherwise) */
            if (!renameContext.hasNumber) {
                for (var i = 0; i < artboardCount; i++) {
                    if (buildArtboardName(i, renameContext, originalArtboardNames) === "") {
                        alert(getLabel("alert", "emptyName"), getLabel("alert", "title"));
                        return null;
                    }
                }
            }
            return renameContext;
        }

        /**
         * プレビューリストを現在の設定で更新する
         * @returns {void}
         */
        function refreshPreviewList() {
            var renameContext = buildRenameContext();
            dialogUI.previewListBox.removeAll();

            if (!renameContext.isValid) {
                dialogUI.previewListBox.add("item", getLabel("preview", "invalid"));
                return;
            }

            var previewCount = Math.min(PREVIEW_MAX_ROWS, artboardCount);
            for (var i = 0; i < previewCount; i++) {
                dialogUI.previewListBox.add("item", buildArtboardName(i, renameContext, originalArtboardNames));
            }
            if (artboardCount > PREVIEW_MAX_ROWS) {
                dialogUI.previewListBox.add("item", getLabel("preview", "truncated"));
            }
        }

        /**
         * 現在のUI設定をプリセット形式のオブジェクトとして取得する
         * @returns {Object} プリセット1件分の設定（PRESET_KEYS の項目を持つ）
         */
        function collectCurrentSettings() {
            return {
                useFilename: prefix.useFilenameRadios[1].value,
                prefixSeparator: getSelectedSeparator(prefix.prefixSeparatorRadios),
                prefix: prefix.prefixTextInput.text,
                nameStyleKey: getSelectedKey(dialogUI.nameStyleDropdown, ARTBOARD_NAME_STYLE_KEYS),
                separator: getSelectedSeparator(suffix.numberSeparatorRadios),
                formatKey: getSelectedKey(suffix.numberingFormatDropdown, NUMBERING_FORMAT_KEYS),
                start: suffix.startValueInput.text,
                increment: suffix.incrementInput.text,
                suffix: suffix.suffixTextInput.text
            };
        }

        /**
         * ファイル名参照の有無に合わせて、接頭辞の区切り文字ラジオの活性状態をそろえる
         * @returns {void}
         */
        function syncPrefixSeparatorControls() {
            setControlsEnabled(prefix.prefixSeparatorRadios, prefix.useFilenameRadios[1].value);
        }

        /**
         * 連番形式に合わせて、区切り文字・開始番号・増分の活性状態をそろえる
         * @param {boolean} resetStartValue - 開始番号を形式ごとの既定値に戻すかどうか
         * @returns {void}
         */
        function syncNumberingControls(resetStartValue) {
            var formatKey = getSelectedKey(suffix.numberingFormatDropdown, NUMBERING_FORMAT_KEYS);
            var hasNumber = (formatKey !== "none");
            /* 区切り文字は連番の直前にしか入らないため、連番なしのときは操作させない
               The separator only precedes a number, so disable it when there is none */
            setControlsEnabled(suffix.numberSeparatorRadios, hasNumber);
            suffix.startValueInput.enabled = hasNumber;
            suffix.incrementInput.enabled = (formatKey === "numeric");
            if (resetStartValue && hasNumber) suffix.startValueInput.text = DEFAULT_START_VALUES[formatKey];
        }

        /**
         * プリセットの値をUIに反映する
         * @param {Object} namingPreset - BUILTIN_NAMING_PRESETS の1件
         * @returns {void}
         */
        function applyPresetToControls(namingPreset) {
            prefix.useFilenameRadios[0].value = !namingPreset.useFilename;
            prefix.useFilenameRadios[1].value = namingPreset.useFilename;
            selectSeparatorRadio(prefix.prefixSeparatorRadios, namingPreset.prefixSeparator);
            prefix.prefixTextInput.text = namingPreset.prefix;

            var nameStyleIndex = findKeyIndex(ARTBOARD_NAME_STYLE_KEYS, namingPreset.nameStyleKey);
            if (nameStyleIndex >= 0) dialogUI.nameStyleDropdown.selection = nameStyleIndex;

            selectSeparatorRadio(suffix.numberSeparatorRadios, namingPreset.separator);
            var formatIndex = findKeyIndex(NUMBERING_FORMAT_KEYS, namingPreset.formatKey);
            if (formatIndex >= 0) suffix.numberingFormatDropdown.selection = formatIndex;

            suffix.startValueInput.text = namingPreset.start;
            suffix.incrementInput.text = namingPreset.increment;
            suffix.suffixTextInput.text = namingPreset.suffix;

            /* 開始番号はプリセットの値を使うため、既定値には戻さない / Keep the preset's start value instead of the per-format default */
            syncPrefixSeparatorControls();
            syncNumberingControls(false);
            refreshPreviewList();
        }

        /**
         * 検証に通ればすべてのアートボードに新しい名前を適用する
         * @returns {boolean} 適用できたかどうか（検証エラー・適用エラーなら false）
         */
        function renameArtboards() {
            var renameContext = validateAndBuildContext();
            if (!renameContext) return false;
            try {
                for (var i = 0; i < artboardCount; i++) {
                    artboards[i].name = buildArtboardName(i, renameContext, originalArtboardNames);
                }
            } catch (e) {
                /* 名前がIllustratorに拒否された場合。途中まで適用された状態で確定させない
                   Illustrator rejected a name; do not commit a partially applied rename */
                alert(getLabel("alert", "generalError") + e.message);
                return false;
            }
            return true;
        }

        return {
            refreshPreviewList: refreshPreviewList,
            collectCurrentSettings: collectCurrentSettings,
            applyPresetToControls: applyPresetToControls,
            syncPrefixSeparatorControls: syncPrefixSeparatorControls,
            syncNumberingControls: syncNumberingControls,
            renameArtboards: renameArtboards
        };
    }

    // =========================================
    // イベント配線 / Event wiring
    // =========================================

    /**
     * ダイアログのコントロールにイベントハンドラを割り当てる
     * @param {DialogUI} dialogUI - 構築済みのダイアログとコントロール
     * @param {Object} renameController - createRenameController() が返す操作一式
     * @returns {void}
     */
    function wireDialogEvents(dialogUI, renameController) {
        var prefix = dialogUI.prefix;
        var suffix = dialogUI.suffix;
        var refreshPreviewList = renameController.refreshPreviewList;

        /* プリセット選択（先頭は「(未選択)」なので読み飛ばす）/ Preset selection (index 0 is the "(None)" entry) */
        dialogUI.preset.presetDropdown.onChange = function () {
            var presetIndex = dialogUI.preset.presetDropdown.selection.index;
            if (presetIndex > 0) renameController.applyPresetToControls(BUILTIN_NAMING_PRESETS[presetIndex - 1]);
        };

        /* プリセット書き出し / Export preset */
        dialogUI.preset.exportPresetButton.onClick = function () {
            exportPresetToFile(renameController.collectCurrentSettings());
        };

        /* ファイル名参照の切り替え：区切り文字ラジオの有効化を連動 / Toggle separator radios with filename usage */
        prefix.useFilenameRadios[0].onClick = prefix.useFilenameRadios[1].onClick = function () {
            renameController.syncPrefixSeparatorControls();
            refreshPreviewList();
        };

        /* 連番形式の切り替え：関連コントロールの有効化と開始番号の既定値を更新 / On format change: sync related controls and reset the start value */
        suffix.numberingFormatDropdown.onChange = function () {
            renameController.syncNumberingControls(true);
            refreshPreviewList();
        };

        /* 入力変更時に即時プレビュー / Live preview on input change */
        bindEventToAll(prefix.prefixSeparatorRadios, "onClick", refreshPreviewList);
        bindEventToAll(suffix.numberSeparatorRadios, "onClick", refreshPreviewList);
        prefix.prefixTextInput.onChanging = refreshPreviewList;
        suffix.startValueInput.onChanging = refreshPreviewList;
        suffix.incrementInput.onChanging = refreshPreviewList;
        suffix.suffixTextInput.onChanging = refreshPreviewList;
        dialogUI.nameStyleDropdown.onChange = refreshPreviewList;

        /* OKボタン：リネームに成功したときだけ閉じる / OK: close only when the rename succeeded */
        dialogUI.okButton.onClick = function () {
            if (renameController.renameArtboards()) dialogUI.dialog.close();
        };
    }

    // =========================================
    // メイン処理 / Main entry
    // =========================================

    /**
     * ダイアログを表示してアートボード名を一括変更する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) return;

        var dialogUI = buildRenameDialog();
        var renameController = createRenameController(dialogUI, app.activeDocument);
        wireDialogEvents(dialogUI, renameController);

        /* 初期状態を反映 / Apply the initial state */
        renameController.syncPrefixSeparatorControls();
        renameController.syncNumberingControls(true);
        renameController.refreshPreviewList();

        dialogUI.dialog.show();
    }

    main();

})();
