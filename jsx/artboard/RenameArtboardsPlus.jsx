#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### スクリプト名 / Script Name

RenameArtboardsPlus.jsx

### 概要 / Overview

- アートボード名を一括で柔軟に変更できるIllustrator用スクリプトです。
- 接頭辞・接尾辞・ファイル名・連番・元アートボード名を自由に組み合わせて、高度な命名ルールを実現できます。
- A script to flexibly batch rename artboards in Illustrator.
- Combine prefixes, suffixes, file name, numbers, and original artboard names to create advanced naming rules.

### 主な機能 / Main Features

- 接頭辞・接尾辞・元アートボード名の有無を選択可能 / Freely choose whether to include prefix, suffix, or original artboard name
- ファイル名参照、区切り文字設定、プリセット保存／選択対応 / File name reference, separator settings, preset save/load
- 連番形式（数字、大文字アルファベット、小文字アルファベット）選択 / Numbering formats: number, uppercase alphabet, lowercase alphabet
- 開始番号に基づく桁数（ゼロ埋め）の自動判定 / Auto-detect padding digits based on start number
- プレビュー機能で変更結果を事前確認 / Preview before applying
- 最大15件のプレビュー表示 / Up to 15 preview items
- ExtendScript（ES3）互換 / Compatible with ExtendScript (ES3)

### 更新履歴 / Update History

- v1.0 (20250420): 初期バージョン / Initial version
- v1.1 (20250430): 開始番号から桁数自動判定追加、プリセットラベル簡素化、ES3対応強化 / Auto-pad detection, simplified preset labels, ES3 hardening
- v1.2.0 (20260507): ローカライズ刷新（L/labelText/labelWithCount）、内部キー化で英語ロケール対応、連番形式に「なし」追加、適用ボタン削除、main()をUIビルダー/プリセットI/O/純粋関数に分割 / Localization overhaul (L/labelText/labelWithCount), key-based internals for EN locale support, "none" numbering format, removed Apply button, split main() into UI builders / preset I/O / pure helpers

*/

(function () {

    // =========================================
    // バージョンとローカライズ / Version & Localization
    // =========================================

    var SCRIPT_VERSION = "v1.2.0";

    /* 現在のロケール判定 / Detect current locale */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var lang = getCurrentLang();

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialogTitle: { ja: "アートボード名の一括設定", en: "Batch Rename Artboards" },
        prefixPanel: { ja: "接頭辞", en: "Prefix" },
        fileName: { ja: "ファイル名", en: "File Name" },
        useFileNo: { ja: "参照しない", en: "Do not use" },
        useFileYes: { ja: "参照する", en: "Use" },
        separator: { ja: "区切り文字", en: "Separator" },
        string: { ja: "文字列", en: "String" },
        namePanel: { ja: "アートボード名と番号", en: "Artboard Name & Number" },
        suffixPanel: { ja: "接尾辞", en: "Suffix" },
        format: { ja: "連番形式", en: "Numbering Format" },
        startNumber: { ja: "開始番号", en: "Start Number" },
        increment: { ja: "増分", en: "Increment" },
        preview: { ja: "プレビュー", en: "Preview" },
        cancel: { ja: "キャンセル", en: "Cancel" },
        ok: { ja: "OK", en: "OK" },
        savePreset: { ja: "プリセット書き出し", en: "Export Preset" },
        presetNone: { ja: "(未選択)", en: "(None)" },
        none: { ja: "なし", en: "None" },
        number: { ja: "番号", en: "Number" },
        name: { ja: "名称", en: "Name" },
        numberDashName: { ja: "番号-名称", en: "Number-Name" },
        numberUnderscoreName: { ja: "番号_名称", en: "Number_Name" },
        numeric: { ja: "数字", en: "Number" },
        alphaUpper: { ja: "アルファベット（大文字）", en: "Alphabet (Upper)" },
        alphaLower: { ja: "アルファベット（小文字）", en: "Alphabet (Lower)" },
        previewInvalid: { ja: "※数値が正しくありません", en: "※ Invalid number" },
        previewTruncated: { ja: "...（以下省略）", en: "... (more)" },
        presetSavePrompt: {
            ja: "プリセットを書き出す場所と名前を指定してください",
            en: "Specify the location and name to export the preset"
        },
        errorTitle: { ja: "エラー", en: "Error" },
        errorInput: {
            ja: "入力内容に誤りがあります。\n正しい数値を指定してください。",
            en: "There is an error in your input.\nPlease enter valid numbers."
        },
        exportSuccess: { ja: "プリセットを書き出しました：\n", en: "Preset exported:\n" },
        exportFailed: { ja: "ファイルを書き込めませんでした。", en: "Could not write the file." },
        exportError: { ja: "プリセットの保存に失敗しました：\n", en: "Failed to save preset:\n" },
        generalError: { ja: "エラーが発生しました：\n", en: "An error occurred:\n" }
    };

    /* ラベル取得 / Get label */
    function L(key) {
        return (LABELS[key] && LABELS[key][lang]) ? LABELS[key][lang] : key;
    }

    /* コロン付きラベル（日本語は全角、英語は半角）/ Label with colon (full-width JA, half-width EN) */
    function labelText(key) {
        return L(key) + (lang === 'ja' ? '：' : ':');
    }

    /* 件数付きラベル（日本語は全角括弧、英語は半角括弧）/ Label with count (full-width JA parentheses, half-width EN parentheses) */
    function labelWithCount(key, count) {
        if (lang === "ja") {
            return L(key) + "（" + count + "）";
        }
        return L(key) + " (" + count + ")";
    }

    // =========================================
    // 定数 / Constants
    // =========================================

    /* 連番形式キー（dropdown の表示順と対応、"none" は連番なし）/ Numbering format keys (matches dropdown order; "none" means no number) */
    var FORMAT_KEYS = ["none", "numeric", "alphaUpper", "alphaLower"];

    /* アートボード名スタイルキー（dropdown の表示順と対応）/ Artboard name style keys (matches dropdown order) */
    var NAME_STYLE_KEYS = ["none", "number", "name", "numberDashName", "numberUnderscoreName"];

    /* 連番形式変更時の開始値デフォルト / Default start value per format */
    var START_DEFAULTS = { numeric: "1", alphaUpper: "A", alphaLower: "a" };

    /* 内蔵プリセット定義（label はローカライズしない固定文字列）/ Built-in presets (label is a fixed, non-localized string) */
    var BUILTIN_PRESETS = [
        {
            label: "ファイル名+連番3",
            useFilename: true,
            prefixSeparator: "-",
            prefix: "",
            nameStyleKey: "none",
            separator: "",
            formatKey: "numeric",
            start: "01",
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

    /* 数字をゼロ埋め / Zero-pad a number */
    function padNumber(num, length) {
        var str = String(num);
        while (str.length < length) str = "0" + str;
        return str;
    }

    /* インデックスからアルファベットラベルを生成 / Generate alpha label from index */
    function getAlphaLabel(index, isLower) {
        var label = "";
        while (index >= 0) {
            label = String.fromCharCode((index % 26) + 65) + label;
            index = Math.floor(index / 26) - 1;
        }
        return isLower ? label.toLowerCase() : label;
    }

    /* アルファベット文字列を1始まりインデックスに変換 / Convert alpha string to 1-based index */
    function getAlphaIndex(str) {
        str = str.toUpperCase();
        var total = 0;
        for (var i = 0; i < str.length; i++) {
            var code = str.charCodeAt(i);
            if (code < 65 || code > 90) return NaN;
            total = total * 26 + (code - 64);
        }
        return total;
    }

    /* 開始値文字列の桁数を取得（数字のみ、ゼロ埋め判定用）/ Pad length from start string */
    function getPadLengthFromStart(str) {
        if (/^\d+$/.test(str)) return str.length;
        return 0;
    }

    /* キー配列から一致インデックスを返す / Find index of value in key array */
    function indexOfKey(arr, val) {
        for (var i = 0; i < arr.length; i++) if (arr[i] === val) return i;
        return -1;
    }

    /* ラジオボタン群から選択中の区切り文字を取得 / Get separator string from radio group */
    function getSeparator(radios) {
        return radios[1].value ? "-" : (radios[2].value ? "_" : "");
    }

    /* 区切り文字値からラジオボタンの状態を設定 / Set radio state from separator value */
    function setSeparatorRadios(radios, value) {
        radios[0].value = (value === "");
        radios[1].value = (value === "-");
        radios[2].value = (value === "_");
    }

    /* 配列要素にイベントハンドラを一括バインド / Bind handler to all elements */
    function bindAll(elements, eventName, handler) {
        for (var i = 0; i < elements.length; i++) {
            elements[i][eventName] = handler;
        }
    }

    /* 名称コンポーネントを構築（番号・名称・組み合わせ）/ Build name component by style key */
    function getNameComponent(index, nameStyleKey, originalNames) {
        var abNum = (index + 1).toString();
        var abName = originalNames[index];
        switch (nameStyleKey) {
            case "number": return abNum;
            case "name": return abName;
            case "numberDashName": return abNum + "-" + abName;
            case "numberUnderscoreName": return abNum + "_" + abName;
            default: return "";
        }
    }

    /* インデックスi番目のアートボード名を生成 / Build artboard name at index i */
    function buildArtboardName(i, context, originalNames) {
        var nameComponent = getNameComponent(i, context.nameStyleKey, originalNames);
        if (!context.hasNumber) {
            return context.composedPrefix + nameComponent + context.suffixText;
        }
        var num = context.startValue + (context.isNumeric ? i * context.increment : i);
        var label = context.isNumeric ? padNumber(num, context.padLength) : getAlphaLabel(num - 1, context.isLower);
        return context.composedPrefix + nameComponent + context.suffixSeparator + label + context.suffixText;
    }

    // =========================================
    // プリセット保存処理 / Preset save handler
    // =========================================

    /* プリセット設定を直書きできるテキスト形式に変換 / Serialize preset settings into a literal-style text */
    function serializePreset(label, s) {
        return '{ label: "' + label + '", ' +
            'useFilename: ' + (s.useFilename ? 'true' : 'false') + ', ' +
            'prefixSeparator: "' + s.prefixSeparator + '", ' +
            'prefix: "' + s.prefix + '", ' +
            'nameStyleKey: "' + s.nameStyleKey + '", ' +
            'separator: "' + s.separator + '", ' +
            'formatKey: "' + s.formatKey + '", ' +
            'start: "' + s.start + '", ' +
            'increment: "' + s.increment + '", ' +
            'suffix: "' + s.suffix + '" }';
    }

    /* プリセットをテキストファイルに書き出し / Export preset to text file */
    function savePresetToFile(settings) {
        var saveFile = File.saveDialog(L("presetSavePrompt"), "*.txt");
        if (!saveFile) return;
        if (saveFile.name.indexOf(".") === -1) saveFile = new File(saveFile.fsName + ".txt");

        var labelName = decodeURIComponent(saveFile.name.replace(/\.txt$/i, "")); // 日本語ファイル名もOK / Handles JA filenames
        var presetString = serializePreset(labelName, settings);

        try {
            if (saveFile.open("w")) {
                saveFile.write(presetString);
                saveFile.close();
                alert(L("exportSuccess") + saveFile.fsName);
            } else {
                alert(L("exportFailed"));
            }
        } catch (e) {
            alert(L("exportError") + e.message);
        }
    }

    // =========================================
    // UIサブセクション構築 / UI subsection builders
    // =========================================

    /* プリセット選択UIを構築 / Build preset selection UI */
    function createPresetSection(parent) {
        var presetGroup = parent.add("group");
        presetGroup.orientation = "row";
        presetGroup.alignChildren = "left";

        var presetItems = [L("presetNone")];
        for (var i = 0; i < BUILTIN_PRESETS.length; i++) {
            presetItems.push(BUILTIN_PRESETS[i].label);
        }

        var presetDropdown = presetGroup.add("dropdownlist", undefined, presetItems);
        presetDropdown.selection = 0;
        presetDropdown.enabled = true;

        var savePresetBtn = presetGroup.add("button", undefined, L("savePreset"));

        return {
            presetDropdown: presetDropdown,
            savePresetBtn: savePresetBtn
        };
    }

    /* 接頭辞パネルを構築 / Build prefix panel */
    function createPrefixPanel(parent) {
        var prefixPanel = parent.add("panel", undefined, L("prefixPanel"));
        prefixPanel.margins = [20, 20, 20, 10];
        prefixPanel.orientation = "column";
        prefixPanel.alignChildren = "left";

        var filenameGroup = prefixPanel.add("group");
        filenameGroup.add("statictext", undefined, labelText("fileName"));
        var useFilenameRadios = [
            filenameGroup.add("radiobutton", undefined, L("useFileNo")),
            filenameGroup.add("radiobutton", undefined, L("useFileYes"))
        ];
        useFilenameRadios[0].value = true;

        var prefixSeparatorGroup = prefixPanel.add("group");
        prefixSeparatorGroup.add("statictext", undefined, labelText("separator"));
        var prefixSeparatorRadios = [
            prefixSeparatorGroup.add("radiobutton", undefined, L("none")),
            prefixSeparatorGroup.add("radiobutton", undefined, "-"),
            prefixSeparatorGroup.add("radiobutton", undefined, "_")
        ];
        prefixSeparatorRadios[0].value = true;
        for (var i = 0; i < 3; i++) prefixSeparatorRadios[i].enabled = false;

        var prefixGroup = prefixPanel.add("group");
        prefixGroup.add("statictext", undefined, labelText("string"));
        var prefixInput = prefixGroup.add("edittext", undefined, "");
        prefixInput.characters = 16;

        return {
            useFilenameRadios: useFilenameRadios,
            prefixSeparatorRadios: prefixSeparatorRadios,
            prefixInput: prefixInput
        };
    }

    /* 名称スタイルパネルを構築 / Build name style panel */
    function createNamePanel(parent) {
        var namePanel = parent.add("panel", undefined, L("namePanel"));
        namePanel.margins = [20, 20, 20, 10];
        namePanel.orientation = "row";
        namePanel.alignChildren = "left";

        var nameStyleDropdown = namePanel.add("dropdownlist", undefined, [
            L("none"), L("number"), L("name"), L("numberDashName"), L("numberUnderscoreName")
        ]);
        nameStyleDropdown.selection = 0;

        return { nameStyleDropdown: nameStyleDropdown };
    }

    /* 接尾辞パネルを構築 / Build suffix panel */
    function createSuffixPanel(parent) {
        var suffixPanel = parent.add("panel", undefined, L("suffixPanel"));
        suffixPanel.margins = [20, 20, 20, 10];
        suffixPanel.orientation = "column";
        suffixPanel.alignChildren = "left";

        var suffixSeparatorGroup = suffixPanel.add("group");
        suffixSeparatorGroup.add("statictext", undefined, labelText("separator"));
        var suffixSeparatorRadios = [
            suffixSeparatorGroup.add("radiobutton", undefined, L("none")),
            suffixSeparatorGroup.add("radiobutton", undefined, "-"),
            suffixSeparatorGroup.add("radiobutton", undefined, "_")
        ];
        suffixSeparatorRadios[0].value = true;

        var formatGroup = suffixPanel.add("group");
        formatGroup.add("statictext", undefined, labelText("format"));
        var formatDropdown = formatGroup.add("dropdownlist", undefined, [
            L("none"), L("numeric"), L("alphaUpper"), L("alphaLower")
        ]);
        formatDropdown.selection = 1; // 既定は「数字」/ Default to numeric

        var startGroup = suffixPanel.add("group");
        startGroup.add("statictext", undefined, labelText("startNumber"));
        var startValueInput = startGroup.add("edittext", undefined, "1");
        startValueInput.characters = 5;

        var incrementGroup = suffixPanel.add("group");
        incrementGroup.add("statictext", undefined, labelText("increment"));
        var incrementInput = incrementGroup.add("edittext", undefined, "1");
        incrementInput.characters = 5;

        var suffixGroup = suffixPanel.add("group");
        suffixGroup.add("statictext", undefined, labelText("string"));
        var suffixInput = suffixGroup.add("edittext", undefined, "");
        suffixInput.characters = 16;

        return {
            suffixSeparatorRadios: suffixSeparatorRadios,
            formatDropdown: formatDropdown,
            startValueInput: startValueInput,
            incrementInput: incrementInput,
            suffixInput: suffixInput
        };
    }

    /* プレビューエリアを構築 / Build preview area */
    function createPreviewPanel(parent) {
        var previewPanel = parent.add("panel", undefined, L("preview"));
        previewPanel.alignChildren = "fill";
        previewPanel.margins = [10, 20, 10, 10];
        previewPanel.preferredSize.width = 250;
        previewPanel.preferredSize.height = 380;

        var previewList = previewPanel.add("listbox", undefined, [], {
            multiselect: false,
            numberOfColumns: 1,
            showHeaders: false
        });
        previewList.preferredSize.height = 320;

        return { previewList: previewList };
    }

    /* 下部ボタン行を構築 / Build bottom button row */
    function createButtonRow(parent) {
        var buttonRow = parent.add("group");
        buttonRow.orientation = "row";
        buttonRow.alignChildren = ["fill", "center"];
        buttonRow.margins = [0, 10, 0, 0];
        buttonRow.spacing = 0;

        var cancelButtonGroup = buttonRow.add("group");
        cancelButtonGroup.orientation = "row";
        cancelButtonGroup.alignChildren = "left";
        cancelButtonGroup.add("button", undefined, L("cancel"), { name: "cancel" });

        var spacer = buttonRow.add("group");
        spacer.alignment = ["fill", "fill"];
        spacer.minimumSize.width = 50;

        var confirmButtonGroup = buttonRow.add("group");
        confirmButtonGroup.orientation = "row";
        confirmButtonGroup.alignChildren = ["right", "center"];
        confirmButtonGroup.spacing = 10;
        var okBtn = confirmButtonGroup.add("button", undefined, L("ok"), { name: "ok" });

        return { okBtn: okBtn };
    }

    // =========================================
    // メイン処理 / Main entry
    // =========================================

    /* メイン関数：ダイアログを表示してアートボード名を一括変更 / Main: show dialog and batch rename artboards */
    function main() {
        if (app.documents.length === 0) return;

        var doc = app.activeDocument;
        var filename = doc.name.replace(/\.[^\.]+$/, "");
        var artboards = doc.artboards;
        var artboardCount = artboards.length;
        var originalNames = [];
        for (var i = 0; i < artboardCount; i++) originalNames.push(artboards[i].name);

        /* ダイアログ枠 / Dialog skeleton */
        var dialog = new Window("dialog", L("dialogTitle") + " " + SCRIPT_VERSION);
        dialog.alignChildren = "fill";

        var bodyRow = dialog.add("group");
        bodyRow.orientation = "row";
        bodyRow.alignChildren = ["top", "fill"];

        var inputGroup = bodyRow.add("group");
        inputGroup.spacing = 20;
        inputGroup.orientation = "column";
        inputGroup.alignChildren = "left";

        /* UIサブセクションを構築 / Build UI subsections */
        var presetUI = createPresetSection(inputGroup);
        var prefixUI = createPrefixPanel(inputGroup);
        var nameUI = createNamePanel(inputGroup);
        var suffixUI = createSuffixPanel(inputGroup);
        var previewUI = createPreviewPanel(bodyRow);
        var buttonUI = createButtonRow(dialog);

        /* 接頭辞文字列を組み立て / Compose prefix from current state */
        function composePrefix() {
            var base = prefixUI.prefixInput.text;
            if (prefixUI.useFilenameRadios[1].value) {
                base = filename + getSeparator(prefixUI.prefixSeparatorRadios) + base;
            }
            return base;
        }

        /* 現在のUI状態からリネーム用コンテキストを生成 / Build rename context from current UI state */
        function getRenameContext() {
            var formatKey = FORMAT_KEYS[suffixUI.formatDropdown.selection.index];
            var hasNumber = (formatKey !== "none");
            var isNumeric = (formatKey === "numeric");
            var rawStart = suffixUI.startValueInput.text;
            var startValue = !hasNumber ? 0 : (isNumeric ? parseInt(rawStart, 10) : getAlphaIndex(rawStart));
            var increment = parseInt(suffixUI.incrementInput.text, 10);
            var valid = !hasNumber || !(isNaN(startValue) || startValue <= 0 || (isNumeric && (isNaN(increment) || increment <= 0)));
            return {
                formatKey: formatKey,
                hasNumber: hasNumber,
                isNumeric: isNumeric,
                isLower: (formatKey === "alphaLower"),
                startValue: startValue,
                increment: increment,
                padLength: getPadLengthFromStart(rawStart),
                nameStyleKey: NAME_STYLE_KEYS[nameUI.nameStyleDropdown.selection.index],
                composedPrefix: composePrefix(),
                suffixSeparator: getSeparator(suffixUI.suffixSeparatorRadios),
                suffixText: suffixUI.suffixInput.text,
                valid: valid
            };
        }

        /* 入力検証＋コンテキスト取得（無効ならエラー表示）/ Validate and return context, alert if invalid */
        function validateContext() {
            var context = getRenameContext();
            if (!context.valid) {
                alert(L("errorInput"), L("errorTitle"));
                return null;
            }
            return context;
        }

        /* すべてのアートボードに名前を適用 / Apply names to all artboards */
        function applyContext(context) {
            for (var i = 0; i < artboards.length; i++) {
                artboards[i].name = buildArtboardName(i, context, originalNames);
            }
        }

        /* プレビューリストを更新 / Refresh preview list */
        function updatePreview() {
            var context = getRenameContext();
            previewUI.previewList.removeAll();

            if (!context.valid) {
                previewUI.previewList.add("item", L("previewInvalid"));
                return;
            }

            var maxCount = Math.min(15, artboardCount);
            for (var i = 0; i < maxCount; i++) {
                previewUI.previewList.add("item", buildArtboardName(i, context, originalNames));
            }

            if (artboardCount > 15) previewUI.previewList.add("item", L("previewTruncated"));
        }

        /* 現在のUI設定を平坦なオブジェクトとして取得 / Snapshot current UI settings */
        function currentSettings() {
            return {
                useFilename: prefixUI.useFilenameRadios[1].value,
                prefixSeparator: getSeparator(prefixUI.prefixSeparatorRadios),
                prefix: prefixUI.prefixInput.text,
                nameStyleKey: NAME_STYLE_KEYS[nameUI.nameStyleDropdown.selection.index],
                separator: getSeparator(suffixUI.suffixSeparatorRadios),
                formatKey: FORMAT_KEYS[suffixUI.formatDropdown.selection.index],
                start: suffixUI.startValueInput.text,
                increment: suffixUI.incrementInput.text,
                suffix: suffixUI.suffixInput.text
            };
        }

        /* プリセットの値をUIに反映 / Apply preset values to UI */
        function applyPreset(preset) {
            prefixUI.useFilenameRadios[0].value = !preset.useFilename;
            prefixUI.useFilenameRadios[1].value = preset.useFilename;
            setSeparatorRadios(prefixUI.prefixSeparatorRadios, preset.prefixSeparator);
            prefixUI.prefixInput.text = preset.prefix;

            var nameIdx = indexOfKey(NAME_STYLE_KEYS, preset.nameStyleKey);
            if (nameIdx >= 0) nameUI.nameStyleDropdown.selection = nameIdx;

            setSeparatorRadios(suffixUI.suffixSeparatorRadios, preset.separator);

            var fmtIdx = indexOfKey(FORMAT_KEYS, preset.formatKey);
            if (fmtIdx >= 0) suffixUI.formatDropdown.selection = fmtIdx;

            suffixUI.startValueInput.text = preset.start;
            suffixUI.incrementInput.text = preset.increment;
            suffixUI.suffixInput.text = preset.suffix;
        }

        // =========================================
        // イベント配線 / Event wiring
        // =========================================

        /* プリセット選択時：UIに反映してプレビュー更新 / On preset selection */
        presetUI.presetDropdown.onChange = function () {
            var index = presetUI.presetDropdown.selection.index;
            if (index <= 0) return;
            applyPreset(BUILTIN_PRESETS[index - 1]);
            updatePreview();
        };

        /* プリセット書き出し / Export preset */
        presetUI.savePresetBtn.onClick = function () {
            savePresetToFile(currentSettings());
        };

        /* ファイル名参照の切り替え：区切り文字ラジオの有効化を連動 / Toggle separator radios with filename usage */
        prefixUI.useFilenameRadios[0].onClick = prefixUI.useFilenameRadios[1].onClick = function () {
            for (var i = 0; i < 3; i++) {
                prefixUI.prefixSeparatorRadios[i].enabled = prefixUI.useFilenameRadios[1].value;
            }
            updatePreview();
        };

        bindAll(prefixUI.prefixSeparatorRadios, "onClick", updatePreview);
        bindAll(suffixUI.suffixSeparatorRadios, "onClick", updatePreview);

        /* 連番形式変更時：開始値をリセット、開始値・増分の有効化を切り替え / On format change: reset start and toggle inputs */
        suffixUI.formatDropdown.onChange = function () {
            var formatKey = FORMAT_KEYS[suffixUI.formatDropdown.selection.index];
            var hasNumber = (formatKey !== "none");
            suffixUI.startValueInput.enabled = hasNumber;
            suffixUI.incrementInput.enabled = (formatKey === "numeric");
            if (hasNumber) suffixUI.startValueInput.text = START_DEFAULTS[formatKey];
            updatePreview();
        };

        /* 入力変更時に即時プレビュー / Live preview on input change */
        prefixUI.prefixInput.onChanging = updatePreview;
        suffixUI.startValueInput.onChanging = updatePreview;
        suffixUI.incrementInput.onChanging = updatePreview;
        suffixUI.suffixInput.onChanging = updatePreview;
        nameUI.nameStyleDropdown.onChange = updatePreview;

        /* OKボタン：リネーム後にダイアログを閉じる / OK: rename and close */
        buttonUI.okBtn.onClick = function () {
            var context = validateContext();
            if (!context) return;
            try {
                applyContext(context);
            } catch (e) {
                alert(L("generalError") + e.message);
            }
            dialog.close();
        };

        suffixUI.formatDropdown.onChange(); // 初期プレビュー反映 / Initial preview
        dialog.show();
    }

    main();

})();