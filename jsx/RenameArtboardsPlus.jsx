#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false); 

/*
### スクリプト名：

RenameArtboardsPlus.jsx

### 概要

- アートボード名を一括で柔軟に変更できるIllustrator用スクリプトです。
- 接頭辞・接尾辞・ファイル名・連番・元アートボード名を自由に組み合わせて、高度な命名ルールを実現できます。

### 主な機能

- 接頭辞・接尾辞・元アートボード名の有無を選択可能
- ファイル名参照、区切り文字設定、プリセット保存／選択対応
- 連番形式（数字、大文字アルファベット、小文字アルファベット）選択
- 開始番号に基づく桁数（ゼロ埋め）の自動判定
- プレビュー機能で変更結果を事前確認
- 最大15件のプレビュー表示、ボタン配置美化
- ExtendScript（ES3）互換

### 処理の流れ

1. ダイアログで設定（接頭辞、接尾辞、連番形式など）を選択
2. プレビューで名称例を確認
3. 「適用」または「OK」でアートボード名を変更

### 更新履歴

- v1.0.0 (20250420) : 初期バージョン作成
- v1.0.1 (20250430) : 開始番号から桁数自動判定追加、プリセットラベル簡素化、ES3対応強化

---

### Script Name:

RenameArtboardsPlus.jsx

### Overview

- A script to flexibly batch rename artboards in Illustrator.
- Combine prefixes, suffixes, file name, numbers, and original artboard names to create advanced naming rules.

### Main Features

- Freely choose whether to include prefix, suffix, or original artboard name
- Supports file name reference, separator settings, preset save/load
- Numbering formats: number, uppercase alphabet, lowercase alphabet
- Auto-detect padding digits based on start number (zero-padding)
- Preview feature to check the renaming result beforehand
- Displays up to 15 preview items, improved button layout
- Compatible with ExtendScript (ES3)

### Process Flow

1. Configure options (prefix, suffix, numbering format, etc.) in dialog
2. Check name examples with the preview
3. Click "Apply" or "OK" to rename artboards

### Update History

- v1.0.0 (20250420): Initial version created
- v1.0.1 (20250430): Added auto-detection of padding digits, simplified preset labels, enhanced ES3 support
*/

function getCurrentLang() {
    return ($.locale && $.locale.indexOf('ja') === 0) ? 'ja' : 'en';
}

// -------------------------------
// 日英ラベル定義 Define label
// -------------------------------
var lang = getCurrentLang();
var LABELS = {
    dialogTitle: { ja: "アートボード名の一括設定", en: "Batch Rename Artboards" },
    prefixPanel: { ja: "接頭辞", en: "Prefix" },
    fileNameLabel: { ja: "ファイル名：", en: "File Name:" },
    useFileNo: { ja: "参照しない", en: "Do not use" },
    useFileYes: { ja: "参照する", en: "Use" },
    separatorLabel: { ja: "区切り文字：", en: "Separator:" },
    stringLabel: { ja: "文字列：", en: "String:" },
    namePanel: { ja: "アートボード名と番号", en: "Artboard Name & Number" },
    suffixPanel: { ja: "接尾辞", en: "Suffix" },
    formatLabel: { ja: "連番形式：", en: "Numbering Format:" },
    startLabel: { ja: "開始番号：", en: "Start Number:" },
    incrementLabel: { ja: "増分：", en: "Increment:" },
    preview: { ja: "プレビュー", en: "Preview" },
    cancel: { ja: "キャンセル", en: "Cancel" },
    apply: { ja: "適用", en: "Apply" },
    ok: { ja: "OK", en: "OK" },
    savePreset: { ja: "プリセット書き出し", en: "Export Preset" },
    presetNone: { ja: "（未選択）", en: "(None)" }
};

(function () {
    if (app.documents.length === 0) return;

    var doc = app.activeDocument;
    var filename = doc.name.replace(/\.[^\.]+$/, "");
    var artboards = doc.artboards;
    var total = artboards.length;
    var originalNames = [];
    for (var i = 0; i < total; i++) originalNames.push(artboards[i].name);

    var dialog = new Window("dialog", LABELS.dialogTitle[lang]);
    dialog.alignChildren = "fill";

    var mainGroup = dialog.add("group");
    mainGroup.orientation = "row";
    mainGroup.alignChildren = ["top", "fill"];

    var inputGroup = mainGroup.add("group");
    inputGroup.spacing = 20;
    inputGroup.orientation = "column";
    inputGroup.alignChildren = "left";

    var presetGroup = inputGroup.add("group");
    presetGroup.orientation = "row";
    presetGroup.alignChildren = "left";

    var builtinPresets = [
{ label: "ファイル名+連番3", useFilename: true, prefixSeparator: "-", prefix: "", nameStyle: "（なし）", separator: "", format: "数字", start: "01", increment: "1", suffix: "" },
{ label: "アートボード名と連番", useFilename: false, prefixSeparator: "-", prefix: "", nameStyle: "名称", separator: "-", format: "数字", start: "1", increment: "1", suffix: "" } ];

    var presetItems = [LABELS.presetNone[lang]];
    for (var i = 0; i < builtinPresets.length; i++) {
        presetItems.push(builtinPresets[i].label);
    }

    var presetDropdown = presetGroup.add("dropdownlist", undefined, presetItems);
    presetDropdown.selection = 0;
    presetDropdown.enabled = true;

    var savePresetBtn = presetGroup.add("button", undefined, LABELS.savePreset[lang]);

    presetDropdown.onChange = function () {
        var index = presetDropdown.selection.index;
        if (index <= 0) return;

        var preset = builtinPresets[index - 1];

        useFilenameRadios[0].value = !preset.useFilename;
        useFilenameRadios[1].value = preset.useFilename;

        prefixSeparatorRadios[0].value = (preset.prefixSeparator === "");
        prefixSeparatorRadios[1].value = (preset.prefixSeparator === "-");
        prefixSeparatorRadios[2].value = (preset.prefixSeparator === "_");

        prefixInput.text = preset.prefix;

        for (var i = 0; i < nameStyleDropdown.items.length; i++) {
            if (nameStyleDropdown.items[i].text === preset.nameStyle) {
                nameStyleDropdown.selection = i;
                break;
            }
        }

        separatorRadios2[0].value = (preset.separator === "");
        separatorRadios2[1].value = (preset.separator === "-");
        separatorRadios2[2].value = (preset.separator === "_");

        for (var i = 0; i < formatDropdown.items.length; i++) {
            if (formatDropdown.items[i].text === preset.format) {
                formatDropdown.selection = i;
                break;
            }
        }

        startNumberInput.text = preset.start;
        incrementInput.text = preset.increment;
        suffixInput.text = preset.suffix;

        updatePreview();
    };

    // ▼ 接頭辞パネル
    var prefixPanel = inputGroup.add("panel", undefined, LABELS.prefixPanel[lang]);
    prefixPanel.margins = [20, 20, 20, 10];
    prefixPanel.orientation = "column";
    prefixPanel.alignChildren = "left";

    var filenameGroup = prefixPanel.add("group");
    filenameGroup.add("statictext", undefined, LABELS.fileNameLabel[lang]);
    var useFilenameRadios = [
        filenameGroup.add("radiobutton", undefined, LABELS.useFileNo[lang]),
        filenameGroup.add("radiobutton", undefined, LABELS.useFileYes[lang])
    ];
    useFilenameRadios[0].value = true;

    var separatorGroup1 = prefixPanel.add("group");
    separatorGroup1.add("statictext", undefined, LABELS.separatorLabel[lang]);
    var prefixSeparatorRadios = [
        separatorGroup1.add("radiobutton", undefined, "なし"),
        separatorGroup1.add("radiobutton", undefined, "-"),
        separatorGroup1.add("radiobutton", undefined, "_")
    ];
    prefixSeparatorRadios[0].value = true;
    for (var i = 0; i < 3; i++) prefixSeparatorRadios[i].enabled = false;

    var prefixGroup = prefixPanel.add("group");
    prefixGroup.add("statictext", undefined, LABELS.stringLabel[lang]);
    var prefixInput = prefixGroup.add("edittext", undefined, "");
    prefixInput.characters = 16;

    // ▼ アートボード名と番号パネル
    var namePanel = inputGroup.add("panel", undefined, LABELS.namePanel[lang]);
    namePanel.margins = [20, 20, 20, 10];
    namePanel.orientation = "row";
    namePanel.alignChildren = "left";
    var nameStyleDropdown = namePanel.add("dropdownlist", undefined, [
        "（なし）", "番号", "名称", "番号-名称", "番号_名称"
    ]);
    nameStyleDropdown.selection = 0;

    // ▼ 接尾辞パネル
    var suffixPanel = inputGroup.add("panel", undefined, LABELS.suffixPanel[lang]);
    suffixPanel.margins = [20, 20, 20, 10];
    suffixPanel.orientation = "column";
    suffixPanel.alignChildren = "left";

    var separatorGroup2 = suffixPanel.add("group");
    separatorGroup2.add("statictext", undefined, LABELS.separatorLabel[lang]);
    var separatorRadios2 = [
        separatorGroup2.add("radiobutton", undefined, "なし"),
        separatorGroup2.add("radiobutton", undefined, "-"),
        separatorGroup2.add("radiobutton", undefined, "_")
    ];
    separatorRadios2[0].value = true;

    var formatGroup = suffixPanel.add("group");
    formatGroup.add("statictext", undefined, LABELS.formatLabel[lang]);
    var formatDropdown = formatGroup.add("dropdownlist", undefined, [
        "数字", "アルファベット（大文字）", "アルファベット（小文字）"
    ]);
    formatDropdown.selection = 0;

    var startGroup = suffixPanel.add("group");
    startGroup.add("statictext", undefined, LABELS.startLabel[lang]);
    var startNumberInput = startGroup.add("edittext", undefined, "1");
    startNumberInput.characters = 5;

    var incrementGroup = suffixPanel.add("group");
    incrementGroup.add("statictext", undefined, LABELS.incrementLabel[lang]);
    var incrementInput = incrementGroup.add("edittext", undefined, "1");
    incrementInput.characters = 5;

    var suffixGroup = suffixPanel.add("group");
    suffixGroup.add("statictext", undefined, LABELS.stringLabel[lang]);
    var suffixInput = suffixGroup.add("edittext", undefined, "");
    suffixInput.characters = 16;

    // ▼ プレビューエリア
    var previewGroup = mainGroup.add("panel", undefined, LABELS.preview[lang]);
    previewGroup.alignChildren = "fill";
    previewGroup.margins = [10, 20, 10, 10];
    previewGroup.preferredSize.width = 250;
    previewGroup.preferredSize.height = 380;
    var previewList = previewGroup.add("listbox", undefined, [], {
        multiselect: false,
        numberOfColumns: 1,
        showHeaders: false
    });
    previewList.preferredSize.height = 320;

    // ▼ 下部ボタンエリア
    var outerGroup = dialog.add("group");
    outerGroup.orientation = "row";
    outerGroup.alignChildren = ["fill", "center"];
    outerGroup.margins = [0, 10, 0, 0];
    outerGroup.spacing = 0;

    var leftGroup = outerGroup.add("group");
    leftGroup.orientation = "row";
    leftGroup.alignChildren = "left";
    var cancelBtn = leftGroup.add("button", undefined, LABELS.cancel[lang], { name: "cancel" });

    var spacer = outerGroup.add("group");
    spacer.alignment = ["fill", "fill"];
    spacer.minimumSize.width = 50;

    var rightGroup = outerGroup.add("group");
    rightGroup.orientation = "row";
    rightGroup.alignChildren = ["right", "center"];
    rightGroup.spacing = 10;
    var applyBtn = rightGroup.add("button", undefined, LABELS.apply[lang]);
    var okBtn = rightGroup.add("button", undefined, LABELS.ok[lang], { name: "ok" });

    // ▼ プリセット保存処理（label: デコード済み）

savePresetBtn.onClick = function () {
    var saveFile = File.saveDialog("プリセットを書き出す場所と名前を指定してください", "*.txt");
    if (!saveFile) return;

    if (saveFile.name.indexOf(".") === -1) {
        saveFile = new File(saveFile.fsName + ".txt");
    }

    var rawName = saveFile.name;
    var fileName = rawName.replace(/\.txt$/i, "");
    var labelName = decodeURIComponent(fileName); // 日本語ファイル名もOK

    var prefixSep = "";
    if (prefixSeparatorRadios[1].value) prefixSep = "-";
    else if (prefixSeparatorRadios[2].value) prefixSep = "_";

    var suffixSep = "";
    if (separatorRadios2[1].value) suffixSep = "-";
    else if (separatorRadios2[2].value) suffixSep = "_";

    var presetString =
        '{ label: "' + labelName + '", ' +
        'useFilename: ' + (useFilenameRadios[1].value ? 'true' : 'false') + ', ' +
        'prefixSeparator: "' + prefixSep + '", ' +
        'prefix: "' + prefixInput.text + '", ' +
        'nameStyle: "' + nameStyleDropdown.selection.text + '", ' +
        'separator: "' + suffixSep + '", ' +
        'format: "' + formatDropdown.selection.text + '", ' +
        'start: "' + startNumberInput.text + '", ' +
        'increment: "' + incrementInput.text + '", ' +
        'suffix: "' + suffixInput.text + '" }';

    try {
        if (saveFile.open("w")) {
            saveFile.write(presetString);
            saveFile.close();
            alert("プリセットを書き出しました：\n" + saveFile.fsName);
        } else {
            alert("ファイルを書き込めませんでした。");
        }
    } catch (e) {
        alert("プリセットの保存に失敗しました：\n" + e.message);
    }
};

    // ▼ 補助関数群

    function getPrefixCombined() {
        var base = prefixInput.text;
        var sep1 = prefixSeparatorRadios[1].value ? "-" : (prefixSeparatorRadios[2].value ? "_" : "");
        if (useFilenameRadios[1].value) base = filename + sep1 + base;
        return base;
    }

    function padNumber(num, length) {
        var str = String(num);
        while (str.length < length) str = "0" + str;
        return str;
    }

    function getAlphaLabel(index, format) {
        var label = "";
        while (index >= 0) {
            label = String.fromCharCode((index % 26) + 65) + label;
            index = Math.floor(index / 26) - 1;
        }
        return (format === "アルファベット（小文字）") ? label.toLowerCase() : label;
    }

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

function getPadLengthFromStart(str) {
    if (/^\d+$/.test(str)) {
        return str.length; // 数字であればその桁数で判断（ゼロ含む）
    }
    return 0;
}

    function getNameComponent(index, nameStyle) {
        var abNum = (index + 1).toString();
        var abName = originalNames[index];
        switch (nameStyle) {
            case "番号": return abNum;
            case "名称": return abName;
            case "番号-名称": return abNum + "-" + abName;
            case "番号_名称": return abNum + "_" + abName;
            default: return "";
        }
    }

    function updatePreview() {
        var format = formatDropdown.selection.text;
        var rawStart = startNumberInput.text;
        var startNum = (format.indexOf("アルファベット") === 0) ? getAlphaIndex(rawStart) : parseInt(rawStart, 10);
        var increment = parseInt(incrementInput.text, 10);
        var padLength = getPadLengthFromStart(rawStart);
        var nameStyle = nameStyleDropdown.selection.text;
        previewList.removeAll();

        if (isNaN(startNum) || startNum <= 0 || (format === "数字" && (isNaN(increment) || increment <= 0))) {
            previewList.add("item", "※数値が正しくありません");
            return;
        }

        var prefixCombined = getPrefixCombined();
        var sep2 = separatorRadios2[1].value ? "-" : (separatorRadios2[2].value ? "_" : "");

        var maxCount = Math.min(15, total);
        for (var i = 0; i < maxCount; i++) {
            var num = startNum + ((format === "数字") ? i * increment : i);
            var label = (format === "数字") ? padNumber(num, padLength) : getAlphaLabel(num - 1, format);
            var nameComponent = getNameComponent(i, nameStyle);
            var finalName = prefixCombined + nameComponent + sep2 + label + suffixInput.text;
            previewList.add("item", finalName);
        }

        if (total > 15) previewList.add("item", "...（以下省略）");
    }

    // ▼ イベントハンドラ

 useFilenameRadios[0].onClick = useFilenameRadios[1].onClick = function () {
    for (var i = 0; i < 3; i++) prefixSeparatorRadios[i].enabled = useFilenameRadios[1].value;
    updatePreview();
};

for (var i = 0; i < prefixSeparatorRadios.length; i++) {
    prefixSeparatorRadios[i].onClick = updatePreview;
}
for (var i = 0; i < separatorRadios2.length; i++) {
    separatorRadios2[i].onClick = updatePreview;
}

formatDropdown.onChange = function () {
    var format = formatDropdown.selection.text;
    var isAlpha = format.indexOf("アルファベット") === 0;
    incrementInput.enabled = !isAlpha;

    if (format === "数字") {
        startNumberInput.text = "1";
    } else if (format === "アルファベット（大文字）") {
        startNumberInput.text = "A";
    } else if (format === "アルファベット（小文字）") {
        startNumberInput.text = "a";
    }

    updatePreview();
};

// ▼ 入力変更時にプレビューを即時更新
prefixInput.onChanging = updatePreview;
startNumberInput.onChanging = updatePreview;
incrementInput.onChanging = updatePreview;
suffixInput.onChanging = updatePreview;
nameStyleDropdown.onChange = updatePreview;


applyBtn.onClick = function () {
    updatePreview();
    try {
        var format = formatDropdown.selection.text;
        var rawStart = startNumberInput.text;
        var startNum = (format.indexOf("アルファベット") === 0) ? getAlphaIndex(rawStart) : parseInt(rawStart, 10);
        var increment = parseInt(incrementInput.text, 10);
        var padLength = getPadLengthFromStart(rawStart);
        var nameStyle = nameStyleDropdown.selection.text;

        if (isNaN(startNum) || startNum <= 0 || (format === "数字" && (isNaN(increment) || increment <= 0))) {
            alert("入力内容に誤りがあります。\n正しい数値を指定してください。", "エラー");
            return;
        }

        var prefixCombined = getPrefixCombined();
        var sep2 = separatorRadios2[1].value ? "-" : (separatorRadios2[2].value ? "_" : "");
        var current = startNum;
        for (var i = 0; i < artboards.length; i++) {
            var label = (format === "数字")
                ? padNumber(current, padLength)
                : getAlphaLabel(current - 1, format);
            var nameComponent = getNameComponent(i, nameStyle);
            var finalName = prefixCombined + nameComponent + sep2 + label + suffixInput.text;
            artboards[i].name = finalName;
            current += (format === "数字") ? increment : 1;
        }

        app.redraw(); // 反映
    } catch (e) {
        alert("エラーが発生しました（適用時）：\n" + e.message);
    }
};

    okBtn.onClick = function () {
        var format = formatDropdown.selection.text;
        var rawStart = startNumberInput.text;
        var startNum = (format.indexOf("アルファベット") === 0) ? getAlphaIndex(rawStart) : parseInt(rawStart, 10);
        var increment = parseInt(incrementInput.text, 10);
        var padLength = getPadLengthFromStart(rawStart);
        var nameStyle = nameStyleDropdown.selection.text;

        if (isNaN(startNum) || startNum <= 0 || (format === "数字" && (isNaN(increment) || increment <= 0))) {
            alert("入力内容に誤りがあります。\n正しい数値を指定してください。", "エラー");
            return;
        }

        try {
            var prefixCombined = getPrefixCombined();
            var sep2 = separatorRadios2[1].value ? "-" : (separatorRadios2[2].value ? "_" : "");
            var current = startNum;
            for (var i = 0; i < artboards.length; i++) {
                var label = (format === "数字")
                    ? padNumber(current, padLength)
                    : getAlphaLabel(current - 1, format);
                var nameComponent = getNameComponent(i, nameStyle);
                var finalName = prefixCombined + nameComponent + sep2 + label + suffixInput.text;
                artboards[i].name = finalName;
                current += (format === "数字") ? increment : 1;
            }
        } catch (e) {
            alert("エラーが発生しました：\n" + e.message);
        }

        dialog.close();
    };

    formatDropdown.onChange(); // 初期プレビュー反映
    dialog.show();
})();