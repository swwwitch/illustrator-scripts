#target illustrator

/*
 * スクリプト名：SmartArtboardRenamer.jsx
 *
 * スクリプトの概要：
 * Adobe Illustrator のアートボード名を一括でリネームするツールです。
 * 接頭辞・接尾辞・参照テキストの組み合わせにより、柔軟な命名が可能です。
 * 「最前面のテキスト」モードでは、アートボードごとに最も上にあるテキストフレーム（Z順で前面）を1つ参照します。
 * 「レイヤーを指定」モードでは、指定レイヤー以下のすべてのテキストフレームを、
 * 見た目の上から順（Z順）に結合して使用します。
 *
 * 【主な機能】（日本語）
 * - 接頭辞・接尾辞には連番（1, 01, A, a）、ファイル名（#FN）を含めることが可能
 * - 「最前面のテキスト」ではアートボードごとにZ順で最前面のテキストフレームを参照
 * - 「レイヤーを指定」ではレイヤー以下の全テキストをZ順で結合
 * - 非表示レイヤーのオブジェクトは無視
 * - アートボードの対象範囲（すべて／一部）を選択可能
 * - 同名が重複する場合は "_2", "_3" などを自動付加
 *
 * Script Name: SmartArtboardRenamer.jsx
 *
 * Description (English):
 * This script batch-renames artboards in Adobe Illustrator.
 * You can flexibly construct artboard names using a prefix, suffix, and reference text.
 *
 * [Key Features] (English)
 * - Prefix/suffix supports sequential formats (1, 01, A, a) and file name (#FN)
 * - Use either the "frontmost text" or "text in a selected layer" as the name
 * - Layer mode includes a popup menu to choose from all layers
 * - Only objects in the selected layer are used when in layer mode
 * - You can target all or specific artboards by number range
 * - If names conflict, "_2", "_3", etc. are appended to make them unique
 * - Hidden layers are ignored
 *
 * 作成日：2025-05-09
 * 最終更新日：2025-05-12（読込確認・調整準備）
 */

//@target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false); 

function getCurrentLang() {
    return ($.locale && $.locale.indexOf('ja') === 0) ? 'ja' : 'en';
}

function main() {
    if (app.documents.length === 0) return;
    var doc = app.activeDocument;
    // Store original artboard names before dialog
    var originalNames = [];
    for (var i = 0; i < doc.artboards.length; i++) {
        originalNames.push(doc.artboards[i].name);
    }
    var dialogResult = showDialog();
    if (!dialogResult) {
        // Restore original artboard names if dialog was cancelled
        for (var i = 0; i < doc.artboards.length; i++) {
            doc.artboards[i].name = originalNames[i];
        }
        return;
    }

    var mode = dialogResult.mode;
    var prefixInput = dialogResult.prefix;
    var suffixInput = dialogResult.suffix;
    var targetType = dialogResult.artboardTarget;
    var numberInput = dialogResult.numberInput;

    var textFrames;
    if (mode === "layer") {
        var namingLayer = getNamingLayerByName(doc, dialogResult.selectedLayerName);
        if (!namingLayer || !namingLayer.visible) {
            alert("指定されたレイヤーが見つからないか、非表示です。");
            return;
        }
        textFrames = getTextFramesInLayer(namingLayer);
    } else if (mode === "frontmost") {
        textFrames = getFrontmostTextFramesPerArtboard(doc);
    } else {
        textFrames = []; // "none" モードでは空の配列
    }

    var targetIndices = getTargetArtboardIndices(doc.artboards.length, targetType, numberInput);
    var artboardTextMap = mapTextToArtboards(textFrames, doc.artboards, mode);
    renameArtboards(doc.artboards, artboardTextMap, prefixInput, suffixInput, targetIndices);
}

function getNamingLayerByName(doc, name) {
    for (var i = 0; i < doc.layers.length; i++) {
        if (doc.layers[i].name === name) return doc.layers[i];
    }
    return null;
}

function getTextFramesInLayer(layer) {
    var results = [];
    if (!layer.visible) return results;
    var items = layer.pageItems;
    for (var i = 0; i < items.length; i++) {
        if (items[i].typename === "TextFrame") results.push(items[i]);
    }
    return results;
}

function getAllTextFrames(doc) {
    var results = [];
    var layers = doc.layers;
    for (var i = 0; i < layers.length; i++) {
        if (!layers[i].visible) continue;
        var items = layers[i].textFrames;
        for (var j = 0; j < items.length; j++) {
            results.push(items[j]);
        }
    }
    return results;
}

function showDialog() {
    var lang = getCurrentLang();
    var LABELS = {
        dialogTitle:       { ja: "アートボードのリネーム", en: "Rename Artboards" },
        prefixLabel:       { ja: "接頭辞", en: "Prefix" },
        suffixLabel:       { ja: "接尾辞", en: "Suffix" },
        sourceLabel:       { ja: "参照するテキスト", en: "Source Text" },
        layerOption:       { ja: "レイヤーを指定", en: "Specify Layer" },
        frontmostOption:   { ja: "最前面のテキスト", en: "Frontmost Text" },
        ignoreOption:      { ja: "参照しない", en: "Ignore" },
        targetLabel:       { ja: "対象：", en: "Target:" },
        allBoards:         { ja: "すべて", en: "All" },
        specificBoards:    { ja: "指定", en: "Range" },
        cancelButton:      { ja: "キャンセル", en: "Cancel" },
        okButton:          { ja: "OK", en: "OK" },
        exampleFormatHint: {
            ja: "例：\n1, 01\n#FN（ファイル名）\n#DT（日付）",
            en: "Example:\n1, 01\n#FN (FileName)\n#DT (Date)"
        },
        applyButton:       { ja: "適用", en: "Apply" }
    };

    var dlg = new Window("dialog", LABELS.dialogTitle[lang]);
    dlg.orientation = "column";
    dlg.alignChildren = "fill";

    var mainGroup = dlg.add("group");
    mainGroup.orientation = "row";
    mainGroup.alignChildren = "top";

    // 接頭辞
    var leftCol = mainGroup.add("panel", undefined, LABELS.prefixLabel[lang]);
    leftCol.orientation = "column";
    leftCol.alignChildren = "left";
    leftCol.margins = [15, 20, 15, 10];
    var prefixInput = leftCol.add("edittext", undefined, "");
    prefixInput.characters = 12;
    prefixInput.active = true;
    var prefixHint = leftCol.add("statictext", undefined, LABELS.exampleFormatHint[lang], { multiline: true });
    // prefixHint.preferredSize.width = 160;
    prefixHint.preferredSize.height = 80;
    prefixHint.enabled = true;

    // 参照テキスト
    var centerCol = mainGroup.add("panel", undefined, LABELS.sourceLabel[lang]);
    centerCol.orientation = "column";
    centerCol.alignChildren = "left";
    centerCol.margins = [15, 25, 15, 15];
    centerCol.preferredSize.height = 65;

    var r1 = centerCol.add("radiobutton", undefined, LABELS.frontmostOption[lang]);
    var r3 = centerCol.add("radiobutton", undefined, LABELS.ignoreOption[lang]);
    var r2 = centerCol.add("radiobutton", undefined, LABELS.layerOption[lang]);
    r1.value = true;
    r3.value = false;

    // ポップアップ：全レイヤー
    // Ensure dropdown is below "レイヤーを指定" radio (r2)
    var layerDropdown = centerCol.add("dropdownlist", undefined, []);
    layerDropdown.minimumSize.width = 130;
    var allLayers = app.activeDocument.layers;
    for (var i = 0; i < allLayers.length; i++) {
        layerDropdown.add("item", allLayers[i].name);
    }
    if (layerDropdown.items.length > 0) layerDropdown.selection = 0;
    layerDropdown.enabled = false;

    // ラジオ切替で有効化
    r1.onClick = function () {
        layerDropdown.enabled = false;
        btnApply.notify();
    };
    r2.onClick = function () {
        layerDropdown.enabled = true;
        btnApply.notify();
    };
    r3.onClick = function () {
        layerDropdown.enabled = false;
        // btnApply.notify();
    };

    // ドロップダウン変更時のイベント
    layerDropdown.onChange = function () {
        if (r2.value) btnApply.notify();
    };

    // 接尾辞
    var rightCol = mainGroup.add("panel", undefined, LABELS.suffixLabel[lang]);
    rightCol.orientation = "column";
    rightCol.alignChildren = "left";
    rightCol.margins = [15, 20, 15, 10];
    var suffixInput = rightCol.add("edittext", undefined, "");
    suffixInput.characters = 12;
    var suffixHint = rightCol.add("statictext", undefined, LABELS.exampleFormatHint[lang], { multiline: true });
    // suffixHint.preferredSize.width = 100;
    suffixHint.preferredSize.height = 80;
    suffixHint.enabled = true;


    // 対象アートボード選択
    var targetGroup = dlg.add("group");
    targetGroup.orientation = "row";
    targetGroup.alignChildren = ["left", "center"];
    targetGroup.margins = [15,0, 15, 0];
    targetGroup.add("statictext", undefined, LABELS.targetLabel[lang]);
    var abAll = targetGroup.add("radiobutton", undefined, LABELS.allBoards[lang]);
    var abNumbered = targetGroup.add("radiobutton", undefined, LABELS.specificBoards[lang]);
    abAll.value = true;
    var abNumbersInput = targetGroup.add("edittext", undefined, "");
    abNumbersInput.characters = 10;
    abNumbersInput.enabled = false;
    abNumbered.onClick = function () { abNumbersInput.enabled = true; };
    abAll.onClick = function () { abNumbersInput.enabled = false; };

   

// ボタンエリア（左右グループ配置：キャンセルは左、適用とOKは右）
var outerGroup = dlg.add("group");
outerGroup.orientation = "column";
outerGroup.alignChildren = ["fill", "center"];
outerGroup.spacing = 10;
outerGroup.margins = 0;

var buttonRow = outerGroup.add("group");
buttonRow.orientation = "row";
buttonRow.alignChildren = ["fill", "center"];
buttonRow.alignment = "fill";
buttonRow.margins = [0, 15, 0, 0];
buttonRow.spacing = 10;

var btnCancel = buttonRow.add("button", undefined, LABELS.cancelButton[lang], { name: "cancel" });
btnCancel.preferredSize.width = 80;

var spacer = buttonRow.add("group");
spacer.alignment = "fill";
spacer.preferredSize.width = 160;

var rightButtons = buttonRow.add("group");
rightButtons.orientation = "row";
rightButtons.alignChildren = ["right", "center"];
var btnApply = rightButtons.add("button", undefined, LABELS.applyButton[lang]);
btnApply.onClick = function () {
    var mode = r1.value ? "frontmost" : (r2.value ? "layer" : "none");
    var prefix = prefixInput.text;
    var suffix = suffixInput.text;
    var targetType = abAll.value ? "all" : "numbered";
    var numberInput = abNumbersInput.text;
    var layerName = r2.value && layerDropdown.selection ? layerDropdown.selection.text : null;

    var textFrames;
    if (mode === "layer") {
        var namingLayer = getNamingLayerByName(app.activeDocument, layerName);
        if (!namingLayer || !namingLayer.visible) {
            alert("指定されたレイヤーが見つからないか、非表示です。");
            return;
        }
        textFrames = getTextFramesInLayer(namingLayer);
    } else if (mode === "frontmost") {
        textFrames = getFrontmostTextFramesPerArtboard(app.activeDocument);
    } else {
        textFrames = []; // "none" モードでは空の配列
    }

    var targetIndices = getTargetArtboardIndices(app.activeDocument.artboards.length, targetType, numberInput);
    var artboardTextMap = mapTextToArtboards(textFrames, app.activeDocument.artboards, mode);

    // Validation: If mode is "none" and both prefix and suffix are empty, alert and stop.
    if (mode === "none" && prefix === "" && suffix === "") {
        alert("いずれの設定が必要です");
        return;
    }

    renameArtboards(app.activeDocument.artboards, artboardTextMap, prefix, suffix, targetIndices);
    // Ensure UI redraw is the last executed statement
    app.redraw();
};
var btnOK = rightButtons.add("button", undefined, LABELS.okButton[lang], { name: "ok" });

    if (dlg.show() !== 1) return null;

    return {
        mode: r1.value ? "frontmost" : (r2.value ? "layer" : "none"),
        prefix: prefixInput.text,
        suffix: suffixInput.text,
        artboardTarget: abAll.value ? "all" : "numbered",
        numberInput: abNumbersInput.text,
        selectedLayerName: r2.value && layerDropdown.selection ? layerDropdown.selection.text : null
    };
}

function getNamingLayer(doc) {
    for (var i = 0; i < doc.layers.length; i++) {
        if (doc.layers[i].name === "_abNames") return doc.layers[i];
    }
    return null;
}

function getTextFramesInLayer(layer) {
    var results = [];

    function collectTextFrames(targetLayer) {
        if (!targetLayer.visible) return;
        var items = targetLayer.pageItems;
        for (var i = items.length - 1; i >= 0; i--) {
            if (items[i].typename === "TextFrame") {
                results.push(items[i]);
            }
        }
        for (var j = targetLayer.layers.length - 1; j >= 0; j--) {
            collectTextFrames(targetLayer.layers[j]);
        }
    }

    collectTextFrames(layer);
    return results;
}

function getAllTextFrames(doc) {
    var results = [];
    var items = doc.textFrames;
    for (var i = 0; i < items.length; i++) results.push(items[i]);
    return results;
}

function getTextCenter(tf) {
    var b = tf.visibleBounds;
    return [(b[0] + b[2]) / 2, (b[1] + b[3]) / 2];
}

function parseArtboardRange(input) {
    var result = [];
    var parts = input.split(",");
    for (var i = 0; i < parts.length; i++) {
        var part = parts[i].replace(/\s+/g, "");
        if (/^\d+$/.test(part)) {
            result.push(parseInt(part, 10) - 1);
        } else if (/^\d+-\d+$/.test(part)) {
            var range = part.split("-");
            var start = parseInt(range[0], 10);
            var end = parseInt(range[1], 10);
            for (var j = start; j <= end; j++) result.push(j - 1);
        }
    }
    return result;
}

function getTargetArtboardIndices(total, mode, input) {
    if (mode === "all") {
        var list = [];
        for (var i = 0; i < total; i++) list.push(i);
        return list;
    }
    return parseArtboardRange(input);
}

function mapTextToArtboards(textFrames, artboards, mode) {
    var map = {};
    for (var i = 0; i < textFrames.length; i++) {
        var tf = textFrames[i];
        var center = getTextCenter(tf);
        for (var j = 0; j < artboards.length; j++) {
            var ab = artboards[j].artboardRect;
            if (center[0] >= ab[0] && center[0] <= ab[2] &&
                center[1] <= ab[1] && center[1] >= ab[3]) {
                if (!map[j]) map[j] = [];
                var cleanedText = tf.contents.replace(/[\r\n\t]/g, "");
                map[j].push(cleanedText);
                break;
            }
        }
    }
    return map;
}

function contains(array, value) {
    for (var i = 0; i < array.length; i++) {
        if (array[i] === value) return true;
    }
    return false;
}

function generateUniqueName(baseName, usedNames) {
    var name = baseName;
    var count = 2;
    while (contains(usedNames, name)) {
        name = baseName + "_" + count;
        count++;
    }
    return name;
}

function formatWithAlphaNumeric(template, index) {
    var fileName = app.activeDocument.name.replace(/\.[^\.]+$/, "");
    var now = new Date();
    var dateString = now.getFullYear().toString() +
                     ("0" + (now.getMonth() + 1)).slice(-2) +
                     ("0" + now.getDate()).slice(-2);

    // Temporarily hide #FN and #DT for matching
    var fnPlaceholder = "<<<FILENAME>>>";
    var dtPlaceholder = "<<<DATETIME>>>";
    var templateForMatch = template.replace(/#FN/g, fnPlaceholder).replace(/#DT/g, dtPlaceholder);

    // Search only in non-placeholder part, only match digit sequences
    var match = templateForMatch.match(/\d+/);
    if (match) {
        var token = match[0];
        var tokenIndex = templateForMatch.indexOf(token);

        var replacement = "";

        if (/^\d+$/.test(token)) {
            var base = parseInt(token, 10);
            var value = base + index - 1;
            replacement = (token.charAt(0) === "0" && token.length > 1)
                ? ("0000000000" + value).slice(-token.length)
                : value.toString();
        } else {
            replacement = token;
        }

        // Replace token in templateForMatch
        var interim = templateForMatch.replace(token, replacement);
        // Revert placeholders
        return interim.replace(fnPlaceholder, fileName).replace(dtPlaceholder, dateString);
    }

    // If no match, still expand placeholders (also handle #FN/#DT in template)
    var finalInterim = templateForMatch;
    finalInterim = finalInterim.replace(fnPlaceholder, fileName).replace(dtPlaceholder, dateString);
    return finalInterim;
}

function renameArtboards(artboards, map, prefixInput, suffixInput, targetIndices) {
    var usedNames = [];
    var count = 1;
    for (var i = 0; i < artboards.length; i++) {
        if (!contains(targetIndices, i)) continue;
        var prefix = formatWithAlphaNumeric(prefixInput, count);
        var suffix = formatWithAlphaNumeric(suffixInput, count);
        var textPart = (map[i] && map[i].length > 0) ? map[i].join(" ") : "";
        if (prefix || suffix || textPart) {
            var baseName = (prefix ? prefix : "") + textPart + (suffix ? suffix : "");
            var uniqueName = generateUniqueName(baseName, usedNames);
            usedNames.push(uniqueName);
            try {
                artboards[i].name = uniqueName;
            } catch (e) {}
            count++;
        }
    }
}

function getFrontmostTextFramesPerArtboard(doc) {
    var result = [];
    var abCount = doc.artboards.length;

    for (var i = 0; i < abCount; i++) {
        var abRect = doc.artboards[i].artboardRect;
        var best = null;
        var layers = doc.layers;
        for (var li = 0; li < layers.length; li++) {
            var layer = layers[li];
            if (!layer.visible) continue;
            var items = layer.pageItems;
            for (var pi = 0; pi < items.length; pi++) {
                var item = items[pi];
                if (item.typename !== "TextFrame") continue;
                var b = item.visibleBounds;
                var cx = (b[0] + b[2]) / 2;
                var cy = (b[1] + b[3]) / 2;
                if (cx >= abRect[0] && cx <= abRect[2] &&
                    cy <= abRect[1] && cy >= abRect[3]) {
                    best = item;
                    break;
                }
            }
            if (best) break;
        }
        if (best) result.push(best);
    }

    return result;
}

main();