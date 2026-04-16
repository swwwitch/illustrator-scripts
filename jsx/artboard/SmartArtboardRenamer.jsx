#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

(function () {

    /*
    ### スクリプト名：
    
    SmartArtboardRenamer.jsx
    
    ### 概要
    
    - Illustratorのアートボード名を、接頭辞・接尾辞・参照テキストを組み合わせて一括リネームするスクリプトです。
    - ダイアログ上で設定を変更すると、アートボード名を自動でプレビュー更新します。
    
    - 接頭辞・接尾辞に連番（1, 01, A, a）、ファイル名（#FN）、日付（#DT）を含めることが可能
    - #FN と #DT の説明は、接頭辞・接尾辞パネル内のツールヒントに表示
    - 「最前面テキスト」モードで各アートボードの最前面テキストを参照
    - 「レイヤー内テキスト」モードで指定レイヤー内のテキストを結合して参照
    - 「ファイル名」モードでドキュメントのファイル名を参照
    - 「使わない」モードで参照テキストなしの命名が可能
    - 対象アートボードを「すべて」または「指定範囲」から選択可能
    - 同名が重複する場合は "_2", "_3" などを自動付加
    - 非表示レイヤーは無視
    - キャンセル時はダイアログ表示前のアートボード名に戻す
    
    ### 処理の流れ
    
    1. 対象アートボードを「すべて」または「指定範囲」から選択
    2. 接頭辞・参照テキスト・接尾辞を設定
    3. 入力内容の変更に応じてアートボード名を自動プレビュー
    4. OKで確定、キャンセルで元の名前に戻す
    
    ### 更新履歴
    
    - v1.0 (20250509) : 初期バージョン
    - v1.1 (20250512) : レイヤー参照改善、UI調整
    
    ---
    
    ### Script Name:
    
    SmartArtboardRenamer.jsx
    
    ### Overview
    
    - A script to batch rename Illustrator artboards by combining prefix, suffix, and reference text.
    - As settings change in the dialog, the artboard names are updated automatically for preview.
    
    - Prefix/suffix can include sequential tokens (1, 01, A, a), file name (#FN), and date (#DT)
    - Explanations for #FN and #DT are shown in the tooltips inside the Prefix and Suffix panels
    - "Frontmost Text" mode uses the frontmost text frame on each artboard
    - "Layer Text" mode combines text from the selected layer
    - "File Name" mode uses the document file name as the reference text
    - "None" mode allows naming without reference text
    - Target artboards can be set to "All" or "Range"
    - Automatically appends "_2", "_3", etc. to avoid duplicate names
    - Ignores hidden layers
    - Cancelling restores the original artboard names from before the dialog opened
    
    ### Process Flow
    
    1. Choose the target artboards as "All" or "Range"
    2. Set the prefix, text source, and suffix
    3. Preview the renamed artboards automatically as settings change
    4. Click OK to confirm, or Cancel to restore the original names
    
    ### Update History
    
    - v1.0 (20250509): Initial version
    - v1.1 (20250512): Improved layer reference and UI adjustments
    */

    var SCRIPT_VERSION = "v1.2.0";

    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var lang = getCurrentLang();

    /* 日英ラベル定義 / Japanese-English label definitions */

    var LABELS = {
        dialogTitle: {
            ja: "アートボード名一括リネーム" + ' ' + SCRIPT_VERSION,
            en: "Smart Artboard Renamer" + ' ' + SCRIPT_VERSION
        },
        prefixLabel: {
            ja: "接頭辞",
            en: "Prefix"
        },
        suffixLabel: {
            ja: "接尾辞",
            en: "Suffix"
        },
        sourceLabel: {
            ja: "参照テキスト",
            en: "Text Source"
        },
        layerOption: {
            ja: "指定レイヤー",
            en: "Selected Layer"
        },
        frontmostOption: {
            ja: "最前面テキスト",
            en: "Frontmost Text"
        },
        fileNameOption: {
            ja: "ファイル名",
            en: "File Name"
        },
        ignoreOption: {
            ja: "使わない",
            en: "None"
        },
        targetLabel: {
            ja: "対象アートボード：",
            en: "Target Artboards:"
        },
        allBoards: {
            ja: "すべて",
            en: "All"
        },
        specificBoards: {
            ja: "指定範囲",
            en: "Range"
        },
        cancelButton: {
            ja: "キャンセル",
            en: "Cancel"
        },
        okButton: {
            ja: "OK",
            en: "OK"
        },
        exampleFormatHint: {
            ja: "例\n1, 01, {A}, {a}\n#FN\n#DT",
            en: "Example\n1, 01, {A}, {a}\n#FN\n#DT"
        },
        exampleFormatToolTip: {
            ja: "#FN：ファイル名\n#DT：日付（YYYYMMDD）\n{A}：英大文字連番\n{a}：英小文字連番",
            en: "#FN: File Name\n#DT: Date (YYYYMMDD)\n{A}: Uppercase letter sequence\n{a}: Lowercase letter sequence"
        },
        alertNoLayer: {
            ja: "指定されたレイヤーが見つからないか、非表示です。",
            en: "The specified layer was not found or is hidden."
        },
        alertNeedSettings: {
            ja: "接頭辞・接尾辞のいずれかを入力するか、参照テキストを指定してください。",
            en: "Enter a prefix or suffix, or choose a text source."
        }
    };

    function restoreOriginalArtboardNames(doc, names) {
        for (var i = 0; i < doc.artboards.length; i++) {
            doc.artboards[i].name = names[i];
        }
    }

    function readDialogSettings(dialogUI) {
        return {
            mode: dialogUI.r1.value ? "frontmost" : (dialogUI.r2.value ? "layer" : (dialogUI.r4.value ? "filename" : "none")),
            prefix: dialogUI.prefixInput.text,
            suffix: dialogUI.suffixInput.text,
            artboardTarget: dialogUI.abAll.value ? "all" : "numbered",
            numberInput: dialogUI.abNumbersInput.text,
            selectedLayerName: dialogUI.r2.value && dialogUI.layerDropdown.selection ? dialogUI.layerDropdown.selection.text : null
        };
    }

    function bindDialogEvents(dialogUI, doc, originalNames) {
        var layerDropdown = dialogUI.layerDropdown;
        var prefixInput = dialogUI.prefixInput;
        var suffixInput = dialogUI.suffixInput;
        var abAll = dialogUI.abAll;
        var abNumbered = dialogUI.abNumbered;
        var abNumbersInput = dialogUI.abNumbersInput;
        var r1 = dialogUI.r1;
        var r2 = dialogUI.r2;
        var r3 = dialogUI.r3;
        var r4 = dialogUI.r4;

        function updatePreview() {
            restoreOriginalArtboardNames(doc, originalNames);
            if (executeRename(doc, readDialogSettings(dialogUI), { silent: true })) {
                app.redraw();
            }
        }

        abNumbered.onClick = function () {
            abNumbersInput.enabled = true;
            updatePreview();
        };
        abAll.onClick = function () {
            abNumbersInput.enabled = false;
            updatePreview();
        };

        r1.onClick = function () {
            layerDropdown.enabled = false;
            updatePreview();
        };
        r2.onClick = function () {
            layerDropdown.enabled = true;
            updatePreview();
        };
        r3.onClick = function () {
            layerDropdown.enabled = false;
            updatePreview();
        };
        r4.onClick = function () {
            layerDropdown.enabled = false;
            updatePreview();
        };
        layerDropdown.onChange = function () {
            if (r2.value) updatePreview();
        };

        prefixInput.onChange = function () {
            updatePreview();
        };
        suffixInput.onChange = function () {
            updatePreview();
        };
        abNumbersInput.onChange = function () {
            if (abNumbered.value) updatePreview();
        };
    }

    function createRenameDialog(doc, originalNames) {
        var dlg = new Window("dialog", LABELS.dialogTitle[lang]);
        dlg.orientation = "column";
        dlg.alignChildren = "fill";

        // 対象アートボード選択（最上部・左右中央）
        var targetOuter = dlg.add("group");
        targetOuter.orientation = "row";
        targetOuter.alignment = "center";

        var targetGroup = targetOuter.add("group");
        targetGroup.orientation = "row";
        targetGroup.alignChildren = ["left", "center"];
        // targetGroup.margins = [0, 0, 0, 0];
        // targetGroup.spacing = 10;

        var targetLabel = targetGroup.add("statictext", undefined, LABELS.targetLabel[lang]);
        targetLabel.alignment = ["left", "center"];

        var abAll = targetGroup.add("radiobutton", undefined, LABELS.allBoards[lang]);
        abAll.alignment = ["left", "center"];

        var abNumbered = targetGroup.add("radiobutton", undefined, LABELS.specificBoards[lang]);
        abNumbered.alignment = ["left", "center"];

        abAll.value = true;
        var abNumbersInput = targetGroup.add("edittext", undefined, "");
        abNumbersInput.alignment = ["left", "center"];
        abNumbersInput.characters = 10;
        abNumbersInput.enabled = false;

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
        var prefixHint = leftCol.add("statictext", undefined, LABELS.exampleFormatHint[lang], {
            multiline: true
        });
        prefixHint.preferredSize.height = 80;
        prefixHint.enabled = true;
        prefixHint.helpTip = LABELS.exampleFormatToolTip[lang];

        // 参照テキスト
        var centerCol = mainGroup.add("panel", undefined, LABELS.sourceLabel[lang]);
        centerCol.orientation = "column";
        centerCol.alignChildren = "left";
        centerCol.margins = [15, 20, 15, 15];
        centerCol.preferredSize.height = 65;

        var r3 = centerCol.add("radiobutton", undefined, LABELS.ignoreOption[lang]);
        var r4 = centerCol.add("radiobutton", undefined, LABELS.fileNameOption[lang]);
        var r1 = centerCol.add("radiobutton", undefined, LABELS.frontmostOption[lang]);
        var r2 = centerCol.add("radiobutton", undefined, LABELS.layerOption[lang]);
        r1.value = true;

        // ポップアップ：全レイヤー（非表示レイヤーは除外）
        var layerDropdown = centerCol.add("dropdownlist", undefined, []);
        layerDropdown.minimumSize.width = 130;
        var allLayers = doc.layers;
        for (var i = 0; i < allLayers.length; i++) {
            if (!allLayers[i].visible) continue;
            layerDropdown.add("item", allLayers[i].name);
        }
        if (layerDropdown.items.length > 0) layerDropdown.selection = 0;
        layerDropdown.enabled = false;

        // 接尾辞
        var rightCol = mainGroup.add("panel", undefined, LABELS.suffixLabel[lang]);
        rightCol.orientation = "column";
        rightCol.alignChildren = "left";
        rightCol.margins = [15, 20, 15, 10];
        var suffixInput = rightCol.add("edittext", undefined, "");
        suffixInput.characters = 12;
        var suffixHint = rightCol.add("statictext", undefined, LABELS.exampleFormatHint[lang], {
            multiline: true
        });
        suffixHint.preferredSize.height = 80;
        suffixHint.enabled = true;
        suffixHint.helpTip = LABELS.exampleFormatToolTip[lang];


        // ボタンエリア（3グループ構成）/ Button area (3-group layout)
        var buttonArea = dlg.add("group");
        buttonArea.orientation = "row";
        buttonArea.alignChildren = ["fill", "center"];
        buttonArea.alignment = "fill";
        // buttonArea.margins = [5, 10, 5, 0];
        // buttonArea.spacing = 10;

        var leftGroup = buttonArea.add("group");
        leftGroup.orientation = "row";
        leftGroup.alignChildren = ["left", "center"];
        var btnCancel = leftGroup.add("button", undefined, LABELS.cancelButton[lang], { name: "cancel" });

        var centerGroup = buttonArea.add("group");
        centerGroup.alignment = ["fill", "fill"];
        centerGroup.minimumSize.width = 0;

        var rightGroup = buttonArea.add("group");
        rightGroup.orientation = "row";
        rightGroup.alignChildren = ["right", "center"];
        var btnOK = rightGroup.add("button", undefined, LABELS.okButton[lang], { name: "ok" });

        return {
            dlg: dlg,
            prefixInput: prefixInput,
            suffixInput: suffixInput,
            r1: r1,
            r2: r2,
            r3: r3,
            r4: r4,
            layerDropdown: layerDropdown,
            abAll: abAll,
            abNumbered: abNumbered,
            abNumbersInput: abNumbersInput,
            btnOK: btnOK
        };
    }

    function showRenameDialog(doc, originalNames) {
        var dialogUI = createRenameDialog(doc, originalNames);
        var dlg = dialogUI.dlg;
        bindDialogEvents(dialogUI, doc, originalNames);

        if (dlg.show() !== 1) {
            return null;
        }

        return readDialogSettings(dialogUI);
    }

    function main() {
        if (app.documents.length === 0) {
            return;
        }

        var doc = app.activeDocument;
        var originalNames = [];
        for (var i = 0; i < doc.artboards.length; i++) {
            originalNames.push(doc.artboards[i].name);
        }

        var dialogResult = showRenameDialog(doc, originalNames);

        if (!dialogResult) {
            restoreOriginalArtboardNames(doc, originalNames);
            return;
        }

        restoreOriginalArtboardNames(doc, originalNames);
        executeRename(doc, dialogResult);
    }

    main();

    function executeRename(doc, settings, options) {
        var mode = settings.mode;
        var prefix = settings.prefix;
        var suffix = settings.suffix;
        var targetType = settings.artboardTarget;
        var numberInput = settings.numberInput;
        var layerName = settings.selectedLayerName;

        var silent = options && options.silent;

        var fileNameText = doc.name.replace(/\.[^\.]+$/, "");

        var textFrames = [];
        var artboardTextMap = {};
        if (mode === "layer") {
            var namingLayer = getNamingLayerByName(doc, layerName);
            if (!namingLayer || !namingLayer.visible) {
                if (!silent) {
                    alert(LABELS.alertNoLayer[lang]);
                }
                return false;
            }
            textFrames = getTextFramesInLayer(namingLayer);
        } else if (mode === "frontmost") {
            textFrames = getFrontmostTextFramesPerArtboard(doc);
        } else if (mode === "filename") {
            for (var ai = 0; ai < doc.artboards.length; ai++) {
                artboardTextMap[ai] = [fileNameText];
            }
        }

        if (mode === "none" && prefix === "" && suffix === "") {
            if (!silent) {
                alert(LABELS.alertNeedSettings[lang]);
            }
            return false;
        }

        var targetIndices = getTargetArtboardIndices(doc.artboards.length, targetType, numberInput);
        if (mode !== "filename") {
            artboardTextMap = mapTextToArtboards(textFrames, doc.artboards, mode);
        }
        renameArtboards(doc.artboards, artboardTextMap, prefix, suffix, targetIndices);
        return true;
    }

    function getReservedArtboardNames(artboards, targetIndices) {
        var reserved = [];
        for (var i = 0; i < artboards.length; i++) {
            if (!contains(targetIndices, i)) {
                reserved.push(artboards[i].name);
            }
        }
        return reserved;
    }

    function getNamingLayerByName(doc, name) {
        for (var i = 0; i < doc.layers.length; i++) {
            if (doc.layers[i].name === name) return doc.layers[i];
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

    function numberToLetters(value, upperCase) {
        var result = "";
        var n = value;

        while (n > 0) {
            n--;
            result = String.fromCharCode((n % 26) + (upperCase ? 65 : 97)) + result;
            n = Math.floor(n / 26);
        }

        return result;
    }

    function formatWithAlphaNumeric(template, index) {
        var fileName = app.activeDocument.name.replace(/\.[^\.]+$/, "");
        var now = new Date();
        var dateString = now.getFullYear().toString() +
            ("0" + (now.getMonth() + 1)).slice(-2) +
            ("0" + now.getDate()).slice(-2);

        var result = template;

        result = result.replace(/\{A\}/g, numberToLetters(index, true));
        result = result.replace(/\{a\}/g, numberToLetters(index, false));
        result = result.replace(/#FN/g, fileName);
        result = result.replace(/#DT/g, dateString);

        var match = result.match(/\d+/);
        if (match) {
            var token = match[0];
            var base = parseInt(token, 10);
            var value = base + index - 1;
            var replacement = (token.charAt(0) === "0" && token.length > 1)
                ? ("0000000000" + value).slice(-token.length)
                : value.toString();
            result = result.replace(token, replacement);
        }

        return result;
    }

    function renameArtboards(artboards, map, prefixInput, suffixInput, targetIndices) {
        var usedNames = getReservedArtboardNames(artboards, targetIndices);
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
                } catch (e) {
                    $.writeln("[SmartArtboardRenamer] Failed to rename artboard index " + i + " to '" + uniqueName + "': " + e);
                }
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
})();