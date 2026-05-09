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

    ### Update History
    
    - v1.0 (20250509): Initial version
    - v1.1 (20250512): Improved layer reference and UI adjustments
    */

    // =========================================
    // バージョンと UI 共通設定
    // =========================================

    var SCRIPT_VERSION = "v1.2.0";

    var PANEL_MARGINS = [15, 20, 15, 10];

    /* パネルの共通プロパティを設定（縦並び・左揃え・横 fill・共通マージン）/ Apply common panel properties (column, left-aligned, horizontal fill, shared margins) */
    function setupPanel(panel, spacing) {
        panel.orientation = "column";
        panel.alignChildren = "left";
        panel.alignment = "fill";
        panel.margins = PANEL_MARGINS;
        if (typeof spacing === "number") {
            panel.spacing = spacing;
        }
    }

    // =========================================
    // ローカライズ
    // =========================================

    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var lang = getCurrentLang();

    /* ラベル取得ヘルパー / Localized label helper */
    function L(key) {
        return LABELS[key][lang];
    }

    /* 日英ラベル定義 / Japanese-English label definitions */

    var LABELS = {
        dialogTitle: {
            ja: "アートボード名一括リネーム",
            en: "Smart Artboard Renamer"
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

    // =========================================
    // アートボード名のキャプチャと復元
    // =========================================

    function captureArtboardNames(doc) {
        var names = [];
        for (var i = 0; i < doc.artboards.length; i++) {
            names.push(doc.artboards[i].name);
        }
        return names;
    }

    function restoreOriginalArtboardNames(doc, names) {
        for (var i = 0; i < doc.artboards.length; i++) {
            doc.artboards[i].name = names[i];
        }
    }

    function getDocBaseName(doc) {
        return doc.name.replace(/\.[^\.]+$/, "");
    }

    // =========================================
    // ダイアログ
    // =========================================

    function readDialogSettings(dialogUI) {
        var mode = dialogUI.radioFrontmost.value ? "frontmost" :
            dialogUI.radioLayer.value ? "layer" :
                dialogUI.radioFilename.value ? "filename" : "none";
        return {
            mode: mode,
            prefix: dialogUI.prefixInput.text,
            suffix: dialogUI.suffixInput.text,
            artboardTarget: dialogUI.targetAllRadio.value ? "all" : "range",
            rangeText: dialogUI.targetRangeInput.text,
            selectedLayerName: (mode === "layer" && dialogUI.layerDropdown.selection) ? dialogUI.layerDropdown.selection.text : null
        };
    }

    /* ラジオ／入力欄の変更でリアルタイムプレビューを再実行 / Wire dialog widgets so changes trigger a live preview rerun */
    function bindDialogEvents(dialogUI, doc, originalNames) {
        var layerDropdown = dialogUI.layerDropdown;
        var targetAllRadio = dialogUI.targetAllRadio;
        var targetRangeRadio = dialogUI.targetRangeRadio;
        var targetRangeInput = dialogUI.targetRangeInput;
        var radioLayer = dialogUI.radioLayer;

        function updatePreview() {
            restoreOriginalArtboardNames(doc, originalNames);
            if (executeRename(doc, readDialogSettings(dialogUI), { silent: true })) {
                app.redraw();
            }
        }

        function setLayerEnabled(enabled) {
            return function () {
                layerDropdown.enabled = enabled;
                updatePreview();
            };
        }
        function setRangeEnabled(enabled) {
            return function () {
                targetRangeInput.enabled = enabled;
                updatePreview();
            };
        }

        dialogUI.radioFrontmost.onClick = setLayerEnabled(false);
        dialogUI.radioLayer.onClick = setLayerEnabled(true);
        dialogUI.radioNone.onClick = setLayerEnabled(false);
        dialogUI.radioFilename.onClick = setLayerEnabled(false);

        targetAllRadio.onClick = setRangeEnabled(false);
        targetRangeRadio.onClick = setRangeEnabled(true);

        layerDropdown.onChange = function () { if (radioLayer.value) updatePreview(); };
        dialogUI.prefixInput.onChange = updatePreview;
        dialogUI.suffixInput.onChange = updatePreview;
        targetRangeInput.onChange = function () { if (targetRangeRadio.value) updatePreview(); };
    }

    /* 接頭辞・接尾辞パネル / Affix panel (prefix or suffix) */
    function addAffixPanel(parent, labelKey, makeActive) {
        var panel = parent.add("panel", undefined, L(labelKey));
        setupPanel(panel);
        var input = panel.add("edittext", undefined, "");
        input.characters = 12;
        if (makeActive) input.active = true;
        var hint = panel.add("statictext", undefined, L("exampleFormatHint"), { multiline: true });
        hint.preferredSize.height = 80;
        hint.helpTip = L("exampleFormatToolTip");
        return input;
    }

    /* リネーム用ダイアログを構築し、各 UI 要素の参照を返す / Build the rename dialog and return references to its widgets */
    function createRenameDialog(doc) {
        /* ダイアログ枠 / Dialog skeleton */
        var dialog = new Window("dialog", L("dialogTitle") + " " + SCRIPT_VERSION);
        dialog.alignChildren = "fill";

        /* 対象アートボード選択（最上部・左右中央）/ Target artboard selection (top, centered) */
        var targetOuter = dialog.add("group");
        targetOuter.orientation = "row";
        targetOuter.alignment = "center";

        var targetGroup = targetOuter.add("group");
        targetGroup.orientation = "row";
        targetGroup.alignChildren = ["left", "center"];

        var targetLabel = targetGroup.add("statictext", undefined, L("targetLabel"));
        targetLabel.alignment = ["left", "center"];

        var targetAllRadio = targetGroup.add("radiobutton", undefined, L("allBoards"));
        targetAllRadio.alignment = ["left", "center"];
        targetAllRadio.value = true;

        var targetRangeRadio = targetGroup.add("radiobutton", undefined, L("specificBoards"));
        targetRangeRadio.alignment = ["left", "center"];

        var targetRangeInput = targetGroup.add("edittext", undefined, "");
        targetRangeInput.alignment = ["left", "center"];
        targetRangeInput.characters = 10;
        targetRangeInput.enabled = false;

        /* 入力エリア / Input area */
        var bodyRow = dialog.add("group");
        bodyRow.orientation = "row";
        bodyRow.alignChildren = ["top", "fill"];

        /* 接頭辞 / Prefix */
        var prefixInput = addAffixPanel(bodyRow, "prefixLabel", true);

        /* 参照テキスト / Text source */
        var sourcePanel = bodyRow.add("panel", undefined, L("sourceLabel"));
        setupPanel(sourcePanel);
        sourcePanel.preferredSize.height = 65;

        var radioNone = sourcePanel.add("radiobutton", undefined, L("ignoreOption"));
        var radioFilename = sourcePanel.add("radiobutton", undefined, L("fileNameOption"));
        var radioFrontmost = sourcePanel.add("radiobutton", undefined, L("frontmostOption"));
        var radioLayer = sourcePanel.add("radiobutton", undefined, L("layerOption"));
        radioFrontmost.value = true;

        /* レイヤードロップダウン（非表示レイヤーは除外）/ Layer dropdown (skip hidden layers) */
        var layerDropdown = sourcePanel.add("dropdownlist", undefined, []);
        layerDropdown.minimumSize.width = 130;
        var allLayers = doc.layers;
        for (var i = 0; i < allLayers.length; i++) {
            if (!allLayers[i].visible) continue;
            layerDropdown.add("item", allLayers[i].name);
        }
        if (layerDropdown.items.length > 0) layerDropdown.selection = 0;
        layerDropdown.enabled = false;

        /* 接尾辞 / Suffix */
        var suffixInput = addAffixPanel(bodyRow, "suffixLabel", false);

        /* ボタンエリア（3グループ構成）/ Button area (3-group layout) */
        var buttonArea = dialog.add("group");
        buttonArea.orientation = "row";
        buttonArea.alignChildren = ["fill", "center"];
        buttonArea.alignment = "fill";

        var cancelGroup = buttonArea.add("group");
        cancelGroup.orientation = "row";
        cancelGroup.alignChildren = ["left", "center"];
        var cancelButton = cancelGroup.add("button", undefined, L("cancelButton"), { name: "cancel" });

        var spacerGroup = buttonArea.add("group");
        spacerGroup.alignment = ["fill", "fill"];
        spacerGroup.minimumSize.width = 0;

        var okGroup = buttonArea.add("group");
        okGroup.orientation = "row";
        okGroup.alignChildren = ["right", "center"];
        var okButton = okGroup.add("button", undefined, L("okButton"), { name: "ok" });

        return {
            dialog: dialog,
            prefixInput: prefixInput,
            suffixInput: suffixInput,
            radioFrontmost: radioFrontmost,
            radioLayer: radioLayer,
            radioNone: radioNone,
            radioFilename: radioFilename,
            layerDropdown: layerDropdown,
            targetAllRadio: targetAllRadio,
            targetRangeRadio: targetRangeRadio,
            targetRangeInput: targetRangeInput,
            okButton: okButton,
            cancelButton: cancelButton
        };
    }

    function showRenameDialog(doc, originalNames) {
        var dialogUI = createRenameDialog(doc);
        bindDialogEvents(dialogUI, doc, originalNames);

        if (dialogUI.dialog.show() !== 1) return null;
        return readDialogSettings(dialogUI);
    }

    // =========================================
    // エントリポイント
    // =========================================

    /* 元のアートボード名を退避→ダイアログ表示→確定なら本リネームを実行 / Capture original names, show the dialog, and commit the rename when confirmed */
    function main() {
        if (app.documents.length === 0) return;

        var doc = app.activeDocument;
        var originalNames = captureArtboardNames(doc);

        var dialogResult = showRenameDialog(doc, originalNames);
        restoreOriginalArtboardNames(doc, originalNames);
        if (dialogResult) executeRename(doc, dialogResult);
    }

    main();

    // =========================================
    // リネーム実行
    // =========================================

    /* モード別にアートボード→参照テキストのマップを構築 / Build per-artboard reference-text map */
    function buildArtboardTextMap(doc, settings, silent) {
        var mode = settings.mode;
        if (mode === "filename") {
            var fileNameText = getDocBaseName(doc);
            var map = {};
            for (var ai = 0; ai < doc.artboards.length; ai++) {
                map[ai] = [fileNameText];
            }
            return map;
        }
        if (mode === "layer") {
            var namingLayer = findLayerByName(doc, settings.selectedLayerName);
            if (!namingLayer || !namingLayer.visible) {
                if (!silent) alert(L("alertNoLayer"));
                return null;
            }
            return mapTextToArtboards(getTextFramesInLayer(namingLayer), doc.artboards);
        }
        if (mode === "frontmost") {
            return mapTextToArtboards(getFrontmostTextFramesPerArtboard(doc), doc.artboards);
        }
        return {};
    }

    /* 設定を検証してテキストマップを生成、対象アートボードへ実リネームを適用 / Validate settings, build the artboard text map, then apply the rename */
    function executeRename(doc, settings, options) {
        var silent = options && options.silent;

        if (settings.mode === "none" && settings.prefix === "" && settings.suffix === "") {
            if (!silent) alert(L("alertNeedSettings"));
            return false;
        }

        var artboardTextMap = buildArtboardTextMap(doc, settings, silent);
        if (artboardTextMap === null) return false;

        var targetIndices = getTargetArtboardIndices(doc.artboards.length, settings.artboardTarget, settings.rangeText);
        renameArtboards(doc.artboards, artboardTextMap, settings.prefix, settings.suffix, targetIndices);
        return true;
    }

    /* リネーム対象外アートボードの名前を予約名として収集（重複回避時の比較対象）/ Collect names of non-target artboards to reserve them when checking for uniqueness */
    function getReservedArtboardNames(artboards, targetIndices) {
        var reserved = [];
        for (var i = 0; i < artboards.length; i++) {
            if (!contains(targetIndices, i)) {
                reserved.push(artboards[i].name);
            }
        }
        return reserved;
    }

    // =========================================
    // レイヤー／テキストフレームの探索
    // =========================================

    function findLayerByName(doc, name) {
        for (var i = 0; i < doc.layers.length; i++) {
            if (doc.layers[i].name === name) return doc.layers[i];
        }
        return null;
    }

    /* 指定レイヤー配下（サブレイヤー含む）から非表示を除いてテキストフレームを再帰収集 / Recursively gather text frames in the layer and its visible sublayers */
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

    function getTextCenter(textFrame) {
        var bounds = textFrame.visibleBounds;
        return [(bounds[0] + bounds[2]) / 2, (bounds[1] + bounds[3]) / 2];
    }

    /* 点 (cx, cy) がアートボード矩形 [left, top, right, bottom] の内側か判定（top > bottom の Illustrator 座標系前提）/ Test whether (cx, cy) lies inside the rect [left, top, right, bottom] (Illustrator's top > bottom convention) */
    function isCenterInArtboard(cx, cy, rect) {
        return cx >= rect[0] && cx <= rect[2] && cy <= rect[1] && cy >= rect[3];
    }

    // =========================================
    // 対象アートボード範囲の解析
    // =========================================

    /* "1, 3-5" 形式の文字列を 0 始まりインデックス配列に変換 / Parse a "1, 3-5" style range string into a 0-based index array */
    function parseArtboardRange(rangeText) {
        var result = [];
        var parts = rangeText.split(",");
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

    /* 対象モード（"all" or "range"）に応じてリネーム対象アートボードのインデックス配列を返す / Return the 0-based index list of artboards to rename, based on target mode ("all" or "range") */
    function getTargetArtboardIndices(total, targetType, rangeText) {
        if (targetType === "all") {
            var list = [];
            for (var i = 0; i < total; i++) list.push(i);
            return list;
        }
        return parseArtboardRange(rangeText);
    }

    // =========================================
    // テキスト→アートボード対応
    // =========================================

    /* 各テキストフレームを、その中心点を含むアートボードへ振り分けて配列に格納 / Group text frames under the artboard whose rect contains their center point */
    function mapTextToArtboards(textFrames, artboards) {
        var map = {};
        for (var i = 0; i < textFrames.length; i++) {
            var textFrame = textFrames[i];
            var center = getTextCenter(textFrame);
            for (var j = 0; j < artboards.length; j++) {
                if (isCenterInArtboard(center[0], center[1], artboards[j].artboardRect)) {
                    if (!map[j]) map[j] = [];
                    map[j].push(textFrame.contents.replace(/[\r\n\t]/g, ""));
                    break;
                }
            }
        }
        return map;
    }

    // =========================================
    // 連番・テンプレ展開ユーティリティ
    // =========================================

    function contains(array, value) {
        for (var i = 0; i < array.length; i++) {
            if (array[i] === value) return true;
        }
        return false;
    }

    /* 既出名と衝突する場合 "_2", "_3" … を末尾に付けて一意な名前を生成 / Generate a name unique among usedNames by appending "_2", "_3", ... when needed */
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

    function todayYYYYMMDD() {
        var d = new Date();
        return d.getFullYear().toString() +
            ("0" + (d.getMonth() + 1)).slice(-2) +
            ("0" + d.getDate()).slice(-2);
    }

    /* テンプレ内のトークンを置換 / Replace literal tokens in template */
    function applyTokens(template, index, fileName, dateString) {
        return template
            .replace(/\{A\}/g, numberToLetters(index, true))
            .replace(/\{a\}/g, numberToLetters(index, false))
            .replace(/#FN/g, fileName)
            .replace(/#DT/g, dateString);
    }

    /* 文字列内の最初の数値トークンに offset を加算（先頭ゼロ桁数を維持）/ Increment first numeric token, preserving zero-padded width */
    function incrementNumericToken(text, offset) {
        var match = text.match(/\d+/);
        if (!match) return text;
        var token = match[0];
        var value = parseInt(token, 10) + offset;
        var replacement = (token.charAt(0) === "0" && token.length > 1)
            ? ("0000000000" + value).slice(-token.length)
            : value.toString();
        return text.replace(token, replacement);
    }

    /* テンプレ内のトークンを展開し、最初の数値トークンに連番オフセット (index-1) を加算 / Expand template tokens and increment the first numeric token by (index - 1) */
    function formatWithAlphaNumeric(template, index) {
        var withTokens = applyTokens(template, index, getDocBaseName(app.activeDocument), todayYYYYMMDD());
        return incrementNumericToken(withTokens, index - 1);
    }

    // =========================================
    // アートボード名の生成と適用
    // =========================================

    /* 1 アートボード分の合成名を組み立て（全要素空なら null）/ Build a single artboard name (null when all parts empty) */
    function buildArtboardName(textMap, index, seq, prefixTemplate, suffixTemplate) {
        var prefix = formatWithAlphaNumeric(prefixTemplate, seq);
        var suffix = formatWithAlphaNumeric(suffixTemplate, seq);
        var textPart = (textMap[index] && textMap[index].length > 0) ? textMap[index].join(" ") : "";
        if (!prefix && !suffix && !textPart) return null;
        return prefix + textPart + suffix;
    }

    /* 対象アートボードを順に走査し、合成名を作って重複回避のうえ実リネーム（連番は実際にリネームしたものだけ進める）/ Walk target artboards, build the composed name, dedupe, then rename (sequence advances only on actual renames) */
    function renameArtboards(artboards, textMap, prefixTemplate, suffixTemplate, targetIndices) {
        var usedNames = getReservedArtboardNames(artboards, targetIndices);
        var seq = 1;
        for (var i = 0; i < artboards.length; i++) {
            if (!contains(targetIndices, i)) continue;
            var baseName = buildArtboardName(textMap, i, seq, prefixTemplate, suffixTemplate);
            if (baseName === null) continue;
            var uniqueName = generateUniqueName(baseName, usedNames);
            usedNames.push(uniqueName);
            try {
                artboards[i].name = uniqueName;
            } catch (e) {
                $.writeln("[SmartArtboardRenamer] Failed to rename artboard index " + i + " to '" + uniqueName + "': " + e);
            }
            seq++;
        }
    }

    // =========================================
    // 最前面テキストフレーム検出
    // =========================================

    /* 指定アートボード矩形に重なる最前面のテキストフレームを 1 つ探す / Find the frontmost text frame whose center falls inside the rect */
    function findFrontmostTextFrameInArtboard(doc, artboardRect) {
        for (var li = 0; li < doc.layers.length; li++) {
            var layer = doc.layers[li];
            if (!layer.visible) continue;
            var items = layer.pageItems;
            for (var pi = 0; pi < items.length; pi++) {
                var item = items[pi];
                if (item.typename !== "TextFrame") continue;
                var center = getTextCenter(item);
                if (isCenterInArtboard(center[0], center[1], artboardRect)) return item;
            }
        }
        return null;
    }

    /* 各アートボードについて、レイヤー順（最前面側から）でアートボード矩形に重なる最初のテキストフレームを集める / For each artboard, collect the first text frame in front-to-back layer order whose center falls inside the artboard rect */
    function getFrontmostTextFramesPerArtboard(doc) {
        var result = [];
        for (var i = 0; i < doc.artboards.length; i++) {
            var found = findFrontmostTextFrameInArtboard(doc, doc.artboards[i].artboardRect);
            if (found) result.push(found);
        }
        return result;
    }
})();