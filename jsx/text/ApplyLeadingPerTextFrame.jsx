#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);


(function () {

    var SCRIPT_VERSION = "v1.0";

    /*
     * 概要 / Overview
     * 選択されたテキストフレームに対して、行送り、行送りの基準、段落前後のアキ、自動行送り量をダイアログで確認・調整して適用します。
     * ダイアログボックスを開く際、行送り値・行送りの基準（autoLeadingType）・段落前後のアキ・自動行送り量の代表値を参照して初期表示に反映します。
     * 行送りの現在値が 110% / 125% / 150% / 自動 に一致する場合は対応するラジオボタンを選択し、それ以外は［その他］を選択します。
     * 行送り値を変更すると［その他］、自動行送り量を変更すると［自動］が選択され、［自動］選択時にも計算上の行送り値を表示します。
     * 段落がひとつだけの場合は、初期状態で［共通］を選択します。
     * テキストの一部（TextRange）が選択されている場合、その親のテキストフレーム（TextFrame）に正規化します。
     * 非テキストオブジェクトや処理不能な要素はスキップし、同じテキストフレームは重複して処理しません。
     * プレビュー時はスナップショットから復元して再適用し、エラーは最初の1回のみ通知されます。
     * アラートメッセージとUIラベルはローカライズ定義（LABELS）を通して取得します。
     * 更新日：2026-04-22
     */


    var LINE_FONT_SIZE_SAMPLE_COUNT = 5; // 行頭で参照する文字数 / Number of leading characters to sample per line

    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }

    var lang = getCurrentLang();

    var LABELS = {
        alertOpenDocument: {
            ja: "ドキュメントを開いてから実行してください。",
            en: "Please open a document before running this script."
        },
        alertSelectTextObject: {
            ja: "テキストオブジェクトを選択してください。",
            en: "Please select a text object."
        },
        alertLeadingApplyError: {
            ja: "行送りの設定中にエラーが発生しました。",
            en: "An error occurred while applying leading."
        },
        alertNoProcessableTextFrame: {
            ja: "処理可能なテキストフレームが選択されていません。",
            en: "No processable text frames are selected."
        },
        dialogTitle: {
            ja: "行送りの設定",
            en: "Leading Settings"
        },
        dialogPanel: {
            ja: "行送り",
            en: "Leading settings"
        },
        dialogPanelType: {
            ja: "行送りの基準",
            en: "Leading basis"
        },
        typeTopToTop: {
            ja: "仮想ボディの上基準",
            en: "Top-to-top (virtual body)"
        },
        typeBottomToBottom: {
            ja: "欧文ベースライン基準",
            en: "Baseline-to-baseline"
        },
        dialogPanelSpace: {
            ja: "段落前後のアキ",
            en: "Paragraph spacing"
        },
        labelSpaceBefore: {
            ja: "段落前：",
            en: "Space before:"
        },
        labelSpaceAfter: {
            ja: "段落後：",
            en: "Space after:"
        },
        btnOk: {
            ja: "OK",
            en: "OK"
        },
        btnCancel: {
            ja: "キャンセル",
            en: "Cancel"
        },
        leadingAuto: {
            ja: "自動",
            en: "Auto"
        },
        leadingOther: {
            ja: "その他",
            en: "Other"
        },
        modePerLine: {
            ja: "行ごとに設定",
            en: "Per line"
        },
        modeCommon: {
            ja: "共通",
            en: "Common value"
        }
    };

    function L(key) {
        return LABELS[key][lang];
    }

    var LEADING_CHOICES = [
        { label: "110%", ratio: 1.1, auto: false, other: false },
        { label: "125%", ratio: 1.25, auto: false, other: false },
        { label: "150%", ratio: 1.5, auto: false, other: false },
        { label: L("leadingOther"), ratio: undefined, auto: false, other: true },
        { label: L("leadingAuto"), ratio: null, auto: true, other: false }
    ];
    function findMatchingLeadingChoiceIndex(isAutoLeading, leadingValuePt, baseFontSizePt) {
        if (isAutoLeading) {
            for (var i = 0; i < LEADING_CHOICES.length; i++) {
                if (LEADING_CHOICES[i].auto) return i;
            }
        }

        if (!isNaN(leadingValuePt) && !isNaN(baseFontSizePt) && baseFontSizePt > 0) {
            var ratio = leadingValuePt / baseFontSizePt;
            for (var j = 0; j < LEADING_CHOICES.length; j++) {
                var choice = LEADING_CHOICES[j];
                if (!choice.auto && !choice.other && typeof choice.ratio === "number") {
                    if (Math.abs(choice.ratio - ratio) < 0.0001) {
                        return j;
                    }
                }
            }
        }

        for (var k = 0; k < LEADING_CHOICES.length; k++) {
            if (LEADING_CHOICES[k].other) return k;
        }
        return 0;
    }

    function getInitialLeadingChoiceState(frames) {
        for (var i = 0; i < frames.length; i++) {
            try {
                var lines = frames[i].lines;
                if (!lines || lines.length === 0 || lines[0].characters.length === 0) continue;

                var firstChar = lines[0].characters[0];
                var isAutoLeading = firstChar.characterAttributes.autoLeading;
                var leadingValuePt = firstChar.characterAttributes.leading;
                var baseFontSizePt = getBaseFontSizeFromLine(lines[0]);
                var selectedIndex = findMatchingLeadingChoiceIndex(isAutoLeading, leadingValuePt, baseFontSizePt);

                return {
                    selectedIndex: selectedIndex,
                    isOther: !!LEADING_CHOICES[selectedIndex].other
                };
            } catch (e) { }
        }
        return {
            selectedIndex: 0,
            isOther: false
        };
    }

    function shouldUseCommonModeInitially(frames) {
        var totalParagraphCount = 0;
        for (var i = 0; i < frames.length; i++) {
            try {
                var paragraphs = frames[i].paragraphs;
                if (paragraphs && paragraphs.length > 0) {
                    totalParagraphCount += paragraphs.length;
                    if (totalParagraphCount > 1) {
                        return false;
                    }
                }
            } catch (e) { }
        }
        return totalParagraphCount === 1;
    }

    var PANEL_MARGINS = [15, 20, 15, 10];

    var UNIT_MAP = {
        0: "in",
        1: "mm",
        2: "pt",
        3: "pica",
        4: "cm",
        6: "px",
        7: "ft/in",
        8: "m",
        9: "yd",
        10: "ft"
    };

    var UNIT_TO_PT = {
        0: 72,
        1: 2.8346456692913386,
        2: 1,
        3: 12,
        4: 28.346456692913386,
        5: 0.7086614173228346,
        6: 1,
        7: 72,
        8: 2834.6456692913386,
        9: 2592,
        10: 864
    };

    function getUnitLabel(code, prefKey) {
        if (code === 5) {
            var hKeys = {
                "text/asianunits": true,
                "rulerType": true,
                "strokeUnits": true
            };
            return hKeys[prefKey] ? "H" : "Q";
        }
        return UNIT_MAP[code] || "pt";
    }

    function getTextUnit() {
        var code = 2;
        try {
            code = app.preferences.getIntegerPreference("text/units");
        } catch (e) { }
        return {
            code: code,
            label: getUnitLabel(code, "text/units"),
            factor: UNIT_TO_PT[code] || 1
        };
    }

    function getLeadingTypeChoices() {
        return [
            { label: L("typeTopToTop"), type: AutoLeadingType.TOPTOTOP },
            { label: L("typeBottomToBottom"), type: AutoLeadingType.BOTTOMTOBOTTOM }
        ];
    }

    function changeValueByArrowKey(editText, allowNegative, onUpdate, decimals) {
        function roundToStep(value, step) {
            return Math.round(value / step) * step;
        }

        editText.addEventListener("keydown", function (event) {
            if (event.keyName !== "Up" && event.keyName !== "Down") return;

            var value = Number(editText.text);
            if (isNaN(value)) return;

            var keyboard = ScriptUI.environment.keyboardState;
            var step = 1;

            if (keyboard.shiftKey) {
                step = 10;
            } else if (keyboard.altKey) {
                step = 0.1;
            }

            if (keyboard.shiftKey) {
                value = roundToStep(value, step);
            }

            if (event.keyName === "Up") {
                value += step;
            } else {
                value -= step;
            }

            if (keyboard.altKey) {
                value = Math.round(value * 10) / 10;
            } else {
                value = Math.round(value);
            }

            if (!allowNegative && value < 0) value = 0;

            event.preventDefault();

            if (typeof decimals === "number" && decimals >= 0) {
                var pow = Math.pow(10, decimals);
                value = Math.round(value * pow) / pow;
                editText.text = value.toFixed(decimals);
            } else {
                editText.text = String(value);
            }

            if (typeof onUpdate === "function") onUpdate();
        });
    }

    function createDialogWindow() {
        var dlg = new Window("dialog", L("dialogTitle") + " " + SCRIPT_VERSION);
        dlg.orientation = "column";
        dlg.alignChildren = "fill";
        dlg.margins = 16;
        dlg.spacing = 12;
        return dlg;
    }

    function buildModeGroup(parent, useCommonInitially) {
        var modeGroup = parent.add("group");
        modeGroup.orientation = "row";
        modeGroup.alignChildren = "center";
        modeGroup.alignment = "center";
        modeGroup.spacing = 12;
        modeGroup.margins = [0, 0, 0, 4];

        var perLineRb = modeGroup.add("radiobutton", undefined, L("modePerLine"));
        var commonRb = modeGroup.add("radiobutton", undefined, L("modeCommon"));
        if (useCommonInitially) {
            commonRb.value = true;
        } else {
            perLineRb.value = true;
        }

        return {
            perLineRb: perLineRb,
            commonRb: commonRb,
            isCommon: function () { return commonRb.value; }
        };
    }

    function buildLeadingPanel(parent, defaultAutoAmount, initialLeadingValuePt, initialLeadingChoiceState) {
        var unit = getTextUnit();
        var panel = parent.add("panel", undefined, L("dialogPanel"));
        panel.orientation = "column";
        panel.alignChildren = "left";
        panel.margins = PANEL_MARGINS;
        panel.spacing = 8;

        var contentGroup = panel.add("group");
        contentGroup.orientation = "row";
        contentGroup.alignChildren = ["left", "top"];
        contentGroup.spacing = 25;

        var leftColumnGroup = contentGroup.add("group");
        leftColumnGroup.orientation = "column";
        leftColumnGroup.alignChildren = "left";
        leftColumnGroup.spacing = 0;

        var leadingGroup = leftColumnGroup.add("group");
        var initialLeadingDisplay = (isNaN(initialLeadingValuePt) || initialLeadingValuePt === null) ? "" : (Math.round((initialLeadingValuePt / unit.factor) * 10) / 10).toFixed(1);
        var leadingInput = leadingGroup.add("edittext", undefined, initialLeadingDisplay);
        leadingInput.characters = 5;
        leadingGroup.add("statictext", undefined, unit.label);

        var radioGroup = contentGroup.add("group");
        radioGroup.orientation = "column";
        radioGroup.alignChildren = "left";
        radioGroup.spacing = 6;

        var radios = [];
        var defaultIndex = (initialLeadingChoiceState && typeof initialLeadingChoiceState.selectedIndex === "number")
            ? initialLeadingChoiceState.selectedIndex
            : 0;
        var autoInput = null;
        for (var i = 0; i < LEADING_CHOICES.length; i++) {
            var choice = LEADING_CHOICES[i];
            var rowGroup = radioGroup.add("group");
            rowGroup.orientation = "row";
            rowGroup.alignChildren = "center";
            rowGroup.spacing = 6;

            var rb = rowGroup.add("radiobutton", undefined, choice.label);
            radios.push(rb);

            if (choice.auto) {
                autoInput = rowGroup.add("edittext", undefined, String(Math.round(defaultAutoAmount)));
                autoInput.characters = 3;
                rowGroup.add("statictext", undefined, "%");
            }
        }
        radios[defaultIndex].value = true;

        function selectExclusiveRadio(index) {
            for (var i = 0; i < radios.length; i++) {
                radios[i].value = (i === index);
            }
        }

        return {
            radios: radios,
            autoInput: autoInput,
            leadingInput: leadingInput,
            unit: unit,
            currentRatio: function () {
                for (var i = 0; i < radios.length; i++) {
                    if (radios[i].value) return LEADING_CHOICES[i].ratio;
                }
                return LEADING_CHOICES[0].ratio;
            },
            currentChoiceIndex: function () {
                for (var i = 0; i < radios.length; i++) {
                    if (radios[i].value) return i;
                }
                return 0;
            },
            selectChoiceByIndex: function (index) {
                if (index >= 0 && index < radios.length) {
                    selectExclusiveRadio(index);
                }
            },
            currentIsOther: function () {
                var index = this.currentChoiceIndex();
                return !!LEADING_CHOICES[index].other;
            },
            currentAutoAmount: function () {
                var n = parseInt(autoInput.text, 10);
                return isNaN(n) ? Math.round(defaultAutoAmount) : n;
            },
            currentLeadingPt: function () {
                var n = parseFloat(leadingInput.text);
                if (isNaN(n)) return NaN;
                return n * unit.factor;
            },
            setLeadingDisplayPt: function (pt) {
                if (isNaN(pt) || pt === null) {
                    leadingInput.text = "";
                    return;
                }
                leadingInput.text = (Math.round((pt / unit.factor) * 10) / 10).toFixed(1);
            }
        };
    }

    function buildLeadingTypePanel(parent, defaultType) {
        var typeChoices = getLeadingTypeChoices();
        var container;

        if (lang === "ja") {
            container = parent.add("panel", undefined, L("dialogPanelType"));
            container.orientation = "column";
            container.alignChildren = "left";
            container.margins = PANEL_MARGINS;
            container.spacing = 8;
        } else {
            container = parent.add("group");
            container.orientation = "column";
            container.alignChildren = "left";
            container.spacing = 8;
            container.margins = [0, 0, 0, 0];
        }

        var radios = [];
        var defaultIndex = 0;
        for (var t = 0; t < typeChoices.length; t++) {
            var trb = container.add("radiobutton", undefined, typeChoices[t].label);
            radios.push(trb);
            if (defaultType !== undefined && defaultType !== null && typeChoices[t].type === defaultType) {
                defaultIndex = t;
            }
        }
        radios[defaultIndex].value = true;

        return {
            radios: radios,
            currentType: function () {
                for (var i = 0; i < radios.length; i++) {
                    if (radios[i].value) return typeChoices[i].type;
                }
                return typeChoices[0].type;
            }
        };
    }

    function buildSpacePanel(parent, initialSpaceBeforePt, initialSpaceAfterPt) {
        var unit = getTextUnit();
        var panel = parent.add("panel", undefined, L("dialogPanelSpace"));
        panel.orientation = "column";
        panel.alignChildren = "left";
        panel.margins = PANEL_MARGINS;
        panel.spacing = 8;

        var initialSpaceBeforeDisplay = (isNaN(initialSpaceBeforePt) || initialSpaceBeforePt === null) ? "0" : (Math.round((initialSpaceBeforePt / unit.factor) * 10) / 10).toFixed(1);
        var initialSpaceAfterDisplay = (isNaN(initialSpaceAfterPt) || initialSpaceAfterPt === null) ? "0" : (Math.round((initialSpaceAfterPt / unit.factor) * 10) / 10).toFixed(1);

        var beforeGroup = panel.add("group");
        beforeGroup.add("statictext", undefined, L("labelSpaceBefore"));
        var beforeInput = beforeGroup.add("edittext", undefined, initialSpaceBeforeDisplay);
        beforeInput.characters = 4;
        beforeGroup.add("statictext", undefined, unit.label);

        var afterGroup = panel.add("group");
        afterGroup.add("statictext", undefined, L("labelSpaceAfter"));
        var afterInput = afterGroup.add("edittext", undefined, initialSpaceAfterDisplay);
        afterInput.characters = 4;
        afterGroup.add("statictext", undefined, unit.label);

        function parseSpacePt(value) {
            var n = parseFloat(value);
            return isNaN(n) ? 0 : n * unit.factor;
        }

        return {
            beforeInput: beforeInput,
            afterInput: afterInput,
            currentSpaceBefore: function () { return parseSpacePt(beforeInput.text); },
            currentSpaceAfter: function () { return parseSpacePt(afterInput.text); }
        };
    }

    function buildButtonGroup(parent) {
        var buttonGroup = parent.add("group");
        buttonGroup.alignment = "center";
        var cancelBtn = buttonGroup.add("button", undefined, L("btnCancel"), { name: "cancel" });
        var okBtn = buttonGroup.add("button", undefined, L("btnOk"), { name: "ok" });
        return { okBtn: okBtn, cancelBtn: cancelBtn };
    }

    function showLeadingDialog(defaultAutoAmount, initialLeadingValuePt, initialLeadingChoiceState, useCommonInitially, initialSpaceBeforePt, initialSpaceAfterPt, defaultLeadingType, computePreviewLeading, onPreview) {
        var dlg = createDialogWindow();
        var modeGroup = buildModeGroup(dlg, useCommonInitially);
        var leadingPanel = buildLeadingPanel(dlg, defaultAutoAmount, initialLeadingValuePt, initialLeadingChoiceState);
        var typePanel = buildLeadingTypePanel(dlg, defaultLeadingType);
        var spacePanel = buildSpacePanel(dlg, initialSpaceBeforePt, initialSpaceAfterPt);
        var buttons = buildButtonGroup(dlg);

        var directMode = !!(initialLeadingChoiceState && initialLeadingChoiceState.isOther);

        var otherChoiceIndex = -1;
        for (var i = 0; i < LEADING_CHOICES.length; i++) {
            if (LEADING_CHOICES[i].other) {
                otherChoiceIndex = i;
                break;
            }
        }
        var autoChoiceIndex = -1;
        for (var j = 0; j < LEADING_CHOICES.length; j++) {
            if (LEADING_CHOICES[j].auto) {
                autoChoiceIndex = j;
                break;
            }
        }

        function refreshPreview() {
            var ratio = leadingPanel.currentRatio();
            var isOther = leadingPanel.currentIsOther();
            var effectiveDirectMode = directMode || isOther;

            if (!effectiveDirectMode) {
                if (typeof computePreviewLeading === "function") {
                    var repPt = computePreviewLeading(ratio, modeGroup.isCommon(), leadingPanel.currentAutoAmount());
                    leadingPanel.setLeadingDisplayPt(repPt);
                }
            }
            if (typeof onPreview === "function") {
                onPreview(
                    ratio,
                    typePanel.currentType(),
                    spacePanel.currentSpaceBefore(),
                    spacePanel.currentSpaceAfter(),
                    leadingPanel.currentAutoAmount(),
                    modeGroup.isCommon(),
                    effectiveDirectMode,
                    leadingPanel.currentLeadingPt()
                );
            }
        }

        function setDirect(flag) { directMode = flag; }

        modeGroup.perLineRb.onClick = function () { setDirect(false); refreshPreview(); };
        modeGroup.commonRb.onClick = function () { setDirect(false); refreshPreview(); };
        for (var k = 0; k < leadingPanel.radios.length; k++) {
            (function (index) {
                leadingPanel.radios[index].onClick = function () {
                    leadingPanel.selectChoiceByIndex(index);
                    setDirect(!!LEADING_CHOICES[index].other);
                    refreshPreview();
                };
            })(k);
        }
        for (var m = 0; m < typePanel.radios.length; m++) {
            typePanel.radios[m].onClick = refreshPreview;
        }
        spacePanel.beforeInput.onChange = refreshPreview;
        spacePanel.afterInput.onChange = refreshPreview;
        leadingPanel.autoInput.onChange = function () {
            if (autoChoiceIndex >= 0) {
                leadingPanel.selectChoiceByIndex(autoChoiceIndex);
            }
            setDirect(false);
            refreshPreview();
        };
        leadingPanel.leadingInput.onChange = function () {
            if (otherChoiceIndex >= 0) {
                leadingPanel.selectChoiceByIndex(otherChoiceIndex);
            }
            setDirect(true);
            refreshPreview();
        };

        changeValueByArrowKey(spacePanel.beforeInput, false, refreshPreview);
        changeValueByArrowKey(spacePanel.afterInput, false, refreshPreview);
        changeValueByArrowKey(leadingPanel.autoInput, false, function () {
            if (autoChoiceIndex >= 0) {
                leadingPanel.selectChoiceByIndex(autoChoiceIndex);
            }
            setDirect(false);
            refreshPreview();
        });
        changeValueByArrowKey(leadingPanel.leadingInput, false, function () {
            if (otherChoiceIndex >= 0) {
                leadingPanel.selectChoiceByIndex(otherChoiceIndex);
            }
            setDirect(true);
            refreshPreview();
        }, 1);

        var result = null;
        buttons.okBtn.onClick = function () {
            result = {
                ratio: leadingPanel.currentRatio(),
                type: typePanel.currentType(),
                spaceBefore: spacePanel.currentSpaceBefore(),
                spaceAfter: spacePanel.currentSpaceAfter(),
                autoAmount: leadingPanel.currentAutoAmount(),
                common: modeGroup.isCommon(),
                directMode: directMode || leadingPanel.currentIsOther(),
                directLeadingPt: leadingPanel.currentLeadingPt()
            };
            dlg.close(1);
        };
        buttons.cancelBtn.onClick = function () { dlg.close(0); };

        refreshPreview();
        try {
            leadingPanel.leadingInput.active = true;
        } catch (e) { }

        if (dlg.show() === 1) {
            return result;
        }
        return null;
    }

    main();

    function main() {
        normalizeTextRangeSelectionToTextFrame();
        var doc;
        try {
            doc = app.activeDocument;
        } catch (e) {
            alert(L("alertOpenDocument"));
            return;
        }

        var selectionItems = doc.selection;
        if (!selectionItems || selectionItems.length === 0) {
            alert(L("alertSelectTextObject"));
            return;
        }

        var targetFrames = [];
        var seenFrames = [];
        for (var i = 0; i < selectionItems.length; i++) {
            var textFrame = getProcessableTextFrame(selectionItems[i]);
            if (!textFrame) continue;

            var alreadyAdded = false;
            for (var j = 0; j < seenFrames.length; j++) {
                if (seenFrames[j] === textFrame) {
                    alreadyAdded = true;
                    break;
                }
            }
            if (alreadyAdded) continue;

            seenFrames.push(textFrame);
            targetFrames.push(textFrame);
        }

        if (targetFrames.length === 0) {
            alert(L("alertNoProcessableTextFrame"));
            return;
        }

        var snapshots = snapshotFrames(targetFrames);
        var errorShown = { flag: false };
        var initialAutoLeadingAmount = getInitialAutoLeadingAmount(targetFrames);
        var initialLeadingValuePt = getInitialLeadingValuePt(targetFrames);
        var initialLeadingChoiceState = getInitialLeadingChoiceState(targetFrames);
        var useCommonInitially = shouldUseCommonModeInitially(targetFrames);
        var initialLeadingType = getInitialLeadingType(targetFrames);
        var initialSpaceBeforePt = getInitialSpaceBeforePt(targetFrames);
        var initialSpaceAfterPt = getInitialSpaceAfterPt(targetFrames);

        function computePreviewLeadingPt(ratio, common, autoAmount) {
            if (targetFrames.length === 0) return NaN;
            var frame = targetFrames[0];
            var baseSize = NaN;
            if (common) {
                baseSize = getCommonBaseFontSize(frame);
            } else {
                var lines = frame.lines;
                for (var i = 0; i < lines.length; i++) {
                    var s = getBaseFontSizeFromLine(lines[i]);
                    if (!isNaN(s)) { baseSize = s; break; }
                }
            }
            if (isNaN(baseSize)) return NaN;
            if (ratio === null) {
                return baseSize * (autoAmount / 100);
            }
            return baseSize * ratio;
        }

        var chosen = showLeadingDialog(
            initialAutoLeadingAmount,
            initialLeadingValuePt,
            initialLeadingChoiceState,
            useCommonInitially,
            initialSpaceBeforePt,
            initialSpaceAfterPt,
            initialLeadingType,
            computePreviewLeadingPt,
            function (ratio, leadingType, spaceBefore, spaceAfter, autoAmount, common, directMode, directLeadingPt) {
                restoreFrames(snapshots);
                applyLeadingToFrames(targetFrames, ratio, leadingType, spaceBefore, spaceAfter, autoAmount, common, directMode, directLeadingPt, errorShown);
                app.redraw();
            }
        );

        if (chosen === null) {
            restoreFrames(snapshots);
            app.redraw();
        }
    }

    function getInitialLeadingType(frames) {
        for (var i = 0; i < frames.length; i++) {
            try {
                var textRange = frames[i].textRange;
                if (!textRange) continue;

                var t = textRange.leadingType;
                if (t !== undefined && t !== null) return t;
            } catch (e) { }
        }
        return null;
    }

    function getInitialLeadingValuePt(frames) {
        for (var i = 0; i < frames.length; i++) {
            try {
                var lines = frames[i].lines;
                if (lines && lines.length > 0 && lines[0].characters.length > 0) {
                    var ch = lines[0].characters[0];
                    var leading = ch.characterAttributes.leading;
                    if (!isNaN(leading)) return leading;
                }
            } catch (e) { }
        }
        return NaN;
    }

    function getInitialAutoLeadingAmount(frames) {
        for (var i = 0; i < frames.length; i++) {
            try {
                var paragraphs = frames[i].paragraphs;
                if (paragraphs && paragraphs.length > 0) {
                    var amount = paragraphs[0].paragraphAttributes.autoLeadingAmount;
                    if (!isNaN(amount)) return amount;
                }
            } catch (e) { }
        }
        return 175;
    }

    function getInitialSpaceBeforePt(frames) {
        for (var i = 0; i < frames.length; i++) {
            try {
                var paragraphs = frames[i].paragraphs;
                if (paragraphs && paragraphs.length > 0) {
                    var spaceBefore = paragraphs[0].paragraphAttributes.spaceBefore;
                    if (!isNaN(spaceBefore)) return spaceBefore;
                }
            } catch (e) { }
        }
        return 0;
    }

    function getInitialSpaceAfterPt(frames) {
        for (var i = 0; i < frames.length; i++) {
            try {
                var paragraphs = frames[i].paragraphs;
                if (paragraphs && paragraphs.length > 0) {
                    var spaceAfter = paragraphs[0].paragraphAttributes.spaceAfter;
                    if (!isNaN(spaceAfter)) return spaceAfter;
                }
            } catch (e) { }
        }
        return 0;
    }

    function snapshotFrames(frames) {
        var snapshots = [];
        for (var i = 0; i < frames.length; i++) {
            var tf = frames[i];
            var chars = tf.textRange.characters;
            var charStates = [];
            for (var c = 0; c < chars.length; c++) {
                try {
                    charStates.push({
                        autoLeading: chars[c].characterAttributes.autoLeading,
                        leading: chars[c].characterAttributes.leading,
                    });
                } catch (e) {
                    charStates.push(null);
                }
            }

            var paragraphs = tf.paragraphs;
            var paraStates = [];
            for (var p = 0; p < paragraphs.length; p++) {
                try {
                    paraStates.push({
                        spaceBefore: paragraphs[p].paragraphAttributes.spaceBefore,
                        spaceAfter: paragraphs[p].paragraphAttributes.spaceAfter,
                        autoLeadingAmount: paragraphs[p].paragraphAttributes.autoLeadingAmount
                    });
                } catch (e) {
                    paraStates.push(null);
                }
            }

            var leadingTypeState = null;
            try {
                leadingTypeState = tf.textRange.leadingType;
            } catch (e) { }

            snapshots.push({ textFrame: tf, charStates: charStates, paraStates: paraStates, leadingType: leadingTypeState });
        }
        return snapshots;
    }

    function restoreFrames(snapshots) {
        for (var i = 0; i < snapshots.length; i++) {
            var snap = snapshots[i];
            var chars = snap.textFrame.textRange.characters;
            var charLen = Math.min(chars.length, snap.charStates.length);
            for (var c = 0; c < charLen; c++) {
                var st = snap.charStates[c];
                if (!st) continue;
                try {
                    chars[c].characterAttributes.autoLeading = st.autoLeading;
                    if (!st.autoLeading) {
                        chars[c].characterAttributes.leading = st.leading;
                    }
                } catch (e) { }
            }

            var paragraphs = snap.textFrame.paragraphs;
            var paraLen = Math.min(paragraphs.length, snap.paraStates.length);
            for (var p = 0; p < paraLen; p++) {
                var ps = snap.paraStates[p];
                if (!ps) continue;
                try {
                    paragraphs[p].paragraphAttributes.spaceBefore = ps.spaceBefore;
                    paragraphs[p].paragraphAttributes.spaceAfter = ps.spaceAfter;
                    if (ps.autoLeadingAmount !== undefined && ps.autoLeadingAmount !== null) {
                        paragraphs[p].paragraphAttributes.autoLeadingAmount = ps.autoLeadingAmount;
                    }
                } catch (e) { }
            }
            try {
                if (snap.leadingType !== undefined && snap.leadingType !== null) {
                    snap.textFrame.textRange.leadingType = snap.leadingType;
                }
            } catch (e) { }
        }
    }

    function getCommonBaseFontSize(textFrame) {
        var allSizes = [];
        var lines = textFrame.lines;
        for (var j = 0; j < lines.length; j++) {
            var sizes = getLineSampleFontSizes(lines[j], LINE_FONT_SIZE_SAMPLE_COUNT);
            for (var s = 0; s < sizes.length; s++) {
                allSizes.push(sizes[s]);
            }
        }
        if (allSizes.length === 0) return NaN;
        return getMostFrequentNumber(allSizes);
    }

    function applyLeadingToFrames(frames, selectedLeadingRatio, leadingType, spaceBefore, spaceAfter, autoAmount, common, directMode, directLeadingPt, errorShown) {
        var useDirect = directMode && !isNaN(directLeadingPt);
        var useAuto = (!useDirect) && (selectedLeadingRatio === null);
        for (var i = 0; i < frames.length; i++) {
            var textFrame = frames[i];
            try {
                var lines = textFrame.lines;
                var commonBaseSize = (!useDirect && common && !useAuto) ? getCommonBaseFontSize(textFrame) : NaN;
                for (var j = 0; j < lines.length; j++) {
                    var line = lines[j];
                    if (useDirect) {
                        line.characterAttributes.autoLeading = false;
                        line.characterAttributes.leading = directLeadingPt;
                    } else if (useAuto) {
                        line.characterAttributes.autoLeading = true;
                    } else {
                        var baseFontSizeFromLine = common ? commonBaseSize : getBaseFontSizeFromLine(line);
                        if (isNaN(baseFontSizeFromLine)) continue;

                        var newLeading = baseFontSizeFromLine * selectedLeadingRatio;
                        line.characterAttributes.autoLeading = false;
                        line.characterAttributes.leading = newLeading;
                    }
                }

                var paragraphs = textFrame.paragraphs;
                for (var p = 0; p < paragraphs.length; p++) {
                    try {
                        paragraphs[p].paragraphAttributes.spaceBefore = spaceBefore;
                        paragraphs[p].paragraphAttributes.spaceAfter = spaceAfter;
                        paragraphs[p].paragraphAttributes.autoLeadingAmount = autoAmount;
                    } catch (e3) { }
                }
                try {
                    textFrame.textRange.leadingType = leadingType;
                } catch (e4) { }
            } catch (e) {
                if (errorShown && !errorShown.flag) {
                    alert(L("alertLeadingApplyError") + "\r" + e.message);
                    errorShown.flag = true;
                }
            }
        }
    }

    function normalizeTextRangeSelectionToTextFrame() {
        // テキストの一部（TextRange）が選択されている場合、親の TextFrame を選択し直す
        if (app.documents.length > 0 && app.selection && app.selection.typename === "TextRange") {
            var story = app.selection.story;
            if (story && story.textFrames.length === 1) {
                var parentTextFrame = story.textFrames[0];
                app.executeMenuCommand("deselectall");
                app.selection = [parentTextFrame];
                try {
                    app.selectTool("Adobe Select Tool");
                } catch (e) { }
            }
        }
    }

    function getLineSampleFontSizes(line, sampleCount) {
        var fontSizes = [];
        if (!line || !line.characters || line.characters.length === 0) {
            return fontSizes;
        }

        var maxCount = Math.min(line.characters.length, sampleCount);
        for (var i = 0; i < maxCount; i++) {
            try {
                var size = line.characters[i].characterAttributes.size;
                if (!isNaN(size)) {
                    fontSizes.push(size);
                }
            } catch (e) { }
        }

        return fontSizes;
    }

    function getMostFrequentNumber(values) {
        if (!values || values.length === 0) {
            return NaN;
        }

        var counts = {};
        var bestValue = values[0];
        var bestCount = 0;

        for (var i = 0; i < values.length; i++) {
            var key = String(values[i]);
            if (!counts[key]) {
                counts[key] = { value: values[i], count: 0 };
            }
            counts[key].count++;

            if (counts[key].count > bestCount) {
                bestCount = counts[key].count;
                bestValue = counts[key].value;
            } else if (counts[key].count === bestCount && counts[key].value > bestValue) {
                bestValue = counts[key].value;
            }
        }

        return bestValue;
    }

    function getBaseFontSizeFromLine(line) {
        var fontSizes = getLineSampleFontSizes(line, LINE_FONT_SIZE_SAMPLE_COUNT);
        if (!fontSizes || fontSizes.length === 0) {
            return NaN;
        }

        return getMostFrequentNumber(fontSizes);
    }

    function getProcessableTextFrame(item) {
        if (!item || item.typename !== "TextFrame") {
            return null;
        }

        if (!item.contents) {
            return null;
        }

        if (!item.lines || item.lines.length === 0) {
            return null;
        }

        return item;
    }

})();