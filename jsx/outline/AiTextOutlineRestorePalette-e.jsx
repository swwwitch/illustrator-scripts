#target illustrator
#targetengine "TextOutlineWithMemo"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
### Script Name

TextOutlineRestorePalette.jsx

English-only build. UI, comments, and the note storage format are all English.

### Overview

- Outline selected text (saving its attributes to the note) and restore it back, from a persistent palette
- Before outlining, character/paragraph attributes are serialized into the object's note as text
- On restore, the note is parsed to recreate the text frame, bringing back font, size, leading,
  kerning, fill color, alignment, horizontal/vertical scale, orientation, position, etc.
- The original outlines are moved to the outlined_text layer (dimmed and locked)
- The selected object's note is listed in the panel (Load button)
- Older notes without newer fields fall back to previous behavior (backward compatible)
- All DOM work is delegated to the main engine via BridgeTalk

### Features

- Outline with Memo: serialize properties into the note
- Restore Text: recreate the text from the note; the original outlines are moved to the outlined_text layer
- Show the selected object's note in the panel (Load button)

### Supported properties

- Text contents (multi-line)
- Orientation (vertical / horizontal)
- Font (PostScript name)
- Font size
- Leading
- Auto leading
- Horizontal & vertical scale
- Kerning method (metrics / optical / roman only / none)
- Proportional metrics (linked to the kerning method)
- Tracking
- Alignment (left / center / right / justify)
- Fill color (CMYK / RGB / Gray / Spot)
- Position (by geometricBounds; not shown in the list)

Note: Gradient/pattern fills and mixed attributes are not supported (the first character's value is used)

### note

https://note.com/dtp_transit/n/n3e0f241508db

*/

(function () {

// ==============================
// Version
// ==============================
var SCRIPT_VERSION = "v2.0.0";

// ==============================
// Labels
// ==============================
var LABELS = {
    dialog: {
        title: 'Outline & Restore Text'
    },
    panel: {
        selected: 'Object with a note',
        commands: 'Commands'
    },
    button: {
        outline: 'Outline with Memo',
        outlineTip: 'Select text and run (Esc to close)',
        restore: 'Restore Text',
        restoreTip: 'Restore text from the outline note (Esc to close)',
        load: 'Load Note',
        loadTip: "Load and show the selected object's note",
        attributes: 'Attributes',
        attributesTip: 'Toggle the Attributes panel'
    },
    memo: {
        empty: '(No note)'
    },
    listCol: {
        item: 'Item',
        value: 'Value'
    },
    status: {
        ready: 'Select text or an outline',
        busy: 'Working…',
        doneOutline: 'Outlined',
        doneRestore: 'Text restored',
        memoLoaded: 'Note loaded',
        fontWarn: 'Some fonts used defaults',
        nodoc: 'No document is open',
        nosel: 'No objects are selected',
        notgt: 'Please select a path or group',
        nonote: 'No usable note found',
        err: 'Error'
    }
};

function L(path) {
    var parts = String(path).split('.');
    var node = LABELS;
    for (var i = 0; i < parts.length; i++) {
        if (node == null) return path;
        node = node[parts[i]];
    }
    return (node == null) ? path : node;
}

// ==============================
// Window & panel margins and spacing
// ==============================
var WINDOW_MARGINS = 16;
var WINDOW_SPACING = 12;
var PANEL_MARGINS = [16, 20, 16, 12];
var PANEL_SPACING = 8;

/* Apply shared window layout */
function setupWindow(win, spacing) {
    win.orientation = "column";
    win.alignChildren = "fill";
    win.margins = WINDOW_MARGINS;
    win.spacing = (typeof spacing === "number") ? spacing : WINDOW_SPACING;
}

/* Apply shared panel layout */
function setupPanel(panel, spacing) {
    panel.orientation = "column";
    panel.alignChildren = ["fill", "top"];
    panel.alignment = "fill";
    panel.margins = PANEL_MARGINS;
    panel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
}

/* Apply a horizontal row group (button rows, etc.) */
function setupRow(group, alignment, spacing) {
    group.orientation = "row";
    group.alignment = alignment || "left";
    group.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
}

// ==============================
// Worker functions (run in main engine)
//   NOTE: inside each worker body, line comments are forbidden; use block comments /* */ only,
//   and always end statements with a semicolon (newlines are stripped when sent via Function.toString()).
//   This heading comment outside the functions is not affected.
// ==============================

/* --- Outline --- */
function workerRound(value) {
    return Math.round(value * 100) / 100;
}

/* --- Toggle the Attributes panel via menu command --- */
function workerToggleAttributesPanel() {
    try {
        app.executeMenuCommand("internal palettes posing as plug-in menus-attributes");
    } catch (e) {
    }
    return "OK";
}

/* --- Build the memo text from gathered values --- */
function workerBuildMemoText(info) {
    return "Text:\n" + info.content + "\n\n" +
        "Font:\n" + info.fontName + "\n\n" +
        "FontSize:\n" + info.fontSize + "\n\n" +
        "Leading:\n" + info.leading + "\n\n" +
        "Kerning:\n" + info.kerning + "\n\n" +
        "ProportionalMetrics:\n" + info.proportionalMetrics + "\n\n" +
        "Tracking:\n" + info.tracking + "\n\n" +
        "Orientation:\n" + info.orientation + "\n\n" +
        "Color:\n" + info.color + "\n\n" +
        "Justification:\n" + info.justification + "\n\n" +
        "AutoLeading:\n" + info.autoLeading + "\n\n" +
        "HorizontalScale:\n" + info.horizontalScale + "\n\n" +
        "VerticalScale:\n" + info.verticalScale + "\n\n" +
        "Bounds:\nL = " + info.left + ", T = " + info.top + ", R = " + info.right + ", B = " + info.bottom;
}

function workerProcessTextFrame(textFrame) {
    var textRange = textFrame.textRange;
    var kerningMethod = textRange.characterAttributes.kerningMethod;
    var kerningMethodText;
    if (kerningMethod == AutoKernType.AUTO) { kerningMethodText = "Metrics"; }
    else if (kerningMethod == AutoKernType.METRICSROMANONLY) { kerningMethodText = "Roman Only"; }
    else if (kerningMethod == AutoKernType.OPTICAL) { kerningMethodText = "Optical"; }
    else { kerningMethodText = "None"; }
    app.redraw();
    var bounds = textFrame.geometricBounds;
    var memoText = workerBuildMemoText({
        content: textRange.contents,
        fontName: textRange.characterAttributes.textFont.name,
        fontSize: workerRound(textRange.characterAttributes.size),
        leading: workerRound(textRange.characterAttributes.leading),
        kerning: kerningMethodText,
        proportionalMetrics: textRange.characterAttributes.proportionalMetrics ? "true" : "false",
        tracking: textRange.characterAttributes.tracking,
        orientation: (textFrame.orientation == TextOrientation.VERTICAL) ? "Vertical" : "Horizontal",
        color: workerColorToText(textRange.characterAttributes.fillColor),
        justification: workerJustificationToText(textRange.paragraphAttributes.justification),
        autoLeading: textRange.characterAttributes.autoLeading ? "true" : "false",
        horizontalScale: workerRound(textRange.characterAttributes.horizontalScale),
        verticalScale: workerRound(textRange.characterAttributes.verticalScale),
        left: workerRound(bounds[0]),
        top: workerRound(bounds[1]),
        right: workerRound(bounds[2]),
        bottom: workerRound(bounds[3])
    });
    app.activeDocument.selection = null;
    textFrame.selected = true;
    textFrame.createOutline();
    /* the selection after createOutline is not always what we expect, so check minimally */
    if (app.activeDocument.selection && app.activeDocument.selection.length > 0) {
        app.activeDocument.selection[0].note = memoText;
    }
    return true;
}

function workerRunOutline() {
    if (app.documents.length < 1) { return "NODOC"; }
    var doc = app.activeDocument;
    var currentSelection = doc.selection;
    if (!currentSelection || currentSelection.length < 1) { return "NOSEL"; }
    var selectedTextFrames = [];
    var loopIndex;
    for (loopIndex = 0; loopIndex < currentSelection.length; loopIndex++) {
        if (currentSelection[loopIndex].typename == "TextFrame") { selectedTextFrames.push(currentSelection[loopIndex]); }
    }
    if (selectedTextFrames.length < 1) { return "NOSEL"; }
    for (loopIndex = 0; loopIndex < selectedTextFrames.length; loopIndex++) {
        workerProcessTextFrame(selectedTextFrames[loopIndex]);
    }
    return "OK:" + selectedTextFrames.length;
}

/* --- Format the note for compact display --- */
function workerFormatNoteForDisplay(noteText) {
    var noteLines = noteText.split("\n");
    var displayLabels = ["Text", "Orientation", "Font", "FontSize", "Leading", "AutoLeading", "HorizontalScale", "VerticalScale", "Kerning", "ProportionalMetrics", "Tracking", "Justification", "Color"];
    var displayLines = [];
    /* the body (Text) can span multiple lines, so grab the whole "Text:\n" .. "\n\nFont:" block and fold newlines into a single line with a return symbol */
    var textStartMarker = "Text:\n";
    var textEndMarker = "\n\nFont:";
    var textStart = noteText.indexOf(textStartMarker);
    var textEnd = noteText.indexOf(textEndMarker);
    var bodyText = null;
    if (textStart >= 0 && textEnd > textStart) { bodyText = noteText.substring(textStart + textStartMarker.length, textEnd); }
    var labelIndex;
    for (labelIndex = 0; labelIndex < displayLabels.length; labelIndex++) {
        var currentLabel = displayLabels[labelIndex];
        if (currentLabel === "Text" && bodyText != null) {
            displayLines.push("Text: " + bodyText.replace(/[\r\n]/g, "↵"));
            continue;
        }
        var scanIndex;
        for (scanIndex = 0; scanIndex < noteLines.length; scanIndex++) {
            if (noteLines[scanIndex].indexOf(currentLabel + ":") === 0 && scanIndex + 1 < noteLines.length) {
                var displayValue = noteLines[scanIndex + 1];
                displayLines.push(currentLabel + ": " + displayValue);
                break;
            }
        }
    }
    return displayLines.join("\n");
}

/* --- Inspect selection state (has text / has note / display note) --- */
function workerInspectSelection() {
    if (app.documents.length < 1) { return "NODOC"; }
    var doc = app.activeDocument;
    var currentSelection = doc.selection;
    if (!currentSelection || currentSelection.length < 1) { return "NOSEL"; }
    var noteHolder = null;
    var scanIndex;
    for (scanIndex = 0; scanIndex < currentSelection.length; scanIndex++) {
        var candidate = currentSelection[scanIndex];
        var candidateType = candidate.typename;
        /* with multiple selection, use the first object that carries a note */
        if (candidateType == "GroupItem" || candidateType == "PathItem" || candidateType == "TextFrame") {
            if (candidate.note && candidate.note.length > 0) { noteHolder = candidate; break; }
        }
    }
    if (!noteHolder) { return "NONOTE"; }
    var formattedNote = workerFormatNoteForDisplay(noteHolder.note);
    if (!formattedNote || formattedNote.length < 1) { formattedNote = noteHolder.note; }
    /* return in the form NOTE:<display note> or NONOTE */
    return "NOTE:" + formattedNote;
}

/* --- Restore: parse note --- */
function workerExtractTextAttributes(noteText) {
    var attributes = {
        text: null,
        font: null,
        fontSize: null,
        leading: null,
        orientation: null,
        tracking: null,
        kerningText: null,
        colorText: null,
        proportionalMetrics: null,
        justificationText: null,
        autoLeading: null,
        horizontalScale: null,
        verticalScale: null,
        x: null,
        y: null,
        savedBounds: null
    };
    /* the body (Text) can span multiple lines, so take the whole "Text:\n" .. "\n\nFont:" block as the body (keeping inner newlines and blank lines) */
    var textStartMarker = "Text:\n";
    var textEndMarker = "\n\nFont:";
    var textStart = noteText.indexOf(textStartMarker);
    var textEnd = noteText.indexOf(textEndMarker);
    var restText = noteText;
    if (textStart >= 0 && textEnd > textStart) {
        attributes.text = noteText.substring(textStart + textStartMarker.length, textEnd);
        restText = noteText.substring(textEnd + 2);
    }
    /* the single-line fields from Font onward are scanned only in the remainder (to avoid matching body lines) */
    var noteLines = restText.split("\n");
    var lineIndex;
    for (lineIndex = 0; lineIndex < noteLines.length; lineIndex++) {
        if (attributes.text == null && noteLines[lineIndex].indexOf("Text:") === 0 && lineIndex + 1 < noteLines.length) { attributes.text = noteLines[lineIndex + 1]; }
        if (noteLines[lineIndex].indexOf("Font:") === 0 && lineIndex + 1 < noteLines.length) { attributes.font = noteLines[lineIndex + 1]; }
        if (noteLines[lineIndex].indexOf("FontSize:") === 0 && lineIndex + 1 < noteLines.length) { attributes.fontSize = parseFloat(noteLines[lineIndex + 1]); }
        if (noteLines[lineIndex].indexOf("Leading:") === 0 && lineIndex + 1 < noteLines.length) { attributes.leading = parseFloat(noteLines[lineIndex + 1]); }
        if (noteLines[lineIndex].indexOf("Orientation:") === 0 && lineIndex + 1 < noteLines.length) { attributes.orientation = noteLines[lineIndex + 1]; }
        if (noteLines[lineIndex].indexOf("Tracking:") === 0 && lineIndex + 1 < noteLines.length) { attributes.tracking = parseFloat(noteLines[lineIndex + 1]); }
        if (noteLines[lineIndex].indexOf("Kerning:") === 0 && lineIndex + 1 < noteLines.length) { attributes.kerningText = noteLines[lineIndex + 1]; }
        if (noteLines[lineIndex].indexOf("Color:") === 0 && lineIndex + 1 < noteLines.length) { attributes.colorText = noteLines[lineIndex + 1]; }
        if (noteLines[lineIndex].indexOf("ProportionalMetrics:") === 0 && lineIndex + 1 < noteLines.length) { attributes.proportionalMetrics = noteLines[lineIndex + 1] === "true"; }
        if (noteLines[lineIndex].indexOf("Justification:") === 0 && lineIndex + 1 < noteLines.length) { attributes.justificationText = noteLines[lineIndex + 1]; }
        if (noteLines[lineIndex].indexOf("AutoLeading:") === 0 && lineIndex + 1 < noteLines.length) { attributes.autoLeading = noteLines[lineIndex + 1] === "true"; }
        if (noteLines[lineIndex].indexOf("HorizontalScale:") === 0 && lineIndex + 1 < noteLines.length) { attributes.horizontalScale = parseFloat(noteLines[lineIndex + 1]); }
        if (noteLines[lineIndex].indexOf("VerticalScale:") === 0 && lineIndex + 1 < noteLines.length) { attributes.verticalScale = parseFloat(noteLines[lineIndex + 1]); }
        if (noteLines[lineIndex].match(/L\s*=\s*([-]?\d+(?:\.\d+)?),\s*T\s*=\s*([-]?\d+(?:\.\d+)?),\s*R\s*=\s*([-]?\d+(?:\.\d+)?),\s*B\s*=\s*([-]?\d+(?:\.\d+)?)/)) {
            attributes.savedBounds = [parseFloat(RegExp.$1), parseFloat(RegExp.$2), parseFloat(RegExp.$3), parseFloat(RegExp.$4)];
        }
    }
    return attributes;
}

/* --- Restore: align by geometricBounds --- */
function workerAlignTextFrameByBounds(target, textFrame) {
    try {
        app.redraw();
        var targetBounds;
        if (target && target.length === 4 && typeof target[0] === 'number') { targetBounds = target; }
        else if (target && target.geometricBounds) { targetBounds = target.geometricBounds; }
        else { return; }
        var frameBounds = textFrame.geometricBounds;
        var dx = targetBounds[0] - frameBounds[0];
        var dy = targetBounds[1] - frameBounds[1];
        textFrame.translate(dx, dy);
    } catch (e) {
        /* keep going on failure */
    }
}

/* --- Restore: make a layer usable (unlock + show) --- */
function workerSetLayerUsable(layer) {
    try { layer.locked = false; } catch (eLock) {}
    try { layer.visible = true; } catch (eVis) {}
}

/* --- Restore: reuse the existing outlined_text layer (unlocked) if present, otherwise create a fresh stash layer --- */
function workerCreateOutlineStashLayer(doc) {
    if (!doc) { return null; }
    var stashLayer = null;
    /* if an outlined_text layer already exists, unlock it and reuse it as the stash (consolidate outlines into one layer) */
    var findIndex;
    for (findIndex = 0; findIndex < doc.layers.length; findIndex++) {
        if (doc.layers[findIndex].name === "outlined_text") { stashLayer = doc.layers[findIndex]; break; }
    }
    if (!stashLayer) {
        stashLayer = doc.layers.add();
        stashLayer.name = "__outlined_text_stash__";
    }
    workerSetLayerUsable(stashLayer);
    try {
        var markerGroup = stashLayer.groupItems.add();
        markerGroup.name = "__outlined_text_marker__";
    } catch (e2) {}
    return stashLayer;
}

/* --- Restore: get the restored_text layer --- */
function workerCreateRestoredTextLayer(doc) {
    var baseName = "restored_text";
    var targetLayer = null;
    var findIndex;
    for (findIndex = 0; findIndex < doc.layers.length; findIndex++) {
        if (doc.layers[findIndex].name === baseName) { targetLayer = doc.layers[findIndex]; break; }
    }
    if (!targetLayer) { targetLayer = doc.layers.add(); targetLayer.name = baseName; }
    workerSetLayerUsable(targetLayer);
    /* restore always reuses the same restored_text layer to consolidate, so do not auto-merge numbered layers
       (avoid swallowing user-made layers like restored_text1) */
    return targetLayer;
}

/* --- Restore: merge existing outlined_text layers (incl. old archive/dup) into the target so only one remains --- */
function workerMergeExistingOutlinedLayers(doc, targetLayer) {
    if (!doc || !targetLayer) { return; }
    var mergePattern = /^outlined_text(_archive\d+|_dup\d+)?$/;
    var mergeIndex;
    for (mergeIndex = doc.layers.length - 1; mergeIndex >= 0; mergeIndex--) {
        var mergeLayer = doc.layers[mergeIndex];
        if (mergeLayer === targetLayer) { continue; }
        if (!mergeLayer.name || !mergePattern.test(mergeLayer.name)) { continue; }
        workerSetLayerUsable(mergeLayer);
        try {
            while (mergeLayer.pageItems.length > 0) { mergeLayer.pageItems[0].move(targetLayer, ElementPlacement.PLACEATBEGINNING); }
        } catch (e1) {}
        try {
            while (mergeLayer.layers && mergeLayer.layers.length > 0) { mergeLayer.layers[0].remove(); }
        } catch (e2) {}
        try { mergeLayer.remove(); } catch (e3) {}
    }
}

/* --- Restore: dedupe outlined_text names --- */
function workerNormalizeOutlinedLayerNames(doc, keepLayer) {
    if (!doc || !keepLayer) { return; }
    var dupCounter = 1;
    var normalizeIndex;
    for (normalizeIndex = 0; normalizeIndex < doc.layers.length; normalizeIndex++) {
        var normalizeLayer = doc.layers[normalizeIndex];
        if (normalizeLayer !== keepLayer && normalizeLayer.name === "outlined_text") {
            workerSetLayerUsable(normalizeLayer);
            try { normalizeLayer.name = "outlined_text_dup" + dupCounter; } catch (e2) {}
            dupCounter++;
        }
    }
}

/* --- Restore: cleanup after the template action --- */
function workerCleanupOutlinedDuplicateNames(doc) {
    if (!doc) { return null; }
    var keepLayer = null;
    var searchIndex;
    for (searchIndex = 0; searchIndex < doc.layers.length; searchIndex++) {
        var searchLayer = doc.layers[searchIndex];
        if (searchLayer.name !== "outlined_text") { continue; }
        try {
            var groupIndex;
            for (groupIndex = 0; groupIndex < searchLayer.groupItems.length; groupIndex++) {
                if (searchLayer.groupItems[groupIndex].name === "__outlined_text_marker__") { keepLayer = searchLayer; break; }
            }
        } catch (e0) {}
        if (keepLayer) { break; }
    }
    if (!keepLayer) {
        var bottomIndex;
        for (bottomIndex = doc.layers.length - 1; bottomIndex >= 0; bottomIndex--) {
            if (doc.layers[bottomIndex].name === "outlined_text") { keepLayer = doc.layers[bottomIndex]; break; }
        }
    }
    /* rename every outlined_text other than keep (reuse the same logic as normalize) */
    workerNormalizeOutlinedLayerNames(doc, keepLayer);
    if (keepLayer) {
        workerSetLayerUsable(keepLayer);
        try {
            var removeIndex;
            for (removeIndex = keepLayer.groupItems.length - 1; removeIndex >= 0; removeIndex--) {
                if (keepLayer.groupItems[removeIndex].name === "__outlined_text_marker__") { keepLayer.groupItems[removeIndex].remove(); }
            }
        } catch (e4) {}
    }
    return keepLayer;
}

/* --- Restore: force the given layer active --- */
function workerSetActiveLayerStrict(doc, layer) {
    if (!doc || !layer) { return false; }
    try { doc.activeLayer = layer; } catch (e0) {}
    try { if (doc.activeLayer === layer) { return true; } } catch (e1) {}
    try {
        var reFindIndex;
        for (reFindIndex = 0; reFindIndex < doc.layers.length; reFindIndex++) {
            if (doc.layers[reFindIndex] === layer) { doc.activeLayer = doc.layers[reFindIndex]; break; }
        }
    } catch (e2) {}
    try { return doc.activeLayer === layer; } catch (e3) { return false; }
}

/* --- Restore: apply the template-layer attribute via action --- */
function workerApplyTemplateLayerAttribute() {
    var actionString =
        '/version 3' +
        '/name [ 5' +
        ' 6c61796572' +
        ' ]' +
        '/isOpen 1' +
        '/actionCount 1' +
        '/action-1 {' +
        ' /name [ 24' +
        ' 6368616e67652d746f2d74656d706c6174652d6c61796572' +
        ' ]' +
        ' /keyIndex 0' +
        ' /colorIndex 0' +
        ' /isOpen 1' +
        ' /eventCount 1' +
        ' /event-1 {' +
        ' /useRulersIn1stQuadrant 0' +
        ' /internalName (ai_plugin_Layer)' +
        ' /localizedName [ 9' +
        ' e8a1a8e7a4ba203a20' +
        ' ]' +
        ' /isOpen 1' +
        ' /isOn 1' +
        ' /hasDialog 1' +
        ' /showDialog 0' +
        ' /parameterCount 10' +
        ' /parameter-1 {' +
        ' /key 1836411236' +
        ' /showInPalette 4294967295' +
        ' /type (integer)' +
        ' /value 4' +
        ' }' +
        ' /parameter-2 {' +
        ' /key 1851878757' +
        ' /showInPalette 4294967295' +
        ' /type (ustring)' +
        ' /value [ 36' +
        ' e383ace382a4e383a4e383bce38391e3838de383abe382aae38397e382b7e383' +
        ' a7e383b3' +
        ' ]' +
        ' }' +
        ' /parameter-3 {' +
        ' /key 1953068140' +
        ' /showInPalette 4294967295' +
        ' /type (ustring)' +
        ' /value [ 13' +
        ' 6f75746c696e65645f74657874' +
        ' ]' +
        ' }' +
        ' /parameter-4 {' +
        ' /key 1953329260' +
        ' /showInPalette 4294967295' +
        ' /type (boolean)' +
        ' /value 1' +
        ' }' +
        ' /parameter-5 {' +
        ' /key 1936224119' +
        ' /showInPalette 4294967295' +
        ' /type (boolean)' +
        ' /value 1' +
        ' }' +
        ' /parameter-6 {' +
        ' /key 1819239275' +
        ' /showInPalette 4294967295' +
        ' /type (boolean)' +
        ' /value 1' +
        ' }' +
        ' /parameter-7 {' +
        ' /key 1886549623' +
        ' /showInPalette 4294967295' +
        ' /type (boolean)' +
        ' /value 1' +
        ' }' +
        ' /parameter-8 {' +
        ' /key 1886547572' +
        ' /showInPalette 4294967295' +
        ' /type (boolean)' +
        ' /value 0' +
        ' }' +
        ' /parameter-9 {' +
        ' /key 1684630830' +
        ' /showInPalette 4294967295' +
        ' /type (boolean)' +
        ' /value 1' +
        ' }' +
        ' /parameter-10 {' +
        ' /key 1885564532' +
        ' /showInPalette 4294967295' +
        ' /type (unit real)' +
        ' /value 50.0' +
        ' /unit 592474723' +
        ' }' +
        ' }' +
        '}';
    /* keep the action loaded until the palette closes (load only on first use) */
    if (!$.global.__outlineTemplateActionLoaded) {
        var actionFile = new File('~/ScriptAction.aia');
        actionFile.encoding = 'UTF-8';
        actionFile.lineFeed = 'Unix';
        actionFile.open('w');
        actionFile.write(actionString);
        actionFile.close();
        app.loadAction(actionFile);
        actionFile.remove();
        $.global.__outlineTemplateActionLoaded = true;
    }
    app.doScript("change-to-template-layer", "layer", false);
    /* unloadAction is run together in workerUnloadTemplateAction when the palette closes */
}

/* --- Restore: unload the loaded template action (on palette close) --- */
function workerUnloadTemplateAction() {
    if ($.global.__outlineTemplateActionLoaded) {
        try { app.unloadAction("layer", ""); } catch (e) {}
        $.global.__outlineTemplateActionLoaded = false;
    }
    return "OK";
}

/* --- Restore: finalize the template layer --- */
function workerFinalizeTemplateLayer(outlinedTextLayer) {
    if (!outlinedTextLayer) { return; }
    workerSetLayerUsable(outlinedTextLayer);
    workerMergeExistingOutlinedLayers(app.activeDocument, outlinedTextLayer);
    if (outlinedTextLayer.name !== "outlined_text") { outlinedTextLayer.name = "outlined_text"; }
    try { outlinedTextLayer.zOrder(ZOrderMethod.SENDTOBACK); } catch (eBack) {}
    workerNormalizeOutlinedLayerNames(app.activeDocument, outlinedTextLayer);
    try {
        var doc = app.activeDocument;
        var madeActive = workerSetActiveLayerStrict(doc, outlinedTextLayer);
        if (!madeActive) { return; }
        workerApplyTemplateLayerAttribute();
        outlinedTextLayer = workerCleanupOutlinedDuplicateNames(app.activeDocument) || outlinedTextLayer;
    } catch (eAct) {}
    try { outlinedTextLayer.locked = false; } catch (eUnlock) {}
    try {
        var markerIndex;
        for (markerIndex = outlinedTextLayer.groupItems.length - 1; markerIndex >= 0; markerIndex--) {
            if (outlinedTextLayer.groupItems[markerIndex].name === "__outlined_text_marker__") { outlinedTextLayer.groupItems[markerIndex].remove(); }
        }
    } catch (eMarker) {}
    try { outlinedTextLayer.zOrder(ZOrderMethod.SENDTOBACK); } catch (e9) {}
    try { outlinedTextLayer.locked = true; } catch (e10) {}
}

/* --- Serialize a fill color to text (CMYK/RGB/Gray/Spot) --- */
/* fillColor type follows the document color mode (CMYK/RGB), so store based on the type */
function workerColorToText(color) {
    if (!color) { return ""; }
    var typeName = color.typename;
    if (typeName == "NoColor") { return "NONE"; }
    if (typeName == "CMYKColor") { return "CMYK " + workerRound(color.cyan) + " " + workerRound(color.magenta) + " " + workerRound(color.yellow) + " " + workerRound(color.black); }
    if (typeName == "RGBColor") { return "RGB " + Math.round(color.red) + " " + Math.round(color.green) + " " + Math.round(color.blue); }
    if (typeName == "GrayColor") { return "GRAY " + workerRound(color.gray); }
    if (typeName == "SpotColor") { return "SPOT " + color.spot.name; }
    /* gradient/pattern etc. cannot be stored -> leave empty and keep the default color on restore (avoid invisible text) */
    return "";
}

/* --- Rebuild a fill color from text --- */
function workerColorFromText(colorText) {
    if (!colorText) { return null; }
    if (colorText === "NONE") { return new NoColor(); }
    var parts = colorText.split(" ");
    var kind = parts[0];
    if (kind === "CMYK" && parts.length >= 5) {
        var cmyk = new CMYKColor();
        cmyk.cyan = parseFloat(parts[1]);
        cmyk.magenta = parseFloat(parts[2]);
        cmyk.yellow = parseFloat(parts[3]);
        cmyk.black = parseFloat(parts[4]);
        return cmyk;
    }
    if (kind === "RGB" && parts.length >= 4) {
        var rgb = new RGBColor();
        rgb.red = parseFloat(parts[1]);
        rgb.green = parseFloat(parts[2]);
        rgb.blue = parseFloat(parts[3]);
        return rgb;
    }
    if (kind === "GRAY" && parts.length >= 2) {
        var gray = new GrayColor();
        gray.gray = parseFloat(parts[1]);
        return gray;
    }
    if (kind === "SPOT") {
        /* reuse the existing spot by swatch name (names may contain spaces, so take everything after the prefix as the name) */
        var spotName = colorText.substring(5);
        try {
            var spotColor = new SpotColor();
            spotColor.spot = app.activeDocument.spots.getByName(spotName);
            return spotColor;
        } catch (e) {
            return null;
        }
    }
    return null;
}

/* --- Restore: resolve kerning label to AutoKernType --- */
/* same mapping as AutoKerning.jsx (numeric assignment is not allowed; always use the enum) */
function workerResolveKernType(kerningText) {
    if (kerningText === "Metrics") { return AutoKernType.AUTO; }
    if (kerningText === "Optical") { return AutoKernType.OPTICAL; }
    if (kerningText === "Roman Only") { return AutoKernType.METRICSROMANONLY; }
    return AutoKernType.NOAUTOKERN;
}

/* --- Serialize justification to text --- */
function workerJustificationToText(justification) {
    if (justification == Justification.CENTER) { return "Center"; }
    if (justification == Justification.RIGHT) { return "Right"; }
    if (justification == Justification.FULLJUSTIFYLASTLINELEFT) { return "Justify (Last Left)"; }
    if (justification == Justification.FULLJUSTIFYLASTLINECENTER) { return "Justify (Last Center)"; }
    if (justification == Justification.FULLJUSTIFYLASTLINERIGHT) { return "Justify (Last Right)"; }
    if (justification == Justification.FULLJUSTIFY) { return "Justify"; }
    return "Left";
}

/* --- Restore: resolve alignment label to Justification --- */
/* the restore target is a new text frame whose default is LEFT; LEFT assignment may be ignored but matches the default, so no harm */
function workerResolveJustification(justificationText) {
    if (justificationText === "Center") { return Justification.CENTER; }
    if (justificationText === "Right") { return Justification.RIGHT; }
    if (justificationText === "Justify (Last Left)") { return Justification.FULLJUSTIFYLASTLINELEFT; }
    if (justificationText === "Justify (Last Center)") { return Justification.FULLJUSTIFYLASTLINECENTER; }
    if (justificationText === "Justify (Last Right)") { return Justification.FULLJUSTIFYLASTLINERIGHT; }
    if (justificationText === "Justify") { return Justification.FULLJUSTIFY; }
    return Justification.LEFT;
}

/* --- Restore: apply kerning (same as applyKerningToRanges in AutoKerning.jsx) --- */
/* turn proportionalMetrics ON only for Metrics (AUTO), OFF otherwise */
function workerApplyKerning(textRange, kerningMethod) {
    var useProportionalMetrics = (kerningMethod === AutoKernType.AUTO);
    try {
        textRange.characterAttributes.kerningMethod = kerningMethod;
        textRange.characterAttributes.proportionalMetrics = useProportionalMetrics;
    } catch (e) {
        /* skip ranges that can't take these */
    }
}

/* --- Restore: recreate the text frame --- */
function workerCreateRestoredTextFrame(sourceItem, attributes, restoredLayer, restoreReport) {
    var targetLayer = restoredLayer || sourceItem.layer;
    var textFrame = targetLayer.textFrames.add();
    textFrame.contents = attributes.text;
    var posX = (attributes.x !== null) ? attributes.x : sourceItem.left;
    var posY = (attributes.y !== null) ? attributes.y : sourceItem.top;
    textFrame.position = [posX, posY];
    /* font is handled separately: if not found, keep the default font. other attributes still apply */
    try {
        textFrame.textRange.characterAttributes.textFont = app.textFonts.getByName(attributes.font);
    } catch (eFont) {
        if (restoreReport) { restoreReport.fontFallback = true; }
    }
    /* apply each attribute individually (one failure must not stop the others) and not depend on whether the font was found */
    var attrs = textFrame.textRange.characterAttributes;
    try { attrs.size = attributes.fontSize; } catch (eSize) {}
    /* when auto leading is ON, do not set leading explicitly and let it auto-compute (old notes lack autoLeading, treated as false) */
    try {
        if (attributes.autoLeading === true) {
            attrs.autoLeading = true;
        } else {
            attrs.autoLeading = false;
            if (attributes.leading !== null) { attrs.leading = attributes.leading; }
        }
    } catch (eLead) {}
    try { attrs.tracking = attributes.tracking; } catch (eTrack) {}
    /* horizontal/vertical scale. old notes lack them, so leave untouched when null */
    try { if (attributes.horizontalScale !== null) { attrs.horizontalScale = attributes.horizontalScale; } } catch (eHScale) {}
    try { if (attributes.verticalScale !== null) { attrs.verticalScale = attributes.verticalScale; } } catch (eVScale) {}
    try {
        if (attributes.kerningText != null) {
            /* restore the kerning method (AutoKerning logic: proportionalMetrics follows the method) */
            workerApplyKerning(textFrame.textRange, workerResolveKernType(attributes.kerningText));
        } else {
            /* old notes (no kerning recorded) restore the stored proportionalMetrics */
            attrs.proportionalMetrics = attributes.proportionalMetrics;
        }
    } catch (eKern) {}
    try { textFrame.orientation = (attributes.orientation === "Vertical") ? TextOrientation.VERTICAL : TextOrientation.HORIZONTAL; } catch (eOri) {}
    /* alignment (paragraph attribute). a new frame defaults to LEFT, so left alignment is effectively unchanged */
    try {
        if (attributes.justificationText != null) {
            textFrame.textRange.paragraphAttributes.justification = workerResolveJustification(attributes.justificationText);
        }
    } catch (eJust) {}
    try {
        if (attributes.colorText != null) {
            var restoredColor = workerColorFromText(attributes.colorText);
            if (restoredColor) { attrs.fillColor = restoredColor; }
        }
    } catch (eColor) {}
    return textFrame;
}

/* --- Restore: move the source to the stash layer (fall back to layer root on failure) --- */
function workerMoveToOutlinedLayer(sourceItem, outlinedTextLayer) {
    try {
        sourceItem.moveToBeginning(outlinedTextLayer.groupItems.add());
    } catch (e) {
        /* if creating/moving into the group fails, stash directly under the layer */
        try { sourceItem.moveToBeginning(outlinedTextLayer); } catch (e2) {}
    }
}

/* --- Restore one item (path or group) --- */
function workerRestoreItem(sourceItem, outlinedTextLayer, restoredTextLayer, restoreReport) {
    var noteText = sourceItem.note;
    if (!noteText) { return false; }
    var attributes = workerExtractTextAttributes(noteText);
    if (!attributes || attributes.text == null) { return false; }
    var textFrame = workerCreateRestoredTextFrame(sourceItem, attributes, restoredTextLayer, restoreReport);
    workerAlignTextFrameByBounds(attributes.savedBounds || sourceItem, textFrame);
    /* stash first, then dim + lock (moving while unlocked is more reliable) */
    workerMoveToOutlinedLayer(sourceItem, outlinedTextLayer);
    /* dim + lock regardless of type. do not stop restore on failure */
    try {
        sourceItem.opacity = 30;
        sourceItem.locked = true;
    } catch (eDim) {}
    if (restoreReport) { restoreReport.restored++; }
    return true;
}

/* --- Restore: entry --- */
function workerRestoreText() {
    if (app.documents.length < 1) { return "NODOC"; }
    var doc = app.activeDocument;
    var currentSelection = doc.selection;
    if (!currentSelection || currentSelection.length < 1) { return "NOSEL"; }
    var restorableItems = [];
    var pickIndex;
    for (pickIndex = 0; pickIndex < currentSelection.length; pickIndex++) {
        var pickType = currentSelection[pickIndex].typename;
        if (pickType === "GroupItem" || pickType === "PathItem") { restorableItems.push(currentSelection[pickIndex]); }
    }
    if (restorableItems.length < 1) { return "NOTGT"; }
    var outlinedTextLayer = workerCreateOutlineStashLayer(doc);
    var restoredTextLayer = workerCreateRestoredTextLayer(doc);
    var restoreReport = { restored: 0, fontFallback: false };
    /* wrap the whole loop in try so an unexpected error does not leave empty layers behind */
    var runError = null;
    try {
        var restoreIndex;
        for (restoreIndex = 0; restoreIndex < restorableItems.length; restoreIndex++) {
            workerRestoreItem(restorableItems[restoreIndex], outlinedTextLayer, restoredTextLayer, restoreReport);
        }
    } catch (eRun) {
        runError = String(eRun);
    }
    if (restoreReport.restored < 1) {
        /* if nothing was restored (including on error), clean up the stash/restore layers we created */
        try { outlinedTextLayer.remove(); } catch (eStash) {}
        try {
            if (restoredTextLayer.pageItems.length < 1 && (!restoredTextLayer.layers || restoredTextLayer.layers.length < 1)) { restoredTextLayer.remove(); }
        } catch (eRestored) {}
        return runError ? ("ERR:" + runError) : "NONOTE";
    }
    /* if at least some succeeded, commit them (never leave the stash layer dangling even after a mid-way error) */
    workerFinalizeTemplateLayer(outlinedTextLayer);
    try { doc.activeLayer = restoredTextLayer; } catch (eActive) {}
    return "OK:" + restoreReport.restored + (restoreReport.fontFallback ? ":FONT" : "");
}

// Register every worker function here (prevents registration misses)
var WORKER_FUNCS = [
    workerRound,
    workerToggleAttributesPanel,
    workerSetLayerUsable,
    workerBuildMemoText,
    workerProcessTextFrame,
    workerRunOutline,
    workerFormatNoteForDisplay,
    workerInspectSelection,
    workerExtractTextAttributes,
    workerAlignTextFrameByBounds,
    workerCreateOutlineStashLayer,
    workerCreateRestoredTextLayer,
    workerMergeExistingOutlinedLayers,
    workerNormalizeOutlinedLayerNames,
    workerCleanupOutlinedDuplicateNames,
    workerSetActiveLayerStrict,
    workerApplyTemplateLayerAttribute,
    workerUnloadTemplateAction,
    workerFinalizeTemplateLayer,
    workerColorToText,
    workerColorFromText,
    workerJustificationToText,
    workerResolveJustification,
    workerResolveKernType,
    workerApplyKerning,
    workerCreateRestoredTextFrame,
    workerMoveToOutlinedLayer,
    workerRestoreItem,
    workerRestoreText
];

// ==============================
// Delegation to main engine
// ==============================
function buildWorkerSource(funcs, entryCall) {
    var source = "";
    for (var i = 0; i < funcs.length; i++) {
        source += funcs[i].toString() + "\n";
    }
    return source + entryCall;
}

/* when funcs is omitted, send all workers. for lightweight calls (e.g. polling), pass only the needed functions */
function callWorker(entryCall, funcs) {
    var workerFuncs = funcs || WORKER_FUNCS;
    var resultHolder = { value: null };
    var bridge = new BridgeTalk();
    bridge.target = "illustrator";
    var code = buildWorkerSource(workerFuncs, entryCall);
    bridge.body = "eval(decodeURIComponent(\"" + encodeURIComponent(code) + "\"));";
    bridge.onResult = function (response) { resultHolder.value = String(response.body); };
    bridge.onError = function (errorResponse) { resultHolder.value = "ERR:" + String(errorResponse.body); };
    bridge.send(60); // sync wait limit (seconds). kept long to allow restoring many objects and running the action
    return resultHolder.value;
}

// ==============================
// Palette
// ==============================
// keep the palette reference on $.global (persist across IIFE runs to prevent GC / multiple launches)
var isBusy = false;

function setStatus(win, message) {
    win.statusText.text = message;
    win.layout.layout(true);
}

function applyResultToStatus(win, result, doneKey) {
    if (result == null) { setStatus(win, L('status.err')); return; }
    if (result.indexOf("OK") === 0) {
        var parts = result.split(":");
        var count = (parts.length > 1) ? parts[1] : "";
        var message = L(doneKey) + (count ? " (" + count + ")" : "");
        if (result.indexOf("FONT") >= 0) { message += " / " + L('status.fontWarn'); }
        setStatus(win, message);
    } else if (result === "NODOC") { setStatus(win, L('status.nodoc'));
    } else if (result === "NOSEL") { setStatus(win, L('status.nosel'));
    } else if (result === "NOTGT") { setStatus(win, L('status.notgt'));
    } else if (result === "NONOTE") { setStatus(win, L('status.nonote'));
    } else if (result.indexOf("ERR") === 0) { setStatus(win, L('status.err') + ": " + result.substring(4));
    } else { setStatus(win, result); }
}

function onOutlineClick(win) {
    if (isBusy) return; // prevent re-entry
    isBusy = true;
    setStatus(win, L('status.busy'));
    var result = callWorker("workerRunOutline();");
    isBusy = false;
    applyResultToStatus(win, result, 'status.doneOutline');
    refreshSelectedNote(win, true); // show the note right after outlining
}

function onRestoreClick(win) {
    if (isBusy) return;
    // run the "load" logic right before restore to show the target note (state before it moves)
    refreshSelectedNote(win, true);
    isBusy = true;
    setStatus(win, L('status.busy'));
    var result = callWorker("workerRestoreText();");
    isBusy = false;
    applyResultToStatus(win, result, 'status.doneRestore');
}

/* Toggle the Attributes panel (delegate the menu command to the main engine) */
function onAttributesClick(win) {
    if (isBusy) return;
    callWorker("workerToggleAttributesPanel();", [workerToggleAttributesPanel]);
}

function populateNoteList(win, formattedNote) {
    win.noteList.removeAll();
    var noteLines = formattedNote.split("\n");
    var lineIndex;
    for (lineIndex = 0; lineIndex < noteLines.length; lineIndex++) {
        var currentLine = noteLines[lineIndex];
        if (!currentLine) { continue; }
        var separatorPos = currentLine.indexOf(": ");
        var itemLabel, itemValue;
        if (separatorPos >= 0) {
            itemLabel = currentLine.substring(0, separatorPos);
            itemValue = currentLine.substring(separatorPos + 2);
        } else {
            itemLabel = currentLine;
            itemValue = "";
        }
        var row = win.noteList.add("item", itemLabel);
        row.subItems[0].text = itemValue;
    }
}

function refreshSelectedNote(win, keepStatus) {
    if (isBusy) return;
    isBusy = true;
    var result = callWorker("workerInspectSelection();");
    isBusy = false;
    win.noteList.removeAll();
    if (result != null && result.indexOf("NOTE:") === 0) {
        populateNoteList(win, result.substring(5));
        if (!keepStatus) { setStatus(win, L('status.memoLoaded')); }
    } else if (result === "NONOTE") {
        if (!keepStatus) { setStatus(win, L('status.nonote')); }
    } else if (result === "NOSEL") {
        if (!keepStatus) { setStatus(win, L('status.nosel')); }
    } else if (result === "NODOC") {
        if (!keepStatus) { setStatus(win, L('status.nodoc')); }
    } else if (!keepStatus) {
        setStatus(win, L('status.err'));
    }
}

/* cleanup on palette close: unload the loaded action */
function performCloseCleanup() {
    try { callWorker("workerUnloadTemplateAction();"); } catch (e) {}
    $.global.__textOutlineMemoPalette = null;
}

function showPalette() {
    // prevent multiple launches: close an existing palette (kept on $.global to persist across IIFE runs)
    if ($.global.__textOutlineMemoPalette) {
        try { $.global.__textOutlineMemoPalette.close(); } catch (e) {}
        $.global.__textOutlineMemoPalette = null;
    }

    var win = new Window("palette", L('dialog.title') + ' ' + SCRIPT_VERSION, undefined, { resizeable: false });
    setupWindow(win);

    // selected object's note
    var selectedObjectPanel = win.add("panel", undefined, L('panel.selected'));
    setupPanel(selectedObjectPanel, 6);

    win.noteList = selectedObjectPanel.add("listbox", undefined, [], {
        numberOfColumns: 2,
        showHeaders: true,
        columnTitles: [L('listCol.item'), L('listCol.value')],
        columnWidths: [140, 170]
    });
    win.noteList.preferredSize = [320, 280];

    // note action row: left = Attributes / center = spacer / right = Load Note
    var noteActionRow = selectedObjectPanel.add("group");
    noteActionRow.orientation = "row";
    noteActionRow.alignment = "fill"; // stretch to full panel width and push children to both edges
    noteActionRow.alignChildren = ["fill", "center"];
    noteActionRow.margins = [0, 6, 0, 0]; // +6 top margin above the row
    noteActionRow.spacing = 8;

    // left: toggle the Attributes panel
    var attributesButton = noteActionRow.add("button", undefined, L('button.attributes'));
    attributesButton.helpTip = L('button.attributesTip');
    attributesButton.alignment = ["left", "center"];
    attributesButton.onClick = function () { onAttributesClick(win); };

    // center: flexible spacer (push the side buttons to both edges)
    var noteActionSpacer = noteActionRow.add("group");
    noteActionSpacer.alignment = ["fill", "center"];

    // right: Load Note
    var loadNoteButton = noteActionRow.add("button", undefined, L('button.load'));
    loadNoteButton.helpTip = L('button.loadTip');
    loadNoteButton.alignment = ["right", "center"];
    loadNoteButton.onClick = function () { refreshSelectedNote(win, false); };

    // command buttons
    var commandPanel = win.add("panel", undefined, L('panel.commands'));
    setupPanel(commandPanel, 6);

    var commandButtonRow = commandPanel.add("group");
    setupRow(commandButtonRow, "left", 8); // left-align the button row (do not stretch to full width)

    var outlineButton = commandButtonRow.add("button", undefined, L('button.outline'));
    outlineButton.helpTip = L('button.outlineTip');
    outlineButton.onClick = function () { onOutlineClick(win); };

    var restoreButton = commandButtonRow.add("button", undefined, L('button.restore'));
    restoreButton.helpTip = L('button.restoreTip');
    restoreButton.onClick = function () { onRestoreClick(win); };

    // status
    win.statusText = win.add("statictext", undefined, L('status.ready'));
    win.statusText.alignment = "left";

    // close on Esc
    win.addEventListener("keydown", function (ev) {
        if (ev.keyName == "Escape") {
            try { win.close(); } catch (e) {}
        }
    });

    // on close (x / Esc), unload the loaded action
    win.onClose = function () {
        performCloseCleanup();
        return true;
    };

    $.global.__textOutlineMemoPalette = win;
    refreshSelectedNote(win, true); // show the selected object's note at launch
    win.show();
    return win;
}

// ==============================
// Run
// ==============================
showPalette();

})();
