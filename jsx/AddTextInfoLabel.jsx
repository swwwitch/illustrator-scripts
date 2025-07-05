// Illustrator用のJavaScript

#target illustrator

/*
 * スクリプトの概要：
 * 選択されたテキストオブジェクトの下または右に、フォント情報を表示します。
 * 表示形式：「詳細」（エリア内文字・右揃え）または「簡易版」（ポイント文字・左揃え）を選択可能。
 * ダイアログでは表示項目を個別にON/OFFでき、すべてON／すべてOFF／最小セットのボタンも用意されています。
 * 「簡易版」を選択中は、表示項目セクションをディム表示します（無効化）。
 *
 * 【詳細表示】（エリア内文字）：
 * - 行数に応じて自動で高さを計算し、さらに2行分の余白を追加します。
 * - 「下」を選ぶとアウトライン化を想定した下端に、「右」を選ぶと左下基準で右に配置します。
 * - spacing に 2mm の余白が自動で追加されます。
 * - 情報は右揃えで整形されます。
 * - 行送りは「フォントサイズ × 1.6」で固定されます。
 *
 * 【簡易表示】（ポイント文字）：
 * - 文字サイズ分下、または右に配置され、左揃えで表示されます。
 * - 文字ツメには「%」を付けて表示します。
 * - 行送りは「フォントサイズ × 1.6」で固定されます。
 *
 * 【共通仕様】：
 * - 行送りは「自動（xx pt）」または「xx pt」で表示。
 * - 行送り（%）は実際の行送り値 ÷ フォントサイズ で計算し、小数第2位で四捨五入・ゼロも保持して表示。
 * - 文字ツメは小数第2位で四捨五入、末尾.0は省略（例：21 → 21、20.1 → 20.1）。
 * - 情報は「フォント情報」レイヤーに配置され、生成されたテキストは選択状態になります。
 *
 * 作成日：2025-04-20
 * 更新日：2025-04-23（表示項目の無効化制御とspacing設定を追加）
 * 最終更新日：2025-04-24（行送り（%）の誤判定を修正し、正確な表示に対応）
 */

main();

function getFontSizeUnitLabel() {
    var textUnit = app.preferences.getIntegerPreference("text/units");
    switch (textUnit) {
        case 0: return "inch";
        case 1: return "mm";
        case 2: return "pt";
        case 3: return "p";
        case 4: return "cm";
        case 5: return "Q";
        case 6: return "px";
        default: return "pt";
    }
}

function getLeadingUnitLabel() {
    var asianUnit = app.preferences.getIntegerPreference("text/asianunits");
    switch (asianUnit) {
        case 0: return "inch";
        case 1: return "mm";
        case 2: return "pt";
        case 3: return "p";
        case 4: return "cm";
        case 5: return "H";
        case 6: return "px";
        default: return "pt";
    }
}

function getFontSize(item) {
    try {
        var rawSize = item.textRange.characterAttributes.size;
        var size = Math.round(rawSize * 10) / 10;
        return size + " " + getFontSizeUnitLabel();
    } catch (e) {
        return "不明";
    }
}

function formatNumber(num) {
    var str = String(num);
    if (str.indexOf(".") === -1) {
        return str;
    }

    var decimals = str.split(".")[1];
    if (decimals.length <= 2) {
        return str;
    } else {
        return String(Math.round(num * 100) / 100);
    }
}

function getLeading(item, displayMode) {
    try {
        var attr = item.textRange.characterAttributes;
        var leading = attr.leading;
        var unit = getLeadingUnitLabel();
        var formatted = formatNumber(leading);

        if (attr.autoLeading) {
            return (displayMode === "compact")
                ? formatted + " " + unit
                : "自動（" + formatted + " " + unit + "）";
        } else {
            return formatted + " " + unit;
        }
    } catch (e) {
        return "不明";
    }
}

function getLeadingPercentage(item) {
    try {
        var attr = item.textRange.characterAttributes;
        var size = attr.size;
        var leading = attr.leading;

        // 安全チェック：数値でない場合は不明
        if (typeof size !== "number" || typeof leading !== "number" || size === 0) {
            return "不明";
        }

        var percent = (leading / size) * 100;
        return roundToTwoDecimalFixed(percent) + " %";
    } catch (e) {
        return "不明";
    }
}

function roundToTwoDecimalFixed(value) {
    var rounded = Math.round(value * 100) / 100;
    var str = rounded.toFixed(2);
    // 小数点第2位が "00" → ".0" or 整数に丸め
    if (str.match(/\.00$/)) return String(parseInt(rounded, 10));
    if (str.match(/\.0$/)) return String(rounded.toFixed(1));
    return str;
}

function main() {
    if (app.documents.length === 0) return;

    var doc = app.activeDocument;

    if (app.selection.constructor.name === "TextRange") {
        var textFramesInStory = app.selection.story.textFrames;
        if (textFramesInStory.length === 1) {
            app.executeMenuCommand("deselectall");
            app.selection = [textFramesInStory[0]];
            try { app.selectTool("Adobe Select Tool"); } catch (e) {}
        }
    }

    var selectedItems = doc.selection;
    if (!selectedItems || selectedItems.length === 0 || selectedItems.length >= 1000) return;

    var dialogOptions = showOptionDialog();
    if (!dialogOptions) return;

    var infoLayer = getOrCreateLayer(doc, "フォント情報");
    var generatedItems = [];

    for (var i = 0; i < selectedItems.length; i++) {
        var originalItem = selectedItems[i];
        if (originalItem.typename !== "TextFrame") continue;

        var fontObj = getFirstAvailableFont(originalItem);
        if (!fontObj) continue;

        var fontSize = getFontSize(originalItem);
        var leadingText = getLeading(originalItem, dialogOptions.displayMode);
        var leadingPercent = getLeadingPercentage(originalItem); 
        var kerningText = getKerningMethodText(originalItem);
        var tracking = getTracking(originalItem);
        var proportionalMetrics = getProportionalMetrics(originalItem);
        var tsume = getTsume(originalItem);
        var fontFamily = fontObj.family;
        var fontStyle = fontObj.style;
        var postScriptName = fontObj.name;
        var displayFontSize = 10;
        var displayFont = "HiraginoSans-W3";

        var textContent;

        if (dialogOptions.displayMode === "compact") {
            var line1 = fontFamily + " " + fontStyle + "、" + fontSize + " ↓" + leadingText;
            var line2 = kerningText + "、プロポーショナルメトリクス：" + proportionalMetrics;
            var line3 = "トラッキング：" + tracking + "、文字ツメ：" + tsume + " %";
            textContent = line1 + String.fromCharCode(13) + line2 + String.fromCharCode(13) + line3;
        } else {
            var infoLines = [];
            if (dialogOptions.includeFontName)     infoLines.push("・フォント名	" + fontFamily);
            if (dialogOptions.includePostScript)   infoLines.push("・PSフォント名	" + postScriptName);
            if (dialogOptions.includeFontStyle)    infoLines.push("・スタイル（ウエイト）	" + fontStyle);
            if (dialogOptions.includeFontSize)     infoLines.push("・フォントサイズ	" + fontSize);
            if (dialogOptions.includeLeading)      infoLines.push("・行送り	" + leadingText);
            if (dialogOptions.includeLeadingPercent) infoLines.push("・行送り（%）\t" + leadingPercent);
            if (dialogOptions.includeKerning)      infoLines.push("・カーニング	" + kerningText);
            if (dialogOptions.includeProportional) infoLines.push("・プロポーショナルメトリクス	" + proportionalMetrics);
            if (dialogOptions.includeTracking)     infoLines.push("・トラッキング	" + tracking);
            if (dialogOptions.includeTsume)        infoLines.push("・文字ツメ	" + tsume + " %");
            textContent = infoLines.join(String.fromCharCode(13));
        }

        var tfInfo;
        var bounds = originalItem.geometricBounds;
        var originalLeft = bounds[0];
        var originalTop = bounds[1];
        var originalRight = bounds[2];
        var originalBottom = bounds[3];

        if (dialogOptions.displayMode === "full") {
            var lineCount = textContent.split(String.fromCharCode(13)).length;
            var lineHeight = displayFontSize * 1.4;
            var height = lineHeight * (lineCount + 2);
            var mmToPt = 72 / 25.4;
            height += 4 * mmToPt;
            var width = 300;

            var rectX = (dialogOptions.position === "right") ? originalRight + 10 : originalLeft;
            var rectY = (dialogOptions.position === "right") ? originalBottom + height : originalBottom - displayFontSize;

            var rectPath = doc.pathItems.rectangle(rectY, rectX, width, height);
            tfInfo = doc.textFrames.areaText(rectPath);
            tfInfo.contents = textContent;

            tfInfo.spacing = 2 * mmToPt;

            // ↓ この行を追加
           var attr = tfInfo.textRange.characterAttributes;
           attr.autoLeading = false;
           attr.leading = 16;

            // タブストップ追加（右揃え、位置400pt、リーダー…）
            var tabStop = new TabStopInfo();
            tabStop.position = 400;
            tabStop.alignment = TabStopAlignment.Right;
            tabStop.leader = "…";
            for (var p = 0; p < tfInfo.paragraphs.length; p++) {
                tfInfo.paragraphs[p].tabStops = [tabStop];
            }
        } else {
            var posX = (dialogOptions.position === "right") ? originalRight + 20 : originalLeft;
            var posY = (dialogOptions.position === "right") ? originalTop : originalBottom - displayFontSize;
            tfInfo = createPointText(doc, textContent, [posX, posY]);
            var infoBounds = tfInfo.visibleBounds;
            tfInfo.translate(0, posY - infoBounds[1]);
        }

        // 行送り設定を追加
        var attr = tfInfo.textRange.characterAttributes;
        attr.autoLeading = false;
        attr.leading = 16;

        tfInfo.textRange.characterAttributes.size = displayFontSize;
        tfInfo.textRange.paragraphAttributes.justification =
            (dialogOptions.displayMode === "full") ? Justification.RIGHT : Justification.LEFT;

        try {
            tfInfo.textRange.characterAttributes.textFont = textFonts.getByName(displayFont);
        } catch (e) {}

        tfInfo.move(infoLayer, ElementPlacement.PLACEATBEGINNING);
        generatedItems.push(tfInfo);
        originalItem.selected = false;
    }

    if (generatedItems.length > 0) {
        app.selection = generatedItems;
    }
    app.redraw(); // 画面再描画
}

function showOptionDialog() {
    var dlg = new Window("dialog", "テキスト情報を追加");
    dlg.alignChildren = "left";

    var topGroup = dlg.add("group");
    topGroup.orientation = "row";
    topGroup.alignChildren = "top";

    var posGroup = topGroup.add("panel", undefined, "位置");
    posGroup.orientation = "row";
    posGroup.margins = [15, 20, 15, 15];
    var posBottom = posGroup.add("radiobutton", undefined, "下");
    var posRight = posGroup.add("radiobutton", undefined, "右");
    posRight.value = true;

    var modeGroup = topGroup.add("panel", undefined, "表示形式");
    modeGroup.orientation = "row";
    modeGroup.margins = [15, 20, 15, 15];
    var modeCompact = modeGroup.add("radiobutton", undefined, "簡易版");
    var modeFull = modeGroup.add("radiobutton", undefined, "詳細");
    modeCompact.value = true;

    var infoGroup = dlg.add("panel", undefined, "表示項目（詳細表示時のみ有効）");
    infoGroup.orientation = "column";
    infoGroup.alignChildren = "left";
    infoGroup.margins = [15, 20, 15, 15];

    var columnsGroup = infoGroup.add("group");
    columnsGroup.orientation = "row";
    columnsGroup.alignChildren = "top";

    var column1 = columnsGroup.add("group");
    column1.orientation = "column";
    column1.alignChildren = "left";

    var column2 = columnsGroup.add("group");
    column2.orientation = "column";
    column2.alignChildren = "left";

    var chkFontName     = column1.add("checkbox", undefined, "フォント名");
    var chkPostScript   = column1.add("checkbox", undefined, "PSフォント名");
    var chkFontStyle    = column1.add("checkbox", undefined, "スタイル（ウエイト）");
    var chkFontSize     = column1.add("checkbox", undefined, "フォントサイズ");
    var chkLeading      = column1.add("checkbox", undefined, "行送り");
    var chkKerning      = column2.add("checkbox", undefined, "カーニング");
    var chkProportional = column2.add("checkbox", undefined, "プロポーショナルメトリクス");
    var chkTracking     = column2.add("checkbox", undefined, "トラッキング");
    var chkTsume        = column2.add("checkbox", undefined, "文字ツメ");
    var chkLeadingPercent = column2.add("checkbox", undefined, "行送り（%）");

    var allToggles = [
        chkFontName, chkPostScript, chkFontStyle, chkFontSize, chkLeading,
        chkKerning, chkProportional, chkTracking, chkTsume, chkLeadingPercent
    ];

    chkFontName.value     = true;
    chkPostScript.value   = false;
    chkFontStyle.value    = true;
    chkFontSize.value     = true;
    chkLeading.value      = true;
    chkKerning.value      = true;
    chkProportional.value = true;
    chkTracking.value     = false;
    chkTsume.value        = false;
    chkLeadingPercent.value = false;

    var toggleGroup = infoGroup.add("group");
    toggleGroup.orientation = "row";
    toggleGroup.alignment = "left";
    var btnAllOn  = toggleGroup.add("button", [0, 0, 80, 24], "すべてON");
    var btnAllOff = toggleGroup.add("button", [0, 0, 80, 24], "すべてOFF");
    var btnMinimal = toggleGroup.add("button", [0, 0, 80, 24], "最小セット");

    btnAllOn.onClick = function () {
        for (var i = 0; i < allToggles.length; i++) allToggles[i].value = true;
    };
    btnAllOff.onClick = function () {
        for (var i = 0; i < allToggles.length; i++) allToggles[i].value = false;
    };

btnMinimal.onClick = function () {
    chkFontName.value     = true;
    chkFontStyle.value    = true;
    chkFontSize.value     = true;
    chkLeading.value      = true;

    chkPostScript.value   = false;
    chkKerning.value      = false;
    chkProportional.value = false;
    chkTracking.value     = false;
    chkTsume.value        = false;
    chkLeadingPercent.value = false;
};

    function updateInfoGroupState() {
        var enabled = modeFull.value;
        infoGroup.enabled = enabled;
    }

    modeFull.onClick = updateInfoGroupState;
    modeCompact.onClick = updateInfoGroupState;
    updateInfoGroupState(); // 初期化

    var btnGroup = dlg.add("group");
    btnGroup.alignment = "right";
    btnGroup.add("button", undefined, "キャンセル", {name: "cancel"});
    btnGroup.add("button", undefined, "OK", {name: "ok"});

    if (dlg.show() !== 1) return null;

    return {
        position: posBottom.value ? "bottom" : "right",
        displayMode: modeFull.value ? "full" : "compact",
        includeFontName:   chkFontName.value,
        includePostScript: chkPostScript.value,
        includeFontStyle:  chkFontStyle.value,
        includeFontSize:   chkFontSize.value,
        includeLeading:    chkLeading.value,
        includeKerning:    chkKerning.value,
        includeTracking:   chkTracking.value,
        includeProportional: chkProportional.value,
        includeTsume: chkTsume.value,
        includeLeadingPercent: chkLeadingPercent.value
    };
}

function getFirstAvailableFont(textFrame) {
    var chars = textFrame.textRange.characters;
    for (var i = 0; i < chars.length; i++) {
        var fnt = chars[i].characterAttributes.textFont;
        if (fnt) return fnt;
    }
    return null;
}

function getKerningMethodText(item) {
    try {
        var method = item.textRange.characterAttributes.kerningMethod;
        switch (method) {
            case AutoKernType.AUTO: return "メトリクス";
            case AutoKernType.METRICSROMANONLY: return "和文等幅";
            case AutoKernType.OPTICAL: return "オプティカル";
            default: return "なし";
        }
    } catch (e) {
        return "不明";
    }
}

function getTracking(item) {
    try {
        return item.textRange.characterAttributes.tracking;
    } catch (e) {
        return "不明";
    }
}

function getProportionalMetrics(item) {
    try {
        return item.textRange.characterAttributes.proportionalMetrics ? "ON" : "OFF";
    } catch (e) {
        return "不明";
    }
}

function getTsume(item) {
    try {
        var tsume = item.textRange.characterAttributes.Tsume;
        if (typeof tsume !== "number" || isNaN(tsume)) return "なし";

        var rounded = Math.round(tsume * 10) / 10;
        return (rounded % 1 === 0) ? String(rounded.toFixed(0)) : String(rounded.toFixed(1));
    } catch (e) {
        return "不明";
    }
}

function getOrCreateLayer(doc, name) {
    for (var i = 0; i < doc.layers.length; i++) {
        if (doc.layers[i].name === name) return doc.layers[i];
    }
    var layer = doc.layers.add();
    layer.name = name;
    return layer;
}

function createPointText(doc, contents, position) {
    var tf = doc.textFrames.add();
    tf.contents = contents;
    tf.position = position;
    return tf;
}