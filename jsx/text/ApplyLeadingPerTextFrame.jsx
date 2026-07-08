#target illustrator
#targetengine "ApplyLeadingPerTextFrame"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

// =========================================
// バージョン / Version
// =========================================
var SCRIPT_VERSION = "v1.1.0";

/*

### ApplyLeadingPerTextFrame （行送りの設定 / 常駐パレット）

選択したテキストフレームの行送りを「自動行送り量（％）」として設定する常駐パレットです。
行送り値を直接書き込まず、フォントサイズに対する割合（％）を求めて autoLeadingAmount に代入し、
行送りは常に自動（autoLeading = true）にします。行送りが自動になるため、行ごとにフォントサイズが
異なっても各行の行送りは自動的に追従します。

- 110% / 125% / 150% は比率×100 をそのまま自動行送り量に代入します。
- ［その他］は行送り値（pt 等）を直接入力する指定で、「指定行送り ÷ フォントサイズ × 100」を段落ごとに算出します。
- ％欄を直接編集するとプリセットの選択は解除され、その値がそのまま自動行送り量になります。
- 段落前後のアキ・行送りの基準（autoLeadingType）も同時に設定します。

常駐パレットの app は表示中に DOM 接続を失うため、DOM を触る処理は worker 関数にまとめ、
操作のたびに BridgeTalk でメインエンジンへ委譲します。値は絶対値で上書きする冪等な処理のため、
UI を変更するたびに選択中のテキストフレームへ即適用します。結果はそのまま残り、取り消しは Cmd+Z です。

### ApplyLeadingPerTextFrame (Leading Settings / persistent palette)

A persistent palette that sets the leading of selected text frames as an auto-leading amount (%).
Instead of writing an explicit leading value, it derives the ratio of leading to font size (%),
assigns it to autoLeadingAmount, and always turns auto leading on. Because leading becomes auto,
each line follows its own font size automatically.

- 110% / 125% / 150% assign ratio×100 directly to the auto leading amount.
- "Other" lets you type an explicit leading value; the % is computed per paragraph as leading ÷ font size × 100.
- Editing the % field clears the preset selection and uses that value as the auto leading amount.
- Paragraph spacing and the leading basis (autoLeadingType) are applied together.

Because a resident palette loses its DOM connection while shown, all DOM work is collected into
worker functions and delegated to the main engine over BridgeTalk on each action. Since the values
are absolute overwrites (idempotent), every UI change is applied live to the selected text frames.
Results stay as-is; undo with Cmd+Z.

更新日 / Updated: 2026-07-08

*/

(function () {

    // =========================================
    // ローカライズ / Localization
    // =========================================
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }

    var lang = getCurrentLang();

    var LABELS = {
        dialog: {
            title: { ja: "行送りの設定", en: "Leading Settings" }
        },
        panel: {
            leading: { ja: "行送り", en: "Leading settings" },
            type: { ja: "行送りの基準", en: "Leading basis" },
            space: { ja: "段落前後のアキ", en: "Paragraph spacing" }
        },
        type: {
            topToTop: { ja: "仮想ボディの上基準", en: "Top-to-top (virtual body)" },
            bottomToBottom: { ja: "欧文ベースライン基準", en: "Baseline-to-baseline" }
        },
        label: {
            spaceBefore: { ja: "段落前：", en: "Space before:" },
            spaceAfter: { ja: "段落後：", en: "Space after:" }
        },
        choice: {
            other: { ja: "その他", en: "Other" }
        },
        tip: {
            amount: { ja: "自動行送り量（％）。↑↓：±1 / Shift＋↑↓：±10 / Option＋↑↓：±0.1", en: "Auto leading amount (%). ↑↓: ±1 / Shift+↑↓: ±10 / Option+↑↓: ±0.1" },
            leading: { ja: "行送り値（［その他］で直接指定）。↑↓：±1 / Shift＋↑↓：±10 / Option＋↑↓：±0.1", en: "Leading value (use Other to set directly). ↑↓: ±1 / Shift+↑↓: ±10 / Option+↑↓: ±0.1" },
            space: { ja: "↑↓：±1 / Shift＋↑↓：±10 / Option＋↑↓：±0.1", en: "↑↓: ±1 / Shift+↑↓: ±10 / Option+↑↓: ±0.1" }
        }
    };

    /**
     * ドットパスでローカライズ文字列を取得する（null 耐性あり）。
     * @param {string} path - "panel.leading" のようなドット区切りキー
     * @returns {string} 該当言語の文字列。無ければ英語、さらに無ければ path をそのまま返す
     */
    function L(path) {
        var parts = String(path).split(".");
        var node = LABELS;
        for (var i = 0; i < parts.length; i++) {
            if (node == null) return path;
            node = node[parts[i]];
        }
        if (node == null) return path;
        if (typeof node[lang] === "string") return node[lang];
        if (typeof node.en === "string") return node.en;
        return path;
    }


    // =========================================
    // 単位 / Units
    // =========================================
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


    // =========================================
    // 行送りの選択肢 / Leading choices
    // =========================================
    var LEADING_CHOICES = [
        { label: "110%", ratio: 1.1, token: "110", other: false },
        { label: "125%", ratio: 1.25, token: "125", other: false },
        { label: "150%", ratio: 1.5, token: "150", other: false },
        { label: L("choice.other"), ratio: undefined, token: "OTHER", other: true }
    ];

    function getLeadingChoiceIndexByToken(token) {
        for (var i = 0; i < LEADING_CHOICES.length; i++) {
            if (LEADING_CHOICES[i].token === token) return i;
        }
        return 0;
    }

    function getLeadingTypeChoices() {
        return [
            { label: L("type.topToTop"), token: "TOPTOTOP" },
            { label: L("type.bottomToBottom"), token: "BOTTOMTOBOTTOM" }
        ];
    }


    // =========================================
    // worker 関数 / Worker functions (run in the MAIN engine via BridgeTalk)
    // 注意: 内部は // 行コメント禁止・/* */ のみ・必ずセミコロンで終える（toString が改行を消すため）
    // =========================================

    function w_normalizeSelection() {
        if (app.documents.length > 0 && app.selection && app.selection.typename === "TextRange") {
            var story = app.selection.story;
            if (story && story.textFrames.length === 1) {
                var parentTextFrame = story.textFrames[0];
                app.executeMenuCommand("deselectall");
                app.selection = [parentTextFrame];
                try { app.selectTool("Adobe Select Tool"); } catch (e) { }
            }
        }
    }

    function w_getProcessableTextFrame(item) {
        if (!item || item.typename !== "TextFrame") { return null; }
        if (!item.contents) { return null; }
        if (!item.lines || item.lines.length === 0) { return null; }
        return item;
    }

    function w_collectFrames() {
        var doc = app.activeDocument;
        var selectionItems = doc.selection;
        var frames = [];
        if (!selectionItems || selectionItems.length === 0) { return frames; }
        var seen = [];
        for (var i = 0; i < selectionItems.length; i++) {
            var tf = w_getProcessableTextFrame(selectionItems[i]);
            if (!tf) { continue; }
            var dup = false;
            for (var j = 0; j < seen.length; j++) { if (seen[j] === tf) { dup = true; break; } }
            if (dup) { continue; }
            seen.push(tf);
            frames.push(tf);
        }
        return frames;
    }

    function w_lineSampleSizes(line, sampleCount) {
        var sizes = [];
        if (!line || !line.characters || line.characters.length === 0) { return sizes; }
        var maxCount = Math.min(line.characters.length, sampleCount);
        for (var i = 0; i < maxCount; i++) {
            try {
                var size = line.characters[i].characterAttributes.size;
                if (!isNaN(size)) { sizes.push(size); }
            } catch (e) { }
        }
        return sizes;
    }

    function w_mostFrequent(values) {
        if (!values || values.length === 0) { return NaN; }
        var counts = {};
        var bestValue = values[0];
        var bestCount = 0;
        for (var i = 0; i < values.length; i++) {
            var key = String(values[i]);
            if (!counts[key]) { counts[key] = { value: values[i], count: 0 }; }
            counts[key].count++;
            if (counts[key].count > bestCount) { bestCount = counts[key].count; bestValue = counts[key].value; }
            else if (counts[key].count === bestCount && counts[key].value > bestValue) { bestValue = counts[key].value; }
        }
        return bestValue;
    }

    function w_baseFontParagraph(paragraph) {
        var all = [];
        try {
            var lines = paragraph.lines;
            for (var j = 0; j < lines.length; j++) {
                var sizes = w_lineSampleSizes(lines[j], LINE_FONT_SIZE_SAMPLE_COUNT);
                for (var s = 0; s < sizes.length; s++) { all.push(sizes[s]); }
            }
        } catch (e) { }
        if (all.length === 0) { return NaN; }
        return w_mostFrequent(all);
    }

    function w_commonBaseFont(textFrame) {
        var all = [];
        var lines = textFrame.lines;
        for (var j = 0; j < lines.length; j++) {
            var sizes = w_lineSampleSizes(lines[j], LINE_FONT_SIZE_SAMPLE_COUNT);
            for (var s = 0; s < sizes.length; s++) { all.push(sizes[s]); }
        }
        if (all.length === 0) { return NaN; }
        return w_mostFrequent(all);
    }

    function w_resolveLeadingType(token) {
        if (token === "BOTTOMTOBOTTOM") { return AutoLeadingType.BOTTOMTOBOTTOM; }
        return AutoLeadingType.TOPTOTOP;
    }

    function w_applyLeading(autoAmount, directMode, directLeadingPt, spaceBefore, spaceAfter, leadingTypeToken) {
        if (app.documents.length === 0) { return "NODOC"; }
        w_normalizeSelection();
        var frames = w_collectFrames();
        if (!frames || frames.length === 0) { return "NOSEL"; }
        var leadingType = w_resolveLeadingType(leadingTypeToken);
        var useDirect = directMode && !isNaN(directLeadingPt);
        var repPt = NaN;
        try {
            for (var i = 0; i < frames.length; i++) {
                var tf = frames[i];
                tf.textRange.characterAttributes.autoLeading = true;
                var paragraphs = tf.paragraphs;
                for (var p = 0; p < paragraphs.length; p++) {
                    try {
                        var paragraph = paragraphs[p];
                        var amount;
                        if (useDirect) {
                            var baseFont = w_baseFontParagraph(paragraph);
                            if (isNaN(baseFont) || baseFont <= 0) { continue; }
                            amount = (directLeadingPt / baseFont) * 100;
                        } else {
                            amount = autoAmount;
                        }
                        paragraph.characterAttributes.autoLeading = true;
                        paragraph.paragraphAttributes.spaceBefore = spaceBefore;
                        paragraph.paragraphAttributes.spaceAfter = spaceAfter;
                        paragraph.paragraphAttributes.autoLeadingAmount = amount;
                    } catch (ep) { }
                }
                try { tf.textRange.leadingType = leadingType; } catch (et) { }
            }
            if (useDirect) {
                repPt = directLeadingPt;
            } else {
                var baseCommon = w_commonBaseFont(frames[0]);
                if (!isNaN(baseCommon)) { repPt = baseCommon * (autoAmount / 100); }
            }
        } catch (e) {
            return "ERR:" + e.message;
        }
        app.redraw();
        return "OK|" + repPt;
    }

    function w_readInitial() {
        if (app.documents.length === 0) { return "NODOC"; }
        w_normalizeSelection();
        var frames = w_collectFrames();
        if (!frames || frames.length === 0) { return "NOSEL"; }
        var autoAmount = 175;
        var leadingPt = NaN;
        var leadingTypeToken = "TOPTOTOP";
        var spaceBefore = 0;
        var spaceAfter = 0;
        var choiceToken = "OTHER";
        var isAuto = false;
        for (var i = 0; i < frames.length; i++) {
            try {
                var lines = frames[i].lines;
                if (lines && lines.length > 0 && lines[0].characters.length > 0) {
                    var chAttr = lines[0].characters[0].characterAttributes;
                    isAuto = chAttr.autoLeading;
                    if (!isNaN(chAttr.leading)) { leadingPt = chAttr.leading; }
                }
            } catch (e) { }
            if (!isNaN(leadingPt)) { break; }
        }
        for (var k = 0; k < frames.length; k++) {
            try {
                var paras = frames[k].paragraphs;
                if (paras && paras.length > 0) {
                    var pa = paras[0].paragraphAttributes;
                    if (!isNaN(pa.autoLeadingAmount)) { autoAmount = pa.autoLeadingAmount; }
                    if (!isNaN(pa.spaceBefore)) { spaceBefore = pa.spaceBefore; }
                    if (!isNaN(pa.spaceAfter)) { spaceAfter = pa.spaceAfter; }
                    break;
                }
            } catch (e2) { }
        }
        for (var t = 0; t < frames.length; t++) {
            try {
                var lt = frames[t].textRange.leadingType;
                if (lt !== undefined && lt !== null) {
                    if (lt === AutoLeadingType.BOTTOMTOBOTTOM) { leadingTypeToken = "BOTTOMTOBOTTOM"; }
                    else { leadingTypeToken = "TOPTOTOP"; }
                    break;
                }
            } catch (e3) { }
        }
        if (isAuto) {
            if (Math.abs(autoAmount - 110) < 0.5) { choiceToken = "110"; }
            else if (Math.abs(autoAmount - 125) < 0.5) { choiceToken = "125"; }
            else if (Math.abs(autoAmount - 150) < 0.5) { choiceToken = "150"; }
            else { choiceToken = "OTHER"; }
        } else {
            choiceToken = "OTHER";
        }
        return "OK|" + autoAmount + "|" + leadingPt + "|" + leadingTypeToken + "|" + spaceBefore + "|" + spaceAfter + "|" + choiceToken;
    }

    // worker 関数はすべてここに登録（追加漏れ防止） / Register every worker function here
    var WORKER_FUNCS = [
        w_normalizeSelection,
        w_getProcessableTextFrame,
        w_collectFrames,
        w_lineSampleSizes,
        w_mostFrequent,
        w_baseFontParagraph,
        w_commonBaseFont,
        w_resolveLeadingType,
        w_applyLeading,
        w_readInitial
    ];


    // =========================================
    // BridgeTalk 委譲 / Delegation to the main engine
    // =========================================
    var LINE_FONT_SIZE_SAMPLE_COUNT = 5; // 行内で参照する文字数 / Characters sampled per line
    var isBusy = false;

    /**
     * worker 関数群のソースを連結して 1 つのコード文字列にする。
     * 先頭で共有定数を宣言し、後続の呼び出し式から参照できるようにする。
     * @returns {string} メインエンジンで eval するソース
     */
    function buildWorkerSource() {
        var src = "var LINE_FONT_SIZE_SAMPLE_COUNT=" + LINE_FONT_SIZE_SAMPLE_COUNT + ";";
        for (var i = 0; i < WORKER_FUNCS.length; i++) {
            src += WORKER_FUNCS[i].toString();
        }
        return src;
    }

    /**
     * worker 呼び出し式をメインエンジンへ同期委譲し、マーカー文字列を受け取る。
     * @param {string} callExpr - "w_applyLeading(...)" 等、文字列を返す呼び出し式
     * @returns {string} マーカー（"OK"/"OK|..."/"NODOC"/"NOSEL"/"ERR:..."）
     */
    function runWorker(callExpr) {
        if (isBusy) { return "ERR:busy"; }
        isBusy = true;
        var holder = { value: null };
        try {
            var code = buildWorkerSource() + "String(" + callExpr + ");";
            var payload = encodeURIComponent(code);
            var bridge = new BridgeTalk();
            bridge.target = "illustrator";
            bridge.body = "eval(decodeURIComponent(\"" + payload + "\"));";
            bridge.onResult = function (resObj) { holder.value = resObj.body; };
            bridge.onError = function (errObj) { holder.value = "ERR:" + errObj.body; };
            bridge.send(10);
        } catch (e) {
            holder.value = "ERR:" + e.message;
        } finally {
            isBusy = false;
        }
        return holder.value;
    }

    /**
     * マーカー文字列を解析する。
     * @param {string} res - runWorker の戻り値
     * @returns {object} { ok: boolean, code: string, extra: string, msg: string }
     */
    function parseMarker(res) {
        if (res == null) { return { ok: false, code: "ERR", msg: "no response" }; }
        var head = res;
        var extra = null;
        var barIndex = res.indexOf("|");
        if (barIndex >= 0) {
            head = res.substring(0, barIndex);
            extra = res.substring(barIndex + 1);
        }
        if (head === "OK") { return { ok: true, code: "OK", extra: extra }; }
        if (head === "NODOC") { return { ok: false, code: "NODOC" }; }
        if (head === "NOSEL") { return { ok: false, code: "NOSEL" }; }
        if (head.indexOf("ERR") === 0) { return { ok: false, code: "ERR", msg: res.substring(4) }; }
        return { ok: false, code: "ERR", msg: res };
    }


    // =========================================
    // 数値ユーティリティ / Numeric helpers
    // =========================================
    function toNumber(value, fallback) {
        var n = parseFloat(value);
        return isNaN(n) ? fallback : n;
    }

    /**
     * 手入力値をパースし、下限でクランプする（負数の手入力対策）。
     * @param {string} text - 入力欄の文字列
     * @param {number} minValue - 下限
     * @param {number} fallback - パース失敗時の値
     * @returns {number} クランプ後の値
     */
    function clampMinNumber(text, minValue, fallback) {
        var n = parseFloat(text);
        if (isNaN(n)) { n = fallback; }
        if (n < minValue) { n = minValue; }
        return n;
    }

    function formatByUnit(pt, unit) {
        if (isNaN(pt) || pt === null) { return ""; }
        return (Math.round((pt / unit.factor) * 10) / 10).toFixed(1);
    }


    // =========================================
    // UI: 矢印キーによる数値増減 / Arrow-key stepping
    // =========================================
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


    // =========================================
    // 常駐パレット / Persistent palette
    // =========================================
    var PANEL_MARGINS = [15, 20, 15, 10];

    function showPalette() {
        // 多重起動防止：既存パレットがあれば閉じる / Prevent multiple launches
        if ($.global.__ALPTF_PALETTE__) {
            try { $.global.__ALPTF_PALETTE__.close(); } catch (e) { }
            $.global.__ALPTF_PALETTE__ = null;
        }

        var unit = getTextUnit();

        // 現在の選択から初期値を読む（委譲） / Read initial values from the selection
        var initResult = parseMarker(runWorker("w_readInitial()"));
        var initData = {
            autoAmount: 175,
            leadingPt: NaN,
            leadingTypeToken: "TOPTOTOP",
            spaceBefore: 0,
            spaceAfter: 0,
            choiceToken: "110"
        };
        if (initResult.ok && initResult.extra != null) {
            var fields = initResult.extra.split("|");
            initData.autoAmount = toNumber(fields[0], 175);
            initData.leadingPt = toNumber(fields[1], NaN);
            initData.leadingTypeToken = fields[2] || "TOPTOTOP";
            initData.spaceBefore = toNumber(fields[3], 0);
            initData.spaceAfter = toNumber(fields[4], 0);
            initData.choiceToken = fields[5] || "110";
        }

        var win = new Window("palette", L("dialog.title") + " " + SCRIPT_VERSION, undefined, { resizeable: false });
        win.orientation = "column";
        win.alignChildren = "fill";
        win.margins = 16;
        win.spacing = 12;

        // ---- 行送りパネル / Leading panel ----
        var leadingPanel = win.add("panel", undefined, L("panel.leading"));
        leadingPanel.orientation = "column";
        leadingPanel.alignChildren = "left";
        leadingPanel.margins = PANEL_MARGINS;
        leadingPanel.spacing = 8;

        var contentGroup = leadingPanel.add("group");
        contentGroup.orientation = "row";
        contentGroup.alignChildren = ["left", "top"];
        contentGroup.spacing = 25;

        // 左カラム：行送り値（pt 等） / Left column: leading value
        var leftColumnGroup = contentGroup.add("group");
        leftColumnGroup.orientation = "column";
        leftColumnGroup.alignChildren = "left";
        leftColumnGroup.spacing = 0;

        var leadingGroup = leftColumnGroup.add("group");
        var leadingInput = leadingGroup.add("edittext", undefined, formatByUnit(initData.leadingPt, unit));
        leadingInput.characters = 3;
        leadingInput.helpTip = L("tip.leading");
        leadingGroup.add("statictext", undefined, unit.label);

        // 右カラム：自動行送り量（％）＋プリセット / Right column: auto amount (%) and presets
        var rightColumnGroup = contentGroup.add("group");
        rightColumnGroup.orientation = "column";
        rightColumnGroup.alignChildren = "left";
        rightColumnGroup.spacing = 6;

        var autoAmountGroup = rightColumnGroup.add("group");
        autoAmountGroup.orientation = "row";
        autoAmountGroup.alignChildren = "center";
        autoAmountGroup.spacing = 6;
        var autoInput = autoAmountGroup.add("edittext", undefined, String(Math.round(initData.autoAmount)));
        autoInput.characters = 3;
        autoInput.helpTip = L("tip.amount");
        autoAmountGroup.add("statictext", undefined, "%");

        var radioButtonsGroup = rightColumnGroup.add("group");
        radioButtonsGroup.orientation = "column";
        radioButtonsGroup.alignChildren = "left";
        radioButtonsGroup.spacing = 6;
        radioButtonsGroup.margins = [0, 5, 0, 0];

        var leadingRadios = [];
        for (var i = 0; i < LEADING_CHOICES.length; i++) {
            leadingRadios.push(radioButtonsGroup.add("radiobutton", undefined, LEADING_CHOICES[i].label));
        }
        var initialChoiceIndex = getLeadingChoiceIndexByToken(initData.choiceToken);
        leadingRadios[initialChoiceIndex].value = true;

        // ---- 行送りの基準パネル / Leading type panel ----
        var typeChoices = getLeadingTypeChoices();
        var typeContainer;
        if (lang === "ja") {
            typeContainer = win.add("panel", undefined, L("panel.type"));
            typeContainer.margins = PANEL_MARGINS;
        } else {
            typeContainer = win.add("group");
            typeContainer.margins = [0, 0, 0, 0];
        }
        typeContainer.orientation = "column";
        typeContainer.alignChildren = "left";
        typeContainer.spacing = 8;

        var typeRadios = [];
        var initialTypeIndex = 0;
        for (var t = 0; t < typeChoices.length; t++) {
            typeRadios.push(typeContainer.add("radiobutton", undefined, typeChoices[t].label));
            if (typeChoices[t].token === initData.leadingTypeToken) { initialTypeIndex = t; }
        }
        typeRadios[initialTypeIndex].value = true;

        // ---- 段落前後のアキパネル / Space panel ----
        var spacePanel = win.add("panel", undefined, L("panel.space"));
        spacePanel.orientation = "column";
        spacePanel.alignChildren = "left";
        spacePanel.margins = PANEL_MARGINS;
        spacePanel.spacing = 8;

        var beforeGroup = spacePanel.add("group");
        beforeGroup.add("statictext", undefined, L("label.spaceBefore"));
        var spaceBeforeInput = beforeGroup.add("edittext", undefined, formatByUnit(initData.spaceBefore, unit));
        spaceBeforeInput.characters = 4;
        spaceBeforeInput.helpTip = L("tip.space");
        beforeGroup.add("statictext", undefined, unit.label);

        var afterGroup = spacePanel.add("group");
        afterGroup.add("statictext", undefined, L("label.spaceAfter"));
        var spaceAfterInput = afterGroup.add("edittext", undefined, formatByUnit(initData.spaceAfter, unit));
        spaceAfterInput.characters = 4;
        spaceAfterInput.helpTip = L("tip.space");
        afterGroup.add("statictext", undefined, unit.label);

        // ---- 状態 / State ----
        var isSyncingUI = false;

        function isOtherSelected() {
            for (var r = 0; r < leadingRadios.length; r++) {
                if (leadingRadios[r].value) { return !!LEADING_CHOICES[r].other; }
            }
            return false;
        }

        function selectLeadingChoice(index) {
            for (var r = 0; r < leadingRadios.length; r++) { leadingRadios[r].value = (r === index); }
        }

        function clearLeadingSelection() {
            for (var r = 0; r < leadingRadios.length; r++) { leadingRadios[r].value = false; }
        }

        function currentLeadingTypeToken() {
            for (var r = 0; r < typeRadios.length; r++) {
                if (typeRadios[r].value) { return typeChoices[r].token; }
            }
            return typeChoices[0].token;
        }

        function otherChoiceIndex() {
            for (var r = 0; r < LEADING_CHOICES.length; r++) { if (LEADING_CHOICES[r].other) { return r; } }
            return -1;
        }

        /**
         * UI から適用オプションを読み取る（負数はクランプ）。
         * @returns {object} { invalid, directMode, autoAmount, directLeadingPt, spaceBefore, spaceAfter, leadingTypeToken }
         */
        function readOptions() {
            var directMode = isOtherSelected();
            var autoAmount = clampMinNumber(autoInput.text, 0, 175);
            var spaceBefore = clampMinNumber(spaceBeforeInput.text, 0, 0) * unit.factor;
            var spaceAfter = clampMinNumber(spaceAfterInput.text, 0, 0) * unit.factor;
            var leadingTypeToken = currentLeadingTypeToken();
            var directLeadingPt = NaN;
            if (directMode) {
                var lv = parseFloat(leadingInput.text);
                if (isNaN(lv)) { return { invalid: true }; }
                if (lv < 0) { lv = 0; }
                directLeadingPt = lv * unit.factor;
            }
            return {
                invalid: false,
                directMode: directMode,
                autoAmount: autoAmount,
                directLeadingPt: directLeadingPt,
                spaceBefore: spaceBefore,
                spaceAfter: spaceAfter,
                leadingTypeToken: leadingTypeToken
            };
        }

        /**
         * 現在の UI 値を選択中のテキストフレームへ即適用する。
         * 絶対値で上書きする冪等な処理なので、操作のたびに呼んでも累積しない。
         * @returns {void}
         */
        function applyToSelection() {
            if (isSyncingUI) { return; }
            var opts = readOptions();
            if (opts.invalid) { return; }

            var call = "w_applyLeading(" +
                opts.autoAmount + "," +
                opts.directMode + "," +
                opts.directLeadingPt + "," +
                opts.spaceBefore + "," +
                opts.spaceAfter + ",'" +
                opts.leadingTypeToken + "'" +
                ")";
            var parsed = parseMarker(runWorker(call));
            if (parsed.ok && !opts.directMode && parsed.extra != null) {
                var repPt = parseFloat(parsed.extra);
                if (!isNaN(repPt)) {
                    isSyncingUI = true;
                    leadingInput.text = formatByUnit(repPt, unit);
                    isSyncingUI = false;
                }
            }
        }

        // ---- イベント配線 / Wiring ----
        function onAutoAmountEdited() {
            clearLeadingSelection();
            applyToSelection();
        }
        function onLeadingPtEdited() {
            var oi = otherChoiceIndex();
            if (oi >= 0) { selectLeadingChoice(oi); }
            applyToSelection();
        }

        for (var rk = 0; rk < leadingRadios.length; rk++) {
            (function (index) {
                leadingRadios[index].onClick = function () {
                    selectLeadingChoice(index);
                    var ratio = LEADING_CHOICES[index].ratio;
                    if (typeof ratio === "number") {
                        isSyncingUI = true;
                        autoInput.text = String(Math.round(ratio * 100));
                        isSyncingUI = false;
                    }
                    applyToSelection();
                };
            })(rk);
        }

        for (var tk = 0; tk < typeRadios.length; tk++) {
            typeRadios[tk].onClick = applyToSelection;
        }

        autoInput.onChange = onAutoAmountEdited;
        leadingInput.onChange = onLeadingPtEdited;
        spaceBeforeInput.onChange = applyToSelection;
        spaceAfterInput.onChange = applyToSelection;

        changeValueByArrowKey(autoInput, false, onAutoAmountEdited);
        changeValueByArrowKey(leadingInput, false, onLeadingPtEdited, 1);
        changeValueByArrowKey(spaceBeforeInput, false, applyToSelection);
        changeValueByArrowKey(spaceAfterInput, false, applyToSelection);

        // Esc で閉じる / Close on Esc
        win.addEventListener("keydown", function (kbEvent) {
            if (kbEvent.keyName === "Escape") { win.close(); }
        });

        win.onClose = function () {
            $.global.__ALPTF_PALETTE__ = null;
            return true;
        };

        $.global.__ALPTF_PALETTE__ = win;
        win.center();
        win.show();
    }

    showPalette();

})();
