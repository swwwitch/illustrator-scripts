#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

(function () {

    /*
     * 概要 / Overview
     * 選択された各テキストフレーム内の各行ごとに、行頭数文字のフォントサイズを基準に行送りを再計算して適用します。
     * テキストの一部（TextRange）が選択されている場合、その親のテキストフレーム（TextFrame）に正規化します。
     * 自動行送りを有効化し、行送りの値そのものではなく自動行送りの値（％）を変更して行送りを調整します。
     * 非テキストオブジェクトや処理不能な要素はスキップされます。
     * エラーは最初の1回のみ通知され、処理は継続されます。
     * アラートメッセージはローカライズ定義（LABELS）を通して取得します。
     * 更新日：2026-07-08
     */

    var DEFAULT_LEADING_RATIO = 1.1; // fallback

    var fileName = File($.fileName).name;
    var match = fileName.match(/(\d+)\.jsx$/);

    if (match) {
        var num = Number(match[1]);
        if (!isNaN(num) && num > 0) {
            DEFAULT_LEADING_RATIO = num / 100;
        }
    }
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
        alertNoProcessableLine: {
            ja: "処理可能な行が見つかりませんでした。",
            en: "No processable lines were found."
        },
        alertNoProcessableTextFrame: {
            ja: "処理可能なテキストフレームが選択されていません。",
            en: "No processable text frames are selected."
        }
    };

    function L(key) {
        return LABELS[key][lang];
    }

    applyLeadingPerTextFrame(DEFAULT_LEADING_RATIO);

    function ensureTextFrameSelection() {
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

    function getProcessableTextFrameInfo(item) {
        if (!item || item.typename !== "TextFrame") {
            return null;
        }

        if (!item.contents) {
            return null;
        }

        if (!item.lines || item.lines.length === 0) {
            return null;
        }

        return {
            textFrame: item
        };
    }

    function applyLeadingPerTextFrame(leadingRatio) {
        ensureTextFrameSelection();
        var doc;
        try {
            doc = app.activeDocument;
        } catch (e) {
            alert(L("alertOpenDocument"));
            return;
        }

        var selectedItems = doc.selection;

        if (!selectedItems || selectedItems.length === 0) {
            alert(L("alertSelectTextObject"));
            return;
        }

        var hasValidTextFrame = false;
        var hasProcessableLine = false;
        var hasShownLeadingError = false;

        for (var i = 0; i < selectedItems.length; i++) {
            var textFrameInfo = getProcessableTextFrameInfo(selectedItems[i]);
            if (!textFrameInfo) {
                continue;
            }

            var textFrame = textFrameInfo.textFrame;
            hasValidTextFrame = true;

            try {
                // 各テキストフレーム内の各行ごとに、自動行送りをONにして自動行送りの値（％）を変更する（行送り値そのものやフォントなど他の属性は触らない）
                var autoLeadingPercent = leadingRatio * 100;
                var lines = textFrame.lines;
                for (var j = 0; j < lines.length; j++) {
                    var line = lines[j];
                    var baseFontSizeFromLine = getBaseFontSizeFromLine(line);
                    if (isNaN(baseFontSizeFromLine)) {
                        continue;
                    }

                    hasProcessableLine = true;
                    line.characterAttributes.autoLeading = true;
                    line.paragraphAttributes.autoLeadingAmount = autoLeadingPercent;
                }

            } catch (e) {
                if (!hasShownLeadingError) {
                    alert(L("alertLeadingApplyError") + "\r" + e.message);
                    hasShownLeadingError = true;
                }
            }
        }

        if (!hasValidTextFrame) {
            alert(L("alertNoProcessableTextFrame"));
            return;
        }

        if (!hasProcessableLine) {
            alert(L("alertNoProcessableLine"));
            return;
        }

        redraw();
    }
})();