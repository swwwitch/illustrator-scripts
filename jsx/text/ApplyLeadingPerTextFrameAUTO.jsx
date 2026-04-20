#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

(function () {

    /*
     * 概要 / Overview
     * 選択された各テキストフレーム内の各行ごとに行送りを設定します。
     * テキストの一部（TextRange）が選択されている場合、その親のテキストフレーム（TextFrame）に正規化します。
     * ファイル名に AUTO または 自動 を含む場合は、自動行送り（autoLeading）を有効にします（サイズ調査は行いません）。
     * それ以外の場合は、各行の先頭最大5文字のフォントサイズをサンプリングし、最頻値（同数時は最大値）を基準に手動行送りを設定します。
     * ファイル名末尾の数値（例: 110）は倍率として解釈され、1.10 のように適用されます。数値が無い場合は 1.2 を使用します。
     * 非テキストオブジェクトや処理不能な要素はスキップされます。
     * エラーは最初の1回のみ通知され、処理は継続されます。
     * アラートメッセージはローカライズ定義（LABELS）を通して取得します。
     * 更新日：2026-04-21
     */

    var DEFAULT_LEADING_RATIO = 1.2; // fallback

    var fileName = File($.fileName).name;
    var SHOULD_USE_AUTO_LEADING = /AUTO|自動/i.test(fileName);
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
        alertAutoLeadingApplyError: {
            ja: "自動行送りの設定中にエラーが発生しました。",
            en: "An error occurred while applying auto leading."
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
            }
        }
    }

    function getBaseFontSizeFromLine(line) {
        if (!line || !line.characters || line.characters.length === 0) return NaN;

        var counts = {};
        var bestValue = NaN;
        var bestCount = 0;

        var maxCount = Math.min(line.characters.length, LINE_FONT_SIZE_SAMPLE_COUNT);

        for (var i = 0; i < maxCount; i++) {
            try {
                var size = line.characters[i].characterAttributes.size;
                if (isNaN(size)) continue;

                var key = String(size);
                if (!counts[key]) counts[key] = { value: size, count: 0 };
                counts[key].count++;

                if (counts[key].count > bestCount || (counts[key].count === bestCount && size > bestValue)) {
                    bestCount = counts[key].count;
                    bestValue = size;
                }
            } catch (e) { }
        }

        return bestValue;
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
            var item = selectedItems[i];
            if (!item || item.typename !== "TextFrame" || !item.contents || !item.lines || item.lines.length === 0) {
                continue;
            }
            var textFrame = item;
            hasValidTextFrame = true;

            try {
                // 各テキストフレーム内の各行ごとに行送りのみ変更する（フォントなど他の属性は触らない）
                var lines = textFrame.lines;

                if (SHOULD_USE_AUTO_LEADING) {
                    for (var j = 0; j < lines.length; j++) {
                        hasProcessableLine = true;
                        lines[j].characterAttributes.autoLeading = true;
                    }
                } else {
                    for (var j = 0; j < lines.length; j++) {
                        var line = lines[j];
                        var baseFontSizeFromLine = getBaseFontSizeFromLine(line);
                        if (isNaN(baseFontSizeFromLine)) {
                            continue;
                        }

                        hasProcessableLine = true;
                        var newLeading = baseFontSizeFromLine * leadingRatio;
                        line.characterAttributes.autoLeading = false;
                        line.characterAttributes.leading = newLeading;
                    }
                }

            } catch (e) {
                if (!hasShownLeadingError) {
                    alert(L(SHOULD_USE_AUTO_LEADING ? "alertAutoLeadingApplyError" : "alertLeadingApplyError") + " : " + e.message);
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