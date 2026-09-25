#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したエリア内文字・パス上文字のあふれ（オーバーセット）を、文字サイズの縮小・拡大、
またはエリア内文字の高さの調整で解消します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AutoFitTextFrame.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n8c2e2568a6b7

### Overview

Resolves overset text in the selected area type and path type, either by shrinking or growing
the font size, or by adjusting the height of the area type.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AutoFitTextFrame.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "FitAreaText";                  /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v2.3.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-03-03";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-23";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AutoFitTextFrame.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AutoFitTextFrame.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n8c2e2568a6b7"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function() {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* 選択にエリア内文字があるとき、［エリア内文字の高さ調整］を初期状態でONにするか / Turn the height mode on by default when the selection contains area type */
    var DEFAULT_HEIGHT_MODE = true;

    /* 高さ調整のオプションの初期選択（"autoSize" = 自動サイズ調整／"adjustHeight" = 高さを調整）/ Default option of the height mode */
    var DEFAULT_HEIGHT_OPTION = "autoSize";

    /* ［文字サイズ：あふれ処理］の初期状態 / Default state of "Shrink Text to Fit" */
    var DEFAULT_SHRINK_TO_FIT = true;

    /* ［文字サイズ：ぴったり］の初期状態 / Default state of "Maximize Text Size" */
    var DEFAULT_MAXIMIZE_SIZE = true;

    /* 文字サイズを縮小する刻み（pt）/ Step used when shrinking the font size */
    var FONT_SIZE_STEP = 0.1;

    /* 縮小できる最小の文字サイズ（pt）/ Smallest font size the shrink loop may reach */
    var MIN_FONT_SIZE = 0.1;

    /* 縮小処理の上限回数（安全弁）/ Safety limit for the shrink loop */
    var MAX_SHRINK_ITERATIONS = 2000;

    /* 上限回数に達したときに警告を出すか / Alert when the shrink limit is reached */
    var ALERT_ON_SHRINK_LIMIT = true;

    /* ［ぴったり］で倍々に拡大する上限回数 / Limit of the doubling loop used by "Maximize" */
    var MAX_GROW_ITERATIONS = 25;

    /* ［ぴったり］で拡大できる上限の文字サイズ（pt）/ Largest font size the grow loop may reach */
    var MAX_FONT_SIZE = 100000;

    /* 元の文字サイズ・高さを控えるタグ名（データセットごとのリセットに使う）/ Tags holding the original values */
    var FONT_SIZE_TAG_NAME = "overset_text_default_size";
    var HEIGHT_TAG_NAME = "overset_text_default_height";

    // =========================================
    // レイアウト / Layout
    // =========================================

    var WINDOW_MARGINS        = 15;                /* ウィンドウ外周の余白 / window margin */
    var WINDOW_SPACING        = 10;                /* ウィンドウ内の要素間隔 / window spacing */
    var PANEL_MARGINS         = [15, 20, 15, 10];  /* パネル余白 [左,上,右,下] / panel margins */
    var PANEL_SPACING         = 6;                 /* パネル内の要素間隔 / panel spacing */
    var OPTION_INDENT_MARGINS = [18, 0, 0, 7];     /* 入れ子オプションの字下げ [左,上,右,下] / indent of nested options */
    var OPTION_SPACING        = 4;                 /* 入れ子オプションの間隔 / spacing of nested options */
    var BUTTON_SPACING        = 10;                /* ボタンの間隔 / spacing between buttons */
    var BUTTON_ROW_TOP_MARGIN = 8;                 /* ボタン行の上余白 / top margin of the button row */

    /**
     * ウィンドウの共通レイアウトを設定する
     * @param {Window} targetWindow - 対象のウィンドウ
     * @returns {void}
     */
    function setupWindow(targetWindow) {
        targetWindow.orientation = "column";
        targetWindow.alignChildren = ["left", "top"];
        targetWindow.margins = WINDOW_MARGINS;
        targetWindow.spacing = WINDOW_SPACING;
    }

    /**
     * パネルを追加し、共通レイアウトを設定する
     * @param {Object} parentContainer - 追加先のウィンドウまたはグループ
     * @param {Object} titleSet - ja/en を持つパネル名
     * @returns {Panel} 追加したパネル
     */
    function addPanel(parentContainer, titleSet) {
        var newPanel = parentContainer.add("panel", undefined, getLabel(titleSet));
        newPanel.orientation = "column";
        newPanel.alignChildren = ["left", "top"];
        newPanel.alignment = ["fill", "top"];
        newPanel.margins = PANEL_MARGINS;
        newPanel.spacing = PANEL_SPACING;
        return newPanel;
    }

    /**
     * 字下げした縦並びグループを追加する（入れ子のオプション用）
     * @param {Object} parentContainer - 追加先のパネルまたはグループ
     * @returns {Group} 追加したグループ
     */
    function addIndentedColumn(parentContainer) {
        var indentedGroup = parentContainer.add("group");
        indentedGroup.orientation = "column";
        indentedGroup.alignment = ["left", "top"];
        indentedGroup.alignChildren = ["left", "top"];
        indentedGroup.margins = OPTION_INDENT_MARGINS;
        indentedGroup.spacing = OPTION_SPACING;
        return indentedGroup;
    }

    // =========================================
    // ローカライズ / Localization
    // =========================================

    var uiLang = ($.locale.indexOf("ja") === 0) ? "ja" : "en";

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "テキストの自動調整", en: "Auto Fit Text Frame" }
        },
        panel: {
            processing: { ja: "調整方法", en: "Adjustment Method" }
        },
        checkbox: {
            shrinkToFit: { ja: "文字サイズ：あふれを解消", en: "Shrink Text to Fit" },
            maximizeSize: { ja: "文字サイズ：最大まで拡大", en: "Maximize Text Size" },
            heightMode: { ja: "エリア内文字の高さ調整", en: "Adjust Area Text Height" }
        },
        radio: {
            adjustHeight: { ja: "高さを広げて固定", en: "Expand and fix" },
            autoSize: { ja: "自動サイズ調整", en: "Auto size" }
        },
        tooltip: {
            shrinkToFit: {
                ja: "あふれ（オーバーセット）がなくなるまで文字サイズを縮小します。",
                en: "Shrink the font size until the text no longer oversets."
            },
            maximizeSize: {
                ja: "いったん文字サイズを拡大してから、あふれない最大サイズまで詰めます。",
                en: "Grow the font size first, then shrink it to the largest size that still fits."
            },
            heightMode: {
                ja: "文字サイズではなく、エリア内文字の高さで調整します。選択にエリア内文字があるときだけ選べます。",
                en: "Adjust the height of area type instead of the font size. Available only when the selection contains area type."
            },
            adjustHeight: {
                ja: "自動サイズ調整を一時的にONにして、必要な分だけ高さを広げてから固定します（以後は自動で変わりません）。",
                en: "Turn Auto Size on and off again, expanding the frame just enough and then fixing that height."
            },
            autoSize: {
                ja: "エリア内文字に自動サイズ調整を適用します（拡張のみ）。",
                en: "Apply Auto Size to Area Text (expand only)."
            }
        },
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" }
        },
        alert: {
            selectObject: {
                ja: "エリア内文字またはパス上文字を選択してください。",
                en: "Please select area type or path type."
            },
            noValidText: {
                ja: "選択中に処理可能なテキストがありません。\n（グループ内のテキストはグループごと選択でもOK）",
                en: "No valid text frames found in the selection.\n(Text inside groups can be processed by selecting the group.)"
            },
            noAreaText: {
                ja: "エリア内文字が選択されていません。\n高さの調整はエリア内文字のみ対応です。",
                en: "No area type found in the selection.\nHeight adjustment is supported for area type only."
            },
            selectMode: {
                ja: "調整方法を選択してください。\n（文字サイズ：あふれを解消 / 文字サイズ：最大まで拡大）",
                en: "Please select an adjustment method.\n(Shrink Text to Fit / Maximize Text Size)"
            },
            hardReturn: {
                ja: "改行コードが含まれているテキストには対応していません。\n対象テキスト：",
                en: "Text containing line breaks is not supported.\nTarget: "
            },
            shrinkLimit: {
                ja: "文字サイズの縮小が上限回数に達しました：\n",
                en: "Shrink iteration limit reached:\n"
            }
        },
        fallbackName: {
            unnamedText: { ja: "［名前なし］", en: "[Unnamed Text]" }
        }
    };

    /**
     * 現在の言語のラベルを取得する
     * @param {Object} labelSet - ja/en を持つラベル
     * @returns {string} 表示用の文字列
     */
    function getLabel(labelSet) {
        return (labelSet && labelSet[uiLang]) || "";
    }

    // =========================================
    // 定数 / Constants
    // =========================================

    /* 調整方法 / Adjustment modes */
    var ADJUST_MODE = {
        FONT_SIZE: "fontSize",   /* 文字サイズで調整 / adjust the font size */
        HEIGHT: "height",        /* エリア内文字の高さで調整 / adjust the area type height */
        AUTO_SIZE: "autoSize"    /* 自動サイズ調整を適用 / apply Auto Size */
    };

    /* 自動サイズ調整アクションの値 / Values of the Auto Size action */
    var AUTO_SIZE_ON = 1;
    var AUTO_SIZE_OFF = 2;

    // =========================================
    // テキストの収集 / Collecting text frames
    // =========================================

    /**
     * 処理対象にできるテキストフレームか判定する
     * @param {TextFrame} textFrame - 判定するテキストフレーム
     * @returns {boolean} 対象にできるとき true
     */
    function isAdjustableTextFrame(textFrame) {
        return ((textFrame.kind == TextType.PATHTEXT || textFrame.kind == TextType.AREATEXT) &&
            textFrame.editable && !textFrame.locked && !textFrame.hidden);
    }

    /**
     * 選択項目を再帰的にたどってテキストフレームを集める
     * @param {Object} selectedItem - 選択項目（TextRange／TextFrame／GroupItem など）
     * @param {TextFrame[]} collectedFrames - 集めたテキストフレームの配列（破壊的に追加）
     * @returns {void}
     */
    function collectTextFramesFromItem(selectedItem, collectedFrames) {
        if (!selectedItem) return;

        /* 選択の種類によっては parent や pageItems を読めない / Some selections cannot expose parent or pageItems */
        try {
            /* 文字カーソルでの選択はフレームに読み替える / A TextRange selection is read as its frame */
            if (selectedItem.typename === "TextRange") {
                if (selectedItem.parent && selectedItem.parent.typename === "TextFrame") {
                    collectTextFramesFromItem(selectedItem.parent, collectedFrames);
                }
                return;
            }

            if (selectedItem.typename === "TextFrame") {
                if (isAdjustableTextFrame(selectedItem)) collectedFrames.push(selectedItem);
                return;
            }

            /* グループ・複合パス・レイヤーなどは中身をたどる / Containers are traversed */
            if (selectedItem.typename === "CompoundPathItem" && selectedItem.pathItems) {
                for (var i = 0; i < selectedItem.pathItems.length; i++) {
                    collectTextFramesFromItem(selectedItem.pathItems[i], collectedFrames);
                }
                return;
            }

            if (selectedItem.pageItems) {
                for (var j = 0; j < selectedItem.pageItems.length; j++) {
                    collectTextFramesFromItem(selectedItem.pageItems[j], collectedFrames);
                }
            }
        } catch (e) { }
    }

    /**
     * 選択から処理対象のテキストフレームを重複なく集める
     * @param {Document} doc - 対象のドキュメント
     * @returns {TextFrame[]} 処理対象のテキストフレーム
     */
    function getSelectedTextFrames(doc) {
        var collectedFrames = [];
        if (!doc || !doc.selection || doc.selection.length === 0) return collectedFrames;

        for (var i = 0; i < doc.selection.length; i++) {
            collectTextFramesFromItem(doc.selection[i], collectedFrames);
        }

        /* 参照そのもので重複を除く（文字列化では区別できない）/ De-duplicate by object reference */
        var uniqueFrames = [];
        for (var j = 0; j < collectedFrames.length; j++) {
            var isDuplicate = false;
            for (var k = 0; k < uniqueFrames.length; k++) {
                if (uniqueFrames[k] === collectedFrames[j]) {
                    isDuplicate = true;
                    break;
                }
            }
            if (!isDuplicate) uniqueFrames.push(collectedFrames[j]);
        }
        return uniqueFrames;
    }

    /**
     * テキストフレームの配列からエリア内文字だけを取り出す
     * @param {TextFrame[]} textFrames - 対象のテキストフレーム
     * @returns {TextFrame[]} エリア内文字のみの配列
     */
    function filterAreaTextFrames(textFrames) {
        var areaFrames = [];
        for (var i = 0; i < textFrames.length; i++) {
            if (textFrames[i].kind == TextType.AREATEXT) areaFrames.push(textFrames[i]);
        }
        return areaFrames;
    }

    /**
     * テキストフレームの表示名を返す
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @returns {string} 名前（未設定のときは代替名）
     */
    function getTextFrameName(textFrame) {
        return textFrame.name ? textFrame.name : getLabel(LABELS.fallbackName.unnamedText);
    }

    // =========================================
    // あふれの判定 / Overset detection
    // =========================================

    /**
     * 表示されている行に収まらない文字があるか判定する
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @returns {boolean} あふれているとき true
     */
    function hasHiddenCharacters(textFrame) {
        var lineCount = textFrame.lines.length;
        if (lineCount === 0) return (textFrame.characters.length > 0);

        var visibleCharacters = 0;
        for (var i = 0; i < lineCount; i++) {
            visibleCharacters += textFrame.lines[i].characters.length;
        }
        return (visibleCharacters < textFrame.characters.length);
    }

    /**
     * テキストフレームがあふれているか判定する
     * エリア内文字は overflows を優先し、パス上文字は表示行の文字数で判定する
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @returns {boolean} あふれているとき true
     */
    function isOversetFrame(textFrame) {
        try {
            if (textFrame.kind == TextType.AREATEXT && typeof textFrame.overflows !== "undefined") {
                return !!textFrame.overflows;
            }
            return hasHiddenCharacters(textFrame);
        } catch (e) {
            return false;
        }
    }

    /**
     * 改行コードを含むテキストなら警告を出す
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @returns {boolean} 改行コードを含み、処理を中止すべきとき true
     */
    function stopIfHardReturn(textFrame) {
        var hasHardReturn;
        try {
            hasHardReturn = /[\r\n]/.test(textFrame.contents);
        } catch (e) {
            return false;
        }
        if (hasHardReturn) {
            alert(getLabel(LABELS.alert.hardReturn) + getTextFrameName(textFrame));
            return true;
        }
        return false;
    }

    // =========================================
    // 元の値の記録とリセット / Recording and resetting the original values
    // =========================================

    /**
     * タグを名前で探す（getByName は見つからないと例外になるためここで受ける）
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @param {string} tagName - タグ名
     * @returns {Tag|null} 見つかったタグ（ないときは null）
     */
    function findTag(textFrame, tagName) {
        try {
            return textFrame.tags.getByName(tagName);
        } catch (e) {
            return null;
        }
    }

    /**
     * 値をタグに書き込む（同名のタグがあれば上書き）
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @param {string} tagName - タグ名
     * @param {number} value - 控える値
     * @returns {void}
     */
    function saveValueToTag(textFrame, tagName, value) {
        var valueTag = findTag(textFrame, tagName);
        if (!valueTag) {
            valueTag = textFrame.tags.add();
            valueTag.name = tagName;
        }
        valueTag.value = value;
    }

    /**
     * タグに控えた値を読み出す
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @param {string} tagName - タグ名
     * @returns {number|null} 控えた値（ないときは null）
     */
    function readValueFromTag(textFrame, tagName) {
        var valueTag = findTag(textFrame, tagName);
        return valueTag ? (valueTag.value * 1) : null;
    }

    /**
     * タグを削除する
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @param {string} tagName - タグ名
     * @returns {void}
     */
    function removeValueTag(textFrame, tagName) {
        var valueTag = findTag(textFrame, tagName);
        if (valueTag) valueTag.remove();
    }

    /**
     * 処理中のデータセットが1件目か判定する
     * @param {Document} doc - 対象のドキュメント
     * @returns {boolean} 1件目のとき true
     */
    function isFirstDataSet(doc) {
        return (doc.dataSets.length > 0 && doc.activeDataSet == doc.dataSets[0]);
    }

    /**
     * 処理中のデータセットが最後の1件か判定する
     * @param {Document} doc - 対象のドキュメント
     * @returns {boolean} 最後の1件のとき true
     */
    function isLastDataSet(doc) {
        return (doc.dataSets.length > 0 && doc.activeDataSet == doc.dataSets[doc.dataSets.length - 1]);
    }

    /**
     * データセットの1件目なら元の値をタグに控え、控えた値があれば毎回そこへ戻す
     * @param {Document} doc - 対象のドキュメント
     * @param {TextFrame[]} textFrames - 対象のテキストフレーム
     * @param {string} tagName - タグ名
     * @param {Function} readCurrentValue - テキストフレームから現在値を読む処理
     * @param {Function} writeSavedValue - テキストフレームへ値を書き戻す処理
     * @returns {void}
     */
    function resetToOriginalValue(doc, textFrames, tagName, readCurrentValue, writeSavedValue) {
        var i;
        if (isFirstDataSet(doc)) {
            for (i = 0; i < textFrames.length; i++) {
                saveValueToTag(textFrames[i], tagName, readCurrentValue(textFrames[i]));
            }
        }
        for (i = 0; i < textFrames.length; i++) {
            if (textFrames[i].contents === "") continue;
            var savedValue = readValueFromTag(textFrames[i], tagName);
            if (savedValue !== null) writeSavedValue(textFrames[i], savedValue);
        }
    }

    /**
     * 最後のデータセットまで終わったらタグを片付ける
     * @param {Document} doc - 対象のドキュメント
     * @param {TextFrame[]} textFrames - 対象のテキストフレーム
     * @param {string} tagName - タグ名
     * @returns {void}
     */
    function removeTagsAfterLastDataSet(doc, textFrames, tagName) {
        if (!isLastDataSet(doc)) return;
        for (var i = 0; i < textFrames.length; i++) {
            removeValueTag(textFrames[i], tagName);
        }
    }

    // =========================================
    // 文字サイズの調整 / Adjusting the font size
    // =========================================

    /**
     * 文字サイズを読む
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @returns {number} 文字サイズ（pt）
     */
    function getFontSize(textFrame) {
        return textFrame.textRange.characterAttributes.size;
    }

    /**
     * 文字サイズを書き込む
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @param {number} fontSize - 文字サイズ（pt）
     * @returns {void}
     */
    function setFontSize(textFrame, fontSize) {
        textFrame.textRange.characterAttributes.size = fontSize;
    }

    /**
     * 手動行送りのときだけ、行送りと文字サイズの比率を返す
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @returns {number|null} 行送りの比率（自動行送りのときは null）
     */
    function getLeadingRatio(textFrame) {
        try {
            var textAttributes = textFrame.textRange.characterAttributes;
            if (textAttributes.autoLeading) return null;
            if (textAttributes.size > 0 && textAttributes.leading > 0) return textAttributes.leading / textAttributes.size;
        } catch (e) { }
        return null;
    }

    /**
     * 文字サイズに合わせて行送りを追従させる
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @param {number} fontSize - 変更後の文字サイズ（pt）
     * @param {number|null} leadingRatio - 行送りの比率（null のときは何もしない）
     * @returns {void}
     */
    function applyLeading(textFrame, fontSize, leadingRatio) {
        if (leadingRatio === null) return;
        try {
            textFrame.textRange.characterAttributes.leading = fontSize * leadingRatio;
        } catch (e) { }
    }

    /**
     * 文字サイズを書き込み、手動行送りなら比率を保って追従させる
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @param {number} fontSize - 文字サイズ（pt）
     * @param {number|null} leadingRatio - 行送りの比率（null のときは行送りに触らない）
     * @returns {void}
     */
    function setFontSizeWithLeading(textFrame, fontSize, leadingRatio) {
        setFontSize(textFrame, fontSize);
        applyLeading(textFrame, fontSize, leadingRatio);
    }

    /**
     * あふれがなくなるまで文字サイズを縮小する
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @returns {boolean} 続行してよいとき true（改行コードを含むときは false）
     */
    function shrinkFontToFit(textFrame) {
        if (stopIfHardReturn(textFrame)) return false;
        if (textFrame.characters.length <= 0 || !isOversetFrame(textFrame)) return true;

        var leadingRatio = getLeadingRatio(textFrame);
        var iteration = 0;
        while (isOversetFrame(textFrame)) {
            var currentSize = getFontSize(textFrame);
            if (currentSize <= MIN_FONT_SIZE) break;

            var reducedSize = Math.max(MIN_FONT_SIZE, currentSize - FONT_SIZE_STEP);
            setFontSizeWithLeading(textFrame, reducedSize, leadingRatio);

            iteration++;
            if (iteration >= MAX_SHRINK_ITERATIONS) {
                if (ALERT_ON_SHRINK_LIMIT) alert(getLabel(LABELS.alert.shrinkLimit) + getTextFrameName(textFrame));
                break;
            }
        }
        return true;
    }

    /**
     * いったんあふれるまで拡大してから、あふれない最大サイズまで詰める
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @returns {boolean} 続行してよいとき true（改行コードを含むときは false）
     */
    function maximizeFontToFit(textFrame) {
        if (stopIfHardReturn(textFrame)) return false;
        if (textFrame.characters.length <= 0) return true;

        var leadingRatio = getLeadingRatio(textFrame);
        var originalSize = getFontSize(textFrame);

        /* あふれるまで倍々に拡大する / Double the size until it oversets */
        var grownSize = originalSize;
        for (var i = 0; i < MAX_GROW_ITERATIONS && !isOversetFrame(textFrame); i++) {
            grownSize = grownSize * 2;
            if (grownSize > MAX_FONT_SIZE) break;
            setFontSizeWithLeading(textFrame, grownSize, leadingRatio);
        }

        /* それでもあふれないときは元に戻す / Restore the original size when it never oversets */
        if (!isOversetFrame(textFrame)) {
            setFontSizeWithLeading(textFrame, originalSize, leadingRatio);
            return true;
        }

        return shrinkFontToFit(textFrame);
    }

    // =========================================
    // 高さの調整（エリア内文字）/ Adjusting the height (area type)
    // =========================================

    /**
     * 自動サイズ調整をアクション経由で切り替える
     * @param {number} autoSizeValue - AUTO_SIZE_ON（ON）または AUTO_SIZE_OFF（OFF）
     * @returns {void}
     */
    function setAutoSizeByAction(autoSizeValue) {
        /* アクション定義（セット名 AreaType／アクション名 AutoSize）/ Action definition */
        var actionCode = [
            '/version 3',
            '/name [ 8 4172656154797065]',
            '/isOpen 1',
            '/actionCount 1',
            '/action-1 {',
            '  /name [ 8 4175746f53697a65 ]',
            '  /keyIndex 0',
            '  /colorIndex 0',
            '  /isOpen 1',
            '  /eventCount 1',
            '  /event-1 {',
            '    /useRulersIn1stQuadrant 0',
            '    /internalName (adobe_SLOAreaTextDialog)',
            '    /localizedName [ 33',
            '      e382a8e383aae382a2e58685e69687e5ad97e382aae38397e382b7e383a7e383b3',
            '    ]',
            '    /isOpen 1',
            '    /isOn 1',
            '    /hasDialog 0',
            '    /parameterCount 1',
            '    /parameter-1 {',
            '      /key 1952539754',
            '      /showInPalette 4294967295',
            '      /type (integer)',
            '      /value ' + String(autoSizeValue),
            '    }',
            '  }',
            '}'
        ].join("\n");

        var actionFile = new File('~/ScriptAction.aia');
        actionFile.open('w');
        actionFile.write(actionCode);
        actionFile.close();
        app.loadAction(actionFile);
        actionFile.remove();

        /* 実行に失敗しても読み込んだアクションは必ず外す / Always unload the action, even if it fails */
        try {
            app.doScript("AutoSize", "AreaType", false);
        } finally {
            app.unloadAction("AreaType", "");
        }
    }

    /**
     * エリア内文字に自動サイズ調整を適用する（拡張のみ・OFFには戻さない）
     * @param {TextFrame} textFrame - 対象のエリア内文字
     * @returns {void}
     */
    function applyAutoSize(textFrame) {
        app.activeDocument.selection = [textFrame];
        setAutoSizeByAction(AUTO_SIZE_ON);
    }

    /**
     * 自動サイズ調整をON→OFFして、必要な分だけ高さを広げて固定する
     * @param {TextFrame} textFrame - 対象のエリア内文字
     * @returns {void}
     */
    function adjustHeightToFit(textFrame) {
        if (textFrame.characters.length <= 0) return;
        applyAutoSize(textFrame);
        setAutoSizeByAction(AUTO_SIZE_OFF);
    }

    // =========================================
    // 実行 / Processing
    // =========================================

    /**
     * エリア内文字だけを取り出す（1つもなければ警告する）
     * @param {TextFrame[]} textFrames - 選択から集めたテキストフレーム
     * @returns {TextFrame[]|null} エリア内文字の配列（1つもないときは null）
     */
    function getAreaTextTargets(textFrames) {
        var areaFrames = filterAreaTextFrames(textFrames);
        if (areaFrames.length === 0) {
            alert(getLabel(LABELS.alert.noAreaText));
            return null;
        }
        return areaFrames;
    }

    /**
     * テキストフレームを順に処理する（中止が返ったらそこで止める）
     * @param {TextFrame[]} textFrames - 対象のテキストフレーム
     * @param {Function} adjustFrame - 1つのテキストフレームを処理する関数
     * @returns {boolean} 最後まで処理できたとき true
     */
    function adjustEachFrame(textFrames, adjustFrame) {
        for (var i = 0; i < textFrames.length; i++) {
            if (adjustFrame(textFrames[i]) === false) return false;
        }
        return true;
    }

    /**
     * 自動サイズ調整を適用する（エリア内文字のみ）
     * @param {TextFrame[]} textFrames - 選択から集めたテキストフレーム
     * @returns {void}
     */
    function runAutoSize(textFrames) {
        var areaFrames = getAreaTextTargets(textFrames);
        if (!areaFrames) return;
        adjustEachFrame(areaFrames, applyAutoSize);
    }

    /**
     * エリア内文字の高さを中身に合わせる
     * @param {Document} doc - 対象のドキュメント
     * @param {TextFrame[]} textFrames - 選択から集めたテキストフレーム
     * @returns {void}
     */
    function runHeightAdjust(doc, textFrames) {
        var areaFrames = getAreaTextTargets(textFrames);
        if (!areaFrames) return;

        resetToOriginalValue(doc, areaFrames, HEIGHT_TAG_NAME,
            function(textFrame) { return textFrame.height; },
            function(textFrame, height) { textFrame.height = height; });

        adjustEachFrame(areaFrames, adjustHeightToFit);
        removeTagsAfterLastDataSet(doc, areaFrames, HEIGHT_TAG_NAME);
    }

    /**
     * 文字サイズであふれを調整する（両方ONなら「最大まで拡大」→「あふれを解消」の順）
     * @param {Document} doc - 対象のドキュメント
     * @param {TextFrame[]} textFrames - 選択から集めたテキストフレーム
     * @param {boolean} doMaximize - ［文字サイズ：最大まで拡大］を実行するか
     * @param {boolean} doShrink - ［文字サイズ：あふれを解消］を実行するか
     * @returns {void}
     */
    function runFontSizeAdjust(doc, textFrames, doMaximize, doShrink) {
        if (textFrames.length === 0) {
            alert(getLabel(LABELS.alert.noValidText));
            return;
        }

        resetToOriginalValue(doc, textFrames, FONT_SIZE_TAG_NAME, getFontSize, setFontSize);

        if (doMaximize && !adjustEachFrame(textFrames, maximizeFontToFit)) return;
        if (doShrink && !adjustEachFrame(textFrames, shrinkFontToFit)) return;

        removeTagsAfterLastDataSet(doc, textFrames, FONT_SIZE_TAG_NAME);
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * チェックボックスかラジオボタンを、LABELS のキーで tooltip 付きで追加する
     * @param {Panel|Group} parentContainer - 追加先のパネルかグループ
     * @param {string} controlType - "checkbox" / "radiobutton"
     * @param {string} labelKey - LABELS.checkbox（または LABELS.radio）と LABELS.tooltip のキー
     * @returns {Checkbox|RadioButton} 追加したコントロール
     */
    function addLabeledControl(parentContainer, controlType, labelKey) {
        var labelCategory = (controlType === "radiobutton") ? LABELS.radio : LABELS.checkbox;
        var addedControl = parentContainer.add(controlType, undefined, getLabel(labelCategory[labelKey]));
        addedControl.helpTip = getLabel(LABELS.tooltip[labelKey]);
        return addedControl;
    }

    /**
     * ［調整方法］パネルを組み立て、中のコントロールを返す
     * @param {Window} parentWindow - 追加先のダイアログ
     * @param {boolean} hasAreaText - 選択にエリア内文字があるか
     * @returns {Object} パネル内のコントロール
     */
    function addProcessingPanel(parentWindow, hasAreaText) {
        var processingPanel = addPanel(parentWindow, LABELS.panel.processing);

        var shrinkCheckbox = addLabeledControl(processingPanel, "checkbox", "shrinkToFit");
        shrinkCheckbox.value = DEFAULT_SHRINK_TO_FIT;

        var maximizeCheckbox = addLabeledControl(processingPanel, "checkbox", "maximizeSize");
        maximizeCheckbox.value = DEFAULT_MAXIMIZE_SIZE;

        var heightModeCheckbox = addLabeledControl(processingPanel, "checkbox", "heightMode");
        heightModeCheckbox.enabled = hasAreaText;
        heightModeCheckbox.value = (hasAreaText && DEFAULT_HEIGHT_MODE);

        var heightOptionGroup = addIndentedColumn(processingPanel);
        var adjustHeightRadio = addLabeledControl(heightOptionGroup, "radiobutton", "adjustHeight");
        var autoSizeRadio = addLabeledControl(heightOptionGroup, "radiobutton", "autoSize");

        /* 高さ調整をOFFに戻したとき用に、文字サイズの選択を控える / Remember the font-size choices */
        var previousShrinkState = shrinkCheckbox.value;
        var previousMaximizeState = maximizeCheckbox.value;

        /**
         * 高さ調整のON/OFFに合わせて、各項目の有効・無効と選択状態を切り替える
         * @returns {void}
         */
        function updateHeightOptionState() {
            var isHeightMode = (heightModeCheckbox.value && hasAreaText);
            heightOptionGroup.enabled = isHeightMode;

            if (isHeightMode) {
                /* 高さ調整中は文字サイズの処理を止める / The font-size options are off while adjusting the height */
                previousShrinkState = shrinkCheckbox.value;
                previousMaximizeState = maximizeCheckbox.value;
                shrinkCheckbox.value = false;
                maximizeCheckbox.value = false;
                shrinkCheckbox.enabled = false;
                maximizeCheckbox.enabled = false;
                autoSizeRadio.value = (DEFAULT_HEIGHT_OPTION === "autoSize");
                adjustHeightRadio.value = !autoSizeRadio.value;
            } else {
                shrinkCheckbox.enabled = true;
                maximizeCheckbox.enabled = true;
                shrinkCheckbox.value = previousShrinkState;
                maximizeCheckbox.value = previousMaximizeState;
                adjustHeightRadio.value = false;
                autoSizeRadio.value = false;
            }
        }

        heightModeCheckbox.onClick = updateHeightOptionState;
        updateHeightOptionState();

        return {
            shrinkCheckbox: shrinkCheckbox,
            maximizeCheckbox: maximizeCheckbox,
            heightModeCheckbox: heightModeCheckbox,
            autoSizeRadio: autoSizeRadio
        };
    }

    /**
     * ボタンエリア（左右中央）を組み立て、ボタンを返す
     * @param {Window} parentWindow - 追加先のダイアログ
     * @returns {Object} キャンセルボタンとOKボタン
     */
    function addButtonRow(parentWindow) {
        var btnRowGroup = parentWindow.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.margins = [0, BUTTON_ROW_TOP_MARGIN, 0, 0];
        btnRowGroup.alignment = ["center", "bottom"];
        btnRowGroup.alignChildren = ["center", "center"];
        btnRowGroup.spacing = BUTTON_SPACING;

        return {
            btnCancel: btnRowGroup.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" }),
            btnOK: btnRowGroup.add("button", undefined, getLabel(LABELS.button.ok), { name: "ok" })
        };
    }

    /**
     * パネルの選択内容を設定にまとめる（選択が足りないときは警告する）
     * @param {Object} processingControls - ［調整方法］パネルのコントロール
     * @returns {Object|null} 実行する処理の設定（選択が足りないときは null）
     */
    function readAdjustSettings(processingControls) {
        if (processingControls.heightModeCheckbox.value) {
            return { mode: processingControls.autoSizeRadio.value ? ADJUST_MODE.AUTO_SIZE : ADJUST_MODE.HEIGHT };
        }

        if (!processingControls.shrinkCheckbox.value && !processingControls.maximizeCheckbox.value) {
            alert(getLabel(LABELS.alert.selectMode));
            return null;
        }

        return {
            mode: ADJUST_MODE.FONT_SIZE,
            doMaximize: processingControls.maximizeCheckbox.value,
            doShrink: processingControls.shrinkCheckbox.value
        };
    }

    /**
     * ダイアログを表示し、選ばれた処理を返す
     * @param {boolean} hasAreaText - 選択にエリア内文字があるか
     * @returns {Object|null} 実行する処理の設定（キャンセル時は null）
     */
    function showDialog(hasAreaText) {
        var adjustDialog = new Window("dialog", getLabel(LABELS.dialog.title) + " " + SCRIPT_VERSION);
        setupWindow(adjustDialog);

        var processingControls = addProcessingPanel(adjustDialog, hasAreaText);
        var dialogButtons = addButtonRow(adjustDialog);
        var selectedSettings = null;

        dialogButtons.btnOK.onClick = function() {
            selectedSettings = readAdjustSettings(processingControls);
            if (selectedSettings) adjustDialog.close(1);
        };

        dialogButtons.btnCancel.onClick = function() {
            adjustDialog.close(0);
        };

        adjustDialog.show();
        return selectedSettings;
    }

    // =========================================
    // メイン / Main
    // =========================================

    /**
     * 選択からテキストを集め、ダイアログで選ばれた処理を実行する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) return;

        var doc = app.activeDocument;
        if (!doc.selection || doc.selection.length === 0) {
            alert(getLabel(LABELS.alert.selectObject));
            return;
        }

        /* 処理中に選択が変わるため、ダイアログの前に対象を確定させる / Collect the targets before the selection changes */
        var textFrames = getSelectedTextFrames(doc);
        var hasAreaText = (filterAreaTextFrames(textFrames).length > 0);

        var adjustSettings = showDialog(hasAreaText);
        if (!adjustSettings) return;

        if (adjustSettings.mode === ADJUST_MODE.AUTO_SIZE) {
            runAutoSize(textFrames);
        } else if (adjustSettings.mode === ADJUST_MODE.HEIGHT) {
            runHeightAdjust(doc, textFrames);
        } else {
            runFontSizeAdjust(doc, textFrames, adjustSettings.doMaximize, adjustSettings.doShrink);
        }
    }

    main();

})();
