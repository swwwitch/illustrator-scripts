#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

グループ内のテキストから数値を抽出し、その数値でグループを並び替えて縦方向に整列します。
フォント情報によるグループ分けと、昇順・降順・ランダム順に対応します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SortByNumbers.md

### Overview

Extracts numbers from the text inside groups, sorts the groups by those numbers, and aligns them vertically.
Groups can be split by font, and the order can be ascending, descending or random.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SortByNumbers.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SortByNumbers";                /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-06-15";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-22";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SortByNumbers.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SortByNumbers.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* 「指定」の値が読めないときに使う間隔（pt） / Gap used when the custom value cannot be read, in points */
    var FALLBACK_GAP = 20;

    // =========================================
    // ローカライズ / Localization
    // =========================================

    var uiLang = ($.locale.indexOf("ja") === 0) ? "ja" : "en";

    var LABELS = {
        dialog: {
            title: { ja: "グループの数値で整列", en: "Align Groups by Number" }
        },
        panel: {
            sortGroup: { ja: "数値グループ", en: "Number Group" },
            spacing:   { ja: "間隔", en: "Spacing" }
        },
        radio: {
            asc:    { ja: "昇順", en: "Ascending" },
            desc:   { ja: "降順", en: "Descending" },
            random: { ja: "ランダム", en: "Random" },
            fit:    { ja: "ぴったり", en: "Fit" },
            custom: { ja: "指定", en: "Custom" }
        },
        tooltip: {
            sortGroup: { ja: "並べ替えの基準に使う数値のまとまりを選びます。", en: "Which set of numbers to sort by." },
            asc:       { ja: "数値の小さい順に並べます。", en: "Sorts from the smallest number up." },
            desc:      { ja: "数値の大きい順に並べます。", en: "Sorts from the largest number down." },
            random:    { ja: "数値と関係なく、順序をシャッフルします。", en: "Shuffles the order regardless of the numbers." },
            fit:       { ja: "現在の並びの間隔を保ったまま詰め直します。", en: "Keeps the current spacing and repacks the objects." },
            custom:    { ja: "間隔を数値で指定します。", en: "Sets the spacing to a value you type." },
            spacingInput: {
                ja: "オブジェクト間にあける間隔です。「指定」を選んだときだけ使われます。",
                en: "Gap left between objects. Used only when Custom is selected."
            }
        },
        button: {
            ok:     { ja: "ソート", en: "Sort" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        }
    };

    /**
     * ラベルを取得する（ドット区切りキー）
     * @param {string} labelPath - "panel.spacing" のようなドット区切りキー
     * @returns {string} 現在のUI言語のラベル（見つからなければキーそのもの）
     */
    function getLabel(labelPath) {
        var pathKeys = String(labelPath).split(".");
        var labelNode = LABELS;
        for (var i = 0; i < pathKeys.length; i++) {
            labelNode = labelNode[pathKeys[i]];
            if (!labelNode) return labelPath;
        }
        return (labelNode[uiLang] != null) ? labelNode[uiLang] : labelPath;
    }

    // =========================================
    // 単位 / Units
    // =========================================

    /**
     * ドキュメントの定規の単位の表示名と、1単位あたりのポイント数を返す
     * @param {Document} targetDocument - 対象ドキュメント
     * @returns {{label: string, pointsPerUnit: number}} 単位の情報（未対応の単位は pt）
     */
    function getRulerUnit(targetDocument) {
        switch (targetDocument.rulerUnits) {
            case RulerUnits.Millimeters:
                return { label: "mm", pointsPerUnit: 2.83464567 };
            case RulerUnits.Centimeters:
                return { label: "cm", pointsPerUnit: 28.3464567 };
            case RulerUnits.Inches:
                return { label: "inch", pointsPerUnit: 72 };
            case RulerUnits.Pixels:
                return { label: "px", pointsPerUnit: 1 };
            case RulerUnits.Picas:
                return { label: "pica", pointsPerUnit: 12 };
            default:
                return { label: "pt", pointsPerUnit: 1 };
        }
    }

    // =========================================
    // 数値の収集 / Number collection
    // =========================================

    /**
     * テキストフレームの数値をフォント（ファミリー＋スタイル）ごとに集める（グループは再帰的にたどる）
     * @param {PageItem} pageItem - 対象オブジェクト
     * @param {Object} fontMap - フォント名 → { value, group } の配列（ここに追加する）
     * @param {GroupItem} ownerGroup - 数値を持たせるグループ（入れ子のグループではそのグループ）
     * @param {Object} [firstNumberRef] - 最初に見つかった数値を value に入れる入れ物
     * @returns {void}
     */
    function collectNumbersByFont(pageItem, fontMap, ownerGroup, firstNumberRef) {
        if (pageItem.typename === "TextFrame") {
            var numberValue = parseFloat(pageItem.contents.replace(/,/g, ""));
            var fontName = "不明";
            try {
                /* 空のテキストでは textRanges[0] が取れないことがある / textRanges[0] may be missing on empty text */
                var textFont = pageItem.textRanges[0].characterAttributes.textFont;
                if (textFont && textFont.family && textFont.style) {
                    fontName = textFont.family + " " + textFont.style;
                }
            } catch (e) {}
            if (!isNaN(numberValue)) {
                if (!fontMap[fontName]) fontMap[fontName] = [];
                fontMap[fontName].push({ value: numberValue, group: ownerGroup });
                if (firstNumberRef && typeof firstNumberRef.value === "undefined") {
                    firstNumberRef.value = numberValue;
                }
            }
        } else if (pageItem.typename === "GroupItem") {
            for (var i = 0; i < pageItem.pageItems.length; i++) {
                var childItem = pageItem.pageItems[i];
                collectNumbersByFont(childItem, fontMap, (childItem.typename === "GroupItem" ? childItem : ownerGroup), firstNumberRef);
            }
        }
    }

    /**
     * 選択の中から、数値のテキストを含むグループだけを取り出す
     * @param {PageItem[]} selectedItems - 選択中のオブジェクト
     * @returns {GroupItem[]} 数値を含むグループ
     */
    function findNumberedGroups(selectedItems) {
        var numberedGroups = [];
        for (var i = 0; i < selectedItems.length; i++) {
            if (selectedItems[i].typename === "GroupItem") {
                var firstNumberRef = { value: undefined };
                collectNumbersByFont(selectedItems[i], {}, selectedItems[i], firstNumberRef);
                if (!isNaN(firstNumberRef.value)) {
                    numberedGroups.push(selectedItems[i]);
                }
            }
        }
        return numberedGroups;
    }

    /**
     * 同じグループの重複を除き、グループごとに最初の数値だけを残す
     * @param {Object[]} fontEntries - collectNumbersByFont() が集めた1フォント分の配列
     * @returns {Object[]} { group, value } の配列
     */
    function getUniqueGroupEntries(fontEntries) {
        var groupEntries = [];
        for (var i = 0; i < fontEntries.length; i++) {
            var alreadyExists = false;
            for (var j = 0; j < groupEntries.length; j++) {
                if (groupEntries[j].group === fontEntries[i].group) {
                    alreadyExists = true;
                    break;
                }
            }
            if (!alreadyExists) {
                groupEntries.push({ group: fontEntries[i].group, value: fontEntries[i].value });
            }
        }
        return groupEntries;
    }

    // =========================================
    // 並べ替えと配置 / Sorting and placement
    // =========================================

    /**
     * 順序をシャッフルする（Fisher–Yates）。先頭が元の先頭のままなら2番目と入れ替える
     * @param {Object[]} groupEntries - 並べ替える配列（直接書き換える）
     * @returns {void}
     */
    function shuffleGroupEntries(groupEntries) {
        var originalFirst = groupEntries[0];
        for (var i = groupEntries.length - 1; i > 0; i--) {
            var j = Math.floor(Math.random() * (i + 1));
            var swapEntry = groupEntries[i];
            groupEntries[i] = groupEntries[j];
            groupEntries[j] = swapEntry;
        }
        if (groupEntries.length > 1 && groupEntries[0].group === originalFirst.group) {
            groupEntries[0] = groupEntries[1];
            groupEntries[1] = originalFirst;
        }
    }

    /**
     * 並べ始める位置（上端がいちばん低いグループの左上）を求める
     * グループ名をキーにして控えるため、同じ名前のグループは後のものだけが候補になる
     * @param {GroupItem[]} numberedGroups - 数値を含むグループ
     * @returns {{left: number, top: number}} 並べ始める左上の座標
     */
    function getStackOrigin(numberedGroups) {
        var boundsByName = {};
        for (var i = 0; i < numberedGroups.length; i++) {
            boundsByName[numberedGroups[i].name] = numberedGroups[i].visibleBounds.concat();
        }
        var startTop = null;
        for (var groupName in boundsByName) {
            if (startTop === null || boundsByName[groupName][1] < startTop) {
                startTop = boundsByName[groupName][1];
            }
        }
        var startLeft = null;
        for (var candidateName in boundsByName) {
            if (boundsByName[candidateName][1] === startTop) {
                startLeft = boundsByName[candidateName][0];
                break;
            }
        }
        return { left: startLeft, top: startTop };
    }

    /**
     * グループを並べた順に、起点から下へ積み上げる
     * @param {Object[]} groupEntries - { group, value } の配列（並べる順）
     * @param {{left: number, top: number}} stackOrigin - 起点の左上
     * @param {boolean} fitsTightly - true なら間隔 0（ぴったり）
     * @param {number} gap - グループ間の間隔（pt）
     * @returns {void}
     */
    function stackGroups(groupEntries, stackOrigin, fitsTightly, gap) {
        var currentTop = stackOrigin.top;
        for (var i = 0; i < groupEntries.length; i++) {
            var groupItem = groupEntries[i].group;
            groupItem.locked = false;
            groupItem.hidden = false;

            var boundsBefore = groupItem.visibleBounds;
            groupItem.translate(stackOrigin.left - boundsBefore[0], currentTop - boundsBefore[1]);

            /* 移動後の高さで次の上端を決める / next top from the height after moving */
            var boundsAfter = groupItem.visibleBounds;
            var groupHeight = boundsAfter[1] - boundsAfter[3];
            if (fitsTightly) {
                currentTop -= groupHeight;
            } else {
                currentTop -= (groupHeight + gap);
            }
        }

        /* 先頭のグループが起点に来るよう全体をずらす / shift everything so the first group sits on the origin */
        var firstBounds = groupEntries[0].group.visibleBounds;
        var shiftX = stackOrigin.left - firstBounds[0];
        var shiftY = stackOrigin.top - firstBounds[1];
        for (var j = 0; j < groupEntries.length; j++) {
            groupEntries[j].group.translate(shiftX, shiftY);
        }
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 数値のまとまり（フォント）ごとのラジオボタンに出す、先頭3つの数値を返す
     * @param {Object[]} fontEntries - 1フォント分の { value, group } の配列
     * @returns {string} "1, 2, 3…" のような表示
     */
    function getNumberPreviewLabel(fontEntries) {
        var numberValues = [];
        for (var i = 0; i < fontEntries.length; i++) {
            numberValues.push(fontEntries[i].value);
        }
        numberValues.sort(function (a, b) {
            return a - b;
        });
        return (numberValues.length > 3) ? numberValues.slice(0, 3).join(", ") + "…" : numberValues.join(", ");
    }

    /**
     * 数値のまとまり・並び順・間隔を選ぶダイアログを出す
     * @param {Object} fontMap - フォント名 → 数値の配列
     * @param {{label: string, pointsPerUnit: number}} rulerUnit - 定規の単位
     * @returns {Object|null} { font, descending, random, spacingMode, spacingValue }。キャンセル時は null
     */
    function showFontChoiceDialog(fontMap, rulerUnit) {
        var sortDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        sortDialog.orientation = "column";
        sortDialog.alignChildren = "fill";

        var numberGroupPanel = sortDialog.add("panel", undefined, getLabel("panel.sortGroup"));
        numberGroupPanel.orientation = "column";
        numberGroupPanel.alignChildren = "left";
        numberGroupPanel.margins = [10, 20, 10, 10];

        var fontNames = [];
        for (var fontName in fontMap) {
            fontNames.push(fontName);
        }
        fontNames.sort();

        var fontRadios = [];
        for (var i = 0; i < fontNames.length; i++) {
            var fontRadio = numberGroupPanel.add("radiobutton", undefined, getNumberPreviewLabel(fontMap[fontNames[i]]));
            fontRadio.helpTip = getLabel("tooltip.sortGroup");
            fontRadios.push({ button: fontRadio, key: fontNames[i] });
        }
        if (fontRadios.length > 0) {
            fontRadios[0].button.value = true;
        }

        var sortOrderPanel = sortDialog.add("panel", undefined);
        sortOrderPanel.orientation = "row";
        sortOrderPanel.alignChildren = "left";
        sortOrderPanel.margins = [10, 20, 10, 10];

        var ascRadio = sortOrderPanel.add("radiobutton", undefined, getLabel("radio.asc"));
        ascRadio.helpTip = getLabel("tooltip.asc");
        var descRadio = sortOrderPanel.add("radiobutton", undefined, getLabel("radio.desc"));
        descRadio.helpTip = getLabel("tooltip.desc");
        var randomRadio = sortOrderPanel.add("radiobutton", undefined, getLabel("radio.random"));
        randomRadio.helpTip = getLabel("tooltip.random");
        ascRadio.value = true;

        var spacingPanel = sortDialog.add("panel", undefined, getLabel("panel.spacing"));
        spacingPanel.orientation = "row";
        spacingPanel.alignChildren = "left";
        spacingPanel.margins = [10, 20, 10, 10];

        var fitRadio = spacingPanel.add("radiobutton", undefined, getLabel("radio.fit"));
        fitRadio.helpTip = getLabel("tooltip.fit");
        var customRadio = spacingPanel.add("radiobutton", undefined, getLabel("radio.custom"));
        customRadio.helpTip = getLabel("tooltip.custom");
        var spacingInput = spacingPanel.add("edittext", undefined, (rulerUnit.label === "mm") ? "1" : "20");
        spacingInput.helpTip = getLabel("tooltip.spacingInput");
        spacingInput.characters = 5;
        spacingInput.enabled = false;
        spacingPanel.add("statictext", undefined, rulerUnit.label);
        fitRadio.value = true;

        customRadio.onClick = function () {
            spacingInput.enabled = true;
        };
        fitRadio.onClick = function () {
            spacingInput.enabled = false;
        };

        var btnRowGroup = sortDialog.add("group");
        btnRowGroup.alignment = "center";
        var btnCancel = btnRowGroup.add("button", undefined, getLabel("button.cancel"));
        var btnOk = btnRowGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
        btnCancel.alignment = "left";
        btnOk.alignment = "right";

        var sortOptions = null;
        btnOk.onClick = function () {
            for (var i = 0; i < fontRadios.length; i++) {
                if (fontRadios[i].button.value) {
                    sortOptions = {
                        font: fontRadios[i].key,
                        descending: descRadio.value,
                        random: randomRadio.value,
                        spacingMode: fitRadio.value ? "fit" : "custom",
                        spacingValue: parseFloat(spacingInput.text) * rulerUnit.pointsPerUnit
                    };
                    break;
                }
            }
            sortDialog.close();
        };
        btnCancel.onClick = function () {
            sortDialog.close();
        };

        sortDialog.show();
        return sortOptions;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 数値を含むグループを選んだ順に並べ替え、縦に積み上げる
     * @param {{label: string, pointsPerUnit: number}} rulerUnit - 定規の単位
     * @returns {void}
     */
    function main(rulerUnit) {
        if (app.documents.length === 0) {
            alert("ドキュメントが開かれていません。");
            return;
        }

        var docSelection = app.activeDocument.selection;
        if (docSelection.length === 0) {
            alert("グループオブジェクトを選択してください。");
            return;
        }

        var numberedGroups = findNumberedGroups(docSelection);
        if (numberedGroups.length === 0) {
            alert("数値を含むグループが見つかりません。");
            return;
        }

        var fontMap = {};
        for (var i = 0; i < numberedGroups.length; i++) {
            collectNumbersByFont(numberedGroups[i], fontMap, numberedGroups[i]);
        }
        var stackOrigin = getStackOrigin(numberedGroups);

        var sortOptions = showFontChoiceDialog(fontMap, rulerUnit);
        if (!sortOptions) return;

        var groupEntries = getUniqueGroupEntries(fontMap[sortOptions.font]);
        if (sortOptions.random) {
            shuffleGroupEntries(groupEntries);
        } else {
            groupEntries.sort(function (a, b) {
                return sortOptions.descending ? b.value - a.value : a.value - b.value;
            });
        }

        var gap = FALLBACK_GAP;
        if (sortOptions.spacingMode === "custom" && !isNaN(sortOptions.spacingValue)) {
            gap = sortOptions.spacingValue;
        }
        stackGroups(groupEntries, stackOrigin, sortOptions.spacingMode === "fit", gap);

        app.redraw();
    }

    /* 定規の単位はドキュメントの有無を確かめる前に読む（従来どおり） / read before the document check, as before */
    main(getRulerUnit(app.activeDocument));

})();
