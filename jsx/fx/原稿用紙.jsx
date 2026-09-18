#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択した縦組みテキストの文字送りをそろえ、1文字分の間隔で横罫線を引き、全体の左右に縦罫線を添えます。
縦罫の伸張、横罫の線種、太くする間隔、十字線の有無はダイアログで指定します。

詳細は README を参照してください。

### Overview

Normalizes the character advance of the selected vertical text, draws horizontal rules
at one-character intervals, and adds vertical rules along both sides of the grid.
A dialog sets the vertical rule extension, the horizontal rule style, the emphasis interval
and whether to add crosshairs.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "原稿用紙";                      /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-09-19";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "20260919";                             /* 更新日 / last updated */

var SCRIPT_README_JA   = ""; /* README（日本語） */
var SCRIPT_README_EN   = ""; /* README (English) */
var SCRIPT_ARTICLE_URL = ""; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // 言語とラベル
    // Language and labels
    // =========================================

    var uiLang = ($.locale && $.locale.indexOf("ja") === 0) ? "ja" : "en";

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noTextFrame: { ja: "テキストオブジェクトを1つ選択してください。", en: "Please select a single text object." },
            noCharacter: { ja: "テキストに文字が入力されていません。", en: "The selected text is empty." },
            noFontSize: { ja: "文字サイズを取得できませんでした。", en: "Could not read the font size." },
            skipped: { ja: "設定できなかった属性:", en: "Attributes that could not be set:" }
        },
        item: {
            layerName: { ja: "原稿用紙罫線", en: "Manuscript Grid" },
            groupName: { ja: "罫線", en: "Rules" }
        },
        dialog: {
            title: { ja: "原稿用紙", en: "Manuscript Grid" }
        },
        label: {
            extension: { ja: "縦罫の上下の伸張", en: "Vertical rule extension" },
            horizontalRule: { ja: "横罫", en: "Horizontal rules" },
            emphasisPrefix: { ja: "", en: "Thicker every" },
            emphasisSuffix: { ja: "文字ごとに太く", en: "characters" }
        },
        checkbox: {
            cross: { ja: "各文字に十字線を追加", en: "Add a crosshair to every cell" }
        },
        radio: {
            solid: { ja: "実線", en: "Solid" },
            dashed: { ja: "破線", en: "Dashed" }
        },
        unit: {
            mm: { ja: "mm", en: "mm" }
        },
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" }
        }
    };

    /**
     * ラベルを現在の言語で取り出す
     * @param {object} entry ja / en を持つラベル定義
     * @return {string} 表示用の文字列
     */
    function getLabel(entry) {
        return (typeof entry[uiLang] === "string") ? entry[uiLang] : entry.en;
    }

    /**
     * 項目名にコロンを付ける（日本語は全角、英語は半角）
     * @param {object} entry ja / en を持つラベル定義
     * @return {string} コロン付きの項目名
     */
    function labelText(entry) {
        return getLabel(entry) + ((uiLang === "ja") ? "：" : ":");
    }

    /**
     * ミリメートルをポイントに変換する
     * @param {number} mm ミリメートル
     * @return {number} ポイント
     */
    function mmToPt(mm) {
        return mm * 72 / 25.4;
    }

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    var CONFIG = {
        horizontalScale: 90,    /* 水平比率（%）/ horizontal scale (%) */
        verticalScale: 90,      /* 垂直比率（%）/ vertical scale (%) */
        forceVertical: true     /* 横組みなら縦組みへ変換する / convert horizontal text to vertical */
    };

    // =========================================
    // レイアウト / Layout
    // =========================================

    var LAYOUT = {
        verticalStrokeMM: 0.25,     /* 縦罫の太さ（mm）/ vertical rule weight (mm) */
        horizontalStrokeMM: 0.1,    /* 横罫の太さ（mm）/ horizontal rule weight (mm) */
        horizontalDashMM: [1, 1],   /* 横罫を破線にしたときの線分と間隔（mm）/ dash and gap of dashed horizontal rules (mm) */
        emphasisStrokeMM: 0.25,     /* 太くする横罫の太さ（mm）/ weight of the emphasized horizontal rules (mm) */
        crossStrokeMM: 0.1,         /* 十字線の太さ（mm）/ crosshair weight (mm) */
        crossDashMM: [1, 1],        /* 十字線の線分と間隔（mm）/ dash and gap of the crosshairs (mm) */
        crossGray: 30,              /* 十字線の濃度（%、100で黒）/ crosshair density (%, 100 is black) */
        strokeGray: 50,             /* 罫線の濃度（%、100で黒）/ rule density (%, 100 is black) */
        gridTopOffset: 0            /* 1本目の位置の微調整（pt、マイナスで下へ）/ nudge of the first rule (pt) */
    };

    /* ダイアログの初期値 / Dialog defaults */
    var DEFAULTS = {
        extensionMM: 2,          /* 縦罫を上下へ伸ばす量（mm）/ how far the vertical rules run past the grid (mm) */
        dashedHorizontal: false, /* 横罫を破線にするか / draw the horizontal rules dashed */
        emphasisEnabled: true,   /* 一定間隔の横罫を太くするか / thicken the horizontal rules at a fixed interval */
        emphasisEvery: 5,        /* 太くする間隔（文字数）/ interval of the thickened rules (characters) */
        showCross: true          /* 各マスに十字線を引くか / draw a crosshair in every cell */
    };

    /* ダイアログの寸法 / Dialog metrics */
    var UI = {
        labelWidth: (uiLang === "ja") ? 120 : 150 /* 項目名の幅（px）/ width of the row labels (px) */
    };

    // =========================================
    // 文字属性 / Character attributes
    // =========================================

    /**
     * テキスト全体に文字属性を1つ設定する
     * 取得済みの characterAttributes を使い回すと Error 9503 や片側無視が起きるため、毎回取り直す
     * @param {TextFrame} textFrame 対象のテキストフレーム
     * @param {string} attributeName 属性名
     * @param {number|boolean|object} value 設定する値
     * @return {boolean} 設定できたかどうか
     */
    function setCharacterAttribute(textFrame, attributeName, value) {
        try {
            textFrame.textRange.characterAttributes[attributeName] = value;
            return true;
        } catch (attributeError) {
            return false;
        }
    }

    /**
     * マス目に合わせて文字属性をそろえる
     * @param {TextFrame} textFrame 対象のテキストフレーム
     * @return {array} 設定できなかった属性名の配列
     */
    function normalizeCharacterAttributes(textFrame) {
        var failed = [];

        var settings = [
            ["tracking", 0],
            ["Tsume", 0],
            ["proportionalMetrics", false],
            ["akiLeft", 0],
            ["akiRight", 0],
            ["horizontalScale", CONFIG.horizontalScale],
            ["verticalScale", CONFIG.verticalScale]
        ];

        for (var i = 0; i < settings.length; i++) {
            if (!setCharacterAttribute(textFrame, settings[i][0], settings[i][1])) {
                failed.push(settings[i][0]);
            }
        }

        /* 手動カーニングはDOMから設定できないため、自動カーニングを「和文等幅」にする */
        if (!setCharacterAttribute(textFrame, "kerningMethod", AutoKernType.NOAUTOKERN)) {
            failed.push("kerningMethod");
        }

        return failed;
    }

    // =========================================
    // 罫線 / Rules
    // =========================================

    /**
     * 罫線用のレイヤーを用意する（同名があれば再利用）
     * @param {Document} doc 対象のドキュメント
     * @param {string} layerName レイヤー名
     * @return {Layer} 罫線用のレイヤー
     */
    function prepareRuleLayer(doc, layerName) {
        for (var i = 0; i < doc.layers.length; i++) {
            if (doc.layers[i].name === layerName) {
                var found = doc.layers[i];
                found.locked = false;
                found.visible = true;
                return found;
            }
        }

        var created = doc.layers.add();
        created.name = layerName;
        return created;
    }

    /**
     * 前回引いた罫線グループが残っていれば削除する
     * @param {Layer} layer 対象のレイヤー
     * @param {string} groupName グループ名
     * @return {void}
     */
    function removeExistingGroup(layer, groupName) {
        for (var i = layer.groupItems.length - 1; i >= 0; i--) {
            if (layer.groupItems[i].name === groupName) {
                layer.groupItems[i].remove();
            }
        }
    }

    /**
     * ドキュメントのカラーモードに合わせた罫線の色をつくる
     * @param {Document} doc 対象のドキュメント
     * @param {number} grayPercent 濃度（%、100で黒）
     * @return {object} GrayColor または RGBColor
     */
    function createRuleColor(doc, grayPercent) {
        if (doc.documentColorSpace === DocumentColorSpace.RGB) {
            var level = Math.round(255 * (100 - grayPercent) / 100);
            var rgbColor = new RGBColor();
            rgbColor.red = level;
            rgbColor.green = level;
            rgbColor.blue = level;
            return rgbColor;
        }

        var grayColor = new GrayColor();
        grayColor.gray = grayPercent;
        return grayColor;
    }

    /**
     * 罫線を1本引く
     * @param {GroupItem} group 罫線を入れるグループ
     * @param {array} from 始点 [x, y]
     * @param {array} to 終点 [x, y]
     * @param {number} strokeWidth 線幅（pt）
     * @param {array} dashes 破線パターン（実線は空配列）
     * @param {object} color 罫線の色
     * @return {void}
     */
    function drawRule(group, from, to, strokeWidth, dashes, color) {
        var line = group.pathItems.add();
        line.setEntirePath([from, to]);

        line.filled = false;
        line.stroked = true;
        line.strokeWidth = strokeWidth;
        line.strokeDashes = dashes;
        line.strokeColor = color;
    }

    /**
     * 縦組み1列あたりの最大文字数を返す
     * @param {TextFrame} textFrame 対象のテキストフレーム
     * @return {number} 最大文字数（取得できないときは0）
     */
    function getMaxCharactersPerLine(textFrame) {
        var maxCount = 0;

        try {
            var lines = textFrame.lines;

            for (var i = 0; i < lines.length; i++) {
                var count = lines[i].characters.length;

                if (count > maxCount) {
                    maxCount = count;
                }
            }
        } catch (lineError) {
            return 0;
        }

        return maxCount;
    }

    /**
     * 各マスの中央に十字線を引く
     * @param {GroupItem} group 罫線を入れるグループ
     * @param {array} bounds テキストの境界 [left, top, right, bottom]
     * @param {number} cellHeight 1文字分の送り（pt）
     * @param {number} cellCount マスの数
     * @param {object} color 十字線の色
     * @return {void}
     */
    function drawCrosshairs(group, bounds, cellHeight, cellCount, color) {
        var left = bounds[0];
        var top = bounds[1] + LAYOUT.gridTopOffset;
        var right = bounds[2];

        var strokeWidth = mmToPt(LAYOUT.crossStrokeMM);
        var dashes = [mmToPt(LAYOUT.crossDashMM[0]), mmToPt(LAYOUT.crossDashMM[1])];
        var centerX = (left + right) / 2;

        for (var i = 0; i < cellCount; i++) {
            var cellTop = top - cellHeight * i;
            var cellBottom = cellTop - cellHeight;
            var centerY = cellTop - cellHeight / 2;

            drawRule(group, [centerX, cellTop], [centerX, cellBottom], strokeWidth, dashes, color);
            drawRule(group, [left, centerY], [right, centerY], strokeWidth, dashes, color);
        }
    }

    /**
     * 1文字分の間隔で横罫線を引き、全体の左右に縦罫線を添える
     * @param {GroupItem} group 罫線を入れるグループ
     * @param {array} bounds テキストの境界 [left, top, right, bottom]
     * @param {number} cellHeight 1文字分の送り（pt）
     * @param {number} cellCount マスの数
     * @param {object} settings ダイアログで決めた設定
     * @param {object} colors 罫線と十字線の色 { rule, cross }
     * @return {void}
     */
    function drawGrid(group, bounds, cellHeight, cellCount, settings, colors) {
        var left = bounds[0];
        var top = bounds[1] + LAYOUT.gridTopOffset;
        var right = bounds[2];

        /* 誤差を溜めないよう、都度 top から引いた位置を使う */
        var gridBottom = top - cellHeight * cellCount;

        var horizontalWidth = mmToPt(LAYOUT.horizontalStrokeMM);
        var emphasisWidth = mmToPt(LAYOUT.emphasisStrokeMM);
        var verticalWidth = mmToPt(LAYOUT.verticalStrokeMM);

        var horizontalDashes = settings.dashedHorizontal ?
            [mmToPt(LAYOUT.horizontalDashMM[0]), mmToPt(LAYOUT.horizontalDashMM[1])] : [];

        /* 十字線は罫線より先に引いて背面へ回す */
        if (settings.showCross) {
            drawCrosshairs(group, bounds, cellHeight, cellCount, colors.cross);
        }

        /* 横罫線：先頭文字の上から最後の文字の下まで引く */
        for (var i = 0; i <= cellCount; i++) {
            var y = top - cellHeight * i;
            var isEmphasized = (settings.emphasisEvery > 0) && (i % settings.emphasisEvery === 0);

            drawRule(group, [left, y], [right, y], isEmphasized ? emphasisWidth : horizontalWidth, horizontalDashes, colors.rule);
        }

        /* 縦罫線：列の区切りは引かず、全体の左右のみ。上下へ伸ばす */
        var verticalTop = top + settings.extension;
        var verticalBottom = gridBottom - settings.extension;

        drawRule(group, [left, verticalTop], [left, verticalBottom], verticalWidth, [], colors.rule);
        drawRule(group, [right, verticalTop], [right, verticalBottom], verticalWidth, [], colors.rule);
    }

    /**
     * レイヤーごと最背面へ送る
     * @param {Layer} layer 対象のレイヤー
     * @return {void}
     */
    function sendLayerToBack(layer) {
        try {
            layer.zOrder(ZOrderMethod.SENDTOBACK);
        } catch (zOrderError) {
            layer.move(layer.parent, ElementPlacement.PLACEATEND);
        }
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 数値入力欄を↑↓キーで増減できるようにする（Shiftで10の倍数にスナップ）
     * @param {object} editText 対象の edittext
     * @param {number} minimum 下限値
     * @return {void}
     */
    function changeValueByArrowKey(editText, minimum) {
        editText.addEventListener("keydown", function (event) {
            var fieldValue = Number(editText.text);
            if (isNaN(fieldValue)) return;

            var keyboard = ScriptUI.environment.keyboardState;

            if (event.keyName == "Up" || event.keyName == "Down") {
                var isUp = event.keyName == "Up";
                var delta = 1;

                if (keyboard.shiftKey) {
                    fieldValue = Math.floor(fieldValue / 10) * 10;
                    delta = 10;
                }

                fieldValue += isUp ? delta : -delta;
                if (fieldValue < minimum) fieldValue = minimum;

                event.preventDefault();
                editText.text = fieldValue;
            }
        });
    }

    /**
     * 項目名のラベルを1つ追加する
     * @param {object} parent 追加先のグループ
     * @param {object} entry ja / en を持つラベル定義
     * @return {object} 追加した statictext
     */
    function addRowLabel(parent, entry) {
        var label = parent.add("statictext", undefined, labelText(entry));
        label.justify = "right";
        label.preferredSize.width = UI.labelWidth;
        return label;
    }

    /**
     * ダイアログの1行分のグループを追加する
     * @param {object} parent 追加先のウィンドウ
     * @return {object} 追加したグループ
     */
    function addRow(parent) {
        var row = parent.add("group");
        row.orientation = "row";
        row.alignment = ["fill", "center"];
        row.alignChildren = ["left", "center"];
        return row;
    }

    /**
     * 設定ダイアログを表示する
     * @return {object|null} 設定値。キャンセルしたときは null
     */
    function showSettingsDialog() {
        var dialog = new Window("dialog", getLabel(LABELS.dialog.title));
        dialog.orientation = "column";
        dialog.alignChildren = ["fill", "top"];
        dialog.margins = 15;
        dialog.spacing = 10;

        /* 縦罫の上下の伸張 */
        var extensionRow = addRow(dialog);
        addRowLabel(extensionRow, LABELS.label.extension);

        var extensionInput = extensionRow.add("edittext", undefined, String(DEFAULTS.extensionMM));
        extensionInput.characters = 5;
        changeValueByArrowKey(extensionInput, 0);

        extensionRow.add("statictext", undefined, getLabel(LABELS.unit.mm));

        /* 横罫の線種 */
        var styleRow = addRow(dialog);
        addRowLabel(styleRow, LABELS.label.horizontalRule);

        /* ラジオは同じ親の中でしか排他にならないため、2つを同じグループへ入れる */
        var styleGroup = styleRow.add("group");
        styleGroup.orientation = "row";
        styleGroup.alignChildren = ["left", "center"];

        var solidRadio = styleGroup.add("radiobutton", undefined, getLabel(LABELS.radio.solid));
        var dashedRadio = styleGroup.add("radiobutton", undefined, getLabel(LABELS.radio.dashed));
        solidRadio.value = !DEFAULTS.dashedHorizontal;
        dashedRadio.value = DEFAULTS.dashedHorizontal;

        /* n文字ごとに太く */
        var emphasisRow = addRow(dialog);
        emphasisRow.add("statictext", undefined, "").preferredSize.width = UI.labelWidth;

        var emphasisCheckbox = emphasisRow.add("checkbox", undefined, "");
        emphasisCheckbox.value = DEFAULTS.emphasisEnabled;

        var emphasisPrefix = getLabel(LABELS.label.emphasisPrefix);

        if (emphasisPrefix.length > 0) {
            emphasisRow.add("statictext", undefined, emphasisPrefix);
        }

        var emphasisInput = emphasisRow.add("edittext", undefined, String(DEFAULTS.emphasisEvery));
        emphasisInput.characters = 3;
        emphasisInput.enabled = emphasisCheckbox.value;
        changeValueByArrowKey(emphasisInput, 1);

        emphasisRow.add("statictext", undefined, getLabel(LABELS.label.emphasisSuffix));

        emphasisCheckbox.onClick = function () {
            emphasisInput.enabled = emphasisCheckbox.value;
        };

        /* 各文字に十字線 */
        var crossRow = addRow(dialog);
        crossRow.add("statictext", undefined, "").preferredSize.width = UI.labelWidth;

        var crossCheckbox = crossRow.add("checkbox", undefined, getLabel(LABELS.checkbox.cross));
        crossCheckbox.value = DEFAULTS.showCross;

        /* ボタンエリア */
        var btnRowGroup = dialog.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.alignChildren = ["left", "center"];
        btnRowGroup.alignment = ["fill", "center"];

        /* スペーサー（右側のボタンを押し出す） */
        var spacer = btnRowGroup.add("group");
        spacer.alignment = ["fill", "fill"];
        spacer.minimumSize.width = 0;

        var btnRightGroup = btnRowGroup.add("group");
        btnRightGroup.alignment = ["right", "center"];
        btnRightGroup.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });
        btnRightGroup.add("button", undefined, getLabel(LABELS.button.ok), { name: "ok" });

        if (dialog.show() !== 1) {
            return null;
        }

        var extensionMM = Number(extensionInput.text);

        if (isNaN(extensionMM) || extensionMM < 0) {
            extensionMM = 0;
        }

        var emphasisEvery = Math.round(Number(emphasisInput.text));

        if (isNaN(emphasisEvery) || emphasisEvery < 1) {
            emphasisEvery = 0;
        }

        return {
            extension: mmToPt(extensionMM),
            dashedHorizontal: dashedRadio.value,
            emphasisEvery: emphasisCheckbox.value ? emphasisEvery : 0,
            showCross: crossCheckbox.value
        };
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択した縦組みテキストに罫線を引く
     * @return {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel(LABELS.alert.noDocument));
            return;
        }

        var doc = app.activeDocument;
        var selectedItems = doc.selection;

        if (!selectedItems || selectedItems.length !== 1 || selectedItems[0].typename !== "TextFrame") {
            alert(getLabel(LABELS.alert.noTextFrame));
            return;
        }

        var textFrame = selectedItems[0];

        if (textFrame.textRange.characters.length === 0) {
            alert(getLabel(LABELS.alert.noCharacter));
            return;
        }

        /* ドキュメントを触る前に設定を決める / Settle the settings before touching the document */
        var settings = showSettingsDialog();

        if (!settings) {
            return;
        }

        if (CONFIG.forceVertical && textFrame.orientation !== TextOrientation.VERTICAL) {
            textFrame.orientation = TextOrientation.VERTICAL;
        }

        var failedAttributes = normalizeCharacterAttributes(textFrame);

        /* 先頭文字の文字サイズを基準にする（比率を変えても size は元の値のまま） */
        var fontSize = textFrame.textRange.characters[0].characterAttributes.size;

        if (!fontSize || fontSize <= 0) {
            alert(getLabel(LABELS.alert.noFontSize));
            return;
        }

        /* 罫線の間隔は文字サイズそのもの。垂直比率は反映しない */
        var cellHeight = fontSize;

        /* 属性変更後の再組版を反映させてから境界を読む */
        app.redraw();

        var bounds = textFrame.geometricBounds;

        var layerName = getLabel(LABELS.item.layerName);
        var groupName = getLabel(LABELS.item.groupName);

        var ruleLayer = prepareRuleLayer(doc, layerName);
        removeExistingGroup(ruleLayer, groupName);

        var group = ruleLayer.groupItems.add();
        group.name = groupName;

        /* 最後の文字の下にも罫線を引くため、マス数は列の文字数から求める */
        var cellCount = getMaxCharactersPerLine(textFrame);

        if (cellCount <= 0) {
            cellCount = Math.ceil((bounds[1] - bounds[3]) / cellHeight);
        }

        var colors = {
            rule: createRuleColor(doc, LAYOUT.strokeGray),
            cross: createRuleColor(doc, LAYOUT.crossGray)
        };

        drawGrid(group, bounds, cellHeight, cellCount, settings, colors);

        /* グループの zOrder だけではテキストの下に回らないため、レイヤーごと最背面へ送る */
        sendLayerToBack(ruleLayer);

        /* 成功時は通知しない。設定できなかった属性があるときだけ知らせる */
        if (failedAttributes.length > 0) {
            alert(getLabel(LABELS.alert.skipped) + " " + failedAttributes.join(", "));
        }
    }

    main();
})();
