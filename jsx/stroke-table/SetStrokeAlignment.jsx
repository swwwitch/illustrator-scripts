#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したオブジェクトに、線幅・線端・角の形状・線の位置をまとめて設定します。
DOMからは変更できない線の位置は、一時アクション（ダイナミックアクション）で適用します。

詳細は README を参照してください。

### Overview

Applies the stroke weight, cap, corner join, and alignment to the selected objects at once.
The alignment, which the DOM cannot change, is applied through a temporary action.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SetStrokeAlignment";           /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-09-11";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-11";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SetStrokeAlignment.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SetStrokeAlignment.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================
    var CONFIG = {
        defaultStrokeWidth: "1",          /* 線幅の初期値（pt）/ default stroke weight in pt */
        defaultStrokeCap: "butt",         /* 線端の初期値（butt / round / projecting）/ default cap */
        defaultCornerJoin: "miter",       /* 角の形状の初期値（miter / round / bevel）/ default join */
        defaultStrokeAlignment: "center", /* 線の位置の初期値（center / inside / outside）/ default alignment */
        previewDefault: true              /* プレビューを既定でON / preview on by default */
    };

    // =========================================
    // 一時アクション / Temporary action
    // =========================================
    var ACTION_SET_NAME   = "StrokeAlign";  /* アクションセット名（英数字）/ action set name (ASCII) */
    var ACTION_NAME       = "SetStroke";    /* アクション名（英数字）/ action name (ASCII) */
    var ACTION_PARAM_NAME = "Alignment";    /* パラメーターの表示名 / parameter display name */

    /* 線幅・線端・角の形状はDOMで設定できるので、アクションは線の位置だけを担当する
       / the DOM covers weight, cap and join, so the action only carries the alignment */
    var ACTION_KEY_STROKE_ALIGNMENT = 1634494318;  /* 'algn' 線の位置 / stroke alignment */

    /* 線の位置の識別子 → アクションの enumerated 値 / alignment keys to action enum values */
    var STROKE_ALIGNMENT_VALUES = {
        center: 0,
        inside: 1,
        outside: 2
    };

    // =========================================
    // ローカライズ / Localize
    // =========================================

    /* 実行環境の言語を判定（ja / en）/ Detect UI language */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = getCurrentLang();

    /* ローカライズ文字列を取得（キー漏れ時は英語へフォールバック）/ Get localized text, fallback to English
       @param {object} labelSet - { ja, en } のラベルセット
       @returns {string} 表示する文字列 */
    function getLabel(labelSet) {
        if (!labelSet) return "";
        if (labelSet[uiLang] != null) return labelSet[uiLang];
        return (labelSet.en != null) ? labelSet.en : "";
    }

    /* コロン付きラベル（日本語は全角、英語は半角）/ Label with colon (full-width JA, half-width EN)
       @param {object} labelSet - { ja, en } のラベルセット
       @returns {string} コロンを付けた文字列 */
    function labelText(labelSet) {
        return getLabel(labelSet) + (uiLang === "ja" ? "：" : ":");
    }

    var LABELS = {
        dialog: {
            title: { ja: "線の設定", en: "Stroke Settings" }
        },
        fieldLabel: {
            strokeWidth: { ja: "線幅", en: "Weight" },
            strokeCap: { ja: "線端", en: "Cap" },
            cornerJoin: { ja: "角の形状", en: "Corner" },
            strokeAlignment: { ja: "線の位置", en: "Alignment" }
        },
        radio: {
            buttCap: { ja: "なし", en: "Butt" },
            roundCap: { ja: "丸型", en: "Round" },
            projectingCap: { ja: "突出", en: "Projecting" },
            miterJoin: { ja: "マイター", en: "Miter" },
            roundJoin: { ja: "ラウンド", en: "Round" },
            bevelJoin: { ja: "ベベル", en: "Bevel" },
            centerAlign: { ja: "中央", en: "Center" },
            insideAlign: { ja: "内側", en: "Inside" },
            outsideAlign: { ja: "外側", en: "Outside" }
        },
        // 線パネルのツールチップと同じ表記 / same wording as the Stroke panel tooltips
        tooltip: {
            buttCap: { ja: "線端なし", en: "Butt Cap" },
            roundCap: { ja: "丸型線端", en: "Round Cap" },
            projectingCap: { ja: "突出線端", en: "Projecting Cap" },
            miterJoin: { ja: "マイター結合", en: "Miter Join" },
            roundJoin: { ja: "ラウンド結合", en: "Round Join" },
            bevelJoin: { ja: "ベベル結合", en: "Bevel Join" },
            centerAlign: { ja: "線を中央に揃える", en: "Align Stroke to Center" },
            insideAlign: { ja: "線を内側に揃える", en: "Align Stroke to Inside" },
            outsideAlign: { ja: "線を外側に揃える", en: "Align Stroke to Outside" }
        },
        checkbox: {
            preview: { ja: "プレビュー", en: "Preview" }
        },
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" }
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noSelection: { ja: "オブジェクトを選択してください。", en: "Select one or more objects." },
            invalidWidth: { ja: "線幅には0より大きい数値を入力してください。", en: "Enter a stroke weight greater than 0." }
        }
    };

    /* ラジオボタンの並び（線パネルと同じ順序）/ Radio button order, matching the Stroke panel */
    var STROKE_CAP_OPTIONS = [
        { key: "butt", label: LABELS.radio.buttCap, tooltip: LABELS.tooltip.buttCap },
        { key: "round", label: LABELS.radio.roundCap, tooltip: LABELS.tooltip.roundCap },
        { key: "projecting", label: LABELS.radio.projectingCap, tooltip: LABELS.tooltip.projectingCap }
    ];
    var CORNER_JOIN_OPTIONS = [
        { key: "miter", label: LABELS.radio.miterJoin, tooltip: LABELS.tooltip.miterJoin },
        { key: "round", label: LABELS.radio.roundJoin, tooltip: LABELS.tooltip.roundJoin },
        { key: "bevel", label: LABELS.radio.bevelJoin, tooltip: LABELS.tooltip.bevelJoin }
    ];
    var STROKE_ALIGNMENT_OPTIONS = [
        { key: "center", label: LABELS.radio.centerAlign, tooltip: LABELS.tooltip.centerAlign },
        { key: "inside", label: LABELS.radio.insideAlign, tooltip: LABELS.tooltip.insideAlign },
        { key: "outside", label: LABELS.radio.outsideAlign, tooltip: LABELS.tooltip.outsideAlign }
    ];

    // =========================================
    // レイアウト / Layout
    // =========================================
    var WINDOW_MARGINS        = 16;                          /* ウィンドウ外周の余白 */
    var WINDOW_SPACING        = 10;                          /* ウィンドウ内の要素間隔 */
    var ROW_SPACING           = 8;                           /* 行内の要素間隔 */
    var BUTTON_ROW_TOP_MARGIN = 8;                           /* ボタン行の上余白 */
    var LABEL_WIDTH           = (uiLang === "ja") ? 76 : 74; /* 行ラベルの共通幅（コロンまで収まる幅）*/

    /* ダイアログの余白・間隔をそろえる / Apply common window metrics
       @param {Window} win - 対象のダイアログ
       @returns {void} */
    function setupWindow(win) {
        win.orientation = "column";
        win.alignChildren = ["fill", "top"];
        win.margins = WINDOW_MARGINS;
        win.spacing = WINDOW_SPACING;
    }

    /* 行グループの向き・整列をそろえる / Apply common row metrics
       @param {Group} targetGroup - 対象のグループ
       @param {string} [alignment] - 横方向の整列（既定 "left"）
       @param {number} [spacing] - 要素間隔（既定 ROW_SPACING）
       @returns {void} */
    function setupRow(targetGroup, alignment, spacing) {
        targetGroup.orientation = "row";
        targetGroup.alignment = [alignment || "left", "center"];
        targetGroup.alignChildren = ["left", "center"];
        targetGroup.spacing = (typeof spacing === "number") ? spacing : ROW_SPACING;
    }

    /* ラベル付きの行を追加し、コントロールを入れるグループを返す / Add a labeled row and return its control group
       @param {Window|Group} parentContainer - 追加先
       @param {object} labelSet - 行ラベルのラベルセット
       @returns {Group} コントロールを入れるグループ */
    function addControlRow(parentContainer, labelSet) {
        var controlRow = parentContainer.add("group");
        setupRow(controlRow);

        var rowLabel = controlRow.add("statictext", undefined, labelText(labelSet));
        rowLabel.preferredSize.width = LABEL_WIDTH;
        rowLabel.justify = "right";

        var rowControlGroup = controlRow.add("group");
        setupRow(rowControlGroup);
        return rowControlGroup;
    }

    /* ラジオボタンの行を追加する / Add a labeled row of radio buttons
       @param {Window|Group} parentContainer - 追加先
       @param {object} labelSet - 行ラベルのラベルセット
       @param {Array} optionDefinitions - 選択肢の定義（key / label / tooltip）
       @param {string} selectedOptionKey - 初期選択のキー
       @param {function} onSelect - 選択が変わったときに呼ぶ処理
       @returns {Group} 選択状態を selectedOptionKey に持つグループ */
    function addRadioRow(parentContainer, labelSet, optionDefinitions, selectedOptionKey, onSelect) {
        var radioGroup = addControlRow(parentContainer, labelSet);
        radioGroup.selectedOptionKey = selectedOptionKey;
        for (var i = 0; i < optionDefinitions.length; i++) {
            var optionRadio = radioGroup.add("radiobutton", undefined, getLabel(optionDefinitions[i].label));
            optionRadio.helpTip = getLabel(optionDefinitions[i].tooltip);
            optionRadio.optionKey = optionDefinitions[i].key;
            optionRadio.value = (optionDefinitions[i].key === selectedOptionKey);
            optionRadio.onClick = function () {
                radioGroup.selectedOptionKey = this.optionKey;
                onSelect();
            };
        }
        return radioGroup;
    }

    /* 数値入力欄に↑↓キーでの増減を付ける / Arrow-key stepping for a numeric field
       @param {EditText} editText - 対象の入力欄
       @param {function} [onUpdate] - 値更新後に呼ぶコールバック
       @returns {void} */
    function changeValueByArrowKey(editText, onUpdate) {
        editText.addEventListener("keydown", function (event) {
            var value = Number(editText.text);
            if (isNaN(value)) return;

            // 修飾キーは event から読む（keyboardState は macOS で誤報あり）。取得不可時のみフォールバック。
            var keyboardState = ScriptUI.environment.keyboardState;
            var isShiftPressed = (event.shiftKey !== undefined) ? event.shiftKey : keyboardState.shiftKey;
            var isOptionPressed = (event.altKey !== undefined) ? event.altKey : keyboardState.altKey;

            if (event.keyName == "Up" || event.keyName == "Down") {
                var isUp = (event.keyName == "Up");
                if (isShiftPressed) {
                    // 10の倍数にスナップ（倍数上ならさらに±10） / snap to the next multiple of 10
                    value = isUp ? (Math.floor(value / 10) * 10 + 10) : (Math.ceil(value / 10) * 10 - 10);
                } else {
                    var step = isOptionPressed ? 0.1 : 1;
                    value = value + (isUp ? step : -step);
                }
                // 浮動小数の誤差を丸める（0.1刻み対応） / trim float error for 0.1 steps
                value = Math.round(value * 10000) / 10000;

                event.preventDefault();
                editText.text = value;
                if (typeof onUpdate === "function") onUpdate(editText.text);
            }
        });
    }

    // =========================================
    // プレビュー / Preview
    // =========================================

    /* プレビュー管理（実行した手数を数え、undoで巻き戻す）/ Preview manager that rolls back by counting undo steps */
    function PreviewManager() {
        this.undoDepth = 0;

        /* 変更処理を実行して1手として数える（false を返した場合は数えない）
           / Run a change and count one undo step; a callback returning false is not counted
           @param {function} changeAction - 実行する処理
           @returns {void} */
        this.addStep = function (changeAction) {
            try {
                if (changeAction() === false) return;
                this.undoDepth++;
                app.redraw();
            } catch (e) {
                $.writeln("[PreviewManager] addStep error: " + e);
            }
        };

        /* プレビューによる変更をすべて取り消す / Undo every preview change
           @returns {void} */
        this.rollback = function () {
            while (this.undoDepth > 0) {
                app.undo();
                this.undoDepth--;
            }
            app.redraw();
        };

        /* 巻き戻してから本処理を一度だけ実行する / Roll back, then run the final action once
           @param {function} confirmAction - 巻き戻したあとに実行する処理
           @returns {void} */
        this.confirm = function (confirmAction) {
            this.rollback();
            confirmAction();
            this.undoDepth = 0;
        };
    }

    // =========================================
    // アクションの生成と実行 / Build and run the action
    // =========================================

    /* ASCII文字列を .aia の名前用16進表記に変換 / Convert an ASCII string to .aia name hex
       @param {string} asciiText - 変換する文字列（英数字）
       @returns {string} 16進表記 */
    function toActionNameHex(asciiText) {
        var hexText = "";
        for (var i = 0; i < asciiText.length; i++) {
            var charHex = asciiText.charCodeAt(i).toString(16);
            hexText += (charHex.length < 2) ? "0" + charHex : charHex;
        }
        return hexText;
    }

    /* .aia の名前ブロック（バイト数＋16進表記）を組み立てる / Build a name block: byte count plus hex
       @param {string} nameKey - /name または /localizedName
       @param {string} asciiText - 名前（英数字）
       @param {string} indent - 行頭のタブ
       @returns {Array<string>} 行の配列 */
    function buildActionNameLines(nameKey, asciiText, indent) {
        return [
            indent + nameKey + " [ " + asciiText.length,
            indent + "\t" + toActionNameHex(asciiText),
            indent + "]"
        ];
    }

    /* 線の位置のパラメーターブロックを組み立てる / Build the alignment parameter block
       @param {number} alignmentValue - 線の位置（0=中央 / 1=内側 / 2=外側）
       @returns {Array<string>} 行の配列 */
    function buildAlignmentParameterLines(alignmentValue) {
        return [
            "\t\t/parameter-1 {",
            "\t\t\t/key " + ACTION_KEY_STROKE_ALIGNMENT,
            "\t\t\t/showInPalette 4294967295",
            "\t\t\t/type (enumerated)"
        ]
            .concat(buildActionNameLines("/name", ACTION_PARAM_NAME, "\t\t\t"))
            .concat([
                "\t\t\t/value " + alignmentValue,
                "\t\t}"
            ]);
    }

    /* 線の位置を設定する一時アクションのコードを組み立てる / Build the .aia source that sets the stroke alignment
       @param {number} alignmentValue - 線の位置（0=中央 / 1=内側 / 2=外側）
       @returns {string} .aia のソース */
    function buildStrokeActionCode(alignmentValue) {
        return ["/version 3"]
            .concat(buildActionNameLines("/name", ACTION_SET_NAME, ""))
            .concat([
                "/isOpen 1",
                "/actionCount 1",
                "/action-1 {"
            ])
            .concat(buildActionNameLines("/name", ACTION_NAME, "\t"))
            .concat([
                "\t/keyIndex 0",
                "\t/colorIndex 0",
                "\t/isOpen 1",
                "\t/eventCount 1",
                "\t/event-1 {",
                "\t\t/useRulersIn1stQuadrant 0",
                "\t\t/internalName (ai_plugin_setStroke)"
            ])
            .concat(buildActionNameLines("/localizedName", ACTION_NAME, "\t\t"))
            .concat([
                "\t\t/isOpen 1",
                "\t\t/isOn 1",
                "\t\t/hasDialog 0",
                "\t\t/parameterCount 1"
            ])
            .concat(buildAlignmentParameterLines(alignmentValue))
            .concat(["\t}", "}"])
            .join("\n");
    }

    /* 一時アクションを読み込み、選択オブジェクトに線の位置を適用する / Apply the stroke alignment via a temporary action
       @param {number} alignmentValue - 線の位置（0=中央 / 1=内側 / 2=外側）
       @returns {void} */
    function applyStrokeByAction(alignmentValue) {
        /* 前回の実行が落ちて同名セットが残っていれば先に破棄（無ければ例外）
           / Unload a leftover set with the same name; missing is the normal case and throws */
        try { app.unloadAction(ACTION_SET_NAME, ""); } catch (e) { }

        var actionFile = new File(Folder.temp + "/" + ACTION_SET_NAME + ".aia");
        actionFile.encoding = "UTF-8";
        actionFile.lineFeed = "Unix";
        if (!actionFile.open("w")) {
            throw new Error("一時アクションファイルを作成できませんでした。");
        }
        actionFile.write(buildStrokeActionCode(alignmentValue));
        actionFile.close();

        /* 読み込み時点でパース済みなので、直後に削除して一時ファイルを残さない / parsed on load, so remove it right away */
        app.loadAction(actionFile);
        actionFile.remove();

        try {
            app.doScript(ACTION_NAME, ACTION_SET_NAME, false);
        } finally {
            app.unloadAction(ACTION_SET_NAME, "");
        }
    }

    // =========================================
    // 線の適用 / Apply strokes
    // =========================================

    /* 線端・角の形状の識別子 → DOMの列挙値 / option keys to DOM enums */
    var STROKE_CAP_ENUMS = {
        butt: StrokeCap.BUTTENDCAP,
        round: StrokeCap.ROUNDENDCAP,
        projecting: StrokeCap.PROJECTINGENDCAP
    };
    var CORNER_JOIN_ENUMS = {
        miter: StrokeJoin.MITERENDJOIN,
        round: StrokeJoin.ROUNDENDJOIN,
        bevel: StrokeJoin.BEVELENDJOIN
    };

    /* 黒のカラーを取得（スウォッチが無い場合はカラーモードから作る）/ Get black, falling back to the document color space
       @param {Document} targetDocument - 対象ドキュメント
       @returns {Color} 黒のカラー */
    function getBlackColor(targetDocument) {
        try {
            return targetDocument.swatches["[Black]"].color;
        } catch (e) {
            if (targetDocument.documentColorSpace === DocumentColorSpace.RGB) {
                var rgbBlack = new RGBColor();
                rgbBlack.red = 0;
                rgbBlack.green = 0;
                rgbBlack.blue = 0;
                return rgbBlack;
            }
            var cmykBlack = new CMYKColor();
            cmykBlack.cyan = 0;
            cmykBlack.magenta = 0;
            cmykBlack.yellow = 0;
            cmykBlack.black = 100;
            return cmykBlack;
        }
    }

    /* 選択アイテムを再帰的にたどり、線を持てるものに処理を適用する / Walk the selection, applying a callback to every item that can take a stroke
       @param {Array} targetItems - 対象アイテムの配列
       @param {function} applyToItem - 1アイテムに対する処理（false を返すと数えない）
       @returns {number} 処理したアイテム数 */
    function forEachStrokableItem(targetItems, applyToItem) {
        var appliedCount = 0;
        for (var i = 0; i < targetItems.length; i++) {
            var targetItem = targetItems[i];
            if (targetItem.typename === "GroupItem") {
                appliedCount += forEachStrokableItem(targetItem.pageItems, applyToItem);
                continue;
            }
            if (targetItem.typename === "CompoundPathItem") {
                appliedCount += forEachStrokableItem(targetItem.pathItems, applyToItem);
                continue;
            }
            try {
                if (applyToItem(targetItem) !== false) appliedCount++;
            } catch (e) {
                /* 線を持てないオブジェクトはスキップ / skip items that cannot take a stroke */
            }
        }
        return appliedCount;
    }

    /* 閉じたパスが含まれるかを調べる（グループ・複合パスは再帰）/ Does the selection contain a closed path?
       @param {Array} targetItems - 対象アイテムの配列
       @returns {boolean} 閉じたパスがあれば true */
    function hasClosedPath(targetItems) {
        for (var i = 0; i < targetItems.length; i++) {
            var targetItem = targetItems[i];
            if (targetItem.typename === "GroupItem") {
                if (hasClosedPath(targetItem.pageItems)) return true;
                continue;
            }
            if (targetItem.typename === "CompoundPathItem") {
                if (hasClosedPath(targetItem.pathItems)) return true;
                continue;
            }
            if (targetItem.typename === "PathItem" && targetItem.closed) return true;
        }
        return false;
    }

    /* 線が無いアイテムに黒の線を付ける / Give unstroked items a black stroke
       @param {Array} targetItems - 対象アイテムの配列
       @param {Color} blackColor - 適用する黒
       @returns {number} 線を付けたアイテム数 */
    function ensureStrokeColor(targetItems, blackColor) {
        return forEachStrokableItem(targetItems, function (targetItem) {
            /* すでに線があれば色は触らない / existing stroke colors stay */
            if (targetItem.stroked && targetItem.strokeColor.typename !== "NoColor") return false;
            targetItem.stroked = true;
            targetItem.strokeColor = blackColor;
        });
    }

    /* 線幅・線端・角の形状をDOMで設定する / Set weight, cap and join through the DOM
       @param {Array} targetItems - 対象アイテムの配列
       @param {object} strokeSettings - { strokeWidth, strokeCap, cornerJoin }
       @returns {number} 設定できたアイテム数 */
    function applyStrokeProperties(targetItems, strokeSettings) {
        return forEachStrokableItem(targetItems, function (targetItem) {
            targetItem.strokeWidth = strokeSettings.strokeWidth;
            targetItem.strokeCap = STROKE_CAP_ENUMS[strokeSettings.strokeCap];
            targetItem.strokeJoin = CORNER_JOIN_ENUMS[strokeSettings.cornerJoin];
        });
    }

    /* 線の設定を適用する / Apply the stroke settings
       アクションはパラメーターに書いていない項目（線幅・線端・角の形状）を既定値に戻すため、
       線の位置をアクションで決めたあとにDOMで上書きする。
       / the action resets the settings it does not carry, so the DOM pass has to run after it
       @param {Array} targetItems - 対象アイテムの配列
       @param {Color} blackColor - 線が無いアイテムに付ける黒
       @param {object} strokeSettings - { strokeWidth, strokeCap, cornerJoin, strokeAlignment }
       @param {PreviewManager} [previewManager] - 指定するとプレビューの手数として記録する
       @returns {void} */
    function applyStroke(targetItems, blackColor, strokeSettings, previewManager) {
        /* strokeAlignment が null＝線の位置を使えない選択なので、アクションは流さない
           / a null alignment means the selection cannot take one, so the action is skipped */
        var usesAlignment = (strokeSettings.strokeAlignment != null);
        var alignmentValue = usesAlignment ? STROKE_ALIGNMENT_VALUES[strokeSettings.strokeAlignment] : 0;

        if (!previewManager) {
            ensureStrokeColor(targetItems, blackColor);
            app.redraw();
            if (usesAlignment) applyStrokeByAction(alignmentValue);
            applyStrokeProperties(targetItems, strokeSettings);
            return;
        }
        /* 何も変えなかったパスは1手として数えない / a pass that changed nothing is not counted */
        previewManager.addStep(function () {
            return ensureStrokeColor(targetItems, blackColor) > 0;
        });
        if (usesAlignment) {
            previewManager.addStep(function () {
                applyStrokeByAction(alignmentValue);
            });
        }
        previewManager.addStep(function () {
            return applyStrokeProperties(targetItems, strokeSettings) > 0;
        });
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /* 線の設定ダイアログを表示し、OKなら適用する / Show the stroke dialog and apply on OK
       @param {Array} targetItems - 対象アイテムの配列
       @param {Color} blackColor - 線が無いアイテムに付ける黒
       @returns {object} 適用した設定（キャンセル時は null） */
    function showStrokeDialog(targetItems, blackColor) {
        var dialog = new Window("dialog", getLabel(LABELS.dialog.title) + " " + SCRIPT_VERSION);
        setupWindow(dialog);

        /* 線幅 / Stroke weight */
        var strokeWidthGroup = addControlRow(dialog, LABELS.fieldLabel.strokeWidth);
        var strokeWidthInput = strokeWidthGroup.add("edittext", undefined, CONFIG.defaultStrokeWidth);
        strokeWidthInput.characters = 5;
        strokeWidthGroup.add("statictext", undefined, "pt");

        /* 線端・角の形状・線の位置 / Cap, join, alignment */
        var strokeCapGroup = addRadioRow(dialog, LABELS.fieldLabel.strokeCap,
            STROKE_CAP_OPTIONS, CONFIG.defaultStrokeCap, function () { updatePreview(); });
        var cornerJoinGroup = addRadioRow(dialog, LABELS.fieldLabel.cornerJoin,
            CORNER_JOIN_OPTIONS, CONFIG.defaultCornerJoin, function () { updatePreview(); });
        var strokeAlignmentGroup = addRadioRow(dialog, LABELS.fieldLabel.strokeAlignment,
            STROKE_ALIGNMENT_OPTIONS, CONFIG.defaultStrokeAlignment, function () { updatePreview(); });

        /* 線の位置は閉じたパスにしか効かないので、無ければ行ごとディムにする（ラベルも含めるため親の行を無効化）
           / alignment only works on closed paths; disable the whole row (parent, so the label dims too) */
        var strokeAlignmentRow = strokeAlignmentGroup.parent;
        strokeAlignmentRow.enabled = hasClosedPath(targetItems);

        /* ボタンエリア（左：プレビュー／右：キャンセル・OK）/ Button area */
        var btnRowGroup = dialog.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.margins = [0, BUTTON_ROW_TOP_MARGIN, 0, 0];
        btnRowGroup.alignment = ["fill", "bottom"];

        var btnLeftGroup = btnRowGroup.add("group");
        btnLeftGroup.alignChildren = ["left", "center"];
        var previewCheckbox = btnLeftGroup.add("checkbox", undefined, getLabel(LABELS.checkbox.preview));
        previewCheckbox.value = CONFIG.previewDefault;

        var spacer = btnRowGroup.add("group");
        spacer.alignment = ["fill", "fill"];
        spacer.minimumSize.width = 0;

        var btnRightGroup = btnRowGroup.add("group");
        btnRightGroup.alignChildren = ["right", "center"];
        var btnCancel = btnRightGroup.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });
        var btnOK = btnRightGroup.add("button", undefined, getLabel(LABELS.button.ok), { name: "ok" });

        /* 入力値を読み取る（不正なら null）/ Read the current input, null when invalid
           @returns {object} 線の設定 */
        function readStrokeSettings() {
            var strokeWidth = parseFloat(strokeWidthInput.text);
            if (isNaN(strokeWidth) || strokeWidth <= 0) return null;
            return {
                strokeWidth: strokeWidth,
                strokeCap: strokeCapGroup.selectedOptionKey,
                cornerJoin: cornerJoinGroup.selectedOptionKey,
                strokeAlignment: strokeAlignmentRow.enabled ? strokeAlignmentGroup.selectedOptionKey : null
            };
        }

        var previewManager = new PreviewManager();

        /* Undo履歴を汚さずにプレビューを更新する / Refresh the preview without polluting the undo history
           @returns {void} */
        function updatePreview() {
            previewManager.rollback();
            if (!previewCheckbox.value) return;
            var strokeSettings = readStrokeSettings();
            if (!strokeSettings) return; /* 入力途中はプレビューしない / skip while the input is incomplete */
            applyStroke(targetItems, blackColor, strokeSettings, previewManager);
        }

        changeValueByArrowKey(strokeWidthInput, updatePreview);
        strokeWidthInput.onChanging = updatePreview;
        previewCheckbox.onClick = updatePreview;

        btnOK.onClick = function () {
            if (!readStrokeSettings()) {
                alert(getLabel(LABELS.alert.invalidWidth));
                return;
            }
            dialog.close(1);
        };
        btnCancel.onClick = function () {
            dialog.close(0);
        };

        updatePreview();
        strokeWidthInput.active = true;

        if (dialog.show() !== 1) {
            /* キャンセル：プレビューを巻き戻す / Cancel: roll back the preview */
            previewManager.rollback();
            return null;
        }

        /* プレビューを巻き戻してから、確定として一度だけ適用する / Roll back the preview, then apply once */
        var confirmedSettings = readStrokeSettings();
        previewManager.confirm(function () {
            applyStroke(targetItems, blackColor, confirmedSettings);
            app.redraw();
        });
        return confirmedSettings;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /* 選択オブジェクトに線の設定を適用する / Apply the stroke settings to the selection
       @returns {void} */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel(LABELS.alert.noDocument));
            return;
        }
        var activeDocument = app.activeDocument;
        var targetItems = activeDocument.selection;
        /* 文字ツールでの文字選択（TextRange）は対象外 / a type-tool selection is not a page-item array */
        if (!(targetItems instanceof Array) || targetItems.length === 0) {
            alert(getLabel(LABELS.alert.noSelection));
            return;
        }

        showStrokeDialog(targetItems, getBlackColor(activeDocument));
    }

    main();

})();
