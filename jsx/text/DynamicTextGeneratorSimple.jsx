#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択した1つのテキストフレームの各行をアウトライン幅で測定し、最長行の幅にそろうよう行ごとの文字サイズを変倍します。ダイアログを開いたまま結果をプレビューでき、行送りを自動（既定100%）に切り替えるかどうかも選べます。

詳細はREADMEを参照。

*/

/*

### Overview

Measures each line of a single selected text frame by its outline width and scales every line's font size so all lines match the widest one. The result can be previewed while the dialog stays open, and leading can optionally be switched to auto (100% by default).

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "DynamicTextGeneratorSimple";   /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-08-11";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-08-11";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/DynamicTextGeneratorSimple.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/DynamicTextGeneratorSimple.md
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/xxxxxxxx"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* 幅をそろえたあとに行送りを自動へ切り替えるか（ダイアログの初期値）
       Switch leading to auto after fitting the widths (initial state in the dialog) */
    var APPLY_AUTO_LEADING = true;

    /* 自動行送りの比率（％）。applyAutoLeading() の既定値でもある
       Auto-leading ratio in percent; also the default used by applyAutoLeading() */
    var AUTO_LEADING_AMOUNT = 100;

    /* ダイアログを開いた直後からプレビューを表示するか / Show the preview as soon as the dialog opens */
    var PREVIEW_ON_OPEN = true;

    /* 変倍率がこの範囲内なら誤差とみなして変更しない / Treat ratios within this range as no change */
    var RATIO_EPSILON = 0.001;

    /* ダイアログの透明度 / Dialog opacity */
    var DIALOG_OPACITY = 0.98;

    // =========================================
    // レイアウト / Layout
    // =========================================
    var PANEL_MARGINS   = [16, 20, 16, 12];  /* パネル余白 [左,上,右,下] / panel margins */
    var PANEL_SPACING   = 8;                 /* パネル内の標準間隔 / default spacing inside panels */
    var FIELD_SPACING   = 6;                 /* 入力行どうしの間隔 / spacing between field rows */
    var FIELD_INDENT    = 20;                /* チェックボックスに従属する行の字下げ / indent for rows owned by a checkbox */
    var BUTTON_WIDTH    = 90;                /* OK・キャンセルの幅 / width of OK and Cancel */
    var EDIT_CHARACTERS = 4;                 /* 入力欄の文字数 / width of an edittext in characters */

    // =========================================
    // ローカライズ / Localization
    // =========================================
    var uiLang = ($.locale.indexOf("ja") === 0) ? "ja" : "en";

    var LABELS = {
        dialog: {
            title: { ja: "行の幅そろえ", en: "Fit Lines to Widest" }
        },
        panel: {
            leading: { ja: "行送り", en: "Leading" }
        },
        checkbox: {
            autoLeading: { ja: "行送りを自動にする", en: "Switch leading to auto" },
            preview: { ja: "プレビュー", en: "Preview" }
        },
        fieldLabel: {
            autoLeadingAmount: { ja: "自動行送り", en: "Auto leading" }
        },
        unit: {
            percent: { ja: "%", en: "%" }
        },
        button: {
            ok: { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        tooltip: {
            preview: {
                ja: "ONの間は結果を仮表示します。OFFにするかキャンセルすると元に戻ります。",
                en: "Shows a temporary result while enabled. Turning it off or cancelling restores the original."
            },
            autoLeading: {
                ja: "幅をそろえたあとに、行送りを自動へ切り替えます。",
                en: "Switches leading to auto after the widths are fitted."
            },
            autoLeadingAmount: {
                ja: "自動行送りの比率です。文字サイズに対する行送りの割合を指定します。",
                en: "The auto-leading ratio, given as a percentage of the font size."
            }
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            selectOneTextFrame: { ja: "テキストフレームを1つだけ選択してください。", en: "Select exactly one text frame." },
            needTwoLines: { ja: "2行以上のテキストを選択してください。", en: "Select text with two or more lines." },
            noMeasurableLine: { ja: "有効なテキスト行が見つかりませんでした。", en: "No measurable text line was found." },
            previewError: { ja: "プレビュー更新エラー：", en: "Preview update error: " }
        }
    };

    /**
     * LABELS を上から順にたどって現在のUI言語のラベルを返す
     * @param {...string} - LABELS をたどるキー
     * @returns {string} ラベル文字列（見つからない場合は空文字）
     */
    function getLabel() {
        var node = LABELS;
        for (var i = 0; i < arguments.length; i++) {
            if (node == null) break;
            node = node[arguments[i]];
        }
        return (node && node[uiLang] != null) ? node[uiLang] : "";
    }

    /**
     * コロン付きのラベルを返す（日本語は全角コロン、英語は半角コロン）
     * @param {...string} - LABELS をたどるキー
     * @returns {string} コロンを付けたラベル文字列
     */
    function getLabelWithColon() {
        return getLabel.apply(null, arguments) + (uiLang === "ja" ? "：" : ":");
    }

    // =========================================
    // 数値ユーティリティ / Number helpers
    // =========================================

    /**
     * 小数第1位に丸める
     * @param {number} value - 丸める値
     * @returns {number} 丸めた値
     */
    function roundToTenth(value) {
        return Math.round(value * 10) / 10;
    }

    /**
     * 入力欄の文字列を数値として読み取る
     * @param {EditText} editText - 対象の入力欄
     * @param {number} fallback - 数値として読めない場合に返す値
     * @returns {number} 読み取った数値
     */
    function readNumber(editText, fallback) {
        var value = Number(editText.text);
        return (editText.text !== "" && !isNaN(value)) ? value : fallback;
    }

    // =========================================
    // UI部品 / UI helpers
    // =========================================

    /**
     * パネルに共通のレイアウト設定を適用する
     * @param {Panel} panel - 対象のパネル
     * @param {number} [spacing] - パネル内の間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function setupPanel(panel, spacing) {
        panel.orientation = "column";
        panel.alignChildren = ["fill", "top"];
        panel.alignment = "fill";
        panel.margins = PANEL_MARGINS;
        panel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /**
     * グループに共通のレイアウト設定を適用する
     * @param {Group} group - 対象のグループ
     * @param {string} [orientation] - "row" または "column"（省略時は "column"）
     * @param {number} [spacing] - グループ内の間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function setupGroup(group, orientation, spacing) {
        var groupOrientation = orientation || "column";
        group.orientation = groupOrientation;
        /* row は横並びなので縦中央、column は縦並びなので左揃え / row: vertically centered, column: left-aligned */
        group.alignChildren = (groupOrientation === "row") ? ["left", "center"] : ["left", "top"];
        group.alignment = "fill";
        group.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /**
     * ラベル＋入力欄＋単位の行を追加する
     * @param {Panel|Group} parent - 追加先のコンテナ
     * @param {string} labelText - 行の見出し（コロン付き）
     * @param {string} initialText - 入力欄の初期値
     * @param {string} unitText - 単位表示の文字列
     * @returns {{label: StaticText, input: EditText, unit: StaticText}} 生成した行の各コントロール
     */
    function addFieldRow(parent, labelText, initialText, unitText) {
        var row = parent.add("group");
        setupGroup(row, "row", FIELD_SPACING);
        row.margins = [FIELD_INDENT, 0, 0, 0];

        var label = row.add("statictext", undefined, labelText);
        var input = row.add("edittext", undefined, initialText);
        input.characters = EDIT_CHARACTERS;
        input.justify = "right";
        var unit = row.add("statictext", undefined, unitText);

        return { label: label, input: input, unit: unit };
    }

    /**
     * 行（ラベル＋入力欄＋単位）をまとめて有効／無効にする
     * @param {{label: StaticText, input: EditText, unit: StaticText}} row - addFieldRow が返した行オブジェクト
     * @param {boolean} enabled - 有効にするなら true
     * @returns {void}
     */
    function setFieldRowEnabled(row, enabled) {
        row.label.enabled = enabled;
        row.input.enabled = enabled;
        row.unit.enabled = enabled;
    }

    /**
     * 行（ラベル＋入力欄＋単位）にまとめてヘルプチップを設定する
     * @param {{label: StaticText, input: EditText, unit: StaticText}} row - addFieldRow が返した行オブジェクト
     * @param {string} tooltip - 設定するヘルプチップ文字列
     * @returns {void}
     */
    function setFieldRowTooltip(row, tooltip) {
        row.label.helpTip = tooltip;
        row.input.helpTip = tooltip;
        row.unit.helpTip = tooltip;
    }

    /**
     * 入力欄で上下キーによる数値の増減を有効にする
     * @param {EditText} editText - 対象の入力欄
     * @param {{step: number, shiftStep: number, altStep: number}} stepOptions - 通常／shift／option時の増減幅
     * @returns {void}
     */
    function changeValueByArrowKey(editText, stepOptions) {
        editText.addEventListener("keydown", function (event) {
            if (event.keyName != "Up" && event.keyName != "Down") return;
            var value = Number(editText.text);
            if (isNaN(value)) return;

            var keyboard = ScriptUI.environment.keyboardState;
            var isUp = (event.keyName == "Up");
            var isFine = (!keyboard.shiftKey && keyboard.altKey);
            var delta;
            if (keyboard.shiftKey) {
                /* shiftキーは刻み幅の倍数に丸めてから増減 / snap to a multiple of the step, then move */
                delta = stepOptions.shiftStep;
                value = isUp ? Math.floor(value / delta) * delta : Math.ceil(value / delta) * delta;
            } else {
                delta = isFine ? stepOptions.altStep : stepOptions.step;
            }
            value = isUp ? value + delta : value - delta;
            /* optionキーは小数第1位まで、それ以外は整数に丸める / option keeps one decimal, otherwise round to an integer */
            value = isFine ? roundToTenth(value) : Math.round(value);
            if (value < 0) value = 0;

            event.preventDefault();
            editText.text = value;
            editText.notify("onChange");
        });
    }

    // =========================================
    // PreviewManager
    // プレビュー時にUndo履歴を汚さないための小さな管理クラス。
    // - addStep(): 変更処理を実行してundoDepthをカウント
    // - rollback(): 適用済みのプレビューをすべて取り消し
    // undoDepth = 適用してまだ取り消していないステップ数 / steps applied and not yet undone
    // 適用中の例外でダイアログごと落ちないよう addStep だけ try で囲み、
    // 上下キー連打で同じアラートが溢れないよう同一メッセージは1度だけ表示する。
    // =========================================

    /**
     * プレビュー適用とUndoの深さを管理するクラス
     * @returns {void}
     */
    function PreviewManager() {
        this.undoDepth = 0;
        var lastReportedError = "";

        /**
         * 変更処理を1ステップとして実行し、Undoの深さを数える
         * @param {function} func - 実行する変更処理
         * @returns {void}
         */
        this.addStep = function (func) {
            try {
                func();
                lastReportedError = "";
                this.undoDepth++;
                app.redraw();
            } catch (e) {
                var message = getLabel("alert", "previewError") + e;
                if (message === lastReportedError) return;
                lastReportedError = message;
                alert(message);
            }
        };

        /**
         * 適用済みのプレビューをすべて取り消す
         * @returns {void}
         */
        this.rollback = function () {
            while (this.undoDepth > 0) {
                try {
                    app.undo();
                } catch (e) {
                    /* Undoに失敗したらカウンターのずれを防ぐため打ち切る / stop on failure so the counter cannot drift */
                    this.undoDepth = 0;
                    break;
                }
                this.undoDepth--;
            }
            app.redraw();
        };
    }

    // =========================================
    // 測定と変倍 / Measure & scale
    // =========================================

    /**
     * 行の測定結果
     * @typedef {object} LineMetric
     * @property {number} start - ストーリー内の開始文字インデックス
     * @property {number} end - ストーリー内の終了文字インデックス（この位置は含まない）
     * @property {number} width - 行の外形幅（pt）。測定できない行は0
     */

    /**
     * 改行や空白しか含まない行かどうかを判定する
     * @param {TextRange} line - 判定する行
     * @returns {boolean} 内容が空とみなせる場合 true
     */
    function isBlankLine(line) {
        return line.contents.replace(/[\r\n\x03\s　]/g, "").length === 0;
    }

    /**
     * 行の内容を一時テキストフレームへ複製し、アウトライン化して外形幅を測る
     * 一時オブジェクトは成否にかかわらず必ず削除する
     * @param {Document} doc - 対象ドキュメント
     * @param {TextRange} line - 測定する行
     * @returns {number} 行の外形幅（pt）。測定できない場合は0
     */
    function measureLineWidth(doc, line) {
        var tempTextFrame = null;
        var outlineGroup = null;
        var width = 0;

        try {
            tempTextFrame = doc.textFrames.add();
            line.duplicate(tempTextFrame, ElementPlacement.INSIDE);
            /* createOutline() は元のテキストフレームを消費するので参照を手放す
               createOutline() consumes the source frame, so drop the reference */
            outlineGroup = tempTextFrame.createOutline();
            tempTextFrame = null;
            width = outlineGroup.width;
        } catch (e) {
            /* アウトライン化できないときはフレーム幅で代用 / Fall back to the frame width */
            try {
                width = (tempTextFrame !== null) ? tempTextFrame.width : 0;
            } catch (err) {
                width = 0;
            }
        }

        /* 残っている一時オブジェクトを後始末する / Clean up whichever temporary object survived */
        try { if (outlineGroup !== null) outlineGroup.remove(); } catch (e) {}
        try { if (tempTextFrame !== null) tempTextFrame.remove(); } catch (e) {}

        return width;
    }

    /**
     * ストーリー内の指定範囲の文字サイズを変倍する（固定行送りの場合は行送りも追従させる）
     * @param {Story} story - 対象ストーリー
     * @param {number} startIndex - 開始文字インデックス
     * @param {number} endIndex - 終了文字インデックス（この位置は含まない）
     * @param {number} ratio - 変倍率
     * @returns {void}
     */
    function scaleCharacterSizes(story, startIndex, endIndex, ratio) {
        for (var i = startIndex; i < endIndex; i++) {
            try {
                var charAttr = story.characters[i].characterAttributes;
                charAttr.size *= ratio;
                /* 固定行送りのときだけ行が重ならないよう行送りも変倍する
                   Scale leading as well, but only when it is fixed */
                if (!charAttr.autoLeading) {
                    charAttr.leading *= ratio;
                }
            } catch (e) {
                /* 設定できない文字はスキップ / Skip characters that reject the change */
            }
        }
    }

    /**
     * 各行の文字サイズを最長行の幅にそろえる
     * 先に全行を測ってから変倍する（変倍で行が再合成されても対象がずれないようにするため）
     * @param {Document} doc - 対象ドキュメント
     * @param {TextFrame} textFrame - 対象テキストフレーム
     * @returns {boolean} 1行でも測定できた場合 true
     */
    function fitLinesToWidestLine(doc, textFrame) {
        var lines = textFrame.lines;
        /** @type {LineMetric[]} */
        var lineMetrics = [];
        var maxWidth = 0;

        /* 1. 全行の外形幅と最大幅を測る（この時点ではテキストを変更しない）
           1. Measure every line and the widest width; the text is not touched yet */
        for (var i = 0; i < lines.length; i++) {
            var line = lines[i];
            var width = isBlankLine(line) ? 0 : measureLineWidth(doc, line);
            lineMetrics.push({ start: line.start, end: line.end, width: width });
            if (width > maxWidth) {
                maxWidth = width;
            }
        }

        if (maxWidth === 0) {
            return false;
        }

        /* 2. 行ごとに変倍する。行オブジェクトではなくストーリー内の文字インデックスで指定して
              変倍による行の再合成の影響を受けないようにする
           2. Scale line by line, addressing characters by story index rather than by line object
              so re-composition during scaling cannot shift the target range */
        var story = textFrame.story;
        for (var i = 0; i < lineMetrics.length; i++) {
            var metric = lineMetrics[i];
            if (metric.width <= 0) continue;

            var ratio = maxWidth / metric.width;
            /* 幅がほぼ同等の行はスキップ / Skip lines that already match the widest one */
            if (Math.abs(ratio - 1) < RATIO_EPSILON) continue;

            scaleCharacterSizes(story, metric.start, metric.end, ratio);
        }

        return true;
    }

    /**
     * テキストフレームの行送りを自動に切り替え、各段落の自動行送り比率を設定する
     * 他のスクリプトからも単体で呼べるよう、対象と比率を引数で受け取る
     * @param {TextFrame} textFrame - 対象テキストフレーム
     * @param {number} autoLeadingAmount - 自動行送りの比率（％）。省略時は100
     * @returns {void}
     */
    function applyAutoLeading(textFrame, autoLeadingAmount) {
        var amount = (typeof autoLeadingAmount === "number" && !isNaN(autoLeadingAmount)) ? autoLeadingAmount : 100;

        /* フレーム全体を自動行送りにする / Switch the whole frame to auto leading */
        textFrame.textRange.characterAttributes.autoLeading = true;

        var paragraphs = textFrame.paragraphs;
        for (var i = 0; i < paragraphs.length; i++) {
            try {
                paragraphs[i].characterAttributes.autoLeading = true;
                paragraphs[i].paragraphAttributes.autoLeadingAmount = amount;
            } catch (e) {
                /* 空段落など設定できないものはスキップ / Skip paragraphs that reject the setting */
            }
        }
    }

    /**
     * 幅そろえと行送りの設定をまとめて適用する（プレビューでも本適用でも共通）
     * @param {Document} doc - 対象ドキュメント
     * @param {TextFrame} textFrame - 対象テキストフレーム
     * @param {{applyAutoLeading: boolean, autoLeadingAmount: number}} settings - ダイアログの設定値
     * @returns {boolean} 1行でも測定できた場合 true
     */
    function applySettings(doc, textFrame, settings) {
        if (!fitLinesToWidestLine(doc, textFrame)) {
            return false;
        }
        if (settings.applyAutoLeading) {
            applyAutoLeading(textFrame, settings.autoLeadingAmount);
        }
        return true;
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 設定ダイアログのUIを構築する（イベントの配線は呼び出し側で行う）
     * @returns {object} ダイアログとコントロール群
     */
    function createSettingsDialog() {
        var dialog = new Window("dialog", getLabel("dialog", "title") + " " + SCRIPT_VERSION);
        dialog.orientation = "column";
        dialog.alignChildren = "fill";
        dialog.opacity = DIALOG_OPACITY;

        var leadingPanel = dialog.add("panel", undefined, getLabel("panel", "leading"));
        setupPanel(leadingPanel, FIELD_SPACING);

        var autoLeadingCheckbox = leadingPanel.add("checkbox", undefined, getLabel("checkbox", "autoLeading"));
        autoLeadingCheckbox.value = APPLY_AUTO_LEADING;
        autoLeadingCheckbox.helpTip = getLabel("tooltip", "autoLeading");

        var amountRow = addFieldRow(
            leadingPanel,
            getLabelWithColon("fieldLabel", "autoLeadingAmount"),
            AUTO_LEADING_AMOUNT + "",
            getLabel("unit", "percent")
        );
        setFieldRowTooltip(amountRow, getLabel("tooltip", "autoLeadingAmount"));
        setFieldRowEnabled(amountRow, autoLeadingCheckbox.value);

        /* 下段：左にプレビュー、右にキャンセル・OK / Footer: preview on the left, Cancel and OK on the right */
        var footerGroup = dialog.add("group");
        footerGroup.orientation = "row";
        footerGroup.alignment = "fill";
        footerGroup.alignChildren = ["fill", "center"];

        var leftGroup = footerGroup.add("group");
        leftGroup.alignment = ["left", "center"];
        var previewCheckbox = leftGroup.add("checkbox", undefined, getLabel("checkbox", "preview"));
        previewCheckbox.value = PREVIEW_ON_OPEN;
        previewCheckbox.helpTip = getLabel("tooltip", "preview");

        /* 左右のボタンを両端に押し広げるスペーサー / spacer that pushes both sides apart */
        var spacerGroup = footerGroup.add("group");
        spacerGroup.alignment = ["fill", "center"];

        var rightGroup = footerGroup.add("group");
        rightGroup.alignment = ["right", "center"];
        var cancelButton = rightGroup.add("button", undefined, getLabel("button", "cancel"), { name: "cancel" });
        var okButton = rightGroup.add("button", undefined, getLabel("button", "ok"), { name: "ok" });
        cancelButton.preferredSize.width = BUTTON_WIDTH;
        okButton.preferredSize.width = BUTTON_WIDTH;

        return {
            dialog: dialog,
            autoLeadingCheckbox: autoLeadingCheckbox,
            amountRow: amountRow,
            previewCheckbox: previewCheckbox,
            cancelButton: cancelButton,
            okButton: okButton
        };
    }

    /**
     * 設定ダイアログを表示し、プレビューを見ながら結果を確定する
     * OKのときはプレビュー分を取り消してから1回だけ適用するので、Undoは1回で戻る
     * @param {Document} doc - 対象ドキュメント
     * @param {TextFrame} textFrame - 対象テキストフレーム
     * @returns {void}
     */
    function showSettingsDialog(doc, textFrame) {
        var ui = createSettingsDialog();
        var previewManager = new PreviewManager();

        /**
         * ダイアログの現在の入力値を読み取る
         * @returns {{applyAutoLeading: boolean, autoLeadingAmount: number}} 設定値
         */
        function readSettings() {
            var amount = readNumber(ui.amountRow.input, AUTO_LEADING_AMOUNT);
            /* 0以下や数値でない入力は既定値に読み替える / fall back to the default for non-positive or invalid input */
            if (!(amount > 0)) amount = AUTO_LEADING_AMOUNT;
            return {
                applyAutoLeading: ui.autoLeadingCheckbox.value,
                autoLeadingAmount: amount
            };
        }

        /**
         * プレビューを最新の設定で貼り直す（Undo履歴は汚さない）
         * @returns {void}
         */
        function updatePreview() {
            previewManager.rollback();
            if (!ui.previewCheckbox.value) return;
            previewManager.addStep(function () {
                applySettings(doc, textFrame, readSettings());
            });
        }

        ui.previewCheckbox.onClick = updatePreview;

        ui.autoLeadingCheckbox.onClick = function () {
            setFieldRowEnabled(ui.amountRow, ui.autoLeadingCheckbox.value);
            updatePreview();
        };

        ui.amountRow.input.onChange = updatePreview;
        changeValueByArrowKey(ui.amountRow.input, { step: 1, shiftStep: 10, altStep: 0.1 });

        ui.okButton.onClick = function () {
            /* プレビュー分を戻し、本適用を1回だけ実行して確定（Undo履歴は1つ）
               undo the preview, then apply once so it lands as a single undo entry */
            previewManager.rollback();
            var applied = applySettings(doc, textFrame, readSettings());
            ui.dialog.close(1);
            if (!applied) {
                alert(getLabel("alert", "noMeasurableLine"));
                return;
            }
            app.redraw();
        };

        ui.cancelButton.onClick = function () {
            /* 開いてから適用した分をすべて取り消してから閉じる / undo everything applied since open, then close */
            previewManager.rollback();
            ui.dialog.close(2);
        };

        updatePreview();
        ui.dialog.show();
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * メイン処理
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert", "noDocument"));
            return;
        }

        var doc = app.activeDocument;
        var selectionItems = doc.selection;

        /* 選択オブジェクトのチェック / Validate the selection */
        if (!selectionItems || selectionItems.length !== 1 || selectionItems[0].typename !== "TextFrame") {
            alert(getLabel("alert", "selectOneTextFrame"));
            return;
        }

        var targetTextFrame = selectionItems[0];
        if (targetTextFrame.lines.length <= 1) {
            alert(getLabel("alert", "needTwoLines"));
            return;
        }

        showSettingsDialog(doc, targetTextFrame);
    }

    main();

})();
