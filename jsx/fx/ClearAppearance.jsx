#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したオブジェクトに「アピアランスの消去」を実行します。
実行前にダイアログを表示し、復元するかどうかと、選択オブジェクトに応じた復元方法を選べます。

詳細は README を参照してください。

### Overview

Runs Clear Appearance on the selected objects.
A dialog first asks whether to restore anything, and which restore mode fits the selection.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ClearAppearance";              /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-04-14";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-19";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ClearAppearance.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ClearAppearance.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/na4c70c5acd60"; /* 紹介記事 / article URL */

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
            noDocument: { ja: "ドキュメントが開かれていません", en: "No document is open." },
            noSelection: { ja: "オブジェクトを選択してください", en: "Please select at least one object." },
            pathFailures: { ja: "パス処理の失敗", en: "Path processing failures" },
            textFailures: { ja: "テキスト処理の失敗", en: "Text processing failures" },
            actionFailures: { ja: "アクション実行の失敗", en: "Action execution failures" },
            selectionRestoreFailures: { ja: "選択の復元失敗", en: "Selection restoration failures" },
            details: { ja: "詳細：", en: "Details:" }
        },
        detailCategory: {
            path: { ja: "パス", en: "Path" },
            text: { ja: "テキスト", en: "Text" },
            action: { ja: "アクション", en: "Action" },
            actionUnload: { ja: "アクション解除", en: "Action unload" },
            selectionRestore: { ja: "選択復元", en: "Selection restore" }
        },
        dialog: {
            title: { ja: "アピアランスの消去", en: "Clear Appearance" }
        },
        panel: {
            clearBehavior: { ja: "消去後の挙動", en: "After Clearing" },
            pathRestore: { ja: "パス・複合パス", en: "Paths and Compound Paths" },
            textRestore: { ja: "テキスト", en: "Text" },
            attributeOption: { ja: "オブジェクト属性", en: "Object Attributes" }
        },
        radio: {
            restore: { ja: "属性を復元", en: "Restore attributes" },
            noRestore: { ja: "消去したまま", en: "Leave cleared" },
            pathColorOnly: { ja: "塗り・線・線幅", en: "Fill, Stroke and Weight" },
            pathColorAndStrokeSettings: { ja: "塗り・線＋線の詳細設定", en: "Fill, Stroke + Stroke Details" },
            textNoRestore: { ja: "塗りを復元しない", en: "Do not restore fill" },
            textFillFirst: { ja: "テキストの塗り（1文字目）", en: "Text Fill (First Character)" },
            textFillPerChar: { ja: "テキストの塗り（文字単位）", en: "Text Fill (Per Character)" }
        },
        checkbox: {
            opacity: { ja: "不透明度", en: "Opacity" },
            blendingMode: { ja: "描画モード", en: "Blending Mode" },
            overprint: { ja: "オーバープリント", en: "Overprint" }
        },
        button: {
            ok: { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        tooltip: {
            restore: {
                ja: "アピアランスを消去したあと、控えておいた塗り・線・不透明度などを戻します。",
                en: "Clears the appearance, then puts the captured fill, stroke and opacity back."
            },
            noRestore: {
                ja: "アピアランスを消去したままにします。塗りも線もなくなります。",
                en: "Leaves the appearance cleared. Both fill and stroke are removed."
            },
            pathColorOnly: {
                ja: "塗り色・線色・線幅だけを戻します。線端・角の形状・破線は戻りません。",
                en: "Restores fill color, stroke color and stroke weight only. Caps, joins and dashes are not restored."
            },
            pathColorAndStrokeSettings: {
                ja: "塗り色・線色・線幅に加えて、線端・角の形状・破線・破線オフセット・角の比率も戻します。",
                en: "Also restores caps, joins, dashes, dash offset and miter limit."
            },
            textNoRestore: {
                ja: "テキストの塗りは戻しません。消去後のままになります。",
                en: "Leaves the text fill cleared."
            },
            textFillFirst: {
                ja: "1文字目の色をテキスト全体に適用します。文字ごとの色分けは失われます。",
                en: "Applies the first character's color to the whole text. Per-character colors are lost."
            },
            textFillPerChar: {
                ja: "文字ごとの色をそのまま戻します。文字数が多いテキストでは時間がかかります。",
                en: "Restores each character's own color. Slow on long text."
            },
            opacity: {
                ja: "オブジェクトの不透明度を消去前の値に戻します。",
                en: "Restores the object's opacity to the value it had before clearing."
            },
            blendingMode: {
                ja: "オブジェクトの描画モードを消去前の設定に戻します。",
                en: "Restores the object's blending mode to the setting it had before clearing."
            },
            overprint: {
                ja: "塗りと線のオーバープリント設定を消去前の状態に戻します。",
                en: "Restores the fill and stroke overprint settings."
            }
        }
    };

    /**
     * ラベル定義から現在の言語の文言を取得する
     * @param {object} labelEntry - ja / en を持つラベル定義
     * @returns {string} 現在の言語の文言
     */
    function getLabel(labelEntry) {
        return labelEntry[uiLang] || labelEntry.en;
    }

    /**
     * 失敗の詳細行に使うカテゴリ名を取得する
     * @param {string} categoryKey - LABELS.detailCategory のキー
     * @returns {string} カテゴリ名（未定義のキーはそのまま返す）
     */
    function getDetailCategoryLabel(categoryKey) {
        var categoryEntry = LABELS.detailCategory[categoryKey];
        return categoryEntry ? getLabel(categoryEntry) : String(categoryKey);
    }

    // =========================================
    // レイアウト
    // Layout
    // =========================================

    var WINDOW_MARGINS        = [15, 20, 15, 15];  /* ウィンドウ余白 [左,上,右,下] / window margins */
    var WINDOW_SPACING        = 10;                /* ウィンドウ内の要素間隔 / window spacing */
    var PANEL_MARGINS         = [15, 20, 15, 10];  /* パネル余白 [左,上,右,下] / panel margins */
    var PANEL_SPACING         = 8;                 /* パネル内の要素間隔 / panel spacing */
    var BUTTON_SPACING        = 10;                /* ボタンの間隔 / spacing between buttons */
    var BUTTON_ROW_TOP_MARGIN = 8;                 /* ボタン行の上余白 / top margin of the button row */

    // =========================================
    // 失敗の集計
    // Failure logging
    // =========================================

    /* 詳細行として残す失敗の上限 / Maximum number of failure detail lines */
    var MAX_FAILURE_DETAILS = 8;

    /* 詳細カテゴリと集計先カウンターの対応 / Mapping from detail category to counter */
    var FAILURE_COUNTER_KEYS = {
        path: "pathFailureCount",
        text: "textFailureCount",
        action: "actionFailureCount",
        actionUnload: "actionFailureCount",
        selectionRestore: "selectionRestoreFailureCount"
    };

    /* アクション実行の失敗に付ける印（二重計上を防ぐ）/ Flag marking a failure already counted by the action runner */
    var ACTION_FAILURE_FLAG = "isClearAppearanceActionFailure";

    /**
     * @typedef {object} FailureLog
     * @property {number} pathFailureCount - パス処理の失敗件数
     * @property {number} textFailureCount - テキスト処理の失敗件数
     * @property {number} actionFailureCount - アクション実行・解除の失敗件数
     * @property {number} selectionRestoreFailureCount - 選択復元の失敗件数
     * @property {Array<string>} details - 失敗の詳細行
     */

    /**
     * 失敗の集計オブジェクトを作る
     * @returns {FailureLog} 初期化した集計オブジェクト
     */
    function createFailureLog() {
        return {
            pathFailureCount: 0,
            textFailureCount: 0,
            actionFailureCount: 0,
            selectionRestoreFailureCount: 0,
            details: []
        };
    }

    /**
     * 失敗したオブジェクトを詳細行用の文字列にする
     * @param {object} item - 対象オブジェクト（null 可）
     * @returns {string} 種別名と名前を並べた文字列
     */
    function describeFailureItem(item) {
        try {
            if (!item) {
                return "Unknown";
            }
            if (!item.typename) {
                return "Unknown";
            }
            var description = String(item.typename);
            if (item.name) {
                description += " [" + item.name + "]";
            }
            return description;
        } catch (e) {
            return "Unknown";
        }
    }

    /**
     * 例外を詳細行用の文字列にする
     * @param {object} error - 捕捉した例外
     * @returns {string} エラーメッセージ
     */
    function describeFailureError(error) {
        if (!error) {
            return "Unknown error";
        }
        return String(error.message || error);
    }

    /**
     * 失敗の詳細行を追加する（上限に達したら何もしない）
     * @param {FailureLog} failureLog - 集計オブジェクト
     * @param {string} categoryKey - LABELS.detailCategory のキー
     * @param {object} item - 対象オブジェクト（null 可）
     * @param {object} error - 捕捉した例外
     * @returns {void}
     */
    function addFailureDetail(failureLog, categoryKey, item, error) {
        if (!failureLog || failureLog.details.length >= MAX_FAILURE_DETAILS) {
            return;
        }
        failureLog.details.push(
            getDetailCategoryLabel(categoryKey) + ": " +
            describeFailureItem(item) + " - " +
            describeFailureError(error)
        );
    }

    /**
     * 失敗を件数と詳細行の両方に記録する
     * @param {FailureLog} failureLog - 集計オブジェクト
     * @param {string} categoryKey - LABELS.detailCategory のキー
     * @param {object} item - 対象オブジェクト（null 可）
     * @param {object} error - 捕捉した例外
     * @returns {void}
     */
    function recordFailure(failureLog, categoryKey, item, error) {
        if (!failureLog) {
            return;
        }
        failureLog[FAILURE_COUNTER_KEYS[categoryKey]]++;
        addFailureDetail(failureLog, categoryKey, item, error);
    }

    /**
     * オブジェクト単位の失敗を記録する（アクション実行で計上済みのものは重ねない）
     * @param {FailureLog} failureLog - 集計オブジェクト
     * @param {string} categoryKey - LABELS.detailCategory のキー
     * @param {object} item - 対象オブジェクト（null 可）
     * @param {object} error - 捕捉した例外
     * @returns {void}
     */
    function recordItemFailure(failureLog, categoryKey, item, error) {
        if (isActionFailure(error)) {
            return;
        }
        recordFailure(failureLog, categoryKey, item, error);
    }

    /**
     * 例外にアクション実行の失敗であることを示す印を付ける
     * @param {object} error - 捕捉した例外
     * @returns {object} 印を付けた例外
     */
    function markActionFailure(error) {
        if (error) {
            error[ACTION_FAILURE_FLAG] = true;
        }
        return error;
    }

    /**
     * 例外がアクション実行の失敗かどうかを判定する
     * @param {object} error - 捕捉した例外
     * @returns {boolean} アクション実行の失敗なら true
     */
    function isActionFailure(error) {
        return !!(error && error[ACTION_FAILURE_FLAG]);
    }

    /**
     * 失敗の集計を報告用のメッセージに組み立てる
     * @param {FailureLog} failureLog - 集計オブジェクト
     * @returns {string} 報告メッセージ（失敗がなければ空文字）
     */
    function buildFailureMessage(failureLog) {
        var summaryRows = [
            { count: failureLog.pathFailureCount, labelEntry: LABELS.alert.pathFailures },
            { count: failureLog.textFailureCount, labelEntry: LABELS.alert.textFailures },
            { count: failureLog.actionFailureCount, labelEntry: LABELS.alert.actionFailures },
            { count: failureLog.selectionRestoreFailureCount, labelEntry: LABELS.alert.selectionRestoreFailures }
        ];

        var messageLines = [];
        for (var i = 0; i < summaryRows.length; i++) {
            if (summaryRows[i].count > 0) {
                messageLines.push(getLabel(summaryRows[i].labelEntry) + ": " + summaryRows[i].count);
            }
        }

        if (failureLog.details.length > 0) {
            messageLines.push("");
            messageLines.push(getLabel(LABELS.alert.details));
            for (var j = 0; j < failureLog.details.length; j++) {
                messageLines.push("- " + failureLog.details[j]);
            }
        }

        return messageLines.join("\n");
    }

    /**
     * 失敗があればまとめて通知する
     * @param {FailureLog} failureLog - 集計オブジェクト
     * @returns {void}
     */
    function alertFailures(failureLog) {
        var message = buildFailureMessage(failureLog);
        if (message !== "") {
            alert(message);
        }
    }

    // =========================================
    // 選択内容の判定
    // Selection inspection
    // =========================================

    /**
     * @typedef {object} RestoreAvailability
     * @property {boolean} hasPathObjects - 選択にパス・複合パスが含まれるか
     * @property {boolean} hasTextFrames - 選択にテキストが含まれるか
     */

    /**
     * 選択を再帰的に走査して、復元対象の種別が含まれるかを調べる
     * @param {Array} items - 走査するオブジェクト
     * @param {RestoreAvailability} availability - 結果の書き込み先
     * @returns {void}
     */
    function collectRestoreAvailability(items, availability) {
        if (!items) {
            return;
        }

        for (var i = 0; i < items.length; i++) {
            var item = items[i];
            if (!item) {
                continue;
            }

            switch (item.typename) {
                case "GroupItem":
                    /* クリップグループは中身を個別に扱わないので走査しない / Clipped groups are handled as a whole */
                    if (!item.clipped) {
                        collectRestoreAvailability(item.pageItems, availability);
                    }
                    break;

                case "PathItem":
                case "CompoundPathItem":
                    availability.hasPathObjects = true;
                    break;

                case "TextFrame":
                    availability.hasTextFrames = true;
                    break;
            }
        }
    }

    /**
     * 選択に含まれる復元対象の種別を調べる
     * @param {Array} items - 選択オブジェクト
     * @returns {RestoreAvailability} 復元対象の有無
     */
    function getRestoreAvailability(items) {
        var availability = {
            hasPathObjects: false,
            hasTextFrames: false
        };
        collectRestoreAvailability(items, availability);
        return availability;
    }

    /**
     * 選択オブジェクトを配列に控える（処理中に選択が変わっても参照を保てるようにする）
     * @param {object} selection - app.selection
     * @returns {Array} オブジェクトの配列
     */
    function toItemArray(selection) {
        var items = [];
        for (var i = 0; i < selection.length; i++) {
            items.push(selection[i]);
        }
        return items;
    }

    // =========================================
    // ダイアログ
    // Dialog
    // =========================================

    /**
     * @typedef {object} RestoreOptions
     * @property {boolean} fillStroke - パスの塗り・線を復元するか
     * @property {boolean} strokeSettings - 線端・破線などの線の設定も復元するか
     * @property {boolean} textFillFirst - テキストの塗りを1文字目の色で復元するか
     * @property {boolean} textFillPerChar - テキストの塗りを文字単位で復元するか
     * @property {boolean} opacity - 不透明度を復元するか
     * @property {boolean} blendingMode - 描画モードを復元するか
     * @property {boolean} overprint - オーバープリントを復元するか
     */

    /**
     * ダイアログの共通レイアウトを設定する
     * @param {Window} targetWindow - 対象のウィンドウ
     * @returns {void}
     */
    function setupDialogWindow(targetWindow) {
        targetWindow.orientation = "column";
        targetWindow.alignChildren = ["fill", "top"];
        targetWindow.margins = WINDOW_MARGINS;
        targetWindow.spacing = WINDOW_SPACING;
    }

    /**
     * パネルを追加し、共通レイアウトを設定する
     * @param {Window} parentWindow - 追加先のウィンドウ
     * @param {object} titleEntry - ja / en を持つパネル名
     * @returns {Panel} 追加したパネル
     */
    function addSettingPanel(parentWindow, titleEntry) {
        var settingPanel = parentWindow.add("panel", undefined, getLabel(titleEntry));
        settingPanel.orientation = "column";
        settingPanel.alignChildren = ["left", "top"];
        settingPanel.alignment = ["fill", "top"];
        settingPanel.margins = PANEL_MARGINS;
        settingPanel.spacing = PANEL_SPACING;
        return settingPanel;
    }

    /**
     * コントロール群の有効・無効をまとめて切り替える
     * @param {object} controlSet - パネルとコントロールを持つオブジェクト
     * @param {boolean} enabled - 有効にするか
     * @returns {void}
     */
    function setControlSetEnabled(controlSet, enabled) {
        for (var controlKey in controlSet) {
            if (controlSet[controlKey]) {
                controlSet[controlKey].enabled = enabled;
            }
        }
    }

    /**
     * 消去後の挙動を選ぶパネルを追加する
     * @param {Window} parentWindow - 追加先のウィンドウ
     * @returns {object} パネルとラジオを持つコントロール群
     */
    function addClearBehaviorPanel(parentWindow) {
        var clearBehaviorPanel = addSettingPanel(parentWindow, LABELS.panel.clearBehavior);

        var behaviorRowGroup = clearBehaviorPanel.add("group");
        behaviorRowGroup.orientation = "row";
        behaviorRowGroup.alignment = ["left", "top"];
        behaviorRowGroup.alignChildren = ["left", "center"];

        var restoreRadio = behaviorRowGroup.add("radiobutton", undefined, getLabel(LABELS.radio.restore));
        var noRestoreRadio = behaviorRowGroup.add("radiobutton", undefined, getLabel(LABELS.radio.noRestore));
        restoreRadio.helpTip = getLabel(LABELS.tooltip.restore);
        noRestoreRadio.helpTip = getLabel(LABELS.tooltip.noRestore);
        restoreRadio.value = true;

        return {
            panel: clearBehaviorPanel,
            restoreRadio: restoreRadio,
            noRestoreRadio: noRestoreRadio
        };
    }

    /**
     * パスオブジェクトの復元方法パネルを追加する
     * @param {Window} parentWindow - 追加先のウィンドウ
     * @returns {object} パネルとラジオを持つコントロール群
     */
    function addPathRestorePanel(parentWindow) {
        var pathRestorePanel = addSettingPanel(parentWindow, LABELS.panel.pathRestore);

        var pathColorOnlyRadio = pathRestorePanel.add("radiobutton", undefined, getLabel(LABELS.radio.pathColorOnly));
        var pathColorAndStrokeRadio = pathRestorePanel.add("radiobutton", undefined, getLabel(LABELS.radio.pathColorAndStrokeSettings));
        pathColorOnlyRadio.helpTip = getLabel(LABELS.tooltip.pathColorOnly);
        pathColorAndStrokeRadio.helpTip = getLabel(LABELS.tooltip.pathColorAndStrokeSettings);
        pathColorAndStrokeRadio.value = true;

        return {
            panel: pathRestorePanel,
            pathColorOnlyRadio: pathColorOnlyRadio,
            pathColorAndStrokeRadio: pathColorAndStrokeRadio
        };
    }

    /**
     * テキストの復元方法パネルを追加する
     * @param {Window} parentWindow - 追加先のウィンドウ
     * @returns {object} パネルとラジオを持つコントロール群
     */
    function addTextRestorePanel(parentWindow) {
        var textRestorePanel = addSettingPanel(parentWindow, LABELS.panel.textRestore);

        var textNoRestoreRadio = textRestorePanel.add("radiobutton", undefined, getLabel(LABELS.radio.textNoRestore));
        var textFillFirstRadio = textRestorePanel.add("radiobutton", undefined, getLabel(LABELS.radio.textFillFirst));
        var textFillPerCharRadio = textRestorePanel.add("radiobutton", undefined, getLabel(LABELS.radio.textFillPerChar));
        textNoRestoreRadio.helpTip = getLabel(LABELS.tooltip.textNoRestore);
        textFillFirstRadio.helpTip = getLabel(LABELS.tooltip.textFillFirst);
        textFillPerCharRadio.helpTip = getLabel(LABELS.tooltip.textFillPerChar);
        textFillPerCharRadio.value = true;

        return {
            panel: textRestorePanel,
            textNoRestoreRadio: textNoRestoreRadio,
            textFillFirstRadio: textFillFirstRadio,
            textFillPerCharRadio: textFillPerCharRadio
        };
    }

    /**
     * オブジェクト属性の復元オプションパネルを追加する
     * @param {Window} parentWindow - 追加先のウィンドウ
     * @returns {object} パネルとチェックボックスを持つコントロール群
     */
    function addAttributeOptionPanel(parentWindow) {
        var attributeOptionPanel = addSettingPanel(parentWindow, LABELS.panel.attributeOption);

        var opacityCheckbox = attributeOptionPanel.add("checkbox", undefined, getLabel(LABELS.checkbox.opacity));
        var blendingModeCheckbox = attributeOptionPanel.add("checkbox", undefined, getLabel(LABELS.checkbox.blendingMode));
        var overprintCheckbox = attributeOptionPanel.add("checkbox", undefined, getLabel(LABELS.checkbox.overprint));
        opacityCheckbox.helpTip = getLabel(LABELS.tooltip.opacity);
        blendingModeCheckbox.helpTip = getLabel(LABELS.tooltip.blendingMode);
        overprintCheckbox.helpTip = getLabel(LABELS.tooltip.overprint);
        opacityCheckbox.value = true;
        blendingModeCheckbox.value = true;
        overprintCheckbox.value = true;

        return {
            panel: attributeOptionPanel,
            opacityCheckbox: opacityCheckbox,
            blendingModeCheckbox: blendingModeCheckbox,
            overprintCheckbox: overprintCheckbox
        };
    }

    /**
     * ボタン行を追加する
     * @param {Window} parentWindow - 追加先のウィンドウ
     * @returns {void}
     */
    function addDialogButtonRow(parentWindow) {
        var btnRowGroup = parentWindow.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.margins = [0, BUTTON_ROW_TOP_MARGIN, 0, 0];
        btnRowGroup.alignment = ["center", "bottom"];
        btnRowGroup.alignChildren = ["center", "center"];
        btnRowGroup.spacing = BUTTON_SPACING;
        btnRowGroup.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });
        btnRowGroup.add("button", undefined, getLabel(LABELS.button.ok), { name: "ok" });
    }

    /**
     * ダイアログの選択状態から復元オプションを組み立てる
     * @param {RestoreAvailability} availability - 復元対象の有無
     * @param {object} clearBehaviorControls - 消去後の挙動のコントロール群
     * @param {object} pathRestoreControls - パスの復元方法のコントロール群
     * @param {object} textRestoreControls - テキストの復元方法のコントロール群
     * @param {object} attributeOptionControls - オブジェクト属性のコントロール群
     * @returns {RestoreOptions} 復元オプション
     */
    function buildRestoreOptions(availability, clearBehaviorControls, pathRestoreControls, textRestoreControls, attributeOptionControls) {
        if (!clearBehaviorControls.restoreRadio.value) {
            return {
                fillStroke: false,
                strokeSettings: false,
                textFillFirst: false,
                textFillPerChar: false,
                opacity: false,
                blendingMode: false,
                overprint: false
            };
        }

        return {
            fillStroke: availability.hasPathObjects,
            strokeSettings: availability.hasPathObjects && pathRestoreControls.pathColorAndStrokeRadio.value,
            textFillFirst: availability.hasTextFrames && textRestoreControls.textFillFirstRadio.value,
            textFillPerChar: availability.hasTextFrames && textRestoreControls.textFillPerCharRadio.value,
            opacity: attributeOptionControls.opacityCheckbox.value,
            blendingMode: attributeOptionControls.blendingModeCheckbox.value,
            overprint: attributeOptionControls.overprintCheckbox.value
        };
    }

    /**
     * 復元オプションダイアログを表示する
     * @param {RestoreAvailability} availability - 復元対象の有無
     * @returns {RestoreOptions} 復元オプション（キャンセル時は null）
     */
    function showRestoreDialog(availability) {
        var restoreDialog = new Window("dialog", getLabel(LABELS.dialog.title) + "  " + SCRIPT_VERSION);
        setupDialogWindow(restoreDialog);

        var clearBehaviorControls = addClearBehaviorPanel(restoreDialog);
        var pathRestoreControls = addPathRestorePanel(restoreDialog);
        var textRestoreControls = addTextRestorePanel(restoreDialog);
        var attributeOptionControls = addAttributeOptionPanel(restoreDialog);

        /**
         * 復元しないときと、対象が選択にないときはパネルを無効にする
         * @returns {void}
         */
        function updateEnabledState() {
            var restoreEnabled = clearBehaviorControls.restoreRadio.value;
            setControlSetEnabled(pathRestoreControls, restoreEnabled && availability.hasPathObjects);
            setControlSetEnabled(textRestoreControls, restoreEnabled && availability.hasTextFrames);
            setControlSetEnabled(attributeOptionControls, restoreEnabled);
        }

        clearBehaviorControls.restoreRadio.onClick = updateEnabledState;
        clearBehaviorControls.noRestoreRadio.onClick = updateEnabledState;
        updateEnabledState();

        addDialogButtonRow(restoreDialog);

        if (restoreDialog.show() !== 1) {
            return null;
        }

        return buildRestoreOptions(availability, clearBehaviorControls, pathRestoreControls, textRestoreControls, attributeOptionControls);
    }

    // =========================================
    // 色処理
    // Color handling
    // =========================================

    /**
     * 色を複製する（消去後に同じ色を再適用できるようにする）
     * @param {object} sourceColor - 複製元の色
     * @returns {object} 複製した色（未対応の種別は null）
     */
    function cloneColor(sourceColor) {
        if (!sourceColor) {
            return null;
        }

        switch (sourceColor.typename) {

            case "RGBColor":
                var rgbColor = new RGBColor();
                rgbColor.red = sourceColor.red;
                rgbColor.green = sourceColor.green;
                rgbColor.blue = sourceColor.blue;
                return rgbColor;

            case "CMYKColor":
                var cmykColor = new CMYKColor();
                cmykColor.cyan = sourceColor.cyan;
                cmykColor.magenta = sourceColor.magenta;
                cmykColor.yellow = sourceColor.yellow;
                cmykColor.black = sourceColor.black;
                return cmykColor;

            case "GrayColor":
                var grayColor = new GrayColor();
                grayColor.gray = sourceColor.gray;
                return grayColor;

            case "SpotColor":
                var spotColor = new SpotColor();
                spotColor.spot = sourceColor.spot;
                spotColor.tint = sourceColor.tint;
                return spotColor;

            case "GradientColor":
                var gradientColor = new GradientColor();
                gradientColor.gradient = sourceColor.gradient;
                gradientColor.angle = sourceColor.angle;
                gradientColor.length = sourceColor.length;
                gradientColor.origin = sourceColor.origin;
                gradientColor.matrix = sourceColor.matrix;
                /* ハイライトは円形グラデーションのみ有効 / Hilite applies to radial gradients only */
                try {
                    gradientColor.hiliteAngle = sourceColor.hiliteAngle;
                    gradientColor.hiliteLength = sourceColor.hiliteLength;
                } catch (e) { }
                return gradientColor;

            case "PatternColor":
                var patternColor = new PatternColor();
                patternColor.pattern = sourceColor.pattern;
                try {
                    patternColor.matrix = sourceColor.matrix;
                } catch (err) { }
                return patternColor;

            case "NoColor":
                return new NoColor();

            default:
                return null;
        }
    }

    /**
     * テキストの塗りとして控える値があるかを判定する
     * @param {object} fillColor - 判定する色
     * @returns {boolean} 塗りとして扱える色なら true
     */
    function isUsableTextColor(fillColor) {
        return !!(fillColor && fillColor.typename && fillColor.typename !== "NoColor");
    }

    /**
     * テキスト範囲の塗り色を複製して控える
     * @param {TextRange} textRange - 対象のテキスト範囲
     * @returns {object} 複製した塗り色（取得できないときは null）
     */
    function captureTextRangeFill(textRange) {
        var fillColor = null;
        try {
            fillColor = textRange.characterAttributes.fillColor;
        } catch (e) {
            return null;
        }
        return isUsableTextColor(fillColor) ? cloneColor(fillColor) : null;
    }

    /**
     * テキスト範囲に塗りだけを戻す（線はなしにする）
     * @param {TextRange} textRange - 対象のテキスト範囲
     * @param {object} fillColor - 戻す塗り色（null なら塗りなし）
     * @returns {void}
     */
    function applyTextRangeFill(textRange, fillColor) {
        var characterAttributes = textRange.characterAttributes;
        characterAttributes.fillColor = fillColor ? cloneColor(fillColor) : new NoColor();
        characterAttributes.strokeColor = new NoColor();
    }

    // =========================================
    // 属性の控えと再適用
    // Capturing and reapplying attributes
    // =========================================

    /**
     * @typedef {object} PathStyle
     * @property {boolean} filled - 塗りがあったか
     * @property {boolean} stroked - 線があったか
     * @property {object} fillColor - 控えた塗り色
     * @property {object} strokeColor - 控えた線色
     * @property {number} strokeWidth - 控えた線幅
     * @property {boolean} fillOverprint - 控えた塗りのオーバープリント
     * @property {boolean} strokeOverprint - 控えた線のオーバープリント
     * @property {StrokeCap} strokeCap - 控えた線端
     * @property {StrokeJoin} strokeJoin - 控えた角の形状
     * @property {Array<number>} strokeDashes - 控えた破線パターン
     * @property {number} strokeDashOffset - 控えた破線オフセット
     * @property {number} strokeMiterLimit - 控えた角の比率
     */

    /**
     * @typedef {object} ObjectAttributes
     * @property {number} opacity - 控えた不透明度
     * @property {BlendModes} blendingMode - 控えた描画モード
     */

    /**
     * 塗り・線のプロパティを持つオブジェクトを返す
     * 複合パスは構成パスの1本目が塗り・線を保持する
     * @param {PageItem} item - パスまたは複合パス
     * @returns {PathItem} 塗り・線を保持するパス（見つからないときは null）
     */
    function resolvePathStyleHost(item) {
        if (item.typename !== "CompoundPathItem") {
            return item;
        }
        try {
            return item.pathItems.length > 0 ? item.pathItems[0] : null;
        } catch (e) {
            return null;
        }
    }

    /**
     * 塗り・線の設定を控える
     * @param {PathItem} styleHost - 塗り・線を保持するパス
     * @param {RestoreOptions} restoreOptions - 復元オプション
     * @returns {PathStyle} 控えた設定
     */
    function capturePathStyle(styleHost, restoreOptions) {
        var pathStyle = {
            filled: styleHost.filled,
            stroked: styleHost.stroked,
            fillColor: null,
            strokeColor: null,
            strokeWidth: 1,
            fillOverprint: null,
            strokeOverprint: null,
            strokeCap: null,
            strokeJoin: null,
            strokeDashes: null,
            strokeDashOffset: null,
            strokeMiterLimit: null
        };

        if (pathStyle.filled) {
            pathStyle.fillColor = cloneColor(styleHost.fillColor);
            if (restoreOptions.overprint) {
                pathStyle.fillOverprint = styleHost.fillOverprint;
            }
        }

        if (pathStyle.stroked) {
            pathStyle.strokeColor = cloneColor(styleHost.strokeColor);
            pathStyle.strokeWidth = styleHost.strokeWidth;
            if (restoreOptions.overprint) {
                pathStyle.strokeOverprint = styleHost.strokeOverprint;
            }
            if (restoreOptions.strokeSettings) {
                pathStyle.strokeCap = styleHost.strokeCap;
                pathStyle.strokeJoin = styleHost.strokeJoin;
                pathStyle.strokeDashes = styleHost.strokeDashes ? styleHost.strokeDashes.slice(0) : null;
                pathStyle.strokeDashOffset = styleHost.strokeDashOffset;
                pathStyle.strokeMiterLimit = styleHost.strokeMiterLimit;
            }
        }

        return pathStyle;
    }

    /**
     * 控えた塗り・線の設定を再適用する
     * 控えていない項目は null のままなので、そのまま読み飛ばす
     * @param {PathItem} styleHost - 塗り・線を保持するパス
     * @param {PathStyle} pathStyle - 控えた設定
     * @returns {void}
     */
    function applyPathStyle(styleHost, pathStyle) {
        if (pathStyle.filled && pathStyle.fillColor) {
            styleHost.filled = true;
            styleHost.fillColor = pathStyle.fillColor;
            if (pathStyle.fillOverprint !== null) {
                styleHost.fillOverprint = pathStyle.fillOverprint;
            }
        } else {
            styleHost.filled = false;
            styleHost.fillColor = new NoColor();
        }

        if (!pathStyle.stroked || !pathStyle.strokeColor) {
            styleHost.stroked = false;
            styleHost.strokeColor = new NoColor();
            return;
        }

        styleHost.stroked = true;
        styleHost.strokeColor = pathStyle.strokeColor;
        styleHost.strokeWidth = pathStyle.strokeWidth;
        if (pathStyle.strokeOverprint !== null) {
            styleHost.strokeOverprint = pathStyle.strokeOverprint;
        }
        if (pathStyle.strokeCap !== null) {
            styleHost.strokeCap = pathStyle.strokeCap;
        }
        if (pathStyle.strokeJoin !== null) {
            styleHost.strokeJoin = pathStyle.strokeJoin;
        }
        if (pathStyle.strokeDashes !== null) {
            styleHost.strokeDashes = pathStyle.strokeDashes;
        }
        if (pathStyle.strokeDashOffset !== null) {
            styleHost.strokeDashOffset = pathStyle.strokeDashOffset;
        }
        if (pathStyle.strokeMiterLimit !== null) {
            styleHost.strokeMiterLimit = pathStyle.strokeMiterLimit;
        }
    }

    /**
     * オブジェクト単位の属性（不透明度・描画モード）を控える
     * @param {PageItem} item - 対象オブジェクト
     * @param {RestoreOptions} restoreOptions - 復元オプション
     * @returns {ObjectAttributes} 控えた属性
     */
    function captureObjectAttributes(item, restoreOptions) {
        return {
            opacity: restoreOptions.opacity ? item.opacity : null,
            blendingMode: restoreOptions.blendingMode ? item.blendingMode : null
        };
    }

    /**
     * 控えたオブジェクト属性を再適用する
     * @param {PageItem} item - 対象オブジェクト
     * @param {ObjectAttributes} objectAttributes - 控えた属性
     * @returns {void}
     */
    function applyObjectAttributes(item, objectAttributes) {
        if (objectAttributes.opacity !== null) {
            item.opacity = objectAttributes.opacity;
        }
        if (objectAttributes.blendingMode !== null) {
            item.blendingMode = objectAttributes.blendingMode;
        }
    }

    /**
     * @typedef {object} TextFillState
     * @property {number} characterCount - 控えた時点の文字数
     * @property {Array} characterFills - 文字ごとの塗り色
     * @property {object} firstCharacterFill - 1文字目の塗り色
     * @property {object} rangeFill - テキスト全体の塗り色
     */

    /**
     * テキストの塗りを控える
     * @param {TextFrame} textFrame - 対象のテキスト
     * @param {RestoreOptions} restoreOptions - 復元オプション
     * @returns {TextFillState} 控えた塗り
     */
    function captureTextFillState(textFrame, restoreOptions) {
        var textRange = textFrame.textRange;
        var characterCount = 0;
        try {
            characterCount = textRange.characters.length;
        } catch (e) {
            characterCount = 0;
        }

        var characterFills = [];
        if (restoreOptions.textFillPerChar && characterCount > 0) {
            var characters = textRange.characters;
            for (var i = 0; i < characterCount; i++) {
                characterFills.push(captureTextRangeFill(characters[i]));
            }
        }

        var firstCharacterFill = null;
        if (restoreOptions.textFillFirst && characterCount > 0) {
            firstCharacterFill = captureTextRangeFill(textRange.characters[0]);
        }

        return {
            characterCount: characterCount,
            characterFills: characterFills,
            firstCharacterFill: firstCharacterFill,
            rangeFill: captureTextRangeFill(textRange)
        };
    }

    /**
     * 控えたテキストの塗りを再適用する
     * @param {TextFrame} textFrame - 対象のテキスト
     * @param {TextFillState} textFillState - 控えた塗り
     * @param {RestoreOptions} restoreOptions - 復元オプション
     * @returns {void}
     */
    function applyTextFillState(textFrame, textFillState, restoreOptions) {
        var textRange = textFrame.textRange;

        if (restoreOptions.textFillPerChar && textFillState.characterCount > 0) {
            /* 消去後に文字を取り直す / Re-resolve characters after clearing */
            var characters = textRange.characters;
            var restoreCount = Math.min(textFillState.characterCount, characters.length);
            for (var i = 0; i < restoreCount; i++) {
                applyTextRangeFill(characters[i], textFillState.characterFills[i]);
            }
            return;
        }

        if (restoreOptions.textFillFirst) {
            applyTextRangeFill(textRange, textFillState.firstCharacterFill || textFillState.rangeFill);
        }
    }

    // =========================================
    // Illustratorアクション
    // Illustrator action
    // =========================================

    var ACTION_SET_NAME = "Appearance";  /* アクションセット名 / action set name */
    var ACTION_NAME     = "Clear";       /* アクション名 / action name */

    /**
     * 「アピアランスを消去」アクションの定義文を組み立てる
     * @returns {string} .aia ファイルに書き出す定義文
     */
    function buildClearAppearanceActionDefinition() {
        return [
            '/version 3',
            '/name [ 10',
            ' 417070656172616e6365',
            ']',
            '/isOpen 1',
            '/actionCount 1',
            '/action-1 {',
            ' /name [ 5',
            ' 436c656172',
            ' ]',
            ' /keyIndex 0',
            ' /colorIndex 0',
            ' /isOpen 1',
            ' /eventCount 1',
            ' /event-1 {',
            ' /useRulersIn1stQuadrant 0',
            ' /internalName (ai_plugin_appearance)',
            ' /localizedName [ 18',
            ' e382a2e38394e382a2e383a9e383b3e382b9',
            ' ]',
            ' /isOpen 1',
            ' /isOn 1',
            ' /hasDialog 0',
            ' /parameterCount 1',
            ' /parameter-1 {',
            ' /key 1835363957',
            ' /showInPalette 4294967295',
            ' /type (enumerated)',
            ' /name [ 27',
            ' e382a2e38394e382a2e383a9e383b3e382b9e38292e6b688e58ebb',
            ' ]',
            ' /value 6',
            ' }',
            ' }',
            '}'
        ].join('');
    }

    /**
     * 「アピアランスを消去」アクションを一時ファイル経由で読み込む
     * オブジェクトごとに読み込み直すと遅いので、処理の前に一度だけ呼ぶ
     * @returns {File} 読み込みに使った一時ファイル（解除時に削除する）
     */
    function loadClearAppearanceAction() {
        var tempFileName = "ClearAppearance_" + (new Date().getTime()) + "_" + Math.floor(Math.random() * 100000) + ".aia";
        var actionFile = new File(Folder.temp.fsName + "/" + tempFileName);

        actionFile.open("w");
        actionFile.write(buildClearAppearanceActionDefinition());
        actionFile.close();
        app.loadAction(actionFile);

        return actionFile;
    }

    /**
     * 読み込んだアクションを解除し、一時ファイルを削除する
     * @param {File} actionFile - 読み込みに使った一時ファイル
     * @param {FailureLog} failureLog - 集計オブジェクト
     * @returns {void}
     */
    function unloadClearAppearanceAction(actionFile, failureLog) {
        try {
            app.unloadAction(ACTION_SET_NAME, "");
        } catch (e) {
            recordFailure(failureLog, "actionUnload", null, e);
        }

        /* 消せなくても一時ファイルなので報告しない / A leftover temp file is not worth reporting */
        try {
            if (actionFile && actionFile.exists) {
                actionFile.remove();
            }
        } catch (err) { }
    }

    /**
     * 読み込み済みの「アピアランスを消去」アクションを実行する
     * @param {FailureLog} failureLog - 集計オブジェクト
     * @returns {void}
     */
    function runClearAppearanceAction(failureLog) {
        try {
            app.doScript(ACTION_NAME, ACTION_SET_NAME, false);
        } catch (e) {
            recordFailure(failureLog, "action", null, e);
            throw markActionFailure(e);
        }
    }

    // =========================================
    // メイン処理
    // Main processing
    // =========================================

    /**
     * 対象オブジェクトだけを選択する
     * @param {PageItem} targetItem - 選択するオブジェクト
     * @returns {void}
     */
    function selectOnlyItem(targetItem) {
        app.activeDocument.selection = null;
        targetItem.selected = true;
    }

    /**
     * 失敗を記録するときのカテゴリを、オブジェクトの種別から決める
     * @param {PageItem} item - 対象オブジェクト
     * @returns {string} LABELS.detailCategory のキー
     */
    function getItemFailureCategory(item) {
        try {
            return (item && item.typename === "TextFrame") ? "text" : "path";
        } catch (e) {
            return "path";
        }
    }

    /**
     * アピアランスを消去し、オブジェクト属性だけを戻す
     * クリップグループや復元対象外の種別に使う
     * @param {PageItem} item - 対象オブジェクト
     * @param {FailureLog} failureLog - 集計オブジェクト
     * @param {RestoreOptions} restoreOptions - 復元オプション
     * @returns {void}
     */
    function clearAppearanceOnly(item, failureLog, restoreOptions) {
        try {
            var objectAttributes = captureObjectAttributes(item, restoreOptions);

            selectOnlyItem(item);
            runClearAppearanceAction(failureLog);

            applyObjectAttributes(item, objectAttributes);
        } catch (e) {
            recordItemFailure(failureLog, getItemFailureCategory(item), item, e);
        }
    }

    /**
     * パスのアピアランスを消去し、塗り・線を復元する
     * @param {PageItem} item - パスまたは複合パス
     * @param {PathItem} styleHost - 塗り・線を保持するパス
     * @param {FailureLog} failureLog - 集計オブジェクト
     * @param {RestoreOptions} restoreOptions - 復元オプション
     * @returns {void}
     */
    function clearPathAppearance(item, styleHost, failureLog, restoreOptions) {
        try {
            var pathStyle = capturePathStyle(styleHost, restoreOptions);
            var objectAttributes = captureObjectAttributes(item, restoreOptions);

            selectOnlyItem(item);
            runClearAppearanceAction(failureLog);

            /* 消去後にパスを取り直す（複合パスは構成パスの参照が変わりうる）/ Re-resolve the style host after clearing */
            applyPathStyle(resolvePathStyleHost(item) || styleHost, pathStyle);
            applyObjectAttributes(item, objectAttributes);
        } catch (e) {
            recordItemFailure(failureLog, "path", item, e);
        }
    }

    /**
     * テキストのアピアランスを消去し、設定に応じて塗りだけを復元する
     * @param {TextFrame} textFrame - 対象のテキスト
     * @param {FailureLog} failureLog - 集計オブジェクト
     * @param {RestoreOptions} restoreOptions - 復元オプション
     * @returns {void}
     */
    function clearTextAppearance(textFrame, failureLog, restoreOptions) {
        try {
            var textFillState = captureTextFillState(textFrame, restoreOptions);
            var objectAttributes = captureObjectAttributes(textFrame, restoreOptions);

            selectOnlyItem(textFrame);
            runClearAppearanceAction(failureLog);

            applyTextFillState(textFrame, textFillState, restoreOptions);
            applyObjectAttributes(textFrame, objectAttributes);
        } catch (e) {
            recordItemFailure(failureLog, "text", textFrame, e);
        }
    }

    /**
     * 選択オブジェクトを再帰的に処理する
     * @param {Array} items - 処理するオブジェクト
     * @param {FailureLog} failureLog - 集計オブジェクト
     * @param {RestoreOptions} restoreOptions - 復元オプション
     * @returns {void}
     */
    function processItems(items, failureLog, restoreOptions) {
        for (var i = 0; i < items.length; i++) {
            var item = items[i];
            if (!item) {
                continue;
            }

            var styleHost = null;

            switch (item.typename) {
                case "GroupItem":
                    /* クリップグループは中身をばらすと見た目が壊れるのでまとめて処理 / Keep clipped groups intact */
                    if (item.clipped) {
                        clearAppearanceOnly(item, failureLog, restoreOptions);
                    } else {
                        processItems(item.pageItems, failureLog, restoreOptions);
                    }
                    break;

                case "PathItem":
                case "CompoundPathItem":
                    styleHost = restoreOptions.fillStroke ? resolvePathStyleHost(item) : null;
                    if (styleHost) {
                        clearPathAppearance(item, styleHost, failureLog, restoreOptions);
                    } else {
                        clearAppearanceOnly(item, failureLog, restoreOptions);
                    }
                    break;

                case "TextFrame":
                    if (restoreOptions.textFillFirst || restoreOptions.textFillPerChar) {
                        clearTextAppearance(item, failureLog, restoreOptions);
                    } else {
                        clearAppearanceOnly(item, failureLog, restoreOptions);
                    }
                    break;

                default:
                    /* 画像・シンボル・ブレンドなども消去自体は有効 / Clearing works on images, symbols, blends, etc. */
                    clearAppearanceOnly(item, failureLog, restoreOptions);
                    break;
            }
        }
    }

    /**
     * 処理前の選択状態を復元する
     * @param {Array} items - 処理前に控えた選択オブジェクト
     * @param {FailureLog} failureLog - 集計オブジェクト
     * @returns {void}
     */
    function restoreSelection(items, failureLog) {
        if (!items || items.length === 0) {
            return;
        }

        app.activeDocument.selection = null;
        for (var i = 0; i < items.length; i++) {
            try {
                if (items[i]) {
                    items[i].selected = true;
                }
            } catch (e) {
                recordItemFailure(failureLog, "selectionRestore", items[i], e);
            }
        }
    }

    // =========================================
    // 実行
    // Run
    // =========================================

    /**
     * スクリプトの入口
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel(LABELS.alert.noDocument));
            return;
        }

        if (app.selection.length === 0) {
            alert(getLabel(LABELS.alert.noSelection));
            return;
        }

        var originalSelection = toItemArray(app.selection);
        var restoreOptions = showRestoreDialog(getRestoreAvailability(originalSelection));
        if (!restoreOptions) {
            return;
        }

        var failureLog = createFailureLog();
        var actionFile = null;

        try {
            actionFile = loadClearAppearanceAction();
        } catch (e) {
            recordFailure(failureLog, "action", null, e);
            alertFailures(failureLog);
            return;
        }

        try {
            processItems(originalSelection, failureLog, restoreOptions);
        } finally {
            unloadClearAppearanceAction(actionFile, failureLog);
            restoreSelection(originalSelection, failureLog);
        }

        alertFailures(failureLog);
    }

    main();
})();
