#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択オブジェクトを複製して背面に置き、オフセットパス→アウトライン→合体→拡張の順で処理して縁取りを作ります。
元オブジェクトと結果をグループ化し、Subtract を実行して白で塗りつぶします。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AddOutlineOffsetPath.md

### Overview

Duplicates the selection behind itself and runs Offset Path, Outline Stroke, Unite and Expand to build an outline around it.
The original and the result are grouped, then Subtract is run and the result filled with white.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AddOutlineOffsetPath.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AddOutlineOffsetPath";         /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-08-13";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-27";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AddOutlineOffsetPath.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AddOutlineOffsetPath.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // レイアウト / Layout
    // =========================================
    var PANEL_MARGINS   = [15, 20, 15, 10];  /* パネル余白 [左,上,右,下] */
    var BUTTON_SPACING  = 10;                /* ボタンの間隔 */
    var DIALOG_OFFSET_X = 20;                /* ダイアログを右へずらす量 */
    var DIALOG_OPACITY  = 0.98;              /* ダイアログの不透明度 */

    // =========================================
    // 単位 / Units
    // =========================================

    /* 単位コードに対応する表示ラベルと、1単位あたりのポイント数
       Unit code -> display label and points per unit */
    var UNITS = [
        { label: "in",    pointsPerUnit: 72 },                /* 0 */
        { label: "mm",    pointsPerUnit: 72 / 25.4 },         /* 1 */
        { label: "pt",    pointsPerUnit: 1 },                 /* 2 */
        { label: "pica",  pointsPerUnit: 12 },                /* 3 */
        { label: "cm",    pointsPerUnit: 72 / 2.54 },         /* 4 */
        { label: "Q",     pointsPerUnit: 72 / 25.4 * 0.25 },  /* 5 */
        { label: "px",    pointsPerUnit: 1 },                 /* 6 */
        { label: "ft/in", pointsPerUnit: 72 * 12 },           /* 7 */
        { label: "m",     pointsPerUnit: 72 / 25.4 * 1000 },  /* 8 */
        { label: "yd",    pointsPerUnit: 72 * 36 },           /* 9 */
        { label: "ft",    pointsPerUnit: 72 * 12 }            /* 10 */
    ];

    /* 単位コード5を「歯（H）」と表示する環境設定キー。文字サイズ（text/units）だけ「級（Q）」
       Preference keys that show unit code 5 as H; only the type size (text/units) shows Q */
    var HA_UNIT_PREF_KEYS = { "rulerType": true, "strokeUnits": true, "text/asianunits": true };

    /**
     * 環境設定キーの単位を返す
     * @param {string} [prefKey] - "rulerType"（既定）/ "strokeUnits" / "text/units" / "text/asianunits"
     * @returns {{code: number, label: string, pointsPerUnit: number}} 単位の情報
     */
    function getUnitInfo(prefKey) {
        var unitKey = prefKey || "rulerType";
        var unitCode = app.preferences.getIntegerPreference(unitKey);
        /* 未知のコードは pt に寄せる / unknown codes fall back to points */
        var unit = UNITS[unitCode] || UNITS[2];
        /* 級（Q）と歯（H）は同じ長さだが、文字サイズは「Q」、距離は「H」と呼び分ける */
        var label = (unitCode === 5 && HA_UNIT_PREF_KEYS[unitKey]) ? "H" : unit.label;
        return { code: unitCode, label: label, pointsPerUnit: unit.pointsPerUnit };
    }

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * Illustrator の UI 言語から表示言語を判定する
     * @returns {string} "ja" または "en"
     */
    function detectUILang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = detectUILang();

    var LABELS = {
        dialog: {
            title: { ja: "アウトライン オフセットパス", en: "Outline Offset Path" }
        },
        panel: {
            offset: { ja: "オフセット設定", en: "Offset settings" },
            corner: { ja: "角の形状", en: "Corner" }
        },
        radio: {
            joinMiter: { ja: "マイター結合", en: "Miter Join" },
            joinRound: { ja: "ラウンド結合", en: "Round Join" },
            joinBevel: { ja: "ベベル結合", en: "Bevel Join" }
        },
        tooltip: {
            offset: {
                ja: "元のパスから外側（マイナスで内側）へ離す距離です。",
                en: "How far the new path sits outside the original. A negative value goes inside."
            },
            joinMiter: {
                ja: "角を尖らせたまま結合します。鋭角では飛び出すことがあります。",
                en: "Keeps the corners pointed. Sharp angles can spike out."
            },
            joinRound: { ja: "角を丸めて結合します。", en: "Rounds off the corners." },
            joinBevel: { ja: "角を面取りして結合します。", en: "Cuts the corners off flat." }
        },
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" }
        },
        alert: {
            noDocument: { ja: "ドキュメントを開いてください。", en: "Please open a document." },
            noSelection: { ja: "オブジェクトが選択されていません。", en: "Please select at least one object." },
            enterNumeric: { ja: "数値を入力してください。", en: "Enter a numeric value." }
        }
    };

    /**
     * LABELS からドット区切りのパスで表示言語の文字列を引く
     * @param {string} labelPath - "alert.noSelection" のようなドット区切りのキー
     * @returns {string} 表示言語のテキスト（見つからない場合は labelPath をそのまま返す）
     */
    function getLabel(labelPath) {
        var labelPathKeys = labelPath.split(".");
        var labelNode = LABELS;
        for (var i = 0; i < labelPathKeys.length; i++) {
            labelNode = labelNode[labelPathKeys[i]];
            if (!labelNode) return labelPath;
        }
        return labelNode[uiLang] || labelNode["en"] || labelPath;
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * ↑↓キーで数値欄の値を増減する（shift で10刻み、option で0.1刻み）
     * @param {EditText} editText - 対象の数値欄
     * @returns {void}
     */
    function changeValueByArrowKey(editText) {
        editText.addEventListener("keydown", function (event) {
            /* ↑↓以外のキーは素通し / Let other keys through */
            if (event.keyName != "Up" && event.keyName != "Down") return;

            var value = Number(editText.text);
            if (isNaN(value)) return;

            var keyboard = ScriptUI.environment.keyboardState;
            var isUp = (event.keyName == "Up");

            if (keyboard.shiftKey) {
                /* 10刻み（10の倍数にスナップ、↓は0で止める）/ Step by 10, snapping; Down stops at 0 */
                value = isUp ? Math.ceil((value + 1) / 10) * 10 : Math.max(0, Math.floor((value - 1) / 10) * 10);
            } else if (keyboard.altKey) {
                /* 0.1刻み / Step by 0.1 */
                value += isUp ? 0.1 : -0.1;
            } else {
                /* 1刻み（↓は0で止める）/ Step by 1; Down stops at 0 */
                value = isUp ? value + 1 : Math.max(0, value - 1);
            }
            event.preventDefault();

            /* option は小数第1位、それ以外は整数に丸める / Round to 0.1 with option, otherwise to integers */
            value = keyboard.altKey ? Math.round(value * 10) / 10 : Math.round(value);
            editText.text = value;
        });
    }

    /**
     * ダイアログを表示時に右へずらす
     * @param {Window} dlg - 対象のダイアログ
     * @param {number} offsetX - 横方向のずらし量
     * @param {number} offsetY - 縦方向のずらし量
     * @returns {void}
     */
    function shiftDialogPosition(dlg, offsetX, offsetY) {
        dlg.onShow = function () {
            var currentLocation = dlg.location;
            dlg.location = [currentLocation[0] + offsetX, currentLocation[1] + offsetY];
        };
    }

    /**
     * ダイアログの不透明度を設定する
     * @param {Window} dlg - 対象のダイアログ
     * @param {number} opacityValue - 不透明度（0〜1）
     * @returns {void}
     */
    function setDialogOpacity(dlg, opacityValue) {
        /* 環境によっては opacity が使えない / opacity is not supported on some platforms */
        try {
            dlg.opacity = opacityValue;
        } catch (e) { }
    }

    /**
     * オフセット量と角の形状を尋ねるダイアログを表示する
     * @param {number} defaultOffset - オフセット量の初期値（定規単位）
     * @param {string} unitLabel - 単位の表示ラベル
     * @returns {{offsetText: string, joinType: number}|null} 入力値（キャンセル時は null）。joinType は 0=ラウンド、1=ベベル、2=マイター
     */
    function showOffsetDialog(defaultOffset, unitLabel) {
        var dlg = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        dlg.alignChildren = ["left", "top"];
        setDialogOpacity(dlg, DIALOG_OPACITY);
        shiftDialogPosition(dlg, DIALOG_OFFSET_X, 0);

        /* オフセット量 / Offset amount */
        var offsetPanel = dlg.add("panel", undefined, getLabel("panel.offset"));
        offsetPanel.orientation = "row";
        offsetPanel.alignChildren = ["left", "center"];
        offsetPanel.margins = PANEL_MARGINS;
        var offsetInput = offsetPanel.add("edittext", undefined, String(defaultOffset));
        offsetInput.helpTip = getLabel("tooltip.offset");
        offsetInput.characters = 4;
        offsetInput.active = true;
        changeValueByArrowKey(offsetInput);
        offsetPanel.add("statictext", undefined, unitLabel);

        /* 角の形状（jntp）/ Corner join type */
        var cornerPanel = dlg.add("panel", undefined, getLabel("panel.corner"));
        cornerPanel.orientation = "column";
        cornerPanel.alignChildren = ["left", "center"];
        cornerPanel.margins = PANEL_MARGINS;
        var rbMiter = cornerPanel.add("radiobutton", undefined, getLabel("radio.joinMiter"));
        rbMiter.helpTip = getLabel("tooltip.joinMiter");
        var rbRound = cornerPanel.add("radiobutton", undefined, getLabel("radio.joinRound"));
        rbRound.helpTip = getLabel("tooltip.joinRound");
        var rbBevel = cornerPanel.add("radiobutton", undefined, getLabel("radio.joinBevel"));
        rbBevel.helpTip = getLabel("tooltip.joinBevel");
        rbRound.value = true; /* 既定はラウンド / Round by default */

        /* ボタン（中央寄せ）/ Buttons (centered) */
        var btnRowGroup = dlg.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.alignment = ["center", "center"];
        btnRowGroup.spacing = BUTTON_SPACING;
        btnRowGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        btnRowGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });

        if (dlg.show() !== 1) {
            return null;
        }
        return {
            offsetText: offsetInput.text,
            joinType: rbRound.value ? 0 : (rbBevel.value ? 1 : 2)
        };
    }

    // =========================================
    // 縁取りの作成 / Outline creation
    // =========================================

    /**
     * 選択範囲全体の幅を返す
     * @param {PageItem[]} items - 対象のオブジェクト
     * @param {boolean} useVisible - true なら visibleBounds、false なら geometricBounds
     * @returns {number} 幅（pt）。取得できなければ 0
     */
    function getSelectionWidthPt(items, useVisible) {
        var left = null, right = null;
        for (var i = 0; i < items.length; i++) {
            /* 境界を持たないオブジェクトは読み飛ばす / Skip items whose bounds cannot be read */
            try {
                var bounds = useVisible ? items[i].visibleBounds : items[i].geometricBounds; /* [left, top, right, bottom] */
                if (!bounds) continue;
                if (left === null || bounds[0] < left) left = bounds[0];
                if (right === null || bounds[2] > right) right = bounds[2];
            } catch (e) { }
        }
        if (left === null || right === null) return 0;
        return right - left;
    }

    /**
     * パス・複合パス・グループ内を再帰的に白で塗る
     * @param {PageItem} node - 対象のオブジェクト
     * @returns {void}
     */
    function fillWithWhite(node) {
        if (!node) return;
        var white = new RGBColor();
        white.red = 255; white.green = 255; white.blue = 255;
        /* ロック中などで塗れないものは飛ばす / Skip items that cannot be painted (e.g. locked) */
        try {
            if (node.typename === "PathItem") {
                node.filled = true;
                node.fillColor = white;
            } else if (node.typename === "CompoundPathItem") {
                for (var i = 0; i < node.pathItems.length; i++) {
                    node.pathItems[i].filled = true;
                    node.pathItems[i].fillColor = white;
                }
            } else if (node.typename === "GroupItem") {
                for (var j = 0; j < node.pageItems.length; j++) {
                    fillWithWhite(node.pageItems[j]);
                }
            }
        } catch (e) { }
    }

    /**
     * 1つのオブジェクトに縁取りを付ける（複製を背面に置き、オフセット→アウトライン→合体→拡張、元とグループ化して前面で型抜き、白塗り）
     * @param {Document} doc - 対象ドキュメント
     * @param {PageItem} original - 元のオブジェクト
     * @param {number} offsetValue - オフセット量（pt）
     * @param {number} joinType - 角の形状（0=ラウンド、1=ベベル、2=マイター）
     * @returns {void}
     */
    function addOutlineToItem(doc, original, offsetValue, joinType) {
        doc.selection = null;

        /* 複製して背面へ / Duplicate and send to back */
        var offsetCopy = original.duplicate();
        offsetCopy.zOrder(ZOrderMethod.SENDTOBACK);
        doc.selection = [offsetCopy];

        /* オフセットを Live Effect で適用 / Apply Offset Path (Live Effect) */
        var offsetEffectXml = '<LiveEffect name="Adobe Offset Path"><Dict data="R mlim 4 R ofst ' + offsetValue + ' I jntp ' + joinType + ' "/><\/LiveEffect>';
        doc.selection[0].applyEffect(offsetEffectXml);

        /* アウトライン→合体→拡張 / Outline → Unite → Expand */
        app.executeMenuCommand("Live Outline Stroke");
        app.executeMenuCommand("Live Pathfinder Add");
        app.executeMenuCommand("expandStyle");
        app.executeMenuCommand("noCompoundPath");
        app.executeMenuCommand("Live Pathfinder Add");
        app.executeMenuCommand("expandStyle");

        /* 元と結果をグループ→前面オブジェクトで型抜き→白塗り / Group original & result → Subtract → Fill white */
        var resultItem = (doc.selection && doc.selection.length) ? doc.selection[0] : null;
        if (!resultItem) return;
        doc.selection = [original, resultItem];
        app.executeMenuCommand("group");
        app.executeMenuCommand("Live Pathfinder Subtract");
        app.executeMenuCommand("expandStyle");

        /* グループは残したまま白で塗る / Keep the group and fill it white */
        if (doc.selection && doc.selection.length) {
            resultItem = doc.selection[0];
            doc.selection = [resultItem];
            fillWithWhite(resultItem);
        }
    }

    /**
     * 選択中の各オブジェクトに縁取りを付ける（アラートは処理中だけ抑止）
     * @param {PageItem[]} items - 対象のオブジェクト
     * @param {number} offsetValue - オフセット量（pt）
     * @param {number} joinType - 角の形状（0=ラウンド、1=ベベル、2=マイター）
     * @returns {void}
     */
    function addOutline(items, offsetValue, joinType) {
        var doc = app.activeDocument;
        var prevInteractionLevel = app.userInteractionLevel;

        /* アラート抑止（終了時に復元）/ Suppress alerts (restore on exit) */
        try {
            app.userInteractionLevel = UserInteractionLevel.DONTDISPLAYALERTS;

            var sourceItems = [].concat(items);
            for (var i = 0; i < sourceItems.length; i++) {
                if (!sourceItems[i]) continue;
                /* 失敗したオブジェクトは飛ばして続行 / Skip an item that fails and continue */
                try {
                    addOutlineToItem(doc, sourceItems[i], offsetValue, joinType);
                } catch (e) { }
            }
        } finally {
            app.userInteractionLevel = prevInteractionLevel;
        }
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択を確認し、ダイアログの入力に従って縁取りを付ける
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }
        var doc = app.activeDocument;
        /* 未選択時のガード / Guard when nothing is selected */
        if (doc.selection.length === 0) {
            alert(getLabel("alert.noSelection"));
            return;
        }

        /* 選択幅の1/30を初期値にする（定規単位、小数第1位）/ Default offset: 1/30 of the selection width in ruler units */
        var rulerUnit = getUnitInfo("rulerType");
        var selectionWidthPt = getSelectionWidthPt(doc.selection, true);
        var defaultOffset = Math.max(0, Math.round(((selectionWidthPt / 30) / rulerUnit.pointsPerUnit) * 10) / 10);

        var dialogResult = showOffsetDialog(defaultOffset, rulerUnit.label);
        if (!dialogResult) {
            return;
        }
        var offsetValue = parseFloat(dialogResult.offsetText);
        if (isNaN(offsetValue)) {
            alert(getLabel("alert.enterNumeric"));
            return;
        }
        addOutline(doc.selection, offsetValue * rulerUnit.pointsPerUnit, dialogResult.joinType);
    }

    main();

})();
