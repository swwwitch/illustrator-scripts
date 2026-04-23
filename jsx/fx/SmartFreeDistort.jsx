#target illustrator

// =========================================
// SmartFreeDistort.jsx
// バージョン: v1.1.1
// 更新日: 2026-04-24
// =========================================
// ### 概要:
// 「自由変形」効果を、プリセット・変形量・強度の組み合わせで適用するスクリプト。
// 選択オブジェクトに対して、ダイアログUIから変形プリセットを選択し、プレビューで確認しながら適用できます。
//
// ### 主な機能:
// ・調整変形
//   - 台形（上辺を狭く / 広く、下辺を狭く / 広く）
//   - 平行四辺形（右シアー / 左シアー / 上シアー / 下シアー）
//   - 変形量スライダー（0.00〜0.49、ドラッグと↑↓キー操作に対応）
//   - 強度切替（マイルド / ノーマル / ブースト）
//     - 台形 / 平行四辺形で有効
//     - 三角形 / 対角線ではディム表示
//     - 選択対象にテキストが含まれる場合、初期値はマイルド
// ・固定変形
//   - 三角形（左下 / 右下 / 左上 / 右上）
//   - 対角線（＼ / ／）
// ・プレビュー機能（Undoベースで一時適用）
//   - スライダーのドラッグ中は更新頻度を抑制
//   - 同一条件では再計算をスキップ
// ・日英ローカライズ対応UI
//
// ### 仕様・注意:
// ・三角形および対角線プリセットでは、変形量と強度は使用されません。
// ・台形 / 平行四辺形の強度は、マイルド / ノーマル / ブーストの3段階です。
// ・台形の強度倍率は、マイルド=0.25、ノーマル=0.5、ブースト=1.0 です。
// ・平行四辺形の強度倍率は、旧基準をベースに定数化しています。
// ・プレビューは Undo を利用して制御しているため、他の操作と混在すると不整合が起こる可能性があります。
// ・複数オブジェクト選択時は、すべてに同一のライブ効果を適用します。
// ・シアー系の名称は、座標の増減方向そのものではなく見た目基準です。
// 
// ### GitHub
// https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/fx/SmartFreeDistort.jsx
//
// ### 更新履歴:
// v1.1.1 (2026-04-24)
// ・台形 / 平行四辺形 / 三角形 / 対角線のプリセットUIを実装
// ・変形量スライダーとプレビューを実装
// ・強度切替（ノーマル / マイルド / ブースト）を追加
// ・台形 / 平行四辺形で強度切替を有効化
// ・台形 / 平行四辺形ともに強度切替の3段階を利用可能
// ・選択対象にテキストが含まれる場合、強度初期値をマイルドに設定
// ・プレビュー更新の間引きと同一条件スキップを追加
// ・日英ローカライズ対応UIを実装

var SCRIPT_VERSION = "v1.1.1";

function getCurrentLang() {
    return ($.locale && $.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

var LABELS = {
    dialogTitle: { ja: "スマート自由変形", en: "Smart Free Distort" },

    alertNoDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
    alertNoSelection: { ja: "オブジェクトを選択してください。", en: "Please select one or more objects." },

    pnlType: { ja: "調整変形", en: "Adjustable Transform" },
    pnlTypeB: { ja: "固定変形", en: "Fixed Transform" },
    pnlTrapezoid: { ja: "台形", en: "Trapezoid" },
    pnlShear: { ja: "平行四辺形", en: "Parallelogram" },
    pnlTriangle: { ja: "三角形", en: "Triangle" },
    pnlDiagonal: { ja: "対角線", en: "Diagonal" },
    pnlAmount: { ja: "変形量", en: "Amount" },

    rbTrapNarrowTop: { ja: "上辺を狭く", en: "Narrow Top" },
    rbTrapWideTop: { ja: "上辺を広く", en: "Wide Top" },
    rbTrapNarrowBottom: { ja: "下辺を狭く", en: "Narrow Bottom" },
    rbTrapWideBottom: { ja: "下辺を広く", en: "Wide Bottom" },
    rbShearRight: { ja: "右シアー", en: "Right Shear" },
    rbShearLeft: { ja: "左シアー", en: "Left Shear" },
    rbShearUp: { ja: "上シアー", en: "Up Shear" },
    rbShearDown: { ja: "下シアー", en: "Down Shear" },
    rbTriBL: { ja: "左下", en: "Bottom Left" },
    rbTriBR: { ja: "右下", en: "Bottom Right" },
    rbTriTL: { ja: "左上", en: "Top Left" },
    rbTriTR: { ja: "右上", en: "Top Right" },
    rbDiagonalRight: { ja: "＼", en: "\\" },
    rbDiagonalLeft: { ja: "／", en: "/" },
    tipAmount: { ja: "0.00-0.49", en: "0.00-0.49" },

    pnlStrength: { ja: "強度", en: "Strength" },
    rbStrengthNormal: { ja: "ノーマル", en: "Normal" },
    rbStrengthMild: { ja: "マイルド", en: "Mild" },
    rbStrengthBoost: { ja: "ブースト", en: "Boost" },

    chkPreview: { ja: "プレビュー", en: "Preview" },

    btnCancel: { ja: "キャンセル", en: "Cancel" },
    btnOK: { ja: "OK", en: "OK" }
};

function L(key) {
    if (!LABELS[key]) return key;
    return LABELS[key][lang] || LABELS[key].ja || LABELS[key].en || key;
}

(function () {

    if (app.documents.length === 0) {
        alert(L("alertNoDocument"));
        return;
    }

    var doc = app.activeDocument;
    if (!doc.selection || doc.selection.length === 0) {
        alert(L("alertNoSelection"));
        return;
    }

    // =========================================================
    // Smart Free Distort main
    // =========================================================
    function applySmartFreeDistort(item, options) {
        var defaults = {
            distortRect: [[0, 0], [1, 0], [0, 1], [1, 1]], // [TL, TR, BL, BR]
            sourceRect: [[0, 0], [1, 0], [0, 1], [1, 1]]
        };

        var o = SFD.mergeOptions(defaults, options);

        var xml = '<LiveEffect name="Adobe Free Distort"><Dict data="R src0h #1 R src0v #2 R src1h #3 R src1v #4 R src2h #5 R src2v #6 R src3h #7 R src3v #8 R dst0h #9 R dst0v #10 R dst1h #11 R dst1v #12 R dst2h #13 R dst2v #14 R dst3h #15 R dst3v #16 "/></LiveEffect>'
            .replace(/#1/, o.sourceRect[0][0])
            .replace(/#2/, -o.sourceRect[0][1])
            .replace(/#3/, o.sourceRect[1][0])
            .replace(/#4/, -o.sourceRect[1][1])
            .replace(/#5/, o.sourceRect[2][0])
            .replace(/#6/, -o.sourceRect[2][1])
            .replace(/#7/, o.sourceRect[3][0])
            .replace(/#8/, -o.sourceRect[3][1])
            .replace(/#9/, o.distortRect[0][0])
            .replace(/#10/, -o.distortRect[0][1])
            .replace(/#11/, o.distortRect[1][0])
            .replace(/#12/, -o.distortRect[1][1])
            .replace(/#13/, o.distortRect[2][0])
            .replace(/#14/, -o.distortRect[2][1])
            .replace(/#15/, o.distortRect[3][0])
            .replace(/#16/, -o.distortRect[3][1]);

        SFD.applyLiveEffect(item, xml);
    }

    var SFD = {
        debug: false,

        mergeOptions: function (defaults, options) {
            try {
                var merged = {};
                var key;

                options = options || {};

                for (key in defaults) {
                    if (defaults.hasOwnProperty(key)) {
                        merged[key] = defaults[key];
                    }
                }

                for (key in options) {
                    if (options.hasOwnProperty(key)) {
                        merged[key] = options[key];
                    }
                }

                if (options.debug) SFD.debug = true;
                return merged;
            } catch (error) {
                throw new Error('SmartFreeDistort failed to parse options object. ' + error);
            }
        },

        applyLiveEffect: function (item, xml) {
            var items = [];

            if (item == undefined) {
                throw new Error('SmartFreeDistort failed. No item available.');
            }

            if (item.typename != undefined) {
                items = [item];
            } else if (item.length != undefined) {
                for (var i = 0; i < item.length; i++) {
                    items.push(item[i]);
                }
            } else {
                throw new Error('SmartFreeDistort failed. Unexpected item type.');
            }

            for (var j = 0; j < items.length; j++) {
                if (!items[j] || items[j].typename == undefined) {
                    throw new Error('SmartFreeDistort failed. Unexpected item type in collection.');
                }
                items[j].applyEffect(xml);
            }

            if (SFD.debug) $.writeln('SmartFreeDistort:\n' + xml);
        },

        showError: function (error) {
            alert(error.message);
        }
    };

    // =========================================================
    // distortion preset generator
    // amount: 0.0 〜 0.49 程度
    // =========================================================

    var SHEAR_BASE_FACTOR = 2;
    var SHEAR_STRENGTH_FACTORS = {
        normal: 0.5,
        mild: 0.25,
        boost: 1.0
    };

    var TRAPEZOID_STRENGTH_FACTORS = {
        mild: 0.25,
        normal: 0.5,
        boost: 1.0
    };

    function applyShearStrength(amount, strengthKey) {
        var factor = SHEAR_STRENGTH_FACTORS[strengthKey];
        if (factor == null) factor = SHEAR_STRENGTH_FACTORS.normal;
        return amount * SHEAR_BASE_FACTOR * factor;
    }

    function applyTrapezoidStrength(amount, strengthKey) {
        var factor = TRAPEZOID_STRENGTH_FACTORS[strengthKey];
        if (factor == null) factor = TRAPEZOID_STRENGTH_FACTORS.normal;
        return amount * factor;
    }

    function makeDistortRect(type, amount, strengthKey) {
        var trapezoidAmount = applyTrapezoidStrength(amount, strengthKey);
        var shearAmount = applyShearStrength(amount, strengthKey);

        // シアー系の名称は、座標の増減方向そのものではなく、
        // ダイアログ上で確認される見た目の方向を基準にしています。
        // そのため内部座標の増減方向とラベル名は一致しない場合があります。

        switch (type) {
            case "trap_narrow_top":
                // 台形変形（上辺を狭く）
                return [
                    [0 + trapezoidAmount, 0],
                    [1 - trapezoidAmount, 0],
                    [0, 1],
                    [1, 1]
                ];

            case "trap_wide_top":
                // 台形変形（上辺を広く）
                return [
                    [0 - trapezoidAmount, 0],
                    [1 + trapezoidAmount, 0],
                    [0, 1],
                    [1, 1]
                ];

            case "trap_narrow_bottom":
                // 台形変形（下辺を狭く）
                return [
                    [0, 0],
                    [1, 0],
                    [0 + trapezoidAmount, 1],
                    [1 - trapezoidAmount, 1]
                ];

            case "trap_wide_bottom":
                // 台形変形（下辺を広く）
                return [
                    [0, 0],
                    [1, 0],
                    [0 - trapezoidAmount, 1],
                    [1 + trapezoidAmount, 1]
                ];

            case "shear_right":
                // 平行四辺形（右シアー）※見た目基準の名称
                return [
                    [0 + shearAmount, 0],
                    [1 + shearAmount, 0],
                    [0, 1],
                    [1, 1]
                ];

            case "shear_left":
                // 平行四辺形（左シアー）※見た目基準の名称
                return [
                    [0, 0],
                    [1, 0],
                    [0 + shearAmount, 1],
                    [1 + shearAmount, 1]
                ];

            case "shear_up":
                // 平行四辺形（上シアー）※見た目基準の名称
                return [
                    [0, 0],
                    [1, 0 + shearAmount],
                    [0, 1],
                    [1, 1 + shearAmount]
                ];

            case "shear_down":
                // 平行四辺形（下シアー）※見た目基準の名称
                return [
                    [0, 0 + shearAmount],
                    [1, 0],
                    [0, 1 + shearAmount],
                    [1, 1]
                ];

            case "tri_bl":
                // 三角形（左下）：右上の角をたたむ
                return [
                    [0, 0],
                    [0, 0],
                    [0, 1],
                    [1, 1]
                ];

            case "tri_br":
                // 三角形（右下）：左上の角をたたむ
                return [
                    [1, 0],
                    [1, 0],
                    [0, 1],
                    [1, 1]
                ];

            case "tri_tl":
                // 三角形（左上）：右下の角をたたむ
                return [
                    [0, 0],
                    [1, 0],
                    [0, 1],
                    [0, 1]
                ];

            case "tri_tr":
                // 三角形（右上）：左下の角をたたむ
                return [
                    [0, 0],
                    [1, 0],
                    [1, 1],
                    [1, 1]
                ];

            case "diagonal_right":
                // 対角線（＼）
                return [
                    [0, 0],
                    [0, 0],
                    [1, 1],
                    [1, 1]
                ];

            case "diagonal_left":
                // 対角線（／）
                return [
                    [1, 0],
                    [1, 0],
                    [0, 1],
                    [0, 1]
                ];

            default:
                return [
                    [0, 0],
                    [1, 0],
                    [0, 1],
                    [1, 1]
                ];
        }
    }

    // =========================================================
    // dialog
    // =========================================================
    // =========================================================
    // helper: 選択内にテキストが含まれるか判定
    // =========================================================
    function selectionHasText(items) {
        if (!items || items.length == null) return false;

        function hasTextInItem(item) {
            if (!item || !item.typename) return false;
            if (item.typename === "TextFrame") return true;
            if (item.typename === "GroupItem") {
                var children = item.pageItems;
                for (var i = 0; i < children.length; i++) {
                    if (hasTextInItem(children[i])) return true;
                }
            }
            return false;
        }

        for (var i = 0; i < items.length; i++) {
            if (hasTextInItem(items[i])) return true;
        }
        return false;
    }

    // =========================================================
    // build dialog UI
    // =========================================================
    function buildDialog(hasTextSelection) {
        var dlg = new Window("dialog", L("dialogTitle") + " " + SCRIPT_VERSION);
        dlg.orientation = "column";
        dlg.alignChildren = ["fill", "top"];

        var grpTypes = dlg.add("group");
        grpTypes.orientation = "row";
        grpTypes.alignChildren = ["fill", "top"];

        // --- 変形タイプA ---
        var pnlType = grpTypes.add("panel", undefined, L("pnlType"));
        pnlType.orientation = "column";
        pnlType.alignChildren = ["fill", "top"];
        pnlType.margins = [15, 20, 15, 10];

        var colLeft = pnlType.add("group");
        colLeft.orientation = "column";
        colLeft.alignChildren = ["fill", "top"];

        var grpTypeA = colLeft.add("group");
        grpTypeA.orientation = "row";
        grpTypeA.alignChildren = ["fill", "top"];

        var pnlTrapezoid = grpTypeA.add("panel", undefined, L("pnlTrapezoid"));
        pnlTrapezoid.orientation = "column";
        pnlTrapezoid.alignChildren = ["left", "top"];
        pnlTrapezoid.margins = [15, 20, 15, 10];

        var rbTypes = [];
        rbTypes[0] = pnlTrapezoid.add("radiobutton", undefined, L("rbTrapNarrowTop"));
        rbTypes[1] = pnlTrapezoid.add("radiobutton", undefined, L("rbTrapWideTop"));
        rbTypes[2] = pnlTrapezoid.add("radiobutton", undefined, L("rbTrapNarrowBottom"));
        rbTypes[3] = pnlTrapezoid.add("radiobutton", undefined, L("rbTrapWideBottom"));

        var pnlShear = grpTypeA.add("panel", undefined, L("pnlShear"));
        pnlShear.orientation = "column";
        pnlShear.alignChildren = ["left", "top"];
        pnlShear.margins = [15, 20, 15, 10];

        rbTypes[4] = pnlShear.add("radiobutton", undefined, L("rbShearRight"));
        rbTypes[5] = pnlShear.add("radiobutton", undefined, L("rbShearLeft"));
        rbTypes[6] = pnlShear.add("radiobutton", undefined, L("rbShearUp"));
        rbTypes[7] = pnlShear.add("radiobutton", undefined, L("rbShearDown"));

        // --- 変形タイプB ---
        var pnlTypeB = grpTypes.add("panel", undefined, L("pnlTypeB"));
        pnlTypeB.orientation = "column";
        pnlTypeB.alignChildren = ["fill", "top"];
        pnlTypeB.margins = [15, 20, 15, 10];

        var colRight = pnlTypeB.add("group");
        colRight.orientation = "column";
        colRight.alignChildren = ["fill", "top"];

        var pnlTriangle = colRight.add("panel", undefined, L("pnlTriangle"));
        pnlTriangle.orientation = "column";
        pnlTriangle.alignChildren = ["left", "top"];
        pnlTriangle.margins = [15, 20, 15, 10];

        rbTypes[8] = pnlTriangle.add("radiobutton", undefined, L("rbTriBL"));
        rbTypes[9] = pnlTriangle.add("radiobutton", undefined, L("rbTriBR"));
        rbTypes[10] = pnlTriangle.add("radiobutton", undefined, L("rbTriTL"));
        rbTypes[11] = pnlTriangle.add("radiobutton", undefined, L("rbTriTR"));

        var pnlDiagonal = colRight.add("panel", undefined, L("pnlDiagonal"));
        pnlDiagonal.orientation = "column";
        pnlDiagonal.alignChildren = ["fill", "top"];
        pnlDiagonal.margins = [15, 20, 15, 10];

        rbTypes[12] = pnlDiagonal.add("radiobutton", undefined, L("rbDiagonalRight"));
        rbTypes[13] = pnlDiagonal.add("radiobutton", undefined, L("rbDiagonalLeft"));
        rbTypes[0].value = true;

        // --- 変形量 ---
        var pnlAmount = colLeft.add("panel", undefined, L("pnlAmount"));
        pnlAmount.orientation = "column";
        pnlAmount.alignChildren = ["fill", "top"];
        pnlAmount.margins = [15, 20, 15, 10];

        var grpAmount = pnlAmount.add("group");
        grpAmount.orientation = "column";
        grpAmount.alignChildren = ["fill", "top"];

        var slAmount = grpAmount.add("slider", undefined, 20, 0, 49);
        slAmount.helpTip = L("tipAmount");
        slAmount.preferredSize.width = 220;

        var stAmountValue = grpAmount.add("statictext", undefined, "0.20");
        stAmountValue.alignment = ["center", "center"];

        var pnlStrength = grpAmount.add("panel", undefined, L("pnlStrength"));
        pnlStrength.orientation = "row";
        pnlStrength.alignChildren = ["left", "center"];
        pnlStrength.margins = [10, 18, 10, 10];
        var rbStrengthMild = pnlStrength.add("radiobutton", undefined, L("rbStrengthMild"));
        var rbStrengthNormal = pnlStrength.add("radiobutton", undefined, L("rbStrengthNormal"));
        var rbStrengthBoost = pnlStrength.add("radiobutton", undefined, L("rbStrengthBoost"));
        rbStrengthNormal.value = !hasTextSelection;
        rbStrengthMild.value = !!hasTextSelection;
        rbStrengthBoost.value = false;

        // --- ボタン ---
        var btns = dlg.add("group");
        btns.orientation = "row";
        btns.alignChildren = ["fill", "center"];
        btns.alignment = ["fill", "top"];

        var btnsLeft = btns.add("group");
        btnsLeft.orientation = "row";
        btnsLeft.alignChildren = ["left", "center"];
        btnsLeft.alignment = ["left", "center"];

        var chkPreview = btnsLeft.add("checkbox", undefined, L("chkPreview"));
        chkPreview.value = false;

        var btnsSpacer = btns.add("group");
        btnsSpacer.alignment = ["fill", "fill"];

        var btnsRight = btns.add("group");
        btnsRight.orientation = "row";
        btnsRight.alignChildren = ["right", "center"];
        btnsRight.alignment = ["right", "center"];

        btnsRight.add("button", undefined, L("btnCancel"), { name: "cancel" });
        btnsRight.add("button", undefined, L("btnOK"), { name: "ok" });

        return {
            dlg: dlg,
            rbTypes: rbTypes,
            pnlAmount: pnlAmount,
            slAmount: slAmount,
            stAmountValue: stAmountValue,
            pnlStrength: pnlStrength,
            rbStrengthNormal: rbStrengthNormal,
            rbStrengthMild: rbStrengthMild,
            rbStrengthBoost: rbStrengthBoost,
            chkPreview: chkPreview
        };
    }

    // =========================================================
    // show dialog and handle interaction
    // =========================================================
    function showDialog() {
        var hasTextSelection = selectionHasText(doc.selection);
        var ui = buildDialog(hasTextSelection);
        var dlg = ui.dlg;
        var rbTypes = ui.rbTypes;
        var pnlAmount = ui.pnlAmount;
        var slAmount = ui.slAmount;
        var stAmountValue = ui.stAmountValue;
        var pnlStrength = ui.pnlStrength;
        var rbStrengthNormal = ui.rbStrengthNormal;
        var rbStrengthMild = ui.rbStrengthMild;
        var rbStrengthBoost = ui.rbStrengthBoost;
        var chkPreview = ui.chkPreview;

        // プレビュー管理
        var previewed = false;
        var lastPreviewUpdateTime = 0;
        var PREVIEW_THROTTLE_MS = 150;
        var lastPreviewSignature = null;

        function getSelectedIndex() {
            for (var i = 0; i < rbTypes.length; i++) {
                if (rbTypes[i].value) return i;
            }
            return 0;
        }

        function updateAmountText() {
            stAmountValue.text = (slAmount.value / 100).toFixed(2);
        }

        function getStrengthKey() {
            if (rbStrengthMild.value) return "mild";
            if (rbStrengthBoost.value) return "boost";
            return "normal";
        }

        function buildPreviewSignature() {
            return [
                getSelectedIndex(),
                slAmount.value,
                getStrengthKey(),
                chkPreview.value ? 1 : 0
            ].join("|");
        }

        function updatePreview(force) {
            if (!chkPreview.value) return;

            var signature = buildPreviewSignature();
            if (signature === lastPreviewSignature) return;

            var now = new Date().getTime();
            if (!force && (now - lastPreviewUpdateTime) < PREVIEW_THROTTLE_MS) {
                return;
            }
            lastPreviewUpdateTime = now;

            var amount = slAmount.value / 100;

            removePreview();

            var typeKey = getTypeKey(getSelectedIndex());
            var strengthKey = getStrengthKey();
            var rect = makeDistortRect(typeKey, amount, strengthKey);
            applySmartFreeDistort(doc.selection, { distortRect: rect });
            app.redraw();
            previewed = true;
            lastPreviewSignature = signature;
        }

        function removePreview() {
            if (previewed) {
                app.undo();
                app.redraw();
                previewed = false;
            }
            lastPreviewSignature = null;
        }

        function isTrapezoidSelected(idx) {
            return idx >= 0 && idx <= 3;
        }

        function isShearSelected(idx) {
            return idx >= 4 && idx <= 7;
        }

        function updateAmountEnabled() {
            var idx = getSelectedIndex();
            var amountEnabled = idx <= 7;
            var strengthEnabled = idx <= 7;
            var boostEnabled = idx <= 7;

            pnlAmount.enabled = amountEnabled;
            pnlStrength.enabled = strengthEnabled;
            rbStrengthNormal.enabled = strengthEnabled;
            rbStrengthMild.enabled = strengthEnabled;
            rbStrengthBoost.enabled = boostEnabled;
        }

        // イベントハンドラ
        slAmount.onChanging = function () {
            updateAmountText();
            updatePreview(false);
        };

        slAmount.onChange = function () {
            updateAmountText();
            updatePreview(true);
        };

        chkPreview.onClick = function () {
            if (chkPreview.value) {
                updatePreview(true);
            } else {
                removePreview();
            }
        };

        rbStrengthNormal.onClick = function () {
            updatePreview(true);
        };

        rbStrengthMild.onClick = function () {
            updatePreview(true);
        };

        rbStrengthBoost.onClick = function () {
            updatePreview(true);
        };

        for (var r = 0; r < rbTypes.length; r++) {
            rbTypes[r].onClick = function () {
                for (var k = 0; k < rbTypes.length; k++) {
                    if (rbTypes[k] !== this) rbTypes[k].value = false;
                }
                updateAmountEnabled();
                updateAmountText();
                updatePreview(true);
            };
        }

        updateAmountText();
        updateAmountEnabled();

        // ダイアログ表示
        var dialogResult = dlg.show();

        if (dialogResult != 1) {
            removePreview();
            return null;
        }

        // OK時：プレビュー中なら一度戻してから正式適用
        removePreview();

        var amount = slAmount.value / 100;

        return {
            typeIndex: getSelectedIndex(),
            amount: amount,
            strengthKey: getStrengthKey()
        };
    }

    // =========================================================
    // map dialog selection -> preset key
    // =========================================================
    function getTypeKey(index) {
        switch (index) {
            case 0: return "trap_narrow_top";
            case 1: return "trap_wide_top";
            case 2: return "trap_narrow_bottom";
            case 3: return "trap_wide_bottom";
            case 4: return "shear_right";
            case 5: return "shear_left";
            case 6: return "shear_up";
            case 7: return "shear_down";
            case 8: return "tri_bl";
            case 9: return "tri_br";
            case 10: return "tri_tl";
            case 11: return "tri_tr";
            case 12: return "diagonal_right";
            case 13: return "diagonal_left";
            default: return "trap_narrow_top";
        }
    }

    // =========================================================
    // run
    // =========================================================
    var result = showDialog();
    if (!result) return;

    var typeKey = getTypeKey(result.typeIndex);
    var rect = makeDistortRect(typeKey, result.amount, result.strengthKey);

    applySmartFreeDistort(doc.selection, {
        distortRect: rect
    });

})();
