#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);
/*
  グリッド複製ツール / Duplicate in Grid
  Version: SCRIPT_VERSION（変数にて管理）

  概要 / Overview
  - 選択中のオブジェクトを、指定した「繰り返し数 × 繰り返し数」のグリッドで複製します。
  - 間隔は現在の定規単位（rulerType）で入力できます。内部では pt に変換して処理します。
  - ダイアログの値変更（入力・↑↓キー・Shift/Option併用）でライブプレビューが更新されます。
  - ローカライズ（日本語/英語）とバージョン表記に対応。ダイアログタイトルにバージョンを表示します。

  操作 / Usage
  1) 複製したいオブジェクトを選択してスクリプトを実行
  2) ダイアログで「繰り返し数」「間隔」を入力
     - ↑↓ で ±1、Shift+↑↓ で ±10、Option(Alt)+↑↓ で ±0.1
     - 「繰り返し数」は常に整数として扱われます
  3) OK で確定（プレビューは自動消去）、キャンセルで中止（プレビューを消去）

  変更履歴 / Changelog
  - 2025-10-23: ローカライズ（JA/EN）、SCRIPT_VERSION 導入、ダイアログタイトルにバージョン表示を追加
  - 2025-10-23: rulerType に基づく単位ラベル表示と、単位→pt 変換を実装
  - 2025-10-23: ライブプレビュー（_preview レイヤー）を追加、再描画の安定化（app.redraw）
  - 2025-10-23: ↑↓／Shift／Option による数値増減、繰り返し数の整数化を実装
  - 2025-10-23: UI 調整（ラベル幅統一・右揃え、OK を右側、透明度・位置補正）
*/

/* バージョン / Version */
var SCRIPT_VERSION = "v1.2";

/* 言語判定 / Locale detection */
function getCurrentLang() {
    return ($.locale && $.locale.toLowerCase().indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

/* ラベル定義 / Label definitions (JA/EN) */
var LABELS = {
    dialogTitle: {
        ja: "グリッド状に複製",
        en: "Duplicate in Grid"
    },
    repeatCount: {
        ja: "繰り返し数",
        en: "Count"
    },
    repeatCountH: {
        ja: "横",
        en: "Count (X)"
    },
    repeatCountV: {
        ja: "縦",
        en: "Count (Y)"
    },
    gap: {
        ja: "間隔",
        en: "Gap"
    },
    unitFmt: {
        ja: "間隔（{unit}）:",
        en: "Gap ({unit}):"
    },
    ok: {
        ja: "OK",
        en: "OK"
    },
    cancel: {
        ja: "キャンセル",
        en: "Cancel"
    },
    alertNoDoc: {
        ja: "ドキュメントが開かれていません。",
        en: "No document is open."
    },
    alertNoSel: {
        ja: "オブジェクトを選択してください。",
        en: "Please select an object."
    },
    alertCountInvalid: {
        ja: "繰り返し数は1以上の整数を入力してください。",
        en: "Enter an integer count of 1 or more."
    },
    alertGapInvalid: {
        ja: "間隔は数値で入力してください。",
        en: "Enter a numeric gap value."
    },
    directionTitle: {
        ja: "方向",
        en: "Direction"
    },
    dirRight: {
        ja: "右",
        en: "Right"
    },
    dirLeft: {
        ja: "左",
        en: "Left"
    },
    verticalTitle: {
        ja: "縦方向",
        en: "Vertical"
    },
    dirUp: {
        ja: "上",
        en: "Up"
    },
    dirDown: {
        ja: "下",
        en: "Down"
    },
};

/* ラベル取得関数 / Label resolver */
function L(key, params) {
    var table = LABELS[key];
    var text = (table && table[lang]) ? table[lang] : key;
    if (params) {
        for (var k in params) {
            if (params.hasOwnProperty(k)) {
                text = text.replace(new RegExp("\\{" + k + "\\}", "g"), params[k]);
            }
        }
    }
    return text;
}

/* テキストフィールドの数値操作 / Arrow key increment–decrement for numeric fields */

function changeValueByArrowKey(edittext) {
    edittext.addEventListener("keydown", function(e) {
        var v = Number(edittext.text);
        if (isNaN(v)) return;

        var keyboard = ScriptUI.environment.keyboardState;
        var delta = 1;

        if (keyboard.shiftKey) {
            delta = 10;
            if (e.keyName == "Up") {
                v = Math.ceil((v + 1) / delta) * delta;
                e.preventDefault();
            } else if (e.keyName == "Down") {
                v = Math.floor((v - 1) / delta) * delta;
                if (v < 0) v = 0;
                e.preventDefault();
            }
        } else if (keyboard.altKey) {
            delta = 0.1;
            if (e.keyName == "Up") {
                v += delta;
                e.preventDefault();
            } else if (e.keyName == "Down") {
                v -= delta;
                e.preventDefault();
            }
        } else {
            delta = 1;
            if (e.keyName == "Up") {
                v += delta;
                e.preventDefault();
            } else if (e.keyName == "Down") {
                v -= delta;
                if (v < 0) v = 0;
                e.preventDefault();
            }
        }

        if (keyboard.altKey) {
            v = Math.round(v * 10) / 10;
        } else {
            v = Math.round(v);
        }

        if (edittext.isInteger) {
            v = Math.round(v);
        }

        edittext.text = v;
        if (typeof edittext.onChanging === "function") {
            try {
                edittext.onChanging();
            } catch (_) {}
        }
    });
}

// 単位コードとラベルのマップ
var unitLabelMap = {
    0: "in",
    1: "mm",
    2: "pt",
    3: "pica",
    4: "cm",
    5: "Q/H",
    6: "px",
    7: "ft/in",
    8: "m",
    9: "yd",
    10: "ft"
};
// 現在の単位ラベルを取得
function getCurrentUnitLabel() {
    var unitCode = app.preferences.getIntegerPreference("rulerType");
    return unitLabelMap[unitCode] || "pt";
}
// 単位コードを取得
function getCurrentUnitCode() {
    return app.preferences.getIntegerPreference("rulerType");
}
// 単位→ポイント変換（座標系はpt基準）
function unitToPoints(unitCode, value) {
    var PT_PER_IN = 72;
    var MM_TO_PT = PT_PER_IN / 25.4;
    switch (unitCode) {
        case 0: // in
            return value * PT_PER_IN;
        case 1: // mm
            return value * MM_TO_PT;
        case 2: // pt
            return value;
        case 3: // pica
            return value * 12; // 1 pica = 12 pt
        case 4: // cm
            return value * (MM_TO_PT * 10);
        case 5: // Q/H （Q=0.25mm換算）
            return value * (MM_TO_PT * 0.25);
        case 6: // px（Illustratorは1px≒1ptとして扱う）
            return value;
        case 7: // ft/in（複合だが便宜上 in と同等に扱う）
            return value * PT_PER_IN;
        case 8: // m
            return value * (MM_TO_PT * 1000);
        case 9: // yd
            return value * (PT_PER_IN * 36); // 1yd=36in
        case 10: // ft
            return value * (PT_PER_IN * 12); // 1ft=12in
        default:
            return value; // fallback pt
    }
}

/* プレビュー用ユーティリティ / Preview utilities */
function getPreviewLayer(doc) {
    var name = "_preview";
    var lyr;
    try {
        lyr = doc.layers.getByName(name);
    } catch (e) {
        lyr = doc.layers.add();
        lyr.name = name;
    }
    lyr.visible = true;
    lyr.locked = false;
    try {
        lyr.zOrder(ZOrderMethod.BRINGTOFRONT);
    } catch (_) {}
    return lyr;
}

function clearPreview(doc) {
    try {
        var lyr = doc.layers.getByName("_preview");
        while (lyr.pageItems.length > 0) {
            lyr.pageItems[0].remove();
        }
        app.redraw();
    } catch (e) {
        // layer may not exist; ignore
    }
}

function buildPreview(doc, sourceItem, rows, cols, gapX, gapY, width, height, direction, vDirection) {
    var lyr = getPreviewLayer(doc);
    clearPreview(doc);
    for (var row = 0; row < rows; row++) {
        for (var col = 0; col < cols; col++) {
            if (row === 0 && col === 0) continue; // keep original as-is
            var dup = sourceItem.duplicate(lyr, ElementPlacement.PLACEATBEGINNING);
            var offsetX = (width + gapX) * col;
            if (direction === "left") offsetX = -offsetX;
            dup.left = sourceItem.left + offsetX;
            var offsetY = (height + gapY) * row;
            dup.top = (vDirection === "up") ?
                (sourceItem.top + offsetY) :
                (sourceItem.top - offsetY);
        }
    }
    app.redraw();
}

/* ダイアログの作成とプレビュー制御
   / Build dialog and preview wiring */
function showDialog(doc, sourceItem, width, height) {
    var dlg = new Window("dialog", L("dialogTitle") + " " + SCRIPT_VERSION);
    var offsetX = 300;
    var dialogOpacity = 0.98;

    function shiftDialogPosition(dlg, offsetX, offsetY) {
        dlg.onShow = function() {
            var currentX = dlg.location[0];
            var currentY = dlg.location[1];
            dlg.location = [currentX + offsetX, currentY + offsetY];
        };
    }

    function setDialogOpacity(dlg, opacityValue) {
        dlg.opacity = opacityValue;
    }

    setDialogOpacity(dlg, dialogOpacity);
    shiftDialogPosition(dlg, offsetX, 0);
    dlg.orientation = "column";
    dlg.alignChildren = "fill";

    var unitCode = getCurrentUnitCode();
    var unitLabel = getCurrentUnitLabel();

    // パネル：繰り返し数 / Panel: Repeat Count
    var repeatPanel = dlg.add("panel", undefined, L("repeatCount"));
    repeatPanel.orientation = "row"; // 2カラム（左右）
    repeatPanel.alignChildren = "top";
    repeatPanel.margins = [20, 20, 20, 10];
    repeatPanel.spacing = 20;

    // 左右カラム用のグループ
    var repeatLeftCol = repeatPanel.add("group");
    repeatLeftCol.orientation = "column";
    repeatLeftCol.alignChildren = "left";

    var repeatRightCol = repeatPanel.add("group");
    repeatRightCol.orientation = "column";
    repeatRightCol.alignChildren = "left";
    /* 繰り返し数パネル内での縦位置センター / Vertically center within the Repeat panel */
    repeatRightCol.alignment = ["left", "center"];

    // 繰り返し数（横）
    var repeatXGroup = repeatLeftCol.add("group");
    var repeatLabelX = repeatXGroup.add("statictext", undefined, L("repeatCountH") + ":");
    var countXInput = repeatXGroup.add("edittext", undefined, "2");
    countXInput.characters = 4;
    countXInput.isInteger = true;
    changeValueByArrowKey(countXInput);

    // 繰り返し数（縦）
    var repeatYGroup = repeatLeftCol.add("group");
    var repeatLabelY = repeatYGroup.add("statictext", undefined, L("repeatCountV") + ":");
    var countYInput = repeatYGroup.add("edittext", undefined, "2");
    countYInput.characters = 4;
    countYInput.isInteger = true;
    changeValueByArrowKey(countYInput);

    // 連動チェック / Link checkbox
    var linkGroup = repeatRightCol.add("group");
    /* 右カラム内で上下センター / Center vertically inside the right column */
    linkGroup.alignment = ["left", "center"];
    var linkCheck = linkGroup.add("checkbox", undefined, (lang === "ja" ? "連動" : "Link X and Y"));

    linkCheck.value = true; // ← これを追加（デフォルトON）

    // 連動の有効/無効と同期
    function syncCounts() {
        if (linkCheck.value) {
            countYInput.enabled = false;
            countYInput.text = countXInput.text;
            if (typeof countYInput.onChanging === "function") {
                try {
                    countYInput.onChanging();
                } catch (_) {}
            }
        } else {
            countYInput.enabled = true;
        }
    }

    // チェック切り替えで同期＋プレビュー
    linkCheck.onClick = function() {
        syncCounts();
        applyPreview();
    };

    syncCounts(); // 起動時に一度状態反映

    /* 間隔入力（現在の定規単位で表示し、内部では pt に変換）
   / Gap input (shown in current ruler units, converted to pt internally) */

    var gapGroup = dlg.add("group");
    var gapLabel = gapGroup.add("statictext", undefined, L("unitFmt", {
        unit: unitLabel
    }));
    gapLabel.preferredSize = [80, 20];
    gapLabel.justify = "right";
    var gapInput = gapGroup.add("edittext", undefined, "10");
    gapInput.characters = 4;
    changeValueByArrowKey(gapInput);

    /* 方向パネル（横・縦の展開方向を指定）
       / Direction panel (horizontal & vertical placement) */
    var dirPanel = dlg.add("panel", undefined, L("directionTitle"));
    dirPanel.orientation = "column";
    dirPanel.alignChildren = "left";
    dirPanel.margins = [20, 15, 20, 10];

    // 横方向ラベル＆ラジオ / Horizontal direction
    var hGroup = dirPanel.add("group");
    hGroup.orientation = "row";
    hGroup.alignChildren = "left";
    var hLabel = hGroup.add("statictext", undefined, (lang === "ja" ? "横方向" : "Horizontal"));
    var dirRight = hGroup.add("radiobutton", undefined, L("dirRight"));
    var dirLeft = hGroup.add("radiobutton", undefined, L("dirLeft"));
    dirRight.value = true; // default: right
    dirRight.onClick = applyPreview;
    dirLeft.onClick = applyPreview;


    // 縦方向ラベル＆ラジオ / Vertical direction
    var vGroup = dirPanel.add("group");
    vGroup.orientation = "row";
    vGroup.alignChildren = "left";

    var vLabel = vGroup.add("statictext", undefined, (lang === "ja" ? "縦方向" : "Vertical"));

    var dirUp = vGroup.add("radiobutton", undefined, L("dirUp"));
    var dirDown = vGroup.add("radiobutton", undefined, L("dirDown"));
    dirDown.value = true; // デフォルトは「下」

    // クリックでプレビュー更新
    dirUp.onClick = applyPreview;
    dirDown.onClick = applyPreview;

    function applyPreview() {
        var cx = parseInt(countXInput.text, 10);
        var cy = parseInt(countYInput.text, 10);
        var g = parseFloat(gapInput.text);
        if (isNaN(cx) || cx < 1) return;
        if (isNaN(cy) || cy < 1) return;
        if (isNaN(g)) return;

        var gpt = unitToPoints(unitCode, g);
        var hDir = dirRight.value ? "right" : "left";
        var vDir = dirUp && dirUp.value ? "up" : "down";
        buildPreview(doc, sourceItem, cy, cx, gpt, gpt, width, height, hDir, vDir);
    }
    // X変更：連動ONならYへミラー、その後プレビュー
    countXInput.onChanging = function() {
        if (linkCheck.value) {
            countYInput.text = countXInput.text;
        }
        applyPreview();
    };
    countXInput.onChange = function() {
        if (linkCheck.value) {
            countYInput.text = countXInput.text;
        }
        applyPreview();
    };

    // Y変更：普通にプレビュー（※ onChanging も必須）
    countYInput.onChanging = applyPreview;
    countYInput.onChange = applyPreview;

    // Gap
    gapInput.onChanging = applyPreview;
    gapInput.onChange = applyPreview;

    var btnGroup = dlg.add("group");
    btnGroup.alignment = "center";
    var cancelBtn = btnGroup.add("button", undefined, L("cancel"), {
        name: "cancel"
    });
    var okBtn = btnGroup.add("button", undefined, L("ok"));

    var result = null;
    okBtn.onClick = function() {
        var cx = parseInt(countXInput.text, 10);
        var cy = parseInt(countYInput.text, 10);
        var g = parseFloat(gapInput.text);
        if (isNaN(cx) || cx < 1) {
            alert(L("alertCountInvalid"));
            return;
        }
        if (isNaN(cy) || cy < 1) {
            alert(L("alertCountInvalid"));
            return;
        }
        if (isNaN(g)) {
            alert(L("alertGapInvalid"));
            return;
        }
        result = {
            cols: cx,
            rows: cy,
            gap: unitToPoints(unitCode, g), // store as pt
            direction: dirRight.value ? "right" : "left",
            vDirection: (dirUp && dirUp.value) ? "up" : "down"
        };
        clearPreview(doc);
        dlg.close();
    };
    cancelBtn.onClick = function() {
        clearPreview(doc);
        dlg.close();
    };

    dlg.onShow = (function(origOnShow) {
        return function() {
            if (typeof origOnShow === "function") {
                try {
                    origOnShow();
                } catch (_) {}
            }
            // Let Illustrator settle the dialog placement first, then draw
            $.sleep(0);
            applyPreview();
            // ダイアログ表示時に「繰り返し数」フィールドをアクティブに
            countXInput.active = true;
        };
    })(dlg.onShow);

    dlg.show();
    return result;
}

/* 選択オブジェクトをグリッド状に複製（横2×縦2） */
/* メイン処理：検証→ダイアログ→複製実行
   / Main flow: validate → dialog → duplication */
function main() {
    if (app.documents.length === 0) {
        alert(L("alertNoDoc"));
        return;
    }
    var doc = app.activeDocument;
    if (doc.selection.length === 0) {
        alert(L("alertNoSel"));
        return;
    }

    var sel = doc.selection;
    var bounds = sel[0].visibleBounds; // [左, 上, 右, 下]
    var width = bounds[2] - bounds[0];
    var height = bounds[1] - bounds[3];

    var settings = showDialog(doc, sel[0], width, height);
    if (!settings) return;

    var rows = settings.rows;
    var cols = settings.cols;
    var gapX = settings.gap;
    var gapY = settings.gap;
    var direction = settings.direction || "right";
    var vDirection = settings.vDirection || "down";

    for (var row = 0; row < rows; row++) {
        for (var col = 0; col < cols; col++) {
            if (row === 0 && col === 0) continue; // 元の位置はスキップ
            var dup = sel[0].duplicate();
            var offsetX = (width + gapX) * col;
            if (direction === "left") offsetX = -offsetX;
            dup.left = sel[0].left + offsetX;
            var offsetY = (height + gapY) * row;
            dup.top = (vDirection === "up") ?
                (sel[0].top + offsetY) :
                (sel[0].top - offsetY);
        }
    }
}
main();