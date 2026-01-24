#target illustrator
//#targetengine "RulesBetweenObjects"
 #targetengine "RulesBetweenObjects"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
  DrawLinesBetween.jsx

  選択したオブジェクト（図形／テキスト）を上から順に並べ、間に水平の罫線（ケイ線）を描画します。
  / Sort selected objects (shapes/text) from top to bottom and draw horizontal rules between them.

  主なポイント / Key features:
  - 入力単位は「線（strokeUnits）」の環境設定に追従（表示ラベル＆内部pt換算）
    / Follows "strokeUnits" preference (UI label + internal pt conversion)
  - 「延長」で左右方向に罫線を伸縮（+で延長、-で短縮）
    / "Extension" expands/contracts the rule length horizontally (+ extends, - shrinks)
  - 線端（なし／丸型／突出）を選択可能
    / Select line cap (Butt / Round / Projecting)
  - ケイ線の長さ：共通（全体幅）／オブジェクトに合わせる（行ペア幅）
    / Rule length: Common (overall width) / Match Objects (per-row-pair width)
  - テキストは複製→アウトライン化した境界で計算し、終了時に一時オブジェクトを削除
    / Text is measured via duplicate→outline bounds, then temporary items are removed
  - 左右に並ぶ要素は同一行として束ねて扱う（行グルーピング）
    / Horizontally aligned items are treated as one row (row grouping)
  - 実行後、作成した罫線を選択状態に
    / Select created rules after execution
  - 最後に使った設定を保存して次回起動時に復元（線幅／延長／線端など）
    / Persist last-used settings and restore on next run (stroke/extension/cap etc.)
*/

// ----------------------------------------
// Version / Localization
// ----------------------------------------

var SCRIPT_VERSION = "v1.0";

function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

/* 日英ラベル定義 / Japanese-English label definitions */
var LABELS = {
    dialogTitle: {
        ja: "オブジェクト間に罫線",
        en: "Rules Between Objects"
    },
    lineWidth: {
        ja: "線幅",
        en: "Stroke"
    },
    extension: {
        ja: "延長",
        en: "Extension"
    },
    lineCap: {
        ja: "線端",
        en: "Line Cap"
    },
    capButt: {
        ja: "なし",
        en: "Butt"
    },
    capRound: {
        ja: "丸型線端",
        en: "Round"
    },
    capProjecting: {
        ja: "突出線端",
        en: "Projecting"
    },
    ruleLength: {
        ja: "ケイ線の長さ",
        en: "Rule Length"
    },
    ruleLengthObject: {
        ja: "オブジェクトに合わせる",
        en: "Match Objects"
    },
    ruleLengthCommon: {
        ja: "共通",
        en: "Common"
    },
    ok: {
        ja: "OK",
        en: "OK"
    },
    cancel: {
        ja: "キャンセル",
        en: "Cancel"
    },
    alertNeedPositiveStroke: {
        ja: "線幅は正の数値を入力してください。",
        en: "Please enter a positive number for stroke."
    },
    alertNeedNumberExtension: {
        ja: "延長は数値を入力してください。",
        en: "Please enter a number for extension."
    },
    alertNoDoc: {
        ja: "ドキュメントが開かれていません。",
        en: "No document is open."
    },
    alertNeedSelection2: {
        ja: "2つ以上のオブジェクトを選択してください。",
        en: "Please select at least two objects."
    }
};

function L(key) {
    if (!LABELS[key]) return key;
    return LABELS[key][lang] || LABELS[key].en || key;
}

// ----------------------------------------
// Settings persistence (last used values)
// ----------------------------------------

var SETTINGS_KEY = "RulesBetweenObjectsSettings";

function sid(s) {
    return stringIDToTypeID(s);
}

function loadSettings() {
    try {
        var d = app.getCustomOptions(SETTINGS_KEY);
        return {
            lineWeightPt: d.getReal(sid("lineWeightPt")),
            marginPt: d.getReal(sid("marginPt")),
            capIndex: d.getInteger(sid("capIndex"))
        };
    } catch (e) {
        return null;
    }
}

function saveSettings(lineWeightPt, marginPt, capIndex) {
    try {
        var d = new ActionDescriptor();
        d.putReal(sid("lineWeightPt"), lineWeightPt);
        d.putReal(sid("marginPt"), marginPt);
        d.putInteger(sid("capIndex"), capIndex);
        app.putCustomOptions(SETTINGS_KEY, d, true);
    } catch (e) {
        // ignore
    }
}

function formatNumberForUI(n) {
    // UI表示用に適度に丸める（単位に依存しすぎない）
    var v = Math.round(n * 100) / 100;
    // -0 を防ぐ
    if (Math.abs(v) < 0.000001) v = 0;
    return String(v);
}

var DEFAULT_LINE_WEIGHT = 0.25; // デフォルト線幅（pt）
var DEFAULT_MARGIN = 0; // デフォルト延長量（pt）左右にどれだけ伸ばすか

// 単位（strokeUnits）ユーティリティ / Units (strokeUnits) utilities

// 単位コード → ラベル（Q/H分岐あり）
var unitMap = {
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

function getUnitLabel(code, prefKey) {
    if (code === 5) {
        // Q/H は設定キーによってラベルが変わる（ここでは strokeUnits は H 扱い）
        var hKeys = {
            "text/asianunits": true,
            "rulerType": true,
            "strokeUnits": true
        };
        return hKeys[prefKey] ? "H" : "Q";
    }
    return unitMap[code] || "pt";
}

function getStrokeUnitsCode() {
    try {
        return app.preferences.getIntegerPreference("strokeUnits");
    } catch (e) {
        // 取得できない環境向けフォールバック
        return 2; // pt
    }
}

function getCurrentStrokeUnitLabel() {
    var code = getStrokeUnitsCode();
    return getUnitLabel(code, "strokeUnits");
}

// 単位コード → pt換算係数（value * factor = pt）
function getUnitToPtFactor(code, prefKey) {
    // 基本単位
    if (code === 2) return 1; // pt
    if (code === 0) return 72; // in
    if (code === 1) return 72 / 25.4; // mm
    if (code === 4) return 72 / 2.54; // cm
    if (code === 3) return 12; // pica (1pica = 12pt)
    if (code === 6) return 1; // px（Illustratorでは概ねpt相当として扱う）
    if (code === 10) return 72 * 12; // ft
    if (code === 9) return 72 * 36; // yd
    if (code === 8) return (72 / 2.54) * 100; // m
    if (code === 5) {
        // Q/H：1Q=0.25mm、1H=0.25mm（扱いは同じ）
        return (72 / 25.4) * 0.25;
    }

    // ft/in は入力形式が複合になり得るため、ここではpt扱いにフォールバック
    return 1;
}

function toPt(value, code, prefKey) {
    return value * getUnitToPtFactor(code, prefKey);
}

function fromPt(ptValue, code, prefKey) {
    var f = getUnitToPtFactor(code, prefKey);
    if (!f) return ptValue;
    return ptValue / f;
}

// ダイアログ外観ユーティリティ / Dialog appearance utilities

var DIALOG_OFFSET_X = 300;
var DIALOG_OFFSET_Y = 0;
var DIALOG_OPACITY = 0.98;

function shiftDialogPosition(dlg, offsetX, offsetY) {
    dlg.onShow = function () {
        try {
            var currentX = dlg.location[0];
            var currentY = dlg.location[1];
            dlg.location = [currentX + offsetX, currentY + offsetY];
        } catch (e) {
            // location が取得できない環境向けフォールバック
        }
    };
}

function setDialogOpacity(dlg, opacityValue) {
    try {
        dlg.opacity = opacityValue;
    } catch (e) {
        // opacity 未対応環境向けフォールバック
    }
}

function changeValueByArrowKey(editText) {
    editText.addEventListener("keydown", function(event) {
        var value = Number(editText.text);
        if (isNaN(value)) return;

        var keyboard = ScriptUI.environment.keyboardState;
        var delta = 1;

        if (keyboard.shiftKey) {
            delta = 10;
            // Shiftキー押下時は10の倍数にスナップ
            if (event.keyName == "Up") {
                value = Math.ceil((value + 1) / delta) * delta;
                event.preventDefault();
            } else if (event.keyName == "Down") {
                value = Math.floor((value - 1) / delta) * delta;
                if (value < 0) value = 0;
                event.preventDefault();
            }
        } else if (keyboard.altKey) {
            delta = 0.1;
            // Optionキー押下時は0.1単位で増減
            if (event.keyName == "Up") {
                value += delta;
                event.preventDefault();
            } else if (event.keyName == "Down") {
                value -= delta;
                event.preventDefault();
            }
        } else {
            delta = 1;
            if (event.keyName == "Up") {
                value += delta;
                event.preventDefault();
            } else if (event.keyName == "Down") {
                value -= delta;
                if (value < 0) value = 0;
                event.preventDefault();
            }
        }

        if (keyboard.altKey) {
            // 小数第1位までに丸め
            value = Math.round(value * 10) / 10;
        } else {
            // 整数に丸め
            value = Math.round(value);
        }

        editText.text = value;
    });
}

function showDialog() {
    var dlg = new Window('dialog', L('dialogTitle') + ' ' + SCRIPT_VERSION);
    setDialogOpacity(dlg, DIALOG_OPACITY);
    shiftDialogPosition(dlg, DIALOG_OFFSET_X, DIALOG_OFFSET_Y);
    dlg.orientation = 'column';
    dlg.alignChildren = ['fill', 'top'];
    var saved = loadSettings();
    var strokeUnitCodeForUI = getStrokeUnitsCode();
    var strokeUnitLabel = getUnitLabel(strokeUnitCodeForUI, "strokeUnits");

    // 既定値（pt）→ 現在の strokeUnits 表示値へ変換
    var defaultLineWeightUI = fromPt(DEFAULT_LINE_WEIGHT, strokeUnitCodeForUI, "strokeUnits");
    var defaultMarginUI = fromPt(DEFAULT_MARGIN, strokeUnitCodeForUI, "strokeUnits");

    // 保存済みがあればそれを優先
    if (saved) {
        if (typeof saved.lineWeightPt === "number") defaultLineWeightUI = fromPt(saved.lineWeightPt, strokeUnitCodeForUI, "strokeUnits");
        if (typeof saved.marginPt === "number") defaultMarginUI = fromPt(saved.marginPt, strokeUnitCodeForUI, "strokeUnits");
    }

    var lineGroup = dlg.add('group');
    lineGroup.add('statictext', undefined, L('lineWidth'));
    var lineWidthInput = lineGroup.add('edittext', undefined, formatNumberForUI(defaultLineWeightUI));
    lineWidthInput.characters = 6;
    changeValueByArrowKey(lineWidthInput);
    lineGroup.add('statictext', undefined, '(' + strokeUnitLabel + ')');

    var marginGroup = dlg.add('group');
    marginGroup.add('statictext', undefined, L('extension'));
    var marginInput = marginGroup.add('edittext', undefined, formatNumberForUI(defaultMarginUI));
    marginInput.characters = 6;
    changeValueByArrowKey(marginInput);
    marginGroup.add('statictext', undefined, '(' + strokeUnitLabel + ')');

    var capPanel = dlg.add('panel', undefined, L('lineCap'));
    capPanel.orientation = 'column';
    capPanel.alignChildren = ['left', 'top'];
    capPanel.margins = [15, 20, 15, 10];

    var capGroup = capPanel.add('group');
    capGroup.orientation = 'column';
    capGroup.alignChildren = ['left', 'center'];

    var capNoneRadio = capGroup.add('radiobutton', undefined, L('capButt'));
    var capRoundRadio = capGroup.add('radiobutton', undefined, L('capRound'));
    var capProjectingRadio = capGroup.add('radiobutton', undefined, L('capProjecting'));

    // デフォルト：なし（BUTT）/ Default: Butt
    capNoneRadio.value = true;
    // 保存済みの線端があれば反映 / Apply saved line cap selection
    if (saved && typeof saved.capIndex === "number") {
        if (saved.capIndex === 1) {
            capRoundRadio.value = true;
        } else if (saved.capIndex === 2) {
            capProjectingRadio.value = true;
        } else {
            capNoneRadio.value = true;
        }
    }

    // --- Rule Length Panel 追加 ---
    var lengthPanel = dlg.add('panel', undefined, L('ruleLength'));
    lengthPanel.orientation = 'column';
    lengthPanel.alignChildren = ['left', 'top'];
    lengthPanel.margins = [15, 20, 15, 10];

    var lengthGroup = lengthPanel.add('group');
    lengthGroup.orientation = 'column';
    lengthGroup.alignChildren = ['left', 'center'];

    var lengthObjectRadio = lengthGroup.add('radiobutton', undefined, L('ruleLengthObject'));
    var lengthCommonRadio = lengthGroup.add('radiobutton', undefined, L('ruleLengthCommon'));

    // デフォルト：共通
    lengthCommonRadio.value = true;

    var btnGroup = dlg.add('group');
    btnGroup.alignment = ['right', 'center'];
    btnGroup.orientation = 'row';
    var cancelBtn = btnGroup.add('button', undefined, L('cancel'), {name: 'cancel'});
    var okBtn = btnGroup.add('button', undefined, L('ok'), {name: 'ok'});
    okBtn.active = true;          // Enterキー＝OK
    dlg.defaultElement = okBtn;  // 環境依存対策

    if (dlg.show() !== 1) {
        return null;
    }

    var lineWeightVal = Number(lineWidthInput.text);
    if (isNaN(lineWeightVal) || lineWeightVal <= 0) {
        alert(L('alertNeedPositiveStroke'));
        return null;
    }

    var marginVal = Number(marginInput.text);
    if (isNaN(marginVal)) {
        alert(L('alertNeedNumberExtension'));
        return null;
    }
    var strokeUnitCode = strokeUnitCodeForUI;
    var lineWeightPt = toPt(lineWeightVal, strokeUnitCode, "strokeUnits");
    var marginPt = toPt(marginVal, strokeUnitCode, "strokeUnits");

    var lineCap = StrokeCap.BUTTENDCAP;
    if (capRoundRadio.value) {
        lineCap = StrokeCap.ROUNDENDCAP;
    } else if (capProjectingRadio.value) {
        lineCap = StrokeCap.PROJECTINGENDCAP;
    }

    var capIndex = 0;
    if (capRoundRadio.value) {
        capIndex = 1;
    } else if (capProjectingRadio.value) {
        capIndex = 2;
    }

    // 最後に使った設定を保存（内部はptで保持）/ Persist last-used settings (stored in pt)
    saveSettings(lineWeightPt, marginPt, capIndex);

    // --- ruleLengthMode 取得 ---
    var ruleLengthMode = lengthObjectRadio.value ? "object" : "common";

    return {
        lineWeight: lineWeightPt,
        margin: marginPt,
        lineCap: lineCap,
        ruleLengthMode: ruleLengthMode
    };
}

function main() {
    var settings = showDialog();
    if (settings === null) return;

    var lineWeight = settings.lineWeight;
    var margin = settings.margin;
    var lineCap = settings.lineCap;
    var ruleLengthMode = settings.ruleLengthMode || "object";

    // 線の色設定（デフォルトは黒 K=100）/ Line color setting (default black K=100)
    var lineColor = new CMYKColor();
    lineColor.cyan = 0;
    lineColor.magenta = 0;
    lineColor.yellow = 0;
    lineColor.black = 100;

    // 線の長さの調整（0ならオブジェクトと同じ幅、正の数なら長く、負の数なら短く）/ Adjust line length (0 = same as objects, positive = longer, negative = shorter)
    var padding = margin; 

    // メイン処理 / Main process
    if (app.documents.length === 0) {
        alert(L('alertNoDoc'));
        return;
    }

    var doc = app.activeDocument;
    var sel = doc.selection;

    // テキストは複製→アウトライン化してから bounds を参照し、最後に複製物を削除
    // / For TextFrame: duplicate -> create outlines -> use outlined bounds -> remove duplicates at the end
    var tempOutlinedItems = [];

    function outlineTextFramesInContainer(container) {
        // container 配下の TextFrame を複製アウトライン化して置換（bounds 安定用）
        // / Outline all TextFrames inside the container (for stable bounds)
        try {
            if (!container || !container.textFrames || container.textFrames.length === 0) return;

            // textFrames はライブコレクションになり得るので、いったん配列化
            var tfs = [];
            for (var i = 0; i < container.textFrames.length; i++) {
                tfs.push(container.textFrames[i]);
            }

            for (var j = 0; j < tfs.length; j++) {
                var tf = tfs[j];
                try {
                    // createOutline は新しい GroupItem を作成して返す
                    var outlined = tf.createOutline();
                    // 元テキストを削除（アウトライン側を残す）
                    try { tf.remove(); } catch (e1) {}
                    // 目視されないように隠す（可能なら）
                    try { outlined.hidden = true; } catch (e2) {}
                } catch (e) {
                    // 変換できない TextFrame は無視
                }
            }
        } catch (eOuter) {}
    }

    function makeBoundsProxy(item) {
        if (!item) return item;

        if (item.typename === "TextFrame") {
            try {
                var dupText = item.duplicate();
                var outlined = dupText.createOutline(); // GroupItem (outlined paths) is created in the document
                try { dupText.remove(); } catch (e1) {}
                try { outlined.hidden = true; } catch (e2) {}
                outlineTextFramesInContainer(outlined);
                tempOutlinedItems.push(outlined);
                return outlined;
            } catch (e) {
                // フォールバック：変換できない場合は元のテキストを使う
                return item;
            }
        }

        if (item.typename === "GroupItem") {
            try {
                var dupGroup = item.duplicate();
                // グループ内テキストもアウトライン化してから bounds を参照する
                outlineTextFramesInContainer(dupGroup);
                try { dupGroup.hidden = true; } catch (e3) {}
                tempOutlinedItems.push(dupGroup);
                return dupGroup;
            } catch (eG) {
                return item;
            }
        }

        return item;
    }

    // 同じ行（左右に並ぶ要素）を1グループとして扱うためのユーティリティ
    // / Treat horizontally aligned items as a single row group
    var ROW_OVERLAP_RATIO = 0.5; // 縦方向の重なり率のしきい値（0.0–1.0）

    function getVerticalOverlapRatio(a, b) {
        // visibleBounds: [left, top, right, bottom]
        var top = Math.min(a.visibleBounds[1], b.visibleBounds[1]);
        var bottom = Math.max(a.visibleBounds[3], b.visibleBounds[3]);
        var overlap = Math.min(a.visibleBounds[1], b.visibleBounds[1]) - Math.max(a.visibleBounds[3], b.visibleBounds[3]);
        if (overlap <= 0) return 0;

        var heightA = a.visibleBounds[1] - a.visibleBounds[3];
        var heightB = b.visibleBounds[1] - b.visibleBounds[3];
        var minHeight = Math.min(heightA, heightB);
        return overlap / minHeight;
    }

    function mergeBounds(items) {
        var left = items[0].visibleBounds[0];
        var top = items[0].visibleBounds[1];
        var right = items[0].visibleBounds[2];
        var bottom = items[0].visibleBounds[3];

        for (var i = 1; i < items.length; i++) {
            left = Math.min(left, items[i].visibleBounds[0]);
            top = Math.max(top, items[i].visibleBounds[1]);
            right = Math.max(right, items[i].visibleBounds[2]);
            bottom = Math.min(bottom, items[i].visibleBounds[3]);
        }

        return {
            visibleBounds: [left, top, right, bottom]
        };
    }

    function groupItemsByRow(items) {
        var rows = [];

        for (var i = 0; i < items.length; i++) {
            var placed = false;
            for (var r = 0; r < rows.length; r++) {
                // 行の代表要素（最初の1つ）と縦方向の重なりを比較
                if (getVerticalOverlapRatio(items[i], rows[r][0]) >= ROW_OVERLAP_RATIO) {
                    rows[r].push(items[i]);
                    placed = true;
                    break;
                }
            }
            if (!placed) {
                rows.push([items[i]]);
            }
        }

        // 各行を1つの bounds にまとめる
        var merged = [];
        for (var j = 0; j < rows.length; j++) {
            merged.push(mergeBounds(rows[j]));
        }
        return merged;
    }

    var proxyItems = [];
    for (var i = 0; i < sel.length; i++) {
        proxyItems.push(makeBoundsProxy(sel[i]));
    }

    // 左右に並ぶ要素を行単位でグループ化
    var sortedItems = groupItemsByRow(proxyItems);

    sortedItems.sort(function(a, b) {
        // topの値が大きいほうが上（Illustratorの座標系に依存するが通常はtopが大きい＝上）
        // ※定規の原点設定によっては逆になる場合もあるが、一般的なDTP設定を想定
        return b.visibleBounds[1] - a.visibleBounds[1];
    });

    var createdLines = []; // 生成した線を保持 / Store created lines

    // 順番に処理して間を計算
    for (var i = 0; i < sortedItems.length - 1; i++) {
        var upperObj = sortedItems[i];
        var lowerObj = sortedItems[i + 1];

        // 上のオブジェクトの下端
        var upperBottom = upperObj.visibleBounds[3];
        // 下のオブジェクトの上端
        var lowerTop = lowerObj.visibleBounds[1];

        // 中間のY座標
        var midY = (upperBottom + lowerTop) / 2;

        // 線のX座標（左右の幅）を決める
        // ruleLengthMode に応じてロジックを切り替え
        var leftX, rightX;

        if (ruleLengthMode === "common") {
            // 共通：全体で共通の左右幅（最初に計算）
            if (i === 0) {
                var commonLeft = upperObj.visibleBounds[0];
                var commonRight = upperObj.visibleBounds[2];
                for (var k = 1; k < sortedItems.length; k++) {
                    commonLeft = Math.min(commonLeft, sortedItems[k].visibleBounds[0]);
                    commonRight = Math.max(commonRight, sortedItems[k].visibleBounds[2]);
                }
                // ループ外参照用に保持
                main._commonLeft = commonLeft;
                main._commonRight = commonRight;
            }
            leftX = main._commonLeft - padding;
            rightX = main._commonRight + padding;
        } else {
            // オブジェクトに合わせる（従来挙動）
            leftX = Math.min(upperObj.visibleBounds[0], lowerObj.visibleBounds[0]) - padding;
            rightX = Math.max(upperObj.visibleBounds[2], lowerObj.visibleBounds[2]) + padding;
        }

        // 線を描画
        var pathLine = doc.pathItems.add();
        pathLine.setEntirePath([[leftX, midY], [rightX, midY]]);
        
        pathLine.filled = false;
        pathLine.stroked = true;
        pathLine.strokeWidth = lineWeight;
        pathLine.strokeColor = lineColor;
        pathLine.strokeCap = lineCap;
        createdLines.push(pathLine);
    }

    // 複製してアウトライン化した一時オブジェクトを削除 / Remove temporary outlined duplicates
    for (var t = 0; t < tempOutlinedItems.length; t++) {
        try {
            if (tempOutlinedItems[t] && tempOutlinedItems[t].typename) {
                tempOutlinedItems[t].remove();
            }
        } catch (e) {}
    }

    // 描画した線を選択状態に / Select created lines
    if (createdLines.length > 0) {
        doc.selection = null;
        for (var s = 0; s < createdLines.length; s++) {
            try {
                createdLines[s].selected = true;
            } catch (e) {}
        }
    }
}

main();
