#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

# ペア間隔調整スクリプト

選択したオブジェクトを最も近いもの同士でペアにし、各ペアの水平方向の間隔を指定値にそろえます。

## 使い方

- 間隔をそろえたいオブジェクトを偶数個（2つずつのペア）選択して実行
- ダイアログで「固定オブジェクト（左 / 右）」と「間隔」を指定（間隔は↑↓キーで±1、Shiftで±10、Optionで±0.1）
- 「プレビュー境界」をオンにすると、線幅や効果を含む見た目の境界で間隔を測ります
- 設定はダイアログを閉じずにライブプレビューで確認できます

## 注意

- 選択数が奇数の場合は処理しません

*/

// =========================================
// バージョン / Version
// =========================================
var SCRIPT_VERSION = "v1.0.0";

// =========================================
// ユーザー設定 / User settings
// =========================================
var DEFAULT_GAP = 30; // 既定の間隔（pt）/ Default gap (pt)

// =========================================
// ローカライズ / Localization
// =========================================
var LABELS = {
    dialogTitle: { ja: "ペア間隔の調整", en: "Adjust Pair Gap" },
    fixedSide: { ja: "固定オブジェクト", en: "Fixed Object" },
    sideLeft: { ja: "左", en: "Left" },
    sideRight: { ja: "右", en: "Right" },
    spacing: { ja: "間隔", en: "Gap" },
    options: { ja: "オプション", en: "Options" },
    previewBounds: { ja: "プレビュー境界", en: "Preview Bounds" },
    cancel: { ja: "キャンセル", en: "Cancel" },
    // ツールチップ / Tooltips
    tipFixedSide: {
        ja: "選んだ側のオブジェクトは動かさず、もう一方を移動して間隔をそろえます。",
        en: "Keeps the chosen side in place and moves the other object to set the gap."
    },
    tipSpacing: {
        ja: "ペアの間隔。↑↓キーで±1、Shiftで±10、Optionで±0.1。",
        en: "Gap between paired objects. Arrow keys ±1, Shift ±10, Option ±0.1."
    },
    tipPreviewBounds: {
        ja: "オンで線幅や効果を含む見た目の境界、オフで幾何境界を基準に間隔を測ります。",
        en: "On: measure by visible bounds (incl. stroke/effects). Off: geometric bounds."
    },
    // アラート / Alerts
    alertNoDocument: {
        ja: "ドキュメントが開かれていません。",
        en: "No document is open."
    },
    alertSelectTwo: {
        ja: "オブジェクトを2つ以上選択してください。",
        en: "Select at least two objects."
    },
    alertOddCount: {
        ja: "選択されたオブジェクトの数が奇数です。\n2つずつのペア（偶数個）になるように選択してください。",
        en: "The number of selected objects is odd.\nSelect an even number so they form pairs."
    }
    // OK はローカライズしない / "OK" is not localized
};

var currentLanguage = ($.locale && $.locale.indexOf("ja") === 0) ? "ja" : "en";

/* キーからロケールに応じたラベルを返す / Return a localized label by key */
function getLocalizedText(key) {
    var entry = LABELS[key];
    return entry ? (entry[currentLanguage] || entry.en) : key;
}

/* コロン付きラベル（日本語は全角、英語は半角）/ Label with colon (full-width JA, half-width EN) */
function labelText(key) {
    return getLocalizedText(key) + (currentLanguage === 'ja' ? '：' : ':');
}

// =========================================
// 単位 / Units
// =========================================
// ルーラー単位を取得し、1単位あたりの pt 係数を求める / Resolve ruler unit and pt factor
var UNIT_LABELS = ["in", "mm", "pt", "pc", "cm", "Q", "px"];
var UNIT_FACTORS = [72, 72 / 25.4, 1, 12, 72 / 2.54, 72 / 25.4 * 0.25, 1];
var rulerType = app.preferences.getIntegerPreference("rulerType");
if (rulerType < 0 || rulerType > 6) { rulerType = 2; } // 不明時は pt / Fallback to pt
var rulerUnitLabel = UNIT_LABELS[rulerType];
var pointsPerUnit = UNIT_FACTORS[rulerType]; // 1単位 = pointsPerUnit pt

// =========================================
// ユーティリティ / Utility
// =========================================
/* テキストフィールドで↑↓キーによる増減を有効化する / Enable arrow-key value stepping
   onUpdate を渡すと値変更後に呼ぶ（プレビュー更新用）/ onUpdate runs after each change */
function changeValueByArrowKey(editText, onUpdate) {
    editText.addEventListener("keydown", function (event) {
        // ↑↓キー以外（Enter・数字入力など）は通常処理に任せる。
        // ここで毎回 onUpdate を呼ぶと、Enter 等でもプレビューが走り不安定になる。
        // Only handle Up/Down; let other keys (Enter, digits...) fall through.
        if (event.keyName !== "Up" && event.keyName !== "Down") return;

        var value = Number(editText.text);
        if (isNaN(value)) return;

        var keyboard = ScriptUI.environment.keyboardState;
        var isUp = (event.keyName === "Up");

        if (keyboard.shiftKey) {
            // Shift：±10、10の倍数にスナップ / Shift: ±10 snapped to multiples of 10
            var step = 10;
            value = isUp ? Math.ceil((value + 1) / step) * step
                         : Math.floor((value - 1) / step) * step;
        } else if (keyboard.altKey) {
            // Option：±0.1 / Option: ±0.1
            value = isUp ? value + 0.1 : value - 0.1;
        } else {
            // 通常：±1 / Default: ±1
            value = isUp ? value + 1 : value - 1;
        }

        if (value < 0) value = 0;

        // 丸め（Option は小数第1位、その他は整数）/ Round (Option: 1 decimal, else integer)
        value = keyboard.altKey ? Math.round(value * 10) / 10 : Math.round(value);

        event.preventDefault();
        editText.text = value;

        // プレビューを更新 / Refresh live preview
        if (typeof onUpdate === "function") {
            onUpdate();
        } else {
            editText.notify("onChanging");
        }
    });
}

/* パネルの余白と間隔 / Panel margins and spacing */
var PANEL_MARGINS = [16, 20, 16, 12];
var PANEL_SPACING = 8;

/* パネルの共通設定 / Apply shared panel layout */
function setupPanel(panel, spacing) {
    panel.orientation = "column";
    panel.alignChildren = ["fill", "top"];
    panel.alignment = "fill";
    panel.margins = PANEL_MARGINS;
    panel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
}

/* グループの共通設定（orientation は呼び出し側で指定）/ Apply shared group layout (orientation passed in) */
function setupGroup(group, orientation, spacing) {
    group.orientation = orientation || "column";
    group.alignChildren = ["fill", "top"];
    group.alignment = "fill";
    group.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
}

// =========================================
// ペアリング / Pairing
// =========================================
/* 選択オブジェクトを最も近いもの同士でペアにする / Pair nearest objects.
   ペアリングは常に geometricBounds（中心）で判定する。
   「プレビュー境界」を ON にしても間隔計算の基準が変わるだけで、
   ペアの組み合わせ自体は変わらない。
   Pairing always uses geometricBounds (centers); the Preview Bounds option
   only changes how the gap is measured, never which objects are paired. */
function createNearestPairs(selectedItems) {
    /* オブジェクトの中心座標を取得する（常に幾何境界）/ Object center (always geometric bounds) */
    function getObjectCenter(item) {
        var bounds = item.geometricBounds;
        return {
            x: (bounds[0] + bounds[2]) / 2,
            y: (bounds[1] + bounds[3]) / 2
        };
    }

    /* 2点間の距離を計算する / Distance between two points */
    function getPointDistance(point1, point2) {
        var dx = point1.x - point2.x;
        var dy = point1.y - point2.y;
        return Math.sqrt(dx * dx + dy * dy);
    }

    // 中心座標付きの作業リスト / Working list with centers
    var remainingItems = [];
    for (var i = 0; i < selectedItems.length; i++) {
        remainingItems.push({
            obj: selectedItems[i],
            center: getObjectCenter(selectedItems[i])
        });
    }

    var pairs = [];
    while (remainingItems.length > 0) {
        var currentItem = remainingItems.shift();
        var nearestIndex = -1;
        var nearestDistance = Infinity;
        for (var j = 0; j < remainingItems.length; j++) {
            var distance = getPointDistance(currentItem.center, remainingItems[j].center);
            if (distance < nearestDistance) {
                nearestDistance = distance;
                nearestIndex = j;
            }
        }
        if (nearestIndex !== -1) {
            var partnerItem = remainingItems.splice(nearestIndex, 1)[0];
            pairs.push({ a: currentItem.obj, b: partnerItem.obj });
        }
    }
    return pairs;
}

(function () {
    if (app.documents.length === 0) {
        alert(getLocalizedText('alertNoDocument'));
        return;
    }

    var activeDocument = app.activeDocument;

    // doc.selection を配列にコピーしておく（後の選択変更や undo の影響を受けないように）
    // Copy doc.selection into a plain array (immune to later selection changes / undo)
    var liveSelection = activeDocument.selection;
    // selection が無効・未選択のケースを先に弾く（Illustrator では稀に null になる）
    // Guard against an invalid / empty selection (selection can be null in rare cases)
    if (!liveSelection || liveSelection.length === 0) {
        alert(getLocalizedText('alertSelectTwo'));
        return;
    }
    var selectedItems = [];
    for (var i = 0; i < liveSelection.length; i++) {
        selectedItems.push(liveSelection[i]);
    }

    if (selectedItems.length < 2) {
        alert(getLocalizedText('alertSelectTwo'));
        return;
    }

    // 選択数が奇数の場合はエラーを出す / Reject an odd number of objects
    if (selectedItems.length % 2 !== 0) {
        alert(getLocalizedText('alertOddCount'));
        return;
    }

    var objectPairs = createNearestPairs(selectedItems);

    // 各オブジェクトの元の境界を一度だけキャッシュする。
    // プレビューは適用前に必ず元位置へ巻き戻すため、適用時は常に元位置 →
    // キャッシュした境界で計算でき、プレビューごとの境界再取得を避けられる（高速化）。
    // Cache original bounds once. Preview always reverts to the original position
    // before re-applying, so cached bounds stay valid and we avoid re-reading
    // bounds on every preview (the slow part).
    for (var p = 0; p < objectPairs.length; p++) {
        var pair = objectPairs[p];
        pair.geometricA = pair.a.geometricBounds;
        pair.visibleA = pair.a.visibleBounds;
        pair.geometricB = pair.b.geometricBounds;
        pair.visibleB = pair.b.visibleBounds;
    }

    // =========================================
    // プレビュー / Live preview
    //   巻き戻しは移動量の逆適用で行う（app.undo() は使わない）。
    //   app.undo() をキーボードイベント内で同期実行すると Illustrator が
    //   不安定になり得るため、ここでは記録した移動を translate で打ち消す。
    //   Revert by reversing recorded moves — never app.undo(), which is
    //   unstable when called synchronously inside keyboard event handlers.
    // =========================================
    var appliedMoves = []; // 適用済みの移動 / Applied moves: { obj, dx }

    /* 記録した移動を逆向きに適用して元に戻す / Reverse recorded moves */
    function undoPreview() {
        for (var m = appliedMoves.length - 1; m >= 0; m--) {
            appliedMoves[m].obj.translate(-appliedMoves[m].dx, 0);
        }
        appliedMoves = [];
    }

    /* 設定値で各ペアの間隔を調整する / Apply the gap to every pair */
    function applySpacing(fixedSide, gapInPoints, boundsType) {
        var useVisible = (boundsType === "visibleBounds");
        for (var k = 0; k < objectPairs.length; k++) {
            var pair = objectPairs[k];
            // キャッシュ済みの元の境界を使用 / Use cached original bounds
            var boundsA = useVisible ? pair.visibleA : pair.geometricA;
            var boundsB = useVisible ? pair.visibleB : pair.geometricB;

            // X座標（左端）で左右を判定 / Decide left/right by x position
            var leftObject, rightObject, leftBounds, rightBounds;
            if (boundsA[0] < boundsB[0]) {
                leftObject = pair.a; leftBounds = boundsA;
                rightObject = pair.b; rightBounds = boundsB;
            } else {
                leftObject = pair.b; leftBounds = boundsB;
                rightObject = pair.a; rightBounds = boundsA;
            }

            var dx;
            if (fixedSide === "right") {
                // 右を固定：左オブジェクトを動かす / Fix right, move left
                dx = (rightBounds[0] - gapInPoints) - leftBounds[2];
                if (dx !== 0) { leftObject.translate(dx, 0); appliedMoves.push({ obj: leftObject, dx: dx }); }
            } else {
                // 左を固定：右オブジェクトを動かす / Fix left, move right
                dx = (leftBounds[2] + gapInPoints) - rightBounds[0];
                if (dx !== 0) { rightObject.translate(dx, 0); appliedMoves.push({ obj: rightObject, dx: dx }); }
            }
        }
    }

    /* 直前のプレビューを巻き戻してから再適用する / Re-run preview from scratch */
    function runPreview(fixedSide, gapInPoints, boundsType) {
        undoPreview();
        applySpacing(fixedSide, gapInPoints, boundsType);
        app.redraw();
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================
    /* ダイアログを生成して表示し、OK されたら true を返す / Build, show dialog; return true on OK */
    function showSettingsDialog() {
        var dialog = new Window("dialog", getLocalizedText('dialogTitle') + ' ' + SCRIPT_VERSION);
        dialog.orientation = "column";
        dialog.alignChildren = "fill";

        // 固定オブジェクト / Fixed object
        var fixedSidePanel = dialog.add("panel", undefined, getLocalizedText('fixedSide'));
        setupPanel(fixedSidePanel, 6);

        var fixedSideRow = fixedSidePanel.add("group");
        setupGroup(fixedSideRow, "row");
        // ラジオは広げず、左寄せでまとめる / Keep radios at natural width, packed left
        fixedSideRow.alignment = "left";
        fixedSideRow.alignChildren = ["left", "center"];
        var leftRadio = fixedSideRow.add("radiobutton", undefined, getLocalizedText('sideLeft'));   // 左 / Left
        var rightRadio = fixedSideRow.add("radiobutton", undefined, getLocalizedText('sideRight')); // 右 / Right
        rightRadio.value = true; // 既定：右を固定（左が動く）/ Default: fix right
        leftRadio.helpTip = getLocalizedText('tipFixedSide');
        rightRadio.helpTip = getLocalizedText('tipFixedSide');

        // オプション / Options
        var optionsPanel = dialog.add("panel", undefined, getLocalizedText('options'));
        setupPanel(optionsPanel, 6);

        // 行：間隔（ラベル＋入力＋単位）/ Row: gap (label + input + unit)
        var spacingRow = optionsPanel.add("group");
        setupGroup(spacingRow, "row");
        spacingRow.alignment = "left"; // 広げず左寄せ / Keep at natural width, packed left
        spacingRow.add("statictext", undefined, labelText('spacing'));
        var defaultSpacingDisplay = Math.round((DEFAULT_GAP / pointsPerUnit) * 100) / 100;
        var spacingInput = spacingRow.add("edittext", undefined, String(defaultSpacingDisplay));
        spacingInput.characters = 4;
        spacingInput.helpTip = getLocalizedText('tipSpacing');
        spacingRow.add("statictext", undefined, rulerUnitLabel);
        changeValueByArrowKey(spacingInput, refreshPreview); // ↑↓キーで増減＋プレビュー更新 / Arrow keys + preview

        // チェックボックス：プレビュー境界（groupに入れて左添え）/ Checkbox in a left-aligned group
        var previewBoundsGroup = optionsPanel.add("group");
        previewBoundsGroup.orientation = "row";
        previewBoundsGroup.alignment = "left";   // パネル内で左添え / Left within panel
        previewBoundsGroup.margins = [0, 5, 0, 0]; // 上マージン5 / Top margin 5
        var previewBoundsCheckbox = previewBoundsGroup.add("checkbox", undefined, getLocalizedText('previewBounds'));
        previewBoundsCheckbox.value = false; // OFF=幾何境界 / ON=プレビュー境界
        previewBoundsCheckbox.helpTip = getLocalizedText('tipPreviewBounds');

        /* 現在固定する側を取得する / Get the currently fixed side */
        function getFixedSide() {
            return leftRadio.value ? "left" : "right";
        }

        /* 入力値を pt に換算して取得する / Get the gap in pt from the input */
        function getSpacingInPoints() {
            var value = parseFloat(spacingInput.text);
            if (isNaN(value)) { value = DEFAULT_GAP / pointsPerUnit; }
            return value * pointsPerUnit;
        }

        /* チェックに応じた境界プロパティ名を返す / Bounds property per checkbox */
        function getBoundsType() {
            return previewBoundsCheckbox.value ? "visibleBounds" : "geometricBounds";
        }

        /* 現在の設定でプレビューを更新する / Refresh preview with current settings */
        function refreshPreview() {
            runPreview(getFixedSide(), getSpacingInPoints(), getBoundsType());
        }

        // 設定変更でライブプレビュー / Update preview on change
        leftRadio.onClick = refreshPreview;
        rightRadio.onClick = refreshPreview;
        spacingInput.onChanging = refreshPreview;
        previewBoundsCheckbox.onClick = refreshPreview;

        // ボタン（Mac 規約：Cancel → OK）/ Buttons (Mac order: Cancel → OK)
        var buttonRow = dialog.add("group");
        buttonRow.alignment = "right";
        buttonRow.add("button", undefined, getLocalizedText('cancel'), { name: "cancel" });
        buttonRow.add("button", undefined, "OK", { name: "ok" });

        // ダイアログ表示時に初回プレビュー（同期側 undo を避けて onShow から起動）
        // Initial preview from onShow (avoid synchronous undo)
        dialog.onShow = function () {
            refreshPreview();
        };

        return dialog.show() === 1;
    }

    if (showSettingsDialog()) {
        // OK：プレビューをそのまま確定 / Keep the applied preview
    } else {
        // キャンセル：プレビューを巻き戻す / Revert the preview
        undoPreview();
        app.redraw();
    }
})();
