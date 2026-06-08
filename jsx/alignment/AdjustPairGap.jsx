#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

# ペア間隔調整スクリプト

オブジェクトの水平方向の間隔を指定値にそろえます。2つのモードがあります。

- 自動ペア認識：選択オブジェクトを最も近いもの同士でペアにし、各ペアの間隔をそろえます
- グループ：選択した各グループの中身を左端順に並べ、隣り合う間隔をすべて指定値にそろえます（2個に限らず、3個以上もカスケードで分配）

選択がすべてグループのときは起動時に自動で「グループ」、それ以外は「自動ペア認識」が選ばれます（ダイアログで切り替え可能）。

## 使い方

- 間隔をそろえたいオブジェクト、またはグループを選択して実行
- ダイアログで「モード」「固定オブジェクト（左 / 右）」「間隔」を指定（間隔は↑↓キーで±1、Shiftで±10、Optionで±0.1）
- 「固定オブジェクト」で基準にする側を選びます。グループモードでは、固定した端（左 or 右）を起点に残りをカスケード配置します
- 「プレビュー境界」をオンにすると、線幅や効果を含む見た目の境界で間隔を測ります
- 設定はダイアログを閉じずにライブプレビューで確認できます

## 注意

- 自動ペア認識：選択数が奇数のときは末尾の1つがペアにならず、そのまま残ります
- グループ：子が2個未満のグループ、およびグループ以外のオブジェクトは対象外です

*/

// =========================================
// バージョン / Version
// =========================================
var SCRIPT_VERSION = "v1.1.0";

// =========================================
// ユーザー設定 / User settings
// =========================================
var DEFAULT_GAP = 30; // 平均間隔を測れないときのフォールバック（pt）/ Fallback gap when no average can be measured (pt)

// =========================================
// ローカライズ / Localization
// =========================================
var LABELS = {
    dialogTitle: { ja: "ペア間隔の調整", en: "Adjust Pair Gap" },
    mode: { ja: "モード", en: "Mode" },
    modeGroup: { ja: "グループ", en: "Group" },
    modeAutoPair: { ja: "自動ペア認識", en: "Auto Pair Detection" },
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

/* GroupItem の直下オブジェクトを配列で返す。グループ以外は null。
   Return a GroupItem's direct children as an array; null for non-groups. */
function getGroupChildren(item) {
    if (item.typename !== "GroupItem") return null;
    var children = [];
    for (var i = 0; i < item.pageItems.length; i++) {
        children.push(item.pageItems[i]);
    }
    return children;
}

/* 選択した各グループの一番左・一番右を取り出してペアにする / Pair the leftmost and
   rightmost object inside each selected group.
   ペア認識の代わりにグループ単位で組む。左右は geometricBounds の左端 x で判定し、
   どちらを固定するかは固定オブジェクト（左/右）ラジオに従う（applySpacing 側で処理）。
   Built per group instead of nearest-neighbor pairing; which side stays fixed is
   decided by the Fixed Object (left/right) radio, handled in applySpacing. */
function createGroupPairs(selectedItems) {
    var pairs = [];
    for (var i = 0; i < selectedItems.length; i++) {
        var children = getGroupChildren(selectedItems[i]);
        // グループでない、または子が2個未満なら間隔を調整できない / Need a group with 2+ children
        if (!children || children.length < 2) continue;

        // グループは2つに限らない。含まれる全オブジェクトを members として持ち、
        // applySpacing 側で左端順に等間隔へ分配する（固定側を基準に配置）。
        // A group may hold any number of objects; keep them all as members and let
        // applySpacing distribute them with a uniform gap, anchored at the fixed side.
        pairs.push({ members: children });
    }
    return pairs;
}

/* 選択オブジェクトの現在の水平間隔の平均（pt）を求める。測れない場合は null。
   モードに合わせて測る：自動ペア認識は各ペアの間隔、グループは各グループ内の隣接間隔。
   常に geometricBounds（幾何境界）基準。Average current horizontal gap (pt) of the
   selection, matching the mode; null if nothing measurable. Always geometric bounds. */
function computeAverageGap(selectedItems, mode) {
    var gaps = [];

    if (mode === "group") {
        // 各グループ内：左端順に並べて隣り合う間隔を測る / Adjacent gaps inside each group
        for (var i = 0; i < selectedItems.length; i++) {
            var children = getGroupChildren(selectedItems[i]);
            if (!children || children.length < 2) continue;
            var sorted = children.slice(0);
            sorted.sort(function (a, b) { return a.geometricBounds[0] - b.geometricBounds[0]; });
            for (var j = 1; j < sorted.length; j++) {
                // 次の左端 - 前の右端 / next left edge - previous right edge
                gaps.push(sorted[j].geometricBounds[0] - sorted[j - 1].geometricBounds[2]);
            }
        }
    } else {
        // 自動ペア認識：各ペアの間隔を測る / Gap of each nearest pair
        var pairs = createNearestPairs(selectedItems);
        for (var p = 0; p < pairs.length; p++) {
            var boundsA = pairs[p].a.geometricBounds;
            var boundsB = pairs[p].b.geometricBounds;
            var leftBounds = (boundsA[0] < boundsB[0]) ? boundsA : boundsB;
            var rightBounds = (boundsA[0] < boundsB[0]) ? boundsB : boundsA;
            gaps.push(rightBounds[0] - leftBounds[2]);
        }
    }

    if (gaps.length === 0) return null;
    var sum = 0;
    for (var g = 0; g < gaps.length; g++) sum += gaps[g];
    return sum / gaps.length;
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

    if (selectedItems.length < 1) {
        alert(getLocalizedText('alertSelectTwo'));
        return;
    }

    // ペアはモード（自動ペア認識 / グループ）に応じて後で組み直すため、ここでは器だけ用意する。
    // 自動ペア認識は最近傍で偶数個を組むので奇数なら末尾が1つ余るが、createNearestPairs が
    // 黙って捨てる（プレビューで何も動かないだけ）。グループモードは選択数の偶奇に依らない。
    // Pairs are (re)built per mode later; declare the holder here. Odd selections in
    // auto mode simply drop the last item; group mode is independent of the count.
    var objectPairs = [];

    // 選択がすべてグループなら既定でグループモードにする（それ以外は自動ペア認識）。
    // Default to group mode when every selected object is a group; otherwise auto pair.
    var selectionIsGroupsOnly = true;
    for (var gi = 0; gi < selectedItems.length; gi++) {
        if (selectedItems[gi].typename !== "GroupItem") { selectionIsGroupsOnly = false; break; }
    }
    var defaultMode = selectionIsGroupsOnly ? "group" : "auto";

    // 間隔の初期値は選択オブジェクトの現在の平均間隔。測れなければ DEFAULT_GAP を使う。
    // 負（重なり）の場合は 0 にクランプ。Initial gap = current average gap of the
    // selection (clamped to >= 0); fall back to DEFAULT_GAP when nothing measurable.
    var measuredGap = computeAverageGap(selectedItems, defaultMode);
    var initialGapPoints = (measuredGap !== null) ? Math.max(0, measuredGap) : DEFAULT_GAP;

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

    /* グループ内の全オブジェクトを左端順に並べ、隣り合う間隔を gap にそろえる。
       固定側（左/右）のオブジェクトは動かさず、そこを起点にカスケードで再配置する。
       Distribute all objects in a group left-to-right with a uniform gap, anchored at
       the fixed side (the fixed-side object stays; the rest cascade from it). */
    function distributeGroup(pair, fixedSide, gapInPoints, useVisible) {
        var members = pair.members;
        var cachedBounds = useVisible ? pair.memberVis : pair.memberGeo;
        var count = members.length;

        // 左端の昇順に並べたインデックス / Indices ordered by left edge ascending
        var order = [];
        for (var i = 0; i < count; i++) order.push(i);
        order.sort(function (x, y) { return cachedBounds[x][0] - cachedBounds[y][0]; });

        if (fixedSide === "right") {
            // 右端を固定し、右から左へ配置 / Anchor the rightmost; walk leftward
            var nextLeft = cachedBounds[order[count - 1]][0]; // 右端オブジェクトの左端（不動）
            for (var r = count - 2; r >= 0; r--) {
                var idxR = order[r];
                var desiredRight = nextLeft - gapInPoints;
                var dxR = desiredRight - cachedBounds[idxR][2];
                if (dxR !== 0) { members[idxR].translate(dxR, 0); appliedMoves.push({ obj: members[idxR], dx: dxR }); }
                nextLeft = cachedBounds[idxR][0] + dxR; // この要素の新しい左端 / its new left edge
            }
        } else {
            // 左端を固定し、左から右へ配置 / Anchor the leftmost; walk rightward
            var prevRight = cachedBounds[order[0]][2]; // 左端オブジェクトの右端（不動）
            for (var f = 1; f < count; f++) {
                var idxF = order[f];
                var desiredLeft = prevRight + gapInPoints;
                var dxF = desiredLeft - cachedBounds[idxF][0];
                if (dxF !== 0) { members[idxF].translate(dxF, 0); appliedMoves.push({ obj: members[idxF], dx: dxF }); }
                prevRight = cachedBounds[idxF][2] + dxF; // この要素の新しい右端 / its new right edge
            }
        }
    }

    /* 設定値で各ペアの間隔を調整する / Apply the gap to every pair */
    function applySpacing(fixedSide, gapInPoints, boundsType) {
        var useVisible = (boundsType === "visibleBounds");
        for (var k = 0; k < objectPairs.length; k++) {
            var pair = objectPairs[k];

            if (pair.members) {
                // グループモード：全オブジェクトを等間隔に分配（固定側を基準）
                // Group mode: distribute all objects with a uniform gap, anchored at the fixed side
                distributeGroup(pair, fixedSide, gapInPoints, useVisible);
                continue;
            }

            // 自動ペア認識：2オブジェクトの間隔を調整 / Auto mode: adjust the gap of a pair
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

    /* 各ペアの元の境界をキャッシュする / Cache original bounds of every pair.
       適用時は常に元位置へ巻き戻してから計算するので、ここで一度取れば使い回せる。 */
    function cachePairBounds(pairs) {
        for (var p = 0; p < pairs.length; p++) {
            var pair = pairs[p];
            if (pair.members) {
                // グループ：全メンバーの境界をキャッシュ / Group: cache every member's bounds
                pair.memberGeo = [];
                pair.memberVis = [];
                for (var m = 0; m < pair.members.length; m++) {
                    pair.memberGeo.push(pair.members[m].geometricBounds);
                    pair.memberVis.push(pair.members[m].visibleBounds);
                }
            } else {
                pair.geometricA = pair.a.geometricBounds;
                pair.visibleA = pair.a.visibleBounds;
                pair.geometricB = pair.b.geometricBounds;
                pair.visibleB = pair.b.visibleBounds;
            }
        }
    }

    /* モードに応じてペアを組み直す / Rebuild pairs for the given mode ("auto" | "group").
       境界キャッシュは必ず元位置で取るため、先にプレビューを巻き戻してから組む。
       Always revert the preview first so cached bounds reflect original positions. */
    function buildPairs(mode) {
        undoPreview();
        var pairs = (mode === "group")
            ? createGroupPairs(selectedItems)
            : createNearestPairs(selectedItems);
        cachePairBounds(pairs);
        objectPairs = pairs;
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================
    /* ダイアログを生成して表示し、OK されたら true を返す / Build, show dialog; return true on OK */
    function showSettingsDialog() {
        var dialog = new Window("dialog", getLocalizedText('dialogTitle') + ' ' + SCRIPT_VERSION);
        dialog.orientation = "column";
        dialog.alignChildren = "fill";

        // モード / Mode（ラジオ縦並び）
        var modePanel = dialog.add("panel", undefined, getLocalizedText('mode'));
        setupPanel(modePanel, 6);
        modePanel.alignChildren = ["left", "top"]; // ラジオは左寄せで縦並び / Stack radios, left-aligned
        var modeGroupRadio = modePanel.add("radiobutton", undefined, getLocalizedText('modeGroup'));       // グループ
        var modeAutoPairRadio = modePanel.add("radiobutton", undefined, getLocalizedText('modeAutoPair')); // 自動ペア認識
        // 既定モードは選択内容で決める：グループのみ→グループ、それ以外→自動ペア認識。
        // Default mode follows the selection: groups only → group, otherwise → auto pair.
        modeGroupRadio.value = selectionIsGroupsOnly;
        modeAutoPairRadio.value = !selectionIsGroupsOnly;

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
        var defaultSpacingDisplay = Math.round((initialGapPoints / pointsPerUnit) * 100) / 100;
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

        /* 現在のモードを取得する / Get the current mode ("group" | "auto") */
        function getMode() {
            return modeGroupRadio.value ? "group" : "auto";
        }

        /* モードを切り替えてペアを組み直し、プレビューを更新する / Switch mode, rebuild pairs, refresh */
        function onModeChange() {
            buildPairs(getMode());
            refreshPreview();
        }

        /* 現在固定する側を取得する / Get the currently fixed side */
        function getFixedSide() {
            return leftRadio.value ? "left" : "right";
        }

        /* 入力値を pt に換算して取得する / Get the gap in pt from the input */
        function getSpacingInPoints() {
            var value = parseFloat(spacingInput.text);
            if (isNaN(value)) { value = initialGapPoints / pointsPerUnit; }
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
        modeGroupRadio.onClick = onModeChange;
        modeAutoPairRadio.onClick = onModeChange;
        leftRadio.onClick = refreshPreview;
        rightRadio.onClick = refreshPreview;
        spacingInput.onChanging = refreshPreview;
        previewBoundsCheckbox.onClick = refreshPreview;

        // ボタン（Mac 規約：Cancel → OK）/ Buttons (Mac order: Cancel → OK)
        var buttonRow = dialog.add("group");
        buttonRow.alignment = "right";
        buttonRow.add("button", undefined, getLocalizedText('cancel'), { name: "cancel" });
        buttonRow.add("button", undefined, "OK", { name: "ok" });

        // ダイアログ表示時に既定モードでペアを組んで初回プレビュー（同期側 undo を避けて onShow から起動）
        // Build pairs for the default mode, then run the first preview (from onShow to avoid sync undo)
        dialog.onShow = function () {
            buildPairs(getMode());
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
