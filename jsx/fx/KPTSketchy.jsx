#target illustrator
#targetengine "KPTSketchy"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

- 選択したオブジェクトにランダムな変形を加え、手書き・スケッチ風の見た目に整えます。
- 塗り／線の分離、角丸、パスのオフセット、ラフ効果（ギザギザ／歪曲）、グループ化に対応します。
- 変形はすべてライブ効果として適用するため、あとから編集・解除できます。
- 常駐パレットとして表示され、開いたままドキュメントを操作できます。

### 使い方

1. スクリプトを実行するとパレットが表示されます。
2. オブジェクトを選択し、各パネルで効果を調整します（変更するたびに結果が作り直されます）。
3. ［再計算］で乱数を振り直します。気に入ったらそのままにしておくだけで確定です。
4. パレットから離れる、または閉じる（× / Esc）と、その時点の結果が確定します。

### 仕様

- 角丸・オフセット・移動は、環境設定の単位（rulerType）で入力します。
- 適用順は「角丸 → パスのオフセット → 変形 → ラフ効果（歪曲 → ギザギザ）」です。
- オフセットは負方向 → 正方向の順に 2 回適用します。
- 塗りと線の両方を持つオブジェクトのみ分離できます（塗りのみ／線のみは分離しません）。
- グループは「グループ内を個別に処理」ON のとき、中の各オブジェクトへ展開します。
- クリッピングパスは処理対象から除外します。
- 値の変更・再計算では、直前の結果を app.undo() で取り消してから適用し直します。
- パレットからフォーカスが外れた時点で結果を確定し、以降は取り消しません
  （ユーザーがドキュメント側で行った操作を誤って取り消さないため）。

### 実装メモ

- 常駐パレットの app はドキュメントへの接続を失うため、DOM を触る処理は
  すべて worker 関数にまとめ、BridgeTalk でメインエンジンへ委譲します。
- worker 関数は toString() で連結し、encodeURIComponent したうえで
  eval(decodeURIComponent(...)) の形で送信します。
- worker 関数内では行コメントを使わず、必ずセミコロンで文を終えます。

---

### Overview

- Applies randomized transformations to the selected objects for a hand-drawn / sketchy look.
- Supports fill/stroke splitting, round corners, offset path, roughen (jagged / distortion) and grouping.
- Everything is applied as a live effect, so it stays editable and removable afterwards.
- Runs as a persistent palette, so the document stays usable while it is open.

### Usage

1. Run the script to show the palette.
2. Select objects and adjust the panels (the result is rebuilt on every change).
3. Press Recalculate to reroll the random values; just leave it as is to keep it.
4. Leaving the palette, or closing it (X / Esc), finalizes the current result.

### Notes

- Round corners, offset and move use the current ruler unit (rulerType).
- Effect order: Round Corners -> Offset Path -> Transform -> Roughen (Distortion -> Jagged).
- Offset Path is applied twice, negative first and then positive.
- Only objects having both a visible fill and a visible stroke can be split.
- Groups are expanded into their children when "Process group items individually" is on.
- Clipping paths are excluded from processing.
- Changing a value or recalculating undoes the previous result before applying a new one.
- Once the palette loses focus the result is finalized and never undone again,
  so that operations the user performed in the document are never reverted by mistake.

### Implementation notes

- A palette's app object loses its document connection, so every DOM operation lives in a
  worker function and is delegated to the main engine through BridgeTalk.
- Worker functions are concatenated with toString(), encoded with encodeURIComponent and
  sent as eval(decodeURIComponent(...)).
- Worker functions must avoid line comments and must terminate every statement with a semicolon.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME = "KPTSketchy"; /* スクリプト名 / script name */
var SCRIPT_VERSION = "v1.2.0"; /* バージョン / version */
var SCRIPT_AUTHOR = "Masahiro Takano (@swwwitch)"; /* 作者 / author */
var SCRIPT_RELEASED = "2026-04-14"; /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED = "2026-07-22"; /* 更新日 / last updated */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

/*
常駐エンジンに保持するパレット参照（GC回避・多重起動防止）
Palette reference kept alive in the persistent engine (prevents GC and double launch)
*/
var paletteWindow = (typeof paletteWindow !== "undefined") ? paletteWindow : null;

// =========================================
// DOM操作 worker / DOM workers (run in the main engine)
// =========================================
/*
以下の wk* 関数は BridgeTalk でメインエンジンへ送られます。
行コメントは使わず、必ずセミコロンで文を終えてください（toString で改行が失われるため）。
The wk* functions below are sent to the main engine through BridgeTalk.
Never use line comments and always terminate statements with a semicolon,
because toString() may collapse the newlines.
*/

/* ランダムな小数を返す / Return a random decimal in the range */
function wkRandomBetween(minValue, maxValue) {
    return minValue + Math.random() * (maxValue - minValue);
}

/* ライブ効果を適用（失敗時は無視） / Apply a live effect, ignoring failures */
function wkApplyEffectXml(targetItem, effectXml) {
    try {
        targetItem.applyEffect(effectXml);
    } catch (e) {}
}

/* ラフ効果のXMLを生成 / Build the Roughen effect XML */
function wkBuildRoughenEffectXml(sizePercent, detailPerInch, roundness) {
    return '<LiveEffect name="Adobe Roughen"><Dict data="' +
        'R asiz ' + sizePercent +
        ' R size ' + sizePercent +
        ' R absoluteness 0' +
        ' R dtal ' + detailPerInch +
        ' R roundness ' + roundness +
        ' "/></LiveEffect>';
}

/* 角を丸くする効果のXMLを生成 / Build the Round Corners effect XML */
function wkBuildRoundCornersEffectXml(radiusPoint) {
    return '<LiveEffect name="Adobe Round Corners"><Dict data="R radius ' + radiusPoint + ' "/></LiveEffect>';
}

/* パスのオフセット効果のXMLを生成 / Build the Offset Path effect XML */
function wkBuildOffsetPathEffectXml(offsetPoint, joinType, miterLimit) {
    return '<LiveEffect name="Adobe Offset Path"><Dict data="' +
        'R ofst ' + offsetPoint +
        ' I jntp ' + joinType +
        ' R mlim ' + miterLimit +
        ' "/></LiveEffect>';
}

/* 変形効果のXMLを生成 / Build the Transform effect XML */
function wkBuildTransformEffectXml(scalePercent, moveHorizontalPoint, moveVerticalPoint, rotateDegrees) {
    return '<LiveEffect name="Adobe Transform"><Dict data="' +
        'R scaleH_Percent ' + scalePercent +
        ' R scaleV_Percent ' + scalePercent +
        ' R scaleH_Factor ' + (scalePercent / 100) +
        ' R scaleV_Factor ' + (scalePercent / 100) +
        ' R moveH_Pts ' + moveHorizontalPoint +
        ' R moveV_Pts ' + (-moveVerticalPoint) +
        ' R rotate_Degrees ' + rotateDegrees +
        ' R rotate_Radians ' + (rotateDegrees * Math.PI / 180) +
        ' I numCopies 0' +
        ' I pinPoint 4' +
        ' B scaleLines 1' +
        ' B transformPatterns 1' +
        ' B transformObjects 1' +
        ' B reflectX 0' +
        ' B reflectY 0' +
        ' B randomize 0' +
        ' "/></LiveEffect>';
}

/* 表示される塗りを持つか / Whether the object has a visible fill */
function wkHasVisibleFill(pageItem) {
    try {
        return !!(pageItem.filled && pageItem.fillColor && pageItem.fillColor.typename !== "NoColor");
    } catch (e) {
        return false;
    }
}

/* 表示される線を持つか / Whether the object has a visible stroke */
function wkHasVisibleStroke(pageItem) {
    try {
        return !!(pageItem.stroked && pageItem.strokeColor && pageItem.strokeColor.typename !== "NoColor");
    } catch (e) {
        return false;
    }
}

/* 塗りと線の両方を持つか（分離可能か） / Whether the object can be split */
function wkCanSplitFillStroke(pageItem) {
    return wkHasVisibleFill(pageItem) && wkHasVisibleStroke(pageItem);
}

/* クリッピングパスか / Whether the item is a clipping path */
function wkIsClippingPathItem(pageItem) {
    return !!(pageItem && pageItem.typename === "PathItem" && pageItem.clipping);
}

/* グループを展開して処理対象を集める / Collect processable items, expanding groups */
function wkCollectTargetItems(pageItem, collected) {
    if (!pageItem) {
        return;
    }
    if (pageItem.typename === "GroupItem") {
        for (var i = 0; i < pageItem.pageItems.length; i++) {
            var childItem = pageItem.pageItems[i];
            if (childItem.parent === pageItem) {
                wkCollectTargetItems(childItem, collected);
            }
        }
        return;
    }
    if (!wkIsClippingPathItem(pageItem)) {
        collected.push(pageItem);
    }
}

/* 選択範囲から処理対象を取得 / Get the processing targets from the selection */
function wkGetTargetItems(targetDoc, expandsGroups) {
    var targetItems = [];
    if (!targetDoc.selection || targetDoc.selection.length === 0) {
        return targetItems;
    }
    for (var i = 0; i < targetDoc.selection.length; i++) {
        var selectedItem = targetDoc.selection[i];
        if (expandsGroups) {
            wkCollectTargetItems(selectedItem, targetItems);
        } else if (!wkIsClippingPathItem(selectedItem)) {
            targetItems.push(selectedItem);
        }
    }
    return targetItems;
}

/*
選択中のオブジェクトを配列に控える（selection は都度作られるため実体を保持）
Snapshot the selected items into a plain array, since selection is rebuilt on each access
*/
function wkSnapshotSelection(targetDoc) {
    var savedItems = [];
    try {
        for (var i = 0; i < targetDoc.selection.length; i++) {
            savedItems.push(targetDoc.selection[i]);
        }
    } catch (e) {}
    return savedItems;
}

/*
控えておいた選択を復元する。複製・移動・グループ化で選択がずれるのを防ぐ
Restore a snapshotted selection, undoing the drift caused by duplicate/move/group
*/
function wkRestoreSelection(targetDoc, savedItems) {
    try {
        targetDoc.selection = null;
    } catch (e) {}
    for (var i = 0; i < savedItems.length; i++) {
        try {
            savedItems[i].selected = true;
        } catch (e) {}
    }
}

/* 展開できる選択か（複数選択、またはグループ） / Whether the selection can be expanded */
function wkHasExpandableSelection(targetDoc) {
    try {
        if (targetDoc.selection.length > 1) {
            return true;
        }
        return targetDoc.selection.length === 1 && targetDoc.selection[0].typename === "GroupItem";
    } catch (e) {
        return false;
    }
}

/* 分離できるオブジェクトが含まれるか / Whether the selection contains a splittable object */
function wkHasSplittableTargetItem(targetDoc, expandsGroups) {
    var targetItems = wkGetTargetItems(targetDoc, expandsGroups);
    for (var i = 0; i < targetItems.length; i++) {
        if (wkCanSplitFillStroke(targetItems[i])) {
            return true;
        }
    }
    return false;
}

/* クリッピンググループの中にあるか / Whether the item lives inside a clipping group */
function wkIsInsideClippingGroup(pageItem) {
    var parentItem = null;
    try {
        parentItem = pageItem.parent;
        while (parentItem) {
            if (parentItem.typename === "GroupItem" && parentItem.clipped) {
                return true;
            }
            parentItem = parentItem.parent;
        }
    } catch (e) {}
    return false;
}

/* 塗りと線を2つのオブジェクトに分離 / Split fill and stroke into two objects */
function wkSplitIntoStrokeCopy(pageItem) {
    var strokeCopy = null;
    try {
        strokeCopy = pageItem.duplicate(pageItem, ElementPlacement.PLACEBEFORE);
        pageItem.stroked = false;
        strokeCopy.filled = false;
    } catch (e) {
        return null;
    }
    return strokeCopy;
}

/* 分離した塗りと線をグループにまとめる / Group the separated fill and stroke */
function wkGroupFillAndStroke(fillItem, strokeCopy) {
    try {
        var wrapperGroup = fillItem.parent.groupItems.add();
        wrapperGroup.move(fillItem, ElementPlacement.PLACEBEFORE);
        strokeCopy.move(wrapperGroup, ElementPlacement.PLACEATEND);
        fillItem.move(wrapperGroup, ElementPlacement.PLACEATEND);
    } catch (e) {}
}

/* 1オブジェクトに効果を重ねて適用 / Apply the whole effect stack to a single object */
function wkApplyEffectStack(targetItem, options) {
    if (options.radiusPoint > 0) {
        wkApplyEffectXml(targetItem, wkBuildRoundCornersEffectXml(options.radiusPoint));
    }
    if (options.offsetPoint !== 0) {
        wkApplyEffectXml(targetItem, wkBuildOffsetPathEffectXml(-options.offsetPoint, options.joinType, options.miterLimit));
        wkApplyEffectXml(targetItem, wkBuildOffsetPathEffectXml(options.offsetPoint, options.joinType, options.miterLimit));
    }
    if (options.usesTransform) {
        wkApplyEffectXml(targetItem, wkBuildTransformEffectXml(
            wkRandomBetween(100 - options.scaleRange, 100 + options.scaleRange),
            wkRandomBetween(-options.movePoint, options.movePoint),
            wkRandomBetween(-options.movePoint, options.movePoint),
            wkRandomBetween(-options.rotateRange, options.rotateRange)));
    }
    if (options.usesDistort) {
        wkApplyEffectXml(targetItem, wkBuildRoughenEffectXml(options.distortSize, options.distortDetail, options.roundness));
    }
    if (options.usesJagged) {
        wkApplyEffectXml(targetItem, wkBuildRoughenEffectXml(options.jaggedSize, options.jaggedDetail, options.roundness));
    }
}

/*
UIの有効・無効に使う選択状態を "|分離可否|展開可否" の形で返す
Return the selection state used for UI enabling, as "|canSplit|canExpand"
*/
function wkSelectionStateSuffix(targetDoc, expandsGroups) {
    var canSplit = wkHasSplittableTargetItem(targetDoc, expandsGroups) ? "1" : "0";
    var canExpand = wkHasExpandableSelection(targetDoc) ? "1" : "0";
    return "|" + canSplit + "|" + canExpand;
}

/* 選択状態だけを取得 / Read the selection state only */
function wkGetSelectionState(expandsGroups) {
    try {
        if (app.documents.length === 0) {
            return "NODOC";
        }
        var targetDoc = app.activeDocument;
        if (targetDoc.selection.length === 0) {
            return "NOSEL";
        }
        return "OK" + wkSelectionStateSuffix(targetDoc, expandsGroups);
    } catch (e) {
        return "ERR:" + e;
    }
}

/*
選択オブジェクトへランダム変形を適用。undoFirst が true なら前回のプレビューを取り消す
Apply the randomized effects; when undoFirst is true the previous preview is undone first
*/
function wkApplyEffects(options, undoFirst) {
    try {
        /* ドキュメント確認は app.undo() より必ず先 / The document check must precede app.undo() */
        if (app.documents.length === 0) {
            return "NODOC";
        }
        if (undoFirst) {
            app.undo();
            app.redraw();
        }
        var targetDoc = app.activeDocument;
        if (targetDoc.selection.length === 0) {
            return "NOSEL";
        }
        var targetItems = wkGetTargetItems(targetDoc, options.expandsGroups);
        if (targetItems.length === 0) {
            return "NOTARGET";
        }
        /*
        分離すると元オブジェクトの線が外れ、分離可能と判定されなくなるため、
        UIへ返す選択状態は適用前に採取しておく
        Splitting clears the original's stroke, so the state reported to the UI
        must be captured before the effects are applied
        */
        var stateSuffix = wkSelectionStateSuffix(targetDoc, options.expandsGroups);
        /*
        複製・移動・グループ化は選択を動かす。次回の再計算も同じ選択を対象にするため、
        処理前の選択を控えて最後に戻す
        Duplicating, moving and grouping all shift the selection. Snapshot it so the
        next rebuild works on the same objects
        */
        var savedSelection = wkSnapshotSelection(targetDoc);
        for (var i = 0; i < targetItems.length; i++) {
            var targetItem = targetItems[i];
            var strokeCopy = null;
            if (options.splitsFillStroke && wkCanSplitFillStroke(targetItem)) {
                strokeCopy = wkSplitIntoStrokeCopy(targetItem);
            }
            wkApplyEffectStack(targetItem, options);
            if (strokeCopy) {
                wkApplyEffectStack(strokeCopy, options);
                if (options.groupsAtEnd && !wkIsInsideClippingGroup(targetItem)) {
                    wkGroupFillAndStroke(targetItem, strokeCopy);
                }
            }
        }
        wkRestoreSelection(targetDoc, savedSelection);
        app.redraw();
        return "OK" + stateSuffix;
    } catch (e) {
        return "ERR:" + e;
    }
}

/*
送信する worker 関数の一覧（追加漏れ防止のためここに全登録）
Every worker function to be sent; register new ones here so none is forgotten
*/
var WORKER_FUNCS = [
    wkRandomBetween,
    wkApplyEffectXml,
    wkBuildRoughenEffectXml,
    wkBuildRoundCornersEffectXml,
    wkBuildOffsetPathEffectXml,
    wkBuildTransformEffectXml,
    wkHasVisibleFill,
    wkHasVisibleStroke,
    wkCanSplitFillStroke,
    wkIsClippingPathItem,
    wkCollectTargetItems,
    wkGetTargetItems,
    wkSnapshotSelection,
    wkRestoreSelection,
    wkHasExpandableSelection,
    wkHasSplittableTargetItem,
    wkIsInsideClippingGroup,
    wkSplitIntoStrokeCopy,
    wkGroupFillAndStroke,
    wkApplyEffectStack,
    wkSelectionStateSuffix,
    wkGetSelectionState,
    wkApplyEffects
];

(function() {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* 入力欄の初期値 / Default values for the input fields */
    var DEFAULT_VALUES = {
        scale: "3", /* スケールの振れ幅（%）/ scale range (%) */
        move: "0", /* 移動の振れ幅（現在の単位）/ move range (current unit) */
        rotate: "1.5", /* 回転の振れ幅（度）/ rotation range (deg) */
        radius: "1", /* 角丸の半径（現在の単位）/ corner radius (current unit) */
        offset: "1", /* オフセット量（現在の単位）/ offset amount (current unit) */
        jaggedSize: "0.3", /* ギザギザのサイズ（%）/ jagged size (%) */
        jaggedDetail: "20", /* ギザギザの詳細（/inch）/ jagged detail (per inch) */
        distortSize: "2", /* 歪曲のサイズ（%）/ distortion size (%) */
        distortDetail: "4" /* 歪曲の詳細（/inch）/ distortion detail (per inch) */
    };

    /* チェックボックスの初期状態 / Default checkbox states */
    var INITIAL_CHECKED = {
        applyEach: true, /* グループ内を個別に処理 / process group items individually */
        scale: true, /* スケール / scale */
        move: false, /* 移動 / move */
        rotate: true, /* 回転 / rotate */
        radius: true, /* 角を丸くする / round corners */
        offset: false, /* オフセット / offset path */
        jagged: true, /* ラフ効果：ギザギザ / roughen: jagged */
        distort: false /* ラフ効果：歪曲 / roughen: distortion */
    };

    var OFFSET_JOIN_TYPE = 0; /* オフセットの角の形状 0:マイター 1:ラウンド 2:ベベル / offset join 0:miter 1:round 2:bevel */
    var OFFSET_MITER_LIMIT = 4; /* オフセットのマイター制限（長さではなく比率。Illustratorの既定値）/ offset miter limit (a ratio, not a length; Illustrator's default) */
    var ROUGHEN_ROUNDNESS = 1; /* ラフ効果の丸み 0:直線的 1:滑らか / roughen roundness 0:corner 1:smooth */
    var ROUGHEN_DETAIL_MIN = 1; /* ラフ効果の詳細の下限 / minimum roughen detail */
    var BRIDGE_TIMEOUT_SEC = 10; /* BridgeTalk の同期待ち時間（秒）/ synchronous BridgeTalk timeout (sec) */
    var COMMITS_ON_DEACTIVATE = true; /* パレットから離れたら結果を確定する / finalize the result when the palette loses focus */
    var DEACTIVATE_GUARD_MS = 500; /* BridgeTalk 直後の deactivate を無視する時間（ミリ秒）/ ignore deactivate for this long after a BridgeTalk call (ms) */

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /* 表示言語を判定 / Detect the UI language */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var currentLanguage = getCurrentLang();

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "手書き・スケッチ風の調整", en: "Hand-drawn / Sketchy Look" }
        },
        panel: {
            target: { ja: "対象", en: "Target" },
            transform: { ja: "変形", en: "Transform" },
            corner: { ja: "角丸・オフセット", en: "Corners & Offset" },
            jagged: { ja: "ラフ効果：ギザギザ", en: "Roughen: Jagged" },
            distort: { ja: "ラフ効果：歪曲", en: "Roughen: Distortion" }
        },
        checkbox: {
            applyEach: { ja: "グループ内を個別に処理", en: "Process group items individually" },
            groupAtEnd: { ja: "分離後にグループ化", en: "Group after split" },
            radius: { ja: "角丸", en: "Round corners" },
            offset: { ja: "オフセット", en: "Offset path" },
            jagged: { ja: "適用", en: "Apply" },
            distort: { ja: "適用", en: "Apply" }
        },
        radio: {
            splitFillStroke: { ja: "塗り／線を分離", en: "Split fill / stroke" },
            keepFillStroke: { ja: "分離しない", en: "Do not split" }
        },
        fieldLabel: {
            scale: { ja: "倍率", en: "Scale" },
            move: { ja: "移動", en: "Move" },
            rotate: { ja: "回転", en: "Rotate" },
            size: { ja: "サイズ", en: "Size" },
            detail: { ja: "詳細", en: "Detail" }
        },
        button: {
            toggleAll: { ja: "すべてON/OFF", en: "Toggle all" },
            recalc: { ja: "再計算", en: "Recalculate" }
        },
        unit: {
            percent: { ja: "%", en: "%" },
            degree: { ja: "°", en: "°" },
            perInch: { ja: "/inch", en: "/inch" }
        },
        helpTip: {
            applyEach: {
                ja: "選択したグループを展開し、中の各オブジェクトに個別の乱数で適用します。\nOFF にするとグループ全体をひとつのオブジェクトとして扱います。",
                en: "Expands the selected groups and applies an individual random value to each child.\nWhen off, each group is treated as a single object."
            },
            splitFillStroke: {
                ja: "塗りと線を別々のオブジェクトに分け、それぞれに違う乱数を適用します。\n手描きの「線がはみ出た」印象になります。",
                en: "Separates fill and stroke into two objects and randomizes them independently,\nwhich mimics a hand-drawn outline that overshoots the fill."
            },
            keepFillStroke: { ja: "塗りと線を分けずに、そのまま適用します。", en: "Applies the effects without separating fill and stroke." },
            groupAtEnd: { ja: "分離した塗りと線をグループにまとめます。", en: "Groups the separated fill and stroke objects." },
            scale: {
                ja: "±この値の範囲で拡大・縮小します。\n↑↓キーで増減（Shift：10単位／Option：0.1単位）",
                en: "Scales randomly within plus/minus this value.\nUp/Down keys step the value (Shift: by 10, Option: by 0.1)."
            },
            move: {
                ja: "±この値の範囲で水平・垂直に移動します。\n↑↓キーで増減（Shift：10単位／Option：0.1単位）",
                en: "Moves randomly within plus/minus this value.\nUp/Down keys step the value (Shift: by 10, Option: by 0.1)."
            },
            rotate: {
                ja: "±この値の範囲で回転します。\n↑↓キーで増減（Shift：10単位／Option：0.1単位）",
                en: "Rotates randomly within plus/minus this value.\nUp/Down keys step the value (Shift: by 10, Option: by 0.1)."
            },
            toggleAll: { ja: "スケール・移動・回転をまとめて切り替えます。", en: "Turns Scale, Move and Rotate on or off together." },
            radius: {
                ja: "［角を丸くする］効果で角の尖りをやわらげます。\n↑↓キーで増減（Shift：10単位／Option：0.1単位）",
                en: "Softens sharp corners with the Round Corners effect.\nUp/Down keys step the value (Shift: by 10, Option: by 0.1)."
            },
            offset: {
                ja: "［パスのオフセット］を負・正の順に適用し、細かなノイズを削って形を整えます。\n↑↓キーで増減（Shift：10単位／Option：0.1単位）",
                en: "Applies Offset Path negatively then positively to smooth away small bumps.\nUp/Down keys step the value (Shift: by 10, Option: by 0.1)."
            },
            jagged: { ja: "輪郭を細かく震わせ、ペンで描いたような線にします。", en: "Adds fine jitter to the outline for a pen-drawn look." },
            distort: { ja: "輪郭を大きくうねらせ、形そのものを崩します。", en: "Warps the outline on a larger scale to distort the shape itself." },
            size: {
                ja: "オブジェクトのサイズに対する変形量（%）。\n↑↓キーで増減（Shift：10単位／Option：0.1単位）",
                en: "Amount of distortion relative to the object size (%).\nUp/Down keys step the value (Shift: by 10, Option: by 0.1)."
            },
            detail: {
                ja: "1インチあたりのアンカーポイント数。大きいほど細かくなります。\n↑↓キーで増減（Shift：10単位）",
                en: "Anchor points per inch. Higher values give a finer result.\nUp/Down keys step the value (Shift: by 10)."
            },
            recalc: {
                ja: "乱数を振り直してプレビューを作り直します。\nパレットから離れると、その時点の結果が確定します。",
                en: "Rerolls the random values and rebuilds the preview.\nLeaving the palette finalizes the current result."
            }
        },
        status: {
            ready: { ja: "オブジェクトを選択してください。", en: "Select an object to start." },
            applied: {
                ja: "適用しました。値を変えると作り直します。",
                en: "Applied. Changing a value rebuilds it."
            },
            busy: { ja: "処理中です。", en: "Working..." },
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noSelection: { ja: "オブジェクトを選択してください。", en: "Please select an object." },
            noTarget: { ja: "処理できるオブジェクトがありません。", en: "There is no object that can be processed." },
            invalidNumber: { ja: "数値入力が不正です。", en: "One or more numeric values are invalid." },
            timeout: { ja: "応答がありません。処理を中止しました。", en: "No response. The operation was cancelled." },
            error: { ja: "エラーが発生しました：", en: "An error occurred: " }
        }
    };

    /*
    ドットパスでラベルを取得（見つからなければ空文字）
    Get a label by dot path, returning an empty string when missing
    */
    function L(labelPath) {
        var pathParts = String(labelPath).split(".");
        var node = LABELS;

        for (var i = 0; i < pathParts.length; i++) {
            if (!node || !node[pathParts[i]]) {
                return "";
            }
            node = node[pathParts[i]];
        }
        return (node && node[currentLanguage]) ? node[currentLanguage] : "";
    }

    /* ラベルに区切り記号を付けて取得 / Get a label followed by a separator */
    function labelWithSeparator(labelPath) {
        return L(labelPath) + (currentLanguage === "ja" ? "：" : ":");
    }

    // =========================================
    // 単位 / Units
    // =========================================

    /* 単位コード → 表示名とポイント換算係数 / Unit code -> label and point factor */
    var UNIT_TABLE = {
        0: { label: "in", pointFactor: 72.0 },
        1: { label: "mm", pointFactor: 72.0 / 25.4 },
        2: { label: "pt", pointFactor: 1.0 },
        3: { label: "pica", pointFactor: 12.0 },
        4: { label: "cm", pointFactor: 72.0 / 2.54 },
        5: { label: "Q/H", pointFactor: 72.0 / 25.4 * 0.25 },
        6: { label: "px", pointFactor: 1.0 },
        7: { label: "ft/in", pointFactor: 72.0 * 12.0 },
        8: { label: "m", pointFactor: 72.0 / 25.4 * 1000.0 },
        9: { label: "yd", pointFactor: 72.0 * 36.0 },
        10: { label: "ft", pointFactor: 72.0 * 12.0 }
    };

    /* 定規の単位コードを取得 / Get the current ruler unit code */
    function getRulerUnitCode() {
        return app.preferences.getIntegerPreference("rulerType");
    }

    /* 定規の単位表示名を取得 / Get the current ruler unit label */
    function getRulerUnitLabel() {
        var unitEntry = UNIT_TABLE[getRulerUnitCode()];
        return unitEntry ? unitEntry.label : "pt";
    }

    /* 現在の単位の値をポイントに変換 / Convert a value in the current unit to points */
    function convertToPoint(value, unitCode) {
        var unitEntry = UNIT_TABLE[unitCode];
        return value * (unitEntry ? unitEntry.pointFactor : 1.0);
    }

    // =========================================
    // BridgeTalk 委譲 / BridgeTalk delegation
    // =========================================

    var isBusy = false; /* 再入防止 / re-entrancy guard */

    /*
    BridgeTalk の往復でパレットが一時的にフォーカスを失うことがある。
    その onDeactivate は isBusy を戻した後に配送される場合があり、
    ユーザー操作による離脱と区別できないため、直後の一定時間は無視する。
    A BridgeTalk round trip can make the palette lose focus briefly. That
    onDeactivate is sometimes delivered after isBusy has been cleared, and it is
    indistinguishable from the user leaving, so ignore it for a short while.
    */
    var engineCallEndTime = 0;

    /* 現在時刻をミリ秒で取得 / Current time in milliseconds */
    function nowMilliseconds() {
        return (new Date()).getTime();
    }

    /*
    worker 関数群をソース文字列に連結
    Concatenate the worker functions into one source string
    */
    function buildWorkerSource() {
        var sourceParts = [];

        for (var i = 0; i < WORKER_FUNCS.length; i++) {
            sourceParts.push(WORKER_FUNCS[i].toString());
        }
        return sourceParts.join("\n");
    }

    /*
    数値・真偽値だけのオプションをソース表現に変換
    Serialize a numbers-and-booleans option object into a source literal
    */
    function serializeOptions(options) {
        var parts = [];

        for (var key in options) {
            if (options.hasOwnProperty(key)) {
                var value = options[key];
                parts.push(key + ":" + ((typeof value === "boolean") ? (value ? "true" : "false") : String(value)));
            }
        }
        return "{" + parts.join(",") + "}";
    }

    /*
    worker をメインエンジンで同期実行し、戻り値のマーカー文字列を返す
    Run a worker call in the main engine synchronously and return its marker string
    */
    function callMainEngine(callExpression) {
        var payload = buildWorkerSource() + "\n" + callExpression + ";";
        var holder = { value: null, done: false };
        var bridge = new BridgeTalk();

        bridge.target = "illustrator";
        bridge.body = 'eval(decodeURIComponent("' + encodeURIComponent(payload) + '"));';

        bridge.onResult = function(resultMessage) {
            holder.value = String(resultMessage.body);
            holder.done = true;
        };
        bridge.onError = function(errorMessage) {
            holder.value = "ERR:" + String(errorMessage.body);
            holder.done = true;
        };
        bridge.onTimeout = function() {
            holder.value = "TIMEOUT";
            holder.done = true;
        };

        try {
            bridge.send(BRIDGE_TIMEOUT_SEC);
        } catch (e) {
            engineCallEndTime = nowMilliseconds();
            return "ERR:" + e;
        }

        engineCallEndTime = nowMilliseconds();
        return holder.done ? holder.value : "TIMEOUT";
    }

    // =========================================
    // 状態とUI参照 / State and UI references
    // =========================================

    /*
    直前の変更がこのスクリプト自身のもので、取り消してよいか
    Whether the last document change came from this script and may be undone
    */
    var canUndoLastApply = false;

    var applyEachCheckbox, splitFillStrokeRadio, keepFillStrokeRadio, groupAtEndCheckbox;
    var scaleCheckbox, moveCheckbox, rotateCheckbox;
    var scaleInput, moveInput, rotateInput;
    var radiusCheckbox, radiusInput, offsetCheckbox, offsetInput;
    var jaggedCheckbox, jaggedSizeRow, jaggedDetailRow;
    var distortCheckbox, distortSizeRow, distortDetailRow;
    var toggleAllButton, recalcButton, statusText;

    // =========================================
    // UIレイアウトの共通設定 / Shared UI layout
    // =========================================

    /* ウィンドウ・パネルの余白と間隔 / Window & panel margins and spacing */
    var WINDOW_MARGINS = 16; /* ウィンドウ外周の余白 / window margin */
    var WINDOW_SPACING = 12; /* ウィンドウ内の要素間隔 / window spacing */
    var PANEL_MARGINS = [16, 20, 16, 12]; /* パネル余白 [左,上,右,下] / panel margins */
    var PANEL_SPACING = 8; /* パネル内の要素間隔 / panel spacing */
    var COLUMN_SPACING = 12; /* 2カラムの間隔 / gap between columns */
    var ROW_LABEL_WIDTH = 60; /* 行ラベルの幅 / row label width */
    var ROW_CHECK_WIDTH = 90; /* 行チェックボックスの幅（既定）/ row checkbox width (default) */
    /*
    ラベルの長さに合わせた行チェックボックスの幅。
    英語ラベルは日本語より長いため、言語ごとに実測値を持つ
    Row checkbox widths tuned to the label length; the English labels are longer
    than the Japanese ones, so each language gets its own value
    */
    var TRANSFORM_CHECK_WIDTH = (currentLanguage === "ja") ? 58 : 78; /* 倍率・移動・回転（2文字）/ Scale, Move, Rotate */
    var CORNER_CHECK_WIDTH = (currentLanguage === "ja") ? 96 : 118; /* 角丸・オフセット / Round corners, Offset path */
    var FIELD_WIDTH_CHARS = 4; /* 入力欄の文字数 / input field width in characters */
    var TOGGLE_ROW_TOP_MARGIN = 10; /* ［すべてON/OFF］の上の余白 / gap above the Toggle all button */

    /* ウィンドウの共通設定 / Apply shared window layout */
    function setupWindow(win, spacing) {
        win.orientation = "column";
        win.alignChildren = "fill";
        win.margins = WINDOW_MARGINS;
        win.spacing = (typeof spacing === "number") ? spacing : WINDOW_SPACING;
    }

    /* パネルの共通設定 / Apply shared panel layout */
    function setupPanel(panel, spacing) {
        panel.orientation = "column";
        panel.alignChildren = ["fill", "top"];
        panel.alignment = "fill";
        panel.margins = PANEL_MARGINS;
        panel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /* 行グループの共通設定（ボタン列など） / Apply a horizontal row group */
    function setupRow(group, alignment, spacing) {
        group.orientation = "row";
        group.alignment = alignment || "left";
        group.alignChildren = ["left", "center"];
        group.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /* ボタンの高さを指定 px 詰める（レイアウト確定後に呼ぶ）/ Trim a button's height by the given px (call after layout) */
    function trimButtonHeight(button, px) {
        try {
            button.size = [button.size.width, button.size.height - px];
        } catch (e) {}
    }

    /* ヘルプチップを設定 / Attach a help tip */
    function setHelpTip(control, helpTipPath) {
        control.helpTip = L(helpTipPath);
        return control;
    }

    /*
    「ラベル付きチェックボックス＋入力欄＋単位」の1行を作る
    Build one row: labeled checkbox + input field + unit
    */
    function addCheckboxFieldRow(parentGroup, labelPath, helpTipPath, defaultText, unitText, isChecked, checkWidth) {
        var rowGroup = parentGroup.add("group");
        setupRow(rowGroup, "left");

        var checkbox = rowGroup.add("checkbox", undefined, labelWithSeparator(labelPath));
        checkbox.preferredSize.width = (typeof checkWidth === "number") ? checkWidth : ROW_CHECK_WIDTH;
        checkbox.value = isChecked;

        var input = rowGroup.add("edittext", undefined, defaultText);
        input.characters = FIELD_WIDTH_CHARS;
        input.enabled = isChecked;

        rowGroup.add("statictext", undefined, unitText);

        setHelpTip(checkbox, helpTipPath);
        setHelpTip(input, helpTipPath);

        return { group: rowGroup, checkbox: checkbox, input: input };
    }

    /*
    「右寄せラベル＋入力欄＋単位」の1行を作る
    Build one row: right-aligned label + input field + unit
    */
    function addLabelFieldRow(parentGroup, labelPath, helpTipPath, defaultText, unitText) {
        var rowGroup = parentGroup.add("group");
        setupRow(rowGroup, "left");

        var labelText = rowGroup.add("statictext", undefined, labelWithSeparator(labelPath), { justify: "right" });
        labelText.preferredSize.width = ROW_LABEL_WIDTH;

        var input = rowGroup.add("edittext", undefined, defaultText);
        input.characters = FIELD_WIDTH_CHARS;

        rowGroup.add("statictext", undefined, unitText);

        setHelpTip(labelText, helpTipPath);
        setHelpTip(input, helpTipPath);

        return { group: rowGroup, label: labelText, input: input };
    }

    // =========================================
    // 入力値の読み取り / Reading the input values
    // =========================================

    /*
    入力欄の数値を取得（OFF なら 0、負数は 0 にクランプ）
    Read a numeric field; 0 when disabled, negatives clamped to 0
    */
    function readFieldNumber(input, isEnabled) {
        if (!isEnabled) {
            return 0;
        }

        var value = parseFloat(input.text);
        if (isNaN(value)) {
            return Number.NaN;
        }
        return (value < 0) ? 0 : value;
    }

    /*
    UIの入力値をまとめて読み取る。不正な数値があれば null を返す
    Read every UI value into one options object; returns null when a value is invalid
    */
    function readOptions() {
        var usesScale = scaleCheckbox.value;
        var usesMove = moveCheckbox.value;
        var usesRotate = rotateCheckbox.value;

        var scaleRange = readFieldNumber(scaleInput, usesScale);
        var moveValue = readFieldNumber(moveInput, usesMove);
        var rotateRange = readFieldNumber(rotateInput, usesRotate);
        var radiusValue = readFieldNumber(radiusInput, radiusCheckbox.value);
        var offsetValue = readFieldNumber(offsetInput, offsetCheckbox.value);
        var jaggedSize = readFieldNumber(jaggedSizeRow.input, jaggedCheckbox.value);
        var jaggedDetail = readFieldNumber(jaggedDetailRow.input, jaggedCheckbox.value);
        var distortSize = readFieldNumber(distortSizeRow.input, distortCheckbox.value);
        var distortDetail = readFieldNumber(distortDetailRow.input, distortCheckbox.value);

        var numericValues = [scaleRange, moveValue, rotateRange, radiusValue, offsetValue,
            jaggedSize, jaggedDetail, distortSize, distortDetail
        ];
        for (var i = 0; i < numericValues.length; i++) {
            if (isNaN(numericValues[i])) {
                return null;
            }
        }

        var unitCode = getRulerUnitCode();

        return {
            usesTransform: (usesScale || usesMove || usesRotate),
            scaleRange: scaleRange,
            rotateRange: rotateRange,
            movePoint: convertToPoint(moveValue, unitCode),
            radiusPoint: convertToPoint(radiusValue, unitCode),
            offsetPoint: convertToPoint(offsetValue, unitCode),
            usesJagged: jaggedCheckbox.value,
            jaggedSize: jaggedSize,
            jaggedDetail: Math.max(ROUGHEN_DETAIL_MIN, Math.round(jaggedDetail)),
            usesDistort: distortCheckbox.value,
            distortSize: distortSize,
            distortDetail: Math.max(ROUGHEN_DETAIL_MIN, Math.round(distortDetail)),
            expandsGroups: applyEachCheckbox.value,
            splitsFillStroke: splitFillStrokeRadio.value,
            groupsAtEnd: groupAtEndCheckbox.value,
            joinType: OFFSET_JOIN_TYPE,
            miterLimit: OFFSET_MITER_LIMIT,
            roundness: ROUGHEN_ROUNDNESS
        };
    }

    // =========================================
    // 状況表示と状態更新 / Status and state updates
    // =========================================

    /* 状況表示を更新 / Update the status line */
    function setStatus(statusPath, extraText) {
        if (!statusText) {
            return;
        }
        statusText.text = L(statusPath) + (extraText ? String(extraText) : "");
    }

    /*
    worker の戻り値マーカーを状況表示に反映
    Reflect a worker result marker in the status line
    */
    function showResultStatus(resultMarker) {
        if (resultMarker === "NODOC") {
            setStatus("status.noDocument");
        } else if (resultMarker === "NOSEL") {
            setStatus("status.noSelection");
        } else if (resultMarker === "NOTARGET") {
            setStatus("status.noTarget");
        } else if (resultMarker === "TIMEOUT") {
            setStatus("status.timeout");
        } else if (resultMarker.indexOf("ERR:") === 0) {
            setStatus("status.error", resultMarker.substring(4));
        } else {
            setStatus("status.applied");
        }
    }

    /*
    "OK|1|0" 形式の戻り値から対象パネルの有効・無効を更新
    Update the Target panel from an "OK|1|0" style result
    */
    function updateTargetPanelState(resultMarker) {
        var canSplit = false;
        var canExpand = false;

        if (resultMarker && resultMarker.indexOf("OK|") === 0) {
            var parts = resultMarker.split("|");
            canSplit = (parts[1] === "1");
            canExpand = (parts[2] === "1");
        }

        if (!canSplit && splitFillStrokeRadio.value) {
            splitFillStrokeRadio.value = false;
            keepFillStrokeRadio.value = true;
        }

        splitFillStrokeRadio.enabled = canSplit;
        applyEachCheckbox.enabled = canExpand;
        groupAtEndCheckbox.enabled = canSplit && splitFillStrokeRadio.value;

        if (!groupAtEndCheckbox.enabled) {
            groupAtEndCheckbox.value = false;
        }
    }

    /* 変形パネルの入力欄と［再計算］の有効・無効を更新 / Update the Transform inputs and Recalculate button */
    function updateTransformPanelState() {
        scaleInput.enabled = scaleCheckbox.value;
        moveInput.enabled = moveCheckbox.value;
        rotateInput.enabled = rotateCheckbox.value;
        updateRecalcButtonState();
    }

    /* ［再計算］の有効・無効を更新 / Update the Recalculate button state */
    function updateRecalcButtonState() {
        if (!recalcButton) {
            return;
        }
        recalcButton.enabled = !!(scaleCheckbox.value || moveCheckbox.value || rotateCheckbox.value ||
            jaggedCheckbox.value || distortCheckbox.value);
    }

    /* ラフ効果パネルの入力欄の有効・無効を更新 / Update the Roughen panel inputs */
    function updateRoughenPanelState(applyCheckbox, sizeRow, detailRow) {
        var isEnabled = applyCheckbox.value;
        sizeRow.label.enabled = isEnabled;
        sizeRow.input.enabled = isEnabled;
        detailRow.label.enabled = isEnabled;
        detailRow.input.enabled = isEnabled;
        updateRecalcButtonState();
    }

    // =========================================
    // プレビュー / Preview
    // =========================================

    /*
    プレビューを作り直す（前回分は worker 側で取り消す）
    Rebuild the preview; the previous one is undone inside the worker
    */
    function refreshPreview() {
        if (isBusy) {
            setStatus("status.busy");
            return;
        }

        var options = readOptions();
        if (!options) {
            setStatus("status.invalidNumber");
            return;
        }

        isBusy = true;
        try {
            var callExpression = "wkApplyEffects(" + serializeOptions(options) + ", " +
                (canUndoLastApply ? "true" : "false") + ")";
            var resultMarker = callMainEngine(callExpression);

            canUndoLastApply = (resultMarker.indexOf("OK") === 0);
            updateTargetPanelState(resultMarker);
            showResultStatus(resultMarker);
        } finally {
            isBusy = false;
        }
    }

    /*
    直前の結果を確定し、次回は取り消さないようにする
    Finalize the last result so that the next run does not undo it
    */
    function commitLastApply() {
        canUndoLastApply = false;
    }

    // =========================================
    // 入力欄の補助 / Input field helpers
    // =========================================

    /*
    矢印キーで数値を増減（Shift:10単位 / Option:0.1単位）
    Arrow keys change the value (Shift: by 10, Option/Alt: by 0.1)
    */
    function changeValueByArrowKey(editText, isIntegerOnly) {
        editText.addEventListener("keydown", function(event) {
            if (event.keyName !== "Up" && event.keyName !== "Down") {
                return;
            }

            var value = Number(editText.text);
            if (isNaN(value)) {
                return;
            }

            var keyboardState = ScriptUI.environment.keyboardState;
            var isUp = (event.keyName === "Up");

            if (keyboardState.shiftKey) {
                value = isUp ? Math.ceil((value + 1) / 10) * 10 : Math.floor((value - 1) / 10) * 10;
            } else if (keyboardState.altKey && !isIntegerOnly) {
                value = Math.round((value + (isUp ? 0.1 : -0.1)) * 10) / 10;
            } else {
                value = Math.round(value) + (isUp ? 1 : -1);
            }

            if (value < 0) {
                value = 0;
            }
            if (isIntegerOnly) {
                value = Math.round(value);
            }

            event.preventDefault();
            editText.text = String(value);
            refreshPreview();
        });
    }

    /* 詳細欄を整数に丸める / Round a detail field to an integer */
    function normalizeDetailInput(editText) {
        var value = Number(editText.text);
        if (isNaN(value)) {
            return;
        }
        editText.text = String(Math.max(ROUGHEN_DETAIL_MIN, Math.round(value)));
    }

    // =========================================
    // パレットの構築 / Building the palette
    // =========================================

    /* 対象パネル（分離とグループの扱い） / Target panel: split and group handling */
    function buildTargetPanel(parentGroup) {
        var targetPanel = parentGroup.add("panel", undefined, L("panel.target"));
        setupPanel(targetPanel, 6);
        targetPanel.alignChildren = ["left", "top"];

        applyEachCheckbox = setHelpTip(
            targetPanel.add("checkbox", undefined, L("checkbox.applyEach")), "helpTip.applyEach");
        applyEachCheckbox.value = INITIAL_CHECKED.applyEach;

        splitFillStrokeRadio = setHelpTip(
            targetPanel.add("radiobutton", undefined, L("radio.splitFillStroke")), "helpTip.splitFillStroke");
        keepFillStrokeRadio = setHelpTip(
            targetPanel.add("radiobutton", undefined, L("radio.keepFillStroke")), "helpTip.keepFillStroke");
        splitFillStrokeRadio.value = false;
        keepFillStrokeRadio.value = true;

        groupAtEndCheckbox = setHelpTip(
            targetPanel.add("checkbox", undefined, L("checkbox.groupAtEnd")), "helpTip.groupAtEnd");
        groupAtEndCheckbox.value = false;

        applyEachCheckbox.onClick = refreshPreview;
        splitFillStrokeRadio.onClick = refreshPreview;
        keepFillStrokeRadio.onClick = refreshPreview;
        groupAtEndCheckbox.onClick = refreshPreview;

        return targetPanel;
    }

    /* 変形パネル（スケール・移動・回転） / Transform panel: scale, move, rotate */
    function buildTransformPanel(parentGroup) {
        var transformPanel = parentGroup.add("panel", undefined, L("panel.transform"));
        setupPanel(transformPanel, 6);

        var rulerUnitLabel = getRulerUnitLabel();

        var scaleRow = addCheckboxFieldRow(transformPanel, "fieldLabel.scale", "helpTip.scale",
            DEFAULT_VALUES.scale, L("unit.percent"), INITIAL_CHECKED.scale, TRANSFORM_CHECK_WIDTH);
        var moveRow = addCheckboxFieldRow(transformPanel, "fieldLabel.move", "helpTip.move",
            DEFAULT_VALUES.move, rulerUnitLabel, INITIAL_CHECKED.move, TRANSFORM_CHECK_WIDTH);
        var rotateRow = addCheckboxFieldRow(transformPanel, "fieldLabel.rotate", "helpTip.rotate",
            DEFAULT_VALUES.rotate, L("unit.degree"), INITIAL_CHECKED.rotate, TRANSFORM_CHECK_WIDTH);

        scaleCheckbox = scaleRow.checkbox;
        scaleInput = scaleRow.input;
        moveCheckbox = moveRow.checkbox;
        moveInput = moveRow.input;
        rotateCheckbox = rotateRow.checkbox;
        rotateInput = rotateRow.input;

        /* ボタンは幅いっぱいに広げず、上に余白を入れる / Keep the button narrow and add a gap above it */
        var toggleButtonRow = transformPanel.add("group");
        setupRow(toggleButtonRow, "left");
        toggleButtonRow.margins = [0, TOGGLE_ROW_TOP_MARGIN, 0, 0];
        toggleAllButton = setHelpTip(
            toggleButtonRow.add("button", undefined, L("button.toggleAll")), "helpTip.toggleAll");
        toggleAllButton.alignment = "left";

        toggleAllButton.onClick = function() {
            var nextValue = !(scaleCheckbox.value && moveCheckbox.value && rotateCheckbox.value);
            scaleCheckbox.value = nextValue;
            moveCheckbox.value = nextValue;
            rotateCheckbox.value = nextValue;
            updateTransformPanelState();
            refreshPreview();
        };

        scaleCheckbox.onClick = function() { updateTransformPanelState(); refreshPreview(); };
        moveCheckbox.onClick = function() { updateTransformPanelState(); refreshPreview(); };
        rotateCheckbox.onClick = function() { updateTransformPanelState(); refreshPreview(); };

        scaleInput.onChange = refreshPreview;
        moveInput.onChange = refreshPreview;
        rotateInput.onChange = refreshPreview;

        return transformPanel;
    }

    /* 角丸・オフセットパネル / Corners & Offset panel */
    function buildCornerPanel(parentGroup) {
        var cornerPanel = parentGroup.add("panel", undefined, L("panel.corner"));
        setupPanel(cornerPanel, 6);

        var rulerUnitLabel = getRulerUnitLabel();

        var radiusRow = addCheckboxFieldRow(cornerPanel, "checkbox.radius", "helpTip.radius",
            DEFAULT_VALUES.radius, rulerUnitLabel, INITIAL_CHECKED.radius, CORNER_CHECK_WIDTH);
        var offsetRow = addCheckboxFieldRow(cornerPanel, "checkbox.offset", "helpTip.offset",
            DEFAULT_VALUES.offset, rulerUnitLabel, INITIAL_CHECKED.offset, CORNER_CHECK_WIDTH);

        radiusCheckbox = radiusRow.checkbox;
        radiusInput = radiusRow.input;
        offsetCheckbox = offsetRow.checkbox;
        offsetInput = offsetRow.input;

        radiusCheckbox.onClick = function() {
            radiusInput.enabled = radiusCheckbox.value;
            refreshPreview();
        };
        offsetCheckbox.onClick = function() {
            offsetInput.enabled = offsetCheckbox.value;
            refreshPreview();
        };
        radiusInput.onChange = refreshPreview;
        offsetInput.onChange = refreshPreview;

        return cornerPanel;
    }

    /*
    ラフ効果パネル（ギザギザ／歪曲で共通）
    Roughen panel, shared by the jagged and distortion variants
    */
    function buildRoughenPanel(parentGroup, panelTitlePath, applyLabelPath, applyHelpTipPath,
        isChecked, sizeDefaultText, detailDefaultText) {
        var roughenPanel = parentGroup.add("panel", undefined, L(panelTitlePath));
        setupPanel(roughenPanel, 6);

        var applyCheckbox = setHelpTip(
            roughenPanel.add("checkbox", undefined, L(applyLabelPath)), applyHelpTipPath);
        applyCheckbox.value = isChecked;

        var sizeRow = addLabelFieldRow(roughenPanel, "fieldLabel.size", "helpTip.size",
            sizeDefaultText, L("unit.percent"));
        var detailRow = addLabelFieldRow(roughenPanel, "fieldLabel.detail", "helpTip.detail",
            detailDefaultText, L("unit.perInch"));

        applyCheckbox.onClick = function() {
            updateRoughenPanelState(applyCheckbox, sizeRow, detailRow);
            refreshPreview();
        };
        sizeRow.input.onChange = refreshPreview;
        detailRow.input.onChange = function() {
            normalizeDetailInput(detailRow.input);
            refreshPreview();
        };

        return { panel: roughenPanel, checkbox: applyCheckbox, sizeRow: sizeRow, detailRow: detailRow };
    }

    /* ボタンと状況表示 / Buttons and status line */
    function buildFooter(parentWindow) {
        var buttonArea = parentWindow.add("group");
        setupRow(buttonArea, ["fill", "top"]);

        recalcButton = setHelpTip(
            buttonArea.add("button", undefined, L("button.recalc")), "helpTip.recalc");
        recalcButton.alignment = "left";
        recalcButton.onClick = refreshPreview;

        /*
        幅は上のカラムに合わせる。固定幅にするとこの1行がウィンドウ幅の下限を決めてしまう
        Let the columns above decide the width; a fixed width would make this single
        line dictate the minimum window width
        */
        statusText = parentWindow.add("statictext", undefined, "", { truncate: "end" });
        statusText.alignment = ["fill", "bottom"];

        return buttonArea;
    }

    /*
    パレットを組み立てる
    Assemble the palette window
    */
    function createPalette() {
        var win = new Window("palette", L("dialog.title") + " " + SCRIPT_VERSION, undefined, { resizeable: false });
        setupWindow(win);

        /* 上部エリア（左右2カラム） / Top area with two columns */
        var columnsGroup = win.add("group");
        columnsGroup.orientation = "row";
        columnsGroup.alignChildren = ["fill", "top"];
        columnsGroup.spacing = COLUMN_SPACING;

        var leftColumnGroup = columnsGroup.add("group");
        leftColumnGroup.orientation = "column";
        leftColumnGroup.alignChildren = ["fill", "top"];
        leftColumnGroup.spacing = WINDOW_SPACING;

        var rightColumnGroup = columnsGroup.add("group");
        rightColumnGroup.orientation = "column";
        rightColumnGroup.alignChildren = ["fill", "top"];
        rightColumnGroup.spacing = WINDOW_SPACING;

        buildTargetPanel(leftColumnGroup);
        buildTransformPanel(leftColumnGroup);

        buildCornerPanel(rightColumnGroup);

        var jaggedPanelUI = buildRoughenPanel(rightColumnGroup,
            "panel.jagged", "checkbox.jagged", "helpTip.jagged",
            INITIAL_CHECKED.jagged, DEFAULT_VALUES.jaggedSize, DEFAULT_VALUES.jaggedDetail);
        jaggedCheckbox = jaggedPanelUI.checkbox;
        jaggedSizeRow = jaggedPanelUI.sizeRow;
        jaggedDetailRow = jaggedPanelUI.detailRow;

        var distortPanelUI = buildRoughenPanel(rightColumnGroup,
            "panel.distort", "checkbox.distort", "helpTip.distort",
            INITIAL_CHECKED.distort, DEFAULT_VALUES.distortSize, DEFAULT_VALUES.distortDetail);
        distortCheckbox = distortPanelUI.checkbox;
        distortSizeRow = distortPanelUI.sizeRow;
        distortDetailRow = distortPanelUI.detailRow;

        buildFooter(win);

        /* 矢印キーでの増減を有効化 / Enable arrow-key stepping */
        changeValueByArrowKey(scaleInput, false);
        changeValueByArrowKey(moveInput, false);
        changeValueByArrowKey(rotateInput, false);
        changeValueByArrowKey(radiusInput, false);
        changeValueByArrowKey(offsetInput, false);
        changeValueByArrowKey(jaggedSizeRow.input, false);
        changeValueByArrowKey(jaggedDetailRow.input, true);
        changeValueByArrowKey(distortSizeRow.input, false);
        changeValueByArrowKey(distortDetailRow.input, true);

        /* Esc で閉じる / Close on Escape */
        win.addEventListener("keydown", function(event) {
            if (event.keyName === "Escape") {
                win.close();
            }
        });

        /*
        パレットから離れたら、その時点の結果を確定する。
        こうしておかないと、ユーザーがドキュメント側で行った操作を
        次の再計算の app.undo() が取り消してしまう。
        Finalize the current result when the palette loses focus; otherwise the
        next app.undo() would revert whatever the user did in the document.
        */
        if (COMMITS_ON_DEACTIVATE) {
            win.onDeactivate = function() {
                if (isBusy) {
                    return;
                }
                /* BridgeTalk 往復に伴う一時的な離脱は無視 / Ignore the transient focus loss caused by BridgeTalk */
                if ((nowMilliseconds() - engineCallEndTime) < DEACTIVATE_GUARD_MS) {
                    return;
                }
                commitLastApply();
            };
        }

        /* 閉じても結果はそのまま残す / Closing keeps the current result */
        win.onClose = function() {
            commitLastApply();
            paletteWindow = null;
            return true;
        };

        /* 初期状態を反映 / Apply the initial state */
        updateTransformPanelState();
        updateRoughenPanelState(jaggedCheckbox, jaggedSizeRow, jaggedDetailRow);
        updateRoughenPanelState(distortCheckbox, distortSizeRow, distortDetailRow);
        updateTargetPanelState(null);
        setStatus("status.ready");

        return win;
    }

    /*
    パレットを表示（既存のパレットがあれば閉じてから）
    Show the palette, closing an existing one first
    */
    function showPalette() {
        if (paletteWindow) {
            try {
                paletteWindow.close();
            } catch (e) {}
            paletteWindow = null;
        }

        paletteWindow = createPalette();
        paletteWindow.show();

        /* レイアウト確定後にボタンの高さを詰める / Trim the button height once the layout is settled */
        trimButtonHeight(toggleAllButton, 4);

        /* 選択状態を取得して対象パネルへ反映 / Fetch the selection state for the Target panel */
        var stateMarker = callMainEngine("wkGetSelectionState(" + (applyEachCheckbox.value ? "true" : "false") + ")");
        updateTargetPanelState(stateMarker);

        if (stateMarker.indexOf("OK") === 0) {
            refreshPreview();
        } else {
            showResultStatus(stateMarker);
        }
    }

    showPalette();

})();
