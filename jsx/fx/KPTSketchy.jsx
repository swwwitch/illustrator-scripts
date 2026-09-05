#target illustrator
#targetengine "KPTSketchy"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したオブジェクトにランダムな変形を加え、手書き・スケッチ風の見た目に整える常駐パレットです。
変形はすべてライブ効果として適用するため、あとから編集・解除できます。

詳細は README を参照してください。

### Overview

A persistent palette that adds random distortion to the selected objects for a hand-drawn, sketchy look.
Everything is applied as live effects, so the result stays editable and removable.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "KPTSketchy";                   /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.2.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-04-14";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-05";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/KPTSketchy.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/KPTSketchy.md
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/na808bac430d9"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    /*
    常駐エンジンに保持するパレット参照（GC回避・多重起動防止）。
    関数スコープの var は巻き上げで必ず undefined から始まり、再実行のたびに作り直されるため、
    前回のパレットを覚えておけない。$.global に載せてエンジン側に残す
    Palette reference kept alive in the persistent engine (prevents GC and double launch).
    A function-scoped var is hoisted as undefined and rebuilt on every run, so it cannot
    remember the previous palette; keep it on $.global instead
    */
    var PALETTE_GLOBAL_KEY = "__KPTSketchyPaletteWindow";

    /* 常駐エンジンが保持しているパレットを取得 / Get the palette kept by the persistent engine */
    function getKeptPalette() {
        return $.global[PALETTE_GLOBAL_KEY] || null;
    }

    /* 常駐エンジンが保持するパレットを差し替え / Replace the palette kept by the persistent engine */
    function keepPalette(targetWindow) {
        $.global[PALETTE_GLOBAL_KEY] = targetWindow;
    }

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

    /* 表示される塗りと線の両方を持つか（分離可能か） / Whether the object has both a visible fill and stroke */
    function wkCanSplitFillStroke(pageItem) {
        try {
            return !!(pageItem.filled && pageItem.fillColor && pageItem.fillColor.typename !== "NoColor" &&
                pageItem.stroked && pageItem.strokeColor && pageItem.strokeColor.typename !== "NoColor");
        } catch (e) {
            return false;
        }
    }

    /* クリッピングパスか / Whether the item is a clipping path */
    function wkIsClippingPathItem(pageItem) {
        return !!(pageItem && pageItem.typename === "PathItem" && pageItem.clipping);
    }

    /* グループを展開して処理対象を集める / Collect processable items, expanding groups */
    function wkCollectTargetItems(pageItem, collectedItems) {
        if (!pageItem) {
            return;
        }
        if (pageItem.typename === "GroupItem") {
            for (var i = 0; i < pageItem.pageItems.length; i++) {
                var childItem = pageItem.pageItems[i];
                if (childItem.parent === pageItem) {
                    wkCollectTargetItems(childItem, collectedItems);
                }
            }
            return;
        }
        if (!wkIsClippingPathItem(pageItem)) {
            collectedItems.push(pageItem);
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
    function wkApplyEffectStack(targetItem, effectOptions) {
        if (effectOptions.radiusPoint > 0) {
            wkApplyEffectXml(targetItem, wkBuildRoundCornersEffectXml(effectOptions.radiusPoint));
        }
        if (effectOptions.offsetPoint !== 0) {
            wkApplyEffectXml(targetItem, wkBuildOffsetPathEffectXml(-effectOptions.offsetPoint, effectOptions.joinType, effectOptions.miterLimit));
            wkApplyEffectXml(targetItem, wkBuildOffsetPathEffectXml(effectOptions.offsetPoint, effectOptions.joinType, effectOptions.miterLimit));
        }
        if (effectOptions.usesTransform) {
            wkApplyEffectXml(targetItem, wkBuildTransformEffectXml(
                wkRandomBetween(100 - effectOptions.scaleRange, 100 + effectOptions.scaleRange),
                wkRandomBetween(-effectOptions.movePoint, effectOptions.movePoint),
                wkRandomBetween(-effectOptions.movePoint, effectOptions.movePoint),
                wkRandomBetween(-effectOptions.rotateRange, effectOptions.rotateRange)));
        }
        if (effectOptions.usesDistort) {
            wkApplyEffectXml(targetItem, wkBuildRoughenEffectXml(effectOptions.distortSize, effectOptions.distortDetail, effectOptions.roundness));
        }
        if (effectOptions.usesJagged) {
            wkApplyEffectXml(targetItem, wkBuildRoughenEffectXml(effectOptions.jaggedSize, effectOptions.jaggedDetail, effectOptions.roundness));
        }
    }

    /* 対象すべてに効果を適用（必要なら塗り／線を分離） / Apply the effects to every target, splitting fill and stroke when asked */
    function wkApplyToTargetItems(targetItems, effectOptions) {
        for (var i = 0; i < targetItems.length; i++) {
            var targetItem = targetItems[i];
            var strokeCopy = null;
            if (effectOptions.splitsFillStroke && wkCanSplitFillStroke(targetItem)) {
                strokeCopy = wkSplitIntoStrokeCopy(targetItem);
            }
            wkApplyEffectStack(targetItem, effectOptions);
            if (strokeCopy) {
                wkApplyEffectStack(strokeCopy, effectOptions);
                if (effectOptions.groupsAtEnd && !wkIsInsideClippingGroup(targetItem)) {
                    wkGroupFillAndStroke(targetItem, strokeCopy);
                }
            }
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
    function wkApplyEffects(effectOptions, undoFirst) {
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
            var targetItems = wkGetTargetItems(targetDoc, effectOptions.expandsGroups);
            if (targetItems.length === 0) {
                return "NOTARGET";
            }
            /*
            分離すると元オブジェクトの線が外れ、分離可能と判定されなくなるため、
            UIへ返す選択状態は適用前に採取しておく
            Splitting clears the original's stroke, so the state reported to the UI
            must be captured before the effects are applied
            */
            var stateSuffix = wkSelectionStateSuffix(targetDoc, effectOptions.expandsGroups);
            /*
            複製・移動・グループ化は選択を動かす。次回の再計算も同じ選択を対象にするため、
            処理前の選択を控えて最後に戻す
            Duplicating, moving and grouping all shift the selection. Snapshot it so the
            next rebuild works on the same objects
            */
            var savedSelection = wkSnapshotSelection(targetDoc);
            wkApplyToTargetItems(targetItems, effectOptions);
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
    var WORKER_FUNCTIONS = [
        wkRandomBetween,
        wkApplyEffectXml,
        wkBuildRoughenEffectXml,
        wkBuildRoundCornersEffectXml,
        wkBuildOffsetPathEffectXml,
        wkBuildTransformEffectXml,
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
        wkApplyToTargetItems,
        wkSelectionStateSuffix,
        wkGetSelectionState,
        wkApplyEffects
    ];

    (function() {

        // =========================================
        // ユーザー設定 / User settings
        // =========================================

        /* 入力欄の初期値 / Default values for the input fields */
        var DEFAULT_FIELD_VALUES = {
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
        var INITIAL_CHECKBOX_STATES = {
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
        var DEACTIVATE_GUARD_MS = 500; /* BridgeTalk 直後の deactivate を無視する時間（ミリ秒）/ ignore deactivate for this long after a BridgeTalk call (ms) */

        // =========================================
        // ローカライズ / Localization
        // =========================================

        /* 表示言語を判定 / Detect the UI language */
        function getUiLang() {
            return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
        }
        var uiLang = getUiLang();

        /* 日英ラベル定義 / Japanese-English label definitions */
        var LABELS = {
            dialog: {
                title: { ja: "手書き・スケッチ風", en: "Hand-drawn / Sketchy Look" }
            },
            panel: {
                target: { ja: "適用対象", en: "Apply to" },
                transform: { ja: "ランダム変形", en: "Random transform" },
                corner: { ja: "角丸・オフセット", en: "Corners & Offset" },
                /*
                Illustratorの［ラフ］で「ギザギザ」はポイントの形状（丸く⇔ギザギザ）を指す。
                本スクリプトは丸く固定で、2つのパネルの違いはサイズと詳細の大小なので、その差で呼び分ける
                In Illustrator's Roughen, "Jagged" names the point option (Smooth vs Corner).
                This script pins it to Smooth, so the two panels differ only in size and detail
                */
                jagged: { ja: "ラフ：細かい揺れ", en: "Roughen: Fine" },
                distort: { ja: "ラフ：大きなうねり", en: "Roughen: Coarse" }
            },
            checkbox: {
                applyEach: { ja: "グループ内を個別に処理", en: "Process group items individually" },
                groupAtEnd: { ja: "分離後にグループ化", en: "Group after split" },
                radius: { ja: "角丸", en: "Round corners" },
                offset: { ja: "パスのオフセット", en: "Offset path" },
                jagged: { ja: "有効", en: "Enable" },
                distort: { ja: "有効", en: "Enable" }
            },
            radio: {
                splitFillStroke: { ja: "塗りと線を分離", en: "Split fill and stroke" },
                keepFillStroke: { ja: "分離しない", en: "Keep as is" }
            },
            fieldLabel: {
                scale: { ja: "倍率", en: "Scale" },
                move: { ja: "移動", en: "Move" },
                rotate: { ja: "回転", en: "Rotate" },
                size: { ja: "サイズ", en: "Size" },
                detail: { ja: "詳細", en: "Detail" }
            },
            button: {
                toggleAll: { ja: "まとめて切替", en: "Toggle all" },
                recalc: { ja: "振り直す", en: "Reroll" }
            },
            unit: {
                percent: { ja: "%", en: "%" },
                degree: { ja: "°", en: "°" },
                perInch: { ja: "/インチ", en: "/inch" }
            },
            helpTip: {
                applyEach: {
                    ja: "選択したグループを展開し、中の各オブジェクトに個別の乱数で適用します。\nOFF にするとグループ全体をひとつのオブジェクトとして扱います。\nグループまたは複数のオブジェクトを選択すると有効になります。",
                    en: "Expands the selected groups and applies an individual random value to each child.\nWhen off, each group is treated as a single object.\nAvailable when a group or several objects are selected."
                },
                splitFillStroke: {
                    ja: "塗りと線を別々のオブジェクトに分け、それぞれに違う乱数を適用します。\n手描きの「線がはみ出た」印象になります。\n塗りと線の両方を持つオブジェクトを選択すると有効になります。",
                    en: "Separates fill and stroke into two objects and randomizes them independently,\nwhich mimics a hand-drawn outline that overshoots the fill.\nAvailable when the selection has an object with both a fill and a stroke."
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
                toggleAll: { ja: "倍率・移動・回転をまとめて切り替えます。", en: "Turns Scale, Move and Rotate on or off together." },
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
                ready: { ja: "オブジェクトを選択すると適用されます。", en: "Select an object and the effects are applied." },
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
                undoManually: {
                    ja: "効果が適用済みの場合は手動で取り消してください。",
                    en: "Undo manually if the effects were already applied."
                },
                error: { ja: "エラーが発生しました：", en: "An error occurred: " }
            }
        };

        /*
        ドットパスでラベルを取得（見つからなければ空文字）
        Get a label by dot path, returning an empty string when missing
        */
        function getLabel(labelPath) {
            var pathParts = String(labelPath).split(".");
            var labelNode = LABELS;

            for (var i = 0; i < pathParts.length; i++) {
                if (!labelNode || !labelNode[pathParts[i]]) {
                    return "";
                }
                labelNode = labelNode[pathParts[i]];
            }
            return (labelNode && labelNode[uiLang]) ? labelNode[uiLang] : "";
        }

        /* コロン付きラベル（日本語は全角、英語は半角）/ Label with colon (full-width JA, half-width EN) */
        function labelText(labelPath) {
            return getLabel(labelPath) + (uiLang === "ja" ? "：" : ":");
        }

        // =========================================
        // 単位 / Units
        // =========================================

        /* 単位コード → 表示名とポイント換算係数 / Unit code -> label and point factor */
        var RULER_UNIT_TABLE = {
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

        /* 振れ幅であることを示す接頭辞 / Prefix marking a value as a plus-minus range */
        var RANGE_PREFIX = "\u00b1";

        /* 定規の単位コードを取得 / Get the current ruler unit code */
        function getRulerUnitCode() {
            return app.preferences.getIntegerPreference("rulerType");
        }

        /* 定規の単位表示名を取得 / Get the current ruler unit label */
        function getRulerUnitLabel() {
            var unitEntry = RULER_UNIT_TABLE[getRulerUnitCode()];
            return unitEntry ? unitEntry.label : "pt";
        }

        /* 現在の単位の値をポイントに変換 / Convert a value in the current unit to points */
        function convertToPoint(value, unitCode) {
            var unitEntry = RULER_UNIT_TABLE[unitCode];
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

            for (var i = 0; i < WORKER_FUNCTIONS.length; i++) {
                sourceParts.push(WORKER_FUNCTIONS[i].toString());
            }
            return sourceParts.join("\n");
        }

        /*
        数値・真偽値だけのオプションをソース表現に変換
        Serialize a numbers-and-booleans option object into a source literal
        */
        function serializeEffectOptions(effectOptions) {
            var optionParts = [];

            for (var optionKey in effectOptions) {
                if (effectOptions.hasOwnProperty(optionKey)) {
                    var optionValue = effectOptions[optionKey];
                    optionParts.push(optionKey + ":" +
                        ((typeof optionValue === "boolean") ? (optionValue ? "true" : "false") : String(optionValue)));
                }
            }
            return "{" + optionParts.join(",") + "}";
        }

        /*
        worker をメインエンジンで同期実行し、戻り値のマーカー文字列を返す
        Run a worker call in the main engine synchronously and return its marker string
        */
        function callMainEngine(callExpression) {
            var payload = buildWorkerSource() + "\n" + callExpression + ";";
            var bridgeResult = { value: null, done: false };
            var bridgeMessage = new BridgeTalk();

            bridgeMessage.target = "illustrator";
            bridgeMessage.body = 'eval(decodeURIComponent("' + encodeURIComponent(payload) + '"));';

            bridgeMessage.onResult = function(resultMessage) {
                bridgeResult.value = String(resultMessage.body);
                bridgeResult.done = true;
            };
            bridgeMessage.onError = function(errorMessage) {
                bridgeResult.value = "ERR:" + String(errorMessage.body);
                bridgeResult.done = true;
            };
            bridgeMessage.onTimeout = function() {
                bridgeResult.value = "TIMEOUT";
                bridgeResult.done = true;
            };

            try {
                bridgeMessage.send(BRIDGE_TIMEOUT_SEC);
            } catch (e) {
                /*
                送信前に失敗している＝ドキュメントは無傷。ERR とは区別して返す
                The call never left the palette, so the document is untouched; report it apart from ERR
                */
                engineCallEndTime = nowMilliseconds();
                return "NOTSENT:" + e;
            }

            engineCallEndTime = nowMilliseconds();
            return bridgeResult.done ? bridgeResult.value : "TIMEOUT";
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
        var scaleRow, moveRow, rotateRow;
        var radiusRow, offsetRow;
        var jaggedControls, distortControls;
        var toggleAllButton, recalcButton, statusText;

        // =========================================
        // UIレイアウトの共通設定 / Shared UI layout
        // =========================================

        /* ウィンドウ・パネルの余白と間隔 / Window & panel margins and spacing */
        var WINDOW_MARGINS = 16; /* ウィンドウ外周の余白 / window margin */
        var WINDOW_SPACING = 12; /* ウィンドウ内の要素間隔 / window spacing */
        var PANEL_MARGINS = [16, 20, 16, 12]; /* パネル余白 [左,上,右,下] / panel margins */
        var PANEL_SPACING = 8; /* パネル内の要素間隔 / panel spacing */
        var PANEL_ROW_SPACING = 6; /* パネル内の行間隔 / spacing between rows inside a panel */
        var COLUMN_SPACING = 12; /* 2カラムの間隔 / gap between columns */
        var ROW_LABEL_WIDTH = 60; /* 行ラベルの幅 / row label width */
        var ROW_CHECKBOX_WIDTH = 90; /* 行チェックボックスの幅（既定）/ row checkbox width (default) */
        /*
        ラベルの長さに合わせた行チェックボックスの幅。
        英語ラベルは日本語より長いため、言語ごとに実測値を持つ
        Row checkbox widths tuned to the label length; the English labels are longer
        than the Japanese ones, so each language gets its own value
        */
        var TRANSFORM_CHECKBOX_WIDTH = (uiLang === "ja") ? 58 : 78; /* 倍率・移動・回転（2文字）/ Scale, Move, Rotate */
        var CORNER_CHECKBOX_WIDTH = (uiLang === "ja") ? 132 : 118; /* 角丸・パスのオフセット / Round corners, Offset path */
        var FIELD_WIDTH_CHARS = 4; /* 入力欄の文字数 / input field width in characters */
        var TOGGLE_ROW_TOP_MARGIN = 10; /* ［すべてON/OFF］の上の余白 / gap above the Toggle all button */
        var TOGGLE_BUTTON_TRIM = 4; /* ［すべてON/OFF］の高さを詰める量（px）/ height trimmed off the Toggle all button (px) */

        /* ウィンドウの共通設定 / Apply shared window layout */
        function applyWindowLayout(targetWindow) {
            targetWindow.orientation = "column";
            targetWindow.alignChildren = "fill";
            targetWindow.margins = WINDOW_MARGINS;
            targetWindow.spacing = WINDOW_SPACING;
        }

        /* パネルの共通設定 / Apply shared panel layout */
        function applyPanelLayout(targetPanel, spacing) {
            targetPanel.orientation = "column";
            targetPanel.alignChildren = ["fill", "top"];
            targetPanel.alignment = "fill";
            targetPanel.margins = PANEL_MARGINS;
            targetPanel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
        }

        /* 行グループの共通設定（ボタン列など） / Apply a horizontal row group */
        function applyRowLayout(rowGroup, alignment) {
            rowGroup.orientation = "row";
            rowGroup.alignment = alignment || "left";
            rowGroup.alignChildren = ["left", "center"];
            rowGroup.spacing = PANEL_SPACING;
        }

        /* 縦積みカラムを追加 / Add one vertically stacked column */
        function addColumnGroup(parentGroup) {
            var columnGroup = parentGroup.add("group");
            columnGroup.orientation = "column";
            columnGroup.alignChildren = ["fill", "top"];
            columnGroup.spacing = WINDOW_SPACING;
            return columnGroup;
        }

        /* ボタンの高さを指定 px 詰める（レイアウト確定後に呼ぶ）/ Trim a button's height by the given px (call after layout) */
        function trimButtonHeight(targetButton, trimPixels) {
            try {
                targetButton.size = [targetButton.size.width, targetButton.size.height - trimPixels];
            } catch (e) {}
        }

        /* ヘルプチップを設定 / Attach a help tip */
        function setHelpTip(control, helpTipPath) {
            control.helpTip = getLabel(helpTipPath);
            return control;
        }

        /* 行に入力欄と単位表示を足す / Append the input field and the unit label to a row */
        function addFieldAndUnit(rowGroup, defaultText, unitText, helpTipPath) {
            var inputField = rowGroup.add("edittext", undefined, defaultText);
            inputField.characters = FIELD_WIDTH_CHARS;
            rowGroup.add("statictext", undefined, unitText);
            setHelpTip(inputField, helpTipPath);
            return inputField;
        }

        /*
        「ラベル付きチェックボックス＋入力欄＋単位」の1行を作る
        Build one row: labeled checkbox + input field + unit
        */
        function addCheckboxFieldRow(parentPanel, labelPath, helpTipPath, defaultText, unitText, isChecked, checkboxWidth) {
            var rowGroup = parentPanel.add("group");
            applyRowLayout(rowGroup);

            var checkbox = setHelpTip(rowGroup.add("checkbox", undefined, labelText(labelPath)), helpTipPath);
            checkbox.preferredSize.width = (typeof checkboxWidth === "number") ? checkboxWidth : ROW_CHECKBOX_WIDTH;
            checkbox.value = isChecked;

            var inputField = addFieldAndUnit(rowGroup, defaultText, unitText, helpTipPath);
            inputField.enabled = isChecked;

            return { checkbox: checkbox, input: inputField };
        }

        /*
        「右寄せラベル＋入力欄＋単位」の1行を作る
        Build one row: right-aligned label + input field + unit
        */
        function addLabelFieldRow(parentPanel, labelPath, helpTipPath, defaultText, unitText) {
            var rowGroup = parentPanel.add("group");
            applyRowLayout(rowGroup);

            var rowLabel = setHelpTip(
                rowGroup.add("statictext", undefined, labelText(labelPath), { justify: "right" }), helpTipPath);
            rowLabel.preferredSize.width = ROW_LABEL_WIDTH;

            return { label: rowLabel, input: addFieldAndUnit(rowGroup, defaultText, unitText, helpTipPath) };
        }

        // =========================================
        // 入力値の読み取り / Reading the input values
        // =========================================

        /*
        数値として厳密に解釈する。parseFloat と違い "3abc" のような部分一致は不正とする
        Parse strictly; unlike parseFloat, a partial match such as "3abc" is rejected
        */
        function parseStrictNumber(text) {
            var trimmedText = String(text).replace(/^\s+|\s+$/g, "");
            if (!/^[+-]?(\d+(\.\d*)?|\.\d+)$/.test(trimmedText)) {
                return Number.NaN;
            }
            return parseFloat(trimmedText);
        }

        /*
        入力欄の数値を取得（OFF なら 0、負数は 0 にクランプ、不正値は NaN）
        Read a numeric field; 0 when disabled, negatives clamped to 0, NaN when invalid
        */
        function readFieldNumber(inputField, isEnabled) {
            if (!isEnabled) {
                return 0;
            }

            var value = parseStrictNumber(inputField.text);
            if (isNaN(value)) {
                return Number.NaN;
            }
            return (value < 0) ? 0 : value;
        }

        /* すべての入力欄を読み取る（不正値は NaN のまま）/ Read every numeric field, leaving invalid ones as NaN */
        function readNumericFields() {
            return {
                scaleRange: readFieldNumber(scaleRow.input, scaleRow.checkbox.value),
                moveValue: readFieldNumber(moveRow.input, moveRow.checkbox.value),
                rotateRange: readFieldNumber(rotateRow.input, rotateRow.checkbox.value),
                radiusValue: readFieldNumber(radiusRow.input, radiusRow.checkbox.value),
                offsetValue: readFieldNumber(offsetRow.input, offsetRow.checkbox.value),
                jaggedSize: readFieldNumber(jaggedControls.sizeRow.input, jaggedControls.checkbox.value),
                jaggedDetail: readFieldNumber(jaggedControls.detailRow.input, jaggedControls.checkbox.value),
                distortSize: readFieldNumber(distortControls.sizeRow.input, distortControls.checkbox.value),
                distortDetail: readFieldNumber(distortControls.detailRow.input, distortControls.checkbox.value)
            };
        }

        /* 不正な数値が混じっていないか / Whether every value in the set is a valid number */
        function hasOnlyValidNumbers(fieldValues) {
            for (var fieldKey in fieldValues) {
                if (fieldValues.hasOwnProperty(fieldKey) && isNaN(fieldValues[fieldKey])) {
                    return false;
                }
            }
            return true;
        }

        /*
        UIの入力値をまとめて読み取る。不正な数値があれば null を返す
        Read every UI value into one options object; returns null when a value is invalid
        */
        function readEffectOptions() {
            var fieldValues = readNumericFields();
            if (!hasOnlyValidNumbers(fieldValues)) {
                return null;
            }

            var unitCode = getRulerUnitCode();

            return {
                usesTransform: (scaleRow.checkbox.value || moveRow.checkbox.value || rotateRow.checkbox.value),
                scaleRange: fieldValues.scaleRange,
                rotateRange: fieldValues.rotateRange,
                movePoint: convertToPoint(fieldValues.moveValue, unitCode),
                radiusPoint: convertToPoint(fieldValues.radiusValue, unitCode),
                offsetPoint: convertToPoint(fieldValues.offsetValue, unitCode),
                usesJagged: jaggedControls.checkbox.value,
                jaggedSize: fieldValues.jaggedSize,
                jaggedDetail: normalizeDetailValue(fieldValues.jaggedDetail),
                usesDistort: distortControls.checkbox.value,
                distortSize: fieldValues.distortSize,
                distortDetail: normalizeDetailValue(fieldValues.distortDetail),
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
            statusText.text = getLabel(statusPath) + (extraText ? String(extraText) : "");
            /* 1行に収まらない状況表示はツールチップで読めるようにする / Let a truncated status line be read as a tooltip */
            statusText.helpTip = statusText.text;
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
                /* 日本語は「。」で切れるので区切りの空白を入れない / Japanese ends on a full stop and needs no separating space */
                setStatus("status.timeout", (uiLang === "ja" ? "" : " ") + getLabel("status.undoManually"));
            } else if (resultMarker.indexOf("NOTSENT:") === 0) {
                setStatus("status.error", resultMarker.substring(8));
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
                var stateParts = resultMarker.split("|");
                canSplit = (stateParts[1] === "1");
                canExpand = (stateParts[2] === "1");
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
            scaleRow.input.enabled = scaleRow.checkbox.value;
            moveRow.input.enabled = moveRow.checkbox.value;
            rotateRow.input.enabled = rotateRow.checkbox.value;
            updateRecalcButtonState();
        }

        /* ［再計算］の有効・無効を更新 / Update the Recalculate button state */
        function updateRecalcButtonState() {
            if (!recalcButton) {
                return;
            }
            recalcButton.enabled = !!(scaleRow.checkbox.value || moveRow.checkbox.value || rotateRow.checkbox.value ||
                jaggedControls.checkbox.value || distortControls.checkbox.value);
        }

        /* ラフ効果パネルの入力欄の有効・無効を更新 / Update the Roughen panel inputs */
        function updateRoughenPanelState(roughenControls) {
            var isEnabled = roughenControls.checkbox.value;
            roughenControls.sizeRow.label.enabled = isEnabled;
            roughenControls.sizeRow.input.enabled = isEnabled;
            roughenControls.detailRow.label.enabled = isEnabled;
            roughenControls.detailRow.input.enabled = isEnabled;
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

            var effectOptions = readEffectOptions();
            if (!effectOptions) {
                setStatus("status.invalidNumber");
                return;
            }

            isBusy = true;
            try {
                var callExpression = "wkApplyEffects(" + serializeEffectOptions(effectOptions) + ", " +
                    (canUndoLastApply ? "true" : "false") + ")";
                var resultMarker = callMainEngine(callExpression);

                /*
                送信できていなければドキュメントは変わっていないので、取り消し状態は据え置く。
                TIMEOUT / ERR は worker が適用を終えたかどうか判断できないため、
                ユーザー自身の操作を誤って取り消さないよう自動取り消しを止める
                A call that never went out left the document untouched, so keep the undo state.
                After TIMEOUT / ERR it is unknown whether the worker finished, so stop the
                automatic undo rather than risk reverting the user's own work
                */
                if (resultMarker.indexOf("NOTSENT:") !== 0) {
                    canUndoLastApply = (resultMarker.indexOf("OK") === 0);
                }
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

                /* 空欄は 0 から数え始める / An empty field steps from 0 */
                var value = (/^\s*$/.test(editText.text)) ? 0 : parseStrictNumber(editText.text);
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
                    value = normalizeDetailValue(value);
                }

                event.preventDefault();
                editText.text = String(value);
                refreshPreview();
            });
        }

        /* 詳細の値を整数・下限に丸める / Round a detail value to an integer at or above the minimum */
        function normalizeDetailValue(value) {
            return Math.max(ROUGHEN_DETAIL_MIN, Math.round(value));
        }

        /*
        入力欄の表示を実際に適用される値に合わせる。負値は 0、整数専用の欄は整数・下限に丸める。
        数値として読めない文字列はそのまま残し、状況表示で知らせる
        Make a field show what will actually be applied: negatives become 0 and integer-only
        fields are rounded to the minimum. Unparseable text is left as typed so that the
        status line can report it
        */
        function normalizeNumberInput(editText, isIntegerOnly) {
            var value = parseStrictNumber(editText.text);
            if (isNaN(value)) {
                return;
            }
            if (value < 0) {
                value = 0;
            }
            if (isIntegerOnly) {
                value = normalizeDetailValue(value);
            }
            editText.text = String(value);
        }

        /*
        数値欄の onChange ハンドラを作る（表示を整えてからプレビューを作り直す）
        Build an onChange handler that normalizes the field before rebuilding the preview
        */
        function makeNumberInputHandler(editText, isIntegerOnly) {
            return function() {
                normalizeNumberInput(editText, isIntegerOnly);
                refreshPreview();
            };
        }

        /*
        チェックボックス行を結線する（入力欄の有効・無効とプレビュー更新）
        Wire a checkbox row: toggle its input field, then rebuild the preview
        */
        function bindCheckboxRow(fieldRow, onToggle) {
            fieldRow.checkbox.onClick = function() {
                fieldRow.input.enabled = fieldRow.checkbox.value;
                if (onToggle) {
                    onToggle();
                }
                refreshPreview();
            };
            fieldRow.input.onChange = makeNumberInputHandler(fieldRow.input, false);
        }

        // =========================================
        // パレットの構築 / Building the palette
        // =========================================

        /* 対象パネル（分離とグループの扱い） / Target panel: split and group handling */
        function buildTargetPanel(parentGroup) {
            var targetPanel = parentGroup.add("panel", undefined, getLabel("panel.target"));
            applyPanelLayout(targetPanel, PANEL_ROW_SPACING);
            targetPanel.alignChildren = ["left", "top"];

            applyEachCheckbox = setHelpTip(
                targetPanel.add("checkbox", undefined, getLabel("checkbox.applyEach")), "helpTip.applyEach");
            applyEachCheckbox.value = INITIAL_CHECKBOX_STATES.applyEach;

            splitFillStrokeRadio = setHelpTip(
                targetPanel.add("radiobutton", undefined, getLabel("radio.splitFillStroke")), "helpTip.splitFillStroke");
            keepFillStrokeRadio = setHelpTip(
                targetPanel.add("radiobutton", undefined, getLabel("radio.keepFillStroke")), "helpTip.keepFillStroke");
            splitFillStrokeRadio.value = false;
            keepFillStrokeRadio.value = true;

            groupAtEndCheckbox = setHelpTip(
                targetPanel.add("checkbox", undefined, getLabel("checkbox.groupAtEnd")), "helpTip.groupAtEnd");
            groupAtEndCheckbox.value = false;

            applyEachCheckbox.onClick = refreshPreview;
            splitFillStrokeRadio.onClick = refreshPreview;
            keepFillStrokeRadio.onClick = refreshPreview;
            groupAtEndCheckbox.onClick = refreshPreview;
        }

        /* 変形パネル（スケール・移動・回転） / Transform panel: scale, move, rotate */
        function buildTransformPanel(parentGroup) {
            var transformPanel = parentGroup.add("panel", undefined, getLabel("panel.transform"));
            applyPanelLayout(transformPanel, PANEL_ROW_SPACING);

            var rulerUnitLabel = getRulerUnitLabel();

            /*
            ここの3つは ±値 の振れ幅なので、単位表示に ± を添えて絶対値と区別する
            These three are plus-minus ranges, so the unit label carries a ± to set them
            apart from the absolute values elsewhere
            */
            scaleRow = addCheckboxFieldRow(transformPanel, "fieldLabel.scale", "helpTip.scale",
                DEFAULT_FIELD_VALUES.scale, RANGE_PREFIX + getLabel("unit.percent"), INITIAL_CHECKBOX_STATES.scale, TRANSFORM_CHECKBOX_WIDTH);
            moveRow = addCheckboxFieldRow(transformPanel, "fieldLabel.move", "helpTip.move",
                DEFAULT_FIELD_VALUES.move, RANGE_PREFIX + rulerUnitLabel, INITIAL_CHECKBOX_STATES.move, TRANSFORM_CHECKBOX_WIDTH);
            rotateRow = addCheckboxFieldRow(transformPanel, "fieldLabel.rotate", "helpTip.rotate",
                DEFAULT_FIELD_VALUES.rotate, RANGE_PREFIX + getLabel("unit.degree"), INITIAL_CHECKBOX_STATES.rotate, TRANSFORM_CHECKBOX_WIDTH);

            bindCheckboxRow(scaleRow, updateRecalcButtonState);
            bindCheckboxRow(moveRow, updateRecalcButtonState);
            bindCheckboxRow(rotateRow, updateRecalcButtonState);

            /* ボタンは幅いっぱいに広げず、上に余白を入れる / Keep the button narrow and add a gap above it */
            var toggleButtonRow = transformPanel.add("group");
            applyRowLayout(toggleButtonRow);
            toggleButtonRow.margins = [0, TOGGLE_ROW_TOP_MARGIN, 0, 0];
            toggleAllButton = setHelpTip(
                toggleButtonRow.add("button", undefined, getLabel("button.toggleAll")), "helpTip.toggleAll");
            toggleAllButton.alignment = "left";

            toggleAllButton.onClick = function() {
                var nextValue = !(scaleRow.checkbox.value && moveRow.checkbox.value && rotateRow.checkbox.value);
                scaleRow.checkbox.value = nextValue;
                moveRow.checkbox.value = nextValue;
                rotateRow.checkbox.value = nextValue;
                updateTransformPanelState();
                refreshPreview();
            };
        }

        /* 角丸・オフセットパネル / Corners & Offset panel */
        function buildCornerPanel(parentGroup) {
            var cornerPanel = parentGroup.add("panel", undefined, getLabel("panel.corner"));
            applyPanelLayout(cornerPanel, PANEL_ROW_SPACING);

            var rulerUnitLabel = getRulerUnitLabel();

            radiusRow = addCheckboxFieldRow(cornerPanel, "checkbox.radius", "helpTip.radius",
                DEFAULT_FIELD_VALUES.radius, rulerUnitLabel, INITIAL_CHECKBOX_STATES.radius, CORNER_CHECKBOX_WIDTH);
            offsetRow = addCheckboxFieldRow(cornerPanel, "checkbox.offset", "helpTip.offset",
                DEFAULT_FIELD_VALUES.offset, rulerUnitLabel, INITIAL_CHECKBOX_STATES.offset, CORNER_CHECKBOX_WIDTH);

            bindCheckboxRow(radiusRow);
            bindCheckboxRow(offsetRow);
        }

        /*
        ラフ効果パネル（ギザギザ／歪曲で共通）
        Roughen panel, shared by the jagged and distortion variants
        */
        function buildRoughenPanel(parentGroup, panelTitlePath, applyLabelPath, applyHelpTipPath,
            isChecked, sizeDefaultText, detailDefaultText) {
            var roughenPanel = parentGroup.add("panel", undefined, getLabel(panelTitlePath));
            applyPanelLayout(roughenPanel, PANEL_ROW_SPACING);

            var applyCheckbox = setHelpTip(
                roughenPanel.add("checkbox", undefined, getLabel(applyLabelPath)), applyHelpTipPath);
            applyCheckbox.value = isChecked;

            var sizeRow = addLabelFieldRow(roughenPanel, "fieldLabel.size", "helpTip.size",
                sizeDefaultText, getLabel("unit.percent"));
            var detailRow = addLabelFieldRow(roughenPanel, "fieldLabel.detail", "helpTip.detail",
                detailDefaultText, getLabel("unit.perInch"));
            var roughenControls = { checkbox: applyCheckbox, sizeRow: sizeRow, detailRow: detailRow };

            applyCheckbox.onClick = function() {
                updateRoughenPanelState(roughenControls);
                refreshPreview();
            };
            sizeRow.input.onChange = makeNumberInputHandler(sizeRow.input, false);
            detailRow.input.onChange = makeNumberInputHandler(detailRow.input, true);

            return roughenControls;
        }

        /* ［再計算］ボタンと状況表示 / The Recalculate button and the status line */
        function buildFooterControls(parentWindow) {
            var recalcButtonRow = parentWindow.add("group");
            applyRowLayout(recalcButtonRow, ["fill", "top"]);

            recalcButton = setHelpTip(
                recalcButtonRow.add("button", undefined, getLabel("button.recalc")), "helpTip.recalc");
            recalcButton.alignment = "left";
            recalcButton.onClick = refreshPreview;

            /*
            幅は上のカラムに合わせる。固定幅にするとこの1行がウィンドウ幅の下限を決めてしまう
            Let the columns above decide the width; a fixed width would make this single
            line dictate the minimum window width
            */
            statusText = parentWindow.add("statictext", undefined, "", { truncate: "end" });
            statusText.alignment = ["fill", "bottom"];
        }

        /* 上部エリア（左右2カラム）を組み立てる / Build the two-column top area */
        function buildPanelColumns(parentWindow) {
            var columnsGroup = parentWindow.add("group");
            columnsGroup.orientation = "row";
            columnsGroup.alignChildren = ["fill", "top"];
            columnsGroup.spacing = COLUMN_SPACING;

            var leftColumnGroup = addColumnGroup(columnsGroup);
            var rightColumnGroup = addColumnGroup(columnsGroup);

            buildTargetPanel(leftColumnGroup);
            buildTransformPanel(leftColumnGroup);

            buildCornerPanel(rightColumnGroup);
            jaggedControls = buildRoughenPanel(rightColumnGroup,
                "panel.jagged", "checkbox.jagged", "helpTip.jagged",
                INITIAL_CHECKBOX_STATES.jagged, DEFAULT_FIELD_VALUES.jaggedSize, DEFAULT_FIELD_VALUES.jaggedDetail);
            distortControls = buildRoughenPanel(rightColumnGroup,
                "panel.distort", "checkbox.distort", "helpTip.distort",
                INITIAL_CHECKBOX_STATES.distort, DEFAULT_FIELD_VALUES.distortSize, DEFAULT_FIELD_VALUES.distortDetail);
        }

        /* 数値欄に↑↓キーでの増減を割り当てる / Enable arrow-key stepping on the numeric fields */
        function enableArrowKeySteppers() {
            var decimalInputs = [scaleRow.input, moveRow.input, rotateRow.input, radiusRow.input, offsetRow.input,
                jaggedControls.sizeRow.input, distortControls.sizeRow.input
            ];
            var integerInputs = [jaggedControls.detailRow.input, distortControls.detailRow.input];

            for (var i = 0; i < decimalInputs.length; i++) {
                changeValueByArrowKey(decimalInputs[i], false);
            }
            for (var j = 0; j < integerInputs.length; j++) {
                changeValueByArrowKey(integerInputs[j], true);
            }
        }

        /* パレット自体のイベントを結線 / Wire the palette-level events */
        function attachPaletteEvents(targetWindow) {
            /* Esc で閉じる / Close on Escape */
            targetWindow.addEventListener("keydown", function(event) {
                if (event.keyName === "Escape") {
                    targetWindow.close();
                }
            });

            /*
            パレットから離れたら、その時点の結果を確定する。
            こうしておかないと、ユーザーがドキュメント側で行った操作を
            次の再計算の app.undo() が取り消してしまう。
            Finalize the current result when the palette loses focus; otherwise the
            next app.undo() would revert whatever the user did in the document.
            */
            targetWindow.onDeactivate = function() {
                if (isBusy) {
                    return;
                }
                /* BridgeTalk 往復に伴う一時的な離脱は無視 / Ignore the transient focus loss caused by BridgeTalk */
                if ((nowMilliseconds() - engineCallEndTime) < DEACTIVATE_GUARD_MS) {
                    return;
                }
                commitLastApply();
            };

            /* 閉じても結果はそのまま残す / Closing keeps the current result */
            targetWindow.onClose = function() {
                commitLastApply();
                keepPalette(null);
                return true;
            };
        }

        /*
        パレットを組み立てる
        Assemble the palette window
        */
        function createPalette() {
            var sketchPalette = new Window("palette", getLabel("dialog.title") + " " + SCRIPT_VERSION, undefined, { resizeable: false });
            applyWindowLayout(sketchPalette);

            buildPanelColumns(sketchPalette);
            buildFooterControls(sketchPalette);
            enableArrowKeySteppers();
            attachPaletteEvents(sketchPalette);

            /* 初期状態を反映 / Apply the initial state */
            updateTransformPanelState();
            updateRoughenPanelState(jaggedControls);
            updateRoughenPanelState(distortControls);
            updateTargetPanelState(null);
            setStatus("status.ready");

            return sketchPalette;
        }

        /*
        パレットを表示（既存のパレットがあれば閉じてから）
        Show the palette, closing an existing one first
        */
        function showPalette() {
            var keptPalette = getKeptPalette();
            if (keptPalette) {
                try {
                    keptPalette.close();
                } catch (e) {}
            }

            var sketchPalette = createPalette();
            keepPalette(sketchPalette);
            sketchPalette.show();

            /* レイアウト確定後にボタンの高さを詰める / Trim the button height once the layout is settled */
            trimButtonHeight(toggleAllButton, TOGGLE_BUTTON_TRIM);

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

})();
