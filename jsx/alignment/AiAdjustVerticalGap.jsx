#target illustrator
#targetengine "AdjustVerticalGap"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択した2つのオブジェクトの上下の間隔を、指定した値にそろえる常駐パレットです。
ライブプレビューに対応し、設定を変えるたびに結果を確認できます。

詳細は README を参照してください。

### Overview

A persistent palette that sets the vertical gap between two selected objects to a value you specify.
A live preview shows the result as you change the settings.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AiAdjustVerticalGap";          /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.3.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-06-28";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-08-26";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AiAdjustVerticalGap.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AiAdjustVerticalGap.md
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n8201294835f9"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    var DEFAULT_GAP_VALUE = "3"; /* 間隔の初期値（定規の単位）/ Default gap (ruler unit) */

    /* パネルの余白と間隔 / Panel margins and spacing */
    var PANEL_MARGINS = [16, 20, 16, 12];
    var PANEL_SPACING = 8;
    var COLUMN_SPACING = 12; /* 2カラムの間隔 / Gap between the two columns */

    /* パレットの参照を常駐エンジンに保持 / Keep the palette reference alive in the resident engine */
    var paletteWindow = null;

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /* 現在の言語を判定（ja / en）/ Detect current language (ja / en) */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var currentLanguage = getCurrentLang();

    var LABELS = {
        dialog: {
            title: { ja: "上下間隔を調整", en: "Adjust Vertical Gap" }
        },
        anchor: {
            title: { ja: "キーオブジェクト", en: "Key object" },
            top: { ja: "上", en: "Top" },
            bottom: { ja: "下", en: "Bottom" },
            auto: { ja: "自動判定", en: "Auto" }
        },
        gap: {
            title: { ja: "上下間隔", en: "Vertical Gap" }
        },
        checkbox: {
            previewBounds: { ja: "プレビュー境界", en: "Use preview bounds" }
        },
        align: {
            title: { ja: "横方向の整列", en: "Horizontal Alignment" },
            none: { ja: "なし", en: "None" },
            left: { ja: "左", en: "Left" },
            center: { ja: "中央", en: "Center" },
            right: { ja: "右", en: "Right" },
            adjust: { ja: "横調整", en: "Offset" }
        },
        justify: {
            title: { ja: "テキストの行揃え", en: "Text alignment" },
            none: { ja: "変更しない", en: "Keep current" },
            link: { ja: "整列に連動", en: "Match alignment" },
            full: { ja: "均等配置（最終行左）", en: "Justify (last line left-aligned)" }
        },
        button: {
            record: { ja: "記録", en: "Record" },
            edit: { ja: "編集", en: "Edit" },
            apply: { ja: "適用", en: "Apply" }
        },
        tooltip: {
            anchorTop: {
                ja: "基準にする（動かさない）キーオブジェクト。上を選択（ショートカット: T）。",
                en: "Key object to keep in place. Select top (shortcut: T)."
            },
            anchorBottom: {
                ja: "基準にする（動かさない）キーオブジェクト。下を選択（ショートカット: B）。",
                en: "Key object to keep in place. Select bottom (shortcut: B)."
            },
            anchorAuto: {
                ja: "Illustrator で設定したキーオブジェクトを判定して、上下どちらを基準にするか決めます（ショートカット: K）。整列コマンドで一時的に動かして判定し、元の位置へ戻します。判定できないときは「上」になります。",
                en: "Detect the key object set in Illustrator and use it as the anchor (shortcut: K). Align commands probe temporarily and every item is moved back. Falls back to Top when it cannot be detected."
            },
            gap: {
                ja: "上下に並ぶ2つのオブジェクト間の距離です。負の値で重なります。↑↓キーで増減（Shiftで±10／Optionで±0.1）。単位は定規に従います。",
                en: "Vertical distance between the two objects (negative values overlap them). Arrow keys change it (Shift: ±10 / Option: ±0.1). Unit follows the ruler."
            },
            previewBounds: {
                ja: "線幅や効果を含む見た目の境界で間隔を計算します。オフにすると線幅や効果を含まないパス境界で計算します。",
                en: "Calculate the gap using visual bounds, including strokes and effects. Turn off to use geometric path bounds."
            },
            align: {
                ja: "移動するオブジェクトを、キーオブジェクトの左・中央・右にそろえます（ショートカット: 整列しない=N／左=L／中央=C／右=R）。",
                en: "Align the moving object to the left, center, or right of the key object (shortcuts: none=N / left=L / center=C / right=R)."
            },
            alignAdjust: {
                ja: "移動するオブジェクトを左右へ追加でずらす量です。正の値で右へ、負の値で左へ移動します。整列「なし」でも有効です。↑↓キーで増減（Shiftで±10／Optionで±0.1）。単位は定規に従います。",
                en: "Additional horizontal offset for the moving object. Positive moves right, negative moves left. Also works when alignment is None. Arrow keys change it (Shift: ±10 / Option: ±0.1). Unit follows the ruler."
            },
            justify: {
                ja: "テキストの段落の行揃え。「整列に連動」は左右の整列（左／中央／右）に合わせます。テキスト以外には影響しません（ショートカット: 整列に連動=S／均等配置=J）。",
                en: "Paragraph alignment of text. \"Match alignment\" follows the horizontal align. Non-text objects are unaffected (shortcuts: match=S / justify=J)."
            },
            justifyFull: {
                ja: "段落を均等配置します（最終行は左揃え）。",
                en: "Justify paragraphs (last line left-aligned)."
            },
            record: {
                ja: "現在の設定（間隔・キー・整列・行揃え）を記録し、パネルをロックします。記録後、複数のグループを選択して［適用］で一括適用できます。",
                en: "Record the current settings (gap, key object, align, justify) and lock the panels. Then select multiple groups and Apply to batch-apply."
            },
            edit: {
                ja: "ロックを解除して設定を編集できるようにします。",
                en: "Unlock the panels to edit the settings again."
            },
            apply: {
                ja: "記録した設定を、選択中のすべての対象（各グループの2点／2点選択）に一括適用します（ショートカット: A）。未記録なら現在の設定を使います。",
                en: "Batch-apply the recorded settings to every target in the selection (each group of two, or two selected objects) (shortcut: A). Falls back to current settings if nothing is recorded."
            }
        }
    };

    /* ドット区切りパスでラベルを取得（途中欠落・null にも耐える）/ Look up a label by dotted path (tolerates missing or null nodes) */
    function L(path) {
        var parts = path.split(".");
        var node = LABELS;
        for (var i = 0; i < parts.length; i++) {
            if (node === null || typeof node !== "object") {
                return path; /* 途中階層が辿れない / cannot descend further */
            }
            node = node[parts[i]];
        }
        if (node === null || typeof node !== "object") {
            return path; /* 葉が { ja, en } でない / leaf is not a localized object */
        }
        return node[currentLanguage] || node.en || path;
    }

    /* コロン付きラベル（日本語は全角、英語は半角）/ Label with colon (full-width JA, half-width EN) */
    function labelText(path) {
        return L(path) + (currentLanguage === "ja" ? "：" : ":");
    }

    // =========================================
    // 単位 / Units
    // =========================================

    /* 定規の単位からラベルと pt 換算係数を求める / Resolve ruler unit label and pt factor */
    function getRulerUnitInfo() {
        var rulerUnit = app.preferences.getIntegerPreference("rulerType");
        var unitLabel = "pt";
        var unitFactor = 1.0;

        switch (rulerUnit) {
            case 0: // inch
                unitLabel = "inch";
                unitFactor = 72.0;
                break;
            case 1: // mm
                unitLabel = "mm";
                unitFactor = 72.0 / 25.4;
                break;
            case 2: // pt
                unitLabel = "pt";
                unitFactor = 1.0;
                break;
            case 3: // pica
                unitLabel = "pica";
                unitFactor = 12.0;
                break;
            case 4: // cm
                unitLabel = "cm";
                unitFactor = 72.0 / 2.54;
                break;
            case 5: // Q
                unitLabel = "Q";
                unitFactor = 72.0 / 25.4 * 0.25;
                break;
            case 6: // px
                unitLabel = "px";
                unitFactor = 1.0;
                break;
            default:
                unitLabel = "pt";
                unitFactor = 1.0;
        }

        return { label: unitLabel, factor: unitFactor };
    }

    // =========================================
    // 委譲される処理 / Delegated worker functions
    //   メインエンジンへ送って実行する。toString() で連結するため、
    //   この節の関数では // コメントを使わず /* */ のみ・必ずセミコロンで終える。
    //   Sent to the main engine; use only /* */ comments and explicit semicolons here.
    // =========================================

    /* クリップグループのクリップパスを返す（無ければ null）/ Return the clipping path of a clip group, or null */
    function getClippingPath(groupItem) {
        var paths = groupItem.pathItems;
        for (var i = 0; i < paths.length; i++) {
            if (paths[i].clipping === true) {
                return paths[i];
            }
        }
        /* 複合パスでクリップしている場合 / When the clip path is a compound path */
        var compounds = groupItem.compoundPathItems;
        for (var j = 0; j < compounds.length; j++) {
            if (compounds[j].pathItems.length > 0 && compounds[j].pathItems[0].clipping === true) {
                return compounds[j];
            }
        }
        return null;
    }

    /* 計算に使う境界を取得（クリップグループはクリップパス基準）/ Bounds for calculation (clip group uses its clipping path) */
    function getItemBounds(item, usePreviewBounds) {
        var boundsTarget = item;
        if (item.constructor.name === "GroupItem" && item.clipped === true) {
            var clipPath = getClippingPath(item);
            if (clipPath !== null) {
                boundsTarget = clipPath;
            }
        }
        return usePreviewBounds ? boundsTarget.visibleBounds : boundsTarget.geometricBounds;
    }

    /* UI の行揃え選択を実際のキーに解決（「整列に連動」は整列値を流用）/ Resolve the choice ("link" follows the align value) */
    function resolveJustifyKey(justify, align) {
        if (justify === "link") {
            return (align === "none") ? "none" : align;
        }
        return justify;
    }

    /* 行揃えキーを Justification 列挙値に変換 / Map key to Justification enum */
    function resolveJustification(key) {
        if (key === "left") {
            return Justification.LEFT;
        }
        if (key === "center") {
            return Justification.CENTER;
        }
        if (key === "right") {
            return Justification.RIGHT;
        }
        if (key === "full") {
            return Justification.FULLJUSTIFYLASTLINELEFT;
        }
        return null;
    }

    /* TextFrame なら段落の行揃えを設定。実際に変更したら true を返す（no-op 検出用）/
       Set justification on a TextFrame; return true if it actually changed (for no-op detection) */
    function applyJustification(item, justification, force) {
        if (justification === null) {
            return false;
        }
        if (item.constructor.name !== "TextFrame") {
            return false;
        }
        /* 既に目的の行揃えなら何もしない（無駄な undo ステップを作らない）。
           ただし確定時（force）は取得値に頼らず必ず代入する。undo 直後の
           paragraphAttributes は前の値を返すことがあり、代入を飛ばすと戻ってしまう /
           Skip if already at the target justification (avoids a spurious undo step), except on
           commit (force): paragraphAttributes can report the pre-undo value right after an undo,
           and skipping the assignment there would leave the justification reverted */
        if (force !== true && item.textRange.paragraphAttributes.justification === justification) {
            return false;
        }
        if (justification === Justification.LEFT) {
            /* Illustrator のバグで Justification.LEFT の代入は無視される（RIGHT/CENTER は可）。
               一時的に resize して段落属性をリフレッシュさせると代入が効く。
               200%->代入->50% で実寸は元に戻り、位置も保存して戻す。
               Assigning Justification.LEFT is ignored by Illustrator; a temporary resize
               refreshes the paragraph attributes so the assignment takes effect (200% then
               50% leaves the real size unchanged; position is saved and restored). */
            var savedPosition = [item.position[0], item.position[1]];
            item.resize(200, 200);
            item.textRange.paragraphAttributes.justification = Justification.LEFT;
            item.resize(50, 50);
            try {
                item.position = savedPosition;
            } catch (ePos) {}
            return true;
        }
        item.textRange.paragraphAttributes.justification = justification;
        return true;
    }

    /* 対象2点を選択から解決（2個選択 or 2点入りグループ1個）。解決不可は null /
       Resolve the two target items (two selected, or a single non-clip group of two). Null if unresolved */
    function resolveTargetPair(sel) {
        if (sel.length === 2) {
            return [sel[0], sel[1]];
        }
        /* 2点を含む通常グループ1つ（クリップグループは1オブジェクト扱いなので除外）/
           One regular group of exactly two items (clip groups count as a single object, so excluded) */
        if (sel.length === 1 && sel[0].constructor.name === "GroupItem" && sel[0].clipped !== true) {
            var children = sel[0].pageItems;
            if (children.length === 2) {
                return [children[0], children[1]];
            }
        }
        return null;
    }

    /* 選択2点の現在の上下間隔を pt で返す（メインエンジンで実行）/ Return the current vertical gap (pt) of the two selected items */
    function measureGap(options) {
        if (app.documents.length === 0) {
            return "NODOC";
        }
        var pair = resolveTargetPair(app.activeDocument.selection);
        if (pair === null) {
            return "NOSEL";
        }
        var boundsA = getItemBounds(pair[0], options.usePreviewBounds);
        var boundsB = getItemBounds(pair[1], options.usePreviewBounds);
        /* top が大きい方が上 / The object with the larger top is the upper one */
        var upper, lower;
        if (boundsA[1] >= boundsB[1]) {
            upper = boundsA;
            lower = boundsB;
        } else {
            upper = boundsB;
            lower = boundsA;
        }
        /* 上のオブジェクトの下辺 − 下のオブジェクトの上辺（重なりは負）/
           Upper object's bottom edge − lower object's top edge (negative when overlapping) */
        return String(upper[3] - lower[1]);
    }

    /* 直前のプレビューを取り消す（メインエンジンで実行）/ Undo the previous preview (runs in the main engine) */
    function undoLast() {
        if (app.documents.length === 0) {
            return "NODOC";
        }
        app.undo();
        app.redraw();
        return "OK";
    }

    /* 1ペア（上下2点）に間隔・整列・行揃えを適用。実際に変更したら true /
       Apply gap, align and justify to one pair; return true if anything actually changed */
    function applyToPair(itemA, itemB, options) {
        var changed = false;
        var MOVE_EPSILON = 0.0001;

        /* 行揃えを先に適用（ポイント文字は揃えで境界が変わるため）/ Justify first; point-text bounds depend on it */
        var justifyKey = resolveJustifyKey(options.justify, options.align);
        if (justifyKey !== "none") {
            var justification = resolveJustification(justifyKey);
            if (applyJustification(itemA, justification, options.forceJustify)) {
                changed = true;
            }
            if (applyJustification(itemB, justification, options.forceJustify)) {
                changed = true;
            }
        }

        var boundsA = getItemBounds(itemA, options.usePreviewBounds);
        var boundsB = getItemBounds(itemB, options.usePreviewBounds);

        /* top が大きい方が上 / The object with the larger top is the upper one */
        var upperItem, lowerItem;
        if (boundsA[1] >= boundsB[1]) {
            upperItem = itemA;
            lowerItem = itemB;
        } else {
            upperItem = itemB;
            lowerItem = itemA;
        }

        var anchorItem = options.anchorTop ? upperItem : lowerItem;
        var movingItem = options.anchorTop ? lowerItem : upperItem;

        /* 上下方向の移動（移動量が実質ゼロなら translate しない）/ Vertical move (skip if the delta is effectively zero) */
        var dy;
        if (options.anchorTop) {
            var targetLowerTop = getItemBounds(upperItem, options.usePreviewBounds)[3] - options.gapPoints;
            dy = targetLowerTop - getItemBounds(movingItem, options.usePreviewBounds)[1];
        } else {
            var targetUpperBottom = getItemBounds(lowerItem, options.usePreviewBounds)[1] + options.gapPoints;
            dy = targetUpperBottom - getItemBounds(movingItem, options.usePreviewBounds)[3];
        }
        if (Math.abs(dy) > MOVE_EPSILON) {
            movingItem.translate(0, dy);
            changed = true;
        }

        /* 左右方向の整列（同上）/ Horizontal alignment (same zero-skip) */
        if (options.align !== "none") {
            var anchorBounds = getItemBounds(anchorItem, options.usePreviewBounds);
            var movingBounds = getItemBounds(movingItem, options.usePreviewBounds);
            var dx = 0;
            if (options.align === "left") {
                dx = anchorBounds[0] - movingBounds[0];
            } else if (options.align === "right") {
                dx = anchorBounds[2] - movingBounds[2];
            } else if (options.align === "center") {
                dx = ((anchorBounds[0] + anchorBounds[2]) / 2) - ((movingBounds[0] + movingBounds[2]) / 2);
            }
            if (Math.abs(dx) > MOVE_EPSILON) {
                movingItem.translate(dx, 0);
                changed = true;
            }
        }

        /* 整列後の左右ずらし（正＝右／負＝左）。整列「なし」でも適用 /
           Extra horizontal offset after alignment (positive = right, negative = left); applies even when align is none */
        if (options.adjustPoints && Math.abs(options.adjustPoints) > MOVE_EPSILON) {
            movingItem.translate(options.adjustPoints, 0);
            changed = true;
        }

        return changed;
    }

    /* ライブプレビュー：選択ペア1組に適用（メインエンジンで実行）/ Live preview: apply to the single selected pair */
    function runAdjustment(options) {
        /* ドキュメント確認を最優先（undo より先）/ Check for a document first, before any undo */
        if (app.documents.length === 0) {
            return "NODOC";
        }
        /* ライブプレビュー：前回適用分を取り消してからやり直す / Live preview: undo the previous apply first */
        if (options.undoFirst === true) {
            app.undo();
        }
        var pair = resolveTargetPair(app.activeDocument.selection);
        if (pair === null) {
            app.redraw();
            return "NOSEL";
        }

        var changed = applyToPair(pair[0], pair[1], options);

        app.redraw();
        /* 変更があれば OK（undo ステップ1つ）、無ければ NOCHANGE（undo ステップ無し）/
           OK if something changed (one undo step), otherwise NOCHANGE (no undo step) */
        return changed ? "OK" : "NOCHANGE";
    }

    /* 選択から一括適用の対象ペア群を集める（各グループの2点／グループ無しなら2点選択を1組）/
       Collect target pairs for batch apply (each group's two children; or two loose items as one pair) */
    function collectTargetPairs(sel) {
        var pairs = [];
        for (var i = 0; i < sel.length; i++) {
            if (sel[i].constructor.name === "GroupItem" && sel[i].clipped !== true && sel[i].pageItems.length === 2) {
                pairs.push([sel[i].pageItems[0], sel[i].pageItems[1]]);
            }
        }
        /* グループが1つも無く、ちょうど2点選択なら単一ペア / no qualifying groups but exactly two loose items */
        if (pairs.length === 0 && sel.length === 2) {
            pairs.push([sel[0], sel[1]]);
        }
        return pairs;
    }

    /* 記録した設定を選択中の全対象ペアへ一括適用・確定（プレビューなし）/
       Batch-apply the recorded settings to every target pair in the selection (committed, no preview) */
    function runBatchAdjustment(options) {
        if (app.documents.length === 0) {
            return "NODOC";
        }
        /* プレビュー分はここで取り消す（別送信で取り消すとプレビューと条件が変わる）/
           Undo the preview here, in the same send as the apply */
        if (options.undoFirst === true) {
            app.undo();
        }
        /* 確定なので行揃えは取得値で判定せず必ず適用する / Commit: always assign the justification */
        options.forceJustify = true;
        var pairs = collectTargetPairs(app.activeDocument.selection);
        if (pairs.length === 0) {
            app.redraw();
            return "NOSEL";
        }
        for (var i = 0; i < pairs.length; i++) {
            applyToPair(pairs[i][0], pairs[i][1], options);
        }
        app.redraw();
        return "OK";
    }

    /* オブジェクトの位置を [x, y] で返す（取得できなければ null）/ Return the item position as [x, y], or null */
    function getItemPosition(item) {
        try {
            var itemPosition = item.position;
            if (!itemPosition || itemPosition.length !== 2) {
                return null;
            }
            return [Number(itemPosition[0]), Number(itemPosition[1])];
        } catch (e) {
            return null;
        }
    }

    /* 判定で動かしたオブジェクトを元の位置へ戻す / Move every probed item back to its original position */
    function restoreItemPositions(items, positions) {
        for (var i = 0; i < items.length; i++) {
            try {
                items[i].position = [positions[i][0], positions[i][1]];
            } catch (e) {
                return false;
            }
        }
        return true;
    }

    /* 整列で動いたオブジェクトを候補から外す / Drop candidates that moved during an align probe */
    function rejectMovedCandidates(items, positions, candidates, tolerance) {
        for (var i = 0; i < items.length; i++) {
            if (!candidates[i]) {
                continue;
            }
            var currentPosition = getItemPosition(items[i]);
            if (currentPosition === null ||
                Math.abs(currentPosition[0] - positions[i][0]) > tolerance ||
                Math.abs(currentPosition[1] - positions[i][1]) > tolerance) {
                candidates[i] = false;
            }
        }
    }

    /* キーオブジェクトを判定（4方向の整列で動かなかった1点）。ExtendScript には
       キーオブジェクトを取得する API が無いため、整列コマンドを一時的に実行して
       位置が変わらないオブジェクトを探し、最後に必ず元の位置へ戻す /
       Detect the key object: the single item that stays put through all four align probes.
       ExtendScript has no API for it, so align commands are run temporarily and every
       item is moved back afterwards */
    function findKeyObject(items) {
        var alignCommands = ["Horizontal Align Left", "Vertical Align Top", "Horizontal Align Right", "Vertical Align Bottom"];
        var tolerance = 0.001;
        var positions = [];
        var candidates = [];
        var i;
        for (i = 0; i < items.length; i++) {
            positions[i] = getItemPosition(items[i]);
            if (positions[i] === null) {
                return null;
            }
            candidates[i] = true;
        }
        try {
            for (i = 0; i < alignCommands.length; i++) {
                /* 2回目以降だけ元位置へ戻す（初回は余計な移動をしない）/ Restore only from the second probe on */
                if (i > 0 && !restoreItemPositions(items, positions)) {
                    return null;
                }
                app.executeMenuCommand(alignCommands[i]);
                /* 候補が1つになっても打ち切らず、4方向すべてで動かないことを確認する /
                   Keep probing all four directions to avoid a false positive */
                rejectMovedCandidates(items, positions, candidates, tolerance);
            }
        } catch (e) {
            return null;
        } finally {
            /* 判定中は redraw せず、最後に一度だけ元位置へ戻す / No redraw while probing; restore once at the end */
            restoreItemPositions(items, positions);
        }
        var keyItem = null;
        var candidateCount = 0;
        for (i = 0; i < candidates.length; i++) {
            if (candidates[i]) {
                keyItem = items[i];
                candidateCount++;
            }
        }
        /* 一意に決まったときだけ採用 / Accept only a unique candidate */
        return (candidateCount === 1) ? keyItem : null;
    }

    /* 選択2点のキーオブジェクトが上下どちらかを返す / Report whether the key object is the upper or lower item */
    function findKeyObjectAnchor(options) {
        if (app.documents.length === 0) {
            return "NODOC";
        }
        var sel = app.activeDocument.selection;
        /* 整列コマンドは選択に対して働くため、2点選択のときだけ判定する /
           Align commands act on the selection, so only exactly two selected items qualify */
        if (!sel || sel.length !== 2) {
            return "NOSEL";
        }
        var keyItem = findKeyObject(sel);
        app.redraw();
        if (keyItem === null) {
            return "NONE";
        }
        var keyBounds = getItemBounds(keyItem, options.usePreviewBounds);
        var otherBounds = getItemBounds((sel[0] === keyItem) ? sel[1] : sel[0], options.usePreviewBounds);
        /* top が大きい方が上 / The item with the larger top is the upper one */
        return (keyBounds[1] >= otherBounds[1]) ? "TOP" : "BOTTOM";
    }

    // =========================================
    // BridgeTalk 委譲 / BridgeTalk delegation
    // =========================================

    /* 委譲する worker 関数（追加したらここにも登録）/ Worker functions to delegate (register new ones here) */
    var WORKER_FUNCS = [
        getClippingPath,
        getItemBounds,
        resolveJustifyKey,
        resolveJustification,
        applyJustification,
        resolveTargetPair,
        measureGap,
        undoLast,
        applyToPair,
        runAdjustment,
        collectTargetPairs,
        runBatchAdjustment
    ];

    /* キーオブジェクト判定だけで使う worker 関数（プレビューの送信量を増やさないため別立て）/
       Worker functions used only by key-object detection (kept apart so previews stay small) */
    var KEY_OBJECT_FUNCS = [
        getClippingPath,
        getItemBounds,
        getItemPosition,
        restoreItemPositions,
        rejectMovedCandidates,
        findKeyObject,
        findKeyObjectAnchor
    ];

    /* options を JS リテラル文字列に変換 / Serialize options to a JS object literal */
    function optionsToLiteral(options) {
        return "{"
            + "anchorTop:" + options.anchorTop + ","
            + "gapPoints:" + options.gapPoints + ","
            + "align:\"" + options.align + "\","
            + "adjustPoints:" + (options.adjustPoints || 0) + ","
            + "justify:\"" + options.justify + "\","
            + "usePreviewBounds:" + options.usePreviewBounds + ","
            + "undoFirst:" + (options.undoFirst === true)
            + "}";
    }

    /* worker 関数群 + ディスパッチ式のコードを組み立てる / Build worker source + a dispatch expression */
    function buildWorkerCode(workerFuncs, dispatchExpr) {
        var bodies = [];
        for (var i = 0; i < workerFuncs.length; i++) {
            bodies.push(workerFuncs[i].toString());
        }
        return bodies.join("\n") + "\n" + dispatchExpr + ";";
    }

    /* メインエンジンへ同期委譲（%エンコードで文字化けを防ぐ）/ Delegate synchronously, %-encoded to avoid corruption */
    function delegateToMainEngine(code) {
        var bridge = new BridgeTalk();
        bridge.target = "illustrator";
        bridge.body = "eval(decodeURIComponent(\"" + encodeURIComponent(code) + "\"));";
        var holder = { value: null };
        bridge.onResult = function (response) {
            holder.value = response.body;
        };
        bridge.onError = function (response) {
            holder.value = "ERR:" + response.body;
        };
        bridge.send(10); // 同期送信（最大10秒）/ Synchronous send (up to 10s)
        return (holder.value === null) ? "ERR:timeout" : holder.value;
    }

    /* プレビューを適用（前回分は worker 側で取り消し）/ Apply a preview (worker undoes the previous one) */
    function runAdjustmentPreview(options) {
        return delegateToMainEngine(buildWorkerCode(WORKER_FUNCS, "runAdjustment(" + optionsToLiteral(options) + ")"));
    }

    /* 記録設定を選択中の全対象ペアへ一括適用・確定 / Batch-apply recorded settings to all target pairs */
    function runBatchAdjustmentDelegate(options) {
        return delegateToMainEngine(buildWorkerCode(WORKER_FUNCS, "runBatchAdjustment(" + optionsToLiteral(options) + ")"));
    }

    /* 選択2点の現在の間隔を計測（pt 文字列 / NODOC / NOSEL）/ Measure the current gap (pt string / NODOC / NOSEL) */
    function measureCurrentGap(options) {
        return delegateToMainEngine(buildWorkerCode(WORKER_FUNCS, "measureGap(" + optionsToLiteral(options) + ")"));
    }

    /* キーオブジェクトを判定（TOP / BOTTOM / NONE / NODOC / NOSEL）/ Detect the key object (TOP / BOTTOM / NONE / NODOC / NOSEL) */
    function detectKeyObjectAnchor(usePreviewBounds) {
        return delegateToMainEngine(buildWorkerCode(KEY_OBJECT_FUNCS, "findKeyObjectAnchor({usePreviewBounds:" + (usePreviewBounds === true) + "})"));
    }

    /* 直前のプレビューを取り消す / Revert the last preview */
    function revertLastPreview() {
        return delegateToMainEngine(buildWorkerCode(WORKER_FUNCS, "undoLast()"));
    }

    // =========================================
    // テキストフィールド操作 / Text field helpers
    // =========================================

    /* 現在値とキー・修飾キーから次の値を求める（↑↓以外は null）/
       Compute the next value from the current value and modifiers (null for keys other than Up/Down) */
    function steppedValue(value, keyName, keyboard) {
        if (keyName !== "Up" && keyName !== "Down") {
            return null;
        }
        var up = (keyName === "Up");
        if (keyboard.shiftKey) {
            /* Shiftは10の倍数にスナップ / Shift snaps to multiples of 10 */
            return up ? Math.ceil((value + 1) / 10) * 10 : Math.floor((value - 1) / 10) * 10;
        }
        if (keyboard.altKey) {
            /* Optionは0.1ずつ（小数1桁に丸め）/ Option steps by 0.1 (rounded to 1 decimal) */
            return Math.round((value + (up ? 0.1 : -0.1)) * 10) / 10;
        }
        /* 通常は整数グリッドへ±1スナップ（1.7→↑2.0／↓1.0）/ Default: snap to the integer grid by ±1 */
        return up ? Math.floor(value) + 1 : Math.ceil(value) - 1;
    }

    /* ↑↓キーで値を増減（Shiftで±10・10スナップ、Optionで±0.1、通常は±1）/ Arrow keys change the value (Shift: ±10 snapped, Option: ±0.1, default: ±1) */
    function changeValueByArrowKey(editText, onChangeCallback) {
        editText.addEventListener("keydown", function (event) {
            var value = Number(editText.text);
            if (isNaN(value)) return;

            var next = steppedValue(value, event.keyName, ScriptUI.environment.keyboardState);
            if (next === null) return;

            event.preventDefault();
            editText.text = next;

            /* 値変更を通知 / Notify the change */
            if (typeof onChangeCallback === "function") {
                onChangeCallback();
            }
        });
    }

    /* ロック中はパネル系ショートカットを無効化する述語（true で無効）/ When this returns true, panel shortcuts are disabled */
    function shortcutsDisabled(isLocked) {
        return typeof isLocked === "function" && isLocked();
    }

    /* ラジオを選んでプレビューするアクションを作る / Build an action that selects a radio then previews */
    function pickThenPreview(radio, onPreview) {
        return function () {
            radio.value = true;
            onPreview();
        };
    }

    /* キー→アクションのテーブルを win に登録（gated:true のものはロック中無効）/
       Register a key→action table on the window (entries with gated:true are disabled while locked) */
    function addKeyHandlers(win, isLocked, bindings) {
        win.addEventListener("keydown", function (event) {
            for (var i = 0; i < bindings.length; i++) {
                if (bindings[i].key !== event.keyName) continue;
                if (bindings[i].gated && shortcutsDisabled(isLocked)) return;
                bindings[i].run();
                event.preventDefault();
                return;
            }
        });
    }

    // =========================================
    // レイアウト共通設定 / Shared layout helpers
    // =========================================

    /* パネルの共通設定 / Apply shared panel layout */
    function setupPanel(panel, spacing) {
        panel.orientation = "column";
        panel.alignChildren = ["fill", "top"];
        panel.alignment = "fill";
        panel.margins = PANEL_MARGINS;
        panel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /* グループの共通設定（row/column で整列を切り替え）/ Apply shared group layout (alignChildren switches by orientation) */
    function setupGroup(group, orientation, spacing) {
        var groupOrientation = orientation || "column";
        group.orientation = groupOrientation;
        /* row は横並びなので縦中央、column は縦並びなので左揃え / row: vertically centered, column: left-aligned */
        group.alignChildren = (groupOrientation === "row") ? ["left", "center"] : ["left", "top"];
        group.alignment = "fill";
        group.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    // =========================================
    // パネル生成 / Panel builders
    // =========================================

    /* 固定するオブジェクトパネル / Fixed-object panel */
    function buildAnchorPanel(parent, onPreview, onAutoDetect) {
        var panel = parent.add("panel", undefined, L("anchor.title"));
        setupPanel(panel);
        panel.helpTip = L("tooltip.anchorTop") + " / " + L("tooltip.anchorBottom") + " / " + L("tooltip.anchorAuto");

        /* 上・下・自動判定は横並び（ショートカット T/B/K はラベル非表示）/ Top, bottom and auto in a row (T/B/K shortcuts are not shown) */
        var row = panel.add("group");
        setupGroup(row, "row");
        var anchorTopRadio = row.add("radiobutton", undefined, L("anchor.top"));
        var anchorBottomRadio = row.add("radiobutton", undefined, L("anchor.bottom"));
        var anchorAutoRadio = row.add("radiobutton", undefined, L("anchor.auto"));
        anchorAutoRadio.value = true; /* 既定は自動判定 / Auto by default */
        anchorTopRadio.helpTip = L("tooltip.anchorTop");
        anchorBottomRadio.helpTip = L("tooltip.anchorBottom");
        anchorAutoRadio.helpTip = L("tooltip.anchorAuto");
        anchorTopRadio.onClick = onPreview;
        anchorBottomRadio.onClick = onPreview;
        anchorAutoRadio.onClick = onAutoDetect;

        return {
            anchorTopRadio: anchorTopRadio,
            anchorBottomRadio: anchorBottomRadio,
            anchorAutoRadio: anchorAutoRadio,
            autoAnchorTop: true /* 自動判定の結果（上が基準なら true）/ Detection result (true when the top item is the key object) */
        };
    }

    /* 間隔値＋プレビュー境界パネル / Gap + preview-bounds panel */
    function buildGapPanel(parent, rulerUnit, onPreview) {
        var panel = parent.add("panel", undefined, L("gap.title"));
        setupPanel(panel);
        panel.helpTip = L("tooltip.gap");

        var row = panel.add("group");
        setupGroup(row, "row");
        var gapValueInput = row.add("edittext", undefined, DEFAULT_GAP_VALUE);
        gapValueInput.characters = 5;
        gapValueInput.helpTip = L("tooltip.gap");
        changeValueByArrowKey(gapValueInput, onPreview);
        gapValueInput.onChange = onPreview;
        row.add("statictext", undefined, rulerUnit.label);

        var previewBoundsCheckbox = panel.add("checkbox", undefined, L("checkbox.previewBounds"));
        previewBoundsCheckbox.value = true;
        previewBoundsCheckbox.helpTip = L("tooltip.previewBounds");
        previewBoundsCheckbox.onClick = onPreview;

        return {
            gapValueInput: gapValueInput,
            previewBoundsCheckbox: previewBoundsCheckbox
        };
    }

    /* 左右の整列パネル / Horizontal-alignment panel */
    function buildAlignPanel(parent, rulerUnit, onPreview) {
        var panel = parent.add("panel", undefined, L("align.title"));
        setupPanel(panel);
        panel.helpTip = L("tooltip.align");

        /* ラジオは横並び / Radios in a row */
        var row = panel.add("group");
        setupGroup(row, "row");
        var radios = {
            none: row.add("radiobutton", undefined, L("align.none")),
            left: row.add("radiobutton", undefined, L("align.left")),
            center: row.add("radiobutton", undefined, L("align.center")),
            right: row.add("radiobutton", undefined, L("align.right"))
        };
        radios.none.value = true;
        radios.none.onClick = onPreview;
        radios.left.onClick = onPreview;
        radios.right.onClick = onPreview;
        radios.none.helpTip = L("tooltip.align");
        radios.left.helpTip = L("tooltip.align");
        radios.center.helpTip = L("tooltip.align");
        radios.right.helpTip = L("tooltip.align");

        /* 整列後の左右ずらし量（正＝右／負＝左）/ Extra horizontal offset (positive = right, negative = left) */
        var adjustRow = panel.add("group");
        setupGroup(adjustRow, "row");
        adjustRow.add("statictext", undefined, labelText("align.adjust"));
        var adjustInput = adjustRow.add("edittext", undefined, "0");
        adjustInput.characters = 5;
        adjustInput.helpTip = L("tooltip.alignAdjust");
        changeValueByArrowKey(adjustInput, onPreview);
        adjustInput.onChange = onPreview;
        adjustRow.add("statictext", undefined, rulerUnit.label);
        radios.adjustInput = adjustInput;

        /* 中央に設定したら横調整を0へリセット（マウス・キー操作共通）/
           Selecting center resets the horizontal adjust to 0 (shared by mouse and keyboard) */
        function selectCenter() {
            radios.center.value = true;
            adjustInput.text = "0";
            onPreview();
        }
        radios.center.onClick = selectCenter;
        radios.selectCenter = selectCenter;

        return radios;
    }

    /* テキストの行揃えパネル / Paragraph-justification panel */
    function buildJustifyPanel(parent, onPreview) {
        var panel = parent.add("panel", undefined, L("justify.title"));
        setupPanel(panel);
        panel.helpTip = L("tooltip.justify");

        var radios = {
            none: panel.add("radiobutton", undefined, L("justify.none")),
            link: panel.add("radiobutton", undefined, L("justify.link")),
            full: panel.add("radiobutton", undefined, L("justify.full"))
        };
        radios.link.value = true;
        radios.none.helpTip = L("tooltip.justify");
        radios.link.helpTip = L("tooltip.justify");
        radios.full.helpTip = L("tooltip.justifyFull"); /* （最終行左）はツールチップに / "(last line left)" lives in the tooltip */
        radios.none.onClick = onPreview;
        radios.link.onClick = onPreview;
        radios.full.onClick = onPreview;
        return radios;
    }

    /* 選択中のラジオに対応するキーを返す / Return the key of the selected radio */
    function selectedRadioKey(radios, keys) {
        for (var i = 0; i < keys.length; i++) {
            if (radios[keys[i]].value) {
                return keys[i];
            }
        }
        return keys[0];
    }

    /* 上を基準にするかを解決（自動判定は直近の判定結果を使う）/ Resolve whether the top item is the anchor (auto uses the last detection) */
    function resolveAnchorTop(anchor) {
        if (anchor.anchorAutoRadio.value) {
            return anchor.autoAnchorTop;
        }
        return anchor.anchorTopRadio.value;
    }

    // =========================================
    // パレット / Palette
    // =========================================

    /* UI 値を読み取り、間隔は pt に換算 / Read UI values, convert gap to points */
    function readOptions(controls, rulerUnit) {
        var gapValue = parseFloat(controls.gap.gapValueInput.text);
        if (isNaN(gapValue)) {
            gapValue = parseFloat(DEFAULT_GAP_VALUE);
        }
        /* 負の値は重なり（オーバーラップ）として許容 / Negative values are allowed (objects overlap) */
        /* 正規化した値を表示にも反映 / Reflect the normalized value back to the field */
        if (controls.gap.gapValueInput.text !== String(gapValue)) {
            controls.gap.gapValueInput.text = gapValue;
        }

        var adjustValue = parseFloat(controls.align.adjustInput.text);
        if (isNaN(adjustValue)) {
            adjustValue = 0;
        }

        return {
            anchorTop: resolveAnchorTop(controls.anchor),
            gapPoints: gapValue * rulerUnit.factor,
            align: selectedRadioKey(controls.align, ["none", "left", "center", "right"]),
            adjustPoints: adjustValue * rulerUnit.factor,
            justify: selectedRadioKey(controls.justify, ["none", "link", "full"]),
            usePreviewBounds: controls.gap.previewBoundsCheckbox.value
        };
    }

    /* パレットを表示 / Show the palette */
    function showPalette() {
        /* 多重起動を防ぐ：既存パレットがあれば閉じる / Prevent duplicates: close any existing palette */
        if (paletteWindow) {
            try {
                paletteWindow.close();
            } catch (e) {}
            paletteWindow = null;
        }

        var rulerUnit = getRulerUnitInfo();

        var win = new Window("palette", L("dialog.title") + " " + SCRIPT_VERSION, undefined, { resizeable: false });
        win.orientation = "column";
        win.alignChildren = "fill";

        var controls = {};
        var previewState = { active: false }; /* プレビューが反映中か / Whether a preview is currently applied */
        var isBusy = false; /* 同期委譲中の再入防止 / Guard against re-entry during a synchronous delegation */

        /* Illustrator のキーオブジェクトを判定して自動判定の基準を更新 /
           Detect Illustrator's key object and update the anchor used by Auto */
        function loadAnchorFromKeyObject() {
            var result = detectKeyObjectAnchor(controls.gap.previewBoundsCheckbox.value);
            if (result !== "TOP" && result !== "BOTTOM") {
                /* 判定できなければ「上」へ切り替え、実際に使う基準を見えるようにする /
                   Fall back to Top so the anchor actually in use stays visible */
                controls.anchor.autoAnchorTop = true;
                controls.anchor.anchorTopRadio.value = true;
                return false;
            }
            controls.anchor.autoAnchorTop = (result === "TOP");
            return true;
        }

        /* 「自動判定」：判定し直してからプレビュー（判定は整列コマンドでドキュメントを触るため、
           先にプレビューを戻す）/ Auto: re-detect then preview (the preview is reverted first
           because detection moves objects with align commands) */
        function refreshAutoAnchor() {
            if (isBusy) return; /* 委譲中の割り込みを無視（同期送信中もUIは動く）/ ignore clicks during a delegation */
            isBusy = true;
            try {
                controls.anchor.anchorAutoRadio.value = true;
                revertActivePreview();
                loadAnchorFromKeyObject();
            } finally {
                isBusy = false;
            }
            /* updatePreview 自身が再入ガードを張るのでガードの外で呼ぶ / updatePreview sets the guard itself */
            updatePreview();
        }

        /* 選択2点の現在の間隔をフィールドへ取り込む / Load the current gap of the two selected items into the field */
        function loadGapFromSelection() {
            var result = measureCurrentGap({
                anchorTop: true,
                gapPoints: 0,
                align: "none",
                justify: "none",
                usePreviewBounds: controls.gap.previewBoundsCheckbox.value,
                undoFirst: false
            });
            var gapPoints = parseFloat(result);
            if (isNaN(gapPoints)) {
                return false; /* NODOC / NOSEL など、計測できず初期値のまま / could not measure; keep default */
            }
            /* 重なり（負の間隔）もそのまま取り込む / Keep negative gaps (overlap) as-is */
            /* pt → 定規単位、0.1単位（小数1桁）に丸め / pt → ruler unit, rounded to 0.1 (1 decimal) */
            var gapValue = Math.round((gapPoints / rulerUnit.factor) * 10) / 10;
            controls.gap.gapValueInput.text = String(gapValue);
            return true;
        }

        /* ライブプレビュー更新（前回分は worker 側で取り消し）/ Refresh live preview (worker undoes the previous one) */
        function updatePreview() {
            if (isBusy) return; /* 委譲中に発火した変更は無視 / ignore changes fired mid-delegation */
            isBusy = true;
            try {
                var options = readOptions(controls, rulerUnit);
                options.undoFirst = previewState.active;
                var result = runAdjustmentPreview(options);
                /* OK＝変更あり（undoステップ1つ）。それ以外（NOCHANGE／エラー）は undo ステップが
                   無いので active=false にし、次回 app.undo() で直前のユーザー操作を巻き戻さない /
                   OK = changed (one undo step). Otherwise (NOCHANGE / error) there is no undo step,
                   so keep active false so the next app.undo() won't revert the user's prior action. */
                previewState.active = (result === "OK");
            } finally {
                /* 委譲が例外で抜けても再入ガードを必ず解除 / Always clear the re-entry guard, even on error */
                isBusy = false;
            }
        }

        /* 記録した設定（間隔・キー・整列・行揃え・境界）/ The recorded recipe */
        var recordedOptions = null;

        /* プレビュー中なら戻してドキュメントを元の状態へ / Revert any active preview so the document is clean */
        function revertActivePreview() {
            if (previewState.active) {
                revertLastPreview();
                previewState.active = false;
            }
        }

        /* 記録モード（パネルをロック中）か / Whether we are in recorded/locked mode */
        var locked = false;

        /* パネルのロック表示を切り替え（ボタン名・ツールチップ・ディムを連動。ボタンエリアは常に有効）/
           Toggle the locked (dimmed) state of the setting panels; the button area stays enabled */
        function setLocked(isLocked) {
            locked = isLocked;
            settingsColumn.enabled = !isLocked;
            recordBtn.text = isLocked ? L("button.edit") : L("button.record");
            recordBtn.helpTip = isLocked ? L("tooltip.edit") : L("tooltip.record");
        }

        /* 「記録」⇔「編集」トグル：記録時は設定を保存しパネルをディム、編集時はロック解除（プレビューは戻さない）/
           Record/Edit toggle: on record, save settings and dim panels; on edit, unlock (preview is left as-is) */
        function toggleRecord() {
            if (!locked) {
                recordedOptions = readOptions(controls, rulerUnit);
                /* プレビューを確定（取り消さず保持）。active=false にして、次の［適用］で
                   app.undo() が走り選択が壊れるのを防ぐ / Commit the preview (keep it) and clear
                   active so the next Apply won't app.undo() and clobber the selection */
                previewState.active = false;
                setLocked(true);
            } else {
                setLocked(false);
            }
        }

        /* 「適用」：記録設定（無ければ現在値）を選択中の全対象ペアへ一括適用・確定 /
           Apply: batch-apply the recorded settings (or current values) to every target pair, committed */
        function applyBatch() {
            var options = recordedOptions ? recordedOptions : readOptions(controls, rulerUnit);
            /* プレビュー分の取り消しは worker 側でまとめて行う（二重適用を防ぐ）/
               The worker undoes the preview in the same send (avoids double-applying it) */
            options.undoFirst = previewState.active;
            previewState.active = false;
            runBatchAdjustmentDelegate(options);
        }

        /* 1カラム：間隔値・キーオブジェクト・左右の整列・テキストの行揃えを縦に並べる /
           Single column: gap, key object, horizontal align, text alignment stacked vertically */
        var settingsColumn = win.add("group");
        setupGroup(settingsColumn, "column");
        controls.gap = buildGapPanel(settingsColumn, rulerUnit, updatePreview);
        controls.anchor = buildAnchorPanel(settingsColumn, updatePreview, refreshAutoAnchor);
        controls.align = buildAlignPanel(settingsColumn, rulerUnit, updatePreview);
        controls.justify = buildJustifyPanel(settingsColumn, updatePreview);

        /* ボタン：記録／適用 / Buttons: Record / Apply */
        var btnGroup = win.add("group");
        btnGroup.alignment = "right";
        var recordBtn = btnGroup.add("button", undefined, L("button.record"));
        recordBtn.helpTip = L("tooltip.record");
        recordBtn.onClick = toggleRecord;
        var applyBtn = btnGroup.add("button", undefined, L("button.apply"));
        applyBtn.helpTip = L("tooltip.apply");
        applyBtn.onClick = applyBatch;

        /* 閉じる時：未確定のプレビューは取り消す（×・Esc 共通）/ On close: revert an uncommitted preview (X and Esc) */
        win.onClose = function () {
            if (previewState.active) {
                revertLastPreview();
                previewState.active = false;
            }
            return true;
        };

        /* キー操作：A で適用、T/B で固定対象を選択、N/L/C/R で整列、S/J で行揃え、Esc で閉じる / Keys: A applies, T/B pick the anchor, N/L/C/R align, S/J justify, Esc closes */
        /* ロック判定（記録中は true）。パネル系ショートカットの抑止に使う / Locked predicate to gate panel shortcuts */
        function isLockedNow() {
            return locked;
        }
        addKeyHandlers(win, isLockedNow, [
            { key: "A", run: applyBatch },
            { key: "Escape", run: function () { win.close(); } },
            { key: "T", gated: true, run: pickThenPreview(controls.anchor.anchorTopRadio, updatePreview) },
            { key: "B", gated: true, run: pickThenPreview(controls.anchor.anchorBottomRadio, updatePreview) },
            { key: "K", gated: true, run: refreshAutoAnchor },
            { key: "N", gated: true, run: pickThenPreview(controls.align.none, updatePreview) },
            { key: "L", gated: true, run: pickThenPreview(controls.align.left, updatePreview) },
            { key: "C", gated: true, run: controls.align.selectCenter },
            { key: "R", gated: true, run: pickThenPreview(controls.align.right, updatePreview) },
            { key: "S", gated: true, run: pickThenPreview(controls.justify.link, updatePreview) },
            { key: "J", gated: true, run: pickThenPreview(controls.justify.full, updatePreview) }
        ]);

        win.center();
        win.show();
        loadAnchorFromKeyObject(); /* 「自動判定」の初期値としてキーオブジェクトを判定 / Detect the key object for the initial Auto anchor */
        loadGapFromSelection(); /* 現在の間隔を取り込む（取り込めれば初期プレビューで動かない）/ Load current gap (no movement on initial preview when available) */
        updatePreview(); /* 初期プレビュー / Initial preview */
        controls.gap.gapValueInput.active = true; /* 開いたら間隔値にフォーカス / Focus the gap field on open */
        return win;
    }

    paletteWindow = showPalette();

})();
