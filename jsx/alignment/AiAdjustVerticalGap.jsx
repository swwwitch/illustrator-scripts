#target illustrator
#targetengine "AdjustVerticalGap"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択した2つのオブジェクトの上下の間隔を、指定した値にそろえる常駐パレットです。
ライブプレビューに対応し、設定を変えるたびに結果を確認できます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AiAdjustVerticalGap.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n8201294835f9

### Overview

A persistent palette that sets the vertical gap between two selected objects to a value you specify.
A live preview shows the result as you change the settings.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AiAdjustVerticalGap.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AiAdjustVerticalGap";          /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.3.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-06-28";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-22";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AiAdjustVerticalGap.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AiAdjustVerticalGap.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n8201294835f9"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    var DEFAULT_GAP_VALUE = "3"; /* 間隔の初期値（定規の単位）/ Default gap (ruler unit) */

    // =========================================
    // レイアウト / Layout
    // =========================================

    var PANEL_MARGINS = [16, 20, 16, 12];  /* パネルの余白 / panel margins */
    var PANEL_SPACING = 8;                 /* パネル・グループ内の間隔 / spacing inside panels and groups */
    var FIELD_CHARS = 5;                   /* 数値欄の幅（文字数）/ width of the number fields */

    /**
     * パネルの共通設定
     * @param {Panel} targetPanel - 対象のパネル
     * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function setupPanel(targetPanel, spacing) {
        targetPanel.orientation = "column";
        targetPanel.alignChildren = ["fill", "top"];
        targetPanel.alignment = "fill";
        targetPanel.margins = PANEL_MARGINS;
        targetPanel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /**
     * グループの共通設定（row は縦中央、column は左揃え）
     * @param {Group} targetGroup - 対象のグループ
     * @param {string} [orientation] - "row" または "column"（既定）
     * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function setupGroup(targetGroup, orientation, spacing) {
        var groupOrientation = orientation || "column";
        targetGroup.orientation = groupOrientation;
        /* row は横並びなので縦中央、column は縦並びなので左揃え / row: vertically centered, column: left-aligned */
        targetGroup.alignChildren = (groupOrientation === "row") ? ["left", "center"] : ["left", "top"];
        targetGroup.alignment = "fill";
        targetGroup.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    // =========================================
    // 常駐パレットの参照 / Resident palette reference
    // =========================================

    /* パレットの参照を常駐エンジンに保持 / Keep the palette reference alive in the resident engine */
    var paletteWindow = null;

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * 現在の言語を判定する
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = getCurrentLang();

    var LABELS = {
        dialog: {
            title: { ja: "上下間隔を調整", en: "Adjust Vertical Gap" }
        },
        panel: {
            anchor: { ja: "キーオブジェクト", en: "Key object" },
            gap: { ja: "上下間隔", en: "Vertical Gap" },
            align: { ja: "横方向の整列", en: "Horizontal Alignment" },
            justify: { ja: "テキストの行揃え", en: "Text alignment" }
        },
        radio: {
            anchorTop: { ja: "上", en: "Top" },
            anchorBottom: { ja: "下", en: "Bottom" },
            anchorAuto: { ja: "自動判定", en: "Auto" },
            alignNone: { ja: "なし", en: "None" },
            alignLeft: { ja: "左", en: "Left" },
            alignCenter: { ja: "中央", en: "Center" },
            alignRight: { ja: "右", en: "Right" },
            justifyNone: { ja: "変更しない", en: "Keep current" },
            justifyLink: { ja: "整列に連動", en: "Match alignment" },
            justifyFull: { ja: "均等配置（最終行左）", en: "Justify (last line left-aligned)" }
        },
        checkbox: {
            previewBounds: { ja: "プレビュー境界", en: "Use preview bounds" }
        },
        fieldLabel: {
            alignAdjust: { ja: "横調整", en: "Offset" }
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

    /**
     * ドット区切りのパスでラベルを取得する（途中の欠落・null にも耐える）
     * @param {string} labelPath - "panel.gap" のようなドット区切りのキー
     * @returns {string} 現在の言語のラベル（見つからなければ labelPath そのもの）
     */
    function getLabel(labelPath) {
        var pathKeys = labelPath.split(".");
        var labelNode = LABELS;
        for (var i = 0; i < pathKeys.length; i++) {
            if (labelNode === null || typeof labelNode !== "object") {
                return labelPath; /* 途中階層が辿れない / cannot descend further */
            }
            labelNode = labelNode[pathKeys[i]];
        }
        if (labelNode === null || typeof labelNode !== "object") {
            return labelPath; /* 葉が { ja, en } でない / leaf is not a localized object */
        }
        return labelNode[uiLang] || labelNode.en || labelPath;
    }

    /**
     * コロン付きの項目名を返す（日本語は全角、英語は半角）
     * @param {string} labelPath - ラベルのパス
     * @returns {string} コロン付きの項目名
     */
    function labelText(labelPath) {
        return getLabel(labelPath) + (uiLang === "ja" ? "：" : ":");
    }

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

    /**
     * 環境設定キーの単位を返す
     * @param {string} [prefKey] - "rulerType"（既定）/ "strokeUnits" / "text/units" / "text/asianunits"
     * @returns {{code: number, label: string, pointsPerUnit: number}} 単位の情報
     */
    function getUnitInfo(prefKey) {
        var unitCode = app.preferences.getIntegerPreference(prefKey || "rulerType");
        /* 未知のコードは pt に寄せる / unknown codes fall back to points */
        var unit = UNITS[unitCode] || UNITS[2];
        return { code: unitCode, label: unit.label, pointsPerUnit: unit.pointsPerUnit };
    }

    // =========================================
    // 委譲される処理 / Delegated worker functions
    //   メインエンジンへ送って実行する。toString() で連結するため、
    //   この節の関数には JSDoc を付けず、// コメントを使わず /* */ のみ・必ずセミコロンで終える。
    //   Sent to the main engine: no JSDoc, only /* */ comments and explicit semicolons here.
    // =========================================

    /* クリップグループのクリップパスを返す（無ければ null）/ Return the clipping path of a clip group, or null */
    function getClippingPath(groupItem) {
        var groupPaths = groupItem.pathItems;
        for (var i = 0; i < groupPaths.length; i++) {
            if (groupPaths[i].clipping === true) {
                return groupPaths[i];
            }
        }
        /* 複合パスでクリップしている場合 / When the clip path is a compound path */
        var compoundPaths = groupItem.compoundPathItems;
        for (var j = 0; j < compoundPaths.length; j++) {
            if (compoundPaths[j].pathItems.length > 0 && compoundPaths[j].pathItems[0].clipping === true) {
                return compoundPaths[j];
            }
        }
        return null;
    }

    /* 計算に使う境界を取得（クリップグループはクリップパス基準）/ Bounds for calculation (clip group uses its clipping path) */
    function getItemBounds(targetItem, usePreviewBounds) {
        var boundsTarget = targetItem;
        if (targetItem.constructor.name === "GroupItem" && targetItem.clipped === true) {
            var clipPath = getClippingPath(targetItem);
            if (clipPath !== null) {
                boundsTarget = clipPath;
            }
        }
        return usePreviewBounds ? boundsTarget.visibleBounds : boundsTarget.geometricBounds;
    }

    /* UI の行揃え選択を実際のキーに解決（「整列に連動」は整列値をそのまま使う。「なし」なら "none"）/
       Resolve the choice ("link" reuses the align value; align "none" stays "none") */
    function resolveJustifyKey(justify, align) {
        if (justify === "link") {
            return align;
        }
        return justify;
    }

    /* 行揃えキーを Justification 列挙値に変換 / Map key to Justification enum */
    function resolveJustification(justifyKey) {
        if (justifyKey === "left") {
            return Justification.LEFT;
        }
        if (justifyKey === "center") {
            return Justification.CENTER;
        }
        if (justifyKey === "right") {
            return Justification.RIGHT;
        }
        if (justifyKey === "full") {
            return Justification.FULLJUSTIFYLASTLINELEFT;
        }
        return null;
    }

    /* TextFrame なら段落の行揃えを設定。実際に変更したら true を返す（no-op 検出用）/
       Set justification on a TextFrame; return true if it actually changed (for no-op detection) */
    function applyJustification(targetItem, justification, force) {
        if (justification === null) {
            return false;
        }
        if (targetItem.constructor.name !== "TextFrame") {
            return false;
        }
        /* 既に目的の行揃えなら何もしない（無駄な undo ステップを作らない）。
           ただし確定時（force）は取得値に頼らず必ず代入する。undo 直後の
           paragraphAttributes は前の値を返すことがあり、代入を飛ばすと戻ってしまう /
           Skip if already at the target justification (avoids a spurious undo step), except on
           commit (force): paragraphAttributes can report the pre-undo value right after an undo,
           and skipping the assignment there would leave the justification reverted */
        if (force !== true && targetItem.textRange.paragraphAttributes.justification === justification) {
            return false;
        }
        if (justification === Justification.LEFT) {
            /* Illustrator のバグで Justification.LEFT の代入は無視される（RIGHT/CENTER は可）。
               一時的に resize して段落属性をリフレッシュさせると代入が効く。
               200%->代入->50% で実寸は元に戻り、位置も保存して戻す。
               Assigning Justification.LEFT is ignored by Illustrator; a temporary resize
               refreshes the paragraph attributes so the assignment takes effect (200% then
               50% leaves the real size unchanged; position is saved and restored). */
            var savedPosition = [targetItem.position[0], targetItem.position[1]];
            targetItem.resize(200, 200);
            targetItem.textRange.paragraphAttributes.justification = Justification.LEFT;
            targetItem.resize(50, 50);
            /* 位置を戻せない種類がある / some items refuse the position assignment */
            try {
                targetItem.position = savedPosition;
            } catch (ePos) {}
            return true;
        }
        targetItem.textRange.paragraphAttributes.justification = justification;
        return true;
    }

    /* 対象2点を選択から解決（2個選択 or 2点入りグループ1個）。解決不可は null /
       Resolve the two target items (two selected, or a single non-clip group of two). Null if unresolved */
    function resolveTargetPair(currentSelection) {
        if (currentSelection.length === 2) {
            return [currentSelection[0], currentSelection[1]];
        }
        /* 2点を含む通常グループ1つ（クリップグループは1オブジェクト扱いなので除外）/
           One regular group of exactly two items (clip groups count as a single object, so excluded) */
        if (currentSelection.length === 1 && currentSelection[0].constructor.name === "GroupItem" && currentSelection[0].clipped !== true) {
            var groupChildren = currentSelection[0].pageItems;
            if (groupChildren.length === 2) {
                return [groupChildren[0], groupChildren[1]];
            }
        }
        return null;
    }

    /* 選択2点の現在の上下間隔を pt で返す（メインエンジンで実行）/ Return the current vertical gap (pt) of the two selected items */
    function measureGap(adjustOptions) {
        if (app.documents.length === 0) {
            return "NODOC";
        }
        var targetPair = resolveTargetPair(app.activeDocument.selection);
        if (targetPair === null) {
            return "NOSEL";
        }
        var boundsA = getItemBounds(targetPair[0], adjustOptions.usePreviewBounds);
        var boundsB = getItemBounds(targetPair[1], adjustOptions.usePreviewBounds);
        /* top が大きい方が上 / The object with the larger top is the upper one */
        var upperBounds, lowerBounds;
        if (boundsA[1] >= boundsB[1]) {
            upperBounds = boundsA;
            lowerBounds = boundsB;
        } else {
            upperBounds = boundsB;
            lowerBounds = boundsA;
        }
        /* 上のオブジェクトの下辺 − 下のオブジェクトの上辺（重なりは負）/
           Upper object's bottom edge − lower object's top edge (negative when overlapping) */
        return String(upperBounds[3] - lowerBounds[1]);
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
    function applyToPair(itemA, itemB, adjustOptions) {
        var changed = false;
        var MOVE_EPSILON = 0.0001;

        /* 行揃えを先に適用（ポイント文字は揃えで境界が変わるため）/ Justify first; point-text bounds depend on it */
        var justifyKey = resolveJustifyKey(adjustOptions.justify, adjustOptions.align);
        if (justifyKey !== "none") {
            var justification = resolveJustification(justifyKey);
            if (applyJustification(itemA, justification, adjustOptions.forceJustify)) {
                changed = true;
            }
            if (applyJustification(itemB, justification, adjustOptions.forceJustify)) {
                changed = true;
            }
        }

        var boundsA = getItemBounds(itemA, adjustOptions.usePreviewBounds);
        var boundsB = getItemBounds(itemB, adjustOptions.usePreviewBounds);

        /* top が大きい方が上 / The object with the larger top is the upper one */
        var upperItem, lowerItem;
        if (boundsA[1] >= boundsB[1]) {
            upperItem = itemA;
            lowerItem = itemB;
        } else {
            upperItem = itemB;
            lowerItem = itemA;
        }

        var anchorItem = adjustOptions.anchorTop ? upperItem : lowerItem;
        var movingItem = adjustOptions.anchorTop ? lowerItem : upperItem;

        /* 上下方向の移動（移動量が実質ゼロなら translate しない）/ Vertical move (skip if the delta is effectively zero) */
        var dy;
        if (adjustOptions.anchorTop) {
            var targetLowerTop = getItemBounds(upperItem, adjustOptions.usePreviewBounds)[3] - adjustOptions.gapPoints;
            dy = targetLowerTop - getItemBounds(movingItem, adjustOptions.usePreviewBounds)[1];
        } else {
            var targetUpperBottom = getItemBounds(lowerItem, adjustOptions.usePreviewBounds)[1] + adjustOptions.gapPoints;
            dy = targetUpperBottom - getItemBounds(movingItem, adjustOptions.usePreviewBounds)[3];
        }
        if (Math.abs(dy) > MOVE_EPSILON) {
            movingItem.translate(0, dy);
            changed = true;
        }

        /* 左右方向の整列（同上）/ Horizontal alignment (same zero-skip) */
        if (adjustOptions.align !== "none") {
            var anchorBounds = getItemBounds(anchorItem, adjustOptions.usePreviewBounds);
            var movingBounds = getItemBounds(movingItem, adjustOptions.usePreviewBounds);
            var dx = 0;
            if (adjustOptions.align === "left") {
                dx = anchorBounds[0] - movingBounds[0];
            } else if (adjustOptions.align === "right") {
                dx = anchorBounds[2] - movingBounds[2];
            } else if (adjustOptions.align === "center") {
                dx = ((anchorBounds[0] + anchorBounds[2]) / 2) - ((movingBounds[0] + movingBounds[2]) / 2);
            }
            if (Math.abs(dx) > MOVE_EPSILON) {
                movingItem.translate(dx, 0);
                changed = true;
            }
        }

        /* 整列後の左右ずらし（正＝右／負＝左）。整列「なし」でも適用 /
           Extra horizontal offset after alignment (positive = right, negative = left); applies even when align is none */
        if (adjustOptions.adjustPoints && Math.abs(adjustOptions.adjustPoints) > MOVE_EPSILON) {
            movingItem.translate(adjustOptions.adjustPoints, 0);
            changed = true;
        }

        return changed;
    }

    /* ライブプレビュー：選択ペア1組に適用（メインエンジンで実行）/ Live preview: apply to the single selected pair */
    function runAdjustment(adjustOptions) {
        /* ドキュメント確認を最優先（undo より先）/ Check for a document first, before any undo */
        if (app.documents.length === 0) {
            return "NODOC";
        }
        /* ライブプレビュー：前回適用分を取り消してからやり直す / Live preview: undo the previous apply first */
        if (adjustOptions.undoFirst === true) {
            app.undo();
        }
        var targetPair = resolveTargetPair(app.activeDocument.selection);
        if (targetPair === null) {
            app.redraw();
            return "NOSEL";
        }

        var changed = applyToPair(targetPair[0], targetPair[1], adjustOptions);

        app.redraw();
        /* 変更があれば OK（undo ステップ1つ）、無ければ NOCHANGE（undo ステップ無し）/
           OK if something changed (one undo step), otherwise NOCHANGE (no undo step) */
        return changed ? "OK" : "NOCHANGE";
    }

    /* 選択から一括適用の対象ペア群を集める（各グループの2点／グループ無しなら2点選択を1組）/
       Collect target pairs for batch apply (each group's two children; or two loose items as one pair) */
    function collectTargetPairs(currentSelection) {
        var targetPairs = [];
        for (var i = 0; i < currentSelection.length; i++) {
            if (currentSelection[i].constructor.name === "GroupItem" && currentSelection[i].clipped !== true && currentSelection[i].pageItems.length === 2) {
                targetPairs.push([currentSelection[i].pageItems[0], currentSelection[i].pageItems[1]]);
            }
        }
        /* グループが1つも無く、ちょうど2点選択なら単一ペア / no qualifying groups but exactly two loose items */
        if (targetPairs.length === 0 && currentSelection.length === 2) {
            targetPairs.push([currentSelection[0], currentSelection[1]]);
        }
        return targetPairs;
    }

    /* 記録した設定を選択中の全対象ペアへ一括適用・確定（プレビューなし）/
       Batch-apply the recorded settings to every target pair in the selection (committed, no preview) */
    function runBatchAdjustment(adjustOptions) {
        if (app.documents.length === 0) {
            return "NODOC";
        }
        /* プレビュー分はここで取り消す（別送信で取り消すとプレビューと条件が変わる）/
           Undo the preview here, in the same send as the apply */
        if (adjustOptions.undoFirst === true) {
            app.undo();
        }
        /* 確定なので行揃えは取得値で判定せず必ず適用する / Commit: always assign the justification */
        adjustOptions.forceJustify = true;
        var targetPairs = collectTargetPairs(app.activeDocument.selection);
        if (targetPairs.length === 0) {
            app.redraw();
            return "NOSEL";
        }
        for (var i = 0; i < targetPairs.length; i++) {
            applyToPair(targetPairs[i][0], targetPairs[i][1], adjustOptions);
        }
        app.redraw();
        return "OK";
    }

    /* オブジェクトの位置を [x, y] で返す（取得できなければ null）/ Return the item position as [x, y], or null */
    function getItemPosition(targetItem) {
        /* position を持たない・読めない種類がある / some items have no readable position */
        try {
            var itemPosition = targetItem.position;
            if (!itemPosition || itemPosition.length !== 2) {
                return null;
            }
            return [Number(itemPosition[0]), Number(itemPosition[1])];
        } catch (e) {
            return null;
        }
    }

    /* 判定で動かしたオブジェクトを元の位置へ戻す / Move every probed item back to its original position */
    function restoreItemPositions(targetItems, savedPositions) {
        var restored = true;
        for (var i = 0; i < targetItems.length; i++) {
            try {
                targetItems[i].position = [savedPositions[i][0], savedPositions[i][1]];
            } catch (e) {
                /* 1つ失敗しても残りは必ず戻す / Keep restoring the rest even if one fails */
                restored = false;
            }
        }
        return restored;
    }

    /* 整列で動いたオブジェクトを候補から外す / Drop candidates that moved during an align probe */
    function rejectMovedCandidates(targetItems, savedPositions, candidates, tolerance) {
        for (var i = 0; i < targetItems.length; i++) {
            if (!candidates[i]) {
                continue;
            }
            var currentPosition = getItemPosition(targetItems[i]);
            if (currentPosition === null ||
                Math.abs(currentPosition[0] - savedPositions[i][0]) > tolerance ||
                Math.abs(currentPosition[1] - savedPositions[i][1]) > tolerance) {
                candidates[i] = false;
            }
        }
    }

    /* 他のオブジェクトを完全に内包しているか。内包している側は整列コマンドで動かないため、
       キーオブジェクトが無くても4方向すべてで残ってしまう /
       Whether the item's bounds enclose every other item: such an item never moves during the
       probes, so it survives all four of them even when no key object is set */
    function enclosesAllOthers(targetItem, targetItems) {
        var itemBounds = getItemBounds(targetItem, false);
        for (var i = 0; i < targetItems.length; i++) {
            if (targetItems[i] === targetItem) {
                continue;
            }
            var otherBounds = getItemBounds(targetItems[i], false);
            /* [left, top, right, bottom] */
            if (otherBounds[0] < itemBounds[0] || otherBounds[2] > itemBounds[2] ||
                otherBounds[1] > itemBounds[1] || otherBounds[3] < itemBounds[3]) {
                return false;
            }
        }
        return true;
    }

    /* キーオブジェクトを判定（4方向の整列で動かなかった1点）。ExtendScript には
       キーオブジェクトを取得する API が無いため、整列コマンドを一時的に実行して
       位置が変わらないオブジェクトを探し、最後に必ず元の位置へ戻す /
       Detect the key object: the single item that stays put through all four align probes.
       ExtendScript has no API for it, so align commands are run temporarily and every
       item is moved back afterwards */
    function findKeyObject(targetItems) {
        var alignCommands = ["Horizontal Align Left", "Vertical Align Top", "Horizontal Align Right", "Vertical Align Bottom"];
        var tolerance = 0.001;
        var savedPositions = [];
        var candidates = [];
        var i;
        for (i = 0; i < targetItems.length; i++) {
            savedPositions[i] = getItemPosition(targetItems[i]);
            if (savedPositions[i] === null) {
                return null;
            }
            candidates[i] = true;
        }
        /* executeMenuCommand が失敗したら判定をあきらめる / give up when a menu command fails */
        try {
            for (i = 0; i < alignCommands.length; i++) {
                /* 2回目以降だけ元位置へ戻す（初回は余計な移動をしない）/ Restore only from the second probe on */
                if (i > 0 && !restoreItemPositions(targetItems, savedPositions)) {
                    return null;
                }
                app.executeMenuCommand(alignCommands[i]);
                /* 候補が1つになっても打ち切らず、4方向すべてで動かないことを確認する /
                   Keep probing all four directions to avoid a false positive */
                rejectMovedCandidates(targetItems, savedPositions, candidates, tolerance);
                /* 候補が尽きたら残りのプローブは結果を変えられないので打ち切る /
                   No candidate left: the remaining probes cannot change the result */
                var remaining = 0;
                for (var j = 0; j < candidates.length; j++) {
                    if (candidates[j]) {
                        remaining++;
                    }
                }
                if (remaining === 0) {
                    break;
                }
            }
        } catch (e) {
            return null;
        } finally {
            /* 判定中は redraw せず、最後に一度だけ元位置へ戻す / No redraw while probing; restore once at the end */
            restoreItemPositions(targetItems, savedPositions);
        }
        var keyItem = null;
        var candidateCount = 0;
        for (i = 0; i < candidates.length; i++) {
            if (candidates[i]) {
                keyItem = targetItems[i];
                candidateCount++;
            }
        }
        /* 一意に決まり、かつ内包による居座りでないときだけ採用 /
           Accept only a unique candidate that is not merely enclosing the others */
        if (candidateCount !== 1) {
            return null;
        }
        return enclosesAllOthers(keyItem, targetItems) ? null : keyItem;
    }

    /* 選択2点のキーオブジェクトが上下どちらかを返す / Report whether the key object is the upper or lower item */
    function findKeyObjectAnchor(adjustOptions) {
        if (app.documents.length === 0) {
            return "NODOC";
        }
        var currentSelection = app.activeDocument.selection;
        /* 整列コマンドは選択に対して働くため、2点選択のときだけ判定する /
           Align commands act on the selection, so only exactly two selected items qualify */
        if (!currentSelection || currentSelection.length !== 2) {
            return "NOSEL";
        }
        var keyItem = findKeyObject(currentSelection);
        app.redraw();
        if (keyItem === null) {
            return "NONE";
        }
        var keyBounds = getItemBounds(keyItem, adjustOptions.usePreviewBounds);
        var otherBounds = getItemBounds((currentSelection[0] === keyItem) ? currentSelection[1] : currentSelection[0], adjustOptions.usePreviewBounds);
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
        enclosesAllOthers,
        findKeyObject,
        findKeyObjectAnchor
    ];

    /**
     * 設定を JS のオブジェクトリテラル文字列にする
     * @param {object} adjustOptions - 設定（anchorTop / gapPoints / align / adjustPoints / justify / usePreviewBounds / undoFirst）
     * @returns {string} オブジェクトリテラル
     */
    function optionsToLiteral(adjustOptions) {
        return "{"
            + "anchorTop:" + adjustOptions.anchorTop + ","
            + "gapPoints:" + adjustOptions.gapPoints + ","
            + "align:\"" + adjustOptions.align + "\","
            + "adjustPoints:" + (adjustOptions.adjustPoints || 0) + ","
            + "justify:\"" + adjustOptions.justify + "\","
            + "usePreviewBounds:" + adjustOptions.usePreviewBounds + ","
            + "undoFirst:" + (adjustOptions.undoFirst === true)
            + "}";
    }

    /**
     * worker 関数群のソースと呼び出し式をつなげたコードを作る
     * @param {function[]} workerFuncs - 送る worker 関数
     * @param {string} dispatchExpr - 最後に評価する呼び出し式
     * @returns {string} メインエンジンで評価するコード
     */
    function buildWorkerCode(workerFuncs, dispatchExpr) {
        var functionSources = [];
        for (var i = 0; i < workerFuncs.length; i++) {
            functionSources.push(workerFuncs[i].toString());
        }
        return functionSources.join("\n") + "\n" + dispatchExpr + ";";
    }

    /**
     * メインエンジンへ同期で委譲する（% エンコードで文字化けを防ぐ）
     * @param {string} workerCode - 評価するコード
     * @returns {string} 結果の文字列（失敗時は "ERR:" で始まる）
     */
    function delegateToMainEngine(workerCode) {
        var bridgeTalk = new BridgeTalk();
        bridgeTalk.target = "illustrator";
        bridgeTalk.body = "eval(decodeURIComponent(\"" + encodeURIComponent(workerCode) + "\"));";
        var resultHolder = { value: null };
        bridgeTalk.onResult = function (response) {
            resultHolder.value = response.body;
        };
        bridgeTalk.onError = function (response) {
            resultHolder.value = "ERR:" + response.body;
        };
        bridgeTalk.send(10); /* 同期送信（最大10秒）/ Synchronous send (up to 10s) */
        return (resultHolder.value === null) ? "ERR:timeout" : resultHolder.value;
    }

    /**
     * プレビューを適用する（前回分は worker 側で取り消す）
     * @param {object} adjustOptions - 設定
     * @returns {string} OK / NOCHANGE / NODOC / NOSEL / ERR:…
     */
    function runAdjustmentPreview(adjustOptions) {
        return delegateToMainEngine(buildWorkerCode(WORKER_FUNCS, "runAdjustment(" + optionsToLiteral(adjustOptions) + ")"));
    }

    /**
     * 設定を選択中の全対象ペアへ一括適用・確定する
     * @param {object} adjustOptions - 設定
     * @returns {string} OK / NODOC / NOSEL / ERR:…
     */
    function runBatchAdjustmentDelegate(adjustOptions) {
        return delegateToMainEngine(buildWorkerCode(WORKER_FUNCS, "runBatchAdjustment(" + optionsToLiteral(adjustOptions) + ")"));
    }

    /**
     * 選択2点の現在の間隔を測る
     * @param {object} adjustOptions - 設定（usePreviewBounds を使う）
     * @returns {string} pt の数値文字列 / NODOC / NOSEL
     */
    function measureCurrentGap(adjustOptions) {
        return delegateToMainEngine(buildWorkerCode(WORKER_FUNCS, "measureGap(" + optionsToLiteral(adjustOptions) + ")"));
    }

    /**
     * キーオブジェクトが上下どちらかを判定する
     * @param {boolean} usePreviewBounds - プレビュー境界で比べるか
     * @returns {string} TOP / BOTTOM / NONE / NODOC / NOSEL
     */
    function detectKeyObjectAnchor(usePreviewBounds) {
        return delegateToMainEngine(buildWorkerCode(KEY_OBJECT_FUNCS, "findKeyObjectAnchor({usePreviewBounds:" + (usePreviewBounds === true) + "})"));
    }

    /**
     * 直前のプレビューを取り消す
     * @returns {string} OK / NODOC / ERR:…
     */
    function revertLastPreview() {
        return delegateToMainEngine(buildWorkerCode(WORKER_FUNCS, "undoLast()"));
    }

    // =========================================
    // キー操作 / Keyboard helpers
    // =========================================

    /**
     * 現在値とキー・修飾キーから次の値を求める
     * @param {number} currentValue - 現在値
     * @param {string} keyName - 押されたキー名
     * @param {object} keyboardState - ScriptUI.environment.keyboardState
     * @returns {number|null} 次の値（↑↓以外は null）
     */
    function steppedValue(currentValue, keyName, keyboardState) {
        if (keyName !== "Up" && keyName !== "Down") {
            return null;
        }
        var isUp = (keyName === "Up");
        if (keyboardState.shiftKey) {
            /* Shiftは10の倍数にスナップ / Shift snaps to multiples of 10 */
            return isUp ? Math.ceil((currentValue + 1) / 10) * 10 : Math.floor((currentValue - 1) / 10) * 10;
        }
        if (keyboardState.altKey) {
            /* Optionは0.1ずつ（小数1桁に丸め）/ Option steps by 0.1 (rounded to 1 decimal) */
            return Math.round((currentValue + (isUp ? 0.1 : -0.1)) * 10) / 10;
        }
        /* 通常は整数グリッドへ±1スナップ（1.7→↑2.0／↓1.0）/ Default: snap to the integer grid by ±1 */
        return isUp ? Math.floor(currentValue) + 1 : Math.ceil(currentValue) - 1;
    }

    /**
     * ↑↓キーで値を増減する（Shift で ±10・10 スナップ、Option で ±0.1、通常は ±1）
     * @param {EditText} editText - 対象の入力欄
     * @param {function} [onChangeCallback] - 値を変えたあとに呼ぶ処理
     * @returns {void}
     */
    function changeValueByArrowKey(editText, onChangeCallback) {
        editText.addEventListener("keydown", function (event) {
            var currentValue = Number(editText.text);
            if (isNaN(currentValue)) return;

            var nextValue = steppedValue(currentValue, event.keyName, ScriptUI.environment.keyboardState);
            if (nextValue === null) return;

            event.preventDefault();
            editText.text = nextValue;

            /* 値変更を通知 / Notify the change */
            if (typeof onChangeCallback === "function") {
                onChangeCallback();
            }
        });
    }

    /**
     * ロック中かを返す述語を評価する（true ならパネル系ショートカットを無効にする）
     * @param {function} [isLockedFn] - ロック中なら true を返す関数
     * @returns {boolean} 無効にするなら true
     */
    function shortcutsDisabled(isLockedFn) {
        return typeof isLockedFn === "function" && isLockedFn();
    }

    /**
     * ラジオを選んでからプレビューする処理を作る
     * @param {RadioButton} targetRadio - 選ぶラジオ
     * @param {function} onPreview - プレビューの処理
     * @returns {function} キーに割り当てる処理
     */
    function pickThenPreview(targetRadio, onPreview) {
        return function () {
            targetRadio.value = true;
            onPreview();
        };
    }

    /**
     * キー → 処理の表をウィンドウに登録する（gated: true のものはロック中は無効）
     * @param {Window} targetWindow - 対象のウィンドウ
     * @param {function} isLockedFn - ロック中なら true を返す関数
     * @param {Array} keyBindings - { key, run, gated } の配列
     * @returns {void}
     */
    function addKeyHandlers(targetWindow, isLockedFn, keyBindings) {
        targetWindow.addEventListener("keydown", function (event) {
            for (var i = 0; i < keyBindings.length; i++) {
                if (keyBindings[i].key !== event.keyName) continue;
                if (keyBindings[i].gated && shortcutsDisabled(isLockedFn)) return;
                keyBindings[i].run();
                event.preventDefault();
                return;
            }
        });
    }

    // =========================================
    // パネル生成 / Panel builders
    // =========================================

    /**
     * 定規単位の数値欄（↑↓キー対応）と単位ラベルを行に追加する
     * @param {Group} parentRow - 追加先の行
     * @param {string} initialText - 初期値
     * @param {string} tooltipText - ツールチップ
     * @param {object} rulerUnit - getUnitInfo() の結果
     * @param {function} onPreview - 値が変わったときの処理
     * @returns {EditText} 追加した入力欄
     */
    function addUnitField(parentRow, initialText, tooltipText, rulerUnit, onPreview) {
        var unitField = parentRow.add("edittext", undefined, initialText);
        unitField.characters = FIELD_CHARS;
        unitField.helpTip = tooltipText;
        changeValueByArrowKey(unitField, onPreview);
        unitField.onChange = onPreview;
        parentRow.add("statictext", undefined, rulerUnit.label);
        return unitField;
    }

    /**
     * ラジオの組すべてに同じツールチップと onClick を付ける
     * @param {RadioButton[]} radioList - 対象のラジオ
     * @param {string} tooltipText - ツールチップ
     * @param {function} [clickHandler] - クリック時の処理（省略時は付けない）
     * @returns {void}
     */
    function setupRadios(radioList, tooltipText, clickHandler) {
        for (var i = 0; i < radioList.length; i++) {
            radioList[i].helpTip = tooltipText;
            if (clickHandler) radioList[i].onClick = clickHandler;
        }
    }

    /**
     * キーオブジェクトのパネル（上・下・自動判定）を追加する
     * @param {Group} parentGroup - 追加先のグループ
     * @param {function} onPreview - 上・下を選んだときの処理
     * @param {function} onAutoDetect - 自動判定を選んだときの処理
     * @returns {{anchorTopRadio: RadioButton, anchorBottomRadio: RadioButton, anchorAutoRadio: RadioButton, autoAnchorTop: boolean}} ラジオと自動判定の結果
     */
    function buildAnchorPanel(parentGroup, onPreview, onAutoDetect) {
        var anchorPanel = parentGroup.add("panel", undefined, getLabel("panel.anchor"));
        setupPanel(anchorPanel);
        anchorPanel.helpTip = getLabel("tooltip.anchorTop") + " / " + getLabel("tooltip.anchorBottom") + " / " + getLabel("tooltip.anchorAuto");

        /* 上・下・自動判定は横並び（ショートカット T/B/K はラベル非表示）/ Top, bottom and auto in a row (T/B/K shortcuts are not shown) */
        var anchorRow = anchorPanel.add("group");
        setupGroup(anchorRow, "row");
        var anchorTopRadio = anchorRow.add("radiobutton", undefined, getLabel("radio.anchorTop"));
        var anchorBottomRadio = anchorRow.add("radiobutton", undefined, getLabel("radio.anchorBottom"));
        var anchorAutoRadio = anchorRow.add("radiobutton", undefined, getLabel("radio.anchorAuto"));
        anchorAutoRadio.value = true; /* 既定は自動判定 / Auto by default */
        setupRadios([anchorTopRadio], getLabel("tooltip.anchorTop"), onPreview);
        setupRadios([anchorBottomRadio], getLabel("tooltip.anchorBottom"), onPreview);
        setupRadios([anchorAutoRadio], getLabel("tooltip.anchorAuto"), onAutoDetect);

        return {
            anchorTopRadio: anchorTopRadio,
            anchorBottomRadio: anchorBottomRadio,
            anchorAutoRadio: anchorAutoRadio,
            autoAnchorTop: true /* 自動判定の結果（上が基準なら true）/ Detection result (true when the top item is the key object) */
        };
    }

    /**
     * 上下間隔の値とプレビュー境界のパネルを追加する
     * @param {Group} parentGroup - 追加先のグループ
     * @param {object} rulerUnit - getUnitInfo() の結果
     * @param {function} onPreview - 値が変わったときの処理
     * @returns {{gapValueInput: EditText, previewBoundsCheckbox: Checkbox}} 入力欄とチェックボックス
     */
    function buildGapPanel(parentGroup, rulerUnit, onPreview) {
        var gapPanel = parentGroup.add("panel", undefined, getLabel("panel.gap"));
        setupPanel(gapPanel);
        gapPanel.helpTip = getLabel("tooltip.gap");

        var gapRow = gapPanel.add("group");
        setupGroup(gapRow, "row");
        var gapValueInput = addUnitField(gapRow, DEFAULT_GAP_VALUE, getLabel("tooltip.gap"), rulerUnit, onPreview);

        var previewBoundsCheckbox = gapPanel.add("checkbox", undefined, getLabel("checkbox.previewBounds"));
        previewBoundsCheckbox.value = true;
        previewBoundsCheckbox.helpTip = getLabel("tooltip.previewBounds");
        previewBoundsCheckbox.onClick = onPreview;

        return {
            gapValueInput: gapValueInput,
            previewBoundsCheckbox: previewBoundsCheckbox
        };
    }

    /**
     * 左右の整列のパネル（なし・左・中央・右と横調整）を追加する
     * @param {Group} parentGroup - 追加先のグループ
     * @param {object} rulerUnit - getUnitInfo() の結果
     * @param {function} onPreview - 値が変わったときの処理
     * @returns {object} none / left / center / right のラジオ、adjustInput、selectCenter()
     */
    function buildAlignPanel(parentGroup, rulerUnit, onPreview) {
        var alignPanel = parentGroup.add("panel", undefined, getLabel("panel.align"));
        setupPanel(alignPanel);
        alignPanel.helpTip = getLabel("tooltip.align");

        /* ラジオは横並び / Radios in a row */
        var alignRow = alignPanel.add("group");
        setupGroup(alignRow, "row");
        var alignRadios = {
            none: alignRow.add("radiobutton", undefined, getLabel("radio.alignNone")),
            left: alignRow.add("radiobutton", undefined, getLabel("radio.alignLeft")),
            center: alignRow.add("radiobutton", undefined, getLabel("radio.alignCenter")),
            right: alignRow.add("radiobutton", undefined, getLabel("radio.alignRight"))
        };
        alignRadios.none.value = true;
        setupRadios([alignRadios.none, alignRadios.left, alignRadios.right], getLabel("tooltip.align"), onPreview);
        setupRadios([alignRadios.center], getLabel("tooltip.align"));

        /* 整列後の左右ずらし量（正＝右／負＝左）/ Extra horizontal offset (positive = right, negative = left) */
        var adjustRow = alignPanel.add("group");
        setupGroup(adjustRow, "row");
        adjustRow.add("statictext", undefined, labelText("fieldLabel.alignAdjust"));
        var adjustInput = addUnitField(adjustRow, "0", getLabel("tooltip.alignAdjust"), rulerUnit, onPreview);
        alignRadios.adjustInput = adjustInput;

        /**
         * 「中央」を選び、横調整を0へ戻してプレビューする（マウス・キー操作共通）
         * @returns {void}
         */
        function selectCenter() {
            alignRadios.center.value = true;
            adjustInput.text = "0";
            onPreview();
        }
        alignRadios.center.onClick = selectCenter;
        alignRadios.selectCenter = selectCenter;

        return alignRadios;
    }

    /**
     * テキストの行揃えのパネル（変更しない・整列に連動・均等配置）を追加する
     * @param {Group} parentGroup - 追加先のグループ
     * @param {function} onPreview - 選択が変わったときの処理
     * @returns {{none: RadioButton, link: RadioButton, full: RadioButton}} ラジオ
     */
    function buildJustifyPanel(parentGroup, onPreview) {
        var justifyPanel = parentGroup.add("panel", undefined, getLabel("panel.justify"));
        setupPanel(justifyPanel);
        justifyPanel.helpTip = getLabel("tooltip.justify");

        var justifyRadios = {
            none: justifyPanel.add("radiobutton", undefined, getLabel("radio.justifyNone")),
            link: justifyPanel.add("radiobutton", undefined, getLabel("radio.justifyLink")),
            full: justifyPanel.add("radiobutton", undefined, getLabel("radio.justifyFull"))
        };
        justifyRadios.link.value = true;
        setupRadios([justifyRadios.none, justifyRadios.link], getLabel("tooltip.justify"), onPreview);
        /* （最終行左）はツールチップに / "(last line left)" lives in the tooltip */
        setupRadios([justifyRadios.full], getLabel("tooltip.justifyFull"), onPreview);
        return justifyRadios;
    }

    /**
     * ［記録］［適用］のボタン行を追加する
     * @param {Window} parentWindow - 追加先のパレット
     * @param {function} onRecord - ［記録］／［編集］の処理
     * @param {function} onApply - ［適用］の処理
     * @returns {Button} ［記録］ボタン（ロック時に表示名を切り替える）
     */
    function buildButtonRow(parentWindow, onRecord, onApply) {
        var btnRowGroup = parentWindow.add("group");
        btnRowGroup.alignment = "right";
        var btnRecord = btnRowGroup.add("button", undefined, getLabel("button.record"));
        btnRecord.helpTip = getLabel("tooltip.record");
        btnRecord.onClick = onRecord;
        var btnApply = btnRowGroup.add("button", undefined, getLabel("button.apply"));
        btnApply.helpTip = getLabel("tooltip.apply");
        btnApply.onClick = onApply;
        return btnRecord;
    }

    /**
     * 選択中のラジオに対応するキーを返す
     * @param {object} radioSet - キー → ラジオ
     * @param {string[]} radioKeys - 調べる順のキー（どれも選ばれていなければ先頭）
     * @returns {string} 選ばれているラジオのキー
     */
    function selectedRadioKey(radioSet, radioKeys) {
        for (var i = 0; i < radioKeys.length; i++) {
            if (radioSet[radioKeys[i]].value) {
                return radioKeys[i];
            }
        }
        return radioKeys[0];
    }

    /**
     * 上のオブジェクトを基準にするかを決める（自動判定は直近の判定結果を使う）
     * @param {object} anchorControls - buildAnchorPanel() の結果
     * @returns {boolean} 上を基準にするなら true
     */
    function resolveAnchorTop(anchorControls) {
        if (anchorControls.anchorAutoRadio.value) {
            return anchorControls.autoAnchorTop;
        }
        return anchorControls.anchorTopRadio.value;
    }

    /**
     * パレットの値を読み取り、間隔と横調整を pt に換算する
     * @param {object} paletteControls - 各パネルのコントロール
     * @param {object} rulerUnit - getUnitInfo() の結果
     * @returns {object} 設定（anchorTop / gapPoints / align / adjustPoints / justify / usePreviewBounds）
     */
    function readAdjustOptions(paletteControls, rulerUnit) {
        var gapValue = parseFloat(paletteControls.gap.gapValueInput.text);
        if (isNaN(gapValue)) {
            gapValue = parseFloat(DEFAULT_GAP_VALUE);
        }
        /* 負の値は重なり（オーバーラップ）として許容し、正規化した値を表示にも反映 /
           Negative values are allowed (objects overlap); reflect the normalized value back to the field */
        if (paletteControls.gap.gapValueInput.text !== String(gapValue)) {
            paletteControls.gap.gapValueInput.text = gapValue;
        }

        var adjustValue = parseFloat(paletteControls.align.adjustInput.text);
        if (isNaN(adjustValue)) {
            adjustValue = 0;
        }

        return {
            anchorTop: resolveAnchorTop(paletteControls.anchor),
            gapPoints: gapValue * rulerUnit.pointsPerUnit,
            align: selectedRadioKey(paletteControls.align, ["none", "left", "center", "right"]),
            adjustPoints: adjustValue * rulerUnit.pointsPerUnit,
            justify: selectedRadioKey(paletteControls.justify, ["none", "link", "full"]),
            usePreviewBounds: paletteControls.gap.previewBoundsCheckbox.value
        };
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * パレットを組み立てて表示する
     * @returns {Window} 表示したパレット
     */
    function showPalette() {
        /* 多重起動を防ぐ：既存パレットがあれば閉じる / Prevent duplicates: close any existing palette */
        if (paletteWindow) {
            /* 破棄済みのウィンドウは close() が失敗することがある / close() may fail on a disposed window */
            try {
                paletteWindow.close();
            } catch (e) {}
            paletteWindow = null;
        }

        var rulerUnit = getUnitInfo();

        var gapPalette = new Window("palette", getLabel("dialog.title") + " " + SCRIPT_VERSION, undefined, { resizeable: false });
        gapPalette.orientation = "column";
        gapPalette.alignChildren = "fill";

        var paletteControls = {};
        var previewState = { active: false }; /* プレビューが反映中か / Whether a preview is currently applied */
        var isBusy = false; /* 同期委譲中の再入防止 / Guard against re-entry during a synchronous delegation */
        var recordedOptions = null; /* 記録した設定（間隔・キー・整列・行揃え・境界）/ The recorded recipe */
        var isRecordLocked = false; /* 記録モード（パネルをロック中）か / Whether we are in recorded/locked mode */
        var settingsColumn = null;
        var btnRecord = null;

        /**
         * Illustrator のキーオブジェクトを判定して、自動判定の基準を更新する
         * @returns {boolean} 上下どちらかに決まったら true
         */
        function loadAnchorFromKeyObject() {
            var detectResult = detectKeyObjectAnchor(paletteControls.gap.previewBoundsCheckbox.value);
            if (detectResult === "TOP" || detectResult === "BOTTOM") {
                paletteControls.anchor.autoAnchorTop = (detectResult === "TOP");
                return true;
            }
            paletteControls.anchor.autoAnchorTop = true;
            if (detectResult === "NOSEL" || detectResult === "NODOC") {
                /* 2点選択以外は判定そのものができない。グループ選択や複数ペアでも「自動」を
                   残したまま既定（上）を使う / Detection needs exactly two selected items;
                   keep Auto selected for groups and batches and fall back to the default (top) */
                return false;
            }
            /* 2点選択でキーオブジェクトが無いときだけ「上」へ切り替え、実際に使う基準を見せる /
               Switch to Top only when two items are selected but no key object exists */
            paletteControls.anchor.anchorTopRadio.value = true;
            return false;
        }

        /**
         * 「自動判定」：プレビューを戻してから判定し直し、プレビューする
         * （判定は整列コマンドでドキュメントを触るため、先にプレビューを戻す）
         * @returns {void}
         */
        function refreshAutoAnchor() {
            if (isBusy) return; /* 委譲中の割り込みを無視（同期送信中もUIは動く）/ ignore clicks during a delegation */
            isBusy = true;
            try {
                paletteControls.anchor.anchorAutoRadio.value = true;
                revertActivePreview();
                loadAnchorFromKeyObject();
            } finally {
                isBusy = false;
            }
            /* updatePreview 自身が再入ガードを張るのでガードの外で呼ぶ / updatePreview sets the guard itself */
            updatePreview();
        }

        /**
         * 選択2点の現在の間隔を入力欄へ取り込む
         * @returns {boolean} 取り込めたら true
         */
        function loadGapFromSelection() {
            var measureResult = measureCurrentGap({
                anchorTop: true,
                gapPoints: 0,
                align: "none",
                justify: "none",
                usePreviewBounds: paletteControls.gap.previewBoundsCheckbox.value,
                undoFirst: false
            });
            var gapPoints = parseFloat(measureResult);
            if (isNaN(gapPoints)) {
                return false; /* NODOC / NOSEL など、計測できず初期値のまま / could not measure; keep default */
            }
            /* 重なり（負の間隔）もそのまま取り込む。pt → 定規単位、0.1単位（小数1桁）に丸め /
               Keep negative gaps (overlap) as-is; pt → ruler unit, rounded to 0.1 (1 decimal) */
            var gapValue = Math.round((gapPoints / rulerUnit.pointsPerUnit) * 10) / 10;
            paletteControls.gap.gapValueInput.text = String(gapValue);
            return true;
        }

        /**
         * ライブプレビューを更新する（前回分は worker 側で取り消す）
         * @returns {void}
         */
        function updatePreview() {
            if (isBusy) return; /* 委譲中に発火した変更は無視 / ignore changes fired mid-delegation */
            isBusy = true;
            try {
                var adjustOptions = readAdjustOptions(paletteControls, rulerUnit);
                adjustOptions.undoFirst = previewState.active;
                var previewResult = runAdjustmentPreview(adjustOptions);
                /* OK＝変更あり（undoステップ1つ）。それ以外（NOCHANGE／エラー）は undo ステップが
                   無いので active=false にし、次回 app.undo() で直前のユーザー操作を巻き戻さない /
                   OK = changed (one undo step). Otherwise (NOCHANGE / error) there is no undo step,
                   so keep active false so the next app.undo() won't revert the user's prior action. */
                previewState.active = (previewResult === "OK");
            } finally {
                /* 委譲が例外で抜けても再入ガードを必ず解除 / Always clear the re-entry guard, even on error */
                isBusy = false;
            }
        }

        /**
         * プレビュー中なら戻して、ドキュメントを元の状態にする
         * @returns {void}
         */
        function revertActivePreview() {
            if (previewState.active) {
                revertLastPreview();
                previewState.active = false;
            }
        }

        /**
         * パネルのロック表示を切り替える（ボタン名・ツールチップ・ディムを連動。ボタンエリアは常に有効）
         * @param {boolean} shouldLock - ロックするなら true
         * @returns {void}
         */
        function setLocked(shouldLock) {
            isRecordLocked = shouldLock;
            settingsColumn.enabled = !shouldLock;
            btnRecord.text = shouldLock ? getLabel("button.edit") : getLabel("button.record");
            btnRecord.helpTip = shouldLock ? getLabel("tooltip.edit") : getLabel("tooltip.record");
        }

        /**
         * 「記録」⇔「編集」の切り替え：記録時は設定を控えてパネルをディム、編集時はロック解除（プレビューは戻さない）
         * @returns {void}
         */
        function toggleRecord() {
            if (!isRecordLocked) {
                recordedOptions = readAdjustOptions(paletteControls, rulerUnit);
                /* プレビューを確定（取り消さず保持）。active=false にして、次の［適用］で
                   app.undo() が走り選択が壊れるのを防ぐ / Commit the preview (keep it) and clear
                   active so the next Apply won't app.undo() and clobber the selection */
                previewState.active = false;
                setLocked(true);
            } else {
                setLocked(false);
            }
        }

        /**
         * 「適用」：記録した設定（無ければ現在値）を選択中の全対象ペアへ一括適用・確定する
         * @returns {void}
         */
        function applyBatch() {
            var adjustOptions = recordedOptions ? recordedOptions : readAdjustOptions(paletteControls, rulerUnit);
            /* プレビュー分の取り消しは worker 側でまとめて行う（二重適用を防ぐ）/
               The worker undoes the preview in the same send (avoids double-applying it) */
            adjustOptions.undoFirst = previewState.active;
            var batchResult = runBatchAdjustmentDelegate(adjustOptions);
            /* 送信に失敗した場合はプレビューが適用されたまま残るので active を維持する /
               On a failed send the preview is still applied, so keep tracking it */
            if (String(batchResult).indexOf("ERR:") !== 0) {
                previewState.active = false;
            }
        }

        /**
         * ロック中か（記録中は true）。パネル系ショートカットの抑止に使う
         * @returns {boolean} ロック中なら true
         */
        function isLockedNow() {
            return isRecordLocked;
        }

        /* 1カラム：間隔値・キーオブジェクト・左右の整列・テキストの行揃えを縦に並べる /
           Single column: gap, key object, horizontal align, text alignment stacked vertically */
        settingsColumn = gapPalette.add("group");
        setupGroup(settingsColumn, "column");
        paletteControls.gap = buildGapPanel(settingsColumn, rulerUnit, updatePreview);
        paletteControls.anchor = buildAnchorPanel(settingsColumn, updatePreview, refreshAutoAnchor);
        paletteControls.align = buildAlignPanel(settingsColumn, rulerUnit, updatePreview);
        paletteControls.justify = buildJustifyPanel(settingsColumn, updatePreview);

        /* ボタン：記録／適用 / Buttons: Record / Apply */
        btnRecord = buildButtonRow(gapPalette, toggleRecord, applyBatch);

        /* 閉じる時：未確定のプレビューは取り消す（×・Esc 共通）/ On close: revert an uncommitted preview (X and Esc) */
        gapPalette.onClose = function () {
            revertActivePreview();
            return true;
        };

        /* キー操作：A で適用、T/B で固定対象を選択、K で自動判定、N/L/C/R で整列、S/J で行揃え、Esc で閉じる
           Keys: A applies, T/B pick the anchor, K re-detects, N/L/C/R align, S/J justify, Esc closes */
        addKeyHandlers(gapPalette, isLockedNow, [
            { key: "A", run: applyBatch },
            { key: "Escape", run: function () { gapPalette.close(); } },
            { key: "T", gated: true, run: pickThenPreview(paletteControls.anchor.anchorTopRadio, updatePreview) },
            { key: "B", gated: true, run: pickThenPreview(paletteControls.anchor.anchorBottomRadio, updatePreview) },
            { key: "K", gated: true, run: refreshAutoAnchor },
            { key: "N", gated: true, run: pickThenPreview(paletteControls.align.none, updatePreview) },
            { key: "L", gated: true, run: pickThenPreview(paletteControls.align.left, updatePreview) },
            { key: "C", gated: true, run: paletteControls.align.selectCenter },
            { key: "R", gated: true, run: pickThenPreview(paletteControls.align.right, updatePreview) },
            { key: "S", gated: true, run: pickThenPreview(paletteControls.justify.link, updatePreview) },
            { key: "J", gated: true, run: pickThenPreview(paletteControls.justify.full, updatePreview) }
        ]);

        gapPalette.center();
        gapPalette.show();
        loadAnchorFromKeyObject(); /* 「自動判定」の初期値としてキーオブジェクトを判定 / Detect the key object for the initial Auto anchor */
        loadGapFromSelection(); /* 現在の間隔を取り込む（取り込めれば初期プレビューで動かない）/ Load current gap (no movement on initial preview when available) */
        updatePreview(); /* 初期プレビュー / Initial preview */
        paletteControls.gap.gapValueInput.active = true; /* 開いたら間隔値にフォーカス / Focus the gap field on open */
        return gapPalette;
    }

    paletteWindow = showPalette();

})();
