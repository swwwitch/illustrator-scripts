#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

2つのオブジェクトを選択し、大きい方を「容器」、小さい方を「タイル」として、容器のバウンディングボックス内にタイルを等間隔で敷き詰めます。
完全に内側に収まるタイルのみを残し、元の2オブジェクトはそのまま残します。

詳細は README を参照してください。

### Overview

With two objects selected, treats the larger as the container and the smaller as the tile, then fills the container's bounding box with an even grid of tiles.
Only the tiles that fit entirely inside are kept, and both originals are left in place.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "PatternFill";                  /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.4.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-10-26";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-19";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/PatternFill.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/PatternFill.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    /* 現在のロケールを判定 / Detect current locale */
    function getCurrentLang() {
      return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = getCurrentLang();

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
      dialogTitle: {
        // ダイアログボックスのタイトルバーに必ずバージョンを表示する形式
        // Always show version in dialog title
        ja: "敷き詰め設定",
        en: "Tile Fill Settings"
      },
      spacing: {
        ja: "間隔",
        en: "Spacing"
      },
      margin: {
        ja: "マージン",
        en: "Margin"
      },
      brick: {
        ja: "レンガ状",
        en: "Brick pattern"
      },
      symbolize: {
        ja: "シンボル化",
        en: "Symbolize"
      },
      ok: {
        ja: "OK",
        en: "OK"
      },
      cancel: {
        ja: "キャンセル",
        en: "Cancel"
      },
      tipSpacing: {
        ja: "隣り合うタイルのアキです。縦横とも同じ値になります。",
        en: "Space between neighbouring tiles, applied both horizontally and vertically."
      },
      tipMargin: {
        ja: "容器の内側に空ける余白です。負の値も入力できます。",
        en: "Inset kept inside the container. Negative values are allowed."
      },
      tipBrick: {
        ja: "1行おきに半個分ずらして、レンガのように並べます。",
        en: "Offsets every other row by half a tile, like brickwork."
      },
      tipSymbolize: {
        ja: "タイルをシンボルとして複製します。あとからまとめて差し替えられます。",
        en: "Duplicates the tile as a symbol, so every copy can be swapped later at once."
      }
    };

    /* ラベル取得関数 / Label resolver */
    function getLabel(key) {
      var entry = LABELS[key];
      if (!entry) return key;
      return entry[uiLang] || entry.en || key;
    }

    /* コロン付きの項目名を返す（日本語は全角、英語は半角） / Return a label with a colon */
    function labelText(key) {
      return getLabel(key) + (uiLang === "ja" ? "：" : ": ");
    }

    /* 単位を括弧でくくった表記を返す（日本語は全角括弧） / Return a parenthesised unit label */
    function unitSuffix(unitLabel) {
      return (uiLang === "ja") ? ("（" + unitLabel + "）") : (" (" + unitLabel + ")");
    }

    // =========================================
    // 単位 / Units
    // =========================================

    /* 単位テーブル（配列の添字が rulerType コードと一致：0=in, 1=mm, 2=pt …）/ Unit table; the array index equals the rulerType code */
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

    /* Q ではなく H と表示する設定キー / Preference keys that display H instead of Q */
    var HA_UNIT_PREF_KEYS = { "rulerType": true, "strokeUnits": true, "text/asianunits": true };

    /**
     * 設定キーごとの単位情報を取得する
     * @param {string} prefKey - 環境設定キー（省略時は "rulerType"）
     * @returns {{code: number, label: string, pointsPerUnit: number}} 単位情報
     */
    function getUnitInfo(prefKey) {
        var unitKey = prefKey || "rulerType";
        var unitCode = app.preferences.getIntegerPreference(unitKey);
        var unit = UNITS[unitCode] || UNITS[2];
        var label = (unitCode === 5 && HA_UNIT_PREF_KEYS[unitKey]) ? "H" : unit.label;
        return { code: unitCode, label: label, pointsPerUnit: unit.pointsPerUnit };
    }

    /* タイルグリッド共通ユーティリティ / Tile Grid Common Utilities */
    // ---- Namespace: TG (Tile Grid) common utilities ----
    var TG = {
        /* 境界情報を取得 / Get visible bounds info */
        boundsInfo: function(item) {
            var b = item.visibleBounds; // [l,t,r,b]
            var w = b[2] - b[0];
            var h = b[1] - b[3];
            return {
                left: b[0],
                top: b[1],
                right: b[2],
                bottom: b[3],
                width: w,
                height: h,
                area: Math.abs(w * h)
            };
        },
        /* ルーラー単位マップ / Ruler unit map */
        isRectanglePath: function(p) {
            if (p.typename !== 'PathItem' || !p.closed || p.pathPoints.length !== 4) return false;
            for (var i = 0; i < 4; i++) {
                if (p.pathPoints[i].pointType !== PointType.CORNER) return false;
            }
            return true;
        },
        /* 楕円っぽいパスかどうか判定 / Check if path is ellipse-like */
        isEllipseLike: function(p) {
            if (p.typename !== 'PathItem' || !p.closed || p.pathPoints.length !== 4) return false;
            for (var i = 0; i < 4; i++) {
                if (p.pathPoints[i].pointType !== PointType.SMOOTH) return false;
            }
            return true;
        }
    };

    function main() {
        var doc = app.documents.length ? app.activeDocument : null;
        if (!doc) {
            alert('ドキュメントが開いていません / No active document.');
            return;
        }

        if (!doc.selection || doc.selection.length !== 2) {
            alert('ちょうど2つを選択してください。/ Please select exactly two objects.');
            return;
        }

        /* 容器とタイルを面積で判定 / Detect container and tile by area */
        var a = doc.selection[0];
        var b = doc.selection[1];
        var infoA = TG.boundsInfo(a);
        var infoB = TG.boundsInfo(b);
        var container = infoA.area >= infoB.area ? a : b;
        var tile = (container === a) ? b : a;
        var cInfo = TG.boundsInfo(container);
        var tInfo = TG.boundsInfo(tile);

        if (tInfo.width <= 0 || tInfo.height <= 0) {
            alert('タイルのサイズが不正です / Invalid tile size.');
            return;
        }

        // --- プレビュー用変数 / Preview tracking ---
        var _previewGroup = null;

        /* プレビューを削除 / Clear preview */
        function clearPreview() {
            try {
                if (_previewGroup && _previewGroup.isValid) {
                    _previewGroup.remove();
                }
            } catch (e1) {}
            _previewGroup = null;

            // 名前で余分なプレビューを掃除 / Sweep stray previews
            try {
                var lay = container.layer;
                for (var i = lay.groupItems.length - 1; i >= 0; i--) {
                    var gi = lay.groupItems[i];
                    if (gi.name === 'TiledGrid_preview' || gi.name.indexOf('TiledGrid_preview') === 0) {
                        try {
                            gi.remove();
                        } catch (e2) {}
                    }
                }
            } catch (e3) {}
        }

        /* プレビューを描画 / Render preview */
        function renderPreview(gapXpt, gapYpt, marginPt, brick) {
            marginPt = marginPt || 0;
            brick = !!brick;
            clearPreview();

            var tileW = tInfo.width;
            var tileH = tInfo.height;
            var stepXp = tileW + gapXpt;
            var stepYp = tileH + gapYpt;
            if (stepXp <= 0 || stepYp <= 0) return;
            var colsPrev = Math.max(1, Math.ceil((cInfo.width + (brick ? stepXp * 0.5 : 0)) / stepXp));
            var rowsPrev = Math.max(1, Math.ceil(cInfo.height / stepYp));

            var targetLayer = container.layer;
            var group = targetLayer.groupItems.add();
            group.name = 'TiledGrid_preview';
            _previewGroup = group;

            var originLeft = cInfo.left;
            var originTop = cInfo.top;

            var adjLeft = cInfo.left + marginPt;
            var adjRight = cInfo.right - marginPt;
            var adjTop = cInfo.top - marginPt;
            var adjBottom = cInfo.bottom + marginPt;

            var containerIsRect = (container.typename === 'PathItem') && TG.isRectanglePath(container);
            var containerIsEllipse = (container.typename === 'PathItem') && TG.isEllipseLike(container);

            var cxC = (cInfo.left + cInfo.right) / 2.0;
            var cyC = (cInfo.top + cInfo.bottom) / 2.0;
            var rxC = Math.abs(cInfo.right - cInfo.left) / 2.0;
            var ryC = Math.abs(cInfo.top - cInfo.bottom) / 2.0;
            var rxAdj = Math.max(0, rxC - marginPt);
            var ryAdj = Math.max(0, ryC - marginPt);

            function pointInEllipse(x, y) {
                if (rxAdj === 0 || ryAdj === 0) return false;
                var dx = (x - cxC) / rxAdj;
                var dy = (y - cyC) / ryAdj;
                return (dx * dx + dy * dy) <= 1.000001;
            }

            for (var r = 0; r < rowsPrev; r++) {
                var yTop = originTop - r * stepYp;
                for (var c = 0; c < colsPrev; c++) {
                    var xLeft = originLeft + c * stepXp + ((brick && (r % 2 === 1)) ? stepXp * 0.5 : 0);
                    var dup = tile.duplicate(group, ElementPlacement.PLACEATBEGINNING);
                    dup.position = [xLeft, yTop];
                }
            }

            // プレビューもトリム / Trim preview too
            for (var i = group.pageItems.length - 1; i >= 0; i--) {
                var it = group.pageItems[i];
                var ib = TG.boundsInfo(it);
                var keep = false;
                if (containerIsRect) {
                    keep = (ib.left >= adjLeft && ib.right <= adjRight && ib.top <= adjTop && ib.bottom >= adjBottom);
                } else if (containerIsEllipse) {
                    var corners = [
                        [ib.left, ib.top],
                        [ib.right, ib.top],
                        [ib.left, ib.bottom],
                        [ib.right, ib.bottom]
                    ];
                    keep = true;
                    for (var k = 0; k < 4; k++) {
                        if (!pointInEllipse(corners[k][0], corners[k][1])) {
                            keep = false;
                            break;
                        }
                    }
                } else {
                    var cxTile = (ib.left + ib.right) / 2.0;
                    var cyTile = (ib.top + ib.bottom) / 2.0;
                    keep = (cxTile >= adjLeft && cxTile <= adjRight && cyTile <= adjTop && cyTile >= adjBottom);
                }
                if (!keep) it.remove();
            }

            try {
                group.opacity = 60;
            } catch (e) {}
            app.redraw();
        }

        // --- ダイアログの値 / Dialog values ---
        var gapX = 0,
            gapY = 0,
            marginVal = 0,
            brickMode = false,
            useSymbolDup = false;
        var isCancelled = false;

        /* ダイアログボックス生成 / Create Dialog Box */
        (function createSpacingDialog() {
            /* ダイアログのタイトルはラベル＋バージョン / Dialog title = label + version */
            var dlg = new Window('dialog', getLabel('dialogTitle') + ' ' + SCRIPT_VERSION);

            // ウィンドウ見た目 / Window appearance
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
            dlg.alignChildren = 'fill';

            var rulerUnit = getUnitInfo("rulerType");
            var unitLabel = rulerUnit.label;
            var labelWidth = 60;

            /* 間隔 / Spacing */
            var rowS = dlg.add('group');
            rowS.alignment = ['fill', 'top'];
            rowS.alignChildren = ['left', 'center'];
            var lblS = rowS.add('statictext', undefined, labelText('spacing'));
            lblS.justify = 'right';
            lblS.preferredSize.width = labelWidth;
            var tileWidthInUnits = tInfo.width / rulerUnit.pointsPerUnit;
            var defaultGapVal = Math.round(tileWidthInUnits * 0.2);
            var gapEdit = rowS.add('edittext', undefined, String(defaultGapVal));
            gapEdit.characters = 4;
            gapEdit.helpTip = getLabel('tipSpacing');
            rowS.add('statictext', undefined, unitSuffix(unitLabel));
            gapEdit.active = true;

            /* キー操作で値を変更するヘルパー / Helper to change value by arrow keys */
            function changeValueByArrowKey(editText, allowNegative) {
                allowNegative = !!allowNegative;
                editText.addEventListener("keydown", function(event) {
                    var value = Number(editText.text);
                    if (isNaN(value)) return;
                    var keyboard = ScriptUI.environment.keyboardState;
                    var delta = 1;

                    if (keyboard.shiftKey) {
                        delta = 10;
                        if (event.keyName == "Up") {
                            value = Math.ceil((value + 1) / delta) * delta;
                            event.preventDefault();
                        } else if (event.keyName == "Down") {
                            value = Math.floor((value - 1) / delta) * delta;
                            if (!allowNegative && value < 0) value = 0;
                            event.preventDefault();
                        }
                    } else if (keyboard.altKey) {
                        delta = 0.1;
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
                            if (!allowNegative && value < 0) value = 0;
                            event.preventDefault();
                        }
                    }

                    if (keyboard.altKey) {
                        value = Math.round(value * 10) / 10;
                    } else {
                        value = Math.round(value);
                    }

                    editText.text = value;
                    updatePreviewFromFields();
                });
            }
            changeValueByArrowKey(gapEdit);

            /* マージン / Margin */
            var rowM = dlg.add('group');
            rowM.alignment = ['fill', 'top'];
            rowM.alignChildren = ['left', 'center'];
            var lblM = rowM.add('statictext', undefined, labelText('margin'));
            lblM.justify = 'right';
            lblM.preferredSize.width = labelWidth;
            var marginEdit = rowM.add('edittext', undefined, '0');
            marginEdit.characters = 4;
            marginEdit.helpTip = getLabel('tipMargin');
            rowM.add('statictext', undefined, unitSuffix(unitLabel));
            changeValueByArrowKey(marginEdit, true); // マージンは負OK / margin can be negative

            /* プレビュー更新 / Update preview */
            function updatePreviewFromFields() {
                var pointsPerUnit = getUnitInfo("rulerType").pointsPerUnit;
                var s = Math.max(0, parseFloat(gapEdit.text) || 0);
                var m = parseFloat(marginEdit.text);
                if (isNaN(m)) m = 0;
                s *= pointsPerUnit;
                m *= pointsPerUnit;
                var b = false;
                try { b = !!(brickChk && brickChk.value); } catch (e) {}
                renderPreview(s, s, m, b);
            }
            gapEdit.onChanging = function() {
                updatePreviewFromFields();
            };
            updatePreviewFromFields();

            /* レンガ状 / Brick pattern */
            var rowBWrap = dlg.add('group');
            rowBWrap.alignment = ['fill', 'top'];
            rowBWrap.alignChildren = ['center', 'center'];

            var rowB = rowBWrap.add('group');
            rowB.alignChildren = ['left', 'center'];
            var brickChk = rowB.add('checkbox', undefined, getLabel('brick'));
            brickChk.helpTip = getLabel('tipBrick');
            brickChk.value = false;
            brickChk.onClick = function() {
                updatePreviewFromFields();
            };

            /* シンボル化して複製 / Duplicate as Symbol */
            var rowSym = rowBWrap.add('group');
            rowSym.alignChildren = ['left', 'center'];
            var symChk = rowSym.add('checkbox', undefined, getLabel('symbolize'));
            symChk.helpTip = getLabel('tipSymbolize');
            symChk.value = false;

            /* ボタン / Buttons */
            var btns = dlg.add('group');
            btns.alignment = 'right';
            btns.add('button', undefined, getLabel('cancel'), { name: 'cancel' });
            btns.add('button', undefined, getLabel('ok'), { name: 'ok' });

            var dialogResult = dlg.show();
            if (dialogResult !== 1) {
                clearPreview();
                isCancelled = true;
                return;
            }

            // --- OKで確定値をptに変換 / Convert to pt on OK ---
            var pointsPerUnit = getUnitInfo("rulerType").pointsPerUnit;
            var sVal = Math.max(0, parseFloat(gapEdit.text) || 0);
            brickMode = !!(brickChk && brickChk.value);
            useSymbolDup = !!(symChk && symChk.value);

            marginVal = parseFloat(marginEdit.text);
            if (isNaN(marginVal)) marginVal = 0;
            sVal *= pointsPerUnit;
            gapX = sVal;
            gapY = sVal;
            marginVal *= pointsPerUnit;

        })();
        if (isCancelled) {
            app.userInteractionLevel = UserInteractionLevel.DISPLAYALERTS;
            return;
        }

        clearPreview();

        // --- 必要ならシンボル作成 / Symbolize if needed ---
        var symDef = null;
        if (useSymbolDup) {
            try {
                var tmpForSym = tile.duplicate();
                symDef = doc.symbols.add(tmpForSym);
                try { tmpForSym.remove(); } catch (_eTmp) {}
            } catch (_eSym) {
                symDef = null; // 失敗時は通常複製 / fallback
            }
        }

        // --- 配置数を計算 / Compute counts ---
        var tileWFinal = tInfo.width;
        var tileHFinal = tInfo.height;
        var stepX = tileWFinal + gapX;
        var stepY = tileHFinal + gapY;
        var cols = Math.max(1, Math.ceil((cInfo.width + (brickMode ? stepX * 0.5 : 0)) / stepX));
        var rows = Math.max(1, Math.ceil(cInfo.height / stepY));

        app.userInteractionLevel = UserInteractionLevel.DONTDISPLAYALERTS;
        app.redraw();

        /* 複製グループを作成 / Create group for final duplicates */
        var targetLayer = container.layer;
        var gridGroup = targetLayer.groupItems.add();
        gridGroup.name = 'TiledGrid';

        var originLeft = cInfo.left;
        var originTop = cInfo.top;

        // 元タイルの位置を記録して上書き重複を防ぐ / Remember original tile position
        var origTileLeft = tInfo.left;
        var origTileTop  = tInfo.top;

        /* 敷き詰め実行 / Duplicate and place in grid */
        for (var r = 0; r < rows; r++) {
            var yTop = originTop - r * stepY;
            for (var c = 0; c < cols; c++) {
                var xLeft = originLeft + c * stepX + ((brickMode && (r % 2 === 1)) ? stepX * 0.5 : 0);
                // 同じ場所に元タイルがある場合はスキップ / Skip if same as original tile
                if (Math.abs(xLeft - origTileLeft) < 0.01 && Math.abs(yTop - origTileTop) < 0.01) {
                    continue;
                }
                if (useSymbolDup && symDef) {
                    var si = doc.symbolItems.add(symDef);
                    si.move(gridGroup, ElementPlacement.PLACEATBEGINNING);
                    si.position = [xLeft, yTop];
                } else {
                    var dup = tile.duplicate(gridGroup, ElementPlacement.PLACEATBEGINNING);
                    dup.position = [xLeft, yTop];
                }
            }
        }

        // --- マスクなしトリム / Trim without mask ---
        var containerIsRect = (container.typename === 'PathItem') && TG.isRectanglePath(container);
        var containerIsEllipse = (container.typename === 'PathItem') && TG.isEllipseLike(container);

        var cx = (cInfo.left + cInfo.right) / 2.0;
        var cy = (cInfo.top + cInfo.bottom) / 2.0;
        var rx = Math.abs(cInfo.right - cInfo.left) / 2.0;
        var ry = Math.abs(cInfo.top - cInfo.bottom) / 2.0;
        var adjLeftF = cInfo.left + marginVal;
        var adjRightF = cInfo.right - marginVal;
        var adjTopF = cInfo.top - marginVal;
        var adjBottomF = cInfo.bottom + marginVal;
        var rxAdjF = Math.max(0, rx - marginVal);
        var ryAdjF = Math.max(0, ry - marginVal);

        function pointInEllipse(x, y) {
            if (rx === 0 || ry === 0) return false;
            var dx = (x - cx) / rx;
            var dy = (y - cy) / ry;
            return (dx * dx + dy * dy) <= 1.000001;
        }

        for (var i = gridGroup.pageItems.length - 1; i >= 0; i--) {
            var it = gridGroup.pageItems[i];
            var ib = TG.boundsInfo(it);
            var keep = false;

            if (containerIsRect) {
                /* 矩形：バウンディングが内側に完全に入っているか / Rect: full bbox inside */
                keep = (ib.left >= adjLeftF && ib.right <= adjRightF && ib.top <= adjTopF && ib.bottom >= adjBottomF);
            } else if (containerIsEllipse) {
                /* 楕円：4隅が楕円の内側にあるか / Ellipse: 4 corners inside */
                var corners = [
                    [ib.left, ib.top],
                    [ib.right, ib.top],
                    [ib.left, ib.bottom],
                    [ib.right, ib.bottom]
                ];
                keep = (rxAdjF > 0 && ryAdjF > 0);
                if (keep) {
                    for (var k = 0; k < 4; k++) {
                        var dx = (corners[k][0] - cx) / rxAdjF;
                        var dy = (corners[k][1] - cy) / ryAdjF;
                        if ((dx * dx + dy * dy) > 1.000001) {
                            keep = false;
                            break;
                        }
                    }
                }
            } else {
                /* その他：中心点が内側にあるか / Others: center point inside */
                var cxTile = (ib.left + ib.right) / 2.0;
                var cyTile = (ib.top + ib.bottom) / 2.0;
                keep = (cxTile >= adjLeftF && cxTile <= adjRightF && cyTile <= adjTopF && cyTile >= adjBottomF);
            }

            if (!keep) it.remove();
        }

        app.userInteractionLevel = UserInteractionLevel.DISPLAYALERTS;
        app.redraw();
    }

    /* メインを1アクションで実行 / Run main in single undo */
    function __runMain() {
        try {
            main();
        } catch (e) {
            try {
                alert('[TileSmallIntoLarge] Error:\n' + e);
            } catch (e) {}
        }
    }
    try {
        if (typeof app.doScript === 'function' && typeof ScriptLanguage !== 'undefined' && typeof UndoModes !== 'undefined') {
            app.doScript(__runMain, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, 'TileSmallIntoLarge');
        } else {
            __runMain();
        }
    } catch (e) {
        __runMain();
    }

})();
