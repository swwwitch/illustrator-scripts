#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
Script: TileSmallIntoLarge.jsx
Purpose (JP):
  2つのオブジェクトを選択し、面積が大きい方を「容器」、小さい方を「タイル」とみなし、
  容器のバウンディングボックスをタイルで均等に敷き詰めます（グリッド配置）。
  結果はマスクなしで、容器の内側に収まるタイルのみ残します。元の2オブジェクトは残します。
Purpose (EN):
  With two selected objects, treats the larger one as the container and the smaller one as the tile.
  Fills the container's bounding box with an even grid of tile duplicates, then keeps only tiles fully inside the container (no mask).
  The two original items remain untouched.

更新日 / Updated: 2025-10-30

更新履歴 / Changelog
- 2025-10-30 v1.4.0: ローカライズ処理を整理し、ダイアログタイトルにバージョンを明示する形式に変更。コメントを日英併記に統一。/ Refined localization, dialog title now shows version explicitly, unified JP/EN comments.
- 2025-10-29 v1.3.0: 「シンボル化して複製」オプションを追加（プレビューは従来どおり） / Added "Duplicate as Symbol" option (preview still uses direct duplicates).
- 2025-10-29 v1.2.0: UIラベル「レンガ」→「レンガ状」に変更（機能は不変） / Changed UI label from "レンガ" to "レンガ状" (no functional change).
- 2025-10-29 v1.1.0: スケールUIとロジックを削除 / Removed Scale UI and logic.
- 2025-10-26 v1.0.0: 初版 / Initial release.
*/

var SCRIPT_VERSION = "v1.4.0"; // バージョン / Script version

/* 現在のロケールを判定 / Detect current locale */
function getCurrentLang() {
  return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

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
    ja: "Cancel",
    en: "Cancel"
  }
};

/* ラベル取得関数 / Label resolver */
function L(key) {
  try {
    return LABELS[key][lang] || LABELS[key].en;
  } catch (e) {
    return key;
  }
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
    unitLabelMap: {
        0: 'in',
        1: 'mm',
        2: 'pt',
        3: 'pica',
        4: 'cm',
        5: 'Q/H',
        6: 'px',
        7: 'ft/in',
        8: 'm',
        9: 'yd',
        10: 'ft'
    },
    /* 現在のルーラー単位コードを取得 / Get current ruler type */
    getRulerType: function() {
        return app.preferences.getIntegerPreference('rulerType');
    },
    /* 現在の単位のラベルを取得 / Get current unit label */
    getCurrentUnitLabel: function() {
        var unitCode = TG.getRulerType();
        return TG.unitLabelMap[unitCode] || 'pt';
    },
    /* 任意の単位をptに変換 / Convert units to points */
    toPoints: function(val, unitCode) {
        if (val === 0) return 0;
        switch (unitCode) {
            case 0:
                return val * 72.0; // in
            case 1:
                return val * (72.0 / 25.4); // mm
            case 2:
                return val; // pt
            case 3:
                return val * 12.0; // pica
            case 4:
                return val * (72.0 / 2.54); // cm
            case 5:
                return val * (72.0 / 25.4) * 0.25; // Q/H (1Q=0.25mm)
            case 6:
                return val * 1.0; // px ≈ pt (Illustrator geometry)
            case 7:
                return val * 72.0; // ft/in -> treat as inch
            case 8:
                return val * (72.0 / 0.0254); // m
            case 9:
                return val * (72.0 * 36.0); // yd
            case 10:
                return val * (72.0 * 12.0); // ft
            default:
                return val;
        }
    },
    /* ptを現在単位に変換 / Convert points to current units */
    pointsToUnits: function(valPt, unitCode) {
        var perUnit = TG.toPoints(1, unitCode);
        if (!perUnit || perUnit === 0) return valPt; // fallback
        return valPt / perUnit;
    },
    /* 長方形パスかどうか判定 / Check if path is rectangle */
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
    var _cancel = false;

    /* ダイアログボックス生成 / Create Dialog Box */
    (function createSpacingDialog() {
        /* ダイアログのタイトルはラベル＋バージョン / Dialog title = label + version */
        var dlg = new Window('dialog', L('dialogTitle') + ' ' + SCRIPT_VERSION);

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

        var unitLabel = TG.getCurrentUnitLabel();
        var labelWidth = 60;

        /* 間隔 / Spacing */
        var rowS = dlg.add('group');
        rowS.alignment = ['fill', 'top'];
        rowS.alignChildren = ['left', 'center'];
        var lblS = rowS.add('statictext', undefined, L('spacing'));
        try { lblS.justify = 'right'; } catch (e) {}
        lblS.preferredSize.width = labelWidth;
        var unitCodeNowForDefault = TG.getRulerType();
        var tileWidthInUnits = TG.pointsToUnits(tInfo.width, unitCodeNowForDefault);
        var defaultGapVal = Math.round(tileWidthInUnits * 0.2);
        var gapEdit = rowS.add('edittext', undefined, String(defaultGapVal));
        gapEdit.characters = 4;
        rowS.add('statictext', undefined, '（' + unitLabel + '）');
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
        var lblM = rowM.add('statictext', undefined, L('margin'));
        try { lblM.justify = 'right'; } catch (e) {}
        lblM.preferredSize.width = labelWidth;
        var marginEdit = rowM.add('edittext', undefined, '0');
        marginEdit.characters = 4;
        rowM.add('statictext', undefined, '（' + unitLabel + '）');
        changeValueByArrowKey(marginEdit, true); // マージンは負OK / margin can be negative

        /* プレビュー更新 / Update preview */
        function updatePreviewFromFields() {
            var unitCodeNow = TG.getRulerType();
            var s = Math.max(0, parseFloat(gapEdit.text) || 0);
            var m = parseFloat(marginEdit.text);
            if (isNaN(m)) m = 0;
            s = TG.toPoints(s, unitCodeNow);
            m = TG.toPoints(m, unitCodeNow);
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
        var brickChk = rowB.add('checkbox', undefined, L('brick'));
        brickChk.value = false;
        brickChk.onClick = function() {
            updatePreviewFromFields();
        };

        /* シンボル化して複製 / Duplicate as Symbol */
        var rowSym = rowBWrap.add('group');
        rowSym.alignChildren = ['left', 'center'];
        var symChk = rowSym.add('checkbox', undefined, L('symbolize'));
        symChk.value = false;

        /* ボタン / Buttons */
        var btns = dlg.add('group');
        btns.alignment = 'right';
        btns.add('button', undefined, L('cancel'), { name: 'cancel' });
        btns.add('button', undefined, L('ok'), { name: 'ok' });

        var res = dlg.show();
        if (res !== 1) {
            clearPreview();
            _cancel = true;
            return;
        }

        // --- OKで確定値をptに変換 / Convert to pt on OK ---
        var unitCode = TG.getRulerType();
        var sVal = Math.max(0, parseFloat(gapEdit.text) || 0);
        brickMode = !!(brickChk && brickChk.value);
        useSymbolDup = !!(symChk && symChk.value);

        marginVal = parseFloat(marginEdit.text);
        if (isNaN(marginVal)) marginVal = 0;
        sVal = TG.toPoints(sVal, unitCode);
        gapX = sVal;
        gapY = sVal;
        marginVal = TG.toPoints(marginVal, unitCode);

    })();
    if (_cancel) {
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
        } catch (_) {}
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