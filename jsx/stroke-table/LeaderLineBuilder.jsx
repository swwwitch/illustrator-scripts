#target illustrator
#targetengine "AngledLeaderLineMakerEngine"

/*
 * 角度指定の引き出し線メーカー.jsx
 * v1.4.2
 * 更新日: 20260310
 *
 * 選択したパスまたはグループから、指定角度の引き出し線を生成するスクリプトです。
 * プレビュー表示、線端の丸、白フチ、線色、線幅、グループ化に対応しています。
 * v1.4.2では bounds 情報の都度再取得、プレビュー／確定の生成構造の共通化、非グループ時の選択状態の整理、生成先レイヤーの安定化を行いました。
 *
 * ■生成構造
 * 白フチONかつグループ化時は次の構造で生成します。
 * 親グループ（note="leader_line"）
 *   ├─ 本体サブグループ
 *   │    ├─ A：引き出し線
 *   │    └─ B：線端の丸
 *   └─ フチサブグループ
 *        ├─ C：フチ線
 *        └─ D：フチの丸
 *
 * ■再適用
 * 生成された leader_line グループに再実行すると、
 * 生成時に保存した元の bounds を基準に再計算します。
 * これにより再適用を繰り返しても長さが変化しません。
 * 元オブジェクトは削除せず、非表示の退避レイヤーへ移動します。
 * これにより生成途中の失敗や再適用対象の想定ズレがあっても、
 * 元オブジェクトを破壊しにくくしています。
 *
 * 古い leader_line データの場合は、A線と線端丸から
 * 元の長さを推定して再計算します。
 *
 * ■色処理
 * 線色と白フチ色は、アクティブドキュメントの
 * カラーモード（RGB / CMYK）に合わせて生成します。
 *
 * ■単位
 * 線幅と線端サイズの表示は Illustrator の
 * strokeUnits 設定に従い、内部計算は pt で行います。
 *
 * Supports creating angled leader lines from selected paths or groups
 * with preview, end circles, white edge, line color, line width, and grouping.
 * Generated groups are tagged with note="leader_line" and store the
 * original bounds so reapplying the script preserves the original length.
 */
var SCRIPT_VERSION = "v1.4.2";

// セッション中の設定を記憶
if (typeof $.global._leaderLineSettings === "undefined") {
    $.global._leaderLineSettings = {
        angle: "45",
        radioAngle: 45,
        applyScope: "all",
        diagDir: "upperLeft",
        hDir: "left",
        vDir: "up",
        capType: "circle",
        capStyle: "fill",
        capSize: "3",
        strokeCapType: "none",
        groupEnabled: true,
        whiteEdge: false,
        preview: false,
        dialogBounds: null, // ダイアログ位置 [x, y] を保存
        lineColor: "black",
        lineColorHex: "#ffcc00",
        lineWidth: "1"
    };
}

// 現在のロケールを取得 / Get current locale
function getCurrentLang() {
    var locale = $.locale || '';
    if (locale.indexOf("ja") === 0) return "ja";
    return "en";
}
var lang = getCurrentLang();

/* 日英ラベル定義 / Japanese-English label definitions */
var LABELS = {
    dialogTitle: {
        ja: "角度指定の引き出し線メーカー",
        en: "Angled Leader Line Maker"
    },
    alertNoDocument: {
        ja: "ドキュメントが開かれていません。",
        en: "No document is open."
    },
    alertNoSelection: {
        ja: "パスまたはグループを選択して実行してください。",
        en: "Please select a path or group and run the script."
    },
    alertNoValidTargets: {
        ja: "パス（2点以上）またはグループが選択されていません。",
        en: "No valid path (2 or more points) or group is selected."
    },
    alertInvalidAngle: {
        ja: "角度は0より大きく90未満の値を入力してください。",
        en: "Enter an angle greater than 0 and less than 90."
    },
    panelAngle: {
        ja: "角度",
        en: "Angle"
    },
    panelDirection: {
        ja: "斜線の方向",
        en: "Diagonal Direction"
    },
    panelApplyScope: {
        ja: "適用範囲",
        en: "Apply Scope"
    },
    scopeAll: {
        ja: "すべて",
        en: "All"
    },
    scopeExceptDirection: {
        ja: "斜線の方向以外",
        en: "Except Direction"
    },
    panelLineStyle: {
        ja: "線のスタイル",
        en: "Line Style"
    },
    panelLineEnd: {
        ja: "線端",
        en: "Line End"
    },
    capShapeLabel: {
        ja: "形状",
        en: "Shape"
    },
    capStrokeCapLabel: {
        ja: "線端",
        en: "Stroke Cap"
    },
    capStrokeCapNone: {
        ja: "なし",
        en: "None"
    },
    panelOptions: {
        ja: "オプション",
        en: "Options"
    },
    dirUpperLeft: {
        ja: "左上",
        en: "Upper Left"
    },
    dirLowerLeft: {
        ja: "左下",
        en: "Lower Left"
    },
    dirUpperRight: {
        ja: "右上",
        en: "Upper Right"
    },
    dirLowerRight: {
        ja: "右下",
        en: "Lower Right"
    },
    colorBlack: {
        ja: "黒",
        en: "Black"
    },
    colorWhite: {
        ja: "白",
        en: "White"
    },
    colorOther: {
        ja: "その他",
        en: "Other"
    },
    lineWidth: {
        ja: "線幅",
        en: "Line Width"
    },
    capNone: {
        ja: "なし",
        en: "None"
    },
    capRound: {
        ja: "丸型",
        en: "Round"
    },
    capCircle: {
        ja: "円",
        en: "Circle"
    },
    capArrow: {
        ja: "矢印",
        en: "Arrow"
    },
    capFill: {
        ja: "塗り",
        en: "Fill"
    },
    capStroke: {
        ja: "線",
        en: "Stroke"
    },
    capSize: {
        ja: "大きさ",
        en: "Size"
    },
    groupEnabled: {
        ja: "グループ化",
        en: "Group"
    },
    whiteEdge: {
        ja: "フチ",
        en: "White Edge"
    },
    preview: {
        ja: "プレビュー",
        en: "Preview"
    },
    cancel: {
        ja: "キャンセル",
        en: "Cancel"
    },
    ok: {
        ja: "OK",
        en: "OK"
    }
};

/* ラベル取得 / Get localized label */
function L(key) {
    var item = LABELS[key];
    if (!item) return key;
    return item[lang] || item.en || key;
}

/* 単位ユーティリティ / Unit utilities */
var unitMap = {
    0: "in",
    1: "mm",
    2: "pt",
    3: "pica",
    4: "cm",
    6: "px",
    7: "ft/in",
    8: "m",
    9: "yd",
    10: "ft"
};

/* 単位コードと設定キーから適切な単位ラベルを返す / Get unit label from unit code and preference key */
function getUnitLabel(code, prefKey) {
    if (code === 5) {
        var hKeys = {
            "text/asianunits": true,
            "rulerType": true,
            "strokeUnits": true
        };
        return hKeys[prefKey] ? "H" : "Q";
    }
    return unitMap[code] || "pt";
}

/* 設定キーから現在の単位ラベルを取得 / Get current unit label from preference key */
function getCurrentUnitLabel(prefKey) {
    var unitCode = app.preferences.getIntegerPreference(prefKey);
    return getUnitLabel(unitCode, prefKey);
}

/* 単位コードからpt換算係数を取得 / Get pt conversion factor from unit code */
function getUnitToPtFactor(code, prefKey) {
    switch (code) {
        case 0: return 72;          // in
        case 1: return 72 / 25.4;   // mm
        case 2: return 1;           // pt
        case 3: return 12;          // pica
        case 4: return 72 / 2.54;   // cm
        case 5: return 0.25;        // Q/H
        case 6: return 1;           // px
        case 7: return 72;          // ft/in（数値入力はinch扱い）
        case 8: return 72 / 0.0254; // m
        case 9: return 72 * 36;     // yd
        case 10: return 72 * 12;    // ft
        default: return 1;
    }
}

/* 設定キーから現在のpt換算係数を取得 / Get current pt conversion factor from preference key */
function getCurrentUnitToPtFactor(prefKey) {
    var unitCode = app.preferences.getIntegerPreference(prefKey);
    return getUnitToPtFactor(unitCode, prefKey);
}

/* 単位値をptに変換 / Convert unit value to points */
function unitValueToPt(value, prefKey) {
    var n = parseFloat(value);
    if (isNaN(n)) return NaN;
    return n * getCurrentUnitToPtFactor(prefKey);
}

/* ptを単位値に変換 / Convert points to unit value */
function ptValueToUnit(valuePt, prefKey) {
    var n = parseFloat(valuePt);
    if (isNaN(n)) return NaN;
    return n / getCurrentUnitToPtFactor(prefKey);
}

function roundTo(value, digits) {
    var p = Math.pow(10, digits || 0);
    return Math.round(value * p) / p;
}

function formatNumber(value, digits) {
    if (isNaN(value)) return "";
    var s = String(roundTo(value, digits || 0));
    s = s.replace(/\.0+$/, "");
    s = s.replace(/(\.\d*?)0+$/, "$1");
    return s;
}

function formatUnitValue(valuePt, prefKey) {
    return formatNumber(ptValueToUnit(valuePt, prefKey), 2);
}

function main() {
    // ドキュメントが開かれているか確認
    if (app.documents.length === 0) {
        alert(L("alertNoDocument"));
        return;
    }

    var doc = app.activeDocument;
    var sel = doc.selection;

    // 選択オブジェクトのチェック
    if (sel.length === 0) {
        alert(L("alertNoSelection"));
        return;
    }

    // 対象を収集（PathItem：2点以上、GroupItem：通常は全体参照、leader_line 再適用時は A線優先）
    var targets = [];
    for (var i = 0; i < sel.length; i++) {
        var t = sel[i].typename;
        if (t === "PathItem" && sel[i].pathPoints.length >= 2) {
            targets.push(sel[i]);
        } else if (t === "GroupItem") {
            targets.push(sel[i]);
        }
    }

    if (targets.length === 0) {
        alert(L("alertNoValidTargets"));
        return;
    }

    // GroupItem の参照用 PathItem を取得する関数
    // 通常は closed path を優先して返す（leader_line 再適用時の A線優先ロジックは getLeaderLineBasePath() 側で処理）
    function getGroupReferencePathItems(groupItem) {
        var closedItems = [];
        var allPathItems = [];
        for (var k = 0; k < groupItem.pathItems.length; k++) {
            var pi = groupItem.pathItems[k];
            allPathItems.push(pi);
            if (pi.closed) closedItems.push(pi);
        }
        return closedItems.length ? closedItems : allPathItems;
    }

    // PathItem のアンカー座標から外接矩形を取得する関数
    function getPathAnchorBounds(pathItem) {
        var pts = pathItem.pathPoints;
        if (!pts || pts.length === 0) {
            return {
                x_left: 0,
                x_right: 0,
                y_top: 0,
                y_bottom: 0
            };
        }

        var xMin = pts[0].anchor[0], xMax = pts[0].anchor[0];
        var yMin = pts[0].anchor[1], yMax = pts[0].anchor[1];
        for (var i = 1; i < pts.length; i++) {
            var ax = pts[i].anchor[0];
            var ay = pts[i].anchor[1];
            if (ax < xMin) xMin = ax;
            if (ax > xMax) xMax = ax;
            if (ay < yMin) yMin = ay;
            if (ay > yMax) yMax = ay;
        }

        return {
            x_left: xMin,
            x_right: xMax,
            y_top: yMax,
            y_bottom: yMin
        };
    }

    // GroupItem の bounds を取得する関数 / 通常はグループ全体、leader_line 再適用時は保存済み bounds を優先し、なければ旧構造から復元する
    function getGroupBounds(groupItem) {
        if (isLeaderLineGroup(groupItem)) {
            var storedBounds = getLeaderLineStoredBounds(groupItem);
            if (storedBounds) {
                return storedBounds;
            }

            var basePath = getLeaderLineBasePath(groupItem);
            if (basePath) {
                var mainCap = getLeaderLineMainCap(groupItem);
                return getPathAnchorBoundsWithLeaderCap(basePath, mainCap);
            }
        }

        var gb = groupItem.geometricBounds; // [left, top, right, bottom]
        return {
            x_left: gb[0],
            x_right: gb[2],
            y_top: gb[1],
            y_bottom: gb[3]
        };
    }
    function getLeaderLineMainCap(groupItem) {
        for (var i = 0; i < groupItem.pathItems.length; i++) {
            var pi = groupItem.pathItems[i];
            if (pi.note === "leader_line_main_cap") return pi;
        }
        return null;
    }

    function getCapCenter(capItem) {
        if (!capItem) return null;
        var capGB = capItem.geometricBounds; // [left, top, right, bottom]
        return [
            (capGB[0] + capGB[2]) / 2,
            (capGB[1] + capGB[3]) / 2
        ];
    }

    function getPathAnchorBoundsWithLeaderCap(pathItem, capItem) {
        var bounds = getPathAnchorBounds(pathItem);
        if (!capItem) return bounds;

        var pts = pathItem.pathPoints;
        if (!pts || pts.length < 2) return bounds;

        var p0 = pts[0].anchor;
        var p1 = pts[1].anchor;
        var p2 = pts[pts.length - 1].anchor;
        var capCenter = getCapCenter(capItem);
        if (!capCenter) return bounds;

        // 線端に丸がある場合、丸の中心（= 丸の半分位置）を期待する端として扱う
        var cx = capCenter[0];
        var cy = capCenter[1];

        var d0 = Math.abs(p0[0] - cx) + Math.abs(p0[1] - cy);
        var d2 = Math.abs(p2[0] - cx) + Math.abs(p2[1] - cy);

        var tip0 = [p0[0], p0[1]];
        var bend = [p1[0], p1[1]];
        var tip2 = [p2[0], p2[1]];

        if (d0 <= d2) {
            tip0 = [cx, cy];
        } else {
            tip2 = [cx, cy];
        }

        return {
            x_left: Math.min(tip0[0], bend[0], tip2[0]),
            x_right: Math.max(tip0[0], bend[0], tip2[0]),
            y_top: Math.max(tip0[1], bend[1], tip2[1]),
            y_bottom: Math.min(tip0[1], bend[1], tip2[1])
        };
    }

    // 線の属性を取得する関数（GroupItem は通常 closed path 優先、leader_line 再適用時は A線優先）
    function getStrokeInfo(item) {
        if (item.typename === "PathItem") {
            return {
                stroked: item.stroked,
                strokeWidth: item.stroked ? item.strokeWidth : 1,
                strokeColor: item.stroked ? item.strokeColor : undefined
            };
        }
        // GroupItemの場合、closed pathを優先して配下のPathItemを探す
        if (isLeaderLineGroup(item)) {
            var leaderBasePath = getLeaderLineBasePath(item);
            if (leaderBasePath) {
                return {
                    stroked: leaderBasePath.stroked,
                    strokeWidth: leaderBasePath.stroked ? leaderBasePath.strokeWidth : 1,
                    strokeColor: leaderBasePath.stroked ? leaderBasePath.strokeColor : undefined
                };
            }
        }

        // GroupItemの場合、closed pathを優先して配下のPathItemを探す
        var refItems = getGroupReferencePathItems(item);
        for (var k = 0; k < refItems.length; k++) {
            var pi = refItems[k];
            if (pi.stroked) {
                return {
                    stroked: true,
                    strokeWidth: pi.strokeWidth,
                    strokeColor: pi.strokeColor
                };
            }
        }
        return { stroked: false, strokeWidth: 1, strokeColor: undefined };
    }

    // 各ターゲットの座標情報を都度取得する関数
    function collectBoundsData(targetItems) {
        var data = [];
        for (var i = 0; i < targetItems.length; i++) {
            var item = targetItems[i];
            var x_left, x_right, y_bottom, y_top;

            if (item.typename === "PathItem") {
                // PathItem：アンカーポイントから外接矩形を算出
                var pathBounds = getPathAnchorBounds(item);
                x_left = pathBounds.x_left;
                x_right = pathBounds.x_right;
                y_bottom = pathBounds.y_bottom;
                y_top = pathBounds.y_top;
            } else {
                // GroupItem：通常はグループ全体、leader_line 再適用時は保存済み bounds を優先して使用
                var groupBounds = getGroupBounds(item);
                x_left = groupBounds.x_left;
                x_right = groupBounds.x_right;
                y_top = groupBounds.y_top;
                y_bottom = groupBounds.y_bottom;
            }

            var si = getStrokeInfo(item);
            data.push({
                target: item,
                x_left: x_left,
                x_right: x_right,
                y_bottom: y_bottom,
                y_top: y_top,
                stroked: si.stroked,
                strokeWidth: si.strokeWidth,
                strokeColor: si.strokeColor
            });
        }
        return data;
    }

    // プレビュー用アイテムの配列（親グループ単位で管理）
    var previewItems = [];

    // HEXカラーをRGBColorに変換する関数
    function hexToRGBColor(hex) {
        hex = hex.replace(/^#/, "");
        if (hex.length === 3) {
            hex = hex.charAt(0) + hex.charAt(0) + hex.charAt(1) + hex.charAt(1) + hex.charAt(2) + hex.charAt(2);
        }
        var r = parseInt(hex.substring(0, 2), 16);
        var g = parseInt(hex.substring(2, 4), 16);
        var b = parseInt(hex.substring(4, 6), 16);
        if (isNaN(r) || isNaN(g) || isNaN(b)) return null;
        var color = new RGBColor();
        color.red = r;
        color.green = g;
        color.blue = b;
        return color;
    }

    function isCMYKDocument() {
        return doc.documentColorSpace === DocumentColorSpace.CMYK;
    }

    function createBlackColor() {
        if (isCMYKDocument()) {
            var black = new CMYKColor();
            black.cyan = 0;
            black.magenta = 0;
            black.yellow = 0;
            black.black = 100;
            return black;
        }
        var rgbBlack = new RGBColor();
        rgbBlack.red = 0;
        rgbBlack.green = 0;
        rgbBlack.blue = 0;
        return rgbBlack;
    }

    function createWhiteColor() {
        if (isCMYKDocument()) {
            var white = new CMYKColor();
            white.cyan = 0;
            white.magenta = 0;
            white.yellow = 0;
            white.black = 0;
            return white;
        }
        var rgbWhite = new RGBColor();
        rgbWhite.red = 255;
        rgbWhite.green = 255;
        rgbWhite.blue = 255;
        return rgbWhite;
    }

    function rgbToCMYKColor(r, g, b) {
        var rr = Math.max(0, Math.min(255, r)) / 255;
        var gg = Math.max(0, Math.min(255, g)) / 255;
        var bb = Math.max(0, Math.min(255, b)) / 255;
        var k = 1 - Math.max(rr, gg, bb);
        var c = 0, m = 0, y = 0;

        if (k < 1) {
            c = (1 - rr - k) / (1 - k);
            m = (1 - gg - k) / (1 - k);
            y = (1 - bb - k) / (1 - k);
        }

        var color = new CMYKColor();
        color.cyan = Math.round(c * 100);
        color.magenta = Math.round(m * 100);
        color.yellow = Math.round(y * 100);
        color.black = Math.round(k * 100);
        return color;
    }

    function hexToDocumentColor(hex) {
        var rgb = hexToRGBColor(hex);
        if (!rgb) return null;
        if (isCMYKDocument()) {
            return rgbToCMYKColor(rgb.red, rgb.green, rgb.blue);
        }
        return rgb;
    }

    // 線のカラーを取得する関数
    function getLineColor(b) {
        if (colorBlack.value) {
            return createBlackColor();
        } else if (colorWhite.value) {
            return createWhiteColor();
        } else {
            // その他：HEXカラーコードから現在のドキュメント色空間に合わせて変換
            var docColor = hexToDocumentColor(hexInput.text);
            if (docColor) return docColor;
            // 無効な値の場合は元のオブジェクトの線色を使用
            if (b.strokeColor) return b.strokeColor;
            return createBlackColor();
        }
    }

    // 線幅を取得する関数
    function getLineWidth(b) {
        var w = unitValueToPt(lineWidthInput.text, "strokeUnits");
        if (!isNaN(w) && w > 0) return w;
        return b.strokeWidth;
    }

    // 引き出し線の座標を計算する関数
    function calcLeaderPoints(b, angleRad, hDir, vDir) {
        var height = b.y_top - b.y_bottom;
        var offset = height / Math.tan(angleRad);
        if (vDir === "down") {
            // 水平線が上、斜線が下に向かう
            if (hDir === "left") {
                var ix = b.x_left + offset;
                return [[b.x_left, b.y_bottom], [ix, b.y_top], [b.x_right, b.y_top]];
            } else {
                var ix = b.x_right - offset;
                return [[b.x_left, b.y_top], [ix, b.y_top], [b.x_right, b.y_bottom]];
            }
        } else {
            // 水平線が下、斜線が上に向かう
            if (hDir === "left") {
                var ix = b.x_left + offset;
                return [[b.x_left, b.y_top], [ix, b.y_bottom], [b.x_right, b.y_bottom]];
            } else {
                var ix = b.x_right - offset;
                return [[b.x_left, b.y_bottom], [ix, b.y_bottom], [b.x_right, b.y_top]];
            }
        }
    }

    // 斜線の先端座標を取得する関数
    function getTipPoint(points, hDir) {
        return (hDir === "left") ? points[0] : points[2];
    }

    // 斜線の先端を円の縁まで短縮する関数
    function shortenTip(points, hDir, r) {
        var tipIdx = (hDir === "left") ? 0 : 2;
        var bendIdx = 1;
        var tip = points[tipIdx];
        var bend = points[bendIdx];
        var dx = bend[0] - tip[0];
        var dy = bend[1] - tip[1];
        var len = Math.sqrt(dx * dx + dy * dy);
        if (len > 0) {
            points[tipIdx] = [tip[0] + dx / len * r, tip[1] + dy / len * r];
        }
    }

    // 先端に円を追加する関数
    function addTipCircle(tipPt, b, diameter, fillOnly) {
        var r = diameter / 2;
        var lineColor = getLineColor(b);
        var sw = getLineWidth(b);
        var circle = doc.pathItems.ellipse(
            tipPt[1] + r, tipPt[0] - r, diameter, diameter
        );
        if (fillOnly) {
            // ●（塗りのみ）
            circle.filled = true;
            circle.fillColor = lineColor;
            circle.stroked = false;
        } else {
            // ○（線のみ）
            circle.filled = false;
            circle.stroked = true;
            circle.strokeWidth = sw;
            circle.strokeColor = lineColor;
        }
        return circle;
    }

    // 先端の円にフチを追加する関数
    function addTipCircleWhiteEdge(tipPt, b, diameter, fillOnly) {
        var sw = getLineWidth(b);
        var whiteColor = createWhiteColor();
        if (fillOnly) {
            // ●のフチ：塗りを白で大きめ円
            var expand = sw * 2;
            var newDiam = diameter + expand;
            var newR = newDiam / 2;
            var whiteCircle = doc.pathItems.ellipse(
                tipPt[1] + newR, tipPt[0] - newR, newDiam, newDiam
            );
            whiteCircle.filled = true;
            whiteCircle.fillColor = whiteColor;
            whiteCircle.stroked = false;
        } else {
            // ○のフチ：線ではなく白い塗り円を背面に置く
            var expand = sw * 2;
            var newDiam = diameter + expand;
            var newR = newDiam / 2;
            var whiteCircle = doc.pathItems.ellipse(
                tipPt[1] + newR, tipPt[0] - newR, newDiam, newDiam
            );
            whiteCircle.filled = true;
            whiteCircle.fillColor = whiteColor;
            whiteCircle.stroked = false;
        }
        return whiteCircle;
    }

    // 先端に矢印を追加する関数
    function addTipArrow(tipPt, bendPt, b, arrowSize, fillOnly) {
        var lineColor = getLineColor(b);
        var sw = getLineWidth(b);
        // tipPt → bendPt 方向のベクトル
        var dx = bendPt[0] - tipPt[0];
        var dy = bendPt[1] - tipPt[1];
        var len = Math.sqrt(dx * dx + dy * dy);
        if (len === 0) return null;
        var ux = dx / len;
        var uy = dy / len;
        // 矢印の2つの翼端を計算
        var halfW = arrowSize / 2;
        var backX = tipPt[0] + ux * arrowSize;
        var backY = tipPt[1] + uy * arrowSize;
        var wing1 = [backX + uy * halfW, backY - ux * halfW];
        var wing2 = [backX - uy * halfW, backY + ux * halfW];

        var arrow = doc.pathItems.add();
        arrow.setEntirePath([tipPt, wing1, wing2]);
        arrow.closed = true;
        if (fillOnly) {
            arrow.filled = true;
            arrow.fillColor = lineColor;
            arrow.stroked = false;
        } else {
            arrow.filled = false;
            arrow.stroked = true;
            arrow.strokeWidth = sw;
            arrow.strokeColor = lineColor;
        }
        return arrow;
    }

    // 先端の矢印にフチを追加する関数（本体と同じ重心を基準にスケールアップ）
    function addTipArrowWhiteEdge(tipPt, bendPt, b, arrowSize, fillOnly) {
        var sw = getLineWidth(b);
        var whiteColor = createWhiteColor();
        var dx = bendPt[0] - tipPt[0];
        var dy = bendPt[1] - tipPt[1];
        var len = Math.sqrt(dx * dx + dy * dy);
        if (len === 0) return null;
        var ux = dx / len;
        var uy = dy / len;
        // 本体の矢印の3頂点を計算
        var halfW = arrowSize / 2;
        var backX = tipPt[0] + ux * arrowSize;
        var backY = tipPt[1] + uy * arrowSize;
        var p0 = tipPt;
        var p1 = [backX + uy * halfW, backY - ux * halfW];
        var p2 = [backX - uy * halfW, backY + ux * halfW];
        // 重心を求める
        var cx = (p0[0] + p1[0] + p2[0]) / 3;
        var cy = (p0[1] + p1[1] + p2[1]) / 3;
        // 重心を基準にスケールアップ
        var scale = (arrowSize + sw * 3) / arrowSize;
        var fp0 = [cx + (p0[0] - cx) * scale, cy + (p0[1] - cy) * scale];
        var fp1 = [cx + (p1[0] - cx) * scale, cy + (p1[1] - cy) * scale];
        var fp2 = [cx + (p2[0] - cx) * scale, cy + (p2[1] - cy) * scale];

        var whiteArrow = doc.pathItems.add();
        whiteArrow.setEntirePath([fp0, fp1, fp2]);
        whiteArrow.closed = true;
        whiteArrow.filled = true;
        whiteArrow.fillColor = whiteColor;
        whiteArrow.stroked = false;
        return whiteArrow;
    }

    function assembleLeaderParts(container, whitePath, newPath, whiteCircle, circle, hasWhiteEdge) {
        var edgeGrp = null;
        var mainGrp = container;

        if (hasWhiteEdge) {
            edgeGrp = container.groupItems.add();
            mainGrp = container.groupItems.add();

            if (newPath) newPath.move(mainGrp, ElementPlacement.PLACEATEND);
            if (circle) circle.move(mainGrp, ElementPlacement.PLACEATEND);

            if (whitePath) whitePath.move(edgeGrp, ElementPlacement.PLACEATEND);
            if (whiteCircle) whiteCircle.move(edgeGrp, ElementPlacement.PLACEATEND);
        } else {
            if (newPath) newPath.move(container, ElementPlacement.PLACEATEND);
            if (circle) circle.move(container, ElementPlacement.PLACEATEND);
        }

        return {
            edgeGroup: edgeGrp,
            mainGroup: mainGrp
        };
    }

    function tagLeaderParts(parts, whitePath, newPath, whiteCircle, circle, hasWhiteEdge) {
        if (hasWhiteEdge) {
            if (parts.mainGroup) {
                try { parts.mainGroup.note = "leader_line_main"; } catch (e) { }
            }
            if (parts.edgeGroup) {
                try { parts.edgeGroup.note = "leader_line_edge"; } catch (e) { }
            }
        }

        if (newPath) {
            try { newPath.note = "leader_line_main_path"; } catch (e) { }
        }
        if (circle) {
            try { circle.note = "leader_line_main_cap"; } catch (e) { }
        }
        if (whitePath) {
            try { whitePath.note = "leader_line_edge_path"; } catch (e) { }
        }
        if (whiteCircle) {
            try { whiteCircle.note = "leader_line_edge_cap"; } catch (e) { }
        }
    }

    function setLeaderLineTag(item) {
        if (!item) return;
        try {
            item.note = "leader_line";
        } catch (e) { }
    }

    function setTagValue(item, name, value) {
        if (!item) return;
        try {
            for (var i = 0; i < item.tags.length; i++) {
                if (item.tags[i].name === name) {
                    item.tags[i].value = String(value);
                    return;
                }
            }
            var t = item.tags.add();
            t.name = name;
            t.value = String(value);
        } catch (e) { }
    }

    function getTagValue(item, name) {
        if (!item) return null;
        try {
            for (var i = 0; i < item.tags.length; i++) {
                if (item.tags[i].name === name) return item.tags[i].value;
            }
        } catch (e) { }
        return null;
    }

    function setLeaderLineBoundsTags(item, b) {
        if (!item || !b) return;
        setTagValue(item, "leader_line_x_left", formatNumber(b.x_left, 4));
        setTagValue(item, "leader_line_x_right", formatNumber(b.x_right, 4));
        setTagValue(item, "leader_line_y_top", formatNumber(b.y_top, 4));
        setTagValue(item, "leader_line_y_bottom", formatNumber(b.y_bottom, 4));
    }

    function setLeaderLineDirTags(item, hDir, vDir) {
        if (!item) return;
        setTagValue(item, "leader_line_hDir", hDir);
        setTagValue(item, "leader_line_vDir", vDir);
    }

    function getLeaderLineStoredDir(item) {
        var hDir = getTagValue(item, "leader_line_hDir");
        var vDir = getTagValue(item, "leader_line_vDir");
        if (!hDir || !vDir) return null;
        return { hDir: hDir, vDir: vDir };
    }

    function getLeaderLineStoredBounds(item) {
        var xLeft = parseFloat(getTagValue(item, "leader_line_x_left"));
        var xRight = parseFloat(getTagValue(item, "leader_line_x_right"));
        var yTop = parseFloat(getTagValue(item, "leader_line_y_top"));
        var yBottom = parseFloat(getTagValue(item, "leader_line_y_bottom"));
        if (isNaN(xLeft) || isNaN(xRight) || isNaN(yTop) || isNaN(yBottom)) return null;
        return {
            x_left: xLeft,
            x_right: xRight,
            y_top: yTop,
            y_bottom: yBottom
        };
    }

    function isLeaderLineGroup(item) {
        return item && item.typename === "GroupItem" && item.note === "leader_line";
    }

    // (getOrCreateLeaderLineBackupLayer, stashOriginalTarget, getTargetCreationLayer removed)

    function isWhiteCMYKColor(color) {
        if (!color) return false;
        if (color.typename !== "CMYKColor") return false;
        return color.cyan === 0 && color.magenta === 0 && color.yellow === 0 && color.black === 0;
    }

    function getLeaderLineBasePath(groupItem) {
        var i, pi;

        // 再適用時の基準線は常に A線のみとし、B/C/D は参照しない
        // 新構造なら note 付きの A線を最優先
        for (i = 0; i < groupItem.pathItems.length; i++) {
            pi = groupItem.pathItems[i];
            if (pi.note === "leader_line_main_path") return pi;
        }

        // 旧構造用フォールバック：
        // 白でない stroked の open path を優先
        for (i = 0; i < groupItem.pathItems.length; i++) {
            pi = groupItem.pathItems[i];
            if (pi.stroked && !pi.closed && !isWhiteCMYKColor(pi.strokeColor)) return pi;
        }

        // 最後のフォールバック：stroked の open path
        for (i = 0; i < groupItem.pathItems.length; i++) {
            pi = groupItem.pathItems[i];
            if (pi.stroked && !pi.closed) return pi;
        }

        return null;
    }

    // プレビューを生成する関数
    function createPreview(angleDeg) {
        removePreview();
        var angleRad = angleDeg * Math.PI / 180;
        var boundsData = collectBoundsData(targets);
        var diagDirPreview = getDiagDirValues();
        var hDir = diagDirPreview.hDir;
        var vDir = diagDirPreview.vDir;
        var addCircle = capCircle.value;
        var addArrow = capArrow.value;
        var addCap = addCircle || addArrow;
        var diameter = unitValueToPt(capSizeInput.text, "strokeUnits");
        if (isNaN(diameter) || diameter <= 0) diameter = 3;

        for (var i = 0; i < boundsData.length; i++) {
            var b = boundsData[i];

            // 「斜線の方向以外」のとき、各オブジェクトの保存済み方向を使用
            var objHDir = hDir;
            var objVDir = vDir;
            if (scopeExceptDir.value) {
                var storedDir = getLeaderLineStoredDir(targets[i]);
                if (storedDir) {
                    objHDir = storedDir.hDir;
                    objVDir = storedDir.vDir;
                }
            }

            var sw = getLineWidth(b);
            var points = calcLeaderPoints(b, angleRad, objHDir, objVDir);

            // 先端座標は短縮前に取得
            var origTipPt = addCap ? getTipPoint(points, objHDir).slice(0) : null;
            var bendPt = addArrow ? points[1].slice(0) : null;

            if (addCircle && !capFill.value) {
                // 円の半径 + 円の線幅の半分で短縮（線が円の縁に接する）
                var shortenR = diameter / 2 + sw / 2;
                shortenTip(points, objHDir, shortenR);
            }
            if (addArrow) {
                // 矢印の長さ分だけ短縮
                shortenTip(points, objHDir, diameter);
            }

            var whitePath = null;
            var newPath = null;
            var whiteCircle = null;
            var circle = null;

            // フチ（最背面）
            if (whiteEdgeCheck.value) {
                whitePath = doc.pathItems.add();
                whitePath.setEntirePath(points);
                whitePath.stroked = true;
                whitePath.strokeWidth = sw * 3;
                whitePath.strokeColor = createWhiteColor();
                whitePath.filled = false;
                whitePath.strokeCap = capRound.value ? StrokeCap.ROUNDENDCAP : StrokeCap.BUTTENDCAP;
            }

            newPath = doc.pathItems.add();
            newPath.setEntirePath(points);
            newPath.stroked = true;
            newPath.strokeWidth = sw;
            newPath.strokeColor = getLineColor(b);
            newPath.filled = false;
            newPath.strokeCap = capRound.value ? StrokeCap.ROUNDENDCAP : StrokeCap.BUTTENDCAP;

            if (addCircle) {
                if (whiteEdgeCheck.value) {
                    whiteCircle = addTipCircleWhiteEdge(origTipPt, b, diameter, capFill.value);
                }
                circle = addTipCircle(origTipPt, b, diameter, capFill.value);
            }
            if (addArrow) {
                if (whiteEdgeCheck.value) {
                    whiteCircle = addTipArrowWhiteEdge(origTipPt, bendPt, b, diameter, capFill.value);
                }
                circle = addTipArrow(origTipPt, bendPt, b, diameter, capFill.value);
            }

            var previewGrp = doc.groupItems.add();
            assembleLeaderParts(previewGrp, whitePath, newPath, whiteCircle, circle, whiteEdgeCheck.value);
            previewItems.push(previewGrp);
        }

        // 元のパスを非表示にする
        for (var j = 0; j < targets.length; j++) {
            targets[j].hidden = true;
        }
        app.redraw();
    }

    // プレビューを削除する関数
    function removePreview() {
        for (var i = previewItems.length - 1; i >= 0; i--) {
            try { previewItems[i].remove(); } catch (e) { }
        }
        previewItems = [];
    }

    // ↑↓キーで値を増減する関数
    function changeValueByArrowKey(editText, options) {
        options = options || {};
        var smallStep = (typeof options.smallStep === "number") ? options.smallStep : 1;
        var largeStep = (typeof options.largeStep === "number") ? options.largeStep : 10;
        var fineStep = (typeof options.fineStep === "number") ? options.fineStep : 0.1;
        var minValue = (typeof options.minValue === "number") ? options.minValue : 0;
        var digits = (typeof options.digits === "number") ? options.digits : 0;

        editText.addEventListener("keydown", function (event) {
            if (event.keyName !== "Up" && event.keyName !== "Down") return;

            var value = parseFloat(editText.text);
            if (isNaN(value)) return;

            var keyboard = ScriptUI.environment.keyboardState;
            var step = smallStep;
            if (keyboard.shiftKey) {
                step = largeStep;
            } else if (keyboard.altKey) {
                step = fineStep;
            }

            if (event.keyName === "Up") {
                value += step;
            } else {
                value -= step;
                if (value < minValue) value = minValue;
            }

            editText.text = formatNumber(value, digits);
            event.preventDefault();
            updatePreview();
        });
    }

    // 前回の設定を取得 / Load previous session settings
    var s = $.global._leaderLineSettings;
    var strokeUnitLabel = getCurrentUnitLabel("strokeUnits");

    // ダイアログボックスで角度を指定
    var dlg = new Window("dialog", L("dialogTitle") + " " + SCRIPT_VERSION);
    dlg.orientation = "column";
    dlg.alignChildren = ["fill", "top"];
    // ダイアログ位置を復元（サイズではなく位置のみ）
    if (s.dialogBounds && s.dialogBounds.length === 2) {
        try {
            dlg.location = s.dialogBounds;
        } catch (e) { }
    }

    // 適用範囲ラジオボタン
    var scopePanel = dlg.add("panel", undefined, L("panelApplyScope"));
    scopePanel.margins = [15, 20, 15, 10];
    scopePanel.orientation = "row";
    scopePanel.alignChildren = ["center", "center"];
    var scopeAll = scopePanel.add("radiobutton", undefined, L("scopeAll"));
    var scopeExceptDir = scopePanel.add("radiobutton", undefined, L("scopeExceptDirection"));
    // 複数オブジェクト選択時は「斜線の方向以外」を自動選択
    if (targets.length > 1) {
        scopeExceptDir.value = true;
    } else if (s.applyScope === "exceptDirection") {
        scopeExceptDir.value = true;
    } else {
        scopeAll.value = true;
    }

    // 上段：2カラム
    var topRow = dlg.add("group");
    topRow.orientation = "row";
    topRow.alignChildren = ["fill", "top"];

    // 左カラム
    var leftCol = topRow.add("group");
    leftCol.orientation = "column";
    leftCol.alignChildren = ["fill", "top"];

    // 角度パネル
    var anglePanel = leftCol.add("panel", undefined, L("panelAngle"));
    anglePanel.margins = [15, 20, 15, 10];
    anglePanel.orientation = "column";
    anglePanel.alignChildren = ["fill", "top"];

    var angleInputGroup = anglePanel.add("group");
    angleInputGroup.alignment = ["center", "top"];
    var angleInput = angleInputGroup.add("edittext", undefined, s.angle);
    angleInput.characters = 4;
    angleInputGroup.add("statictext", undefined, "\u00B0");
    angleInput.active = true;
    changeValueByArrowKey(angleInput, { smallStep: 1, largeStep: 10, fineStep: 0.1, minValue: 0, digits: 1 });

    // ラジオボタンで角度を選択
    var radioGroup = anglePanel.add("group");
    radioGroup.orientation = "row";
    var radio30 = radioGroup.add("radiobutton", undefined, "30");
    var radio45 = radioGroup.add("radiobutton", undefined, "45");
    var radio60 = radioGroup.add("radiobutton", undefined, "60");
    if (s.radioAngle === 30) radio30.value = true;
    else if (s.radioAngle === 60) radio60.value = true;
    else radio45.value = true;

    radio30.onClick = function () { angleInput.text = "30"; updatePreview(); };
    radio45.onClick = function () { angleInput.text = "45"; updatePreview(); };
    radio60.onClick = function () { angleInput.text = "60"; updatePreview(); };

    // 斜線の方向パネル（左カラム）
    var dirPanel = leftCol.add("panel", undefined, L("panelDirection"));
    dirPanel.margins = [15, 20, 15, 10];
    dirPanel.orientation = "column";
    dirPanel.alignChildren = ["fill", "top"];

    var dirGroup = dirPanel.add("group");
    dirGroup.orientation = "row";
    dirGroup.alignChildren = ["fill", "top"];

    // 左列
    var dirLeftCol = dirGroup.add("group");
    dirLeftCol.orientation = "column";
    dirLeftCol.alignChildren = ["fill", "top"];
    var dirUpperLeft = dirLeftCol.add("radiobutton", undefined, L("dirUpperLeft"));
    dirLeftCol.add("statictext", undefined, " ");
    var dirLowerLeft = dirLeftCol.add("radiobutton", undefined, L("dirLowerLeft"));

    // 中央列（「対象」ラベル）
    var dirCenterCol = dirGroup.add("group");
    dirCenterCol.orientation = "column";
    dirCenterCol.alignChildren = ["center", "center"];
    dirCenterCol.add("statictext", undefined, " ");
    dirCenterCol.add("statictext", undefined, "対象");

    // 右列
    var dirRightCol = dirGroup.add("group");
    dirRightCol.orientation = "column";
    dirRightCol.alignChildren = ["fill", "top"];
    var dirUpperRight = dirRightCol.add("radiobutton", undefined, L("dirUpperRight"));
    dirRightCol.add("statictext", undefined, " ");
    var dirLowerRight = dirRightCol.add("radiobutton", undefined, L("dirLowerRight"));
    if (s.diagDir === "lowerLeft") dirLowerLeft.value = true;
    else if (s.diagDir === "upperRight") dirUpperRight.value = true;
    else if (s.diagDir === "lowerRight") dirLowerRight.value = true;
    else dirUpperLeft.value = true;

    // diagDir → hDir/vDir マッピング
    function getDiagDirValues() {
        if (dirUpperLeft.value) return { hDir: "right", vDir: "down" };
        if (dirLowerLeft.value) return { hDir: "right", vDir: "up" };
        if (dirUpperRight.value) return { hDir: "left", vDir: "down" };
        if (dirLowerRight.value) return { hDir: "left", vDir: "up" };
        return { hDir: "right", vDir: "up" };
    }

    // 異なるグループ間のラジオボタンを排他制御
    var allDirRadios = [dirUpperLeft, dirLowerLeft, dirUpperRight, dirLowerRight];
    function uncheckOtherDirs(selected) {
        for (var d = 0; d < allDirRadios.length; d++) {
            if (allDirRadios[d] !== selected) allDirRadios[d].value = false;
        }
    }
    dirUpperLeft.onClick = function () { uncheckOtherDirs(dirUpperLeft); updatePreview(); };
    dirLowerLeft.onClick = function () { uncheckOtherDirs(dirLowerLeft); updatePreview(); };
    dirUpperRight.onClick = function () { uncheckOtherDirs(dirUpperRight); updatePreview(); };
    dirLowerRight.onClick = function () { uncheckOtherDirs(dirLowerRight); updatePreview(); };

    // 「斜線の方向以外」選択時に方向パネルをディム
    function updateDirPanelEnabled() {
        var on = scopeAll.value;
        dirPanel.enabled = on;
    }
    updateDirPanelEnabled();

    scopeAll.onClick = function () { updateDirPanelEnabled(); updatePreview(); };
    scopeExceptDir.onClick = function () { updateDirPanelEnabled(); updatePreview(); };

    // 右カラム
    var rightCol = topRow.add("group");
    rightCol.orientation = "column";
    rightCol.alignChildren = ["fill", "top"];

    // 線のスタイルパネル
    var colorPanel = rightCol.add("panel", undefined, L("panelLineStyle"));
    colorPanel.margins = [15, 20, 15, 10];
    colorPanel.orientation = "column";
    colorPanel.alignChildren = ["fill", "top"];

    var colorGroup = colorPanel.add("group");
    var colorBlack = colorGroup.add("radiobutton", undefined, L("colorBlack"));
    var colorWhite = colorGroup.add("radiobutton", undefined, L("colorWhite"));
    var colorOther = colorGroup.add("radiobutton", undefined, L("colorOther"));
    if (s.lineColor === "white") colorWhite.value = true;
    else if (s.lineColor === "other") colorOther.value = true;
    else colorBlack.value = true;

    var hexGroup = colorPanel.add("group");
    hexGroup.orientation = "row";
    hexGroup.alignChildren = ["left", "center"];
    var hexInput = hexGroup.add("edittext", undefined, s.lineColorHex);
    hexInput.characters = 8;
    hexInput.enabled = colorOther.value;

    // カラースウォッチ
    var colorSwatch = hexGroup.add("panel", undefined, "");
    colorSwatch.preferredSize = [20, 20];
    function updateSwatch() {
        var hex = hexInput.text.replace(/^#/, "");
        if (hex.length === 3) {
            hex = hex.charAt(0) + hex.charAt(0) + hex.charAt(1) + hex.charAt(1) + hex.charAt(2) + hex.charAt(2);
        }
        var r = parseInt(hex.substring(0, 2), 16) / 255;
        var g = parseInt(hex.substring(2, 4), 16) / 255;
        var b = parseInt(hex.substring(4, 6), 16) / 255;
        if (isNaN(r) || isNaN(g) || isNaN(b)) return;
        var gfx = colorSwatch.graphics;
        gfx.backgroundColor = gfx.newBrush(gfx.BrushType.SOLID_COLOR, [r, g, b]);
    }
    updateSwatch();

    hexInput.onChanging = function () { updateSwatch(); updatePreview(); };

    colorBlack.onClick = function () { hexInput.enabled = false; updatePreview(); };
    colorWhite.onClick = function () { hexInput.enabled = false; updatePreview(); };
    colorOther.onClick = function () { hexInput.enabled = true; updatePreview(); };

    var lineWidthGroup = colorPanel.add("group");
    lineWidthGroup.add("statictext", undefined, L("lineWidth"));
    var lineWidthDisplay = parseFloat(formatUnitValue(s.lineWidth, "strokeUnits"));
    if (isNaN(lineWidthDisplay) || lineWidthDisplay <= 0) lineWidthDisplay = 1;
    var lineWidthInput = lineWidthGroup.add("edittext", undefined, formatNumber(lineWidthDisplay, 2));
    lineWidthInput.characters = 4;
    changeValueByArrowKey(lineWidthInput, { smallStep: 0.1, largeStep: 1, fineStep: 0.01, minValue: 0, digits: 2 });
    lineWidthGroup.add("statictext", undefined, strokeUnitLabel);
    lineWidthInput.onChanging = function () { updatePreview(); };

    // 線端：なし　丸型
    var strokeCapGroup = colorPanel.add("group");
    var strokeCapNone = strokeCapGroup.add("radiobutton", undefined, L("capStrokeCapNone"));
    var capRound = strokeCapGroup.add("radiobutton", undefined, L("capRound"));
    if (s.strokeCapType === "round") capRound.value = true;
    else strokeCapNone.value = true;
    capRound.onClick = function () { updatePreview(); };
    strokeCapNone.onClick = function () { updatePreview(); };

    // 線端パネル（右カラム）
    var capPanel = rightCol.add("panel", undefined, L("panelLineEnd"));
    capPanel.margins = [15, 20, 15, 10];
    capPanel.orientation = "column";
    capPanel.alignChildren = ["fill", "top"];

    // 形状：なし　円　矢印
    var capGroup = capPanel.add("group");
    var capNone = capGroup.add("radiobutton", undefined, L("capNone"));
    var capCircle = capGroup.add("radiobutton", undefined, L("capCircle"));
    var capArrow = capGroup.add("radiobutton", undefined, L("capArrow"));
    if (s.capType === "circle") capCircle.value = true;
    else if (s.capType === "arrow") capArrow.value = true;
    else capNone.value = true;

    var capStyleGroup = capPanel.add("group");
    var capFill = capStyleGroup.add("radiobutton", undefined, L("capFill"));
    var capStroke = capStyleGroup.add("radiobutton", undefined, L("capStroke"));
    if (s.capStyle === "stroke") capStroke.value = true; else capFill.value = true;

    var capSizeGroup = capPanel.add("group");
    capSizeGroup.add("statictext", undefined, L("capSize"));
    var capSizeDisplay = parseFloat(formatUnitValue(s.capSize, "strokeUnits"));
    if (isNaN(capSizeDisplay) || capSizeDisplay <= 0) capSizeDisplay = 3;
    var capSizeInput = capSizeGroup.add("edittext", undefined, formatNumber(capSizeDisplay, 2));
    capSizeInput.characters = 4;
    changeValueByArrowKey(capSizeInput, { smallStep: 0.1, largeStep: 1, fineStep: 0.01, minValue: 0, digits: 2 });
    capSizeGroup.add("statictext", undefined, strokeUnitLabel);

    var groupCheck = capPanel.add("checkbox", undefined, L("groupEnabled"));
    groupCheck.value = s.groupEnabled;

    // 形状の有無に応じたUI制御
    function updateCapEnabled() {
        var on = capCircle.value || capArrow.value;
        if (capArrow.value) {
            if (capStroke.value) {
                capStroke.value = false;
                capFill.value = true;
            }
            capFill.enabled = false;
            capStroke.enabled = false;
        } else {
            capFill.enabled = on;
            capStroke.enabled = on;
        }
        capSizeInput.enabled = on;
    }
    updateCapEnabled();

    capNone.onClick = function () { updateCapEnabled(); updatePreview(); };
    capCircle.onClick = function () { updateCapEnabled(); updatePreview(); };
    capArrow.onClick = function () { updateCapEnabled(); updatePreview(); };
    capFill.onClick = function () { updatePreview(); };
    capStroke.onClick = function () { updatePreview(); };
    capSizeInput.onChanging = function () { updatePreview(); };

    // オプションパネル
    var optPanel = dlg.add("panel", undefined, L("panelOptions"));
    optPanel.margins = [15, 20, 15, 10];
    optPanel.orientation = "column";
    optPanel.alignChildren = ["fill", "top"];

    var whiteEdgeGroup = optPanel.add("group");
    whiteEdgeGroup.orientation = "row";
    whiteEdgeGroup.alignChildren = ["left", "center"];
    var whiteEdgeCheck = whiteEdgeGroup.add("checkbox", undefined, L("whiteEdge"));
    whiteEdgeCheck.value = s.whiteEdge;
    whiteEdgeCheck.onClick = function () { updatePreview(); };
    var whiteEdgeSwatch = whiteEdgeGroup.add("panel", undefined, "");
    whiteEdgeSwatch.preferredSize = [20, 20];
    var gfx = whiteEdgeSwatch.graphics;
    gfx.backgroundColor = gfx.newBrush(gfx.BrushType.SOLID_COLOR, [1, 1, 1]);

    var previewCheck = dlg.add("checkbox", undefined, L("preview"));
    previewCheck.alignment = ["center", "top"];
    previewCheck.value = !!s.preview;

    var btnGroup = dlg.add("group");
    btnGroup.alignment = ["center", "top"];
    btnGroup.add("button", undefined, L("cancel"), { name: "cancel" });
    btnGroup.add("button", undefined, L("ok"), { name: "ok" });

    // プレビュー更新の共通処理
    function updatePreview() {
        if (!previewCheck.value) return;
        var val = parseFloat(angleInput.text);
        if (isNaN(val) || val <= 0 || val >= 90) return;
        createPreview(val);
    }

    // プレビューチェックボックスの変更
    previewCheck.onClick = function () {
        if (previewCheck.value) {
            updatePreview();
        } else {
            removePreview();
            for (var i = 0; i < targets.length; i++) {
                targets[i].hidden = false;
            }
            app.redraw();
        }
    };

    // 角度入力の変更
    angleInput.onChanging = function () {
        updatePreview();
    };

    var result = dlg.show();
    try {
        if (dlg.location) {
            s.dialogBounds = [dlg.location[0], dlg.location[1]];
        }
    } catch (e) { }

    // 設定を保存
    s.angle = angleInput.text;
    s.radioAngle = radio30.value ? 30 : (radio60.value ? 60 : 45);
    s.applyScope = scopeExceptDir.value ? "exceptDirection" : "all";
    var diagDirVals = getDiagDirValues();
    s.diagDir = dirUpperLeft.value ? "upperLeft" : (dirLowerLeft.value ? "lowerLeft" : (dirUpperRight.value ? "upperRight" : "lowerRight"));
    s.hDir = diagDirVals.hDir;
    s.vDir = diagDirVals.vDir;
    s.capType = capCircle.value ? "circle" : (capArrow.value ? "arrow" : "none");
    s.capStyle = capStroke.value ? "stroke" : "fill";
    var savedCapSizePt = unitValueToPt(capSizeInput.text, "strokeUnits");
    s.capSize = (!isNaN(savedCapSizePt) && savedCapSizePt > 0) ? formatNumber(savedCapSizePt, 4) : "3";
    s.strokeCapType = capRound.value ? "round" : "none";
    s.groupEnabled = groupCheck.value;
    s.whiteEdge = whiteEdgeCheck.value;
    s.preview = previewCheck.value;
    s.lineColor = colorBlack.value ? "black" : (colorWhite.value ? "white" : "other");
    s.lineColorHex = hexInput.text;
    var savedLineWidthPt = unitValueToPt(lineWidthInput.text, "strokeUnits");
    s.lineWidth = (!isNaN(savedLineWidthPt) && savedLineWidthPt > 0) ? formatNumber(savedLineWidthPt, 4) : "1";

    // プレビューを削除して元のパスを復元
    removePreview();
    for (var i = 0; i < targets.length; i++) {
        targets[i].hidden = false;
    }
    app.redraw();

    if (result !== 1) {
        return;
    }

    var angleDeg = parseFloat(angleInput.text);
    if (isNaN(angleDeg) || angleDeg <= 0 || angleDeg >= 90) {
        alert(L("alertInvalidAngle"));
        app.redraw();
        return;
    }
    var angleRad = angleDeg * Math.PI / 180;

    // 確定：引き出し線を生成

    var diagDirResult = getDiagDirValues();
    var hDir = diagDirResult.hDir;
    var vDir = diagDirResult.vDir;
    var addCircle = capCircle.value;
    var addArrow = capArrow.value;
    var addCap = addCircle || addArrow;
    var diameter = unitValueToPt(capSizeInput.text, "strokeUnits");
    if (isNaN(diameter) || diameter <= 0) diameter = 3;
    var boundsData = collectBoundsData(targets);
    var finalSelectionItems = [];
    doc.selection = null;

    for (var i = 0; i < targets.length; i++) {
        var sourceTarget = targets[i];
        var b = boundsData[i];
        var createdItems = [];
        var createdSelectionItems = [];

        try {
            // 「斜線の方向以外」のとき、各オブジェクトの保存済み方向を使用
            var objHDir = hDir;
            var objVDir = vDir;
            if (scopeExceptDir.value) {
                var storedDir = getLeaderLineStoredDir(sourceTarget);
                if (storedDir) {
                    objHDir = storedDir.hDir;
                    objVDir = storedDir.vDir;
                }
            }

            var sw = getLineWidth(b);
            var points = calcLeaderPoints(b, angleRad, objHDir, objVDir);

            // 先端座標は短縮前に取得
            var origTipPt = addCap ? getTipPoint(points, objHDir).slice(0) : null;
            var bendPt = addArrow ? points[1].slice(0) : null;

            if (addCircle && !capFill.value) {
                var shortenR = diameter / 2 + sw / 2;
                shortenTip(points, objHDir, shortenR);
            }
            if (addArrow) {
                shortenTip(points, objHDir, diameter);
            }

            var whitePath = null;
            var newPath = null;
            var whiteCircle = null;
            var circle = null;

            // フチ（最背面）
            if (whiteEdgeCheck.value) {
                whitePath = doc.pathItems.add();
                createdItems.push(whitePath);
                whitePath.setEntirePath(points);
                whitePath.stroked = true;
                whitePath.strokeWidth = sw * 3;
                whitePath.strokeColor = createWhiteColor();
                whitePath.filled = false;
                whitePath.strokeCap = capRound.value ? StrokeCap.ROUNDENDCAP : StrokeCap.BUTTENDCAP;
            }

            newPath = doc.pathItems.add();
            createdItems.push(newPath);
            newPath.setEntirePath(points);
            newPath.stroked = true;
            newPath.strokeWidth = sw;
            newPath.strokeColor = getLineColor(b);
            newPath.filled = false;
            newPath.strokeCap = capRound.value ? StrokeCap.ROUNDENDCAP : StrokeCap.BUTTENDCAP;

            if (addCircle) {
                if (whiteEdgeCheck.value) {
                    whiteCircle = addTipCircleWhiteEdge(origTipPt, b, diameter, capFill.value);
                    createdItems.push(whiteCircle);
                }
                circle = addTipCircle(origTipPt, b, diameter, capFill.value);
                createdItems.push(circle);
            }
            if (addArrow) {
                if (whiteEdgeCheck.value) {
                    whiteCircle = addTipArrowWhiteEdge(origTipPt, bendPt, b, diameter, capFill.value);
                    createdItems.push(whiteCircle);
                }
                circle = addTipArrow(origTipPt, bendPt, b, diameter, capFill.value);
                createdItems.push(circle);
            }

            if (groupCheck.value) {
                var grp = doc.groupItems.add();
                createdItems.push(grp);
                var assembled = assembleLeaderParts(grp, whitePath, newPath, whiteCircle, circle, whiteEdgeCheck.value);
                tagLeaderParts(assembled, whitePath, newPath, whiteCircle, circle, whiteEdgeCheck.value);
                setLeaderLineTag(grp);
                setLeaderLineBoundsTags(grp, b);
                setLeaderLineDirTags(grp, objHDir, objVDir);
                createdSelectionItems.push(grp);
            } else {
                if (whitePath) createdSelectionItems.push(whitePath);
                if (whiteCircle) createdSelectionItems.push(whiteCircle);
                if (newPath) createdSelectionItems.push(newPath);
                if (circle) createdSelectionItems.push(circle);
            }

            for (var cs = 0; cs < createdSelectionItems.length; cs++) {
                finalSelectionItems.push(createdSelectionItems[cs]);
            }

            // 置き換え成功時のみ元オブジェクトを削除
            sourceTarget.remove();
        } catch (eTarget) {
            for (var cr = createdItems.length - 1; cr >= 0; cr--) {
                try { createdItems[cr].remove(); } catch (eRemove) { }
            }
            throw eTarget;
        }
    }

    if (finalSelectionItems.length) {
        doc.selection = null;
        for (var si = 0; si < finalSelectionItems.length; si++) {
            try {
                finalSelectionItems[si].selected = true;
            } catch (e) { }
        }
    }
}

main();
