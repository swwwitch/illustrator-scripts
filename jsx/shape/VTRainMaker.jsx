#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
 * 概要 / Overview:
 * 選択した長方形、または現在のアートボードを対象に、ランダムな雨の罫線や雨粒を生成するIllustrator用スクリプトです。
 * 何も選択していない場合は現在のアートボード全体を対象にし、長方形を1つ選択している場合はその長方形の範囲内を対象にします。
 * 春雨・五月雨・鉄砲雨・雨粒・水玉の5種類のプリセットを切り替えながら、線幅・角度・密度・子罫オフセット・外側拡張・全体スケール・プレビューをダイアログで調整できます。
 * 子罫オフセットは親線に対する法線方向へ平行移動する仕様で、全体スケールは25〜300の範囲で入力欄とスライダーの制約を揃えています。
 * 雨粒プリセットでは塗り・線・A/B/C形状を切り替えられ、大きさと不透明度はランダムに変化します。
 * 単位表示は Illustrator の環境設定を参照し、線幅と子罫オフセットは strokeUnits、外側拡張は rulerType に合わせて表示・入力できます。
 *
 * This Illustrator script generates random rain lines and raindrops inside a selected rectangle or the active artboard.
 * If nothing is selected, it uses the current artboard. If a single rectangle is selected, it uses that rectangle’s bounds as the target area.
 * It provides five presets—Harusame, Samidare, Driving Rain, Raindrops, and Polka Dots—and lets you adjust stroke width, angle, density, child-line offset, outward expansion, overall scale, and preview in a dialog.
 * Child-line offset moves parallel to the parent line along its normal, and overall scale is constrained consistently between 25 and 300 in both the edit field and the slider.
 * In the Raindrops preset, you can switch fill, stroke, and A/B/C shapes, while size and opacity vary randomly.
 * Unit display follows Illustrator preferences: stroke width and child-line offset use strokeUnits, while outward expansion uses rulerType for both labels and input values.
 *
 * 作成日 / Created: 2026-03-27
 * 更新日 / Updated: 2026-03-28（長方形選択対応、子罫オフセットの法線方向化、外側拡張表記、全体スケール制約統一、概要更新 / Added selected-rectangle targeting, normal-based child offset, outward-expansion labeling, unified overall-scale constraints, and updated the overview）
 */

// =========================================
// バージョンとローカライズ / Version and localization
// =========================================
var SCRIPT_VERSION = "v1.0";

function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}

var lang = getCurrentLang();

function L(key) {
    var entry = LABELS[key];
    if (!entry) return key;
    return entry[lang] || entry.ja || key;
}

function labelText(key) {
    return L(key) + (lang === "ja" ? "：" : ":");
}

/* 日英ラベル定義 / Japanese-English label definitions */
var LABELS = {
    dialogTitle: {
        ja: "雨の罫線",
        en: "Rain Lines"
    },
    noDocument: {
        ja: "ドキュメントを開いてから実行してください。",
        en: "Open a document before running this script."
    },
    panelRainType: {
        ja: "雨の種類",
        en: "Rain Type"
    },
    panelOptions: {
        ja: "オプション",
        en: "Options"
    },
    panelRaindrop: {
        ja: "雨粒",
        en: "Raindrop"
    },
    rainHarusame: {
        ja: "春雨",
        en: "Harusame"
    },
    rainSamidare: {
        ja: "五月雨",
        en: "Samidare"
    },
    rainTeppouame: {
        ja: "鉄砲雨",
        en: "Driving Rain"
    },
    rainAmatsubu: {
        ja: "雨粒",
        en: "Raindrops"
    },
    rainMizutama: {
        ja: "水玉",
        en: "Polka Dots"
    },
    labelStrokeWidth: {
        ja: "線幅",
        en: "Stroke Width"
    },
    labelAngle: {
        ja: "角度",
        en: "Angle"
    },
    labelDensity: {
        ja: "密度",
        en: "Density"
    },
    unitLines: {
        ja: "本数",
        en: "Lines"
    },
    unitDegree: {
        ja: "°",
        en: "°"
    },
    labelMargin: {
        ja: "外側拡張",
        en: "Margin Expansion"
    },
    labelSpacing: {
        ja: "子罫オフセット",
        en: "Child Line Offset"
    },
    labelScale: {
        ja: "全体スケール",
        en: "Overall Scale"
    },
    unitPercent: {
        ja: "%",
        en: "%"
    },
    chkRaindropFill: {
        ja: "塗り",
        en: "Fill"
    },
    chkRaindropStroke: {
        ja: "線",
        en: "Stroke"
    },
    raindropShapeA: {
        ja: "A",
        en: "A"
    },
    raindropShapeB: {
        ja: "B",
        en: "B"
    },
    raindropShapeC: {
        ja: "C",
        en: "C"
    },
    invalidScale: {
        ja: "全体スケールには25〜300の数値を入力してください。",
        en: "Enter a value between 25 and 300 for the overall scale."
    },
    invalidAngle: {
        ja: "角度には数値を入力してください。",
        en: "Enter a numeric value for the angle."
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
    },
    invalidStrokeWidth: {
        ja: "線幅には正の数値を入力してください。",
        en: "Enter a positive number for the stroke width."
    },
    invalidDensity: {
        ja: "密度には正の整数を入力してください。",
        en: "Enter a positive integer for the density."
    },
    invalidSpacing: {
        ja: "子罫オフセットには0以上の数値を入力してください。",
        en: "Enter a number greater than or equal to 0 for the child offset."
    },
    invalidMargin: {
        ja: "外側拡張には0以上の数値を入力してください。",
        en: "Enter a number greater than or equal to 0 for the margin expansion."
    }
};


// =========================================
// 単位ユーティリティ / Unit utilities
// =========================================
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

/* 単位コードと設定キーから適切な単位ラベルを返す / Return the proper unit label from a unit code and preference key */
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

/* 単位コードから pt 換算係数を返す / Return the pt conversion factor from a unit code */
function getPtFactorFromUnitCode(code) {
    switch (code) {
        case 0: return 72.0;                 // in
        case 1: return 72.0 / 25.4;          // mm
        case 2: return 1.0;                  // pt
        case 3: return 12.0;                 // pica
        case 4: return 72.0 / 2.54;          // cm
        case 5: return 72.0 / 25.4 * 0.25;   // Q or H
        case 6: return 1.0;                  // px
        case 7: return 72.0 * 12.0;          // ft/in
        case 8: return 72.0 / 25.4 * 1000.0; // m
        case 9: return 72.0 * 36.0;          // yd
        case 10: return 72.0 * 12.0;         // ft
        default: return 1.0;
    }
}

/* 設定キーごとの単位情報を取得 / Get unit info for a preference key */
function getUnitInfo(prefKey) {
    var code = app.preferences.getIntegerPreference(prefKey);
    return {
        code: code,
        label: getUnitLabel(code, prefKey),
        factor: getPtFactorFromUnitCode(code)
    };
}

/* pt 値を現在の単位値へ変換 / Convert a pt value to the current display unit */
function ptToUnitValue(ptValue, prefKey, decimals) {
    var info = getUnitInfo(prefKey);
    var unitValue = ptValue / info.factor;
    if (typeof decimals === "number") {
        return parseFloat(unitValue.toFixed(decimals));
    }
    return unitValue;
}

/* 現在の単位値を pt 値へ変換 / Convert a current-unit value to pt */

function unitValueToPt(unitValue, prefKey) {
    var info = getUnitInfo(prefKey);
    return unitValue * info.factor;
}


function parseMarginPtFromText(text) {
    return unitValueToPt(parseFloat(text), "rulerType");
}

function parseScaleValueFromText(text) {
    return parseFloat(text);
}


function isValidMarginPt(marginPt) {
    return !isNaN(marginPt) && marginPt >= 0;
}

function clampScaleValue(scaleValue) {
    if (scaleValue < 25) return 25;
    if (scaleValue > 300) return 300;
    return scaleValue;
}


function isValidScaleValue(scaleValue) {
    return !isNaN(scaleValue) && scaleValue >= 25 && scaleValue <= 300;
}

/* ↑↓キーで数値を増減 / Increment or decrement numeric values with the arrow keys */
function changeValueByArrowKey(editText, allowDecimal, allowNegative) {
    editText.addEventListener("keydown", function (event) {
        var value = Number(editText.text);
        if (isNaN(value)) return;

        var keyboard = ScriptUI.environment.keyboardState;
        var useDecimalStep = !!allowDecimal && keyboard.altKey;
        var delta = useDecimalStep ? 0.1 : (keyboard.shiftKey ? 10 : 1);

        if (event.keyName == "Up") {
            if (keyboard.shiftKey) {
                value = Math.ceil((value + 1) / delta) * delta;
            } else {
                value += delta;
            }
            event.preventDefault();
        } else if (event.keyName == "Down") {
            if (keyboard.shiftKey) {
                value = Math.floor((value - 1) / delta) * delta;
            } else {
                value -= delta;
            }
            if (!allowNegative && value < 0) value = 0;
            event.preventDefault();
        } else {
            return;
        }

        if (useDecimalStep) {
            value = Math.round(value * 10) / 10;
        } else {
            value = Math.round(value);
        }

        editText.text = useDecimalStep ? value.toFixed(1).replace(/\.0$/, "") : String(value);

        if (typeof editText.onChanging === "function") {
            editText.onChanging();
        }
    });
}

function isRectanglePathItem(item) {
    if (!item || item.typename !== "PathItem") return false;
    if (!item.closed || item.guides || item.clipping) return false;
    if (item.pathPoints.length !== 4) return false;

    var gb = item.geometricBounds;
    var left = gb[0];
    var top = gb[1];
    var right = gb[2];
    var bottom = gb[3];
    var eps = 0.01;
    var hasLeft = false;
    var hasRight = false;
    var hasTop = false;
    var hasBottom = false;

    for (var i = 0; i < item.pathPoints.length; i++) {
        var anchor = item.pathPoints[i].anchor;
        var x = anchor[0];
        var y = anchor[1];

        if (Math.abs(x - left) <= eps) hasLeft = true;
        if (Math.abs(x - right) <= eps) hasRight = true;
        if (Math.abs(y - top) <= eps) hasTop = true;
        if (Math.abs(y - bottom) <= eps) hasBottom = true;
    }

    return hasLeft && hasRight && hasTop && hasBottom;
}

function getTargetBounds(doc) {
    var target = null;
    if (doc.selection && doc.selection.length === 1 && isRectanglePathItem(doc.selection[0])) {
        target = doc.selection[0].geometricBounds;
    } else {
        target = doc.artboards[doc.artboards.getActiveArtboardIndex()].artboardRect;
    }

    return {
        left: target[0],
        top: target[1],
        right: target[2],
        bottom: target[3]
    };
}

function drawRandomParentChildLines() {
    if (app.documents.length === 0) {
        alert(L("noDocument"));
        return;
    }

    var doc = app.activeDocument;

    // 対象範囲を取得（選択長方形があればその範囲、なければ現在のアートボード）
    var targetBounds = getTargetBounds(doc);
    var left = targetBounds.left;
    var top = targetBounds.top;
    var right = targetBounds.right;
    var bottom = targetBounds.bottom;

    var width = right - left;
    var height = top - bottom;

    var layer = doc.activeLayer;

    var rulerUnitInfo = getUnitInfo("rulerType");
    var strokeUnitInfo = getUnitInfo("strokeUnits");

    function parseMarginPt() {
        return parseMarginPtFromText(inputMargin.text);
    }

    function parseScaleValue() {
        return parseScaleValueFromText(inputScale.text);
    }

    var strokeColor = new CMYKColor();
    strokeColor.cyan = 0;
    strokeColor.magenta = 0;
    strokeColor.yellow = 0;
    strokeColor.black = 100;

    /* プリセット定義 / Preset definitions */
    // { lineLength, angleDeg, defaultStrokeWidth, defaultAngle, defaultDensity, defaultSpacing, defaultScale, roundCap, noChild, colors, penetrate, raindrop, polkadot }
    var presets = {
        harusame: { lineLength: 60, angleDeg: 45, defaultStrokeWidth: "2.0", defaultAngle: "45", defaultDensity: "50", defaultSpacing: "3", defaultScale: "100", avoidOverlap: true, placementGap: 0.95 },
        samidare: { lineLength: 80, angleDeg: 45, defaultStrokeWidth: "10.0", defaultAngle: "45", defaultDensity: "35", defaultSpacing: "0", defaultScale: "100", roundCap: true, noChild: true, colors: [[66, 122, 190], [207, 167, 204]], avoidOverlap: true, placementGap: 1.05 },
        teppouame: { lineLength: 0, angleDeg: 0, defaultStrokeWidth: "100", defaultAngle: "45", defaultDensity: "15", defaultSpacing: "0", defaultScale: "100", noChild: true, penetrate: true },
        amatsubu: { defaultStrokeWidth: "8.0", defaultAngle: "0", defaultDensity: "40", defaultSpacing: "0", defaultScale: "100", noChild: true, raindrop: true },
        mizutama: { defaultStrokeWidth: "8.0", defaultAngle: "0", defaultDensity: "35", defaultSpacing: "0", defaultScale: "100", noChild: true, polkadot: true, colors: [[177, 208, 236], [102, 197, 236]] }
    };

    /* プレビュー用グループの参照 / References to preview groups */
    var previewGroups = [];

    /* 現在のプリセットを取得 / Get the current preset */
    function getSelectedPreset() {
        if (rbSamidare.value) return presets.samidare;
        if (rbTeppouame.value) return presets.teppouame;
        if (rbAmatsubu.value) return presets.amatsubu;
        if (rbMizutama.value) return presets.mizutama;
        return presets.harusame;
    }

    /* プリセットのcolorsからランダムに1色選択 / Pick one random color from preset colors */
    function pickColor(colors) {
        var rgb = colors[Math.floor(Math.random() * colors.length)];
        var c = new RGBColor();
        c.red = rgb[0];
        c.green = rgb[1];
        c.blue = rgb[2];
        return c;
    }

    /* 雨粒用の塗り色を作成 / Create a fill color for raindrops */
    function createRaindropColor() {
        var c = new RGBColor();
        c.red = 0;
        c.green = 150;
        c.blue = 255;
        return c;
    }

    /* 雨粒用の線色を作成 / Create a stroke color for raindrops */
    function createRaindropStrokeColor() {
        var c = new RGBColor();
        c.red = 0;
        c.green = 100;
        c.blue = 200;
        return c;
    }

    /* ティアドロップ形状を作成 / Create a teardrop shape */
    function createRaindropShape(parentGroup, centerX, centerY, widthPt, heightPt, fillColor, angleDeg, styleOptions) {
        function rotatePoint(px, py) {
            var rad = (angleDeg || 0) * Math.PI / 180;
            var dx = px - centerX;
            var dy = py - centerY;
            var cosA = Math.cos(rad);
            var sinA = Math.sin(rad);
            return [
                centerX + dx * cosA - dy * sinA,
                centerY + dx * sinA + dy * cosA
            ];
        }

        var opt = {
            heightRatio: 1.8,
            midRatio: 0.55,
            bottomRatio: 0.80,
            downHandleScale: 0.35,
            topHandleLength: 0.40,
            topHandleInset: 0.15,
            strokeWidth: 2,
            filled: true,
            stroked: false,
            strokeColor: createRaindropStrokeColor()
        };

        if (styleOptions) {
            for (var key in styleOptions) {
                if (styleOptions.hasOwnProperty(key)) {
                    opt[key] = styleOptions[key];
                }
            }
        }

        var width = widthPt;
        var height = (typeof heightPt === "number" && heightPt > 0) ? heightPt : width * opt.heightRatio;
        var hw = width / 2;
        var hh = height / 2;

        var cx = centerX;
        var cy = centerY + hh * 0.80;

        var top = cy;
        var bottom = cy - height * opt.bottomRatio;
        var mid = cy - height * opt.midRatio;

        var k = 0.5523;
        var rx = hw;
        var ry = hh * 0.85;

        var pathPoints = [
            {
                anchor: [cx, top],
                leftDir: [cx - hw * opt.topHandleInset, top - hh * opt.topHandleLength],
                rightDir: [cx + hw * opt.topHandleInset, top - hh * opt.topHandleLength],
                pointType: PointType.CORNER
            },
            {
                anchor: [cx + rx, mid],
                leftDir: [cx + rx, mid + ry * k],
                rightDir: [cx + rx, mid - ry * opt.downHandleScale],
                pointType: PointType.SMOOTH
            },
            {
                anchor: [cx, bottom],
                leftDir: [cx + rx * k, bottom],
                rightDir: [cx - rx * k, bottom],
                pointType: PointType.SMOOTH
            },
            {
                anchor: [cx - rx, mid],
                leftDir: [cx - rx, mid - ry * opt.downHandleScale],
                rightDir: [cx - rx, mid + ry * k],
                pointType: PointType.SMOOTH
            }
        ];

        var item = parentGroup.pathItems.add();
        item.stroked = !!opt.stroked;
        if (item.stroked) {
            item.strokeColor = opt.strokeColor;
            item.strokeWidth = opt.strokeWidth;
        }
        item.filled = !!opt.filled;
        if (item.filled) {
            item.fillColor = fillColor;
        }
        item.closed = true;
        item.setEntirePath([
            rotatePoint(pathPoints[0].anchor[0], pathPoints[0].anchor[1]),
            rotatePoint(pathPoints[1].anchor[0], pathPoints[1].anchor[1]),
            rotatePoint(pathPoints[2].anchor[0], pathPoints[2].anchor[1]),
            rotatePoint(pathPoints[3].anchor[0], pathPoints[3].anchor[1])
        ]);

        var pts = item.pathPoints;
        for (var i = 0; i < pts.length; i++) {
            pts[i].leftDirection = rotatePoint(pathPoints[i].leftDir[0], pathPoints[i].leftDir[1]);
            pts[i].rightDirection = rotatePoint(pathPoints[i].rightDir[0], pathPoints[i].rightDir[1]);
            pts[i].pointType = pathPoints[i].pointType;
        }

        return item;
    }

    /* 水玉を作成 / Create a polka dot */
    function createPolkaDot(parentGroup, centerX, centerY, diameterPt, fillColor) {
        var ellipse = parentGroup.pathItems.ellipse(
            centerY + diameterPt / 2,
            centerX - diameterPt / 2,
            diameterPt,
            diameterPt
        );
        ellipse.stroked = false;
        ellipse.filled = true;
        ellipse.fillColor = fillColor;
        return ellipse;
    }

    /* 描画関数 / Draw lines */
    function drawLines(mainStrokeWidth, numLines) {

        var scaleValue = parseScaleValue();
        if (!isValidScaleValue(scaleValue)) return;
        var scaleFactor = scaleValue / 100;
        var preset = getSelectedPreset();
        var margin = parseMarginPt();
        if (!isValidMarginPt(margin)) return;

        if (preset.raindrop) {
            for (var r = 0; r < numLines; r++) {
                var dropX = (left - margin) + Math.random() * (width + margin * 2);
                var dropY = (bottom - margin) + Math.random() * (height + margin * 2);
                var baseDropWidth = Math.max((mainStrokeWidth * scaleFactor) * 1.1, 2);
                var dropWidth = Math.max((baseDropWidth * 0.6) + Math.random() * (baseDropWidth * 0.8), 2);
                var dropHeightRatio = 1.5;
                var dropOpacity = 35 + Math.random() * 65;
                var dropGroup = layer.groupItems.add();
                var raindropStyle = getRaindropAppearanceOptions();

                if (rbRaindropB.value) {
                    dropHeightRatio = 1.8;
                } else if (rbRaindropC.value) {
                    dropHeightRatio = 1.3;
                }

                var dropHeight = Math.max(dropWidth * dropHeightRatio, 3);

                createRaindropShape(
                    dropGroup,
                    dropX,
                    dropY,
                    dropWidth,
                    dropHeight,
                    createRaindropColor(),
                    parseFloat(inputAngle.text) || 0,
                    raindropStyle
                );
                dropGroup.opacity = dropOpacity;
                previewGroups.push(dropGroup);

            }
            app.redraw();
            return;
        }

        if (preset.polkadot) {
            for (var p = 0; p < numLines; p++) {
                var dotX = (left - margin) + Math.random() * (width + margin * 2);
                var dotY = (bottom - margin) + Math.random() * (height + margin * 2);
                var dotBase = mainStrokeWidth * scaleFactor;
                var dotDiameter = Math.max((dotBase * 0.5) + Math.random() * (dotBase * 1.5), 2);
                var dotGroup = layer.groupItems.add();
                createPolkaDot(dotGroup, dotX, dotY, dotDiameter, pickColor(preset.colors));
                previewGroups.push(dotGroup);
            }
            app.redraw();
            return;
        }

        var angleDeg = parseFloat(inputAngle.text);
        if (isNaN(angleDeg)) angleDeg = parseFloat(preset.defaultAngle || preset.angleDeg || 45);
        var angleRad = angleDeg * Math.PI / 180;
        var dx = preset.lineLength * scaleFactor * Math.cos(angleRad);
        var dy = preset.lineLength * scaleFactor * Math.sin(angleRad);
        var childOffset = unitValueToPt(parseFloat(inputSpacing.text) || 0, "strokeUnits") * scaleFactor;
        var childOffsetX = 0;
        var childOffsetY = 0;
        if (childOffset !== 0) {
            var lineLength = Math.sqrt(dx * dx + dy * dy);
            if (lineLength > 0) {
                childOffsetX = -dy / lineLength * childOffset;
                childOffsetY = dx / lineLength * childOffset;
            }
        }
        var placedCenters = [];
        var placementRadius = Math.max(preset.lineLength * scaleFactor * (preset.placementGap || 1), mainStrokeWidth * scaleFactor * 4);
        var maxPlacementTries = 200;

        for (var i = 0; i < numLines; i++) {
            var startX, startY, endX, endY;
            var currentStrokeWidth = mainStrokeWidth;
            var scaledStrokeWidth = mainStrokeWidth * scaleFactor;

            if (preset.penetrate) {
                var randAngle = (75 + Math.random() * 30) * Math.PI / 180;
                var penetrateLen = (height + margin * 2) / Math.sin(randAngle);
                startX = (left - margin) + Math.random() * (width + margin * 2);
                startY = top + margin;
                endX = startX + penetrateLen * Math.cos(randAngle);
                endY = startY - penetrateLen * Math.sin(randAngle);
                currentStrokeWidth = 0.5 + Math.random() * 4.5;
            } else {
                var placed = false;
                var tryCount;
                for (tryCount = 0; tryCount < maxPlacementTries; tryCount++) {
                    startX = (left - margin) + Math.random() * (width + margin * 2);
                    startY = (bottom - margin) + Math.random() * (height + margin * 2);
                    endX = startX + dx;
                    endY = startY + dy;

                    if (!preset.avoidOverlap) {
                        placed = true;
                        break;
                    }

                    var centerX = (startX + endX) / 2;
                    var centerY = (startY + endY) / 2;
                    var overlaps = false;
                    for (var j = 0; j < placedCenters.length; j++) {
                        var pc = placedCenters[j];
                        var ddx = centerX - pc.x;
                        var ddy = centerY - pc.y;
                        if (Math.sqrt(ddx * ddx + ddy * ddy) < placementRadius) {
                            overlaps = true;
                            break;
                        }
                    }
                    if (!overlaps) {
                        placed = true;
                        placedCenters.push({ x: centerX, y: centerY });
                        break;
                    }
                }

                if (!placed) {
                    startX = (left - margin) + Math.random() * (width + margin * 2);
                    startY = (bottom - margin) + Math.random() * (height + margin * 2);
                    endX = startX + dx;
                    endY = startY + dy;
                    if (preset.avoidOverlap) {
                        placedCenters.push({ x: (startX + endX) / 2, y: (startY + endY) / 2 });
                    }
                }
            }

            var lineColor = preset.colors ? pickColor(preset.colors) : strokeColor;

            var group = layer.groupItems.add();

            /* 子罫線（点線）— noChild の場合はスキップ / Child line (dashed) — skip when noChild is enabled */
            if (!preset.noChild) {
                var childLine = group.pathItems.add();
                /* 親線に対する法線方向へ平行移動 / Offset parallel to the parent line along its normal */
                childLine.setEntirePath([[startX + childOffsetX, startY + childOffsetY], [endX + childOffsetX, endY + childOffsetY]]);
                childLine.stroked = true;
                childLine.strokeWidth = scaledStrokeWidth;
                childLine.filled = false;
                childLine.strokeColor = lineColor;
                childLine.strokeCap = StrokeCap.ROUNDENDCAP;
                childLine.strokeDashes = [0, childOffset > 0 ? childOffset : 3];
            }

            /* 親罫線（実線） / Parent line (solid) */
            var mainLine = group.pathItems.add();
            mainLine.setEntirePath([[startX, startY], [endX, endY]]);
            mainLine.stroked = true;
            mainLine.strokeWidth = preset.penetrate ? currentStrokeWidth : scaledStrokeWidth;
            mainLine.filled = false;
            mainLine.strokeColor = lineColor;
            mainLine.strokeDashes = [];
            mainLine.strokeCap = preset.roundCap ? StrokeCap.ROUNDENDCAP : StrokeCap.BUTTENDCAP;

            previewGroups.push(group);
        }
        app.redraw();
    }

    /* プレビュー削除関数 / Remove preview */
    function removePreview() {
        for (var i = previewGroups.length - 1; i >= 0; i--) {
            try {
                previewGroups[i].remove();
            } catch (e) {
                $.writeln("[VTRainMaker] Failed to remove preview group: " + e);
            }
        }
        previewGroups = [];
        app.redraw();
    }

    /* プレビュー更新 / Update preview */
    function updatePreview() {
        removePreview();
        if (!chkPreview.value) return;

        var preset = getSelectedPreset();
        var sw = unitValueToPt(parseFloat(inputStrokeWidth.text), "strokeUnits");
        var angle = parseFloat(inputAngle.text);
        var nl = parseInt(inputDensity.text, 10);
        var spacing = unitValueToPt(parseFloat(inputSpacing.text), "strokeUnits");
        var margin = parseMarginPt();
        var scale = parseScaleValue();
        if ((!preset.penetrate && (isNaN(sw) || sw <= 0)) || (!preset.penetrate && !preset.polkadot && isNaN(angle)) || isNaN(nl) || nl <= 0 || !isValidScaleValue(scale) || !isValidMarginPt(margin) || (!preset.noChild && !preset.raindrop && !preset.polkadot && (isNaN(spacing) || spacing < 0))) return;
        drawLines(sw, nl);
    }

    /* ラジオボタン切り替え時：デフォルト値を反映 / Apply preset defaults when the radio button changes */
    function onPresetChange() {
        var preset = getSelectedPreset();
        inputStrokeWidth.text = ptToUnitValue(parseFloat(preset.defaultStrokeWidth), "strokeUnits", 2);
        inputAngle.text = preset.defaultAngle || String(preset.angleDeg || 45);
        inputDensity.text = preset.defaultDensity;
        inputSpacing.text = ptToUnitValue(parseFloat(preset.defaultSpacing || "0"), "strokeUnits", 2);
        inputScale.text = preset.defaultScale || "100";
        syncScaleFromInput();
        normalizeRaindropAppearance();
        reflectEnabledUI();
        updatePreview();
    }

    /* UIの有効状態を反映 / Reflect enabled UI state */
    function reflectEnabledUI() {
        var preset = getSelectedPreset();
        var strokeEnabled = !preset.penetrate;
        var angleEnabled = !preset.penetrate && !preset.polkadot;
        var spacingEnabled = !preset.noChild && !preset.raindrop && !preset.polkadot;
        inputStrokeWidth.enabled = strokeEnabled;
        lblStroke.enabled = strokeEnabled;
        txtStrokeUnit.enabled = strokeEnabled;
        inputAngle.enabled = angleEnabled;
        lblAngle.enabled = angleEnabled;
        txtAngleUnit.enabled = angleEnabled;
        inputSpacing.enabled = spacingEnabled;
        lblSpacing.enabled = spacingEnabled;
        txtSpacingUnit.enabled = spacingEnabled;
        var raindropEnabled = !!preset.raindrop;
        raindropPanel.enabled = raindropEnabled;
        chkRaindropFill.enabled = raindropEnabled;
        chkRaindropStroke.enabled = raindropEnabled;
        rbRaindropA.enabled = raindropEnabled;
        rbRaindropB.enabled = raindropEnabled;
        rbRaindropC.enabled = raindropEnabled;
    }

    function normalizeRaindropAppearance(changedControl) {
        if (chkRaindropFill.value || chkRaindropStroke.value) return;

        if (changedControl === chkRaindropFill) {
            chkRaindropStroke.value = true;
        } else if (changedControl === chkRaindropStroke) {
            chkRaindropFill.value = true;
        } else {
            chkRaindropFill.value = true;
        }
    }

    function getRaindropAppearanceOptions() {
        var fillOn = chkRaindropFill.value;
        var strokeOn = chkRaindropStroke.value;

        if (fillOn && strokeOn) {
            if (Math.random() < 0.5) {
                fillOn = true;
                strokeOn = false;
            } else {
                fillOn = false;
                strokeOn = true;
            }
        }

        return {
            filled: fillOn,
            stroked: strokeOn,
            strokeColor: createRaindropStrokeColor(),
            strokeWidth: 2
        };
    }

    function syncScaleFromInput() {
        var v = parseScaleValue();
        if (isNaN(v)) return;
        v = clampScaleValue(v);
        if (Math.round(v) !== v) {
            inputScale.text = String(Math.round(v));
            v = Math.round(v);
        } else {
            inputScale.text = String(v);
        }
        sldScale.value = v;
    }

    function syncScaleFromSlider() {
        inputScale.text = String(Math.round(sldScale.value));
    }

    /* ダイアログボックス / Dialog box */
    var dlg = new Window('dialog', L('dialogTitle') + ' ' + SCRIPT_VERSION);
    dlg.orientation = "column";
    dlg.alignChildren = ["fill", "top"];

    var mainColumns = dlg.add("group");
    mainColumns.orientation = "row";
    mainColumns.alignChildren = ["fill", "top"];
    mainColumns.alignment = ["fill", "top"];

    var leftCol = mainColumns.add("group");
    leftCol.orientation = "column";
    leftCol.alignChildren = ["fill", "top"];
    leftCol.alignment = ["left", "fill"];

    var rightCol = mainColumns.add("group");
    rightCol.orientation = "column";
    rightCol.alignChildren = ["left", "top"];
    rightCol.alignment = ["fill", "fill"];

    /* 雨の種類ラジオボタン / Rain type radio buttons */
    var radioPanel = leftCol.add("panel", undefined, L("panelRainType"));
    radioPanel.margins = [15, 20, 15, 10];
    radioPanel.orientation = "column";
    radioPanel.alignment = "fill";
    radioPanel.alignChildren = "left";
    var rbHarusame = radioPanel.add("radiobutton", undefined, L("rainHarusame"));
    var rbSamidare = radioPanel.add("radiobutton", undefined, L("rainSamidare"));
    var rbTeppouame = radioPanel.add("radiobutton", undefined, L("rainTeppouame"));
    var rbAmatsubu = radioPanel.add("radiobutton", undefined, L("rainAmatsubu"));
    var rbMizutama = radioPanel.add("radiobutton", undefined, L("rainMizutama"));
    rbHarusame.value = true;

    var optionPanel = rightCol.add("panel", undefined, L("panelOptions"));
    optionPanel.margins = [15, 20, 15, 10];
    optionPanel.orientation = "column";
    optionPanel.alignChildren = ["left", "top"];
    optionPanel.alignment = ["fill", "top"];

    var labelWidth = 96;

    var grpScale = optionPanel.add("group");
    grpScale.alignment = "left";
    var lblScale = grpScale.add("statictext", undefined, labelText("labelScale"), { justify: "right" });
    lblScale.preferredSize = [labelWidth, -1];
    lblScale.justify = "right";
    var inputScale = grpScale.add("edittext", undefined, presets.harusame.defaultScale || "100");
    inputScale.characters = 4;
    var txtScaleUnit = grpScale.add("statictext", undefined, L("unitPercent"));

    var grpScaleSlider = optionPanel.add("group");
    grpScaleSlider.alignment = ["fill", "top"];
    var sldScale = grpScaleSlider.add("slider", undefined, 100, 25, 300);
    sldScale.preferredSize.width = 180;

    var grpStroke = optionPanel.add("group");
    grpStroke.alignment = "left";
    var lblStroke = grpStroke.add("statictext", undefined, labelText("labelStrokeWidth"), { justify: "right" });
    lblStroke.preferredSize = [labelWidth, -1];
    lblStroke.justify = "right";
    var inputStrokeWidth = grpStroke.add("edittext", undefined, ptToUnitValue(parseFloat(presets.harusame.defaultStrokeWidth), "strokeUnits", 2));
    inputStrokeWidth.characters = 6;
    var txtStrokeUnit = grpStroke.add("statictext", undefined, strokeUnitInfo.label);
    var grpAngle = optionPanel.add("group");
    grpAngle.alignment = "left";
    var lblAngle = grpAngle.add("statictext", undefined, labelText("labelAngle"), { justify: "right" });
    lblAngle.preferredSize = [labelWidth, -1];
    lblAngle.justify = "right";
    var inputAngle = grpAngle.add("edittext", undefined, presets.harusame.defaultAngle || String(presets.harusame.angleDeg || 45));
    inputAngle.characters = 6;
    var txtAngleUnit = grpAngle.add("statictext", undefined, L("unitDegree"));

    var grpDensity = optionPanel.add("group");
    grpDensity.alignment = "left";
    var lblDensity = grpDensity.add("statictext", undefined, labelText("labelDensity"), { justify: "right" });
    lblDensity.preferredSize = [labelWidth, -1];
    lblDensity.justify = "right";
    var inputDensity = grpDensity.add("edittext", undefined, presets.harusame.defaultDensity);
    inputDensity.characters = 6;
    grpDensity.add("statictext", undefined, L("unitLines"));

    var grpSpacing = optionPanel.add("group");
    grpSpacing.alignment = "left";
    var lblSpacing = grpSpacing.add("statictext", undefined, labelText("labelSpacing"), { justify: "right" });
    lblSpacing.preferredSize = [labelWidth, -1];
    lblSpacing.justify = "right";
    var inputSpacing = grpSpacing.add("edittext", undefined, ptToUnitValue(parseFloat(presets.harusame.defaultSpacing || "0"), "strokeUnits", 2));
    inputSpacing.characters = 6;
    var txtSpacingUnit = grpSpacing.add("statictext", undefined, strokeUnitInfo.label);

    var grpMargin = optionPanel.add("group");
    grpMargin.alignment = "left";
    var lblMargin = grpMargin.add("statictext", undefined, labelText("labelMargin"), { justify: "right" });
    lblMargin.preferredSize = [labelWidth, -1];
    lblMargin.justify = "right";
    var inputMargin = grpMargin.add("edittext", undefined, ptToUnitValue(20, "rulerType", 2));
    inputMargin.characters = 6;
    grpMargin.add("statictext", undefined, rulerUnitInfo.label);

    var raindropPanel = optionPanel.add("panel", undefined, L("panelRaindrop"));
    raindropPanel.orientation = "column";
    raindropPanel.alignChildren = ["left", "top"];
    raindropPanel.alignment = ["fill", "top"];
    raindropPanel.margins = [15, 20, 15, 10];

    var grpRaindropAppearance = raindropPanel.add("group");
    grpRaindropAppearance.orientation = "row";
    grpRaindropAppearance.alignChildren = ["left", "center"];

    var chkRaindropFill = grpRaindropAppearance.add("checkbox", undefined, L("chkRaindropFill"));
    chkRaindropFill.value = true;

    var chkRaindropStroke = grpRaindropAppearance.add("checkbox", undefined, L("chkRaindropStroke"));
    chkRaindropStroke.value = false;

    var grpRaindropShapes = raindropPanel.add("group");
    grpRaindropShapes.orientation = "row";
    grpRaindropShapes.alignChildren = ["left", "center"];

    var rbRaindropA = grpRaindropShapes.add("radiobutton", undefined, L("raindropShapeA"));
    var rbRaindropB = grpRaindropShapes.add("radiobutton", undefined, L("raindropShapeB"));
    var rbRaindropC = grpRaindropShapes.add("radiobutton", undefined, L("raindropShapeC"));
    rbRaindropB.value = true;

    changeValueByArrowKey(inputStrokeWidth, true, false);
    changeValueByArrowKey(inputAngle, true, false);
    changeValueByArrowKey(inputDensity, false, false);
    changeValueByArrowKey(inputSpacing, true, false);
    changeValueByArrowKey(inputMargin, true, false);
    changeValueByArrowKey(inputScale, false, false);

    /* 2カラムにしないUI / UI outside the two-column layout */
    var bottomBar = dlg.add("group");
    bottomBar.orientation = "row";
    bottomBar.alignChildren = ["left", "center"];
    bottomBar.alignment = ["fill", "top"];

    var chkPreview = bottomBar.add("checkbox", undefined, L("preview"));
    chkPreview.value = false;

    var bottomSpacer = bottomBar.add("group");
    bottomSpacer.alignment = ["fill", "fill"];
    bottomSpacer.minimumSize.width = 0;

    var btnGroup = bottomBar.add("group");
    btnGroup.orientation = "row";
    btnGroup.alignChildren = ["right", "center"];
    btnGroup.add("button", undefined, L("cancel"), { name: "cancel" });
    btnGroup.add("button", undefined, L("ok"), { name: "ok" });

    /* イベントハンドラ / Event handlers */
    rbHarusame.onClick = onPresetChange;
    rbSamidare.onClick = onPresetChange;
    rbTeppouame.onClick = onPresetChange;
    rbAmatsubu.onClick = onPresetChange;
    rbMizutama.onClick = onPresetChange;
    chkPreview.onClick = updatePreview;
    inputStrokeWidth.onChanging = updatePreview;
    inputAngle.onChanging = updatePreview;
    inputDensity.onChanging = updatePreview;
    inputSpacing.onChanging = updatePreview;
    inputMargin.onChanging = updatePreview;
    chkRaindropFill.onClick = function () {
        normalizeRaindropAppearance(chkRaindropFill);
        updatePreview();
    };
    chkRaindropStroke.onClick = function () {
        normalizeRaindropAppearance(chkRaindropStroke);
        updatePreview();
    };
    rbRaindropA.onClick = updatePreview;
    rbRaindropB.onClick = updatePreview;
    rbRaindropC.onClick = updatePreview;
    reflectEnabledUI();

    inputScale.onChanging = function () {
        syncScaleFromInput();
        updatePreview();
    };

    sldScale.onChanging = function () {
        syncScaleFromSlider();
        updatePreview();
    };

    var result = dlg.show();

    if (result !== 1) {
        removePreview();
        return;
    }

    var preset = getSelectedPreset();
    var mainStrokeWidth = preset.penetrate ? 1 : unitValueToPt(parseFloat(inputStrokeWidth.text), "strokeUnits");
    var angle = parseFloat(inputAngle.text);
    var numLines = parseInt(inputDensity.text, 10);
    var spacing = unitValueToPt(parseFloat(inputSpacing.text), "strokeUnits");
    var margin = parseMarginPt();
    var scale = parseScaleValue();
    normalizeRaindropAppearance();

    if (!preset.penetrate && (isNaN(mainStrokeWidth) || mainStrokeWidth <= 0)) {
        removePreview();
        alert(L("invalidStrokeWidth"));
        return;
    }
    if (!preset.penetrate && !preset.polkadot && isNaN(angle)) {
        removePreview();
        alert(L("invalidAngle"));
        return;
    }
    if (isNaN(numLines) || numLines <= 0) {
        removePreview();
        alert(L("invalidDensity"));
        return;
    }
    if (!isValidScaleValue(scale)) {
        removePreview();
        alert(L("invalidScale"));
        return;
    }
    if (!isValidMarginPt(margin)) {
        removePreview();
        alert(L("invalidMargin"));
        return;
    }
    if (!preset.noChild && !preset.raindrop && !preset.polkadot && (isNaN(spacing) || spacing < 0)) {
        removePreview();
        alert(L("invalidSpacing"));
        return;
    }

    /* OK → プレビューがあればそのまま確定、なければ描画 / On OK, keep the preview as final if it exists; otherwise draw */
    if (previewGroups.length === 0) {
        drawLines(mainStrokeWidth, numLines);
    }
}

drawRandomParentChildLines();