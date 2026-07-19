#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

# convert2separategradient.jsx

選択したオブジェクトのグラデーションに、指定した数の中間カラーストップを追加するスクリプトです。
通常モードでは滑らかなグラデーション、セパレートモードでは色が混ざらない縞模様グラデーションを生成します。
特色（スポットカラー）はドキュメントのカラーモードに合わせて自動変換されます。

## 主な機能

- 追加するカラーストップ数を指定（1 以上の整数）
- セパレートグラデーション（縞模様）への変換
- 特色（スポットカラー）の RGB / CMYK 自動変換
- RGB / CMYK ドキュメントに対応

## 既定の挙動

- 単一オブジェクトのみ対象
- 塗りがグラデーション以外、またはストップが 2 つ未満の場合はアラートで終了

## Overview (English)

Adds a specified number of intermediate color stops to the gradient fill of a selected object.
Smooth mode produces a continuous gradient while Separate mode produces a banded (striped) gradient.
Spot colors are automatically converted to the document color mode.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "convert2separategradient";     /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "";                             /* 更新日 / last updated */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

// =========================================
// ユーザー設定 / User Settings
// =========================================

/* 追加ストップ数の初期値 / Default number of stops to add */
var DEFAULT_STOP_COUNT = 2;

/* セパレートモードで隣接ストップの追い越しを防ぐ微小オフセット（%）/ Tiny offset to avoid stop overlap (%) */
var SEPARATE_BOUNDARY_OFFSET = 0.01;

// =========================================
// ローカライズ / Localization
// =========================================

function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

var LABELS = {
    dialog: {
        title: { ja: "グラデーションカスタム追加", en: "Add Custom Gradient Stops" }
    },
    panel: {
        settings: { ja: "設定", en: "Settings" }
    },
    label: {
        stopCount: { ja: "追加するカラーストップ数", en: "Stops to add" }
    },
    checkbox: {
        separate: { ja: "セパレートグラデーションにする", en: "Separate gradient" }
    },
    tooltip: {
        separate: {
            ja: "色が混ざらない境界（縞模様）になります。",
            en: "Colors do not blend across stops (banded look)."
        }
    },
    button: {
        cancel: { ja: "キャンセル", en: "Cancel" }
    },
    alert: {
        noDocument: {
            ja: "ドキュメントが開かれていません。",
            en: "No document is open."
        },
        noSelection: {
            ja: "オブジェクトを選択してください。",
            en: "Please select at least one object."
        },
        noGradient: {
            ja: "選択にグラデーション塗りのオブジェクトが含まれていません。",
            en: "The selection contains no objects with a gradient fill."
        },
        invalidCount: {
            ja: "追加するストップ数には 1 以上の整数を入力してください。",
            en: "Enter an integer of 1 or greater for the stop count."
        }
    }
};

/* ドット区切りパスで多言語ラベルを取得 / Resolve a localized label by dot-path */
function L(path) {
    var parts = path.split(".");
    var node = LABELS;
    for (var i = 0; i < parts.length; i++) {
        node = node && node[parts[i]];
    }
    if (node && node[lang]) return node[lang];
    if (node && node.en) return node.en;
    return path;
}

/* コロン付きラベル（日本語は全角、英語は半角）/ Label with colon (full-width JA, half-width EN) */
function labelText(path) {
    return L(path) + (lang === 'ja' ? '：' : ':');
}

// =========================================
// メイン処理 / Main
// =========================================

main();

/* スクリプトのエントリポイント / Script entry point */
function main() {
    if (app.documents.length === 0) {
        alert(L('alert.noDocument'));
        return;
    }
    var doc = app.activeDocument;
    var selection = doc.selection;

    if (!selection || selection.length === 0) {
        alert(L('alert.noSelection'));
        return;
    }

    /* 選択の中からグラデーション塗り＋ストップ 2 つ以上のオブジェクトのみ抽出 / Filter to gradient-filled objects with 2+ stops */
    var targets = collectGradientTargets(selection);
    if (targets.length === 0) {
        alert(L('alert.noGradient'));
        return;
    }

    var options = showStopDialog();
    if (!options) return;

    var useRGB = (doc.documentColorSpace === DocumentColorSpace.RGB);
    for (var i = 0; i < targets.length; i++) {
        processTarget(targets[i], options, useRGB);
    }
    app.redraw();
}

/* 選択からグラデーション塗りオブジェクトのみを抽出 / Collect gradient-filled targets from a selection */
function collectGradientTargets(selection) {
    var targets = [];
    for (var i = 0; i < selection.length; i++) {
        var item = selection[i];
        if (!item.filled) continue;
        if (item.fillColor.typename !== "GradientColor") continue;
        if (item.fillColor.gradient.gradientStops.length < 2) continue;
        targets.push(item);
    }
    return targets;
}

/* 1 オブジェクト分のストップ生成と書き戻し / Build and apply stops for a single target */
function processTarget(targetObject, options, useRGB) {
    var gradientFillColor = targetObject.fillColor;
    var gradient = gradientFillColor.gradient;

    var originalStops = snapshotStops(gradient.gradientStops, useRGB);
    var finalStops = options.useSeparate
        ? buildSeparateStops(originalStops, options.stopCount)
        : buildSmoothStops(originalStops, options.stopCount);

    finalStops.sort(byRampPoint);
    applyStops(gradient, finalStops);

    /* 塗りの再設定で表示を更新 / Reassign the fill to refresh the display */
    targetObject.fillColor = gradientFillColor;
}

// =========================================
// ダイアログ / Dialog
// =========================================

/* オプション入力ダイアログを表示し入力結果を返す / Show the options dialog and return user input */
function showStopDialog() {
    var dlg = new Window("dialog", L('dialog.title') + ' ' + SCRIPT_VERSION);
    dlg.orientation = "column";
    dlg.alignChildren = ["fill", "top"];
    dlg.spacing = 15;
    dlg.margins = 20;

    var settingsPanel = dlg.add("panel", undefined, L('panel.settings'));
    settingsPanel.orientation = "column";
    settingsPanel.alignChildren = ["left", "center"];
    settingsPanel.margins = 15;
    settingsPanel.spacing = 10;

    var stopCountGroup = settingsPanel.add("group");
    stopCountGroup.add("statictext", undefined, labelText('label.stopCount'));
    var stopCountInput = stopCountGroup.add("edittext", undefined, String(DEFAULT_STOP_COUNT));
    stopCountInput.characters = 5;

    var separateCheckbox = settingsPanel.add("checkbox", undefined, L('checkbox.separate'));
    separateCheckbox.value = false;
    separateCheckbox.helpTip = L('tooltip.separate');

    var buttonGroup = dlg.add("group");
    buttonGroup.alignment = "right";
    buttonGroup.add("button", undefined, L('button.cancel'), { name: "cancel" });
    buttonGroup.add("button", undefined, "OK", { name: "ok" });

    if (dlg.show() !== 1) return null;

    var stopCount = parseInt(stopCountInput.text, 10);
    if (isNaN(stopCount) || stopCount < 1) {
        alert(L('alert.invalidCount'));
        return null;
    }
    return { stopCount: stopCount, useSeparate: separateCheckbox.value };
}

// =========================================
// ストップ生成 / Stop Building
// =========================================

/* rampPoint 昇順比較 / Comparator by rampPoint */
function byRampPoint(a, b) {
    return a.rampPoint - b.rampPoint;
}

/* 小数第 3 位に丸める / Round to 3 decimal places */
function roundTo3(value) {
    return Math.round(value * 1000) / 1000;
}

/* 既存ストップを正規化済みカラーで複製しソートして返す / Snapshot existing stops with normalized colors, sorted */
function snapshotStops(gradientStops, useRGB) {
    var snapshot = [];
    for (var i = 0; i < gradientStops.length; i++) {
        snapshot.push({
            rampPoint: gradientStops[i].rampPoint,
            color: normalizeColor(gradientStops[i].color, useRGB),
            opacity: gradientStops[i].opacity
        });
    }
    snapshot.sort(byRampPoint);
    return snapshot;
}

/* 通常（スムーズ）モードのストップ列を構築 / Build stop list for smooth mode */
function buildSmoothStops(originalStops, addCount) {
    var result = [];
    for (var i = 0; i < originalStops.length; i++) {
        result.push(originalStops[i]);
    }
    var divisor = addCount + 1;
    for (var k = 1; k <= addCount; k++) {
        var targetPoint = roundTo3((k / divisor) * 100);
        var sample = getInterpolatedStopAt(originalStops, targetPoint);
        if (sample) {
            result.push({ rampPoint: targetPoint, color: sample.color, opacity: sample.opacity });
        }
    }
    return result;
}

/* セパレートモードのストップ列を構築 / Build stop list for separate mode */
function buildSeparateStops(originalStops, addCount) {
    var result = [];
    var segmentCount = addCount + 1;
    for (var j = 0; j < segmentCount; j++) {
        var segmentLeft = roundTo3((j / segmentCount) * 100);
        var segmentRight = roundTo3(((j + 1) / segmentCount) * 100);
        /* 元グラデーションの両端を保持できるよう 0〜100% を等分した位置からサンプル / Sample at uniform positions including both endpoints */
        var samplePoint = roundTo3((j / addCount) * 100);
        var sample = getInterpolatedStopAt(originalStops, samplePoint);
        if (!sample) continue;

        /* 隣接ストップの追い越しを避けるため左端を微小ずらし / Offset the left edge to avoid overlap */
        var adjustedLeft = (j > 0) ? segmentLeft + SEPARATE_BOUNDARY_OFFSET : segmentLeft;
        result.push({ rampPoint: adjustedLeft, color: sample.color, opacity: sample.opacity });
        result.push({ rampPoint: segmentRight, color: sample.color, opacity: sample.opacity });
    }
    return result;
}

/* 任意位置（targetPoint）における補間色と不透明度を求める / Compute interpolated color and opacity at a position */
function getInterpolatedStopAt(stops, targetPoint) {
    if (stops.length === 0) return null;
    if (targetPoint <= stops[0].rampPoint) {
        return { color: stops[0].color, opacity: stops[0].opacity };
    }
    var last = stops[stops.length - 1];
    if (targetPoint >= last.rampPoint) {
        return { color: last.color, opacity: last.opacity };
    }

    var leftStop = stops[0];
    var rightStop = last;
    for (var i = 0; i < stops.length - 1; i++) {
        if (stops[i].rampPoint <= targetPoint && stops[i + 1].rampPoint >= targetPoint) {
            leftStop = stops[i];
            rightStop = stops[i + 1];
            break;
        }
    }

    var range = rightStop.rampPoint - leftStop.rampPoint;
    var ratio = (range > 0) ? (targetPoint - leftStop.rampPoint) / range : 0;
    return {
        color: interpolateColor(leftStop.color, rightStop.color, ratio),
        opacity: leftStop.opacity + (rightStop.opacity - leftStop.opacity) * ratio
    };
}

// =========================================
// グラデーション書き戻し / Apply Stops
// =========================================

/* 計算済みストップ列をグラデーションへ反映 / Write computed stops back to the gradient */
function applyStops(gradient, finalStops) {
    var requiredCount = finalStops.length;
    while (gradient.gradientStops.length < requiredCount) {
        gradient.gradientStops.add();
    }
    while (gradient.gradientStops.length > requiredCount) {
        gradient.gradientStops[gradient.gradientStops.length - 1].remove();
    }

    /* 重なりエラー回避のため一旦等間隔に散らす / Temporarily space evenly to avoid overlap errors */
    var totalStops = gradient.gradientStops.length;
    for (var i = 0; i < totalStops; i++) {
        gradient.gradientStops[i].rampPoint = (i / (totalStops - 1)) * 100;
    }

    /* 左から順に位置・色・不透明度を書き込む / Write rampPoint, color and opacity from left to right */
    for (var i = 0; i < totalStops; i++) {
        var data = finalStops[i];
        var stop = gradient.gradientStops[i];
        stop.rampPoint = data.rampPoint;
        stop.color = createIllustratorColor(data.color);
        stop.opacity = data.opacity;
    }
}

// =========================================
// カラー変換 / Color Conversion
// =========================================

/* 任意カラーをドキュメントカラーの生データへ正規化 / Normalize any color to document raw color data */
function normalizeColor(color, useRGB) {
    if (color.typename === "SpotColor") {
        return convertSpotColor(color, useRGB);
    }
    return useRGB ? toRGBRaw(color) : toCMYKRaw(color);
}

/* スポットカラーを濃度（tint）込みで生データへ展開 / Expand spot color (with tint) to raw data */
function convertSpotColor(spotColor, useRGB) {
    var baseRaw = normalizeColor(spotColor.spot.color, useRGB);
    var tint = spotColor.tint / 100;
    if (useRGB) {
        return {
            typename: "RGBColor",
            red: 255 - (255 - baseRaw.red) * tint,
            green: 255 - (255 - baseRaw.green) * tint,
            blue: 255 - (255 - baseRaw.blue) * tint
        };
    }
    return {
        typename: "CMYKColor",
        cyan: baseRaw.cyan * tint,
        magenta: baseRaw.magenta * tint,
        yellow: baseRaw.yellow * tint,
        black: baseRaw.black * tint
    };
}

/* 任意カラーを RGB 生データへ変換 / Convert any color to RGB raw data */
function toRGBRaw(color) {
    var type = color.typename;
    if (type === "RGBColor") {
        return { typename: "RGBColor", red: color.red, green: color.green, blue: color.blue };
    }
    if (type === "CMYKColor") {
        var c = color.cyan / 100;
        var m = color.magenta / 100;
        var y = color.yellow / 100;
        var k = color.black / 100;
        return {
            typename: "RGBColor",
            red: 255 * (1 - c) * (1 - k),
            green: 255 * (1 - m) * (1 - k),
            blue: 255 * (1 - y) * (1 - k)
        };
    }
    if (type === "GrayColor") {
        var grayValue = 255 * (1 - color.gray / 100);
        return { typename: "RGBColor", red: grayValue, green: grayValue, blue: grayValue };
    }
    return { typename: "RGBColor", red: 0, green: 0, blue: 0 };
}

/* 任意カラーを CMYK 生データへ変換 / Convert any color to CMYK raw data */
function toCMYKRaw(color) {
    var type = color.typename;
    if (type === "CMYKColor") {
        return { typename: "CMYKColor", cyan: color.cyan, magenta: color.magenta, yellow: color.yellow, black: color.black };
    }
    if (type === "RGBColor") {
        var r = color.red / 255;
        var g = color.green / 255;
        var b = color.blue / 255;
        var k = 1 - Math.max(r, g, b);
        var divisor = 1 - k;
        var c = (k === 1) ? 0 : (1 - r - k) / divisor;
        var m = (k === 1) ? 0 : (1 - g - k) / divisor;
        var y = (k === 1) ? 0 : (1 - b - k) / divisor;
        return {
            typename: "CMYKColor",
            cyan: c * 100,
            magenta: m * 100,
            yellow: y * 100,
            black: k * 100
        };
    }
    if (type === "GrayColor") {
        return { typename: "CMYKColor", cyan: 0, magenta: 0, yellow: 0, black: color.gray };
    }
    return { typename: "CMYKColor", cyan: 0, magenta: 0, yellow: 0, black: 100 };
}

/* 2 色を比率（ratio）で線形補間 / Linearly interpolate between two raw colors */
function interpolateColor(fromColor, toColor, ratio) {
    if (fromColor.typename === "RGBColor") {
        return {
            typename: "RGBColor",
            red: fromColor.red + (toColor.red - fromColor.red) * ratio,
            green: fromColor.green + (toColor.green - fromColor.green) * ratio,
            blue: fromColor.blue + (toColor.blue - fromColor.blue) * ratio
        };
    }
    return {
        typename: "CMYKColor",
        cyan: fromColor.cyan + (toColor.cyan - fromColor.cyan) * ratio,
        magenta: fromColor.magenta + (toColor.magenta - fromColor.magenta) * ratio,
        yellow: fromColor.yellow + (toColor.yellow - fromColor.yellow) * ratio,
        black: fromColor.black + (toColor.black - fromColor.black) * ratio
    };
}

/* 生データから Illustrator の Color オブジェクトを生成 / Build an Illustrator Color from raw data */
function createIllustratorColor(rawColor) {
    if (rawColor.typename === "RGBColor") {
        var rgb = new RGBColor();
        rgb.red = rawColor.red;
        rgb.green = rawColor.green;
        rgb.blue = rawColor.blue;
        return rgb;
    }
    var cmyk = new CMYKColor();
    cmyk.cyan = rawColor.cyan;
    cmyk.magenta = rawColor.magenta;
    cmyk.yellow = rawColor.yellow;
    cmyk.black = rawColor.black;
    return cmyk;
}
