#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

#targetengine "MyScriptEngine"

/*

### 概要

選択したテキスト（またはオブジェクト）の背面に、見た目寸法に基づく図形を生成して配置します。
既存の背面図形があれば検出して置き換え、プレビューは［OK］時に1ステップで確定します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AddBackdrop.md

### Overview

Generates a shape sized to the visual bounds of the selected text (or objects) and places it behind them.
An existing backdrop is detected and replaced, and the undo-based preview commits in one step on OK.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AddBackdrop.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AddBackdrop";                  /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.6.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-19";                             /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AddBackdrop.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AddBackdrop.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // -----------------------------------------------------------------------------
    // Dialog state persistence (engine-global)
    // - With #targetengine, $.global persists while Illustrator is running.
    // - We store the last UI values here so they restore on next run.
    // -----------------------------------------------------------------------------
    var __GLOBAL = $.global;
    if (!__GLOBAL.__AddBackdrop_State) {
        __GLOBAL.__AddBackdrop_State = {
            _hasSaved: false,

            // Shape
            shape: 'circle', // 'circle' | 'super' | 'rect'

            // Scale
            scale: '90',
            oneChar: false,

            // Margin
            marginV: '0',
            marginH: '0',
            marginLink: true,
            marginSquare: false,
            _userSetMargin: false,

            // Round / Pill
            roundEnable: false,
            roundValue: '2',
            pill: false,
            _userSetRound: false,

            // Group
            groupWithText: true,
            exclude: false,

            // Axis
            offsetX: '0',
            offsetY: '0',

            // Kind
            kind: 'fill', // 'fill' | 'stroke'
            strokeWidth: '1',

            // Color
            colorMode: 'black', // 'text' | 'black' | 'white' | 'cmyk'
            cmykC: '0',
            cmykM: '0',
            cmykY: '0',
            cmykK: '0',

            // Opacity
            opacityApply: false,
            opacity: '60'
        };
    }
    var __STATE = __GLOBAL.__AddBackdrop_State;

    // --- UI Labels (JP/EN) ---
    function getCurrentLang() {
        return ($.locale && $.locale.indexOf('ja') === 0) ? 'ja' : 'en';
    }
    var __lang = getCurrentLang();
    var __BTN_LABELS = {
        ok: {
            ja: 'OK',
            en: 'OK'
        },
        cancel: {
            ja: 'キャンセル',
            en: 'Cancel'
        }
    };

    // ラベル定義 / Labels
    var LABELS = {
        ui: {
            dialogTitle: {
                ja: "テキストの背面に図形を作成",
                en: "Create Shape Behind Text"
            },
            optionsPanel: {
                ja: "形状",
                en: "Shape"
            },
            marginTitle: {
                ja: "マージン",
                en: "Margin"
            },
            axisPanel: {
                ja: "座標",
                en: "Axis"
            },
            scalePanel: {
                ja: "スケール",
                en: "Scale"
            },
            magnification: {
                ja: "倍率",
                en: "Scale"
            },
            groupPanel: {
                ja: "グループ",
                en: "Group"
            },
            kindPanel: {
                ja: "種別",
                en: "Kind"
            },
            colorPanel: {
                ja: "カラー",
                en: "Color"
            },
            opacityPanel: {
                ja: "不透明度",
                en: "Opacity"
            },
            roundTitle: {
                ja: "角丸",
                en: "Round"
            }
        },
        tooltip: {
            perfectCircle: { ja: "背景を正円にします。", en: "Makes the backdrop a circle." },
            superEllipse:  { ja: "背景をスーパー楕円（角の丸い四角に近い形）にします。", en: "Makes the backdrop a superellipse, between a circle and a rounded square." },
            rectangle:     { ja: "背景を長方形にします。", en: "Makes the backdrop a rectangle." },
            scale:         { ja: "文字に対する背景の大きさ（％）です。", en: "Size of the backdrop relative to the text, in percent." },
            oneChar:       { ja: "1文字ずつに背景を付けます。", en: "Draws a backdrop behind each character." },
            marginV:       { ja: "文字の上下に足す余白です。", en: "Space added above and below the text." },
            marginH:       { ja: "文字の左右に足す余白です。", en: "Space added left and right of the text." },
            marginLink:    { ja: "上下と左右の余白を同じ値にそろえます。", en: "Uses the same value for the vertical and horizontal margins." },
            marginSquare:  { ja: "背景を正方形にします。", en: "Makes the backdrop a square." },
            round:         { ja: "背景の角を丸めます。半径は右の欄で指定します。", en: "Rounds the corners of the backdrop. The field on the right sets the radius." },
            pillShape:     { ja: "左右の端を半円にして、丸いピル型にします。", en: "Rounds both ends into a pill shape." },
            groupWithText: { ja: "背景と文字を1つのグループにまとめます。", en: "Groups the backdrop with the text." },
            exclude:       { ja: "背景と文字を「中マド」にして、文字部分を抜きます。", en: "Knocks the text out of the backdrop using Exclude." },
            offset:        { ja: "背景の位置を、この値だけずらします。", en: "Nudges the backdrop by this amount." },
            fill:          { ja: "背景に塗りを付けます。", en: "Fills the backdrop." },
            stroke:        { ja: "背景に線を付けます。太さは右の欄で指定します。", en: "Strokes the backdrop. The field on the right sets the weight." },
            opacity:       { ja: "背景の不透明度（％）です。", en: "Opacity of the backdrop, in percent." },
            textColorRef:  { ja: "文字の色をそのまま背景の色に使います。", en: "Uses the text color for the backdrop." },
            black:         { ja: "背景を黒にします。", en: "Makes the backdrop black." },
            white:         { ja: "背景を白にします。", en: "Makes the backdrop white." },
            cmyk:          { ja: "背景の色をCMYKで指定します。", en: "Sets the backdrop color in CMYK." },
            cmykValue:     { ja: "この版の濃度（％）です。", en: "Ink percentage for this plate." }
        },
        option: {
            perfectCircle: {
                ja: "正円",
                en: "Circle"
            },
            superEllipse: {
                ja: "スーパー楕円",
                en: "Superellipse"
            },
            rectangle: {
                ja: "長方形",
                en: "Rectangle"
            },
            marginV: {
                ja: "上下",
                en: "V"
            },
            marginH: {
                ja: "左右",
                en: "H"
            },
            link: {
                ja: "連動",
                en: "Link"
            },
            pillShape: {
                ja: "ピル形状",
                en: "Pill shape"
            },
            groupWithText: {
                ja: "テキストとグループ化",
                en: "Group with Text"
            },
            exclude: {
                ja: "中マド処理",
                en: "Exclude"
            },
            fill: {
                ja: "塗り",
                en: "Fill"
            },
            stroke: {
                ja: "線",
                en: "Stroke"
            },
            strokeWidth: {
                ja: "線幅",
                en: "Stroke Width"
            },
            textColorRef: {
                ja: "テキストカラー",
                en: "Use Text Color"
            },
            black: {
                ja: "ブラック",
                en: "Black"
            },
            white: {
                ja: "ホワイト",
                en: "White"
            },
            cmyk: {
                ja: "CMYK",
                en: "CMYK"
            },
            oneChar: {
                ja: "1文字",
                en: "Single Character"
            },
            square: {
                ja: "正方形",
                en: "Square"
            }
        },
        message: {
            noDocument: {
                ja: "ドキュメントが開かれていません。",
                en: "No document is open."
            },
            noSelection: {
                ja: "オブジェクトを選択してください。",
                en: "Please select an object."
            },
            genericError: {
                ja: "エラーが発生しました",
                en: "An error occurred"
            },
            previewError: {
                ja: "プレビュー中にエラーが発生しました",
                en: "Preview error occurred"
            },
            targetResolveError: {
                ja: "対象オブジェクトが取得できません",
                en: "Could not resolve target object"
            },
            groupFailed: {
                ja: "グループ化に失敗しました",
                en: "Failed to group objects"
            },
            excludeFailed: {
                ja: "中マド処理（Exclude）の適用に失敗しました",
                en: "Failed to apply Pathfinder Exclude"
            }
        }
    };

    function getLabel(key) {
        try {
            var parts = key.split('.');
            var t = LABELS;
            for (var i = 0; i < parts.length; i++) {
                t = t[parts[i]];
                if (!t) return key;
            }
            return (t && (t[__lang] || t.ja || t.en)) || key;
        } catch (e) {
            return key;
        }
    }

    function safeRemove(item, label) {
        try {
            if (item && item.remove) item.remove();
        } catch (e) {
            try { $.writeln('[AddBackdrop] remove failed' + (label ? ' (' + label + ')' : '') + ': ' + e); } catch (e) { }
        }
    }

    function clearSelection() {
        try {
            app.selection = null;
        } catch (e) {
            try { $.writeln('[AddBackdrop] clearSelection failed: ' + e); } catch (e) { }
        }
    }

    // --- Geometry helpers ---
    function sign(x) {
        return ((x > 0) - (x < 0)) || +x;
    }

    /**
     * スーパー楕円（Superellipse）のアンカーポイント配列を生成（固定点数サンプリング）
     * - 旧: t を細かく刻んで大量ポイント
     * - 新: 少ない点数 + スムーズハンドル付与で軽量化
     */
    function buildSuperellipseAnchorPoints(cx, cy, width, height, n, numPoints) {
        var pts = [];
        var TWO_PI = Math.PI * 2;
        var N = (numPoints && numPoints > 3) ? Math.round(numPoints) : 32;

        for (var i = 0; i < N; i++) {
            var theta = (TWO_PI * i) / N;
            var ct = Math.cos(theta);
            var st = Math.sin(theta);
            var x = Math.pow(Math.abs(ct), 2 / n) * (width / 2) * sign(ct);
            var y = Math.pow(Math.abs(st), 2 / n) * (height / 2) * sign(st);
            pts.push([cx + x, cy + y]);
        }
        return pts;
    }

    /**
     * PathItem をスーパー楕円形状に変形し、スムーズなハンドルを付与
     */
    function morphPathToSuperellipse(pathItem, cx, cy, width, height, n) {
        if (!pathItem) return;

        // 扱いやすさと滑らかさのバランス（必要なら後でUI化）
        var NUM_POINTS = 8;

        var anchorPoints = buildSuperellipseAnchorPoints(cx, cy, width, height, n, NUM_POINTS);
        pathItem.setEntirePath(anchorPoints);
        pathItem.closed = true;

        // 角を減らすために各ポイントをスムーズ化し、接線方向からハンドルを推定
        try {
            var pp = pathItem.pathPoints;
            var len = pp.length;
            if (len >= 4) {
                for (var i = 0; i < len; i++) {
                    var prev = anchorPoints[(i - 1 + len) % len];
                    var cur = anchorPoints[i];
                    var next = anchorPoints[(i + 1) % len];

                    // 接線（next - prev）
                    var tx = next[0] - prev[0];
                    var ty = next[1] - prev[1];
                    var tlen = Math.sqrt(tx * tx + ty * ty);
                    if (tlen === 0) continue;
                    tx /= tlen;
                    ty /= tlen;

                    // 前後セグメント長
                    var d1x = cur[0] - prev[0];
                    var d1y = cur[1] - prev[1];
                    var d2x = next[0] - cur[0];
                    var d2y = next[1] - cur[1];
                    var d1 = Math.sqrt(d1x * d1x + d1y * d1y);
                    var d2 = Math.sqrt(d2x * d2x + d2y * d2y);

                    // ハンドル長係数（小さいほど角ばる／大きいほど丸くなる）
                    var h = Math.min(d1, d2) * 0.35;

                    var left = [cur[0] - tx * h, cur[1] - ty * h];
                    var right = [cur[0] + tx * h, cur[1] + ty * h];

                    pp[i].anchor = cur;
                    pp[i].leftDirection = left;
                    pp[i].rightDirection = right;
                    pp[i].pointType = PointType.SMOOTH;
                }
            }
        } catch (e) {
            // ignore (環境差など)
        }
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

    /* 単位コード5を「歯（H）」と表示する環境設定キー。文字サイズ（text/units）だけ「級（Q）」
       Preference keys that show unit code 5 as H; only the type size (text/units) shows Q */
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

    // --- Dialog visual adjustments (opacity & initial shift) ---
    // 初期シフト量と不透明度
    var __DIALOG_OFFSET_X = 300;
    var __DIALOG_OFFSET_Y = 0;
    var __DIALOG_OPACITY = 0.98;

    // 透明度設定（tryガード付き）
    function setDialogOpacity(dlg, opacityValue) {
        try {
            dlg.opacity = opacityValue;
        } catch (e) { }
    }

    /**
     * 初回表示時だけ位置をずらすヘルパ
     * すでに保存済みのboundsがある場合は、ズレ続けを避けるためシフトしない
     */
    function shiftDialogPositionOnceOnShow(dlg, dx, dy) {
        dlg.onShow = (function (prev) {
            return function () {
                try {
                    if (typeof prev === 'function') prev();
                } catch (e) { }
                try {
                    var l = dlg.location;
                    dlg.location = [l[0] + (dx | 0), l[1] + (dy | 0)];
                } catch (e) { }
            };
        })(dlg.onShow);
    }

    // --- Arrow-key value changer for EditText ---
    function changeValueByArrowKey(editText, onUpdate) {
        if (!editText) return;

        // フォールバック：値が変わったら常にプレビュー更新（単純入力やスクロールでも反応）
        editText.onChanging = function () {
            try {
                if (typeof onUpdate === 'function') onUpdate();
            } catch (e) { }
        };

        // ScriptUI の addEventListener が使える場合は、矢印キー専用の高速インクリメント処理
        if (editText.addEventListener) {
            editText.addEventListener('keydown', function (event) {
                var keyboard = ScriptUI.environment.keyboardState;
                var key = event && event.keyName;
                if (key !== 'Up' && key !== 'Down') return; // 矢印以外は無視

                var value = Number(editText.text);
                if (isNaN(value)) return;

                var delta;
                if (keyboard.shiftKey) {
                    // Shift: 10刻み（10の倍数にスナップ）
                    delta = 10;
                    if (key === 'Up') {
                        value = Math.ceil((value + 1) / delta) * delta;
                    } else {
                        value = Math.floor((value - 1) / delta) * delta;
                    }
                } else if (keyboard.altKey) {
                    // Option: 0.1刻み
                    delta = 0.1;
                    value += (key === 'Up') ? delta : -delta;
                } else {
                    // 通常: 1刻み
                    delta = 1;
                    value += (key === 'Up') ? delta : -delta;
                }

                // 丸め：Option時は小数1位、それ以外は整数
                if (keyboard.altKey) {
                    value = Math.round(value * 10) / 10;
                } else {
                    value = Math.round(value);
                }

                // 最小値クランプ（必要なフィールドのみ）
                // 例: strokeWInput.__minValue = 0;
                try {
                    if (editText && editText.__minValue !== undefined && editText.__minValue !== null) {
                        var minV = Number(editText.__minValue);
                        if (!isNaN(minV) && value < minV) value = minV;
                    }
                } catch (e) { }
                // 最大値クランプ（必要なフィールドのみ）
                try {
                    if (editText && editText.__maxValue !== undefined && editText.__maxValue !== null) {
                        var maxV = Number(editText.__maxValue);
                        if (!isNaN(maxV) && value > maxV) value = maxV;
                    }
                } catch (e) { }

                editText.text = value;

                // キー操作でも即時プレビュー
                try {
                    if (typeof onUpdate === 'function') onUpdate();
                } catch (e) { }

                // 既定動作の抑止（可能な環境のみ）
                try {
                    if (event && event.preventDefault) event.preventDefault();
                } catch (e) { }
            });
        }
    }

    // 入力フィールドにプレビュー更新を一括でバインド / Bind preview updates to an input
    function bindPreview(editText, onUpdate) {
        if (!editText) return;

        // 既存の onChanging を尊重しつつ、プレビュー更新も追加
        var _prevOnChanging = null;
        try { _prevOnChanging = editText.onChanging; } catch (e) { _prevOnChanging = null; }

        editText.onChanging = function () {
            try { if (typeof _prevOnChanging === 'function') _prevOnChanging(); } catch (e) { }
            try { if (typeof onUpdate === 'function') onUpdate(); } catch (e) { }
        };

        // 矢印キー（↑↓ / Shift / Option）も同じ処理
        changeValueByArrowKey(editText, onUpdate);
    }

    // キー入力で形状ラジオを切り替え（E: 正円 / S: スーパー楕円 / R: 長方形）
    function addShapeKeyHandler(dialog, rbCircle, rbSuper, rbRect, onChange) {
        if (!dialog || !dialog.addEventListener) return;
        dialog.addEventListener('keydown', function (event) {
            try {
                var k = event && event.keyName;
                if (!k) return;
                if (k === 'E') {
                    rbCircle.value = true;
                } else if (k === 'S') {
                    rbSuper.value = true;
                } else if (k === 'R') {
                    rbRect.value = true;
                } else {
                    return;
                }
                try { if (typeof onChange === 'function') onChange(); } catch (e) { }
                try { if (event.preventDefault) event.preventDefault(); } catch (e) { }
            } catch (e) { }
        });
    }

    /* =============================================================================
     * PreviewManager (Undo-safe preview)
     * - preview updates do NOT leave multiple undo steps
     * - Cancel: rollback to original
     * - OK: rollback then re-run once (single undo step)
     * ========================================================================== */
    function PreviewManager() {
        this.undoDepth = 0;

        this.addStep = function (func) {
            try {
                func();
                this.undoDepth++;
                app.redraw();
            } catch (e) {
                alert(getLabel('message.previewError') + ": " + e);
            }
        };

        this.rollback = function () {
            try {
                while (this.undoDepth > 0) {
                    try {
                        app.undo();
                    } catch (eUndo) {
                        try { $.writeln('[AddBackdrop] rollback undo failed: ' + eUndo); } catch (e) { }
                        break;
                    }
                    this.undoDepth--;
                }
                try {
                    app.redraw();
                } catch (eRedraw) {
                    try { $.writeln('[AddBackdrop] rollback redraw failed: ' + eRedraw); } catch (e) { }
                }
            } catch (eRollback) {
                try { $.writeln('[AddBackdrop] rollback failed: ' + eRollback); } catch (e) { }
            }
        };

        this.confirm = function (finalAction) {
            if (typeof finalAction === 'function') {
                this.rollback();
                finalAction();
                // finalAction is a single step; do not count it as preview
                this.undoDepth = 0;
            } else {
                // adopt preview as-is (not used in this script)
                this.undoDepth = 0;
            }
        };
    }

    /* =============================================================================
     * main entry
     * ========================================================================== */
    function main() {
        // Preview/Undo manager
        var previewMgr = new PreviewManager();
        // ダイアログ作成
        var dlg = new Window('dialog', getLabel('ui.dialogTitle') + ' ' + SCRIPT_VERSION);
        setDialogOpacity(dlg, __DIALOG_OPACITY);
        shiftDialogPositionOnceOnShow(dlg, __DIALOG_OFFSET_X, __DIALOG_OFFSET_Y);

        dlg.orientation = 'column';
        dlg.alignChildren = ['fill', 'top'];

        // 形状（ダイアログ上部・2カラムから独立）
        var grpShapeTop = dlg.add('group');
        grpShapeTop.orientation = 'row';
        grpShapeTop.alignChildren = ['center', 'center'];
        grpShapeTop.alignment = 'center';

        var rbPerfectCircle = grpShapeTop.add('radiobutton', undefined, getLabel('option.perfectCircle'));
        rbPerfectCircle.helpTip = getLabel('tooltip.perfectCircle');
        var rbSuperEllipse = grpShapeTop.add('radiobutton', undefined, getLabel('option.superEllipse'));
        rbSuperEllipse.helpTip = getLabel('tooltip.superEllipse');
        var rbRectangle = grpShapeTop.add('radiobutton', undefined, getLabel('option.rectangle'));
        rbRectangle.helpTip = getLabel('tooltip.rectangle');

        rbPerfectCircle.value = true; // デフォルトは正円

        // どのオプションを選んでもプレビュー更新
        var _onOptionChanged = function () {
            try { if (typeof syncRoundPanelUI === 'function') syncRoundPanelUI(); } catch (e) { }
            try { if (typeof syncMarginPanelUI === 'function') syncMarginPanelUI(); } catch (e) { }
            try { if (typeof syncScalePanelUI === 'function') syncScalePanelUI(); } catch (e) { }
            if (typeof updatePreview === 'function') updatePreview();
        };
        rbPerfectCircle.onClick = _onOptionChanged;
        rbSuperEllipse.onClick = _onOptionChanged;
        rbRectangle.onClick = _onOptionChanged;

        rbPerfectCircle.onChanging = _onOptionChanged;
        rbSuperEllipse.onChanging = _onOptionChanged;
        rbRectangle.onChanging = _onOptionChanged;

        // キー操作で形状切替（E/S/R）
        addShapeKeyHandler(dlg, rbPerfectCircle, rbSuperEllipse, rbRectangle, _onOptionChanged);

        // 現在のアプリ環境の単位ラベルを取得
        var rulerUnitLabel = getUnitInfo('rulerType').label;
        var strokeUnitLabel = getUnitInfo('strokeUnits').label;
        // 2カラムコンテナ

        // --- マージン正方形チェックボックスを先に宣言してスコープを広げる ---
        var cbMarginSquare = null; // will be assigned in the margin panel

        // 2カラムコンテナ
        var cols = dlg.add('group');
        cols.orientation = 'row';
        cols.alignChildren = ['fill', 'top'];

        var leftCol = cols.add('group');
        leftCol.orientation = 'column';
        leftCol.alignChildren = ['fill', 'top'];

        var rightCol = cols.add('group');
        rightCol.orientation = 'column';
        rightCol.alignChildren = ['fill', 'top'];

        // パネル：スケール（先頭へ移動）
        var pnlScale = leftCol.add('panel', undefined, getLabel('ui.scalePanel'));
        pnlScale.orientation = 'column';
        pnlScale.alignChildren = ['fill', 'top'];
        pnlScale.margins = [15, 20, 15, 10];

        // 倍率（%表示）
        var grpScale = pnlScale.add('group');
        grpScale.add('statictext', undefined, getLabel('ui.magnification'));
        var scaleInput = grpScale.add('edittext', undefined, '90');
        scaleInput.helpTip = getLabel('tooltip.scale');
        scaleInput.characters = 4;
        grpScale.add('statictext', undefined, '%');

        // 1文字チェック（倍率の直下）
        var grpOneChar = pnlScale.add('group');
        grpOneChar.orientation = 'row';
        grpOneChar.alignChildren = ['left', 'center'];
        var cbOneChar = grpOneChar.add('checkbox', undefined, getLabel('option.oneChar'));
        cbOneChar.helpTip = getLabel('tooltip.oneChar');
        cbOneChar.value = false;

        // スケールパネルは「長方形」選択時は 100% 固定 + ディム表示（ただし正方形ON時は例外で有効）
        function syncScalePanelUI() {
            try {
                var squareOn = false;
                try { squareOn = !!(cbMarginSquare && cbMarginSquare.value); } catch (e) { squareOn = false; }

                // 長方形選択時は通常 100%固定 + ディム。正方形ONのときだけ例外で有効化。
                if (rbRectangle && rbRectangle.value && !squareOn) {
                    // 強制 100%
                    scaleInput.text = '100';
                    // ついでに 1文字は意味が薄いのでOFF（パネル全体がディムなので安全側）
                    cbOneChar.value = false;

                    pnlScale.enabled = false;
                } else {
                    pnlScale.enabled = true;
                }
            } catch (e) { }
        }
        syncScalePanelUI();

        // パネル：マージン（形状パネルの直下）
        var marginPanel = leftCol.add('panel', undefined, getLabel('ui.marginTitle'));
        marginPanel.orientation = 'column';
        marginPanel.alignChildren = ['fill', 'top'];
        marginPanel.margins = [15, 20, 15, 10];

        // LEFTCOL container (row)
        var marginLeftCol = marginPanel.add('group');
        marginLeftCol.orientation = 'row';
        marginLeftCol.alignChildren = ['left', 'top'];
        marginLeftCol.spacing = 10;
        marginLeftCol.margins = 0;

        // GROUP1 (column for two rows: 上下 / 左右)
        var marginGroup1 = marginLeftCol.add('group');
        marginGroup1.orientation = 'column';
        marginGroup1.alignChildren = ['left', 'center'];
        marginGroup1.spacing = 10;
        marginGroup1.margins = 0;

        // GROUP2 (row: 上下)
        var groupMV = marginGroup1.add('group');
        groupMV.orientation = 'row';
        groupMV.alignChildren = ['left', 'center'];
        groupMV.spacing = 10;
        groupMV.margins = 0;

        groupMV.add('statictext', undefined, getLabel('option.marginV'));
        var marginVInput = groupMV.add('edittext', undefined, '0');
        marginVInput.helpTip = getLabel('tooltip.marginV');
        marginVInput.characters = 4;
        groupMV.add('statictext', undefined, rulerUnitLabel);

        // GROUP3 (row: 左右)
        var groupMH = marginGroup1.add('group');
        groupMH.orientation = 'row';
        groupMH.alignChildren = ['left', 'center'];
        groupMH.spacing = 10;
        groupMH.margins = 0;

        groupMH.add('statictext', undefined, getLabel('option.marginH'));
        var marginHInput = groupMH.add('edittext', undefined, '0');
        marginHInput.helpTip = getLabel('tooltip.marginH');
        marginHInput.characters = 4;
        groupMH.add('statictext', undefined, rulerUnitLabel);

        // RIGHT side: Link checkbox (center-ish)
        var marginRightCol = marginLeftCol.add('group');
        marginRightCol.orientation = 'column';
        marginRightCol.alignChildren = ['left', 'center'];
        marginRightCol.spacing = 0;
        marginRightCol.margins = 0;

        // small spacer to align checkbox between two rows
        try { marginRightCol.add('statictext', undefined, ''); } catch (e) { }

        var cbMarginLink = marginRightCol.add('checkbox', undefined, getLabel('option.link'));
        cbMarginLink.helpTip = getLabel('tooltip.marginLink');
        cbMarginLink.value = true; // デフォルトは連動

        // マージン：正方形
        cbMarginSquare = marginPanel.add('checkbox', undefined, getLabel('option.square'));
        cbMarginSquare.helpTip = getLabel('tooltip.marginSquare');
        cbMarginSquare.value = false; // デフォルトOFF

        cbMarginSquare.onClick = function () {
            // 正方形ON: マージン無視 + スケールを有効化しアクティブに
            // 正方形ON → ピル形状OFF
            try { if (cbMarginSquare.value && cbPill) { cbPill.value = false; syncRoundInputsUI(); } } catch (e) { }
            try { if (typeof syncScalePanelUI === 'function') syncScalePanelUI(); } catch (e) { }
            try { if (cbMarginSquare.value) { scaleInput.active = true; } } catch (e) { }
            if (typeof updatePreview === 'function') updatePreview();
        };
        cbMarginSquare.onChanging = cbMarginSquare.onClick;

        function syncMarginUI() {
            var linked = !!cbMarginLink.value;
            try {
                if (linked) {
                    marginHInput.text = String(marginVInput.text);
                    marginHInput.enabled = false;
                } else {
                    marginHInput.enabled = true;
                }
            } catch (e) { }
        }
        syncMarginUI();

        cbMarginLink.onClick = function () {
            syncMarginUI();
            if (typeof updatePreview === 'function') updatePreview();
        };
        cbMarginLink.onChanging = cbMarginLink.onClick;

        // 入力変更で常に同期 + プレビュー
        marginVInput.onChanging = function () {
            try { syncMarginUI(); } catch (e) { }
            if (typeof updatePreview === 'function') updatePreview();
        };
        marginHInput.onChanging = function () {
            if (typeof updatePreview === 'function') updatePreview();
        };

        // 矢印キー増減（↑↓ / Shift / Option）
        changeValueByArrowKey(marginVInput, function () {
            try { syncMarginUI(); } catch (e) { }
            if (typeof updatePreview === 'function') updatePreview();
        });
        changeValueByArrowKey(marginHInput, function () {
            if (typeof updatePreview === 'function') updatePreview();
        });

        // マージンパネルは「長方形」選択時のみ有効（正円/スーパー楕円ではディム表示）
        function syncMarginPanelUI() {
            try {
                marginPanel.enabled = !!rbRectangle.value;
            } catch (e) { }
        }
        syncMarginPanelUI();

        function getMarginValues() {
            var mv = parseFloat(marginVInput && marginVInput.text);
            var mh = parseFloat(marginHInput && marginHInput.text);
            if (isNaN(mv)) mv = 0;
            if (isNaN(mh)) mh = 0;
            // 連動時は左右=上下
            try {
                if (cbMarginLink && cbMarginLink.value) mh = mv;
            } catch (e) { }
            return { mx: mh, my: mv };
        }

        // パネル：角丸（形状パネルの直下）
        var roundPanel = leftCol.add('panel', undefined, getLabel('ui.roundTitle'));
        roundPanel.orientation = 'column';
        roundPanel.alignChildren = ['left', 'top'];
        roundPanel.margins = [15, 20, 15, 10];

        var roundRow = roundPanel.add('group');
        roundRow.orientation = 'row';
        roundRow.alignChildren = ['left', 'center'];
        // Checkbox before the numeric field (UI only for now)
        var cbRoundEnable = roundRow.add('checkbox', undefined, '');
        cbRoundEnable.helpTip = getLabel('tooltip.round');
        var __lastRoundValue = '2'; // will be updated after roundInput is created

        var roundInput = roundRow.add('edittext', undefined, '2');
        roundInput.helpTip = getLabel('tooltip.round');
        roundInput.characters = 4;
        __lastRoundValue = String(roundInput.text);
        roundRow.add('statictext', undefined, rulerUnitLabel);

        // --- Pill shape option: now on its own row ---
        var pillRow = roundPanel.add('group');
        pillRow.orientation = 'row';
        pillRow.alignChildren = ['left', 'center'];
        var cbPill = pillRow.add('checkbox', undefined, getLabel('option.pillShape'));
        cbPill.helpTip = getLabel('tooltip.pillShape');
        cbPill.value = false;

        // 角丸パネルは「長方形」選択時のみ有効（正円/スーパー楕円ではディム表示）
        function syncRoundPanelUI() {
            try {
                roundPanel.enabled = !!rbRectangle.value;
            } catch (e) { }
            syncRoundInputsUI();
        }

        function syncRoundInputsUI() {
            try {
                var pillOn = !!cbPill.value;
                var roundOn = !!cbRoundEnable.value;

                // pill ON のときは半径が自動なので入力は無効
                if (pillOn) {
                    roundInput.enabled = false;
                    return;
                }

                if (roundOn) {
                    // OFF→ON で値を復元
                    if (!roundInput.enabled) {
                        roundInput.text = String(__lastRoundValue || roundInput.text || '0');
                    }
                    roundInput.enabled = true;
                } else {
                    // ON→OFF で値を保存して無効化
                    try { __lastRoundValue = String(roundInput.text); } catch (e) { }
                    roundInput.enabled = false;
                }
            } catch (e) { }
        }

        // 初期状態
        try { syncRoundPanelUI(); } catch (e) { }

        // イベント：角丸ON/OFF
        cbRoundEnable.onClick = function () {
            try { syncRoundInputsUI(); } catch (e) { }
            try {
                if (cbRoundEnable.value && !cbPill.value) {
                    roundInput.active = true;
                }
            } catch (e) { }
            if (typeof updatePreview === 'function') updatePreview();
        };
        cbRoundEnable.onChanging = cbRoundEnable.onClick;

        // イベント：ピル
        cbPill.onClick = function () {
            // pill ON のときは入力を無効化（半径は自動）
            try { syncRoundInputsUI(); } catch (e) { }
            if (typeof updatePreview === 'function') updatePreview();
        };
        cbPill.onChanging = cbPill.onClick;

        // 入力：矢印キー増減 + プレビュー
        changeValueByArrowKey(roundInput, function () {
            try { __lastRoundValue = String(roundInput.text); } catch (e) { }
            if (typeof updatePreview === 'function') updatePreview();
        });

        // パネル：グループ（スケールの直下）
        var pnlGroup = leftCol.add('panel', undefined, getLabel('ui.groupPanel'));
        pnlGroup.orientation = 'column';
        pnlGroup.alignChildren = ['fill', 'top'];
        pnlGroup.margins = [15, 20, 15, 10];
        var cbGroup = pnlGroup.add('checkbox', undefined, getLabel('option.groupWithText'));
        cbGroup.helpTip = getLabel('tooltip.groupWithText');
        cbGroup.value = true;

        var cbExclude = pnlGroup.add('checkbox', undefined, getLabel('option.exclude'));
        cbExclude.helpTip = getLabel('tooltip.exclude');
        cbExclude.value = false;

        // ロジック：中マド=ONなら自動でグループ化し、カラーを「テキストカラー」に強制
        cbExclude.onClick = function () {
            if (cbExclude.value) {
                // 中マド=ONなら自動でグループ化し、カラーを「テキストカラー」に強制
                cbGroup.value = true;
                try {
                    rbTextColor.value = true;
                    rbBlack.value = false;
                    rbWhite.value = false;
                    rbCMYK.value = false;
                    if (typeof syncColorUI === 'function') syncColorUI();
                } catch (e) { }
            }
            if (typeof updatePreview === 'function') updatePreview();
        };
        cbExclude.onChanging = cbExclude.onClick;

        // ロジック：「テキストとグループ化」をOFFにしたら「中マド処理」もOFF
        cbGroup.onClick = function () {
            if (!cbGroup.value) cbExclude.value = false;
            if (typeof updatePreview === 'function') updatePreview();
        };
        cbGroup.onChanging = cbGroup.onClick;

        // パネル：座標調整（右カラムへ移動）
        var pnlOffset = rightCol.add('panel', undefined, getLabel('ui.axisPanel') + '（' + rulerUnitLabel + '）');
        pnlOffset.orientation = 'column';
        pnlOffset.alignChildren = ['fill', 'top'];
        pnlOffset.margins = [15, 20, 15, 10];

        // X と Y を横並びに
        var grpOffset = pnlOffset.add('group');
        grpOffset.orientation = 'row';
        grpOffset.alignChildren = ['left', 'center'];

        grpOffset.add('statictext', undefined, 'X');
        var offsetXInput = grpOffset.add('edittext', undefined, '0');
        offsetXInput.helpTip = getLabel('tooltip.offset');
        offsetXInput.characters = 4;

        grpOffset.add('statictext', undefined, 'Y');
        var offsetYInput = grpOffset.add('edittext', undefined, '0');
        offsetYInput.helpTip = getLabel('tooltip.offset');
        offsetYInput.characters = 4;

        // パネル：種別（塗り / 線 / 線幅を1行に） — 右カラム最下部へ移動
        var pnlKind = rightCol.add('panel', undefined, getLabel('ui.kindPanel'));
        pnlKind.orientation = 'column';
        pnlKind.alignChildren = ['fill', 'top'];
        pnlKind.margins = [15, 20, 15, 10];

        // 1行：塗り / 線 / 線幅
        var grpKind = pnlKind.add('group');
        grpKind.orientation = 'row';
        grpKind.alignChildren = ['left', 'center'];
        grpKind.spacing = 10;

        var rbFill = grpKind.add('radiobutton', undefined, getLabel('option.fill'));
        rbFill.helpTip = getLabel('tooltip.fill');
        var rbStroke = grpKind.add('radiobutton', undefined, getLabel('option.stroke'));
        rbStroke.helpTip = getLabel('tooltip.stroke');
        rbFill.value = true; // デフォルトは塗り

        var strokeWInput = grpKind.add('edittext', undefined, '1');
        strokeWInput.helpTip = getLabel('tooltip.stroke');
        strokeWInput.characters = 4;
        grpKind.add('statictext', undefined, strokeUnitLabel);

        // 線幅は負の値を許容しない
        // bindPreview() が onChanging をラップするので、ここではクランプだけ行う
        strokeWInput.__minValue = 0;
        strokeWInput.onChanging = function () {
            var v = parseFloat(strokeWInput.text);
            if (isNaN(v)) return;
            if (v < 0) strokeWInput.text = '0';
        };

        function syncKindUI() {
            var strokeOn = rbStroke.value;
            // 線のときのみ線幅を有効
            strokeWInput.enabled = strokeOn;
            // 線のときは中マド処理をディム＆OFF
            try {
                if (strokeOn) {
                    cbExclude.value = false;
                    cbExclude.enabled = false;
                } else {
                    cbExclude.enabled = true;
                }
            } catch (e) { }
        }
        syncKindUI();
        rbFill.onClick = function () {
            syncKindUI();
            if (typeof updatePreview === 'function') updatePreview();
        };
        rbStroke.onClick = rbFill.onClick;
        rbFill.onChanging = rbFill.onClick;
        rbStroke.onChanging = rbFill.onClick;

        // パネル：カラー（右カラム）
        var pnlColor = rightCol.add('panel', undefined, getLabel('ui.colorPanel'));
        pnlColor.orientation = 'column';
        pnlColor.alignChildren = ['fill', 'top'];
        pnlColor.margins = [15, 20, 15, 10];

        // パネル：不透明度（右カラム）
        var pnlOpacity = rightCol.add('panel', undefined, getLabel('ui.opacityPanel'));
        pnlOpacity.orientation = 'column';
        pnlOpacity.alignChildren = ['fill', 'top'];
        pnlOpacity.margins = [15, 20, 15, 10];

        var grpOpacity = pnlOpacity.add('group');
        grpOpacity.orientation = 'row';
        var cbOpacityApply = grpOpacity.add('checkbox', undefined, '');
        cbOpacityApply.helpTip = getLabel('tooltip.opacity');
        cbOpacityApply.value = false;

        var opacityInput = grpOpacity.add('edittext', undefined, '60');
        opacityInput.helpTip = getLabel('tooltip.opacity');
        opacityInput.characters = 3; // 0-100
        opacityInput.enabled = cbOpacityApply.value;
        grpOpacity.add('statictext', undefined, '%');

        // 色モード（ラジオ）
        var grpMode = pnlColor.add('group');
        grpMode.orientation = 'column';
        grpMode.alignChildren = ['left', 'top'];

        var rbTextColor = grpMode.add('radiobutton', undefined, getLabel('option.textColorRef'));
        rbTextColor.helpTip = getLabel('tooltip.textColorRef');

        var rbBlack = grpMode.add('radiobutton', undefined, getLabel('option.black'));
        rbBlack.helpTip = getLabel('tooltip.black');
        var rbWhite = grpMode.add('radiobutton', undefined, getLabel('option.white'));
        rbWhite.helpTip = getLabel('tooltip.white');
        var rbCMYK = grpMode.add('radiobutton', undefined, getLabel('option.cmyk'));
        rbCMYK.helpTip = getLabel('tooltip.cmyk');

        // デフォルトはブラック
        rbBlack.value = true;
        rbTextColor.value = rbWhite.value = rbCMYK.value = false;

        // (RGB UI removed)

        // CMYK 入力（各チャンネルを縦積み：ラベル→入力）
        var grpCMYK = pnlColor.add('group');
        grpCMYK.orientation = 'row';
        grpCMYK.alignChildren = ['center', 'top'];

        function addCMYKColumn(parent, label) {
            var col = parent.add('group');
            col.orientation = 'column';
            col.alignChildren = ['fill', 'top'];
            var st = col.add('statictext', undefined, label);
            st.justify = 'center';
            var et = col.add('edittext', undefined, '0');
            et.helpTip = getLabel('tooltip.cmykValue');
            et.characters = 4;
            return et;
        }

        var fillC = addCMYKColumn(grpCMYK, 'C');
        var fillM = addCMYKColumn(grpCMYK, 'M');
        var fillY = addCMYKColumn(grpCMYK, 'Y');
        var fillK = addCMYKColumn(grpCMYK, 'K');

        // Clamp CMYK inputs to 0–100
        function clampCMYKInput(et) {
            if (!et) return;
            et.__minValue = 0;
            et.__maxValue = 100;
            et.onChanging = function () {
                var v = parseFloat(et.text);
                if (isNaN(v)) return;
                if (v < 0) et.text = '0';
                else if (v > 100) et.text = '100';
            };
        }

        clampCMYKInput(fillC);
        clampCMYKInput(fillM);
        clampCMYKInput(fillY);
        clampCMYKInput(fillK);

        function syncColorUI() {
            var cmykOn = rbCMYK.value;
            // 個別フィールドの有効/無効
            fillC.enabled = cmykOn;
            fillM.enabled = cmykOn;
            fillY.enabled = cmykOn;
            fillK.enabled = cmykOn;
            // 非表示にはせず、CMYK以外のときはグループごとディム表示
            try {
                grpCMYK.visible = true; // 常に表示
                grpCMYK.enabled = cmykOn; // CMYK選択時のみ操作可（ラベルも含めてディム制御）
            } catch (e) { }
        }
        syncColorUI();

        // -------------------------------------------------------------------------
        // Persist / restore UI state
        // -------------------------------------------------------------------------
        function saveUIStateFromControls() {
            try {
                __STATE.shape = (rbRectangle.value ? 'rect' : (rbSuperEllipse.value ? 'super' : 'circle'));
            } catch (e) { }

            // Scale
            try { __STATE.scale = String(scaleInput.text); } catch (e) { }
            try { __STATE.oneChar = !!cbOneChar.value; } catch (e) { }

            // Margin
            try { __STATE.marginV = String(marginVInput.text); } catch (e) { }
            try { __STATE.marginH = String(marginHInput.text); } catch (e) { }
            try { __STATE.marginLink = !!cbMarginLink.value; } catch (e) { }
            try { __STATE.marginSquare = !!(cbMarginSquare && cbMarginSquare.value); } catch (e) { }

            // Round / Pill
            try { __STATE.roundEnable = !!cbRoundEnable.value; } catch (e) { }
            try { __STATE.roundValue = String(roundInput.text); } catch (e) { }
            try { __STATE.pill = !!cbPill.value; } catch (e) { }

            // Group
            try { __STATE.groupWithText = !!cbGroup.value; } catch (e) { }
            try { __STATE.exclude = !!cbExclude.value; } catch (e) { }

            // Axis
            try { __STATE.offsetX = String(offsetXInput.text); } catch (e) { }
            try { __STATE.offsetY = String(offsetYInput.text); } catch (e) { }

            // Kind
            try { __STATE.kind = (rbStroke.value ? 'stroke' : 'fill'); } catch (e) { }
            try { __STATE.strokeWidth = String(strokeWInput.text); } catch (e) { }

            // Color
            try {
                __STATE.colorMode = (rbTextColor.value ? 'text' : (rbWhite.value ? 'white' : (rbCMYK.value ? 'cmyk' : 'black')));
            } catch (e) { }
            try { __STATE.cmykC = String(fillC.text); } catch (e) { }
            try { __STATE.cmykM = String(fillM.text); } catch (e) { }
            try { __STATE.cmykY = String(fillY.text); } catch (e) { }
            try { __STATE.cmykK = String(fillK.text); } catch (e) { }

            // Opacity
            try { __STATE.opacityApply = !!cbOpacityApply.value; } catch (e) { }
            try { __STATE.opacity = String(opacityInput.text); } catch (e) { }

            __STATE._hasSaved = true;
        }

        function restoreUIStateToControls() {
            if (!__STATE || !__STATE._hasSaved) return;

            // Shape
            try {
                rbPerfectCircle.value = (__STATE.shape === 'circle');
                rbSuperEllipse.value = (__STATE.shape === 'super');
                rbRectangle.value = (__STATE.shape === 'rect');
            } catch (e) { }

            // Scale
            scaleInput.text = String(__STATE.scale);
            cbOneChar.value = !!__STATE.oneChar;

            // Margin
            marginVInput.text = String(__STATE.marginV);
            marginHInput.text = String(__STATE.marginH);
            cbMarginLink.value = !!__STATE.marginLink;
            try { if (cbMarginSquare) cbMarginSquare.value = !!__STATE.marginSquare; } catch (e) { }

            // Round / Pill
            cbRoundEnable.value = !!__STATE.roundEnable;
            roundInput.text = String(__STATE.roundValue);
            cbPill.value = !!__STATE.pill;

            // Group
            cbGroup.value = !!__STATE.groupWithText;
            cbExclude.value = !!__STATE.exclude;

            // Axis
            offsetXInput.text = String(__STATE.offsetX);
            offsetYInput.text = String(__STATE.offsetY);

            // Kind
            try {
                rbFill.value = (__STATE.kind !== 'stroke');
                rbStroke.value = (__STATE.kind === 'stroke');
            } catch (e) { }
            strokeWInput.text = String(__STATE.strokeWidth);

            // Color
            try {
                rbTextColor.value = (__STATE.colorMode === 'text');
                rbBlack.value = (__STATE.colorMode === 'black');
                rbWhite.value = (__STATE.colorMode === 'white');
                rbCMYK.value = (__STATE.colorMode === 'cmyk');
            } catch (e) { }
            fillC.text = String(__STATE.cmykC);
            fillM.text = String(__STATE.cmykM);
            fillY.text = String(__STATE.cmykY);
            fillK.text = String(__STATE.cmykK);

            // Opacity
            cbOpacityApply.value = !!__STATE.opacityApply;
            opacityInput.text = String(__STATE.opacity);
            opacityInput.enabled = cbOpacityApply.value;

            // Sync dependent UIs
            try { syncColorUI(); } catch (e) { }
            try { syncKindUI(); } catch (e) { }
            try { syncMarginUI(); } catch (e) { }
            try { syncRoundPanelUI(); } catch (e) { }
            try { syncMarginPanelUI(); } catch (e) { }
            try { syncScalePanelUI(); } catch (e) { }
        }

        // Restore UI values (if previously saved)
        try { restoreUIStateToControls(); } catch (e) { }
        // ラジオと入力のイベントでプレビュー更新
        var rbList = [rbTextColor, rbBlack, rbWhite, rbCMYK];
        for (var i = 0; i < rbList.length; i++) {
            rbList[i].onClick = function () {
                syncColorUI();
                if (typeof updatePreview === 'function') updatePreview();
            };
        }
        fillC.onChanging = updatePreview;
        fillM.onChanging = updatePreview;
        fillY.onChanging = updatePreview;
        fillK.onChanging = updatePreview;
        cbOpacityApply.onClick = function () {
            opacityInput.enabled = cbOpacityApply.value;
            updatePreview();
        };
        cbOpacityApply.onChanging = cbOpacityApply.onClick;
        opacityInput.onChanging = updatePreview;

        // 必要オブジェクト参照（プレビュー用に先に取得）
        var doc = app.activeDocument;
        if (!doc) {
            alert(getLabel('message.noDocument'));
            return;
        }

        // カラーのデフォルトは「ブラック」
        // ※ただし、前回値をリストア済みの場合は上書きしない
        try {
            if (!(__STATE && __STATE._hasSaved)) {
                rbBlack.value = true;
                rbTextColor.value = rbWhite.value = rbCMYK.value = false;
                if (typeof syncColorUI === 'function') syncColorUI();
            }
        } catch (e) { }

        var currentSelection = app.selection;
        if (!currentSelection || currentSelection.length === 0) {
            alert(getLabel('message.noSelection'));
            return;
        }

        // テキストがあれば従来ロジック（テキスト優先）
        var textItem = findFirstTextItem(currentSelection);
        var isTextMode = !!textItem;

        // テキストが無い場合：選択オブジェクト（複数なら矩形合成）を対象にする
        var targetItem = isTextMode ? textItem : null;

        // --- 既存の背面図形を検出 / Detect existing backdrop in selection ---
        var existingBackdrop = null;
        var ungroupTarget = null; // グループ解除対象

        function preparePreviewReplacementState() {
            // プレビュー用：グループ解除と既存背面図形の非表示だけを行う（Undoで巻き戻し可能）
            if (ungroupTarget) {
                try {
                    for (var pui = ungroupTarget.pageItems.length - 1; pui >= 0; pui--) {
                        var pChild = ungroupTarget.pageItems[pui];
                        if (pChild === existingBackdrop) continue;
                        pChild.move(ungroupTarget, ElementPlacement.PLACEBEFORE);
                    }
                    ungroupTarget.hidden = true;
                } catch (ePrepareUngroup) {
                    try { $.writeln('[AddBackdrop] preparePreviewReplacementState ungroup failed: ' + ePrepareUngroup); } catch (e) { }
                }
            }

            if (existingBackdrop) {
                try {
                    existingBackdrop.hidden = true;
                } catch (ePrepareBackdropHide) {
                    try { $.writeln('[AddBackdrop] preparePreviewReplacementState hide failed: ' + ePrepareBackdropHide); } catch (e) { }
                }
            }
        }

        function applyFinalReplacement() {
            // 確定用：既存背面図形の削除と、必要ならグループ自体の解体・削除を行う
            if (ungroupTarget) {
                try {
                    for (var ui = ungroupTarget.pageItems.length - 1; ui >= 0; ui--) {
                        var uChild = ungroupTarget.pageItems[ui];
                        if (uChild === existingBackdrop) continue;
                        uChild.move(ungroupTarget, ElementPlacement.PLACEBEFORE);
                    }
                } catch (eFinalUngroup) {
                    try { $.writeln('[AddBackdrop] applyFinalReplacement ungroup failed: ' + eFinalUngroup); } catch (e) { }
                }

                safeRemove(existingBackdrop, 'existingBackdrop');
                safeRemove(ungroupTarget, 'ungroupTarget');
                existingBackdrop = null;
                ungroupTarget = null;
                return;
            }

            if (existingBackdrop) {
                safeRemove(existingBackdrop, 'existingBackdrop');
                existingBackdrop = null;
            }
        }

        // パターン1: 単一グループ選択（テキスト＋図形/グループを内包）
        if (currentSelection.length === 1 && currentSelection[0].typename === 'GroupItem' && isTextMode) {
            var outerGrp = currentSelection[0];
            for (var gi = 0; gi < outerGrp.pageItems.length; gi++) {
                var gItem = outerGrp.pageItems[gi];
                if (gItem === textItem) continue;
                if (gItem.typename === 'TextFrame') continue;
                // テキスト以外のアイテム（PathItem / CompoundPathItem / GroupItem）= 既存背面図形
                existingBackdrop = gItem;
                break;
            }
            if (existingBackdrop) {
                ungroupTarget = outerGrp;

                // 形状タイプを検出してラジオボタンに反映
                var detectedShape0 = detectShapeType(existingBackdrop);
                rbPerfectCircle.value = (detectedShape0 === 'circle');
                rbSuperEllipse.value = (detectedShape0 === 'super');
                rbRectangle.value = (detectedShape0 === 'rect');

                // 依存UIを同期
                try { syncRoundPanelUI(); } catch (e) { }
                try { syncMarginPanelUI(); } catch (e) { }
                try { syncScalePanelUI(); } catch (e) { }
            }
        }

        // パターン2: 複数選択（テキスト+図形、グループ+図形）
        if (!existingBackdrop && currentSelection.length >= 2) {
            var _targets = [];
            var _backdrops = [];
            for (var si = 0; si < currentSelection.length; si++) {
                var siItem = currentSelection[si];
                if (siItem === textItem) continue;
                if (siItem.typename === 'TextFrame' || siItem.typename === 'GroupItem') {
                    _targets.push(siItem);
                } else if (siItem.typename === 'PathItem' || siItem.typename === 'CompoundPathItem') {
                    _backdrops.push(siItem);
                }
            }

            // テキスト+図形、またはグループ+図形のパターン
            if (_backdrops.length > 0 && (isTextMode || _targets.length > 0)) {
                existingBackdrop = _backdrops[0];

                // 非テキスト時：ターゲットをグループ等に設定
                if (!isTextMode && _targets.length > 0) {
                    targetItem = _targets[0];
                }

                // 形状タイプを検出してラジオボタンに反映
                var detectedShape = detectShapeType(existingBackdrop);
                rbPerfectCircle.value = (detectedShape === 'circle');
                rbSuperEllipse.value = (detectedShape === 'super');
                rbRectangle.value = (detectedShape === 'rect');

                // 依存UIを同期
                try { syncRoundPanelUI(); } catch (e) { }
                try { syncMarginPanelUI(); } catch (e) { }
                try { syncScalePanelUI(); } catch (e) { }
            }
        }

        // 非テキスト時は「1文字」や「テキストカラー参照」を無効化して破綻を避ける
        if (!isTextMode) {
            try {
                cbOneChar.value = false;
                cbOneChar.enabled = false;
            } catch (e) { }
            try {
                // テキストカラー参照が選ばれていたらブラックへフォールバック
                if (rbTextColor && rbTextColor.value) {
                    rbTextColor.value = false;
                    rbBlack.value = true;
                }
                rbTextColor.enabled = false;
            } catch (e) { }

            // 対象は選択の先頭（move/place の基準）― 既に設定済みでなければ
            if (!targetItem) {
                try {
                    targetItem = currentSelection[0];
                } catch (e) {
                    targetItem = null;
                }
            }

            if (!targetItem) {
                alert(getLabel('message.targetResolveError'));
                return;
            }
        }

        // 自動判定：選択中テキストが「1文字」ならチェックON（テキスト時のみ）
        if (isTextMode) {
            try {
                var raw = textItem.contents || "";
                // 改行・タブ・スペースを除外してカウント
                var count = String(raw).replace(/[\r\n\t\s]+/g, "").length;
                if (count === 1) {
                    cbOneChar.value = true;
                }
            } catch (e) { }
        }

        // プレビュー管理（Undoで巻き戻す）
        var previewCircle = null;

        function removePreview() {
            // 互換のため関数名は残すが、実体はUndo巻き戻し
            try { previewMgr.rollback(); } catch (eRollbackRemovePreview) {
                try { $.writeln('[AddBackdrop] rollback failed in removePreview: ' + eRollbackRemovePreview); } catch (e) { }
            }
            previewCircle = null;
        }

        // --- 見た目寸法キャッシュ（プレビュー最適化） ---
        var __measured = null; // {left, top, right, bottom, w, h}
        function measureTextVisualBoundsOnce() {
            if (__measured) return __measured;
            var dup = null;
            var outlinedGroup = null;
            // 初回のみ、見た目寸法を取得（アウトライン→bounds→破棄）
            try {
                dup = textItem.duplicate();
                outlinedGroup = dup.createOutline();
                var ogb = outlinedGroup.geometricBounds; // [l, t, r, b]
                __measured = {
                    left: ogb[0],
                    top: ogb[1],
                    right: ogb[2],
                    bottom: ogb[3],
                    w: ogb[2] - ogb[0],
                    h: ogb[1] - ogb[3]
                };
            } catch (e) {
                // フォールバック：フレームの幾何境界
                var gb0 = textItem.geometricBounds;
                __measured = {
                    left: gb0[0],
                    top: gb0[1],
                    right: gb0[2],
                    bottom: gb0[3],
                    w: gb0[2] - gb0[0],
                    h: gb0[1] - gb0[3]
                };
            } finally {
                safeRemove(outlinedGroup, 'outlinedGroup');
                safeRemove(dup, 'dup');
            }
            return __measured;
        }

        // 非テキスト時：選択オブジェクトの矩形を合成（可能なら visibleBounds を優先）
        // - visibleBounds: 線幅など見た目を含みやすい
        // - geometricBounds: フォールバック
        function measureSelectionBounds() {
            var b = null; // [l,t,r,b]
            try {
                for (var i = 0; i < currentSelection.length; i++) {
                    var it = currentSelection[i];
                    if (!it) continue;
                    // 置換対象の既存図形はバウンド計算から除外
                    if (existingBackdrop && it === existingBackdrop) continue;

                    var gb = null;
                    try {
                        if (it.visibleBounds) gb = it.visibleBounds;
                    } catch (eVisible) {
                        try { $.writeln('[AddBackdrop] visibleBounds read failed: ' + eVisible); } catch (e) { }
                    }
                    if (!gb) {
                        try { gb = it.geometricBounds; } catch (eGeom) {
                            try { $.writeln('[AddBackdrop] geometricBounds read failed: ' + eGeom); } catch (e) { }
                        }
                    }

                    if (!gb) continue;

                    if (!b) {
                        b = [gb[0], gb[1], gb[2], gb[3]];
                    } else {
                        b[0] = Math.min(b[0], gb[0]);
                        b[1] = Math.max(b[1], gb[1]);
                        b[2] = Math.max(b[2], gb[2]);
                        b[3] = Math.min(b[3], gb[3]);
                    }
                }
            } catch (e) {
                b = null;
            }
            return b;
        }

        // マージンのデフォルト：選択オブジェクトの短辺の 1/4
        var __didInitMarginDefault = !!(__STATE && __STATE._hasSaved);
        function initMarginDefaultFromSelection() {
            if (__didInitMarginDefault) return;
            __didInitMarginDefault = true;

            var ww = 0, hh = 0;
            try {
                if (isTextMode) {
                    // テキスト：通常は見た目寸法キャッシュ、未計測なら幾何境界
                    var m2 = null;
                    try { m2 = measureTextVisualBoundsOnce(); } catch (e) { m2 = null; }
                    if (m2) {
                        ww = m2.w;
                        hh = m2.h;
                    } else {
                        var gbT = textItem.geometricBounds;
                        ww = gbT[2] - gbT[0];
                        hh = gbT[1] - gbT[3];
                    }
                } else {
                    // 非テキスト：選択範囲の矩形
                    var bb = measureSelectionBounds();
                    if (!bb) {
                        var gbb2 = null;
                        try { if (targetItem.visibleBounds) gbb2 = targetItem.visibleBounds; } catch (e) { }
                        if (!gbb2) {
                            try { gbb2 = targetItem.geometricBounds; } catch (e) { }
                        }
                        bb = [gbb2[0], gbb2[1], gbb2[2], gbb2[3]];
                    }
                    ww = bb[2] - bb[0];
                    hh = bb[1] - bb[3];
                }
            } catch (e) {
                ww = 0; hh = 0;
            }

            var s = Math.min(Math.abs(ww), Math.abs(hh));
            if (!isFinite(s) || s <= 0) return;

            var def = s / 4;
            // 見た目が良いように：小数1桁まで（整数に近い場合は整数）
            var v = Math.round(def * 10) / 10;
            if (Math.abs(v - Math.round(v)) < 1e-6) v = Math.round(v);

            try {
                marginVInput.text = String(v);
                marginHInput.text = String(v);
            } catch (e) { }

            try { syncMarginUI(); } catch (e) { }
        }

        // 角丸のデフォルト：選択オブジェクトの短辺の 1/5
        var __didInitRoundDefault = !!(__STATE && __STATE._hasSaved);
        function initRoundDefaultFromSelection() {
            if (__didInitRoundDefault) return;
            __didInitRoundDefault = true;

            var ww = 0, hh = 0;
            try {
                if (isTextMode) {
                    var m2 = null;
                    try { m2 = measureTextVisualBoundsOnce(); } catch (e) { m2 = null; }
                    if (m2) {
                        ww = m2.w;
                        hh = m2.h;
                    } else {
                        var gbT = textItem.geometricBounds;
                        ww = gbT[2] - gbT[0];
                        hh = gbT[1] - gbT[3];
                    }
                } else {
                    var bb = measureSelectionBounds();
                    if (!bb) {
                        var gbb2 = null;
                        try { if (targetItem.visibleBounds) gbb2 = targetItem.visibleBounds; } catch (e) { }
                        if (!gbb2) {
                            try { gbb2 = targetItem.geometricBounds; } catch (e) { }
                        }
                        bb = [gbb2[0], gbb2[1], gbb2[2], gbb2[3]];
                    }
                    ww = bb[2] - bb[0];
                    hh = bb[1] - bb[3];
                }
            } catch (e) {
                ww = 0; hh = 0;
            }

            var s = Math.min(Math.abs(ww), Math.abs(hh));
            if (!isFinite(s) || s <= 0) return;

            var def = s / 5;
            // 見た目が良いように：小数1桁まで（整数に近い場合は整数）
            var v = Math.round(def * 10) / 10;
            if (Math.abs(v - Math.round(v)) < 1e-6) v = Math.round(v);

            try {
                roundInput.text = String(v);
                __lastRoundValue = String(roundInput.text);
            } catch (e) { }
        }

        // 計算（UI状態に応じてサイズ・中心を返す）
        function computeParams() {
            var left, top, right, bottom, w, h;

            if (isTextMode) {
                // 基準寸法：
                //  - 1文字モード: アウトライン不要（フォントサイズベース）
                //  - 通常モード  : 初回のみアウトラインして見た目寸法をキャッシュ
                if (!cbOneChar.value) {
                    // 通常：見た目寸法を使用（アウトライン計測のキャッシュ）
                    var m = measureTextVisualBoundsOnce();
                    left = m.left;
                    top = m.top;
                    right = m.right;
                    bottom = m.bottom;
                    w = m.w;
                    h = m.h;
                } else {
                    // 1文字モード：幾何境界＋フォントサイズベース
                    var gb0 = textItem.geometricBounds; // [l,t,r,b]
                    left = gb0[0];
                    top = gb0[1];
                    right = gb0[2];
                    bottom = gb0[3];
                    w = right - left;
                    h = top - bottom; // 参照用（A取得失敗時のフォールバック）
                }
            } else {
                // 非テキスト：選択範囲の矩形（geometricBounds 合成）
                var b = measureSelectionBounds();
                if (!b) {
                    // 最後の保険：targetItem の bounds（visibleBounds 優先）
                    var gbb = null;
                    try { if (targetItem.visibleBounds) gbb = targetItem.visibleBounds; } catch (e) { }
                    if (!gbb) {
                        try { gbb = targetItem.geometricBounds; } catch (e) { }
                    }
                    b = [gbb[0], gbb[1], gbb[2], gbb[3]];
                }
                left = b[0];
                top = b[1];
                right = b[2];
                bottom = b[3];
                w = right - left;
                h = top - bottom;
            }

            // マージン適用（w/h を拡張）
            var mg = getMarginValues ? getMarginValues() : { mx: 0, my: 0 };
            var effectiveW = w + (mg.mx * 2);
            var effectiveH = h + (mg.my * 2);

            var cx = left + w / 2;
            var cy = top - h / 2;

            var userScaleVal = parseFloat(scaleInput.text);
            if (isNaN(userScaleVal) || userScaleVal <= 0) userScaleVal = 100;
            userScaleVal = userScaleVal / 100; // percent → multiplier

            var d;
            if (isTextMode && cbOneChar.value) {
                // 1文字モード：A=フォントサイズ、中心Y=上端−A/2、直径=1.5×A
                var A;
                try {
                    A = textItem.textRange.characterAttributes.size;
                } catch (e) {
                    A = Math.max(w, h);
                }
                var B = A / 2;
                d = A * 1.5 * userScaleVal;
                cy = top - B;
            } else {
                // 正方形がすっぽり入る円の直径 = 一辺 × √2 × ユーザー倍率
                var squareSide = Math.max(effectiveW, effectiveH);
                d = squareSide * Math.SQRT2 * userScaleVal;
            }

            // オフセット適用（負値可）
            var offX = parseFloat(offsetXInput.text);
            var offY = parseFloat(offsetYInput.text);
            if (!isNaN(offX)) cx += offX;
            if (!isNaN(offY)) cy += offY;

            // 長方形用（w/h をスケール適用した値も返す）
            var rectW = effectiveW * userScaleVal;
            var rectH = effectiveH * userScaleVal;

            // 正方形ONのときは「正円と同様のロジック」で d を一辺として正方形を描く
            try {
                if (cbMarginSquare && cbMarginSquare.value) {
                    rectW = d;
                    rectH = d;
                }
            } catch (e) { }

            return {
                cx: cx,
                cy: cy,
                d: d,
                rectW: rectW,
                rectH: rectH,
                // for pill sizing
                baseW: w,
                baseH: h,
                effectiveW: effectiveW,
                effectiveH: effectiveH,
                scale: userScaleVal
            };
        }

        function resolveBackdropColor() {
            // テキストカラー参照
            if (rbTextColor.value) {
                var tcol = null;
                try {
                    tcol = textItem.textRange.characterAttributes.fillColor;
                } catch (e) { }
                if (tcol) return tcol;
            }

            if (rbBlack.value || rbTextColor.value) {
                var kcol = new GrayColor();
                kcol.gray = 100;
                return kcol;
            }
            if (rbWhite.value) {
                var wcol = new GrayColor();
                wcol.gray = 0;
                return wcol;
            }
            if (rbCMYK.value) {
                var c = Math.min(100, Math.max(0, parseInt(fillC.text, 10) || 0));
                var m = Math.min(100, Math.max(0, parseInt(fillM.text, 10) || 0));
                var y = Math.min(100, Math.max(0, parseInt(fillY.text, 10) || 0));
                var k = Math.min(100, Math.max(0, parseInt(fillK.text, 10) || 0));
                var cmyk = new CMYKColor();
                cmyk.cyan = c;
                cmyk.magenta = m;
                cmyk.yellow = y;
                cmyk.black = k;
                return cmyk;
            }

            var fallback = new GrayColor();
            fallback.gray = 100;
            return fallback;
        }

        function applyStyleToItem(item) {
            if (!item) return;
            try {
                var col = resolveBackdropColor();

                if (rbFill.value) {
                    item.filled = true;
                    item.stroked = false;
                    item.fillColor = col;
                } else {
                    item.filled = false;
                    item.stroked = true;
                    var sw = parseFloat(strokeWInput.text);
                    if (isNaN(sw) || sw < 0) sw = 1;
                    item.strokeWidth = sw;
                    item.strokeColor = col;
                }
            } catch (e) {
                try {
                    var g2 = new GrayColor();
                    g2.gray = 20;
                    item.fillColor = g2;
                    item.filled = true;
                    item.stroked = false;
                } catch (e) { }
            }

            try {
                var opv = parseFloat(opacityInput.text);
                if (isNaN(opv)) opv = 100;
                opv = Math.max(0, Math.min(100, opv));
                item.opacity = cbOpacityApply.value ? opv : 100;
            } catch (e) { }
        }

        function reapplyStyleAfterConvertToShape(item) {
            applyStyleToItem(item);
        }

        function reapplyStyleAfterPathfinder(item) {
            applyStyleToItem(item);
        }

        // Live Effect: Round Corners（ライブ効果で角丸を適用）
        function applyRoundCornersLive(item, r) {
            try {
                if (!item || !r || r <= 0) return;
                var xml = '<LiveEffect name="Adobe Round Corners"><Dict data="R radius ' + r + ' "/></LiveEffect>';
                item.applyEffect(xml);
            } catch (e) {
                try { $.writeln('[AddBackdrop] applyRoundCornersLive failed: ' + e); } catch (e) { }
            }
        }

        function buildBackdropOnce() {
            var p = computeParams();
            var created = null;

            // --- shape creation ---
            if (rbRectangle.value) {
                var rectLeft = p.cx - p.rectW / 2;
                var rectTop = p.cy + p.rectH / 2;

                var pillOn = false;
                var roundOn = false;
                var r = 0;
                try {
                    pillOn = !!cbPill.value;
                    roundOn = !!cbRoundEnable.value;
                    if (!pillOn && roundOn) {
                        r = parseFloat(roundInput.text);
                        if (isNaN(r) || r < 0) r = 0;
                    }
                } catch (e) { }

                if (pillOn) {
                    var pr = p.rectH / 2;
                    var minW = (Math.abs(p.baseW) * (p.scale || 1)) + p.rectH;
                    var pillW = Math.max(p.rectW, minW);
                    var pillLeft = p.cx - pillW / 2;

                    created = doc.pathItems.rectangle(rectTop, pillLeft, pillW, p.rectH);
                    applyRoundCornersLive(created, pr);
                } else if (roundOn && r > 0) {
                    created = doc.pathItems.rectangle(rectTop, rectLeft, p.rectW, p.rectH);
                    applyRoundCornersLive(created, r);
                } else {
                    created = doc.pathItems.rectangle(rectTop, rectLeft, p.rectW, p.rectH);
                }
            } else {
                var circleLeft = p.cx - p.d / 2;
                var circleTop = p.cy + p.d / 2;
                created = doc.pathItems.ellipse(circleTop, circleLeft, p.d, p.d);

                if (rbSuperEllipse.value) {
                    var n = 2.5;
                    morphPathToSuperellipse(created, p.cx, p.cy, p.d, p.d, n);
                }
            }

            // --- stacking ---
            try {
                var base = isTextMode ? textItem : targetItem;
                created.move(base, ElementPlacement.PLACEAFTER);
            } catch (e) {
                try {
                    created.zOrder(ZOrderMethod.SENDTOBACK);
                } catch (zerr) {
                    try { $.writeln('[AddBackdrop] zOrder fallback failed: ' + zerr); } catch (e) { }
                }
            }

            applyStyleToItem(created);
            return created;
        }

        function finalizeBackdropShape(finalShape) {
            try {
                if (!finalShape) return finalShape;

                if (rbPerfectCircle.value) {
                    clearSelection();
                    finalShape.selected = true;
                    app.executeMenuCommand('Convert to Shape');
                    var converted = null;
                    try {
                        converted = app.selection && app.selection.length ? app.selection[0] : null;
                    } catch (eConvertedSelection) {
                        try { $.writeln('[AddBackdrop] converted selection read failed: ' + eConvertedSelection); } catch (e) { }
                    }
                    if (!converted) converted = finalShape;
                    finalShape = converted;
                    reapplyStyleAfterConvertToShape(finalShape);
                    clearSelection();
                    return finalShape;
                }

                if (rbRectangle.value) {
                    try {
                        clearSelection();
                        finalShape.selected = true;
                        app.executeMenuCommand('Live Pathfinder Add');
                        var pathfinderAdded = null;
                        try {
                            pathfinderAdded = app.selection && app.selection.length ? app.selection[0] : null;
                        } catch (ePathfinderSelection) {
                            try { $.writeln('[AddBackdrop] pathfinder selection read failed: ' + ePathfinderSelection); } catch (e) { }
                        }
                        if (!pathfinderAdded) pathfinderAdded = finalShape;
                        finalShape = pathfinderAdded;
                        reapplyStyleAfterPathfinder(finalShape);
                        clearSelection();
                    } catch (ePathfinderAdd) {
                        try { $.writeln('[AddBackdrop] Live Pathfinder Add failed: ' + ePathfinderAdd); } catch (e) { }
                    }
                }
            } catch (eFinalizeShape) {
                try { $.writeln('[AddBackdrop] final shape finalize step failed: ' + eFinalizeShape); } catch (e) { }
            }
            return finalShape;
        }

        function placeFinalResult(finalShape) {
            var baseItem = isTextMode ? textItem : targetItem;

            if (cbGroup.value) {
                try {
                    var parent = baseItem.parent;
                    var resultGroup = parent.groupItems.add();
                    try { resultGroup.move(baseItem, ElementPlacement.PLACEAFTER); } catch (eMoveGroup) {
                        try { $.writeln('[AddBackdrop] group move failed: ' + eMoveGroup); } catch (e) { }
                    }
                    finalShape.move(resultGroup, ElementPlacement.PLACEATEND);
                    baseItem.move(resultGroup, ElementPlacement.PLACEATEND);
                    try { finalShape.zOrder(ZOrderMethod.SENDTOBACK); } catch (eGroupBack) {
                        try { $.writeln('[AddBackdrop] grouped finalShape zOrder failed: ' + eGroupBack); } catch (e) { }
                    }
                    return resultGroup;
                } catch (eGroup) {
                    alert(getLabel('message.groupFailed') + ": " + eGroup);
                    return finalShape;
                }
            }

            try {
                finalShape.move(baseItem, ElementPlacement.PLACEAFTER);
            } catch (ePlaceUngrouped) {
                try { $.writeln('[AddBackdrop] ungrouped finalShape move failed: ' + ePlaceUngrouped); } catch (e) { }
            }

            return finalShape;
        }

        function finalizeExclude(resultItem) {
            if (!(cbExclude.value && resultItem)) return resultItem;

            try {
                clearSelection();
                resultItem.selected = true;
                app.executeMenuCommand('Live Pathfinder Exclude');

                var excludedResult = null;
                try {
                    excludedResult = app.selection && app.selection.length ? app.selection[0] : null;
                } catch (eExcludedSelection) {
                    try { $.writeln('[AddBackdrop] exclude selection read failed: ' + eExcludedSelection); } catch (e) { }
                }
                if (!excludedResult) excludedResult = resultItem;
                return excludedResult;
            } catch (eExclude) {
                alert(getLabel('message.excludeFailed') + ": " + eExclude);
                return resultItem;
            } finally {
                clearSelection();
            }
        }

        function selectFinalResult(resultItem) {
            try {
                clearSelection();
                if (resultItem) {
                    resultItem.selected = true;
                }
            } catch (eSelectResult) {
                try { $.writeln('[AddBackdrop] result selection failed: ' + eSelectResult); } catch (e) { }
            }
        }

        function updatePreview() {
            // 1) rollback previous preview changes
            try { previewMgr.rollback(); } catch (eRollbackUpdatePreview) {
                try { $.writeln('[AddBackdrop] rollback failed in updatePreview: ' + eRollbackUpdatePreview); } catch (e) { }
            }
            previewCircle = null;

            // 2) apply one preview step (counts as 1 undo step)
            previewMgr.addStep(function () {
                preparePreviewReplacementState();
                previewCircle = buildBackdropOnce();
            });

            // 3) Save UI state for next run (no document edits)
            try { saveUIStateFromControls(); } catch (e) { }
        }

        // 入力変更でプレビュー更新（ユーティリティで一括）
        bindPreview(offsetXInput, updatePreview);
        bindPreview(offsetYInput, updatePreview);
        bindPreview(scaleInput, updatePreview);
        bindPreview(strokeWInput, updatePreview);
        bindPreview(fillC, updatePreview);
        bindPreview(fillM, updatePreview);
        bindPreview(fillY, updatePreview);
        bindPreview(fillK, updatePreview);
        bindPreview(opacityInput, updatePreview);

        // チェック系
        cbOneChar.onClick = updatePreview;
        cbOneChar.onChanging = updatePreview;

        var btns = dlg.add('group');
        btns.alignment = 'center';
        // Order: Cancel (left), OK (right)
        var cancelBtn = btns.add('button', undefined, __BTN_LABELS.cancel[__lang]);
        cancelBtn.name = 'cancel';
        var okBtn = btns.add('button', undefined, __BTN_LABELS.ok[__lang]);
        okBtn.name = 'ok';

        // 明示的にキャンセルを処理（プレビュー掃除→ダイアログを閉じる）
        cancelBtn.onClick = function () {
            // Cancel: rollback all preview edits
            try { previewMgr.rollback(); } catch (eRollbackCancelBtn) {
                try { $.writeln('[AddBackdrop] rollback failed in cancelBtn.onClick: ' + eRollbackCancelBtn); } catch (e) { }
            }
            try { saveUIStateFromControls(); } catch (e) { }
            try { dlg.close(0); } catch (e) { }
        };

        // 初回表示後に初期化と初期プレビューを行う
        var __didRunInitialPreview = false;
        try {
            dlg.onShow = (function (prev) {
                return function () {
                    try { if (typeof prev === 'function') prev(); } catch (e) { }
                    try { scaleInput.active = true; } catch (e) { }

                    if (!__didRunInitialPreview) {
                        __didRunInitialPreview = true;
                        try { initMarginDefaultFromSelection(); } catch (e) { }
                        try { initRoundDefaultFromSelection(); } catch (e) { }
                        try { updatePreview(); } catch (e) { }
                    }
                };
            })(dlg.onShow);
        } catch (e) { }

        if (dlg.show() !== 1) {
            try { saveUIStateFromControls(); } catch (e) { }
            // Cancel path: rollback preview
            try { previewMgr.rollback(); } catch (eRollbackDialogCancel) {
                try { $.writeln('[AddBackdrop] rollback failed in dialog cancel path: ' + eRollbackDialogCancel); } catch (e) { }
            }
            return;
        }

        // OK was pressed: persist final UI state
        try { saveUIStateFromControls(); } catch (e) { }

        // Confirm: rollback preview, then run final action ONCE (single undo step)
        previewMgr.confirm(function () {
            // 0) 確定用の置換処理（既存背面図形の削除と必要なグループ解体）
            applyFinalReplacement();

            // 1) build final shape
            var finalShape = buildBackdropOnce();
            if (!finalShape) return;

            // 2) finalize shape
            finalShape = finalizeBackdropShape(finalShape);

            // 3) place result (group if enabled)
            var resultItem = placeFinalResult(finalShape);

            // 4) Pathfinder Exclude
            resultItem = finalizeExclude(resultItem);

            // 5) Select final result
            selectFinalResult(resultItem);
        });

        return;
    }

    /* -----------------------------------------------------------------------------
     * 選択配列から最初の TextFrame を見つける / Find first TextFrame in selection
     * --------------------------------------------------------------------------- */
    function findFirstTextItem(selectionArray) {
        for (var i = 0; i < selectionArray.length; i++) {
            var it = selectionArray[i];
            // テキストそのもの
            if (it.typename === "TextFrame") return it;

            // グループ等の中を掘る場合
            if (it.typename === "GroupItem") {
                var tf = findTextInGroup(it);
                if (tf) return tf;
            }
        }
        return null;
    }

    /* 再帰的に GroupItem 内を検索 / Recursively search inside GroupItem */
    function findTextInGroup(groupItem) {
        for (var i = 0; i < groupItem.pageItems.length; i++) {
            var it = groupItem.pageItems[i];
            if (it.typename === "TextFrame") return it;
            if (it.typename === "GroupItem") {
                var tf = findTextInGroup(it);
                if (tf) return tf;
            }
        }
        return null;
    }

    /**
     * 既存の図形から形状タイプを判定 / Detect shape type from an existing PathItem
     * @returns 'circle' | 'super' | 'rect'
     */
    function detectShapeType(item) {
        if (!item) return 'rect';
        try {
            var path = item;
            // CompoundPathItem: 最初のサブパスをチェック
            if (item.typename === 'CompoundPathItem') {
                try { path = item.pathItems[0]; } catch (e) { return 'rect'; }
            }
            // GroupItem: 内部の最初の PathItem を探す
            if (item.typename === 'GroupItem') {
                var foundPath = null;
                try {
                    for (var gi = 0; gi < item.pageItems.length; gi++) {
                        var gpi = item.pageItems[gi];
                        if (gpi.typename === 'PathItem') { foundPath = gpi; break; }
                        if (gpi.typename === 'CompoundPathItem') {
                            try { foundPath = gpi.pathItems[0]; } catch (e) { }
                            if (foundPath) break;
                        }
                    }
                } catch (e) { }
                if (foundPath) {
                    path = foundPath;
                } else {
                    // PathItem が見つからなければバウンディングボックスで判定
                    try {
                        var igb = item.geometricBounds;
                        var iw = igb[2] - igb[0];
                        var ih = igb[1] - igb[3];
                        if (Math.abs(iw - ih) < Math.max(iw, ih) * 0.05) return 'circle';
                    } catch (e) { }
                    return 'rect';
                }
            }
            if (!path || path.typename !== 'PathItem') return 'rect';

            var pp = path.pathPoints;
            var len = pp.length;

            // スーパー楕円（このスクリプトでは8点で生成）
            if (len === 8) return 'super';

            // 4点: 楕円/正円 or 長方形
            if (len === 4) {
                var hasBezier = false;
                for (var i = 0; i < 4; i++) {
                    var a = pp[i].anchor;
                    var ld = pp[i].leftDirection;
                    var rd = pp[i].rightDirection;
                    var dl = Math.abs(a[0] - ld[0]) + Math.abs(a[1] - ld[1]);
                    var dr = Math.abs(a[0] - rd[0]) + Math.abs(a[1] - rd[1]);
                    if (dl > 0.5 || dr > 0.5) {
                        hasBezier = true;
                        break;
                    }
                }
                // ベジェハンドルがあれば楕円系 → 正円扱い
                if (hasBezier) return 'circle';
                // ハンドルなし → 長方形
                return 'rect';
            }

            return 'rect';
        } catch (e) {
            return 'rect';
        }
    }

    // run
    try {
        main();
    } catch (err) {
        alert(getLabel('message.genericError') + ": " + err);
    }

})();
