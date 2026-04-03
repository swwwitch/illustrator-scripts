#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

// バージョン / Version
var SCRIPT_VERSION = "v1.3";

/*
【概要 / Overview】
選択したテキスト（またはオブジェクト）の背面に、見た目寸法に基づく「図形」を自動生成して背面配置します。正円／スーパー楕円／長方形に対応し、マージン、角丸、ピル形状、正方形オプションなどを指定できます。OK 時にはプレビューで生成した図形を確定します。

- 対象 / Target: pointText / areaText。非テキスト選択時は選択範囲を対象
- 形状 / Shape: 正円 / スーパー楕円（n=2.5） / 長方形
  - キー操作: E=正円, S=スーパー楕円, R=長方形
- マージン / Margin: 上下・左右（連動可）。デフォルトは短辺の 1/4
  - 長方形選択時のみ有効
  - 「正方形」ON: マージンを無視し、正円と同等ロジックのサイズで正方形を作成（スケール有効）
- 角丸 / Round: 半径入力で角丸長方形。デフォルトは短辺の 1/5
- ピル形状 / Pill: 高さの半分を半径とするライブ角丸でカプセル形状を作成（長方形時）
- 位置 / Position: 対象の直下（背面）へ配置。X/Y オフセット対応
- スケール / Scale: ％指定（通常の長方形時は 100% 固定）。正方形ON時や円系では有効。＋「1文字」モード
- 種別 / Kind: 塗り / 線（線幅指定）
- カラー / Color: ブラック / ホワイト / CMYK / テキストカラー参照
- 不透明度 / Opacity: ［適用］チェック＋％
- グループ / Group: 「テキストとグループ化」／「中マド処理（Exclude）」対応
- プレビュー / Preview: 入力変更・矢印キー（↑↓ / Shift=±10 / Option=±0.1）で即時更新
- ローカライズ / Localization: すべてのラベルは日英対応。ダイアログタイトルに SCRIPT_VERSION を併記
*/

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
    dialogTitle: {
        ja: "テキストの背面に図形を作成",
        en: "Create Shape Behind Text"
    },
    // --- New Option Panel Labels ---
    optionsPanel: {
        ja: "形状",
        en: "Shape"
    },
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
    marginTitle: {
        ja: "マージン",
        en: "Margin"
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
    // --- 角丸・ピル形状
    roundTitle: {
        ja: "角丸",
        en: "Round"
    },
    pillShape: {
        ja: "ピル形状",
        en: "Pill shape"
    },
    // ---
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
    groupWithText: {
        ja: "テキストとグループ化",
        en: "Group with Text"
    },
    exclude: {
        ja: "中マド処理",
        en: "Exclude"
    },
    kindPanel: {
        ja: "種別",
        en: "Kind"
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
    colorPanel: {
        ja: "カラー",
        en: "Color"
    },
    opacityPanel: {
        ja: "不透明度",
        en: "Opacity"
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
    }
};

function L(key) {
    try {
        var t = LABELS[key];
        return (t && (t[__lang] || t.ja || t.en)) || key;
    } catch (e) {
        return key;
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

// --- Illustrator 単位ユーティリティ関数群 ---
// ▼ 各設定キーの意味：
// - "rulerType"       ：一般（定規の単位）
// - "strokeUnits"     ：線
// - "text/units"      ：文字
// - "text/asianunits" ：東アジア言語のオプション

// 単位コード → ラベル
var unitMap = {
    0: "in",
    1: "mm",
    2: "pt",
    3: "pica",
    4: "cm",
    5: "Q/H", // 後段で Q/H に分岐
    6: "px",
    7: "ft/in",
    8: "m",
    9: "yd",
    10: "ft"
};

// 単位コードと設定キーから適切な単位ラベルを返す（Q/H分岐含む）
function getUnitLabel(code, prefKey) {
    if (code === 5) {
        var hKeys = {
            "text/asianunits": true,
            "rulerType": true,
            "strokeUnits": true
        };
        return hKeys[prefKey] ? "H" : "Q";
    }
    return unitMap[code] || "?";
}

// --- Dialog visual adjustments (opacity & initial shift) ---
// 初期シフト量と不透明度
var __DIALOG_OFFSET_X = 300;
var __DIALOG_OFFSET_Y = 0;
var __DIALOG_OPACITY = 0.97;

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
            } catch (_) { }
            try {
                var l = dlg.location;
                dlg.location = [l[0] + (dx | 0), l[1] + (dy | 0)];
            } catch (_) { }
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
        } catch (_) { }
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

            editText.text = value;

            // キー操作でも即時プレビュー
            try {
                if (typeof onUpdate === 'function') onUpdate();
            } catch (_) { }

            // 既定動作の抑止（可能な環境のみ）
            try {
                if (event && event.preventDefault) event.preventDefault();
            } catch (_) { }
        });
    }
}

// 入力フィールドにプレビュー更新を一括でバインド / Bind preview updates to an input
function bindPreview(editText, onUpdate) {
    if (!editText) return;
    // 通常入力の変更でプレビュー
    editText.onChanging = onUpdate;
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
            try { if (typeof onChange === 'function') onChange(); } catch (_) {}
            try { if (event.preventDefault) event.preventDefault(); } catch (_) {}
        } catch (_) { }
    });
}

/* =============================================================================
 * main entry
 * ========================================================================== */
function main() {
    // ダイアログ作成
    var dlg = new Window('dialog', L('dialogTitle') + ' ' + SCRIPT_VERSION);
    setDialogOpacity(dlg, __DIALOG_OPACITY);
    shiftDialogPositionOnceOnShow(dlg, __DIALOG_OFFSET_X, __DIALOG_OFFSET_Y);

    dlg.orientation = 'column';
    dlg.alignChildren = ['fill', 'top'];

    // 形状（ダイアログ上部・2カラムから独立）
    var grpShapeTop = dlg.add('group');
    grpShapeTop.orientation = 'row';
    grpShapeTop.alignChildren = ['center', 'center'];
    grpShapeTop.alignment = 'center';

    var rbPerfectCircle = grpShapeTop.add('radiobutton', undefined, L('perfectCircle'));
    var rbSuperEllipse = grpShapeTop.add('radiobutton', undefined, L('superEllipse'));
    var rbRectangle = grpShapeTop.add('radiobutton', undefined, L('rectangle'));

    rbPerfectCircle.value = true; // デフォルトは正円

    // どのオプションを選んでもプレビュー更新
    var _onOptionChanged = function () {
        try { if (typeof syncRoundPanelUI === 'function') syncRoundPanelUI(); } catch (_) { }
        try { if (typeof syncMarginPanelUI === 'function') syncMarginPanelUI(); } catch (_) { }
        try { if (typeof syncScalePanelUI === 'function') syncScalePanelUI(); } catch (_) { }
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
    var __rulerCode = app.preferences.getIntegerPreference('rulerType');
    var __strokeCode = app.preferences.getIntegerPreference('strokeUnits');
    var __rulerLabel = getUnitLabel(__rulerCode, 'rulerType');
    var __strokeLabel = getUnitLabel(__strokeCode, 'strokeUnits');
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
    var pnlScale = leftCol.add('panel', undefined, L('scalePanel'));
    pnlScale.orientation = 'column';
    pnlScale.alignChildren = ['fill', 'top'];
    pnlScale.margins = [15, 20, 15, 10];

    // 倍率（%表示）
    var grpScale = pnlScale.add('group');
    grpScale.add('statictext', undefined, L('magnification'));
    var scaleInput = grpScale.add('edittext', undefined, '90');
    scaleInput.characters = 4;
    grpScale.add('statictext', undefined, '%');

    // 1文字チェック（倍率の直下）
    var grpOneChar = pnlScale.add('group');
    grpOneChar.orientation = 'row';
    grpOneChar.alignChildren = ['left', 'center'];
    var cbOneChar = grpOneChar.add('checkbox', undefined, L('oneChar'));
    cbOneChar.value = false;

    // スケールパネルは「長方形」選択時は 100% 固定 + ディム表示（ただし正方形ON時は例外で有効）
    function syncScalePanelUI() {
        try {
            var squareOn = false;
            try { squareOn = !!(cbMarginSquare && cbMarginSquare.value); } catch (_) { squareOn = false; }

            // 長方形選択時は通常 100%固定 + ディム。正方形ONのときだけ例外で有効化。
            if (rbRectangle && rbRectangle.value && !squareOn) {
                // 強制 100%
                scaleInput.text = '100';
                // ついでに 1文字は意味が薄いのでOFF（パネル全体がディムなので安全側）
                try { cbOneChar.value = false; } catch (_) {}

                pnlScale.enabled = false;
            } else {
                pnlScale.enabled = true;
            }
        } catch (_) { }
    }
    syncScalePanelUI();

    // パネル：マージン（形状パネルの直下）
    var marginPanel = leftCol.add('panel', undefined, L('marginTitle'));
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

    groupMV.add('statictext', undefined, L('marginV'));
    var marginVInput = groupMV.add('edittext', undefined, '0');
    marginVInput.characters = 4;
    groupMV.add('statictext', undefined, __rulerLabel);

    // GROUP3 (row: 左右)
    var groupMH = marginGroup1.add('group');
    groupMH.orientation = 'row';
    groupMH.alignChildren = ['left', 'center'];
    groupMH.spacing = 10;
    groupMH.margins = 0;

    groupMH.add('statictext', undefined, L('marginH'));
    var marginHInput = groupMH.add('edittext', undefined, '0');
    marginHInput.characters = 4;
    groupMH.add('statictext', undefined, __rulerLabel);

    // RIGHT side: Link checkbox (center-ish)
    var marginRightCol = marginLeftCol.add('group');
    marginRightCol.orientation = 'column';
    marginRightCol.alignChildren = ['left', 'center'];
    marginRightCol.spacing = 0;
    marginRightCol.margins = 0;

    // small spacer to align checkbox between two rows
    try { marginRightCol.add('statictext', undefined, ''); } catch (_) {}

    var cbMarginLink = marginRightCol.add('checkbox', undefined, L('link'));
    cbMarginLink.value = true; // デフォルトは連動

    // マージン：正方形
    cbMarginSquare = marginPanel.add('checkbox', undefined, '正方形');
    cbMarginSquare.value = false; // デフォルトOFF

    cbMarginSquare.onClick = function () {
        // 正方形ON: マージン無視 + スケールを有効化しアクティブに
        try { if (typeof syncScalePanelUI === 'function') syncScalePanelUI(); } catch (_) { }
        try { if (cbMarginSquare.value) { scaleInput.active = true; } } catch (_) { }
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
        } catch (_) {}
    }
    syncMarginUI();

    cbMarginLink.onClick = function () {
        syncMarginUI();
        if (typeof updatePreview === 'function') updatePreview();
    };
    cbMarginLink.onChanging = cbMarginLink.onClick;

    // 入力変更で常に同期 + プレビュー
    marginVInput.onChanging = function () {
        try { syncMarginUI(); } catch (_) {}
        if (typeof updatePreview === 'function') updatePreview();
    };
    marginHInput.onChanging = function () {
        if (typeof updatePreview === 'function') updatePreview();
    };

    // 矢印キー増減（↑↓ / Shift / Option）
    changeValueByArrowKey(marginVInput, function () {
        try { syncMarginUI(); } catch (_) {}
        if (typeof updatePreview === 'function') updatePreview();
    });
    changeValueByArrowKey(marginHInput, function () {
        if (typeof updatePreview === 'function') updatePreview();
    });

// マージンパネルは「長方形」選択時のみ有効（正円/スーパー楕円ではディム表示）
function syncMarginPanelUI() {
    try {
        marginPanel.enabled = !!rbRectangle.value;
    } catch (_) { }
}
syncMarginPanelUI();

    function getMarginValues() {
        try {
            if (cbMarginSquare && cbMarginSquare.value) {
                return { mx: 0, my: 0 };
            }
        } catch (_) { }

        var mv = parseFloat(marginVInput && marginVInput.text);
        var mh = parseFloat(marginHInput && marginHInput.text);
        if (isNaN(mv)) mv = 0;
        if (isNaN(mh)) mh = 0;
        // 連動時は左右=上下
        try {
            if (cbMarginLink && cbMarginLink.value) mh = mv;
        } catch (_) {}
        return { mx: mh, my: mv };
    }

    // パネル：角丸（形状パネルの直下）
    var roundPanel = leftCol.add('panel', undefined, L('roundTitle'));
    roundPanel.orientation = 'column';
    roundPanel.alignChildren = ['left', 'top'];
    roundPanel.margins = [15, 20, 15, 10];

    var roundRow = roundPanel.add('group');
    roundRow.orientation = 'row';
    roundRow.alignChildren = ['left', 'center'];
    // Checkbox before the numeric field (UI only for now)
    var cbRoundEnable = roundRow.add('checkbox', undefined, '');
    var __lastRoundValue = '2'; // will be updated after roundInput is created

    var roundInput = roundRow.add('edittext', undefined, '2');
    roundInput.characters = 4;
    __lastRoundValue = String(roundInput.text);
    roundRow.add('statictext', undefined, __rulerLabel);

    // --- Pill shape option: now on its own row ---
    var pillRow = roundPanel.add('group');
    pillRow.orientation = 'row';
    pillRow.alignChildren = ['left', 'center'];
    var cbPill = pillRow.add('checkbox', undefined, L('pillShape'));
    cbPill.value = false;

    // 角丸パネルは「長方形」選択時のみ有効（正円/スーパー楕円ではディム表示）
    function syncRoundPanelUI() {
        try {
            roundPanel.enabled = !!rbRectangle.value;
        } catch (_) { }
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
                    try { roundInput.text = String(__lastRoundValue || roundInput.text || '0'); } catch (_) {}
                }
                roundInput.enabled = true;
            } else {
                // ON→OFF で値を保存して無効化
                try { __lastRoundValue = String(roundInput.text); } catch (_) {}
                roundInput.enabled = false;
            }
        } catch (_) { }
    }

    // 初期状態
    try { syncRoundPanelUI(); } catch (_) {}

    // イベント：角丸ON/OFF
    cbRoundEnable.onClick = function () {
        try { syncRoundInputsUI(); } catch (_) {}
        try {
            if (cbRoundEnable.value && !cbPill.value) {
                roundInput.active = true;
            }
        } catch (_) {}
        if (typeof updatePreview === 'function') updatePreview();
    };
    cbRoundEnable.onChanging = cbRoundEnable.onClick;

    // イベント：ピル
    cbPill.onClick = function () {
        // pill ON のときは入力を無効化（半径は自動）
        try { syncRoundInputsUI(); } catch (_) {}
        if (typeof updatePreview === 'function') updatePreview();
    };
    cbPill.onChanging = cbPill.onClick;

    // 入力：矢印キー増減 + プレビュー
    changeValueByArrowKey(roundInput, function () {
        try { __lastRoundValue = String(roundInput.text); } catch (_) {}
        if (typeof updatePreview === 'function') updatePreview();
    });


    // パネル：グループ（スケールの直下）
    var pnlGroup = leftCol.add('panel', undefined, L('groupPanel'));
    pnlGroup.orientation = 'column';
    pnlGroup.alignChildren = ['fill', 'top'];
    pnlGroup.margins = [15, 20, 15, 10];
    var cbGroup = pnlGroup.add('checkbox', undefined, L('groupWithText'));
    cbGroup.value = true;

    var cbExclude = pnlGroup.add('checkbox', undefined, L('exclude'));
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
            } catch (_) { }
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
    var pnlOffset = rightCol.add('panel', undefined, L('axisPanel') + '（' + __rulerLabel + '）');
    pnlOffset.orientation = 'column';
    pnlOffset.alignChildren = ['fill', 'top'];
    pnlOffset.margins = [15, 20, 15, 10];

    // X と Y を横並びに
    var grpOffset = pnlOffset.add('group');
    grpOffset.orientation = 'row';
    grpOffset.alignChildren = ['left', 'center'];

    grpOffset.add('statictext', undefined, 'X');
    var offsetXInput = grpOffset.add('edittext', undefined, '0');
    offsetXInput.characters = 4;

    grpOffset.add('statictext', undefined, 'Y');
    var offsetYInput = grpOffset.add('edittext', undefined, '0');
    offsetYInput.characters = 4;

// パネル：種別（塗り / 線 / 線幅を1行に） — 右カラム最下部へ移動
var pnlKind = rightCol.add('panel', undefined, L('kindPanel'));
pnlKind.orientation = 'column';
pnlKind.alignChildren = ['fill', 'top'];
pnlKind.margins = [15, 20, 15, 10];

// 1行：塗り / 線 / 線幅
var grpKind = pnlKind.add('group');
grpKind.orientation = 'row';
grpKind.alignChildren = ['left', 'center'];
grpKind.spacing = 10;

var rbFill = grpKind.add('radiobutton', undefined, L('fill'));
var rbStroke = grpKind.add('radiobutton', undefined, L('stroke'));
rbFill.value = true; // デフォルトは塗り

var strokeWInput = grpKind.add('edittext', undefined, '1');
strokeWInput.characters = 4;
grpKind.add('statictext', undefined, __strokeLabel);

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
        } catch (_) { }
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
    var pnlColor = rightCol.add('panel', undefined, L('colorPanel'));
    pnlColor.orientation = 'column';
    pnlColor.alignChildren = ['fill', 'top'];
    pnlColor.margins = [15, 20, 15, 10];

    // パネル：不透明度（右カラム）
    var pnlOpacity = rightCol.add('panel', undefined, L('opacityPanel'));
    pnlOpacity.orientation = 'column';
    pnlOpacity.alignChildren = ['fill', 'top'];
    pnlOpacity.margins = [15, 20, 15, 10];

    var grpOpacity = pnlOpacity.add('group');
    grpOpacity.orientation = 'row';
    var cbOpacityApply = grpOpacity.add('checkbox', undefined, '');
    cbOpacityApply.value = false;

    var opacityInput = grpOpacity.add('edittext', undefined, '60');
    opacityInput.characters = 3; // 0-100
    opacityInput.enabled = cbOpacityApply.value;
    grpOpacity.add('statictext', undefined, '%');

    // 色モード（ラジオ）
    var grpMode = pnlColor.add('group');
    grpMode.orientation = 'column';
    grpMode.alignChildren = ['left', 'top'];

    var rbTextColor = grpMode.add('radiobutton', undefined, L('textColorRef'));

    var rbBlack = grpMode.add('radiobutton', undefined, L('black'));
    var rbWhite = grpMode.add('radiobutton', undefined, L('white'));
    var rbCMYK = grpMode.add('radiobutton', undefined, L('cmyk'));

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
        et.characters = 4;
        return et;
    }

    var fillC = addCMYKColumn(grpCMYK, 'C');
    var fillM = addCMYKColumn(grpCMYK, 'M');
    var fillY = addCMYKColumn(grpCMYK, 'Y');
    var fillK = addCMYKColumn(grpCMYK, 'K');

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
        } catch (_) { }
    }
    syncColorUI();

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
        alert("ドキュメントが開かれていません。\nNo document is open.");
        return;
    }

    // カラーのデフォルトは「ブラック」
    try {
        rbBlack.value = true;
        rbTextColor.value = rbWhite.value = rbCMYK.value = false;
        if (typeof syncColorUI === 'function') syncColorUI();
    } catch (_) { }

    var sel = app.selection;
    if (!sel || sel.length === 0) {
        alert("オブジェクトを選択してください。\nPlease select an object.");
        return;
    }

    // テキストがあれば従来ロジック（テキスト優先）
    var textItem = findFirstTextItem(sel);
    var isTextMode = !!textItem;

    // テキストが無い場合：選択オブジェクト（複数なら矩形合成）を対象にする
    var targetItem = isTextMode ? textItem : null;

    // 非テキスト時は「1文字」や「テキストカラー参照」を無効化して破綻を避ける
    if (!isTextMode) {
        try {
            cbOneChar.value = false;
            cbOneChar.enabled = false;
        } catch (_) { }
        try {
            // テキストカラー参照が選ばれていたらブラックへフォールバック
            if (rbTextColor && rbTextColor.value) {
                rbTextColor.value = false;
                rbBlack.value = true;
            }
            rbTextColor.enabled = false;
        } catch (_) { }

        // 対象は選択の先頭（move/place の基準）
        try {
            targetItem = sel[0];
        } catch (e) {
            targetItem = null;
        }

        if (!targetItem) {
            alert("対象オブジェクトが取得できません。\nCould not resolve target object.");
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

    // プレビュー管理
    var previewCircle = null;

    function removePreview() {
        try {
            if (previewCircle) previewCircle.remove();
        } catch (e) { }
        previewCircle = null;
    }

    // --- 見た目寸法キャッシュ（プレビュー最適化） ---
    var __measured = null; // {left, top, right, bottom, w, h}
    function measureTextVisualBoundsOnce() {
        if (__measured) return __measured;
        // 初回のみ、見た目寸法を取得（アウトライン→bounds→破棄）
        try {
            var dup = textItem.duplicate();
            var outlinedGroup = dup.createOutline();
            var ogb = outlinedGroup.geometricBounds; // [l, t, r, b]
            __measured = {
                left: ogb[0],
                top: ogb[1],
                right: ogb[2],
                bottom: ogb[3],
                w: ogb[2] - ogb[0],
                h: ogb[1] - ogb[3]
            };
        } catch (_) {
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
            try {
                outlinedGroup.remove();
            } catch (e1) { }
            try {
                dup.remove();
            } catch (e2) { }
        }
        return __measured;
    }

    // 非テキスト時：選択オブジェクトの矩形を合成（可能なら visibleBounds を優先）
    // - visibleBounds: 線幅など見た目を含みやすい
    // - geometricBounds: フォールバック
    function measureSelectionBounds() {
        var b = null; // [l,t,r,b]
        try {
            for (var i = 0; i < sel.length; i++) {
                var it = sel[i];
                if (!it) continue;

                var gb = null;
                try {
                    if (it.visibleBounds) gb = it.visibleBounds;
                } catch (_) { }
                if (!gb) {
                    try { gb = it.geometricBounds; } catch (_) { }
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
    var __didInitMarginDefault = false;
    function initMarginDefaultFromSelection() {
        if (__didInitMarginDefault) return;
        __didInitMarginDefault = true;

        var ww = 0, hh = 0;
        try {
            if (isTextMode) {
                // テキスト：通常は見た目寸法キャッシュ、未計測なら幾何境界
                var m2 = null;
                try { m2 = measureTextVisualBoundsOnce(); } catch (_) { m2 = null; }
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
                    try { if (targetItem.visibleBounds) gbb2 = targetItem.visibleBounds; } catch (_) { }
                    if (!gbb2) {
                        try { gbb2 = targetItem.geometricBounds; } catch (_) { }
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
        } catch (_) { }

        try { syncMarginUI(); } catch (_) { }
    }

    // 角丸のデフォルト：選択オブジェクトの短辺の 1/5
    var __didInitRoundDefault = false;
    function initRoundDefaultFromSelection() {
        if (__didInitRoundDefault) return;
        __didInitRoundDefault = true;

        var ww = 0, hh = 0;
        try {
            if (isTextMode) {
                var m2 = null;
                try { m2 = measureTextVisualBoundsOnce(); } catch (_) { m2 = null; }
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
                    try { if (targetItem.visibleBounds) gbb2 = targetItem.visibleBounds; } catch (_) { }
                    if (!gbb2) {
                        try { gbb2 = targetItem.geometricBounds; } catch (_) { }
                    }
                    bb = [gbb2[0], gbb2[1], gbb2[2], gbb2[3]];
                }
                ww = bb[2] - bb[0];
                hh = bb[1] - bb[3];
            }
        } catch (_) {
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
        } catch (_) { }
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
                try { if (targetItem.visibleBounds) gbb = targetItem.visibleBounds; } catch (_) { }
                if (!gbb) {
                    try { gbb = targetItem.geometricBounds; } catch (_) { }
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
        } catch (_) { }

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

    // Live Effect: Round Corners（ライブ効果で角丸を適用）
    function applyRoundCornersLive(item, r) {
        try {
            if (!item || !r || r <= 0) return;
            var xml = '<LiveEffect name="Adobe Round Corners"><Dict data="R radius ' + r + ' "/></LiveEffect>';
            item.applyEffect(xml);
        } catch (_) { }
    }

    function updatePreview() {
        removePreview();
        var p = computeParams();
        var circleLeft = p.cx - p.d / 2;
        var circleTop = p.cy + p.d / 2;

        // 形状作成：正円 / スーパー楕円 / 長方形
        if (rbRectangle.value) {
            // 長方形：中心合わせ（cx,cy）
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
            } catch (_) { }

            if (pillOn) {
                // ピル形状：ライブ角丸（半径=高さ/2）。
                // 横幅は「元のオブジェクト幅（baseW）に対して、左右の半円（直径=rectH）がかからない」最小値にする。
                // pillW = max(現在の長方形幅 rectW, baseW*scale + rectH)
                var pr = p.rectH / 2;
                var minW = (Math.abs(p.baseW) * (p.scale || 1)) + p.rectH; // baseW + 2R
                var pillW = Math.max(p.rectW, minW);
                var pillLeft = p.cx - pillW / 2;

                previewCircle = doc.pathItems.rectangle(rectTop, pillLeft, pillW, p.rectH);
                applyRoundCornersLive(previewCircle, pr);
            } else if (roundOn && r > 0) {
                // 角丸長方形：パス自体を角丸で作成
                previewCircle = doc.pathItems.roundedRectangle(rectTop, rectLeft, p.rectW, p.rectH, r, r);
            } else {
                // 通常の長方形
                previewCircle = doc.pathItems.rectangle(rectTop, rectLeft, p.rectW, p.rectH);
            }
        } else {
            // まずは通常の円（Ellipse）を作成し、その形状を必要に応じて変形する
            previewCircle = doc.pathItems.ellipse(circleTop, circleLeft, p.d, p.d);

            // 形状変更（スーパー楕円）
            if (rbSuperEllipse.value) {
                var n = 2.5; // スーパー楕円の指数（固定。後でUI化）
                morphPathToSuperellipse(previewCircle, p.cx, p.cy, p.d, p.d, n);
            }
        }
        // 塗り色：カラーpanelに従う（ブラック／ホワイト／CMYK）

        // スタイル適用：種別（塗り/線）とカラー設定
        try {
            // updatePreview() 内の applyColor(function(col){...}) を呼ぶ直前の分岐群
            var applyColor = function (setter) {
                // ★ 追加: テキストカラーを参照
                if (rbTextColor.value) {
                    var tcol = null;
                    try {
                        tcol = textItem.textRange.characterAttributes.fillColor;
                    } catch (e) { }
                    if (tcol) {
                        setter(tcol);
                        return;
                    }
                    // 取得できなかった場合はブラック等にフォールバック
                }

                // 既存分岐（ブラック/ホワイト/CMYK）
                if (rbBlack.value || rbTextColor.value) {
                    var kcol = new GrayColor();
                    kcol.gray = 100;
                    setter(kcol);
                } else if (rbWhite.value) {
                    var wcol = new GrayColor();
                    wcol.gray = 0;
                    setter(wcol);
                } else if (rbCMYK.value) {
                    var c = Math.min(100, Math.max(0, parseInt(fillC.text, 10) || 0));
                    var m = Math.min(100, Math.max(0, parseInt(fillM.text, 10) || 0));
                    var y = Math.min(100, Math.max(0, parseInt(fillY.text, 10) || 0));
                    var k = Math.min(100, Math.max(0, parseInt(fillK.text, 10) || 0));
                    var cmyk = new CMYKColor();
                    cmyk.cyan = c;
                    cmyk.magenta = m;
                    cmyk.yellow = y;
                    cmyk.black = k;
                    setter(cmyk);
                }
            };

            if (rbFill.value) {
                previewCircle.filled = true;
                previewCircle.stroked = false;
                applyColor(function (col) {
                    previewCircle.fillColor = col;
                });
            } else { // 線
                previewCircle.filled = false;
                previewCircle.stroked = true;
                var sw = parseFloat(strokeWInput.text);
                if (isNaN(sw) || sw < 0) sw = 1;
                previewCircle.strokeWidth = sw;
                applyColor(function (col) {
                    previewCircle.strokeColor = col;
                });
            }
        } catch (e) {
            // フォールバック：塗り20%グレー、線OFF
            try {
                var g2 = new GrayColor();
                g2.gray = 20;
                previewCircle.fillColor = g2;
                previewCircle.filled = true;
                previewCircle.stroked = false;
            } catch (_) { }
        }

        // 不透明度の適用（チェックONなら%値、OFFなら100%）
        try {
            var opv = parseFloat(opacityInput.text);
            if (isNaN(opv)) opv = 100;
            opv = Math.max(0, Math.min(100, opv));
            previewCircle.opacity = cbOpacityApply.value ? opv : 100;
        } catch (_) { }

        // テキストまたは対象の直下（背面）へ
        try {
            var base = isTextMode ? textItem : targetItem;
            previewCircle.move(base, ElementPlacement.PLACEAFTER);
        } catch (e) {
            try {
                previewCircle.zOrder(ZOrderMethod.SENDTOBACK);
            } catch (_) { }
        }
        // 即時描画更新
        try {
            app.redraw();
        } catch (_) { }
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
        try {
            removePreview();
        } catch (_) { }

        try {
            dlg.close(0);
        } catch (_) { }
    };

    // 初期プレビュー
    try { initMarginDefaultFromSelection(); } catch (_) { }
    try { initRoundDefaultFromSelection(); } catch (_) { }
    updatePreview();

    // ダイアログ表示時：倍率フィールドをアクティブに
    try {
        dlg.onShow = (function (prev) {
            return function () {
                try { if (typeof prev === 'function') prev(); } catch (_) { }
                try { scaleInput.active = true; } catch (_) { }
            };
        })(dlg.onShow);
    } catch (_) { }

    // 最終確定時にスタイルを再適用するヘルパ
    function applyStyleToItem(item) {
        if (!item) return;
        // 色適用ロジック（updatePreview 内の applyColor と同等）
        function _applyColor(setter) {
            // テキストカラー参照
            if (rbTextColor.value) {
                var tcol = null;
                try {
                    tcol = textItem.textRange.characterAttributes.fillColor;
                } catch (e) { }
                if (tcol) {
                    setter(tcol);
                    return;
                }
            }
            if (rbBlack.value || rbTextColor.value) {
                var kcol = new GrayColor();
                kcol.gray = 100;
                setter(kcol);
            } else if (rbWhite.value) {
                var wcol = new GrayColor();
                wcol.gray = 0;
                setter(wcol);
            } else if (rbCMYK.value) {
                var c = Math.min(100, Math.max(0, parseInt(fillC.text, 10) || 0));
                var m = Math.min(100, Math.max(0, parseInt(fillM.text, 10) || 0));
                var y = Math.min(100, Math.max(0, parseInt(fillY.text, 10) || 0));
                var k = Math.min(100, Math.max(0, parseInt(fillK.text, 10) || 0));
                var cmyk = new CMYKColor();
                cmyk.cyan = c;
                cmyk.magenta = m;
                cmyk.yellow = y;
                cmyk.black = k;
                setter(cmyk);
            }
        }
        if (rbFill.value) {
            item.filled = true;
            item.stroked = false;
            _applyColor(function (col) {
                item.fillColor = col;
            });
        } else {
            item.filled = false;
            item.stroked = true;
            var sw = parseFloat(strokeWInput.text);
            if (isNaN(sw) || sw < 0) sw = 1;
            item.strokeWidth = sw;
            _applyColor(function (col) {
                item.strokeColor = col;
            });
        }
        // 不透明度
        try {
            var opv = parseFloat(opacityInput.text);
            if (isNaN(opv)) opv = 100;
            opv = Math.max(0, Math.min(100, opv));
            item.opacity = cbOpacityApply.value ? opv : 100;
        } catch (_) { }
    }

    if (dlg.show() !== 1) {
        removePreview();
        return;
    }

    // OK：プレビューを確定。必要ならグループ化＋中マド処理
    var g = null;
    if (cbGroup.value) {
        try {
            var baseItem = isTextMode ? textItem : targetItem;
            var parent = baseItem.parent;
            g = parent.groupItems.add();
            // 先に空グループを対象の直後（同じスタック位置）に置く
            try {
                g.move(baseItem, ElementPlacement.PLACEAFTER);
            } catch (_) { }
            // グループに追加（順序は一旦どちらでもOK）
            previewCircle.move(g, ElementPlacement.PLACEATEND);
            baseItem.move(g, ElementPlacement.PLACEATEND);
            // 念のため、グループ内で円を最背面＝テキストの背面へ固定
            try {
                previewCircle.zOrder(ZOrderMethod.SENDTOBACK);
            } catch (_) { }
        } catch (e) {
            alert("グループ化に失敗しました: " + e);
        }
    }
    // グループ化しない場合は、念のため対象の直後（背面）へ再配置
    if (!cbGroup.value) {
        try {
            var baseItem2 = isTextMode ? textItem : targetItem;
            previewCircle.move(baseItem2, ElementPlacement.PLACEAFTER);
        } catch (_) { }
    }
    // 正円のみライブシェイプに変換（Convert to Shape）し、外観を再適用
    try {
        if (previewCircle) {
            // スーパー楕円は PathItem のため Convert to Shape は不要（実行すると崩れる可能性がある）
            if (rbPerfectCircle.value) {
                app.selection = null;
                previewCircle.selected = true;
                app.executeMenuCommand('Convert to Shape');
                // 変換後のオブジェクト（選択状態の先頭）を取得
                var converted = null;
                try {
                    converted = app.selection && app.selection.length ? app.selection[0] : null;
                } catch (_) { }
                if (!converted) converted = previewCircle; // フォールバック
                // 外観を再適用（線が消える/初期化されるのを防ぐ）
                applyStyleToItem(converted);
                app.selection = null;
            } else {
                // スーパー楕円：外観はそのまま／念のため再適用
                applyStyleToItem(previewCircle);
            }
        }
    } catch (e) {
        // 失敗しても処理継続
    }
    // 中マド処理（Pathfinder Exclude）：グループが生成できている場合のみ適用
    if (cbExclude.value && g) {
        try {
            app.selection = null;
            g.selected = true;
            app.executeMenuCommand('Live Pathfinder Exclude');
        } catch (e) {
            alert("中マド処理（Exclude）の適用に失敗しました: " + e);
        } finally {
            try {
                app.selection = null;
            } catch (_) { }
        }
    }

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

// run
try {
    main();
} catch (err) {
    alert("エラーが発生しました: " + err);
}