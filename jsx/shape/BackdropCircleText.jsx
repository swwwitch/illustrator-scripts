#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

// バージョン / Version
var SCRIPT_VERSION = "v1.0";

/*
【概要 / Overview】
選択したテキストの背面に、見た目寸法に基づく「正円」を自動生成して背面配置します。直径はテキストの幅・高さのうち大きい方を基準に算出し、テキストの中心に揃えます。OK 時に作成した円は［Convert to Shape］でライブシェイプ化されます。

- 対象 / Target: pointText または areaText（複数選択時は最初のテキストを対象）
- 位置 / Position: テキスト直後（背面側 / PLACEAFTER）に配置。X/Y オフセットはアプリの rulerType 単位（負値可）
- スケール / Scale: ％入力（例: 100 = 等倍）。専用パネルに「倍率」「1文字」チェックを配置
- 種別 / Kind: 塗り or 線（線のときは線幅を指定）。線選択中は「中マド処理」を自動でディム＆OFF
- カラー / Color: ブラック / ホワイト / CMYK（各値入力）/ テキストカラーを参照（取得失敗時はフォールバック）
- 不透明度 / Opacity: ［適用］チェック＋数値（%）
- グループ / Group: 「テキストとグループ化」「中マド処理（Exclude）」に対応（Exclude を ON した場合は自動でグループ化）
- プレビュー / Preview: 入力変更・矢印キー（↑↓ / Shift=±10 / Option=±0.1）で即時更新。計測結果はキャッシュして負荷を軽減
- ローカライズ / Localization: すべてのラベルは日英対応。ダイアログタイトルに SCRIPT_VERSION を併記

- 2025-11-11: 初版。テキスト背面に正円を自動作成・配置、OK 時に Convert to Shape。
- 2025-11-11: 座標パネルのタイトルへ単位を表示、X/Y 負値を許容、矢印キーでの増減に対応。
- 2025-11-11: スケールを 1.0 倍入力 → 100% 入力に変更。スケール専用パネルを追加し「1文字」チェックを直下に配置。
- 2025-11-11: 種別パネルを右カラム下部へ移動。線選択時に「中マド処理」をディム＆OFF。線幅の単位は入力欄の後ろに表示。
- 2025-11-11: カラーに「テキストカラーを参照」を追加。RGB UI をいったん削除し、CMYK 入力を整理。
- 2025-11-11: 不透明度パネルを追加（［適用］＋％）。
- 2025-11-11: プレビュー最適化：見た目寸法の再アウトライン化は初回のみ計測してキャッシュ。
- 2025-11-11: 全 UI ラベルをローカライズ対応（Axis / Scale / Exclude など）。
- 2025-11-11: ダイアログのウィンドウ位置（bounds）を記憶・復元する機能を追加。
（更新日 / Updated: 2025-11-11）
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
        ja: "テキストの背面に正円を作成",
        en: "Create Circle Behind Text"
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
var __DIALOG_OPACITY = 0.95;

// 透明度設定（tryガード付き）
function setDialogOpacity(dlg, opacityValue) {
    try {
        dlg.opacity = opacityValue;
    } catch (e) {}
}

/**
 * 初回表示時だけ位置をずらすヘルパ
 * すでに保存済みのboundsがある場合は、ズレ続けを避けるためシフトしない
 */
function shiftDialogPositionOnceOnShow(dlg, dx, dy) {
    dlg.onShow = (function(prev) {
        return function() {
            try {
                if (typeof prev === 'function') prev();
            } catch (_) {}
            try {
                var l = dlg.location;
                dlg.location = [l[0] + (dx | 0), l[1] + (dy | 0)];
            } catch (_) {}
        };
    })(dlg.onShow);
}

// --- Arrow-key value changer for EditText ---
function changeValueByArrowKey(editText, onUpdate) {
    if (!editText) return;

    // フォールバック：値が変わったら常にプレビュー更新（単純入力やスクロールでも反応）
    editText.onChanging = function() {
        try {
            if (typeof onUpdate === 'function') onUpdate();
        } catch (_) {}
    };

    // ScriptUI の addEventListener が使える場合は、矢印キー専用の高速インクリメント処理
    if (editText.addEventListener) {
        editText.addEventListener('keydown', function(event) {
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
            } catch (_) {}

            // 既定動作の抑止（可能な環境のみ）
            try {
                if (event && event.preventDefault) event.preventDefault();
            } catch (_) {}
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
    // 現在のアプリ環境の単位ラベルを取得
    var __rulerCode = app.preferences.getIntegerPreference('rulerType');
    var __strokeCode = app.preferences.getIntegerPreference('strokeUnits');
    var __rulerLabel = getUnitLabel(__rulerCode, 'rulerType');
    var __strokeLabel = getUnitLabel(__strokeCode, 'strokeUnits');
    // 2カラムコンテナ

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

    // パネル：座標調整
    var pnlOffset = leftCol.add('panel', undefined, L('axisPanel') + '（' + __rulerLabel + '）');
    pnlOffset.orientation = 'column';
    pnlOffset.alignChildren = ['fill', 'top'];
    pnlOffset.margins = [15, 20, 15, 10];

    // 新パネル：スケール（座標パネルの直下）
    var pnlScale = leftCol.add('panel', undefined, L('scalePanel'));
    pnlScale.orientation = 'column';
    pnlScale.alignChildren = ['fill', 'top'];
    pnlScale.margins = [15, 20, 15, 10];

    // X と Y を縦並びに
    var grpOffset = pnlOffset.add('group');
    grpOffset.orientation = 'column';
    grpOffset.alignChildren = ['fill', 'top'];

    var rowX = grpOffset.add('group');
    rowX.orientation = 'row';
    rowX.add('statictext', undefined, 'X');
    var offsetXInput = rowX.add('edittext', undefined, '0');
    offsetXInput.characters = 4;

    var rowY = grpOffset.add('group');
    rowY.orientation = 'row';
    rowY.add('statictext', undefined, 'Y');
    var offsetYInput = rowY.add('edittext', undefined, '0');
    offsetYInput.characters = 4;

    // 倍率（%表示）
    var grpScale = pnlScale.add('group');
    grpScale.add('statictext', undefined, L('magnification'));
    var scaleInput = grpScale.add('edittext', undefined, '100');
    scaleInput.characters = 4;
    grpScale.add('statictext', undefined, '%');

    // 1文字チェック（倍率の直下）
    var grpOneChar = pnlScale.add('group');
    grpOneChar.orientation = 'row';
    grpOneChar.alignChildren = ['left', 'center'];
    var cbOneChar = grpOneChar.add('checkbox', undefined, L('oneChar'));
    cbOneChar.value = false;

    // パネル：グループ
    var pnlGroup = leftCol.add('panel', undefined, L('groupPanel'));
    pnlGroup.orientation = 'column';
    pnlGroup.alignChildren = ['fill', 'top'];
    pnlGroup.margins = [15, 20, 15, 10];
    var cbGroup = pnlGroup.add('checkbox', undefined, L('groupWithText'));
    cbGroup.value = true;

    var cbExclude = pnlGroup.add('checkbox', undefined, L('exclude'));
    cbExclude.value = false;

    // ロジック：中マド=ONなら自動で「テキストとグループ化」をON
    cbExclude.onClick = function() {
        if (cbExclude.value) cbGroup.value = true;
        if (typeof updatePreview === 'function') updatePreview();
    };
    cbExclude.onChanging = cbExclude.onClick;

    // ロジック：「テキストとグループ化」をOFFにしたら「中マド処理」もOFF
    cbGroup.onClick = function() {
        if (!cbGroup.value) cbExclude.value = false;
        if (typeof updatePreview === 'function') updatePreview();
    };
    cbGroup.onChanging = cbGroup.onClick;

    // パネル：種別（塗り or 線） — 右カラム最下部へ移動
    var pnlKind = rightCol.add('panel', undefined, L('kindPanel'));
    pnlKind.orientation = 'column';
    pnlKind.alignChildren = ['fill', 'top'];
    pnlKind.margins = [15, 20, 15, 10];

    var grpKind = pnlKind.add('group');
    grpKind.orientation = 'row';
    var rbFill = grpKind.add('radiobutton', undefined, L('fill'));
    var rbStroke = grpKind.add('radiobutton', undefined, L('stroke'));
    rbFill.value = true; // デフォルトは塗り

    var grpStrokeW = pnlKind.add('group');
    grpStrokeW.orientation = 'row';
    grpStrokeW.add('statictext', undefined, L('strokeWidth'));
    var strokeWInput = grpStrokeW.add('edittext', undefined, '1');
    strokeWInput.characters = 4;
    grpStrokeW.add('statictext', undefined, __strokeLabel);

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
        } catch (_) {}
    }
    syncKindUI();
    rbFill.onClick = function() {
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
        fillC.enabled = cmykOn;
        fillM.enabled = cmykOn;
        fillY.enabled = cmykOn;
        fillK.enabled = cmykOn;
        try {
            grpCMYK.visible = cmykOn;
        } catch (_) {}
    }
    syncColorUI();

    // ラジオと入力のイベントでプレビュー更新
    var rbList = [rbTextColor, rbBlack, rbWhite, rbCMYK];
    for (var i = 0; i < rbList.length; i++) {
        rbList[i].onClick = function() {
            syncColorUI();
            if (typeof updatePreview === 'function') updatePreview();
        };
    }
    fillC.onChanging = updatePreview;
    fillM.onChanging = updatePreview;
    fillY.onChanging = updatePreview;
    fillK.onChanging = updatePreview;
    cbOpacityApply.onClick = function() {
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

    // ドキュメントのカラーモードに応じてカラーUIの初期値を設定
    try {
        if (doc.documentColorSpace === DocumentColorSpace.CMYK) {
            rbCMYK.value = true;
            rbBlack.value = rbWhite.value = false;
        }
        if (typeof syncColorUI === 'function') syncColorUI();
    } catch (_) {}

    var sel = app.selection;
    if (!sel || sel.length === 0) {
        alert("テキストオブジェクトを選択してください。\nPlease select a text object.");
        return;
    }
    var textItem = findFirstTextItem(sel);
    if (!textItem) {
        alert("選択内にテキストオブジェクトが見つかりません。\nNo text object found in selection.");
        return;
    }

    // 自動判定：選択中テキストが「1文字」ならチェックON
    try {
        var raw = textItem.contents || "";
        // 改行・タブ・スペースを除外してカウント
        var count = String(raw).replace(/[\r\n\t\s]+/g, "").length;
        if (count === 1) {
            cbOneChar.value = true;
        }
    } catch (e) {}

    // プレビュー管理
    var previewCircle = null;

    function removePreview() {
        try {
            if (previewCircle) previewCircle.remove();
        } catch (e) {}
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
            } catch (e1) {}
            try {
                dup.remove();
            } catch (e2) {}
        }
        return __measured;
    }

    // 計算（UI状態に応じてサイズ・中心を返す）
    function computeParams() {
        // 基準寸法：
        //  - 1文字モード: アウトライン不要（フォントサイズベース）
        //  - 通常モード  : 初回のみアウトラインして見た目寸法をキャッシュ
        var left, top, right, bottom, w, h;
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

        var cx = left + w / 2;
        var cy = top - h / 2;

        var userScaleVal = parseFloat(scaleInput.text);
        if (isNaN(userScaleVal) || userScaleVal <= 0) userScaleVal = 100;
        userScaleVal = userScaleVal / 100; // percent → multiplier

        var d;
        if (cbOneChar.value) {
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
            var squareSide = Math.max(w, h);
            d = squareSide * Math.SQRT2 * userScaleVal;
        }

        // オフセット適用（負値可）
        var offX = parseFloat(offsetXInput.text);
        var offY = parseFloat(offsetYInput.text);
        if (!isNaN(offX)) cx += offX;
        if (!isNaN(offY)) cy += offY;

        return {
            cx: cx,
            cy: cy,
            d: d
        };
    }

    function updatePreview() {
        removePreview();
        var p = computeParams();
        var circleLeft = p.cx - p.d / 2;
        var circleTop = p.cy + p.d / 2;
        previewCircle = doc.pathItems.ellipse(circleTop, circleLeft, p.d, p.d);
        // 塗り色：カラーpanelに従う（ブラック／ホワイト／CMYK）

        // スタイル適用：種別（塗り/線）とカラー設定
        try {
            // updatePreview() 内の applyColor(function(col){...}) を呼ぶ直前の分岐群
            var applyColor = function(setter) {
                // ★ 追加: テキストカラーを参照
                if (rbTextColor.value) {
                    var tcol = null;
                    try {
                        tcol = textItem.textRange.characterAttributes.fillColor;
                    } catch (e) {}
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
                applyColor(function(col) {
                    previewCircle.fillColor = col;
                });
            } else { // 線
                previewCircle.filled = false;
                previewCircle.stroked = true;
                var sw = parseFloat(strokeWInput.text);
                if (isNaN(sw) || sw < 0) sw = 1;
                previewCircle.strokeWidth = sw;
                applyColor(function(col) {
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
            } catch (_) {}
        }

        // 不透明度の適用（チェックONなら%値、OFFなら100%）
        try {
            var opv = parseFloat(opacityInput.text);
            if (isNaN(opv)) opv = 100;
            opv = Math.max(0, Math.min(100, opv));
            previewCircle.opacity = cbOpacityApply.value ? opv : 100;
        } catch (_) {}

        // テキストの直下（背面）へ
        try {
            previewCircle.move(textItem, ElementPlacement.PLACEAFTER);
        } catch (e) {
            try {
                previewCircle.zOrder(ZOrderMethod.SENDTOBACK);
            } catch (_) {}
        }
        // 即時描画更新
        try {
            app.redraw();
        } catch (_) {}
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
    cancelBtn.onClick = function() {
        try {
            removePreview();
        } catch (_) {}

        try {
            dlg.close(0);
        } catch (_) {}
    };

    // 初期プレビュー
    updatePreview();

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
                } catch (e) {}
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
            _applyColor(function(col) {
                item.fillColor = col;
            });
        } else {
            item.filled = false;
            item.stroked = true;
            var sw = parseFloat(strokeWInput.text);
            if (isNaN(sw) || sw < 0) sw = 1;
            item.strokeWidth = sw;
            _applyColor(function(col) {
                item.strokeColor = col;
            });
        }
        // 不透明度
        try {
            var opv = parseFloat(opacityInput.text);
            if (isNaN(opv)) opv = 100;
            opv = Math.max(0, Math.min(100, opv));
            item.opacity = cbOpacityApply.value ? opv : 100;
        } catch (_) {}
    }

    if (dlg.show() !== 1) {
        removePreview();
        return;
    }

    // OK：プレビューを確定。必要ならグループ化＋中マド処理
    var g = null;
    if (cbGroup.value) {
        try {
            var parent = textItem.parent;
            g = parent.groupItems.add();
            // 先に空グループをテキストの直後（同じスタック位置）に置く
            try {
                g.move(textItem, ElementPlacement.PLACEAFTER);
            } catch (_) {}
            // グループに追加（順序は一旦どちらでもOK）
            previewCircle.move(g, ElementPlacement.PLACEATEND);
            textItem.move(g, ElementPlacement.PLACEATEND);
            // 念のため、グループ内で円を最背面＝テキストの背面へ固定
            try {
                previewCircle.zOrder(ZOrderMethod.SENDTOBACK);
            } catch (_) {}
        } catch (e) {
            alert("グループ化に失敗しました: " + e);
        }
    }
    // グループ化しない場合は、念のためテキストの直後（背面）へ再配置
    if (!cbGroup.value) {
        try {
            previewCircle.move(textItem, ElementPlacement.PLACEAFTER);
        } catch (_) {}
    }
    // 正円をライブシェイプに変換（Convert to Shape）し、外観を再適用
    try {
        if (previewCircle) {
            app.selection = null;
            previewCircle.selected = true;
            app.executeMenuCommand('Convert to Shape');
            // 変換後のオブジェクト（選択状態の先頭）を取得
            var converted = null;
            try {
                converted = app.selection && app.selection.length ? app.selection[0] : null;
            } catch (_) {}
            if (!converted) converted = previewCircle; // フォールバック
            // 外観を再適用（線が消える/初期化されるのを防ぐ）
            applyStyleToItem(converted);
            app.selection = null;
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
            } catch (_) {}
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