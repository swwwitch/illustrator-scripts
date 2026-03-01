#targetengine "TabularizeEngine"
#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
  tabularize.jsx
  更新日: 2026-02-10

  選択オブジェクトを「表」として解釈し、表組み用の塗りと線（横ケイ/縦ケイ）を生成します。
  - プレビュー機能: ダイアログ操作時にリアルタイムでプレビューを表示
  - 塗り: 通常／ゼブラ／行方向に連結／ヘッダー行のみ
  - オプション: ガター（rulerType単位）、1行目をヘッダー行に
  - 線: 縦ケイ（なし／列間のみ／すべて）＋ ガター0時は連結描画
  - プリセット: 代表的な組み合わせを一括適用
  - ダイアログ値はセッション内で復元（Illustrator再起動でリセット）

  Interpret the selection as a table grid and generate fills and rules (horizontal/vertical).
  - Live preview: Shows a preview in real time as you change dialog options
  - Fill: normal / zebra / join by row / header-only
  - Options: gutter (rulerType units), treat first row as header
  - Rules: vertical modes (none / gaps only / all) + continuous drawing when gutter is 0
  - Presets: apply common combinations at once
  - Dialog values persist within the session (reset on Illustrator restart)
*/

// =========================
// セッション内の状態保持（Illustrator再起動でリセット） / Session-only persistence
// =========================
// Engine スコープ（#targetengine）で $.global を使って保持

$.global.__tabularizeState = $.global.__tabularizeState || {
    presetIndex: 0,
    // Options
    useGutter: true,
    gutterText: "",
    headerRow: true,
    // Fill
    doFill: false,
    zebra: false,
    fillJoinRow: false,
    fillHeaderOnly: false,
    // Lines
    doRule: true,
    vRuleMode: "gapsOnly" // 'none' | 'gapsOnly' | 'all'
};

(function () {
    // 単位変換用の定数 (ポイント換算) / Unit conversion constant (pt)
    var MM_TO_PT = 2.83464567;

    // バージョン / Version
    var SCRIPT_VERSION = "v1.1";

    // 言語判定 / Language detection
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var lang = getCurrentLang();

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialogTitle: {
            ja: "表組み化" + ' ' + SCRIPT_VERSION,
            en: "Tabularize" + ' ' + SCRIPT_VERSION
        },
        vRulePanel: {
            ja: "線",
            en: "Strokes"
        },
        fillPanel: {
            ja: "塗り",
            en: "Fill"
        },
        fillCheck: {
            ja: "塗り",
            en: "Fill"
        },
        fillOptionPanel: {
            ja: "オプション",
            en: "Options"
        },
        zebra: {
            ja: "ゼブラ",
            en: "Zebra"
        },
        fillJoinRow: {
            ja: "行方向に連結",
            en: "Join by row"
        },
        fillHeaderOnly: {
            ja: "ヘッダー行のみ",
            en: "Header only"
        },
        ruleCheck: {
            ja: "線",
            en: "Rules"
        },
        fillNone: {
            ja: "塗りなし",
            en: "No fill"
        },
        fillOnly: {
            ja: "塗りのみ",
            en: "Fill only"
        },
        fillAndRule: {
            ja: "塗りとケイ",
            en: "Fill + rules"
        },
        gutter: {
            ja: "ガター",
            en: "Gutter"
        },
        useGutter: {
            ja: "ガター",
            en: "Gutter"
        },
        headerRow: {
            ja: "1行目をヘッダー行にする",
            en: "Treat first row as header"
        },
        gapsOnly: {
            ja: "列間のみ",
            en: "Gaps only"
        },
        all: {
            ja: "すべて",
            en: "All"
        },
        none: {
            ja: "なし",
            en: "None"
        },
        vRuleLabel: {
            ja: "縦ケイ",
            en: "Vertical"
        },
        cancel: {
            ja: "キャンセル",
            en: "Cancel"
        },
        ok: {
            ja: "OK",
            en: "OK"
        },
        preset: {
            ja: "プリセット",
            en: "Preset"
        },
        alertOpenDoc: {
            ja: "ドキュメントを開いてください。",
            en: "Please open a document."
        },
        alertSelectObj: {
            ja: "オブジェクトを選択してください。",
            en: "Please select objects."
        },
        optionPanel: {
            ja: "オプション",
            en: "Options"
        }
    };

    function L(key) {
        var o = LABELS[key];
        if (!o) return key;
        return o[lang] || o.ja || key;
    }

    /* 単位ユーティリティ / Unit utilities */

    // --- 外部定義：共通単位マップ ---
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

    /**
     * 単位コードと設定キーから適切な単位ラベルを返す（Q/H分岐含む） / Get unit label (with Q/H)
     */
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

    /**
     * 単位コードから pt 換算係数を返す / Get pt factor from unit code
     */
    function getPtFactorFromUnitCode(code) {
        switch (code) {
            case 0: return 72.0;                        // in
            case 1: return 72.0 / 25.4;                 // mm
            case 2: return 1.0;                         // pt
            case 3: return 12.0;                        // pica
            case 4: return 72.0 / 2.54;                 // cm
            case 5: return 72.0 / 25.4 * 0.25;          // Q or H
            case 6: return 1.0;                         // px
            case 7: return 72.0 * 12.0;                 // ft/in
            case 8: return 72.0 / 25.4 * 1000.0;        // m
            case 9: return 72.0 * 36.0;                 // yd
            case 10: return 72.0 * 12.0;                // ft
            default: return 1.0;
        }
    }

    /**
     * rulerType の単位コードを取得 / Get rulerType unit code
     */
    function getRulerUnitCode() {
        return app.preferences.getIntegerPreference("rulerType");
    }

    /* 設定項目 / Settings */
    var lineWeightMM = 0.1; // 線の太さ (mm) ※細めが良い場合は0.1など
    var paddingMM = 0.0;    // テキストの左右に少し余白を持たせるか (mm)
    // ----------------

    var lineWeightPt = lineWeightMM * MM_TO_PT;
    var paddingPt = paddingMM * MM_TO_PT;
    var HEADER_LINE_WEIGHT_MM = 0.25; // チェックON時、上から1本目と2本目だけ太くする(mm)
    var headerLineWeightPt = HEADER_LINE_WEIGHT_MM * MM_TO_PT;

    // ドキュメントチェックと選択チェックをダイアログより前に移動
    // 横罫ガターなど単位周辺の初期値計算はこのまま

    /* ダイアログ位置・透明度 / Dialog position & opacity */
    var offsetX = 300;
    var offsetY = 0;
    var dialogOpacity = 0.98;

    function shiftDialogPosition(dlg, offsetX, offsetY) {
        var origOnShow = dlg.onShow;
        dlg.onShow = function () {
            if (origOnShow) try { origOnShow(); } catch (_) {}
            try {
                var currentX = dlg.location[0];
                var currentY = dlg.location[1];
                dlg.location = [currentX + offsetX, currentY + offsetY];
            } catch (_) { }
        };
    }

    function setDialogOpacity(dlg, opacityValue) {
        try {
            dlg.opacity = opacityValue;
        } catch (_) { }
    }

    /* ダイアログ / Dialog */
    var dlg = new Window('dialog', L('dialogTitle'));
    dlg.orientation = 'column';
    dlg.alignChildren = 'left';

    setDialogOpacity(dlg, dialogOpacity);
    shiftDialogPosition(dlg, offsetX, offsetY);

    /* ダイアログ状態の復元/保存 / Restore & save dialog state */
    function restoreDialogState(controls) {
        var st = $.global.__tabularizeState || {};
        // ガード：プリセット変更で手動へ戻す処理を抑制
        isApplyingPreset = true;
        try {
            // プリセット
            if (typeof st.presetIndex === 'number' && controls.ddPreset.items && controls.ddPreset.items.length > st.presetIndex) {
                controls.ddPreset.selection = st.presetIndex;
            }

            // Options
            if (typeof st.useGutter === 'boolean') controls.cbUseGutter.value = st.useGutter;
            if (typeof st.gutterText === 'string' && st.gutterText !== "") controls.etHGutter.text = st.gutterText;
            if (typeof st.headerRow === 'boolean') controls.cbHeader.value = st.headerRow;

            // Fill
            if (typeof st.doFill === 'boolean') controls.cbFill.value = st.doFill;
            if (typeof st.zebra === 'boolean') controls.cbZebra.value = st.zebra;
            if (typeof st.fillJoinRow === 'boolean') controls.cbFillJoinRow.value = st.fillJoinRow;
            if (typeof st.fillHeaderOnly === 'boolean') controls.cbFillHeaderOnly.value = st.fillHeaderOnly;

            // Lines
            if (typeof st.doRule === 'boolean') controls.cbRule.value = st.doRule;
            if (st.vRuleMode === 'none') {
                controls.rbVruleNone.value = true;
            } else if (st.vRuleMode === 'all') {
                controls.rbVruleAll.value = true;
            } else {
                controls.rbVruleGapsOnly.value = true;
            }

        } catch (_) {
        } finally {
            isApplyingPreset = false;
        }

        // 依存UIの反映
        try { controls.applyGutterEnabled(); } catch (_) { }
        try { controls.applyFillEnabled(); } catch (_) { }
        try { controls.applyRuleEnabled(); } catch (_) { }
        try { controls.applyFillJoinRow(); } catch (_) { }
        try { controls.applyFillHeaderOnly(); } catch (_) { }
    }

    function saveDialogState(controls) {
        $.global.__tabularizeState = $.global.__tabularizeState || {};
        var st = $.global.__tabularizeState;

        // プリセット
        st.presetIndex = (controls.ddPreset.selection) ? controls.ddPreset.selection.index : 0;

        // Options
        st.useGutter = !!controls.cbUseGutter.value;
        st.gutterText = String(controls.etHGutter.text || "");
        st.headerRow = !!controls.cbHeader.value;

        // Fill
        st.doFill = !!controls.cbFill.value;
        st.zebra = !!controls.cbZebra.value;
        st.fillJoinRow = !!controls.cbFillJoinRow.value;
        st.fillHeaderOnly = !!controls.cbFillHeaderOnly.value;

        // Lines
        st.doRule = !!controls.cbRule.value;
        st.vRuleMode = controls.rbVruleNone.value ? 'none' : (controls.rbVruleAll.value ? 'all' : 'gapsOnly');
    }

    /* プリセット / Preset */
    var gPreset = dlg.add('group');
    gPreset.orientation = 'row';
    gPreset.alignChildren = ['left', 'center'];
    var ddPreset = gPreset.add('dropdownlist', undefined, [
        (lang === 'ja') ? '（手動）' : '(Manual)',
        (lang === 'ja') ? '塗りON / 線OFF（ガター1mm）' : 'Fill ON / Rules OFF (Gutter 1mm)', // 1
        (lang === 'ja') ? '線ON（ガター2mm）/ 縦ケイ=すべて' : 'Rules ON (Gutter 2mm) / V=All', // 4
        (lang === 'ja') ? 'ヘッダー行のみ（塗り+線/ ガター0 / 縦ケイ=列間）' : 'Header only (Fill+Rules / Gutter 0 / V=Gaps)', // 6
        (lang === 'ja') ? '線ON（ガター1mm）/ 縦ケイ=列間' : 'Rules ON (Gutter 1mm) / V=Gaps', // 3
        (lang === 'ja') ? '線ON（ガター0）/ 縦ケイ=列間' : 'Rules ON (Gutter 0) / V=Gaps', // 2
        (lang === 'ja') ? '線ON（ガター0）/ 縦ケイ=なし' : 'Rules ON (Gutter 0) / V=None' // 5
    ]);
    ddPreset.selection = 0;
    // 初期状態は手動（＝何もしない）

    // 手動操作が入ったらプリセットを「手動」に戻す / Switch preset to Manual on any manual change
    var isApplyingPreset = false;
    function setPresetManual() {
        if (isApplyingPreset) return;
        if (ddPreset.selection && ddPreset.selection.index !== 0) {
            ddPreset.selection = 0;
        }
    }

    /* オプション / Options */
    var pOpt = dlg.add('panel', undefined, L('optionPanel'));
    pOpt.orientation = 'column';
    pOpt.alignChildren = 'left';
    pOpt.margins = [15, 20, 15, 10];

    /* ガター設定 / Gutter */
    var gGutter = pOpt.add('group');
    gGutter.orientation = 'row';
    gGutter.alignChildren = ['left', 'center'];

    // チェックOFF時はガター=0＆ディム表示 / When OFF, set gutter=0 and dim
    var cbUseGutter = gGutter.add('checkbox', undefined, L('useGutter'));
    cbUseGutter.value = true;

    var rulerUnitCode = getRulerUnitCode();
    var rulerFactorPt = getPtFactorFromUnitCode(rulerUnitCode);
    var rulerUnitLabel = getUnitLabel(rulerUnitCode, 'rulerType');
    // デフォルトガター：1mm 相当を rulerType に変換 / Default gutter ≈ 1mm in rulerType
    var defaultGutterMm = 1;
    var defaultGutterPt = defaultGutterMm * MM_TO_PT;
    var defaultGutterVal = defaultGutterPt / rulerFactorPt;

    // 表示用に丸め（pt/pxは整数、その他は小数1桁）
    if (rulerUnitLabel === 'pt' || rulerUnitLabel === 'px') {
        defaultGutterVal = Math.round(defaultGutterVal);
    } else {
        defaultGutterVal = Math.round(defaultGutterVal * 10) / 10;
    }

    var etHGutter = gGutter.add('edittext', undefined, String(defaultGutterVal)); // rulerType
    etHGutter.characters = 3;
    changeValueByArrowKey(etHGutter);

    var stGutterUnit = gGutter.add('statictext', undefined, rulerUnitLabel);

    // OFF→0固定 & ディム / ON→復帰
    var lastGutterText = etHGutter.text;
    function applyGutterEnabled() {
        if (!cbUseGutter.value) {
            lastGutterText = etHGutter.text;
            etHGutter.text = '0';
            etHGutter.enabled = false;
            stGutterUnit.enabled = false;
        } else {
            etHGutter.enabled = true;
            stGutterUnit.enabled = true;
            // 0のまま戻したくない場合は直前値に戻す
            if (etHGutter.text === '0' && lastGutterText && lastGutterText !== '0') {
                etHGutter.text = lastGutterText;
            }
        }
    }
    cbUseGutter.onClick = applyGutterEnabled;
    applyGutterEnabled();

    // ガター値をmm指定でセット（rulerTypeに変換して入力欄へ）
    function setGutterByMm(mmVal) {
        var val = (mmVal * MM_TO_PT) / rulerFactorPt;
        // 表示用に丸め（pt/pxは整数、その他は小数1桁）
        if (rulerUnitLabel === 'pt' || rulerUnitLabel === 'px') {
            val = Math.round(val);
        } else {
            val = Math.round(val * 10) / 10;
        }
        etHGutter.text = String(val);
    }

    /* 1行目をヘッダー行にする / Treat first row as header */
    var cbHeader = pOpt.add('checkbox', undefined, L('headerRow'));
    cbHeader.value = true;

    /* 2カラムレイアウト / Two-column layout */
    var gCols = dlg.add('group');
    gCols.orientation = 'row';
    gCols.alignChildren = ['fill', 'top'];

    var gLeft = gCols.add('group');
    gLeft.orientation = 'column';
    gLeft.alignChildren = 'fill';

    var gRight = gCols.add('group');
    gRight.orientation = 'column';
    gRight.alignChildren = 'fill';

    /* 塗り / Fill */
    var pFill = gLeft.add('panel', undefined, L('fillPanel'));
    pFill.orientation = 'column';
    pFill.alignChildren = 'left';
    pFill.margins = [15, 20, 15, 10];

    var gFill = pFill.add('group');
    gFill.orientation = 'row';
    gFill.alignChildren = ['left', 'center'];

    var cbFill = gFill.add('checkbox', undefined, L('fillCheck'));

    // デフォルト：塗りOFF
    cbFill.value = false;

    /* 塗りオプション / Fill options */
    var pFillOpt = pFill.add('panel', undefined, L('fillOptionPanel'));
    pFillOpt.orientation = 'column';
    pFillOpt.alignChildren = 'left';
    pFillOpt.margins = [15, 20, 15, 10];

    // 行方向に連結（UI）
    var cbFillJoinRow = pFillOpt.add('checkbox', undefined, L('fillJoinRow'));
    cbFillJoinRow.value = false;

    // ゼブラ（UI）
    var cbZebra = pFillOpt.add('checkbox', undefined, L('zebra'));
    cbZebra.value = false;

    // ヘッダー行のみ（UI）
    var cbFillHeaderOnly = pFillOpt.add('checkbox', undefined, L('fillHeaderOnly'));
    cbFillHeaderOnly.value = false;

    // 塗りOFFならゼブラ/行方向に連結/ヘッダー行のみ はディム表示
    function applyFillEnabled() {
        var on = !!cbFill.value;

        pFillOpt.enabled = on;
        cbZebra.enabled = on;
        cbFillJoinRow.enabled = on;
        cbFillHeaderOnly.enabled = on;

        if (!on) {
            cbZebra.value = false;
            cbFillJoinRow.value = false;
            cbFillHeaderOnly.value = false;
        }
    }

    // ON時：ガターを0にして横方向に連結 / When ON: force gutter=0 and join horizontally
    function applyFillJoinRow() {
        if (cbFillJoinRow.value) {
            cbUseGutter.value = false;
            applyGutterEnabled();
        }
    }
    cbFillJoinRow.onClick = applyFillJoinRow;

    // ON時：ガター0 + ヘッダーON + 塗りは1行目のみ
    function applyFillHeaderOnly() {
        if (cbFillHeaderOnly.value) {
            // ガターを0に
            cbUseGutter.value = false;
            applyGutterEnabled();

            // 1行目をヘッダーに
            cbHeader.value = true;

            // 競合回避：行方向に連結はOFF
            cbFillJoinRow.value = false;
        }
    }
    cbFillHeaderOnly.onClick = applyFillHeaderOnly;

    cbFill.onClick = function() {
        applyFillEnabled();
        // 塗りOFFになったら状態をリセット
        if (!cbFill.value) {
            cbZebra.value = false;
            cbFillJoinRow.value = false;
            cbFillHeaderOnly.value = false;
        }
    };

    // 初期反映
    applyFillEnabled();


    /* 縦罫 / Vertical rules */
    var pVrule = gRight.add('panel', undefined, L('vRulePanel'));
    pVrule.orientation = 'column';
    pVrule.alignChildren = 'left';
    pVrule.margins = [15, 20, 15, 10];

    // 線（横ケイ＋縦ケイの有効/無効）
    var cbRule = pVrule.add('checkbox', undefined, L('ruleCheck'));
    cbRule.value = true;

    // 縦ケイ
    var pVkei = pVrule.add('panel', undefined, L('vRuleLabel'));
    pVkei.orientation = 'column';
    pVkei.alignChildren = 'left';
    pVkei.margins = [15, 20, 15, 10];

    var gVrule = pVkei.add('group');
    gVrule.orientation = 'column';
    gVrule.alignChildren = ['left', 'top'];

    var rbVruleNone = gVrule.add('radiobutton', undefined, L('none'));
    var rbVruleGapsOnly = gVrule.add('radiobutton', undefined, L('gapsOnly'));
    var rbVruleAll = gVrule.add('radiobutton', undefined, L('all'));


    // デフォルト：列間のみ
    rbVruleGapsOnly.value = true;

    // 「線」OFF時は縦ケイを「なし」にしてディム表示 / If Rules OFF, force vertical rules to None and dim
    function applyRuleEnabled() {
        if (!cbRule.value) {
            rbVruleNone.value = true;
            pVkei.enabled = false;
        } else {
            pVkei.enabled = true;
        }
    }
    cbRule.onClick = applyRuleEnabled;
    applyRuleEnabled();

    // プリセット適用（UIのみ。設定ロジックは後で追加可能）
    function applyPreset() {
        isApplyingPreset = true;
        try {
            var idx = ddPreset.selection ? ddPreset.selection.index : 0;
            if (idx === 0) return; // 手動

        // 共通：1行目ON
        cbHeader.value = true;
        cbFillJoinRow.value = false;

        // プリセット用：ヘッダー行のみ（現状のプリセットはすべてOFF）
        var presetHeaderOnly = false;
        cbFillHeaderOnly.value = presetHeaderOnly;

        // 1) 塗り：ON / 線：OFF / ガター：1mm / 縦ケイ：OFF
        if (idx === 1) {
            cbFill.value = true;
            applyFillEnabled();
            cbFillHeaderOnly.value = presetHeaderOnly;
            cbRule.value = false;

            cbUseGutter.value = true;
            applyGutterEnabled();
            setGutterByMm(1);

            rbVruleNone.value = true;
            applyRuleEnabled();
            return;
        }

        // 2) 塗り：OFF / 線：ON / ガター：2mm / 縦ケイ：すべて
        if (idx === 2) {
            cbFill.value = false;
            applyFillEnabled();
            cbFillHeaderOnly.value = presetHeaderOnly;
            cbRule.value = true;

            cbUseGutter.value = true;
            applyGutterEnabled();
            setGutterByMm(2);

            rbVruleAll.value = true;
            applyRuleEnabled();
            return;
        }

        // 3) 塗り：ON / 線：ON / ガター：0 / 縦ケイ：列間 / ヘッダー行のみ：ON
        if (idx === 3) {
            cbFill.value = true;
            applyFillEnabled();

            cbRule.value = true;

            cbUseGutter.value = false; // 0mm
            applyGutterEnabled();

            cbHeader.value = true;

            rbVruleGapsOnly.value = true;
            applyRuleEnabled();

            cbFillHeaderOnly.value = true;
            applyFillHeaderOnly();
            return;
        }

        // 4) 塗り：OFF / 線：ON / ガター：1mm / 縦ケイ：列間のみ
        if (idx === 4) {
            cbFill.value = false;
            applyFillEnabled();
            cbFillHeaderOnly.value = presetHeaderOnly;
            cbRule.value = true;

            cbUseGutter.value = true;
            applyGutterEnabled();
            setGutterByMm(1);

            rbVruleGapsOnly.value = true;
            applyRuleEnabled();
            return;
        }

        // 5) 塗り：OFF / 線：ON / ガター：0 / 縦ケイ：列間のみ
        if (idx === 5) {
            cbFill.value = false;
            applyFillEnabled();
            cbFillHeaderOnly.value = presetHeaderOnly;
            cbRule.value = true;

            cbUseGutter.value = false; // 0扱い（連結）
            applyGutterEnabled();

            rbVruleGapsOnly.value = true;
            applyRuleEnabled();
            return;
        }

        // 6) 塗り：OFF / 線：ON / ガター：0 / 縦ケイ：なし
        if (idx === 6) {
            cbFill.value = false;
            applyFillEnabled();
            cbFillHeaderOnly.value = presetHeaderOnly;
            cbRule.value = true;

            cbUseGutter.value = false; // 0扱い（連結）
            applyGutterEnabled();

            rbVruleNone.value = true;
            applyRuleEnabled();
            return;
        }

        // NOTE: 今後プリセットでヘッダー行のみをONにする場合は presetHeaderOnly=true にしてから
        // cbFillHeaderOnly.value を反映し、必要なら applyFillHeaderOnly() を呼ぶ。
        } finally {
            isApplyingPreset = false;
        }
    }

    ddPreset.onChange = function() {
        applyPreset();
    };


    // --- 手動変更検知 / Manual change detection ---
    function hookManual(control) {
        var prev = control.onClick;
        control.onClick = function() {
            if (prev) prev();
            setPresetManual();
        };
    }

    // チェックボックス類
    hookManual(cbFill);
    hookManual(cbZebra);
    hookManual(cbFillJoinRow);
    hookManual(cbFillHeaderOnly);
    hookManual(cbUseGutter);
    hookManual(cbHeader);
    hookManual(cbRule);

    // ラジオボタン
    hookManual(rbVruleNone);
    hookManual(rbVruleGapsOnly);
    hookManual(rbVruleAll);

    // ガター数値の手動変更
    etHGutter.onChanging = function() {
        setPresetManual();
    };



    // セッション状態を復元
    restoreDialogState({
        ddPreset: ddPreset,
        cbUseGutter: cbUseGutter,
        etHGutter: etHGutter,
        cbHeader: cbHeader,
        cbFill: cbFill,
        cbZebra: cbZebra,
        cbFillJoinRow: cbFillJoinRow,
        cbFillHeaderOnly: cbFillHeaderOnly,
        cbRule: cbRule,
        rbVruleNone: rbVruleNone,
        rbVruleGapsOnly: rbVruleGapsOnly,
        rbVruleAll: rbVruleAll,
        applyGutterEnabled: applyGutterEnabled,
        applyFillEnabled: applyFillEnabled,
        applyRuleEnabled: applyRuleEnabled,
        applyFillJoinRow: applyFillJoinRow,
        applyFillHeaderOnly: applyFillHeaderOnly
    });


    var btnGroup = dlg.add('group');
    btnGroup.alignment = 'right';
    var btnCancel = btnGroup.add('button', undefined, L('cancel'), { name: 'cancel' });
    var btnOK = btnGroup.add('button', undefined, L('ok'), { name: 'ok' });
    btnCancel.onClick = function() {
        saveDialogState({
            ddPreset: ddPreset,
            cbUseGutter: cbUseGutter,
            etHGutter: etHGutter,
            cbHeader: cbHeader,
            cbFill: cbFill,
            cbZebra: cbZebra,
            cbFillJoinRow: cbFillJoinRow,
            cbFillHeaderOnly: cbFillHeaderOnly,
            cbRule: cbRule,
            rbVruleNone: rbVruleNone,
            rbVruleGapsOnly: rbVruleGapsOnly,
            rbVruleAll: rbVruleAll
        });
        dlg.close(0);
    };
    btnOK.onClick = function() {
        saveDialogState({
            ddPreset: ddPreset,
            cbUseGutter: cbUseGutter,
            etHGutter: etHGutter,
            cbHeader: cbHeader,
            cbFill: cbFill,
            cbZebra: cbZebra,
            cbFillJoinRow: cbFillJoinRow,
            cbFillHeaderOnly: cbFillHeaderOnly,
            cbRule: cbRule,
            rbVruleNone: rbVruleNone,
            rbVruleGapsOnly: rbVruleGapsOnly,
            rbVruleAll: rbVruleAll
        });
        dlg.close(1);
    };
    dlg.onClose = function() {
        try {
            saveDialogState({
                ddPreset: ddPreset,
                cbUseGutter: cbUseGutter,
                etHGutter: etHGutter,
                cbHeader: cbHeader,
                cbFill: cbFill,
                cbZebra: cbZebra,
                cbFillJoinRow: cbFillJoinRow,
                cbFillHeaderOnly: cbFillHeaderOnly,
                cbRule: cbRule,
                rbVruleNone: rbVruleNone,
                rbVruleGapsOnly: rbVruleGapsOnly,
                rbVruleAll: rbVruleAll
            });
        } catch (_) { }
    };

    // ドキュメントチェック（ダイアログより前に移動）
    if (app.documents.length === 0) {
        alert(L('alertOpenDoc'));
        return;
    }
    var doc = app.activeDocument;
    var sel = doc.selection;
    var baseLayer = doc.activeLayer;
    if (sel.length === 0) {
        alert(L('alertSelectObj'));
        return;
    }
    // 選択オブジェクトの正規化
    var items = [];
    for (var i = 0; i < sel.length; i++) {
        var it = sel[i];
        var g = getSelectedAncestorGroup(it);
        if (g) it = g;
        pushUniqueRef(items, it);
    }
    // 列（カラム）のグループ化
    var columns = buildColumnsFromItems(items);

    // テキストが含まれる場合は、一度だけアウトライン参照を作って以降の計算に使う（プレビュー含む）
    var refItems = buildReferenceItems(items);
    if (refItems && refItems.length) {
        items = refItems;
        columns = buildColumnsFromItems(items);
    }

    // レイヤー作成（現在のレイヤーの背面） / Create layers behind the current layer
    // 罫線レイヤー
    var lineLayer = null;
    // 塗りレイヤー
    var fillLayer = null;

    // 線の色設定（黒）
    var blackColor = new CMYKColor();
    blackColor.cyan = 0;
    blackColor.magenta = 0;
    blackColor.yellow = 0;
    blackColor.black = 100;

    // 塗り色設定（薄いグレー）/ Fill color (light gray)
    var fillGray = new CMYKColor();
    fillGray.cyan = 0;
    fillGray.magenta = 0;
    fillGray.yellow = 0;
    fillGray.black = 15;

    // ヘッダー塗り色（濃いグレー）/ Header fill color (darker gray)
    var fillGrayHeader = new CMYKColor();
    fillGrayHeader.cyan = 0;
    fillGrayHeader.magenta = 0;
    fillGrayHeader.yellow = 0;
    fillGrayHeader.black = 40;

    // ヘッダー塗り色（ゼブラON時）/ Header fill color when Zebra is ON
    var fillGrayHeaderZebra = new CMYKColor();
    fillGrayHeaderZebra.cyan = 0;
    fillGrayHeaderZebra.magenta = 0;
    fillGrayHeaderZebra.yellow = 0;
    fillGrayHeaderZebra.black = 50;

    // ゼブラ塗り色（奇数行）/ Zebra fill color (odd rows)
    var fillGrayZebra = new CMYKColor();
    fillGrayZebra.cyan = 0;
    fillGrayZebra.magenta = 0;
    fillGrayZebra.yellow = 0;
    fillGrayZebra.black = 30;


    // --- プレビュー/ライブ生成ヘルパー ---
    // プレビュー用レイヤー名
    var PREVIEW_LAYER_NAME = "__TabularizePreview__";
    var PREVIEW_GROUP_NAME = "__TabularizePreviewGroup__";

    // =========================
    // 参照ジオメトリ（テキストをアウトライン化して一度だけ計算）
    // =========================
    var REF_LAYER_NAME = "__TabularizeRef__";

    function getOrCreateRefLayer(doc) {
        // Always resolve to a Document (defensive: avoid passing a Layer/other object)
        var d = null;
        try {
            if (doc && doc.typename === 'Document') d = doc;
        } catch (_) { d = null; }
        if (!d) {
            try { d = app.activeDocument; } catch (_) { d = null; }
        }
        if (!d) return null;

        var lyr;
        try {
            lyr = d.layers.getByName(REF_LAYER_NAME);
        } catch (e) {
            lyr = d.layers.add();
            lyr.name = REF_LAYER_NAME;
        }
        try { lyr.visible = false; } catch (_) {}
        try { lyr.locked = false; } catch (_) {}
        try { lyr.template = false; } catch (_) {}
        // 最背面へ
        try {
            if (d.layers.length > 0 && d.layers[d.layers.length - 1] !== lyr) {
                lyr.move(d.layers[d.layers.length - 1], ElementPlacement.PLACEAFTER);
            }
        } catch (_) {}
        return lyr;
    }

    function clearRefLayer() {
        var d;
        try { d = app.activeDocument; } catch (_) { d = null; }
        if (!d) return;
        try {
            var lyr = d.layers.getByName(REF_LAYER_NAME);
            try { lyr.locked = false; } catch (_) {}
            lyr.remove();
        } catch (_) {}
    }

    // TextFrame をアウトライン化した参照アイテム配列を作る（1回だけ）
    function buildReferenceItems(srcItems) {
        var hasText = false;
        for (var i = 0; i < srcItems.length; i++) {
            if (srcItems[i] && srcItems[i].typename === 'TextFrame') { hasText = true; break; }
        }
        if (!hasText) return null;

        // 既存参照があれば消して作り直し
        clearRefLayer();
        var refLayer = getOrCreateRefLayer();
        if (!refLayer) return null;

        var refItems = [];
        for (var j = 0; j < srcItems.length; j++) {
            var it = srcItems[j];
            if (!it) continue;

            try {
                var dup = it.duplicate(refLayer, ElementPlacement.PLACEATBEGINNING);

                // テキストはアウトライン化
                if (dup && dup.typename === 'TextFrame') {
                    var outlined = null;
                    try {
                        outlined = dup.createOutline();
                    } catch (e1) {
                        // 保険（メニューコマンド）
                        var _oldSel = null;
                        try { _oldSel = doc.selection; } catch (_) { _oldSel = null; }
                        try {
                            doc.selection = [dup];
                            app.executeMenuCommand('outline');
                            outlined = (doc.selection && doc.selection.length) ? doc.selection[0] : null;
                        } catch (e2) { outlined = null; }
                        try { if (_oldSel !== null) doc.selection = _oldSel; } catch (_) {}
                    }
                    try { dup.remove(); } catch (_) {}
                    if (outlined) refItems.push(outlined);
                } else {
                    // テキスト以外は複製そのまま
                    refItems.push(dup);
                }
            } catch (_) {}
        }

        try { refLayer.visible = false; } catch (_) {}
        return refItems;
    }

    // items配列から columns を再構築（既存ロジックと同じ）
    function buildColumnsFromItems(itemsArr) {
        var cols = [];
        if (!itemsArr || itemsArr.length === 0) return cols;

        itemsArr.sort(function (a, b) {
            return a.geometricBounds[0] - b.geometricBounds[0];
        });

        var currentColumn = [itemsArr[0]];
        cols.push(currentColumn);

        for (var k = 1; k < itemsArr.length; k++) {
            var item2 = itemsArr[k];
            var itemLeft = item2.geometricBounds[0];
            var prevColMaxRight = getMaxRightInColumn(currentColumn);

            if (itemLeft < prevColMaxRight) {
                currentColumn.push(item2);
            } else {
                currentColumn = [item2];
                cols.push(currentColumn);
            }
        }
        return cols;
    }

    // プレビュー用の描画先（Layer優先。Layerが作れない場合はGroupにフォールバック）
    function getOrCreatePreviewLayer(doc) {
        // Resolve document
        var d = null;
        try { if (doc && doc.typename === 'Document') d = doc; } catch (_) { d = null; }
        if (!d) { try { d = app.activeDocument; } catch (_) { d = null; } }
        if (!d) return null;

        // まずレイヤーを取得
        var lyr = null;
        try {
            lyr = d.layers.getByName(PREVIEW_LAYER_NAME);
        } catch (_) {
            lyr = null;
        }

        // レイヤーが無ければ作成を試みる（失敗したらGroupへ）
        if (!lyr) {
            try {
                lyr = d.layers.add();
                lyr.name = PREVIEW_LAYER_NAME;
            } catch (eAdd) {
                lyr = null;
            }
        }

        // Layerが作れた場合
        if (lyr) {
            try { lyr.visible = true; } catch (_) {}
            try { lyr.locked = false; } catch (_) {}
            try { lyr.template = false; } catch (_) {}
            try {
                if (d.layers.length > 0 && d.layers[0] !== lyr) {
                    lyr.move(d.layers[0], ElementPlacement.PLACEBEFORE);
                }
            } catch (_) {}
            try { lyr.zOrder(ZOrderMethod.BRINGTOFRONT); } catch (_) {}
            return lyr;
        }

        // --- Fallback: GroupItem in active layer ---
        var host = null;
        try { host = d.activeLayer; } catch (_) { host = null; }
        if (!host) return null;

        // 既存プレビューGroupがあれば削除
        try {
            var pg = null;
            for (var i = host.groupItems.length - 1; i >= 0; i--) {
                if (host.groupItems[i].name === PREVIEW_GROUP_NAME) { pg = host.groupItems[i]; break; }
            }
            if (pg) pg.remove();
        } catch (_) {}

        // 新規作成
        var g;
        try {
            g = host.groupItems.add();
            g.name = PREVIEW_GROUP_NAME;
        } catch (_) {
            return host; // 最終手段
        }
        return g;
    }
    // プレビューを削除（LayerまたはGroupフォールバックの両対応）
    function clearPreview() {
        var d;
        try { d = app.activeDocument; } catch (_) { d = null; }
        if (!d) return;

        // Remove preview layer if exists
        try {
            var lyr = d.layers.getByName(PREVIEW_LAYER_NAME);
            lyr.remove();
            return;
        } catch (_) {}

        // Remove fallback group if exists
        try {
            var host = d.activeLayer;
            for (var i = host.groupItems.length - 1; i >= 0; i--) {
                if (host.groupItems[i].name === PREVIEW_GROUP_NAME) {
                    host.groupItems[i].remove();
                    break;
                }
            }
        } catch (_) {}
    }
    // プレビューを生成
    function updatePreview(params) {
        clearPreview();
        generateToLayers(params, true);
        try { app.redraw(); } catch (_) {}
    }
    // UI値を読み取ってパラメータを返す
    function readParamsFromUI() {
        var isZebra = cbZebra.value;
        var isFillJoinRow = cbFillJoinRow.value;
        var isFillHeaderOnly = cbFillHeaderOnly.value;
        var isHeaderRow = cbHeader.value;
        var hGutterVal = cbUseGutter.value ? parseFloat(etHGutter.text) : 0;
        if (isNaN(hGutterVal) || hGutterVal < 0) hGutterVal = 0;
        var hGutterPt = hGutterVal * rulerFactorPt;
        var vRuleMode = rbVruleNone.value ? 'none' : (rbVruleAll.value ? 'all' : 'gapsOnly');
        var doFill = cbFill.value;
        var doRule = cbRule.value;
        var fillMode = (doFill && doRule) ? 'fillAndRule' : (doFill ? 'fillOnly' : 'none');
        var KEEP_GAP_PT = hGutterPt;
        return {
            isZebra: isZebra,
            isFillJoinRow: isFillJoinRow,
            isFillHeaderOnly: isFillHeaderOnly,
            isHeaderRow: isHeaderRow,
            hGutterVal: hGutterVal,
            hGutterPt: hGutterPt,
            vRuleMode: vRuleMode,
            doFill: doFill,
            doRule: doRule,
            fillMode: fillMode,
            KEEP_GAP_PT: KEEP_GAP_PT
        };
    }
    // プレビュー更新を安全に
    function safeUpdatePreview() {
        try {
            updatePreview(readParamsFromUI());
        } catch (e) {}
    }
    // プレビュー呼び出しをUIイベントにフック
    dlg.onShow = (function(orig) {
        return function() {
            if (orig) try { orig(); } catch (_) {}
            safeUpdatePreview();
        };
    })(dlg.onShow);
    // UI変更時にプレビュー
    function addPreviewHook(ctrl, evtName) {
        if (ctrl && typeof ctrl.addEventListener === "function") {
            ctrl.addEventListener(evtName, safeUpdatePreview);
        } else if (ctrl && evtName in ctrl) {
            var prev = ctrl[evtName];
            ctrl[evtName] = function() {
                if (prev) prev();
                safeUpdatePreview();
            };
        }
    }
    // チェックボックス
var previewCtrls = [
    cbFill, cbZebra, cbFillJoinRow, cbFillHeaderOnly,
    cbUseGutter, cbHeader, cbRule,
    rbVruleNone, rbVruleGapsOnly, rbVruleAll
];

for (var i = 0; i < previewCtrls.length; i++) {
    addPreviewHook(previewCtrls[i], "onClick");
}
    // ガター数値
    if (etHGutter && typeof etHGutter.onChanging !== "undefined") {
        var prevChanging = etHGutter.onChanging;
        etHGutter.onChanging = function() {
            if (prevChanging) prevChanging();
            safeUpdatePreview();
        };
    }
    // プリセット
    if (ddPreset) {
        var prevPreset = ddPreset.onChange;
        ddPreset.onChange = function() {
            if (prevPreset) prevPreset();
            safeUpdatePreview();
        };
    }
    // Cancel/OK/Close時にプレビュー消去＆参照レイヤー消去
    var origCancel = btnCancel.onClick;
    btnCancel.onClick = function() {
        clearPreview();
        clearRefLayer();
        if (origCancel) origCancel();
    };
    var origOK = btnOK.onClick;
    btnOK.onClick = function() {
        clearPreview();
        clearRefLayer();
        if (origOK) origOK();
    };
    var origClose = dlg.onClose;
    dlg.onClose = function() {
        clearPreview();
        clearRefLayer();
        if (origClose) origClose();
    };
    // dlg.show()でキャンセル時にプレビュー消去してreturn
    var dlgResult = dlg.show();
    if (dlgResult !== 1) {
        clearPreview();
        clearRefLayer();
        return;
    }
    clearPreview();
    // --- 生成パラメータを取得し、確定生成 ---
    generateToLayers(readParamsFromUI(), false);

    // 生成処理本体 / Main generation
    function generateToLayers(params, isPreview) {
        // isPreview: trueならプレビュー用レイヤー, falseなら本番レイヤー
        var targetLineLayer, targetFillLayer;
        if (isPreview) {
            var previewLayer = getOrCreatePreviewLayer();
            if (!previewLayer) return; // ドキュメントが取れない場合は何もしない
            // プレビュー時は fill/lineともこのレイヤーに
            targetLineLayer = previewLayer;
            targetFillLayer = previewLayer;
        } else {
            // 本番は既存レイヤー取得または作成
            // 罫線レイヤー
            try {
                targetLineLayer = doc.layers.getByName("罫線レイヤー");
            } catch (e) {
                targetLineLayer = doc.layers.add();
                targetLineLayer.name = "罫線レイヤー";
                try { targetLineLayer.move(baseLayer, ElementPlacement.PLACEAFTER); } catch (_) { }
            }
            // 塗りレイヤー
            try {
                targetFillLayer = doc.layers.getByName("塗りレイヤー");
            } catch (e2) {
                targetFillLayer = doc.layers.add();
                targetFillLayer.name = "塗りレイヤー";
                try { targetFillLayer.move(targetLineLayer, ElementPlacement.PLACEAFTER); } catch (_) { }
            }
            try {
                if (targetLineLayer !== baseLayer) targetLineLayer.move(baseLayer, ElementPlacement.PLACEAFTER);
            } catch (_) { }
            try {
                if (targetFillLayer !== targetLineLayer) targetFillLayer.move(targetLineLayer, ElementPlacement.PLACEAFTER);
            } catch (_) { }
        }
        // 事前クリア
        function clearLayer(layer) {
            try {
                for (var i = layer.pageItems.length - 1; i >= 0; i--) {
                    layer.pageItems[i].remove();
                }
            } catch (_) {}
        }
        clearLayer(targetLineLayer);
        if (targetFillLayer !== targetLineLayer) clearLayer(targetFillLayer);

        // --- 以降は元のgenerateMain()のローカル変数をparams経由で参照 ---
        var isZebra = params.isZebra;
        var isFillJoinRow = params.isFillJoinRow;
        var isFillHeaderOnly = params.isFillHeaderOnly;
        var isHeaderRow = params.isHeaderRow;
        var hGutterVal = params.hGutterVal;
        var hGutterPt = params.hGutterPt;
        var vRuleMode = params.vRuleMode;
        var doFill = params.doFill;
        var doRule = params.doRule;
        var fillMode = params.fillMode;
        var KEEP_GAP_PT = params.KEEP_GAP_PT;

        // 列ごとの左右端を先に計算しておく（列間Aを求めるため）
        var colBounds = []; // {minX, maxX}
        for (var c0 = 0; c0 < columns.length; c0++) {
            var colItems0 = columns[c0];
            var minX0 = 999999;
            var maxX0 = -999999;
            for (var k0 = 0; k0 < colItems0.length; k0++) {
                var bb0 = colItems0[k0].geometricBounds; // [left, top, right, bottom]
                if (bb0[0] < minX0) minX0 = bb0[0];
                if (bb0[2] > maxX0) maxX0 = bb0[2];
            }
            colBounds.push({ minX: minX0, maxX: maxX0 });
        }

        // --- 行（ロウ）を全体で共通化 ---
        var rowTol = 2.0; // 同一行とみなす中心Yの許容値(pt)
        var baseColIdx = 0;
        var maxCount = -1;
        for (var bc = 0; bc < columns.length; bc++) {
            if (columns[bc].length > maxCount) {
                maxCount = columns[bc].length;
                baseColIdx = bc;
            }
        }
        var allRows = [];
        var baseItemsSorted = columns[baseColIdx].slice(0);
        baseItemsSorted.sort(function (a, b) { return b.geometricBounds[1] - a.geometricBounds[1]; });
        for (var ai = 0; ai < baseItemsSorted.length; ai++) {
            var it = baseItemsSorted[ai];
            var bb = it.geometricBounds; // [left, top, right, bottom]
            var top = bb[1];
            var bottom = bb[3];
            var h = top - bottom;
            var center = (top + bottom) / 2;
            var ridx = findRowIndex(allRows, center, rowTol);
            if (ridx === -1) {
                allRows.push({ centerY: center, maxH: h });
            } else {
                if (h > allRows[ridx].maxH) allRows[ridx].maxH = h;
                allRows[ridx].centerY = (allRows[ridx].centerY + center) / 2;
            }
        }
        allRows.sort(function (a, b) { return b.centerY - a.centerY; });
        var yListGlobal = [];
        function rowTop(i) {
            return allRows[i].centerY + (allRows[i].maxH / 2);
        }
        function rowBottom(i) {
            return allRows[i].centerY - (allRows[i].maxH / 2);
        }
        if (allRows.length >= 2) {
            for (var r = 0; r < allRows.length - 1; r++) {
                var upperBottom = rowBottom(r);
                var lowerTop = rowTop(r + 1);
                var gap = upperBottom - lowerTop;
                var mid = upperBottom - (gap / 2);
                yListGlobal.push(mid);
            }
            var firstGap = rowBottom(0) - rowTop(1);
            yListGlobal.push(rowTop(0) + (firstGap / 2));
            var last = allRows.length - 1;
            var lastGap = rowBottom(last - 1) - rowTop(last);
            yListGlobal.push(rowBottom(last) - (lastGap / 2));
        } else if (allRows.length === 1) {
            var t0 = rowTop(0);
            var b0 = rowBottom(0);
            var rowH = t0 - b0;
            yListGlobal.push(t0 + (rowH / 2));
            yListGlobal.push(b0 - (rowH / 2));
        }
        var yTol = 0.4;
        var yUniqGlobal = [];
        for (var yiG = 0; yiG < yListGlobal.length; yiG++) {
            addUniqueY(yUniqGlobal, yListGlobal[yiG], yTol);
        }
        yUniqGlobal.sort(function (a, b) { return b - a; });
        var yTopBorder = (yUniqGlobal.length > 0) ? yUniqGlobal[0] : 0;
        var yBottomBorder = (yUniqGlobal.length > 0) ? yUniqGlobal[yUniqGlobal.length - 1] : 0;
        var colFillLeft = [];
        var colFillRight = [];
        var halfGap = 0;
        for (var cc2 = 0; cc2 < columns.length; cc2++) {
            if (cc2 === 0) {
                var minX_0 = colBounds[cc2].minX;
                var maxX_0 = colBounds[cc2].maxX;
                var extL0 = 0;
                if (columns.length >= 2) {
                    var A0_ = colBounds[1].minX - maxX_0;
                    extL0 = (A0_ - KEEP_GAP_PT) / 2;
                    if (extL0 < 0) extL0 = 0;
                }
                colFillLeft[cc2] = minX_0 - paddingPt - extL0;
            } else {
                var xMidL = (colBounds[cc2 - 1].maxX + colBounds[cc2].minX) / 2;
                colFillLeft[cc2] = xMidL;
            }
            if (cc2 === columns.length - 1) {
                var minX_n = colBounds[cc2].minX;
                var maxX_n = colBounds[cc2].maxX;
                var extRn = 0;
                if (columns.length >= 2) {
                    var An_ = minX_n - colBounds[cc2 - 1].maxX;
                    extRn = (An_ - KEEP_GAP_PT) / 2;
                    if (extRn < 0) extRn = 0;
                }
                colFillRight[cc2] = maxX_n + paddingPt + extRn;
            } else {
                var xMidR = (colBounds[cc2].maxX + colBounds[cc2 + 1].minX) / 2;
                colFillRight[cc2] = xMidR;
            }
        }
        // --- 塗り / Fill ---
        if (doFill && yUniqGlobal.length >= 2) {
            if (isFillHeaderOnly) {
                var xL0 = colFillLeft[0];
                var xR0 = colFillRight[columns.length - 1];
                var yTopH = yUniqGlobal[0];
                var yBotH = yUniqGlobal[1];
                var wH = xR0 - xL0;
                var hH = yTopH - yBotH;
                if (wH > 0 && hH > 0) {
                    var rectH = targetFillLayer.pathItems.rectangle(yTopH, xL0, wH, hH);
                    rectH.stroked = false;
                    rectH.filled = true;
                    rectH.fillColor = isZebra ? fillGrayHeaderZebra : fillGrayHeader;
                }
            } else if (isFillJoinRow) {
                var xL0 = colFillLeft[0];
                var xR0 = colFillRight[columns.length - 1];
                for (var fr = 0; fr < yUniqGlobal.length - 1; fr++) {
                    var yTop = yUniqGlobal[fr];
                    var yBot = yUniqGlobal[fr + 1];
                    var left = xL0;
                    var right = xR0;
                    var top = yTop;
                    var bottom = yBot;
                    var w = right - left;
                    var h = top - bottom;
                    if (w <= 0 || h <= 0) continue;
                    var rect = targetFillLayer.pathItems.rectangle(top, left, w, h);
                    rect.stroked = false;
                    rect.filled = true;
                    if (isHeaderRow && fr === 0) {
                        rect.fillColor = isZebra ? fillGrayHeaderZebra : fillGrayHeader;
                    } else if (isZebra && ((fr + 1) % 2 === 1)) {
                        rect.fillColor = fillGrayZebra;
                    } else {
                        rect.fillColor = fillGray;
                    }
                }
            } else {
                var insetX = KEEP_GAP_PT / 2;
                var insetY = KEEP_GAP_PT / 2;
                for (var fc = 0; fc < columns.length; fc++) {
                    var xL = colFillLeft[fc];
                    var xR = colFillRight[fc];
                    for (var fr2 = 0; fr2 < yUniqGlobal.length - 1; fr2++) {
                        var yTop2 = yUniqGlobal[fr2];
                        var yBot2 = yUniqGlobal[fr2 + 1];
                        var left2 = xL + insetX;
                        var right2 = xR - insetX;
                        var top2 = yTop2 - insetY;
                        var bottom2 = yBot2 + insetY;
                        var w2 = right2 - left2;
                        var h2 = top2 - bottom2;
                        if (w2 <= 0 || h2 <= 0) continue;
                        var rect2 = targetFillLayer.pathItems.rectangle(top2, left2, w2, h2);
                        rect2.stroked = false;
                        rect2.filled = true;
                        if (isHeaderRow && fr2 === 0) {
                            rect2.fillColor = isZebra ? fillGrayHeaderZebra : fillGrayHeader;
                        } else if (isZebra && ((fr2 + 1) % 2 === 1)) {
                            rect2.fillColor = fillGrayZebra;
                        } else {
                            rect2.fillColor = fillGray;
                        }
                    }
                }
            }
        }
        // --- 横罫（行ごと） ---
        if (doRule) {
            if (!cbUseGutter.value || KEEP_GAP_PT === 0) {
                var minXLeft = colBounds[0].minX;
                var maxXLeft = colBounds[0].maxX;
                var minXRight = colBounds[columns.length - 1].minX;
                var maxXRight = colBounds[columns.length - 1].maxX;
                var extLeft = 0;
                var extRight = 0;
                if (columns.length >= 2) {
                    var A0 = colBounds[1].minX - maxXLeft;
                    extLeft = (A0 - KEEP_GAP_PT) / 2;
                    if (extLeft < 0) extLeft = 0;
                    var lastIdx = columns.length - 1;
                    var An = minXRight - colBounds[lastIdx - 1].maxX;
                    extRight = (An - KEEP_GAP_PT) / 2;
                    if (extRight < 0) extRight = 0;
                }
                var globalStartX = minXLeft - paddingPt - extLeft;
                var globalEndX = maxXRight + paddingPt + extRight;
                for (var iLine = 0; iLine < yUniqGlobal.length; iLine++) {
                    var strokeW = lineWeightPt;
                    if (isHeaderRow && iLine < 2) strokeW = headerLineWeightPt;
                    drawLineToLayer(globalStartX, yUniqGlobal[iLine], globalEndX, yUniqGlobal[iLine], strokeW, targetLineLayer);
                }
            } else {
                for (var c = 0; c < columns.length; c++) {
                    var colItems = columns[c];
                    var minX = colBounds[c].minX;
                    var maxX = colBounds[c].maxX;
                    var extendLeft = 0;
                    var extendRight = 0;
                    if (c > 0) {
                        var prevMaxX = colBounds[c - 1].maxX;
                        var A_left = minX - prevMaxX;
                        extendLeft = (A_left - KEEP_GAP_PT) / 2;
                        if (extendLeft < 0) extendLeft = 0;
                    }
                    if (c < columns.length - 1) {
                        var nextMinX = colBounds[c + 1].minX;
                        var A_right = nextMinX - maxX;
                        extendRight = (A_right - KEEP_GAP_PT) / 2;
                        if (extendRight < 0) extendRight = 0;
                    }
                    if (c === 0 && columns.length >= 2) {
                        var nextMinX0 = colBounds[1].minX;
                        var A0a = nextMinX0 - maxX;
                        var ext0 = (A0a - KEEP_GAP_PT) / 2;
                        if (ext0 < 0) ext0 = 0;
                        extendLeft = ext0;
                    }
                    if (c === columns.length - 1 && columns.length >= 2) {
                        var prevMaxXn = colBounds[columns.length - 2].maxX;
                        var An2 = minX - prevMaxXn;
                        var extn = (An2 - KEEP_GAP_PT) / 2;
                        if (extn < 0) extn = 0;
                        extendRight = extn;
                    }
                    var lineStartX = minX - paddingPt - extendLeft;
                    var lineEndX = maxX + paddingPt + extendRight;
                    for (var iLine2 = 0; iLine2 < yUniqGlobal.length; iLine2++) {
                        var strokeW2 = lineWeightPt;
                        if (isHeaderRow && iLine2 < 2) strokeW2 = headerLineWeightPt;
                        drawLineToLayer(lineStartX, yUniqGlobal[iLine2], lineEndX, yUniqGlobal[iLine2], strokeW2, targetLineLayer);
                    }
                }
            }
        }
        // --- 縦罫（列間のみ / すべて） ---
        if (doRule && vRuleMode !== 'none' && yUniqGlobal && yUniqGlobal.length > 0 && columns.length >= 2) {
            var vXs = [];
            for (var cc = 0; cc < columns.length - 1; cc++) {
                var xMid = (colBounds[cc].maxX + colBounds[cc + 1].minX) / 2;
                vXs.push(xMid);
            }
            if (vRuleMode === 'all') {
                var leftMinX = colBounds[0].minX;
                var leftExtend = 0;
                {
                    var A0 = colBounds[1].minX - colBounds[0].maxX;
                    leftExtend = (A0 - KEEP_GAP_PT) / 2;
                    if (leftExtend < 0) leftExtend = 0;
                }
                var xLeftBorder = leftMinX - paddingPt - leftExtend;
                var lastIdx = columns.length - 1;
                var rightMaxX = colBounds[lastIdx].maxX;
                var rightExtend = 0;
                {
                    var An = colBounds[lastIdx].minX - colBounds[lastIdx - 1].maxX;
                    rightExtend = (An - KEEP_GAP_PT) / 2;
                    if (rightExtend < 0) rightExtend = 0;
                }
                var xRightBorder = rightMaxX + paddingPt + rightExtend;
                vXs.unshift(xLeftBorder);
                vXs.push(xRightBorder);
            }
            if (!cbUseGutter.value) {
                for (var vx = 0; vx < vXs.length; vx++) {
                    drawLineToLayer(vXs[vx], yTopBorder, vXs[vx], yBottomBorder, lineWeightPt, targetLineLayer);
                }
            } else {
                var vTrim = KEEP_GAP_PT / 2;
                for (var vx = 0; vx < vXs.length; vx++) {
                    for (var ry = 0; ry < yUniqGlobal.length - 1; ry++) {
                        var y1 = yUniqGlobal[ry];
                        var y2 = yUniqGlobal[ry + 1];
                        var segH = y1 - y2;
                        if (segH <= vTrim * 2) {
                            drawLineToLayer(vXs[vx], y1, vXs[vx], y2, lineWeightPt, targetLineLayer);
                            continue;
                        }
                        var ys = y1 - vTrim;
                        var ye = y2 + vTrim;
                        drawLineToLayer(vXs[vx], ys, vXs[vx], ye, lineWeightPt, targetLineLayer);
                    }
                }
            }
        }
    }

    // (duplicate block removed)

    /* ユーティリティ関数 / Utilities */

    /* 選択の正規化 / Normalize selection */

    /* グループを1アイテム扱い / Treat group as a single item */

    // 選択オブジェクトがグループ内にある場合、最上位の親グループを返す（グループ選択の有無は問わない）
    // If the item is inside a group, return the topmost parent group (regardless of group selection state).
    function getSelectedAncestorGroup(item) {
        var p = item;
        var found = null;
        while (p && p.parent && p.parent.typename === 'GroupItem') {
            found = p.parent;
            p = p.parent;
        }
        return found;
    }

    // 参照でユニーク追加（同じオブジェクトを重複追加しない）
    function pushUniqueRef(arr, obj) {
        for (var i = 0; i < arr.length; i++) {
            if (arr[i] === obj) return;
        }
        arr.push(obj);
    }


    /* 値の変更（↑↓キー） / Change value by arrow keys */

    function changeValueByArrowKey(editText) {
        editText.addEventListener("keydown", function (event) {
            var value = Number(editText.text);
            if (isNaN(value)) return;

            var keyboard = ScriptUI.environment.keyboardState;
            var delta = 1;

            if (keyboard.shiftKey) {
                delta = 10;
                // Shiftキー押下時は10の倍数にスナップ
                if (event.keyName == "Up") {
                    value = Math.ceil((value + 1) / delta) * delta;
                    event.preventDefault();
                } else if (event.keyName == "Down") {
                    value = Math.floor((value - 1) / delta) * delta;
                    if (value < 0) value = 0;
                    event.preventDefault();
                }
            } else if (keyboard.altKey) {
                delta = 0.1;
                // Optionキー押下時は0.1単位で増減
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
                    if (value < 0) value = 0;
                    event.preventDefault();
                }
            }

            if (keyboard.altKey) {
                // 小数第1位までに丸め
                value = Math.round(value * 10) / 10;
            } else {
                // 整数に丸め
                value = Math.round(value);
            }

            editText.text = value;
        });
    }


    // 近いYを同一とみなしてユニークに追加する
    function addUniqueY(list, y, tol) {
        for (var i = 0; i < list.length; i++) {
            if (Math.abs(list[i] - y) <= tol) return;
        }
        list.push(y);
    }

    // rows配列から、centerYが近い行のインデックスを返す（見つからなければ-1）
    function findRowIndex(rows, centerY, tol) {
        for (var i = 0; i < rows.length; i++) {
            if (Math.abs(rows[i].centerY - centerY) <= tol) return i;
        }
        return -1;
    }

    // 列内の最大右端座標を取得する関数
    function getMaxRightInColumn(colAry) {
        var maxR = -999999;
        for (var i = 0; i < colAry.length; i++) {
            var r = colAry[i].geometricBounds[2];
            if (r > maxR) maxR = r;
        }
        return maxR;
    }

    // 線描画関数（レイヤー指定）
    function drawLineToLayer(x1, y1, x2, y2, strokeW, layer) {
        var pathItem = layer.pathItems.add();
        pathItem.setEntirePath([[x1, y1], [x2, y2]]);
        pathItem.stroked = true;
        try { pathItem.strokeCap = StrokeCap.BUTTENDCAP; } catch (_) { }
        try { pathItem.strokeJoin = StrokeJoin.MITERENDJOIN; } catch (_) { }
        pathItem.strokeWidth = (typeof strokeW === 'number') ? strokeW : lineWeightPt;
        pathItem.strokeColor = blackColor;
        pathItem.filled = false;
    }

})();
