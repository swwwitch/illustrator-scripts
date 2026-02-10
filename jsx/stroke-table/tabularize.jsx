#targetengine "TabularizeEngine"
#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
  tabularize.jsx

  選択オブジェクトを「表」として解釈し、表組み用の塗りと線（横ケイ/縦ケイ）を生成します。
  - テキストは計算用に複製・アウトライン化されます（元のテキストは編集可能なまま、非破壊です）。
  - 塗り：通常／ゼブラ／行方向に連結／ヘッダー行のみ
  - オプション：ガター（rulerType単位）、1行目をヘッダー行に
  - 線：縦ケイ（なし／列間のみ／すべて）＋ ガター0時は連結描画
  - プリセット：代表的な組み合わせを一括適用
  - ダイアログ値はセッション内で復元（Illustrator再起動でリセット）
  - 更新日: 2026-02-10

  Interpret the selection as a table grid and generate fills and rules (horizontal/vertical).
  - Text is duplicated and outlined for calculation only (original text remains editable).
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
    var SCRIPT_VERSION = "v1.2";

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
        dlg.onShow = function () {
            try { requestPreviewUpdate(); } catch (_) { }

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

    // --- Preview state ---
    var cbPreview = null;
    var isPreviewing = false;
    var previewItems = []; // created items for preview (lines + fills)
    var previewIsCurrent = false;

    function clearPreview() {
        try {
            for (var i = 0; i < previewItems.length; i++) {
                try { previewItems[i].remove(); } catch (_) { }
            }
        } catch (_) { }
        previewItems = [];
        previewIsCurrent = false;
        try { app.redraw(); } catch (_) { }
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

    cbFill.onClick = function () {
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
        try { requestPreviewUpdate(); } catch (_) { }
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

    ddPreset.onChange = function () {
        applyPreset();
    };


    // --- 手動変更検知 / Manual change detection ---
    function hookManual(control) {
        var prev = control.onClick;
        control.onClick = function () {
            if (prev) prev();
            setPresetManual();
            try { requestPreviewUpdate(); } catch (_) { }
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
    etHGutter.onChanging = function () {
        setPresetManual();
        try { requestPreviewUpdate(); } catch (_) { }
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
    btnGroup.orientation = 'row';
    btnGroup.alignChildren = ['left', 'center'];

    // Preview toggle (left)
    cbPreview = btnGroup.add('checkbox', undefined, (lang === 'ja') ? 'プレビュー' : 'Preview');
    cbPreview.value = true;

    // Spacer
    var _sp = btnGroup.add('statictext', undefined, '');
    _sp.alignment = 'fill';

    // Buttons (right)
    var btnCancel = btnGroup.add('button', undefined, L('cancel'), { name: 'cancel' });
    var btnOK = btnGroup.add('button', undefined, L('ok'), { name: 'ok' });
    btnCancel.alignment = 'right';
    btnOK.alignment = 'right';

    btnCancel.onClick = function () {
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
        try { clearPreview(); } catch (_) { }
        dlg.close(0);
    };
    btnOK.onClick = function () {
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
    dlg.onClose = function () {
        try { clearPreview(); } catch (_) { }
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

    function requestPreviewUpdate() {
        try {
            if (cbPreview && cbPreview.value) {
                applyPreviewNow();
            } else {
                clearPreview();
            }
        } catch (_) { }
    }

    function applyPreviewNow() {
        if (!cbPreview || !cbPreview.value) {
            clearPreview();
            return;
        }

        // Rebuild preview from current dialog state
        clearPreview();

        if (app.documents.length === 0) return;
        var _doc = app.activeDocument;
        var _sel = _doc.selection;
        if (!_sel || _sel.length === 0) return;

        try {
            isPreviewing = true;
            previewIsCurrent = false;

            // Use current doc/selection
            doc = _doc;
            sel = _sel;
            baseLayer = doc.activeLayer;

            // Columns + params from UI
            buildCalcColumnsAndParams();

            // Draw preview
            generateMain();
            try { app.redraw(); } catch (_) { }

            // Cleanup calc proxies only (outlined duplicates)
            try {
                var _pc = 0;
                try { _pc = (calcCleanups && typeof calcCleanups.length === 'number') ? calcCleanups.length : 0; } catch (eLen2) { _pc = 0; }
                for (var ii = 0; ii < _pc; ii++) {
                    try {
                        var fn2 = null;
                        try { fn2 = calcCleanups[ii]; } catch (eGet2) { fn2 = null; }
                        if (fn2) { try { fn2(); } catch (eRun2) { } }
                    } catch (_) { }
                }
            } catch (_) { }

            previewIsCurrent = true;
            try { app.redraw(); } catch (_) { }
        } catch (_) {
            try { clearPreview(); } catch (__) { }
            previewIsCurrent = false;
        } finally {
            isPreviewing = false;
        }
    }

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

    // レイヤー作成（現在のレイヤーの背面） / Create layers behind the current layer
    // 罫線レイヤー
    var lineLayer;
    try {
        lineLayer = doc.layers.getByName("罫線レイヤー");
    } catch (e) {
        lineLayer = doc.layers.add();
        lineLayer.name = "罫線レイヤー";
        try { lineLayer.move(baseLayer, ElementPlacement.PLACEAFTER); } catch (_) { }
    }
    // 塗りレイヤー
    var fillLayer;
    try {
        fillLayer = doc.layers.getByName("塗りレイヤー");
    } catch (e2) {
        fillLayer = doc.layers.add();
        fillLayer.name = "塗りレイヤー";
        try { fillLayer.move(lineLayer, ElementPlacement.PLACEAFTER); } catch (_) { }
    }
    try {
        if (lineLayer !== baseLayer) lineLayer.move(baseLayer, ElementPlacement.PLACEAFTER);
    } catch (_) { }
    try {
        if (fillLayer !== lineLayer) fillLayer.move(lineLayer, ElementPlacement.PLACEAFTER);
    } catch (_) { }

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
    // --- Calculation proxy layer and outline proxies for geometricBounds ---
    // Build calculation proxies for selection after dialog confirmation
    // After dlg.show() confirmed:
    // 1. Create temp layer
    // 2. For each selection, duplicate & outline as needed
    // 3. Build srcItems, calcItems, columns based on outlined proxies
    // 4. Clean up all proxies/layer at the end

    // Proxy/calc data
    var srcItems = [];
    var calcProxyList = [];
    var calcCleanups = [];
    var calcLayer = null;
    var columns = [];
    var calcItems = [];

    if (dlg.show() !== 1) return;

    // OK: if preview is on and current, keep it as final (no regeneration)
    if (cbPreview && cbPreview.value && previewIsCurrent) {
        previewItems = [];      // keep objects (do not remove on exit)
        previewIsCurrent = false;
        return;
    }

    function buildCalcColumnsAndParams() {
        // Reset containers
        srcItems = [];
        calcProxyList = [];
        calcCleanups = [];
        calcLayer = null;
        columns = [];
        calcItems = [];

        // 1) temp layer
        var calcLayerName = "__TabularizeCalc__";
        try { calcLayer = doc.layers.getByName(calcLayerName); }
        catch (e) { calcLayer = doc.layers.add(); calcLayer.name = calcLayerName; }
        try { calcLayer.zOrder(ZOrderMethod.SENDTOFRONT); } catch (_) { }
        try { calcLayer.visible = true; } catch (_) { }
        try { calcLayer.locked = false; } catch (_) { }

        // 2) proxies
        var seenSrc = [];
        for (var i = 0; i < sel.length; i++) {
            var it = sel[i];
            var g = getSelectedAncestorGroup(it);
            if (g) it = g;
            if (arrayHasRef(seenSrc, it)) continue;
            seenSrc.push(it);
            srcItems.push(it);

            var proxy = { src: it, calc: it, cleanup: null };

            if (it.typename === "TextFrame") {
                var dup = it.duplicate(calcLayer, ElementPlacement.PLACEATBEGINNING);
                var outlined = null;
                try { outlined = dup.createOutline(); } catch (eO) { outlined = null; }
                try { if (dup && dup.parent) dup.remove(); } catch (_) { }
                if (outlined) {
                    try { outlined.opacity = 0; } catch (_) { }
                    proxy.calc = outlined;
                    proxy.cleanup = (function (gg) { return function () { try { gg.remove(); } catch (_) { } }; })(outlined);
                }
            } else if (it.typename === "GroupItem" && hasAnyTextFrame(it)) {
                var dupg = it.duplicate(calcLayer, ElementPlacement.PLACEATBEGINNING);
                try { while (dupg.textFrames.length > 0) { dupg.textFrames[0].createOutline(); } } catch (_) { }
                try { dupg.opacity = 0; } catch (_) { }
                proxy.calc = dupg;
                proxy.cleanup = (function (gg2) { return function () { try { gg2.remove(); } catch (_) { } }; })(dupg);
            }

            calcProxyList.push(proxy);
            if (proxy.cleanup) calcCleanups.push(proxy.cleanup);
        }

        // 3) columns
        calcItems = [];
        for (var ci = 0; ci < calcProxyList.length; ci++) calcItems.push(calcProxyList[ci].calc);
        calcItems.sort(function (a, b) { return a.geometricBounds[0] - b.geometricBounds[0]; });

        columns = [];
        if (calcItems.length > 0) {
            var currentColumn = [calcItems[0]];
            columns.push(currentColumn);
            for (var j = 1; j < calcItems.length; j++) {
                var item = calcItems[j];
                var itemLeft = item.geometricBounds[0];
                var prevColMaxRight = getMaxRightInColumn(currentColumn);
                if (itemLeft < prevColMaxRight) currentColumn.push(item);
                else { currentColumn = [item]; columns.push(currentColumn); }
            }
        }

        // 4) params from UI
        isZebra = cbZebra.value;
        isFillJoinRow = cbFillJoinRow.value;
        isFillHeaderOnly = cbFillHeaderOnly.value;
        isHeaderRow = cbHeader.value;

        hGutterVal = cbUseGutter.value ? parseFloat(etHGutter.text) : 0;
        if (isNaN(hGutterVal) || hGutterVal < 0) hGutterVal = 0;
        hGutterPt = hGutterVal * rulerFactorPt;

        vRuleMode = rbVruleNone.value ? 'none' : (rbVruleAll.value ? 'all' : 'gapsOnly');

        doFill = cbFill.value;
        doRule = cbRule.value;

        fillMode = (doFill && doRule) ? 'fillAndRule' : (doFill ? 'fillOnly' : 'none');
        KEEP_GAP_PT = hGutterPt;
    }

    try {
        buildCalcColumnsAndParams();
        generateMain();
    } finally {
        // --- Cleanup calc proxies (outlined duplicates only) ---
        try {
            var _cleanupCount = 0;
            try { _cleanupCount = (calcCleanups && typeof calcCleanups.length === 'number') ? calcCleanups.length : 0; } catch (_) { _cleanupCount = 0; }
            for (var cl = 0; cl < _cleanupCount; cl++) {
                try {
                    var fn = null;
                    try { fn = calcCleanups[cl]; } catch (_) { fn = null; }
                    if (fn) { try { fn(); } catch (_) { } }
                } catch (_) { }
            }
        } catch (_) { }
    }


    // 生成処理本体 / Main generation
    function generateMain() {

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
        // 高さが異なるテキストが混在しても、同じ行として扱い、全列で同じYに横罫を引く
        var rowTol = 2.0; // 同一行とみなす中心Yの許容値(pt) ※必要に応じて調整

        // 行（ロウ）定義は「最も行数が多い列」を基準にする
        // 例：別列に複数行をまたぐグループ（背の高い要素）があっても、行グリッドが歪まない
        var baseColIdx = 0;
        var maxCount = -1;
        for (var bc = 0; bc < columns.length; bc++) {
            if (columns[bc].length > maxCount) {
                maxCount = columns[bc].length;
                baseColIdx = bc;
            }
        }

        // 基準列のアイテムだけで行をクラスタリング
        // 行の高さは「その行で一番高さのあるアイテム」を基準にする（基準列内で）
        var allRows = []; // {centerY, maxH}
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
                // centerは軽く追従（安定化）
                allRows[ridx].centerY = (allRows[ridx].centerY + center) / 2;
            }
        }

        // 上→下に並び替え（centerYで）
        allRows.sort(function (a, b) { return b.centerY - a.centerY; });

        // 罫線Yを確定：各行間の中間（全列共通）
        // 行ボックスは centerY ± (maxH/2) で定義（行内で一番高い要素に合わせる）
        var yListGlobal = [];

        function rowTop(i) {
            return allRows[i].centerY + (allRows[i].maxH / 2);
        }
        function rowBottom(i) {
            return allRows[i].centerY - (allRows[i].maxH / 2);
        }

        if (allRows.length >= 2) {
            // 行間の中間（行区切り）
            for (var r = 0; r < allRows.length - 1; r++) {
                var upperBottom = rowBottom(r);
                var lowerTop = rowTop(r + 1);
                var gap = upperBottom - lowerTop;
                var mid = upperBottom - (gap / 2);
                yListGlobal.push(mid);
            }

            // 最上段：次の行間から算出
            var firstGap = rowBottom(0) - rowTop(1);
            yListGlobal.push(rowTop(0) + (firstGap / 2));

            // 最下段：直前の行間から算出
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

        // 近いYを統合してから描画（最後の保険）
        var yTol = 0.4;
        var yUniqGlobal = [];
        for (var yiG = 0; yiG < yListGlobal.length; yiG++) {
            addUniqueY(yUniqGlobal, yListGlobal[yiG], yTol);
        }
        yUniqGlobal.sort(function (a, b) { return b - a; });

        var yTopBorder = (yUniqGlobal.length > 0) ? yUniqGlobal[0] : 0;
        var yBottomBorder = (yUniqGlobal.length > 0) ? yUniqGlobal[yUniqGlobal.length - 1] : 0;

        // 列ごとのセル領域（塗り）の左右境界を作る / Build fill boundaries per column
        // 列境界は中央（xMid）を共有し、塗り側のinsetでガター見かけを作る
        var colFillLeft = [];
        var colFillRight = [];
        var halfGap = 0;

        for (var cc2 = 0; cc2 < columns.length; cc2++) {
            // 左境界
            if (cc2 === 0) {
                // 外側は現行の延長ロジックと整合（padding + 端の伸ばし）
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

            // 右境界
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
                // ヘッダー行のみ：1行目だけ塗る（横方向に連結）
                var xL0 = colFillLeft[0];
                var xR0 = colFillRight[columns.length - 1];

                var yTopH = yUniqGlobal[0];
                var yBotH = yUniqGlobal[1];

                var wH = xR0 - xL0;
                var hH = yTopH - yBotH;
                if (wH > 0 && hH > 0) {
                    var rectH = fillLayer.pathItems.rectangle(yTopH, xL0, wH, hH);
                    rectH.stroked = false;
                    rectH.filled = true;
                    // 塗り色：ヘッダーを優先し、ゼブラONならK50
                    rectH.fillColor = isZebra ? fillGrayHeaderZebra : fillGrayHeader;
                    if (isPreviewing) { try { previewItems.push(rectH); } catch (_) { } }
                }

            } else if (isFillJoinRow) {
                // 行方向に連結：各行1つの矩形（横方向に連結）
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

                    var rect = fillLayer.pathItems.rectangle(top, left, w, h);
                    rect.stroked = false;
                    rect.filled = true;

                    // 塗り色：ヘッダーを優先し、ゼブラONなら奇数行をK30
                    if (isHeaderRow && fr === 0) {
                        rect.fillColor = isZebra ? fillGrayHeaderZebra : fillGrayHeader;
                    } else if (isZebra && ((fr + 1) % 2 === 1)) {
                        rect.fillColor = fillGrayZebra;
                    } else {
                        rect.fillColor = fillGray;
                    }
                    if (isPreviewing) { try { previewItems.push(rect); } catch (_) { } }
                }

            } else {
                // 通常：セルごと
                // 見かけのガターを統一：列間・段間ともに KEEP_GAP_PT/2 にする
                // 各方向をinsetずつ詰めるので、間隔は 2*inset になる
                var insetX = KEEP_GAP_PT / 2;
                var insetY = KEEP_GAP_PT / 2;

                for (var fc = 0; fc < columns.length; fc++) {
                    var xL = colFillLeft[fc];
                    var xR = colFillRight[fc];

                    for (var fr2 = 0; fr2 < yUniqGlobal.length - 1; fr2++) {
                        var yTop2 = yUniqGlobal[fr2];
                        var yBot2 = yUniqGlobal[fr2 + 1];

                        // 内側へ詰める（負のオフセット相当）
                        var left2 = xL + insetX;
                        var right2 = xR - insetX;
                        var top2 = yTop2 - insetY;
                        var bottom2 = yBot2 + insetY;

                        var w2 = right2 - left2;
                        var h2 = top2 - bottom2;
                        if (w2 <= 0 || h2 <= 0) continue;

                        var rect2 = fillLayer.pathItems.rectangle(top2, left2, w2, h2);
                        rect2.stroked = false;
                        rect2.filled = true;

                        // 塗り色：ヘッダーを優先し、ゼブラONなら奇数行をK30
                        if (isHeaderRow && fr2 === 0) {
                            rect2.fillColor = isZebra ? fillGrayHeaderZebra : fillGrayHeader;
                        } else if (isZebra && ((fr2 + 1) % 2 === 1)) {
                            rect2.fillColor = fillGrayZebra;
                        } else {
                            rect2.fillColor = fillGray;
                        }
                        if (isPreviewing) { try { previewItems.push(rect2); } catch (_) { } }
                    }
                }
            }
        }

        // --- 横罫（行ごと） ---
        if (doRule) {
            // ガターOFF（=0）なら、横ケイも表全体で連結して1本にする
            if (!cbUseGutter.value || KEEP_GAP_PT === 0) {
                // 表全体の左右端（外側の伸ばしも考慮）
                var minXLeft = colBounds[0].minX;
                var maxXLeft = colBounds[0].maxX;
                var minXRight = colBounds[columns.length - 1].minX;
                var maxXRight = colBounds[columns.length - 1].maxX;

                var extLeft = 0;
                var extRight = 0;

                if (columns.length >= 2) {
                    // 左端：1列目と2列目の間隔Aから B=(A-KEEP)/2
                    var A0 = colBounds[1].minX - maxXLeft;
                    extLeft = (A0 - KEEP_GAP_PT) / 2;
                    if (extLeft < 0) extLeft = 0;

                    // 右端：最終-1列目と最終列目の間隔Aから B=(A-KEEP)/2
                    var lastIdx = columns.length - 1;
                    var An = minXRight - colBounds[lastIdx - 1].maxX;
                    extRight = (An - KEEP_GAP_PT) / 2;
                    if (extRight < 0) extRight = 0;
                }

                var globalStartX = minXLeft - paddingPt - extLeft;
                var globalEndX = maxXRight + paddingPt + extRight;

                for (var iLine = 0; iLine < yUniqGlobal.length; iLine++) {
                    // ヘッダーON時は「上から1本目と2本目」だけ 0.25mm、それ以外は基本 0.1mm
                    var strokeW = lineWeightPt;
                    if (isHeaderRow && iLine < 2) strokeW = headerLineWeightPt;
                    drawLine(globalStartX, yUniqGlobal[iLine], globalEndX, yUniqGlobal[iLine], strokeW);
                }

            } else {
                // ガターON：従来どおり、列ごとに線長を計算して描画
                for (var c = 0; c < columns.length; c++) {
                    var colItems = columns[c];

                    var minX = colBounds[c].minX;
                    var maxX = colBounds[c].maxX;

                    // A：列間を計算（左/右）
                    // B = (A - 1mm) / 2 （マイナスなら0）
                    var extendLeft = 0;
                    var extendRight = 0;

                    // 内側（隣接列に向かう側）の伸ばし：列間Aから B=(A-1mm)/2 を算出
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

                    // 外側（端）の伸ばし：
                    // 1列目の左は「1列目と2列目の間隔A」から B=(A-1mm)/2
                    if (c === 0 && columns.length >= 2) {
                        var nextMinX0 = colBounds[1].minX;
                        var A0a = nextMinX0 - maxX;
                        var ext0 = (A0a - KEEP_GAP_PT) / 2;
                        if (ext0 < 0) ext0 = 0;
                        extendLeft = ext0;
                    }

                    // 最終列の右は「最終-1列目と最終列目の間隔A」から B=(A-1mm)/2
                    if (c === columns.length - 1 && columns.length >= 2) {
                        var prevMaxXn = colBounds[columns.length - 2].maxX;
                        var An2 = minX - prevMaxXn;
                        var extn = (An2 - KEEP_GAP_PT) / 2;
                        if (extn < 0) extn = 0;
                        extendRight = extn;
                    }

                    // パディング＋延長を適用
                    var lineStartX = minX - paddingPt - extendLeft;
                    var lineEndX = maxX + paddingPt + extendRight;

                    // 横罫線は全列共通のY（yUniqGlobal）で描画する
                    for (var iLine2 = 0; iLine2 < yUniqGlobal.length; iLine2++) {
                        var strokeW2 = lineWeightPt;
                        if (isHeaderRow && iLine2 < 2) strokeW2 = headerLineWeightPt;
                        drawLine(lineStartX, yUniqGlobal[iLine2], lineEndX, yUniqGlobal[iLine2], strokeW2);
                    }
                }
            }
        }

        // --- 縦罫（列間のみ / すべて） ---
        if (doRule && vRuleMode !== 'none' && yUniqGlobal && yUniqGlobal.length > 0 && columns.length >= 2) {
            // 縦罫のX位置（列境界）を作る
            var vXs = [];

            // 列間のみ：列間（ガターの中央）
            for (var cc = 0; cc < columns.length - 1; cc++) {
                var xMid = (colBounds[cc].maxX + colBounds[cc + 1].minX) / 2;
                vXs.push(xMid);
            }

            // すべて：外枠（左/右）も追加
            if (vRuleMode === 'all') {
                // 左端は「1列目と2列目の間隔A」から B=(A-KEEP)/2 を算出して外側に広げる
                var leftMinX = colBounds[0].minX;
                var leftExtend = 0;
                {
                    var A0 = colBounds[1].minX - colBounds[0].maxX;
                    leftExtend = (A0 - KEEP_GAP_PT) / 2;
                    if (leftExtend < 0) leftExtend = 0;
                }
                var xLeftBorder = leftMinX - paddingPt - leftExtend;

                // 右端は「最終-1列目と最終列目の間隔A」から B=(A-KEEP)/2
                var lastIdx = columns.length - 1;
                var rightMaxX = colBounds[lastIdx].maxX;
                var rightExtend = 0;
                {
                    var An = colBounds[lastIdx].minX - colBounds[lastIdx - 1].maxX;
                    rightExtend = (An - KEEP_GAP_PT) / 2;
                    if (rightExtend < 0) rightExtend = 0;
                }
                var xRightBorder = rightMaxX + paddingPt + rightExtend;

                // 先頭/末尾に追加（重複しない順番で）
                vXs.unshift(xLeftBorder);
                vXs.push(xRightBorder);
            }

            // 縦罫を描画
            if (!cbUseGutter.value) {
                // ガターOFF：縦罫は連結して1本で描画
                for (var vx = 0; vx < vXs.length; vx++) {
                    drawLine(vXs[vx], yTopBorder, vXs[vx], yBottomBorder, lineWeightPt);
                }
            } else {
                // ガターON：セルごとに分割して描画（既存挙動）
                // 端をvTrimずつ詰める（段間=2*vTrim）
                var vTrim = KEEP_GAP_PT / 2;
                for (var vx = 0; vx < vXs.length; vx++) {
                    for (var ry = 0; ry < yUniqGlobal.length - 1; ry++) {
                        var y1 = yUniqGlobal[ry];
                        var y2 = yUniqGlobal[ry + 1];

                        // y1 は上、y2 は下（y1 > y2）の想定
                        var segH = y1 - y2;
                        if (segH <= vTrim * 2) {
                            // 短すぎる場合は無理に詰めない（0長さや逆転防止）
                            drawLine(vXs[vx], y1, vXs[vx], y2, lineWeightPt);
                            continue;
                        }

                        var ys = y1 - vTrim;
                        var ye = y2 + vTrim;
                        drawLine(vXs[vx], ys, vXs[vx], ye, lineWeightPt);
                    }
                }
            }
        }

    } // end generateMain

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

    // ExtendScript互換：参照配列にobjが含まれるか（indexOfが無い環境向け）
    function arrayHasRef(arr, obj) {
        for (var i = 0; i < arr.length; i++) {
            if (arr[i] === obj) return true;
        }
        return false;
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

    // グループまたは子孫にTextFrameが含まれるか
    function hasAnyTextFrame(item) {
        if (!item) return false;
        if (item.typename === "TextFrame") return true;
        if (item.typename === "GroupItem") {
            if (item.textFrames && item.textFrames.length > 0) return true;
            // Recursively check subgroups
            for (var i = 0; i < item.groupItems.length; i++) {
                if (hasAnyTextFrame(item.groupItems[i])) return true;
            }
        }
        return false;
    }

    // 線描画関数
    function drawLine(x1, y1, x2, y2, strokeW) {
        var pathItem = lineLayer.pathItems.add();
        pathItem.setEntirePath([[x1, y1], [x2, y2]]);
        pathItem.stroked = true;

        // 端部が重なってガターが黒く見えるのを防ぐ（端部をバットに固定）
        try { pathItem.strokeCap = StrokeCap.BUTTENDCAP; } catch (_) { }
        try { pathItem.strokeJoin = StrokeJoin.MITERENDJOIN; } catch (_) { }

        pathItem.strokeWidth = (typeof strokeW === 'number') ? strokeW : lineWeightPt;
        pathItem.strokeColor = blackColor;
        pathItem.filled = false;
        if (isPreviewing) {
            try { previewItems.push(pathItem); } catch (_) { }
        }
    }

})();
