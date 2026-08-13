#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したパス（オープン／クローズ）の長さをもとに、分割数と間隔から破線の線分長を計算して適用するツール。
線分から間隔を逆算するモードやランダムパターン、開始位置（位相）の指定にも対応する。

詳細はREADMEを参照。

*/

/*

### Overview

Applies dashes to the selected open or closed paths by calculating the dash length from the number
of segments and the gap. Also supports solving the gap from a given dash length, random patterns,
and a dash offset (phase).

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "DashGapCalculator";            /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v2.0.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-02-25";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-08-13";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/DashGapCalculator.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/DashGapCalculator.md
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n868bedb96542"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* ランダムモードで生成する線分・間隔の範囲（現在の線の単位）/ Random dash & gap range in the current stroke unit */
    var RANDOM_DASH_MAX     = 40;    /* 線分の最大値 / max dash */
    var RANDOM_GAP_MIN      = 3;     /* 間隔の最小値 / min gap */
    var RANDOM_GAP_MAX      = 3;     /* 間隔の最大値 / max gap */
    var RANDOM_ROUND_VALUES = false; /* 生成値を整数に丸めるか / round generated values */

    /* 線端「なし」のときに線分が消えないための最小値（単位コード別）/ Minimum dash for butt caps, per unit code */
    var RANDOM_DASH_MIN_BY_UNIT = {
        1: 1, /* mm */
        2: 2, /* pt */
        5: 4  /* Q / H */
    };
    var RANDOM_DASH_MIN_DEFAULT = 2;

    /* ランダムパターンの要素数（線分・間隔を3組）/ Entries in one random pattern */
    var RANDOM_PATTERN_LENGTH = 6;

    // =========================================
    // UIレイアウトの共通設定 / Shared UI layout
    // =========================================

    /* ウィンドウ・パネルの余白と間隔 / Window & panel margins and spacing */
    var WINDOW_MARGINS = 16;                 /* ウィンドウ外周の余白 / window margin */
    var WINDOW_SPACING = 12;                 /* ウィンドウ内の要素間隔 / window spacing */
    var PANEL_MARGINS  = [16, 20, 16, 12];   /* パネル余白 [左,上,右,下] / panel margins */
    var PANEL_SPACING  = 12;                 /* パネル内の要素間隔 / panel spacing */
    var COLUMN_SPACING = 12;                 /* 2カラムの間隔 / gap between columns */

    /* このダイアログ固有の寸法 / Sizes specific to this dialog */
    var BUTTON_ROW_MARGINS       = [0, 5, 0, 0]; /* ボタンエリアの余白 [左,上,右,下] */
    var OPTION_ROW_SPACING       = 20;           /* 下部チェックボックスの間隔 */
    var OFFSET_PANEL_SPACING     = 6;            /* 開始位置パネルの要素間隔（密） */
    var FIELD_LABEL_WIDTH        = 40;           /* 分割数・間隔・線分のラベル幅 */
    var NUMBER_FIELD_CHARS       = 4;            /* 数値入力欄の幅（文字数） */
    var LABELLESS_CHECKBOX_WIDTH = 18;           /* ラベルなしチェックボックスの幅 */

    /**
     * ウィンドウへ共通のレイアウトを適用する
     * @param {Window} win - 対象ダイアログ
     * @param {number} spacing - 要素間隔（省略時は WINDOW_SPACING）
     * @returns {void}
     */
    function setupWindow(win, spacing) {
        win.orientation = "column";
        win.alignChildren = "fill";
        win.margins = WINDOW_MARGINS;
        win.spacing = (typeof spacing === "number") ? spacing : WINDOW_SPACING;
    }

    /**
     * パネルへ共通のレイアウトを適用する
     * @param {Panel} panel - 対象パネル
     * @param {number} spacing - 要素間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function setupPanel(panel, spacing) {
        panel.orientation = "column";
        panel.alignChildren = ["fill", "top"];
        panel.alignment = "fill";
        panel.margins = PANEL_MARGINS;
        panel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /**
     * 行グループへ共通のレイアウトを適用する
     * @param {Group} group - 対象グループ
     * @param {string} alignment - 親の中での配置（省略時は "left"）
     * @param {number} spacing - 要素間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function setupRow(group, alignment, spacing) {
        group.orientation = "row";
        group.alignment = alignment || "left";
        group.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /**
     * ラベル付きパネルを生成し、共通レイアウトを適用する
     * @param {Window|Panel|Group} parent - 追加先のコンテナ
     * @param {string} titleText - パネルのタイトル
     * @param {number} spacing - 要素間隔（省略時は PANEL_SPACING）
     * @returns {Panel} 生成したパネル
     */
    function addPanel(parent, titleText, spacing) {
        var panel = parent.add("panel", undefined, titleText);
        setupPanel(panel, spacing);
        return panel;
    }

    /**
     * 横並びのグループを生成する
     * @param {Window|Panel|Group} parent - 追加先のコンテナ
     * @param {string} alignment - 親の中での配置（省略時は "left"）
     * @param {number} spacing - 要素間隔（省略時は PANEL_SPACING）
     * @returns {Group} 生成したグループ
     */
    function addRow(parent, alignment, spacing) {
        var row = parent.add("group");
        setupRow(row, alignment, spacing);
        row.alignChildren = ["left", "center"];
        return row;
    }

    /**
     * 縦積みのグループを生成する
     * @param {Window|Panel|Group} parent - 追加先のコンテナ
     * @param {Array<string>} alignChildren - 子要素の整列指定（省略時は ["fill", "top"]）
     * @param {string} alignment - 親の中での配置（省略時はコンテナ既定）
     * @returns {Group} 生成したグループ
     */
    function addColumn(parent, alignChildren, alignment) {
        var column = parent.add("group");
        column.orientation = "column";
        column.alignChildren = alignChildren || ["fill", "top"];
        if (alignment) column.alignment = alignment;
        return column;
    }

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * 実行環境のロケールから表示言語を判定する
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        return ($.locale && $.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = getCurrentLang();

    /* カテゴリ分けした日英ラベル定義 / Categorized Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "破線計算機", en: "Dash Calculator (Gap→Dash)" }
        },
        panel: {
            pathInfo:   { ja: "選択中のパス情報", en: "Selected Path Info" },
            dashCalc:   { ja: "破線の計算", en: "Dash Calculation" },
            calcMethod: { ja: "計算方法", en: "Calculation" },
            offset:     { ja: "開始位置", en: "Offset" },
            partial:    { ja: "部分表示", en: "Partial Display" },
            cap:        { ja: "線端", en: "Cap" }
        },
        fieldLabel: {
            pathLength: { ja: "パスの長さ:", en: "Path length:" },
            segments:   { ja: "分割数:", en: "Segments:" },
            gap:        { ja: "間隔:", en: "Gap:" },
            dash:       { ja: "線分:", en: "Dash:" }
        },
        radio: {
            gapToDash:  { ja: "間隔→線分", en: "Gap→Dash" },
            dashToGap:  { ja: "線分→間隔", en: "Dash→Gap" },
            random:     { ja: "ランダム", en: "Random" },
            capButt:    { ja: "なし", en: "Butt" },
            capRound:   { ja: "丸型", en: "Round" },
            capProject: { ja: "突出", en: "Projecting" }
        },
        checkbox: {
            partialDisplay: { ja: "部分表示", en: "Partial Display" },
            adjustEnds:     { ja: "両端を調整", en: "Adjust ends" },
            reversePath:    { ja: "パスの方向反転", en: "Reverse Path Direction" }
        },
        button: {
            ok:        { ja: "OK", en: "OK" },
            cancel:    { ja: "キャンセル", en: "Cancel" },
            clearDash: { ja: "破線クリア", en: "Clear Dashes" }
        },
        tooltip: {
            pathInfo: {
                ja: "先頭のパスの長さ。括弧内は選択しているパスの数",
                en: "Length of the first path. The number in parentheses is the count of selected paths"
            },
            segments: {
                ja: "パスをいくつに分けるか。線分＋間隔の繰り返し回数になります",
                en: "How many parts the path is divided into — the number of dash + gap cycles"
            },
            gap: {
                ja: "破線のすき間の長さ",
                en: "Length of the empty space between dashes"
            },
            dash: {
                ja: "破線の線の長さ",
                en: "Length of each dash"
            },
            gapToDash: {
                ja: "間隔を入力して線分の長さを求めます",
                en: "Enter the gap; the dash length is calculated"
            },
            dashToGap: {
                ja: "線分の長さを入力して間隔を求めます",
                en: "Enter the dash length; the gap is calculated"
            },
            random: {
                ja: "線分をランダムに割り当てます。パスごとに別の乱数を使い、クリックのたびに作り直します",
                en: "Assigns random dashes. Each path gets its own draw, and every click generates a new pattern"
            },
            useOffset: {
                ja: "破線の開始位置（位相）を指定します。パスの始点から数えます",
                en: "Sets the dash offset (phase), measured from the start point of the path"
            },
            offsetPreset: {
                ja: "1周期（線分＋間隔）に対する比率で開始位置を決めます",
                en: "Sets the offset as a fraction of one cycle (dash + gap)"
            },
            partialDisplay: {
                ja: "分割した1本分の線分だけを表示し、残りを隠します",
                en: "Shows only one of the divided dashes and hides the rest"
            },
            cap: {
                ja: "破線の線端の形。丸型・突出は線分の長さより少しはみ出します",
                en: "Shape of the dash ends. Round and Projecting extend slightly beyond the dash length"
            },
            adjustEnds: {
                ja: "オープンパスで、両端が線分で終わるように配分します（クローズパスでは使いません）",
                en: "On an open path, distributes the dashes so both ends finish with a dash (unused for closed paths)"
            },
            reversePath: {
                ja: "パスの向きを反転して、破線の開始位置を反対の端へ移します",
                en: "Reverses the path direction, moving the dash start to the other end"
            },
            clearDash: {
                ja: "破線を解除した状態をプレビューします。OKで確定します",
                en: "Previews the paths without dashes. Click OK to confirm"
            }
        },
        alert: {
            calcError:       { ja: "エラー", en: "Error" },
            noDocument:      { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noSelection:     { ja: "対象となるパス（線や円など）を選択してください。", en: "Select one path (open or closed)." },
            noPathSelected:  { ja: "パス（オープン/クローズ）を1つ選択してください。", en: "Select exactly one path (open or closed)." },
            segmentsInvalid: { ja: "分割数は1以上の整数を入力してください。", en: "Enter an integer of 1 or greater for Segments." },
            gapInvalid:      { ja: "間隔 (Gap) は0以上の数値を入力してください。", en: "Enter a number of 0 or greater for Gap." },
            dashInvalid:     { ja: "線分 (Dash) は0以上の数値を入力してください。", en: "Enter a number of 0 or greater for Dash." },
            offsetInvalid:   { ja: "開始位置 (Offset) は数値を入力してください。", en: "Enter a number for Offset." },
            gapTooLong: {
                ja: "間隔 (Gap) が長すぎます。線分がゼロまたはマイナスになってしまいます。\n(設定可能な最大Gap: ほぼ {0} {1})",
                en: "Gap is too long; dash would be zero or negative.\n(Max allowed Gap: about {0} {1})"
            },
            dashTooLong: {
                ja: "線分 (Dash) が長すぎます。間隔がゼロまたはマイナスになってしまいます。\n(設定可能な最大Dash: ほぼ {0} {1})",
                en: "Dash is too long; gap would be zero or negative.\n(Max allowed Dash: about {0} {1})"
            }
        }
    };

    /**
     * 表示言語に応じたラベル文字列を返す
     * @param {Object} labelNode - LABELS 内の { ja, en } ノード
     * @returns {string} 表示用の文言
     */
    function getLabel(labelNode) {
        if (!labelNode) return "";
        return labelNode[uiLang] || labelNode.en || labelNode.ja || "";
    }

    /**
     * ラベル内の {0} {1} … を引数で置き換える
     * @param {Object} labelNode - LABELS 内の { ja, en } ノード
     * @param {Array<string>} args - 差し込む文字列の配列
     * @returns {string} 置き換え後の文言
     */
    function formatLabel(labelNode, args) {
        var labelText = getLabel(labelNode);
        if (!args) return labelText;
        for (var i = 0; i < args.length; i++) {
            labelText = labelText.split("{" + i + "}").join(String(args[i]));
        }
        return labelText;
    }

    // =========================================
    // 単位ユーティリティ / Unit utilities
    // =========================================

    /* 環境設定の単位コードと表記の対応 / Unit code to label */
    var UNIT_LABELS = {
        0: "in",
        1: "mm",
        2: "pt",
        3: "pica",
        4: "cm",
        5: "Q/H",
        6: "px",
        7: "ft/in",
        8: "m",
        9: "yd",
        10: "ft"
    };

    /**
     * 単位コードから pt 換算係数（1単位あたりのpt数）を返す
     * @param {number} unitCode - 環境設定の単位コード
     * @returns {number} pt 換算係数
     */
    function unitCodeToPtFactor(unitCode) {
        switch (unitCode) {
            case 0: return 72;                 /* in */
            case 1: return 72 / 25.4;          /* mm */
            case 2: return 1;                  /* pt */
            case 3: return 12;                 /* pica */
            case 4: return 72 / 2.54;          /* cm */
            case 5: return (72 / 25.4) * 0.25; /* Q/H（1Q=1H=0.25mm） */
            case 6: return 1;                  /* px（Illustratorでは 1px=1pt） */
            case 7: return 72;                 /* ft/in（複合表記は扱わず in 相当） */
            case 8: return 72 / 0.0254;        /* m */
            case 9: return 72 * 36;            /* yd */
            case 10: return 72 * 12;           /* ft */
            default: return 1;
        }
    }

    /**
     * 線の単位（strokeUnits）の情報を取得する
     * @returns {Object} { code:number, label:string, factor:number }
     */
    function getStrokeUnitInfo() {
        var unitCode = 2; /* 取得できない場合は pt */
        var asianUnits = 0;
        try {
            unitCode = app.preferences.getIntegerPreference("strokeUnits");
            /* Q/H は環境設定（東アジア言語）に合わせて表記を切り替える */
            if (unitCode === 5) asianUnits = app.preferences.getIntegerPreference("text/asianunits");
        } catch (e) { }

        var unitLabel = UNIT_LABELS[unitCode] || "pt";
        if (unitCode === 5) unitLabel = (asianUnits === 1) ? "H" : "Q";

        return {
            code: unitCode,
            label: unitLabel,
            factor: unitCodeToPtFactor(unitCode)
        };
    }

    /**
     * 単位値を pt に変換する
     * @param {number} value - 単位値
     * @param {Object} unitInfo - getStrokeUnitInfo() の戻り値
     * @returns {number} pt 値
     */
    function unitToPt(value, unitInfo) {
        return value * unitInfo.factor;
    }

    /**
     * pt 値を単位値に変換する
     * @param {number} ptValue - pt 値
     * @param {Object} unitInfo - getStrokeUnitInfo() の戻り値
     * @returns {number} 単位値
     */
    function ptToUnit(ptValue, unitInfo) {
        return ptValue / unitInfo.factor;
    }

    /**
     * 入力欄向けに数値を整形する（小数第3位まで、整数はそのまま）
     * @param {number} value - 表示したい数値
     * @returns {string} 整形後の文字列
     */
    function formatFieldNumber(value) {
        if (value == null || isNaN(value)) return "0";
        var rounded = Math.round(value * 1000) / 1000;
        if (Math.abs(rounded - Math.round(rounded)) < 1e-10) return String(Math.round(rounded));
        return String(rounded);
    }

    // =========================================
    // 前回値の記憶 / Session settings
    // =========================================

    var PREF_KEY = "DashCalcPrefs_GapToDash_v1";

    /* ランダムパターン保存用のキー / Keys for the random pattern */
    var RANDOM_PREF_KEYS = ["rand0Pt", "rand1Pt", "rand2Pt", "rand3Pt", "rand4Pt", "rand5Pt"];

    /**
     * 前回値を読み込む
     * @returns {Object} 保存値のオブジェクト（読み込めない場合は null）
     */
    function loadPrefs() {
        try {
            var descriptor = app.getCustomOptions(PREF_KEY);
            var savedPrefs = {};

            var kSegments = stringIDToTypeID("segments");
            var kGapPt = stringIDToTypeID("gapPt");
            var kDashPt = stringIDToTypeID("dashPt");
            var kOffsetPt = stringIDToTypeID("offsetPt");
            var kCapMode = stringIDToTypeID("capMode");
            var kMode = stringIDToTypeID("mode");
            var kReverse = stringIDToTypeID("reversePath");
            var kAdjustEnds = stringIDToTypeID("adjustEnds");
            var kUseOffset = stringIDToTypeID("useOffset");

            if (descriptor.hasKey(kSegments)) savedPrefs.segments = descriptor.getInteger(kSegments);
            if (descriptor.hasKey(kGapPt)) savedPrefs.gapPt = descriptor.getDouble(kGapPt);
            if (descriptor.hasKey(kDashPt)) savedPrefs.dashPt = descriptor.getDouble(kDashPt);
            if (descriptor.hasKey(kOffsetPt)) savedPrefs.offsetPt = descriptor.getDouble(kOffsetPt);
            if (descriptor.hasKey(kCapMode)) savedPrefs.capMode = descriptor.getInteger(kCapMode);
            if (descriptor.hasKey(kMode)) savedPrefs.mode = descriptor.getInteger(kMode);
            if (descriptor.hasKey(kReverse)) savedPrefs.reversePath = descriptor.getBoolean(kReverse);
            if (descriptor.hasKey(kAdjustEnds)) savedPrefs.adjustEnds = descriptor.getBoolean(kAdjustEnds);
            if (descriptor.hasKey(kUseOffset)) savedPrefs.useOffset = descriptor.getBoolean(kUseOffset);

            /* ランダムパターンは先頭から連続している分だけ読む（旧バージョンの4要素も許容） */
            var randomDashes = [];
            for (var i = 0; i < RANDOM_PREF_KEYS.length; i++) {
                var randomKey = stringIDToTypeID(RANDOM_PREF_KEYS[i]);
                if (!descriptor.hasKey(randomKey)) break;
                randomDashes.push(descriptor.getDouble(randomKey));
            }
            if (randomDashes.length > 0) savedPrefs.randPt = randomDashes;

            return savedPrefs;
        } catch (e) {
            return null;
        }
    }

    /**
     * 前回値を保存する
     * @param {Object} newPrefs - 保存する設定値
     * @returns {void}
     */
    function savePrefs(newPrefs) {
        try {
            var descriptor = new ActionDescriptor();
            descriptor.putInteger(stringIDToTypeID("segments"), newPrefs.segments);
            descriptor.putDouble(stringIDToTypeID("gapPt"), newPrefs.gapPt);
            descriptor.putDouble(stringIDToTypeID("dashPt"), newPrefs.dashPt);
            descriptor.putDouble(stringIDToTypeID("offsetPt"), newPrefs.offsetPt);
            descriptor.putInteger(stringIDToTypeID("capMode"), newPrefs.capMode);
            descriptor.putInteger(stringIDToTypeID("mode"), newPrefs.mode);
            descriptor.putBoolean(stringIDToTypeID("reversePath"), !!newPrefs.reversePath);
            descriptor.putBoolean(stringIDToTypeID("adjustEnds"), !!newPrefs.adjustEnds);
            descriptor.putBoolean(stringIDToTypeID("useOffset"), !!newPrefs.useOffset);

            if (newPrefs.randPt && newPrefs.randPt.length === RANDOM_PATTERN_LENGTH) {
                for (var i = 0; i < RANDOM_PREF_KEYS.length; i++) {
                    descriptor.putDouble(stringIDToTypeID(RANDOM_PREF_KEYS[i]), newPrefs.randPt[i]);
                }
            }
            app.putCustomOptions(PREF_KEY, descriptor, true);
        } catch (e) { }
    }

    // =========================================
    // 破線の計算 / Dash calculation
    // =========================================

    /**
     * 間隔から線分長と1周期（線分＋間隔）の長さを求める
     * クローズパスと「両端を調整」OFFは 1周期＝全長÷分割数、
     * オープンパスで「両端を調整」ONは両端が線分で終わるように配分する。
     * @param {number} segments - 分割数
     * @param {number} gapPt - 間隔（pt）
     * @param {number} pathLen - パスの長さ（pt）
     * @param {boolean} isClosed - クローズパスかどうか
     * @param {boolean} adjustEnds - 両端を調整するかどうか
     * @returns {Object} { dashPt:number, cyclePt:number }（計算できない場合は null）
     */
    function calcDashAndCyclePt(segments, gapPt, pathLen, isClosed, adjustEnds) {
        if (!(segments > 0)) return null;
        if (!(gapPt >= 0)) return null;

        if (isClosed || !adjustEnds) {
            var cyclePt = pathLen / segments;
            return { dashPt: cyclePt - gapPt, cyclePt: cyclePt };
        }

        /* 分割数＝線分の本数、間隔は（分割数−1）回 */
        if (segments === 1) return { dashPt: pathLen, cyclePt: pathLen + gapPt };

        var dashPt = (pathLen - gapPt * (segments - 1)) / segments;
        return { dashPt: dashPt, cyclePt: dashPt + gapPt };
    }

    /**
     * 線分長から間隔を逆算する
     * @param {number} segments - 分割数
     * @param {number} dashPt - 線分長（pt）
     * @param {number} pathLen - パスの長さ（pt）
     * @param {boolean} isClosed - クローズパスかどうか
     * @param {boolean} adjustEnds - 両端を調整するかどうか
     * @returns {number} 間隔（pt）。計算できない場合は null
     */
    function calcGapPtFromDashPt(segments, dashPt, pathLen, isClosed, adjustEnds) {
        if (!(segments > 0)) return null;
        if (!(dashPt >= 0)) return null;

        if (isClosed || !adjustEnds) return (pathLen / segments) - dashPt;

        if (segments === 1) return 0;
        return (pathLen - dashPt * segments) / (segments - 1);
    }

    // =========================================
    // 入力欄のキー操作 / Arrow key handling
    // =========================================

    /**
     * 入力欄に↑↓キーでの数値増減を設定する
     * ↑↓は±1、Shift+↑↓は±10（10の倍数にスナップ）、Option(Alt)+↑↓は±0.1。
     * @param {EditText} editText - 対象の入力欄
     * @param {boolean} forceInteger - 整数のみ扱うかどうか
     * @param {number} minValue - 下限値
     * @returns {void}
     */
    function changeValueByArrowKey(editText, forceInteger, minValue) {
        editText._forceInteger = !!forceInteger;
        editText._minValue = minValue;

        editText.addEventListener("keydown", function (event) {
            if (event.keyName !== "Up" && event.keyName !== "Down") return;

            var value = Number(editText.text);
            if (isNaN(value)) return;

            var keyboard = ScriptUI.environment.keyboardState;
            var useDecimal = keyboard.altKey && !editText._forceInteger;
            var isUp = (event.keyName === "Up");

            if (keyboard.shiftKey) {
                value = isUp ? Math.ceil((value + 1) / 10) * 10 : Math.floor((value - 1) / 10) * 10;
            } else if (useDecimal) {
                value += isUp ? 0.1 : -0.1;
            } else {
                value += isUp ? 1 : -1;
            }
            event.preventDefault();

            if (useDecimal) {
                value = Math.round(value * 10) / 10; /* 小数第1位まで */
                if (Math.abs(value) < 0.0000001) value = 0; /* -0 対策 */
            } else {
                value = Math.round(value);
            }
            if (value < editText._minValue) value = editText._minValue;

            editText.text = String(value);

            if (typeof editText._onArrowChange === "function") editText._onArrowChange();
        });
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * バウンディングボックスをリセットし、エッジ表示を切り替える
     * 実行時と終了時に呼び、表示を元へ戻す。
     * @returns {void}
     */
    function resetBoundsAndToggleEdges() {
        app.executeMenuCommand('AI Reset Bounding Box');
        app.executeMenuCommand('edge');
    }

    /**
     * 選択の中から PathItem だけを集める
     * @param {Array} selection - ドキュメントの選択
     * @returns {Array<PathItem>} 対象のパス
     */
    function collectTargetPaths(selection) {
        var paths = [];
        for (var i = 0; i < selection.length; i++) {
            if (selection[i] && selection[i].typename === "PathItem") paths.push(selection[i]);
        }
        return paths;
    }

    /**
     * ダイアログを表示し、対象のパスに破線を適用する
     * @param {Document} doc - 対象ドキュメント
     * @param {Array<PathItem>} targetPaths - 対象のパス
     * @returns {void}
     */
    function showDashDialog(doc, targetPaths) {
        /* 先頭のパスをUI表示・計算の代表として扱う */
        var primaryPath = targetPaths[0];
        var primaryPathLength = primaryPath.length;
        var strokeUnit = getStrokeUnitInfo();

        /* ダイアログを開く前の状態（キャンセル時に復元）*/
        var originalStates = [];
        for (var i = 0; i < targetPaths.length; i++) {
            var pathItem = targetPaths[i];
            originalStates.push({
                item: pathItem,
                stroked: pathItem.stroked,
                strokeCap: pathItem.strokeCap,
                strokeDashes: (pathItem.strokeDashes && pathItem.strokeDashes.length) ? pathItem.strokeDashes.slice(0) : [],
                strokeDashOffset: (typeof pathItem.strokeDashOffset === "number") ? pathItem.strokeDashOffset : 0
            });
        }

        var closedByOK = false;
        var directionReversed = false;
        var isDashCleared = false;

        /* ランダムパターン（線分・間隔を3組、pt）/ Random pattern in pt */
        var randomDashesPt = null;

        // -----------------------------------------
        // 前回値の復元 / Restore previous values
        // -----------------------------------------

        var prefs = loadPrefs();

        /**
         * 前回値の長さ（pt）を現在の単位に変換して返す
         * @param {number} storedPt - 保存されている長さ（pt）
         * @param {number} fallbackUnit - 保存値がないときの値（単位値）
         * @returns {number} 長さ（単位値）
         */
        function toInitialUnit(storedPt, fallbackUnit) {
            return (typeof storedPt === "number" && storedPt >= 0) ? ptToUnit(storedPt, strokeUnit) : fallbackUnit;
        }

        /**
         * 前回値の真偽値を返す
         * @param {boolean} storedFlag - 保存されている値
         * @param {boolean} fallbackFlag - 保存値がないときの値
         * @returns {boolean} 復元した値
         */
        function toInitialFlag(storedFlag, fallbackFlag) {
            return (typeof storedFlag === "boolean") ? storedFlag : fallbackFlag;
        }

        /**
         * 前回値の数値を返す
         * @param {number} storedNumber - 保存されている値
         * @param {number} fallbackNumber - 保存値がないときの値
         * @param {number} minValue - 許容する下限値
         * @returns {number} 復元した値
         */
        function toInitialNumber(storedNumber, fallbackNumber, minValue) {
            return (typeof storedNumber === "number" && storedNumber >= minValue) ? storedNumber : fallbackNumber;
        }

        var initialSegments   = toInitialNumber(prefs && prefs.segments, 3, 1);
        var initialCapMode    = toInitialNumber(prefs && prefs.capMode, 0, 0);
        var initialMode       = toInitialNumber(prefs && prefs.mode, 0, 0); /* 0:間隔→線分 / 1:線分→間隔 / 2:ランダム */
        var initialGapUnit    = toInitialUnit(prefs && prefs.gapPt, 5);
        var initialDashUnit   = toInitialUnit(prefs && prefs.dashPt, 0);
        var initialOffsetUnit = toInitialUnit(prefs && prefs.offsetPt, 0);
        var initialUseOffset  = toInitialFlag(prefs && prefs.useOffset, false);
        var initialAdjustEnds = toInitialFlag(prefs && prefs.adjustEnds, true);
        var initialReverse    = toInitialFlag(prefs && prefs.reversePath, false);

        // -----------------------------------------
        // ダイアログの構築 / Build dialog
        // -----------------------------------------

        var win = new Window("dialog", getLabel(LABELS.dialog.title) + " " + SCRIPT_VERSION);
        setupWindow(win);

        /* 選択中のパス情報（全幅）*/
        var panelPathInfo = addPanel(win, getLabel(LABELS.panel.pathInfo));
        var pathInfoText = getLabel(LABELS.fieldLabel.pathLength) + " " +
            ptToUnit(primaryPathLength, strokeUnit).toFixed(3) + " " + strokeUnit.label;
        if (targetPaths.length > 1) pathInfoText += "  (" + targetPaths.length + ")";
        var lblPathInfo = panelPathInfo.add("statictext", undefined, pathInfoText);
        lblPathInfo.alignment = "center";
        lblPathInfo.helpTip = getLabel(LABELS.tooltip.pathInfo);

        /* 2カラム */
        var mainColumns = win.add("group");
        setupRow(mainColumns, "fill", COLUMN_SPACING);
        mainColumns.alignChildren = ["fill", "top"];

        var columnLeft = addColumn(mainColumns);
        var columnRight = addColumn(mainColumns);

        /* 破線の計算（左カラム）*/
        var panelDashCalc = addPanel(columnLeft, getLabel(LABELS.panel.dashCalc));
        var dashInputColumn = addColumn(panelDashCalc, ["left", "top"], "left");

        /* 分割数 */
        var segmentsRow = addRow(dashInputColumn);
        var lblSegments = segmentsRow.add("statictext", undefined, getLabel(LABELS.fieldLabel.segments));
        lblSegments.preferredSize.width = FIELD_LABEL_WIDTH;
        lblSegments.justify = "right";
        var txtSegments = segmentsRow.add("edittext", undefined, String(initialSegments));
        txtSegments.characters = NUMBER_FIELD_CHARS;
        segmentsRow.helpTip = txtSegments.helpTip = getLabel(LABELS.tooltip.segments);
        changeValueByArrowKey(txtSegments, true, 1);

        /* 間隔（入力欄と結果表示を重ねる）*/
        var gapRow = addRow(dashInputColumn);
        var lblGap = gapRow.add("statictext", undefined, getLabel(LABELS.fieldLabel.gap));
        lblGap.preferredSize.width = FIELD_LABEL_WIDTH;
        lblGap.justify = "right";

        var gapFieldStack = gapRow.add("group");
        gapFieldStack.orientation = "stack";

        var txtGap = gapFieldStack.add("edittext", undefined, formatFieldNumber(initialGapUnit));
        txtGap.characters = NUMBER_FIELD_CHARS;
        changeValueByArrowKey(txtGap, false, 0);

        var lblGapResult = gapFieldStack.add("statictext", undefined, "");
        lblGapResult.justify = "right";
        lblGapResult.preferredSize.width = txtGap.preferredSize.width;
        lblGapResult.visible = false;

        gapRow.add("statictext", undefined, strokeUnit.label);
        gapRow.helpTip = txtGap.helpTip = getLabel(LABELS.tooltip.gap);

        /* 線分（入力欄と結果表示を重ねる）*/
        var dashRow = addRow(dashInputColumn);
        var lblDash = dashRow.add("statictext", undefined, getLabel(LABELS.fieldLabel.dash));
        lblDash.preferredSize.width = FIELD_LABEL_WIDTH;
        lblDash.justify = "right";

        var dashFieldStack = dashRow.add("group");
        dashFieldStack.orientation = "stack";

        var txtDash = dashFieldStack.add("edittext", undefined, formatFieldNumber(initialDashUnit));
        txtDash.characters = NUMBER_FIELD_CHARS;
        changeValueByArrowKey(txtDash, false, 0);

        var lblDashResult = dashFieldStack.add("statictext", undefined, "");
        lblDashResult.justify = "right";
        lblDashResult.preferredSize.width = txtDash.preferredSize.width;

        dashRow.add("statictext", undefined, strokeUnit.label);
        dashRow.helpTip = txtDash.helpTip = getLabel(LABELS.tooltip.dash);

        /* 計算方法（左カラム）*/
        var panelCalcMethod = addPanel(columnLeft, getLabel(LABELS.panel.calcMethod));
        var calcModeColumn = addColumn(panelCalcMethod, ["left", "top"], "left");

        var rbModeGapToDash = calcModeColumn.add("radiobutton", undefined, getLabel(LABELS.radio.gapToDash));
        var rbModeDashToGap = calcModeColumn.add("radiobutton", undefined, getLabel(LABELS.radio.dashToGap));
        var rbModeRandom = calcModeColumn.add("radiobutton", undefined, getLabel(LABELS.radio.random));
        rbModeGapToDash.helpTip = getLabel(LABELS.tooltip.gapToDash);
        rbModeDashToGap.helpTip = getLabel(LABELS.tooltip.dashToGap);
        rbModeRandom.helpTip = getLabel(LABELS.tooltip.random);
        rbModeGapToDash.value = (initialMode === 0);
        rbModeDashToGap.value = (initialMode === 1);
        rbModeRandom.value = (initialMode === 2);

        /* 開始位置（右カラム）*/
        var panelOffset = addPanel(columnRight, getLabel(LABELS.panel.offset), OFFSET_PANEL_SPACING);

        var offsetRow = addRow(panelOffset);
        var chkUseOffset = offsetRow.add("checkbox", undefined, "");
        chkUseOffset.value = initialUseOffset;
        chkUseOffset.preferredSize.width = LABELLESS_CHECKBOX_WIDTH;

        var txtOffset = offsetRow.add("edittext", undefined, formatFieldNumber(initialOffsetUnit));
        txtOffset.characters = NUMBER_FIELD_CHARS;
        txtOffset.enabled = chkUseOffset.value;
        changeValueByArrowKey(txtOffset, false, 0);

        var lblOffsetUnit = offsetRow.add("statictext", undefined, strokeUnit.label);
        lblOffsetUnit.enabled = chkUseOffset.value;
        offsetRow.helpTip = chkUseOffset.helpTip = txtOffset.helpTip = getLabel(LABELS.tooltip.useOffset);

        /* 1周期（線分＋間隔）を基準にしたプリセット */
        var offsetPresetRow = addRow(panelOffset);
        offsetPresetRow.enabled = chkUseOffset.value;
        offsetPresetRow.helpTip = getLabel(LABELS.tooltip.offsetPreset);
        var rbOffsetQuarter = offsetPresetRow.add("radiobutton", undefined, "1/4");
        var rbOffsetHalf = offsetPresetRow.add("radiobutton", undefined, "1/2");
        var rbOffsetThreeQuarter = offsetPresetRow.add("radiobutton", undefined, "3/4");

        /* 部分表示（右カラム）*/
        var panelPartialDisplay = addPanel(columnRight, getLabel(LABELS.panel.partial));
        var chkPartialDisplay = panelPartialDisplay.add("checkbox", undefined, getLabel(LABELS.checkbox.partialDisplay));
        chkPartialDisplay.alignment = "left";
        chkPartialDisplay.value = false;
        chkPartialDisplay.helpTip = getLabel(LABELS.tooltip.partialDisplay);

        /* 線端（右カラム）*/
        var panelCap = addPanel(columnRight, getLabel(LABELS.panel.cap));
        var capRow = addRow(panelCap);
        capRow.helpTip = getLabel(LABELS.tooltip.cap);
        var rbCapButt = capRow.add("radiobutton", undefined, getLabel(LABELS.radio.capButt));
        var rbCapRound = capRow.add("radiobutton", undefined, getLabel(LABELS.radio.capRound));
        var rbCapProject = capRow.add("radiobutton", undefined, getLabel(LABELS.radio.capProject));
        if (initialCapMode === 1) rbCapRound.value = true;
        else if (initialCapMode === 2) rbCapProject.value = true;
        else rbCapButt.value = true;

        /* 両端を調整・パスの方向反転（中央）*/
        var pathOptionRow = addRow(win, "center", OPTION_ROW_SPACING);

        var chkAdjustEnds = pathOptionRow.add("checkbox", undefined, getLabel(LABELS.checkbox.adjustEnds));
        chkAdjustEnds.value = initialAdjustEnds;
        chkAdjustEnds.helpTip = getLabel(LABELS.tooltip.adjustEnds);

        var chkReversePath = pathOptionRow.add("checkbox", undefined, getLabel(LABELS.checkbox.reversePath));
        chkReversePath.value = initialReverse;
        chkReversePath.helpTip = getLabel(LABELS.tooltip.reversePath);

        /* ボタン（左：破線クリア／右：キャンセル・OK）*/
        var buttonRow = win.add("group");
        setupRow(buttonRow, "fill");
        buttonRow.alignChildren = ["fill", "center"];
        buttonRow.margins = BUTTON_ROW_MARGINS;

        var clearButtonGroup = addRow(buttonRow, "left");
        var btnClearDash = clearButtonGroup.add("button", undefined, getLabel(LABELS.button.clearDash));
        btnClearDash.alignment = "left";
        btnClearDash.helpTip = getLabel(LABELS.tooltip.clearDash);

        var buttonSpacer = buttonRow.add("group");
        buttonSpacer.alignment = ["fill", "fill"];
        buttonSpacer.minimumSize.width = 0;

        var commitButtonGroup = addRow(buttonRow, "right");
        var btnCancel = commitButtonGroup.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });
        var btnOK = commitButtonGroup.add("button", undefined, getLabel(LABELS.button.ok), { name: "ok" });

        // -----------------------------------------
        // 表示用の書式 / Display formatting
        // -----------------------------------------

        /**
         * pt 値を結果表示用の文字列にする
         * @param {number} ptValue - pt 値
         * @returns {string} 小数第3位までの文字列
         */
        function ptToResultText(ptValue) {
            return ptToUnit(ptValue, strokeUnit).toFixed(3);
        }

        /**
         * pt 値を入力欄用の文字列にする
         * @param {number} ptValue - pt 値
         * @returns {string} 整形した文字列
         */
        function ptToFieldText(ptValue) {
            return formatFieldNumber(ptToUnit(ptValue, strokeUnit));
        }

        // -----------------------------------------
        // ランダムパターン / Random pattern
        // -----------------------------------------

        /**
         * 指定範囲の乱数（現在の線の単位）を返す
         * @param {number} minValue - 最小値
         * @param {number} maxValue - 最大値
         * @returns {number} 生成した長さ（0以上）
         */
        function getRandomLengthUnit(minValue, maxValue) {
            var value = Math.random() * (maxValue - minValue) + minValue;
            if (RANDOM_ROUND_VALUES) value = Math.round(value);
            /* Illustrator は NaN や負値を受け付けない */
            if (isNaN(value) || value < 0) value = 0;
            return value;
        }

        /**
         * ランダムパターンを生成し直す
         * 線端が「なし」のときは、線分が消えないように単位ごとの最小値を使う。
         * @returns {void}
         */
        function recalcRandomPattern() {
            var dashMinUnit = rbCapButt.value
                ? (RANDOM_DASH_MIN_BY_UNIT[strokeUnit.code] || RANDOM_DASH_MIN_DEFAULT)
                : 0;

            randomDashesPt = [];
            for (var k = 0; k < RANDOM_PATTERN_LENGTH; k += 2) {
                randomDashesPt.push(unitToPt(getRandomLengthUnit(dashMinUnit, RANDOM_DASH_MAX), strokeUnit));
                randomDashesPt.push(unitToPt(getRandomLengthUnit(RANDOM_GAP_MIN, RANDOM_GAP_MAX), strokeUnit));
            }
        }

        /**
         * ランダムパターンを用意する（前回値があれば復元）
         * @returns {void}
         */
        function ensureRandomPattern() {
            if (randomDashesPt && randomDashesPt.length === RANDOM_PATTERN_LENGTH) return;

            var savedDashes = prefs ? prefs.randPt : null;
            if (savedDashes && savedDashes.length >= RANDOM_PATTERN_LENGTH) {
                randomDashesPt = savedDashes.slice(0, RANDOM_PATTERN_LENGTH);
                return;
            }
            /* v1.7以前の4要素は、最後の1組をコピーして6要素へ拡張 */
            if (savedDashes && savedDashes.length >= 4) {
                randomDashesPt = [savedDashes[0], savedDashes[1], savedDashes[2], savedDashes[3], savedDashes[0], savedDashes[1]];
                return;
            }
            recalcRandomPattern();
        }

        /**
         * 入力された間隔をランダムパターンの全ギャップへ反映する
         * @returns {void}
         */
        function applyGapToRandomPattern() {
            var gapUnit = parseFloat(txtGap.text);
            if (isNaN(gapUnit) || gapUnit < 0) gapUnit = 0;
            var gapPt = unitToPt(gapUnit, strokeUnit);

            ensureRandomPattern();
            for (var k = 1; k < RANDOM_PATTERN_LENGTH; k += 2) {
                randomDashesPt[k] = gapPt;
            }
        }

        /**
         * 新しいランダムパターンを作り、入力された間隔を反映して返す
         * @returns {Array<number>} 破線パターン（pt）
         */
        function nextRandomPattern() {
            recalcRandomPattern();
            applyGapToRandomPattern();
            return randomDashesPt.slice(0);
        }

        /**
         * ランダムパターンの線分を表示用テキストにする
         * @returns {string} 「線分1 / 線分2 / 線分3」形式の文字列
         */
        function getRandomDashText() {
            ensureRandomPattern();
            var dashTexts = [];
            for (var k = 0; k < RANDOM_PATTERN_LENGTH; k += 2) {
                dashTexts.push(ptToResultText(randomDashesPt[k]));
            }
            return dashTexts.join(" / ");
        }

        // -----------------------------------------
        // UIの状態 / UI state
        // -----------------------------------------

        /**
         * 選択中の計算方法を返す
         * @returns {number} 0:間隔→線分 / 1:線分→間隔 / 2:ランダム
         */
        function getCalcModeIndex() {
            if (rbModeRandom.value) return 2;
            return rbModeDashToGap.value ? 1 : 0;
        }

        /**
         * 選択中の線端を返す
         * @returns {number} 0:なし / 1:丸型 / 2:突出
         */
        function getCapModeIndex() {
            if (rbCapRound.value) return 1;
            return rbCapProject.value ? 2 : 0;
        }

        /**
         * 計算方法に応じて入力欄・結果表示の切り替えを行う
         * @returns {void}
         */
        function updateModeUI() {
            var isRandom = rbModeRandom.value;
            var isDashToGap = rbModeDashToGap.value;

            /* ランダムは分割数・線分を使わないため、間隔だけ入力できるようにする */
            txtDash.visible = isDashToGap && !isRandom;
            lblDashResult.visible = !txtDash.visible;
            txtGap.visible = !isDashToGap || isRandom;
            lblGapResult.visible = !txtGap.visible;

            segmentsRow.enabled = !isRandom;
            dashRow.enabled = !isRandom;
            chkPartialDisplay.enabled = !isRandom;
            chkAdjustEnds.enabled = !isRandom && !primaryPath.closed;

            if (isRandom) {
                /* 部分表示はランダムパターンと併用できない */
                chkPartialDisplay.value = false;
                applyGapToRandomPattern();
                lblDashResult.text = getRandomDashText();
                txtGap.active = true;
            }
        }

        /**
         * 開始位置（オフセット）の入力可否を切り替える
         * @returns {void}
         */
        function updateOffsetUI() {
            txtOffset.enabled = chkUseOffset.value;
            lblOffsetUnit.enabled = chkUseOffset.value;
            offsetPresetRow.enabled = chkUseOffset.value;
        }

        // -----------------------------------------
        // 入力値の取得 / Read input values
        // -----------------------------------------

        /**
         * 入力された分割数を返す
         * @returns {number} 分割数（不正な場合は null）
         */
        function getSegmentsFromUI() {
            var segments = parseInt(txtSegments.text, 10);
            if (isNaN(segments) || segments <= 0) return null;
            return segments;
        }

        /**
         * 入力された開始位置（単位値）を返す
         * @returns {number} 開始位置。チェックOFFなら0、不正な場合は null
         */
        function getOffsetUnitFromUI() {
            if (!chkUseOffset.value) return 0;
            var offsetUnit = parseFloat(txtOffset.text);
            if (isNaN(offsetUnit) || offsetUnit < 0) return null;
            return offsetUnit;
        }

        /**
         * 現在の計算方法で使う入力欄の値が読めるかどうかを返す
         * @returns {boolean} 読める場合は true
         */
        function hasValidDashGapInput() {
            if (chkPartialDisplay.value) return true;
            var enteredValue = parseFloat(rbModeDashToGap.value ? txtDash.text : txtGap.text);
            return !isNaN(enteredValue) && enteredValue >= 0;
        }

        /**
         * 入力された線分長（pt）を返す
         * オープンパスで「両端を調整」ONかつ分割数1のときは、パス全長を線分長とする。
         * @param {number} pathLen - 対象パスの長さ（pt）
         * @param {boolean} isClosed - クローズパスかどうか
         * @param {number} segments - 分割数
         * @returns {number} 線分長（pt）。不正な場合は null
         */
        function getDashPtForPath(pathLen, isClosed, segments) {
            var dashUnit = parseFloat(txtDash.text);
            if (isNaN(dashUnit) || dashUnit < 0) return null;
            if (!isClosed && chkAdjustEnds.value && segments === 1) return pathLen;
            return unitToPt(dashUnit, strokeUnit);
        }

        /**
         * 1つのパスに適用する破線パターンを計算する
         * @param {number} pathLen - パスの長さ（pt）
         * @param {boolean} isClosed - クローズパスかどうか
         * @param {number} segments - 分割数
         * @returns {Object} { dashPt:number, gapPt:number, dashesPt:Array<number> }。計算できない場合は null
         */
        function calcDashPatternForPath(pathLen, isClosed, segments) {
            var adjustEnds = chkAdjustEnds.value;

            /* 部分表示：間隔0で計算し、線分1本だけを見せて残りは長い間隔で隠す */
            if (chkPartialDisplay.value) {
                var partialResult = calcDashAndCyclePt(segments, 0, pathLen, isClosed, adjustEnds);
                if (!partialResult || partialResult.dashPt <= 0) return null;
                var partialDashPt = partialResult.dashPt;
                return {
                    dashPt: partialDashPt,
                    gapPt: 0,
                    dashesPt: [0, 0, partialDashPt, pathLen + partialDashPt / 2]
                };
            }

            if (rbModeDashToGap.value) {
                var enteredDashPt = getDashPtForPath(pathLen, isClosed, segments);
                if (enteredDashPt == null) return null;
                var solvedGapPt = calcGapPtFromDashPt(segments, enteredDashPt, pathLen, isClosed, adjustEnds);
                if (solvedGapPt == null || solvedGapPt < 0) return null;
                return { dashPt: enteredDashPt, gapPt: solvedGapPt, dashesPt: [enteredDashPt, solvedGapPt] };
            }

            var gapUnit = parseFloat(txtGap.text);
            if (isNaN(gapUnit) || gapUnit < 0) return null;
            var enteredGapPt = unitToPt(gapUnit, strokeUnit);
            var calcResult = calcDashAndCyclePt(segments, enteredGapPt, pathLen, isClosed, adjustEnds);
            if (!calcResult || calcResult.dashPt <= 0) return null;
            return { dashPt: calcResult.dashPt, gapPt: enteredGapPt, dashesPt: [calcResult.dashPt, enteredGapPt] };
        }

        /**
         * 1周期（線分＋間隔）の長さを単位値で返す
         * @returns {number} 1周期の長さ。計算できない場合は null
         */
        function getDashCycleUnit() {
            var segments = getSegmentsFromUI();
            if (segments == null) return null;

            if (rbModeRandom.value) {
                applyGapToRandomPattern();
                var totalPt = 0;
                for (var k = 0; k < RANDOM_PATTERN_LENGTH; k++) totalPt += randomDashesPt[k];
                return ptToUnit(totalPt, strokeUnit);
            }

            var pattern = calcDashPatternForPath(primaryPathLength, primaryPath.closed, segments);
            if (!pattern) return null;
            return ptToUnit(pattern.dashPt + pattern.gapPt, strokeUnit);
        }

        // -----------------------------------------
        // 開始位置のプリセット / Offset presets
        // -----------------------------------------

        /**
         * プリセット（1/4・1/2・3/4）の選択状態を入力値に合わせる
         * @param {number} offsetUnit - 現在の開始位置（単位値）
         * @returns {void}
         */
        function syncOffsetPreset(offsetUnit) {
            rbOffsetQuarter.value = rbOffsetHalf.value = rbOffsetThreeQuarter.value = false;

            var cycleUnit = getDashCycleUnit();
            if (cycleUnit == null) return;

            /* 単位換算で誤差が出るため、判定はゆるめにする */
            var tolerance = Math.max(0.001, Math.abs(cycleUnit) * 0.0005);

            if (Math.abs(offsetUnit - cycleUnit * 0.25) <= tolerance) rbOffsetQuarter.value = true;
            else if (Math.abs(offsetUnit - cycleUnit * 0.50) <= tolerance) rbOffsetHalf.value = true;
            else if (Math.abs(offsetUnit - cycleUnit * 0.75) <= tolerance) rbOffsetThreeQuarter.value = true;
        }

        /**
         * 1周期を基準にした開始位置プリセットを適用する
         * @param {number} fraction - 1周期に対する比率（0.25 / 0.5 / 0.75）
         * @returns {void}
         */
        function applyOffsetPreset(fraction) {
            var cycleUnit = getDashCycleUnit();
            if (cycleUnit == null) return;
            var offsetUnit = cycleUnit * fraction;
            txtOffset.text = formatFieldNumber(offsetUnit < 0 ? 0 : offsetUnit);
            updatePreviewFromInput();
        }

        // -----------------------------------------
        // パスへの適用 / Apply to paths
        // -----------------------------------------

        /**
         * 対象パスすべてに処理を行う（個別の失敗は無視する）
         * @param {function} callback - 各パスに対して実行する処理
         * @returns {void}
         */
        function forEachTargetPath(callback) {
            for (var k = 0; k < targetPaths.length; k++) {
                try {
                    callback(targetPaths[k]);
                } catch (e) { }
            }
        }

        /**
         * 選択中の線端をパスへ設定する
         * @param {PathItem} pathItem - 対象のパス
         * @returns {void}
         */
        function applySelectedStrokeCap(pathItem) {
            if (rbCapRound.value) pathItem.strokeCap = StrokeCap.ROUNDENDCAP;
            else if (rbCapProject.value) pathItem.strokeCap = StrokeCap.PROJECTINGENDCAP;
            else pathItem.strokeCap = StrokeCap.BUTTENDCAP;
        }

        /**
         * パスへ破線設定を適用する
         * @param {PathItem} pathItem - 対象のパス
         * @param {Array<number>} dashesPt - 破線パターン（pt）
         * @param {number} offsetPt - 開始位置（pt）
         * @returns {void}
         */
        function applyStrokeDashes(pathItem, dashesPt, offsetPt) {
            pathItem.stroked = true;
            applySelectedStrokeCap(pathItem);
            pathItem.strokeDashOffset = offsetPt;
            pathItem.strokeDashes = dashesPt;
        }

        /**
         * 計算した破線を対象パスへ適用する（パスごとに長さを見て計算する）
         * @param {number} segments - 分割数
         * @param {number} offsetPt - 開始位置（pt）
         * @returns {void}
         */
        function applyDashesToPaths(segments, offsetPt) {
            forEachTargetPath(function (pathItem) {
                var pattern = calcDashPatternForPath(pathItem.length, pathItem.closed, segments);
                if (pattern) applyStrokeDashes(pathItem, pattern.dashesPt, offsetPt);
            });
        }

        /**
         * ランダムな破線を対象パスへ適用する（パスごとに別の乱数を使う）
         * @param {number} offsetPt - 開始位置（pt）
         * @returns {void}
         */
        function applyRandomDashesToPaths(offsetPt) {
            /* 結果表示用に代表のパターンを1回作る */
            nextRandomPattern();
            lblDashResult.text = getRandomDashText();

            forEachTargetPath(function (pathItem) {
                applyStrokeDashes(pathItem, nextRandomPattern(), offsetPt);
            });
        }

        /**
         * 対象パスの破線設定を解除する
         * @returns {void}
         */
        function clearDashesOnPaths() {
            forEachTargetPath(function (pathItem) {
                applyStrokeDashes(pathItem, [], 0);
            });
            app.redraw();
        }

        /**
         * パスの方向反転を現在の指定に合わせる
         * @param {boolean} shouldReverse - 反転させるかどうか
         * @returns {void}
         */
        function setReversePath(shouldReverse) {
            shouldReverse = !!shouldReverse;
            if (shouldReverse === directionReversed) return;

            try {
                /* パスの方向反転は選択に対して実行されるため、対象を選択してから実行する */
                doc.selection = targetPaths;
                app.executeMenuCommand('Reverse Path Direction');
                directionReversed = shouldReverse;
            } catch (e) {
                /* 失敗した場合はフラグを変更しない */
            }
        }

        /**
         * ダイアログを開く前の状態へ戻す
         * @returns {void}
         */
        function restoreOriginalState() {
            if (directionReversed) setReversePath(false);

            for (var k = 0; k < originalStates.length; k++) {
                var savedState = originalStates[k];
                try {
                    savedState.item.strokeDashes = savedState.strokeDashes.slice(0);
                    savedState.item.strokeDashOffset = savedState.strokeDashOffset;
                    savedState.item.strokeCap = savedState.strokeCap;
                    savedState.item.stroked = savedState.stroked;
                } catch (e) { }
            }
            app.redraw();
        }

        // -----------------------------------------
        // プレビュー / Preview
        // -----------------------------------------

        /**
         * 線分・間隔の結果表示を更新し、隠れている入力欄も同期する
         * @param {number} segments - 分割数
         * @returns {boolean} 計算できた場合は true
         */
        function refreshDashGapDisplay(segments) {
            var resultLabel = rbModeDashToGap.value ? lblGapResult : lblDashResult;

            /* 入力途中で数値として読めないときは、結果表示を空にする */
            if (!hasValidDashGapInput()) {
                resultLabel.text = "";
                return false;
            }

            var pattern = calcDashPatternForPath(primaryPathLength, primaryPath.closed, segments);
            if (!pattern) {
                resultLabel.text = getLabel(LABELS.alert.calcError);
                return false;
            }

            lblDashResult.text = ptToResultText(pattern.dashPt);
            lblGapResult.text = ptToResultText(pattern.gapPt);

            /* 表示を切り替えたときにずれないよう、隠れている入力欄も同期する */
            if (chkPartialDisplay.value) {
                txtGap.text = "0";
                txtDash.text = ptToFieldText(pattern.dashPt);
            } else if (rbModeDashToGap.value) {
                txtGap.text = ptToFieldText(pattern.gapPt);
                /* 全長ダッシュに置き換わった場合だけ入力欄を合わせる（入力中の値は書き換えない）*/
                if (!primaryPath.closed && chkAdjustEnds.value && segments === 1) {
                    txtDash.text = ptToFieldText(pattern.dashPt);
                }
            } else {
                txtDash.text = ptToFieldText(pattern.dashPt);
            }
            return true;
        }

        /**
         * 現在の設定でプレビューを更新する（アラートは出さない）
         * @returns {void}
         */
        function updatePreview() {
            if (isDashCleared) {
                lblDashResult.text = "";
                lblGapResult.text = "";
                clearDashesOnPaths();
                return;
            }

            var offsetUnit = getOffsetUnitFromUI();
            var segments = getSegmentsFromUI();
            if (offsetUnit == null || segments == null) {
                lblDashResult.text = "";
                lblGapResult.text = "";
                return;
            }

            var offsetPt = unitToPt(offsetUnit, strokeUnit);

            if (rbModeRandom.value) {
                applyRandomDashesToPaths(offsetPt);
                app.redraw();
                return;
            }

            if (!refreshDashGapDisplay(segments)) return;

            /* プリセット表示を手入力・分割数の変更に追従させる */
            syncOffsetPreset(offsetUnit);
            applyDashesToPaths(segments, offsetPt);
            app.redraw();
        }

        /**
         * 入力変更時にプレビューを更新する（破線クリア状態は解除する）
         * @returns {void}
         */
        function updatePreviewFromInput() {
            isDashCleared = false;
            updatePreview();
        }

        // -----------------------------------------
        // 確定処理 / Commit
        // -----------------------------------------

        /**
         * OK時の入力チェックを行い、保存用の線分・間隔（pt）を返す
         * @param {number} segments - 分割数
         * @returns {Object} { dashPt:number, gapPt:number }。エラー時は null（アラート表示済み）
         */
        function validateBeforeApply(segments) {
            if (!hasValidDashGapInput()) {
                alert(getLabel(rbModeDashToGap.value ? LABELS.alert.dashInvalid : LABELS.alert.gapInvalid));
                return null;
            }

            /* ランダムは分割数を使わないため、間隔が読めれば確定できる */
            if (rbModeRandom.value) {
                return {
                    dashPt: unitToPt(parseFloat(txtDash.text) || 0, strokeUnit),
                    gapPt: unitToPt(parseFloat(txtGap.text), strokeUnit)
                };
            }

            var pattern = calcDashPatternForPath(primaryPathLength, primaryPath.closed, segments);
            if (pattern) return pattern;

            /* 計算できないときは、設定できる最大値を知らせる */
            if (chkPartialDisplay.value) {
                alert(getLabel(LABELS.alert.calcError));
            } else if (rbModeDashToGap.value) {
                alert(formatLabel(LABELS.alert.dashTooLong, [ptToResultText(primaryPathLength / segments), strokeUnit.label]));
            } else {
                var maxGapPt;
                if (primaryPath.closed || !chkAdjustEnds.value) maxGapPt = primaryPathLength / segments;
                else maxGapPt = (segments <= 1) ? primaryPathLength : primaryPathLength / (segments - 1);
                alert(formatLabel(LABELS.alert.gapTooLong, [ptToResultText(maxGapPt), strokeUnit.label]));
                lblDashResult.text = getLabel(LABELS.alert.calcError);
            }
            return null;
        }

        /**
         * 現在のUI状態を前回値として保存する
         * @param {number} segments - 分割数
         * @param {number} dashPt - 線分長（pt）
         * @param {number} gapPt - 間隔（pt）
         * @param {number} offsetPt - 開始位置（pt）
         * @returns {void}
         */
        function saveCurrentPrefs(segments, dashPt, gapPt, offsetPt) {
            savePrefs({
                segments: segments,
                gapPt: gapPt,
                dashPt: dashPt,
                offsetPt: offsetPt,
                capMode: getCapModeIndex(),
                mode: getCalcModeIndex(),
                reversePath: chkReversePath.value,
                adjustEnds: chkAdjustEnds.value,
                randPt: (rbModeRandom.value && randomDashesPt) ? randomDashesPt.slice(0) : null,
                useOffset: chkUseOffset.value
            });
        }

        // -----------------------------------------
        // イベント / Event handlers
        // -----------------------------------------

        txtSegments._onArrowChange = updatePreviewFromInput;
        txtGap._onArrowChange = updatePreviewFromInput;
        txtDash._onArrowChange = updatePreviewFromInput;
        txtOffset._onArrowChange = updatePreviewFromInput;

        txtSegments.onChanging = txtSegments.onChange = updatePreviewFromInput;
        txtGap.onChanging = txtGap.onChange = updatePreviewFromInput;
        txtDash.onChanging = txtDash.onChange = updatePreviewFromInput;
        txtOffset.onChanging = txtOffset.onChange = updatePreviewFromInput;

        rbCapButt.onClick = rbCapRound.onClick = rbCapProject.onClick = updatePreview;

        rbModeGapToDash.onClick = rbModeDashToGap.onClick = function () {
            updateModeUI();
            updatePreviewFromInput();
        };

        rbModeRandom.onClick = function () {
            /* クリックのたびにパターンを作り直す */
            recalcRandomPattern();
            updateModeUI();
            updatePreviewFromInput();
        };

        rbOffsetQuarter.onClick = function () { applyOffsetPreset(0.25); };
        rbOffsetHalf.onClick = function () { applyOffsetPreset(0.50); };
        rbOffsetThreeQuarter.onClick = function () { applyOffsetPreset(0.75); };

        chkUseOffset.onClick = function () {
            updateOffsetUI();
            updatePreviewFromInput();
        };

        chkPartialDisplay.onClick = function () {
            /* ONにしたら間隔を0にする */
            if (chkPartialDisplay.value) txtGap.text = "0";
            updatePreviewFromInput();
        };

        chkAdjustEnds.onClick = updatePreviewFromInput;

        chkReversePath.onClick = function () {
            setReversePath(chkReversePath.value);
            updatePreview();
        };

        btnClearDash.onClick = function () {
            isDashCleared = true;
            updatePreview();
        };

        /* キャンセル：ダイアログを開く前の状態に戻して閉じる */
        btnCancel.onClick = function () {
            restoreOriginalState();
            win.close(0);
        };

        /* ×ボタンやEscで閉じた場合も、OK以外は復元する */
        win.onClose = function () {
            if (!closedByOK) restoreOriginalState();
            return true;
        };

        btnOK.onClick = function () {
            var segments = getSegmentsFromUI();

            /* 破線クリアを確定（他のUI状態は残して保存する）*/
            if (isDashCleared) {
                clearDashesOnPaths();
                saveCurrentPrefs(
                    (segments == null) ? initialSegments : segments,
                    unitToPt(initialDashUnit, strokeUnit),
                    unitToPt(initialGapUnit, strokeUnit),
                    0
                );
                closedByOK = true;
                win.close(1);
                return;
            }

            if (segments == null) {
                alert(getLabel(LABELS.alert.segmentsInvalid));
                return;
            }
            var offsetUnit = getOffsetUnitFromUI();
            if (offsetUnit == null) {
                alert(getLabel(LABELS.alert.offsetInvalid));
                return;
            }

            var applied = validateBeforeApply(segments);
            if (!applied) return;

            /* プレビューと同じ処理で対象のパスへ適用する */
            updatePreview();
            saveCurrentPrefs(segments, applied.dashPt, applied.gapPt, unitToPt(offsetUnit, strokeUnit));

            closedByOK = true;
            win.close(1);
        };

        // -----------------------------------------
        // 初期化と表示 / Initialize & show
        // -----------------------------------------

        /* 前回値（パス方向）を反映 */
        setReversePath(chkReversePath.value);

        updateOffsetUI();
        updateModeUI();
        updatePreview();

        /* ランダム以外は分割数の入力欄をアクティブにして表示する */
        if (!rbModeRandom.value) txtSegments.active = true;
        win.onShow = function () {
            if (rbModeRandom.value) txtGap.active = true;
            else txtSegments.active = true;
        };

        win.show();
    }

    /**
     * ドキュメントと選択を確認し、ダイアログを開く
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel(LABELS.alert.noDocument));
            return;
        }

        var doc = app.activeDocument;
        if (doc.selection.length === 0) {
            alert(getLabel(LABELS.alert.noSelection));
            return;
        }

        var targetPaths = collectTargetPaths(doc.selection);
        if (targetPaths.length === 0) {
            alert(getLabel(LABELS.alert.noPathSelected));
            return;
        }

        resetBoundsAndToggleEdges();
        try {
            showDashDialog(doc, targetPaths);
        } finally {
            resetBoundsAndToggleEdges();
        }
    }

    main();

})();
