#target illustrator
#targetengine "FontCatalogGeneratorEngine"
app.preferences.setBooleanPreference("ShowExternalJSXWarning", false);

/*

### 概要

システムにインストールされているフォントを一覧化し、アートボード上にフォント見本を自動生成します。
表示する文字列とフォントサイズはダイアログで指定できます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/FontCatalogGenerator.md

note記事も参照してください。
https://note.com/studio_tofu/n/n7b0cf367ec88

### Overview

Lists the fonts installed on the system and generates a specimen sheet for them on the artboard.
The sample string and the font size are set in a dialog.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/FontCatalogGenerator.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "FontCatalogGenerator";         /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.7.3";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-01-27";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-29";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/FontCatalogGenerator.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/FontCatalogGenerator.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/studio_tofu/n/n7b0cf367ec88"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // ボタン行（再利用パーツ） / Button row (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内（ダイアログを作る関数より前）に貼る。
    //    識別子は BUTTON_ROW_* / addButtonRow
    // 2. ダイアログの最後で行を作り、ボタンは btn 接頭辞の変数で左右のグループに足す（キャンセル → OK の順）
    //      var buttonRow = addButtonRow(dialog);
    //      var btnPreferences = buttonRow.leftGroup.add("button", undefined, getLabel("button.preferences"));
    //      var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
    //      var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
    //    左右中央に並べるときは addButtonRow(dialog, { centered: true }) にして、buttonRow.rowGroup に直接足す
    // 3. 行の上の余白は BUTTON_ROW_TOP_MARGIN で決める。左右の余白はダイアログの margins に任せる
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

    var BUTTON_ROW_TOP_MARGIN = 5; /* ボタン行の上の余白 / top margin of the button row */
    var BUTTON_ROW_SPACING = 10;   /* ボタンどうしの間隔 / spacing between buttons */

    /**
     * ダイアログ下部のボタン行を作る。
     * 通常は「左のグループ・伸びるスペーサー・右のグループ」、centered なら行そのものを左右中央に置く
     * @param {Window|Group|Panel} parent - 行を足す先（ふつうはダイアログ）
     * @param {Object} [rowOptions] - { centered: true } で左右中央に並べる
     * @returns {{rowGroup: Group, leftGroup: Group|null, rightGroup: Group|null}} 行と左右のグループ（centered のときは左右が null）
     */
    function addButtonRow(parent, rowOptions) {
        var isCentered = !!(rowOptions && rowOptions.centered);
        var btnRowGroup = parent.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.margins = [0, BUTTON_ROW_TOP_MARGIN, 0, 0];
        btnRowGroup.spacing = BUTTON_ROW_SPACING;

        if (isCentered) {
            btnRowGroup.alignment = ["center", "bottom"];
            btnRowGroup.alignChildren = ["center", "center"];
            return { rowGroup: btnRowGroup, leftGroup: null, rightGroup: null };
        }

        btnRowGroup.alignment = ["fill", "bottom"];

        var btnLeftGroup = btnRowGroup.add("group");
        btnLeftGroup.alignChildren = ["left", "center"];
        btnLeftGroup.spacing = BUTTON_ROW_SPACING;

        /* 余りの幅を吸って、右のグループを右端に寄せる / Absorbs the extra width so the right group sits at the right edge */
        var spacer = btnRowGroup.add("group");
        spacer.alignment = ["fill", "fill"];
        spacer.minimumSize.width = 0;

        var btnRightGroup = btnRowGroup.add("group");
        btnRightGroup.alignChildren = ["right", "center"];
        btnRightGroup.spacing = BUTTON_ROW_SPACING;

        return { rowGroup: btnRowGroup, leftGroup: btnLeftGroup, rightGroup: btnRightGroup };
    }

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // ボタン行（再利用パーツ）ここまで / End of the reusable button row
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // ローカライズ（再利用パーツ） / Localization (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内のローカライズ節（LABELS の直前）に貼る。
    //    uiLang を使うコード（StepperButtons・LinkToggle の部品など）より前に置く
    // 2. 識別子は uiLang / getCurrentLang / getLabel / labelText / labelValueText / fillLabelPlaceholders。
    //    同じ役割の既存の関数・変数（getCurrentLanguage、currentLanguage、formatLabel など）は消して、これに寄せる
    // 3. 呼び出しはどちらの形でもよい（混ぜてもよい）
    //      getLabel("dialog.title")        … パス
    //      getLabel(LABELS.dialog.title)   … { ja, en } を直接
    //      getLabel("alert.count", { count: 3 })  … "{count} 個" の {count} を差し込む
    //      getLabel("alert.range", [1, 10])       … "%1〜%2" の %1・%2 を差し込む
    //      labelText("fieldLabel.width")   … 末尾にコロン（日本語は全角「：」、英語は半角「:」）
    //      labelValueText("message.count", 5) … 「件数：5」／「Count: 5」（値が続く1行。英語はコロンのあとに空白）
    // 4. 見つからないパスはパスの文字列をそのまま返す（表示で気づけるように）。{ ja, en } が無いときは空文字
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

    /**
     * UI の言語を返す（"ja" で始まるロケールは日本語、それ以外は英語）
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        return (String($.locale || "").indexOf("ja") === 0) ? "ja" : "en";
    }

    var uiLang = getCurrentLang();

    /**
     * LABELS から今の UI 言語の文言を取り出す。
     * @param {string|Object} labelRef - "dialog.title" のようなパス、または { ja, en }
     * @param {Object|Array} [placeholderValues] - { name: 値 } なら {name} を、[値, …] なら %1, %2 … を差し込む
     * @returns {string} 文言。パスが見つからなければパスの文字列、{ ja, en } が無ければ空文字
     */
    function getLabel(labelRef, placeholderValues) {
        var labelEntry = labelRef;
        if (typeof labelRef === "string") {
            var labelPathKeys = labelRef.split(".");
            labelEntry = LABELS;
            for (var i = 0; i < labelPathKeys.length && labelEntry != null; i++) {
                labelEntry = labelEntry[labelPathKeys[i]];
            }
        }
        var labelString;
        if (typeof labelEntry === "string") labelString = labelEntry;
        else if (labelEntry != null && labelEntry[uiLang] != null) labelString = labelEntry[uiLang];
        else if (labelEntry != null && labelEntry.en != null) labelString = labelEntry.en;
        else return (typeof labelRef === "string") ? labelRef : "";
        return fillLabelPlaceholders(String(labelString), placeholderValues);
    }

    /**
     * 項目名の文言の末尾にコロンを付ける（日本語は全角「：」、英語は半角「:」）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {Object|Array} [placeholderValues] - getLabel と同じ
     * @returns {string} コロン付きの文言
     */
    function labelText(labelRef, placeholderValues) {
        return getLabel(labelRef, placeholderValues) + (uiLang === "ja" ? "：" : ":");
    }

    /**
     * 「項目名：値」の1行を返す（日本語は「件数：5」、英語は「Count: 5」とコロンのあとに空白を入れる）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {string|number} value - コロンのあとに続ける値
     * @returns {string} 項目名と値をつないだ文字列
     */
    function labelValueText(labelRef, value) {
        return labelText(labelRef) + (uiLang === "ja" ? "" : " ") + value;
    }

    /**
     * 文言の {name} や %1 に値を差し込む
     * @param {string} labelString - 文言
     * @param {Object|Array} [placeholderValues] - { name: 値 } または [値, …]
     * @returns {string} 差し込んだ文言
     */
    function fillLabelPlaceholders(labelString, placeholderValues) {
        if (placeholderValues == null) return labelString;
        if (placeholderValues instanceof Array) {
            /* 大きい番号から置き換え、%1 が %10 の一部を置き換えないようにする / Replace from the highest index so %1 does not eat into %10 */
            for (var i = placeholderValues.length; i >= 1; i--) {
                labelString = labelString.split("%" + i).join(String(placeholderValues[i - 1]));
            }
            return labelString;
        }
        for (var placeholderKey in placeholderValues) {
            if (!placeholderValues.hasOwnProperty(placeholderKey)) continue;
            labelString = labelString.split("{" + placeholderKey + "}").join(String(placeholderValues[placeholderKey]));
        }
        return labelString;
    }

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // ローカライズ（再利用パーツ）ここまで / End of the reusable localization
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
      dialogTitle: {
        ja: "フォント見本作成",
        en: "FontCatalogGenerator"
      },
      labelSampleText: {
        ja: "表示する文字",
        en: "Sample text"
      },
      tipSampleText: {
        ja: "各フォントの見本として並べる文字です。",
        en: "The text shown as the specimen for each font."
      },
      tipFontSize: {
        ja: "見本の文字サイズです。",
        en: "Size of the specimen text."
      },
      labelFontSize: {
        ja: "フォントサイズ",
        en: "Font size"
      },
      panelSampleSettings: {
        ja: "見本の設定",
        en: "Sample settings"
      },
      panelFontNameDisplay: {
        ja: "フォント名の表示",
        en: "Font name display"
      },
      panelExcludeList: {
        ja: "除外リスト",
        en: "Exclude list"
      },
      panelExcludeAdd: {
        ja: "除外リストへの追加",
        en: "Add to exclude list"
      },
      checkboxExcludeItalic: {
        ja: "斜体（Italic/Oblique）",
        en: "Italic/Oblique"
      },
      checkboxExcludeKorean: {
        ja: "韓国語フォント",
        en: "Korean fonts"
      },
      checkboxExcludeChineseSC: {
        ja: "中国語フォント（簡体）",
        en: "Chinese fonts (Simplified)"
      },
      checkboxExcludeChineseTC: {
        ja: "中国語フォント（繁体）",
        en: "Chinese fonts (Traditional)"
      },
      checkboxExcludeHebrew: {
        ja: "ヘブライ語フォント",
        en: "Hebrew fonts"
      },
      checkboxExcludeThai: {
        ja: "タイ語フォント",
        en: "Thai fonts"
      },
      checkboxExcludeArabic: {
        ja: "アラビア系フォント",
        en: "Arabic fonts"
      },
      checkboxExcludeSystemFonts: {
        ja: "システムフォント",
        en: "System fonts"
      },
      checkboxExcludeVariableFonts: {
        ja: "バリアブル",
        en: "Variable fonts"
      },
      checkboxExcludeIllustratorBundled: {
        ja: "Illustrator付属",
        en: "Bundled with Illustrator"
      },
      checkboxExcludeCompositeFonts: {
        ja: "合成フォント",
        en: "Composite fonts"
      },
      checkboxExcludeMorisawa: {
        ja: "モリサワ",
        en: "Morisawa"
      },
      checkboxExcludeFontworks: {
        ja: "フォントワークス",
        en: "Fontworks"
      },
      checkboxAddSelectedFontToExclude: {
        ja: "選択フォントを除外",
        en: "Exclude selected font"
      },
      labelSelectedFontPrefix: {
        ja: "選択フォント: ",
        en: "Selected font: "
      },
      labelSelectedFontNone: {
        ja: "（なし）",
        en: "(none)"
      },
      buttonCancel: {
        ja: "キャンセル",
        en: "Cancel"
      },
      buttonOK: {
        ja: "OK",
        en: "OK"
      },
      progressTitle: {
        ja: "処理中...",
        en: "Processing..."
      },
      progressCancel: {
        ja: "キャンセル",
        en: "Cancel"
      },
      progressCancelled: {
        ja: "キャンセルしました",
        en: "Cancelled"
      },
      checkboxShowFontName: {
        ja: "フォント名（ファミリー＋スタイル）",
        en: "Font name (family + style)"
      },
      checkboxShowPostScriptName: {
        ja: "PostScript名（内部名）",
        en: "PostScript name (internal)"
      },
      tooltip: {
        stepUp: {
          ja: "値を増やす（shift＋クリックで10の倍数へ、option＋クリックで0.1ずつ）",
          en: "Increase (Shift-click to snap to 10s, Option-click by 0.1)"
        },
        stepDown: {
          ja: "値を減らす（shift＋クリックで10の倍数へ、option＋クリックで0.1ずつ）",
          en: "Decrease (Shift-click to snap to 10s, Option-click by 0.1)"
        },
        stepUpInteger: { ja: "値を増やす（shift＋クリックで10の倍数へ）", en: "Increase (Shift-click to snap to 10s)" },
        stepDownInteger: { ja: "値を減らす（shift＋クリックで10の倍数へ）", en: "Decrease (Shift-click to snap to 10s)" }
      }
    };

    /* --------------------------------------------------
     * 単位ユーティリティ（文字単位対応） / Unit utilities for text
     * -------------------------------------------------- */

    /**
     * 環境設定の文字単位のラベルを返す
     * @returns {string} 単位ラベル（取得できない場合は "pt"）
     */
    function getCurrentTextUnitLabel() {
      try {
        return getUnitInfo("text/units").label;
      } catch (e) {
        return "pt";
      }
    }

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

    /**
     * pt の値を、現在の文字単位の値へ換算する
     * @param {number} ptValue - pt の値
     * @returns {number} 現在の文字単位での値
     */
    function ptToCurrentTextUnit(ptValue) {
      return ptValue / getUnitInfo("text/units").pointsPerUnit;
    }

    /**
     * 現在の文字単位の値を pt へ換算する
     * @param {number} unitValue - 現在の文字単位での値
     * @returns {number} pt の値
     */
    function currentTextUnitToPt(unitValue) {
      return unitValue * getUnitInfo("text/units").pointsPerUnit;
    }

    /**
     * 表示用に数値を整形する（小数第3位まで、末尾の0は落とす）
     * @param {number} num - 整形する数値
     * @returns {string} 整形した文字列
     */
    function formatNumberForDisplay(num) {
      try {
        var n = Math.round(num * 1000) / 1000; // up to 3 decimals
        var s = String(n);
        if (s.indexOf(".") !== -1) {
          s = s.replace(/\.?0+$/, "");
        }
        return s;
      } catch (e) {
        return String(num);
      }
    }

    /* 部分一致（大文字小文字無視） / Case-insensitive contains */
    /**
     * 大文字小文字を区別せずに部分一致を判定する
     * @param {string} haystack - 検索対象の文字列
     * @param {string} needle - 探す文字列
     * @returns {boolean} 含まれていれば true
     */
    function containsIgnoreCase(haystack, needle) {
      try {
        if (!haystack || !needle) return false;
        return String(haystack).toLowerCase().indexOf(String(needle).toLowerCase()) !== -1;
      } catch (e) {
        return false;
      }
    }

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // UI の明暗（再利用パーツ） / UI theme (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内に貼る（StepperButtons・LinkToggle の部品より前）。識別子は isDarkUI
    // 2. 配色を明暗で切り替えるときは isDarkUI() を1回だけ呼んで定数に控える
    //      var MY_UI_DARK = isDarkUI();
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

    /**
     * UI がダークテーマかどうかを判定する（Illustrator は uiBrightness、InDesign は uiBrightnessPreference）
     * @returns {boolean} ダークなら true。取得できない環境では false（明るいUI扱い）
     */
    function isDarkUI() {
        try {
            if (app.preferences && app.preferences.getRealPreference) {
                return app.preferences.getRealPreference("uiBrightness") <= 0.5; /* Illustrator */
            }
            return app.generalPreferences.uiBrightnessPreference <= 0.5; /* InDesign */
        } catch (e) {
            return false;
        }
    }

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // UI の明暗（再利用パーツ）ここまで / End of the reusable UI theme
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // ステップボタン（再利用パーツ） / Stepper buttons (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内（ローカライズより前）に貼る。
    //    識別子はすべて STEPPER_* / *Stepper* / *Stepped* の名前なので、既存の名前とはぶつからない
    //    UI の明暗は UITheme 部品の isDarkUI() を使う（先に UITheme の ▼〜▲ も貼っておく）
    // 2. コピー先の LABELS.tooltip に stepUp / stepDown / stepUpInteger / stepDownInteger を足す（このファイルの LABELS から写す）。
    //    getLabel() と uiLang はコピー先のものをそのまま使う
    // 3. 数値欄を addSteppedField() で作る。項目名・∧∨・入力欄がひと組で入り、↑↓キーも∧∨と同じ処理で増減する
    //      var widthInput = addSteppedField(parentPanel, {
    //          label: labelText(LABELS.fieldLabel.width), labelWidth: 60,
    //          text: "210 mm", characters: 8, step: 1, min: 1, unit: " mm",
    //          onStep: function (numberInput) { updatePreview(); }
    //      });
    //    値の種類は options で切り分ける:
    //      小数あり（幅・位置など）   … 指定なし（option＋クリックで0.1ずつ）
    //      整数・1以上（段数・個数など）… integer: true, min: 1（0・小数・負数は受け付けず、option＋クリックも1ずつ）
    //      整数・0以上（間隔の数など）  … integer: true, min: 0
    //      範囲つき（％など）           … min: 0, max: 100, unit: "%"
    // 4. 有効／無効は setSteppedFieldEnabled(widthInput, isEnabled)（∧∨のディム表示も切り替わる）。
    //    行・パネルなど親の enabled を切り替えたときは、そのあとで redrawSteppersIn(親) を呼んで∧∨を描き直す
    //    （∧∨は親をたどって無効を判定し、無効の間はクリックも↑↓キーも効かない）
    // 5. 値は parseFloat(widthInput.text) で読む（unit 付きの欄は「210 mm」の形で入っている）
    // 6. この欄に別の↑↓キー処理を付けない（↑↓キーが二重に効く）
    // 既存の edittext をそのまま使うときは、同じ行の group（spacing 0）に addStepper() → edittext の順で置き、
    // bindSteppedArrowKeys(edittext, stepperGroup) を呼ぶ
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

    // -----------------------------------------
    // ステップボタンの寸法・増減量 / Stepper metrics and steps
    // -----------------------------------------
    var STEPPER_BUTTON_WIDTH   = 20;  /* ∧∨ボタンの幅 / button width */
    var STEPPER_BUTTON_HEIGHT  = 11;  /* ∧∨ボタン1つの高さ（2つ重ねた全体の高さは22） / button height (22 for the pair) */
    var STEPPER_CORNER_RADIUS  = 2;   /* 枠の角丸の半径（ScriptUIは円弧を描けないため短い線分で近似） / corner radius, approximated with segments */
    var STEPPER_FIELD_SPACING  = 3;   /* 項目名と∧∨の間隔 / spacing between the label and the stepper */
    var STEPPER_SIDE_MARGIN    = 3;   /* ∧∨の左に足す余白（右は入力欄に突き合わせる） / extra space left of the stepper */
    var STEPPER_SHIFT_MULTIPLE = 10;  /* shift＋クリックでそろえる倍数 / Shift-click snaps to multiples of this */
    var STEPPER_OPTION_STEP    = 0.1; /* option＋クリックの増減量 / Option-click step */

    // -----------------------------------------
    // ステップボタンの配色 / Stepper colors
    // -----------------------------------------
    var STEPPER_UI_DARK           = isDarkUI();
    /* UIの明るさは4段階あり、段階ごとに背景色が違う。どの段階でも背景に対する差で見せるよう、黒・白の半透明を重ねる。
       ダーク側は Illustrator 標準のスピナー（［グリッドに分割］）で実測、明るい側は最も明るい段階（背景 約0.94）から逆算
       UI brightness has four levels with different backgrounds, so colors are translucent overlays that follow the
       dialog background. Dark values are measured from Illustrator's own spinner; light values derived for the lightest level */
    var STEPPER_FILL_COLOR        = STEPPER_UI_DARK ? [0, 0, 0, 0.10]  : [1, 1, 1, 0.50];  /* 地 / background */
    var STEPPER_FRAME_COLOR       = STEPPER_UI_DARK ? [1, 1, 1, 0.07]  : [0, 0, 0, 0.10];  /* 枠線 / frame */
    var STEPPER_PRESSED_COLOR     = STEPPER_UI_DARK ? [1, 1, 1, 0.12]  : [0, 0, 0, 0.13];  /* 押下中 / pressed */
    var STEPPER_CHEVRON_COLOR     = STEPPER_UI_DARK ? [1, 1, 1, 1]     : [0, 0, 0, 0.70];  /* 山形の線 / chevron */
    var STEPPER_DIM_FILL_COLOR    = STEPPER_UI_DARK ? [1, 1, 1, 0.035] : [1, 1, 1, 0.30];  /* 無効時の地 / background when disabled */
    var STEPPER_DIM_FRAME_COLOR   = STEPPER_UI_DARK ? [1, 1, 1, 0.035] : [0, 0, 0, 0.05];  /* 無効時の枠線（ダークは地と同じで見せない） / frame when disabled */
    var STEPPER_DIM_CHEVRON_COLOR = STEPPER_UI_DARK ? [1, 1, 1, 0.20]  : [0, 0, 0, 0.25];  /* 無効時の山形 / chevron when disabled */

    // -----------------------------------------
    // 数値欄を作る（外から呼ぶ関数） / Public API
    // -----------------------------------------
    /**
     * 「項目名・∧∨・入力欄」をひと組にした数値欄を追加する。
     * ↑↓キーでも∧∨と同じように増減する。直接入力した値も、フォーカスが外れたときに
     * 整数化・下限・上限・単位（「20 mm」の形）へそろえ、数値でなければ直前の値に戻す
     * @param {Group|Panel} parent - 追加先
     * @param {Object} fieldOptions - label（コロン込みの項目名）/ labelWidth / text / characters /
     *     step / min / max / integer（true で整数のみ）/ unit / onStep
     * @returns {EditText} 入力欄（項目名は .fieldLabel、∧∨は .stepperGroup で参照できる）
     */
    function addSteppedField(parent, fieldOptions) {
        var fieldRowGroup = parent.add("group");
        fieldRowGroup.orientation = "row";
        fieldRowGroup.alignChildren = ["left", "center"];
        fieldRowGroup.spacing = STEPPER_FIELD_SPACING;

        var fieldLabel = fieldRowGroup.add("statictext", undefined, fieldOptions.label || "");
        if (fieldOptions.labelWidth) {
            fieldLabel.preferredSize.width = fieldOptions.labelWidth;
            fieldLabel.justify = "right";
        }

        /* ∧∨と入力欄は隙間0で突き合わせる / butt the stepper against the field */
        var stepperInputGroup = fieldRowGroup.add("group");
        stepperInputGroup.orientation = "row";
        stepperInputGroup.alignChildren = ["left", "center"];
        stepperInputGroup.spacing = 0;
        stepperInputGroup.margins = 0;

        var numberInput;
        var stepperGroup = addStepper(stepperInputGroup, function () { return numberInput; }, fieldOptions);
        numberInput = stepperInputGroup.add("edittext", undefined, fieldOptions.text || "");
        numberInput.characters = fieldOptions.characters || 6;
        numberInput.fieldLabel = fieldLabel;
        numberInput.stepperGroup = stepperGroup;

        /* ↑↓キーも∧∨と同じ処理で増減する（増減量・下限・上限・単位・修飾キーをそろえる） / arrow keys share the stepper's logic */
        bindSteppedArrowKeys(numberInput, stepperGroup);

        /* 直接入力をそろえる。数値でなければ直前の値に戻す / normalize typed values; revert non-numbers */
        numberInput.lastValidText = numberInput.text;
        numberInput.onChange = function () {
            var value = parseFloat(numberInput.text);
            if (isNaN(value)) {
                numberInput.text = numberInput.lastValidText;
                return;
            }
            writeSteppedValue(numberInput, value, fieldOptions);
        };
        return numberInput;
    }

    /**
     * 数値欄の有効／無効を、項目名・∧∨ごとまとめて切り替える
     * @param {EditText} numberInput - addSteppedField() で作った入力欄
     * @param {boolean} isEnabled - 有効にするなら true
     * @returns {void}
     */
    function setSteppedFieldEnabled(numberInput, isEnabled) {
        numberInput.enabled = isEnabled;
        numberInput.fieldLabel.enabled = isEnabled;
        numberInput.stepperGroup.enabled = isEnabled;
        /* ∧∨は自作描画なので、描き直してディム表示を切り替える / redraw the custom-drawn buttons to update the dimming */
        for (var i = 0; i < numberInput.stepperGroup.children.length; i++) {
            redrawStepperGroup(numberInput.stepperGroup.children[i]);
        }
    }

    /**
     * 入力欄の値を増減する∧∨ボタンを、隙間なく縦に積んで追加する
     * @param {Group|Panel} parent - 追加先
     * @param {Function} getNumberInput - 対象の入力欄を返す関数（入力欄を∧∨より後に作れるよう、クリック時に引く）
     * @param {Object} stepOptions - step（増減量）/ min / max / integer / unit（例 " mm"）/ onStep(numberInput)
     * @returns {Group} ∧∨をまとめた group（.stepBy(direction) で同じ増減を呼べる）
     */
    function addStepper(parent, getNumberInput, stepOptions) {
        var stepperGroup = parent.add("group");
        stepperGroup.orientation = "column";
        stepperGroup.spacing = 0; /* 2つのボタンをつなげて1つの枠に見せる / join the buttons into one frame */
        stepperGroup.margins = [STEPPER_SIDE_MARGIN, 0, 0, 0]; /* 右は入力欄に突き合わせる / butt against the field on the right */
        stepperGroup.alignment = ["left", "center"];

        /**
         * 入力欄の値を増減する（shift を押しながらなら STEPPER_SHIFT_MULTIPLE の倍数へ、option なら STEPPER_OPTION_STEP ずつ。下限・上限で止める）
         * @param {number} direction - 増やすなら 1、減らすなら -1
         * @returns {void}
         */
        function stepBy(direction) {
            var numberInput = getNumberInput();
            if (!isStepperEnabledInTree(numberInput)) return; /* 入力欄か親が無効の間は動かさない */
            var value = parseFloat(numberInput.text);
            if (isNaN(value)) value = 0;
            writeSteppedValue(numberInput, computeSteppedValue(value, direction, stepOptions), stepOptions);
            if (stepOptions.onStep) stepOptions.onStep(numberInput);
        }

        /* 整数の欄では option＋クリックの0.1刻みが効かないので、説明から外す / integer fields have no 0.1 step */
        var upTooltip = stepOptions.integer ? LABELS.tooltip.stepUpInteger : LABELS.tooltip.stepUp;
        var downTooltip = stepOptions.integer ? LABELS.tooltip.stepDownInteger : LABELS.tooltip.stepDown;
        makeStepperChevronButton(stepperGroup, "up", function () { stepBy(1); }).helpTip = getLabel(upTooltip);
        makeStepperChevronButton(stepperGroup, "down", function () { stepBy(-1); }).helpTip = getLabel(downTooltip);
        stepperGroup.stepBy = stepBy; /* ↑↓キーからも同じ処理で増減できるよう公開 / shared with the arrow keys */
        return stepperGroup;
    }

    /**
     * 入力欄の↑↓キーを、∧∨と同じ処理で増減させる。ほかのキーは素通し
     * @param {EditText} numberInput - 対象の入力欄
     * @param {Group} stepperGroup - addStepper() で作った∧∨
     * @returns {void}
     */
    function bindSteppedArrowKeys(numberInput, stepperGroup) {
        numberInput.addEventListener("keydown", function (event) {
            if (event.keyName !== "Up" && event.keyName !== "Down") return;
            stepperGroup.stepBy(event.keyName === "Up" ? 1 : -1);
            event.preventDefault(); /* カーソル移動を止める / keep the caret from moving */
        });
    }

    // -----------------------------------------
    // 値の計算 / Value helpers
    // -----------------------------------------
    /**
     * 押された修飾キーに応じて、1回分増減した値を返す
     * （shift なら STEPPER_SHIFT_MULTIPLE の倍数へ、option なら STEPPER_OPTION_STEP ずつ、それ以外は step の倍数へ（1.5→2、1.5→1）。
     * 整数の欄では option を無視して step の倍数へ）
     * @param {number} value - 元の値
     * @param {number} direction - 増やすなら 1、減らすなら -1
     * @param {Object} stepOptions - step（通常の増減量。省略時は 1）/ integer
     * @returns {number} 増減した値（下限・上限は未適用）
     */
    function computeSteppedValue(value, direction, stepOptions) {
        var keyState = ScriptUI.environment.keyboardState;
        if (keyState.shiftKey) return snapStepperToNextMultiple(value, STEPPER_SHIFT_MULTIPLE, direction);
        if (keyState.altKey && !stepOptions.integer) return value + direction * STEPPER_OPTION_STEP;
        return snapStepperToNextMultiple(value, stepOptions.step || 1, direction);
    }

    /**
     * 値を、指定した方向にある次の倍数へ移す（230→240、232→240、下げるときは 232→230、230→220）
     * @param {number} value - 元の値
     * @param {number} multiple - 倍数の単位（例 10）
     * @param {number} direction - 上げるなら 1、下げるなら -1
     * @returns {number} 移した値
     */
    function snapStepperToNextMultiple(value, multiple, direction) {
        /* 0.29 / 0.01 = 28.999… のような浮動小数の誤差で同じ値に戻らないよう、商を丸めてから切り捨て・切り上げる
           round the quotient first so float error (0.29 / 0.01 = 28.999…) does not step back to the same value */
        var quotient = Math.round(value / multiple * 1e6) / 1e6;
        if (direction > 0) return Math.round((Math.floor(quotient) + 1) * multiple * 1e6) / 1e6;
        return Math.round((Math.ceil(quotient) - 1) * multiple * 1e6) / 1e6;
    }

    /**
     * 値を下限・上限の範囲に収める
     * @param {number} value - 数値
     * @param {Object} rangeOptions - min / max（どちらも省略可）
     * @returns {number} 範囲に収めた値
     */
    function clampSteppedValue(value, rangeOptions) {
        if (rangeOptions.min !== undefined && value < rangeOptions.min) return rangeOptions.min;
        if (rangeOptions.max !== undefined && value > rangeOptions.max) return rangeOptions.max;
        return value;
    }

    /**
     * 値を整数化・下限・上限でそろえ、単位を付けて入力欄に書き込む（直前の正しい値としても控える）
     * @param {EditText} numberInput - 書き込む入力欄
     * @param {number} value - 数値
     * @param {Object} valueOptions - integer / min / max / unit（どれも省略可）
     * @returns {void}
     */
    function writeSteppedValue(numberInput, value, valueOptions) {
        numberInput.text = formatSteppedValue(value, valueOptions);
        numberInput.lastValidText = numberInput.text;
    }

    /**
     * 値を整数化・下限・上限でそろえ、丸めて単位を付けた表示用の文字列にする。
     * 整数化してから下限で止めるので、「整数・下限1」の欄に 0.4 が入っても 1 になる
     * @param {number} value - 数値
     * @param {Object} valueOptions - integer / min / max / unit（どれも省略可）
     * @returns {string} 入力欄に入れる文字列（例 "20 mm"）
     */
    function formatSteppedValue(value, valueOptions) {
        if (valueOptions.integer) value = Math.round(value);
        return formatStepperNumber(clampSteppedValue(value, valueOptions)) + (valueOptions.unit || "");
    }

    /**
     * 小数第2位で丸めた数値を文字列で返す
     * @param {number} value - 数値
     * @returns {string} 表示用の数値文字列
     */
    function formatStepperNumber(value) {
        return String(Math.round(value * 100) / 100);
    }

    // -----------------------------------------
    // ∧∨ボタンの描画 / Drawing
    // -----------------------------------------
    /**
     * 山形（∧／∨）の極小ボタンを作成する。
     * 上下2つを隙間なく積んで1つの枠に見えるよう、枠線は外側の辺だけ描き（上ボタンは上側、下ボタンは下側）、
     * 継ぎ目に線は引かない
     * @param {Group|Panel} parent - 追加先
     * @param {string} direction - "up" または "down"
     * @param {Function} onClickFn - クリック時の処理
     * @returns {Group} ボタンとして使う group
     */
    function makeStepperChevronButton(parent, direction, onClickFn) {
        var buttonWidth = STEPPER_BUTTON_WIDTH;
        var buttonHeight = STEPPER_BUTTON_HEIGHT;
        var isUp = (direction === "up");
        var chevronBox = parent.add("group");
        chevronBox.margins = 0;
        chevronBox.spacing = 0;
        chevronBox.preferredSize = [buttonWidth, buttonHeight];
        chevronBox.minimumSize = [buttonWidth, buttonHeight];
        chevronBox.maximumSize = [buttonWidth, buttonHeight];
        chevronBox.isPressed = false;
        chevronBox.isStepperButton = true; /* redrawSteppersIn() の目印 / marker for redrawSteppersIn() */

        chevronBox.onDraw = function () {
            var boxGraphics = chevronBox.graphics;
            /* 自作描画は自動でディムにならないため、無効なら薄い色で描く。親の無効化は子の enabled に出ないので親も見る
               Custom drawing is not dimmed automatically; the parent's state does not reach the child's enabled */
            var isDimmed = !isStepperEnabledInTree(chevronBox);

            /* 枠線の内側の地（押下中は押下色） / background inside the frame, pressed color while pressed */
            var fillColor = isDimmed ? STEPPER_DIM_FILL_COLOR : (chevronBox.isPressed ? STEPPER_PRESSED_COLOR : STEPPER_FILL_COLOR);
            boxGraphics.newPath();
            boxGraphics.rectPath(1, isUp ? 1 : 0, buttonWidth - 2, buttonHeight - 1);
            boxGraphics.fillPath(boxGraphics.newBrush(boxGraphics.BrushType.SOLID_COLOR, fillColor));

            drawStepperFrame(boxGraphics, buttonWidth, buttonHeight, isUp, isDimmed ? STEPPER_DIM_FRAME_COLOR : STEPPER_FRAME_COLOR);
            drawStepperChevron(boxGraphics, buttonWidth, buttonHeight, isUp, isDimmed ? STEPPER_DIM_CHEVRON_COLOR : STEPPER_CHEVRON_COLOR);
        };

        /**
         * 押下状態を変えて描き直す
         * @param {boolean} isPressed - 押下中なら true
         * @returns {void}
         */
        function repaint(isPressed) {
            if (chevronBox.isPressed === isPressed) return;
            chevronBox.isPressed = isPressed;
            redrawStepperGroup(chevronBox);
        }
        chevronBox.addEventListener("mousedown", function () {
            if (!isStepperEnabledInTree(chevronBox)) return;
            repaint(true);
            if (onClickFn) onClickFn();
        });
        chevronBox.addEventListener("mouseup", function () { repaint(false); });
        /* 押したまま外へ出たときも押下色を残さない / reset when the pointer leaves while pressed */
        chevronBox.addEventListener("mouseout", function () { repaint(false); });
        return chevronBox;
    }

    /**
     * 外側の辺だけの枠を描く（角は丸める）。継ぎ目側は開けておき、上下2つで1つの枠に見せる。
     * ScriptUI は円弧を描けないため、角丸は短い線分で近似する
     * @param {ScriptUIGraphics} boxGraphics - 描画先
     * @param {number} boxWidth - ボタンの幅
     * @param {number} boxHeight - ボタンの高さ
     * @param {boolean} isUp - 上のボタンなら true（上側に枠を描く）
     * @param {number[]} frameColor - [r, g, b, a]
     * @returns {void}
     */
    function drawStepperFrame(boxGraphics, boxWidth, boxHeight, isUp, frameColor) {
        var frameLeft = 0.5;
        var frameRight = boxWidth - 0.5;
        var outerY = isUp ? 0.5 : boxHeight - 0.5;
        var seamY = isUp ? boxHeight : 0;
        var towardSeam = isUp ? 1 : -1; /* 外側の辺から継ぎ目へ向かう向き / direction from the outer edge to the seam */
        var radius = STEPPER_CORNER_RADIUS;
        var arcSteps = 4; /* 角丸1つを何本の線分で近似するか / segments per corner */
        var angle, k;

        boxGraphics.newPath();
        boxGraphics.moveTo(frameLeft, seamY);
        /* 左の角丸 / left corner */
        for (k = 0; k <= arcSteps; k++) {
            angle = (Math.PI / 2) * k / arcSteps;
            boxGraphics.lineTo(frameLeft + radius - radius * Math.cos(angle), outerY + towardSeam * (radius - radius * Math.sin(angle)));
        }
        /* 右の角丸 / right corner */
        for (k = 0; k <= arcSteps; k++) {
            angle = (Math.PI / 2) * k / arcSteps;
            boxGraphics.lineTo(frameRight - radius + radius * Math.sin(angle), outerY + towardSeam * (radius - radius * Math.cos(angle)));
        }
        boxGraphics.lineTo(frameRight, seamY);
        boxGraphics.strokePath(boxGraphics.newPen(boxGraphics.PenType.SOLID_COLOR, frameColor, 1));
    }

    /**
     * 山形（∧／∨）を描く。文字グリフの▲▼は上下で大きさやベースラインが揃わないため、線で描く
     * @param {ScriptUIGraphics} boxGraphics - 描画先
     * @param {number} boxWidth - ボタンの幅
     * @param {number} boxHeight - ボタンの高さ
     * @param {boolean} isUp - ∧なら true、∨なら false
     * @param {number[]} chevronColor - [r, g, b, a]
     * @returns {void}
     */
    function drawStepperChevron(boxGraphics, boxWidth, boxHeight, isUp, chevronColor) {
        var centerX = boxWidth / 2;
        var centerY = isUp ? boxHeight / 2 + 0.5 : boxHeight / 2 - 0.5; /* 継ぎ目から少し離す / nudged away from the seam */
        var halfWidth = 3.6; /* 山形の半幅（高さ1.8に対して開き約127°） / half width of the chevron */
        var tipOffsetY = isUp ? -1.8 : 1.8; /* 頂点の中心からのずれ（上向きは上、下向きは下） */
        boxGraphics.newPath();
        boxGraphics.moveTo(centerX - halfWidth, centerY - tipOffsetY);
        boxGraphics.lineTo(centerX, centerY + tipOffsetY);
        boxGraphics.lineTo(centerX + halfWidth, centerY - tipOffsetY);
        boxGraphics.strokePath(boxGraphics.newPen(boxGraphics.PenType.SOLID_COLOR, chevronColor, 1.2));
    }

    /**
     * コントロールと、その親をたどってすべて有効かを返す（親の無効化は子の enabled に出ない）
     * @param {Object} control - 対象のコントロール
     * @returns {boolean} すべて有効なら true
     */
    function isStepperEnabledInTree(control) {
        for (var node = control; node; node = node.parent) {
            if (!node.enabled) return false;
        }
        return true;
    }

    /**
     * コンテナ以下にある∧∨ボタンをすべて描き直す。行やパネルの enabled を切り替えたあとに呼ぶ
     * @param {Object} container - 行・グループ・パネルなど
     * @returns {void}
     */
    function redrawSteppersIn(container) {
        if (!container.children) return;
        for (var i = 0; i < container.children.length; i++) {
            var child = container.children[i];
            if (child.isStepperButton) redrawStepperGroup(child);
            else redrawSteppersIn(child);
        }
    }

    /**
     * group の onDraw を呼び直す。group には notify() が無いため、隠して再表示して描き直させる
     * @param {Group} targetGroup - 描き直す group
     * @returns {void}
     */
    function redrawStepperGroup(targetGroup) {
        targetGroup.hide();
        targetGroup.show();
    }

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // ステップボタン（再利用パーツ）ここまで / End of the reusable stepper
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // ダイアログの位置と不透明度（再利用パーツ） / Dialog position and opacity (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内に貼る。
    //    識別子は DIALOG_* / prepareDialogWindow / *DialogLeft* / getSelectionViewSpan の名前
    // 2. スクリプトの先頭（#target の次の行）に #targetengine "<SCRIPT_NAME>Engine" を置く。
    //    #targetengine が無いと $.global が実行ごとに消え、位置を覚えられない。すでにあればそのまま使う
    // 3. ダイアログの show() の直前で prepareDialogWindow(dialog, SCRIPT_NAME) を呼ぶ。
    //    それまでに入れた onShow / onMove / onClose はそのまま生かし、あとに位置の復元・記録をつなぐ
    //      prepareDialogWindow(mainDialog, SCRIPT_NAME);
    //      var dialogResult = mainDialog.show();
    //    同じスクリプトで複数のダイアログを開くときは、2つ目以降のキーを変える（SCRIPT_NAME + "_colorPicker" など）
    //    同じダイアログを何度も開くときも、毎回 show() の直前で呼んでよい（2回目からは選択範囲を測り直すだけ）
    // 4. 初めて開くとき（記録が無いとき）は、スクリプト側の配置（中央・オフセットなど）がそのまま効く
    // 5. 開く位置が選択中のオブジェクトに重なりそうなら左右の反対側へずらす（Illustrator のみ）。
    //    ずらした位置は記録せず、ユーザーが動かしたときだけ記録する
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

    var DIALOG_OPACITY = 0.97;       /* ダイアログの不透明度 / dialog opacity */
    var DIALOG_AVOID_MARGIN = 60;    /* 選択範囲の推定位置の両側に取る余裕（px）/ margin on each side of the estimated selection (px) */
    var DIALOG_AVOID_MAX_ITEMS = 100; /* 選択範囲を測るオブジェクトの上限 / max items measured for the selection bounds */

    /**
     * ダイアログの不透明度を設定し、前回閉じた位置で開いて、動かした位置を記録するようにする。
     * 開く位置が選択中のオブジェクトに重なりそうなときは、左右の反対側へずらす（Illustrator のみ）。
     * 既存の onShow / onMove / onClose は先に呼んでから、位置の復元・記録を行う。
     * @param {Window} dialog - 対象のダイアログ
     * @param {string} storageKey - 位置を覚えるキー（ふつうは SCRIPT_NAME）
     * @returns {void}
     */
    function prepareDialogWindow(dialog, storageKey) {
        /* 同じダイアログを開き直すときは、選択範囲を測り直すだけにする（ハンドラーを重ねない）
           When the same dialog is shown again, only re-measure the selection (don't stack handlers) */
        if (dialog.dialogWindowState) {
            dialog.dialogWindowState.selectionSpan = getSelectionViewSpan();
            dialog.dialogWindowState.avoidedLocation = null;
            return;
        }
        var locationKey = "__" + storageKey + "_DialogLocation";
        var previousOnShow = dialog.onShow;
        var previousOnMove = dialog.onMove;
        var previousOnClose = dialog.onClose;
        var windowState = {
            selectionSpan: getSelectionViewSpan(), /* 選択範囲は show() の前に測る / measured before show() */
            screenWidth: null,                     /* 最初に開いたときに推定する / estimated on the first show */
            avoidedLocation: null                  /* 避けるためにずらした位置（記録しない）/ location set to avoid the selection (not remembered) */
        };
        dialog.dialogWindowState = windowState;

        dialog.opacity = DIALOG_OPACITY;

        /* 今の位置を記録する / Remember the current location */
        function rememberDialogLocation() {
            var currentLocation = [dialog.location[0], dialog.location[1]];
            var avoidedLocation = windowState.avoidedLocation;
            if (avoidedLocation && currentLocation[0] === avoidedLocation[0] && currentLocation[1] === avoidedLocation[1]) return;
            $.global[locationKey] = currentLocation;
        }

        dialog.onShow = function () {
            /* 最初に開くときの既定の位置は画面の横中央なので、画面の幅を逆算できる。2回目からは前回の位置なので使い回す
               On the first show the default location is centered horizontally, which gives the screen width; reuse it afterwards */
            if (windowState.screenWidth === null) windowState.screenWidth = dialog.location[0] * 2 + dialog.bounds.width;
            if (previousOnShow) previousOnShow.apply(this, arguments);
            /* $.screens は実際の画面の大きさと合わない（Mac で 1280×524 など）ので、画面内かは判定しない
               $.screens does not match the real display (e.g. 1280x524 on a Mac), so no on-screen check */
            var savedLocation = $.global[locationKey];
            if (savedLocation) dialog.location = [savedLocation[0], savedLocation[1]];
            if (windowState.selectionSpan) {
                var avoidLeft = findDialogLeftAvoidingSelection(dialog.location[0], dialog.bounds.width, windowState.screenWidth, windowState.selectionSpan);
                if (avoidLeft !== null) {
                    dialog.location = [avoidLeft, dialog.location[1]];
                    /* 代入後の値で比べる（丸められることがある）/ Compare with the value after assignment, which may be rounded */
                    windowState.avoidedLocation = [dialog.location[0], dialog.location[1]];
                }
            }
        };
        dialog.onMove = function () {
            if (previousOnMove) previousOnMove.apply(this, arguments);
            rememberDialogLocation();
        };
        dialog.onClose = function () {
            rememberDialogLocation();
            /* false を返すと閉じるのを取りやめるので、戻り値は元の onClose のものを返す
               Returning false cancels the close, so pass the original onClose result through */
            if (previousOnClose) return previousOnClose.apply(this, arguments);
        };
    }

    /**
     * 選択中のオブジェクトが、ドキュメントの表示域の左端から画面上で何 px の範囲にあるかを返す。
     * @returns {{left: number, right: number, viewWidth: number}|null} 選択が無い・測れないときは null
     */
    function getSelectionViewSpan() {
        try {
            if (app.name !== "Adobe Illustrator" || !app.documents.length) return null;
            var targetDoc = app.activeDocument;
            var selectedItems = targetDoc.selection;
            if (!selectedItems || !selectedItems.length || !selectedItems[0].visibleBounds) return null;
            var itemCount = Math.min(selectedItems.length, DIALOG_AVOID_MAX_ITEMS);
            var spanLeft = Infinity;
            var spanRight = -Infinity;
            for (var i = 0; i < itemCount; i++) {
                var itemBounds = selectedItems[i].visibleBounds;
                if (itemBounds[0] < spanLeft) spanLeft = itemBounds[0];
                if (itemBounds[2] > spanRight) spanRight = itemBounds[2];
            }
            var activeView = targetDoc.activeView; /* 複数ウィンドウで開いていても今のウィンドウ / the current window even with multiple windows */
            var viewBounds = activeView.bounds;
            var zoom = activeView.zoom;
            var viewWidth = (viewBounds[2] - viewBounds[0]) * zoom;
            /* 表示域の外にはみ出した部分は数えない / Ignore the part outside the view */
            var left = Math.max(0, (spanLeft - viewBounds[0]) * zoom);
            var right = Math.min(viewWidth, (spanRight - viewBounds[0]) * zoom);
            if (right <= left) return null;
            return { left: left, right: right, viewWidth: viewWidth };
        } catch (e) {
            /* テキスト編集中など測れないときは避けない / Do not avoid when it cannot be measured, e.g. while editing text */
            return null;
        }
    }

    /**
     * ダイアログが選択範囲に重なるなら、重ならない左端の位置を返す。
     * 表示域は画面の横中央にあるとみなし、ずれは DIALOG_AVOID_MARGIN で吸収する。
     * @param {number} dialogLeft - 今のダイアログの左端
     * @param {number} dialogWidth - ダイアログの幅
     * @param {number} screenWidth - 画面の幅
     * @param {{left: number, right: number, viewWidth: number}} selectionSpan - getSelectionViewSpan() の結果
     * @returns {number|null} ずらした左端。重ならない・どちらにも収まらないときは null
     */
    function findDialogLeftAvoidingSelection(dialogLeft, dialogWidth, screenWidth, selectionSpan) {
        var viewLeft = (screenWidth - selectionSpan.viewWidth) / 2;
        var avoidLeft = viewLeft + selectionSpan.left - DIALOG_AVOID_MARGIN;
        var avoidRight = viewLeft + selectionSpan.right + DIALOG_AVOID_MARGIN;
        if (dialogLeft + dialogWidth <= avoidLeft || dialogLeft >= avoidRight) return null;

        var leftSideLeft = avoidLeft - dialogWidth;   /* 選択範囲の左に置くとき / placed left of the selection */
        var rightSideLeft = avoidRight;               /* 選択範囲の右に置くとき / placed right of the selection */
        var fitsLeft = leftSideLeft >= 0;
        var fitsRight = rightSideLeft + dialogWidth <= screenWidth;
        /* 選択範囲が画面の右寄りなら左へ、左寄りなら右へ逃がす / Move away from the side the selection leans to */
        var preferLeft = (avoidLeft + avoidRight) / 2 > screenWidth / 2;
        if (preferLeft && fitsLeft) return leftSideLeft;
        if (fitsRight) return rightSideLeft;
        if (fitsLeft) return leftSideLeft;
        return null;
    }

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // ダイアログの位置と不透明度（再利用パーツ）ここまで / End of the reusable dialog position and opacity
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    (function() {
        // --- 除外リスト保存ファイル / Exclude list persistence ---
        var EXCLUDE_LIST_FILE = File("~/Library/Preferences/FontCatalogExcludeList.json");

        /**
         * 除外リストをファイルから読み込む
         * @param {Array<string>} defaultList - ファイルが無い、または読めない場合に使う既定のリスト
         * @returns {Array<string>} 読み込んだ除外キーワードの配列
         */
        function loadExcludeList(defaultList) {
            try {
                if (EXCLUDE_LIST_FILE.exists) {
                    EXCLUDE_LIST_FILE.open("r");
                    var json = EXCLUDE_LIST_FILE.read();
                    EXCLUDE_LIST_FILE.close();
                    var data = JSON.parse(json);
                    if (data && data instanceof Array) {
                        return data;
                    }
                }
            } catch (e) {}
            return defaultList.slice(); // fallback (clone)
        }

        /**
         * 除外リストをファイルへ保存する
         * @param {Array<string>} list - 保存する除外キーワードの配列
         * @returns {void}
         */
        function saveExcludeList(list) {
            try {
                EXCLUDE_LIST_FILE.open("w");
                EXCLUDE_LIST_FILE.write(JSON.stringify(list, null, 2));
                EXCLUDE_LIST_FILE.close();
            } catch (e) {}
        }

        // --- 1. 基本設定 ---
        var defaultText = "Sample / サンプル";
        var defaultFontSizePt = 24; // internal pt

        // 選択中のテキストを初期値に流用（可能な場合） / Use selected text as default when possible
        // 併せて、選択中テキストのフォント名を取得（可能な場合） / Also capture selected text font name
        var hasSelectedText = false;
        var selectedFontName = "";
        try {
            if (app.documents.length > 0) {
                var selectionItems = app.activeDocument.selection;
                if (selectionItems && selectionItems.length > 0) {
                    var firstSelectionItem = selectionItems[0];

                    // TextFrame
                    if (firstSelectionItem.typename === "TextFrame") {
                        hasSelectedText = true;
                        if (firstSelectionItem.contents) defaultText = firstSelectionItem.contents;
                        try {
                            selectedFontName = firstSelectionItem.textRange.characters[0].characterAttributes.textFont.name;
                        } catch (e1) { selectedFontName = ""; }

                    // TextRange
                    } else if (firstSelectionItem.typename === "TextRange") {
                        hasSelectedText = true;
                        if (firstSelectionItem.contents) defaultText = firstSelectionItem.contents;
                        try {
                            selectedFontName = firstSelectionItem.characters[0].characterAttributes.textFont.name;
                        } catch (e2) { selectedFontName = ""; }
                    }
                }
            }
        } catch (e) {}

        // --- ダイアログボックス / Dialog ---
        var dialog = new Window("dialog", getLabel("dialogTitle") + " " + SCRIPT_VERSION);
        dialog.orientation = "column";
        dialog.alignChildren = ["fill", "top"];

        // --- 2カラムレイアウト / Two-column layout ---
        var columnsGroup = dialog.add("group");
        columnsGroup.orientation = "row";
        columnsGroup.alignChildren = ["fill", "top"];

        var leftColumn = columnsGroup.add("group");
        leftColumn.orientation = "column";
        leftColumn.alignChildren = ["fill", "top"];

        var rightColumn = columnsGroup.add("group");
        rightColumn.orientation = "column";
        rightColumn.alignChildren = ["fill", "top"];

        // --- 表示 / Display ---
        var displayPanel = leftColumn.add("panel", undefined, getLabel("panelSampleSettings"));
        displayPanel.orientation = "column";
        displayPanel.alignChildren = ["fill", "top"];
        displayPanel.margins = [15, 20, 15, 10];

        var textInputGroup = displayPanel.add("group");
        textInputGroup.orientation = "column";
        textInputGroup.alignChildren = ["fill", "top"];
        textInputGroup.add("statictext", undefined, labelText("labelSampleText"));
        var sampleTextInput = textInputGroup.add("edittext", undefined, defaultText);
        sampleTextInput.helpTip = getLabel("tipSampleText");
        sampleTextInput.characters = 20;

        var fontSizeGroup = displayPanel.add("group");
        fontSizeGroup.orientation = "row";

        var textUnitLabel = getCurrentTextUnitLabel();

        // ラベル末尾の「:」を除去（ja/en両対応） / Remove trailing colon from label
        fontSizeGroup.add("statictext", undefined, labelText("labelFontSize"));

        var defaultFontSizeDisplay = ptToCurrentTextUnit(defaultFontSizePt);
        /* ∧∨と入力欄は隙間0で突き合わせる / butt the stepper against the field */
        var fontSizeStepperGroup = fontSizeGroup.add("group");
        fontSizeStepperGroup.orientation = "row";
        fontSizeStepperGroup.alignChildren = ["left", "center"];
        fontSizeStepperGroup.spacing = 0;
        fontSizeStepperGroup.margins = 0;
        /* 文字サイズは0以下にできないので1で止める / font size cannot go to zero */
        var fontSizeStepper = addStepper(fontSizeStepperGroup, function () { return fontSizeInput; }, { step: 1, min: 1 });
        var fontSizeInput = fontSizeStepperGroup.add("edittext", undefined, formatNumberForDisplay(defaultFontSizeDisplay));
        fontSizeInput.helpTip = getLabel("tipFontSize");
        fontSizeInput.characters = 3;
        bindSteppedArrowKeys(fontSizeInput, fontSizeStepper);

        fontSizeGroup.add("statictext", undefined, "(" + textUnitLabel + ")");

        // --- 表示オプション / Display options ---
        var displayOptionGroup = rightColumn.add("panel", undefined, getLabel("panelFontNameDisplay"));
        displayOptionGroup.orientation = "column";
        displayOptionGroup.alignChildren = ["left", "top"];
        displayOptionGroup.margins = [15, 20, 15, 10];

        var showFontNameCheckbox = displayOptionGroup.add(
            "checkbox",
            undefined,
            getLabel("checkboxShowFontName")
        );
        showFontNameCheckbox.value = true;

        var showPostScriptNameCheckbox = displayOptionGroup.add(
            "checkbox",
            undefined,
            getLabel("checkboxShowPostScriptName")
        );
        showPostScriptNameCheckbox.value = false;

        // --- 全幅（貫通）: 除外リスト / Full width: Exclude list ---
        var fullWidthGroup = dialog.add("group");
        fullWidthGroup.orientation = "column";
        fullWidthGroup.alignChildren = ["fill", "top"];

        var excludeListPanel = fullWidthGroup.add("panel", undefined, getLabel("panelExcludeList"));
        excludeListPanel.orientation = "column";
        excludeListPanel.alignChildren = ["fill", "top"];
        excludeListPanel.margins = [15, 20, 15, 10];

        // --- 除外オプション（3カラム） / Exclusion options (3 columns) ---
        var excludeColumnsGroup = excludeListPanel.add("group");
        excludeColumnsGroup.orientation = "row";
        excludeColumnsGroup.alignChildren = ["left", "top"];

        // 1列目 / Column 1
        var excludeCol1 = excludeColumnsGroup.add("group");
        excludeCol1.orientation = "column";
        excludeCol1.alignChildren = ["left", "top"];

        var excludeItalicCheckbox = excludeCol1.add(
            "checkbox",
            undefined,
            getLabel("checkboxExcludeItalic")
        );
        excludeItalicCheckbox.value = true; // 既定で除外 / Default: exclude

        var excludeSystemFontsCheckbox = excludeCol1.add(
            "checkbox",
            undefined,
            getLabel("checkboxExcludeSystemFonts")
        );
        excludeSystemFontsCheckbox.value = true;

        var excludeVariableFontsCheckbox = excludeCol1.add(
            "checkbox",
            undefined,
            getLabel("checkboxExcludeVariableFonts")
        );
        excludeVariableFontsCheckbox.value = true;

        var excludeIllustratorBundledCheckbox = excludeCol1.add(
            "checkbox",
            undefined,
            getLabel("checkboxExcludeIllustratorBundled")
        );
        excludeIllustratorBundledCheckbox.value = true;

        var excludeCompositeFontsCheckbox = excludeCol1.add(
            "checkbox",
            undefined,
            getLabel("checkboxExcludeCompositeFonts")
        );
        excludeCompositeFontsCheckbox.value = true;

        // 2列目 / Column 2
        var excludeCol2 = excludeColumnsGroup.add("group");
        excludeCol2.orientation = "column";
        excludeCol2.alignChildren = ["left", "top"];

        var excludeKoreanCheckbox = excludeCol2.add(
            "checkbox",
            undefined,
            getLabel("checkboxExcludeKorean")
        );
        excludeKoreanCheckbox.value = true;

        var excludeChineseSCCheckbox = excludeCol2.add(
            "checkbox",
            undefined,
            getLabel("checkboxExcludeChineseSC")
        );
        excludeChineseSCCheckbox.value = true;

        var excludeChineseTCCheckbox = excludeCol2.add(
            "checkbox",
            undefined,
            getLabel("checkboxExcludeChineseTC")
        );
        excludeChineseTCCheckbox.value = true;

        var excludeHebrewCheckbox = excludeCol2.add(
            "checkbox",
            undefined,
            getLabel("checkboxExcludeHebrew")
        );
        excludeHebrewCheckbox.value = true;

        var excludeThaiCheckbox = excludeCol2.add(
            "checkbox",
            undefined,
            getLabel("checkboxExcludeThai")
        );
        excludeThaiCheckbox.value = true;

        var excludeArabicCheckbox = excludeCol2.add(
            "checkbox",
            undefined,
            getLabel("checkboxExcludeArabic")
        );
        excludeArabicCheckbox.value = true;

        // 3列目 / Column 3
        var excludeCol3 = excludeColumnsGroup.add("group");
        excludeCol3.orientation = "column";
        excludeCol3.alignChildren = ["left", "top"];

        var excludeMorisawaCheckbox = excludeCol3.add(
            "checkbox",
            undefined,
            getLabel("checkboxExcludeMorisawa")
        );
        excludeMorisawaCheckbox.value = true;

        var excludeFontworksCheckbox = excludeCol3.add(
            "checkbox",
            undefined,
            getLabel("checkboxExcludeFontworks")
        );
        excludeFontworksCheckbox.value = true;

        // --- 除外リスト / Exclude list ---
        var excludeOptionGroup = excludeCol3.add("panel", undefined, getLabel("panelExcludeAdd"));
        excludeOptionGroup.orientation = "column";
        excludeOptionGroup.alignChildren = ["left", "top"];
        excludeOptionGroup.margins = [15, 20, 15, 10];

        var addSelectedFontToExcludeCheckbox = excludeOptionGroup.add(
            "checkbox",
            undefined,
            getLabel("checkboxAddSelectedFontToExclude")
        );
        addSelectedFontToExcludeCheckbox.value = false;

        var selectedFontLabelText = hasSelectedText && selectedFontName
            ? (getLabel("labelSelectedFontPrefix") + selectedFontName)
            : (getLabel("labelSelectedFontPrefix") + getLabel("labelSelectedFontNone"));
        var selectedFontLabel = excludeOptionGroup.add("statictext", undefined, selectedFontLabelText);

        // 選択テキストがない場合は無効化 / Disable when there is no selected text font
        if (!(hasSelectedText && selectedFontName)) {
            addSelectedFontToExcludeCheckbox.enabled = false;
        }

        var buttonRow = addButtonRow(dialog);
        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("buttonCancel"), {name:"cancel"});
        var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel("buttonOK"), {name:"ok"});

        prepareDialogWindow(dialog, SCRIPT_NAME);
        if (dialog.show() !== 1) return;

        var inputStr = sampleTextInput.text;
        var fontSizeInputValue = Number(fontSizeInput.text); // value in current text unit
        var fontSize = currentTextUnitToPt(fontSizeInputValue); // internal pt

        var showFontName = showFontNameCheckbox.value === true;
        var showPostScriptName = showPostScriptNameCheckbox.value === true;

        var excludeItalicFonts = excludeItalicCheckbox.value === true;
        var excludeKoreanFonts = excludeKoreanCheckbox.value === true;
        var excludeChineseSCFonts = excludeChineseSCCheckbox.value === true;
        var excludeChineseTCFonts = excludeChineseTCCheckbox.value === true;
        var excludeHebrewFonts = excludeHebrewCheckbox.value === true;
        var excludeThaiFonts = excludeThaiCheckbox.value === true;
        var excludeArabicFonts = excludeArabicCheckbox.value === true;
        var excludeSystemFonts = excludeSystemFontsCheckbox.value === true;
        var excludeVariableFonts = excludeVariableFontsCheckbox.value === true;
        var excludeIllustratorBundledFonts = excludeIllustratorBundledCheckbox.value === true;
        var excludeCompositeFonts = excludeCompositeFontsCheckbox.value === true;
        var excludeMorisawaFonts = excludeMorisawaCheckbox.value === true;
        var excludeFontworksFonts = excludeFontworksCheckbox.value === true;
        var addSelectedFontToExclude = addSelectedFontToExcludeCheckbox.value === true;
        // バリアブルフォント / Variable fonts
        var variableFontKeywords = [
            "Variable",
            "VF",
            "Var"
        ];
        if (!inputStr || isNaN(fontSizeInputValue) || fontSizeInputValue <= 0) return;
        if (isNaN(fontSize) || fontSize <= 0) return;

        var columnGap = 80;   // 列同士の間隔
        var rowGap = 20;      // 行同士の間隔
        var rowLimit = 20;    // 1列に並べる個数（後で自動計算で上書き） / Will be overwritten by auto calc

        var infoFontSize = 9;        // 下部表示テキストのサイズ / Info text size
        var infoTopGap = 2;          // 見本と情報の間隔 / Gap between sample and info
        var infoBottomGap = 6;       // 情報と次行の間隔 / Gap after info block

        var infoFontPostScriptName = "HiraginoSans-W3"; // 情報表示用フォント（PostScript名） / Info font (PostScript)
        var infoTextFont = null;
        try {
            infoTextFont = app.textFonts.getByName(infoFontPostScriptName);
        } catch (e) {
            infoTextFont = null;
        }

        var documentRef = app.documents.add();
        var allFonts = app.textFonts;
        var fontCount = allFonts.length; // パフォーマンス改善のためキャッシュ

        // --- 進捗表示（プログレスバー） / Progress bar ---
        var progressPalette = new Window("palette", getLabel("progressTitle"));
        progressPalette.orientation = "column";
        progressPalette.alignChildren = ["fill", "top"];

        var progressLabel = progressPalette.add("statictext", undefined, "0 / " + fontCount);

        var progressBar = progressPalette.add("progressbar", undefined, 0, fontCount);
        progressBar.preferredSize.width = 320;

        var progressButtonGroup = progressPalette.add("group");
        progressButtonGroup.alignment = "right";

        var progressCancelButton = progressButtonGroup.add("button", undefined, getLabel("progressCancel"));
        var isCancelled = false;

        progressCancelButton.onClick = function () {
            isCancelled = true;
            try {
                progressCancelButton.enabled = false;
                progressLabel.text = getLabel("progressCancelled");
                progressPalette.update();
            } catch (e) {}
        };

        progressPalette.show();

        // アートボードの座標取得 [左, 上, 右, 下]
        var activeArtboard = documentRef.artboards[documentRef.artboards.getActiveArtboardIndex()];
        var artboardRect = activeArtboard.artboardRect;

        // 開始座標を左上に設定（左に10ptの余白、1行目がはみ出ないようfontSize分だけ下にオフセット）
        var leftMarginPt = 10; // pt
        var startPosX = artboardRect[0] + leftMarginPt;
        var startPosY = artboardRect[1] - fontSize;

        // rowLimit をフォントサイズとアートボード高さから自動計算 / Auto-calc rowLimit from font size and artboard height
        // 1行目は startPosY、以降は (fontSize + rowGap) ずつ下げるため、入る行数 = floor(availableHeight / step) + 1
        try {
            var topY = startPosY;
            var bottomY = artboardRect[3];
            var availableHeight = topY - bottomY;
            var stepY = fontSize + rowGap;
            if (stepY > 0) {
                var autoRowLimit = Math.floor(availableHeight / stepY) + 1;
                if (autoRowLimit < 1) autoRowLimit = 1;
                rowLimit = autoRowLimit;
            }
        } catch (e) {}

        var currentX = startPosX;
        var currentY = startPosY;
        var maxColumnWidth = 0;
        var placedItemCount = 0;

        // 言語別キーワード（UIでON/OFF制御） / Language keyword groups (controlled by UI)
        var koreanFontKeywords = [
            "Korean",
            "Hangul",
            "AppleSDGothicNeo",
            "Malgun",
            "Gulim",
            "Dotum",
            "Batang",
            "Nanum",
            "AdobeMyungjoStd",    // Adobe 명조 Std (Korean)
            "AdobeGothicStd"     // Adobe 고딕 Std (Korean)
        ];

        // 中国語（簡体） / Chinese (Simplified)
        var chineseSCFontKeywords = [
            "Chinese",
            "PingFangSC",
            "PingFang SC",
            "HeitiSC",
            "Heiti SC",
            "SongtiSC",
            "Songti SC",
            "SimSun",
            "SimHei",
            "MicrosoftYaHei",
            "Microsoft YaHei",
            "GB",
            "AdobeHeitiStd",      // Adobe 黑体
            "AdobeKaitiStd",      // Adobe 楷体
            "AdobeFangsongStd",   // Adobe 仿宋
            "AdobeSongStd",       // Adobe 宋体
            "文鼎",
            "ARS",
            "Arphic",
            "AR PL"
        ];

        // 中国語（繁体） / Chinese (Traditional)
        var chineseTCFontKeywords = [
            "Chinese",
            "標楷體",
            "標楷體-港澳",
            "標楷體-繁",
            "PingFangTC",
            "PingFang TC",
            "HeitiTC",
            "Heiti TC",
            "SongtiTC",
            "Songti TC",
            "MingLiU",
            "PMingLiU",
            "DFKai",
            "Big5",
            "AdobeMingStd",       // Adobe 明體 Std (Traditional Chinese)
            "AdobeHeitiStd",      // Adobe 黑体
            "AdobeKaitiStd",      // Adobe 楷体
            "AdobeFangsongStd",   // Adobe 仿宋
            "AdobeSongStd",       // Adobe 宋体
            "文鼎",
            "ARS",
            "Arphic",
            "AR PL"
        ];

        // ヘブライ語 / Hebrew
        var hebrewFontKeywords = [
            "Hebrew",
            "MyriadHebrew",
            "AdobeHebrew",
            "Narkisim",
            "Ezra",
            "FrankRuehl",
            "David"
        ];

        // タイ語 / Thai
        var thaiFontKeywords = [
            "Thai",
            "Thonburi",
            "Ayuthaya",
            "Krungthep",
            "Sukhumvit",
            "Silom",
            "NotoSansThai",
            "NotoSerifThai"
        ];

        // アラビア系 / Arabic
        var arabicFontKeywords = [
            "Arabic",
            "AdobeArabic",
            "Geeza",
            "Diwan",
            "Kufi",
            "Naskh",
            "NotoNaskhArabic",
            "NotoKufiArabic",
            "NotoSansArabic"
        ];

        // メーカー別キーワード（UIでON/OFF制御） / Vendor keyword groups (controlled by UI)
        var morisawaFontKeywords = [
            "A-OTF",
            "A P-OTF",
            "AP-OTF",
            "A-SK"
        ];

        var fontworksFontKeywords = [
            "FOT-",   // explicit prefix used by many Fontworks PS names
            "FOT"     // fallback
        ];

        // システムフォント（macOS） / System fonts (macOS)
        var systemFontKeywords = [
            "AdobeCleanUX",
            "ADTNumeric",
            "Apple Braille",
            "Apple Symbols",
            "AppleSDGothicNeo",
            "AquaKana",
            "ArialHB",
            "Avenir",
            "Avenir Next",
            "Avenir Next Condensed",
            "CJKSymbolsFallback",
            "Courier",
            "DecoTypeNastaleeqUrdu",
            "GeezaPro",
            "Geneva",
            "HelveLTMM",
            "Helvetica",
            "HelveticaNeue",
            "Hiragino Sans GB",
            "Keyboard",
            "Kohinoor",
            "KohinoorBangla",
            "KohinoorGujarati",
            "KohinoorTelugu",
            "LastResort",
            "LucidaGrande",
            "MarkerFelt",
            "Menlo",
            "Monaco",
            "MuktaMahee",
            "NewYork",
            "Noteworthy",
            "Optima",
            "Palatino",
            "SF",
            "STHeiti",
            "Symbol",
            "ThonburiUI",
            "Times",
            "TimesLTMM",
            "ZapfDingbats",
            "Zither",
            "ヒラギノ角ゴシック",
            "ヒラギノ丸ゴ",
            "ヒラギノ明朝"
        ];

        // Illustrator付属フォント / Bundled with Illustrator
        // Note: list is based on provided file names; we match by partial keyword against fontItem.name/family/style.
        var illustratorBundledFontKeywords = [
            "AcuminVariableConcept",
            "AdobeArabic",
            "AdobeInvisible",
            "AdobeMingStd",
            "AdobeMyungjoStd",
            "AdobeSongStd",
            "EmojiOneColor",
            "KozGoPr6N",
            "MinionVariableConcept",
            "MyriadHebrew",
            "MyriadPro",
            "MyriadVariableConcept",
            "SourceCodeVariable",
            "SourceHanSansSC",
            "SourceSansVariable",
            "SourceSerifVariable",
            "TrajanColor"
        ];

        // 合成フォント（ATC-） / Composite fonts (ATC-*)
        var compositeFontPrefix = "ATC-";

        /**
         * キーワードがリストに完全一致で含まれるかを判定する
         * @param {string} keyword - 探すキーワード
         * @param {Array<string>} list - 検索対象のリスト
         * @returns {boolean} 含まれていれば true
         */
        function isKeywordInList(keyword, list) {
            for (var ii = 0; ii < list.length; ii++) {
                if (list[ii] === keyword) return true;
            }
            return false;
        }

        /**
         * リスト内のいずれかのキーワードが文字列に含まれるかを、大文字小文字を区別せずに判定する
         * @param {string} text - 検索対象の文字列
         * @param {Array<string>} list - キーワードの配列
         * @returns {boolean} いずれかが含まれていれば true
         */
        function containsAnyKeywordCI(text, list) {
            for (var kk = 0; kk < list.length; kk++) {
                if (containsIgnoreCase(text, list[kk])) return true;
            }
            return false;
        }

        // --- 2. 除外リスト (編集エリア) ---
        // 非表示にしたいフォントのキーワードを自由に足してください。
        // より具体的なパターンを先に配置
        var excludeList = loadExcludeList([
            "HiraMin",           // ヒラギノ明朝
            "HiraKaku",          // ヒラギノ角ゴ
            "Hiragino",          // ヒラギノシリーズ
            "MS-P",              // MS Pゴシック/明朝
            "MS-Gothic",         // MS ゴシック
            "MS-Mincho",         // MS 明朝
            "YuMincho",          // 游明朝
            "YuGothic",          // 游ゴシック
            "Meiryo",            // メイリオ
            "Noto",              // Noto フォント全般
            "NotoSans",          // Noto Sans
            "NotoSerif",         // Noto Serif
            "SourceHan",         // 源ノ角ゴシック/源ノ明朝
            "Koz",               // 小塚ゴシック/明朝
            "AdobeSongStd",      // Adobe 宋体
            "BIZ-UD",            // BIZ UDシリーズ
            "UDDigiKyo",         // UD デジタル教科書体
            "HGP",               // HG系 (P)
            "HGS",               // HG系 (S)
            "HG",                // HG系
            "Font Awesome",      // Font Awesome (icon font)
            "Adobe Clean",       // Adobe Clean (UI/system font)
            "Artifakt Element",  // Artifakt Element (UI font)
            "Ornaments",         // Ornaments (icon/ornament fonts)
            // STIX math fonts
            "STIXIntegralsD",
            "STIXIntegralsUpD",
            "STIXSizeFiveSym",
            "STIXSizeThreeSym",
            "STIXIntegralsSm",
            "STIXIntegralsUpSm",
            "STIXSizeFourSym",
            "STIXSizeTwoSym",
            "STIXIntegralsUp",
            "STIXNonUnicode",
            "STIXSizeOneSym",
            "STIXVariants",
            "AppleColorEmoji",   // 絵文字 (Mac)
            "SegoeUIEmoji"       // 絵文字 (Win)
        ]);

        // 韓国語/中国語/システムフォントキーワードは専用チェックで制御するため、永続リストからは除外 / Keep language/system keywords out of persisted list
        try {
            var cleaned = [];
            for (var c = 0; c < excludeList.length; c++) {
                var kw2 = excludeList[c];
                if (isKeywordInList(kw2, koreanFontKeywords)) continue;
                if (isKeywordInList(kw2, chineseSCFontKeywords)) continue;
                if (isKeywordInList(kw2, chineseTCFontKeywords)) continue;
                if (isKeywordInList(kw2, hebrewFontKeywords)) continue;
                if (isKeywordInList(kw2, thaiFontKeywords)) continue;
                if (isKeywordInList(kw2, arabicFontKeywords)) continue;
                if (isKeywordInList(kw2, systemFontKeywords)) continue;
                if (isKeywordInList(kw2, variableFontKeywords)) continue;
                if (isKeywordInList(kw2, illustratorBundledFontKeywords)) continue;
                if (kw2.indexOf("ATC-") === 0) continue;
                cleaned.push(kw2);
            }
            excludeList = cleaned;
        } catch (eClean) {}

        // 選択フォントを除外リストに追加（任意） / Optionally add selected font to exclude list
        if (addSelectedFontToExclude && selectedFontName) {
            var alreadyExists = false;
            for (var k = 0; k < excludeList.length; k++) {
                if (excludeList[k] === selectedFontName) {
                    alreadyExists = true;
                    break;
                }
            }
            if (!alreadyExists) {
                excludeList.unshift(selectedFontName); // 先頭に追加（優先） / Add to top for priority
            }
        }
        // 除外リストを保存 / Save exclude list
        saveExcludeList(excludeList);

        // --- 3. メイン処理 ---
        for (var i = 0; i < fontCount; i++) {
            // progress update
            try {
                progressBar.value = i;
                progressLabel.text = i + " / " + fontCount;
                progressPalette.update();
            } catch (e) {}
            // cancel check
            if (isCancelled) {
                break;
            }

            var textFrame = null; // テキストフレームの参照を保持 / Keep refs for cleanup
            var infoTextFrame = null; // フォント情報表示用 / For font info display
            try {
                var fontItem = allFonts[i];
                if (!fontItem) continue;

                var fontName = fontItem.name;
                var fontStyle = "";
                try { fontStyle = fontItem.style; } catch(e) { fontStyle = ""; }

                // 斜体を除外（任意） / Optionally exclude italic fonts
                if (excludeItalicFonts) {
                    if (fontName.indexOf("Italic") !== -1 ||
                        fontName.indexOf("Oblique") !== -1 ||
                        fontStyle.indexOf("Italic") !== -1 ||
                        fontStyle.indexOf("Oblique") !== -1) {
                        continue;
                    }
                }

                // メーカー別フォント除外（任意） / Optionally exclude vendor fonts
                // PostScript名だけでなく family/style も含めて判定 / Check name + family/style for robustness
                var fontSearchText = fontName;
                try {
                    var fam2 = "";
                    var sty2 = "";
                    try { fam2 = fontItem.family || ""; } catch (eFam2) { fam2 = ""; }
                    try { sty2 = fontItem.style || ""; } catch (eSty2) { sty2 = ""; }
                    fontSearchText = fontName + " | " + fam2 + " | " + sty2;
                } catch (eFS) {}

                if (excludeMorisawaFonts) {
                    if (containsAnyKeywordCI(fontSearchText, morisawaFontKeywords)) {
                        continue;
                    }
                }
                if (excludeFontworksFonts) {
                    if (containsAnyKeywordCI(fontSearchText, fontworksFontKeywords)) {
                        continue;
                    }
                }

                // システムフォント除外（任意） / Optionally exclude system fonts
                // name/family/style をまとめて大文字小文字無視で判定（空白差異なども吸収） / Case-insensitive across name/family/style
                if (excludeSystemFonts) {
                    if (containsAnyKeywordCI(fontSearchText, systemFontKeywords)) {
                        continue;
                    }
                }

                // バリアブルフォント除外（任意） / Optionally exclude variable fonts
                if (excludeVariableFonts) {
                    if (containsAnyKeywordCI(fontSearchText, variableFontKeywords)) {
                        continue;
                    }
                }

                // Illustrator付属フォント除外（任意） / Optionally exclude Illustrator bundled fonts
                if (excludeIllustratorBundledFonts) {
                    if (containsAnyKeywordCI(fontSearchText, illustratorBundledFontKeywords)) {
                        continue;
                    }
                }

                // 合成フォント除外（任意） / Optionally exclude composite fonts (ATC-*)
                if (excludeCompositeFonts) {
                    if (fontName.indexOf(compositeFontPrefix) === 0) {
                        continue;
                    }
                }

                // 言語別フォント除外（任意） / Optionally exclude language fonts
                if (excludeKoreanFonts) {
                    if (containsAnyKeywordCI(fontSearchText, koreanFontKeywords)) {
                        continue;
                    }
                }
                if (excludeChineseSCFonts) {
                    if (containsAnyKeywordCI(fontSearchText, chineseSCFontKeywords)) {
                        continue;
                    }
                }
                if (excludeChineseTCFonts) {
                    if (containsAnyKeywordCI(fontSearchText, chineseTCFontKeywords)) {
                        continue;
                    }
                }
                if (excludeHebrewFonts) {
                    if (containsAnyKeywordCI(fontSearchText, hebrewFontKeywords)) {
                        continue;
                    }
                }
                if (excludeThaiFonts) {
                    if (containsAnyKeywordCI(fontSearchText, thaiFontKeywords)) {
                        continue;
                    }
                }
                if (excludeArabicFonts) {
                    if (containsAnyKeywordCI(fontSearchText, arabicFontKeywords)) {
                        continue;
                    }
                }

                // 除外リスト照合（言語グループはUIでON/OFF） / Exclude list match (language groups are toggleable)
                // name+family+style で大文字小文字無視、さらに空白差異も吸収 / Case-insensitive across name/family/style + ignore whitespace differences
                var isFontExcluded = false;
                var hay = fontSearchText || fontName || "";
                var hayNoSpace = String(hay).replace(/\s+/g, "");

                for (var j = 0; j < excludeList.length; j++) {
                    var kw = excludeList[j];
                    if (!kw) continue;

                    // 1) 通常（大文字小文字無視） / Normal case-insensitive match
                    if (containsIgnoreCase(hay, kw)) {
                        isFontExcluded = true;
                        break;
                    }

                    // 2) 空白除去して比較（"Font Awesome" vs "FontAwesome" 等） / Ignore whitespace differences
                    var kwNoSpace = String(kw).replace(/\s+/g, "");
                    if (kwNoSpace && containsIgnoreCase(hayNoSpace, kwNoSpace)) {
                        isFontExcluded = true;
                        break;
                    }
                }
                if (isFontExcluded) continue;

                // 描画テスト（合成フォント・エラーフォント回避）
                textFrame = documentRef.textFrames.add();
                textFrame.contents = inputStr;

                try {
                    textFrame.textRange.characterAttributes.textFont = fontItem;
                    textFrame.textRange.characterAttributes.size = fontSize;
                } catch(e) {
                    if (textFrame && textFrame.parent) textFrame.remove();
                    if (infoTextFrame && infoTextFrame.parent) infoTextFrame.remove();
                    continue;
                }

                // フォント適用確認（フォールバック検知）
                if (textFrame.textRange.characters[0].characterAttributes.textFont.name !== fontName) {
                    if (textFrame && textFrame.parent) textFrame.remove();
                    if (infoTextFrame && infoTextFrame.parent) infoTextFrame.remove();
                    continue;
                }

                // --- 配置 / Placement ---
                textFrame.left = currentX;
                textFrame.top = currentY;

                // フォント情報を見本の下に表示（任意） / Optionally show font info under the sample
                if (showFontName || showPostScriptName) {
                    var fontObj = textFrame.textRange.characterAttributes.textFont;
                    var family = "";
                    var style = "";
                    try { family = fontObj.family; } catch (eFam) { family = ""; }
                    try { style = fontObj.style; } catch (eSty) { style = ""; }

                    var postScriptName = "";
                    try { postScriptName = textFrame.textRange.characterAttributes.textFont.name; } catch (ePS) { postScriptName = ""; }

                    var nameWithStyle = family;
                    if (style) {
                        nameWithStyle += " " + style;
                    }

                    var fontLine = "";

                    if (showFontName && showPostScriptName) {
                        // 2 lines:
                        // <family> <style>
                        // <PostScriptName>
                        fontLine = nameWithStyle + "\r" + postScriptName;
                    } else if (showFontName) {
                        // 1 line: <family> <style>
                        fontLine = nameWithStyle;
                    } else if (showPostScriptName) {
                        // 1 line: <PostScriptName>
                        fontLine = postScriptName;
                    }

                    infoTextFrame = documentRef.textFrames.add();
                    infoTextFrame.contents = fontLine;

                    // プロポーショナルメトリクスをON（情報表示のみ） / Enable proportional metrics for info text only
                    try { infoTextFrame.textRange.proportionalMetrics = true; } catch (ePM) {}

                    if (infoTextFont) {
                        try { infoTextFrame.textRange.characterAttributes.textFont = infoTextFont; } catch (eFont) {}
                    }
                    infoTextFrame.textRange.characterAttributes.size = infoFontSize;

                    // 情報は見本の下に配置 / Place info under the sample
                    infoTextFrame.left = currentX;
                    infoTextFrame.top = currentY - (fontSize + infoTopGap);

                    // 列幅計算に反映 / Include in column width
                    if (infoTextFrame.width > maxColumnWidth) {
                        maxColumnWidth = infoTextFrame.width;
                    }

                    // 次の行へ（情報ブロック分も下げる） / Move down including info block height
                    currentY -= (fontSize + rowGap + infoTextFrame.height + infoTopGap + infoBottomGap);
                } else {
                    // 次の行へ / Next row
                    currentY -= (fontSize + rowGap);
                }

                if (textFrame.width > maxColumnWidth) {
                    maxColumnWidth = textFrame.width;
                }

                placedItemCount++;

                // 20個並んだら次の列へ
                if (placedItemCount % rowLimit === 0) {
                    currentX += (maxColumnWidth + columnGap);
                    currentY = startPosY;
                    maxColumnWidth = 0;
                }

            } catch (err) {
                // エラー時はテキストフレームを確実に削除
                try {
                    if (textFrame && textFrame.parent) textFrame.remove();
                    if (infoTextFrame && infoTextFrame.parent) infoTextFrame.remove();
                } catch(e) {}
            }
        }
        // 完了時にプログレスバーを閉じる
        try {
            progressBar.value = fontCount;
            progressLabel.text = fontCount + " / " + fontCount;
            progressPalette.update();
            progressPalette.close();
        } catch (e) {}
    })();

})();
