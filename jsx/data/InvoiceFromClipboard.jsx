#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

クリップボードにある「見出し行＋値行」を空行で区切った形式のテキストを読み取り、テンプレート内の `<タグ>` を置き換えてPDFを書き出す作例スクリプトです。
税込金額から税抜・消費税を計算し、書き出したあとは返信メールの文面をクリップボードへ入れて、保存先フォルダーを開きます。

### 注意

- テンプレートは［指定］で選んだものを記憶します（Illustratorを再起動しても残ります）。置換は作業用の複製ファイルで行うため、テンプレート自体は変更されません。
- 見出しとタグの対応、税率、ファイル名の付け方、メールの文面はユーザー設定で差し替えられます。

詳しい機能・使い方はREADMEを参照してください。

*/

/*

### Overview

A worked example that reads blank-line separated heading-and-value text from the clipboard, replaces `<tag>`
placeholders in an Illustrator template, and exports a PDF.
The ex-tax amount and the tax are derived from the tax-inclusive amount; after the export the reply mail is
placed on the clipboard and the output folder is revealed.

### Notes

- The template picked with Choose is remembered across Illustrator restarts. Replacement happens on a working
  copy, so the template itself is never modified.
- The heading-to-tag map, the tax rate, the file naming and the mail body are all user settings.

See the README for the full feature list and usage.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "InvoiceFromClipboard";         /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-08-16";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-08-16";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/InvoiceFromClipboard.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/InvoiceFromClipboard.md

/**
 * @discussion クリップボードの読み取りは ReplaceTextWithPasteSequential.jsx から流用
 * @discussion タグ置換は VariableDataImport.jsx から流用
 */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

// =========================================
// ユーザー設定 / User settings
// =========================================

/*
   クリップボードの見出しと、テンプレート内のタグの対応。
   valueType の意味：
     documentType        … 書類名。ダイアログの一番上で選ぶ（見出しは持たない）
     text                … 値をそのまま入れる
     application         … 適用（但し書き）。下の APPLICATION_PRESETS から選んだ書き方で組み立てる
     date                … 「2026年8月14日」形式に整える
     amountIncludingTax  … 税込金額
     amountExcludingTax  … 税込金額から求めた税抜金額
     tax                 … 税込金額から求めた消費税額
   Clipboard headings mapped to template tags.
*/
var FIELD_MAPPINGS = [
    { tag: "<タイトル>",         heading: "",               valueType: "documentType" },
    { tag: "<日付>",             heading: "日付",           valueType: "date" },
    { tag: "<御中>",             heading: "領収書の宛先",    valueType: "text" },
    { tag: "<適用>",             heading: "該当イベント名",  valueType: "application" },
    { tag: "<Price>",            heading: "金額",           valueType: "amountIncludingTax" },
    { tag: "<PriceWithoutTax>",  heading: "金額",           valueType: "amountExcludingTax" },
    { tag: "<Tax>",              heading: "金額",           valueType: "tax" }
];

/*
   適用（但し書き）の書き方。ダイアログのラジオボタンで選ぶ。
   `#見出し#` は、その見出しの入力値に置き換わる。先頭が既定の選択。
   テンプレート側は「但  <適用>」のように、一文まるごとをタグにしておく。
   How the "適用" line is worded; the first entry is selected by default.
*/
var APPLICATION_PRESETS = [
    { label: "セミナーイベント", format: "セミナーアーカイブ（#該当イベント名#）" },
    { label: "その他",           format: "#該当イベント名#" }
];

/*
   入力欄に出すラベル。クリップボードの見出しと違う名前で見せたいときだけ書く。
   ここに無い見出しは、見出しの文字がそのままラベルになる。
   Display labels for the input rows (a heading not listed here labels itself).
*/
var FIELD_LABEL_OVERRIDES = {
    "該当イベント名": "適用"
};

/*
   対応表には無いが、見出しとして区切りに使う項目。
   空行が無い形式で貼り付けられても、ここに挙げた見出しで値の終わりを判断できる。
   Headings that are not merged but still act as value boundaries.
*/
var IGNORED_HEADINGS = ["購入くださった窓口", "購入時のユーザー名"];

/*
   領収書に載せない文字を、見出しごとに取り除く。出現位置は問わない。
   Strings dropped from a value before it is merged, per heading.
*/
var VALUE_REMOVAL_PATTERNS = [
    { heading: "該当イベント名", pattern: /【法人向け】/g }
];

/*
   書類名の選択肢。ダイアログの一番上のラジオボタンになる。
   初期選択はテンプレートのファイル名から決まり（「領収書テンプレート.ai」なら領収書）、
   どれも含まれていなければ配列の先頭を使う。選んだ書類名は、テンプレートの `<タイトル>`、
   ダイアログのタイトル、PDFのファイル名、メールの件名・本文に反映される。
   Document types offered as radio buttons; the template's file name picks the initial one.
*/
var DOCUMENT_TYPE_NAMES = ["領収書", "請求書", "納品書"];

/*
   書き出し後にクリップボードへ入れる返信メールの文面。
   `#見出し#` はクリップボードから読み取った値に、`#ファイル名#` は作成したPDFの名前に、
   `#書類名#` はテンプレートから決めた書類名に置き換わる。
   ここで使った見出しは、タグに使っていなくても読み取り対象になる。
   The reply mail placed on the clipboard after the export.
*/
var REPLY_MAIL_FILE_NAME_TOKEN = "ファイル名";
var DOCUMENT_TYPE_TOKEN = "書類名";
var REPLY_MAIL_SUBJECT = "#書類名#をお送りします";
var REPLY_MAIL_TEMPLATE = [
    "#お名前#さん、",
    "アーカイブ購入ありがとうございました！",
    "",
    "#書類名#PDFを添付します。",
    "（#ファイル名#）",
    "",
    "よろしくお願いします。"
].join("\n");

/*
   書き出し後に、既定のメールソフトで下書きを開く。
   宛先は下の見出しから取り、件名と本文は上の設定を流し込む。
   mailtoでは添付を指定できないため、PDFは開いた保存先フォルダーから手で添付する。
   Open a draft in the default mail client after the export (attachments are not possible via mailto).
*/
var OPEN_MAIL_AFTER_EXPORT = true;
var MAIL_ADDRESS_HEADING = "メールアドレス"; /* 宛先に使う見出し（空文字で宛先なし）/ heading used as the To address */

/*
   Illustratorテンプレートの初期パス。ダイアログで［指定］すると、以降はそちらが記憶されて優先される。
   空文字にすると、最初の実行から［指定］で選ばせる。
   Initial template path; whatever is picked in the dialog takes precedence from then on.
*/
var DEFAULT_TEMPLATE_PATH = "/Users/takano/sw Dropbox/takano masahiro/Dropbox-shared/CeeBeeDee/領収書/領収書テンプレート.ai";

/*
   PDFファイル名の組み立て方。区切りをはさんで
   「書類名・発行元・日付（8桁）・宛先」の順につなぐ。
   書類名はテンプレートのファイル名から決まり、発行元を空文字にするとその部分ごと省く。
   例：領収書-CeeBeeDee-20260814-株式会社セガ.pdf
   How the PDF file name is assembled ("" issuer drops that part).
*/
var PDF_FILE_NAME_ISSUER = "CeeBeeDee";      /* 書類名の次に入れる発行元（空文字で省略）/ issuer ("" omits it) */
var PDF_FILE_NAME_SEPARATOR = "-";           /* 各要素の区切り / separator between the parts */
var PDF_FILE_NAME_HEADING = "領収書の宛先";  /* 宛先に使う見出し / heading used as the recipient */
/*
   PDFの書き出しに使うプリセット名の候補。
   プリセット名は環境の言語で変わるため、候補を並べて、実際に用意されているものを使う。
   どれも見つからない環境では、下の個別設定で同等の書き出しを組み立てる。
   Preset names are localised, so list the candidates and use whichever exists.
*/
var PDF_PRESET_CANDIDATES = ["[最小ファイルサイズ]", "[Smallest File Size]"];
var PDF_COMPATIBILITY_NAME = "ACROBAT7";     /* 互換性：ACROBAT5=1.4 / ACROBAT6=1.5 / ACROBAT7=1.6 / ACROBAT8=1.7 */
var PDF_PRESERVE_EDITABILITY = false;        /* Illustratorの編集機能を保持（trueにすると大幅に重くなる）/ keeps AI data */
var PDF_IMAGE_RESOLUTION = 100;              /* カラー・グレースケール画像の解像度（ppi、0で縮小しない）/ downsample target */
var PDF_IMAGE_THRESHOLD = 150;               /* この解像度を超える画像だけ縮小する（ppi）/ downsample above this */
var OVERWRITE_EXISTING_PDF = false;          /* 同名PDFを上書きする（falseなら連番を付ける）/ overwrite an existing PDF */
var OPEN_FOLDER_AFTER_EXPORT = true;         /* 書き出し後に保存先フォルダーを開く / reveal the output folder afterwards */

var TAX_RATE = 0.1;                          /* 消費税率（単一税率）/ tax rate */
var TAX_FRACTION_MODE = "floor";             /* 税抜の端数処理："floor" / "round" / "ceil" */

/* Dropboxのローカルマウントパス。空文字にするとホーム直下から自動検出 / Local Dropbox mount path ("" = auto detect) */
var DROPBOX_MOUNT_PATH = "";

/*
   マウントパスの直下で、さらに表示から省くフォルダー名。
   共有フォルダー（「Dropbox-shared」など）は自動検出では外れないため、ここに挙げる。
   例：/Users/<ユーザー>/<チーム> Dropbox/<アカウント>/Dropbox-shared/CeeBeeDee/領収書/領収書テンプレート.ai
       → CeeBeeDee/領収書/領収書テンプレート.ai
   Extra folders dropped from displayed paths, right under the Dropbox mount.
*/
var DROPBOX_SKIP_FOLDERS = ["Dropbox-shared"];

var MAX_TAG_REPLACEMENTS = 1000;             /* 1フレーム内で同一タグを置換する上限 / replacement guard */
var MAX_FILE_NAME_SERIAL = 1000;             /* ファイル名に付ける連番の上限 / file name serial limit */

// =========================================
// テンプレートの記憶 / Remembered template
// =========================================

/* Illustratorの環境設定に書くため、次回起動時も残る / Stored in Illustrator's preferences, so it survives a restart */
var PREF_KEY_TEMPLATE_PATH = "InvoiceFromClipboard.templatePath";

/* ファイル選択に渡す絞り込み。Windowsは文字列、macOSは判定関数を受け取る / Windows takes a string, macOS a filter function */
var TEMPLATE_FILE_FILTER = ($.os.indexOf("Windows") !== -1) ? "Illustrator:*.ai;*.ait" : function (fileToTest) {
    return (fileToTest instanceof Folder) || /\.(ai|ait)$/i.test(fileToTest.name);
};

/**
 * 記憶しているテンプレートのパスを読み出す
 * @returns {string} 記憶しているパス（未設定なら空文字）
 */
function loadSavedTemplatePath() {
    var savedPath = app.preferences.getStringPreference(PREF_KEY_TEMPLATE_PATH);
    return (savedPath != null) ? String(savedPath) : "";
}

/**
 * テンプレートのパスを記憶する
 * @param {string} templatePath - 記憶するパス
 * @returns {void}
 */
function saveTemplatePath(templatePath) {
    app.preferences.setStringPreference(PREF_KEY_TEMPLATE_PATH, String(templatePath));
}

/**
 * テンプレートのファイル名から書類名を決める
 * 名前に含まれていた語をそのまま使い、どれも含まれていなければ先頭の語を返す
 * @param {File} templateFile - 判定するテンプレート（未指定ならnull）
 * @returns {string} 「領収書」「請求書」など
 */
function resolveDocumentTypeName(templateFile) {
    if (templateFile !== null) {
        var templateName = decodeURI(templateFile.name);
        for (var i = 0; i < DOCUMENT_TYPE_NAMES.length; i++) {
            if (templateName.indexOf(DOCUMENT_TYPE_NAMES[i]) !== -1) return DOCUMENT_TYPE_NAMES[i];
        }
    }
    return DOCUMENT_TYPE_NAMES[0];
}

/**
 * 入力欄に出すラベルを決める（設定が無ければ見出しをそのまま使う）
 * @param {string} headingText - クリップボードの見出し
 * @returns {string} 表示するラベル
 */
function fieldLabelFor(headingText) {
    var overrideLabel = FIELD_LABEL_OVERRIDES[headingText];
    /* Objectが元から持つプロパティを拾わないよう、文字列のときだけ差し替える / Only accept a string */
    return (typeof overrideLabel === "string") ? overrideLabel : headingText;
}

/**
 * `#書類名#` を、テンプレートから決めた書類名に置き換える
 * @param {string} templateText - 置換前の文字列
 * @param {string} documentTypeName - 差し込む書類名
 * @returns {string} 置換後の文字列
 */
function fillDocumentTypeToken(templateText, documentTypeName) {
    /* 全出現を置換（ExtendScript安全）/ Replace every occurrence */
    return String(templateText).split("#" + DOCUMENT_TYPE_TOKEN + "#").join(documentTypeName);
}

/**
 * 使用するテンプレートを決める（記憶しているパスを優先し、無ければ初期パス）
 * どちらも見つからなければnullを返し、ダイアログの［指定］で選ばせる
 * @returns {File} テンプレートファイル（見つからなければnull）
 */
function resolveTemplateFile() {
    var candidatePaths = [loadSavedTemplatePath(), DEFAULT_TEMPLATE_PATH];
    for (var i = 0; i < candidatePaths.length; i++) {
        if (candidatePaths[i] === "") continue;
        var candidateFile = new File(candidatePaths[i]);
        if (candidateFile.exists) return candidateFile;
    }
    return null;
}

// =========================================
// UIレイアウトの共通設定 / Shared UI layout
// =========================================

/* ウィンドウ・パネルの余白と間隔 / Window & panel margins and spacing */
var WINDOW_MARGINS = 16;                 /* ウィンドウ外周の余白 / window margin */
var WINDOW_SPACING = 12;                 /* ウィンドウ内の要素間隔 / window spacing */
var PANEL_MARGINS  = [16, 20, 16, 12];   /* パネル余白 [左,上,右,下] / panel margins */
var PANEL_SPACING  = 12;                 /* パネル内の要素間隔 / panel spacing */
var FIELD_ROW_SPACING = 6;               /* 入力行が続くパネルの間隔 / spacing for stacked field rows */
var BUTTON_BAR_TOP_MARGIN = 10;          /* ボタン領域の上マージン / space above the button row */
var PANEL_BUTTON_TOP_MARGIN = 5;         /* パネル内ボタン行の上マージン / space above a button row inside a panel */

/* コントロールの寸法 / Control sizes */
var FIELD_LABEL_WIDTH = 120;             /* 入力欄ラベルの幅 / field label width */
var AMOUNT_INPUT_SIZE = [110, 25];       /* 金額入力欄 / amount input field */
var READONLY_TEXT_WIDTH = 330;           /* 計算結果・パス・書き出し先の表示幅 / read-only text width */
var INLINE_BUTTON_WIDTH = 70;            /* 行の中に置くボタンの幅 / width of buttons placed inside a row */
var INLINE_SPACER_WIDTH = 12;            /* 行の中で要素を離す固定スペーサーの幅 / fixed spacer width inside a row */
var PATH_DISPLAY_MAX_WIDTH = 56;         /* パス表示の桁数（半角換算）/ path width in half-width units */

/**
 * ウィンドウの共通設定
 * @param {Window} targetWindow - 設定するウィンドウ
 * @param {number} [spacing] - 要素間隔（省略時は WINDOW_SPACING）
 * @returns {void}
 */
function setupWindow(targetWindow, spacing) {
    targetWindow.orientation = "column";
    targetWindow.alignChildren = "fill";
    targetWindow.margins = WINDOW_MARGINS;
    targetWindow.spacing = (typeof spacing === "number") ? spacing : WINDOW_SPACING;
}

/**
 * パネルの共通設定
 * @param {Panel} targetPanel - 設定するパネル
 * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
 * @returns {void}
 */
function setupPanel(targetPanel, spacing) {
    targetPanel.orientation = "column";
    targetPanel.alignChildren = ["fill", "top"];
    targetPanel.alignment = "fill";
    targetPanel.margins = PANEL_MARGINS;
    targetPanel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
}

/**
 * 行グループの共通設定（ボタン列など）
 * @param {Group} targetGroup - 設定するグループ
 * @param {string} [alignment] - 揃え位置（省略時は "left"）
 * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
 * @returns {void}
 */
function setupRow(targetGroup, alignment, spacing) {
    targetGroup.orientation = "row";
    /*
       揃えは横と天地を必ず対で指定する。文字列だけを渡すと天地の指定が外れ、
       行の中で背の低いチェックボックスなどが上端に張り付く。
       Always pass both axes: a bare string drops the vertical one and pins short controls to the top.
    */
    targetGroup.alignment = [alignment || "left", "center"];
    /*
       親パネルの alignChildren（fill）を引き継ぐと、行の中のボタンまで横いっぱいに伸びる。
       行の中身は本来の幅のままにし、伸ばしたい要素だけが個別に alignment を指定する。
       Without this, buttons inherit the panel's fill and stretch across the row.
    */
    targetGroup.alignChildren = ["left", "center"];
    targetGroup.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
}

/**
 * 一定幅の見出しラベルを右揃えで追加し、入力欄の左端を縦に揃える
 * 揃え位置は幅で決まるため、幅を押さえてから右揃えにする
 * @param {Group} parentGroup - ラベルを追加するグループ
 * @param {string} labelString - 表示する見出し
 * @returns {StaticText} 追加したラベル
 */
function addFieldLabel(parentGroup, labelString) {
    var fieldLabel = parentGroup.add("statictext", undefined, labelString);
    fieldLabel.preferredSize.width = FIELD_LABEL_WIDTH;
    fieldLabel.justify = "right";
    return fieldLabel;
}

/**
 * 「ラベル＋読み取り専用テキスト」の行を追加する
 * 値は実行時に入るため、先に幅を押さえておく
 * @param {Panel} parentPanel - 行を追加するパネル
 * @param {string} labelString - 左に置く見出し（空文字なら字下げのみ）
 * @returns {StaticText} 値を入れるテキスト
 */
function addReadOnlyTextRow(parentPanel, labelString) {
    var rowGroup = parentPanel.add("group");
    setupRow(rowGroup, "fill");
    addFieldLabel(rowGroup, labelString);
    var valueText = rowGroup.add("statictext", undefined, "");
    valueText.preferredSize.width = READONLY_TEXT_WIDTH;
    return valueText;
}

// =========================================
// 表示文字列 / UI text
// =========================================

/* ダイアログとメッセージの文言。`#…#` は実行時に値へ置き換える / Dialog and message text */
var LABELS = {
    /* ダイアログ / Dialog */
    dialog: {
        title: "#書類名#作成"
    },
    /* パネル見出し / Panel titles */
    panel: {
        template: "テンプレート",
        parsedFields: "読み取り内容",
        pdfOutput: "書き出し"
    },
    /* フィールド見出し（コロンまで含める）/ Field labels, colon included */
    fieldLabel: {
        documentType: "タイトル：",
        templateFolder: "パス：",
        templateFileName: "ファイル：",
        pdfFileName: "ファイル名：",
        saveFolder: "保存先："
    },
    /* チェックボックス / Checkboxes */
    checkbox: {
        fullPath: "フルパス",
        shortenDropbox: "Dropboxパスを短縮"
    },
    /* 単位・計算結果の表示 / Units and derived values */
    unit: {
        yen: "円"
    },
    breakdown: {
        tax: "税抜 #excluding# 円 ／ 消費税（#rate#％） #tax# 円"
    },
    /* ボタン / Buttons */
    button: {
        pickTemplate: "指定",
        changeTemplate: "変更",
        reloadClipboard: "更新",
        cancel: "キャンセル",
        run: "PDFを作成"
    },
    /* ファイル選択の見出し / File picker prompt */
    prompt: {
        pickTemplate: "#書類名#テンプレートを選択"
    },
    /* ヘルプチップ / Tooltips */
    tooltip: {
        templatePath: "流し込み先のテンプレートです。\n［指定］で選び直すと、次回以降もそのファイルが使われます。",
        fullPath: "パスを省略せずに表示します。\nDropboxパスの短縮中は使えません。",
        shortenDropbox: "Dropboxフォルダーまでのパスを省いて表示します。\nテンプレートと保存先の両方に効きます。",
        amount: "税込金額を入力します。税抜と消費税はここから計算されます。",
        application: "テンプレートの「但」に入れる文面の書き方を選びます。",
        documentType: "書類の種類です。テンプレートの<タイトル>、PDFのファイル名、メールの文面に入ります。\nテンプレートのファイル名から初期選択を決めています。",
        reloadClipboard: "クリップボードを読み直して、各項目を入れ直します。\nダイアログを開いたまま別の内容をコピーして押してください。",
        run: "テンプレートの複製にデータを流し込み、PDFを書き出します。\nテンプレート自体は変更されません。"
    },
    /* 警告・完了メッセージ / Alerts */
    alert: {
        noTemplate: "テンプレートが指定されていません。\n［指定］でテンプレートファイルを選んでください。",
        templateMissing: "記憶していたテンプレートが見つかりません。\n［指定］で選び直してください。\n\n#path#",
        emptyClipboard: "クリップボードが空か、Illustratorに貼り付けられない内容です。",
        noTextInClipboard: "クリップボードにテキストが見つかりませんでした。",
        clipboardError: "クリップボードからの取得に失敗しました：\n",
        noFields: "クリップボードから項目を読み取れませんでした。\n次の見出しのいずれかを含むテキストをコピーしてください。\n\n#headings#",
        workCopyFailed: "作業用の複製ファイルを作成できませんでした。\n\n#detail#",
        exportFailed: "PDFの書き出しに失敗しました。\n\n#detail#",
        done: "PDFを作成しました。\n\n#filename#",
        mailCopied: "返信メールの文面をクリップボードにコピーしました。",
        missingTags: "テンプレートに次のタグが見つかりませんでした。\n#tags#"
    },
    /* 既定の名前 / Fallback names */
    fallbackName: {
        recipient: "宛先なし",
        noTemplate: "（未指定）"
    }
};

// =========================================
// 文字列と数値の整形 / Text and number helpers
// =========================================

/**
 * 文字列先頭のBOM（U+FEFF / 65279）と前後の空白（全角スペースを含む）を除去する
 * @param {string} sourceText - 対象の文字列
 * @returns {string} 整形後の文字列
 */
function trimAndStripBom(sourceText) {
    sourceText = String(sourceText);
    if (sourceText.length && sourceText.charCodeAt(0) === 65279) sourceText = sourceText.substring(1);
    return sourceText.replace(/^[\s　]+|[\s　]+$/g, "");
}

/**
 * 改行を空白に置き換えて1行にまとめる
 * ダイアログの入力欄は1行しか扱えず、領収書の各項目も1行に収める前提のため
 * @param {string} sourceText - 対象の文字列
 * @returns {string} 1行にまとめた文字列
 */
function toSingleLine(sourceText) {
    return trimAndStripBom(String(sourceText).replace(/[\r\n]+/g, " "));
}

/**
 * 全角数字を半角数字に置き換える
 * @param {string} sourceText - 対象の文字列
 * @returns {string} 半角数字に揃えた文字列
 */
function toHalfWidthDigits(sourceText) {
    return String(sourceText).replace(/[０-９]/g, function (fullWidthDigit) {
        return String.fromCharCode(fullWidthDigit.charCodeAt(0) - 0xFEE0);
    });
}

/**
 * 「11,000」「¥11,000」「11000円」などの表記から数値を取り出す
 * @param {string} amountText - 金額の文字列
 * @returns {number} 取り出した数値（読み取れなければ0）
 */
function parseAmount(amountText) {
    var digitsOnly = toHalfWidthDigits(String(amountText)).replace(/[^0-9.\-]/g, "");
    var amountValue = parseFloat(digitsOnly);
    return isNaN(amountValue) ? 0 : amountValue;
}

/**
 * 数値を3桁区切りの文字列にする
 * @param {number} amountValue - 整形する数値
 * @returns {string} 3桁区切りの文字列
 */
function formatAmount(amountValue) {
    var remainingDigits = String(Math.abs(Math.round(amountValue)));
    var groupedDigits = "";
    while (remainingDigits.length > 3) {
        groupedDigits = "," + remainingDigits.substring(remainingDigits.length - 3) + groupedDigits;
        remainingDigits = remainingDigits.substring(0, remainingDigits.length - 3);
    }
    return (amountValue < 0 ? "-" : "") + remainingDigits + groupedDigits;
}

/**
 * 税抜金額の端数を、設定した方法で処理する
 * @param {number} rawValue - 端数を含む数値
 * @returns {number} 端数処理後の整数
 */
function applyTaxFractionMode(rawValue) {
    if (TAX_FRACTION_MODE === "ceil") return Math.ceil(rawValue);
    if (TAX_FRACTION_MODE === "round") return Math.round(rawValue);
    return Math.floor(rawValue);
}

/**
 * 税込金額から税抜金額と消費税額を求める
 * 消費税は「税込－税抜」で出すため、3つを足し引きしても金額がずれない
 * @param {number} amountIncludingTax - 税込金額
 * @returns {{includingTax: number, excludingTax: number, tax: number}} 税込・税抜・消費税
 */
function splitTaxFromTotal(amountIncludingTax) {
    var roundedTotal = Math.round(amountIncludingTax);
    var amountExcludingTax = applyTaxFractionMode(roundedTotal / (1 + TAX_RATE));
    return {
        includingTax: roundedTotal,
        excludingTax: amountExcludingTax,
        tax: roundedTotal - amountExcludingTax
    };
}

/**
 * 1桁の数値を2桁のゼロ埋め文字列にする
 * @param {number} numberValue - 対象の数値
 * @returns {string} 2桁に揃えた文字列
 */
function padTwoDigits(numberValue) {
    return (numberValue < 10 ? "0" : "") + String(numberValue);
}

/**
 * 日付の文字列から年月日を取り出す
 * 「2026/08/14」「2026-08-14」「2026年8月14日」のいずれも解釈する
 * @param {string} dateText - 日付の文字列
 * @returns {{year: number, month: number, day: number}} 年月日（読み取れなければnull）
 */
function parseDateParts(dateText) {
    var dateParts = toHalfWidthDigits(String(dateText)).match(/(\d{4})\D+(\d{1,2})\D+(\d{1,2})/);
    if (!dateParts) return null;
    return {
        year: parseInt(dateParts[1], 10),
        month: parseInt(dateParts[2], 10),
        day: parseInt(dateParts[3], 10)
    };
}

/**
 * 「2026/08/14」形式の日付を「2026年8月14日」に整える
 * 年月日を取り出せない文字列は、そのまま返す
 * @param {string} dateText - 日付の文字列
 * @returns {string} 整形後の日付
 */
function formatDateValue(dateText) {
    var dateParts = parseDateParts(dateText);
    if (dateParts === null) return String(dateText);
    return String(dateParts.year) + "年" + String(dateParts.month) + "月" + String(dateParts.day) + "日";
}

/**
 * 日付をファイル名用の8桁（YYYYMMDD）にする
 * @param {string} dateText - 日付の文字列
 * @returns {string} 8桁の文字列（読み取れなければ空文字）
 */
function formatDateStamp(dateText) {
    var dateParts = parseDateParts(dateText);
    if (dateParts === null) return "";
    return String(dateParts.year) + padTwoDigits(dateParts.month) + padTwoDigits(dateParts.day);
}

/**
 * 実行日を8桁（YYYYMMDD）にする（データの日付が読み取れないときの控え）
 * @returns {string} 8桁の文字列
 */
function todayDateStamp() {
    var currentTime = new Date();
    return String(currentTime.getFullYear()) + padTwoDigits(currentTime.getMonth() + 1) + padTwoDigits(currentTime.getDate());
}

/**
 * ファイル名に使えない文字を下線に置き換える
 * `%` も対象にする。File() はパスをURIとして解釈するため、そのまま渡すと別の文字に化ける
 * @param {string} nameText - 元の文字列
 * @returns {string} ファイル名に使える文字列
 */
function sanitizeFileName(nameText) {
    return trimAndStripBom(String(nameText).replace(/[\\\/:\*\?"<>\|%\r\n\t]/g, "_"));
}

/**
 * ホーム直下から「Dropbox」を含むフォルダーを探す
 * チームフォルダー（「sw Dropbox」など）を優先する。
 * 個人用の「Dropbox」は「~/Dropbox/」として短縮できるため、優先度を下げている
 * @returns {Folder} 見つかったフォルダー（なければnull）
 */
function findDropboxFolder() {
    var homeFolder = Folder("~");
    if (!homeFolder.exists) return null;

    var entryList = homeFolder.getFiles();
    var personalFolder = null;
    var teamFolder = null;

    for (var i = 0; i < entryList.length; i++) {
        if (!(entryList[i] instanceof Folder)) continue;

        var entryName = decodeURI(String(entryList[i].name));
        if (entryName.charAt(0) === ".") continue;
        if (entryName.indexOf("Dropbox") === -1) continue;

        if (entryName === "Dropbox") {
            if (!personalFolder) personalFolder = entryList[i];
        } else if (!teamFolder) {
            teamFolder = entryList[i];
        }
    }
    return teamFolder ? teamFolder : personalFolder;
}

/**
 * フォルダー直下に表示用のサブフォルダーが1つだけあるとき、そのフォルダーを返す
 * チームDropboxのメンバーフォルダー（「takano masahiro」など）の判定に使う
 * @param {Folder} parentFolder - 探索するフォルダー
 * @returns {Folder} 唯一のサブフォルダー（0個または2個以上のときはnull）
 */
function findSingleSubFolder(parentFolder) {
    var entryList = parentFolder.getFiles();
    var foundFolder = null;

    for (var i = 0; i < entryList.length; i++) {
        if (!(entryList[i] instanceof Folder)) continue;

        var entryName = decodeURI(String(entryList[i].name));
        if (entryName.charAt(0) === ".") continue;

        if (foundFolder) return null;
        foundFolder = entryList[i];
    }
    return foundFolder;
}

/**
 * Dropboxのローカルマウントパスを決める
 * 手動指定が空のときは、ホーム直下の「Dropbox」を含むフォルダーを探し、
 * その中にメンバーフォルダーが1つだけあれば、そこまでをプレフィックスとする
 * @param {string} manualPath - 手動で指定するパス（空文字なら自動検出）
 * @returns {string} 末尾に「/」を付けたプレフィックス（見つからない場合は空文字）
 */
function resolveDropboxPrefix(manualPath) {
    if (manualPath) {
        return (manualPath.charAt(manualPath.length - 1) === "/") ? manualPath : manualPath + "/";
    }
    var dropboxFolder = findDropboxFolder();
    if (!dropboxFolder) return "";

    var memberFolder = findSingleSubFolder(dropboxFolder);
    return (memberFolder ? memberFolder : dropboxFolder).fsName + "/";
}

/* Dropboxのローカルマウントパス（起動時に1度だけ判定）/ Local Dropbox mount path, resolved once */
var DROPBOX_PREFIX = "";
try {
    DROPBOX_PREFIX = resolveDropboxPrefix(DROPBOX_MOUNT_PATH);
} catch (e) {
    // ホームを読めない環境では短縮せず、フルパスのまま表示する / fall back to full paths
}

/**
 * ホームフォルダーを「~」に置き換える
 * @param {string} pathText - 元のパス
 * @returns {string} 置き換えたパス（ホーム配下でなければ元のまま）
 */
function toTildePath(pathText) {
    var homePath = Folder("~").fsName;
    if (homePath === "") return pathText;
    if (pathText === homePath) return "~";
    if (pathText.indexOf(homePath + "/") === 0) return "~" + pathText.substring(homePath.length);
    return pathText;
}

/**
 * Dropboxのマウントパスからの相対パスから、先頭の共有フォルダー名を省く
 * @param {string} relativePath - マウントパスからの相対パス
 * @returns {string} 共有フォルダー名を省いたパス（該当しなければ元のまま）
 */
function stripSkippedFolders(relativePath) {
    for (var i = 0; i < DROPBOX_SKIP_FOLDERS.length; i++) {
        var skipPrefix = DROPBOX_SKIP_FOLDERS[i] + "/";
        if (relativePath.indexOf(skipPrefix) === 0) return relativePath.substring(skipPrefix.length);
    }
    return relativePath;
}

/**
 * パスを表示用に整える
 * Dropbox短縮が有効でDropbox配下なら、マウントパスと共有フォルダーまでを省く。
 * そうでなければ、「~」に畳むかフルパスのまま返す
 * @param {string} pathText - 元のパス
 * @param {boolean} useTilde - ホームを「~」に置き換えるならtrue
 * @param {boolean} useDropbox - Dropboxのマウントパスを省くならtrue
 * @returns {string} 表示用のパス
 */
function formatDisplayPath(pathText, useTilde, useDropbox) {
    pathText = String(pathText);
    if (useDropbox && DROPBOX_PREFIX !== "" && pathText.indexOf(DROPBOX_PREFIX) === 0) {
        return stripSkippedFolders(pathText.substring(DROPBOX_PREFIX.length));
    }
    if (useTilde) return toTildePath(pathText);
    return pathText;
}

/**
 * 表示上の桁数を数える（全角は2桁ぶんとして数える）
 * ScriptUIの文字幅は取得できないため、半角換算で見積もる
 * @param {string} sourceText - 対象の文字列
 * @returns {number} 半角に換算した桁数
 */
function displayWidthOf(sourceText) {
    sourceText = String(sourceText);
    var totalWidth = 0;
    for (var i = 0; i < sourceText.length; i++) {
        totalWidth += (sourceText.charCodeAt(i) > 0x2E7F) ? 2 : 1;
    }
    return totalWidth;
}

/**
 * 表示用の文字列を、決め打ちの幅に収まるよう先頭側から省く
 * ScriptUIの文字は作成後に伸びないため、はみ出す分は詰める。省略しても分かるよう、
 * ファイル名側の末尾を残す。全体は helpTip で確認できる。
 * @param {string} displayText - 表示する文字列
 * @param {number} maxWidth - 半角換算の最大桁数
 * @returns {string} 幅に収めた文字列
 */
function shortenToWidth(displayText, maxWidth) {
    displayText = String(displayText);
    if (displayWidthOf(displayText) <= maxWidth) return displayText;

    /* フォルダー名が途中で切れると読めないので、区切りの位置から残す / Cut at a separator, never mid-segment */
    var separatorIndex = displayText.indexOf("/");
    while (separatorIndex !== -1) {
        var tailText = displayText.substring(separatorIndex);
        if (displayWidthOf(tailText) + 2 <= maxWidth) return "…" + tailText;
        separatorIndex = displayText.indexOf("/", separatorIndex + 1);
    }

    /* 最後の区切りでも収まらないときは、文字単位で詰める / Fall back to a character cut */
    var startIndex = 0;
    while (startIndex < displayText.length && displayWidthOf(displayText.substring(startIndex)) + 2 > maxWidth) {
        startIndex++;
    }
    return "…" + displayText.substring(startIndex);
}

// =========================================
// メイン処理 / Main
// =========================================

(function () {

    /* 記憶しているテンプレート。ダイアログの［指定］／［変更］で差し替わる / Swapped by the Choose / Change button */
    var templateFile = resolveTemplateFile();

    /* 記憶はあるのにファイルが無い場合は、選び直しが必要だと知らせる / Tell the user when the remembered file is gone */
    var savedTemplatePath = loadSavedTemplatePath();
    if (templateFile === null && savedTemplatePath !== "") {
        alert(LABELS.alert.templateMissing.replace("#path#", savedTemplatePath));
    }

    // =========================================
    // クリップボードの読み取り / Clipboard access
    // =========================================

    /**
     * ドキュメント内のオブジェクトをすべて削除する
     * @param {Document} targetDocument - 対象のドキュメント
     * @returns {void}
     */
    function removeAllPageItems(targetDocument) {
        for (var i = targetDocument.pageItems.length - 1; i >= 0; i--) {
            targetDocument.pageItems[i].remove();
        }
    }

    /**
     * ドキュメント内のテキストフレームの内容を、改行でつないで取り出す
     * textFrames はグループの中まで含むので、貼り付いた形にかかわらず拾える
     * @param {Document} targetDocument - 対象のドキュメント
     * @returns {string} つないだ文字列（テキストが無ければ空文字）
     */
    function joinTextFrameContents(targetDocument) {
        var contentsList = [];
        for (var i = 0; i < targetDocument.textFrames.length; i++) {
            contentsList.push(String(targetDocument.textFrames[i].contents));
        }
        return contentsList.join("\n");
    }

    /**
     * 使い捨てのドキュメントへ貼り付けて、クリップボードのテキストを読み取る
     * Illustratorには文字列を直接読み出すAPIが無いためペーストを使うが、開いているドキュメントを
     * 汚さないよう、読み取り用の新規ドキュメントを立ててすぐ閉じる。
     * また、Illustratorは自分がコピーした内容を内部に保持していて、他アプリがクリップボードを
     * 書き換えたあとの1回目のペーストでは古い内容が貼り付く。その1回目が内部の更新を促すため、
     * 1回目の結果は消してから2回目を貼り付ける。
     * 読めなかった理由は呼び出し元に返し、知らせるかどうかは呼び出し元に任せる
     * @returns {{text: string, error: string}} 読み取った文字列と、読めなかったときのメッセージ
     */
    function readClipboardText() {
        var scratchDocument = null;
        var pastedItemCount = 0;
        var clipboardText = "";
        var pasteError = null;

        try {
            scratchDocument = app.documents.add();

            /* 1回目は内部クリップボードを最新にするためだけのペースト / The first paste only refreshes the cached clipboard */
            app.paste();
            app.redraw();
            removeAllPageItems(scratchDocument);

            app.paste();
            /* 貼り付け直後は反映が遅れることがあるため、描画を確定させてから読む / Flush the paste before reading */
            app.redraw();
            pastedItemCount = scratchDocument.pageItems.length;
            clipboardText = joinTextFrameContents(scratchDocument);
        } catch (e) {
            pasteError = String(e);
        } finally {
            /* 成否にかかわらず、読み取り用のドキュメントは残さない / Never leave the scratch document behind */
            if (scratchDocument) scratchDocument.close(SaveOptions.DONOTSAVECHANGES);
        }

        if (pasteError !== null) return { text: null, error: LABELS.alert.clipboardError + pasteError };
        /* ペースト自体が起きなかった場合と、貼り付いたがテキストが無い場合を区別する / Tell the two failures apart */
        if (pastedItemCount === 0) return { text: null, error: LABELS.alert.emptyClipboard };
        if (clipboardText === "") return { text: null, error: LABELS.alert.noTextInClipboard };
        return { text: clipboardText, error: null };
    }

    // =========================================
    // クリップボードの解析 / Clipboard parsing
    // =========================================

    /**
     * 配列に、まだ入っていない見出しだけを追加する
     * @param {string[]} headingList - 追加先の配列
     * @param {string[]} headingsToAdd - 追加する見出し
     * @returns {void}
     */
    function appendNewHeadings(headingList, headingsToAdd) {
        for (var i = 0; i < headingsToAdd.length; i++) {
            var isAlreadyListed = false;
            for (var j = 0; j < headingList.length; j++) {
                if (headingList[j] === headingsToAdd[i]) { isAlreadyListed = true; break; }
            }
            if (!isAlreadyListed) headingList.push(headingsToAdd[i]);
        }
    }

    /**
     * FIELD_MAPPINGS に現れる見出しを、重複を除いて定義順に取り出す
     * @returns {string[]} 見出しの配列
     */
    function collectMappedHeadings() {
        var mappedHeadings = [];
        for (var i = 0; i < FIELD_MAPPINGS.length; i++) {
            /* 見出しを持たない項目（書類名など）は入力欄を作らない / Entries without a heading get no input row */
            if (FIELD_MAPPINGS[i].heading === "") continue;
            appendNewHeadings(mappedHeadings, [FIELD_MAPPINGS[i].heading]);
        }
        return mappedHeadings;
    }

    /**
     * 返信メールの件名と本文が使っている見出しを、`#…#` から拾う
     * 見出しを別に列挙せずに済むので、文面を書き替えるだけで読み取り対象が増える
     * @returns {string[]} 見出しの配列
     */
    function collectReplyMailHeadings() {
        var mailHeadings = [];
        var tokenPattern = /#([^#\r\n]+)#/g;
        var searchText = REPLY_MAIL_SUBJECT + "\n" + REPLY_MAIL_TEMPLATE;
        var tokenMatch = tokenPattern.exec(searchText);
        while (tokenMatch !== null) {
            var foundToken = tokenMatch[1];
            /* 差し込み用のトークンは見出しではないので、読み取り対象から外す / These are not headings */
            if (foundToken !== REPLY_MAIL_FILE_NAME_TOKEN && foundToken !== DOCUMENT_TYPE_TOKEN) {
                appendNewHeadings(mailHeadings, [foundToken]);
            }
            tokenMatch = tokenPattern.exec(searchText);
        }
        return mailHeadings;
    }

    /**
     * ダイアログに入力欄を出す見出しを集める
     * タグに使う見出し、メールの件名・本文で使う見出し、メールの宛先の順に並べる
     * @returns {string[]} 見出しの配列
     */
    function collectInputHeadings() {
        var collectedHeadings = collectMappedHeadings();
        appendNewHeadings(collectedHeadings, collectReplyMailHeadings());
        /* 宛先は送る前に確かめたいので、文面で使っていなくても入力欄を出す / Always show the To address */
        if (MAIL_ADDRESS_HEADING !== "") appendNewHeadings(collectedHeadings, [MAIL_ADDRESS_HEADING]);
        return collectedHeadings;
    }

    /**
     * 値の区切りとして扱う見出しをすべて集める（入力欄の見出し＋読み飛ばす見出し）
     * @returns {string[]} 見出しの配列
     */
    function collectKnownHeadings() {
        var collectedHeadings = collectInputHeadings();
        appendNewHeadings(collectedHeadings, IGNORED_HEADINGS);
        return collectedHeadings;
    }

    /**
     * 行が見出しのいずれかと一致するか判定する
     * @param {string} lineText - 判定する行
     * @param {string[]} knownHeadings - 見出しとして扱う文字列の配列
     * @returns {boolean} 一致すればtrue
     */
    function isKnownHeading(lineText, knownHeadings) {
        for (var i = 0; i < knownHeadings.length; i++) {
            if (knownHeadings[i] === lineText) return true;
        }
        return false;
    }

    /**
     * 「見出し行＋値行」を空行で区切った形式のテキストを解析する
     * 空行と既知の見出しの両方を値の終わりとして扱うため、
     * 対応表に無い項目が挟まっていても、その値を隣の項目に取り込んでしまわない
     * @param {string} clipboardText - クリップボードから読み取った文字列
     * @param {string[]} knownHeadings - 見出しとして扱う文字列の配列
     * @returns {Object} 見出しをキー、値を文字列とするオブジェクト
     */
    function parseHeadingValuePairs(clipboardText, knownHeadings) {
        var parsedValues = {};
        var textLines = String(clipboardText).split(/\r\n|\r|\n/);
        var currentHeading = null;
        var valueLines = [];
        var isAwaitingHeading = true; // 空行の直後と文頭では見出しを待つ / a heading is expected here

        /**
         * 読みかけの項目を確定する
         * @returns {void}
         */
        function closeCurrentField() {
            if (currentHeading !== null) parsedValues[currentHeading] = valueLines.join("\n");
            currentHeading = null;
            valueLines = [];
        }

        for (var i = 0; i < textLines.length; i++) {
            var lineText = trimAndStripBom(textLines[i]);

            if (lineText === "") {
                closeCurrentField();
                isAwaitingHeading = true;
                continue;
            }
            if (isKnownHeading(lineText, knownHeadings)) {
                closeCurrentField();
                currentHeading = lineText;
                isAwaitingHeading = false;
                continue;
            }
            /* 見出しを待っている位置に現れた未知の行は、見出しとみなして次の区切りまで読み飛ばす / Skip an unknown block */
            if (isAwaitingHeading) {
                isAwaitingHeading = false;
                continue;
            }
            if (currentHeading !== null) valueLines.push(lineText);
        }
        /* 末尾に空行が無くても最後の項目を確定させる / Close the last field even without a trailing blank line */
        closeCurrentField();
        return parsedValues;
    }

    /**
     * 見出しに対応する除去パターンを値に適用する
     * @param {string} headingText - 値の見出し
     * @param {string} valueText - 元の値
     * @returns {string} 不要な文字を取り除いた値
     */
    function applyValueRemovals(headingText, valueText) {
        var cleanedValue = String(valueText);
        for (var i = 0; i < VALUE_REMOVAL_PATTERNS.length; i++) {
            if (VALUE_REMOVAL_PATTERNS[i].heading !== headingText) continue;
            cleanedValue = cleanedValue.replace(VALUE_REMOVAL_PATTERNS[i].pattern, "");
        }
        return trimAndStripBom(cleanedValue);
    }

    /**
     * 解析結果から見出しの値を1行で取り出す（不要な文字は取り除く）
     * ダイアログに入る前に整えるため、表示された値がそのまま流し込まれる
     * @param {Object} parsedValues - 解析結果
     * @param {string} headingText - 取り出す見出し
     * @returns {string} 見出しの値（無ければ空文字）
     */
    function readFieldValue(parsedValues, headingText) {
        if (parsedValues[headingText] == null) return "";
        return applyValueRemovals(headingText, toSingleLine(parsedValues[headingText]));
    }

    /**
     * 見出しがひとつでも読み取れたか判定する
     * @param {Object} parsedValues - 解析結果
     * @param {string[]} headingsToCheck - 確認する見出しの配列
     * @returns {boolean} ひとつでも読み取れていればtrue
     */
    function hasAnyMappedField(parsedValues, headingsToCheck) {
        for (var i = 0; i < headingsToCheck.length; i++) {
            if (parsedValues[headingsToCheck[i]] != null) return true;
        }
        return false;
    }

    var mappedHeadings = collectMappedHeadings();
    var inputHeadings = collectInputHeadings();

    /**
     * クリップボードを読み取って解析する
     * 対応表の見出しがひとつも無い場合も「読み取れなかった」として扱う
     * @returns {{values: Object, error: string}} 解析結果と、読み取れなかったときのメッセージ
     */
    function readParsedClipboardValues() {
        var clipboardResult = readClipboardText();
        if (clipboardResult.text === null) return { values: {}, error: clipboardResult.error };

        var readValues = parseHeadingValuePairs(clipboardResult.text, collectKnownHeadings());
        if (!hasAnyMappedField(readValues, mappedHeadings)) {
            return { values: {}, error: LABELS.alert.noFields.replace("#headings#", mappedHeadings.join("\n")) };
        }
        return { values: readValues, error: null };
    }

    /*
       読み取れなくてもダイアログは開き、起動時は知らせない。
       空欄のダイアログがそのまま状態を表すうえ、手で入力するか、
       コピーし直して［更新］を押せば続けられるため。
       Open the dialog even when nothing was read, and stay quiet about it on startup.
    */
    var parsedValues = readParsedClipboardValues().values;

    // =========================================
    // ダイアログUIの構築 / Build the dialog UI
    // =========================================

    /**
     * 指定した種類の値を入れる見出しを、対応表から探す
     * @param {string} valueType - 探す valueType
     * @returns {string} 見出しの文字列（対応表に無ければ空文字）
     */
    function findHeadingByValueType(valueType) {
        for (var i = 0; i < FIELD_MAPPINGS.length; i++) {
            if (FIELD_MAPPINGS[i].valueType === valueType) return FIELD_MAPPINGS[i].heading;
        }
        return "";
    }
    var amountHeading = findHeadingByValueType("amountIncludingTax");
    var dateHeading = findHeadingByValueType("date");
    var applicationHeading = findHeadingByValueType("application");

    var mainDialog = new Window("dialog", "");
    setupWindow(mainDialog);

    /* テンプレートパネル：記憶しているファイルと、選び直すボタン / Template panel: the remembered file and its picker */
    var templatePanel = mainDialog.add("panel", undefined, LABELS.panel.template);
    setupPanel(templatePanel, FIELD_ROW_SPACING);

    /* ファイル名とフォルダーは行を分ける。ファイル名が長いパスに押されて切れないようにする */
    var templateFileNameText = addReadOnlyTextRow(templatePanel, LABELS.fieldLabel.templateFileName);
    templateFileNameText.helpTip = LABELS.tooltip.templatePath;
    var templateFolderText = addReadOnlyTextRow(templatePanel, LABELS.fieldLabel.templateFolder);
    templateFolderText.helpTip = LABELS.tooltip.templatePath;

    /*
       オプション行：パス表示の切り替え2つと、ファイル選択のボタンを右揃えで並べる。
       行そのものを右寄せにするので、余った幅は左側にまとまり、要素同士は離れない。
       The row itself is right-aligned, so the leftover width collects on the left.
    */
    var templateOptionGroup = templatePanel.add("group");
    setupRow(templateOptionGroup, "right");
    templateOptionGroup.margins = [0, PANEL_BUTTON_TOP_MARGIN, 0, 0];

    /*
       チェックボックスは専用のグループで囲む。行に直接置くと、余った幅が
       チェックボックス同士の間にも配分されて離れてしまう。
       Wrapping them keeps any leftover width out from between the two boxes.
    */
    var templateToggleGroup = templateOptionGroup.add("group");
    setupRow(templateToggleGroup, "center");

    var fullPathCheckbox = templateToggleGroup.add("checkbox", undefined, LABELS.checkbox.fullPath);
    fullPathCheckbox.helpTip = LABELS.tooltip.fullPath;

    var dropboxCheckbox;
    if (DROPBOX_PREFIX !== "") {
        dropboxCheckbox = templateToggleGroup.add("checkbox", undefined, LABELS.checkbox.shortenDropbox);
        dropboxCheckbox.value = true;
        dropboxCheckbox.helpTip = LABELS.tooltip.shortenDropbox;
    } else {
        /* Dropboxが見つからない環境ではチェックボックスを出さず、常にオフとして扱う / Stand-in when there is no Dropbox */
        dropboxCheckbox = { value: false, enabled: false };
    }

    /*
       固定幅のスペーサー。伸縮させないよう、上限と下限も同じ幅で押さえる。
       A fixed spacer: pin the minimum and maximum too so it never stretches or collapses.
    */
    var templateOptionSpacer = templateOptionGroup.add("group");
    templateOptionSpacer.minimumSize.width = INLINE_SPACER_WIDTH;
    templateOptionSpacer.preferredSize.width = INLINE_SPACER_WIDTH;
    templateOptionSpacer.maximumSize.width = INLINE_SPACER_WIDTH;

    /* ラベルが［指定］と［変更］で入れ替わるので先に幅を押さえる / The label swaps at runtime */
    var templatePickButton = templateOptionGroup.add("button", undefined, LABELS.button.pickTemplate);
    templatePickButton.preferredSize.width = INLINE_BUTTON_WIDTH;
    templatePickButton.helpTip = LABELS.tooltip.templatePath;

    /* 読み取り内容パネル：入力欄の見出しぶんだけ入力欄を作る / One input per heading shown in the dialog */
    var parsedFieldsPanel = mainDialog.add("panel", undefined, LABELS.panel.parsedFields);
    setupPanel(parsedFieldsPanel, FIELD_ROW_SPACING);

    /*
       書類名はクリップボードには無いので、一番上で選ばせる。
       初期選択はテンプレートのファイル名から決め、［変更］で選び直したときも付け直す。
       The document type is chosen here; the template's file name sets the initial pick.
    */
    var documentTypeRadioGroup = parsedFieldsPanel.add("group");
    setupRow(documentTypeRadioGroup, "fill");
    addFieldLabel(documentTypeRadioGroup, LABELS.fieldLabel.documentType);
    var documentTypeRadios = [];
    for (var typeIndex = 0; typeIndex < DOCUMENT_TYPE_NAMES.length; typeIndex++) {
        var documentTypeRadio = documentTypeRadioGroup.add("radiobutton", undefined, DOCUMENT_TYPE_NAMES[typeIndex]);
        documentTypeRadio.helpTip = LABELS.tooltip.documentType;
        documentTypeRadios.push(documentTypeRadio);
    }

    var fieldInputs = {};        // 見出し → 入力欄 / heading to input field
    var amountInput = null;
    var taxBreakdownText = null;
    var applicationRadios = [];  // 適用の書き方を選ぶラジオ / wording choices for the 適用 line

    for (var headingIndex = 0; headingIndex < inputHeadings.length; headingIndex++) {
        var headingText = inputHeadings[headingIndex];

        var fieldRowGroup = parsedFieldsPanel.add("group");
        setupRow(fieldRowGroup, "fill");
        addFieldLabel(fieldRowGroup, fieldLabelFor(headingText) + "：");

        var fieldInput = fieldRowGroup.add("edittext", undefined, readFieldValue(parsedValues, headingText));
        if (headingText === amountHeading) {
            fieldInput.size = AMOUNT_INPUT_SIZE;
            fieldInput.helpTip = LABELS.tooltip.amount;
            fieldRowGroup.add("statictext", undefined, LABELS.unit.yen);
            amountInput = fieldInput;
            /* 税抜と消費税は入力ではなく計算結果なので、見出しを空にして金額欄のすぐ下に置く / Derived values */
            taxBreakdownText = addReadOnlyTextRow(parsedFieldsPanel, "");
        } else {
            /* ダイアログを広げず、余った幅いっぱいまで伸ばす / Fill the remaining width instead of widening the dialog */
            fieldInput.alignment = ["fill", "center"];
        }
        fieldInputs[headingText] = fieldInput;

        /* 適用の書き方はイベント名のすぐ下で選ばせる / The wording choice sits under the event name */
        if (headingText === applicationHeading) {
            var applicationRadioGroup = parsedFieldsPanel.add("group");
            setupRow(applicationRadioGroup, "fill");
            addFieldLabel(applicationRadioGroup, ""); // 入力欄と左端を揃える / line up with the input column
            for (var presetIndex = 0; presetIndex < APPLICATION_PRESETS.length; presetIndex++) {
                var applicationRadio = applicationRadioGroup.add("radiobutton", undefined, APPLICATION_PRESETS[presetIndex].label);
                applicationRadio.helpTip = LABELS.tooltip.application;
                applicationRadios.push(applicationRadio);
            }
            applicationRadios[0].value = true; // 先頭を既定にする / the first entry is the default
        }
    }

    /* クリップボードを差し替えてから読み直せるよう、パネルの最下部・右端に置く / Re-read the clipboard without reopening */
    var reloadButtonGroup = parsedFieldsPanel.add("group");
    setupRow(reloadButtonGroup, "right");
    reloadButtonGroup.margins = [0, PANEL_BUTTON_TOP_MARGIN, 0, 0];
    var reloadClipboardButton = reloadButtonGroup.add("button", undefined, LABELS.button.reloadClipboard);
    reloadClipboardButton.preferredSize.width = INLINE_BUTTON_WIDTH;
    reloadClipboardButton.helpTip = LABELS.tooltip.reloadClipboard;

    /* 書き出しパネル：作成されるPDFの名前と保存先 / Export panel: the resulting PDF name and folder */
    var pdfOutputPanel = mainDialog.add("panel", undefined, LABELS.panel.pdfOutput);
    setupPanel(pdfOutputPanel, FIELD_ROW_SPACING);
    var pdfFileNameText = addReadOnlyTextRow(pdfOutputPanel, LABELS.fieldLabel.pdfFileName);
    var saveFolderText = addReadOnlyTextRow(pdfOutputPanel, LABELS.fieldLabel.saveFolder);

    var buttonBarGroup = mainDialog.add("group");
    setupRow(buttonBarGroup, "right");
    /* パネルとの間隔を、ウィンドウ既定の行間よりも広げる / Widen the gap from the panel above */
    buttonBarGroup.margins = [0, BUTTON_BAR_TOP_MARGIN, 0, 0];
    buttonBarGroup.add("button", undefined, LABELS.button.cancel, { name: "cancel" });
    var runButton = buttonBarGroup.add("button", undefined, LABELS.button.run, { name: "ok" });
    runButton.helpTip = LABELS.tooltip.run;

    // =========================================
    // 書き出し先の決定 / Resolving the output file
    // =========================================

    /**
     * 選ばれている書類名を返す
     * @returns {string} 「領収書」「請求書」など
     */
    function currentDocumentTypeName() {
        for (var i = 0; i < documentTypeRadios.length; i++) {
            if (documentTypeRadios[i].value) return DOCUMENT_TYPE_NAMES[i];
        }
        return DOCUMENT_TYPE_NAMES[0];
    }

    /**
     * 書類名のラジオを、指定した名前に合わせる
     * @param {string} documentTypeName - 選択する書類名
     * @returns {void}
     */
    function selectDocumentTypeRadio(documentTypeName) {
        for (var i = 0; i < documentTypeRadios.length; i++) {
            documentTypeRadios[i].value = (DOCUMENT_TYPE_NAMES[i] === documentTypeName);
        }
    }

    /**
     * ダイアログの入力から、PDFファイル名の元になる宛先を取り出す
     * @returns {string} ファイル名に使う宛先（空なら既定の名前）
     */
    function currentRecipientName() {
        var recipientInput = fieldInputs[PDF_FILE_NAME_HEADING];
        var recipientName = recipientInput ? sanitizeFileName(recipientInput.text) : "";
        return (recipientName !== "") ? recipientName : LABELS.fallbackName.recipient;
    }

    /**
     * ファイル名に付ける8桁の日付を決める
     * 領収書の日付を使うので、同じ内容から作り直しても同じ名前になる
     * @returns {string} 8桁の日付（日付欄が読めないときは実行日）
     */
    function currentDateStamp() {
        var dateInput = (dateHeading !== "") ? fieldInputs[dateHeading] : null;
        var dateStamp = dateInput ? formatDateStamp(dateInput.text) : "";
        return (dateStamp !== "") ? dateStamp : todayDateStamp();
    }

    /**
     * PDFファイル名の拡張子より前を組み立てる
     * 発行元が空文字のときは、その要素ごと省いて区切りも重ならないようにする
     * @returns {string} 「書類名-発行元-宛先-日付」の文字列
     */
    function buildPdfBaseName() {
        var nameParts = [currentDocumentTypeName()];
        if (PDF_FILE_NAME_ISSUER !== "") nameParts.push(sanitizeFileName(PDF_FILE_NAME_ISSUER));
        /* 日付を宛先より前に置く。名前順に並べたときに発行日順で揃う / Date first, so the folder sorts chronologically */
        nameParts.push(currentDateStamp());
        nameParts.push(currentRecipientName());
        return nameParts.join(PDF_FILE_NAME_SEPARATOR);
    }

    /**
     * 書き出すPDFファイルを決める（テンプレートと同じフォルダーへ）
     * 上書きしない設定のときは、既存ファイルを避けて連番を付ける
     * @returns {File} 書き出し先のファイル（テンプレート未指定ならnull）
     */
    function resolvePdfFile() {
        if (templateFile === null) return null;

        var baseName = buildPdfBaseName();
        var outputFolder = templateFile.parent;
        var pdfFile = new File(outputFolder.fsName + "/" + baseName + ".pdf");
        if (OVERWRITE_EXISTING_PDF) return pdfFile;

        var serialNumber = 2;
        while (pdfFile.exists && serialNumber < MAX_FILE_NAME_SERIAL) {
            pdfFile = new File(outputFolder.fsName + "/" + baseName + "-" + serialNumber + ".pdf");
            serialNumber++;
        }
        return pdfFile;
    }

    // =========================================
    // 表示の更新 / Dialog refresh
    // =========================================

    /**
     * 入力中の金額から税抜・消費税の表示を更新する
     * @returns {void}
     */
    function refreshTaxBreakdown() {
        var taxAmounts = splitTaxFromTotal(amountInput ? parseAmount(amountInput.text) : 0);
        taxBreakdownText.text = LABELS.breakdown.tax
            .replace("#excluding#", formatAmount(taxAmounts.excludingTax))
            .replace("#rate#", String(Math.round(TAX_RATE * 1000) / 10))
            .replace("#tax#", formatAmount(taxAmounts.tax));
    }

    /**
     * チェックボックスの状態に合わせて、フォルダーのパスを表示用の文字列にする
     * @param {string} folderPath - 表示するフォルダーのパス
     * @returns {string} 表示用の文字列
     */
    function folderPathForDisplay(folderPath) {
        var displayPath = formatDisplayPath(folderPath, !fullPathCheckbox.value, dropboxCheckbox.value);
        return shortenToWidth(displayPath, PATH_DISPLAY_MAX_WIDTH);
    }

    /**
     * テンプレートのパス・ファイル名の表示と、ボタン名を更新する
     * 未指定なら［指定］、指定済みなら［変更］になる
     * @returns {void}
     */
    function refreshTemplateRow() {
        var isTemplateSet = (templateFile !== null);
        /* fsName は環境依存のパスそのもので、URIエンコードされていないため decodeURI は通さない / fsName is a raw path */
        var folderPath = isTemplateSet ? templateFile.parent.fsName : "";

        /* 書類名はテンプレート次第で変わるので、タイトルも選び直すたびに付け直す / Retitle on every pick */
        mainDialog.text = fillDocumentTypeToken(LABELS.dialog.title, currentDocumentTypeName()) +
            " " + SCRIPT_VERSION;

        templateFolderText.text = isTemplateSet ? folderPathForDisplay(folderPath) : "";
        templateFolderText.helpTip = isTemplateSet ? folderPath : LABELS.tooltip.templatePath;
        templateFileNameText.text = isTemplateSet
            ? shortenToWidth(decodeURI(templateFile.name), PATH_DISPLAY_MAX_WIDTH)
            : LABELS.fallbackName.noTemplate;
        templateFileNameText.helpTip = isTemplateSet ? templateFile.fsName : LABELS.tooltip.templatePath;

        templatePickButton.text = isTemplateSet ? LABELS.button.changeTemplate : LABELS.button.pickTemplate;
        /* テンプレートが無いと書き出せないので、実行できないことを見て分かるようにする / Nothing to export without a template */
        runButton.enabled = isTemplateSet;
    }

    /**
     * 入力中の宛先・日付・テンプレートから、書き出し先の表示を更新する
     * @returns {void}
     */
    function refreshExportRows() {
        var pdfFile = resolvePdfFile();
        pdfFileNameText.text = (pdfFile !== null) ? decodeURI(pdfFile.name) : "";
        saveFolderText.text = (templateFile !== null) ? folderPathForDisplay(templateFile.parent.fsName) : "";
        saveFolderText.helpTip = (templateFile !== null) ? templateFile.parent.fsName : "";
    }

    /**
     * テンプレートと書き出しの行をまとめて更新する（タイトルとファイル名も付け直す）
     * @returns {void}
     */
    function refreshTemplateAndExportRows() {
        refreshTemplateRow();
        refreshExportRows();
    }

    fullPathCheckbox.onClick = refreshTemplateAndExportRows;
    dropboxCheckbox.onClick = function () {
        /* Dropbox短縮が効いている間はフルパス表示が反映されないので、操作できないようにする */
        fullPathCheckbox.enabled = !dropboxCheckbox.value;
        refreshTemplateAndExportRows();
    };
    fullPathCheckbox.enabled = !dropboxCheckbox.value;

    /**
     * クリップボードを読み直し、各入力欄を入れ直す
     * 読めなかったときと項目が拾えなかったときは、今の入力を残したまま知らせるだけにする
     * @returns {void}
     */
    function reloadFromClipboard() {
        var reloadResult = readParsedClipboardValues();
        if (reloadResult.error !== null) {
            alert(reloadResult.error);
            return; // 入力済みの内容を消さない / keep whatever is already typed in
        }

        parsedValues = reloadResult.values;
        for (var i = 0; i < inputHeadings.length; i++) {
            fieldInputs[inputHeadings[i]].text = readFieldValue(parsedValues, inputHeadings[i]);
        }
        refreshTaxBreakdown();
        refreshExportRows();
    }

    /* ［更新］：ダイアログを開いたまま、コピーし直した内容を取り込む / Re-read the clipboard in place */
    reloadClipboardButton.onClick = reloadFromClipboard;

    /* ［指定］／［変更］：テンプレートを選び直し、次回以降のために記憶する / Pick a template and remember it */
    templatePickButton.onClick = function () {
        var promptText = fillDocumentTypeToken(LABELS.prompt.pickTemplate, currentDocumentTypeName());
        var pickedFile = File.openDialog(promptText, TEMPLATE_FILE_FILTER);
        if (!pickedFile) return;
        templateFile = pickedFile;
        saveTemplatePath(pickedFile.fsName);
        /* 新しいテンプレートの名前から、書類名を選び直す / Re-pick the document type from the new name */
        selectDocumentTypeRadio(resolveDocumentTypeName(templateFile));
        refreshTemplateAndExportRows();
    };

    if (amountInput) {
        amountInput.onChanging = refreshTaxBreakdown;
    }
    /* 宛先と日付はどちらもファイル名に出るため、打つそばから書き出し名を追従させる / Both feed the file name */
    if (fieldInputs[PDF_FILE_NAME_HEADING]) {
        fieldInputs[PDF_FILE_NAME_HEADING].onChanging = refreshExportRows;
    }
    if (dateHeading !== "" && fieldInputs[dateHeading]) {
        fieldInputs[dateHeading].onChanging = refreshExportRows;
    }
    /* 書類名はタイトルにもファイル名にも出るため、選び直すたびに両方を付け直す / Retitle and rename on every pick */
    for (var radioIndex = 0; radioIndex < documentTypeRadios.length; radioIndex++) {
        documentTypeRadios[radioIndex].onClick = refreshTemplateAndExportRows;
    }

    selectDocumentTypeRadio(resolveDocumentTypeName(templateFile));
    refreshTaxBreakdown();
    refreshTemplateAndExportRows();

    // =========================================
    // タグの置換 / Tag replacement
    // =========================================

    /**
     * `#見出し#` を、その見出しの入力値に置き換える
     * @param {string} templateText - 置換前の文字列
     * @returns {string} 置換後の文字列
     */
    function fillHeadingTokens(templateText) {
        var filledText = String(templateText);
        for (var i = 0; i < inputHeadings.length; i++) {
            var inputField = fieldInputs[inputHeadings[i]];
            var inputText = inputField ? String(inputField.text) : "";
            /* 全出現を置換（ExtendScript安全）/ Replace every occurrence */
            filledText = filledText.split("#" + inputHeadings[i] + "#").join(inputText);
        }
        return filledText;
    }

    /**
     * 選んだ書き方で、適用（但し書き）の文字列を組み立てる
     * @returns {string} テンプレートの `<適用>` に流し込む文字列
     */
    function buildApplicationText() {
        var selectedPreset = APPLICATION_PRESETS[0];
        for (var i = 0; i < applicationRadios.length; i++) {
            if (applicationRadios[i].value) { selectedPreset = APPLICATION_PRESETS[i]; break; }
        }
        return fillHeadingTokens(selectedPreset.format);
    }

    /**
     * ダイアログの入力から、タグと流し込む値の組を作る
     * @returns {Array<{tag: string, value: string}>} 置換の組
     */
    function buildTagReplacements() {
        var taxAmounts = splitTaxFromTotal(amountInput ? parseAmount(amountInput.text) : 0);
        var tagReplacements = [];

        for (var i = 0; i < FIELD_MAPPINGS.length; i++) {
            var fieldMapping = FIELD_MAPPINGS[i];
            var inputField = fieldInputs[fieldMapping.heading];
            var inputText = inputField ? String(inputField.text) : "";
            var replacementValue;

            if (fieldMapping.valueType === "date") {
                replacementValue = formatDateValue(inputText);
            } else if (fieldMapping.valueType === "documentType") {
                replacementValue = currentDocumentTypeName();
            } else if (fieldMapping.valueType === "application") {
                replacementValue = buildApplicationText();
            } else if (fieldMapping.valueType === "amountIncludingTax") {
                replacementValue = formatAmount(taxAmounts.includingTax);
            } else if (fieldMapping.valueType === "amountExcludingTax") {
                replacementValue = formatAmount(taxAmounts.excludingTax);
            } else if (fieldMapping.valueType === "tax") {
                replacementValue = formatAmount(taxAmounts.tax);
            } else {
                replacementValue = inputText;
            }
            tagReplacements.push({ tag: fieldMapping.tag, value: replacementValue });
        }
        return tagReplacements;
    }

    /**
     * レイヤーのロックを再帰的に解除する（サブレイヤーも含む）
     * @param {Layers} layerCollection - 対象のレイヤーコレクション
     * @returns {void}
     */
    function unlockLayers(layerCollection) {
        for (var i = 0; i < layerCollection.length; i++) {
            layerCollection[i].locked = false;
            unlockLayers(layerCollection[i].layers);
        }
    }

    /**
     * 置換できるよう、レイヤーとテキストフレームのロックを解除する
     * 作業用の複製にだけ行う。ロックはPDFの見た目に影響しないので体裁は変わらず、
     * ロックされたレイヤー上のタグも置換できるようになる。非表示はそのまま残す
     * @param {Document} targetDocument - 対象のドキュメント
     * @returns {void}
     */
    function unlockForReplacement(targetDocument) {
        unlockLayers(targetDocument.layers);
        /* グループがロックされていると中のテキストも編集できないため、オブジェクトも一律で解除する / Locked groups block their children */
        for (var i = 0; i < targetDocument.pageItems.length; i++) {
            targetDocument.pageItems[i].locked = false;
        }
    }

    /**
     * テキストフレームの内容を読む（読めないものはnullを返す）
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @returns {string} フレームの内容。読めなければnull
     */
    function readFrameContents(textFrame) {
        try {
            var frameContents = textFrame.contents;
            return (frameContents == null) ? null : String(frameContents);
        } catch (e) {
            return null; // 内容を取り出せないフレームは対象外 / skip frames we cannot read
        }
    }

    /**
     * ドキュメント内のテキストフレームを配列に写し取る
     * 内容を書き換えるとコレクションの中身が変わるため、先に配列へ控える
     * @param {Document} targetDocument - 対象のドキュメント
     * @returns {Array<TextFrame>} テキストフレームの配列
     */
    function collectDocumentTextFrames(targetDocument) {
        var textFrames = [];
        for (var i = 0; i < targetDocument.textFrames.length; i++) {
            textFrames.push(targetDocument.textFrames[i]);
        }
        return textFrames;
    }

    /**
     * テキストフレーム内のタグを、文字書式を保ったまま置換する
     * タグの1文字目にデータ値を流し込み、残りのタグ文字を削除することで書式を引き継ぐ
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @param {string} placeholderTag - 置換する `<タグ>` 形式の文字列
     * @param {string} replacementValue - 流し込む値
     * @returns {void}
     */
    function replaceTagKeepingStyle(textFrame, placeholderTag, replacementValue) {
        var guardCount = 0;
        var tagPosition = textFrame.contents.indexOf(placeholderTag);
        while (tagPosition !== -1 && guardCount++ < MAX_TAG_REPLACEMENTS) {
            if (replacementValue === "") {
                for (var i = 0; i < placeholderTag.length; i++) {
                    textFrame.characters[tagPosition].remove();
                }
            } else {
                textFrame.characters[tagPosition].contents = replacementValue; // 1文字目の書式を引き継ぐ / inherit the style
                for (var j = 1; j < placeholderTag.length; j++) {
                    textFrame.characters[tagPosition + replacementValue.length].remove();
                }
            }
            tagPosition = textFrame.contents.indexOf(placeholderTag);
        }
    }

    /**
     * 集めた置換をすべて文字列に適用する（書式を保てないときのフォールバック用）
     * @param {string} textContents - 置換前のテキスト
     * @param {Array<{tag: string, value: string}>} tagReplacements - 置換の組
     * @returns {string} 置換後のテキスト
     */
    function applyReplacementsToText(textContents, tagReplacements) {
        var mergedContents = String(textContents);
        for (var i = 0; i < tagReplacements.length; i++) {
            /* 全出現を置換（ExtendScript安全）/ Replace every occurrence */
            mergedContents = mergedContents.split(tagReplacements[i].tag).join(tagReplacements[i].value);
        }
        return mergedContents;
    }

    /**
     * テキストフレーム1つぶんのタグを置換する
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @param {Array<{tag: string, value: string}>} tagReplacements - 置換の組
     * @returns {void}
     */
    function replaceTagsInFrame(textFrame, tagReplacements) {
        var originalContents = readFrameContents(textFrame);
        if (originalContents === null) return;

        var applicableReplacements = [];
        for (var i = 0; i < tagReplacements.length; i++) {
            if (originalContents.indexOf(tagReplacements[i].tag) !== -1) applicableReplacements.push(tagReplacements[i]);
        }
        if (applicableReplacements.length === 0) return; // タグが無いフレームには触らない / leave untagged frames alone

        try {
            for (var j = 0; j < applicableReplacements.length; j++) {
                replaceTagKeepingStyle(textFrame, applicableReplacements[j].tag, applicableReplacements[j].value);
            }
        } catch (e) {
            /* パス上文字など文字単位で扱えない場合は一括代入に切り替える（書式は失われる）/ Fallback: whole-contents assignment */
            textFrame.contents = applyReplacementsToText(originalContents, applicableReplacements);
        }
    }

    /**
     * ドキュメント全体のタグを置換し、見つからなかったタグを返す
     * @param {Document} targetDocument - 対象のドキュメント
     * @param {Array<{tag: string, value: string}>} tagReplacements - 置換の組
     * @returns {string[]} テンプレートに存在しなかったタグ
     */
    function replaceTagsInDocument(targetDocument, tagReplacements) {
        unlockForReplacement(targetDocument);
        var textFrames = collectDocumentTextFrames(targetDocument);

        /* 置換前に走査しておく。置換後はタグが消えていて有無を判定できない / Check before replacing: the tags are gone afterwards */
        var missingTags = [];
        for (var i = 0; i < tagReplacements.length; i++) {
            var isFound = false;
            for (var j = 0; j < textFrames.length; j++) {
                var frameContents = readFrameContents(textFrames[j]);
                if (frameContents !== null && frameContents.indexOf(tagReplacements[i].tag) !== -1) { isFound = true; break; }
            }
            if (!isFound) missingTags.push(tagReplacements[i].tag);
        }

        for (var k = 0; k < textFrames.length; k++) {
            try {
                replaceTagsInFrame(textFrames[k], tagReplacements);
            } catch (e) {
                /* 1つのフレームで失敗しても、残りの置換とPDF書き出しは続ける / One bad frame must not lose the whole PDF */
            }
        }
        return missingTags;
    }

    // =========================================
    // 返信メールの文面 / Reply mail text
    // =========================================

    /**
     * `#見出し#` をダイアログの入力値に、`#ファイル名#` を作成したPDFの名前に置き換える
     * @param {string} templateText - 置換前の文字列（件名または本文）
     * @param {File} pdfFile - 作成したPDFファイル
     * @returns {string} 置換後の文字列
     */
    function fillMailTokens(templateText, pdfFile) {
        var filledText = fillHeadingTokens(templateText);
        filledText = filledText.split("#" + REPLY_MAIL_FILE_NAME_TOKEN + "#").join(decodeURI(pdfFile.name));
        return fillDocumentTypeToken(filledText, currentDocumentTypeName());
    }

    /**
     * メールソフトで下書きを開くための mailto URL を組み立てる
     * 宛先はそのまま渡す（メールアドレスに変換が要る文字は入らない）。
     * 件名と本文は改行や記号を含むため、URL用に変換する
     * @param {File} pdfFile - 作成したPDFファイル
     * @param {string} mailText - 本文
     * @returns {string} mailto URL
     */
    function buildMailToUrl(pdfFile, mailText) {
        var addressInput = (MAIL_ADDRESS_HEADING !== "") ? fieldInputs[MAIL_ADDRESS_HEADING] : null;
        var mailAddress = addressInput ? trimAndStripBom(addressInput.text) : "";
        return "mailto:" + mailAddress +
            "?subject=" + encodeURIComponent(fillMailTokens(REPLY_MAIL_SUBJECT, pdfFile)) +
            "&body=" + encodeURIComponent(mailText);
    }

    /**
     * 返信メールに必要な情報をまとめて取り出す
     * ダイアログを閉じたあとは入力欄を読めないため、閉じる前に呼ぶ
     * @param {File} pdfFile - 作成したPDFファイル
     * @returns {{text: string, mailToUrl: string}} 本文と、メールソフトを開くURL
     */
    function buildReplyMail(pdfFile) {
        var mailText = fillMailTokens(REPLY_MAIL_TEMPLATE, pdfFile);
        return { text: mailText, mailToUrl: buildMailToUrl(pdfFile, mailText) };
    }

    /**
     * 既定のメールソフトで下書きを開く
     * @param {string} mailToUrl - 開く mailto URL
     * @returns {void}
     */
    function openMailDraft(mailToUrl) {
        try {
            new File(mailToUrl).execute();
        } catch (e) {
            // 開けなくても文面はクリップボードに入っている / the body is on the clipboard either way
        }
    }

    /**
     * 一時テキストフレーム経由でクリップボードを書き換える
     * Illustratorには文字列を直接クリップボードへ送るAPIが無いため、内容を持つフレームを作ってコピーする。
     * 追加直後のフレームは再描画しないとコピー対象にならず、app.copy() は黙って無視されることがある。
     * 置くのは書き出し済みの作業用ドキュメントなので、フレームを消さずにそのまま閉じてよい。
     * @param {Document} targetDocument - 一時フレームを置くドキュメント
     * @param {string} textContent - クリップボードに残す文字列
     * @returns {boolean} コピーできたらtrue
     */
    function writeTextToClipboard(targetDocument, textContent) {
        try {
            /* 元のレイヤーはロックされていることがあるため、新しいレイヤーに置く / The original layers may be locked */
            var clipboardFrame = targetDocument.layers.add().textFrames.add();
            clipboardFrame.contents = textContent;

            /* 追加したフレームを画面に反映してから選択する / Flush the new frame before selecting it */
            app.redraw();
            app.executeMenuCommand("deselectall");
            clipboardFrame.selected = true;
            app.redraw();

            app.executeMenuCommand("copy");
            app.redraw();
            return true;
        } catch (e) {
            return false; // コピーできなくてもPDFは作成済み / the PDF is already written
        }
    }

    // =========================================
    // 作業用複製とPDF書き出し / Working copy & PDF export
    // =========================================

    /**
     * 作業用の複製ファイルのパスを作る
     * リンク画像の相対パスを保つため、テンプレートと同じフォルダーに置く
     * @returns {File} 作業用ファイル
     */
    function buildWorkFilePath() {
        var currentTime = new Date();
        var timestamp = todayDateStamp() + "_" + padTwoDigits(currentTime.getHours()) +
            padTwoDigits(currentTime.getMinutes()) + padTwoDigits(currentTime.getSeconds());

        var templateFileName = templateFile.name;
        var dotIndex = templateFileName.lastIndexOf(".");
        var baseName = (dotIndex >= 0) ? templateFileName.substring(0, dotIndex) : templateFileName;
        var fileExtension = (dotIndex >= 0) ? templateFileName.substring(dotIndex) : ".ai";

        return new File(templateFile.parent.fsName + "/" + baseName + "_work_" + timestamp + fileExtension);
    }

    /**
     * テンプレートを複製して開く
     * 警告の抑止中に通知しないよう、失敗はその場で知らせず呼び出し元へ返す
     * @param {File} workFile - 複製先のファイル
     * @returns {{document: Document, error: string}} 開いたドキュメントと、失敗したときの内容
     */
    function openWorkCopy(workFile) {
        try {
            if (!templateFile.copy(workFile)) throw new Error("File copy failed");
            return { document: app.open(workFile), error: null };
        } catch (e) {
            return { document: null, error: String(e) };
        }
    }

    /**
     * 作業用ドキュメントを保存せずに閉じ、複製ファイルも削除する
     * @param {Document} workDocument - 閉じるドキュメント（開けていなければnull）
     * @param {File} workFile - 削除する複製ファイル
     * @returns {void}
     */
    function discardWorkCopy(workDocument, workFile) {
        try {
            if (workDocument) workDocument.close(SaveOptions.DONOTSAVECHANGES);
        } catch (e) {
            // 閉じられなくても、ファイルの削除と結果の通知は続ける / carry on even if it will not close
        }
        try {
            if (workFile.exists) workFile.remove();
        } catch (e) {
            // 消せない場合は複製が残るが、PDFは作成済み / the PDF is already written
        }
    }

    /**
     * 設定した名前からPDFの互換性を決める
     * 名前が合わないときはPDF 1.6（Acrobat 7）にする
     * @returns {PDFCompatibility} 互換性の指定
     */
    function resolvePdfCompatibility() {
        var compatibilityByName = {
            ACROBAT5: PDFCompatibility.ACROBAT5,
            ACROBAT6: PDFCompatibility.ACROBAT6,
            ACROBAT7: PDFCompatibility.ACROBAT7,
            ACROBAT8: PDFCompatibility.ACROBAT8
        };
        var chosenCompatibility = compatibilityByName[PDF_COMPATIBILITY_NAME];
        return chosenCompatibility ? chosenCompatibility : PDFCompatibility.ACROBAT7;
    }

    /**
     * この環境に用意されているプリセット名を、候補の中から選ぶ
     * 名前は言語ごとに違うため、実際のプリセット一覧と突き合わせる
     * @returns {string} 見つかったプリセット名（無ければ空文字）
     */
    function resolvePdfPresetName() {
        var availablePresets = app.PDFPresetsList;
        if (!availablePresets) return "";

        for (var i = 0; i < PDF_PRESET_CANDIDATES.length; i++) {
            for (var j = 0; j < availablePresets.length; j++) {
                if (String(availablePresets[j]) === PDF_PRESET_CANDIDATES[i]) return PDF_PRESET_CANDIDATES[i];
            }
        }
        return "";
    }

    /**
     * 書き出しに使うPDFオプションを組み立てる
     * プリセットが見つかればそれを読み込み、無ければ同等の設定を個別に組む
     * @returns {PDFSaveOptions} 書き出しオプション
     */
    function buildPdfSaveOptions() {
        var pdfOptions = new PDFSaveOptions();
        pdfOptions.viewAfterSaving = false;

        var presetName = resolvePdfPresetName();
        if (presetName !== "") {
            /* 互換性はプリセットの値を上書きするので、読み込んだあとに指定する / Load first, then override */
            pdfOptions.pDFPreset = presetName;
            pdfOptions.compatibility = resolvePdfCompatibility();
            return pdfOptions;
        }

        /* 編集用データとサムネールを外すのが、ファイルサイズでは一番効く / Dropping AI data shrinks it the most */
        pdfOptions.compatibility = resolvePdfCompatibility();
        pdfOptions.preserveEditability = PDF_PRESERVE_EDITABILITY;
        pdfOptions.generateThumbnails = false;
        pdfOptions.acrobatLayers = false;
        pdfOptions.optimization = true;   /* Web表示用に最適化 / optimize for fast web view */
        pdfOptions.compressArt = true;

        pdfOptions.colorDownsampling = PDF_IMAGE_RESOLUTION;
        pdfOptions.colorDownsamplingImageThreshold = PDF_IMAGE_THRESHOLD;
        pdfOptions.colorDownsamplingMethod = DownsampleMethod.BICUBICDOWNSAMPLE;
        pdfOptions.colorCompression = CompressionQuality.JPEGLOW;

        pdfOptions.grayscaleDownsampling = PDF_IMAGE_RESOLUTION;
        pdfOptions.grayscaleDownsamplingImageThreshold = PDF_IMAGE_THRESHOLD;
        pdfOptions.grayscaleDownsamplingMethod = DownsampleMethod.BICUBICDOWNSAMPLE;
        pdfOptions.grayscaleCompression = CompressionQuality.JPEGLOW;

        /* 白黒画像は線がつぶれると読めなくなるため、カラーより高い解像度を保つ / Keep 1-bit art legible */
        pdfOptions.monochromeDownsampling = 300;
        pdfOptions.monochromeDownsamplingImageThreshold = 450;
        pdfOptions.monochromeDownsamplingMethod = DownsampleMethod.BICUBICDOWNSAMPLE;
        pdfOptions.monochromeCompression = MonochromeCompression.CCIT4;

        return pdfOptions;
    }

    /**
     * ドキュメントをPDFとして書き出す
     * プリセット名や設定は環境によって通らないことがあるため、失敗したら既定設定で書き出し直す
     * @param {Document} targetDocument - 書き出すドキュメント
     * @param {File} pdfFile - 書き出し先のファイル
     * @returns {void}
     */
    function exportDocumentAsPdf(targetDocument, pdfFile) {
        try {
            targetDocument.saveAs(pdfFile, buildPdfSaveOptions());
            return;
        } catch (e) {
            // 設定が通らない環境では、何も指定しない書き出しに切り替える / fall back to plain defaults
        }
        var defaultOptions = new PDFSaveOptions();
        defaultOptions.viewAfterSaving = false;
        targetDocument.saveAs(pdfFile, defaultOptions);
    }

    /**
     * 完了メッセージを組み立てる
     * @param {File} pdfFile - 作成したPDFファイル
     * @param {string[]} missingTags - テンプレートに無かったタグ
     * @param {boolean} isMailCopied - 返信メールをコピーできたか
     * @returns {string} 表示するメッセージ
     */
    function buildDoneMessage(pdfFile, missingTags, isMailCopied) {
        var doneMessage = LABELS.alert.done.replace("#filename#", decodeURI(pdfFile.name));
        if (isMailCopied) doneMessage += "\n\n" + LABELS.alert.mailCopied;
        if (missingTags.length > 0) {
            doneMessage += "\n\n" + LABELS.alert.missingTags.replace("#tags#", missingTags.join("\n"));
        }
        return doneMessage;
    }

    /**
     * テンプレートの複製にデータを流し込み、PDFの書き出しと返信メールのコピーまで行う
     * テンプレート自体には触れず、作業用の複製は最後に削除する
     * @param {Array<{tag: string, value: string}>} tagReplacements - 置換の組
     * @param {{text: string, mailToUrl: string}} replyMail - 返信メールの文面と、メールソフトを開くURL
     * @param {File} pdfFile - 書き出し先のファイル
     * @returns {void}
     */
    function createPdfAndReplyMail(tagReplacements, replyMail, pdfFile) {
        var workFile = buildWorkFilePath();

        /* 作業が終わったら、実行前に前面だったドキュメントへ戻す / Come back to whatever was in front before */
        var previousDocument = (app.documents.length > 0) ? app.activeDocument : null;

        /* 複製を開く際のプロファイル警告などで処理が止まらないようにする / Keep Illustrator's own dialogs out of the way */
        var previousInteractionLevel = app.userInteractionLevel;
        app.userInteractionLevel = UserInteractionLevel.DONTDISPLAYALERTS;

        var workCopy = openWorkCopy(workFile);
        if (workCopy.document === null) {
            discardWorkCopy(null, workFile);
            /* 警告を抑止したまま知らせない / Restore before reporting */
            app.userInteractionLevel = previousInteractionLevel;
            alert(LABELS.alert.workCopyFailed.replace("#detail#", workCopy.error));
            return;
        }
        var workDocument = workCopy.document;

        var missingTags = [];
        var isMailCopied = false;
        var exportError = null;
        try {
            missingTags = replaceTagsInDocument(workDocument, tagReplacements);
            workDocument.selection = null;
            app.redraw();
            exportDocumentAsPdf(workDocument, pdfFile);
            /* 書き出したあとに一時フレームを置くので、PDFには入らない / The temp frame is added after the export */
            isMailCopied = writeTextToClipboard(workDocument, replyMail.text);
        } catch (e) {
            exportError = String(e);
        }

        discardWorkCopy(workDocument, workFile);
        if (previousDocument) previousDocument.activate();
        /* 結果を知らせる前に戻す / Restore before reporting the result */
        app.userInteractionLevel = previousInteractionLevel;

        if (exportError !== null) {
            alert(LABELS.alert.exportFailed.replace("#detail#", exportError));
            return;
        }
        alert(buildDoneMessage(pdfFile, missingTags, isMailCopied));

        /* 開くのは知らせたあと。先に開くと、前面に出たFinderやメールソフトの裏にalertが隠れる / Reveal after the alert */
        if (OPEN_FOLDER_AFTER_EXPORT) pdfFile.parent.execute();
        /* 添付はmailtoで指定できないため、開いたフォルダーからPDFをドラッグして添付する / Attach the PDF by hand */
        if (OPEN_MAIL_AFTER_EXPORT) openMailDraft(replyMail.mailToUrl);
    }

    /* 「PDFを作成」：ダイアログを閉じてから書き出す / Run: close the dialog, then export */
    runButton.onClick = function () {
        var pdfFile = resolvePdfFile();
        if (pdfFile === null) {
            alert(LABELS.alert.noTemplate);
            return;
        }
        /* 閉じたあとの入力欄は参照できなくなりうるので、必要な値はすべて先に取り出す / Read every field before closing */
        var tagReplacements = buildTagReplacements();
        var replyMail = buildReplyMail(pdfFile);
        mainDialog.close();
        createPdfAndReplyMail(tagReplacements, replyMail, pdfFile);
    };

    mainDialog.show();
})();
