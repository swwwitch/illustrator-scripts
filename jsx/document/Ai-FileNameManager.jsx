#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
 
### スクリプト名：
 
Ai-FileNameManager.jsx — 現在のドキュメントのファイル名を変更／別名で保存

### 概要：

- アクティブな Illustrator ドキュメントのファイル名を変更（実ファイルのリネーム）／別名で保存／コピーを保存します。
- 保存先は常に現在のファイルと同じフォルダ。未保存ドキュメントは保存先フォルダを選択して保存します。
- 動作モードは 3 種類：「元ファイルをリネーム」（新名で保存後、元ファイルをゴミ箱へ移動）／「別名で保存」（元ファイルを残し、作業ドキュメントを新ファイルへ切替）／「コピー（複製）を保存」（元ファイルへ上書き保存後、別名のコピーを作成）。
- スコープは「全体」（すべての整形を適用）と「バージョンのみ」（元ファイル名の v 番号だけを繰り上げ、他の整形はしない）から選択。元ファイルにタイムスタンプが無く v 番号がある場合は「バージョンのみ」を初期選択します。
- ファイル名を base / title / status / timestamp / page / version の各セグメントに分解し、UI から個別に編集できます。
  - base: 先頭の固定部分。「ベース」入力欄で編集可能（検出値が初期値、空欄可）
  - title: 「サブテキスト」ラジオで「なし / 指定（入力欄）/ 親フォルダー / 2 階層上のフォルダー」を選択
  - status: 「ステータス」dropdown で制作ステータス（wip / draft / review / approved / outlined など）を挿入。`：` より前の値だけがファイル名に入る
  - timestamp: 「タイムスタンプ」ラジオで「なし / YYYYMMDD / YYYY-MM-DD」を選択（デフォルト YYYYMMDD）。「時刻も付与」で末尾に -HHMM を付加
  - page: 「連番」で page01 / page001 のような連番セグメントを追加（プレフィックスと桁数 01 / 001 を指定可）。元ファイル名の pageNN は自動検出して初期値に反映し、保存時は「既存 +1」かつ同フォルダ内の最大連番 +1 へ自動繰り上げ
  - version: 「バージョン番号」ラジオで「なし / v1, v2… / v01, v02… / v001, v002…」を選択（デフォルトは「なし」）。既存の v 番号は選択形式のまま +1、無い場合は選択形式で v1 / v01 / v001 から採番。親フォルダ内に同パターンがあれば最大 v +1 まで自動繰り上げ
- 構成要素の順序は「標準順 / 現在のファイル名に準じる / カスタム順（↑↓ で編集）」から選択できます。
- ステータス検出、サブテキスト選択、連番（pageNN）、タイムスタンプ／バージョン番号の形式切替や自動繰り上げに対応します。
- 区切り記号統一、NFC 正規化、クリーンなファイル名処理（OS 禁止文字・絵文字・機種依存文字）、半角カナ → 全角変換、丸数字や法人略記などの ASCII 化など、ファイル名整形機能を搭載しています。ファイル名が長すぎる場合は確認ダイアログを表示します。

### 解説記事：

https://note.com/dtp_tranist/n/nc88dd887eb1c
 
### 更新履歴：

- v1.0 (2026-05-27) : 初期バージョン
- v1.3.5 (2026-07-12) : mergeFragmentedText が page/date/status/version など認識済みセグメントをタイトルへ再結合しないよう修正（text 値のみ結合）。既存 pageNN 検出時は prefs より優先して連番を必ず ON に。連番の初期候補を「既存番号 +1」とし、フォルダー検索失敗時も既存番号を下回らないように。「ファイル名の設定」パネルのラベルを右揃えに変更。loadPrefs が CRLF 保存された prefs の行末 \r を除去するよう堅牢化。moveToTrash がゴミ箱へコピー後にオリジナル削除へ失敗した場合、コピーを後始末して二重化を防止。先頭トークンは常に base（先頭日付も base）である仕様とバージョン番号の初期値が「なし」であることをコメントへ明記。内部の変数名・関数名を推測しやすい具体名へ整理（挙動不変）。概要コメントを最新実装に更新（スクリプト名、時刻付与 HHMM、連番の自動検出／プレフィックス・桁数指定、半角カナ→全角、長すぎるファイル名の警告）。
 
---
 
### Script name:
 
Ai-FileNameManager.jsx — Rename or Save As the Current Document
 
### Overview:

- Renames the active Illustrator document (true rename), saves as a new file, or saves a copy.
- Destination is always the same folder as the current file; for unsaved docs, a destination folder is chosen.
- Three modes: "Rename Original" (save under the new name, then move the original to the Trash) / "Save As" (keep the original and switch the active document to the new file) / "Save a Copy" (save over the original, then create a copy under the new name).
- Scope is either "Full" (apply all formatting) or "Version Only" (bump just the v-number in the original filename, skipping all other formatting). "Version Only" is preselected when the original has a version number but no timestamp.
- Decomposes the filename into base / title / status / timestamp / page / version segments and lets you edit them.
  - base: editable via the "Base" input (defaults to the detected value; may be empty)
  - title: "Project Name" radio (None / Custom input / Parent Folder / Grandparent Folder)
  - status: "Status" dropdown (wip / draft / review / approved / outlined, etc.); only the text before `:` is written to the filename
  - timestamp: "Timestamp" radio selects "None / YYYYMMDD / YYYY-MM-DD" (default YYYYMMDD); "Append HHMM" adds -HHMM at the end
  - page: "Sequence" adds a sequence segment such as page01 / page001 (prefix and width 01 / 001 are configurable). An existing pageNN in the filename is auto-detected and used as the initial value; on save it uses "existing + 1" and bumps to (max sequence in folder) + 1
  - version: "Version" radio selects "None / v1, v2… / v01, v02… / v001, v002…" (default None). Existing v-numbers are bumped by +1 in the selected format; when there is none, numbering starts from v1 / v01 / v001 in the selected format. If files matching the same pattern exist in the parent folder, the result is bumped further to (max v) + 1.
- Segment order can be set to "Default / Match Current / Custom" (reorder with ↑↓).
- Supports status detection, project name selection from folder hierarchy, sequence numbers (pageNN), timestamp/version formatting, and auto-increment.
- Includes separator unification, NFC normalization, filename cleaning (OS-invalid / emoji / platform-dependent chars), half-width to full-width kana conversion, and ASCII conversion of symbols/corporate abbreviations. Shows a confirmation dialog when the filename is too long.

### Changelog:

- v1.0 (2026-05-27): Initial release
- v1.3.5 (2026-07-12): mergeFragmentedText no longer folds recognized segments (page/date/status/version) back into the title (only text values are merged). When an existing pageNN is detected, the sequence is forced ON over prefs. The sequence starting candidate is now "existing + 1", and never falls below the existing number even if the folder scan fails. Right-aligned the labels in the Filename Settings panel. Hardened loadPrefs to strip trailing \r from CRLF-saved prefs. moveToTrash now cleans up the trashed copy if removing the original fails, avoiding duplicates. Documented that the leading token is always base (leading date stays base) and that the version default is "None". Clarified internal variable/function names for readability (no behavior change). Updated the overview comments to match the current implementation (script name, "Append HHMM", sequence auto-detection / configurable prefix and width, half-width to full-width kana, and the too-long-filename warning).
 
*/

(function () {

    // =========================================
    // バージョン / Version
    // =========================================

    var SCRIPT_VERSION = "v1.3.5";

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* 整形機能のオン/オフ。false にすると該当 UI とロジックを丸ごとスキップ
       / Feature switches. Set to false to omit the UI and skip the logic */
    var FEATURE_STATUS = true;        // ステータス dropdown と検出 / Status dropdown + detection
    var FEATURE_SORT = true;          // ソートパネル + 並び順カスタマイズ / Sort panel + custom segment order
    var FEATURE_SEPARATOR = '-';      // 区切り記号統一: '-' / '_' で有効化＋既定値、false で無効 / '-' or '_' enables with that default; false disables
    var FEATURE_DOT_NORMALIZE = true; // "." を区切り記号にあわせて置換（要 FEATURE_SEPARATOR） / "." normalization with the chosen separator
    var FEATURE_NFC = true;           // 濁点・半濁点の NFC 結合 / Combine separated dakuten/handakuten (NFC)
    var FEATURE_CLEAN = true;         // クリーンなファイル名（OS 禁止文字 + 絵文字・機種依存文字の処理） / Clean filename (OS-invalid + emoji + platform-dependent)
    var FEATURE_TRANSLITERATE = true; // 法人略記・丸数字などを ASCII 相当に変換 / Transliterate corp abbrev / circled numbers / etc.
    var FEATURE_HALFWIDTH_KANA = true; // 半角カナ → 全角カナ変換（クリーンが '-' / '_' のときのみ有効） / Halfwidth-kana → fullwidth (active only when Clean is '-' or '_')
    var FEATURE_PAGE = true;          // 連番（pageNN）セグメント / Page-number segment (pageNN)

    /* ファイル名（拡張子込み）の UTF-8 バイト長の上限。超過時は確認ダイアログを出して続行可
       / Max UTF-8 byte length of the basename + extension; exceeding it triggers a confirm dialog */
    var FEATURE_MAX_FILENAME_BYTES = 240;

    /* rename モードで元ファイルを削除するときに、即削除ではなく ~/.Trash に移動
       / In rename mode, move the original to ~/.Trash instead of removing it outright */
    var FEATURE_USE_TRASH = true;

    /* 文字 → 置換文字列のマップ。法人略記、丸数字（白・黒・括弧）、ローマ数字、略記号、単位、ダッシュ類
       / Char → replacement map: corporate abbrev, circled (white/black/parenthesized), Roman numerals, abbrev symbols, units, dashes */
    var TRANSLITERATE_MAP = {
        // 法人略記
        '㈱': '株', '㈲': '有', '㈹': '代', '㈳': '社',
        '㈵': '特', '㈶': '財', '㈻': '学', '㍿': '株式会社',
        // 白丸数字 ①-⑳ (U+2460-U+2473)
        '①': '1', '②': '2', '③': '3', '④': '4', '⑤': '5',
        '⑥': '6', '⑦': '7', '⑧': '8', '⑨': '9', '⑩': '10',
        '⑪': '11', '⑫': '12', '⑬': '13', '⑭': '14', '⑮': '15',
        '⑯': '16', '⑰': '17', '⑱': '18', '⑲': '19', '⑳': '20',
        // 黒丸数字 ❶-❿ (U+2776-U+277F) + ⓫-⓴ (U+24EB-U+24F4)
        '❶': '1', '❷': '2', '❸': '3', '❹': '4', '❺': '5',
        '❻': '6', '❼': '7', '❽': '8', '❾': '9', '❿': '10',
        '⓫': '11', '⓬': '12', '⓭': '13', '⓮': '14', '⓯': '15',
        '⓰': '16', '⓱': '17', '⓲': '18', '⓳': '19', '⓴': '20',
        // 括弧数字 ⑴-⒇ (U+2474-U+2487)
        '⑴': '1', '⑵': '2', '⑶': '3', '⑷': '4', '⑸': '5',
        '⑹': '6', '⑺': '7', '⑻': '8', '⑼': '9', '⑽': '10',
        '⑾': '11', '⑿': '12', '⒀': '13', '⒁': '14', '⒂': '15',
        '⒃': '16', '⒄': '17', '⒅': '18', '⒆': '19', '⒇': '20',
        // ローマ数字 大文字 Ⅰ-Ⅻ (U+2160-U+216B)
        'Ⅰ': '1', 'Ⅱ': '2', 'Ⅲ': '3', 'Ⅳ': '4', 'Ⅴ': '5',
        'Ⅵ': '6', 'Ⅶ': '7', 'Ⅷ': '8', 'Ⅸ': '9', 'Ⅹ': '10',
        'Ⅺ': '11', 'Ⅻ': '12',
        // ローマ数字 小文字 ⅰ-ⅻ (U+2170-U+217B)
        'ⅰ': '1', 'ⅱ': '2', 'ⅲ': '3', 'ⅳ': '4', 'ⅴ': '5',
        'ⅵ': '6', 'ⅶ': '7', 'ⅷ': '8', 'ⅸ': '9', 'ⅹ': '10',
        'ⅺ': '11', 'ⅻ': '12',
        // 略記号
        '℡': 'TEL', '№': 'No',
        // 単位記号
        '㎜': 'mm', '㎝': 'cm', '㎞': 'km',
        '㎎': 'mg', '㎏': 'kg',
        '㎡': 'm2', '㎥': 'm3',
        // チルダ類
        '〜': '~', '～': '~',  // 〜 WAVE DASH / ～ FULLWIDTH TILDE
        // ハイフン・ダッシュ類
        '－': '-',  // － FULLWIDTH HYPHEN-MINUS
        '‐': '-',  // ‐ HYPHEN
        '‑': '-',  // ‑ NON-BREAKING HYPHEN
        '‒': '-',  // ‒ FIGURE DASH
        '–': '-',  // – EN DASH
        '—': '-',  // — EM DASH
        '―': '-',  // ― HORIZONTAL BAR
        '−': '-'   // − MINUS SIGN
    };

    /* 出力時のセグメント順序。base / title / status / timestamp / page / version。
       FEATURE_STATUS=false なら status を、FEATURE_PAGE=false なら page を除外
       / Output segment order; "status" / "page" are dropped when their FEATURE_* flags are false */
    var SEGMENT_ORDER = (function () {
        var order = ['base', 'title'];
        if (FEATURE_STATUS) order.push('status');
        order.push('timestamp');
        if (FEATURE_PAGE) order.push('page');
        order.push('version');
        return order;
    })();

    /* ステータス選択肢。value がファイル名に入り、ja / en が UI 表示。先頭は「なし」。
       ja が '---' の項目は dropdown 上の区切り線として表示（選択不可）
       / Status choices: `value` is what enters the filename; `ja` / `en` are UI labels.
       Items with `ja === '---'` render as an unselectable divider. */
    var STATUS_ITEMS = [
        { value: '', ja: 'なし', en: 'None' },
        { value: 'wip', ja: 'wip：作業中・仕掛かり中', en: 'wip: Work in progress' },
        { value: 'draft', ja: 'draft：下書き・ラフ・素案', en: 'draft: Draft / Rough' },
        { value: 'review', ja: 'review：レビュー待ち', en: 'review: Awaiting review' },
        { value: 'revised', ja: 'revised：改訂版・修正反映版', en: 'revised: Revised' },
        { value: 'updated', ja: 'updated：更新版', en: 'updated: Updated' },
        { value: 'fixed', ja: 'fixed：修正完了', en: 'fixed: Fix completed' },
        { value: 'approved', ja: 'approved：承認済み・確定', en: 'approved: Approved / Final' },
        { value: 'rejected', ja: 'rejected：ボツ・不採用案', en: 'rejected: Rejected' },
        { value: 'archived', ja: 'archived：保管用', en: 'archived: Archive' },
        { value: '', ja: '---', en: '---' },
        { value: 'flattened', ja: 'flattened：レイヤー結合済み（.psd）', en: 'flattened: Flattened (.psd)' },
        { value: 'outlined', ja: 'outlined：アウトライン済み（.ai）', en: 'outlined: Outlined (.ai)' }
    ];

    /* 区切り線エントリかどうか / Whether a STATUS_ITEMS entry is a divider */
    function isStatusDivider(item) {
        return item && item.ja === '---';
    }

    var NEW_NAME_FIELD_WIDTH = 250;
    var PANEL_MARGINS = [16, 20, 16, 12];
    var PANEL_SPACING = 8;

    var DIALOG_OPACITY = 0.98;

    // =========================================
    // ローカライズ / Localization
    // =========================================

    var currentLanguage = ($.locale.indexOf("ja") === 0) ? "ja" : "en";

    /* 日英ラベル定義 / Japanese-English label definitions */

    var LABELS = {
        dialog: {
            title: { ja: "ファイル名を変更して保存", en: "Rename and Save" }
        },
        panel: {
            mode: { ja: "動作", en: "Mode" },
            opMode: { ja: "モード", en: "Scope" },
            filename: { ja: "ファイル名プレビュー", en: "File Name Preview" },
            options: { ja: "ファイル名の設定", en: "Filename Settings" },
            sort: { ja: "構成要素の順序", en: "Segment Order" }
        },
        radio: {
            rename: { ja: "元ファイルをリネーム", en: "Rename Original" },
            saveAs: { ja: "別名で保存", en: "Save As" },
            saveCopy: { ja: "コピー（複製）を保存", en: "Save a Copy" },
            opVersionOnly: { ja: "バージョンのみ", en: "Version Only" },
            opFull: { ja: "全体", en: "Full" },
            noChange: { ja: "変更しない", en: "No Change" },
            titleNone: { ja: "なし", en: "None" },
            titleParent: { ja: "親フォルダー", en: "Parent Folder" },
            titleGrandparent: { ja: "2 階層上のフォルダー", en: "Grandparent Folder" },
            titleCustom: { ja: "指定", en: "Custom" },
            timestampNone: { ja: "なし", en: "None" },
            timestampDate: { ja: "YYYYMMDD", en: "YYYYMMDD" },
            timestampDateDash: { ja: "YYYY-MM-DD", en: "YYYY-MM-DD" },
            timestampWithTime: { ja: "時刻も付与", en: "Append HHMM" },
            versionNone: { ja: "なし", en: "None" },
            versionShort: { ja: "v1, v2…", en: "v1, v2…" },
            versionPadded: { ja: "v01, v02…", en: "v01, v02…" },
            versionPaddedWide: { ja: "v001, v002…", en: "v001, v002…" },
            changeToDash: { ja: "-に変更", en: "Change to -" },
            changeToUnderscore: { ja: "_に変更", en: "Change to _" },
            nfcCombine: { ja: "結合する", en: "Combine" },
            cleanRemove: { ja: "削除する", en: "Remove" },
            translitRemove: { ja: "削除する", en: "Remove" },
            translitConvert: { ja: "変換する", en: "Convert" },
            halfwidthKanaConvert: { ja: "半角カナ → 全角", en: "Half-width → Full-width Kana" },
            pageEnable: { ja: "連番を付与", en: "Append Sequence" },
            pagePad2: { ja: "01, 02…", en: "01, 02…" },
            pagePad3: { ja: "001, 002…", en: "001, 002…" },
            sortOff: { ja: "標準順", en: "Default" },
            sortCurrent: { ja: "現在のファイル名に準じる", en: "Match Current" },
            sortOn: { ja: "カスタム順", en: "Custom" }
        },
        label: {
            currentName: { ja: "現在", en: "Current" },
            finalName: { ja: "保存後の名前", en: "Saved Name" },
            base: { ja: "ベース", en: "Base" },
            title: { ja: "サブテキスト", en: "Project Name" },
            status: { ja: "ステータス", en: "Status" },
            timestamp: { ja: "タイムスタンプ", en: "Timestamp" },
            version: { ja: "バージョン番号", en: "Version" },
            page: { ja: "連番", en: "Sequence" },
            separator: { ja: "区切り記号", en: "Separator" },
            nfc: { ja: "濁点・半濁点の正規化", en: "NFC Normalization" },
            clean: { ja: "クリーンなファイル名", en: "Clean Filename" },
            translit: { ja: "丸数字や法人略記など", en: "Symbols" }
        },
        tip: {
            rename: {
                ja: "新しい名前で保存したあと、元ファイルをゴミ箱に移します（実質的にリネーム）。元ファイルを参照している他のドキュメント（配置 .ai／InDesign のリンクなど）はリンク切れになります。",
                en: "Saves with the new name, then moves the original to the Trash (effectively a rename). Documents that reference the original file (placed .ai or InDesign links) will lose the link."
            },
            saveAs: {
                ja: "新しい名前で保存します。元ファイルは残り、作業中のドキュメントが新ファイルに切り替わります。",
                en: "Saves with the new name. The original file is kept, and the active document switches to the new file."
            },
            saveCopy: {
                ja: "元ファイルに上書き保存したうえで、別名のコピーを作成します。作業中のドキュメントは元ファイルのまま残ります。",
                en: "Saves the original, then creates a copy with the new name. The active document remains the original file."
            },
            base: {
                ja: "ファイル名の先頭部分。空欄にすると省略されます。",
                en: "Leading part of the filename. Leave empty to omit."
            },
            title: {
                ja: "ファイル名に追加する案件名・補足テキスト。「指定」で入力欄の文字列を使用します。",
                en: "Choose how to set the project name in the filename. With \"Custom\", the entered text is used."
            },
            status: {
                ja: "ファイル名に挿入する制作ステータスを選択します。「：」より前の文字列が入ります。",
                en: "Choose a production status to insert. Only the text before \":\" is used in the filename."
            },
            timestamp: {
                ja: "タイムスタンプの形式を選択。「なし」で元の日付があっても削除します。",
                en: "Choose timestamp format. \"None\" removes any existing date."
            },
            timestampWithTime: {
                ja: "タイムスタンプの末尾に時刻（HHMM）を付加します。1 日に複数版を出すときに便利。",
                en: "Append the current time (HHMM) to the timestamp. Useful for multiple versions per day."
            },
            version: {
                ja: "バージョン番号の形式。v1/v2 はパディング無し、v01/v02 は 2 桁、v001/v002 は 3 桁ゼロ埋め。既存の v 番号は +1、無い場合は v1/v01/v001 を付与。「なし」で削除。",
                en: "Version format. v1/v2 has no padding, v01/v02 is 2-digit, v001/v002 is 3-digit zero-padded. An existing v-number is bumped by +1; otherwise v1/v01/v001 is added. \"None\" removes."
            },
            page: {
                ja: "「page01」「page001」のような連番をファイル名に追加します。プレフィックス（page 等）と桁数（01 / 001）を選択。保存時に同フォルダ内の最大連番 +1 へ自動繰り上げ。",
                en: "Append a sequence such as page01 / page001. Configure prefix (e.g. page) and width (01 / 001). On save, bumps to (max in folder) + 1 if collisions exist."
            },
            separator: {
                ja: "ファイル名全体の区切り記号の扱いを選択します。「-」「_」「.」が対象。YYYY-MM-DD のタイムスタンプは保護されます。",
                en: "Choose how separators in the filename are handled. Targets `-`, `_`, and `.`. YYYY-MM-DD timestamps are preserved."
            },
            sort: {
                ja: "並び順を選択。「現在のファイル名に準じる」は検出された要素のみ、「カスタム順」は［編集］で並び替え。",
                en: "Choose order. \"Match Current\" uses only detected segments; \"Custom\" enables [Edit]."
            },
            nfc: {
                ja: "ファイル名に分離した濁点・半濁点が含まれる場合、結合済み文字（NFC）に正規化します。",
                en: "Normalize separated dakuten/handakuten to combined characters (NFC)."
            },
            clean: {
                ja: "OS で使えない文字（\\ / : * ? \" < > |）、絵文字・機種依存文字（㈱・①・㌔ など）、スペースの扱いをまとめて選択します。連続するスペースは 1 つにまとめてから処理。",
                en: "How to handle OS-invalid characters (\\ / : * ? \" < > |), emoji, platform-dependent chars (㈱, ①, ㌔), and spaces. Consecutive spaces collapse to one before replacement."
            },
            translit: {
                ja: "法人略記（㈱→株）、丸数字（①→1）、ローマ数字（Ⅰ→1）、TEL/No、単位記号（㎜→mm）、ダッシュ類を ASCII 相当に変換します。",
                en: "Convert corporate abbreviations (㈱→株), circled numbers (①→1), Roman numerals, TEL/No, units (㎜→mm), and dashes to ASCII equivalents."
            },
            halfwidthKana: {
                ja: "「-」または「_」を選んだとき、半角カタカナを全角カタカナに変換します（濁点・半濁点付きは結合）。",
                en: "When `-` or `_` is selected, convert half-width katakana to full-width (voiced/semi-voiced marks are merged)."
            }
        },
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            sort: { ja: "編集", en: "Edit" }
        },
        sort: {
            title: { ja: "ファイル名の並び順", en: "Filename Order" },
            hint: { ja: "選択中の項目を ↑↓ で並び替え", en: "Select an item and use ↑↓ to reorder" }
        },
        message: {
            noDoc: { ja: "ドキュメントが開かれていません", en: "No document is open." },
            emptyName: { ja: "ファイル名が空です", en: "File name is empty." },
            chooseDestination: { ja: "保存先フォルダを指定", en: "Choose destination folder" },
            confirmOverwrite: {
                ja: "同名ファイルが存在します。上書きしますか？",
                en: "A file with the same name exists. Overwrite?"
            },
            confirmTooLong: {
                ja: "ファイル名が長すぎる可能性があります（{bytes} バイト / 上限 {limit} バイト）。このまま続行しますか？",
                en: "The filename may be too long ({bytes} bytes / limit {limit}). Continue anyway?"
            },
            saveFailed: { ja: "保存に失敗しました", en: "Failed to save" }
        }
    };

    /* ドット記法でローカライズ済み文字列を取得 / Get the localized string by dotted key (e.g. L('dialog.title')) */
    function L(path) {
        var parts = String(path).split('.');
        var entry = LABELS;
        for (var i = 0; i < parts.length; i++) {
            entry = entry && entry[parts[i]];
            if (!entry) return path;
        }
        return entry[currentLanguage] || entry.en || path;
    }

    /* コロン付きラベル（日本語は全角、英語は半角）/ Label with colon (full-width JA, half-width EN) */
    function labelText(path) {
        return L(path) + (currentLanguage === 'ja' ? '：' : ':');
    }

    // =========================================
    // ヘルパー / Helpers
    // =========================================

    var VERSION_TOKEN_RE = /^[vV]\d+$/;   // v123 / V123
    var PAGE_TOKEN_RE = /^(page)(\d+)$/i;  // page01 / PAGE003（既定プレフィックス page + 数字）

    /* 実在する暦日か（うるう年・月末を含む）を Date で判定。
       year/month/day は文字列でも数値でも可。存在しない日付（例: 2026-02-31）は false
       / Whether year/month/day is a real calendar date (leap years / month-ends), validated via Date */
    function isRealDate(yearStr, monthStr, dayStr) {
        var y = parseInt(yearStr, 10);
        var m = parseInt(monthStr, 10);
        var d = parseInt(dayStr, 10);
        if (isNaN(y) || m < 1 || m > 12 || d < 1 || d > 31) return false;
        var dt = new Date(y, m - 1, d);
        // Date が桁あふれ（2/31 → 3/3 など）を補正するので、元の Y/M/D と一致するかで実在日を判定
        return dt.getFullYear() === y && dt.getMonth() === (m - 1) && dt.getDate() === d;
    }

    /* YYYYMMDD（8 桁）または YYMMDD（6 桁）の日付トークンか。実在する暦日のみ true。
       6 桁の年は 2000 年代として扱い、うるう年判定を行う
       / Whether a token is a real YYYYMMDD or YYMMDD date (6-digit year assumed 20xx for leap checks) */
    function isDateToken(token) {
        var s = String(token);
        if (/^\d{8}$/.test(s)) return isRealDate(s.substring(0, 4), s.substring(4, 6), s.substring(6, 8));
        if (/^\d{6}$/.test(s)) return isRealDate('20' + s.substring(0, 2), s.substring(2, 4), s.substring(4, 6));
        return false;
    }

    /* "v123" 形式のバージョントークンか / Whether a token is a v+digits version */
    function isVersionToken(token) {
        return VERSION_TOKEN_RE.test(String(token));
    }

    /* "page01" / "page003" 形式の連番トークンか（既定プレフィックス page、大文字小文字無視）
       / Whether a token is a page-number token ("page" + digits, case-insensitive) */
    function isPageToken(token) {
        return PAGE_TOKEN_RE.test(String(token));
    }

    /* STATUS_ITEMS のいずれかの value と完全一致（大文字小文字無視）すれば true。一致した正規 value を返す。divider は除外
       / Match against STATUS_ITEMS values (case-insensitive); returns the canonical value or '' (skips dividers) */
    function matchStatusToken(token) {
        var lowerToken = String(token).toLowerCase();
        for (var i = 1; i < STATUS_ITEMS.length; i++) {
            if (isStatusDivider(STATUS_ITEMS[i])) continue;
            if (STATUS_ITEMS[i].value === lowerToken) return STATUS_ITEMS[i].value;
        }
        return '';
    }

    /* ファイル名を順序付きセグメント配列に分解。各 segment は前にあった区切り文字も保持
       / Decompose into ordered segments; each segment stores the preceding separator */
    function parseFileName(name) {
        var splitTokens = String(name).split(/([-_])/); // ["handout","-","Adobe","-","20260422"]
        var segments = [];
        var textBuffer = [];      // token と区切りを交互に蓄積（join('') で結合）
        var textLeadingSep = '';  // テキスト segment の直前に置く区切り
        var hasDate = false, hasVersion = false, hasStatus = false, hasPage = false;
        var currentSep = '';

        function flushText() {
            if (!textBuffer.length) return;
            segments.push({
                kind: 'text',
                value: textBuffer.join(''),
                sep: textLeadingSep
            });
            textBuffer = [];
            textLeadingSep = '';
        }

        var i = 0;
        while (i < splitTokens.length) {
            if (i % 2 === 1) {
                currentSep = splitTokens[i];
                i++;
                continue;
            }
            var token = splitTokens[i];
            // ステータス判定は同一トークンで 1 回だけ評価し、結果を再利用する
            var statusMatch = (FEATURE_STATUS && !hasStatus) ? matchStatusToken(token) : '';

            // YYYY-MM-DD / YYYY_MM_DD（3 トークン + 同じ区切り 2 つ）の日付パターン
            if (!hasDate && i > 0
                && /^\d{4}$/.test(token)
                && i + 4 < splitTokens.length
                && (splitTokens[i + 1] === '-' || splitTokens[i + 1] === '_')
                && /^\d{2}$/.test(splitTokens[i + 2])
                && splitTokens[i + 3] === splitTokens[i + 1]
                && /^\d{2}$/.test(splitTokens[i + 4])
                && isRealDate(token, splitTokens[i + 2], splitTokens[i + 4])) {
                flushText();
                var compositeDate = token + splitTokens[i + 1] + splitTokens[i + 2] + splitTokens[i + 3] + splitTokens[i + 4];
                segments.push({ kind: 'date', value: compositeDate, sep: currentSep });
                hasDate = true;
                i += 5; // 3 tokens + 2 separators
                continue;
            }

            if (i === 0) {
                // 仕様: 先頭トークンは常に base として扱う。
                // したがって "20260712-catalog" のような先頭日付は date ではなく base になる
                segments.push({ kind: 'base', value: token, sep: '' });
            } else if (!hasDate && isDateToken(token)) {
                flushText();
                segments.push({ kind: 'date', value: token, sep: currentSep });
                hasDate = true;
            } else if (!hasVersion && isVersionToken(token)) {
                flushText();
                segments.push({ kind: 'version', value: token, sep: currentSep });
                hasVersion = true;
            } else if (FEATURE_PAGE && !hasPage && isPageToken(token)) {
                flushText();
                segments.push({ kind: 'page', value: token, sep: currentSep });
                hasPage = true;
            } else if (statusMatch) {
                flushText();
                segments.push({ kind: 'status', value: statusMatch, sep: currentSep });
                hasStatus = true;
            } else {
                if (!textBuffer.length) {
                    textLeadingSep = currentSep;
                } else {
                    textBuffer.push(currentSep);
                }
                textBuffer.push(token);
            }
            i++;
        }
        flushText();
        return segments;
    }

    /* 分散した text セグメント（間に date/status/page/version が挟まる）を 1 つの text にまとめる。
       ただし認識済みセグメント（date/status/page/version）はタイトルへ吸収せず、順序を保って残す。
       text の値どうしのみを結合し、最初の text の位置に置く
       / Merge fragmented text values into one (at the first text's position), while keeping recognized
       segments (date/status/page/version) as their own segments — they are NOT folded into the title */
    function mergeFragmentedText(segments) {
        var textIndices = [];
        for (var i = 0; i < segments.length; i++) {
            if (segments[i].kind === 'text') textIndices.push(i);
        }
        if (textIndices.length <= 1) return segments;
        var firstText = textIndices[0];
        // text 値のみを（各 text の直前区切りで）連結。認識済みセグメントの値は含めない
        var combinedText = segments[firstText].value;
        for (var j = 1; j < textIndices.length; j++) {
            var textSegment = segments[textIndices[j]];
            combinedText += textSegment.sep + textSegment.value;
        }
        var mergedSegments = [];
        for (var k = 0; k < segments.length; k++) {
            if (segments[k].kind === 'text') {
                // 最初の text だけを結合値で置き換え、以降の text は破棄（結合済み）
                if (k === firstText) {
                    mergedSegments.push({ kind: 'text', value: combinedText, sep: segments[firstText].sep });
                }
            } else {
                mergedSegments.push(segments[k]);
            }
        }
        return mergedSegments;
    }

    /* parseFileName の kind を SEGMENT_ORDER の kind に変換しつつ、現在のファイル名で
       実際に出現した要素の並びを返す。重複・FEATURE_STATUS=false の status は除外
       / Map parsed segment kinds to SEGMENT_ORDER kinds in appearance order;
       dedupe and drop 'status' when FEATURE_STATUS is false */
    function deriveOrderFromSegments(segments) {
        var kindToOrder = {
            base: 'base',
            text: 'title',
            status: 'status',
            date: 'timestamp',
            page: 'page',
            version: 'version'
        };
        var order = [];
        var seen = {};
        for (var i = 0; i < segments.length; i++) {
            var orderKind = kindToOrder[segments[i].kind];
            if (!orderKind || seen[orderKind]) continue;
            if (orderKind === 'status' && !FEATURE_STATUS) continue;
            if (orderKind === 'page' && !FEATURE_PAGE) continue;
            order.push(orderKind);
            seen[orderKind] = true;
        }
        return order;
    }

    /* baseOrder（検出された並び）に含まれない SEGMENT_ORDER の kind を、正規位置に補完して返す。
       連番（page）は元ファイル名に現れないため常に欠落し、ステータスも未検出だと欠落する。
       これらを正規位置へ挿入することで、有効化されているのに出力から漏れるのを防ぐ
       （無効な要素は buildFinalName が空値としてスキップするので害はない）
       / Fill in SEGMENT_ORDER kinds missing from a detected order, inserting each at its
       canonical position so enabled-but-undetected segments (page/status) aren't dropped.
       Inactive kinds are harmless because buildFinalName skips empty values. */
    function fillMissingSegmentOrder(baseOrder) {
        var filledOrder = baseOrder.slice();
        for (var i = 0; i < SEGMENT_ORDER.length; i++) {
            var kind = SEGMENT_ORDER[i];
            if (arrayContains(filledOrder, kind)) continue;
            // SEGMENT_ORDER 上で直前にあり、かつ filledOrder に既にある kind の直後へ挿入
            var insertAt = filledOrder.length;
            for (var j = i - 1; j >= 0; j--) {
                var foundIndex = -1;
                for (var k = 0; k < filledOrder.length; k++) {
                    if (filledOrder[k] === SEGMENT_ORDER[j]) { foundIndex = k; break; }
                }
                if (foundIndex !== -1) { insertAt = foundIndex + 1; break; }
            }
            filledOrder.splice(insertAt, 0, kind);
        }
        return filledOrder;
    }

    /* segments から最初に出現する kind の value を取得 / Get the value of the first segment of the given kind */
    function getFirstSegmentValue(segments, kind) {
        for (var i = 0; i < segments.length; i++) {
            if (segments[i].kind === kind) return segments[i].value;
        }
        return '';
    }

    /* 指定 kind の segment が存在するか / Whether a segment of the given kind exists */
    function hasSegmentKind(segments, kind) {
        for (var i = 0; i < segments.length; i++) {
            if (segments[i].kind === kind) return true;
        }
        return false;
    }

    /* 文字列を左 0 パディング / Left-pad a string with zeros to the given width */
    function padLeft(text, width) {
        while (text.length < width) text = '0' + text;
        return text;
    }

    /* 今日の日付を返す。partSeparator で日付内区切り、withTime=true で末尾に "-HHMM" を付加
       / Today's date; with `partSeparator` between Y/M/D parts; `withTime=true` appends "-HHMM" */
    function todayTimestamp(partSeparator, withTime) {
        partSeparator = partSeparator || '';
        var now = new Date();
        var dateText = String(now.getFullYear()) + partSeparator +
            padLeft(String(now.getMonth() + 1), 2) + partSeparator +
            padLeft(String(now.getDate()), 2);
        if (!withTime) return dateText;
        return dateText + '-' + padLeft(String(now.getHours()), 2) + padLeft(String(now.getMinutes()), 2);
    }

    /* 現在のファイル名（拡張子なし）の v 番号だけを桁数維持で +1。
       v 番号が無ければ末尾に "-v2" を付与
       / Bump the v-number inside the current basename in place (preserving digit width).
       Appends "-v2" if no v-number exists. */
    function bumpVersionInPlace(currentBaseName) {
        var versionMatch = String(currentBaseName).match(/([vV])(\d+)/);
        if (versionMatch) {
            var letter = versionMatch[1];
            var digits = versionMatch[2];
            var nextNumber = parseInt(digits, 10) + 1;
            var newDigits = padLeft(String(nextNumber), digits.length);
            return currentBaseName.replace(/([vV])\d+/, letter + newDigits);
        }
        return currentBaseName + '-v2';
    }

    /* バージョン文字列を +1。mode='padded' で 2 桁ゼロ埋め、'paddedWide' で 3 桁ゼロ埋め。
       元バージョンが無ければ新規付与（v1 / v01 / v001）
       / Bump the version string by +1. 'padded'=min 2 digits, 'paddedWide'=min 3. Returns v1/v01/v001 if no original */
    function formatVersion(originalVersion, mode) {
        var versionMatch = String(originalVersion || '').match(/^([vV])(\d+)$/);
        var letter = versionMatch ? versionMatch[1] : 'v';
        var nextNumber = versionMatch ? (parseInt(versionMatch[2], 10) + 1) : 1;
        if (mode === 'padded') {
            var width = versionMatch ? Math.max(versionMatch[2].length, 2) : 2;
            return letter + padLeft(String(nextNumber), width);
        }
        if (mode === 'paddedWide') {
            var wideWidth = versionMatch ? Math.max(versionMatch[2].length, 3) : 3;
            return letter + padLeft(String(nextNumber), wideWidth);
        }
        return letter + String(nextNumber);
    }

    /* baseName 内の最初の v+数字パターンを抽出。マッチしなければ null
       / Extract the first v+digits in baseName; null if no v-number */
    function extractVersionParts(baseName) {
        var versionMatch = String(baseName).match(/^(.*?)([vV])(\d+)(.*)$/);
        if (!versionMatch) return null;
        return { prefix: versionMatch[1], letter: versionMatch[2], digits: versionMatch[3], suffix: versionMatch[4] };
    }

    /* baseName 内の最初の {prefix}+数字パターンを抽出。マッチしなければ null
       prefix は "page" などユーザー指定可能（空文字なら null）
       / Extract the first {prefix} + digits in baseName; null if no match (or empty prefix) */
    function extractPageParts(baseName, prefix) {
        var pagePrefix = String(prefix || '');
        if (!pagePrefix) return null;
        var pagePattern = new RegExp('^(.*?)(' + escapeRegExp(pagePrefix) + ')(\\d+)(.*)$', 'i');
        var pageMatch = String(baseName).match(pagePattern);
        if (!pageMatch) return null;
        return { prefix: pageMatch[1], digits: pageMatch[3], suffix: pageMatch[4] };
    }

    /* 正規表現エスケープ / Escape for use in RegExp */
    function escapeRegExp(text) {
        return String(text).replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&');
    }

    /* baseName と同じ prefix / suffix を持つファイルを folder から探し、最大の v 番号を返す
       / Scan folder for files matching baseName's v-pattern; return max v-number or null */
    function findMaxVersionInFolder(baseName, folder, extension) {
        if (!folder) return null;
        var versionParts = extractVersionParts(baseName);
        if (!versionParts) return null;
        var versionPattern = new RegExp(
            '^' + escapeRegExp(versionParts.prefix) + '[vV](\\d+)'
            + escapeRegExp(versionParts.suffix) + escapeRegExp(extension) + '$',
            'i'
        );
        var folderFiles;
        try { folderFiles = folder.getFiles(); } catch (e) { return null; }
        var maxNumber = -1;
        for (var i = 0; i < folderFiles.length; i++) {
            if (!(folderFiles[i] instanceof File)) continue;
            var fileName = decodePercentEncoded(folderFiles[i].name);
            var fileMatch = fileName.match(versionPattern);
            if (!fileMatch) continue;
            var number = parseInt(fileMatch[1], 10);
            if (number > maxNumber) maxNumber = number;
        }
        return maxNumber >= 0 ? maxNumber : null;
    }

    /* 親フォルダー内の同パターンファイルの最大 v 番号 +1 に置き換える。
       無い、または +1 が現在値以下のときは baseName のまま返す。桁数は維持
       / Replace v-number with (folder max + 1); preserves digit width */
    function nextAvailableVersionName(baseName, folder, extension) {
        var maxNumber = findMaxVersionInFolder(baseName, folder, extension);
        if (maxNumber === null) return baseName;
        var versionParts = extractVersionParts(baseName);
        if (!versionParts) return baseName;
        var currentNumber = parseInt(versionParts.digits, 10);
        var targetNumber = maxNumber + 1;
        if (targetNumber <= currentNumber) return baseName;
        var width = Math.max(versionParts.digits.length, String(targetNumber).length);
        return versionParts.prefix + versionParts.letter + padLeft(String(targetNumber), width) + versionParts.suffix;
    }

    /* 親フォルダー内で baseName の連番（{prefix}NN）と同じ prefix / suffix を持つ最大連番を返す
       / Scan folder for files matching baseName's sequence pattern; return max sequence number or null */
    function findMaxPageInFolder(baseName, folder, extension, prefix) {
        if (!folder) return null;
        var pagePrefix = String(prefix || '');
        if (!pagePrefix) return null;
        var pageParts = extractPageParts(baseName, pagePrefix);
        if (!pageParts) return null;
        var pagePattern = new RegExp(
            '^' + escapeRegExp(pageParts.prefix) + escapeRegExp(pagePrefix) + '(\\d+)'
            + escapeRegExp(pageParts.suffix) + escapeRegExp(extension) + '$',
            'i'
        );
        var folderFiles;
        try { folderFiles = folder.getFiles(); } catch (e) { return null; }
        var maxNumber = -1;
        for (var i = 0; i < folderFiles.length; i++) {
            if (!(folderFiles[i] instanceof File)) continue;
            var fileName = decodePercentEncoded(folderFiles[i].name);
            var fileMatch = fileName.match(pagePattern);
            if (!fileMatch) continue;
            var number = parseInt(fileMatch[1], 10);
            if (number > maxNumber) maxNumber = number;
        }
        return maxNumber >= 0 ? maxNumber : null;
    }

    /* baseName 内の連番を、親フォルダー内の同パターンの最大連番 +1 に置き換える
       / Replace sequence number with (folder max + 1); preserves digit width */
    function nextAvailablePageName(baseName, folder, extension, prefix) {
        var pagePrefix = String(prefix || '');
        if (!pagePrefix) return baseName;
        var maxNumber = findMaxPageInFolder(baseName, folder, extension, pagePrefix);
        if (maxNumber === null) return baseName;
        var pageParts = extractPageParts(baseName, pagePrefix);
        if (!pageParts) return baseName;
        var currentNumber = parseInt(pageParts.digits, 10);
        var targetNumber = maxNumber + 1;
        if (targetNumber <= currentNumber) return baseName;
        var width = Math.max(pageParts.digits.length, String(targetNumber).length);
        return pageParts.prefix + pagePrefix + padLeft(String(targetNumber), width) + pageParts.suffix;
    }

    /* Windows 予約名（拡張子の有無を問わず使用不可）。一致したら末尾に "_" を足してエスケープ
       / Windows-reserved basenames (regardless of extension); append "_" to escape them */
    var WINDOWS_RESERVED_NAMES = {
        CON: 1, PRN: 1, AUX: 1, NUL: 1,
        COM1: 1, COM2: 1, COM3: 1, COM4: 1, COM5: 1,
        COM6: 1, COM7: 1, COM8: 1, COM9: 1,
        LPT1: 1, LPT2: 1, LPT3: 1, LPT4: 1, LPT5: 1,
        LPT6: 1, LPT7: 1, LPT8: 1, LPT9: 1
    };

    /* baseName が Windows 予約名と衝突するなら末尾に "_" を足して回避（大文字小文字無視）
       / If baseName matches a Windows reserved name, append "_" to avoid the collision */
    function escapeWindowsReserved(baseName) {
        var upperName = String(baseName).toUpperCase();
        return (WINDOWS_RESERVED_NAMES[upperName] === 1) ? baseName + '_' : baseName;
    }

    /* % エンコードを 1 回デコード / Decode a percent-encoded string once (best-effort) */
    function decodePercentEncoded(text) {
        text = String(text);
        if (text.indexOf('%') === -1) return text;
        try {
            return decodeURIComponent(text);
        } catch (e) {
            return text;
        }
    }

    /* 拡張子を除去 / Strip a trailing file extension */
    function stripExtension(name) {
        var dotIndex = name.lastIndexOf('.');
        return (dotIndex > 0) ? name.substring(0, dotIndex) : name;
    }

    /* 連続する区切り（- _ .）を 1 つに圧縮し、先頭・末尾の区切りと空白をトリム。
       sep は最終的な区切り文字（'-' / '_' / '' のいずれか、'' のときは混在をそのまま残す）
       / Collapse runs of separators (-, _, .) to one and trim leading/trailing separators + spaces */
    function collapseAndTrimSeparators(text, separator) {
        var collapsed = String(text);
        if (separator === '-' || separator === '_') {
            // 異種の混在も含めて連続した区切りを separator 1 つに統一（YYYY-MM-DD 等は呼び出し前に通過済み想定）
            collapsed = collapsed.replace(/[-_.]{2,}/g, separator);
        } else {
            // 区切り未指定でも同種連続だけは圧縮（-- → -, __ → _, .. → .）
            collapsed = collapsed.replace(/--+/g, '-').replace(/__+/g, '_').replace(/\.\.+/g, '.');
        }
        // 先頭末尾の区切り・空白を除去
        collapsed = collapsed.replace(/^[-_.\s]+|[-_.\s]+$/g, '');
        return collapsed;
    }

    /* 文字列の UTF-8 バイト長（拡張子込みのファイル名サイズ判定用）。ES3 / ExtendScript 向け
       / UTF-8 byte length of a string (for file-name size checks); ES3-safe */
    function byteLengthUTF8(text) {
        text = String(text);
        var byteCount = 0;
        for (var i = 0; i < text.length; i++) {
            var code = text.charCodeAt(i);
            if (code < 0x80) byteCount += 1;
            else if (code < 0x800) byteCount += 2;
            else if (code >= 0xD800 && code <= 0xDBFF) { byteCount += 4; i++; } // サロゲートペア（4 バイト）
            else byteCount += 3;
        }
        return byteCount;
    }

    /* 前後空白をトリム。FEATURE_CLEAN=false のときは OS 禁止文字も除去（後段で処理されないため）
       / Trim whitespace. When FEATURE_CLEAN is off, also strip OS-invalid chars (no post-process safety net) */
    function sanitizeFilename(text) {
        var trimmed = String(text).replace(/^\s+|\s+$/g, '');
        if (FEATURE_CLEAN) return trimmed;
        return trimmed.replace(/[\\\/:*?"<>|]/g, '');
    }

    /* 濁点（U+3099）結合マップ: 合成可能な base → 合成済み濁音
       / Base → voiced composed form (combining dakuten U+3099) */
    var NFC_DAKUTEN_MAP = {
        // ひらがな
        'か': 'が', 'き': 'ぎ', 'く': 'ぐ', 'け': 'げ', 'こ': 'ご',
        'さ': 'ざ', 'し': 'じ', 'す': 'ず', 'せ': 'ぜ', 'そ': 'ぞ',
        'た': 'だ', 'ち': 'ぢ', 'つ': 'づ', 'て': 'で', 'と': 'ど',
        'は': 'ば', 'ひ': 'び', 'ふ': 'ぶ', 'へ': 'べ', 'ほ': 'ぼ',
        'う': 'ゔ',
        // カタカナ
        'カ': 'ガ', 'キ': 'ギ', 'ク': 'グ', 'ケ': 'ゲ', 'コ': 'ゴ',
        'サ': 'ザ', 'シ': 'ジ', 'ス': 'ズ', 'セ': 'ゼ', 'ソ': 'ゾ',
        'タ': 'ダ', 'チ': 'ヂ', 'ツ': 'ヅ', 'テ': 'デ', 'ト': 'ド',
        'ハ': 'バ', 'ヒ': 'ビ', 'フ': 'ブ', 'ヘ': 'ベ', 'ホ': 'ボ',
        'ウ': 'ヴ', 'ワ': 'ヷ', 'ヰ': 'ヸ', 'ヱ': 'ヹ', 'ヲ': 'ヺ'
    };

    /* 半濁点（U+309A）結合マップ: 合成可能な base → 合成済み半濁音（ハ行のみ）
       / Base → semi-voiced composed form (combining handakuten U+309A; ha-row only) */
    var NFC_HANDAKUTEN_MAP = {
        // ひらがな
        'は': 'ぱ', 'ひ': 'ぴ', 'ふ': 'ぷ', 'へ': 'ぺ', 'ほ': 'ぽ',
        // カタカナ
        'ハ': 'パ', 'ヒ': 'ピ', 'フ': 'プ', 'ヘ': 'ペ', 'ホ': 'ポ'
    };

    /* HFS+/APFS 由来の NFD 文字列（か + ゛など）を NFC（が）に結合。
       合成可能な組み合わせのみ変換し、合成不能な濁点・半濁点（例: あ゙）は元の文字列を維持する
       / Normalize NFD-style kana + combining marks to composed NFC form (map-based).
       Only combinable pairs are converted; non-combinable sequences are kept as-is */
    function normalizeNFC(text) {
        return String(text).replace(/(.)([゙゚])/g, function (_, base, mark) {
            var markMap = (mark === '゙') ? NFC_DAKUTEN_MAP : NFC_HANDAKUTEN_MAP;
            var composed = markMap[base];
            return (typeof composed === 'string') ? composed : base + mark;
        });
    }

    /* ファイル名で安全に使える「標準」文字か。
       OS 禁止文字（\ / : * ? " < > |）と半角・全角スペースは ASCII / CJK 範囲だが除外。
       ASCII printable / CJK 記号 / 仮名 / 漢字 / 全角 ASCII / 半角カタカナ を許容
       / Whether a code point is "standard" for filenames; OS-invalid chars and spaces are excluded */
    function isStandardFilenameChar(code) {
        // OS 禁止文字: \ 5C / 2F : 3A * 2A ? 3F " 22 < 3C > 3E | 7C
        if (code === 0x5C || code === 0x2F || code === 0x3A || code === 0x2A
            || code === 0x3F || code === 0x22 || code === 0x3C
            || code === 0x3E || code === 0x7C) return false;
        // 半角・全角スペース
        if (code === 0x20 || code === 0x3000) return false;
        if (code >= 0x21 && code <= 0x7E) return true;     // ASCII printable（space を除く）
        if (code >= 0x3001 && code <= 0x303F) return true; // CJK 記号と句読点（、。「」など / 全角スペースを除く）
        if (code >= 0x3040 && code <= 0x309F) return true; // ひらがな
        if (code >= 0x30A0 && code <= 0x30FF) return true; // カタカナ
        if (code >= 0x3400 && code <= 0x4DBF) return true; // CJK 拡張 A
        if (code >= 0x4E00 && code <= 0x9FFF) return true; // CJK 基本
        if (code >= 0xAC00 && code <= 0xD7A3) return true; // ハングル音節
        if (code >= 0xFF00 && code <= 0xFF5E) return true; // 全角 ASCII
        if (code >= 0xFF61 && code <= 0xFF9F) return true; // 半角カタカナ
        return false;
    }

    /* OS 禁止文字 + 絵文字 + 機種依存文字 + スペースを mode に応じて置換／削除。
       連続するスペース（タブ・改行含む）は事前に 1 つに畳んでから処理。サロゲートペアは 2 文字単位
       mode: 'remove'（削除） / 'dash'（-）/ 'underscore'（_）
       / Clean OS-invalid + emoji + platform-dependent chars + spaces per mode */
    function cleanFilenameChars(text, mode) {
        var replacement = (mode === 'dash') ? '-' : (mode === 'underscore') ? '_' : '';
        // 連続する空白を 1 つにまとめてから per-char で置換（mode='dash' で "  " が "--" にならないように）
        var spaceCollapsed = String(text).replace(/\s+/g, ' ');
        var cleaned = '';
        var i = 0;
        while (i < spaceCollapsed.length) {
            var code = spaceCollapsed.charCodeAt(i);
            if (code >= 0xD800 && code <= 0xDBFF && i + 1 < spaceCollapsed.length) {
                // 高サロゲート + 低サロゲートは常に非標準（BMP 外の絵文字など）
                cleaned += replacement;
                i += 2;
                continue;
            }
            if (isStandardFilenameChar(code)) {
                cleaned += spaceCollapsed.charAt(i);
            } else {
                cleaned += replacement;
            }
            i++;
        }
        return cleaned;
    }

    /* 半角カタカナ → 全角カタカナのマップ（清音）/ Half-width → full-width katakana (unvoiced) */
    var HALFWIDTH_KANA_MAP = {
        'ｦ': 'ヲ', 'ｧ': 'ァ', 'ｨ': 'ィ', 'ｩ': 'ゥ', 'ｪ': 'ェ', 'ｫ': 'ォ',
        'ｬ': 'ャ', 'ｭ': 'ュ', 'ｮ': 'ョ', 'ｯ': 'ッ', 'ｰ': 'ー',
        'ｱ': 'ア', 'ｲ': 'イ', 'ｳ': 'ウ', 'ｴ': 'エ', 'ｵ': 'オ',
        'ｶ': 'カ', 'ｷ': 'キ', 'ｸ': 'ク', 'ｹ': 'ケ', 'ｺ': 'コ',
        'ｻ': 'サ', 'ｼ': 'シ', 'ｽ': 'ス', 'ｾ': 'セ', 'ｿ': 'ソ',
        'ﾀ': 'タ', 'ﾁ': 'チ', 'ﾂ': 'ツ', 'ﾃ': 'テ', 'ﾄ': 'ト',
        'ﾅ': 'ナ', 'ﾆ': 'ニ', 'ﾇ': 'ヌ', 'ﾈ': 'ネ', 'ﾉ': 'ノ',
        'ﾊ': 'ハ', 'ﾋ': 'ヒ', 'ﾌ': 'フ', 'ﾍ': 'ヘ', 'ﾎ': 'ホ',
        'ﾏ': 'マ', 'ﾐ': 'ミ', 'ﾑ': 'ム', 'ﾒ': 'メ', 'ﾓ': 'モ',
        'ﾔ': 'ヤ', 'ﾕ': 'ユ', 'ﾖ': 'ヨ',
        'ﾗ': 'ラ', 'ﾘ': 'リ', 'ﾙ': 'ル', 'ﾚ': 'レ', 'ﾛ': 'ロ',
        'ﾜ': 'ワ', 'ﾝ': 'ン'
    };

    /* 半角カナ + ﾞ の結合マップ（濁音）/ Half-width + dakuten → voiced full-width */
    var HALFWIDTH_KANA_VOICED = {
        'ｶ': 'ガ', 'ｷ': 'ギ', 'ｸ': 'グ', 'ｹ': 'ゲ', 'ｺ': 'ゴ',
        'ｻ': 'ザ', 'ｼ': 'ジ', 'ｽ': 'ズ', 'ｾ': 'ゼ', 'ｿ': 'ゾ',
        'ﾀ': 'ダ', 'ﾁ': 'ヂ', 'ﾂ': 'ヅ', 'ﾃ': 'デ', 'ﾄ': 'ド',
        'ﾊ': 'バ', 'ﾋ': 'ビ', 'ﾌ': 'ブ', 'ﾍ': 'ベ', 'ﾎ': 'ボ',
        'ｳ': 'ヴ'
    };

    /* 半角カナ + ﾟ の結合マップ（半濁音）/ Half-width + handakuten → semi-voiced full-width */
    var HALFWIDTH_KANA_SEMI_VOICED = {
        'ﾊ': 'パ', 'ﾋ': 'ピ', 'ﾌ': 'プ', 'ﾍ': 'ペ', 'ﾎ': 'ポ'
    };

    /* 半角カタカナを全角カタカナへ変換。ｶﾞ → ガ、ﾊﾟ → パ のように濁点・半濁点も結合
       / Convert half-width katakana to full-width; merges trailing dakuten / handakuten */
    function convertHalfwidthKana(text) {
        text = String(text);
        var converted = '';
        var i = 0;
        while (i < text.length) {
            var currentChar = text.charAt(i);
            var nextChar = (i + 1 < text.length) ? text.charAt(i + 1) : '';
            if (nextChar === 'ﾞ' && typeof HALFWIDTH_KANA_VOICED[currentChar] === 'string') {
                converted += HALFWIDTH_KANA_VOICED[currentChar];
                i += 2;
                continue;
            }
            if (nextChar === 'ﾟ' && typeof HALFWIDTH_KANA_SEMI_VOICED[currentChar] === 'string') {
                converted += HALFWIDTH_KANA_SEMI_VOICED[currentChar];
                i += 2;
                continue;
            }
            var fullwidthKana = HALFWIDTH_KANA_MAP[currentChar];
            if (typeof fullwidthKana === 'string') {
                converted += fullwidthKana;
            } else if (currentChar === 'ﾞ') {
                converted += '゛';
            } else if (currentChar === 'ﾟ') {
                converted += '゜';
            } else {
                converted += currentChar;
            }
            i++;
        }
        return converted;
    }

    /* TRANSLITERATE_MAP に従って 1 文字ずつ処理。mode='convert' は変換、'remove' は削除、'keep' は無処理。
       値が文字列のときのみ対象（toString などの prototype プロパティ衝突を回避）
       / Per-char transform via TRANSLITERATE_MAP: 'convert' replaces, 'remove' drops, 'keep' returns as-is */
    function transliterate(text, mode) {
        if (mode === 'keep') return String(text);
        text = String(text);
        var converted = '';
        for (var i = 0; i < text.length; i++) {
            var currentChar = text.charAt(i);
            var replacement = TRANSLITERATE_MAP[currentChar];
            if (typeof replacement === 'string') {
                converted += (mode === 'remove') ? '' : replacement;
            } else {
                converted += currentChar;
            }
        }
        return converted;
    }

    /* 保存先ダイアログ。.ai 拡張子を補正して File を返す。キャンセルは null / Show save dialog, normalize to .ai; null on cancel */
    function pickAiDestination(promptLabel) {
        var chosenFile = File.saveDialog(promptLabel, '*.ai');
        if (!chosenFile) return null;
        var aiFile = File(chosenFile.fsName);
        if (!/\.ai$/i.test(aiFile.name)) {
            aiFile = File(aiFile.parent.fsName + '/' + stripExtension(aiFile.name) + '.ai');
        }
        return aiFile;
    }

    /* プリセット保存用ファイル / Preferences file (key=value lines) */
    function getPrefsFile() {
        return File(Folder.userData.fsName + '/FileNameManager-prefs.txt');
    }

    /* 前回保存したプリセットを読み込み（無ければ空オブジェクト） / Load previously saved prefs (or empty if none) */
    function loadPrefs() {
        var prefsFile = getPrefsFile();
        if (!prefsFile.exists) return {};
        prefsFile.encoding = 'UTF-8';
        prefsFile.open('r');
        var rawText = prefsFile.read();
        prefsFile.close();
        var prefs = {};
        var lines = String(rawText).split('\n');
        for (var i = 0; i < lines.length; i++) {
            // CRLF 保存された prefs でも値末尾に \r が残らないよう行末の改行類を除去
            var line = lines[i].replace(/[\r\n]+$/, '');
            var equalsIndex = line.indexOf('=');
            if (equalsIndex > 0) prefs[line.substring(0, equalsIndex)] = line.substring(equalsIndex + 1);
        }
        return prefs;
    }

    /* プリセットを保存（失敗時は黙って継続） / Save prefs (silently ignore failure) */
    function savePrefs(prefs) {
        try {
            var prefsFile = getPrefsFile();
            prefsFile.encoding = 'UTF-8';
            prefsFile.open('w');
            var lines = [];
            for (var key in prefs) {
                if (prefs.hasOwnProperty(key)) lines.push(key + '=' + prefs[key]);
            }
            prefsFile.write(lines.join('\n'));
            prefsFile.close();
        } catch (e) { /* 保存できなくても処理は止めない */ }
    }

    /* ドキュメントから現在のファイル情報を収集。Illustrator は未保存ドキュメントでも fullName を返すため、
       fullName.exists で保存済みかを判定する
       / Collect file info from the active document; in Illustrator fullName exists even when unsaved,
       so use fullName.exists to detect a saved-to-disk file */
    function gatherDocumentInfo(doc) {
        var fullName;
        var savedToDisk = false;
        try {
            fullName = doc.fullName;
            savedToDisk = !!(fullName && fullName.exists);
        } catch (e) { /* unsaved document */ }
        var rawName = (fullName && fullName.name) ? fullName.name : doc.name;
        var currentName = decodePercentEncoded(rawName);
        var parentFolderName = '';
        var grandparentFolderName = '';
        if (savedToDisk && fullName.parent) {
            parentFolderName = decodePercentEncoded(fullName.parent.name);
            if (fullName.parent.parent) {
                grandparentFolderName = decodePercentEncoded(fullName.parent.parent.name);
            }
        }
        return {
            currentName: currentName,
            baseName: stripExtension(currentName),
            fsPath: savedToDisk ? fullName.fsName : null,
            folder: savedToDisk ? fullName.parent : null,
            parentFolderName: parentFolderName,
            grandparentFolderName: grandparentFolderName
        };
    }

    /* 保存先フォルダを確定（未保存なら saveDialog で取得）。キャンセル時は null / Resolve the target folder (prompt if needed); null on cancel */
    function ensureTargetFolder(folder) {
        if (folder) return folder;
        var pickedFile = pickAiDestination(L('message.chooseDestination'));
        return pickedFile ? pickedFile.parent : null;
    }

    /* 既存同名（自分自身を除く）の上書き確認。OK なら true / Confirm overwrite for an existing file (excluding self); true if approved */
    function confirmOverwriteIfExists(destFile, originalFsPath) {
        if (!destFile.exists) return true;
        if (destFile.fsName === originalFsPath) return true;
        return confirm(L('message.confirmOverwrite') + '\n\n' + destFile.fsName);
    }

    /* file を ~/.Trash に移動。同名衝突時は連番でユニーク化。成功で true、失敗で false
       / Move file to ~/.Trash, disambiguating by appending a counter. Returns true on success */
    function moveToTrash(sourceFile) {
        if (!sourceFile || !sourceFile.exists) return false;
        var trashFolder = Folder("~/.Trash");
        if (!trashFolder.exists) return false;
        var originalName = sourceFile.name;
        var dotIndex = originalName.lastIndexOf('.');
        var stemName = (dotIndex > 0) ? originalName.substring(0, dotIndex) : originalName;
        var extension = (dotIndex > 0) ? originalName.substring(dotIndex) : '';
        var trashDest = File(trashFolder.fsName + '/' + originalName);
        var counter = 1;
        while (trashDest.exists) {
            trashDest = File(trashFolder.fsName + '/' + stemName + ' ' + counter + extension);
            counter++;
            if (counter > 9999) return false; // 暴走防止
        }
        try {
            if (!sourceFile.copy(trashDest)) return false;
            if (sourceFile.remove()) return true;
            // オリジナルを消せなかった場合はゴミ箱側のコピーを後始末（二重化を防ぐ）
            try { trashDest.remove(); } catch (removeError) { /* 後始末に失敗しても続行 */ }
            return false;
        } catch (e) {
            return false;
        }
    }

    /* 旧ファイルを削除（FEATURE_USE_TRASH=true ならゴミ箱に移動、失敗時は file.remove() にフォールバック）
       / Remove the original file: move to ~/.Trash when enabled, else (or on failure) hard-remove */
    function removeOriginalFile(originalFsPath, destFsPath) {
        if (!originalFsPath || originalFsPath === destFsPath) return;
        var originalFile = File(originalFsPath);
        if (!originalFile.exists) return;
        if (FEATURE_USE_TRASH && moveToTrash(originalFile)) return;
        try { originalFile.remove(); } catch (e) { /* 削除できない場合は黙って継続 */ }
    }

    /* モード別の出力処理（rename / saveAs / copy） / Execute the output according to the selected mode */
    function executeOutput(doc, destFile, mode, originalFsPath) {
        if (mode === 'copy' && originalFsPath) {
            // 現在の変更を元ファイルへ保存してから、物理ファイルとしてコピー
            if (!doc.saved) doc.save();
            var originalFile = File(originalFsPath);
            if (!originalFile.exists || !originalFile.copy(destFile)) {
                throw new Error(L('message.saveFailed') + '\n' + destFile.fsName);
            }
            return;
        }
        // rename / saveAs / 未保存ドキュメントの copy: 新名で保存
        doc.saveAs(destFile, new IllustratorSaveOptions());
        if (mode === 'rename') {
            removeOriginalFile(originalFsPath, destFile.fsName);
        }
    }

    /* segments と UI 状態から最終ファイル名（拡張子なし）を構築。SEGMENT_ORDER に従う
       / Build the final filename from segments and UI state, following SEGMENT_ORDER */
    function buildFinalName(segments, uiState) {
        // 区切り記号: 明示選択があればそれを、無ければ元のファイル名で優勢な区切りを使う
        var defaultSeparator;
        if (uiState.separator === '-' || uiState.separator === '_') {
            defaultSeparator = uiState.separator;
        } else {
            defaultSeparator = dominantSeparator(segments);
        }

        function valueForKind(kind) {
            if (kind === 'base') {
                return sanitizeFilename(uiState.baseText || '');
            }
            if (kind === 'title') {
                if (uiState.titleMode === 'none') return '';
                if (uiState.titleMode === 'parent') return sanitizeFilename(uiState.parentFolderName);
                if (uiState.titleMode === 'grandparent') return sanitizeFilename(uiState.grandparentFolderName);
                return sanitizeFilename(uiState.titleText);
            }
            if (kind === 'status') {
                return uiState.status || '';
            }
            if (kind === 'timestamp') {
                var withTime = (uiState.timestampTime === 'hhmm');
                if (uiState.timestamp === 'date') return todayTimestamp('', withTime);
                if (uiState.timestamp === 'dateDash') return todayTimestamp('-', withTime);
                return '';
            }
            if (kind === 'version') {
                if (uiState.version === 'none') return '';
                return formatVersion(getFirstSegmentValue(segments, 'version'), uiState.version);
            }
            if (kind === 'page') {
                if (uiState.pageEnable !== 'yes') return '';
                var pageWidth = (uiState.pagePad === '3') ? 3 : 2;
                // 既存の連番があれば「既存 +1」を初期候補にする（無ければ 1）。
                // フォルダー内の最大連番 +1 への繰り上げは後段の nextAvailablePageName が担う
                var pageStart = 1;
                var existingPage = getFirstSegmentValue(segments, 'page');
                if (existingPage) {
                    var existingPageMatch = String(existingPage).match(PAGE_TOKEN_RE);
                    if (existingPageMatch) pageStart = parseInt(existingPageMatch[2], 10) + 1;
                }
                return (uiState.pagePrefix || '') + padLeft(String(pageStart), pageWidth);
            }
            return '';
        }

        // ピースごとに区切り統一（YYYY-MM-DD のタイムスタンプだけは内部 "-" を保護）
        var order = (uiState.segmentOrder && uiState.segmentOrder.length) ? uiState.segmentOrder : SEGMENT_ORDER;
        var nameParts = [];
        for (var i = 0; i < order.length; i++) {
            var kind = order[i];
            var segmentValue = valueForKind(kind);
            if (!segmentValue) continue;
            // YYYY-MM-DD（および YYYY-MM-DD-HHMM、YYYYMMDD-HHMM）は内部の "-" を保護
            var isProtectedDate = (kind === 'timestamp')
                && (uiState.timestamp === 'dateDash' || uiState.timestampTime === 'hhmm');
            if (!isProtectedDate && FEATURE_SEPARATOR) {
                // 区切り記号統一。FEATURE_DOT_NORMALIZE が true なら "." も対象
                if (uiState.separator === '-') {
                    segmentValue = FEATURE_DOT_NORMALIZE ? segmentValue.replace(/[_.]/g, '-') : segmentValue.replace(/_/g, '-');
                } else if (uiState.separator === '_') {
                    segmentValue = FEATURE_DOT_NORMALIZE ? segmentValue.replace(/[-.]/g, '_') : segmentValue.replace(/-/g, '_');
                }
            }
            nameParts.push(segmentValue);
        }
        var joinedName = nameParts.join(defaultSeparator);
        // スペース（半角・全角・タブ・改行）の処理は後段の cleanFilenameChars に集約
        return joinedName;
    }

    /* segments から優勢な区切り記号を返す。同数なら "-" / Dominant separator across segments (defaults to "-") */
    function dominantSeparator(segments) {
        var dashes = 0, underscores = 0;
        for (var i = 0; i < segments.length; i++) {
            if (segments[i].sep === '-') dashes++;
            else if (segments[i].sep === '_') underscores++;
        }
        return (underscores > dashes) ? '_' : '-';
    }

    // =========================================
    // ダイアログビルダー / Dialog builders
    // =========================================

    /* パネルの共通設定（向き・整列・余白・子要素間隔） / Apply shared panel settings (orientation, alignment, margins, spacing) */
    function setupPanel(panel, spacing) {
        panel.orientation = 'column';
        panel.alignChildren = ['fill', 'top'];
        panel.alignment = 'fill';
        panel.margins = PANEL_MARGINS;
        panel.spacing = (typeof spacing === 'number') ? spacing : PANEL_SPACING;
    }

    /* ソートパネルを構築（標準順 / 現在のファイル名に準じる / カスタム順 + [順序を編集] ボタン）。
       ボタンは「カスタム順」のときだけ有効。currentOrderAvailable=false なら「現在...」は無効化
       / Build the sort panel: Default / Match Current / Custom + [Edit order] button.
       The button is enabled only for "Custom". When currentOrderAvailable=false, "Match Current" is disabled */
    function buildSortPanel(parent, currentOrderAvailable) {
        var panel = parent.add('panel', undefined, L('panel.sort'));
        setupPanel(panel);
        var sortOffRadio = panel.add('radiobutton', undefined, L('radio.sortOff'));
        sortOffRadio.helpTip = L('tip.sort');
        var sortCurrentRadio = panel.add('radiobutton', undefined, L('radio.sortCurrent'));
        sortCurrentRadio.helpTip = L('tip.sort');
        if (!currentOrderAvailable) sortCurrentRadio.enabled = false;
        // 「カスタム順」ラジオと「順序を編集...」ボタンを同じ行に並べる
        var customRow = panel.add('group');
        customRow.orientation = 'row';
        customRow.alignment = ['fill', 'top'];
        customRow.alignChildren = ['left', 'center'];
        customRow.spacing = 8;
        var sortOnRadio = customRow.add('radiobutton', undefined, L('radio.sortOn'));
        sortOnRadio.helpTip = L('tip.sort');
        var sortButton = customRow.add('button', undefined, L('button.sort'));
        // 並び順の初期値は prefs を見ず、常に「現在のファイル名に準じる」（不可なら「標準順」）に固定
        var initialSort = currentOrderAvailable ? 'current' : 'off';
        sortOffRadio.value = (initialSort === 'off');
        sortCurrentRadio.value = (initialSort === 'current');
        sortOnRadio.value = (initialSort === 'on');
        sortButton.enabled = (initialSort === 'on');
        function syncSortButtonEnabled() {
            sortButton.enabled = sortOnRadio.value;
        }
        // onClick は呼び出し側で wire（refreshPreviews と組み合わせるため）
        return {
            panel: panel,
            sortOffRadio: sortOffRadio,
            sortCurrentRadio: sortCurrentRadio,
            sortOnRadio: sortOnRadio,
            sortButton: sortButton,
            /* 'off' / 'current' / 'on' */
            getSortMode: function () {
                if (sortOnRadio.value) return 'on';
                if (sortCurrentRadio.value) return 'current';
                return 'off';
            },
            isSortOn: function () { return sortOnRadio.value; },
            syncSortButtonEnabled: syncSortButtonEnabled
        };
    }

    /* 「モード」パネルを構築（バージョン番号のみ / 全体）。
       defaultVersionOnly=true なら「バージョンのみ」、それ以外は「全体」を初期選択
       / Build the scope panel (Version Only / Full); defaults to "Version Only" when defaultVersionOnly is true, else "Full" */
    function buildOpModePanel(parent, defaultVersionOnly) {
        var panel = parent.add('panel', undefined, L('panel.opMode'));
        setupPanel(panel);
        var versionOnlyRadio = panel.add('radiobutton', undefined, L('radio.opVersionOnly'));
        var fullRadio = panel.add('radiobutton', undefined, L('radio.opFull'));
        versionOnlyRadio.value = !!defaultVersionOnly;
        fullRadio.value = !defaultVersionOnly;
        return {
            panel: panel,
            versionOnlyRadio: versionOnlyRadio,
            fullRadio: fullRadio,
            isVersionOnly: function () { return versionOnlyRadio.value; },
            getOpMode: function () { return versionOnlyRadio.value ? 'versionOnly' : 'full'; }
        };
    }

    /* モード選択パネルを構築（リネーム / 別名で保存 / コピーを保存） / Build the mode panel (Rename / Save As / Save a Copy) */
    function buildModePanel(parent) {
        var panel = parent.add('panel', undefined, L('panel.mode'));
        setupPanel(panel);
        var renameRadio = panel.add('radiobutton', undefined, L('radio.rename'));
        renameRadio.helpTip = L('tip.rename');
        var saveAsRadio = panel.add('radiobutton', undefined, L('radio.saveAs'));
        saveAsRadio.helpTip = L('tip.saveAs');
        var saveCopyRadio = panel.add('radiobutton', undefined, L('radio.saveCopy'));
        saveCopyRadio.helpTip = L('tip.saveCopy');
        // 初期選択は常に「別名で保存」
        renameRadio.value = false;
        saveAsRadio.value = true;
        saveCopyRadio.value = false;
        return {
            panel: panel,
            renameRadio: renameRadio,
            saveAsRadio: saveAsRadio,
            saveCopyRadio: saveCopyRadio,
            /* 現在選択中のモード（'rename' / 'saveAs' / 'copy'） / Currently selected mode */
            getMode: function () {
                if (renameRadio.value) return 'rename';
                if (saveCopyRadio.value) return 'copy';
                return 'saveAs';
            }
        };
    }

    /* ファイル名パネルを構築（現在名・変更後名のみ） / Build the file-name panel (current / final only) */
    function buildFilenamePanel(parent, currentName) {
        var panel = parent.add('panel', undefined, L('panel.filename'));
        setupPanel(panel);

        var currentNameRow = panel.add('group');
        currentNameRow.orientation = 'row';
        var currentNameLabel = currentNameRow.add('statictext', undefined, labelText('label.currentName'), { justify: 'right' });
        currentNameRow.add('statictext', undefined, currentName);

        var finalNameRow = panel.add('group');
        finalNameRow.orientation = 'row';
        var finalNameLabel = finalNameRow.add('statictext', undefined, labelText('label.finalName'), { justify: 'right' });
        // 「変更後：」は statictext のためレイアウト後にサイズ固定。
        // 現在のファイル名と「入力フィールド + 余白」の大きい方を確保しておく
        var finalNameValue = finalNameRow.add('statictext', undefined, currentName + '.ai');
        var currentNameWidth = panel.graphics.measureString(currentName + '.ai').width;
        finalNameValue.preferredSize.width = Math.max(currentNameWidth + 20, 340);

        // 個別整列はせず、ラベル参照を呼び出し側に返し、後段で全パネル統一整列する
        return {
            panel: panel,
            finalNameValue: finalNameValue,
            labels: [currentNameLabel, finalNameLabel],
            labelTexts: [labelText('label.currentName'), labelText('label.finalName')]
        };
    }

    /* オプションパネルを構築（ベース表示・タイトル選択・タイムスタンプ・バージョン番号・区切り） / Build the options panel */
    function buildOptionsPanel(parent, segments, prefs, parentFolderName, grandparentFolderName) {
        var panel = parent.add('panel', undefined, L('panel.options'));
        setupPanel(panel);

        // ベース: 検出値を初期表示する入力欄。空欄可、prefs には保存しない
        var detectedBase = getFirstSegmentValue(segments, 'base');
        var baseRow = panel.add('group');
        baseRow.orientation = 'row';
        var baseLabel = baseRow.add('statictext', undefined, labelText('label.base'), { justify: 'right' });
        baseLabel.helpTip = L('tip.base');
        var baseField = baseRow.add('edittext', undefined, detectedBase);
        baseField.preferredSize.width = NEW_NAME_FIELD_WIDTH;
        baseField.helpTip = L('tip.base');

        // 元ファイル名のテキスト部（base / status / date / version 以外）をサブテキストとして検出
        var detectedTitle = getFirstSegmentValue(segments, 'text');

        // サブテキスト: 1 行目 = ラベル + 3 ラジオ、2 行目 = 「指定」用の入力欄
        var titleSection = panel.add('group');
        titleSection.orientation = 'column';
        titleSection.alignChildren = ['fill', 'top'];
        titleSection.spacing = 4;

        var titleRow = addRadioRow(titleSection, 'label.title', 'tip.title', [
            { key: 'none', text: L('radio.titleNone') },
            { key: 'custom', text: L('radio.titleCustom') },
            { key: 'parent', text: L('radio.titleParent') },
            { key: 'grandparent', text: L('radio.titleGrandparent') }
        ]);
        titleRow.group.alignment = ['left', 'top'];
        // 親/2 階層上のフォルダー名がある場合は helpTip にフォルダ名を追記、無ければ無効化
        if (parentFolderName) titleRow.radios.parent.helpTip = L('tip.title') + ' (' + parentFolderName + ')';
        else titleRow.radios.parent.enabled = false;
        if (grandparentFolderName) titleRow.radios.grandparent.helpTip = L('tip.title') + ' (' + grandparentFolderName + ')';
        else titleRow.radios.grandparent.enabled = false;

        // 「指定」用の入力欄は次の行（ラベル列幅だけ左に余白を入れて radios に揃える）
        var titleFieldRow = titleSection.add('group');
        titleFieldRow.orientation = 'row';
        titleFieldRow.alignment = ['left', 'top'];
        var titleFieldSpacer = titleFieldRow.add('statictext', undefined, '');
        var titleField = titleFieldRow.add('edittext', undefined, detectedTitle);
        titleField.preferredSize.width = NEW_NAME_FIELD_WIDTH;
        titleField.helpTip = L('tip.title');

        // 初期モード: 元ファイル名にサブテキストが無ければ「なし」を強制。
        // 検出できた場合は prefs.titleMode を優先、無ければ 'custom'（parent/grandparent は対応フォルダ名必須）
        var initialTitleMode = !detectedTitle ? 'none' :
            pickPref(prefs, 'titleMode', ['none', 'parent', 'grandparent', 'custom'], 'custom');
        if (initialTitleMode === 'parent' && !parentFolderName) {
            initialTitleMode = detectedTitle ? 'custom' : 'none';
        }
        if (initialTitleMode === 'grandparent' && !grandparentFolderName) {
            initialTitleMode = detectedTitle ? 'custom' : 'none';
        }
        titleRow.radios.none.value = (initialTitleMode === 'none');
        titleRow.radios.parent.value = (initialTitleMode === 'parent');
        titleRow.radios.grandparent.value = (initialTitleMode === 'grandparent');
        titleRow.radios.custom.value = (initialTitleMode === 'custom');
        titleField.enabled = (initialTitleMode === 'custom');

        function syncTitleFieldEnabled() {
            titleField.enabled = titleRow.radios.custom.value;
        }

        // ステータス（dropdown。先頭は「なし」）
        var statusLabel = null, statusDropdown = null;
        if (FEATURE_STATUS) {
            var statusRow = panel.add('group');
            statusRow.orientation = 'row';
            statusLabel = statusRow.add('statictext', undefined, labelText('label.status'), { justify: 'right' });
            statusLabel.helpTip = L('tip.status');
            statusDropdown = statusRow.add('dropdownlist', undefined, undefined);
            statusDropdown.helpTip = L('tip.status');
            var parsedStatus = getFirstSegmentValue(segments, 'status');
            var initialStatus = parsedStatus || ((prefs && prefs.status) ? prefs.status : '');
            var initialStatusIndex = -1;
            for (var statusIndex = 0; statusIndex < STATUS_ITEMS.length; statusIndex++) {
                var statusItem = STATUS_ITEMS[statusIndex];
                var dropdownItem = statusDropdown.add('item', (currentLanguage === 'ja' ? statusItem.ja : statusItem.en));
                if (isStatusDivider(statusItem)) {
                    dropdownItem.enabled = false;
                    continue;
                }
                if (initialStatusIndex === -1 && statusItem.value === initialStatus) {
                    initialStatusIndex = statusIndex;
                }
            }
            if (initialStatusIndex === -1) initialStatusIndex = 0;
            statusDropdown.selection = initialStatusIndex;
        }

        // タイムスタンプ（なし / YYYYMMDD / YYYY-MM-DD。デフォルト YYYYMMDD）
        // 末尾に「時刻も付与」(HHMM) チェックボックスを同居
        var timestampRow = addRadioRow(panel, 'label.timestamp', 'tip.timestamp', [
            { key: 'none', text: L('radio.timestampNone') },
            { key: 'date', text: L('radio.timestampDate') },
            { key: 'dateDash', text: L('radio.timestampDateDash') }
        ]);
        var initialTimestamp = pickPref(prefs, 'timestamp', ['none', 'date', 'dateDash'], 'date');
        timestampRow.radios.none.value = (initialTimestamp === 'none');
        timestampRow.radios.date.value = (initialTimestamp === 'date');
        timestampRow.radios.dateDash.value = (initialTimestamp === 'dateDash');

        var timestampHHMMCheckbox = timestampRow.group.add('checkbox', undefined, L('radio.timestampWithTime'));
        timestampHHMMCheckbox.helpTip = L('tip.timestampWithTime');
        timestampHHMMCheckbox.value = pickPref(prefs, 'timestampTime', ['no', 'hhmm'], 'no') === 'hhmm';

        function syncTimestampHHMMEnabled() {
            timestampHHMMCheckbox.enabled = !timestampRow.radios.none.value;
        }
        syncTimestampHHMMEnabled();

        // 連番（チェックボックス + プレフィックス入力 + 桁数ラジオ）。デフォルト OFF / "page" / 2 桁
        // SEGMENT_ORDER 内では timestamp の後・version の前に配置
        var pageRow = null;
        var pageCheckbox = null;
        var pagePrefixField = null;
        var pagePadRadio2 = null;
        var pagePadRadio3 = null;
        if (FEATURE_PAGE) {
            // 元ファイル名から連番セグメント（page03 など）を検出。あればプレフィックス・桁数を初期値に反映し、
            // 既定で ON にして既存の連番が出力から消えないようにする（無ければ prefs / 既定値）
            var detectedPage = getFirstSegmentValue(segments, 'page');
            var detectedPageMatch = detectedPage ? String(detectedPage).match(PAGE_TOKEN_RE) : null;
            var detectedPagePrefix = detectedPageMatch ? detectedPageMatch[1] : '';
            var detectedPagePad = detectedPageMatch ? (detectedPageMatch[2].length >= 3 ? '3' : '2') : '';

            pageRow = panel.add('group');
            pageRow.orientation = 'row';
            var pageLabel = pageRow.add('statictext', undefined, labelText('label.page'), { justify: 'right' });
            pageLabel.helpTip = L('tip.page');
            pageCheckbox = pageRow.add('checkbox', undefined, L('radio.pageEnable'));
            pageCheckbox.helpTip = L('tip.page');
            // 既存の連番を検出したら prefs より優先して必ず ON（既存 pageNN を消さない）
            pageCheckbox.value = detectedPage
                ? true
                : (pickPref(prefs, 'pageEnable', ['no', 'yes'], 'no') === 'yes');
            pagePrefixField = pageRow.add('edittext', undefined, '');
            pagePrefixField.preferredSize.width = 60;
            pagePrefixField.helpTip = L('tip.page');
            pagePrefixField.text = detectedPagePrefix
                ? detectedPagePrefix
                : ((prefs && typeof prefs.pagePrefix === 'string') ? prefs.pagePrefix : 'page');
            pagePadRadio2 = pageRow.add('radiobutton', undefined, L('radio.pagePad2'));
            pagePadRadio2.helpTip = L('tip.page');
            pagePadRadio3 = pageRow.add('radiobutton', undefined, L('radio.pagePad3'));
            pagePadRadio3.helpTip = L('tip.page');
            var initialPagePad = detectedPagePad ? detectedPagePad : pickPref(prefs, 'pagePad', ['2', '3'], '2');
            pagePadRadio2.value = (initialPagePad === '2');
            pagePadRadio3.value = (initialPagePad === '3');
            pageRow.label = pageLabel; // alignLabelWidths 用に統一形にしておく
        }

        function syncPageControlsEnabled() {
            if (!pageCheckbox) return;
            var enabled = pageCheckbox.value;
            if (pagePrefixField) pagePrefixField.enabled = enabled;
            if (pagePadRadio2) pagePadRadio2.enabled = enabled;
            if (pagePadRadio3) pagePadRadio3.enabled = enabled;
        }
        syncPageControlsEnabled();

        // バージョン番号（なし / v1, v2… / v01, v02… / v001, v002…。初期値は「なし」）
        // ES3 で 'short' を裸キーにすると予約語エラーになるため、ローカル変数名は short_ にする
        var versionRow = addRadioRow(panel, 'label.version', 'tip.version', [
            { key: 'none', text: L('radio.versionNone') },
            { key: 'short_', text: L('radio.versionShort') },
            { key: 'padded', text: L('radio.versionPadded') },
            { key: 'paddedWide', text: L('radio.versionPaddedWide') }
        ]);
        // デフォルトは「なし」。元ファイルに v 番号がある場合は prefs を優先（無ければ「なし」）
        var hasOriginalVersion = !!getFirstSegmentValue(segments, 'version');
        var initialVersion = hasOriginalVersion
            ? pickPref(prefs, 'version', ['none', 'short', 'padded', 'paddedWide'], 'none')
            : 'none';
        versionRow.radios.none.value = (initialVersion === 'none');
        versionRow.radios.short_.value = (initialVersion === 'short');
        versionRow.radios.padded.value = (initialVersion === 'padded');
        versionRow.radios.paddedWide.value = (initialVersion === 'paddedWide');

        // 区切り記号（変更しない / - / _ 横並び。デフォルト "-"）
        var separatorRow = null;
        if (FEATURE_SEPARATOR) {
            separatorRow = addRadioRow(panel, 'label.separator', 'tip.separator', [
                { key: 'noChange', text: L('radio.noChange') },
                { key: 'dash', text: '-' },
                { key: 'underscore', text: '_' }
            ]);
            // prefs があればそれを優先、無ければ FEATURE_SEPARATOR の指定文字（'-' or '_'）を初期値に
            var defaultSeparator = (FEATURE_SEPARATOR === '_') ? '_' : '-';
            var initialSeparator = pickPref(prefs, 'separator', ['', '-', '_'], defaultSeparator);
            separatorRow.radios.noChange.value = (initialSeparator === '');
            separatorRow.radios.dash.value = (initialSeparator === '-');
            separatorRow.radios.underscore.value = (initialSeparator === '_');
        }

        // 濁点処理（変更しない / 結合する。デフォルト "結合する"）
        var nfcRow = null;
        if (FEATURE_NFC) {
            nfcRow = addRadioRow(panel, 'label.nfc', 'tip.nfc', [
                { key: 'keep', text: L('radio.noChange') },
                { key: 'combine', text: L('radio.nfcCombine') }
            ]);
            var initialNfc = pickPref(prefs, 'nfc', ['keep', 'combine'], 'combine');
            nfcRow.radios.keep.value = (initialNfc === 'keep');
            nfcRow.radios.combine.value = (initialNfc === 'combine');
        }

        // クリーンなファイル名（削除する / -に変更 / _に変更。デフォルト "-に変更"）
        // OS 禁止文字（\/:*?"<>|）と絵文字・機種依存文字をまとめて処理。
        // 末尾に「半角カナ → 全角」チェックボックスを同居（'-' / '_' のときだけ有効）
        var cleanRow = null;
        var halfwidthKanaCheckbox = null;
        if (FEATURE_CLEAN) {
            cleanRow = addRadioRow(panel, 'label.clean', 'tip.clean', [
                { key: 'remove', text: L('radio.cleanRemove') },
                { key: 'dash', text: L('radio.changeToDash') },
                { key: 'underscore', text: L('radio.changeToUnderscore') }
            ]);
            var initialClean = pickPref(prefs, 'clean', ['remove', 'dash', 'underscore'], 'dash');
            cleanRow.radios.remove.value = (initialClean === 'remove');
            cleanRow.radios.dash.value = (initialClean === 'dash');
            cleanRow.radios.underscore.value = (initialClean === 'underscore');

            if (FEATURE_HALFWIDTH_KANA) {
                halfwidthKanaCheckbox = cleanRow.group.add('checkbox', undefined, L('radio.halfwidthKanaConvert'));
                halfwidthKanaCheckbox.helpTip = L('tip.halfwidthKana');
                halfwidthKanaCheckbox.value = pickPref(prefs, 'halfwidthKana', ['keep', 'convert'], 'convert') === 'convert';
            }
        }

        // 法人略記・丸数字（変更しない / 削除する / 変換する。デフォルト "変換する"）
        var translitRow = null;
        if (FEATURE_TRANSLITERATE) {
            translitRow = addRadioRow(panel, 'label.translit', 'tip.translit', [
                { key: 'keep', text: L('radio.noChange') },
                { key: 'remove', text: L('radio.translitRemove') },
                { key: 'convert', text: L('radio.translitConvert') }
            ]);
            var initialTranslit = pickPref(prefs, 'translit', ['keep', 'remove', 'convert'], 'convert');
            translitRow.radios.keep.value = (initialTranslit === 'keep');
            translitRow.radios.remove.value = (initialTranslit === 'remove');
            translitRow.radios.convert.value = (initialTranslit === 'convert');
        }

        // クリーンが '-' / '_' のときだけ「半角カナ → 全角」を有効化
        function syncHalfwidthKanaEnabled() {
            if (!halfwidthKanaCheckbox || !cleanRow) return;
            halfwidthKanaCheckbox.enabled = (cleanRow.radios.dash.value || cleanRow.radios.underscore.value);
        }
        syncHalfwidthKanaEnabled();

        // ラベル幅を統一（FEATURE で UI 非表示の行は除外）
        var labelTexts = [labelText('label.base'), labelText('label.title')];
        var labelControls = [baseLabel, titleRow.label];
        if (FEATURE_STATUS) { labelTexts.push(labelText('label.status')); labelControls.push(statusLabel); }
        labelTexts.push(labelText('label.timestamp')); labelControls.push(timestampRow.label);
        if (FEATURE_PAGE) { labelTexts.push(labelText('label.page')); labelControls.push(pageRow.label); }
        labelTexts.push(labelText('label.version')); labelControls.push(versionRow.label);
        if (FEATURE_SEPARATOR) { labelTexts.push(labelText('label.separator')); labelControls.push(separatorRow.label); }
        if (FEATURE_NFC) { labelTexts.push(labelText('label.nfc')); labelControls.push(nfcRow.label); }
        if (FEATURE_CLEAN) { labelTexts.push(labelText('label.clean')); labelControls.push(cleanRow.label); }
        if (FEATURE_TRANSLITERATE) { labelTexts.push(labelText('label.translit')); labelControls.push(translitRow.label); }
        // 個別整列はせず、呼び出し側に渡してファイル名パネルと統一整列する
        // titleFieldSpacer の幅は createDialog 側で整列後に設定する

        return {
            panel: panel,
            baseField: baseField,
            titleRow: titleRow,
            titleField: titleField,
            titleFieldSpacer: titleFieldSpacer,
            syncTitleFieldEnabled: syncTitleFieldEnabled,
            labels: labelControls,
            labelTexts: labelTexts,
            statusDropdown: statusDropdown,
            timestampRow: timestampRow,
            timestampHHMMCheckbox: timestampHHMMCheckbox,
            syncTimestampHHMMEnabled: syncTimestampHHMMEnabled,
            pageCheckbox: pageCheckbox,
            pagePrefixField: pagePrefixField,
            pagePadRadio2: pagePadRadio2,
            pagePadRadio3: pagePadRadio3,
            syncPageControlsEnabled: syncPageControlsEnabled,
            versionRow: versionRow,
            separatorRow: separatorRow,
            nfcRow: nfcRow,
            cleanRow: cleanRow,
            halfwidthKanaCheckbox: halfwidthKanaCheckbox,
            translitRow: translitRow,
            syncHalfwidthKanaEnabled: syncHalfwidthKanaEnabled,
            /* '' = 変更しない、'-' / '_' = 統一。FEATURE_SEPARATOR=false なら常に '' */
            getSeparator: function () {
                if (!separatorRow) return '';
                if (separatorRow.radios.noChange.value) return '';
                if (separatorRow.radios.dash.value) return '-';
                return '_';
            },
            /* 'keep' = そのまま、'combine' = 濁点・半濁点を NFC 結合。FEATURE_NFC=false なら常に 'keep' */
            getNfc: function () {
                if (!nfcRow) return 'keep';
                return nfcRow.radios.keep.value ? 'keep' : 'combine';
            },
            /* 'remove' / 'dash' / 'underscore'。FEATURE_CLEAN=false なら常に '' を返す（無処理） */
            getClean: function () {
                if (!cleanRow) return '';
                if (cleanRow.radios.dash.value) return 'dash';
                if (cleanRow.radios.underscore.value) return 'underscore';
                return 'remove';
            },
            /* 'keep' = そのまま、'remove' = 削除、'convert' = TRANSLITERATE_MAP で変換。FEATURE_TRANSLITERATE=false なら常に 'keep' */
            getTranslit: function () {
                if (!translitRow) return 'keep';
                if (translitRow.radios.convert.value) return 'convert';
                if (translitRow.radios.remove.value) return 'remove';
                return 'keep';
            },
            /* 'keep' = そのまま、'convert' = 半角カナを全角カナに。FEATURE_HALFWIDTH_KANA=false なら常に 'keep' */
            getHalfwidthKana: function () {
                if (!halfwidthKanaCheckbox) return 'keep';
                return halfwidthKanaCheckbox.value ? 'convert' : 'keep';
            },
            /* 'none' / 'date' / 'dateDash' */
            getTimestamp: function () {
                if (timestampRow.radios.none.value) return 'none';
                if (timestampRow.radios.dateDash.value) return 'dateDash';
                return 'date';
            },
            /* 'no' / 'hhmm'（時刻 HHMM をタイムスタンプ末尾に付与するか）*/
            getTimestampTime: function () {
                return timestampHHMMCheckbox.value ? 'hhmm' : 'no';
            },
            /* 'no' / 'yes'（ページ番号セグメントを付与するか）。FEATURE_PAGE=false なら常に 'no' */
            getPageEnable: function () {
                if (!pageCheckbox) return 'no';
                return pageCheckbox.value ? 'yes' : 'no';
            },
            /* '2' / '3'（連番のゼロ埋め桁数）。FEATURE_PAGE=false なら '2' */
            getPagePad: function () {
                if (!pagePadRadio3) return '2';
                return pagePadRadio3.value ? '3' : '2';
            },
            /* 連番のプレフィックス（例: "page"）。FEATURE_PAGE=false なら '' */
            getPagePrefix: function () {
                if (!pagePrefixField) return '';
                return pagePrefixField.text;
            },
            /* STATUS_ITEMS の value（'' = なし）。FEATURE_STATUS=false なら常に '' */
            getStatus: function () {
                if (!FEATURE_STATUS || !statusDropdown) return '';
                var selectedIndex = statusDropdown.selection ? statusDropdown.selection.index : 0;
                var selectedItem = STATUS_ITEMS[selectedIndex];
                if (!selectedItem || isStatusDivider(selectedItem)) return '';
                return selectedItem.value;
            },
            /* 'none' / 'short' / 'padded' / 'paddedWide' */
            getVersion: function () {
                if (versionRow.radios.none.value) return 'none';
                if (versionRow.radios.short_.value) return 'short';
                if (versionRow.radios.padded.value) return 'padded';
                return 'paddedWide';
            },
            /* 'none' / 'parent' / 'grandparent' / 'custom' */
            getTitleMode: function () {
                if (titleRow.radios.none.value) return 'none';
                if (titleRow.radios.parent.value) return 'parent';
                if (titleRow.radios.grandparent.value) return 'grandparent';
                return 'custom';
            },
            /* ベース入力欄の現在値 / Current value of the base input */
            getBase: function () {
                return baseField.text;
            }
        };
    }

    /* 並び順を ↑↓ で編集するサブダイアログ。OK で新しい順序を返す。キャンセルで null
       / Sub-dialog to reorder segments with ↑↓; returns the new order, or null on cancel */
    function openSortDialog(initialOrder) {
        var sortDialog = new Window('dialog', L('sort.title'));
        sortDialog.opacity = DIALOG_OPACITY;
        sortDialog.orientation = 'column';
        sortDialog.alignChildren = 'fill';
        sortDialog.margins = 15;
        sortDialog.spacing = 10;

        var hintText = sortDialog.add('statictext', undefined, L('sort.hint'));
        hintText.alignment = 'left';

        var bodyGroup = sortDialog.add('group');
        bodyGroup.orientation = 'row';
        bodyGroup.alignChildren = ['fill', 'fill'];

        var order = initialOrder.slice();
        var orderList = bodyGroup.add('listbox', undefined, []);
        orderList.preferredSize = [200, 140];
        for (var i = 0; i < order.length; i++) {
            orderList.add('item', L('label.' + order[i]));
        }
        orderList.selection = 0;

        var buttonColumn = bodyGroup.add('group');
        buttonColumn.orientation = 'column';
        buttonColumn.alignChildren = 'fill';
        var upButton = buttonColumn.add('button', undefined, '↑');
        upButton.preferredSize = [36, 24];
        var downButton = buttonColumn.add('button', undefined, '↓');
        downButton.preferredSize = [36, 24];

        function refreshList(newIndex) {
            for (var i = 0; i < order.length; i++) {
                orderList.items[i].text = L('label.' + order[i]);
            }
            orderList.selection = newIndex;
        }
        upButton.onClick = function () {
            var selectedIndex = orderList.selection ? orderList.selection.index : -1;
            if (selectedIndex <= 0) return;
            var swapTemp = order[selectedIndex - 1];
            order[selectedIndex - 1] = order[selectedIndex];
            order[selectedIndex] = swapTemp;
            refreshList(selectedIndex - 1);
        };
        downButton.onClick = function () {
            var selectedIndex = orderList.selection ? orderList.selection.index : -1;
            if (selectedIndex < 0 || selectedIndex >= order.length - 1) return;
            var swapTemp = order[selectedIndex + 1];
            order[selectedIndex + 1] = order[selectedIndex];
            order[selectedIndex] = swapTemp;
            refreshList(selectedIndex + 1);
        };

        var buttonRow = sortDialog.add('group');
        buttonRow.alignment = ['center', 'top'];
        buttonRow.add('button', undefined, L('button.cancel'), { name: 'cancel' });
        buttonRow.add('button', undefined, 'OK', { name: 'ok' });

        if (sortDialog.show() !== 1) return null;
        return order;
    }

    /* ES3 互換: 配列に値が含まれるか / ES3-safe array contains check */
    function arrayContains(array, value) {
        for (var i = 0; i < array.length; i++) {
            if (array[i] === value) return true;
        }
        return false;
    }

    /* prefs[key] が validValues 内ならそれを、そうでなければ fallback を返す
       / Returns prefs[key] if it's in validValues, else fallback */
    function pickPref(prefs, key, validValues, fallback) {
        var value = prefs && prefs[key];
        return arrayContains(validValues, value) ? value : fallback;
    }

    /* 保存された並び順文字列を妥当性チェックして配列で返す。不正なら null
       / Parse a comma-separated segment order; returns null if invalid */
    function parseSegmentOrderPref(value) {
        if (!value) return null;
        var parts = String(value).split(',');
        if (parts.length !== SEGMENT_ORDER.length) return null;
        var seen = {};
        for (var i = 0; i < parts.length; i++) {
            var kind = parts[i];
            if (!arrayContains(SEGMENT_ORDER, kind) || seen[kind]) return null;
            seen[kind] = true;
        }
        return parts;
    }

    /* labelTexts の最大幅に controls の幅を揃える。checkbox/radiobutton が含まれていれば
       インジケーター分（+20px）を全コントロールに加算して右端を揃える
       / Align controls' widths to the widest label. If any control is a checkbox/radiobutton,
       add indicator width (+20px) to all so their right edges align */
    function alignLabelWidths(panel, labelTexts, controls) {
        var graphics = panel.graphics;
        var maxWidth = 0;
        for (var i = 0; i < labelTexts.length; i++) {
            var width = graphics.measureString(labelTexts[i]).width;
            if (width > maxWidth) maxWidth = width;
        }
        maxWidth += 12;
        var needsIndicator = false;
        for (var k = 0; k < controls.length; k++) {
            if (controls[k].type === 'checkbox' || controls[k].type === 'radiobutton') {
                needsIndicator = true;
                break;
            }
        }
        if (needsIndicator) maxWidth += 20;
        for (var j = 0; j < controls.length; j++) {
            controls[j].preferredSize = [maxWidth, controls[j].preferredSize.height || 20];
        }
    }

    /* ラベル + ラジオ群を 1 行追加して { group, label, radios: { key: radio, ... } } を返す。
       radioDefs: [{ key, text }, ...]。すべてのコントロールに L(tipKey) の helpTip を設定
       / Add a label + radio row; returns { group, label, radios } */
    function addRadioRow(panel, labelKey, tipKey, radioDefs) {
        var row = panel.add('group');
        row.orientation = 'row';
        var tipText = L(tipKey);
        // ラベルは右揃え（alignLabelWidths で固定列幅になるため右端が入力欄側に揃う）
        var label = row.add('statictext', undefined, labelText(labelKey), { justify: 'right' });
        label.helpTip = tipText;
        var radios = {};
        for (var i = 0; i < radioDefs.length; i++) {
            var radioDef = radioDefs[i];
            var radio = row.add('radiobutton', undefined, radioDef.text);
            radio.helpTip = tipText;
            radios[radioDef.key] = radio;
        }
        return { group: row, label: label, radios: radios };
    }

    /* 各コントロールの種類に応じた変更イベントに callback を割り当てる。null/undefined はスキップ。
       radiobutton/checkbox → onClick、edittext → onChanging、dropdownlist → onChange
       / Wire a refresh callback to each control's appropriate change event */
    function wireRefresh(callback, controls) {
        for (var i = 0; i < controls.length; i++) {
            var control = controls[i];
            if (!control) continue;
            var controlType = control.type;
            if (controlType === 'radiobutton' || controlType === 'checkbox') control.onClick = callback;
            else if (controlType === 'edittext') control.onChanging = callback;
            else if (controlType === 'dropdownlist') control.onChange = callback;
        }
    }

    /* ダイアログ全体を組み立て、イベント配線とプレビューを行う / Compose the full dialog, wire events, and run live preview */
    function createDialog(segments, currentName, prefs, parentFolderName, grandparentFolderName, folder) {
        var dialog = new Window('dialog', L('dialog.title') + ' ' + SCRIPT_VERSION);
        dialog.opacity = DIALOG_OPACITY;
        dialog.orientation = 'column';
        dialog.alignChildren = 'fill';

        // 上部: 動作パネル + モードパネル + ソートパネル（FEATURE_SORT=false ならソートパネル無し）
        var topRow = dialog.add('group');
        topRow.orientation = 'row';
        topRow.alignChildren = ['fill', 'top'];
        var mode = buildModePanel(topRow);
        // 元ファイル名が「タイムスタンプ無し」かつ「バージョン番号あり（-vN）」なら、初期モードを「バージョンのみ」に
        var defaultVersionOnly = !hasSegmentKind(segments, 'date') && hasSegmentKind(segments, 'version');
        var opMode = buildOpModePanel(topRow, defaultVersionOnly);
        // 現在のファイル名から検出した並び順（要素ゼロなら「現在に準じる」は無効化）
        var currentOrder = FEATURE_SORT ? deriveOrderFromSegments(segments) : [];
        var sort = FEATURE_SORT ? buildSortPanel(topRow, currentOrder.length > 0) : null;

        // 並び順カスタム値（prefs に保存されたものを採用、不正・未保存ならデフォルト）
        var customOrder = FEATURE_SORT
            ? (parseSegmentOrderPref(prefs && prefs.segmentOrder) || SEGMENT_ORDER.slice())
            : SEGMENT_ORDER.slice();

        var filename = buildFilenamePanel(dialog, currentName);
        var options = buildOptionsPanel(dialog, segments, prefs, parentFolderName, grandparentFolderName);

        // 2 つのパネルのラベル幅を統一整列（測定基準は options.panel）
        var combinedTexts = filename.labelTexts.concat(options.labelTexts);
        var combinedLabels = filename.labels.concat(options.labels);
        alignLabelWidths(options.panel, combinedTexts, combinedLabels);
        // サブテキスト 2 行目「指定」入力欄の左余白を統一後のラベル列幅に合わせる
        options.titleFieldSpacer.preferredSize = [options.titleRow.label.preferredSize.width, 1];

        // ---- ライブプレビュー ----
        function currentUIState() {
            var sortMode = (FEATURE_SORT && sort) ? sort.getSortMode() : 'off';
            var segmentOrder;
            // 'current' は検出順のみだと連番/未検出ステータスが欠落するため正規位置に補完する
            if (sortMode === 'on') segmentOrder = customOrder;
            else if (sortMode === 'current') segmentOrder = fillMissingSegmentOrder(currentOrder);
            else segmentOrder = SEGMENT_ORDER;
            return {
                opMode: opMode.getOpMode(),
                baseText: options.getBase(),
                titleMode: options.getTitleMode(),
                titleText: options.titleField.text,
                parentFolderName: parentFolderName,
                grandparentFolderName: grandparentFolderName,
                status: options.getStatus(),
                timestamp: options.getTimestamp(),
                timestampTime: options.getTimestampTime(),
                pageEnable: options.getPageEnable(),
                pagePad: options.getPagePad(),
                pagePrefix: options.getPagePrefix(),
                version: options.getVersion(),
                separator: options.getSeparator(),
                nfc: options.getNfc(),
                translit: options.getTranslit(),
                clean: options.getClean(),
                halfwidthKana: options.getHalfwidthKana(),
                sort: sortMode,
                customSegmentOrder: customOrder.slice(),
                segmentOrder: segmentOrder
            };
        }

        // 「バージョン番号のみ」モード時に隠す UI（ソート + ファイル名の設定）
        function syncOpModeVisibility() {
            var versionOnly = opMode.isVersionOnly();
            if (sort) sort.panel.visible = !versionOnly;
            options.panel.visible = !versionOnly;
            dialog.layout.layout(true);
            dialog.layout.resize();
        }

        function refreshPreviews() {
            options.syncTitleFieldEnabled();
            options.syncHalfwidthKanaEnabled();
            options.syncTimestampHHMMEnabled();
            options.syncPageControlsEnabled();
            // 「バージョンのみ」モードでは UI 整形を一切かけず、元ファイル名の v 番号だけ更新
            if (opMode.isVersionOnly()) {
                var versionOnlyBase = bumpVersionInPlace(stripExtension(currentName));
                versionOnlyBase = nextAvailableVersionName(versionOnlyBase, folder, '.ai');
                filename.finalNameValue.text = versionOnlyBase + '.ai';
                return;
            }
            var uiState = currentUIState();
            var finalBase = buildFinalName(segments, uiState);
            if (uiState.version === 'short' || uiState.version === 'padded' || uiState.version === 'paddedWide') {
                finalBase = nextAvailableVersionName(finalBase, folder, '.ai');
            }
            if (uiState.pageEnable === 'yes') {
                finalBase = nextAvailablePageName(finalBase, folder, '.ai', uiState.pagePrefix);
            }
            if (FEATURE_NFC && options.getNfc() === 'combine') {
                finalBase = normalizeNFC(finalBase);
            }
            // 半角カナ → 全角カナ（clean が '-' / '_' のときだけ）。translit より先に行う
            if (FEATURE_HALFWIDTH_KANA && options.getHalfwidthKana() === 'convert') {
                var cleanMode = options.getClean();
                if (cleanMode === 'dash' || cleanMode === 'underscore') {
                    finalBase = convertHalfwidthKana(finalBase);
                }
            }
            // translit はクリーンより先に行う（㈱→株 などを残すため）
            if (FEATURE_TRANSLITERATE) {
                finalBase = transliterate(finalBase, options.getTranslit());
            }
            if (FEATURE_CLEAN) {
                finalBase = cleanFilenameChars(finalBase, options.getClean());
            }
            finalBase = collapseAndTrimSeparators(finalBase, options.getSeparator());
            finalBase = escapeWindowsReserved(finalBase);
            filename.finalNameValue.text = finalBase + '.ai';
        }

        opMode.versionOnlyRadio.onClick = function () {
            syncOpModeVisibility();
            refreshPreviews();
        };
        opMode.fullRadio.onClick = function () {
            syncOpModeVisibility();
            refreshPreviews();
        };

        // ファイル名設定パネル内のすべての入力（無効化中の FEATURE は row が null で来るので skip される）
        var titleRow = options.titleRow, timestampRow = options.timestampRow, versionRow = options.versionRow;
        var separatorRow = options.separatorRow, nfcRow = options.nfcRow, cleanRow = options.cleanRow, translitRow = options.translitRow;

        // 構成要素（ベース／サブテキスト／ステータス／タイムスタンプ／バージョン）を変更したら、
        // 「現在のファイル名に準じる」は前提が崩れるので「標準順」に降格させる
        function demoteSortToDefault() {
            if (!sort) return;
            if (!sort.sortCurrentRadio.value) return;
            sort.sortCurrentRadio.value = false;
            sort.sortOffRadio.value = true;
            sort.syncSortButtonEnabled();
        }
        function refreshAndDemoteSort() {
            demoteSortToDefault();
            refreshPreviews();
        }

        // 構成要素を変える操作（「現在のファイル名に準じる」は解除）
        wireRefresh(refreshAndDemoteSort, [
            titleRow.radios.none, titleRow.radios.parent, titleRow.radios.grandparent, titleRow.radios.custom,
            options.baseField, options.titleField,
            options.statusDropdown,
            timestampRow.radios.none, timestampRow.radios.date, timestampRow.radios.dateDash,
            options.timestampHHMMCheckbox,
            options.pageCheckbox, options.pagePrefixField, options.pagePadRadio2, options.pagePadRadio3,
            versionRow.radios.none, versionRow.radios.short_, versionRow.radios.padded, versionRow.radios.paddedWide
        ]);
        // 整形のみ変える操作（並び順には影響しないので「現在のファイル名に準じる」を維持）
        wireRefresh(refreshPreviews, [
            separatorRow && separatorRow.radios.noChange, separatorRow && separatorRow.radios.dash, separatorRow && separatorRow.radios.underscore,
            nfcRow && nfcRow.radios.keep, nfcRow && nfcRow.radios.combine,
            cleanRow && cleanRow.radios.remove, cleanRow && cleanRow.radios.dash, cleanRow && cleanRow.radios.underscore,
            options.halfwidthKanaCheckbox,
            translitRow && translitRow.radios.keep, translitRow && translitRow.radios.remove, translitRow && translitRow.radios.convert
        ]);

        // ソートパネルの ON/OFF とサブダイアログ起動（FEATURE_SORT のとき）
        if (FEATURE_SORT && sort) {
            var onSortToggle = function () {
                sort.syncSortButtonEnabled();
                refreshPreviews();
            };
            sort.sortOffRadio.onClick = onSortToggle;
            sort.sortCurrentRadio.onClick = onSortToggle;
            sort.sortOnRadio.onClick = onSortToggle;
            sort.sortButton.onClick = function () {
                var newOrder = openSortDialog(customOrder);
                if (newOrder) {
                    customOrder = newOrder;
                    refreshPreviews();
                }
            };
        }

        refreshPreviews();

        // ---- ボタン（右寄せ Cancel / OK） ----
        var buttonGroup = dialog.add('group');
        buttonGroup.alignment = ['right', 'top'];
        buttonGroup.add('button', undefined, L('button.cancel'), { name: 'cancel' });
        buttonGroup.add('button', undefined, 'OK', { name: 'ok' });

        // 初期モードが「バージョンのみ」のときは該当 UI を隠した状態で表示
        if (opMode.isVersionOnly()) syncOpModeVisibility();

        return {
            dialog: dialog,
            getUIState: currentUIState,
            getMode: mode.getMode
        };
    }

    // =========================================
    // メイン / Main
    // =========================================

    /* エントリポイント。ダイアログを開き、選択モードに応じた出力を実行 / Entry point: open the dialog and execute the selected mode */
    function main() {
        if (app.documents.length === 0) {
            alert(L('message.noDoc'));
            return;
        }

        var doc = app.activeDocument;
        var documentInfo = gatherDocumentInfo(doc);
        var segments = mergeFragmentedText(parseFileName(documentInfo.baseName));
        var prefs = loadPrefs();

        // 保存先はダイアログ表示前に確定する。未保存ドキュメントはここで保存先フォルダを選択。
        // 同じフォルダをプレビュー（連番/バージョンの繰り上げ）と実処理の両方で使い、両者を一致させる
        var targetFolder = ensureTargetFolder(documentInfo.folder);
        if (!targetFolder) return; // キャンセル

        var dialogUI = createDialog(segments, documentInfo.currentName, prefs, documentInfo.parentFolderName, documentInfo.grandparentFolderName, targetFolder);
        if (dialogUI.dialog.show() !== 1) return; // キャンセル

        var uiState = dialogUI.getUIState();
        var newBaseName = (uiState.opMode === 'versionOnly')
            ? bumpVersionInPlace(documentInfo.baseName)
            : buildFinalName(segments, uiState);
        if (!newBaseName) {
            alert(L('message.emptyName'));
            return;
        }

        if (uiState.opMode === 'versionOnly' || uiState.version === 'short' || uiState.version === 'padded' || uiState.version === 'paddedWide') {
            newBaseName = nextAvailableVersionName(newBaseName, targetFolder, '.ai');
        }
        if (uiState.opMode !== 'versionOnly' && uiState.pageEnable === 'yes') {
            newBaseName = nextAvailablePageName(newBaseName, targetFolder, '.ai', uiState.pagePrefix);
        }

        // 「バージョンのみ」モードでは UI 整形をスキップして元ファイル名をそのまま尊重
        if (uiState.opMode !== 'versionOnly') {
            if (FEATURE_NFC && uiState.nfc === 'combine') {
                newBaseName = normalizeNFC(newBaseName);
            }
            if (FEATURE_HALFWIDTH_KANA && uiState.halfwidthKana === 'convert'
                && (uiState.clean === 'dash' || uiState.clean === 'underscore')) {
                newBaseName = convertHalfwidthKana(newBaseName);
            }
            if (FEATURE_TRANSLITERATE) {
                newBaseName = transliterate(newBaseName, uiState.translit);
            }
            if (FEATURE_CLEAN) {
                newBaseName = cleanFilenameChars(newBaseName, uiState.clean);
            }
            newBaseName = collapseAndTrimSeparators(newBaseName, uiState.separator);
            newBaseName = escapeWindowsReserved(newBaseName);
        }

        var destFile = File(targetFolder.fsName + '/' + newBaseName + '.ai');

        // 長さチェック（拡張子込み）: 上限超過なら確認ダイアログを出して続行可
        var fullByteLength = byteLengthUTF8(newBaseName + '.ai');
        if (fullByteLength > FEATURE_MAX_FILENAME_BYTES) {
            var warningMessage = L('message.confirmTooLong')
                .replace('{bytes}', String(fullByteLength))
                .replace('{limit}', String(FEATURE_MAX_FILENAME_BYTES));
            if (!confirm(warningMessage + '\n\n' + newBaseName + '.ai')) return;
        }

        if (!confirmOverwriteIfExists(destFile, documentInfo.fsPath)) return;

        try {
            executeOutput(doc, destFile, dialogUI.getMode(), documentInfo.fsPath);
            // 成功したら今回の選択をプリセットとして保存（versionOnly モードでは保存しない）
            if (uiState.opMode !== 'versionOnly') {
                var prefsToSave = {
                    titleMode: uiState.titleMode,
                    timestamp: uiState.timestamp,
                    timestampTime: uiState.timestampTime,
                    version: uiState.version
                };
                if (FEATURE_PAGE) {
                    prefsToSave.pageEnable = uiState.pageEnable;
                    prefsToSave.pagePad = uiState.pagePad;
                    prefsToSave.pagePrefix = uiState.pagePrefix;
                }
                if (FEATURE_STATUS) prefsToSave.status = uiState.status;
                if (FEATURE_SEPARATOR) prefsToSave.separator = uiState.separator;
                if (FEATURE_NFC) prefsToSave.nfc = uiState.nfc;
                if (FEATURE_CLEAN) prefsToSave.clean = uiState.clean;
                if (FEATURE_HALFWIDTH_KANA && FEATURE_CLEAN) prefsToSave.halfwidthKana = uiState.halfwidthKana;
                if (FEATURE_TRANSLITERATE) prefsToSave.translit = uiState.translit;
                if (FEATURE_SORT) {
                    // sort モード自体は復元しない（毎回「現在のファイル名に準じる」を初期値）
                    prefsToSave.segmentOrder = uiState.customSegmentOrder.join(',');
                }
                savePrefs(prefsToSave);
            }
        } catch (e) {
            alert(L('message.saveFailed') + '\n' + e);
        }
    }

    main();

})();
