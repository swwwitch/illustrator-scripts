#target illustrator

    /*
    
    ### スクリプト名：
    
    FileNameManager.jsx — 現在のドキュメントのファイル名を変更／別名で保存

    ### 概要：

    - アクティブな Illustrator ドキュメントのファイル名を変更（実ファイルのリネーム）／別名で保存／コピーを保存します。
    - 保存先は常に現在のファイルと同じフォルダ。未保存ドキュメントは保存先フォルダを選択して保存します。
    - ファイル名を base / title / status / timestamp / version の各セグメントに分解し、UI から個別に編集できます。
      - base: 先頭の固定部分。「ベース」入力欄で編集可能（検出値が初期値、空欄可）
      - title: 「サブテキスト」ラジオで「なし / 指定（入力欄）/ 親フォルダー / 2 階層上のフォルダー」を選択
      - status: 「ステータス」dropdown で制作ステータス（wip / draft / review / approved / outlined など）を挿入。`：` より前の値だけがファイル名に入る
      - timestamp: 「タイムスタンプ」ラジオで「なし / YYYYMMDD / YYYY-MM-DD」を選択（デフォルト YYYYMMDD）
      - version: 「バージョン番号」ラジオで「なし / v1, v2… / v01, v02… / v001, v002…」を選択（元ファイルに v 番号が無ければ強制的に v01 形式）。既存の v 番号は +1、親フォルダ内に同パターンがあれば最大 v +1 まで自動繰り上げ
    - 整形ルール：
      - 区切り記号: 「変更しない / `-` / `_`」。`-` または `_` を選ぶとピース内の `-` `_` `.`（FEATURE_DOT_NORMALIZE）を統一。`YYYY-MM-DD` のタイムスタンプは対象外（内部 `-` を保護）
      - 濁点・半濁点の正規化: 「変更しない / 結合する」。HFS+/APFS 由来の分離した濁点・半濁点（NFD）を結合済み文字（NFC）に正規化（デフォルト 結合する）
      - クリーンなファイル名: 「削除する / -に変更 / _に変更」。OS 禁止文字（`\ / : * ? " < > |`）／絵文字／機種依存文字／半角・全角スペース・タブ・改行をまとめて処理（連続するスペースは 1 つに畳んでから置換）
      - 丸数字や法人略記など: 「変更しない / 変換する」。法人略記（㈱→株 等）／白丸 ①-⑳ / 黒丸 ❶-⓴ / 括弧 ⑴-⒇ / ローマ数字 Ⅰ-Ⅻ・ⅰ-ⅻ / 略記号（℡ № ）/ 単位（㎜ ㎝ ㎞ ㎎ ㎏ ㎡ ㎥）/ チルダ（〜 ～ → `~`）/ ハイフン類（― – — − 等 → `-`）を ASCII 相当に変換
    - 構成要素順のカスタマイズ：
      - 上部「構成要素の順序」パネルで「標準順 / 現在のファイル名に準じる / カスタム順」を選択
      - 「現在のファイル名に準じる」は、現在のファイル名で検出された要素のみをその並び順で出力
      - 「カスタム順」で［順序を編集...］ボタンから ↑↓ で並び替え
      - 設定はプリセットに保存され次回起動時に復元
    - 動作モード: 「元ファイルをリネーム」（新名で保存後に元ファイルを削除）／「別名で保存」（元ファイルは残り、作業中のドキュメントが新ファイルに切り替わる。デフォルト）／「コピー（複製）を保存」（元ファイルを上書き保存したうえで別名のコピーを作成。作業中のドキュメントは元ファイルのまま）
    - UI 構成: 上部「動作 + モード + 構成要素の順序」の 3 カラム → 「ファイル名プレビュー」（現在 / 保存後の名前）→ 「ファイル名の設定」 → Cancel / OK（右寄せ）
    - トップ部の FEATURE_STATUS / FEATURE_SORT / FEATURE_SEPARATOR / FEATURE_DOT_NORMALIZE / FEATURE_NFC / FEATURE_CLEAN / FEATURE_TRANSLITERATE で各機能を個別オフ可能。FEATURE_SEPARATOR は `'-'` / `'_'` の文字列で既定の区切りも指定

    ### 主な機能：

    - 元ファイルをリネーム（新名で保存後に元ファイルを削除して実質的にリネーム）
    - 別名で保存（元ファイルは保持。アクティブドキュメントは新ファイルに切り替わる。デフォルト）
    - コピー（複製）を保存（元ファイルに上書き保存後、別名のコピーを物理ファイルとして作成）
    - セグメント分解による独立編集（base / title / status / timestamp / version。base は UI 入力で編集、status / date / version は自動検出）
    - ステータス dropdown（11 種類 + 区切り線、ファイル名に入るのは `：` より前のみ。既存ファイル名から自動検出）
    - サブテキストは「指定／親フォルダー／2 階層上のフォルダー」から選択可能（フォルダ名が存在しないと該当ラジオは無効化）
    - 構成要素順のカスタマイズ（標準順／現在のファイル名に準じる／カスタム順。プリセット保存）
    - タイムスタンプ／バージョン番号は「なし」で削除、形式切替（YYYYMMDD / YYYY-MM-DD / v1, v2… / v01, v02… / v001, v002…）。バージョン番号は親フォルダ内の同パターン最大 v +1 へ自動繰り上げ
    - 区切り記号統一（`-` `_` `.`）／NFC 正規化（分離した濁点・半濁点の結合、デフォルトオン）
    - クリーンなファイル名（OS 禁止文字 + 絵文字 + 機種依存文字 + スペースを統一処理）
    - 丸数字や法人略記など（㈱→株 / ①→1 / Ⅰ→1 / ㎜→mm / ― → - 等の ASCII 化、デフォルトオフ）
    - 各機能を FEATURE_* 定数で個別にオフ可能

    ### 処理の流れ：

    1) ドキュメント検証 → gatherDocumentInfo() でファイル情報を取得
    2) parseFileName() でセグメント配列に分解（区切り文字も保持、status も検出）
    3) createDialog() でダイアログを構築（buildModePanel + buildSortPanel + buildFilenamePanel + buildOptionsPanel）
    4) buildFinalName() で UI 状態と segments、現在の SEGMENT_ORDER（カスタム順含む）から最終ファイル名を構築
    5) ensureTargetFolder() で未保存ドキュメントは保存先フォルダを取得
    6) confirmOverwriteIfExists() で既存同名の上書きを確認
    7) executeOutput() で選択モード（rename / saveAs / copy）に応じた出力を実行

    ### 更新履歴：

    - v1.0 (2026-05-27) : 初期バージョン
    - v1.1 (2026-06-01) : コロン付きラベル関数追加
    - (2026-05-27 追記) : ステータス dropdown、構成要素順カスタマイズ、スペース置換、YYYY-MM-DD タイムスタンプ、`.` 区切り正規化、整形機能の FEATURE スイッチ化、UI 整理（ベース行非表示、動作・構成要素 2 カラム、ボタン右寄せ）
    - (2026-05-28 追記) : ベース行 UI 復活（入力欄）、「タイトル」→「サブテキスト」（英語: Project Name）改名、バージョン番号表記を v1, v2… / v01, v02… / v001, v002… の 3 段階に拡張、親フォルダ内の同パターンから最大 v +1 を採用、NFC 正規化（濁点・半濁点の結合、FEATURE_NFC）追加、「コピー（複製）を保存」表記、「現在 / 保存後の名前」ラベル右揃え
    - v1.2.5 (2026-05-28) : 「現在のファイル名に準じる」並び順モード追加、元ファイルに v 番号が無い場合のデフォルトを v01 形式に固定、サブテキストに「2 階層上のフォルダー」追加・順序を「なし／指定／親フォルダー／2 階層上」に変更、「クリーンなファイル名」（OS 禁止文字 + 絵文字 + 機種依存文字 + スペースの統一処理。「スペースの扱い」を吸収）、「丸数字や法人略記など」（㈱→株 / ①→1 / Ⅰ→1 / ㎜→mm / ハイフン類 → `-` 等）追加、「濁点処理」→「濁点・半濁点の正規化」改名、内部リファクタ（addRadioRow / pickPref / wireRefresh ヘルパー導入で重複削減）
    - v1.2.6 (2026-05-29) : 「クリーンなファイル名」行に「半角カナ → 全角」チェックボックス追加（クリーンが `-` または `_` のときだけ有効・濁点／半濁点も結合）、「丸数字や法人略記など」に「削除する」を追加（変更しない／削除する／変換する）、UI デフォルト見直し（並び順「現在のファイル名に準じる」／クリーン「-に変更」／丸数字や法人略記など「変換する」）、「バージョン番号のみ」→「バージョンのみ」改名、ソートパネルの編集ボタンを「カスタム順」の右に移動し、ラベルを「順序を編集...」→「編集」に短縮
    - v1.2.7 (2026-05-29) : 自動整形（連続する `-` `_` `.` を 1 つに圧縮・先頭末尾のトリム）、ファイル名長の事前チェック（拡張子込み UTF-8 240 バイト超過で確認ダイアログ・FEATURE_MAX_FILENAME_BYTES）、rename 時の元ファイル削除を `~/.Trash` 移動へ変更（FEATURE_USE_TRASH、失敗時は file.remove() にフォールバック）、「元ファイルをリネーム」の helpTip に配置リンク切れ注意を追記、Windows 予約名（CON / PRN / AUX / NUL / COM1-9 / LPT1-9）を末尾 `_` で回避、ハングル音節（U+AC00–U+D7A3）を「標準」文字に追加、タイムスタンプに「時刻も付与」(HHMM) チェックボックス追加
    - v1.2.8 (2026-05-29) : 「バージョンのみ」モードでは UI 整形（NFC / 半角カナ / translit / クリーン / 連続区切り圧縮 / Windows 予約名回避）を一切かけず、元ファイル名の v 番号だけ更新するように変更、ソートパネルの初期値を prefs と無関係に常に「現在のファイル名に準じる」（不可なら「標準順」）に固定（カスタム並び順 segmentOrder は引き続き保存）、Ai 版に欠落していた `LABELS.sort`（サブダイアログのタイトル／ヒント）を追加、`tip.sort` をボタン名「編集」に揃えて短縮、ラベル表記を Id 版に統一（「2 階層上のフォルダー」「-に変更」「_に変更」）
    
    ---
    
    ### Script name:
    
    FileNameManager.jsx — Rename or Save As the Current Document
    
    ### Overview:

    - Renames the active Illustrator document (true rename), saves as a new file, or saves a copy.
    - Destination is always the same folder as the current file; for unsaved docs, a destination folder is chosen.
    - Decomposes the filename into base / title / status / timestamp / version segments and lets you edit them.
      - base: editable via the "Base" input (defaults to the detected value; may be empty)
      - title: "Project Name" radio (None / Custom input / Parent Folder / Grandparent Folder)
      - status: "Status" dropdown (wip / draft / review / approved / outlined, etc.); only the text before `:` is written to the filename
      - timestamp: "Timestamp" radio selects "None / YYYYMMDD / YYYY-MM-DD" (default YYYYMMDD)
      - version: "Version" radio selects "None / v1, v2… / v01, v02… / v001, v002…". When the original has no v-number, defaults to v01 form. Existing v-numbers are bumped by +1; if files matching the same pattern exist in the parent folder, the result is bumped further to (max v) + 1.
    - Formatting rules:
      - Separator: "Don't change / `-` / `_`". Selecting `-` or `_` unifies `-`, `_`, and `.` (FEATURE_DOT_NORMALIZE) inside each piece. `YYYY-MM-DD` timestamps are exempt (internal `-` is preserved)
      - NFC Normalization: "Don't change / Combine". Normalizes separated dakuten/handakuten (NFD) to combined characters (NFC). Default: Combine.
      - Clean Filename: "Remove / Change to - / Change to _". Handles OS-invalid chars (`\ / : * ? " < > |`), emoji, platform-dependent chars, and spaces (incl. tabs/newlines/fullwidth space) in one pass. Consecutive whitespace is collapsed to one before replacement.
      - Symbols (translit): "Don't change / Convert". Converts corporate abbreviations (㈱→株 etc.), circled numbers (white ①-⑳ / black ❶-⓴ / parenthesized ⑴-⒇), Roman numerals (Ⅰ-Ⅻ / ⅰ-ⅻ), abbrev symbols (℡ →TEL / № →No), units (㎜→mm / ㎏→kg etc.), tildes (〜 ～ → `~`), and dashes (― – — − etc. → `-`) to ASCII equivalents.
    - Segment order:
      - In the top "Segment Order" panel, choose "Default" / "Match Current" / "Custom"
      - "Match Current" outputs only the elements present in the current filename, in their original order
      - "Custom" enables [Edit order...] for ↑↓ reordering
      - The setting is persisted to prefs
    - Modes: "Rename Original" (saves with the new name, then deletes the original), "Save As" (keeps the original; the active document switches to the new file; default), and "Save a Copy" (saves over the original, then creates a separate copy; the active document remains the original file).
    - UI layout: top row "Mode + Scope + Segment Order" (3 columns) → "File Name Preview" (Current / Saved Name) → "Filename Settings" → Cancel / OK (right-aligned)
    - Features can be toggled individually via FEATURE_STATUS / FEATURE_SORT / FEATURE_SEPARATOR / FEATURE_DOT_NORMALIZE / FEATURE_NFC / FEATURE_CLEAN / FEATURE_TRANSLITERATE. FEATURE_SEPARATOR accepts `'-'` or `'_'` to also pick the default separator.

    ### Key features:

    - Rename Original (saves with the new name, then removes the original)
    - Save As (keeps the original; the active document switches to the new file; default)
    - Save a Copy (saves over the original, then creates a separate copy; the active document remains the original file)
    - Segment-based editing (base / title / status / timestamp / version; base edited via input, status / date / version auto-detected)
    - Status dropdown (11 entries + divider; only the text before `:` is used; detected from existing filenames)
    - Project Name selectable from Custom / Parent Folder / Grandparent Folder (radios are disabled when no such folder exists)
    - Segment order via Default / Match Current / Custom (persisted to prefs)
    - Timestamp / version: "None" removes the segment; switch format (YYYYMMDD / YYYY-MM-DD / v1, v2… / v01, v02… / v001, v002…). Version auto-bumps to (max v in folder) + 1 when collisions exist.
    - Separator unification (`-` `_` `.`) / NFC normalization for separated dakuten/handakuten (default on)
    - Clean Filename (unified handling of OS-invalid chars + emoji + platform-dependent chars + spaces)
    - Symbols (㈱→株 / ①→1 / Ⅰ→1 / ㎜→mm / ― → - etc., default off)
    - Each feature toggleable via FEATURE_* constants

    ### Flow:

    1) gatherDocumentInfo() collects file info
    2) parseFileName() splits the name into ordered segments (separators preserved; status detected)
    3) createDialog() composes the dialog (buildModePanel + buildSortPanel + buildFilenamePanel + buildOptionsPanel)
    4) buildFinalName() composes the final filename from UI state, segments, and the current SEGMENT_ORDER (custom order included)
    5) ensureTargetFolder() prompts for a destination folder if the doc is unsaved
    6) confirmOverwriteIfExists() confirms before overwriting an existing file
    7) executeOutput() runs the action for the selected mode (rename / saveAs / copy)

    ### Changelog:

    - v1.0 (2026-05-27): Initial release
    - (2026-05-27 update): Added status dropdown, custom segment order, space replacement, YYYY-MM-DD timestamp, `.` separator normalization, FEATURE switches, UI cleanup (no base row, 2-col top, right-aligned buttons)
    - (2026-05-28 update): Restored editable Base input, renamed "Title" → "Project Name", expanded version format to v1, v2… / v01, v02… / v001, v002…, auto-bump to (folder max v) + 1, added NFC normalization (FEATURE_NFC, default on), "Save a Copy" wording, right-aligned Current/Saved Name labels
    - v1.2.5 (2026-05-28): Added "Match Current" segment-order mode, forced v01 default when no v-number exists in original, added "Grandparent Folder" as Project Name option with reordered radios (None / Custom / Parent / Grandparent), added "Clean Filename" (OS-invalid + emoji + platform-dependent chars + spaces in one pass; absorbs the former "Spaces" option), added "Symbols" transliteration (㈱→株 / ①→1 / Ⅰ→1 / ㎜→mm / dashes → `-` etc.), renamed "NFC" label to "NFC Normalization", internal refactor (addRadioRow / pickPref / wireRefresh helpers to reduce duplication)
    - v1.2.6 (2026-05-29): Added "Half-width → Full-width Kana" checkbox inside the Clean Filename row (active only when Clean is `-` or `_`; merges voiced/semi-voiced marks), expanded "Symbols" to 3-way (Keep / Remove / Convert), updated UI defaults (Sort → Match Current, Clean → `-`, Symbols → Convert), renamed "Version Only" to drop the word "Number" in JA, moved the order-edit button next to the "Custom" radio and shortened "Edit order..." → "Edit"
    - v1.2.7 (2026-05-29): Auto cleanup (collapse `-` / `_` / `.` runs to one; trim leading/trailing separators), pre-save length check (confirm dialog when basename+ext exceeds 240 UTF-8 bytes; FEATURE_MAX_FILENAME_BYTES), rename moves the original to `~/.Trash` instead of hard-removing it (FEATURE_USE_TRASH; falls back to file.remove() on failure), added a placed-link warning to the rename helpTip, escape Windows reserved names (CON / PRN / AUX / NUL / COM1-9 / LPT1-9) with a trailing `_`, treat Hangul syllables (U+AC00–U+D7A3) as standard characters, added an "Append HHMM" checkbox to the timestamp row
    - v1.2.8 (2026-05-29): "Version Only" mode no longer applies any cleanup (NFC / half-width kana / translit / clean / collapse separators / Windows reserved escape); the original filename is preserved and only its v-number is bumped, the sort panel default is now always "Match Current" (or "Default" when unavailable) regardless of prefs (custom `segmentOrder` is still persisted), added the missing `LABELS.sort` (reorder sub-dialog title / hint) to the AI version, shortened `tip.sort` to match the "Edit" button label, and unified label spelling with the InDesign version ("Grandparent Folder" radio / "Change to -" / "Change to _")
    
    */

    (function () {

        // =========================================
        // バージョン / Version
        // =========================================

        var SCRIPT_VERSION = "v1.2.8";

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

        /* 出力時のセグメント順序。base / title / status / timestamp / version。FEATURE_STATUS=false なら status は除外
           / Output segment order; "status" is dropped when FEATURE_STATUS is false */
        var SEGMENT_ORDER = FEATURE_STATUS
            ? ['base', 'title', 'status', 'timestamp', 'version']
            : ['base', 'title', 'timestamp', 'version'];

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
        var PANEL_MARGINS = [15, 20, 15, 10];
        var PANEL_SPACING = 8;

        var DIALOG_OPACITY = 0.98;

        // =========================================
        // ローカライズ / Localization
        // =========================================

        var lang = ($.locale.indexOf("ja") === 0) ? "ja" : "en";

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
            return entry[lang] || entry.en || path;
        }

        /* コロン付きラベル（日本語は全角、英語は半角）/ Label with colon (full-width JA, half-width EN) */
        function labelText(path) {
            return L(path) + (lang === 'ja' ? '：' : ':');
        }

        // =========================================
        // ヘルパー / Helpers
        // =========================================

        var VERSION_TOKEN_RE = /^[vV]\d+$/;   // v123 / V123

        /* 月（1-12）と日（1-31）の妥当性 / Validate month (1-12) and day (1-31) */
        function isValidMonthDay(monthStr, dayStr) {
            var m = parseInt(monthStr, 10);
            var d = parseInt(dayStr, 10);
            return m >= 1 && m <= 12 && d >= 1 && d <= 31;
        }

        /* YYYYMMDD（8 桁）または YYMMDD（6 桁）の日付トークンか / Whether a token is a YYYYMMDD or YYMMDD date */
        function isDateToken(token) {
            var s = String(token);
            if (/^\d{8}$/.test(s)) return isValidMonthDay(s.substring(4, 6), s.substring(6, 8));
            if (/^\d{6}$/.test(s)) return isValidMonthDay(s.substring(2, 4), s.substring(4, 6));
            return false;
        }

        /* "v123" 形式のバージョントークンか / Whether a token is a v+digits version */
        function isVersionToken(token) {
            return VERSION_TOKEN_RE.test(String(token));
        }

        /* STATUS_ITEMS のいずれかの value と完全一致（大文字小文字無視）すれば true。一致した正規 value を返す。divider は除外
           / Match against STATUS_ITEMS values (case-insensitive); returns the canonical value or '' (skips dividers) */
        function matchStatusToken(token) {
            var s = String(token).toLowerCase();
            for (var i = 1; i < STATUS_ITEMS.length; i++) {
                if (isStatusDivider(STATUS_ITEMS[i])) continue;
                if (STATUS_ITEMS[i].value === s) return STATUS_ITEMS[i].value;
            }
            return '';
        }

        /* ファイル名を順序付きセグメント配列に分解。各 segment は前にあった区切り文字も保持
           / Decompose into ordered segments; each segment stores the preceding separator */
        function parseFileName(name) {
            var split = String(name).split(/([-_])/); // ["handout","-","Adobe","-","20260422"]
            var segments = [];
            var textBuffer = [];      // token と区切りを交互に蓄積（join('') で結合）
            var textLeadingSep = '';  // テキスト segment の直前に置く区切り
            var hasDate = false, hasVersion = false, hasStatus = false;
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
            while (i < split.length) {
                if (i % 2 === 1) {
                    currentSep = split[i];
                    i++;
                    continue;
                }
                var token = split[i];

                // YYYY-MM-DD / YYYY_MM_DD（3 トークン + 同じ区切り 2 つ）の日付パターン
                if (!hasDate && i > 0
                    && /^\d{4}$/.test(token)
                    && i + 4 < split.length
                    && (split[i + 1] === '-' || split[i + 1] === '_')
                    && /^\d{2}$/.test(split[i + 2])
                    && split[i + 3] === split[i + 1]
                    && /^\d{2}$/.test(split[i + 4])
                    && isValidMonthDay(split[i + 2], split[i + 4])) {
                    flushText();
                    var compositeDate = token + split[i + 1] + split[i + 2] + split[i + 3] + split[i + 4];
                    segments.push({ kind: 'date', value: compositeDate, sep: currentSep });
                    hasDate = true;
                    i += 5; // 3 tokens + 2 separators
                    continue;
                }

                if (i === 0) {
                    segments.push({ kind: 'base', value: token, sep: '' });
                } else if (!hasDate && isDateToken(token)) {
                    flushText();
                    segments.push({ kind: 'date', value: token, sep: currentSep });
                    hasDate = true;
                } else if (!hasVersion && isVersionToken(token)) {
                    flushText();
                    segments.push({ kind: 'version', value: token, sep: currentSep });
                    hasVersion = true;
                } else if (FEATURE_STATUS && !hasStatus && matchStatusToken(token)) {
                    flushText();
                    segments.push({ kind: 'status', value: matchStatusToken(token), sep: currentSep });
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

        /* 複数の text セグメントが分散している（間に date/version/status が挟まる）場合、
           それらをまとめて 1 つの text として再結合。subtext として完全な文字列を扱えるようにする
           / Merge fragmented text segments back into one (with the in-between non-text segments inlined) */
        function mergeFragmentedText(segments) {
            var firstText = -1, lastText = -1;
            for (var i = 0; i < segments.length; i++) {
                if (segments[i].kind === 'text') {
                    if (firstText === -1) firstText = i;
                    lastText = i;
                }
            }
            if (firstText === -1 || firstText === lastText) return segments;
            var combined = segments[firstText].value;
            for (var j = firstText + 1; j <= lastText; j++) {
                combined += segments[j].sep + segments[j].value;
            }
            var merged = { kind: 'text', value: combined, sep: segments[firstText].sep };
            var result = segments.slice(0, firstText);
            result.push(merged);
            for (var k = lastText + 1; k < segments.length; k++) result.push(segments[k]);
            return result;
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
                version: 'version'
            };
            var order = [];
            var seen = {};
            for (var i = 0; i < segments.length; i++) {
                var k = kindToOrder[segments[i].kind];
                if (!k || seen[k]) continue;
                if (k === 'status' && !FEATURE_STATUS) continue;
                order.push(k);
                seen[k] = true;
            }
            return order;
        }

        /* segments から最初に出現する kind の value を取得 / Get the value of the first segment of the given kind */
        function getFirstSegmentValue(segments, kind) {
            for (var i = 0; i < segments.length; i++) {
                if (segments[i].kind === kind) return segments[i].value;
            }
            return '';
        }

        /* 文字列を左 0 パディング / Left-pad a string with zeros to the given width */
        function padLeft(str, width) {
            while (str.length < width) str = '0' + str;
            return str;
        }

        /* 今日の日付を返す。sep で日付内区切り、withTime=true で末尾に "-HHMM" を付加
           / Today's date; with `sep` between Y/M/D parts; `withTime=true` appends "-HHMM" */
        function todayTimestamp(sep, withTime) {
            sep = sep || '';
            var d = new Date();
            var date = String(d.getFullYear()) + sep +
                padLeft(String(d.getMonth() + 1), 2) + sep +
                padLeft(String(d.getDate()), 2);
            if (!withTime) return date;
            return date + '-' + padLeft(String(d.getHours()), 2) + padLeft(String(d.getMinutes()), 2);
        }

        /* 現在のファイル名（拡張子なし）の v 番号だけを桁数維持で +1。
           v 番号が無ければ末尾に "-v2" を付与
           / Bump the v-number inside the current basename in place (preserving digit width).
           Appends "-v2" if no v-number exists. */
        function bumpVersionInPlace(currentBaseName) {
            var match = String(currentBaseName).match(/([vV])(\d+)/);
            if (match) {
                var letter = match[1];
                var digits = match[2];
                var nextNum = parseInt(digits, 10) + 1;
                var newDigits = padLeft(String(nextNum), digits.length);
                return currentBaseName.replace(/([vV])\d+/, letter + newDigits);
            }
            return currentBaseName + '-v2';
        }

        /* バージョン文字列を +1。mode='padded' で 2 桁ゼロ埋め、'paddedWide' で 3 桁ゼロ埋め。
           元バージョンが無ければ新規付与（v1 / v01 / v001）
           / Bump the version string by +1. 'padded'=min 2 digits, 'paddedWide'=min 3. Returns v1/v01/v001 if no original */
        function formatVersion(originalVersion, mode) {
            var match = String(originalVersion || '').match(/^([vV])(\d+)$/);
            var letter = match ? match[1] : 'v';
            var nextNum = match ? (parseInt(match[2], 10) + 1) : 1;
            if (mode === 'padded') {
                var width = match ? Math.max(match[2].length, 2) : 2;
                return letter + padLeft(String(nextNum), width);
            }
            if (mode === 'paddedWide') {
                var widthWide = match ? Math.max(match[2].length, 3) : 3;
                return letter + padLeft(String(nextNum), widthWide);
            }
            return letter + String(nextNum);
        }

        /* baseName 内の最初の v+数字パターンを抽出。マッチしなければ null
           / Extract the first v+digits in baseName; null if no v-number */
        function extractVersionParts(baseName) {
            var m = String(baseName).match(/^(.*?)([vV])(\d+)(.*)$/);
            if (!m) return null;
            return { prefix: m[1], letter: m[2], digits: m[3], suffix: m[4] };
        }

        /* 正規表現エスケープ / Escape for use in RegExp */
        function escapeRegExp(s) {
            return String(s).replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&');
        }

        /* baseName と同じ prefix / suffix を持つファイルを folder から探し、最大の v 番号を返す
           / Scan folder for files matching baseName's v-pattern; return max v-number or null */
        function findMaxVersionInFolder(baseName, folder, ext) {
            if (!folder) return null;
            var parts = extractVersionParts(baseName);
            if (!parts) return null;
            var re = new RegExp(
                '^' + escapeRegExp(parts.prefix) + '[vV](\\d+)'
                + escapeRegExp(parts.suffix) + escapeRegExp(ext) + '$',
                'i'
            );
            var files;
            try { files = folder.getFiles(); } catch (e) { return null; }
            var maxNum = -1;
            for (var i = 0; i < files.length; i++) {
                if (!(files[i] instanceof File)) continue;
                var fname = decodePercentEncoded(files[i].name);
                var fm = fname.match(re);
                if (!fm) continue;
                var num = parseInt(fm[1], 10);
                if (num > maxNum) maxNum = num;
            }
            return maxNum >= 0 ? maxNum : null;
        }

        /* 親フォルダー内の同パターンファイルの最大 v 番号 +1 に置き換える。
           無い、または +1 が現在値以下のときは baseName のまま返す。桁数は維持
           / Replace v-number with (folder max + 1); preserves digit width */
        function nextAvailableVersionName(baseName, folder, ext) {
            var maxNum = findMaxVersionInFolder(baseName, folder, ext);
            if (maxNum === null) return baseName;
            var parts = extractVersionParts(baseName);
            if (!parts) return baseName;
            var current = parseInt(parts.digits, 10);
            var target = maxNum + 1;
            if (target <= current) return baseName;
            var width = Math.max(parts.digits.length, String(target).length);
            return parts.prefix + parts.letter + padLeft(String(target), width) + parts.suffix;
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
            var key = String(baseName).toUpperCase();
            return (WINDOWS_RESERVED_NAMES[key] === 1) ? baseName + '_' : baseName;
        }

        /* % エンコードを 1 回デコード / Decode a percent-encoded string once (best-effort) */
        function decodePercentEncoded(str) {
            str = String(str);
            if (str.indexOf('%') === -1) return str;
            try {
                return decodeURIComponent(str);
            } catch (e) {
                return str;
            }
        }

        /* 拡張子を除去 / Strip a trailing file extension */
        function stripExtension(name) {
            var dot = name.lastIndexOf('.');
            return (dot > 0) ? name.substring(0, dot) : name;
        }

        /* 連続する区切り（- _ .）を 1 つに圧縮し、先頭・末尾の区切りと空白をトリム。
           sep は最終的な区切り文字（'-' / '_' / '' のいずれか、'' のときは混在をそのまま残す）
           / Collapse runs of separators (-, _, .) to one and trim leading/trailing separators + spaces */
        function collapseAndTrimSeparators(str, sep) {
            var s = String(str);
            if (sep === '-' || sep === '_') {
                // 異種の混在も含めて連続した区切りを sep 1 つに統一（YYYY-MM-DD 等は呼び出し前に通過済み想定）
                s = s.replace(/[-_.]{2,}/g, sep);
            } else {
                // 区切り未指定でも同種連続だけは圧縮（-- → -, __ → _, .. → .）
                s = s.replace(/--+/g, '-').replace(/__+/g, '_').replace(/\.\.+/g, '.');
            }
            // 先頭末尾の区切り・空白を除去
            s = s.replace(/^[-_.\s]+|[-_.\s]+$/g, '');
            return s;
        }

        /* 文字列の UTF-8 バイト長（拡張子込みのファイル名サイズ判定用）。ES3 / ExtendScript 向け
           / UTF-8 byte length of a string (for file-name size checks); ES3-safe */
        function byteLengthUTF8(str) {
            var s = String(str);
            var n = 0;
            for (var i = 0; i < s.length; i++) {
                var code = s.charCodeAt(i);
                if (code < 0x80) n += 1;
                else if (code < 0x800) n += 2;
                else if (code >= 0xD800 && code <= 0xDBFF) { n += 4; i++; } // サロゲートペア（4 バイト）
                else n += 3;
            }
            return n;
        }

        /* 前後空白をトリム。FEATURE_CLEAN=false のときは OS 禁止文字も除去（後段で処理されないため）
           / Trim whitespace. When FEATURE_CLEAN is off, also strip OS-invalid chars (no post-process safety net) */
        function sanitizeFilename(str) {
            var trimmed = String(str).replace(/^\s+|\s+$/g, '');
            if (FEATURE_CLEAN) return trimmed;
            return trimmed.replace(/[\\\/:*?"<>|]/g, '');
        }

        /* HFS+/APFS 由来の NFD 文字列（か + ゛など）を NFC（が）に結合。
           ひらがな・カタカナ + U+3099/U+309A を主に対象。う/ウ/ワ/ヰ/ヱ/ヲ は特例マッピング
           / Normalize NFD-style hiragana/katakana + combining marks to composed NFC form */
        function normalizeNFC(str) {
            return String(str).replace(/(.)([゙゚])/g, function (_, base, mark) {
                var voiced = (mark === '゙');
                if (voiced) {
                    if (base === 'う') return 'ゔ';
                    if (base === 'ウ') return 'ヴ';
                    if (base === 'ワ') return 'ヷ';
                    if (base === 'ヰ') return 'ヸ';
                    if (base === 'ヱ') return 'ヹ';
                    if (base === 'ヲ') return 'ヺ';
                }
                var code = base.charCodeAt(0);
                var isHiragana = code >= 0x3041 && code <= 0x3093;
                var isKatakana = code >= 0x30A1 && code <= 0x30F3;
                if (isHiragana || isKatakana) {
                    return String.fromCharCode(code + (voiced ? 1 : 2));
                }
                return base + mark;
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
        function cleanFilenameChars(str, mode) {
            var replacement = (mode === 'dash') ? '-' : (mode === 'underscore') ? '_' : '';
            // 連続する空白を 1 つにまとめてから per-char で置換（mode='dash' で "  " が "--" にならないように）
            var s = String(str).replace(/\s+/g, ' ');
            var result = '';
            var i = 0;
            while (i < s.length) {
                var code = s.charCodeAt(i);
                if (code >= 0xD800 && code <= 0xDBFF && i + 1 < s.length) {
                    // 高サロゲート + 低サロゲートは常に非標準（BMP 外の絵文字など）
                    result += replacement;
                    i += 2;
                    continue;
                }
                if (isStandardFilenameChar(code)) {
                    result += s.charAt(i);
                } else {
                    result += replacement;
                }
                i++;
            }
            return result;
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
        function convertHalfwidthKana(str) {
            var s = String(str);
            var result = '';
            var i = 0;
            while (i < s.length) {
                var ch = s.charAt(i);
                var next = (i + 1 < s.length) ? s.charAt(i + 1) : '';
                if (next === 'ﾞ' && typeof HALFWIDTH_KANA_VOICED[ch] === 'string') {
                    result += HALFWIDTH_KANA_VOICED[ch];
                    i += 2;
                    continue;
                }
                if (next === 'ﾟ' && typeof HALFWIDTH_KANA_SEMI_VOICED[ch] === 'string') {
                    result += HALFWIDTH_KANA_SEMI_VOICED[ch];
                    i += 2;
                    continue;
                }
                var mapped = HALFWIDTH_KANA_MAP[ch];
                if (typeof mapped === 'string') {
                    result += mapped;
                } else if (ch === 'ﾞ') {
                    result += '゛';
                } else if (ch === 'ﾟ') {
                    result += '゜';
                } else {
                    result += ch;
                }
                i++;
            }
            return result;
        }

        /* TRANSLITERATE_MAP に従って 1 文字ずつ処理。mode='convert' は変換、'remove' は削除、'keep' は無処理。
           値が文字列のときのみ対象（toString などの prototype プロパティ衝突を回避）
           / Per-char transform via TRANSLITERATE_MAP: 'convert' replaces, 'remove' drops, 'keep' returns as-is */
        function transliterate(str, mode) {
            if (mode === 'keep') return String(str);
            var s = String(str);
            var result = '';
            for (var i = 0; i < s.length; i++) {
                var ch = s.charAt(i);
                var mapped = TRANSLITERATE_MAP[ch];
                if (typeof mapped === 'string') {
                    result += (mode === 'remove') ? '' : mapped;
                } else {
                    result += ch;
                }
            }
            return result;
        }

        /* 保存先ダイアログ。.ai 拡張子を補正して File を返す。キャンセルは null / Show save dialog, normalize to .ai; null on cancel */
        function pickAiDestination(promptLabel) {
            var chosen = File.saveDialog(promptLabel, '*.ai');
            if (!chosen) return null;
            var file = File(chosen.fsName);
            if (!/\.ai$/i.test(file.name)) {
                file = File(file.parent.fsName + '/' + stripExtension(file.name) + '.ai');
            }
            return file;
        }

        /* プリセット保存用ファイル / Preferences file (key=value lines) */
        function getPrefsFile() {
            return File(Folder.userData.fsName + '/FileNameManager-prefs.txt');
        }

        /* 前回保存したプリセットを読み込み（無ければ空オブジェクト） / Load previously saved prefs (or empty if none) */
        function loadPrefs() {
            var file = getPrefsFile();
            if (!file.exists) return {};
            file.encoding = 'UTF-8';
            file.open('r');
            var raw = file.read();
            file.close();
            var prefs = {};
            var lines = String(raw).split('\n');
            for (var i = 0; i < lines.length; i++) {
                var line = lines[i];
                var eq = line.indexOf('=');
                if (eq > 0) prefs[line.substring(0, eq)] = line.substring(eq + 1);
            }
            return prefs;
        }

        /* プリセットを保存（失敗時は黙って継続） / Save prefs (silently ignore failure) */
        function savePrefs(prefs) {
            try {
                var file = getPrefsFile();
                file.encoding = 'UTF-8';
                file.open('w');
                var lines = [];
                for (var key in prefs) {
                    if (prefs.hasOwnProperty(key)) lines.push(key + '=' + prefs[key]);
                }
                file.write(lines.join('\n'));
                file.close();
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
            var picked = pickAiDestination(L('message.chooseDestination'));
            return picked ? picked.parent : null;
        }

        /* 既存同名（自分自身を除く）の上書き確認。OK なら true / Confirm overwrite for an existing file (excluding self); true if approved */
        function confirmOverwriteIfExists(destFile, originalFsPath) {
            if (!destFile.exists) return true;
            if (destFile.fsName === originalFsPath) return true;
            return confirm(L('message.confirmOverwrite') + '\n\n' + destFile.fsName);
        }

        /* file を ~/.Trash に移動。同名衝突時は連番でユニーク化。成功で true、失敗で false
           / Move file to ~/.Trash, disambiguating by appending a counter. Returns true on success */
        function moveToTrash(file) {
            if (!file || !file.exists) return false;
            var trash = Folder("~/.Trash");
            if (!trash.exists) return false;
            var origName = file.name;
            var dot = origName.lastIndexOf('.');
            var stem = (dot > 0) ? origName.substring(0, dot) : origName;
            var ext = (dot > 0) ? origName.substring(dot) : '';
            var dest = File(trash.fsName + '/' + origName);
            var counter = 1;
            while (dest.exists) {
                dest = File(trash.fsName + '/' + stem + ' ' + counter + ext);
                counter++;
                if (counter > 9999) return false; // 暴走防止
            }
            try {
                if (!file.copy(dest)) return false;
                return file.remove();
            } catch (e) {
                return false;
            }
        }

        /* 旧ファイルを削除（FEATURE_USE_TRASH=true ならゴミ箱に移動、失敗時は file.remove() にフォールバック）
           / Remove the original file: move to ~/.Trash when enabled, else (or on failure) hard-remove */
        function removeOriginalFile(originalFsPath, destFsPath) {
            if (!originalFsPath || originalFsPath === destFsPath) return;
            var file = File(originalFsPath);
            if (!file.exists) return;
            if (FEATURE_USE_TRASH && moveToTrash(file)) return;
            try { file.remove(); } catch (e) { /* 削除できない場合は黙って継続 */ }
        }

        /* モード別の出力処理（rename / saveAs / copy） / Execute the output according to the selected mode */
        function executeOutput(doc, destFile, mode, originalFsPath) {
            if (mode === 'copy' && originalFsPath) {
                // 現在の変更を元ファイルへ保存してから、物理ファイルとしてコピー
                if (!doc.saved) doc.save();
                File(originalFsPath).copy(destFile);
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
            var defaultSep;
            if (uiState.separator === '-' || uiState.separator === '_') {
                defaultSep = uiState.separator;
            } else {
                defaultSep = dominantSeparator(segments);
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
                return '';
            }

            // ピースごとに区切り統一（YYYY-MM-DD のタイムスタンプだけは内部 "-" を保護）
            var order = (uiState.segmentOrder && uiState.segmentOrder.length) ? uiState.segmentOrder : SEGMENT_ORDER;
            var pieces = [];
            for (var i = 0; i < order.length; i++) {
                var kind = order[i];
                var value = valueForKind(kind);
                if (!value) continue;
                // YYYY-MM-DD（および YYYY-MM-DD-HHMM、YYYYMMDD-HHMM）は内部の "-" を保護
                var isProtectedDate = (kind === 'timestamp')
                    && (uiState.timestamp === 'dateDash' || uiState.timestampTime === 'hhmm');
                if (!isProtectedDate && FEATURE_SEPARATOR) {
                    // 区切り記号統一。FEATURE_DOT_NORMALIZE が true なら "." も対象
                    if (uiState.separator === '-') {
                        value = FEATURE_DOT_NORMALIZE ? value.replace(/[_.]/g, '-') : value.replace(/_/g, '-');
                    } else if (uiState.separator === '_') {
                        value = FEATURE_DOT_NORMALIZE ? value.replace(/[-.]/g, '_') : value.replace(/-/g, '_');
                    }
                }
                pieces.push(value);
            }
            var result = pieces.join(defaultSep);
            // スペース（半角・全角・タブ・改行）の処理は後段の cleanFilenameChars に集約
            return result;
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
        function buildSortPanel(parent, prefs, currentOrderAvailable) {
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

        /* 「モード」パネルを構築（バージョン番号のみ / 全体）。常に「全体」が初期値
           / Build the scope panel (Version Only / Full); always defaults to "Full" */
        function buildOpModePanel(parent) {
            var panel = parent.add('panel', undefined, L('panel.opMode'));
            setupPanel(panel);
            var versionOnlyRadio = panel.add('radiobutton', undefined, L('radio.opVersionOnly'));
            var fullRadio = panel.add('radiobutton', undefined, L('radio.opFull'));
            versionOnlyRadio.value = false;
            fullRadio.value = true;
            return {
                panel: panel,
                versionOnlyRadio: versionOnlyRadio,
                fullRadio: fullRadio,
                isVersionOnly: function () { return versionOnlyRadio.value; },
                getOpMode: function () { return versionOnlyRadio.value ? 'versionOnly' : 'full'; }
            };
        }

        /* モード選択パネルを構築（リネーム / 別名で保存 / コピーを保存） / Build the mode panel (Rename / Save As / Save a Copy) */
        function buildModePanel(parent, prefs) {
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
            finalNameValue.preferredSize.width = Math.max(currentNameWidth + 20, NEW_NAME_FIELD_WIDTH + 150);

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
            var baseLabel = baseRow.add('statictext', undefined, labelText('label.base'));
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
                statusLabel = statusRow.add('statictext', undefined, labelText('label.status'));
                statusLabel.helpTip = L('tip.status');
                statusDropdown = statusRow.add('dropdownlist', undefined, undefined);
                statusDropdown.helpTip = L('tip.status');
                var parsedStatus = getFirstSegmentValue(segments, 'status');
                var initialStatus = parsedStatus || ((prefs && prefs.status) ? prefs.status : '');
                var initialStatusIdx = -1;
                for (var siItem = 0; siItem < STATUS_ITEMS.length; siItem++) {
                    var sItem = STATUS_ITEMS[siItem];
                    var ddItem = statusDropdown.add('item', (lang === 'ja' ? sItem.ja : sItem.en));
                    if (isStatusDivider(sItem)) {
                        ddItem.enabled = false;
                        continue;
                    }
                    if (initialStatusIdx === -1 && sItem.value === initialStatus) {
                        initialStatusIdx = siItem;
                    }
                }
                if (initialStatusIdx === -1) initialStatusIdx = 0;
                statusDropdown.selection = initialStatusIdx;
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

            // バージョン番号（なし / v1, v2… / v01, v02… / v001, v002…。デフォルト v1, v2…）
            // ES3 で 'short' を裸キーにすると予約語エラーになるため、ローカル変数名は short_ にする
            var versionRow = addRadioRow(panel, 'label.version', 'tip.version', [
                { key: 'none', text: L('radio.versionNone') },
                { key: 'short_', text: L('radio.versionShort') },
                { key: 'padded', text: L('radio.versionPadded') },
                { key: 'paddedWide', text: L('radio.versionPaddedWide') }
            ]);
            // 元ファイルに v 番号が無ければ常に 'padded' を強制（保存済み prefs を上書き）
            var hasOriginalVersion = !!getFirstSegmentValue(segments, 'version');
            var initialVersion = !hasOriginalVersion ? 'padded' :
                pickPref(prefs, 'version', ['none', 'short', 'padded', 'paddedWide'], 'padded');
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
                /* STATUS_ITEMS の value（'' = なし）。FEATURE_STATUS=false なら常に '' */
                getStatus: function () {
                    if (!FEATURE_STATUS || !statusDropdown) return '';
                    var idx = statusDropdown.selection ? statusDropdown.selection.index : 0;
                    var item = STATUS_ITEMS[idx];
                    if (!item || isStatusDivider(item)) return '';
                    return item.value;
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
            var dlg = new Window('dialog', L('sort.title'));
            dlg.opacity = DIALOG_OPACITY;
            dlg.orientation = 'column';
            dlg.alignChildren = 'fill';
            dlg.margins = 15;
            dlg.spacing = 10;

            var hint = dlg.add('statictext', undefined, L('sort.hint'));
            hint.alignment = 'left';

            var body = dlg.add('group');
            body.orientation = 'row';
            body.alignChildren = ['fill', 'fill'];

            var order = initialOrder.slice();
            var list = body.add('listbox', undefined, []);
            list.preferredSize = [200, 140];
            for (var i = 0; i < order.length; i++) {
                list.add('item', L('label.' + order[i]));
            }
            list.selection = 0;

            var btnCol = body.add('group');
            btnCol.orientation = 'column';
            btnCol.alignChildren = 'fill';
            var upBtn = btnCol.add('button', undefined, '↑');
            upBtn.preferredSize = [36, 24];
            var downBtn = btnCol.add('button', undefined, '↓');
            downBtn.preferredSize = [36, 24];

            function refreshList(newIdx) {
                for (var i = 0; i < order.length; i++) {
                    list.items[i].text = L('label.' + order[i]);
                }
                list.selection = newIdx;
            }
            upBtn.onClick = function () {
                var idx = list.selection ? list.selection.index : -1;
                if (idx <= 0) return;
                var tmp = order[idx - 1];
                order[idx - 1] = order[idx];
                order[idx] = tmp;
                refreshList(idx - 1);
            };
            downBtn.onClick = function () {
                var idx = list.selection ? list.selection.index : -1;
                if (idx < 0 || idx >= order.length - 1) return;
                var tmp = order[idx + 1];
                order[idx + 1] = order[idx];
                order[idx] = tmp;
                refreshList(idx + 1);
            };

            var buttons = dlg.add('group');
            buttons.alignment = ['center', 'top'];
            buttons.add('button', undefined, L('button.cancel'), { name: 'cancel' });
            buttons.add('button', undefined, 'OK', { name: 'ok' });

            if (dlg.show() !== 1) return null;
            return order;
        }

        /* ES3 互換: 配列に値が含まれるか / ES3-safe array contains check */
        function arrayContains(arr, val) {
            for (var i = 0; i < arr.length; i++) {
                if (arr[i] === val) return true;
            }
            return false;
        }

        /* prefs[key] が validValues 内ならそれを、そうでなければ fallback を返す
           / Returns prefs[key] if it's in validValues, else fallback */
        function pickPref(prefs, key, validValues, fallback) {
            var v = prefs && prefs[key];
            return arrayContains(validValues, v) ? v : fallback;
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
            var tip = L(tipKey);
            var label = row.add('statictext', undefined, labelText(labelKey));
            label.helpTip = tip;
            var radios = {};
            for (var i = 0; i < radioDefs.length; i++) {
                var def = radioDefs[i];
                var r = row.add('radiobutton', undefined, def.text);
                r.helpTip = tip;
                radios[def.key] = r;
            }
            return { group: row, label: label, radios: radios };
        }

        /* 各コントロールの種類に応じた変更イベントに callback を割り当てる。null/undefined はスキップ。
           radiobutton/checkbox → onClick、edittext → onChanging、dropdownlist → onChange
           / Wire a refresh callback to each control's appropriate change event */
        function wireRefresh(callback, controls) {
            for (var i = 0; i < controls.length; i++) {
                var c = controls[i];
                if (!c) continue;
                var t = c.type;
                if (t === 'radiobutton' || t === 'checkbox') c.onClick = callback;
                else if (t === 'edittext') c.onChanging = callback;
                else if (t === 'dropdownlist') c.onChange = callback;
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
            var mode = buildModePanel(topRow, prefs);
            var opMode = buildOpModePanel(topRow);
            // 現在のファイル名から検出した並び順（要素ゼロなら「現在に準じる」は無効化）
            var currentOrder = FEATURE_SORT ? deriveOrderFromSegments(segments) : [];
            var sort = FEATURE_SORT ? buildSortPanel(topRow, prefs, currentOrder.length > 0) : null;

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
                if (sortMode === 'on') segmentOrder = customOrder;
                else if (sortMode === 'current') segmentOrder = currentOrder;
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
                // 「バージョンのみ」モードでは UI 整形を一切かけず、元ファイル名の v 番号だけ更新
                if (opMode.isVersionOnly()) {
                    var versionOnlyBase = bumpVersionInPlace(stripExtension(currentName));
                    versionOnlyBase = nextAvailableVersionName(versionOnlyBase, folder, '.ai');
                    filename.finalNameValue.text = versionOnlyBase + '.ai';
                    return;
                }
                var st = currentUIState();
                var finalBase = buildFinalName(segments, st);
                if (st.version === 'short' || st.version === 'padded' || st.version === 'paddedWide') {
                    finalBase = nextAvailableVersionName(finalBase, folder, '.ai');
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
            var tr = options.titleRow, ts = options.timestampRow, vr = options.versionRow;
            var sr = options.separatorRow, nr = options.nfcRow, cr = options.cleanRow, lr = options.translitRow;
            wireRefresh(refreshPreviews, [
                tr.radios.none, tr.radios.parent, tr.radios.grandparent, tr.radios.custom,
                options.baseField, options.titleField,
                options.statusDropdown,
                ts.radios.none, ts.radios.date, ts.radios.dateDash,
                options.timestampHHMMCheckbox,
                vr.radios.none, vr.radios.short_, vr.radios.padded, vr.radios.paddedWide,
                sr && sr.radios.noChange, sr && sr.radios.dash, sr && sr.radios.underscore,
                nr && nr.radios.keep, nr && nr.radios.combine,
                cr && cr.radios.remove, cr && cr.radios.dash, cr && cr.radios.underscore,
                options.halfwidthKanaCheckbox,
                lr && lr.radios.keep, lr && lr.radios.remove, lr && lr.radios.convert
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
            var info = gatherDocumentInfo(doc);
            var segments = mergeFragmentedText(parseFileName(info.baseName));
            var prefs = loadPrefs();

            var ui = createDialog(segments, info.currentName, prefs, info.parentFolderName, info.grandparentFolderName, info.folder);
            if (ui.dialog.show() !== 1) return; // キャンセル

            var uiState = ui.getUIState();
            var newBaseName = (uiState.opMode === 'versionOnly')
                ? bumpVersionInPlace(info.baseName)
                : buildFinalName(segments, uiState);
            if (!newBaseName) {
                alert(L('message.emptyName'));
                return;
            }

            var targetFolder = ensureTargetFolder(info.folder);
            if (!targetFolder) return; // キャンセル

            if (uiState.opMode === 'versionOnly' || uiState.version === 'short' || uiState.version === 'padded' || uiState.version === 'paddedWide') {
                newBaseName = nextAvailableVersionName(newBaseName, targetFolder, '.ai');
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
                var msg = L('message.confirmTooLong')
                    .replace('{bytes}', String(fullByteLength))
                    .replace('{limit}', String(FEATURE_MAX_FILENAME_BYTES));
                if (!confirm(msg + '\n\n' + newBaseName + '.ai')) return;
            }

            if (!confirmOverwriteIfExists(destFile, info.fsPath)) return;

            try {
                executeOutput(doc, destFile, ui.getMode(), info.fsPath);
                // 成功したら今回の選択をプリセットとして保存（versionOnly モードでは保存しない）
                if (uiState.opMode !== 'versionOnly') {
                    var prefsToSave = {
                        titleMode: uiState.titleMode,
                        timestamp: uiState.timestamp,
                        timestampTime: uiState.timestampTime,
                        version: uiState.version
                    };
                    if (FEATURE_STATUS) prefsToSave.status = uiState.status;
                    if (FEATURE_SEPARATOR) prefsToSave.separator = uiState.separator;
                    if (FEATURE_NFC) prefsToSave.nfc = uiState.nfc;
                    if (FEATURE_CLEAN) prefsToSave.clean = uiState.clean;
                    if (FEATURE_HALFWIDTH_KANA && FEATURE_CLEAN) prefsToSave.halfwidthKana = uiState.halfwidthKana;
                    if (FEATURE_TRANSLITERATE) prefsToSave.translit = uiState.translit;
                    if (FEATURE_SORT) {
                        // sort モード自体は復元しない（毎回「現在のファイル名に準じる」が初期値）
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
