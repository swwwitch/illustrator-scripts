# AGENTS.md

This repository contains Adobe Illustrator ExtendScript (`.jsx`).

## Language and compatibility

- ES3 only. Use `var`; never `let`/`const`
- Wrap the script body in an IIFE
- Compatible with Illustrator 2024–2026
- Never use ES3 reserved or future-reserved words as identifiers or as `LABELS` keys
- Prefer `doc.selection` over `app.selection`
- Preserve existing functionality; do not remove features

## File structure

Split the top of the file into separate `// ====` comment blocks, one per concern:

1. **基本情報 / Basic info** — the metadata block below
2. **ユーザー設定 / User Settings** — values you tweak to change behavior
   (protected layer names, options that start unchecked, external folder paths)
3. **レイアウト / Layout** — dialog metrics (panel margins/spacing, control widths, indents)
   and shared layout helpers such as `setupPanel()`
4. Any other concern gets its own block (session memory, temporary action settings, path display, …)
5. **ローカライズ / Localization**
6. **メイン処理 / Main**

Keep layout metrics out of User Settings: pixel values are not what a user edits to change behavior.

## Header comment (`### 概要` / `### Overview`)

- A few bullets, then point at the README for the full feature list
- **No GitHub direct links** (no `### GitHub` section, no source URLs)
- **No version history** — the changelog lives in `readme-ja/` and `readme-en/`
- Keep it to 概要 / 注意 / Overview / Notes

## Basic info block

`=` aligns at column 20 (16 characters after `var `), comments at column 54.
`SCRIPT_ARTICLE_URL` is exempt from the alignment and sits right after the README comments.

```js
// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SmartGridMaker";               /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.6.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-02-24";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-07-26";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/<ScriptName>.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/<ScriptName>.md
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/xxxxxxxx"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php
```

When adapting someone else's script, keep the original author in `SCRIPT_AUTHOR` and add
`var SCRIPT_MODIFIED = "Masahiro Takano (@swwwitch)";  /* 改変 / modified by */` on the same alignment.
Put the original credit (`@author` / `@discussion`) in a JSDoc block right after the metadata.

## SCRIPT_VERSION

**Never bump `SCRIPT_VERSION` on your own.** Change it only when explicitly told to.
Version numbers are tied to the READMEs and published articles.

- On a feature change, update `SCRIPT_UPDATED` only
- If a bump seems warranted, propose it and wait
- `jsx/link/LinkedImageManager.jsx` is fixed at v1.5.0

## Comments

- Inline comments are Japanese and English on one line: `/* ガイドの設定 / Set guide properties */`
- Add JSDoc to functions: `@param {型} 名前 - 説明`, `@returns {型} 説明`, `@returns {void}` when there is none.
  Lowercase primitives (`string` / `number` / `boolean`), element types on arrays (`string[]`),
  Adobe DOM type names (`File`, `Folder`, `Document`, `Window`, `PlacedItem`, …). Descriptions in Japanese
- **Exception:** functions serialized with `toString()` and sent through BridgeTalk to the main engine must
  **not** carry JSDoc blocks — `toString()` mangles them and the receiving `eval` fails with an
  "illegal return statement" error. Use one-line comments there and say so in the section comment
- Collapse runs of two or more blank lines into one

## LABELS and localization

- Localization goes through a nested `LABELS` object plus a `getLabel()` lookup helper.
  `getLabel()` is the standard name — not `L()`, not `getLocalizedText()`
- Nest by UI part: `dialog` / `panel` / `radio` / `checkbox` / `dropdown` / `fieldLabel` /
  `tooltip` / `button` / `alert` / `fallbackName`
- Short entries on one line: `key: { ja: "...", en: "..." }`
- Entries containing `\n`, or long text, expand across multiple lines

## Naming

- Names a third party can guess. Avoid bare generic nouns; expand names that are too short
  (`lang` → `uiLang`, `options` → `organizeOptions`, `panel` → `exclusionPanel`, `group` → `separatorGroup`)
- Applies to variables, panels, groups, and functions alike
- Leave these as they are: `i` / `j` / `k` (loop counters), `e` / `err` (catch clauses),
  `dx` / `dy` (coordinate deltas)

## UI

- **Do not change UI wording on your own — propose it first.**
- A 3×3 anchor / alignment picker is a custom widget drawn in a `button`'s `onDraw`, never nine radio buttons.
  Reference implementation: `addAnchorWidget()` / `drawAnchorWidget()` / `drawAnchorCell()` in
  `jsx/transform/QuickTransformPalette.jsx`
- ScriptUI controls do not grow after creation. Reserve the width a label will need
  (`preferredSize.width`) before text is assigned at runtime
- Radio buttons are only mutually exclusive within the same container. Radios split across containers
  must have their exclusivity managed by hand
