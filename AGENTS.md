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
- Put the README URL on the line right after the pointer: `SCRIPT_README_JA` under 概要,
  `SCRIPT_README_EN` under Overview. When `SCRIPT_ARTICLE_URL` exists, add the article under 概要 only.
  Copy the values from the basic info block and keep them there too. Skip the lines when the README
  files do not exist
- **No other GitHub links** (no `### GitHub` section, no source URLs)
- **No version history** — the changelog lives in `readme-ja/` and `readme-en/`
- Keep it to 概要 / 注意 / Overview / Notes

```js
/*

### 概要

（1〜2文の説明）

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/<ScriptName>.md

note記事も参照してください。
https://note.com/dtp_tranist/n/xxxxxxxx

### Overview

(one or two sentences)

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/<ScriptName>.md

*/
```

Bare URLs are safe inside this `/* */` block; the Error 11 label problem only affects `//` lines.

## Basic info block

`=` aligns at column 20 (16 characters after `var `), comments at column 54.
The README links and `SCRIPT_ARTICLE_URL` are exempt from that alignment: put them in a separate
group after a blank line and align `=` across those lines only (one space when `SCRIPT_ARTICLE_URL`
is absent).

Hold the README links in string variables, not in line comments. Two consecutive
`// https://…` lines make ExtendScript read `https:` as a duplicate label and the script dies with
Error 11 ("label not found"). Include the links only when the README files actually exist.

```js
// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SmartGridMaker";               /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.6.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-02-24";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-07-26";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/<ScriptName>.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/<ScriptName>.md"; /* README (English) */
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

## Units

Unit handling goes through one table and one accessor. Do not add per-script label maps or
`getPtFactorFromUnitCode()`-style helpers.

```js
// =========================================
// 単位 / Units
// =========================================

/* 単位コードに対応する表示ラベルと、1単位あたりのポイント数
   Unit code -> display label and points per unit */
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
 * 環境設定キーの単位を返す
 * @param {string} [prefKey] - "rulerType"（既定）/ "strokeUnits" / "text/units" / "text/asianunits"
 * @returns {{code: number, label: string, pointsPerUnit: number}} 単位の情報
 */
function getUnitInfo(prefKey) {
    var unitKey = prefKey || "rulerType";
    var unitCode = app.preferences.getIntegerPreference(unitKey);
    /* 未知のコードは pt に寄せる / unknown codes fall back to points */
    var unit = UNITS[unitCode] || UNITS[2];
    /* 級（Q）と歯（H）は同じ長さだが、文字サイズは「Q」、距離は「H」と呼び分ける */
    var label = (unitCode === 5 && HA_UNIT_PREF_KEYS[unitKey]) ? "H" : unit.label;
    return { code: unitCode, label: label, pointsPerUnit: unit.pointsPerUnit };
}
```

- Display label: `getUnitInfo().label`
- Points per unit: `getUnitInfo().pointsPerUnit`
- Another preference: `getUnitInfo("strokeUnits")`, `getUnitInfo("text/units")`
- Include only `UNITS` and `getUnitInfo()` when the script never shows a Q/H label;
  keep `HA_UNIT_PREF_KEYS` whenever a label reaches the UI.
- **Exception:** `jsx/preference/single-function/PreferenceManager-print-pt.jsx` keeps its own `["pt","pc",…,"Q/H","px"]`
  list. One dropdown serves four preference keys there, so unit code 5 needs the neutral `Q/H` label,
  and the file does no pt conversion at all. Leave it as it is.

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
