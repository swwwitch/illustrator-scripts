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

1. **基本情報 / Basic info** — the Script info part that reads SCRIPTMETA (see below)
2. **ユーザー設定 / User Settings** — values you tweak to change behavior
   (protected layer names, options that start unchecked, external folder paths)
3. **レイアウト / Layout** — dialog metrics (panel margins/spacing, control widths, indents)
   and shared layout helpers such as `setupPanel()`
4. Any other concern gets its own block (session memory, temporary action settings, path display, …)
5. **ローカライズ / Localization**
6. **メイン処理 / Main**

Keep layout metrics out of User Settings: pixel values are not what a user edits to change behavior.

## SCRIPTMETA (basic info)

A script's basic info lives only in the SCRIPTMETA block at the very top of the file, in the SCRIPTMETA 2.1
multilingual form ([MetadataAuthoring.md](https://github.com/Yamonov/SCRIPTMETAKit/blob/main/docs/MetadataAuthoring.md)).
There is no hand-written `var SCRIPT_* = "…"` block and no `### 概要` / `### Overview` header comment any more.
Update checking is not used: write no `Meta-URL` and no `SCRIPTMETA.txt`.

```js
/*
SCRIPTMETA-BEGIN
Script-ID=com.swwwitch.illustrator.<ScriptName>
Meta-Default-language=ja
Name=<README の見出し>
Name[en]=<English README title>
Author=Masahiro Takano (@swwwitch)
Target-App=Illustrator
Version=1.6.0
Release-Date=2026-02-24
Tags=<folder tags>
X-com.swwwitch-Script-Name=<ScriptName>
X-com.swwwitch-Updated=2026-07-26
X-com.swwwitch-README-JA=https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/<ScriptName>.md
X-com.swwwitch-README-EN=https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/<ScriptName>.md
X-com.swwwitch-Article-URL=https://note.com/dtp_tranist/n/xxxxxxxx
X-com.swwwitch-Source-URL=https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/<category>/<ScriptName>.jsx
Description-BEGIN
[ja]
（1〜2文の説明）
[en]
(one or two sentences)
Description-END
SCRIPTMETA-END
*/
```

- Edit by hand: `Version` (no `v`), `Author`, `Release-Date` (first release), the Description, and the
  `X-com.swwwitch-*` fields `Script-Name`, `Modified-By`, `Updated`, `README-JA` / `README-EN` (only when the README
  files exist), `Article-URL`, `Reference-URL`, `Pro-URL`. Leave out a field that has no value
- `X-com.swwwitch-Script-Name` becomes `SCRIPT_NAME`, which is also the key of saved settings and dialog positions.
  Changing it loses what users have saved
- Description: `[ja]` and `[en]` labels on their own lines, one or two sentences each, then point nowhere else
  (the README links are in the fields). Write a literal `[`, `]`, `\` as `\[`, `\]`, `\\`. Never `*/`
- Notes that do not fit in the Description (`### 注意` / `### Notes`) go in a `/* */` comment right after the SCRIPTMETA block
- When adapting someone else's script, `Author` is the original author and `X-com.swwwitch-Modified-By=Masahiro Takano (@swwwitch)`.
  Put the original credit (`@author` / `@discussion`) in a JSDoc block after the Script info part
- `python3 helpers/sync-scriptmeta.py` (`--check` reports files out of sync) fills in the rest. Run it after editing
  SCRIPTMETA or a README title, and after adding, moving or renaming a script:
  - `Script-ID` only when it is missing (`com.swwwitch.illustrator.<Script-Name>`; names with characters other than
    `A-Z a-z 0-9 . _ -` get an entry in `ID_OVERRIDES`). Never change a published Script-ID, even when renaming a script
  - `Name` / `Name[en]` from the `# ` titles of the READMEs that `README-JA` / `README-EN` point to. To change them, edit
    the README titles. A script without a README keeps a hand-written `Name` (a short phrase saying what it does)
  - `Meta-Default-language`, `Target-App`, `Tags` (from the top folder under `jsx/`, `FOLDER_TAGS` in the tool; a new folder
    needs an entry there), `X-com.swwwitch-Source-URL` (from the file's path), percent-encoding of non-ASCII in URLs,
    and the field order
  - the Script info part below, rewritten to match `jsx/_templates/ScriptMeta.jsx`
- `jsx/_templates/` is excluded

## Basic info (Script info part)

The 基本情報 / Basic info block holds the Script info part (`jsx/_templates/ScriptMeta.jsx`), at the top level outside the IIFE:

```js
// =========================================
// 基本情報 / Basic info
// =========================================
// スクリプト情報（再利用パーツ） / Script info (reusable)
…
// スクリプト情報（再利用パーツ）ここまで / End of the reusable script info

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php
```

It reads the SCRIPTMETA of `$.fileName` when the script loads and sets `SCRIPT_NAME`, `SCRIPT_VERSION` (`"v" + Version`),
`SCRIPT_AUTHOR`, `SCRIPT_RELEASED`, `SCRIPT_UPDATED`, `SCRIPT_DESCRIPTION_JA` (`Name`), `SCRIPT_README_JA` / `SCRIPT_README_EN`,
`SCRIPT_ARTICLE_URL` and `SCRIPT_SOURCE_URL` (empty string when a field is missing). Use those variables as before; do not
assign them by hand and do not edit the part (the sync tool rewrites it). In code, write the SCRIPTMETA marker strings split
(`"SCRIPTMETA" + "-BEGIN"`) so apps that scan for SCRIPTMETA do not pick them up.

## Version

Bump `Version` in SCRIPTMETA whenever a change alters behavior; do not wait to be told.
Version numbers are tied to the READMEs and published articles, so keep them in step.

- Bug fix or small behavior change: bump the patch (`1.6.0` → `1.6.1`; two-part versions grow a third part, `1.1` → `1.1.1`)
- Larger feature addition: bump the minor (`1.6.1` → `1.7.0`)
- Set `X-com.swwwitch-Updated` to that day and add the same version to the change log in `readme-ja/` and `readme-en/`
  (create the section when it is missing)
- Do not bump for links, comments, alignment or JSDoc only
- Bump once per unreleased change set: count from the version in the last commit, not from today's working copy
- New scripts start at `1.0.0`

## Units

Unit handling goes through one table, one accessor and two converters. Do not add per-script label maps,
`getPtFactorFromUnitCode()`-style helpers, or other conversion functions (`convertToPt`, `mmToPt`, `toPoints`, …).

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

/**
 * 単位の値を pt に換算する
 * @param {number} value - 単位の値
 * @param {{pointsPerUnit: number}} unit - getUnitInfo() の戻り値か UNITS[i]
 * @returns {number} pt の値
 */
function unitToPt(value, unit) {
    return value * unit.pointsPerUnit;
}

/**
 * pt の値を単位の値に換算する
 * @param {number} ptValue - pt の値
 * @param {{pointsPerUnit: number}} unit - getUnitInfo() の戻り値か UNITS[i]
 * @returns {number} 単位の値
 */
function ptToUnit(ptValue, unit) {
    return ptValue / unit.pointsPerUnit;
}

/**
 * 単位コードから、getUnitInfo() と同じ形の単位の情報を作る（コードを別に読んだとき・単位を選ぶポップアップ用）
 * @param {number} unitCode - 単位コード
 * @param {string} prefKey - そのコードの環境設定キー（"rulerType" など。Q/H の表示を決める）
 * @returns {{code: number, label: string, pointsPerUnit: number}} 単位の情報（未知のコードは pt 相当）
 */
function getUnitInfoFromCode(unitCode, prefKey) {
    /* 未知のコードは pt に寄せる / unknown codes fall back to points */
    var unit = UNITS[unitCode] || UNITS[2];
    /* 級（Q）と歯（H）は同じ長さだが、文字サイズは「Q」、距離は「H」と呼び分ける */
    var label = (unitCode === 5 && HA_UNIT_PREF_KEYS[prefKey]) ? "H" : unit.label;
    return { code: unitCode, label: label, pointsPerUnit: unit.pointsPerUnit };
}

/**
 * 表示ラベルから単位を探す（「H」は「Q」と同じ単位）
 * @param {string} label - 単位の表示ラベル（"mm" など）
 * @returns {{label: string, pointsPerUnit: number}|null} UNITS の要素（見つからなければ null）
 */
function findUnitByLabel(label) {
    var unitLabel = (label === "H") ? "Q" : label;
    for (var i = 0; i < UNITS.length; i++) {
        if (UNITS[i].label === unitLabel) return UNITS[i];
    }
    return null;
}
```

- Display label: `getUnitInfo().label`
- Convert with the unit object, not a bare factor: `unitToPt(value, getUnitInfo())`, `ptToUnit(ptValue, getUnitInfo("strokeUnits"))`.
  Keep the object in a variable named after its preference (`rulerUnit`, `strokeUnit`, `textUnit`), not `factor` / `ptPerUnit`
- Fixed units: `unitToPt(3, UNITS[1])` (3 mm). No `POINTS_PER_MM`-style constants or `72 / 25.4` literals
- Unit chosen in a dropdown: index into `UNITS` when the list follows the unit codes; otherwise `findUnitByLabel(label)`
- A unit code read elsewhere (a palette receives it from the main engine through BridgeTalk) or listed in a unit
  popup: `getUnitInfoFromCode(code, "rulerType")`, which gives the same label (H) and shape as `getUnitInfo()`
- `UNITS` rows carry only `label` and `pointsPerUnit`. Per-script extras (popup membership, display decimals)
  go in a separate array indexed by unit code (`UNIT_POPUP_CODES = [0, 1, 2, …]`, `KEY_INCREMENT_DECIMALS = […]`)
- Show the labels as `UNITS` / `getUnitInfo()` give them (`in`, `H`); do not rename them per script
- Another preference: `getUnitInfo("strokeUnits")`, `getUnitInfo("text/units")`
- `pointsPerUnit` appears only in `UNITS`, `getUnitInfo()`, `unitToPt()` and `ptToUnit()`
- Include only the parts the script uses: `getUnitInfo()` (always together with `HA_UNIT_PREF_KEYS`, which it reads)
  when it reads a preference, and `unitToPt()` / `ptToUnit()` / `getUnitInfoFromCode()` / `findUnitByLabel()` when called
  (`getUnitInfoFromCode()` also needs `HA_UNIT_PREF_KEYS`).
  Place them in this order right after `getUnitInfo()` (or after `UNITS` when there is no `getUnitInfo()`)
- Do not trim `getUnitInfo()` into a variant without the H label; copy it as written
- Rounding and display formatting stay outside the converters (`formatUnitValue()` etc. call them)
- Values typed into a stepper field (`1in`, `1p6`, `10 mm`) are read by the StepperButtons part
  (`evaluateArithmetic()` with `UnitValue`); outside that part, do not call `UnitValue` for conversion
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
- Labels carry no colon; add it at display time with `labelText()`: Japanese ends with ` :` (half-width space + colon,
  as in Illustrator's Stroke panel), English with `:`. `labelValueText()` adds a space after the colon (`件数 : 5` / `Count: 5`).
  Never write a full-width `：` into a label

## Naming

- Names a third party can guess. Avoid bare generic nouns; expand names that are too short
  (`lang` → `uiLang`, `options` → `organizeOptions`, `panel` → `exclusionPanel`, `group` → `separatorGroup`)
- Applies to variables, panels, groups, and functions alike
- Leave these as they are: `i` / `j` / `k` (loop counters), `e` / `err` (catch clauses),
  `dx` / `dy` (coordinate deltas)

## UI

- **Do not change UI wording on your own — propose it first.**
- Window and panel margins/spacing go through the UI layout part (`jsx/_templates/UILayout.jsx`,
  pasted from "（再利用パーツ）" to "ここまで" into the Layout block). Apply it whenever a dialog is created or
  modified, and in every cleanup pass: `setupWindow()` for the window, `setupPanel()` for **every** panel
  (`setupPanel(panel, 6)` for a column of radio buttons or checkboxes), no hand-written `margins` / `spacing`.
  `setupPanel()` sets `alignChildren` to fill, so give buttons inside a panel `alignment = "left"` so they do not stretch
- Build the bottom button row with the ButtonRow part (`jsx/_templates/ButtonRow.jsx`) and call
  `alignRightOnlyButtonRow(buttonRow)` after the last button is added (before `show()`): a row with only
  right-side buttons (Cancel / OK) is centered when the dialog's inner width (margins excluded) is 240 px or less
  (`BUTTON_ROW_CENTER_MAX_WIDTH`) and stays right-aligned in wider dialogs. The width is measured on the `show`
  event, so it works even when sizes are set after the call. A row with left-side buttons keeps the split layout
- Do not put `SCRIPT_VERSION` in a window title. Show it with the VersionFooter part (`jsx/_templates/VersionFooter.jsx`):
  call `addVersionFooter(window)` after every control is added, right before `show()` (before `prepareDialogWindow()`).
  It draws the script name and version (left) and "DTP Transit" (right) small and gray at the bottom left (dot glyphs in `onDraw`; needs the UITheme part for `isDarkUI()`) and takes over the button row's bottom margin.
  Progress, confirm and other small sub-windows get no footer
  Clicking the name and version opens the part's About This Script dialog (script name with a copy button, the Japanese README title,
  version, README / source code / article links, Show in Finder — in Path Finder when it is running, through
  `/Applications/OpenInFileViewer.app`). It reads the `SCRIPT_*` basic info with `typeof` and carries its own
  labels (`VERSION_INFO_LABELS`), so paste the whole part unchanged; do not add its labels to `LABELS`
- A 3×3 anchor / alignment picker is a custom widget drawn in a `button`'s `onDraw`, never nine radio buttons.
  Reference implementation: `addAnchorWidget()` / `drawAnchorWidget()` / `drawAnchorCell()` in
  `jsx/transform/QuickTransformPalette.jsx`
- ScriptUI controls do not grow after creation. Reserve the width a label will need
  (`preferredSize.width`) before text is assigned at runtime
- Radio buttons are only mutually exclusive within the same container. Radios split across containers
  must have their exclusivity managed by hand
