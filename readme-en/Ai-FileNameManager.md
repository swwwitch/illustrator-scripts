# Rename and Save

[![Direct](https://img.shields.io/badge/Direct%20Link-Ai--FileNameManager.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/document/Ai-FileNameManager.jsx)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

You open `catalog-v3.ai`, do your work, and then wonder how to make the next version. Switching to Finder, duplicating the file, renaming it to `catalog-v4.ai`, reopening it… that round trip is quietly tedious.

And filenames, left alone, tend to grow into things like `ｶﾀﾛｸﾞ 最新㈱ (2).ai` — half-width kana, full-width spaces, circled numbers, emoji. None of them hurt in the moment, but all of them hurt later.

So this script **breaks the open document's filename into segments** and lets you reassemble it from the UI. Tune it while watching the preview, and save right away.

<img alt="" src="" width="50%" />

## How to use

1. Open the document whose filename you want to change
2. Run the script
3. Adjust the settings in the dialog while watching the preview
4. Click [OK]

The destination is **always the same folder as the current file**. For an unsaved document, you are asked for a destination folder before the dialog opens.

## Mode: three ways to save

| Mode | Behavior |
| --- | --- |
| Rename Original | Save under the new name, then move the original to the Trash |
| Save As | Save under the new name. The original is kept, and the active document switches to the new file |
| Save a Copy | Save over the original, then create a copy under the new name. The active document stays on the original |

The default selection is always "Save As".

The difference shows up in **which file you keep editing after saving**. "Save As" moves you to the new file, while "Save a Copy" keeps you on the original. If you just want to leave a snapshot at a milestone and keep working, use the latter.

Only "Rename Original" needs care. Other documents that reference the original file (placed .ai, InDesign links, etc.) will lose their link. Note that it moves the file to the Trash rather than deleting it, so it is recoverable. If a file of the same name already exists in the Trash, a counter is appended to move it aside safely; and in the rare case where "the copy to the Trash succeeded but the original could not be removed", the copy is cleaned up so the file is not left duplicated.

## Scope: Full / Version Only

| Scope | Behavior |
| --- | --- |
| Full | Apply all formatting |
| Version Only | Bump only the v-number in the original filename; apply no other formatting |

"Version Only" is for cases like turning `report-final-v03.ai` into `report-final-v04.ai` and nothing else. The digit width is preserved. If there is no v-number, `-v2` is appended.

In this scope, the "Filename Settings" panel and the "Segment Order" panel are hidden entirely. Since no formatting is applied, configuring them would be meaningless. Presets are not saved either.

If the original file **has a v-number but no timestamp**, "Version Only" is preselected. Otherwise it defaults to "Full".

## Filename Settings

The filename is broken into six segments — base / project name / status / timestamp / sequence / version — and each can be edited independently.

### Base

The leading part of the filename. The value detected from the original filename is used as the initial value. Leave it empty to omit it.

**The leading token is always treated as base.** That is, the `20260712` in `20260712-catalog.ai` is base, not a timestamp. Interpreting it as a date would break the filenames of anyone who uses a date-first naming convention.

### Project Name

| Choice | Value |
| --- | --- |
| None | Add nothing |
| Custom | The text in the input field |
| Parent Folder | The name of the folder the file lives in |
| Grandparent Folder | The name of the folder one level above that |

The part of the original filename that is not base / status / date / sequence / version is detected as the project name and placed in the "Custom" input field. If nothing is detected, "None" is forced. If the corresponding folder does not exist, "Parent Folder" / "Grandparent Folder" are disabled.

### Status

Choose a production status from the dropdown. It is shown as `wip: Work in progress`, but **only the `wip` before the ":" is written to the filename**.

The provided values are wip / draft / review / revised / updated / fixed / approved / rejected / archived, then — across a divider — flattened / outlined. If the original filename contains one of these words, it is auto-detected.

### Timestamp

Choose from "None", "YYYYMMDD", and "YYYY-MM-DD" (default YYYYMMDD). "Append HHMM" adds `-HHMM` at the end — handy for multiple versions in a single day.

Choosing "None" removes any existing date in the original filename.

The date is **always today's date**. It does not carry over the original date.

### Sequence

Adds a sequence segment such as `page01` / `page001`. You can configure the prefix (default `page`) and the width (01 / 001).

If the original filename contains `pageNN`, it is auto-detected — the prefix and width are reflected as initial values, and the option is **forced on, overriding the preset**. This prevents an existing sequence from being silently dropped.

### Version

| Choice | Behavior |
| --- | --- |
| None | Add nothing (removes any existing v-number) |
| v1, v2… | No padding |
| v01, v02… | 2-digit zero-padding |
| v001, v002… | 3-digit zero-padding |

The default is "None". Only when the original file has a v-number does the previous selection (preset) take priority.

An existing v-number is bumped by +1; if there is none, numbering starts from v1 / v01 / v001.

## Auto-increment

Sequence and version numbers are incremented by **actually scanning the same folder**.

For example, if you have `report-v03.ai` open and `report-v07.ai` already exists in the same folder, the result becomes `report-v08.ai`, not `report-v04.ai`. The scan targets ".ai files with the same prefix / suffix", so unrelated files do not pull the number around.

If the computed number would be less than or equal to the current one, no bump is applied. And if scanning the folder fails, the result never falls below the existing number.

The bump is matched against the **finalized (formatted) name** — after cleaning, transliteration, and separator unification — so the number always agrees with the file that is actually saved. (Since existing filenames in the folder are stored already-formatted, matching against a pre-formatting name could miss the maximum and fail to bump.)

## Filename formatting

Everything below handles "things that feel wrong in a filename" in one pass.

| Item | Choices | Default |
| --- | --- | --- |
| Separator | No change / `-` / `_` | `-` |
| NFC normalization | No change / Combine | Combine |
| Clean filename | Remove / Change to `-` / Change to `_` | Change to `-` |
| Half-width → Full-width kana | Checkbox | On |
| Symbols (circled numbers, corporate abbrev., etc.) | No change / Remove / Convert | Convert |

### Separator

Unifies `-`, `_`, and `.` to the chosen symbol. `YYYY-MM-DD` timestamps (and those with `-HHMM` appended) have their internal `-` protected, so unifying to `_` does not turn them into `2026_07_17`.

Consecutive separators are collapsed to one, and leading/trailing separators and spaces are removed.

### NFC normalization

macOS filenames sometimes have "が" decomposed into "か" + a combining dakuten (two characters — NFD). This restores it to the single combined character (NFC). Non-combinable sequences such as "あ゙" are left as-is.

### Clean filename

Handles OS-invalid characters (`\ / : * ? " < > |`), emoji, platform-dependent characters, and spaces in one pass. "Remove" deletes them; "Change to `-`" replaces them with `-`.

Consecutive spaces are **collapsed to one first**, then replaced. Otherwise, with `-` selected, "A  B" would become `A--B`.

The check is allow-list based: ASCII printable characters, CJK symbols, kana, kanji, Hangul, full-width ASCII, and half-width katakana are treated as "standard", and everything else is a replacement target.

### Half-width → Full-width kana

Converts `ｶﾀﾛｸﾞ` to `カタログ`. Voiced/semi-voiced forms such as `ｶﾞ` / `ﾊﾟ` are merged as well.

This feature is active **only when Clean filename is set to "Change to `-`" or "Change to `_`"**. With "Remove", the half-width kana would be deleted, so converting them would be pointless.

### Symbols (circled numbers, corporate abbrev., etc.)

Converts to ASCII equivalents: ㈱→株, ①→1, Ⅰ→1, ℡→TEL, №→No, ㎜→mm, ～→~, —→-, and so on. Circled numbers cover white (①–⑳), black (❶–❿, ⓫–⓴), and parenthesized (⑴–⒇). Eight kinds of dashes are folded to a plain hyphen.

Choosing "Remove" drops these characters instead of converting them.

The processing order is **half-width kana → convert → clean**. If clean ran first, ㈱ would be removed as a "platform-dependent character", losing the chance to convert it to 株.

## Segment Order

| Choice | Behavior |
| --- | --- |
| Default | base → project name → status → timestamp → sequence → version |
| Match Current | The order the segments actually appear in the original filename |
| Custom | Reorder with ↑↓ via the [Edit] button |

The default is "Match Current" (or "Default" if nothing is detectable). This one is reset every time regardless of the preset.

When "Match Current" is selected and you change a segment (base / project name / status / timestamp / sequence / version), it is **automatically demoted to "Default"**. The order is based on the original filename, so changing its contents breaks that premise. Formatting options (separator, clean, etc.) do not affect the order and therefore do not trigger a demotion.

Note that while "Match Current" reflects the order of detected segments, a segment that was not in the original filename (such as the sequence) is filled in at its canonical position. This ensures a segment you turned on does not go missing from the output.

## Target

The active Illustrator document (saved as .ai).

## Notes

**When the filename is too long**: If the UTF-8 byte length including the extension exceeds 240 bytes, a confirmation dialog appears. You can continue.

**Windows reserved names**: If the name ends up as `CON` `PRN` `AUX` `NUL` `COM1`–`COM9` `LPT1`–`LPT9`, a `_` is appended to avoid the collision.

**The destination folder is fixed before the dialog opens.** This keeps the incremented number shown in the preview consistent with the file actually saved.

**Your previous settings are saved as a preset.** They are stored in `FileNameManager-prefs.txt` under `Folder.userData` as `key=value` text. It is written only when the save succeeds.

**Feature switches at the top of the script**: The status, ordering, separator, NFC, clean, convert, half-width kana, and sequence features each disappear — UI and all — when the corresponding `FEATURE_*` at the top of the script is set to `false`. Useful if you dislike a crowded dialog full of features you never use.

## Changelog

- v1.0 (2026-05-27) Initial release
- v1.3.5 (2026-07-12) Fixed mergeFragmentedText so it no longer folds recognized segments (page/date/status/version) back into the title (only text values are merged). When an existing pageNN is detected, the sequence is forced on over the preset. The sequence starting candidate is now "existing + 1", and never falls below the existing number even if the folder scan fails. Right-aligned the labels in the "Filename Settings" panel. Hardened loadPrefs to strip trailing line endings from CRLF-saved prefs. moveToTrash now cleans up the trashed copy if removing the original fails, preventing duplicates. Documented that the leading token is always base (a leading date stays base) and that the version default is "None". Renamed internal variables/functions to more self-explanatory names (no behavior change). Updated the overview comments to match the current implementation.
- v1.3.6 (2026-07-23) Fixed version/sequence auto-increment to run **after** formatting (clean, convert, separator unification). Since existing filenames in the folder are stored already-formatted, matching against a pre-formatting name could miss the maximum and fail to bump — resolved for both the preview and the actual save. Consolidated the shared UI layout settings (`setupWindow` / `setupPanel`, etc. and the margin constants) into one place, and unified the window margins of the main dialog and the segment-order sub-dialog.

### note

- [An Illustrator script to rename and save the current document](https://note.com/dtp_tranist/n/nc88dd887eb1c)
