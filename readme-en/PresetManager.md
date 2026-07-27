# PresetManager

[![Direct](https://img.shields.io/badge/Direct%20Link-PresetManager.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/preference/PresetManager.jsx)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Description

- A utility that gathers the Illustrator preferences you actually touch often into a single dialog.
- Items that normally live in separate categories — General, Type, Guides, Performance, File Handling and so on — can be changed in one place without switching tabs.
- The dialog opens with the current preferences loaded. Choosing [Default] or [Preset 1] from the dropdown fills the UI with a whole set of values.
- Nothing is written while you work; every change is saved at once when [OK] is pressed (pressing [Cancel] changes nothing).

### Usage

1. Run the script; the dialog opens with the current preferences applied.
2. Change what you need. Choosing [Default] or [Preset 1] from the dropdown loads a complete set of values into the UI (still not saved).
3. Press [OK] to save.

### Preferences covered

#### [General] category

| Item | Preference key |
| --- | --- |
| Show Rich Tool Tips | showRichToolTips |
| Show the Home Screen When No Documents Are Open | Hello/ShowHomeScreenWS |
| Use Legacy "File > New" Interface | Hello/NewDoc |
| Show 'Print Bleed' generative AI buttons on Bleed | enablePrintBleedWidget |

#### [Selection & Anchor Display] category

| Item | Preference key |
| --- | --- |
| Zoom to Selection | zoomToSelection |

#### Artboard

| Item | Preference key |
| --- | --- |
| Move Locked and Hidden Artwork | moveLockedAndHiddenArt |
| Show Artboard Name | showArtboardLabelOnCanvas |
| Highlight Color | ArtboardBBColorRed / Green / Blue |
| Stroke Width (1–4) | ArtboardBBWidth |

The highlight color is picked from nine presets: Light Blue, Light Red, Green, Medium Blue, Magenta, Cyan, Light Gray, Black and Yellow. When the current value matches none of them, the preset with the smallest RGB difference is shown as selected.

#### [Type] category

| Item | Preference key |
| --- | --- |
| Auto Size New Area Type | text/autoSizing |
| Number of Recent Fonts | text/recentFontMenu/showNEntries |
| Enable Missing Glyph Protection | text/doFontLocking |
| Show Character Alternates | text/enableAlternateGlyph |

"Number of Recent Fonts" is a checkbox plus a numeric field. Unchecking it disables the field and saves 0 (list hidden). The accepted range is 0–30; out-of-range or non-numeric input falls back to 15.

#### [User Interface] category

| Item | Preference key |
| --- | --- |
| Canvas Color (Match Brightness / White) | uiCanvasIsWhite |

The four-swatch brightness control (`uiBrightness`) is **currently commented out and not shown**. Its values are discrete presets (0.0 / 0.5 / 0.50999999 / 1.0) rather than a continuous scale, and writing `uiBrightness` alone does not apply the change on screen — Preferences (User Interface) has to be opened after [OK] and confirmed with the arrow keys + Return. The code is kept in place, so removing the comment markers brings it back.

#### Guides

| Item | Preference key |
| --- | --- |
| Color (Cyan / Light Blue) | Guide/Color/red, green, blue |
| Style (Lines / Dots) | Guide/Style (0 = Lines / 1 = Dots) |

The color is either Cyan (0, 1, 1) or Light Blue (0.29, 0.52, 1.0). When the current value does not match Light Blue, Cyan is shown as selected.

#### Smart Guides

| Item | Preference key |
| --- | --- |
| Object Highlighting | smartGuides/showObjectHighlighting |

#### [Performance] category

| Item | Preference key |
| --- | --- |
| Animated Zoom | Performance/AnimZoom |
| History States | maximumUndoDepth |
| Real-Time Drawing and Editing | LiveEdit_State_Machine |

History States accepts 1–1000; out-of-range or non-numeric input falls back to 100.

#### [File Management] category

| Item | Preference key |
| --- | --- |
| Use System Defaults for 'Edit Original' | useSysDefEdit |
| Auto-activate Adobe Fonts | AutoActivateMissingFont |
| Save Location (Computer / Cloud) | AdobeSaveAsCloudDocumentPreference |
| Update Links (Automatic / Manual / Ask When Modified) | plugin/FileClipboard/linkoptions (0 = Automatic / 1 = Manual / 2 = Ask) |

#### Clipboard Handling

| Item | Preference key |
| --- | --- |
| Include SVG Code | plugin/FileClipboard/copySVGCode |

#### Limit to Path

| Item | Preference key |
| --- | --- |
| Object Selection by Path Only | hitShapeOnPreview |
| Type Object Selection by Path Only | hitTypeShapeOnPreview |

For these two keys 0 means ON and 1 means OFF, so the script inverts the value when reading and writing.

### Preset contents

| Item | Default | Preset 1 |
| --- | --- | --- |
| Show Rich Tool Tips | ON | OFF |
| Show the Home Screen | ON | OFF |
| Use Legacy "File > New" Interface | OFF | ON |
| 'Print Bleed' generative AI buttons | ON | OFF |
| Move Locked and Hidden Artwork | OFF | ON |
| Zoom to Selection | ON | OFF |
| Object Selection by Path Only | OFF | OFF |
| Type Object Selection by Path Only | OFF | OFF |
| Show Artboard Name | ON | OFF |
| Highlight Color | Light Blue | Black |
| Stroke Width | 1 | 2 |
| Auto Size New Area Type | OFF | ON |
| Number of Recent Fonts | 10 | 15 |
| Enable Missing Glyph Protection | ON | OFF |
| Show Character Alternates | ON | OFF |
| Canvas Color | Match Brightness | White |
| Guide Color | Cyan | Light Blue |
| Object Highlighting (Smart Guides) | ON | OFF |
| Animated Zoom | ON | OFF |
| History States | 100 | 50 |
| Real-Time Drawing and Editing | ON | OFF |
| Use System Defaults for 'Edit Original' | OFF | ON |
| Auto-activate Adobe Fonts | OFF | ON |
| Save Location | Cloud | Computer |
| Update Links | Ask When Modified | Automatic |
| Include SVG Code | OFF | ON |

[Default] roughly corresponds to Illustrator's out-of-the-box state; [Preset 1] is tuned for production work. Neither preset includes **UI brightness or guide style**, so those two items stay as they are when a preset is selected.

### Flow

1. Build the dialog and lay out the panels in two columns
2. Read the current preferences into the UI (`loadPreferencesIntoUI`)
3. Selecting a preset overwrites the UI with predefined values (nothing is written to preferences)
4. [OK] writes every preference key and opens the Preferences panels when required

### Notes

- Preference changes are not always reflected on screen right away, so [OK] runs zoom out → zoom in to force a redraw.
- Every preference read and write is wrapped in try/catch, so keys missing in a given Illustrator version do not stop the script; a failed read falls back to a default value.
- Closing with [Cancel] leaves the preferences untouched, even if a preset was selected.

### Change Log

- v1.0 (20250807): Initial release
- v1.6 (20260323): UI layout adjustments, localization improvements, removed incomplete features, fixed preference key (useSysDefEdit)
- v1.7.0 (20260422): Refined localization, naming, comments, and structure
- v1.8.0 (20260727): Reorganized naming, label definitions and UI layout (brightness swatches commented out, [Open File Handling] button removed)
