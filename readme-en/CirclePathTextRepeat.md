# Repeat text around a circle

[![Direct](https://img.shields.io/badge/Direct%20Link-CirclePathTextRepeat.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/CirclePathTextRepeat.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/CirclePathTextRepeat.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- Select one circle (path) and one text object; the text is repeated a given number of times and converted into type on a duplicate of the circle
- The separator is either a space or any character you type (a bullet `•` by default), and it is appended after the last repetition as well
- The number of half-width spaces around the separator can be set (with a space-only separator this becomes the gap between repetitions)
- For a non-space separator, its scale (horizontal and vertical) and baseline (in the preference text unit) can be adjusted
- Automatic font-size fitting to the circumference can be switched on or off, with a correction factor to fine-tune the gap between start and end
- The result is rotated about the circle's center
- Numeric fields step with the stepper buttons (∧∨) and the arrow keys (to the next whole number, e.g. 1.5 → 2; to the next multiple of 10 with Shift; ±0.1 with Option)
- Preview is supported: the original text and circle are kept until OK, which then deletes them and selects the result

### Article

[DTP Transit 別館 (Japanese)](https://note.com/dtp_tranist/n/na9334a217ec3)

### Update history

- v1.0.0 (2026-06-12) : Initial release
- v1.0.1 (2026-09-07) : Separator scale and baseline are now applied by position instead of by character content (identical characters inside the source text are no longer affected). Esc and Enter now trigger Cancel and OK. Font size is clamped to Illustrator's limits. Line breaks in the source text are replaced with spaces. UI wording adjusted to match the actual behavior ("Circumference correction" is now "Correction", and the separator fields name the separator explicitly), and the arrow-key stepping is documented in every number field's tooltip. Naming, JSDoc, and layout aligned with the house rules
- v1.1.0 (2026-09-27) : Added stepper buttons to the number fields. The arrow keys now share the steppers' logic (to the next whole number; Shift to the next multiple of ten)
- v1.1.1 (2026-09-28) : The dialog now reopens where it was last closed and moves sideways to avoid covering the selection; opacity unified at 97%
- v1.1.2 (2026-09-28) : The button row is now built with the shared part
- v1.1.3 (2026-09-29) : Dialog opacity changed to 98%
- v1.1.4 (2026-09-30) : Fixed an error when running with characters selected by the Type tool
- v1.1.5 (2026-10-01) Unified the window and panel margins and spacing with the shared layout part
- v1.1.6 (2026-10-01) Added space below the button row to match Illustrator's own dialogs
- v1.1.7 (2026-10-01) Typed values in number fields are now rounded and clamped like the stepper, and non-numbers revert to the previous value. Disabled fields dim their labels too. Shortened "Separator scale/baseline" to "Scale/Baseline". Integer fields no longer mention the Option 0.1 step in their tooltips
- v1.1.8 (2026-10-01) Fixed rotation pivoting on the text bounds instead of the circle center. Fitting now measures the actual path length, so non-circular paths fit too. The preview is now built from a duplicate with the originals hidden instead of with undo, so no measurement steps are left in the undo history. The scale field now has a 1% minimum (0 used to produce 100%). Forced line breaks are replaced with spaces too. Widened the rotation field. Arrow-key tooltips now say "snap to 10s" and "by 0.1"
- v1.1.9 (2026-10-04) Japanese labels now end with " :" (half-width space and colon) (shared part update)
