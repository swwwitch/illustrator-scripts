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
- Numeric fields step with the arrow keys (±10 snapping to multiples of 10 with Shift, ±0.1 with Option)
- Preview is supported: the original text and circle are kept until OK, which then deletes them and selects the result

### Article

[DTP Transit 別館 (Japanese)](https://note.com/dtp_tranist/n/na9334a217ec3)

### Update history

- v1.0.0 (2026-06-12) : Initial release
- v1.0.1 (2026-09-07) : Separator scale and baseline are now applied by position instead of by character content (identical characters inside the source text are no longer affected). Esc and Enter now trigger Cancel and OK. Font size is clamped to Illustrator's limits. Line breaks in the source text are replaced with spaces. UI wording adjusted to match the actual behavior ("Circumference correction" is now "Correction", and the separator fields name the separator explicitly), and the arrow-key stepping is documented in every number field's tooltip. Naming, JSDoc, and layout aligned with the house rules
