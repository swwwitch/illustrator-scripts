# SmartFreeDistort.jsx

[![Direct Link](https://img.shields.io/badge/Direct%20Link-SmartFreeDistort.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/fx/SmartFreeDistort.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-e95464.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartFreeDistort.md)

[![Back to home](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

## Overview

This script applies Illustrator's Free Distort live effect to selected objects. Choose from 18 icon-based trapezoid, parallelogram, triangle, and diagonal presets.

Adjustable presets support amount and strength controls, with an Undo-based preview for checking the result before applying the effect.

## Main features

- Shape icons for every preset
- Four trapezoid presets
- Eight parallelogram presets (four anchor corners by two axes)
- Four triangle presets
- Two diagonal presets
- Adjustable amount from `0.00` to `0.49`
- Mild, Normal, and Boost strength options
- Undo-based preview
- Application to multiple selected objects
- Japanese and English UI

## Usage

1. Select the objects to distort.
2. Run `SmartFreeDistort.jsx`.
3. Choose a preset in the dialog.
4. For a trapezoid or parallelogram preset, adjust the amount and strength.
5. Enable Preview if you want to check the result.
6. Click OK to apply the live effect.

## Presets

| Type | Presets | Amount and strength |
| --- | ---: | --- |
| Trapezoid | 4 | Used |
| Parallelogram | 8 | Used |
| Triangle | 4 | Not used |
| Diagonal | 2 | Not used |

## Strength

| Setting | Factor |
| --- | ---: |
| Mild | 0.25 |
| Normal | 0.5 |
| Boost | 1.0 |

When the selection contains text, Mild is selected by default to reduce distortion of the letterforms.

## Notes

- Compatible with Illustrator 2024–2026.
- Selection entries that cannot receive a live effect are excluded.
- For multiple selected objects, the same live effect is applied to each object individually.
- Preview uses Illustrator's Undo history. Mixing it with other operations may produce an unexpected history state.
- Amount and strength settings are ignored by the triangle and diagonal presets.

## Article

[Apply Free Distort easily with an Illustrator script (Japanese)](https://note.com/dtp_tranist/n/n15a7ae196a23)

## Changelog

- v1.5.4 (2026-07-21): Added Japanese and English READMEs and linked them from the script's basic information
- v1.5.3 (2026-07-21): Updated the overview for the current preset count and application flow
- v1.5.2 (2026-07-21): Improved preview cleanup, target detection, and partial-application reporting
- v1.5.1 (2026-07-21): Added icons for every preset and expanded parallelograms to eight presets
- v1.5.0 (2026-07-21): Refined preview handling, target filtering, and preset definitions
- v1.1.1 (2026-04-24): Initial version
