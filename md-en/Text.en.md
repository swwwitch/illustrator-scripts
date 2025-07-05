# Text | Adobe Illustrator Scripts

![Twitter URL](https://img.shields.io/twitter/url?label=%40DTP_Transit&style=social&url=https%3A%2F%2Ftwitter.com%2FDTP_Tranist) ![Twitter URL](https://img.shields.io/twitter/url?label=%40swwwitch&style=social&url=https%3A%2F%2Ftwitter.com%2Fswwwitch)

- Adjust Baseline Vertical Center

## Adjust Baseline Vertical Center

[![Direct](https://img.shields.io/badge/Direct%20Link-AdjustBaselineVerticalCenter.jsx-FFcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/b62f2d91d5347a0c1208b9d92bd44a98e8d90938/jsx/AdjustBaselineVerticalCenter.jsx)

Adjusts the baseline (vertical position) of one or more specified characters in selected text frames to align vertically with a reference character. You can specify one or multiple target characters; the script automatically suggests the most frequent symbol if multiple are present, and manual override is also possible.

### Workflow:

1.	Specify the target character(s) and reference character in a dialog (the target character is auto-filled; if there are multiple options, the most frequent symbol is chosen. Manual override is also supported).
2.	Duplicate and create outlines to compare center Y coordinates.
3.	Apply the calculated offset to all target characters.

### Applicable to:

- Text frames (multiple selections supported; batch application possible)

### Original idea:

Egor Chistyakov https://x.com/tchegr

### Changes from the original:

- The target character is auto-filled (the most frequent symbol is chosen if there are multiple).
- Manual override of the target character is possible.
- Supports batch adjustment across multiple text objects.
- Supports specifying multiple target characters at the same time.

![](../png-en/ss-536-392-72-20250704-053323.png)

