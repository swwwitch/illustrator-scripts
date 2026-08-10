# AiSmartRotateView.jsx

[![Direct](https://img.shields.io/badge/Direct%20Link-AiSmartRotateView.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/misc/AiSmartRotateView.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AiSmartRotateView.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

## Overview

A palette for checking and changing both the active view rotation and the "constrain angle" preference (the angle used while the Shift key is held).

With "Link to view rotation" on, working with the view rotated 30 degrees automatically seeds the constrain angle with the same 30 degrees, so the view and operation angles match for easier work.

You can also read the angle of a selected object and rotate the view to it, or rotate the selection to match the view rotation.

## Features

- **Active view rotation angle**
  - Shows the current rotation angle
  - Slider for -180° to 180° (the display updates while dragging; the view is rotated on release)
  - Hold Shift while dragging to snap to 15° steps
  - [Reset] returns only the view rotation to 0° (dimmed when it is already 0°)
- **Constrain angle**
  - Set it with the number field (↑↓ for ±1, Shift+↑↓ for ±10) or the slider
  - The slider updates the field while dragging and applies the preference on release (15° steps with Shift)
  - [Change constrain angle] applies the field value to the preference
  - With "Link to view rotation" on, the same angle is applied automatically whenever the view rotation changes
  - [Reset] returns only the constrain angle to 0° (dimmed when it is already 0°)
- **Selected object**
  - Shows the selected path's angle (slope of the line between the first two anchor points)
  - [Rotate view to match selection]
  - [Rotate selection to match view] (dimmed while the view rotation is 0°)
  - [Reset selection rotation] (undoes the accumulated rotation stored in the `BBAccumRotation` tag)
- **Reset**
  - [Reset text tilt] (runs `transform/ResetText.jsx`)
  - [Reset image tilt] (runs `transform/ResetRotation.jsx`)
- A [Refresh] button, plus an automatic refresh whenever the palette becomes active
- Japanese / English UI

## How to use

1. Open a document.
2. Run `AiSmartRotateView.jsx` to show the palette.
3. Use the sliders and buttons to adjust the view rotation, the constrain angle, and the selected object's angle.
4. Keep the palette open while you work. Changes made in Illustrator itself are picked up when you click the palette to activate it or press [Refresh].

## About the constrain angle

The constrain angle is the Preferences > General setting that defines the base angle for moving, rotating, and drawing while the Shift key is held.

Illustrator stores the actual constraint direction in `constrain/sin` and `constrain/cos`, so this script writes all three values, including `constrain/angle` (which is what the Preferences dialog displays). Writing `constrain/angle` alone does not affect the constraint behavior.

## Notes

- Because this is a palette (persistent window), DOM access (reading the view rotation, applying the preference, and so on) is delegated to the main engine via BridgeTalk on each action.
- [Reset text tilt] and [Reset image tilt] run scripts in the `jsx/transform` folder of the same repository, so keep the folder structure intact.

## Script info

- Version: v1.0.0
- First release: 2026-06-05
- Last updated: 2026-08-11
