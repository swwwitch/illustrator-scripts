# ZoomToSelection

[![Direct](https://img.shields.io/badge/Direct%20Link-ZoomToSelection.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/misc/ZoomToSelection.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ZoomToSelection.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- Zooms and centers the active view on the selection.
- With several objects selected it fits their combined bounding box, leaving a little margin.
- With nothing selected it returns to 100% at the current position.
- The zoom is interpolated in small steps rather than jumping, giving an animated feel.

### Settings

- `ZOOM_FIT_RATIO`: margin factor when fitting (1.0 is exact, 0.9 leaves a 10% margin)
- `MAX_ANIMATION_STEP_COUNT`: upper bound on the number of interpolation steps
- `FRAME_DELAY_MS`: wait in milliseconds per frame

### Original idea

John Wundes - Zoom and Center to Selection v2.

http://www.wundes.com/js4ai/copyright.txt

The animation interpolation is based on ArtboardNavigator.jsx by Yuki Furushima.
