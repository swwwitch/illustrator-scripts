# Generate a long shadow from distance, angle and scale

[![Direct](https://img.shields.io/badge/Direct%20Link-LongShadowMaker.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/misc/LongShadowMaker.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/LongShadowMaker.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Generates a long shadow from the selected object using a distance, an angle and a scale. Adjust the values against a live preview and click OK to commit.

The shadow is built by bridging the original shape and a moved, scaled copy of it with faces. Its fill is the object's own fill with the saturation dialed down.

### Features

- Distance, angle and scale set from a number field and a slider (arrow keys step the value)
- Eight presets that set the scale and the angle in one go
- An offset that grows the shadow (miter, round or bevel corners)
- A Simplify pass that reduces the number of anchor points
- A preview built from stacked translucent copies
- Japanese and English UI

### Usage

1. Select a single object.
2. Run `LongShadowMaker.jsx`.
3. Set the distance, angle and scale. Picking a preset sets the scale and the angle together; the distance is left as it is.
4. Turn on Offset and Simplify if you need them.
5. Check the preview and click OK.

### Settings

| Setting | What it does | Default |
| --- | --- | --- |
| Distance | Length of the shadow, in points | Width plus height of the selection |
| Angle | Direction the shadow extends (-180° to 180°) | 45° |
| Scale | Size of the far end (1% to 300%); 100% matches the original | 100% |
| Offset | How much the generated shadow is grown, in points | Off (average of width and height divided by 20) |
| Join | Corner treatment used when growing the shadow | Round |
| Simplify | Reduces the anchor points in the shadow | On |

### Notes

- Works on a closed PathItem, a CompoundPathItem, a GroupItem made only of closed paths, or a TextFrame.
- Groups and text are merged or outlined into a temporary shape first; the original object is left untouched.
- The offset is not shown in the preview. It is applied when you click OK.
- The `1% / 90°` preset actually runs at 0.01%, so that the shadow narrows to a point.
- Turning Simplify on opens Illustrator's Simplify dialog; that is an Illustrator limitation.
- The preview relies on Illustrator's undo. Mixing it with other operations can leave the history in an unexpected state.
- Original idea: [こじらせたクマー](https://note.com/nice_lotus120/n/nf406fb3ae2b4)

### Article

[【Illustrator】ロングシャドウをスクリプトで作成する｜DTP Transit 別館](https://note.com/dtp_tranist/n/n0be484dab7fc)

### Update History

- v1.2.1 (2026-09-20): Added tooltips to the controls. Fixed the preview (nothing appeared for a group selection, an opaque copy was left behind for a text selection, copies survived closing the dialog by its close box) and the slider, preset and arrow-key values not staying in sync
- v1.2 (2026-02-25): First release
