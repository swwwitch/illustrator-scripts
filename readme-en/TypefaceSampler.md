# List fonts in weight order by category and instantly create font samples

[![Direct](https://img.shields.io/badge/Direct%20Link-TypefaceSampler.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/fonts/TypefaceSampler.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/TypefaceSampler.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Groups the fonts available in Illustrator by family (`font.family`), orders each group by weight and style, and lays them out on the artboard. Keyword, weight and style-category filters let you build a specimen sheet for just the range you need.

<img alt="" src="https://www.dtp-transit.jp/images/ss-776-1010-72-20250713-082053.png" width="70%" />

### Main Features

- Grouping by family, with each group sorted by a weight/style score
- Keyword search with AND / OR / NOT / prefix matching
- Checkbox filters for weight (5 ranks) and style category (5 groups)
- Selectable output: font name + weight/style, PostScript name, alphabet sample, numeral sample, or custom text
- Japanese and English UI

### Usage

1. Run the script with a document open.
2. Set the keyword, the output content and the filters in the dialog.
3. Press OK — the samples are drawn from the top-left of the active artboard.

Leaving the keyword blank targets every installed font, so a confirmation dialog appears first. Depending on how many fonts are installed, this can take a long time.

### Options

**Keyword**

The search covers `font.name`, `font.family` and `font.style`.

| Syntax | Meaning | Example |
| --- | --- | --- |
| space, `,` | OR (matches any) | `Helvetica Futura` |
| `+` | AND (matches all) | `DIN+Bold` |
| `^` | Prefix match | `^DIN` |
| `-` | Exclude (NOT) | `Helvetica -Now` |

Full-width spaces and commas are handled as well, and the forms combine (for example `^DIN+Bold -Condensed` means: starts with DIN, contains Bold, does not contain Condensed).

**Output content**

| Item | What is drawn |
| --- | --- |
| Font Name + Weight/Style | `font.family` and `font.style` |
| PostScript Name | `font.name` |
| The quick brown fox… | Alphabet sample |
| 1234567890 | Numeral sample |
| Custom | The text in the input field |

**Display options**

| Item | Description |
| --- | --- |
| Weight Count | Appends the number of fonts to the category name |
| Weight List | When off, draws only the category names in a single column, set in that category's lightest weight |
| Columns | Number of columns of categories (arrow keys step the value, shift steps by 10) |
| Debug Score | Appends the sort score (only for the Font Name + Weight/Style output) |

**Weight**

Classifies `font.style` into five ranks. Numeric styles such as `W3` / `W600` and leading-number styles such as `25 Ultra Light` are supported.

| Rank | Example styles |
| --- | --- |
| Hairline / Thin | Hairline, Ultra Thin, Thin |
| Light | Ultra Light, Extra Light, Light |
| Regular | Book, Normal, Regular, Roman |
| Medium / SemiBold | Medium, SemiBold, DemiBold |
| Bold / Black | Bold, Extra Bold, Heavy, Black, Ultra |

**Style**

Classifies by the decoration words found in `font.style`. A font can fall into more than one category.

| Category | Matching words |
| --- | --- |
| Basic | Text, Headline |
| Condensed | Cond, Condensed, Compressed, Comp |
| Expanded | Expanded, Extended |
| Display / Special | Compact, Display |
| Size / Proportion | Micro, Low, Wide |

If nothing is checked in a group, that filter is ignored. When both groups have selections, they combine with AND — only fonts satisfying both are kept.

### Notes

- Running with a blank keyword can cover several thousand faces and take a long time.
- Basic only matches fonts whose `font.style` contains Text or Headline; ordinary Regular and Bold cuts do not qualify.
- Fonts that cannot be applied are skipped and logged to the ExtendScript console.

### Article

- [Listing fonts by category in weight order to build a specimen sheet in one step (Japanese)](https://note.com/dtp_tranist/n/n103ac6622657)

### Update History

- v1.0.0 (20250420): Initial version
- v1.3.0 (20250508): Improved evaluation logic and filter features
- v1.3.1 (20250706): Localization adjustments
- v1.3.2 (20260902): Added weight and style-category filters (merged TypefaceSampler-text.jsx)
