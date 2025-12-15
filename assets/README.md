# Cellify Brand Assets

This folder contains the official Cellify logo and branding assets.

## Logo Files

| File | Description | Size |
|------|-------------|------|
| `logo.svg` | Primary logo (emerald green) | 32x32 |
| `logo-large.svg` | Large logo for headers/social | 128x128 |
| `logo-dark.svg` | Dark mode variant (lighter green) | 32x32 |

## Brand Colors

| Color | Hex | Usage |
|-------|-----|-------|
| Primary (Emerald) | `#059669` | Main brand color, logo background |
| Primary Light | `#10b981` | Dark mode, hover states |
| Primary Dark | `#047857` | Active states, accents |
| White | `#ffffff` | Logo cells, text on primary |

## Usage

### In HTML
```html
<img src="assets/logo.svg" alt="Cellify" width="32" height="32">
```

### In Markdown
```markdown
![Cellify Logo](assets/logo.svg)
```

### Copying to other locations
The logo is stored here as the single source of truth. When needed in other locations (demo, docs), copy from here:

```bash
cp assets/logo.svg demo/logo.svg
cp assets/logo.svg docs-site/static/img/logo.svg
```

## Design Notes

- The logo represents a spreadsheet grid with 6 cells (2 columns x 3 rows)
- Uses rounded corners (rx=4 for outer, rx=1 for cells) for a modern look
- SVG format ensures crisp rendering at any size
- Viewbox is 32x32 but can be scaled to any dimension
