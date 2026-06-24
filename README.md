# Auto Rake Tally

Consolidate messy, multi-sheet steel **rake/wagon** Excel workbooks into one clean,
standard 24-column CSV — one row per physical item (plate / coil / bar), with the
11-digit rail wagon number resolved deterministically and missing fields left blank.

Built for the four common workbook shapes (CR, HR coil, Plate, TMT), each of which
names its columns differently and hides the wagon number in a different place.

## Features

- **One header schema, many source layouts.** A per-destination synonym table maps
  varied headers (`THK` / `THICK` / `Thickness`) onto the fixed 24 columns without
  collapsing distinct fields together.
- **Deterministic wagon resolution — no guessing.** The 11-digit number is taken from
  the cell (handling a `JSPL…` prefix), back/forward-filled across a wagon group, or
  joined from a companion sheet by item id (HR coil). If it genuinely isn't present,
  the cell is left blank rather than fuzzy-matched to a wrong wagon.
- **Clean output.** Integer ids lose their `.0` tail, scientific notation is expanded,
  footer/subtotal/trailer rows are dropped, and exact duplicate rows are removed.
- **Modern, dependency-free GUI.** A ttk desktop app — no third-party UI library.

## Install

```bash
pip install -r requirements.txt
```

Requires Python 3.10+.

## Usage

**GUI** — pick the workbook, pick an output path, click *Convert*:

```bash
python gui.py        # or: run_gui.bat (Windows) / ./run_gui.sh (macOS/Linux)
```

**Command line / import:**

```python
from wagon_converter import SCHEMA, run_conversion

stats = run_conversion("RAKE.xlsx", SCHEMA, "output_cleaned.csv")
print(stats)  # {'primary_sheet': ..., 'total': ..., 'matched': ..., 'failed': ...}
```

Running `python wagon_converter.py` executes a built-in self-check.

## Project structure

```
auto-rake-tally/
├── wagon_converter.py   # engine: read → map → resolve wagon → clean → CSV
├── gui.py               # ttk desktop front end
├── run_gui.bat          # Windows launcher
├── run_gui.sh           # macOS/Linux launcher
├── requirements.txt
├── skills/              # AI-assist skills used while building this
│   ├── Ponytail.md      # "laziest solution that works" coding discipline
│   └── skills-lock.json # pinned: pandas-pro
├── LICENSE
└── README.md
```

## How wagon resolution works per material

| Material | Detail sheet | Wagon source |
|----------|-------------|--------------|
| CR       | richest sheet | embedded in the cell (`JSPL` prefix stripped) |
| HR coil  | richest sheet | joined from the annexure by coil id |
| Plate    | richest sheet | not present in the workbook → left blank |
| TMT      | richest sheet | trailer row at the bottom of each group → back-filled |

## Assumptions / limitations

- Output grain is **one row per item** (plate/coil/bar).
- Wagon groups are assumed **contiguous**; the wagon marks the group edge.
- Tables grow row-wise; processing is row-by-row (fine well past a few thousand rows).

## Skills used

This project was built with two AI-assist skills (kept in `skills/`):

- **Ponytail** — a discipline favouring the simplest solution that works: stdlib over
  dependencies, deletion over addition, root-cause over symptom. It's why the fuzzy
  matcher was removed and the GUI uses only tkinter.
- **pandas-pro** — pinned pandas guidance (`skills-lock.json`).

## License

MIT — see [LICENSE](LICENSE).
