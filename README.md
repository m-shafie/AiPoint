# AiPoint

Repair PowerPoint files exported from AI tools — fix Persian typography, watermarks, slide sizes, and more.

AI-generated PowerPoint presentations often have broken layout, wrong slide dimensions, missing Persian fonts, mismatched language tags, or leftover watermarks from the AI tool. AiPoint fixes all of these in bulk.

## Features

- **Slide Size Correction** — Rescale all shapes/Elements (EMU) to a standard 16:9 target (or custom) size
- **Persian Typography Fix** — Set Persian (cs) font for Persian/rune text, Latin font for English, fix parentheses sizing, assign language tags (fa-IR / en-US)
- **Font Size Override** — Force a global font size across all slides
- **Line Spacing Control** — Apply single spacing or a custom multiplier
- **Watermark Removal** — Remove AI-tool watermarks (strict pixel-box or heuristic position-based mode)
- **Multi-File Merge** — Append multiple decks into one, preserving images and charts
- **Two Merge Engines** — PowerPoint COM automation (Windows, preserves everything) or pure Python fallback (no Office required)
- **Graphical & Command-Line Interfaces**

## Modes

### GUI (Tkinter)

```bash
pip install -r requirements.txt
python fix_pptx_gui.py
```

Features a drag-to-reorder file queue, font picker, watermark toggle, and merge/file selection dialogs.

### CLI

```bash
pip install python-pptx
python fix_pptx_layout.py -i input.pptx -o output.pptx --persian-font "B Nazanin"
```

#### CLI Options

| Option | Description |
|--------|-------------|
| `-i, --input` | Input .pptx file |
| `-o, --output` | Output .pptx file |
| `--target-cx, --target-cy` | Target slide dimensions (default: 12192000 x 6858000 = 16:9) |
| `--persian-font` | Persian font name (default: "B Nazanin") |
| `--latin-font` | Latin font name (default: "Times New Roman") |
| `--font-size` | Force font size in points |
| `--line-spacing-single` | Apply single line spacing |
| `--line-spacing-multiple` | Apply custom line spacing multiplier |
| `--remove-watermark` | Enable watermark removal |
| `--parts` | Which parts to process (default: slides,layouts,masters,notes,diagrams,handouts) |

### Installed Entry Points

```bash
pip install -e .
aipoint       # launches the GUI
aipoint-cli   # launches the CLI parser
```

## Tech Stack

| Layer | Technology |
|-------|------------|
| Language | Python 3 |
| Libraries | python-pptx, pywin32 (Windows COM) |
| GUI | Tkinter |
| File Format | OpenXML (ZIP-based .pptx) |

## Requirements

- Python 3.8+
- `python-pptx` (required for both CLI and GUI)
- `pywin32` (Windows only — optional, enables PowerPoint COM merge)

---

## فارسی

**تعمیر فایل‌های پاورپوینت خروجی هوش مصنوعی** — رفع مشکلات تایپوگرافی فارسی، واترمارک و اندازه اسلایدها

فایل‌های پاورپوینتی که توسط ابزارهای هوش مصنوعی ساخته می‌شوند اغلب مشکل دارند: فونت فارسی ندارند، اندازه اسلایدها درست نیست، یا واترمارک ابزار روی آنها باقی مانده. این ابزار همه این مشکلات را رفع می‌کند.

### قابلیت‌ها
- **تنظیم فونت فارسی** — اعمال فونت فارسی (مثلاً ب نازنین) روی متن‌های فارسی و فونت لاتین روی متن‌های انگلیسی
- **رفع اندازه اسلایدها** — تغییر ابعاد به ۱۶:۹ استاندارد
- **حذف واترمارک** — حذف واترمارک ابزارهای هوش مصنوعی
- **تنظیم فاصله خطوط** — اعمال فاصله خطوط یکسان در کل فایل
- **رفع پرانتزها** — اصلاح سایز پرانتزها در متن فارسی
- **ادغام فایل‌ها** — ادغام چند فایل پاورپوینت با حفظ تصاویر
- **رابط گرافیکی و خط فرمان** — هم CLI و هم GUI (Tkinter)
