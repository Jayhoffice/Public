# Admin

## script_analyzer_files_category.py
A small always-on-top status window that scans filenames in the desktop (or a chosen folder) for the `primary-category_secondary-category_...` pattern and shows a live count per category. (Recommended to rename to `.pyw` before use.)
Double-clicking a row in the table opens Everything with that row's "primary_secondary" string as the search term.

## script_compressor_pdf.py
Select multiple PDF files and compress them using Ghostscript (choose quality: screen / ebook / print). Requires Ghostscript installed.

## script_converter_doc2pdf.py
Select multiple doc/docx files and convert them to PDF using Microsoft Word's ExportAsFixedFormat feature. Requires Windows + Word + pywin32.

## script_converter_hwp2md.py
Converts HWP/HWPX files into Markdown that's easier for AI to read (equations are converted to LaTeX). Can be run via file picker or with a file path as a command-line argument.

## script_converter_hwp2pdf.py
Select multiple HWP/HWPX files and convert them to PDF using Hangul (Hancom Office)'s built-in PDF export. Requires Windows + Hangul + pyhwpx.

## script_merger_pdf.py
Select multiple PDF files, arrange their order, and merge them into one PDF. Only needs pypdf, no OS restriction.

## script_merger_ppt.py
Select multiple PPT/PPTX files, arrange their order, and merge them into one file while preserving each slide's background/master design. Requires Windows + PowerPoint + pywin32.

## page_calculator_neis_byte.html
Local-only web page: checks character/byte count for text input (e.g. NEIS student records).

# physics_simulations

## page_sim_mech_oscillation.html
Damped harmonic oscillator simulation. Adjust mass (m), spring constant (k), and damping coefficient (c) and view the resulting graph.

## page_sim_thermo_collapsingbottle.html
Simulation of a water-cooler bottle collapsing, based on the ideal gas law.

---

**Note:** `page_mech_damped_oscillation.html` and `page_thermo_collapsingbottle.html` at the repo root are byte-identical duplicates of the files above in `physics_simulations/`. Left as-is; let me know if you'd like them removed.