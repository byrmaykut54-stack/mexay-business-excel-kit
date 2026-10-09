# MexAy Business Excel Kit

Offline workbook for appointment, customer and actual income/expense tracking. The delivered workbook starts empty. Each input sheet supports 500 records (rows 2-501); extending capacity requires extending the tables, dashboard formulas and validations together. Appointment fees are not automatically booked as revenue. Customer visit/spend totals are manual. All financial values are TRY. Business hours are informational.

## Build

Run in the Codex primary runtime with `CODEX_PRIMARY_RUNTIME_NODE` and `CODEX_PRIMARY_RUNTIME_NODE_MODULES` configured and Python dependencies from requirements.txt available:

```bash
python generator/build_product.py
```

The workbook is authored with @oai/artifact-tool. The builder checks recalculation at the first, middle and last supported rows, restores the empty state, and renders every sheet to `qa/`. PDF generation requires DejaVu Sans fonts (default `/usr/share/fonts/truetype/dejavu`; override with `MEXAY_FONT_DIR`). Final distribution files are in `output/`; the ZIP includes exactly the workbook, PDF guide, three sales visuals and license. QA files are excluded.

Review all rendered sheets/PDF pages and reopen the workbook in Microsoft Excel before changing its supported-engine claim or uploading a new sale package. Do not automatically publish a package during a build.
