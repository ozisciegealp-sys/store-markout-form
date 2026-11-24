# Store End-of-Day (Mark-Out), Thaw Planning and Waste Analysis System

[Türkçe](README.md) | **English**

An Excel + VBA workbook for multi-store retail and F&B operations that automates end-of-day frozen/daily stock counts, ERP (Oracle) SKU sales reconciliation, transaction-based (ADT/UPH) thaw planning, shelf-life/expiry (`SKT`) tracking, and monthly waste classification.

> **Confidentiality & Synthetic Data Note:** This repository presents a **working version rebuilt entirely with synthetic data** of the store-level reporting and inventory control workbook deployed across a 16-store area. All brand names, store identifiers, and company data have been removed; product names, SKUs, unit costs, and sales figures are synthetic.
>
> In live operations, the daily sales report was exported from the ERP system (Oracle) and read into this workbook with a single shortcut. Here, `ornek-satis-raporu.xlsx` represents that exported ERP report.

## What problem it solves

At closing time each evening—and when planning the next day's pull—a store team needs fast, error-free answers: How much stock remains in the freezer and display case? How much was sold vs. wasted? Does the physical count match expected closing stock? Based on tomorrow's forecast customer transactions, how many units should be pulled to thaw and with what expiry date?

In this workbook:

- Sales quantities are pulled by SKU from the exported ERP report with a single shortcut (`Ctrl+Shift+S`).
- Expected frozen/daily closing stocks and count variances (shrinkage) calculate automatically.
- Tomorrow's forecast transactions (`ADT`) and per-SKU units per hundred transactions (`UPH`) drive suggested pull-to-thaw quantities and freezer stock alerts.
- Expiry dates (`SKT`) for newly pulled and carried-over items compute automatically.
- 30-day sales and waste figures classify items into **Star Items** (high sales, low waste), **Balanced**, and **Problem Items** (above target waste — reduce pull).
- Closing the day (`Ctrl+Shift+K`) archives the day's rows and rolls closing counts into tomorrow's opening stock.

## Before

Sales quantities were looked up item by item from ERP exports and typed by hand, variances and thaw pulls were calculated manually, and next-day opening balances were copied over manually (~60 min/day → ~1 min/day per store).

## Files

| File | Contents |
|---|---|
| `markout-formu.xlsx` | Main workbook with all formulas, planning/analysis tabs, and synthetic sample data |
| `markout.bas` | VBA module handling ERP sales import, PDF export, and day-close rollover |
| `ornek-satis-raporu.xlsx` | Synthetic daily ERP SKU sales export + period KPI summary for testing |

`markout.bas` intentionally uses ASCII-only identifiers so that the module renders cleanly both on GitHub and inside the Windows ANSI-based VBA editor. All formula and analysis sheets work out-of-the-box without importing macros.

## Sheets

| Sheet | Purpose |
|---|---|
| `Nasil_Kullanilir` | Daily workflow, formula rules, and colour coding |
| `Markout` | Main sheet: frozen and daily stock flows, expected closing stock, and variance |
| `SKT` | Thaw timestamps and expiry dates (`SKT`) for items pulled today and carried over |
| `Cozundurme_Plani` | Transaction (`ADT`) and `UPH`-driven daily pull-to-thaw recommendations and freezer stock alerts |
| `Ayarlar` | Store name, averaging basis, thaw/mark-out hours, waste threshold, and ERP column mapping |
| `Satis_Ortalama` | Average daily sales derived from 7-day and 30-day SKU sales |
| `Fire_Analizi` | 30-day waste rate, waste cost, and Star / Balanced / Problem item classification |
| `Bos_Form` | Printable blank sheet for floor closing counts |
| `Arsiv` | Historical log of closed days |

## Core Calculations

- **Expected frozen closing** = opening frozen + received − pulled to thaw
- **Expected daily closing** = opening daily + thawed from prior day − sales − sampling − waste
- **Variance** = counted closing − expected closing (negative = loss)
- **UPH (Units per 100 transactions)** = average daily sales ÷ reference daily transactions × 100
- **Suggested thaw quantity** = `MAX(0, ROUNDUP(forecast demand − expected daily closing, 0))`
- **Expiry date (SKT)** = thaw completion date + shelf life − 1 (removed at mark-out hour)
- **Waste rate %** = 30-day waste ÷ (30-day sales + 30-day waste)

## Macros

| Shortcut | Macro | Action |
|---|---|---|
| `Ctrl+Shift+S` | `SatislariAl` | Prompts for the ERP sales export, aggregates quantities by SKU, populates `Markout`, and flags any unmatched SKUs in the report. |
| `Ctrl+Shift+P` | `GunSonuPDF` | Exports the form as `Markout_<store>_<date>.pdf` in the workbook folder. |
| `Ctrl+Shift+K` | `GunuKapat` | Warns on blank counts, appends rows to `Arsiv`, rolls closing counts into tomorrow's opening columns, clears daily inputs, and advances the date. |

## Enabling Macros & Quick Test

1. Download `markout-formu.xlsx`, `markout.bas`, and `ornek-satis-raporu.xlsx` into the same folder.
2. All formula sheets (`Markout`, `SKT`, `Cozundurme_Plani`, `Fire_Analizi`) can be explored immediately in `.xlsx`.
3. To test the VBA automation in Windows Excel, press **Alt+F11** > **File > Import File…**, select `markout.bas`, save as `.xlsm`, and run `KisayollariAta` once via **Alt+F8**.
4. Press **Ctrl+Shift+S** on `Markout` and select `ornek-satis-raporu.xlsx` — sales populate automatically and SKU `109999` triggers the unmatched-SKU alert.

## License

MIT. See `LICENSE` for details.
