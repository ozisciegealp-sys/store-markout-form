# Store End-of-Day (Mark-Out) and Thaw Tracking Form

[Türkçe](README.md) | **English**

An Excel + VBA workbook that manages end-of-day frozen and daily food stock counts, sales reconciliation, thaw scheduling, and next-day carryover in a coffee store.

> **Note:** This repository is a **working version rebuilt with synthetic data** of the end-of-day form I developed from the standard store sheet and shared across 16 stores in the area at a coffee chain. Each store adapted the form to its own product range. The structure of the original file (columns, calculations, shelf-life/expiry sheet, printable blank form) is preserved. Products, SKUs, and sales figures are synthetic.
>
> In the original system, the daily sales report was exported from the company's ERP system (Oracle) and imported into this form. In this version, `ornek-satis-raporu.xlsx` stands in for the exported report; the macro selects and reads the file in the exact same way.

## What problem it solves

At closing time each evening, every item requires answers to four questions: How much is left in frozen storage? How much was sold or wasted during the day? Does the physical count match the expected closing stock? Which items go on the shelf tomorrow and with what expiry date?

In this form:

- Sales quantities are pulled from the report by SKU with a single shortcut.
- Expected closing stock and variance (shrinkage/loss) are calculated automatically.
- Closing the day rolls today's closing count into tomorrow's opening stock, resets daily entries, and archives the day's record.
- Expiry dates (SKT) for items pulled to thaw are calculated on a dedicated sheet.

## Before

Sales quantities were looked up item by item from the report and entered by hand, closing variances were calculated manually, and next-day opening stocks were carried over manually.

## Files

| File | Contents |
|---|---|
| `markout-formu.xlsx` | Workbook (macro-free) containing formulas and sample data |
| `markout.bas` | VBA module to be imported into Excel |
| `ornek-satis-raporu.xlsx` | Synthetic daily sales report for testing the macro |

`markout.bas` avoids non-ASCII Turkish characters so that the file renders properly both on GitHub and after importing into the Windows ANSI-based VBA editor.

## Sheets

| Sheet | Purpose |
|---|---|
| `Nasil_Kullanilir` | Daily workflow and calculation notes |
| `Markout` | Main form: frozen and daily sections, expected closing stock, and variance |
| `SKT` | Thaw times and expiry dates for items pulled today and carried over from yesterday |
| `Bos_Form` | Printable blank count sheet for manual closing count |
| `Satis_Ortalama` | Average daily sales from weekly and monthly sales |
| `Ayarlar` | Store name, averaging basis, thaw hours, and sales report layout |
| `Arsiv` | Archived rows for every closed day |

## Calculations

- **Expected frozen closing** = opening frozen + received − pulled to thaw
- **Expected daily closing** = opening daily + thawed from previous day − sales − sampling − waste
- **Variance** = counted closing − expected closing (negative value indicates loss)
- **Thawed from previous day**: quantity pulled to thaw the prior day (thaw starts at 16:00, finishes at 08:00 the next morning)
- **Expiry date (SKT)** = thaw completion date + shelf life − 1 (removed from sale at mark-out time, 22:00)
- **Average daily sales** = last 7 days ÷ 7 or last 30 days ÷ 30 (weekly for short peaks such as holidays, monthly for regular planning; thaw quantity remains a manager decision assisted by this column)

## Macros

| Shortcut | Macro | Action |
|---|---|---|
| Ctrl+Shift+S | `SatislariAl` | Prompts to select the sales report file, reads SKU and quantity columns, sums duplicate SKUs, writes them to the form, and lists any SKUs present in the report but missing from the form. |
| Ctrl+Shift+P | `GunSonuPDF` | Saves the form as `Markout_<store>_<date>.pdf` in the workbook's folder. |
| Ctrl+Shift+K | `GunuKapat` | Warns if any counts are blank, archives the day's rows, rolls counted closing stocks into tomorrow's opening stock and thawed items into tomorrow's thawed column, clears daily inputs, and advances the date by one day. |

Macros never overwrite formula columns; all calculations remain transparent on the sheet.

## Setup

1. Download `markout-formu.xlsx` and `markout.bas`.
2. Open the workbook and press **Alt+F11** to open the VBA editor.
3. Choose **File > Import File…** and select `markout.bas`.
4. Close the editor and save the file as **Excel Macro-Enabled Workbook (\*.xlsm)**.
5. To activate shortcuts automatically on open, double-click **ThisWorkbook** in the VBA Project Explorer and paste:

   ```vb
   Private Sub Workbook_Open()
       KisayollariSessizAta
   End Sub
   ```

   Save, close, and reopen the workbook. (If skipped, run **Alt+F8** > `KisayollariAta` on open, or run macros directly via Alt+F8.)

## Quick Test

1. On the `Markout` sheet, press **Ctrl+Shift+S** and select `ornek-satis-raporu.xlsx`. The sales column populates and alerts that SKU `109999` is not in the form.
2. Enter a few frozen and daily "Sayılan Kapanış" (Counted Closing) values; variance columns update automatically.
3. Press **Ctrl+Shift+K** to close the day. Rows append to `Arsiv`, the date advances by one day, and counted closing values carry over as opening stock.

## Adapting to Your Data

| Location | What to adjust |
|---|---|
| `Markout` | Product rows: category header (blank SKU), SKU, product name, shelf life (days) |
| `Satis_Ortalama` | Weekly and monthly sales for the same SKUs |
| `Ayarlar` B8–B10 | SKU column, quantity column, and starting data row in the sales report |
| `Ayarlar` B5–B7 | Thaw and mark-out hours |
| `markout.bas` config block | Column numbers if sheet layout changes |

When adding a new product row, copy the corresponding formulas to the `SKT` and `Bos_Form` sheets as well.

## Limitations

- The sales report is exported manually from the ERP; the macro reads the exported file rather than downloading it.
- `GunuKapat` (Close Day) cannot be undone; exporting a PDF beforehand is recommended.
- Written for Excel 2010+ on Windows; not tested on macOS or LibreOffice.

## License

MIT. See `LICENSE` for details.
