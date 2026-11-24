# Capacity Planning and Automated Par Level Model

[Türkçe](README.md) | **English**

An Excel model that calculates **per-item order quantities in cases** from store product sales, checks **whether the resulting stock fits dry storage and freezer capacity**, and determines **minimum shelf/freezer unit requirements** for new store planning.

> **Confidentiality & Synthetic Data Note:** This workbook is a **demonstration version rebuilt entirely with synthetic data** of the model I built for a 16-store area and rolled out across 6 areas. All brand identifiers, store names, and company data have been omitted; recipes, sales figures, shelf dimensions, and delivery schedules are synthetic.

## What problem it solves

A store's order depends on three variables: product sales velocity, how much of each stock item a product consumes, and how many days remain until the next delivery. This model links all three in one formula chain:

- Sales arrive by **menu product** (latte, croissant…), while orders are placed by **stock SKU** (cups, lids, coffee beans, frozen pastry…). A recipe matrix handles the conversion automatically.
- Order quantities scale by the delivery cycle days and delivery shift (`SABAH` morning vs. `AKŞAM` afternoon).
- Total case volume (`m³`) at par level is benchmarked against dry-storage shelves and freezer doors, outputting recommended minimum unit counts for store capacity sizing.

## Before

Every store calculated its order item by item by hand. After deployment across 16 stores in the area, the model was rolled out to 6 areas through regional leadership.

## How it works

```
Satis_Raporu ──────► Receteler ──────► Par_Level ──────► Kapasite & Kapasite_Ozeti
 (last 30-day /       (product sales    (daily usage ×    (case volume ÷ shelf &
  7-day sales)         × recipe =        period × 115%     freezer capacity + min.
                       daily usage)      − count = cases)  required units)
       │
       └───────────► Bolge_Ozet (sales mix across 5 store formats)
```

| Sheet | Purpose |
|---|---|
| `Nasıl Kullanılır` | Workflow, colour coding, and period-day logic |
| `Ayarlar` | Store selector, order/delivery dates, and report basis (`Aylık (30 gün)` monthly default / `Haftalık (7 gün)` weekly) |
| `Par_Level` | Per-item daily usage, par level, physical count, and order in cases |
| `Kapasite` | Dry storage and freezer utilization + minimum recommended unit counts for the selected store |
| `Kapasite_Ozeti` | Side-by-side capacity and unit sizing benchmark across 5 store formats (Mall, High Street, Campus, Airport, Highway) |
| `Magazalar` | Store master data: format, route, delivery shift, shelf and freezer door counts |
| `Stok_Kalemleri` | Units per case, case volume (`m³`), storage zone (`Depo` / `Donuk`) |
| `Sevkiyat_Takvimi` | Weekly delivery schedule per store |
| `Satis_Raporu` | Paste sheet holding both last 30-day and last 7-day POS/ERP product sales |
| `Receteler` | Product × stock item consumption matrix and daily usage totals |
| `Bolge_Ozet` | Category sales mix across all stores |

### Core formulas

- **Daily sales** (`Receteler`, col C): divides either **last 30-day sales** (default in `Ayarlar!B9`, used for regular ordering and new-store capacity planning) or **last 7-day sales** (used for short holiday/campaign peaks) by the selected day count.
- **Daily usage** (`Receteler`, total row): `SUMPRODUCT(daily sales; unit recipe consumption)`
- **Period days** (`Ayarlar`): `next delivery − order date + shift adjustment` (`+1` day for afternoon deliveries because that day's sales happen before the truck arrives).
- **Par level**: `daily usage × period × (1 + safety stock)` (`15%` default safety stock)
- **Order (cases)**: `ROUNDUP(MAX(0; par − physical count) ÷ units per case; 0)`
- **Utilization & unit sizing**: `par-level case volume ÷ (unit count × unit volume)`

## Setup and usage

Macro-free; opens directly in Excel 2010+ and LibreOffice.

1. Open `par-level-modeli.xlsx`.
2. Select the store, order/delivery dates, and report period (`Aylık (30 gün)` or `Haftalık (7 gün)`) in the yellow cells on `Ayarlar`.
3. Enter current physical counts in **Sayım (eldeki)** on `Par_Level` to read **Sipariş (koli)**.
4. Check `Kapasite` and `Kapasite_Ozeti` for storage utilization and unit sizing.

Colour coding: **blue** text = static input, **yellow** fill = active user input, **green** text = cross-sheet link, **black** text = formula.

## License

MIT. See `LICENSE` for details.
