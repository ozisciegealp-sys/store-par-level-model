# Par Level and Capacity Planning Model

[Türkçe](README.md) | **English**

An Excel model that calculates **how much to order for every stock item** from a store's product sales report and checks **whether the resulting stock fits the stockroom and freezer capacity**.

> **Note:** This file is a **demonstration version rebuilt with synthetic data** of the model I built and used across multi-store regional operations at a coffee chain. The original workbook contained proprietary company data and is not shared. Store names, products, sales quantities, shelf dimensions, and delivery schedules are synthetic. The calculation logic and sheet chain match the original; formulas were rewritten to work in older Excel versions and LibreOffice as well.

## What problem it solves

A store's order quantity depends on three variables: how much each product sells, how much of each consumable/stock item a product uses, and how many days remain until the next delivery. This model connects all three in a single chain:

- Sales arrive by **product** (latte, croissant…), while orders are placed by **stock item** (large cup, lid, coffee beans…). A recipe matrix handles the conversion.
- Order quantities scale automatically by the number of days until the next delivery (period).
- Total cases and volume of the calculated stock are compared against the store's shelf and freezer capacity, planning inventory together with physical space.

## Before

Each store calculated and placed orders manually, item by item. After rollout, the model was used across 16 stores in the area, with stores transitioning to ordering directly from its output, and was later rolled out across six areas through the area manager.

## How it works

```
Satis_Raporu ──► Receteler ──► Par_Level ──► Kapasite
 (weekly         (product sales  (daily usage ×   (case volume ÷
  product sales)  × recipe =      period ×         shelf/freezer
                  daily usage)    115% − count     capacity)
                                  = order)
       │
       └──────► Bolge_Ozet (store sales mix)
```

| Sheet | Purpose |
|---|---|
| `Nasıl Kullanılır` | Workflow, colour coding, and period calculation guide |
| `Ayarlar` | Select store, order date, and next delivery date; calculates period days |
| `Par_Level` | Per-item par level, order quantity, and case count |
| `Kapasite` | Dry storage and freezer utilization ratio |
| `Magazalar` | Store master data: format, route, delivery shift, shelf and freezer door counts |
| `Stok_Kalemleri` | Units per case, case volume, storage area |
| `Sevkiyat_Takvimi` | Weekly delivery schedule per store |
| `Satis_Raporu` | Paste area for the weekly/monthly POS/ERP sales report |
| `Receteler` | Product × stock item usage matrix and daily usage totals |
| `Bolge_Ozet` | Main-group sales mix across all stores |

### Core formulas

- **Daily sales** (`Receteler`, col C): pulls the selected store's sales via `SUMIFS` and divides by the report period days.
  - **Monthly** sales report (`Ayarlar` B10 = 30) is used for baseline ordering and new-store capacity sizing.
  - **Weekly** sales report (`Ayarlar` B10 = 7) is used during short demand shifts such as holidays. The sample file ships with this setting.
- **Daily usage** (`Receteler`, total row): `SUMPRODUCT(daily sales; usage per unit)`
- **Period days** (`Ayarlar`): `next delivery − order date + shift adjustment`. Stores receiving afternoon deliveries get `+1` day because that day's sales occur before the truck arrives.
- **Par level**: `daily usage × period × (1 + safety stock)`, with 15% safety stock
- **Order**: `MAX(0; par − physical count)`, divided by case pack size and rounded up
- **Utilization**: `par-level case volume ÷ (unit count × unit volume)`

## Setup and usage

No macros or setup required. Works in Excel 2010+ and LibreOffice.

1. Download and open `par-level-modeli.xlsx`.
2. Fill the yellow cells on `Ayarlar`: store (from dropdown), order date, and next delivery date (refer to `Sevkiyat_Takvimi`).
3. Enter physical stock counts in the **Sayım** column on `Par_Level`.
4. Read the **Sipariş (koli)** column for cases to order, and check `Kapasite` to confirm physical fit.

To refresh for a new period, paste the new sales report into `Satis_Raporu` using the same column layout.

## Adapting to your data

| Location | What to enter |
|---|---|
| `Magazalar` | Store name, format, delivery shift (`SABAH` / `AKŞAM`), shelf and freezer door counts |
| `Stok_Kalemleri` | Item name, units per case, case volume (m³), storage area (`Depo` / `Donuk`) |
| `Receteler` | Rows: products; columns: stock items; cells: quantity consumed per unit sold |
| `Sevkiyat_Takvimi` | Weekly delivery days per store |
| `Ayarlar` B10 / B11 | Days covered by sales report (7 weekly, 30 monthly) and safety stock ratio |
| `Kapasite` B4 / B5 | Volume of one shelf unit and one freezer door |

When adding a new stock item, use the exact same name in `Stok_Kalemleri` and the `Receteler` header row (`Par_Level` matches by name), and expand the formula ranges on `Par_Level` and `Receteler`.

Colour coding: **blue** text = manual input, **yellow** fill = filled on every use, **green** text = cross-sheet reference, **black** text = formula.

## Limitations

- The next delivery date is selected manually rather than auto-resolved from the calendar.
- Demand reflects the daily average of the loaded report period; seasonal peaks require loading that period's weekly report (no standalone forecasting).
- Shelf and freezer volumes are sample figures and should be measured on site for live use.

## License

MIT. See `LICENSE` for details.
