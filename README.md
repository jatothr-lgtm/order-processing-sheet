# Order Processing Sheet

**How much was ordered per Group this month?**

A static page that takes the **Proj vs SO** file and the month's **fill rate
report** and returns the *Orders Placed* sheet: one row per Group, one column
per month, in kg, with Grand Total Kgs and Grand Total Tons underneath.

Everything runs in the browser. No file is uploaded, there is no backend and
nothing is stored anywhere except master edits, which stay in the browser that
made them.

## Inputs

| Slot | Required columns | Notes |
|---|---|---|
| **1 · Proj vs SO** | `Month`, `Year`, `Item Group`, `item_parent`, `SO Qty` | Any number of months. Only the fill rate's month is used. |
| **2 · Fill rate** | `Sales Order Date`, `Stock Qty in KGS`, `Item Group`, `NEW MIS ITEM GROUP`, `Parent Item`, `Item Name` | The month comes from `Sales Order Date`. |
| **3 · Previous output** *(optional)* | An Orders Placed sheet: `Group` in column A, months across that row | Its history is kept and the new month is added (or replaces that month if already there). |

`.xlsx`, `.xlsm`, `.xls` and `.csv` are accepted. The header row is found
anywhere in the first 15 rows, column names are matched ignoring case, spaces
and underscores, and a workbook that holds both the Proj vs SO and fill rate
tabs fills both slots at once. When several sheets match, the largest is taken
and the others are offered in a dropdown.

Dates are read from the raw Excel serial in UTC. Text dates may be
`2026-09-30`, `30-09-2026`, `30/09/2026` (day first), `30-Sep-2026` or `Sep-26`.
A fill-rate line with kg but no readable `Sales Order Date` cannot be placed in
a month; those are left out and reported with their kg.

## The calculation

The month is the fill rate's month. Proj vs SO is read for that same month.

**Proj vs SO Groups** sum `SO Qty` over the Proj vs SO rows for that month whose
`Item Group` or `item_parent` matches the Group's mapping rows. A row that
matches more than one mapping row of the same Group counts once.

**Fill Rate Groups** sum `Stock Qty in KGS` over the fill-rate lines that match
**every** filled filter column (`Item Group`, `NEW MIS ITEM GROUP`, `Parent Item`,
`Item Name contains`). A cell listing several values separated by commas matches
**any** of them. Text is compared ignoring case and extra spaces, so
`Trail mix` matches `Trail Mix` and `Satva Mix ` matches `Satva Mix`.

**Item Name contains:** a number (`16` or `16 g`) means a pack size: that number
followed by g / gm / gms / gram, and not part of a bigger number. So
`Pillow Pouch 16 g` and `16g-Ladi` count, while `Pouch 160 g` and `1.16 g` do
not. Any other text is a plain "contains".

**Grand Total Kgs** is a plain `SUM` of every Group row and **Grand Total Tons**
is `ROUND(kg / 1000, 0)`, as in the history. The size rows sit inside their parent
row and are counted again in the total. For example, Makhana Roasted 20 g is
also in Makhana Roasted.

Values are kept unrounded (to the gram) and displayed as whole kg, so the total
is exact. Months carried over from a previous output keep the totals they had
there. A few of those were taken before rounding and differ from the sum of
their rows by a kg or two.

### The default master

Built from the `ops` and `ops nm1` tabs of the working file, with these
decisions:

| Group | Source | Rule |
|---|---|---|
| Berry Mix, Seed Mix | Proj vs SO | `Item Group` **or** `item_parent` = the Group (listed under both in `ops nm1`) |
| Almonds, Dates, Munchies, Raisins, Pistachio, Walnut | Proj vs SO | `Item Group` = the Group |
| Dry Fruit Combo, Gifting, Satva Mix, Trail mix, Mexican Peri Peri Snack Mix | Proj vs SO | `item_parent` = the Group |
| Seeds | Fill Rate | NEW MIS ITEM GROUP in Other Seeds, Chia Seeds |
| Tropical Savoury Snack Mix | Fill Rate | Parent Item = Tropical Savoury Mix |
| Apricot | Fill Rate | NEW MIS ITEM GROUP = Apricots |
| Fig | Fill Rate | NEW MIS ITEM GROUP = Figs |
| Cashewnut Retail | Fill Rate | Item Group = Cashewnut and Parent Item in Premium W320 / Popular W400 / Premium W240 |
| Cashewnut Roasted | Fill Rate | Item Group = Cashewnut and Parent Item in Roasted and Salted Cashew (+ Airlines), Roasted Thai Chilli Cashew — **includes** the 50 g packs |
| Cashewnut Roasted 50 g | Fill Rate | NEW MIS ITEM GROUP = Roasted & Flavoured Cashew and Item Name has 50 g |
| Desserts | Fill Rate | Item Group = Desserts |
| Dried Fruits | Fill Rate | Item Group = Dried Fruits (overall, separate from Apricot and Fig) |
| Nut Mix | Fill Rate | Parent Item in Protein Nut Mix - Airlines, Nut Mix - Airlines |
| Premium Panchmeva | Fill Rate | Parent Item in Premium Panchmewa Superfood (+ Airlines) |
| Makhana Retail | Fill Rate | NEW MIS ITEM GROUP = Makhana Raw |
| Makhana Roasted | Fill Rate | NEW MIS ITEM GROUP = Roasted & Flavoured Makhana — **all** sizes |
| Makhana Roasted 14 / 16 / 20 / 51 / 55 / 57 g | Fill Rate | Roasted & Flavoured Makhana and Item Name has that size |
| Premium Nuts, Dry Fruit Mix Combo 18 g, Oudh Paan Snack Mix, Makhana Pasta | Fill Rate | no rule, so 0 |

Fill-rate lines are included whatever their Workflow State or SO Status.

Two corrections to `ops` as written:

- `Cashewnut` is an *Item Group* value in the fill rate. NEW MIS ITEM GROUP says `Cashew`.
- `Roasted & Flavoured Cashew` is a *NEW MIS ITEM GROUP* value, not a Parent Item.

The "Roasted & Flavoured Makhana" written against Walnut, Mexican Peri Peri
Snack Mix and Seed Mix in `ops` is ignored: those Groups come from Proj vs SO.

## The master is editable

The **Master** tab has two tables: *Groups* (output rows, their order, source
and fill-rate filters) and *Proj vs SO mapping*. Edits apply at once and are
kept in that browser. **Download master** saves them as an `.xlsx` and **Upload
master** loads one back. Every output file also carries the master it was built
with, so an output can be uploaded as a master too. **Reset to default** returns
to [`master.js`](master.js); change that file to change the default for everyone.

Mistakes are flagged rather than turned into a silent 0:

- a filter value that appears nowhere in the uploaded file
- a Proj vs SO Group with no mapping
- a mapping that points at a missing Group
- a Group listed twice

## The output workbook

| Sheet | What it holds |
|---|---|
| **Orders Placed** | The Sheet1 layout: title, year labels, `Group` with `mmm-yy` months, one row per Group, Grand Total Kgs (`SUM`) and Tons (`ROUND`) |
| **Workings** | Each Group's rule, the number of source rows it matched and the unrounded kg |
| **Notes** | The files and sheets used, and every check raised on screen |
| **Master - Groups**, **Master - Proj vs SO** | The master used for this output |

## Files

| File | Role |
|---|---|
| `index.html` | The page |
| `calc.js` | Reading the inputs and the calculation. No DOM, so the same file is checked in Node |
| `master.js` | The default master |
| `reader.js` | Web Worker that parses workbooks off the main thread (≈45,000 fill-rate lines takes a few seconds). If it cannot start, the page parses on the main thread with the same code |
| `app.js` | Slots, master editor, preview, Excel export |
| `styles.css` | Light and dark |

[SheetJS](https://sheetjs.com) 0.18.5 reads and ExcelJS 4.4.0 writes, both from
cdnjs.

## Deploying

Static: no build, no dependencies, no environment variables.

On Vercel, import this repository with Framework Preset **Other**, no build
command, no output directory and the repository root as the root.
`vercel.json` sets the cache and security headers. Any static host works the
same way: serve the files as they are.

`styles.css`, `calc.js` and the worker load by relative path. If the app is
served under a sub-path and opened without its trailing slash, the page adds the
slash so those files still resolve.

Locally, serve the folder (for example `python3 -m http.server`) and open it.
The worker does not start from `file://`, but the page still works there,
parsing on the main thread.
