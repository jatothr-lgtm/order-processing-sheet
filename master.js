/* Orders Placed — the default master.
 *
 * Built from the `ops` and `ops nm1` tabs of the working file, with the
 * decisions taken when the rules were reviewed:
 *
 *  - Groups whose source is "Proj vs SO" take SO Qty from the Proj vs SO file,
 *    matched through `pvsMap` (the `ops nm1` tab). A Group listed there under
 *    both Item Group and item_parent counts a row once if it matches either.
 *    The stray "Roasted & Flavoured Makhana" written against Walnut, Mexican
 *    Peri Peri Snack Mix and Seed Mix in `ops` is ignored.
 *  - Every other Group sums Stock Qty in KGS from the fill rate. Filled columns
 *    must all match; a comma-separated cell matches any of its values.
 *  - Makhana Roasted covers all Roasted & Flavoured Makhana, and Cashewnut
 *    Roasted includes the 50 g packs. The size rows are reported on top, not
 *    carved out — Grand Total is a plain sum of the rows, as in the history.
 *  - "Cashewnut" is an Item Group value in the fill rate (NEW MIS has "Cashew"),
 *    so the cashew rows filter on Item Group rather than NEW MIS ITEM GROUP.
 *  - "Roasted & Flavoured Cashew" is a NEW MIS ITEM GROUP value, so the 50 g
 *    row filters on that column rather than Parent Item.
 *  - Groups with no rule output 0, as the history does.
 *
 * Edits made on the page are kept in that browser and can be exported as an
 * .xlsx. Change this file to change the default for everyone.
 */
(function (root) {
  "use strict";

  const PVS = "Proj vs SO";
  const FR = "Fill Rate";
  const RFM = "Roasted & Flavoured Makhana";

  const g = (group, source, rule = {}) => ({
    group,
    source,
    itemGroup: rule.itemGroup || "",
    newMis: rule.newMis || "",
    parentItem: rule.parentItem || "",
    itemName: rule.itemName || "",
  });

  const groups = [
    g("Seeds", FR, { newMis: "Other Seeds, Chia Seeds" }),
    g("Tropical Savoury Snack Mix", FR, { parentItem: "Tropical Savoury Mix" }),
    g("Berry Mix", PVS),
    g("Almonds", PVS),
    g("Apricot", FR, { newMis: "Apricots" }),
    g("Cashewnut Retail", FR, { itemGroup: "Cashewnut", parentItem: "Premium W320 Cashew, Popular W400 Cashew, Premium W240 cashew" }),
    g("Cashewnut Roasted", FR, { itemGroup: "Cashewnut", parentItem: "Roasted and Salted Cashew, Roasted and Salted Cashew - Airlines, Roasted Thai Chilli Cashew" }),
    g("Dates", PVS),
    g("Desserts", FR, { itemGroup: "Desserts" }),
    g("Dry Fruit Combo", PVS),
    g("Dried Fruits", FR, { itemGroup: "Dried Fruits" }),
    g("Fig", FR, { newMis: "Figs" }),
    g("Gifting", PVS),
    g("Nut Mix", FR, { parentItem: "Protein Nut Mix - Airlines, Nut Mix - Airlines" }),
    g("Makhana Retail", FR, { newMis: "Makhana Raw" }),
    g("Makhana Roasted", FR, { newMis: RFM }),
    g("Makhana Roasted 16 g", FR, { newMis: RFM, itemName: "16 g" }),
    g("Makhana Roasted 51 g", FR, { newMis: RFM, itemName: "51 g" }),
    g("Munchies", PVS),
    g("Premium Panchmeva", FR, { parentItem: "Premium Panchmewa Superfood, Premium Panchmewa Superfood - Airlines" }),
    g("Premium Nuts", FR),
    g("Raisins", PVS),
    g("Satva Mix", PVS),
    g("Trail mix", PVS),
    g("Makhana Roasted 20 g", FR, { newMis: RFM, itemName: "20 g" }),
    g("Makhana Roasted 55 g", FR, { newMis: RFM, itemName: "55 g" }),
    g("Pistachio", PVS),
    g("Cashewnut Roasted 50 g", FR, { newMis: "Roasted & Flavoured Cashew", itemName: "50 g" }),
    g("Dry Fruit Mix Combo 18 g", FR),
    g("Oudh Paan Snack Mix", FR),
    g("Walnut", PVS),
    g("Mexican Peri Peri Snack Mix", PVS),
    g("Seed Mix", PVS),
    g("Makhana Pasta", FR),
    g("Makhana Roasted 14 g", FR, { newMis: RFM, itemName: "14 g" }),
    g("Makhana Roasted 57 g", FR, { newMis: RFM, itemName: "57 g" }),
  ];

  const m = (column, value, group = value) => ({ column, value, group });

  const pvsMap = [
    m("Item Group", "Berry Mix"),
    m("Item Group", "Almonds"),
    m("Item Group", "Dates"),
    m("Item Group", "Munchies"),
    m("Item Group", "Raisins"),
    m("Item Group", "Pistachio"),
    m("Item Group", "Walnut"),
    m("Item Group", "Seed Mix"),
    m("item_parent", "Berry Mix"),
    m("item_parent", "Dry Fruit Combo"),
    m("item_parent", "Gifting"),
    m("item_parent", "Satva Mix"),
    m("item_parent", "Trail mix"),
    m("item_parent", "Mexican Peri Peri Snack Mix"),
    m("item_parent", "Seed Mix"),
  ];

  const DEFAULT_MASTER = { groups, pvsMap };

  if (typeof module === "object" && module.exports) module.exports = DEFAULT_MASTER;
  else root.ORDERS_DEFAULT_MASTER = DEFAULT_MASTER;
})(typeof self !== "undefined" ? self : this);
