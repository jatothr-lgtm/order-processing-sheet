/* Orders Placed — reads a workbook off the page's main thread.
 * A month's fill rate is ~45,000 lines and takes several seconds to unzip and
 * parse; doing it here keeps the page responsive while it does. If this
 * worker cannot start, app.js reads on the main thread with the same code. */
importScripts("https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js", "calc.js");

self.onmessage = e => {
  const { id, buffer } = e.data;
  try {
    self.postMessage({ id, ok: true, result: self.OrdersCalc.readWorkbook(self.XLSX, buffer) });
  } catch (err) {
    self.postMessage({ id, ok: false, error: String((err && err.message) || err) });
  }
};
