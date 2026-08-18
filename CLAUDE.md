# Provisions London — Stock/Ops Tool

This is **provisionslondon**, a stock/products/orders tool for Provisions
UK. It is a separate project from `carisma-catering` — do not mix data
models, sheet schemas, or conventions between the two.

- **Repo:** `Plato-uk/provisionslondon` on GitHub. Branch `main` only.
- **Architecture:** static site, no server/build step. Every page talks
  directly to a single Google Sheet via the Sheets API v4, authenticated
  client-side with Google OAuth (`js/config.js`, `js/sheets.js`). Current
  OAuth scope is `spreadsheets` only — no Drive scope yet, so there's no
  file-upload/storage capability wired up today.
- Generic CRUD table framework in `js/crud-page.js` — every list page
  (Products, Suppliers, Customers, Deliveries, Orders, OrderLines,
  Allocations) configures one `initCrudTable()` call instead of hand-rolling
  fetch/append/update/delete logic. Read the comment block at the top of
  that file before adding a new page or field type.

## Sheet schema (see `setup.html`'s `REQUIRED_TABS` for the source of truth)
- **Products** — item master, one row per SKU keyed by `Product code`
  (natural key, not an auto-id). Mirrors the Xero-linked schema: cost/landed
  cost, cut-to-order fields, Xero item/account codes.
- **Suppliers**, **Customers** — reference lists.
- **Deliveries** — goods-in, linked to Products by `PRODUCT CODE` and
  tracked by batch/lot number + best-before date.
- **Orders** / **OrderLines** / **Allocations** — order header, order lines,
  and stock allocated to each line (FEFO-aware — see Allocations page logic).

**Important convention:** when adding a column to an existing tab, always
append it to the END of that tab's `headers` array in `setup.html` — never
insert in the middle. Setup only ever rewrites row 1 if the stored header is
shorter than required; inserting a column midway silently relabels an
existing column and misaligns every row's data under the new header.

## Pages
```
index.html          Dashboard: config, sign-in, nav
setup.html            Auto-creates required tabs + headers — run first on a new sheet
products.html           Item master
suppliers.html            Supplier list
customers.html              Customer list
deliveries.html                Goods-in log
stock.html                       Stock view
orders.html                        Orders + lines
lists.html                           Pick/pack lists
```

## Working here
- Always confirm you're in `~/provisionslondon` (not a Downloads copy or
  the loose `provisions-*.html` prototype files in `~/Downloads`) and on a
  clean, up-to-date `main` before editing.
- Setup is safe to re-run: it only writes a tab's header row if that tab's
  stored header is shorter than required — never overwrites existing data.
