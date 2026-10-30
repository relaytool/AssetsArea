# Assets Area — nested warehouse storage blocks

This is the updated GitHub Pages project. It adds an independent, two-column **A / B storage-block map inside each warehouse area**, based on the original Assets Area layout. Everything remains in the browser; live data is in Google Sheets.

## Deploy

Replace these files in your GitHub Pages repository, keeping them in the same folder:

- `index.html`
- `styles.css`
- `app.js`
- `config.js`

No Apps Script, build tool, or new Google OAuth client is needed. The configured spreadsheet and owner email remain the same.

## Enable nested storage blocks (once)

1. Backup your existing spreadsheet.
2. Open the deployed website, select **Owner edit**, and sign in with the configured owner account.
3. Select **Enable area blocks**. This creates two new tabs if they are missing, while leaving existing tabs, areas and inventory untouched.
4. Open any holding area. It now shows **A01, B01, A02, B02**, etc. The default is **8 rows per lane** (16 blocks per area). Change the row count between **2 and 30** and select **Save block rows**.
5. Select a block to view or update its stock. Each block can contain multiple distinct asset types and quantities.
6. Existing inventory in `Warehouse Stock` appears as **Stock not assigned to a block**. Choose a destination block and select **Move unassigned stock into block** to allocate it without changing quantities.

If the prior warehouse upgrade was never activated, select **Activate warehouse layout** first, then **Enable area blocks**.

## Storage and access

- **Warehouse** → **area** → **storage block** → multiple asset types.
- Block IDs are scoped by area. For example, `WH-A01/B02` and `WH-B02/B02` are different locations.
- Open `?area=WH-A01` to see an area and its blocks. Open `?area=WH-A01&block=B02` to open a specific block directly.
- Open a block to **add**, **remove**, **transfer all to another block**, or **empty the entire block** with confirmation. All actions update Google Sheets and attempt to append to `Area Activity`.
- **Empty all blocks in this area** clears all nested blocks **and** any unassigned stock, with confirmation.
- The warehouse overview totals items per asset type, counting both nested blocks and any unassigned inventory, without double-counting.
- Users with the public link are view-only. Only the configured owner, authorized through Google OAuth and Google Sheets file permissions, can edit data.

## Workbook structure

| Sheet | Columns | Purpose |
| --- | --- | --- |
| `Warehouse Areas` | Area ID, Name, Zone, Type, X, Y, Width, Height, Color, Updated At, Active | Existing warehouse layout |
| `Warehouse Stock` | Area ID, Item, Quantity, Updated At | Existing area-level stock, now treated as *unassigned* until allocated |
| `Warehouse Block Settings` | Area ID, Rows, Updated At | New: row count per area; 8 if absent |
| `Warehouse Block Stock` | Area ID, Block ID, Item, Quantity, Updated At | New: multiple asset records for each storage block |
| `Area Activity` | Timestamp, Action, Item, Quantity, From, To, Note, By | Existing audit history |

**Important:** New block tabs are created by the signed-in owner via **Enable area blocks**. They are *not* created by public viewers. If creating them manually, the header names and their order must match the table above exactly. Do not delete `Warehouse Stock` or move its rows without using the app to assign the stock.

## Security and caveats

Share the warehouse spreadsheet as **Anyone with the link → Viewer**, with Editor access restricted to the configured owner. Because a viewer can access all tabs in a publicly shared Google Sheet, keep private information in a separate workbook.

The Google Sheets API has no transactional row locking. The app compares remote values before writes and verifies the result, but if multiple owners edit the same blocks concurrently they should refresh and reconcile changes. Save failures do not silently replace quantities. An activity-log write can fail separately from an inventory write; if that happens, the app warns you after saving stock.

The original legacy tabs, existing warehouse areas, IDs, and existing stock remain unchanged when nested storage is first enabled.
