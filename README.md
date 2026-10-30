# Assets Area — warehouse layout

Static GitHub Pages site backed by Google Sheets. No server or build step is required.

## Publish

Upload `index.html`, `styles.css`, `app.js`, and `config.js` together to the existing GitHub Pages repository. The spreadsheet ID, existing Google OAuth client ID, and owner account are preserved in `config.js`.

## First-time upgrade

1. Keep a backup of the original Google Sheet before upgrading.
2. Open the website and select **Owner edit**, signing in with the account specified in `config.js`.
3. Select **Activate warehouse layout**. The site creates `Warehouse Areas` and `Warehouse Stock` tabs and copies the current legacy blocks and their inventories. The existing `Area Positions`, `Area Activity`, and `Area Settings` tabs remain untouched.
4. Select **Design layout** to drag/resize existing areas and add new holding, staging, loading, aisle, or other areas. Each new area has a stable ID.
5. Select a holding area to add/remove different assets, move all contents to another area, or empty it (with confirmation). Owners can also long-press a stocked block and drag its contents onto another storage block; matching asset types merge. **Copy area link** produces `?area=AREA-ID` for read-only sharing.

The site uses **Warehouse Stock** as the live source of truth after activation. Do not continue editing the old `Area Positions` tab after migrating; edits there will not be reflected in the new layout. Run the activation **once**, and verify counts against the physical inventory before regular operation.

## Google Sheets access

For an anonymous public view, share the **dedicated warehouse spreadsheet** as **Anyone with the link → Viewer**. Make sure only the owner has **Editor** permission. The website verifies the signed-in Google account is the configured owner before providing edit controls, and Google Sheets API enforces that account's actual file permissions.

**Privacy:** All tabs of a publicly shared spreadsheet may be accessible to viewers. Do not put sensitive or private data into this public workbook. If you need private operational sheets, keep those in a different, restricted workbook.

## New sheets

| Tab | Columns |
| --- | --- |
| `Warehouse Areas` | Area ID, Name, Zone, Type, X, Y, Width, Height, Color, Updated At, Active |
| `Warehouse Stock` | Area ID, Item, Quantity, Updated At |
| `Area Activity` (existing) | Timestamp, Action, Item, Quantity, From, To, Note, By |

Area IDs are permanent. Deleting an empty area archives its record (Active = FALSE) so past links are not reassigned. One asset type per row in `Warehouse Stock` allows unlimited different types per area.

## Notes

- The floor planner uses a 16-column, 80-row maximum grid; empty spaces can represent aisles/roads. Walkways are read-only areas and cannot hold stock.
- The public viewer refreshes roughly every 90 seconds and when returning to the tab. Manual refresh is also available.
- Writes use the signed-in owner's Google OAuth token, check current sheet values for changes, and verify saved values before reporting success. The backing spreadsheet itself is not changed until the owner activates the upgrade on the live site.
- Public users cannot create, remove, move, or edit stock from the website. The live Google Sheets sharing settings must also remain **Viewer** for the public.
- Google Sheets does not provide transactional row locking: use a single owner editor at a time and refresh if someone else edits the sheet directly.
