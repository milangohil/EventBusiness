# EventBusiness
Handles Challenges for Event Business

## Stage 1 app: enquiries, bookings, availability, payments

A mobile-friendly web app for all halls. **No domain, server or database needed.**
All data is saved in a Google Sheet you own, and the app is hosted free by Google Apps Script.

### What it does
- **Home:** greeting, key numbers, a "Needs attention" list (overdue follow-ups, balances due before an event, past events to close) and the next 14 days of events.
- **Bookings:** search, hall filter and tabs (Active, Enquiries, Booked, Completed, Lost & cancelled), grouped by month.
- **+ New:** tap-to-choose form. Hall cards show live **Free / Booked / Enquiries** for the chosen date and slot, and the total is worked out from guests × rate.
- **Booking page:** one-tap Call and WhatsApp, a progress tracker with next-step buttons (Confirm booking, Mark as lost, Mark event done...), follow-up reminders, a payment progress bar, a quick "Record payment" sheet, and ready-made WhatsApp messages.
- **Calendar:** month view for all halls or one hall. Tap a day to see every slot, then tap **Book** on a free slot to start a booking with the hall, date and slot filled in.
- **Payments:** collected this month, outstanding and overdue amounts, a "To collect" list with reminder buttons, and the full payment history.
- Works on phone (bottom tabs) and desktop (sidebar), with automatic dark mode.
- **Built-in rules:**
  - A hall cannot be double-booked.
  - A booking needs a hall, date, slot and guest count.
  - A lost enquiry needs a reason.
  - If two people edit the same record at once, the second save is stopped instead of overwriting the first.

### Kitchen (Stage 2)
Replaces the 3-page Word file made for every event.
- **Dishes & recipes:** add each dish once (Hindi name, category) with its ingredients **for 100 guests**.
- **Ingredients & suppliers:** each ingredient has a unit (kg, litre, piece...) and a supplier with a WhatsApp number.
- **Menu builder:** on any enquiry or booking, tick dishes (for example during food tasting) and mark the menu **Final**.
- **Kitchen sheet:** shows the Hindi menu card and a purchase list grouped by supplier, scaled to the guest count and rounded up to sensible buying amounts. Devji bhai checks it, edits quantities, adds extra items (gas, disposables), then taps **Approve**.
- **Supplier orders unlock only after approval.** Each supplier gets a ready WhatsApp order. **Share list** sends the full list to Vinod bhai, and **Print / PDF** prints the menu, purchase list and ingredient list.
- **Safety checks:** if the guest count or menu changes after approval, the sheet is flagged **Re-check** and needs approving again.
- **The 6-day rule:** Home and Kitchen > Events flag every booking within 6 days whose menu or quantities aren't approved yet.

### Where the data lives
| Sheet tab | One row per |
|---|---|
| `Records` | enquiry or booking |
| `Payments` | payment received |
| `Halls` | hall. Edit this tab to rename, add or deactivate halls (set Active to `No`) |
| `Suppliers`, `Ingredients`, `Dishes` | supplier / ingredient / dish (best edited in the app) |
| `Recipes` | ingredient used in a dish, with quantity per 100 guests |
| `Kitchen` | booking's menu and approved purchase list |

Rensi and the owner can open the Google Sheet at any time to view or export data.

### One-time setup (about 15 minutes)
1. Go to <https://sheets.google.com> and create a blank sheet. Name it **Event Business Data**.
2. In the sheet, open **Extensions > Apps Script**.
3. In the editor:
   - Replace the contents of `Code.gs` with [`apps-script/Code.gs`](apps-script/Code.gs).
   - Click **+ > HTML**, name the file `Index`, and paste in [`apps-script/Index.html`](apps-script/Index.html).
   - Click **Project Settings** (gear icon), tick **Show "appsscript.json" manifest file**, then replace that file's contents with [`apps-script/appsscript.json`](apps-script/appsscript.json).
4. Choose `setup` from the function menu at the top and click **Run**. Allow the permissions it asks for. The `Records`, `Payments` and `Halls` tabs appear in your sheet.
5. Click **Deploy > New deployment**, choose the type **Web app**, and set:
   - Execute as: **User accessing the web app**
   - Who has access: **Anyone with Google account**

   Click **Deploy** and copy the **Web app URL**. That link is your app.
6. **Give staff access:** share the Google Sheet (Share button, **Editor**) with each staff member's Gmail address. Only people the sheet is shared with can use the app.
7. Send staff the Web app URL. On their phone they open it, sign in with Google, and use **Add to Home screen** so it works like an app.

> The first time each person opens the app, Google shows *"Google hasn't verified this app"*. This is normal for private business apps. Click **Advanced > Go to Event Business (unsafe)** and allow access. It's asked only once.

### Updating the app later
Paste the new code into the editor. **If `Code.gs` changed, run `setup` once more** (it adds any new tabs and never touches existing data). Then click **Deploy > Manage deployments > Edit (pencil) > Version: New version > Deploy**. The URL stays the same.

### Try it without Google
Open `apps-script/Index.html` directly in a browser. It runs in **demo mode** and saves data only in that browser. Tap **Load sample data** to see it filled with example bookings.

### For developers
`node --test tests/*.test.js` runs the server logic in `Code.gs` against an in-memory fake of Google Sheets.

### Coming next
- **Stage 3:** supplier orders and automatic WhatsApp messages and payment links.
- **Stage 4:** inventory and owner dashboard.
