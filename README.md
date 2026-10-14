# EventBusiness
Handles Challenges for Event Business

## Stage 1 app: enquiries, bookings, availability, payments

A mobile-friendly web app for all halls. **No domain, server or database needed.**
All data is saved in a Google Sheet you own, and the app is hosted free by Google Apps Script.

### What it does
- **Home:** open enquiries, follow-ups due, upcoming events, balance to collect.
- **+ New:** add an enquiry or booking. It shows straight away whether the hall is free for that date and slot.
- **Records:** search and filter every enquiry and booking. Tap one to edit it, record payments, or send WhatsApp messages (thank-you, booking confirmation, receipt, balance reminder).
- **Availability:** one date across all halls, or one hall for a whole month.
- **Built-in rules:**
  - A hall cannot be double-booked.
  - A booking needs a hall, date, slot and guest count.
  - A lost enquiry needs a reason.
  - If two people edit the same record at once, the second save is stopped instead of overwriting the first.

### Where the data lives
| Sheet tab | One row per |
|---|---|
| `Records` | enquiry or booking |
| `Payments` | payment received |
| `Halls` | hall. Edit this tab to rename, add or deactivate halls (set Active to `No`) |

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
Paste the new code into the editor, then click **Deploy > Manage deployments > Edit (pencil) > Version: New version > Deploy**. The URL stays the same.

### Try it without Google
Open `apps-script/Index.html` directly in a browser. It runs in **demo mode** and saves data only in that browser, which is useful for a quick look.

### For developers
`node --test tests/*.test.js` runs the server logic in `Code.gs` against an in-memory fake of Google Sheets.

### Coming next
- **Stage 2:** menu builder (Hindi), recipe master, automatic ingredient list 6 days before each event.
- **Stage 3:** supplier orders and automatic WhatsApp messages and payment links.
- **Stage 4:** inventory and owner dashboard.
