# Slot Booking

A free online booking page for shared slots in an apartment building (laundry room, 07:00–21:00, 1-hour slots).

Residents book online, get a confirmation email, and get a reminder before their slot asking them to
confirm or release it. Unconfirmed slots are freed automatically so other residents can use them.

## How it's built

```
Booking page (index.html, on GitHub Pages)
        │  "which slots are free?" / "book this slot"
        ▼
Backend (apps-script/Code.gs, on Google Apps Script)
        │  reads & writes                    │  sends emails
        ▼                                    ▼
Google Sheet ("Bookings" tab)        Confirmation · Reminder · Slot released
```

| Part | File | What it does |
|------|------|--------------|
| Frontend | `index.html` | The page residents see: day cards, slot buttons, booking form |
| Backend | `apps-script/Code.gs` | Checks free slots, saves bookings, sends emails, runs the 15-minute timer |
| Database | Google Sheet | One row per booking; the building can open it any time |

**Booking status flow:** `booked` → `confirmed` (clicked "Yes") / `released` (clicked "Release") / `expired` (didn't answer in time).

**Timing rules** (change them in `CONFIG` at the top of `Code.gs`):
- reminder 2 hours before the slot
- not confirmed 30 minutes before → slot released
- max 2 upcoming bookings per apartment

## Setup (about 15 minutes)

### 1. Backend: Google Sheet + Apps Script
1. Create a new Google Sheet (ideally on a shared building account, not a personal Gmail).
2. In the Sheet: **Extensions → Apps Script**.
3. Delete the sample code, paste in everything from `apps-script/Code.gs`, and save.
4. **Project Settings (gear icon) → Time zone → (GMT+02:00) Helsinki**.
5. Choose the `setup` function in the toolbar and click **Run**. Allow the permissions Google asks for.
   This creates the "Bookings" tab and the 15-minute timer.
6. **Deploy → New deployment → type: Web app**
   - Execute as: **Me**
   - Who has access: **Anyone**
7. Click **Deploy** and copy the **Web app URL** (it ends with `/exec`).

### 2. Frontend: GitHub Pages
1. Open `index.html` and paste the Web app URL into `API_URL`.
2. In the GitHub repo: **Settings → Pages → Branch: main → / (root) → Save**.
3. After a minute your page is live at `https://<your-username>.github.io/slot-booking/`.

### 3. Test
Book a slot with your own email, check the confirmation email, and check the row in the Sheet.

## Changing the backend later
After editing `Code.gs` in Apps Script: **Deploy → Manage deployments → Edit (pencil) → Version: New version → Deploy**.
This keeps the same URL, so the page keeps working.

## Limits
A free Gmail account can send about 100 emails a day from Apps Script — plenty for one building.
