# Shared address book: setup

Every computer running the Warehouse app keeps its own address book and syncs it with one shared copy in a Google
Sheet. The app never waits for the internet: it works on its own copy and syncs when it can.

- **Push:** changes made on a computer go up within about 15 seconds.
- **Pull:** changes from the others come down when the app starts, every 5 minutes, and before a preview (if the
  last pull is over a minute old).
- **Two computers changed the same entry:** a verified entry (confirmed by the courier portal) beats one that isn't;
  otherwise the latest change wins. The losing change is written in the log.
- **Deleting** removes an entry from that computer only. The shared copy and the other computers keep it. It comes
  back on that computer only if someone changes it in the shared copy after the delete.
- **Log:** every change, with the name typed on the Address Book page and the computer's name. There's one log
  spreadsheet per month, listed in the **Logs** sheet.

What is shared: addresses, verified marks and learned Excel spellings. Not shared: Keep as is choices, courier
services, packing specs, service usage, and who collects from the warehouse (Kyle yellow, Gessie blue and "client"
green are the same everywhere anyway).

## 1. The Google Sheet and its script (once)

1. In the company Google Drive, create a Google Sheet named, for example, **Warehouse address book**. Share it only
   with the people who look after it. The app doesn't need it shared: the script runs as you.
2. In the sheet: **Extensions → Apps Script**. Delete what's in `Code.gs` and paste in the whole of
   `cloud/apps_script/Code.gs` from the app folder. Save.
3. Optional, makes busy syncs faster: **Services (+) → Google Sheets API → Add**.
4. **Project Settings (⚙) → Script properties → Add script property**: name `SECRET`, value a long random key.
   To make one, run this on any computer with the app:
   `python -c "import secrets; print(secrets.token_urlsafe(32))"`
5. **Deploy → New deployment → Select type: Web app.**
   - Execute as: **Me**
   - Who has access: **Anyone**. The app signs in with the key, not a Google account, so "Anyone within your
     organisation" would turn it away. Without the key the script answers nothing but "running".
   - Deploy, allow the permissions it asks for (the sheet, and creating the monthly log spreadsheets), and copy the
     **Web app URL** (it ends in `/exec`).

To change the script later, edit it and use **Deploy → Manage deployments → ✏️ → Version: New version**. The URL
stays the same.

## 2. Each computer

1. In the app folder's `.env` file (copy `.env.example` if there isn't one), add:
   ```
   WAREHOUSE_SYNC_URL=https://script.google.com/macros/s/…/exec
   WAREHOUSE_SYNC_KEY=the SECRET from step 4
   ```
2. Restart the app (`python autostart.py restart`, or close and start it again).
3. Open **📇 Address Book** and type your name in the **☁ Shared address book** bar. Nothing goes up until a name is
   given. Pulling works without one.

**The first sync** sends the computer's whole address book up. Entries that are already there (same receiver, Attn
and address) become one entry, and the rule above decides which version stays. Any order works, but connect one
computer at a time, starting with the one whose address book is most complete. That keeps the first log readable.

## 3. When something's wrong

The bar on the Address Book page shows the last pull and push, what's waiting to go up, and the last problem.
**⟳ Sync now** tries straight away. The 📜 Logs page has the details under "sync".

| The bar says | What to do |
|---|---|
| Can't reach the shared address book | No internet: the app keeps working and syncs when it's back. |
| the key … is wrong | `WAREHOUSE_SYNC_KEY` doesn't match the script's `SECRET`. |
| answered 404 / couldn't be read | `WAREHOUSE_SYNC_URL` isn't the web app URL ending in `/exec`, or the deployment was removed. |
| Type your name | Type it in the bar. |

Size: the Addresses sheet holds about half a million entries (Google's 10 million cells per spreadsheet). The log
gets a new spreadsheet each month, in the Drive of whoever deployed the script.
