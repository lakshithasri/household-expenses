# Household Expenses – phone app

A small web app you add to your phone's home screen. Each person signs in with their own Microsoft
account and adds expenses or income. Each entry is written straight into the shared
`Household_Expenses.xlsx` on OneDrive (tables `tblExpenses` / `tblIncome`) through the Microsoft Graph
Excel API. That API is safe to use while someone else has the file open in Excel.

```
Phone ─ sign in ─▶ Microsoft account ─ add row ─▶ OneDrive Excel ─ sync ─▶ PC dashboard (auto-reloads every 20 s)
```

- **Privacy:** the hosted page is only code. The Excel link, sign-in tokens and cached entries stay on
  each phone (browser storage). Data goes only between the phone and Microsoft.
- **Offline:** entries saved without a connection are queued on the phone and sent automatically when
  you're back online. The header shows "N waiting to sync".
- **Undo:** tap **Undo** on the confirmation message right after saving to delete that row.
- **Demo:** open the page with `?demo=1` at the end of the address to try it with fake data. Nothing is saved.

## One-time setup

### 1. Host the page (free, GitHub Pages)
Publish the contents of this `mobile/` folder to a GitHub repository with Pages turned on. The address
will look like `https://<github-user>.github.io/<repo>/`. Use that exact address, including the
trailing `/`, in step 2.

### 2. Register the app with Microsoft (free)
Microsoft only lets apps it knows about sign people in, so the app needs a registration once.

1. Go to <https://entra.microsoft.com> → **Applications → App registrations → New registration**.
   - If it says you can't create apps outside a directory, first create a free Azure account at
     <https://azure.microsoft.com/free> using the same Microsoft account. This creates a "Default
     Directory". The card is only for identity verification, and app registrations cost nothing.
2. **Name:** `Household Expenses`
3. **Supported account types:** *Personal Microsoft accounts only*
4. **Redirect URI:** platform **Single-page application (SPA)**, value = your Pages URL from step 1.
5. Click **Register**, then copy the **Application (client) ID**.
6. Optional, for testing on the PC: **Authentication → Add URI** → `http://localhost:8611/`
7. **API permissions → Add → Microsoft Graph → Delegated:** `Files.ReadWrite.All`, `User.Read`
   (`openid`, `profile` and `offline_access` are requested automatically). There is no admin consent
   for personal accounts; each person approves once at first sign-in.
   `Files.ReadWrite.All` is needed so your wife's account can write to a file *shared with her*.
   The app only ever opens the one file whose link you paste.

### 3. Put the client ID in `config.js`
```js
clientId: "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
tenant: "consumers",
```
Push the change to GitHub.

### 4. On each phone
1. Share the Excel file with your wife (**Can edit**, **Specific people**) if you haven't already.
2. Open the Pages URL in Safari (iPhone) or Chrome (Android) → **Sign in with Microsoft**.
3. Paste the Excel file's OneDrive link: Share → **Copy link**. Your wife can also open it from
   OneDrive → **Shared** → … → **Copy link**.
4. Pick who's using the phone.
5. Add the page to the home screen. iPhone: Share → **Add to Home Screen**. Android: ⋮ → **Add to Home screen** / **Install app**.

## Notes
- Categories, people, payment methods and budgets are read from the **Settings** sheet each time the
  app opens. Edit them there, not in the app.
- Microsoft limits sign-ins for this kind of app to about 24 hours. After that the app briefly
  redirects to Microsoft and comes straight back. You normally won't need your password again.
- Test locally: `python -m http.server 8611 --bind 127.0.0.1` in this folder, then open
  `http://localhost:8611/` (or `http://localhost:8611/?demo=1`).
