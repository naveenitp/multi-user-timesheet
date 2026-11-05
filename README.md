# Timesheet: multi-user version

This folder is a **separate version**. Your existing single-user app (`index.html` + `Code.gs` in the folder above)
is not changed and keeps working with its own Google Sheet and its own web app URL.

## What is different
| | Single-user (existing) | Multi-user (this folder) |
|---|---|---|
| Sign-in | none | username + password |
| Where entries are saved | one Google Sheet | **each person gets their own Google Sheet file** ("Timesheet - Name") |
| Masters (business names, products, status, source) | yours | shared lists, **edited by admins only**; users can use them in dropdowns but cannot open the Masters tab |
| SQL Library | yours | **admins only**; it is never sent to normal users |
| Quick templates | in Masters | personal, saved in each person's own file |
| Accounts | n/a | created by an admin in the **Users** tab |

Everything else (dashboard, records, export, themes, pets, weather, overlap warning, duplicate yesterday ...) is the same.

## Set up (about 10 minutes)
1. Create a **new** Google Sheet (name it e.g. "Timesheet Control"). This is NOT your current timesheet sheet.
2. In it: **Extensions > Apps Script**. Delete the sample code and paste all of `multiuser/Code.gs`.
3. At the top change `ADMIN_USER` and `ADMIN_PASS` (use a strong password, at least 8 characters).
4. Select **setupAdmin** in the function dropdown and click **Run**. Approve the permissions (it needs to create sheets).
   This creates the Users, Masters and SQL Library tabs and the admin's own timesheet file.
5. **Deploy > New deployment > Web app**. Set **Execute as: Me** and **Who has access: Anyone**. Deploy and copy the URL ending in `/exec`.
   (New deployments default to "Only myself", which makes sign-in fail. Change it to Anyone.)
6. Host `multiuser/index.html` (see below) and open it. Click **Server address** under the sign-in box, paste the URL, then **Test connection**.
   You should see "Connection is fine".
7. Sign in as the admin, change the password (Password button), then open **Users** and create accounts.
8. Optional: in **Masters** add your business names, products, statuses and sources. Everyone sees them in their dropdowns.

Each person's file is created in your Google Drive and stays private to you. Use **Share sheet** in the Users tab only if you want
a person to be able to open their own file (view only).

## Hosting it next to the existing site
GitHub Pages publishes one branch per repository, so pointing Pages at a new branch would replace your current site. Pick one:

**A. Second repository (safest)**
```
git init timesheet-multiuser && cd timesheet-multiuser
cp /path/to/multiuser/index.html .
git add . && git commit -m "Multi-user timesheet"
git branch -M main && git remote add origin https://github.com/<you>/timesheet-multiuser.git && git push -u origin main
```
Then Settings > Pages > main.

**B. A branch in your current repository (to keep the code together, not live)**
```
git checkout -b multi-user
cp /path/to/multiuser/index.html multiuser-index.html     # or replace index.html on this branch only
git add . && git commit -m "Multi-user version" && git push -u origin multi-user
```
`main` stays exactly as it is.

**C. Same site, extra folder:** put `multiuser/index.html` in a `multiuser` folder of the current repository. The existing `index.html` is untouched.

## Good to know
- The admin is the owner of every file, so the admin (you) can open all of them. "No access" applies to normal users.
- Passwords are stored as salted, repeated SHA-256 hashes. Use long passwords. Five wrong tries lock an account for 15 minutes.
- Sessions last 6 hours. Google may clear them sooner, in which case people simply sign in again.
- Forgot the admin password: run **resetAdminPassword** in the Apps Script editor.
- If you change `Code.gs` later, redeploy with **Deploy > Manage deployments > Edit > New version** so the URL stays the same.
- This version stores its own browser data under separate keys, so it does not mix with the single-user app even on the same website.
