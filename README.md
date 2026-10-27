# Workhours IT

A local-first browser app for recording IT work for **EC, ENAB, 4SOME and SEABOW**. Split a task across companies, see monthly company totals, search/filter entries, edit/delete work, and switch between light and dark modes.

## Main and test versions

- `main`: stable release branch. It currently retains the original repository scaffold until the owner explicitly requests promotion.
- `test`: the working app and all development changes. Work only on this branch.
- Do not merge or push to `main` without the owner's explicit instruction. See `AGENTS.md`.

## GitHub Pages test website

This app can be hosted like LoadingListChecker. The deployment workflow builds and tests **only `test`**, then publishes **only `dist/`** through GitHub Pages. It never merges, commits or pushes to `main`. GitHub Actions used by the workflow are pinned to verified upstream commits.

The expected website address after a successful deployment is **https://ricoro.github.io/Workhours-IT/**. Saving the workflow alone does not make this address live.

One-time owner setup:

1. Open **https://github.com/Ricoro/Workhours-IT/settings/pages**. Under **Build and deployment → Source**, choose **GitHub Actions**.
2. In **Settings → Environments → github-pages**, ensure the deployment branch rules allow **`test`** (selected branches: `test`). If the environment does not exist yet, create it with that name and allow `test`.
3. In the repository's **Actions** tab, open **Deploy test app to GitHub Pages** and rerun the run triggered by the latest push to `test` if it failed before Pages was enabled. Later pushes to `test` deploy automatically after all tests pass.
4. Use the website URL shown by the successful deployment. The owner has made the repository public so it can use GitHub Pages without private-repository plan restrictions.

The website is normally publicly accessible, even from a private repository. It contains the app, not your recorded hours: entries stay in your browser and Excel writes go to your chosen local folder. Publishing is not a login system. Actual OneDrive upload remains the responsibility of the desktop sync client.

The hosted site and localhost have separate browser storage. To move existing entries, download a JSON backup from the old address and restore it at the new address. Continue making backups. This cloud environment cannot access the Pages settings API or verify the hosted URL; the owner must enable Pages and inspect deployment success in GitHub.

Reproduce the deployed-path checks locally:

```sh
npm run build
PLAYWRIGHT_APP_PATH=/Workhours-IT/ PLAYWRIGHT_SERVER_COMMAND="npm run preview -- --base=/Workhours-IT/ --port 5174 --strictPort" npm run test:e2e
```

## Open the test app on your computer

Install Node.js 22+ and Git, then use a terminal:

```sh
git clone --branch test --single-branch https://github.com/Ricoro/Workhours-IT.git
cd Workhours-IT
npm ci
npm run dev -- --port 5173 --strictPort --open
```

If you already have the repository, preserve any local changes, switch with `git switch test`, and run the npm commands from that checkout. The repository is public; its source code can be downloaded without signing in.

The app opens in your browser at **http://127.0.0.1:5173/**. Keep the terminal running while using it; stop with Ctrl+C. This address works on the computer running the command. It is not an externally hosted site. On later visits, start it using the same command and browser profile so your saved entries remain available.

## Run locally

Requires Node.js 22+ and npm.

```sh
npm ci
npm run dev
```

Open the local address printed by Vite. Use the **same address, port and browser profile** every time: browser storage belongs to that origin. The default binds only to your computer. No backend, database server, account, API key or Microsoft registration is required.

```sh
npm test          # Validation and real Excel roundtrip tests
npm run test:e2e  # Browser persistence, editing, export, backup/restore and folder writing
npm run build    # Production files in dist/
npm run preview  # Serve the production build locally
PLAYWRIGHT_SERVER_COMMAND="npm run preview -- --port 5174 --strictPort" npm run test:e2e
                 # Test the built production app on its own server
```

Tests start their own server on port 5174 and never reuse your development server. Browser tests use `/usr/bin/chromium` in the cloud environment. GitHub Actions explicitly uses the Playwright-managed browser installed for the pinned package version, so a preinstalled system browser cannot override it. Elsewhere set `PLAYWRIGHT_CHROMIUM_EXECUTABLE_PATH` to your Chrome/Chromium executable, or run `npx playwright install chromium` to use Playwright's browser.

## Record and export hours

1. Choose the date and describe the work. Employee ID is automatically fixed to `12` in Excel; there is no employee field to fill in.
2. Enter decimal hours for each company involved. `1.5` is one hour and 30 minutes. Each task can total up to 24 hours, with two decimal places.
3. Click **Save entry**. The committed entry is stored in IndexedDB; unfinished form changes are saved as a draft in localStorage. Editing a task replaces it rather than adding duplicate hours.
4. Select a month and choose **Export month** for a real `.xlsx` workbook. Export includes all companies for that month, regardless of activity search/filter settings.

Exports preserve the example workbook's seven columns: `Account`, `Date`, `Quantity`, `Employee`, `Item`, `Notes`, `Project`. Each company allocation is a separate Excel row, with a numeric quantity, Excel date, `WH` item and `WH_<company>` project. Tabs use full English month names and years, such as `October 2026` and `January 2027`. A Summary sheet and monthly company totals are included. Existing entries and restored backups also export with employee `12`. The app generates a new `.xlsx`; it does not modify or import the supplied `.xlsb` workbook, historical records, formatting or macros. No example company records are bundled in the app or committed.

## Import your existing Excel locally

Click **Import Excel** next to Export month and select an `.xlsb` or `.xlsx` file (up to 20 MB). Your file is parsed in the browser, never uploaded. The parser is bundled locally and loaded only when needed. It understands the original Account, Date, Quantity and Notes columns, shifted column positions, Dutch/English month tabs, and the new app's exports.

Before saving, the review shows:

- Entries to add, proposed date corrections, unfinished rows and ignored non-work sheets.
- A choice between the worksheet's month/year (keeping the day) or original cell dates. Review date corrections carefully; impossible dates and invalid rows block the whole import.
- Exact matches and previously imported source rows are skipped. Same-company/date/hours rows with different descriptions are flagged as **possible duplicates** and skipped by default. Select **Import this row anyway** only when it represents different work.

Use **Download existing backup** before confirming. Import adds new records without replacing existing ones. If another tab changes entries during review, the entire commit is rejected and the file must be selected again. Reimporting the same file does not add duplicate source rows. Legitimate repeated rows in the source are kept according to their occurrence counts.

Under **Backup & restore**, **Undo last Excel import** removes unchanged entries from the most recent Excel import. Entries edited afterwards and unrelated/manual entries are kept. Only the latest import can be undone; JSON restore is a separate merge operation. Linked Excel output refreshes after import and undo.

To migrate the old workbook into your current app, select **IT UREN.xlsb**, review date differences and possible overlaps, then confirm. No download from this chat is required. For a fresh browser with no entries, import the current `IT-hours.xlsx` first, then the old `.xlsb`. Company records and original files must never be committed to the public repository.

## Automatic Excel file and OneDrive

In **desktop Chrome or Edge**, use **Link Excel folder** and grant access to a folder you choose. After every committed save, edit, deletion or restore, the app rewrites `IT-hours.xlsx` with **all stored months**, including summary totals. Existing files with that name require confirmation before replacement. Export failure does not discard the saved browser entry; the status explains the failure. If Excel locks the file, close it and reconnect the folder to retry.

Choose a folder inside your locally synced **OneDrive** directory to let the OneDrive desktop client upload the Excel file. This is desktop folder synchronization, not direct Microsoft Graph integration. The app cannot verify OneDrive's cloud sync status. Excel writing occurs while the app is open and folder permission is granted; reconnect after a browser restart if prompted. Use a different empty folder for another independent installation; avoid simultaneous writers to the same Excel file.

Browsers without folder access still support saved entries, Excel downloads and JSON backups. Folder access requires localhost or HTTPS.

## Backup and privacy

- Entries and drafts remain in this browser profile. There is no analytics, remote font request, data upload, backend or public database. Application dependencies are bundled locally.
- Browser storage is **not a backup** and is not application-encrypted. Clearing site data, changing origins/profiles, private browsing, or losing the device can lose entries. Use a managed company device and appropriate OS protections.
- Use **Backup & restore → Download backup** regularly. Restore validates the entire backup before importing, merges entries by ID, and replaces matching IDs. JSON backups retain entry IDs; Excel exports are not restorable through the app.
- Browser entries do not sync between devices. OneDrive receives the generated workbook only through its desktop client; it does not synchronize the app's browser database. Concurrent edits across devices are not supported.
- Do not commit workbooks, backups or confidential records. `.gitignore` excludes `.xlsb`/`.xlsx`; keep JSON backups outside the checkout.
- `private: true` in `package.json` prevents npm publishing. **It does not make the GitHub repository private.** The owner has made **Ricoro/Workhours-IT public**; source code and Git history are public, while entries and chosen local Excel files are not uploaded to GitHub. Also verify deployment access separately; a private GitHub repository does not protect a publicly hosted app.

## Cloud development

Use the existing `/workspace/Workhours-IT` checkout; cloud tasks are already isolated, so no worktree is needed. Install with `npm ci --cache /workspace/.npm-cache`, then `npm run dev`. Live processes must be started again in a new task. The uploaded workbook is outside the checkout and is not needed at runtime.
