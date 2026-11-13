# Tests

These run the panel in a headless browser against a stand-in for Word (`mock-office.js`), so they need no Word and no Mac.

GitHub runs all of them on every push (`.github/workflows/test-and-deploy.yml`). The live panel is only updated from `main`, and only when every check passes.

Requirements: Python 3 with Playwright (`pip install playwright` then `playwright install chromium`), Node.js, and ESLint 9 or later for the lint step.

From the repository folder:

```
node tests/tools-test.js           # address, CSV, barcode and scan-measuring helpers
node tests/worker-test.mjs         # the Cloudflare Worker's request handling (worker/)
python3 tests/e2e/run.py           # sheets, printing, addresses, table window, shipping labels
python3 tests/e2e/run15.py         # stock, fill other sheets first, numbers and codes, scans, backups
python3 tests/e2e/run16.py         # sheet list clean-up, editing addresses, undoing older prints
python3 tests/e2e/run17.py         # alignment squares, recent addresses, groups, labels saved
python3 tests/e2e/run18.py         # linking a phone, labels sent from it, the phone page
python3 tests/e2e/run19.py         # phone, email and second addresses; choosing contacts to import
python3 tests/e2e/run20.py         # address suggestions in the panel's address boxes
python3 tests/e2e/run21.py         # label text where Word puts it, suggestions on labels, print preview
python3 tests/e2e/run22.py         # signing in and sync in the panel, and the sign-in window
node tests/sync-test.mjs           # several pretend computers syncing through the Worker's code
npx eslint -c tests/eslint.config.cjs taskpane.js table.js lsk-tools.js sheetdoc.js preview.js lsk-account.js lsk-config.js auth.js
```

Each script ends with "N checks, 0 failed" and exits with an error if any check failed. Screenshots and files they make go in `tests/out/`, which isn't committed. Each script serves the repository on its own local port (8790, 8791, 8794 to 8800).
