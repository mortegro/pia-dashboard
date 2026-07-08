# PIA Dashboard

Dashboard for reviewing PIA Leistungen (services) per Mitarbeiter, loaded from
an `.xls`/`.xlsx`/`.xml` export and matched against the Vergütungstabellen in
[`pia-verträge/`](pia-verträge).

## Structure

- `index.html`, `css/style.css`, `js/app.js` — the frontend.
- `main.ts` — Deno server: serves the static files and exposes
  `/api/tarif-periods`, which reads the dated
  `PIA-Leistungen-Preise-<start>-<end>.csv` files in `pia-verträge/`.
- `pia-verträge/` — source contract documents (PDF/docx) and the CSV price
  tables the app parses; a new rate period only needs a new CSV named
  `PIA-Leistungen-Preise-YYYYMMDD-YYYYMMDD.csv` dropped in this folder.
- `vendor/` — third-party libraries (Bootstrap, SheetJS), vendored as-is.
- `dist/` — build output for desktop packages (git-ignored).

## Development

```sh
deno task dev
```

Serves the app at the address printed by Deno; open it in a browser.

```sh
deno task fmt   # format main.ts and js/app.js
deno task lint  # lint main.ts and js/app.js
```

## Desktop packaging

```sh
deno task desktop:mac        # dist/macos/PIA-Dashboard.app
deno task desktop:mac:dmg    # dist/macos/PIA-Dashboard.dmg
deno task desktop:windows    # dist/windows/PIA-Dashboard.msi
deno task desktop:all        # all targets
```
