# AC-34 Matiala — SIR 2026 Hearing Dashboard

A browser-based hearing and notice dashboard for AC-34 Matiala.

## Normal workflow

1. Keep the AC-34 master/schedule information loaded in the browser using the existing dashboard workbook once.
2. Every time a fresh ECI `NCT OF Delhi_NOTICE_REPORT_PART_WISE` Excel is downloaded, click **Upload ECI Excel**.
3. The app reads the ECI `sirNoticeGenerate` sheet (or compatible `ECI_INPUT` sheet), filters AC 34 / MATIALA, joins the live ECI figures to the PS master, and rebuilds the dashboard.
4. Select any hearing date to see that day's scheduled workload.

The ECI source workbook is processed locally in the browser and is not stored in GitHub.

## Hearing schedule

The application also accepts the separate **Part Wise Hearing Summary** workbook containing:
- `Part No.`
- `Hearing Date`
- `Total Hearings`

Use **Update from Part Wise Hearing Summary** when the hearing schedule changes. After the schedule is loaded, the routine update remains the ECI upload only.

## Dashboard views

- Overview with workload and delivery charts
- Officer-wise summary
- PS-wise detail including officer, BLO, supervisor, locality, polling area, anomaly/mapping and voter fields
- Raw ECI status/audit table
- Hearing-date selector and latest-date shortcut
- Search across PS/officer/centre/BLO/supervisor/locality/polling area

## Data model

The supplied AC-34 workbook provides the 430-row PS master and existing hearing schedule. The separate Part Wise Hearing Summary can update the schedule. The ECI report supplies the live notice and hearing status fields.

## Local development

```bash
npm install
npm run dev
```

Then open `http://localhost:3000`.

## Deploy on Vercel

Import this GitHub repository (`anacambridge-max/hearing_dashboard`) into Vercel as a Next.js project. The default build command is `next build`.
