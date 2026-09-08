# AC-34 Matiala — SIR 2026 Hearing Dashboard

A browser-based hearing and notice dashboard for AC-34 Matiala.

## Workflow

1. Download the latest ECI `NOTICE_REPORT_PART_WISE` Excel export.
2. Open the deployed dashboard.
3. Upload the Excel file using **Upload latest Excel**.
4. Select any hearing date.
5. Review overview, officer-wise summary, PS-wise details, and the complete ECI status table.

The workbook is processed locally in the browser. The source ECI data does not need to be stored in GitHub.

## Expected workbook

The preferred input workbook contains:
- `ECI_INPUT` — current ECI report, with the standard AC-34 columns.
- `PS_MASTER` — polling-station/officer/BLO/supervisor master information.
- `HEARING_DATA` — hearing dates and scheduled PS records.

The current sample workbook has 430 polling stations and 19 hearing dates. The app filters the ECI report to AC Number 34 / MATIALA and uses the master/schedule sheets when present.

## Local development

```bash
npm install
npm run dev
```

Then open `http://localhost:3000`.

## Deploy on Vercel

Import this GitHub repository (`anacambridge-max/hearing_dashboard`) into Vercel as a Next.js project. The default build command is `next build` and the output is ready for Vercel hosting.
