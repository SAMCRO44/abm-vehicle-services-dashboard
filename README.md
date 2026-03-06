# ABM Rental Operations Dashboard

A static dashboard for ABM fleet/rental operations with Excel import, charts, and staffing trends.

## Deploy to Vercel

1. Install Vercel CLI (if needed): `npm i -g vercel`
2. From this folder, run: `vercel`
3. Follow the prompts to create a new project named "abm-rental-operations-dashboard"

Or connect this folder to Vercel via [vercel.com](https://vercel.com) → New Project → Import.

## Local

Open `index.html` in a browser, or serve with:

```bash
npx serve .
```

Then visit http://localhost:3000

## Features

- **Overview**: Moves, workers, hourly activity, task mix, top workers
- **Workers**: Sortable table with volume, speed, sub-min, gap
- **Task Types**: Avg time by type, full breakdown
- **Locations**: Start/end locations, worker × destination matrix
- **Hours & Pay**: Payroll summary, hours distribution
- **Performance**: Sub-min moves, gap stats, scans heatmap
- **Staffing Trends**: Actual vs needed hours, FTE variance, monthly summary
- **Compare**: Week-over-week and day-of-week patterns
- **Date Filter**: Presets (Today, Week, Month) + custom range
- **Import**: Drag-and-drop or click to import `.xlsx` / `.xls` files

## Supported Data Files

- **Field Ops Task** (full or Export format) — Workers, task types, locations, hourly activity, gaps
- **Summary** — Task type breakdown with avg/slowest/fastest
- **Scans 2.7.26** — Worker stats (completed by name, count, avg time)
- **<1 Minute Pivot** — Sub-minute move counts per worker
- **Pivot** — Worker × destination matrix
- **Names-IDs** — ID to name mapping
- **Hours 2.7.26** / Payroll — Rate type and hours (Reg, OT, Lunch, PTO)
- **HoursOverview** — Monthly sheets (May 2025, FEB 2026, etc.) or Sheet1 with Actual Driver Hours
- **Scans per hour** — Worker × date heatmap

Upload any combination of these files; data is merged automatically.

## Future Development

- **Location**: `C:\Users\luiss\ABM-Rental-Operations-Dashboard`
- **Live**: https://abm-rental-operations-dashboard.vercel.app
- **Stack**: Static HTML/CSS/JS, Chart.js, SheetJS (XLSX)
- **Deploy**: `vercel --prod --yes`
