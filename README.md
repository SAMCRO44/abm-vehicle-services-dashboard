# ABM Rental Operations Dashboard

Dashboard for ABM fleet/rental operations. Import one or more Excel exports to view moves, task types, locations, workers, and performance.

## Run locally

```bash
npm run dev
```

Then open http://localhost:3000 and use **Import Excel** to load your files.

## Supported Excel files

- **Field Ops Task** — Raw move data (Status, Task Type, Name, ID, Start/End Location, timestamps, Duration Taken Seconds, Is Blocked By Foundry). Used for overview, workers, and locations.
- **Pivot Table** — Task type summary (Task Type Description, Count, avg/slowest/fastest times), or Start/End Location Title + Count.
- **Summary** — Same shape as task-type Pivot (task type, count, times).
- **Names-IDs** — ID → Name mapping for workers.
- **&lt;1 Minute Pivot** — Sub-minute move counts per worker (Row Labels, Count of End Location).
- **&lt;1 Minute DATA** — Detailed sub-minute moves (same columns as Field Ops Task).
- **Scans 3.5.26** — Completed by Name, COUNT, Average Time (secs).

You can upload multiple files at once; data is merged (e.g. all Field Ops Task rows combined, task types and locations merged by name).

## Deploy (Vercel)

```bash
npx vercel
```

Or connect the repo at [vercel.com](https://vercel.com) → New Project → Import.
