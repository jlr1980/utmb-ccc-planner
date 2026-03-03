# UTMB CCC Qualification Planner

An Excel-based race planning tool built for Kathleen Reece to qualify for the **CCC (100K)** at UTMB Mont-Blanc 2027 in Chamonix, France.

## What This Is

The CCC is one of three championship races at the annual UTMB Mont-Blanc festival. Getting in requires navigating a qualification system built around two things:

1. **Running Stones** — lottery tickets earned by finishing races on the UTMB World Series circuit. More stones = more entries in the draw.
2. **A valid UTMB Index** — proof you've completed a race of sufficient distance in the last 24 months. For the CCC, that means a 50K, 100K, or 100M finish.

Demand for the CCC is roughly 5x available spots. The lottery drew ~6,000 applicants for ~1,200 spots in 2025, and it's growing ~20-30% per year. The average selected runner had 5.7 stones.

This workbook maps out all 65+ UTMB World Series events worldwide for 2026, lets you pick races, tallies your stones, and estimates your lottery odds.

## What's in the Workbook

| Tab | Purpose |
|-----|---------|
| **All UTMB Events 2026** | Every World Series event globally with dates, location, distances, stone values, CCC index eligibility, distance from Logan UT, and estimated costs. Mark "Y" to select races. |
| **Race Schedule & Stones** | Enter your selected races in chronological order. Auto-calculates cumulative stones and tracks index status. |
| **CCC Lottery Odds** | Probability calculator using 2025 lottery data. Editable assumptions. Includes a scenario table showing odds for 2-12 stones. |
| **Strategy Guide** | Four pre-built race strategies tailored to Logan, UT — from budget-friendly local races to a high-stone plan using the Americas Major. |
| **How UTMB Qualification Works** | Plain-language explainer of the Running Stones and Index system, lottery mechanics, key dates, and alternative qualification paths. |

## Key Details

- **Target race:** CCC (100K) at UTMB Mont-Blanc, August 29, 2027
- **Home base:** Logan, Utah
- **Starting stones:** Zero
- **Lottery pre-registration:** December 2026
- **Lottery draw:** Mid-January 2027

## Running Stone Values

| Category | World Series Event | World Series Major |
|----------|-------------------|--------------------|
| 20K | 1 stone | 2 stones |
| 50K | 2 stones | 4 stones |
| 100K | 3 stones | 6 stones |
| 100M | 4 stones | 8 stones |

## Local Advantage

Two UTMB World Series events are within an hour of Logan:

- **Speedgoat Mountain Races** — Snowbird, UT (Jul 23-25, 2026)
- **Snowbasin by UTMB** — Ogden, UT (Sep 10-12, 2026) — brand new for 2026

## Regenerating the Workbook

The Python script that builds the Excel file is included. Requires `openpyxl`:

```bash
pip install openpyxl
python build_utmb.py
```

## Sources

- [UTMB World Series Sports System](https://utmb.world/sports-system)
- [UTMB Running Stones FAQ](https://help.utmb.world/running-stones)
- [UTMB World Series Events Calendar](https://utmb.world/utmb-world-series-events)
- [2025 Lottery Numbers (Electric Cable Car)](https://electriccablecar.com/utmb-2025-lottery-numbers-and-observations/)

Race dates, event details, and pricing are based on information available as of March 2026 and may change.
