# MATA Avenew - Jargas Gabungan Dashboard

An interactive dashboard for monitoring the Pembangunan Jaringan Gas Bumi untuk Rumah Tangga project across 5 packages and 14 locations, built for the MATA (Monitoring and Tracking Activities) system.

## Features

- Overview: combined KPIs, package progress summary, S-curve (planned vs actual)
- Per-Package Detail: S-curve chart, construction progress bars, SK sector table
- Safety and Manpower: safe hours, workforce breakdown per location
- Permits: permit status badges (Open / Proses / Closed) per location
- Interactive Navigation: sidebar navigation with smooth transitions
- Responsive Design: works on desktop and tablet

## Data Source

Data is sourced exclusively from:
Dashboard Gabungan Jargas Week 11.pptx - Direktorat Perencanaan dan Pembangunan Infrastruktur Migas, KESDM
Reporting period: 28 Januari - 04 Februari 2026 (Minggu ke-11)

## Tech Stack

- Structure: HTML5
- Styling: Vanilla CSS (glassmorphism design system)
- Logic: JavaScript ES6+
- Charts: Chart.js v4
- Server: npx http-server (local dev)

## Project Structure

`
MATA Dashboard/
├── index.html          # App shell
├── style.css           # Design system
├── app.js              # Data + chart logic
├── verify_data2.py     # Data consistency verification
└── README.md
`

## Getting Started

### Local Development

`ash
npx -y http-server . -p 3355 -o --cors
`

Then open http://localhost:3355

## Packages

| # | Package | Region | EPC | SR |
|---|---------|--------|-----|----|
| 1 | Jambi Region | Kota Jambi, Kab. TJB dan Pelalawan | KSO NOOREL-KSM | 22,972 |
| 2 | Sumatra + Jabar | Kab. MUBA, OKU Timur dan Indramayu | PT. Pratiwi Putri Sulung | 22,087 |
| 3 | Jateng + Sulsel | Kab. Batang, Kendal dan Wajo | KSO NOOREL-KSM | 22,213 |
| 4 | Jateng + Kaltara | Kab. Demak, Kota Bontang dan Tana Tidung | PT. Noorel Idea | 26,137 |
| 5 | Jatim + Kaltim | Kab. Gresik, Sidoarjo dan Samarinda | PT. Panca Indah Jayamahe | 21,855 |

Total: 115,264 SR | Contract Value: Rp 961.9 Billion

## License

MIT License
