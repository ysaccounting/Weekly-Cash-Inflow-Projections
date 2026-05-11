# Weekly Ticket Reconciliation

Combines Part 1 (StubHub) and Part 2 (Six Networks) into a single weekly report.

## Output

`Weekly_Reconciliation.xlsx` with two tabs:
- **StubHub Details** — Part 1: all Viagogo transactions with InInvoice flag (1/0)
- **Network Summary** — Part 2: Pay on Confirmation + Pay on Delivery summary

## Inputs each week

### Part 1 — StubHub
| Input | Description |
|---|---|
| StubHub portal screenshot | Enter Payment ID, Date, Proceeds, Charges, Credit manually |
| Viagogo CSV files | One per Payment ID, downloaded from the StubHub portal |
| StubHub Invoice Details (.xlsx) | TicketVault Invoice Details Report for StubHub — used to match TransactionIDs |

### Part 2 — Six Networks
| Input | Description |
|---|---|
| Invoice Details Report (.xlsx) | TicketVault — prior 7 invoice date days ending Sunday. Covers Gametime, TickPick, SeatGeek, Vivid Seats, GoTickets, TicketsNow |
| Invoices Report (.xlsx) | TicketVault — prior 7 event date days ending Sunday. Unpaid sales for Pay on Delivery balance |

## Payment groups
- **Pay on Confirmation:** Gametime, TickPick, SeatGeek
- **Pay on Delivery:** Vivid Seats, GoTickets, Ticketmaster (TicketsNow)

## Local setup

```bash
git clone <your-repo-url>
cd ticket-reconciliation

python -m venv venv
source venv/bin/activate      # Windows: venv\Scripts\activate
pip install -r requirements.txt

python app.py
# Open http://localhost:5000
```

## Deploy to Railway

1. Push this repo to GitHub
2. Go to [railway.app](https://railway.app) → New Project → Deploy from GitHub repo
3. Select this repo — Railway auto-detects the `Procfile` and deploys
4. Your app will be live at a `*.railway.app` URL

## Project structure

```
ticket-reconciliation/
├── app.py              # Flask web app (routes)
├── reconcile.py        # Core logic — Part 1 + Part 2
├── requirements.txt
├── Procfile
├── .gitignore
├── templates/
│   └── index.html
└── README.md
```
