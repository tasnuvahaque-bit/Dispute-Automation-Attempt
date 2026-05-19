# FN–CFD Dispute Generator

Generates a fully formatted `.docx` dispute document entirely on GitHub — no local device required.

---

## How to use (GitHub only)

### Step 1 — Upload per-dispute files

Go to your repository on GitHub → `pdfs/` folder → **Add file → Upload files**

| File name | What it is | Required? |
|---|---|---|
| `journal.pdf` | MT5 Manager Journal export | Yes |
| `invoice.pdf` or `invoice.png` | FundedNext invoice | Optional |
| `intercom.png` or `intercom.pdf` | Intercom customer email screenshot | Optional |
| `gmail.png` or `gmail.pdf` | Gmail confirmation email screenshot | Optional |
| `admin.png` or `admin.pdf` | Admin panel screenshot | Optional |
| `drawdown_chart.png` | Daily drawdown chart | Optional |

> **The three policy documents are already in the repo and never need to be re-uploaded.**
> (`checkout_flow.pdf`, `refund_cancellation.pdf`, `terms_of_service.pdf`)

---

### Step 2 — Run the workflow

1. Click the **Actions** tab at the top of the repository
2. Click **Generate Dispute DOCX** in the left sidebar
3. Click **Run workflow** (top right)
4. Fill in the form:

| Field | Example |
|---|---|
| Customer Full Name | John Doe |
| Customer Email | john.doe@example.com |
| ARN / Chargeback Reference | 02308446110110167267 |
| Transaction Amount | 549.00 |
| Card Type | Visa |
| Card Number (masked) | 4111 **** **** 1111 |
| Purchase Date | 2026-03-15 |
| Account / Login ID | 123456 |
| Merchant Payment Reference | FN-2026-031500001 |
| Payment Descriptor | FUNDEDNEXT*CHALLENGE |
| Customer IP | 192.168.1.100 |
| IP Location | United Arab Emirates |
| Access Date & Time | 2026-03-15 09:42:11 UTC |
| Customer Device | Windows 10 — Chrome 123 |
| Device ID | d4e5f6a7b8c9 |

5. Click **Run workflow** (green button)

---

### Step 3 — Download the DOCX

1. Wait ~60 seconds for the workflow to finish (green checkmark ✓)
2. Click on the completed run
3. Scroll to **Artifacts** at the bottom
4. Download **FN-Dispute-\<ARN\>-\<date\>**

The zip contains your `FN_Dispute_....docx` file, ready to submit.

---

## What's in the document

| Section | Format |
|---|---|
| Checkout Flow | **Real text** — searchable, copyable |
| Refund & Cancellation Policy | **Real text** — searchable, copyable |
| Terms of Service | **Real text** — searchable, copyable |
| Invoice | High-resolution image |
| Customer Communication (Intercom) | High-resolution image |
| Proof of Service Delivery (Gmail) | High-resolution image |
| Customer Purchase Details + Admin | Text + high-resolution image |
| Daily Drawdown | Text + chart image |
| Customer IP & Access Log | Text |
| Trading History / Journal | High-resolution PDF pages (3× scale) |

---

## Repository structure

```
FN--CFD-Dispute/
  .github/
    workflows/
      generate_docx.yml     ← the workflow
  pdfs/
    checkout_flow.pdf       ← permanent, never changes
    refund_cancellation.pdf ← permanent, never changes
    terms_of_service.pdf    ← permanent, never changes
    journal.pdf             ← upload per dispute
    invoice.pdf / .png      ← upload per dispute (optional)
    intercom.png / .pdf     ← upload per dispute (optional)
    gmail.png / .pdf        ← upload per dispute (optional)
    admin.png / .pdf        ← upload per dispute (optional)
    drawdown_chart.png      ← upload per dispute (optional)
  generate_docx.py
  requirements.txt
  README.md
```

---

## Running locally (optional)

```bash
pip install -r requirements.txt
python generate_docx.py     # reads from config.json if present
```
