FundedNext Dispute Document Generator
======================================

SETUP (one time only)
----------------------
1. Make sure Python 3.8+ is installed.
2. Open a terminal in this folder and run:

       pip install -r requirements.txt

HOW TO USE
----------
1. Edit config.json with your dispute details (customer name, ARN, amount, etc.)

2. Place files in the pdfs/ folder:

   REQUIRED (generate as text — put the real PDFs here):
     checkout_flow.pdf          — the Checkout Flow document
     refund_cancellation.pdf    — the Refund & Cancellation Policy document
     terms_of_service.pdf       — the Terms of Service document
     journal.pdf                — MetaTrader 5 Manager Journal export

   OPTIONAL (embed as images — can be .pdf or .png/.jpg):
     invoice.pdf or invoice.png              — FundedNext invoice
     intercom.png or intercom.pdf            — Intercom customer email screenshot
     gmail.png or gmail.pdf                  — Gmail confirmation email screenshot
     admin.png or admin.pdf                  — Admin panel screenshot
     drawdown_chart.png                      — Daily drawdown chart image

3. Run:

       python generate_docx.py

4. The output file will be saved in this folder:

       FN_Dispute_<ARN>_<date>.docx

WHAT GETS EMBEDDED AS TEXT (not images)
----------------------------------------
- Checkout Flow          → actual paragraphs, section headers, numbered steps
- Refund & Cancellation  → actual paragraphs, all sections and sub-clauses
- Terms of Service       → actual paragraphs, all 24 sections

WHAT GETS EMBEDDED AS HIGH-RESOLUTION IMAGES
---------------------------------------------
- Journal pages are rendered at 3× resolution for crisp, clear output
- Invoice, Intercom, Gmail, Admin images scale to full page width

NOTES
-----
- Journal MUST be a .pdf file. All pages are rendered at 3× scale (very sharp).
- Policy documents MUST be .pdf files so text can be extracted.
- Optional image zones accept .pdf (renders page 1) or .png/.jpg/.jpeg.
- If a file is missing the script prints [SKIP] and adds a placeholder.
