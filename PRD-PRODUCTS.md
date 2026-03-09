# PRD: BuiltRight Products & Stripe Checkout

## Problem
We have payment links and product listings in Stripe, but:
1. No actual downloadable product files (templates, spreadsheets)
2. No proper checkout flow (server-side Stripe Checkout Sessions)
3. No secure download delivery after payment
4. No product preview/demo pages that sell

## Product Lineup

### Tier 1: Invoice Template Pro — $9
**Target:** Solo contractors who need one professional template NOW
**Contents:**
- Professional contractor invoice spreadsheet (Excel + Google Sheets compatible)
- Auto-calculating line items, subtotals, tax, and totals
- Sections: Company logo/info, client info, job description, itemized labor & materials, payment terms, notes
- Pre-formatted for printing
- Instructions sheet

### Tier 2: Pro Contractor Template Bundle — $29 (HERO PRODUCT)
**Target:** Contractors who want to run their business professionally
**Contents:**
- Everything in Tier 1, PLUS:
- **Estimate/Bid Template** — Professional estimate with itemized costs, markup %, profit calculation, terms & conditions
- **Job Costing Tracker** — Track actual vs estimated costs per job, materials, labor hours, overhead, profit/loss per job
- **Profit & Loss Tracker** — Monthly P&L statement, income categories, expense tracking, running totals, year summary
- Each template: Excel file, instructions, example data

### Tier 3: Complete Contractor Business Kit — $99
**Target:** Contractors starting or scaling a business
**Contents:**
- Everything in Tier 1 + Tier 2, PLUS:
- **Client Tracker/CRM Spreadsheet** — Track leads, clients, job status, follow-ups, revenue per client
- **Contractor Proposal Template** — Professional proposal with scope of work, timeline, pricing, terms
- **Change Order Template** — Document scope changes, cost adjustments, client approval
- **Weekly Timesheet Template** — Track employee hours, overtime, job allocation
- **Contractor Checklist Pack** — Job site safety, pre-job walkthrough, final inspection, punch list
- **Cash Flow Forecast Template** — 12-month cash flow projection, income vs expenses, cash runway

## Technical Architecture

### Checkout Flow
1. User clicks "Buy Now" button on any page
2. Button hits `/api/create-checkout` (Vercel serverless function)
3. Server creates Stripe Checkout Session with:
   - Product details
   - Success URL: `/thank-you?session_id={CHECKOUT_SESSION_ID}`
   - Cancel URL: back to product page
4. User redirected to Stripe Checkout (hosted by Stripe — no card data touches us)
5. On success → redirected to thank-you page
6. Thank-you page calls `/api/verify-session` to confirm payment
7. If paid → shows download links (time-limited signed URLs or direct links)

### File Delivery
Since we're on static hosting (Vercel), simplest secure approach:
- Store template files in a `/downloads/` directory on the server
- Vercel `vercel.json` blocks direct access to `/downloads/*`
- `/api/download` endpoint verifies payment via session_id, then serves the file
- OR: Use Stripe's built-in file delivery (after_completion redirect)

### Simpler Alternative (MVP — DO THIS FIRST):
- Create real Excel template files with actual useful content
- Host on a hidden URL path
- Stripe payment link → success_url includes a unique token
- Thank-you page shows download links
- Not perfectly secure but good enough for MVP (nobody's going to guess the URL)

### API Endpoints Needed
1. `POST /api/create-checkout` — Creates Stripe Checkout Session
   - Input: { product: "tier1" | "tier2" | "tier3" }
   - Output: { url: "https://checkout.stripe.com/..." }
   - Uses MACK_STRIPE_KEY (server-side only, NEVER exposed to client)

2. `GET /api/verify-session?session_id=xxx` — Verifies payment completed
   - Checks session status with Stripe
   - Returns download links if paid
   - Uses MACK_STRIPE_KEY (server-side only)

3. `GET /api/download?file=xxx&session_id=xxx` — Serves file after payment verification
   - Verifies session_id is paid
   - Serves the requested file
   - Uses MACK_STRIPE_KEY (server-side only)

### Security Rules (NON-NEGOTIABLE)
- MACK_STRIPE_KEY is set as a Vercel environment variable, accessed only via process.env
- NO API keys in any client-side code (.html, client .js)
- All Stripe operations happen in /api/ serverless functions
- Download links are only revealed after payment verification

## Template Files to Create

Each template should be a real, professional, USEFUL Excel spreadsheet:
- Professional formatting (branded header, clean layout, proper fonts)
- Working formulas (auto-calc totals, tax, markup, etc.)
- Example data pre-filled (so users see how to use it)
- Instructions tab
- Print-ready formatting
- Compatible with Excel, Google Sheets, and Numbers

### File Format
- Primary: .xlsx (Excel)
- Create using openpyxl Python library
- Each template gets its own .xlsx file
- Bundle tiers get a .zip containing all relevant files

## Pages to Build/Update

1. **Product landing pages** (update existing template pages):
   - Show preview screenshots of templates
   - Feature comparison table (Tier 1 vs 2 vs 3)
   - Social proof / trust signals
   - Clear pricing with "Buy Now" buttons
   - FAQ section

2. **Thank-you/download page** (`/templates/thank-you-download`):
   - Verify payment
   - Show download links for purchased tier
   - Upsell to higher tier

3. **Update all CTAs across the site** to use the new checkout flow

## Success Metrics
- Checkout completion rate > 5%
- Download rate after payment > 90%
- Zero API key leaks (audit before every deploy)

## Priority
1. Create the actual Excel template files (this is the product!)
2. Build the checkout API endpoints
3. Build the download/delivery flow
4. Update product pages
5. Test end-to-end: click buy → pay → download → open file → verify it works
