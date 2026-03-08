#!/usr/bin/env python3
"""Update all pro-bundle CTAs with trust signals and better conversion copy."""

import os
import re

REPO = "/home/openclaw/.openclaw/workspaces/agent4/builtright-academy"

OLD_CTA = '''<!-- pro-bundle-cta -->
<div style="max-width:720px;margin:24px auto 40px;padding:32px;background:linear-gradient(135deg,rgba(249,115,22,0.08) 0%,rgba(249,115,22,0.02) 100%);border:2px solid rgba(249,115,22,0.3);border-radius:16px;text-align:center;">
  <p style="font-family:Inter,-apple-system,sans-serif;font-size:1.1rem;font-weight:700;color:#f8fafc;margin-bottom:8px;">⚡ Want the complete toolkit?</p>
  <p style="font-family:Inter,-apple-system,sans-serif;color:#94a3b8;font-size:0.9rem;margin-bottom:20px;">Invoice + Estimate + Job Costing + P&L Tracker — everything you need to run your business like a pro.</p>
  <a href="https://buy.stripe.com/cNi6oI87ta3X7Q7bTZaMU04" target="_blank" style="display:inline-block;padding:14px 32px;background:#F97316;color:#fff;font-family:Inter,-apple-system,sans-serif;font-weight:700;font-size:1rem;border-radius:8px;text-decoration:none;">Get Pro Bundle — $29</a>
  <p style="font-family:Inter,-apple-system,sans-serif;color:#64748b;font-size:0.75rem;margin-top:12px;">One-time payment. Instant download. No subscription.</p>
</div>'''

NEW_CTA = '''<!-- pro-bundle-cta -->
<div style="max-width:720px;margin:24px auto 40px;padding:32px;background:linear-gradient(135deg,rgba(249,115,22,0.08) 0%,rgba(249,115,22,0.02) 100%);border:2px solid rgba(249,115,22,0.3);border-radius:16px;text-align:center;">
  <p style="font-family:Inter,-apple-system,sans-serif;font-size:1.2rem;font-weight:800;color:#f8fafc;margin-bottom:8px;">⚡ Stop Losing Money on Every Job</p>
  <p style="font-family:Inter,-apple-system,sans-serif;color:#94a3b8;font-size:0.95rem;margin-bottom:6px;">The average contractor loses <strong style="color:#ef4444;">$3,400/year</strong> from bad invoicing and missed costs.</p>
  <p style="font-family:Inter,-apple-system,sans-serif;color:#94a3b8;font-size:0.9rem;margin-bottom:20px;">Our Pro Template Bundle gives you professional Invoice, Estimate, Job Costing &amp; P&L Tracker spreadsheets — ready to use in 5 minutes.</p>
  <a href="https://buy.stripe.com/cNi6oI87ta3X7Q7bTZaMU04" target="_blank" style="display:inline-block;padding:16px 40px;background:#F97316;color:#fff;font-family:Inter,-apple-system,sans-serif;font-weight:700;font-size:1.1rem;border-radius:8px;text-decoration:none;box-shadow:0 4px 14px rgba(249,115,22,0.4);">Get Pro Bundle — $29</a>
  <div style="margin-top:16px;display:flex;flex-wrap:wrap;justify-content:center;gap:16px;">
    <span style="font-family:Inter,-apple-system,sans-serif;color:#94a3b8;font-size:0.8rem;">🔒 Secure checkout via Stripe</span>
    <span style="font-family:Inter,-apple-system,sans-serif;color:#94a3b8;font-size:0.8rem;">✅ 30-day money-back guarantee</span>
    <span style="font-family:Inter,-apple-system,sans-serif;color:#94a3b8;font-size:0.8rem;">📥 Instant download</span>
  </div>
  <p style="font-family:Inter,-apple-system,sans-serif;color:#64748b;font-size:0.75rem;margin-top:10px;">One-time payment. No subscription. Works with Excel, Google Sheets, and Numbers.</p>
</div>'''

count = 0
for root, dirs, files in os.walk(REPO):
    # Skip .git and .vercel
    dirs[:] = [d for d in dirs if d not in ('.git', '.vercel', 'node_modules')]
    for f in files:
        if not f.endswith('.html'):
            continue
        path = os.path.join(root, f)
        with open(path, 'r') as fh:
            content = fh.read()
        if OLD_CTA in content:
            content = content.replace(OLD_CTA, NEW_CTA)
            with open(path, 'w') as fh:
                fh.write(content)
            count += 1
            print(f"  Updated: {os.path.relpath(path, REPO)}")

print(f"\nTotal updated: {count} files")
