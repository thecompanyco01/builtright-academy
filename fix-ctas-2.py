#!/usr/bin/env python3
"""Second pass: fix remaining waitlist text and orphaned scripts."""

import os
import re
import glob

STRIPE_BUNDLE = "https://buy.stripe.com/cNi6oI87ta3X7Q7bTZaMU04"

changes = 0

for f in glob.glob('blog/*.html') + glob.glob('tools/*.html') + glob.glob('templates/*.html') + glob.glob('licensing/*.html') + ['index.html']:
    if not os.path.exists(f):
        continue
    with open(f, 'r', encoding='utf-8') as fh:
        content = fh.read()
    original = content

    # Fix "Join the Waitlist — Free" button text (already pointing to Stripe)
    content = content.replace('>Join the Waitlist — Free</a>', '>Get Pro Bundle — $29</a>')
    content = content.replace('>Join the Waitlist</a>', '>Get Pro Bundle — $29</a>')
    content = content.replace('>Download Free Toolkit</a>', '>Get Pro Bundle — $29</a>')
    
    # Fix surrounding cta-box context: replace "Free Contractor Toolkit" heading and description
    content = re.sub(
        r'<h3>📋 Free Contractor Toolkit</h3>\s*<p>[^<]*(?:<[^>]*>[^<]*)*</p>',
        '<h3>⚡ Pro Contractor Template Bundle — $29</h3>\n      <p>Professional Invoice, Estimate, Job Costing &amp; P&L Tracker spreadsheets — ready to use in 5 minutes.</p>',
        content
    )
    
    # Remove orphaned handleCapture functions (multi-line, various formats)
    # Pattern 1: function handleCapture(e) { ... return false; } (with newlines)
    content = re.sub(r'\s*function handleCapture\(e\)\s*\{[^}]*(?:\{[^}]*\}[^}]*)*\}[^}]*\}', '', content)
    # Catch remaining inline versions
    content = re.sub(r'function handleCapture\(e\)\{.*?return false;\}', '', content, flags=re.DOTALL)
    
    # Remove orphaned capture-form/capture-email references if the form was removed but script remains  
    # Remove empty script tags
    content = re.sub(r'<script>\s*</script>', '', content)
    
    # Remove "Join 500+ contractors" email capture text that might remain
    content = re.sub(r'<p[^>]*>Join 500\+ contractors getting[^<]*<strong[^>]*>free business templates</strong>[^<]*</p>', '', content)

    if content != original:
        with open(f, 'w', encoding='utf-8') as fh:
            fh.write(content)
        changes += 1
        print(f"  ✅ Fixed: {f}")

print(f"\n🔥 Second pass done! Fixed {changes} files.")
