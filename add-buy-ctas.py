#!/usr/bin/env python3
"""Add buy CTA to all HTML pages missing Stripe buy links."""
import os, glob

BUY_LINK_29 = "https://buy.stripe.com/3cI14pfgy88SddhdMK9R605"
BUY_LINK_9 = "https://buy.stripe.com/5kQ00lc4m0Gq7SXfUS9R604"

CTA_HTML = f'''
<!-- Sticky CTA Banner -->
<div style="position:fixed;bottom:0;left:0;right:0;background:linear-gradient(135deg,#1E293B,#0F172A);border-top:2px solid #F97316;padding:12px 20px;z-index:1000;display:flex;align-items:center;justify-content:center;gap:16px;flex-wrap:wrap;">
  <span style="color:#F8FAFC;font-size:0.95rem;font-weight:600;">📋 Need professional contractor templates?</span>
  <a href="{BUY_LINK_9}" style="background:#F97316;color:#fff;padding:8px 20px;border-radius:8px;font-size:0.9rem;font-weight:700;text-decoration:none;white-space:nowrap;">Invoice Template — $9</a>
  <a href="{BUY_LINK_29}" style="background:transparent;color:#F97316;padding:8px 20px;border-radius:8px;font-size:0.9rem;font-weight:600;text-decoration:none;border:1.5px solid #F97316;white-space:nowrap;">Pro Bundle — $29</a>
</div>
'''

# Skip these directories/files
SKIP = {'node_modules', '.git', 'download.html', 'embed', 'assets'}

count = 0
for html_file in glob.glob('**/*.html', recursive=True):
    # Skip excluded paths
    parts = html_file.split('/')
    if any(s in parts for s in SKIP):
        continue
    if 'embed/' in html_file:
        continue
    
    with open(html_file, 'r') as f:
        content = f.read()
    
    # Skip if already has buy links
    if 'buy.stripe.com' in content:
        continue
    
    # Skip if no </body> tag
    if '</body>' not in content:
        continue
    
    # Add CTA before </body>
    content = content.replace('</body>', CTA_HTML + '\n</body>')
    
    with open(html_file, 'w') as f:
        f.write(content)
    count += 1

print(f'Added buy CTAs to {count} pages')
