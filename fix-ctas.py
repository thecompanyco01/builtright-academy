#!/usr/bin/env python3
"""Replace all waitlist/email-capture CTAs with direct Stripe buy buttons across the entire site."""

import os
import re
import glob

STRIPE_BUNDLE = "https://buy.stripe.com/cNi6oI87ta3X7Q7bTZaMU04"
STRIPE_KIT = "https://buy.stripe.com/fZu7sMafB8ZT6M3bTZaMU05"
STRIPE_SINGLE = "https://buy.stripe.com/fZu4gA3Rddg95HZ4rxaMU06"

# The new buy CTA block to replace email captures and waitlist forms
BUY_CTA_BLOCK = f'''<!-- buy-cta -->
<div style="max-width:720px;margin:24px auto 40px;padding:32px;background:linear-gradient(135deg,rgba(249,115,22,0.08) 0%,rgba(249,115,22,0.02) 100%);border:2px solid rgba(249,115,22,0.3);border-radius:16px;text-align:center;">
  <p style="font-family:Inter,-apple-system,sans-serif;font-size:1.2rem;font-weight:800;color:#f8fafc;margin-bottom:8px;">⚡ Stop Losing Money on Every Job</p>
  <p style="font-family:Inter,-apple-system,sans-serif;color:#94a3b8;font-size:0.95rem;margin-bottom:6px;">The average contractor loses <strong style="color:#ef4444;">$3,400/year</strong> from bad invoicing and missed costs.</p>
  <p style="font-family:Inter,-apple-system,sans-serif;color:#94a3b8;font-size:0.9rem;margin-bottom:20px;">Our Pro Template Bundle gives you professional Invoice, Estimate, Job Costing &amp; P&L Tracker spreadsheets — ready to use in 5 minutes.</p>
  <a href="{STRIPE_BUNDLE}" target="_blank" style="display:inline-block;padding:16px 40px;background:#F97316;color:#fff;font-family:Inter,-apple-system,sans-serif;font-weight:700;font-size:1.1rem;border-radius:8px;text-decoration:none;box-shadow:0 4px 14px rgba(249,115,22,0.4);">Get Pro Bundle — $29</a>
  <div style="margin-top:12px;"><a href="{STRIPE_KIT}" target="_blank" style="font-family:Inter,-apple-system,sans-serif;color:#3B82F6;font-size:0.9rem;text-decoration:underline;">Or get the Complete Business Kit for $99 →</a></div>
  <div style="margin-top:16px;display:flex;flex-wrap:wrap;justify-content:center;gap:16px;">
    <span style="font-family:Inter,-apple-system,sans-serif;color:#94a3b8;font-size:0.8rem;">🔒 Secure checkout via Stripe</span>
    <span style="font-family:Inter,-apple-system,sans-serif;color:#94a3b8;font-size:0.8rem;">✅ 30-day money-back guarantee</span>
    <span style="font-family:Inter,-apple-system,sans-serif;color:#94a3b8;font-size:0.8rem;">📥 Instant download</span>
  </div>
  <p style="font-family:Inter,-apple-system,sans-serif;color:#64748b;font-size:0.75rem;margin-top:10px;">One-time payment. No subscription. Works with Excel, Google Sheets, and Numbers.</p>
</div>'''

# Smaller inline CTA for mid-article placement
INLINE_BUY_CTA = f'''<div class="cta-box" style="text-align:center;">
      <h3>⚡ Pro Contractor Template Bundle — $29</h3>
      <p>Professional Invoice, Estimate, Job Costing &amp; P&L Tracker spreadsheets. Ready to use in 5 minutes. Used by 500+ contractors.</p>
      <a href="{STRIPE_BUNDLE}" target="_blank" class="cta-btn" style="display:inline-block;padding:14px 32px;background:#F97316;color:#fff;font-weight:700;border-radius:8px;text-decoration:none;">Buy Now — $29</a>
      <p style="margin-top:8px;font-size:0.85rem;color:#64748b;">🔒 Secure Stripe checkout · 30-day money-back guarantee · Instant download</p>
    </div>'''

changes = 0

def process_file(filepath):
    global changes
    with open(filepath, 'r', encoding='utf-8') as f:
        content = f.read()
    
    original = content
    
    # 1. Replace mid-article cta-box divs that link to waitlistForm
    # Pattern: <div class="cta-box">...<a href="/#waitlistForm"...>...</a>...</div>
    pattern_cta_box = r'<div class="cta-box">\s*<h3>[^<]*</h3>\s*<p>[^<]*(?:<[^>]*>[^<]*)*</p>\s*<a href="/#waitlistForm[^"]*"[^>]*>[^<]*</a>\s*</div>'
    content = re.sub(pattern_cta_box, INLINE_BUY_CTA, content, flags=re.DOTALL)
    
    # 2. Replace the email-cta capture blocks (the big email form at bottom of articles/tools)
    # Match from <div id="email-cta" to the closing </div> before <!-- pro-bundle-cta -->
    pattern_email_cta = r'<div id="email-cta"[^>]*>.*?</div>\s*\n'
    # More precise: match the entire email-cta block including nested divs
    # Use a simpler approach: find and remove the block
    if 'id="email-cta"' in content:
        # Find start
        start = content.find('<div id="email-cta"')
        if start != -1:
            # Count divs to find matching close
            depth = 0
            i = start
            while i < len(content):
                if content[i:i+4] == '<div':
                    depth += 1
                elif content[i:i+6] == '</div>':
                    depth -= 1
                    if depth == 0:
                        end = i + 6
                        # Remove this block (it's redundant if pro-bundle-cta exists after it)
                        # Check if pro-bundle-cta already follows
                        if '<!-- pro-bundle-cta -->' in content[end:end+200] or '<!-- buy-cta -->' in content[end:end+200]:
                            # Just remove the email capture, keep the buy CTA
                            content = content[:start] + content[end:]
                        else:
                            # Replace email capture with buy CTA
                            content = content[:start] + BUY_CTA_BLOCK + content[end:]
                        break
                i += 1
    
    # 3. Replace any remaining handleCapture script blocks (orphaned after removing the form)
    content = re.sub(r'\n<script>\s*function handleCapture\(e\)\{.*?\}\s*</script>', '', content, flags=re.DOTALL)
    
    # 4. On homepage/templates: replace waitlist-form sections with buy CTAs
    # Pattern: <form class="waitlist-form"...>...</form>
    if 'waitlist-form' in content or 'waitlistForm' in content:
        # Replace the bottom-cta section on homepage
        pattern_bottom_cta = r'<section class="bottom-cta">.*?</section>'
        replacement_bottom = f'''<section class="bottom-cta">
  <div class="container" style="text-align:center;">
    <h2>Stop guessing. Start building right.</h2>
    <p>Professional templates used by 500+ contractors to win more jobs and boost profits.</p>
    <div style="display:flex;gap:16px;justify-content:center;flex-wrap:wrap;margin-top:24px;">
      <a href="{STRIPE_BUNDLE}" target="_blank" style="display:inline-block;padding:16px 40px;background:#F97316;color:#fff;font-family:Inter,-apple-system,sans-serif;font-weight:700;font-size:1.1rem;border-radius:8px;text-decoration:none;box-shadow:0 4px 14px rgba(249,115,22,0.4);">Get Pro Bundle — $29</a>
      <a href="{STRIPE_KIT}" target="_blank" style="display:inline-block;padding:16px 40px;background:transparent;color:#F97316;font-family:Inter,-apple-system,sans-serif;font-weight:700;font-size:1.1rem;border-radius:8px;text-decoration:none;border:2px solid #F97316;">Complete Business Kit — $99</a>
    </div>
    <div style="margin-top:16px;display:flex;flex-wrap:wrap;justify-content:center;gap:16px;">
      <span style="font-family:Inter,-apple-system,sans-serif;color:#94a3b8;font-size:0.8rem;">🔒 Secure Stripe checkout</span>
      <span style="font-family:Inter,-apple-system,sans-serif;color:#94a3b8;font-size:0.8rem;">✅ 30-day money-back guarantee</span>
      <span style="font-family:Inter,-apple-system,sans-serif;color:#94a3b8;font-size:0.8rem;">📥 Instant download</span>
    </div>
  </div>
</section>'''
        content = re.sub(pattern_bottom_cta, replacement_bottom, content, flags=re.DOTALL)
        
        # Replace any remaining standalone waitlist forms
        pattern_waitlist_form = r'<form class="waitlist-form"[^>]*>.*?</form>'
        buy_btn = f'<a href="{STRIPE_BUNDLE}" target="_blank" style="display:inline-block;padding:16px 40px;background:#F97316;color:#fff;font-family:Inter,-apple-system,sans-serif;font-weight:700;font-size:1.1rem;border-radius:8px;text-decoration:none;box-shadow:0 4px 14px rgba(249,115,22,0.4);">Get Pro Bundle — $29</a>'
        content = re.sub(pattern_waitlist_form, buy_btn, content, flags=re.DOTALL)
    
    # 5. Replace any remaining /#waitlistForm links
    content = content.replace('href="/#waitlistForm"', f'href="{STRIPE_BUNDLE}" target="_blank"')
    content = content.replace("href='/#waitlistForm'", f'href="{STRIPE_BUNDLE}" target="_blank"')
    
    # 6. Remove form-success divs (no longer needed)
    content = re.sub(r'<div class="form-success"[^>]*>.*?</div>', '', content, flags=re.DOTALL)
    
    # 7. Clean up waitlist JS handlers
    content = re.sub(r"\s*handleSubmit\('waitlistForm[^']*',\s*'formSuccess[^']*'\);?\s*", '', content)
    
    if content != original:
        with open(filepath, 'w', encoding='utf-8') as f:
            f.write(content)
        changes += 1
        print(f"  ✅ Fixed: {filepath}")
    else:
        pass  # No changes needed

# Process all HTML files
dirs = ['blog', 'tools', 'templates', 'licensing', '.']
for d in dirs:
    for f in glob.glob(os.path.join(d, '*.html')):
        process_file(f)

print(f"\n🔥 Done! Fixed {changes} files.")
