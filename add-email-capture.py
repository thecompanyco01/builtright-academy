#!/usr/bin/env python3
"""Add email capture CTA to all calculator/tool pages."""
import os, glob

EMAIL_CTA = '''
<!-- Email Capture CTA -->
<div id="email-cta" style="max-width:720px;margin:40px auto;padding:40px 32px;background:linear-gradient(135deg,#1e293b 0%,#0f172a 100%);border:1px solid rgba(249,115,22,0.3);border-radius:16px;text-align:center;">
  <div style="font-size:2rem;margin-bottom:12px;">📋</div>
  <h3 style="font-family:Inter,-apple-system,sans-serif;font-size:1.4rem;font-weight:700;color:#f8fafc;margin-bottom:8px;">Free Contractor Toolkit</h3>
  <p style="font-family:Inter,-apple-system,sans-serif;color:#94a3b8;font-size:0.95rem;margin-bottom:24px;max-width:480px;margin-left:auto;margin-right:auto;">Get our <strong style="color:#f97316;">free spreadsheet templates</strong> &mdash; invoice, estimate, job costing &amp; profit tracker. Used by 500+ contractors.</p>
  <form id="capture-form" onsubmit="return handleCapture(event)" style="display:flex;gap:10px;max-width:420px;margin:0 auto;flex-wrap:wrap;justify-content:center;">
    <input type="email" id="capture-email" placeholder="you@company.com" required style="flex:1;min-width:220px;padding:12px 16px;border-radius:8px;border:1px solid rgba(255,255,255,0.1);background:rgba(255,255,255,0.05);color:#f8fafc;font-size:0.95rem;font-family:Inter,-apple-system,sans-serif;outline:none;">
    <button type="submit" style="padding:12px 24px;background:#f97316;color:#fff;border:none;border-radius:8px;font-weight:600;font-size:0.95rem;cursor:pointer;font-family:Inter,-apple-system,sans-serif;white-space:nowrap;">Send Me the Templates</button>
  </form>
  <p id="capture-msg" style="font-family:Inter,-apple-system,sans-serif;color:#22c55e;font-size:0.85rem;margin-top:12px;display:none;"></p>
  <p style="font-family:Inter,-apple-system,sans-serif;color:#64748b;font-size:0.75rem;margin-top:12px;">No spam. Unsubscribe anytime.</p>
</div>
<script>
function handleCapture(e){
  e.preventDefault();
  var em=document.getElementById("capture-email").value;
  var msg=document.getElementById("capture-msg");
  var btn=e.target.querySelector("button");
  btn.textContent="Sending...";btn.disabled=true;
  fetch("/api/subscribe",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({email:em,source:window.location.pathname,lead_magnet:"contractor-toolkit"})})
  .then(function(r){return r.json()})
  .then(function(d){msg.textContent="\\u2705 Check your inbox! Templates are on the way.";msg.style.display="block";document.getElementById("capture-form").style.display="none"})
  .catch(function(){msg.textContent="Something went wrong. Try again.";msg.style.color="#ef4444";msg.style.display="block";btn.textContent="Send Me the Templates";btn.disabled=false});
  return false;
}
</script>
'''

# Also add to blog articles
BLOG_CTA = '''
<!-- Email Capture CTA -->
<div id="email-cta" style="max-width:720px;margin:40px auto;padding:40px 32px;background:linear-gradient(135deg,#1e293b 0%,#0f172a 100%);border:1px solid rgba(249,115,22,0.3);border-radius:16px;text-align:center;">
  <div style="font-size:2rem;margin-bottom:12px;">🎯</div>
  <h3 style="font-family:Inter,-apple-system,sans-serif;font-size:1.4rem;font-weight:700;color:#f8fafc;margin-bottom:8px;">Grow Your Trade Business</h3>
  <p style="font-family:Inter,-apple-system,sans-serif;color:#94a3b8;font-size:0.95rem;margin-bottom:24px;max-width:480px;margin-left:auto;margin-right:auto;">Join 500+ contractors getting <strong style="color:#f97316;">free business templates</strong> and weekly tips to win more jobs and boost profits.</p>
  <form id="capture-form" onsubmit="return handleCapture(event)" style="display:flex;gap:10px;max-width:420px;margin:0 auto;flex-wrap:wrap;justify-content:center;">
    <input type="email" id="capture-email" placeholder="you@company.com" required style="flex:1;min-width:220px;padding:12px 16px;border-radius:8px;border:1px solid rgba(255,255,255,0.1);background:rgba(255,255,255,0.05);color:#f8fafc;font-size:0.95rem;font-family:Inter,-apple-system,sans-serif;outline:none;">
    <button type="submit" style="padding:12px 24px;background:#f97316;color:#fff;border:none;border-radius:8px;font-weight:600;font-size:0.95rem;cursor:pointer;font-family:Inter,-apple-system,sans-serif;white-space:nowrap;">Get Free Templates</button>
  </form>
  <p id="capture-msg" style="font-family:Inter,-apple-system,sans-serif;color:#22c55e;font-size:0.85rem;margin-top:12px;display:none;"></p>
  <p style="font-family:Inter,-apple-system,sans-serif;color:#64748b;font-size:0.75rem;margin-top:12px;">No spam. Unsubscribe anytime.</p>
</div>
<script>
function handleCapture(e){
  e.preventDefault();
  var em=document.getElementById("capture-email").value;
  var msg=document.getElementById("capture-msg");
  var btn=e.target.querySelector("button");
  btn.textContent="Sending...";btn.disabled=true;
  fetch("/api/subscribe",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({email:em,source:window.location.pathname,lead_magnet:"contractor-toolkit"})})
  .then(function(r){return r.json()})
  .then(function(d){msg.textContent="\\u2705 Check your inbox! Templates are on the way.";msg.style.display="block";document.getElementById("capture-form").style.display="none"})
  .catch(function(){msg.textContent="Something went wrong. Try again.";msg.style.color="#ef4444";msg.style.display="block";btn.textContent="Get Free Templates";btn.disabled=false});
  return false;
}
</script>
'''

def add_cta(filepath, cta_html):
    with open(filepath, 'r') as f:
        content = f.read()
    
    if 'email-cta' in content:
        return False  # Already has CTA
    
    # Insert before Vercel analytics or before </body>
    if '/_vercel/insights/script.js' in content:
        content = content.replace(
            '<script defer src="/_vercel/insights/script.js"></script>',
            cta_html + '\n  <script defer src="/_vercel/insights/script.js"></script>'
        )
    elif '</body>' in content:
        content = content.replace('</body>', cta_html + '\n</body>')
    else:
        return False
    
    with open(filepath, 'w') as f:
        f.write(content)
    return True

# Process calculator pages
tools_count = 0
for f in sorted(glob.glob('tools/*.html')):
    if os.path.basename(f) == 'index.html':
        continue
    if add_cta(f, EMAIL_CTA):
        tools_count += 1
        print(f'  ✅ {os.path.basename(f)}')
    else:
        print(f'  ⏭️  {os.path.basename(f)} (already has CTA)')

# Process blog articles
blog_count = 0
for f in sorted(glob.glob('blog/*.html')):
    if os.path.basename(f) == 'index.html':
        continue
    if add_cta(f, BLOG_CTA):
        blog_count += 1
        print(f'  ✅ {os.path.basename(f)}')
    else:
        print(f'  ⏭️  {os.path.basename(f)} (already has CTA)')

# Process template pages  
template_count = 0
for f in sorted(glob.glob('templates/*.html')):
    if os.path.basename(f) == 'index.html':
        continue
    if add_cta(f, EMAIL_CTA):
        template_count += 1
        print(f'  ✅ {os.path.basename(f)}')
    else:
        print(f'  ⏭️  {os.path.basename(f)} (already has CTA)')

print(f'\nDone! Added email capture to {tools_count} tools, {blog_count} blog posts, {template_count} templates.')
print(f'Total: {tools_count + blog_count + template_count} pages now have lead capture.')
