#!/bin/bash
# Add email capture CTA to all calculator/tool pages
# Injects a lead capture form before the closing </body> tag

TOOLS_DIR="tools"
cd "$(dirname "$0")"

EMAIL_CTA='
<!-- Email Capture CTA -->
<div id="email-cta" style="max-width:720px;margin:40px auto;padding:40px 32px;background:linear-gradient(135deg,#1e293b 0%,#0f172a 100%);border:1px solid rgba(249,115,22,0.3);border-radius:16px;text-align:center;">
  <div style="font-size:2rem;margin-bottom:12px;">📋</div>
  <h3 style="font-family:Inter,-apple-system,sans-serif;font-size:1.4rem;font-weight:700;color:#f8fafc;margin-bottom:8px;">Free Contractor Toolkit</h3>
  <p style="font-family:Inter,-apple-system,sans-serif;color:#94a3b8;font-size:0.95rem;margin-bottom:24px;max-width:480px;margin-left:auto;margin-right:auto;">Get our <strong style="color:#f97316;">free spreadsheet templates</strong> — invoice, estimate, job costing &amp; profit tracker. Used by 500+ contractors.</p>
  <form id="capture-form" onsubmit="return handleCapture(event)" style="display:flex;gap:10px;max-width:420px;margin:0 auto;flex-wrap:wrap;justify-content:center;">
    <input type="email" id="capture-email" placeholder="you@company.com" required style="flex:1;min-width:220px;padding:12px 16px;border-radius:8px;border:1px solid rgba(255,255,255,0.1);background:rgba(255,255,255,0.05);color:#f8fafc;font-size:0.95rem;font-family:Inter,-apple-system,sans-serif;outline:none;" onfocus="this.style.borderColor=&apos;#f97316&apos;" onblur="this.style.borderColor=&apos;rgba(255,255,255,0.1)&apos;">
    <button type="submit" style="padding:12px 24px;background:#f97316;color:#fff;border:none;border-radius:8px;font-weight:600;font-size:0.95rem;cursor:pointer;font-family:Inter,-apple-system,sans-serif;white-space:nowrap;" onmouseover="this.style.background=&apos;#ea580c&apos;" onmouseout="this.style.background=&apos;#f97316&apos;">Send Me the Templates</button>
  </form>
  <p id="capture-msg" style="font-family:Inter,-apple-system,sans-serif;color:#22c55e;font-size:0.85rem;margin-top:12px;display:none;"></p>
  <p style="font-family:Inter,-apple-system,sans-serif;color:#64748b;font-size:0.75rem;margin-top:12px;">No spam. Unsubscribe anytime.</p>
</div>
<script>
function handleCapture(e){
  e.preventDefault();
  var em=document.getElementById("capture-email").value;
  var msg=document.getElementById("capture-msg");
  var src=window.location.pathname;
  fetch("/api/subscribe",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({email:em,source:src,lead_magnet:"contractor-toolkit"})}).then(function(r){return r.json()}).then(function(d){msg.textContent="✅ Check your inbox! Templates are on the way.";msg.style.display="block";document.getElementById("capture-form").style.display="none"}).catch(function(){msg.textContent="Something went wrong. Try again.";msg.style.color="#ef4444";msg.style.display="block"});
  return false;
}
</script>
'

count=0
for file in $TOOLS_DIR/*.html; do
  filename=$(basename "$file")
  # Skip index
  if [ "$filename" = "index.html" ]; then continue; fi
  
  # Check if already has email capture
  if grep -q "email-cta" "$file"; then
    echo "SKIP (already has CTA): $filename"
    continue
  fi
  
  # Insert before the Vercel analytics script (near end of body)
  if grep -q "_vercel/insights" "$file"; then
    sed -i "s|<script defer src=\"/_vercel/insights/script.js\"></script>|${EMAIL_CTA//$'\n'/\\n}\n  <script defer src=\"/_vercel/insights/script.js\"></script>|" "$file"
  else
    # Fallback: insert before </body>
    sed -i "s|</body>|${EMAIL_CTA//$'\n'/\\n}\n</body>|" "$file"
  fi
  
  count=$((count + 1))
  echo "ADDED: $filename"
done

echo ""
echo "Done. Added email capture to $count calculator pages."
