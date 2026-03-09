#!/usr/bin/env python3
"""Add Google Analytics G-5R1VM2SZE1 to all HTML pages."""

import glob, os

GA_TAG = """<!-- Google Analytics -->
<script async src="https://www.googletagmanager.com/gtag/js?id=G-5R1VM2SZE1"></script>
<script>window.dataLayer=window.dataLayer||[];function gtag(){dataLayer.push(arguments);}gtag('js',new Date());gtag('config','G-5R1VM2SZE1');</script>"""

count = 0
for f in glob.glob('**/*.html', recursive=True):
    if 'node_modules' in f or '.next' in f:
        continue
    with open(f, 'r') as fh:
        content = fh.read()
    
    # Skip if already has this GA tag
    if 'G-5R1VM2SZE1' in content:
        continue
    
    # Insert after <head> tag
    if '<head>' in content:
        content = content.replace('<head>', '<head>\n' + GA_TAG, 1)
        with open(f, 'w') as fh:
            fh.write(content)
        count += 1

print(f"✅ Added GA to {count} pages")
