#!/usr/bin/env python3
"""Submit ALL sitemap URLs to Google Indexing API."""

import xml.etree.ElementTree as ET
from google.oauth2 import service_account
from google.auth.transport.requests import AuthorizedSession

SERVICE_ACCOUNT_FILE = 'google-service-account.json'
SCOPES = ['https://www.googleapis.com/auth/indexing']
ENDPOINT = 'https://indexing.googleapis.com/v3/urlNotifications:publish'

credentials = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES)
session = AuthorizedSession(credentials)

# Parse sitemap for all URLs
tree = ET.parse('sitemap.xml')
root = tree.getroot()
ns = {'sm': 'http://www.sitemaps.org/schemas/sitemap/0.9'}
urls = [loc.text for loc in root.findall('.//sm:loc', ns)]

# Filter out the 35 already submitted
already = set([
    'https://builtrighthq.com/',
    'https://builtrighthq.com/tools/concrete-calculator',
    'https://builtrighthq.com/tools/paint-calculator',
    'https://builtrighthq.com/tools/square-footage-calculator',
    'https://builtrighthq.com/tools/roof-pitch-calculator',
    'https://builtrighthq.com/tools/drywall-calculator',
    'https://builtrighthq.com/tools/asphalt-calculator',
    'https://builtrighthq.com/tools/cubic-yard-calculator',
    'https://builtrighthq.com/tools/board-foot-calculator',
    'https://builtrighthq.com/tools/deck-calculator',
    'https://builtrighthq.com/tools/btu-calculator',
    'https://builtrighthq.com/tools/wire-size-calculator',
    'https://builtrighthq.com/tools/paver-calculator',
    'https://builtrighthq.com/tools/siding-calculator',
    'https://builtrighthq.com/tools/rebar-calculator',
    'https://builtrighthq.com/tools/plywood-calculator',
    'https://builtrighthq.com/tools/carpet-calculator',
    'https://builtrighthq.com/tools/brick-calculator',
    'https://builtrighthq.com/tools/flooring-calculator',
    'https://builtrighthq.com/tools/insulation-calculator',
    'https://builtrighthq.com/blog/cost-to-build-a-house',
    'https://builtrighthq.com/blog/how-to-start-a-landscaping-business',
    'https://builtrighthq.com/blog/how-to-start-a-plumbing-business',
    'https://builtrighthq.com/blog/how-to-start-a-roofing-company',
    'https://builtrighthq.com/blog/electrician-advertising',
    'https://builtrighthq.com/blog/contractor-tax-deductions',
    'https://builtrighthq.com/blog/how-to-start-an-electrical-business',
    'https://builtrighthq.com/blog/how-to-start-an-hvac-business',
    'https://builtrighthq.com/blog/bathroom-remodel-cost',
    'https://builtrighthq.com/blog/kitchen-remodel-cost',
    'https://builtrighthq.com/templates/',
    'https://builtrighthq.com/templates/contractor-invoice-template',
    'https://builtrighthq.com/blog/',
    'https://builtrighthq.com/tools/',
    'https://builtrighthq.com/licensing/contractor-license-by-state',
])

remaining = [u for u in urls if u not in already]
print(f"Submitting {len(remaining)} remaining URLs (of {len(urls)} total)...\n")

success = 0
errors = 0
quota_hit = False

for url in remaining:
    body = {'url': url, 'type': 'URL_UPDATED'}
    try:
        response = session.post(ENDPOINT, json=body)
        if response.status_code == 200:
            print(f'  ✅ {url}')
            success += 1
        elif response.status_code == 429:
            print(f'  ⚠️  Quota limit reached at {success} URLs')
            quota_hit = True
            break
        else:
            print(f'  ❌ {url}: {response.status_code}')
            errors += 1
    except Exception as e:
        print(f'  ❌ {url}: {str(e)[:80]}')
        errors += 1

print(f'\n🔥 Submitted {success} URLs, {errors} errors')
if quota_hit:
    print(f'⚠️  Daily quota hit. Remaining URLs will be submitted tomorrow.')
print(f'📊 Total submitted today: {success + 35} of {len(urls)}')
