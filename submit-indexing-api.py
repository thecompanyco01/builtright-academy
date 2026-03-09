#!/usr/bin/env python3
"""Submit URLs to Google Indexing API for faster crawling."""

import json
from google.oauth2 import service_account
from googleapiclient.discovery import build
import httplib2

SERVICE_ACCOUNT_FILE = 'google-service-account.json'
SCOPES = ['https://www.googleapis.com/auth/indexing']
SITE = 'https://builtrighthq.com'

# Load credentials
credentials = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES)

# Build the service
http = httplib2.Http()
authed_http = google_auth_httplib2 = None

from google.auth.transport.requests import AuthorizedSession
session = AuthorizedSession(credentials)

# Top priority URLs to submit
urls = [
    '/',
    '/tools/concrete-calculator',
    '/tools/paint-calculator',
    '/tools/square-footage-calculator',
    '/tools/roof-pitch-calculator',
    '/tools/drywall-calculator',
    '/tools/asphalt-calculator',
    '/tools/cubic-yard-calculator',
    '/tools/board-foot-calculator',
    '/tools/deck-calculator',
    '/tools/btu-calculator',
    '/tools/wire-size-calculator',
    '/tools/paver-calculator',
    '/tools/siding-calculator',
    '/tools/rebar-calculator',
    '/tools/plywood-calculator',
    '/tools/carpet-calculator',
    '/tools/brick-calculator',
    '/tools/flooring-calculator',
    '/tools/insulation-calculator',
    '/blog/cost-to-build-a-house',
    '/blog/how-to-start-a-landscaping-business',
    '/blog/how-to-start-a-plumbing-business',
    '/blog/how-to-start-a-roofing-company',
    '/blog/electrician-advertising',
    '/blog/contractor-tax-deductions',
    '/blog/how-to-start-an-electrical-business',
    '/blog/how-to-start-an-hvac-business',
    '/blog/bathroom-remodel-cost',
    '/blog/kitchen-remodel-cost',
    '/templates/',
    '/templates/contractor-invoice-template',
    '/blog/',
    '/tools/',
    '/licensing/contractor-license-by-state',
]

ENDPOINT = 'https://indexing.googleapis.com/v3/urlNotifications:publish'

success = 0
errors = 0

for path in urls:
    url = SITE + path
    body = {
        'url': url,
        'type': 'URL_UPDATED'
    }
    try:
        response = session.post(ENDPOINT, json=body)
        if response.status_code == 200:
            print(f'  ✅ {url}')
            success += 1
        else:
            print(f'  ❌ {url}: {response.status_code} - {response.text[:100]}')
            errors += 1
    except Exception as e:
        print(f'  ❌ {url}: {str(e)[:100]}')
        errors += 1

print(f'\n🔥 Submitted {success} URLs, {errors} errors')
