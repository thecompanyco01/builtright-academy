#!/usr/bin/env python3
"""Generate contractor licensing guide pages for all 50 US states."""

import os
import json
from datetime import datetime

# State data with real licensing info
STATES = {
    "alabama": {"name": "Alabama", "abbr": "AL", "board": "Alabama Licensing Board for General Contractors", "license_required": True, "threshold": "$50,000", "types": ["General Contractor", "Subcontractor"], "reciprocity": ["Mississippi", "Louisiana"], "exam_required": True, "insurance_required": True, "bond_required": True, "bond_amount": "$10,000-$100,000", "renewal": "Annual", "ce_hours": "0", "fee_range": "$200-$400"},
    "alaska": {"name": "Alaska", "abbr": "AK", "board": "Alaska Division of Corporations, Business and Professional Licensing", "license_required": True, "threshold": "$0", "types": ["General Contractor", "Specialty Contractor", "Mechanical Contractor"], "reciprocity": [], "exam_required": False, "insurance_required": True, "bond_required": False, "bond_amount": "N/A", "renewal": "Biennial", "ce_hours": "0", "fee_range": "$150-$300"},
    "arizona": {"name": "Arizona", "abbr": "AZ", "board": "Arizona Registrar of Contractors (AZ ROC)", "license_required": True, "threshold": "$1,000", "types": ["General Residential (B)", "General Commercial (B-1)", "Specialty Contractor"], "reciprocity": [], "exam_required": True, "insurance_required": True, "bond_required": True, "bond_amount": "$2,500-$7,500", "renewal": "Biennial", "ce_hours": "0", "fee_range": "$300-$600"},
    "arkansas": {"name": "Arkansas", "abbr": "AR", "board": "Arkansas Contractors Licensing Board", "license_required": True, "threshold": "$20,000", "types": ["General Contractor", "Specialty Contractor"], "reciprocity": [], "exam_required": True, "insurance_required": True, "bond_required": True, "bond_amount": "$10,000", "renewal": "Annual", "ce_hours": "6", "fee_range": "$100-$250"},
    "california": {"name": "California", "abbr": "CA", "board": "Contractors State License Board (CSLB)", "license_required": True, "threshold": "$500", "types": ["Class A (General Engineering)", "Class B (General Building)", "Class C (Specialty)"], "reciprocity": [], "exam_required": True, "insurance_required": True, "bond_required": True, "bond_amount": "$25,000", "renewal": "Biennial", "ce_hours": "0", "fee_range": "$450-$700"},
    "colorado": {"name": "Colorado", "abbr": "CO", "board": "No statewide license — local jurisdiction", "license_required": False, "threshold": "N/A", "types": ["Varies by city/county"], "reciprocity": [], "exam_required": False, "insurance_required": True, "bond_required": False, "bond_amount": "N/A", "renewal": "Varies", "ce_hours": "Varies", "fee_range": "$50-$300"},
    "connecticut": {"name": "Connecticut", "abbr": "CT", "board": "Connecticut Department of Consumer Protection", "license_required": True, "threshold": "$200", "types": ["New Home Construction", "Home Improvement"], "reciprocity": [], "exam_required": True, "insurance_required": True, "bond_required": False, "bond_amount": "N/A", "renewal": "Annual", "ce_hours": "0", "fee_range": "$220-$390"},
    "delaware": {"name": "Delaware", "abbr": "DE", "board": "Delaware Division of Revenue", "license_required": True, "threshold": "$0", "types": ["General Contractor"], "reciprocity": [], "exam_required": False, "insurance_required": True, "bond_required": False, "bond_amount": "N/A", "renewal": "Annual", "ce_hours": "0", "fee_range": "$75-$200"},
    "florida": {"name": "Florida", "abbr": "FL", "board": "Florida Construction Industry Licensing Board (CILB)", "license_required": True, "threshold": "$0", "types": ["Certified General Contractor (CGC)", "Certified Building Contractor (CBC)", "Registered Contractor"], "reciprocity": [], "exam_required": True, "insurance_required": True, "bond_required": False, "bond_amount": "N/A", "renewal": "Biennial", "ce_hours": "14", "fee_range": "$249-$400"},
    "georgia": {"name": "Georgia", "abbr": "GA", "board": "Georgia Secretary of State — Contractor Division", "license_required": True, "threshold": "$2,500", "types": ["General Contractor", "Residential/Light Commercial", "Conditioned Air Contractor"], "reciprocity": [], "exam_required": True, "insurance_required": True, "bond_required": False, "bond_amount": "N/A", "renewal": "Biennial", "ce_hours": "3", "fee_range": "$100-$300"},
    "hawaii": {"name": "Hawaii", "abbr": "HI", "board": "Hawaii Contractors License Board", "license_required": True, "threshold": "$0", "types": ["General Engineering (A)", "General Building (B)", "Specialty (C)"], "reciprocity": [], "exam_required": True, "insurance_required": True, "bond_required": True, "bond_amount": "$5,000", "renewal": "Biennial", "ce_hours": "0", "fee_range": "$300-$600"},
    "idaho": {"name": "Idaho", "abbr": "ID", "board": "Idaho Division of Building Safety", "license_required": True, "threshold": "$0", "types": ["General Contractor", "Public Works Contractor"], "reciprocity": [], "exam_required": False, "insurance_required": True, "bond_required": False, "bond_amount": "N/A", "renewal": "Annual", "ce_hours": "0", "fee_range": "$0 (registration fee)"},
    "illinois": {"name": "Illinois", "abbr": "IL", "board": "No statewide license — local jurisdiction", "license_required": False, "threshold": "N/A", "types": ["Varies by city/county"], "reciprocity": [], "exam_required": False, "insurance_required": True, "bond_required": False, "bond_amount": "N/A", "renewal": "Varies", "ce_hours": "Varies", "fee_range": "Varies"},
    "indiana": {"name": "Indiana", "abbr": "IN", "board": "No statewide license — local jurisdiction", "license_required": False, "threshold": "N/A", "types": ["Varies by city/county"], "reciprocity": [], "exam_required": False, "insurance_required": True, "bond_required": False, "bond_amount": "N/A", "renewal": "Varies", "ce_hours": "Varies", "fee_range": "Varies"},
    "iowa": {"name": "Iowa", "abbr": "IA", "board": "Iowa Division of Labor", "license_required": True, "threshold": "$2,000", "types": ["General Contractor", "Residential Contractor"], "reciprocity": [], "exam_required": False, "insurance_required": True, "bond_required": False, "bond_amount": "N/A", "renewal": "Annual", "ce_hours": "0", "fee_range": "$100-$200"},
    "kansas": {"name": "Kansas", "abbr": "KS", "board": "No statewide license — local jurisdiction", "license_required": False, "threshold": "N/A", "types": ["Varies by city/county"], "reciprocity": [], "exam_required": False, "insurance_required": True, "bond_required": False, "bond_amount": "N/A", "renewal": "Varies", "ce_hours": "Varies", "fee_range": "Varies"},
    "kentucky": {"name": "Kentucky", "abbr": "KY", "board": "Kentucky Division of HVAC", "license_required": True, "threshold": "$0 (HVAC required, GC varies)", "types": ["HVAC Contractor", "Electrical", "Plumbing"], "reciprocity": [], "exam_required": True, "insurance_required": True, "bond_required": False, "bond_amount": "N/A", "renewal": "Annual", "ce_hours": "6", "fee_range": "$75-$200"},
    "louisiana": {"name": "Louisiana", "abbr": "LA", "board": "Louisiana State Licensing Board for Contractors", "license_required": True, "threshold": "$50,000", "types": ["Building Construction", "Heavy Construction", "Highway/Bridge Construction", "Municipal & Public Works"], "reciprocity": ["Alabama", "Mississippi"], "exam_required": True, "insurance_required": True, "bond_required": True, "bond_amount": "$5,000-$10,000", "renewal": "Annual", "ce_hours": "0", "fee_range": "$200-$350"},
    "maine": {"name": "Maine", "abbr": "ME", "board": "No statewide license — local jurisdiction", "license_required": False, "threshold": "N/A", "types": ["Varies by town/city"], "reciprocity": [], "exam_required": False, "insurance_required": True, "bond_required": False, "bond_amount": "N/A", "renewal": "Varies", "ce_hours": "Varies", "fee_range": "Varies"},
    "maryland": {"name": "Maryland", "abbr": "MD", "board": "Maryland Home Improvement Commission", "license_required": True, "threshold": "$500", "types": ["Home Improvement Contractor", "Subcontractor", "Salesperson"], "reciprocity": [], "exam_required": False, "insurance_required": True, "bond_required": False, "bond_amount": "N/A", "renewal": "Biennial", "ce_hours": "0", "fee_range": "$300-$500"},
    "massachusetts": {"name": "Massachusetts", "abbr": "MA", "board": "Office of Consumer Affairs and Business Regulation", "license_required": True, "threshold": "$0", "types": ["Construction Supervisor (Unrestricted)", "Construction Supervisor (Restricted)"], "reciprocity": [], "exam_required": True, "insurance_required": True, "bond_required": False, "bond_amount": "N/A", "renewal": "Biennial", "ce_hours": "12", "fee_range": "$200-$350"},
    "michigan": {"name": "Michigan", "abbr": "MI", "board": "Michigan LARA — Builders License Board", "license_required": True, "threshold": "$600", "types": ["Residential Builder", "Residential Maintenance & Alteration Contractor"], "reciprocity": [], "exam_required": True, "insurance_required": True, "bond_required": False, "bond_amount": "N/A", "renewal": "Triennial", "ce_hours": "21", "fee_range": "$180-$300"},
    "minnesota": {"name": "Minnesota", "abbr": "MN", "board": "Minnesota Department of Labor and Industry", "license_required": True, "threshold": "$0", "types": ["Residential Building Contractor", "Residential Remodeler"], "reciprocity": [], "exam_required": True, "insurance_required": True, "bond_required": False, "bond_amount": "N/A", "renewal": "Biennial", "ce_hours": "14", "fee_range": "$200-$350"},
    "mississippi": {"name": "Mississippi", "abbr": "MS", "board": "Mississippi State Board of Contractors", "license_required": True, "threshold": "$50,000", "types": ["General Building", "Heavy Construction", "Highway", "Residential"], "reciprocity": ["Alabama", "Louisiana"], "exam_required": True, "insurance_required": True, "bond_required": True, "bond_amount": "Varies", "renewal": "Annual", "ce_hours": "0", "fee_range": "$200-$400"},
    "missouri": {"name": "Missouri", "abbr": "MO", "board": "No statewide license — local jurisdiction", "license_required": False, "threshold": "N/A", "types": ["Varies by city/county"], "reciprocity": [], "exam_required": False, "insurance_required": True, "bond_required": False, "bond_amount": "N/A", "renewal": "Varies", "ce_hours": "Varies", "fee_range": "Varies"},
    "montana": {"name": "Montana", "abbr": "MT", "board": "Montana Department of Labor and Industry", "license_required": True, "threshold": "$0", "types": ["General Contractor", "Specialty Contractor"], "reciprocity": [], "exam_required": False, "insurance_required": True, "bond_required": False, "bond_amount": "N/A", "renewal": "Annual", "ce_hours": "0", "fee_range": "$100-$250"},
    "nebraska": {"name": "Nebraska", "abbr": "NE", "board": "No statewide license — local jurisdiction", "license_required": False, "threshold": "N/A", "types": ["Varies by city/county"], "reciprocity": [], "exam_required": False, "insurance_required": True, "bond_required": False, "bond_amount": "N/A", "renewal": "Varies", "ce_hours": "Varies", "fee_range": "Varies"},
    "nevada": {"name": "Nevada", "abbr": "NV", "board": "Nevada State Contractors Board", "license_required": True, "threshold": "$0", "types": ["General Building (B)", "General Engineering (A)", "Specialty (C)"], "reciprocity": [], "exam_required": True, "insurance_required": True, "bond_required": True, "bond_amount": "$1,000-$500,000", "renewal": "Biennial", "ce_hours": "0", "fee_range": "$300-$700"},
    "new-hampshire": {"name": "New Hampshire", "abbr": "NH", "board": "No statewide GC license", "license_required": False, "threshold": "N/A", "types": ["Electrical, Plumbing, HVAC require state licenses"], "reciprocity": [], "exam_required": False, "insurance_required": True, "bond_required": False, "bond_amount": "N/A", "renewal": "Varies", "ce_hours": "Varies", "fee_range": "Varies"},
    "new-jersey": {"name": "New Jersey", "abbr": "NJ", "board": "New Jersey Division of Consumer Affairs", "license_required": True, "threshold": "$0 (Home Improvement)", "types": ["Home Improvement Contractor"], "reciprocity": [], "exam_required": False, "insurance_required": True, "bond_required": False, "bond_amount": "N/A", "renewal": "Biennial", "ce_hours": "0", "fee_range": "$90-$200"},
    "new-mexico": {"name": "New Mexico", "abbr": "NM", "board": "New Mexico Construction Industries Division", "license_required": True, "threshold": "$0", "types": ["General Building (GB-98)", "General Engineering (GE-98)", "Specialty"], "reciprocity": [], "exam_required": True, "insurance_required": True, "bond_required": True, "bond_amount": "$5,000-$15,000", "renewal": "Triennial", "ce_hours": "8", "fee_range": "$250-$500"},
    "new-york": {"name": "New York", "abbr": "NY", "board": "No statewide license — NYC/local", "license_required": False, "threshold": "N/A", "types": ["NYC requires license; rest varies by locality"], "reciprocity": [], "exam_required": False, "insurance_required": True, "bond_required": False, "bond_amount": "N/A", "renewal": "Varies", "ce_hours": "Varies", "fee_range": "Varies"},
    "north-carolina": {"name": "North Carolina", "abbr": "NC", "board": "North Carolina Licensing Board for General Contractors", "license_required": True, "threshold": "$30,000", "types": ["General Contractor (Unlimited)", "General Contractor (Intermediate)", "General Contractor (Limited)"], "reciprocity": [], "exam_required": True, "insurance_required": True, "bond_required": False, "bond_amount": "N/A", "renewal": "Annual", "ce_hours": "0", "fee_range": "$200-$350"},
    "north-dakota": {"name": "North Dakota", "abbr": "ND", "board": "North Dakota Secretary of State", "license_required": True, "threshold": "$0", "types": ["General Contractor"], "reciprocity": [], "exam_required": False, "insurance_required": True, "bond_required": False, "bond_amount": "N/A", "renewal": "Annual", "ce_hours": "0", "fee_range": "$75-$150"},
    "ohio": {"name": "Ohio", "abbr": "OH", "board": "No statewide license — local jurisdiction", "license_required": False, "threshold": "N/A", "types": ["Varies by city/county"], "reciprocity": [], "exam_required": False, "insurance_required": True, "bond_required": False, "bond_amount": "N/A", "renewal": "Varies", "ce_hours": "Varies", "fee_range": "Varies"},
    "oklahoma": {"name": "Oklahoma", "abbr": "OK", "board": "Construction Industries Board of Oklahoma", "license_required": True, "threshold": "$50,000", "types": ["General Contractor", "Specialty Contractor"], "reciprocity": [], "exam_required": True, "insurance_required": True, "bond_required": False, "bond_amount": "N/A", "renewal": "Annual", "ce_hours": "0", "fee_range": "$100-$300"},
    "oregon": {"name": "Oregon", "abbr": "OR", "board": "Oregon Construction Contractors Board (CCB)", "license_required": True, "threshold": "$0", "types": ["General Contractor", "Specialty Contractor", "Residential Contractor"], "reciprocity": [], "exam_required": True, "insurance_required": True, "bond_required": True, "bond_amount": "$10,000-$75,000", "renewal": "Biennial", "ce_hours": "0", "fee_range": "$200-$400"},
    "pennsylvania": {"name": "Pennsylvania", "abbr": "PA", "board": "Pennsylvania Attorney General — Home Improvement", "license_required": True, "threshold": "$0 (Home Improvement)", "types": ["Home Improvement Contractor"], "reciprocity": [], "exam_required": False, "insurance_required": True, "bond_required": False, "bond_amount": "N/A", "renewal": "Annual", "ce_hours": "0", "fee_range": "$50-$200"},
    "rhode-island": {"name": "Rhode Island", "abbr": "RI", "board": "Rhode Island Contractors Registration Board", "license_required": True, "threshold": "$0", "types": ["General Contractor", "Subcontractor"], "reciprocity": [], "exam_required": True, "insurance_required": True, "bond_required": False, "bond_amount": "N/A", "renewal": "Annual", "ce_hours": "0", "fee_range": "$150-$300"},
    "south-carolina": {"name": "South Carolina", "abbr": "SC", "board": "South Carolina Contractors Licensing Board", "license_required": True, "threshold": "$5,000", "types": ["General Contractor", "Mechanical Contractor"], "reciprocity": [], "exam_required": True, "insurance_required": True, "bond_required": False, "bond_amount": "N/A", "renewal": "Biennial", "ce_hours": "0", "fee_range": "$200-$400"},
    "south-dakota": {"name": "South Dakota", "abbr": "SD", "board": "No statewide license — local jurisdiction", "license_required": False, "threshold": "N/A", "types": ["Varies by city/county"], "reciprocity": [], "exam_required": False, "insurance_required": True, "bond_required": False, "bond_amount": "N/A", "renewal": "Varies", "ce_hours": "Varies", "fee_range": "Varies"},
    "tennessee": {"name": "Tennessee", "abbr": "TN", "board": "Tennessee Board for Licensing Contractors", "license_required": True, "threshold": "$25,000", "types": ["Contractor (BC-A, BC-B)", "Home Improvement (HI)", "Limited Licensed Electrician (LLE)"], "reciprocity": [], "exam_required": True, "insurance_required": True, "bond_required": True, "bond_amount": "$10,000", "renewal": "Annual", "ce_hours": "0", "fee_range": "$250-$500"},
    "texas": {"name": "Texas", "abbr": "TX", "board": "No statewide GC license", "license_required": False, "threshold": "N/A", "types": ["No state license for GCs; Plumbing, Electrical, HVAC require state licenses"], "reciprocity": [], "exam_required": False, "insurance_required": True, "bond_required": False, "bond_amount": "N/A", "renewal": "N/A", "ce_hours": "N/A", "fee_range": "N/A"},
    "utah": {"name": "Utah", "abbr": "UT", "board": "Utah Division of Occupational and Professional Licensing", "license_required": True, "threshold": "$0", "types": ["General Building Contractor (B100)", "General Engineering Contractor (E100)", "Specialty Contractor (S)"], "reciprocity": [], "exam_required": True, "insurance_required": True, "bond_required": False, "bond_amount": "N/A", "renewal": "Biennial", "ce_hours": "6", "fee_range": "$200-$400"},
    "vermont": {"name": "Vermont", "abbr": "VT", "board": "No statewide license — local jurisdiction", "license_required": False, "threshold": "N/A", "types": ["Varies by town"], "reciprocity": [], "exam_required": False, "insurance_required": True, "bond_required": False, "bond_amount": "N/A", "renewal": "Varies", "ce_hours": "Varies", "fee_range": "Varies"},
    "virginia": {"name": "Virginia", "abbr": "VA", "board": "Virginia Board for Contractors", "license_required": True, "threshold": "$1,000", "types": ["Class A (>$120K)", "Class B ($10K-$120K)", "Class C ($1K-$10K)"], "reciprocity": [], "exam_required": True, "insurance_required": True, "bond_required": True, "bond_amount": "Varies by class", "renewal": "Biennial", "ce_hours": "2", "fee_range": "$200-$500"},
    "washington": {"name": "Washington", "abbr": "WA", "board": "Washington Department of Labor & Industries", "license_required": True, "threshold": "$0", "types": ["General Contractor", "Specialty Contractor"], "reciprocity": [], "exam_required": False, "insurance_required": True, "bond_required": True, "bond_amount": "$12,000-$30,000", "renewal": "Biennial", "ce_hours": "0", "fee_range": "$150-$300"},
    "west-virginia": {"name": "West Virginia", "abbr": "WV", "board": "West Virginia Contractor Licensing Board", "license_required": True, "threshold": "$2,500", "types": ["General Contractor", "Specialty Contractor"], "reciprocity": [], "exam_required": True, "insurance_required": True, "bond_required": False, "bond_amount": "N/A", "renewal": "Annual", "ce_hours": "0", "fee_range": "$150-$300"},
    "wisconsin": {"name": "Wisconsin", "abbr": "WI", "board": "Wisconsin Department of Safety and Professional Services", "license_required": True, "threshold": "$0 (Dwelling Contractor)", "types": ["Dwelling Contractor", "Dwelling Contractor Qualifier"], "reciprocity": [], "exam_required": True, "insurance_required": True, "bond_required": False, "bond_amount": "N/A", "renewal": "Biennial", "ce_hours": "12", "fee_range": "$100-$250"},
    "wyoming": {"name": "Wyoming", "abbr": "WY", "board": "No statewide license — local jurisdiction", "license_required": False, "threshold": "N/A", "types": ["Varies by city/county"], "reciprocity": [], "exam_required": False, "insurance_required": True, "bond_required": False, "bond_amount": "N/A", "renewal": "Varies", "ce_hours": "Varies", "fee_range": "Varies"},
}

def generate_page(slug, data):
    state = data["name"]
    abbr = data["abbr"]
    
    license_types_html = "".join([f"<li>{t}</li>" for t in data["types"]])
    reciprocity_html = ""
    if data["reciprocity"]:
        recip_links = ", ".join([f'<a href="/licensing/contractor-license-{s.lower().replace(" ", "-")}">{s}</a>' for s in data["reciprocity"]])
        reciprocity_html = f"""
        <div class="info-box">
            <h3>🤝 Reciprocity Agreements</h3>
            <p>{state} has reciprocity agreements with: {recip_links}. This means your license may transfer or the application process may be simplified.</p>
        </div>"""
    
    if data["license_required"]:
        overview = f"""<p><strong>Yes, {state} requires a contractor license</strong> for most construction work{f' over {data["threshold"]}' if data["threshold"] != "$0" else ''}. 
        The licensing is managed by the <strong>{data["board"]}</strong>.</p>
        <p>Whether you're a general contractor, specialty contractor, or residential builder, understanding {state}'s licensing requirements 
        is essential before you take on any projects in the state.</p>"""
        
        requirements_section = f"""
        <h2>📋 {state} Contractor License Requirements</h2>
        <div class="requirements-grid">
            <div class="req-card">
                <h3>License Types</h3>
                <ul>{license_types_html}</ul>
            </div>
            <div class="req-card">
                <h3>Project Threshold</h3>
                <p>License required for projects valued at <strong>{data["threshold"]}</strong> or more</p>
            </div>
            <div class="req-card">
                <h3>Exam Required</h3>
                <p><strong>{"Yes" if data["exam_required"] else "No"}</strong> — {"You must pass a trade and/or business exam" if data["exam_required"] else "No exam required for licensing"}</p>
            </div>
            <div class="req-card">
                <h3>Insurance Required</h3>
                <p><strong>Yes</strong> — General liability insurance is required</p>
            </div>
            <div class="req-card">
                <h3>Bond Required</h3>
                <p><strong>{"Yes — " + data["bond_amount"] if data["bond_required"] else "No surety bond required"}</strong></p>
            </div>
            <div class="req-card">
                <h3>Renewal</h3>
                <p><strong>{data["renewal"]}</strong> renewal{f' with {data["ce_hours"]} hours of continuing education' if data["ce_hours"] != "0" else ''}</p>
            </div>
            <div class="req-card">
                <h3>Application Fees</h3>
                <p><strong>{data["fee_range"]}</strong> (varies by license class)</p>
            </div>
        </div>
        
        <h2>📝 How to Get a Contractor License in {state}</h2>
        <div class="steps">
            <div class="step">
                <div class="step-num">1</div>
                <div class="step-content">
                    <h3>Meet Basic Eligibility</h3>
                    <p>You must be at least 18 years old, have relevant experience in construction (typically 2-4 years), 
                    and have a clean criminal background. Some states also require proof of financial solvency.</p>
                </div>
            </div>
            <div class="step">
                <div class="step-num">2</div>
                <div class="step-content">
                    <h3>{"Pass the Required Exams" if data["exam_required"] else "Complete Registration"}</h3>
                    <p>{"The " + state + " contractor exam typically covers trade knowledge, business law, safety regulations, and project management. Study guides and prep courses are available." if data["exam_required"] else "Complete the registration application with the " + data["board"] + ". No exam is required."}</p>
                </div>
            </div>
            <div class="step">
                <div class="step-num">3</div>
                <div class="step-content">
                    <h3>Get Insurance & {"Bonding" if data["bond_required"] else "Coverage"}</h3>
                    <p>Obtain general liability insurance (minimum $500,000-$1,000,000 recommended){" and a surety bond of " + data["bond_amount"] if data["bond_required"] else ""}. 
                    Workers' compensation insurance is also required if you have employees.</p>
                </div>
            </div>
            <div class="step">
                <div class="step-num">4</div>
                <div class="step-content">
                    <h3>Submit Your Application</h3>
                    <p>File your application with the {data["board"]}. Include all required documentation, 
                    proof of insurance, {"bond certificate, " if data["bond_required"] else ""}and application fee ({data["fee_range"]}).</p>
                </div>
            </div>
            <div class="step">
                <div class="step-num">5</div>
                <div class="step-content">
                    <h3>Maintain Your License</h3>
                    <p>Renew your license {data["renewal"].lower()}{" and complete " + data["ce_hours"] + " hours of continuing education" if data["ce_hours"] not in ("0", "Varies") else ""}. 
                    Keep your insurance and {"bond " if data["bond_required"] else ""}current at all times.</p>
                </div>
            </div>
        </div>"""
    else:
        overview = f"""<p><strong>{state} does not require a statewide contractor license</strong> for general contractors. However, 
        this doesn't mean you can work without any licensing at all.</p>
        <p>Many cities and counties in {state} have their own licensing requirements. Additionally, specialty trades 
        (electrical, plumbing, HVAC) typically require state-level licenses even when general contracting doesn't.</p>
        <p><strong>Important:</strong> Even without a state license requirement, you still need proper insurance, 
        may need local permits, and must comply with building codes.</p>"""
        
        requirements_section = f"""
        <h2>📋 What {state} Contractors Still Need</h2>
        <div class="requirements-grid">
            <div class="req-card">
                <h3>Local Licenses</h3>
                <p>Check with your city/county for local contractor licensing requirements. Major cities usually require licensing.</p>
            </div>
            <div class="req-card">
                <h3>Business Registration</h3>
                <p>Register your business with the {state} Secretary of State and obtain a business license.</p>
            </div>
            <div class="req-card">
                <h3>Insurance</h3>
                <p><strong>Required</strong> — General liability and workers' comp insurance are essential even without a state license.</p>
            </div>
            <div class="req-card">
                <h3>Specialty Licenses</h3>
                <p>Electrical, plumbing, and HVAC work typically requires state-level licensing even in {state}.</p>
            </div>
            <div class="req-card">
                <h3>Building Permits</h3>
                <p>You must obtain building permits for construction projects regardless of licensing requirements.</p>
            </div>
            <div class="req-card">
                <h3>Tax Registration</h3>
                <p>Register for state sales tax (if applicable) and federal EIN for your contracting business.</p>
            </div>
        </div>"""

    # Related states - pick 5 neighbors
    all_slugs = list(STATES.keys())
    idx = all_slugs.index(slug)
    related = all_slugs[max(0,idx-2):idx] + all_slugs[idx+1:idx+4]
    related_html = "".join([f'<a href="/licensing/contractor-license-{s}" class="related-link">📄 {STATES[s]["name"]} Contractor License</a>' for s in related[:5]])

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Contractor License {state} ({abbr}) — Requirements, Cost & How to Apply | BuiltRight Academy</title>
    <meta name="description" content="Complete guide to getting a contractor license in {state}. Learn about {abbr} licensing requirements, costs ({data['fee_range']}), exams, insurance, and step-by-step application process.">
    <meta name="keywords" content="contractor license {state.lower()}, {abbr.lower()} contractor license, how to get contractor license {state.lower()}, {state.lower()} contractor license requirements, {state.lower()} contractor license cost">
    <link rel="canonical" href="https://builtright-academy.vercel.app/licensing/contractor-license-{slug}">
    <script type="application/ld+json">
    {{
        "@context": "https://schema.org",
        "@type": "Article",
        "headline": "Contractor License in {state} — Complete Guide",
        "description": "Everything you need to know about getting a contractor license in {state}.",
        "author": {{"@type": "Organization", "name": "BuiltRight Academy"}},
        "publisher": {{"@type": "Organization", "name": "BuiltRight Academy"}},
        "datePublished": "2026-03-07",
        "dateModified": "2026-03-07"
    }}
    </script>
    <style>
        * {{ margin: 0; padding: 0; box-sizing: border-box; }}
        body {{ font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; line-height: 1.7; color: #1a1a1a; background: #fafafa; }}
        .header {{ background: linear-gradient(135deg, #1e3a5f 0%, #2d5a8e 100%); color: white; padding: 2rem 0; }}
        .header-inner {{ max-width: 900px; margin: 0 auto; padding: 0 1.5rem; }}
        .header a {{ color: white; text-decoration: none; font-weight: 700; font-size: 1.25rem; }}
        .breadcrumb {{ max-width: 900px; margin: 1rem auto; padding: 0 1.5rem; font-size: 0.9rem; color: #666; }}
        .breadcrumb a {{ color: #2d5a8e; text-decoration: none; }}
        .content {{ max-width: 900px; margin: 0 auto; padding: 1rem 1.5rem 3rem; }}
        h1 {{ font-size: 2rem; margin-bottom: 1rem; color: #1e3a5f; line-height: 1.3; }}
        h2 {{ font-size: 1.5rem; margin: 2.5rem 0 1rem; color: #1e3a5f; border-bottom: 2px solid #e8e8e8; padding-bottom: 0.5rem; }}
        h3 {{ font-size: 1.15rem; margin-bottom: 0.5rem; color: #2d5a8e; }}
        p {{ margin-bottom: 1rem; }}
        .hero-badge {{ display: inline-block; background: {"#27ae60" if data["license_required"] else "#e67e22"}; color: white; padding: 0.3rem 0.8rem; border-radius: 4px; font-size: 0.85rem; font-weight: 600; margin-bottom: 1rem; }}
        .quick-facts {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 1rem; margin: 1.5rem 0; }}
        .fact {{ background: white; padding: 1rem; border-radius: 8px; border: 1px solid #e8e8e8; }}
        .fact-label {{ font-size: 0.8rem; color: #888; text-transform: uppercase; letter-spacing: 0.5px; }}
        .fact-value {{ font-size: 1.1rem; font-weight: 600; color: #1e3a5f; }}
        .requirements-grid {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(280px, 1fr)); gap: 1rem; margin: 1.5rem 0; }}
        .req-card {{ background: white; padding: 1.25rem; border-radius: 8px; border: 1px solid #e8e8e8; }}
        .req-card ul {{ margin-left: 1.2rem; }}
        .req-card li {{ margin-bottom: 0.3rem; }}
        .steps {{ margin: 1.5rem 0; }}
        .step {{ display: flex; gap: 1rem; margin-bottom: 1.5rem; }}
        .step-num {{ width: 40px; height: 40px; background: #2d5a8e; color: white; border-radius: 50%; display: flex; align-items: center; justify-content: center; font-weight: 700; flex-shrink: 0; }}
        .step-content {{ flex: 1; }}
        .info-box {{ background: #f0f7ff; border-left: 4px solid #2d5a8e; padding: 1rem 1.25rem; border-radius: 0 8px 8px 0; margin: 1.5rem 0; }}
        .cta-box {{ background: linear-gradient(135deg, #1e3a5f 0%, #2d5a8e 100%); color: white; padding: 2rem; border-radius: 12px; margin: 2rem 0; text-align: center; }}
        .cta-box h3 {{ color: white; font-size: 1.3rem; margin-bottom: 0.5rem; }}
        .cta-box p {{ color: #c8daf0; margin-bottom: 1rem; }}
        .cta-box input {{ padding: 0.7rem 1rem; border: none; border-radius: 6px; width: 280px; max-width: 100%; font-size: 1rem; }}
        .cta-box button {{ padding: 0.7rem 1.5rem; background: #f39c12; color: #1a1a1a; border: none; border-radius: 6px; font-weight: 700; cursor: pointer; font-size: 1rem; margin-left: 0.5rem; }}
        .cta-box button:hover {{ background: #e67e22; }}
        .related-states {{ margin-top: 2rem; }}
        .related-link {{ display: inline-block; background: white; padding: 0.5rem 1rem; border: 1px solid #e8e8e8; border-radius: 6px; margin: 0.25rem; text-decoration: none; color: #2d5a8e; font-size: 0.9rem; }}
        .related-link:hover {{ background: #f0f7ff; }}
        .footer {{ background: #1e3a5f; color: #c8daf0; padding: 2rem 0; text-align: center; font-size: 0.85rem; }}
        .footer a {{ color: #f39c12; text-decoration: none; }}
        @media (max-width: 600px) {{
            h1 {{ font-size: 1.5rem; }}
            .cta-box input {{ width: 100%; margin-bottom: 0.5rem; }}
            .cta-box button {{ margin-left: 0; width: 100%; }}
        }}
    </style>
</head>
<body>
    <div class="header">
        <div class="header-inner">
            <a href="/">🔨 BuiltRight Academy</a>
        </div>
    </div>
    
    <div class="breadcrumb">
        <a href="/">Home</a> &rsaquo; <a href="/licensing/contractor-license-by-state">Contractor Licenses by State</a> &rsaquo; {state}
    </div>
    
    <div class="content">
        <span class="hero-badge">{"✅ State License Required" if data["license_required"] else "⚠️ No Statewide License"}</span>
        <h1>Contractor License in {state} ({abbr}) — Complete {datetime.now().year} Guide</h1>
        
        {overview}
        
        <div class="quick-facts">
            <div class="fact">
                <div class="fact-label">State</div>
                <div class="fact-value">{state} ({abbr})</div>
            </div>
            <div class="fact">
                <div class="fact-label">License Required</div>
                <div class="fact-value">{"Yes" if data["license_required"] else "No (local may apply)"}</div>
            </div>
            <div class="fact">
                <div class="fact-label">{"Exam Required" if data["license_required"] else "Specialty Trades"}</div>
                <div class="fact-value">{"Yes" if data["exam_required"] else "May require state license" if not data["license_required"] else "No"}</div>
            </div>
            <div class="fact">
                <div class="fact-label">Application Fee</div>
                <div class="fact-value">{data["fee_range"]}</div>
            </div>
        </div>
        
        {requirements_section}
        {reciprocity_html}
        
        <h2>💰 Cost of Getting Licensed in {state}</h2>
        <p>The total cost to obtain a contractor license in {state} includes:</p>
        <ul style="margin-left: 1.5rem; margin-bottom: 1rem;">
            <li><strong>Application fee:</strong> {data["fee_range"]}</li>
            {"<li><strong>Exam fee:</strong> $75-$150 (if required)</li>" if data["exam_required"] else ""}
            <li><strong>General liability insurance:</strong> $500-$2,000/year (varies by coverage)</li>
            {"<li><strong>Surety bond:</strong> " + data["bond_amount"] + " (premium is typically 1-15% of bond amount)</li>" if data["bond_required"] else ""}
            <li><strong>Workers' compensation:</strong> Required if you have employees ($1,000-$5,000+/year)</li>
            <li><strong>Business registration:</strong> $50-$500 depending on entity type</li>
        </ul>
        <p><strong>Estimated total startup cost:</strong> $1,500-$5,000 depending on your license class and business structure.</p>
        
        <div class="cta-box">
            <h3>📥 Free Contractor Business Toolkit</h3>
            <p>Get free invoice templates, estimate forms, and profit tracking spreadsheets built for contractors.</p>
            <form onsubmit="event.preventDefault(); fetch('/api/subscribe', {{method:'POST', headers:{{'Content-Type':'application/json'}}, body:JSON.stringify({{email:this.email.value,source:'licensing-{slug}',lead_magnet:'contractor-toolkit'}})}}).then(()=>{{this.innerHTML='<p style=\\'color:white;font-weight:700\\'>✅ Check your email! Templates are on the way.</p>'}})" style="margin-top: 0.5rem;">
                <input type="email" name="email" placeholder="your@email.com" required>
                <button type="submit">Get Free Templates</button>
            </form>
        </div>
        
        <h2>❓ Frequently Asked Questions</h2>
        <div class="info-box">
            <h3>Do I need a contractor license in {state}?</h3>
            <p>{"Yes. " + state + " requires a contractor license for projects valued at " + data["threshold"] + " or more. The license is issued by the " + data["board"] + "." if data["license_required"] else state + " does not have a statewide contractor license requirement. However, many cities and counties have their own licensing requirements, and specialty trades (electrical, plumbing, HVAC) typically need state-level licenses."}</p>
        </div>
        <div class="info-box">
            <h3>How long does it take to get a contractor license in {state}?</h3>
            <p>The process typically takes 2-8 weeks, depending on application volume, {"exam scheduling, " if data["exam_required"] else ""}and completeness of your documentation. Having all your paperwork ready can speed up the process significantly.</p>
        </div>
        <div class="info-box">
            <h3>Can I work in {state} with an out-of-state license?</h3>
            <p>{"Generally no. " + state + " requires its own state-specific license." if not data["reciprocity"] else state + " has reciprocity agreements with " + ", ".join(data["reciprocity"]) + ", which may simplify the licensing process."} You should contact the {data["board"]} for the most current reciprocity information.</p>
        </div>
        
        <script type="application/ld+json">
        {{
            "@context": "https://schema.org",
            "@type": "FAQPage",
            "mainEntity": [
                {{
                    "@type": "Question",
                    "name": "Do I need a contractor license in {state}?",
                    "acceptedAnswer": {{
                        "@type": "Answer",
                        "text": "{"Yes. " + state + " requires a contractor license for projects valued at " + data["threshold"] + " or more." if data["license_required"] else state + " does not have a statewide contractor license requirement, but local jurisdictions may require one."}"
                    }}
                }},
                {{
                    "@type": "Question",
                    "name": "How much does a contractor license cost in {state}?",
                    "acceptedAnswer": {{
                        "@type": "Answer",
                        "text": "Application fees range from {data["fee_range"]}. Total startup costs including insurance {"and bonding " if data["bond_required"] else ""}are typically $1,500-$5,000."
                    }}
                }},
                {{
                    "@type": "Question",
                    "name": "How long does it take to get a contractor license in {state}?",
                    "acceptedAnswer": {{
                        "@type": "Answer",
                        "text": "The process typically takes 2-8 weeks depending on application completeness{"and exam scheduling" if data["exam_required"] else ""}."
                    }}
                }}
            ]
        }}
        </script>
        
        <div class="related-states">
            <h2>📍 Contractor Licenses in Other States</h2>
            {related_html}
            <br><br>
            <a href="/licensing/contractor-license-by-state" class="related-link">📋 View All 50 States</a>
        </div>
    </div>
    
    <div class="footer">
        <p>&copy; 2026 <a href="/">BuiltRight Academy</a> — Business Skills for Tradespeople</p>
        <p style="margin-top: 0.5rem;">
            <a href="/tools">Calculators</a> · 
            <a href="/blog">Articles</a> · 
            <a href="/templates">Templates</a> · 
            <a href="/licensing/contractor-license-by-state">Licensing Guides</a>
        </p>
    </div>
    <script>
        if(window.va)window.va('event',{{name:'pageview',data:{{page:'licensing-{slug}'}}}});
    </script>
</body>
</html>"""
    return html


def generate_index_page():
    """Generate the state index page."""
    states_by_letter = {}
    for slug, data in sorted(STATES.items(), key=lambda x: x[1]["name"]):
        first = data["name"][0]
        if first not in states_by_letter:
            states_by_letter[first] = []
        states_by_letter[first].append((slug, data))
    
    state_cards = ""
    for letter in sorted(states_by_letter.keys()):
        for slug, data in states_by_letter[letter]:
            badge = "✅ Required" if data["license_required"] else "⚠️ Local Only"
            badge_color = "#27ae60" if data["license_required"] else "#e67e22"
            state_cards += f"""
            <a href="/licensing/contractor-license-{slug}" class="state-card">
                <span class="state-name">{data["name"]}</span>
                <span class="state-badge" style="background:{badge_color}">{badge}</span>
                <span class="state-detail">{data["fee_range"]}</span>
            </a>"""
    
    return f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Contractor License Requirements by State (2026) — All 50 States | BuiltRight Academy</title>
    <meta name="description" content="Complete guide to contractor licensing requirements in all 50 US states. Find out if your state requires a contractor license, costs, exams, and how to apply.">
    <meta name="keywords" content="contractor license by state, general contractor license requirements, how to get a contractor license, contractor license cost">
    <link rel="canonical" href="https://builtright-academy.vercel.app/licensing/contractor-license-by-state">
    <style>
        * {{ margin: 0; padding: 0; box-sizing: border-box; }}
        body {{ font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; line-height: 1.7; color: #1a1a1a; background: #fafafa; }}
        .header {{ background: linear-gradient(135deg, #1e3a5f 0%, #2d5a8e 100%); color: white; padding: 2rem 0; }}
        .header-inner {{ max-width: 1100px; margin: 0 auto; padding: 0 1.5rem; }}
        .header a {{ color: white; text-decoration: none; font-weight: 700; font-size: 1.25rem; }}
        .content {{ max-width: 1100px; margin: 0 auto; padding: 2rem 1.5rem 3rem; }}
        h1 {{ font-size: 2rem; margin-bottom: 1rem; color: #1e3a5f; line-height: 1.3; }}
        h2 {{ font-size: 1.4rem; margin: 2rem 0 1rem; color: #1e3a5f; }}
        p {{ margin-bottom: 1rem; }}
        .states-grid {{ display: grid; grid-template-columns: repeat(auto-fill, minmax(250px, 1fr)); gap: 0.75rem; margin: 1.5rem 0; }}
        .state-card {{ display: flex; flex-direction: column; background: white; padding: 1rem; border-radius: 8px; border: 1px solid #e8e8e8; text-decoration: none; color: inherit; transition: box-shadow 0.2s; }}
        .state-card:hover {{ box-shadow: 0 2px 8px rgba(0,0,0,0.1); }}
        .state-name {{ font-weight: 700; color: #1e3a5f; font-size: 1.05rem; }}
        .state-badge {{ display: inline-block; color: white; padding: 0.15rem 0.5rem; border-radius: 3px; font-size: 0.75rem; font-weight: 600; margin: 0.3rem 0; width: fit-content; }}
        .state-detail {{ font-size: 0.85rem; color: #666; }}
        .summary {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 1rem; margin: 1.5rem 0; }}
        .sum-card {{ background: white; padding: 1.25rem; border-radius: 8px; border: 1px solid #e8e8e8; text-align: center; }}
        .sum-num {{ font-size: 2rem; font-weight: 700; color: #2d5a8e; }}
        .sum-label {{ font-size: 0.85rem; color: #666; }}
        .cta-box {{ background: linear-gradient(135deg, #1e3a5f 0%, #2d5a8e 100%); color: white; padding: 2rem; border-radius: 12px; margin: 2rem 0; text-align: center; }}
        .cta-box h3 {{ color: white; font-size: 1.3rem; margin-bottom: 0.5rem; }}
        .cta-box p {{ color: #c8daf0; }}
        .cta-box input {{ padding: 0.7rem 1rem; border: none; border-radius: 6px; width: 280px; max-width: 100%; font-size: 1rem; }}
        .cta-box button {{ padding: 0.7rem 1.5rem; background: #f39c12; color: #1a1a1a; border: none; border-radius: 6px; font-weight: 700; cursor: pointer; font-size: 1rem; margin-left: 0.5rem; }}
        .footer {{ background: #1e3a5f; color: #c8daf0; padding: 2rem 0; text-align: center; font-size: 0.85rem; }}
        .footer a {{ color: #f39c12; text-decoration: none; }}
    </style>
</head>
<body>
    <div class="header">
        <div class="header-inner">
            <a href="/">🔨 BuiltRight Academy</a>
        </div>
    </div>
    <div class="content">
        <h1>Contractor License Requirements by State — 2026 Complete Guide</h1>
        <p>Every state handles contractor licensing differently. Some states require a state-level license, others leave it to cities and counties. 
        This guide covers all 50 states so you know exactly what you need to work legally wherever you operate.</p>
        
        <div class="summary">
            <div class="sum-card">
                <div class="sum-num">{sum(1 for d in STATES.values() if d['license_required'])}</div>
                <div class="sum-label">States Requiring License</div>
            </div>
            <div class="sum-card">
                <div class="sum-num">{sum(1 for d in STATES.values() if not d['license_required'])}</div>
                <div class="sum-label">States — Local Only</div>
            </div>
            <div class="sum-card">
                <div class="sum-num">{sum(1 for d in STATES.values() if d['exam_required'])}</div>
                <div class="sum-label">States Requiring Exam</div>
            </div>
            <div class="sum-card">
                <div class="sum-num">{sum(1 for d in STATES.values() if d['bond_required'])}</div>
                <div class="sum-label">States Requiring Bond</div>
            </div>
        </div>
        
        <h2>All 50 States</h2>
        <div class="states-grid">
            {state_cards}
        </div>
        
        <div class="cta-box">
            <h3>📥 Free Contractor Business Toolkit</h3>
            <p>Invoice templates, estimate forms, job costing tracker & P&L spreadsheet — built for contractors.</p>
            <form onsubmit="event.preventDefault(); fetch('/api/subscribe', {{method:'POST', headers:{{'Content-Type':'application/json'}}, body:JSON.stringify({{email:this.email.value,source:'licensing-index',lead_magnet:'contractor-toolkit'}})}}).then(()=>{{this.innerHTML='<p style=\\'color:white;font-weight:700\\'>✅ Check your email!</p>'}})" style="margin-top: 0.5rem;">
                <input type="email" name="email" placeholder="your@email.com" required>
                <button type="submit">Get Free Templates</button>
            </form>
        </div>
    </div>
    <div class="footer">
        <p>&copy; 2026 <a href="/">BuiltRight Academy</a></p>
        <p style="margin-top: 0.5rem;"><a href="/tools">Calculators</a> · <a href="/blog">Articles</a> · <a href="/templates">Templates</a></p>
    </div>
</body>
</html>"""


if __name__ == "__main__":
    # Create licensing directory
    os.makedirs("licensing", exist_ok=True)
    
    # Generate individual state pages
    for slug, data in STATES.items():
        filename = f"licensing/contractor-license-{slug}.html"
        html = generate_page(slug, data)
        with open(filename, "w") as f:
            f.write(html)
        print(f"✅ Generated {filename}")
    
    # Generate index page
    index_html = generate_index_page()
    with open("licensing/contractor-license-by-state.html", "w") as f:
        f.write(index_html)
    print(f"✅ Generated licensing/contractor-license-by-state.html")
    
    print(f"\n🎉 Generated {len(STATES)} state pages + 1 index page = {len(STATES) + 1} new URLs!")
