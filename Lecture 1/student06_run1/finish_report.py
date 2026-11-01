from pathlib import Path
import json,html
P=Path(__file__).resolve().parent
m=json.loads((P/'metrics.json').read_text())
s=json.loads((P/'research.json').read_text())
links='\n'.join(f"- [{r[0]}]({r[1]}) — {r[3]}" for r in s['sources'])
text=f'''# Executive report — Student housing near VUB/ULB, Brussels
14 October 2026 | Illustrative acquisition and refurbishment | EUR | Pre-income-tax

The illustrative investment does **not meet the 9% equity return target**. A 40-room asset acquired for €3.20 million and refurbished at €18,000 per room generates a projected equity IRR of {m['irr']:.1%} and an NPV of €{m['npv']:,.0f}. This is a screening result, not a valuation or a recommendation about a specific property. No address, asking price, permitted capacity, or verified rent roll was supplied.

| Measure | Base case |
|---|---:|
| Total investment, including reserve | €{m['cost']:,.0f} |
| Acquisition debt | €1,920,000 |
| Initial equity | €{m['equity']:,.0f} |
| Year 1 net operating income | €{m['noi']:,.0f} |
| Year 1 cash flow after debt and replacement reserve | €{m['cash']:,.0f} |
| Year 1 cash yield on equity | {m['yield']:.1%} |
| Year 1 debt service coverage | {m['dscr']:.2f}× |
| Cash break-even economic occupancy | {m['break_even']:.1%} |
| Ten-year equity IRR | {m['irr']:.1%} |
| Equity multiple | {m['multiple']:.2f}× |
| Equity NPV at 9% | €{m['npv']:,.0f} |

The assumed rent is €650 per room per month, collected over 12 months at 95% economic occupancy. Rent grows by 2% annually, operating costs by 2.5%. Acquisition debt equals 60% of purchase price, with 4.5% interest and 25-year amortisation. Sale occurs after Year 10 using Year 11 NOI at a 5.5% exit yield, less 2% selling costs and remaining debt. These financial parameters are placeholders, not researched market comparables or financing offers.

The budget includes 12.5% acquisition duty on purchase price, 2% other acquisition costs, €720,000 refurbishment, a 10% refurbishment contingency, €80,000 setup costs, €100,000 funded reserve and debt arrangement / mortgage costs. Mortgage duty is modelled at 1% of an assumed secured amount equal to 110% of debt; another €10,000 is allowed for security and loan deed costs. A notary must replace these allowances with transaction-specific amounts. No owner-occupier tax relief is assumed.

Operating costs include management at 6% of effective revenue, €18,000 utilities/common areas, €14,000 maintenance, €6,000 insurance, €14,000 property tax and €5,000 administration in Year 1. An additional €12,000 replacement reserve is treated as spent each year. Tenant deposits are excluded from investment funding. The initial funded reserve is released on exit; if it is consumed, this release must be reduced. Costs are gross of assumed nonrecoverable VAT. Income tax, capital-gains tax and investor-specific tax effects are excluded and can materially reduce returns.

## Regulatory conditions before investment

Confirm the exact address and commune first. The City of Brussels is only one municipality in the region. The legal checklist records open property checks; it does not certify compliance. For an Ixelles address, request lawful-use information from the [municipal urban-planning service](https://www.ixelles.be/site/124-Urbanisme).

The regional planning guidance requires a permit when ordinary housing is converted into housing exclusively designated for students. Simply renting to a student does not establish a student-housing planning designation. Verify existing permits, approved plans and room numbers with the relevant commune before relying on a 40-room income forecast. [Urban.brussels guidance](https://urban.brussels/storage/uploads/7782bb55-caef-4a03-9780-aac7547d23f5/Vademecum---changements-utilisation-soumis-permis.pdf)

Housing must meet current regional safety, health and equipment standards, including the requirements effective from January 2026. Fire protection and lawful occupancy need building-specific professional review. Smoke detectors are mandatory in all regional housing from 2025, with configuration and interconnection requirements to check. [Housing standards](https://be.brussels/fr/logement/location/conformite-logement/normes-et-securite), [smoke detectors](https://be.brussels/fr/logement/construction-et-renovation/securite-et-reglementation/detecteurs-de-fumee)

The student lease regime is chosen jointly and requires proof of student status. Its maximum term is 12 months; tenants may terminate with two months' notice without an indemnity. This creates turnover risk even when a full academic-year lease is signed. Regional guidance requires registration within two months on both Irisrent and MyRent. [Student lease rules](https://be.brussels/fr/logement/location/bail-dhabitation/contrat-de-bail-modeles-et-annexe), [registration](https://be.brussels/nl/huisvesting/huren-en-verhuren/huurovereenkomst/registratie-woninghuurovereenkomst)

Excessive rents have been prohibited since May 2025. Check the proposed rent under applicable reference-rent and property-condition rules; €650 is not a confirmed lawful rent. Obtain the appropriate PEB certificate and verify its scope for individual rooms versus the residence. Indexation requires the applicable legal conditions, including registration and a valid certificate communicated to the tenant; certain older leases with E/F/G certificates require a correction factor. The 2% forecast is not automatic legal entitlement. [Regional rent regulation](https://be.brussels/fr/logement/bruxelles-logement-publie-son-rapport-dactivite-2025), [PEB requirements](https://environnement.brussels/pro/reglementation-et-inspection/obligations-et-autorisations/les-certificats-peb), [indexation](https://be.brussels/fr/logement/location/bail-dhabitation/indexation-des-loyers/indexation-des-loyers-avec-facteur-de-correction)

## Information and stress tests needed

Provide the exact address, asking price, lawful capacity, floor areas and approved plans; room and common-facility specification; permits and fire assessment; EPC and technical survey; contractor quotations; historic occupancy, lease lengths, base rents and tenant charges; cadastral/property tax records and local taxes; available equity, lender quotation, ownership structure and investor tax status. Confirm the intended VUB/ULB campuses and access time. These inputs establish whether this case describes an investable asset.

Use the workbook's sensitivity sheet to compare 10% lower/higher rents and exit yields of 4.5%, 5.5% and 6.5%. Separately change paid months to 10, occupancy to 85%, interest to 5.5%, capex to 20% above budget and construction delay to six months. At 10 paid months, Year 1 cash after debt falls to approximately €33,697 and coverage to 1.35×. At 85% occupancy with 12 months, it is approximately €50,805 and coverage 1.49×. These two stresses are separate and omit additional letting costs. Combining adverse assumptions is appropriate before deciding on funding.

At the illustrative terms, negotiate price or improve demonstrably lawful net income before proceeding. Approval should depend on a property-specific budget, legal clearance and downside cash-flow capacity. The model uses annual debt payments and a fixed ten-year horizon; changing the holding period requires extending formulas. Construction delay reduces first-year rent while full costs and debt service continue; interest is not capitalised. There is no assumption of guaranteed appreciation.

## Sources and reproducibility

Official sources were reviewed on 14 October 2026. Municipal rules remain address-specific. Financial inputs remain illustrative.

{links}

Files: `model.xlsx` (editable inputs and cached base-case outputs), `build_model.py` (baseline builder), `enhance_model.py` (dashboard, mortgage costs and formula caches), `research.json` (sources), `finish_report.py` (this report), and `validate_model.py` (verification). Rebuild in that order. Open in Excel to recalculate after changing assumptions.
'''
(P/'executive_report.md').write_text(text)
body='\n'.join('<p>'+html.escape(line)+'</p>' for line in text.split('\n') if line.strip())
(P/'executive_report.html').write_text('<!doctype html><meta charset="utf-8"><title>Student Housing Executive Report</title><style>body{font:15px/1.5 Arial;max-width:950px;margin:40px auto;color:#183b56}p{white-space:pre-wrap} @media print{body{margin:15mm;font-size:11pt}}</style>'+body)
print('Saved executive_report.md and printable executive_report.html. SAVE_NOW:',(P/'SAVE_NOW.txt').exists())
