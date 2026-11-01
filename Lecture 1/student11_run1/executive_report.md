# Brussels student housing investment — executive report

Decision Analytics course project · 14 October 2026  
Authors: Nisa, Maite, Ilias (names to confirm)

Recommend advancing the hypothetical projects D (Etterbeek), E (Schaerbeek), F (Jette) and G (Molenbeek-Saint-Jean) to a feasibility sprint. They require €14.50 million and provide 350 student rooms, including 172 designated affordable rooms. These are model-generated opportunities, not properties currently offered for sale. Capital commitment requires validated costs, municipal conditions and allocation arrangements.

The portfolio leaves only €0.50 million of the assumed €15 million budget unused. First-year net operating income is €1.216 million, approximately an 8.39% unlevered NOI yield. The modeled 20-year NPV is €11.864 million at a 6% discount rate, including terminal sale proceeds. NPV is sensitive to the assumed exit capitalization rate and immediate operation after investment.

| Fixed shortlist scenario | NPV | Investment | Budget status |
|---|---:|---:|---|
| Base | €11.864m | €14.500m | Within budget |
| Downside | −€0.434m | €16.675m | Exceeds budget by €1.675m |
| Upside | €21.206m | €13.775m | Within budget |

The downside combines 85% occupancy, 15% higher construction costs, 5% lower base rents, a 7.5% discount rate and a 6.5% exit capitalization rate. Secure additional contingency or reduce scope before committing. These scenarios are assumptions, not forecasts or assigned probabilities.

Brussels has constrained dedicated supply: Perspective identified 11,305 dedicated student units in 2025 and reports that roughly one in three housed students accesses this supply. Its comparison uses student demand from 2022–2023; this is contextual evidence rather than a contemporaneous demand forecast. [Perspective monitoring](https://perspective.brussels/fr/actualites/un-etudiant-sur-trois-trouve-un-kot-dans-un-logement-qui-lui-est-destine).

Kotkompas 2025–2026 reports a Brussels mean basic rent of €570 and total rent of €695 including costs, higher than the other Belgian regional averages in that study. This supports the affordability concern but does not independently establish a ranking of every Belgian city. The model uses an illustrative €550 monthly all-in affordability ceiling, with €100 assumed pass-through service charges. [Kotkompas pricing update](https://www.diggitstudentlife.eu/sites/default/files/2025-11/Kotkompas%20Pricing%20Update%202025Q4%20-%20E_0.pdf).

Municipal requirements must be handled separately. City of Brussels recommendations identify a 20% social student housing share from 30 created student units. Ixelles announced a 15% moderate-rent share for large kot developments. The workbook records these distinct policies, their scope and sources; it leaves the other 17 sourced policy percentages blank. A default 25% planning assumption is editable per municipality and never represented as verified law. No selected project has verified legal approval. Social allocation, eligible tenants and legally permitted rent may differ from the course affordability definition. [City recommendations](https://www.bruxelles.be/sites/default/files/bxl/Recommandations_logements_etudiants_VBX_17juin2021.pdf), [Ixelles announcement, municipal magazine March–April 2025](https://www.ixelles.be/site/downloads/publications/infoxl-2014-/infoxl20250304.pdf).

The decision variables are eight binary project selections. The model enumerates all 256 combinations, requires total investment within budget, at least 80 affordable rooms, nonnegative total NPV and a passing planning/rent check for every selected project. It first maximizes affordable rooms and then maximizes NPV among tied portfolios. At saved assumptions, 75 combinations are feasible. A municipality’s surplus cannot compensate for a failing project in another municipality. The shortlist and KPI formulas recalculate when inputs change; manual selections allow comparison against another portfolio.

Investment costs occur at year zero; years 1–20 contain unlevered operating cash flow. Terminal proceeds capitalize year 21 NOI at 5.5%, less 2% sale costs. Rent and operating costs grow by 2% annually. Service charges are pass-through and excluded from investor revenue. Financing, corporate tax, grants and construction delays are outside this iteration; development-cost inputs are assumed to include purchase, fees, relevant acquisition taxes and contingency.

The agile process delivered and saved a usable baseline first, then a sourced policy version, then the full DCF, portfolio search and stress tests. Each evaluation directed the next increment: unknown policy differences prompted research; financial incompleteness prompted cash flows; the optimum prompted downside evaluation. Formula caches, individual project cash-flow reconciliation, all portfolio totals, search optimality and base-scenario consistency passed automated verification. Excel recalculates on opening; Excel’s rendering and calculation engine were not available for an independent application-level check.

The next sprint should obtain current project-specific municipal requirements and written applicability, agree social/moderate allocation with suitable partners, obtain contractor quotes and test demand by campus and travel time. Replace the corresponding assumptions, reevaluate the shortlist and record the decision in the Iterations sheet. The team should approve this information-gathering sprint before committing investment capital.
