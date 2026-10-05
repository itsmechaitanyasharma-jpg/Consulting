# India Infusion Pumps: Product Classification (B. Braun, Fresenius Kabi)

- `Infusion_Pump_Classification_India.xlsx` contains five tabs:
  - **Slide Matrix**: the corrected version of the sample slide
  - **Product Evidence**: quoted proof and URLs for each product
  - **Revenue**: revenue cross-check, with an editable FX input
  - **Clean URL Map**: which URLs to keep, add or drop for the client
  - **Method & Caveats**: assumptions and limitations
- `evidence/`: screenshots of each India product page, captured October 2026
- `build_matrix.py`: script that regenerates the workbook

Colour legend: blue = official India website · peach = Indian distributor only (incl. IndiaMart) · grey = not disclosed.

## Slide 2/8 - Terumo, Mindray, ICU Medical

`Infusion_Pump_Classification_India_Slide2.xlsx` (built by `build_matrix_slide2.py`) uses the same tabs, criteria and colour legend as slide 1.
Revenue uses official sources only (Terumo FY2025 results, Mindray 2025 annual report filed on SZSE/CNINFO, ICU Medical Q4/FY2025 release).
USD figures use ECB 5-Oct-2026 reference rates crossed via EUR. Evidence screenshots and brochure pages are in `evidence/` (prefixes Terumo_, Mindray_, ICU_, Honmed_, Medikabazaar_, Graseby_).
