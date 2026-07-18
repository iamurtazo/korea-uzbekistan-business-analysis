# Korea–Uzbekistan Business Expansion Analysis

A personal, long-running research project tracking Korean companies expanding into, investing in, or operating in Uzbekistan. The project is designed to identify active sectors, understand company profiles and investment activity, and discover potential employers for a Korean-speaking professional interested in working in Uzbekistan.

## Research goals

The dataset is intended to answer questions such as:

- Which sectors in Uzbekistan are attracting Korean companies?
- Which Korean companies have active, planned, completed, or withdrawn projects?
- How much are these companies investing, and through which investment structures?
- Where in Uzbekistan are their projects located?
- Which companies have local offices, career pages, or active recruitment?
- What are the companies' establishment years, employee counts, revenue, and net profit?

## Sectors covered

- Automotive and auto parts
- Semiconductors and electronics
- Healthcare and pharmaceuticals
- Energy, oil and gas, and mining
- Textile and cotton value chains
- Construction and infrastructure
- IT, software, and telecommunications
- Banking and finance
- Agriculture and food processing
- Logistics and trade
- Education

The sector list can be expanded as new patterns emerge.

## Repository structure

```text
.
├── data/
│   └── korea_uzbekistan_companies.xlsx
├── docs/
│   ├── search_strategy.md
│   └── sources.md
├── scripts/
│   └── generate_template.py
├── .gitignore
├── README.md
└── requirements.txt
```

## Excel workbook

The main tracker is `data/korea_uzbekistan_companies.xlsx`. It contains four worksheets:

1. **Companies** — one row per company or Uzbekistan project, with company profile, investment, location, status, careers, and source fields.
2. **Data Dictionary** — definitions and instructions for every column.
3. **Sources** — a dated log of individual reports, articles, disclosures, and websites consulted.
4. **Search Log** — a record of searches performed, including language, search engine, query, sector, and companies found.

The `Companies` sheet includes dropdowns for sector, investment type, project status, and source type. The header is frozen to make a growing dataset easier to navigate.

## Local setup

The project uses Python 3 and `openpyxl` to generate the workbook. Create an isolated environment and install the pinned dependencies:

```sh
python3 -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
```

Generate the workbook:

```sh
.venv/bin/python scripts/generate_template.py
```

The generator deliberately refuses to overwrite an existing workbook. This protects manually collected research data. If a completely new blank workbook is required, first preserve or deliberately remove the existing file.

## Research workflow

1. Choose one sector and a small set of queries from `docs/search_strategy.md`.
2. Record the research session in the workbook's **Search Log** sheet.
3. Add one row to **Companies** for each distinct company/project relationship.
4. Record all supporting pages in **Sources** and place the most relevant URLs in the company row.
5. Prefer official disclosures and government sources; cross-check Korean and Uzbek reporting.
6. Mark uncertain, estimated, announced, or conflicting figures clearly in **Notes**.
7. Revisit older rows periodically and update **Last Updated**.

A manageable cadence is one 30–60 minute research session per week, focusing on one sector at a time.

## Source priority

Recommended order of reliability:

1. Korean or Uzbek government records and official statistics
2. DART disclosures and company annual reports
3. Company websites and official press releases
4. KOTRA country and market reports
5. Established Korean and Uzbek news organizations
6. LinkedIn, job boards, and secondary databases

See `docs/search_strategy.md` for English and Korean keywords and `docs/sources.md` for the reference source list.

## Data quality principles

- Do not record a figure without retaining its source.
- Keep the original currency and reporting year clear in `Notes` when appropriate.
- Distinguish announced/committed investment from money actually invested.
- Distinguish global company employees from Uzbekistan-local employees.
- Do not treat an MOU as an operational project.
- Use `Unknown` rather than guessing.
- Preserve historical and withdrawn projects because they are still useful for sector analysis.

## Job-search use

For each company, check its official careers page, LinkedIn presence, and Uzbekistan job boards such as `hh.uz` and `jobs.uz`. Local hiring activity can both confirm that an operation is active and reveal opportunities where Korean-language ability is valuable.

## Disclaimer

This is personal research for career planning, not investment advice. Public figures may be approximate, outdated, reported in different currencies, or based on announcements that were never completed. Verify important information before relying on it.
