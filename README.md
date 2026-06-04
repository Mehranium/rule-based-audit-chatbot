# Rule-Based Audit Chatbot

A rule-based chatbot for querying internal audit data using natural language. Built for an internal audit department context, it translates structured natural language questions into pandas DataFrame queries against a relational Excel database.

## Overview

Auditors can ask **54 pre-defined question patterns** (customizable) in plain English. The chatbot matches the input against regex patterns and returns answers by querying four linked Excel files — effectively performing SQL-like operations without a database engine.

## Architecture

```
User question (natural language)
        │
        ▼
Regex pattern matching (54 patterns)
        │
        ▼
Pandas query (JOIN + WHERE equivalent)
        │
        ▼
Formatted markdown answer
```

### Data model (4 Excel files, relational via `Report File Name`)

| File | Description |
|---|---|
| `Main.xlsx` | One row per audit — company, country, year, auditors, KF/EVP responsibilities |
| `Introduction.xlsx` | Audit metadata — start/end dates, duration, participants |
| `Finding.xlsx` | Audit findings — category, title, finding text, risk, recommendation |
| `Suggestion.xlsx` | Post-audit suggestions — issue title, issue, suggestion |

## Question categories

- **Main table**: company locations, auditor lookups, KF/EVP responsibilities, country rankings
- **Finding table**: keyword search across findings/risks/recommendations, top-N frequencies, company-specific views, composite filters
- **Introduction table**: audit dates, durations, participants, date-range queries
- **Suggestion table**: issue titles, issues, suggestions — by company or keyword

## Two interfaces

| File | Interface |
|---|---|
| `Rule_Based_Streamlit.py` | Web UI via Streamlit |
| `Rule-Based CLI.py` | Command-line chatbot loop |

## Setup

1. Install dependencies:
   ```bash
   pip install streamlit pandas openpyxl
   ```

2. Place your four Excel files in a `data/` folder:
   ```
   data/
     Main.xlsx
     Introduction.xlsx
     Finding.xlsx
     Suggestion.xlsx
   ```

3. Run:
   ```bash
   # Streamlit web UI
   streamlit run Rule_Based_Streamlit.py

   # CLI
   python "Rule-Based CLI.py"
   ```

## Example questions

```
Which companies are located in Germany?
Who were the auditors for Würth USA Inc. in 2021?
Which companies have "inventory" in the finding title from 2020 to 2021?
Show all findings for a company with "IT" in the finding category in 2022.
List the top 5 frequent finding categories from 2020 to 2022.
What are the issues, issue titles, and suggestions for a company in 2021?
Which companies had audit durations of more than 15 days from 2020 to 2022?
```

## Project evolution: From rule-based to LLM-powered

This project went through two distinct phases, each a deliberate engineering decision.

### Phase 1 — Rule-based chatbot (this repository)

The first version was built under practical constraints: no external API access, no LLM integration, and a need for fast, predictable, explainable answers for non-technical auditors.

The rule-based approach was the right call at that stage:
- **Zero latency** — answers return instantly from in-memory DataFrames
- **Full explainability** — every answer traces back to an exact regex pattern and a pandas filter
- **No API cost or dependency** — runs entirely offline on local Excel files
- **Forces deep domain understanding** — designing 54 query patterns required mapping every question auditors actually ask, which produced a precise mental model of the data

The limitation was coverage: any question phrased outside the 54 patterns returned a fallback message. For a controlled internal tool with trained users, this was acceptable.

### Phase 2 — LLM agent with Claude Sonnet 4.6 (Langdock)

After Würth Group signed a contract with [Langdock](https://www.langdock.com/), the same four Excel files were migrated into a Langdock agent powered by **Claude Sonnet 4.6**.

The deep domain knowledge built during Phase 1 — the data model, the column relationships, the edge cases in company name spelling, the query patterns auditors actually need — translated directly into a precise, effective system prompt for the LLM agent.

The result: auditors can now ask **any question in any phrasing**, including multi-step and comparative questions that no rule-based system could handle, while the agent answers with the same accuracy as Phase 1 because the underlying data model was already well-understood.

### Key insight

> Building the rule-based version first was not wasted effort — it was the research phase. The constraints of writing 54 explicit patterns forced a level of domain understanding that made the LLM agent's system prompt precise and reliable from day one. The two phases are complementary, not competing.

This progression reflects a core principle in applied AI engineering: **choose the right tool for the current constraints, and design for the evolution you can see coming.**

## Key technical features

- **Tolerant string matching**: `canon()` normalizes accents, case, and spelling variants (e.g., ü → u)
- **Relational joins**: `Report File Name` links all four tables at load time
- **Streamlit caching**: `@st.cache_data` loads Excel files once per session
- **54 regex patterns** covering single-year and year-range variants for all query types
