
import streamlit as st
import pandas as pd
import re
import unicodedata

# =============================
# Configuration
# =============================
st.set_page_config(page_title="Audit Chatbot", page_icon=":robot_face:")

# --- File paths ---
# Update these paths to point to your local Excel data files.
main_path = "data/Main.xlsx"
intro_path = "data/Introduction.xlsx"
finding_path = "data/Finding.xlsx"
suggestion_path = "data/Suggestion.xlsx"

# =============================
# Canonicalization helpers
# =============================
def strip_accents(s: str) -> str:
    """Remove accents/diacritics (e.g., 'ü' -> 'u')."""
    return ''.join(
        c for c in unicodedata.normalize('NFKD', str(s))
        if not unicodedata.combining(c)
    )

def canon(s: str) -> str:
    """
    Canonicalize strings for tolerant matching:
      - lowercase
      - remove accents
      - unify 'wuerth' and 'würth' -> 'wurth'
    """
    if s is None:
        return ''
    s0 = str(s).strip().lower()
    s1 = strip_accents(s0)              # 'würth' -> 'wurth'
    s2 = s1.replace('wuerth', 'wurth')  # 'wuerth' -> 'wurth'
    return s2

# =============================
# Utility formatting helpers
# =============================
def split_people(s: str):
    """Split a free-text list of names in 'Auditors' using common separators."""
    if pd.isna(s) or str(s).strip() == '':
        return []
    s0 = str(s)
    s0 = re.sub(r'\s*(?:/|&| and | und |;|\|)\s*', ',', s0, flags=re.IGNORECASE)
    parts = [p.strip() for p in s0.split(',') if p.strip()]
    return parts

def format_list(items):
    """Simple bullet list (unique, sorted, non-empty values) for markdown."""
    cleaned = [str(x).strip() for x in items if pd.notna(x) and str(x).strip()]
    if not cleaned:
        return "*(none found)*"
    return "\n".join(f"- {x}" for x in sorted(set(cleaned)))

def format_company_year_lines(df: pd.DataFrame, company_col: str, year_col: str) -> str:
    rows = (df.dropna(subset=[company_col, year_col])
              [[company_col, year_col]]
              .drop_duplicates()
              .sort_values([year_col, company_col]))
    if rows.empty:
        return ""
    return "\n".join(f"- {c} ({int(y)})" for c, y in rows.itertuples(index=False))

def format_title_year_lines(df: pd.DataFrame, title_col: str, year_col: str) -> str:
    rows = (df.dropna(subset=[title_col, year_col])
              [[title_col, year_col]]
              .drop_duplicates()
              .sort_values([year_col, title_col]))
    if rows.empty:
        return ""
    return "\n".join(f"- {t} ({int(y)})" for t, y in rows.itertuples(index=False))

def format_enriched_rows(df_rows: pd.DataFrame, value_col: str) -> str:
    """
    Render rows as:
    - <Finding Category> | <Finding Title>: <value_col text>
    """
    if df_rows.empty:
        return "*(none found)*"
    lines = []
    for _, r in df_rows.iterrows():
        cat = str(r.get('Finding Category', '')).strip()
        ttl = str(r.get('Finding Title', '')).strip()
        val = str(r.get(value_col, '')).strip()
        if not val:
            continue
        left = " | ".join([x for x in [cat, ttl] if x])
        lines.append(f"- {left}: {val}" if left else f"- {val}")
    return "\n".join(lines) if lines else "*(none found)*"

def format_issue_suggestion_year_lines(df: pd.DataFrame, issue_title_col: str, issue_col: str, sugg_col: str, year_col: str) -> str:
    rows = (df.dropna(subset=[year_col])
              [[issue_title_col, issue_col, sugg_col, year_col]]
              .drop_duplicates()
              .sort_values([year_col, issue_title_col]))
    if rows.empty:
        return ""
    lines = []
    for ttl, iss, sgg, yr in rows.itertuples(index=False):
        ttl = ttl or ""
        iss = iss or ""
        sgg = sgg or ""
        lines.append(f"- {ttl} ({int(yr)})\n  Issue: {iss}\n  Suggestion: {sgg}")
    return "\n".join(lines)

# =============================
# Merge helpers
# =============================
def coalesce_columns(df: pd.DataFrame, preferred: str, fallback: str, out: str = None):
    """
    df[out] = df[preferred] if non-empty else df[fallback]
    If preferred not present, just take fallback.
    If out is None -> write back to 'preferred' column name.
    """
    target = out or preferred
    if preferred in df.columns:
        lhs = df[preferred]
        rhs = df[fallback] if fallback in df.columns else pd.Series([None]*len(df))
        df[target] = lhs.where(lhs.notna() & (lhs.astype(str).str.strip() != ''), rhs)
    else:
        df[target] = df[fallback] if fallback in df.columns else pd.Series([None]*len(df))
    return df

# =============================
# Data loading with caching
# =============================
@st.cache_data(show_spinner=True)
def load_data():
    main = pd.read_excel(main_path)
    intro = pd.read_excel(intro_path)
    finding = pd.read_excel(finding_path)
    suggestion = pd.read_excel(suggestion_path)

    # Normalize year in Main
    if 'Audit Year' in main.columns:
        main['Audit Year'] = pd.to_numeric(main['Audit Year'], errors='coerce').fillna(0).astype(int)
    else:
        main['Audit Year'] = 0

    # --- Finding: coalesce Company and Year from Main (robust) ---
    finding = finding.merge(
        main[['Report File Name', 'Company', 'Audit Year']].rename(
            columns={'Company': 'Company_from_main', 'Audit Year': 'AuditYear_from_main'}
        ),
        on='Report File Name',
        how='left'
    )
    finding = coalesce_columns(finding, preferred='Company', fallback='Company_from_main', out='Company')
    finding['Year'] = pd.to_numeric(finding.get('Year'), errors='coerce')
    finding['AuditYear_from_main'] = pd.to_numeric(finding.get('AuditYear_from_main'), errors='coerce')
    finding['Year'] = finding['Year'].fillna(finding['AuditYear_from_main']).fillna(0).astype(int)
    for c in ['Company_from_main', 'AuditYear_from_main']:
        if c in finding.columns:
            finding.drop(columns=[c], inplace=True)

    # --- Introduction: coalesce Company and Audit Year from Main, and parse dates ---
    intro = intro.merge(
        main[['Report File Name', 'Company', 'Audit Year']].rename(
            columns={'Company': 'Company_from_main', 'Audit Year': 'AuditYear_from_main'}
        ),
        on='Report File Name',
        how='left'
    )
    intro = coalesce_columns(intro, preferred='Company', fallback='Company_from_main', out='Company')
    if 'Audit Year' not in intro.columns:
        intro['Audit Year'] = intro['AuditYear_from_main']
    else:
        intro['Audit Year'] = pd.to_numeric(intro['Audit Year'], errors='coerce').fillna(
            pd.to_numeric(intro['AuditYear_from_main'], errors='coerce')
        ).astype('Int64')
    for c in ['Company_from_main', 'AuditYear_from_main']:
        if c in intro.columns:
            intro.drop(columns=[c], inplace=True)

    # Parse start/end to DTS for range filters
    intro['Audit Start Date_dt'] = pd.to_datetime(intro.get('Audit Start Date', ''), errors='coerce', dayfirst=True)
    intro['Audit End Date_dt']   = pd.to_datetime(intro.get('Audit End Date', ''), errors='coerce', dayfirst=True)
    # If 'Audit Year' still missing, fallback to start-year
    if 'Audit Year' in intro.columns:
        intro['Audit Year'] = intro['Audit Year'].fillna(intro['Audit Start Date_dt'].dt.year)

    # --- Suggestion: ensure Company + Audit Year via Main ---
    if ('Company' not in suggestion.columns) or ('Audit Year' not in suggestion.columns):
        suggestion = suggestion.merge(
            main[['Report File Name', 'Company', 'Audit Year']],
            on='Report File Name',
            how='left'
        )

    # --- Precompute normalized cols for tolerant matching ---
    finding['Company_norm'] = finding['Company'].fillna('').map(canon)
    main['Company_norm']    = main.get('Company', pd.Series(dtype=str)).fillna('').map(canon)
    main['Country_norm']    = main.get('Country', pd.Series(dtype=str)).fillna('').map(canon)
    main['KF_norm']         = main.get('KF Responsibility', pd.Series(dtype=str)).fillna('').map(canon)
    main['EVP_SVP_norm']    = main.get('EVP/SVP Responsibility', pd.Series(dtype=str)).fillna('').map(canon)

    return main, intro, finding, suggestion

main, intro, finding, suggestion = load_data()

# =============================
# Core: Answer function (mirrors CLI)
# =============================
def answer_question(user_input, finding_df, main_df, intro_df, suggestion_df):
    patterns = [
        # ==========================================
        # MAIN TABLE QUERIES
        # ==========================================
        (r'which companies are located in (.+?)\??$', 'main_companies_in_country', 'main'),
        (r'which country is (.+?) located in\??$', 'main_country_of_company', 'main'),
        (r'who is the kf responsible for (.+?)\??$', 'main_kf_of_company', 'main'),
        (r'who is the evp/svp responsible for (.+?)\??$', 'main_evp_of_company', 'main'),
        (r'who were the auditors for (.+?) in (\d{4})\??$', 'main_auditors_of_company_year', 'main'),
        (r'which companies have (.+?) as their kf\??$', 'main_companies_by_kf_person', 'main'),
        (r'which companies have (.+?) as their evp/svp\??$', 'main_companies_by_evp_person', 'main'),
        (r'which companies were audited by (.+?) in (\d{4})\??$', 'main_companies_by_auditor_year', 'main'),
        (r'which companies were audited by (.+?) from (\d{4}) to (\d{4})\??$', 'main_companies_by_auditor_range', 'main'),
        (r'how many companies from (.+?) were audited in (\d{4})\??$', 'main_count_companies_country_year', 'main'),
        (r'how many companies from (.+?) were audited from (\d{4}) to (\d{4})\??$', 'main_count_companies_country_range', 'main'),
        (r'list the top (\d+) auditors by number of audits from (\d{4}) to (\d{4})\.?$', 'main_topn_auditors_range', 'main'),
        (r'list the top (\d+) countries by number of audited companies from (\d{4}) to (\d{4})\.?$', 'main_topn_countries_range', 'main'),

        # ==========================================
        # FINDING TABLE QUERIES (keyword -> companies)
        # ==========================================
        (r'which companies have ["“”\']?(.+?)["“”\']? in the findings in (\d{4})\??$', 'Finding', 'single_year'),
        (r'which companies have ["“”\']?(.+?)["“”\']? in the findings from (\d{4}) to (\d{4})\??$', 'Finding', 'year_range'),
        (r'which companies have ["“”\']?(.+?)["“”\']? in the risk in (\d{4})\??$', 'Risk', 'single_year'),
        (r'which companies have ["“”\']?(.+?)["“”\']? in the risk from (\d{4}) to (\d{4})\??$', 'Risk', 'year_range'),
        (r'which companies have ["“”\']?(.+?)["“”\']? in the recommendation in (\d{4})\??$', 'Recommendation', 'single_year'),
        (r'which companies have ["“”\']?(.+?)["“”\']? in the recommendation from (\d{4}) to (\d{4})\??$', 'Recommendation', 'year_range'),
        (r'which companies have ["“”\']?(.+?)["“”\']? in the finding category in (\d{4})\??$', 'Finding Category', 'single_year'),
        (r'which companies have ["“”\']?(.+?)["“”\']? in the finding category from (\d{4}) to (\d{4})\??$', 'Finding Category', 'year_range'),
        (r'which companies have ["“”\']?(.+?)["“”\']? in the finding title in (\d{4})\??$', 'Finding Title', 'single_year'),
        (r'which companies have ["“”\']?(.+?)["“”\']? in the finding title from (\d{4}) to (\d{4})\??$', 'Finding Title', 'year_range'),

        # ==========================================
        # NEW: Company-specific (no keyword)
        # ==========================================
        (r'show (?:me )?all (finding categories|finding titles|findings|risks|recommendations) (?:of|for) (.+?) in (\d{4})\.?$',
         'finding_company_all_single', 'finding_company_all'),
        (r'show (?:me )?all (finding categories|finding titles|findings|risks|recommendations) (?:of|for) (.+?) from (\d{4}) to (\d{4})\.?$',
         'finding_company_all_range', 'finding_company_all'),

        # ==========================================
        # Composite (company + keyword + field + time)
        # ==========================================
        (r'show (?:me )?all (findings?|risks?|recommendations?) (?:of|for) (.+?) (?:with|where) ["“”\']?(.+?)["“”\']?\s+in the (finding category|finding title|finding|risk|recommendation) in (\d{4})\.?$',
         'composite', 'company_keyword_year'),
        (r'show (?:me )?all (findings?|risks?|recommendations?) (?:of|for) (.+?) (?:with|where) ["“”\']?(.+?)["“”\']?\s+in the (finding category|finding title|finding|risk|recommendation) from (\d{4}) to (\d{4})\.?$',
         'composite', 'company_keyword_year_range'),

        # ==========================================
        # Top N frequent categories/titles
        # ==========================================
        (r'list the top (\d+) frequent finding categories in (\d{4})\.?$', 'Finding Category', 'topn_category_single_year'),
        (r'list the top (\d+) frequent finding categories from (\d{4}) to (\d{4})\.?$', 'Finding Category', 'topn_category_year_range'),
        (r'list the top (\d+) frequent finding titles in (\d{4})\.?$', 'Finding Title', 'topn_title_single_year'),
        (r'list the top (\d+) frequent finding titles from (\d{4}) to (\d{4})\.?$', 'Finding Title', 'topn_title_year_range'),

        # ==========================================
        # Top N companies by keyword count
        # ==========================================
        (r'list the top (\d+) companies by the number of (.+?) in finding category in (\d{4})\.?$', 'Finding Category', 'topn_company_single_year'),
        (r'list the top (\d+) companies by the number of (.+?) in finding category from (\d{4}) to (\d{4})\.?$', 'Finding Category', 'topn_company_year_range'),
        (r'list the top (\d+) companies by the number of (.+?) in finding title in (\d{4})\.?$', 'Finding Title', 'topn_company_single_year'),
        (r'list the top (\d+) companies by the number of (.+?) in finding title from (\d{4}) to (\d{4})\.?$', 'Finding Title', 'topn_company_year_range'),

        # ==========================================
        # Counting in Finding Category/Title
        # ==========================================
        (r'how many (.+?) in finding category were reported in (\d{4})\??$', 'Finding Category', 'count_single_year'),
        (r'how many (.+?) in finding category were reported from (\d{4}) to (\d{4})\??$', 'Finding Category', 'count_year_range'),
        (r'how many (.+?) in finding title were reported in (\d{4})\??$', 'Finding Title', 'count_single_year'),
        (r'how many (.+?) in finding title were reported from (\d{4}) to (\d{4})\??$', 'Finding Title', 'count_year_range'),

        # ==========================================
        # INTRODUCTION TABLE QUERIES
        # ==========================================
        (r'what were the start and end dates of the audit for (.+?) in (\d{4})\??$', 'intro_dates', 'intro'),
        (r'what was the audit duration for (.+?) in (\d{4})\??$', 'intro_duration', 'intro'),
        (r'who participated in the audit for (.+?) in (\d{4})\??$', 'intro_participants', 'intro'),
        (r'which companies had audits that started between (\d{2}\.\d{2}\.\d{4}) and (\d{2}\.\d{2}\.\d{4})\??$', 'intro_started_between', 'intro'),
        (r'which companies had audits that ended between (\d{2}\.\d{2}\.\d{4}) and (\d{2}\.\d{2}\.\d{4})\??$', 'intro_ended_between', 'intro'),
        (r'which companies had audit durations of less than (\d+) days from (\d{4}) to (\d{4})\??$', 'intro_duration_less_than', 'intro'),
        (r'which companies had audit durations of more than (\d+) days from (\d{4}) to (\d{4})\??$', 'intro_duration_more_than', 'intro'),

        # =========================================
        # SUGGESTION TABLE QUERIES
        # =========================================
        (r'what are the issue titles for (.+?) in (\d{4})\??$', 'suggestion_issue_titles_company_year', 'suggestion'),
        (r'what are the issue titles for (.+?) from (\d{4}) to (\d{4})\??$', 'suggestion_issue_titles_company_range', 'suggestion'),
        (r'what are the issues, issue titles, and suggestions for (.+?) in (\d{4})\??$', 'suggestion_issues_suggestions_company_year', 'suggestion'),
        (r'what are the issues, issue titles, and suggestions for (.+?) from (\d{4}) to (\d{4})\??$', 'suggestion_issues_suggestions_company_range', 'suggestion'),
        (r'which companies have ["“”\']?(.+?)["“”\']? in their issue titles from (\d{4}) to (\d{4})\??$', 'suggestion_companies_issue_titles_keyword_range', 'suggestion'),
        (r'which companies have ["“”\']?(.+?)["“”\']? in their issues from (\d{4}) to (\d{4})\??$', 'suggestion_companies_issues_keyword_range', 'suggestion'),
        (r'which companies have ["“”\']?(.+?)["“”\']? in their suggestions from (\d{4}) to (\d{4})\??$', 'suggestion_companies_suggestions_keyword_range', 'suggestion'),
        (r'how many companies have ["“”\']?(.+?)["“”\']? in their issue titles from (\d{4}) to (\d{4})\??$', 'suggestion_count_companies_issue_titles_keyword_range', 'suggestion'),
    ]

    sanitized = user_input.rstrip().rstrip('.!?').strip()

    for _, (pattern, tag, qtype) in enumerate(patterns, 1):
        match = re.fullmatch(pattern, sanitized, re.IGNORECASE)
        if not match:
            continue

        # ============================
        # MAIN TABLE HANDLERS
        # ============================
        if qtype == 'main':
            if tag == 'main_companies_in_country':
                country = match.group(1).strip()
                mask = main_df['Country_norm'].str.contains(canon(country), na=False)
                companies = main_df.loc[mask, 'Company'].dropna().unique()
                if len(companies) == 0:
                    return f"No companies found located in **{country}**."
                header = f"**Companies located in {country}:**"
                body = format_list(companies)
                return f"{header}\n\n{body}"

            elif tag == 'main_country_of_company':
                company = match.group(1).strip()
                mask = main_df['Company_norm'].str.contains(canon(company), na=False)
                countries = main_df.loc[mask, 'Country'].dropna().unique()
                if len(countries) == 0:
                    return f"No country found for company **'{company}'**."
                header = f"**Country for {company}:**"
                body = format_list(countries)
                return f"{header}\n\n{body}"

            elif tag == 'main_kf_of_company':
                company = match.group(1).strip()
                mask = main_df['Company_norm'].str.contains(canon(company), na=False)
                kfs = main_df.loc[mask, 'KF Responsibility'].dropna().unique()
                if len(kfs) == 0:
                    return f"No KF responsible found for company **'{company}'**."
                header = f"**KF responsible for {company}:**"
                body = format_list(kfs)
                return f"{header}\n\n{body}"

            elif tag == 'main_evp_of_company':
                company = match.group(1).strip()
                mask = main_df['Company_norm'].str.contains(canon(company), na=False)
                evps = main_df.loc[mask, 'EVP/SVP Responsibility'].dropna().unique()
                if len(evps) == 0:
                    return f"No EVP/SVP responsible found for company **'{company}'**."
                header = f"**EVP/SVP responsible for {company}:**"
                body = format_list(evps)
                return f"{header}\n\n{body}"

            elif tag == 'main_auditors_of_company_year':
                company = match.group(1).strip()
                year = int(match.group(2))
                mask = (main_df['Company_norm'].str.contains(canon(company), na=False)) & (main_df['Audit Year'] == year)
                auditors = main_df.loc[mask, 'Auditors'].dropna().unique()
                names = []
                for a in auditors:
                    names.extend(split_people(a))
                names = sorted(set(names))
                if len(names) == 0:
                    return f"No auditors found for **{company}** in **{year}**."
                header = f"**Auditors for {company} in {year}:**"
                body = format_list(names)
                return f"{header}\n\n{body}"

            elif tag == 'main_companies_by_kf_person':
                person = match.group(1).strip()
                mask = main_df['KF_norm'].str.contains(canon(person), na=False)
                companies = main_df.loc[mask, 'Company'].dropna().unique()
                if len(companies) == 0:
                    return f"No companies found with KF **'{person}'**."
                header = f"**Companies with {person} as KF:**"
                body = format_list(companies)
                return f"{header}\n\n{body}"

            elif tag == 'main_companies_by_evp_person':
                person = match.group(1).strip()
                mask = main_df['EVP_SVP_norm'].str.contains(canon(person), na=False)
                companies = main_df.loc[mask, 'Company'].dropna().unique()
                if len(companies) == 0:
                    return f"No companies found with EVP/SVP **'{person}'**."
                header = f"**Companies with {person} as EVP/SVP:**"
                body = format_list(companies)
                return f"{header}\n\n{body}"

            elif tag == 'main_companies_by_auditor_year':
                auditor = match.group(1).strip()
                year = int(match.group(2))
                rows = main_df[main_df['Audit Year'] == year]
                def row_has_auditor(s):
                    return any(canon(auditor) in canon(p) for p in split_people(s))
                mask = rows['Auditors'].fillna('').map(row_has_auditor)
                companies = rows.loc[mask, 'Company'].dropna().unique()
                if len(companies) == 0:
                    return f"No companies audited by **'{auditor}'** in **{year}**."
                header = f"**Companies audited by {auditor} in {year}:**"
                body = format_list(companies)
                return f"{header}\n\n{body}"

            elif tag == 'main_companies_by_auditor_range':
                auditor = match.group(1).strip()
                y1 = int(match.group(2)); y2 = int(match.group(3))
                rows = main_df[main_df['Audit Year'].between(y1, y2)]
                def row_has_auditor(s):
                    return any(canon(auditor) in canon(p) for p in split_people(s))
                mask = rows['Auditors'].fillna('').map(row_has_auditor)
                result = rows.loc[mask, ['Company', 'Audit Year']]
                if result.empty:
                    return f"No companies audited by '{auditor}' from {y1} to {y2}."
                header = f"Companies audited by {auditor} from {y1} to {y2}:"
                body = format_company_year_lines(result, 'Company', 'Audit Year')
                return f"{header}\n{body}"

            elif tag == 'main_count_companies_country_year':
                country = match.group(1).strip()
                year = int(match.group(2))
                mask = (main_df['Country_norm'].str.contains(canon(country), na=False)) & (main_df['Audit Year'] == year)
                count = main_df.loc[mask, 'Company'].dropna().nunique()
                return f"**Count:** {count} companies from **{country}** audited in **{year}**."

            elif tag == 'main_count_companies_country_range':
                country = match.group(1).strip()
                y1 = int(match.group(2)); y2 = int(match.group(3))
                mask = (main_df['Country_norm'].str.contains(canon(country), na=False)) & (main_df['Audit Year'].between(y1, y2))
                count = main_df.loc[mask, 'Company'].dropna().nunique()
                return f"**Count:** {count} companies from **{country}** audited from **{y1}** to **{y2}**."

            elif tag == 'main_topn_auditors_range':
                n = int(match.group(1)); y1 = int(match.group(2)); y2 = int(match.group(3))
                rows = main_df[main_df['Audit Year'].between(y1, y2)]
                counts = {}
                display = {}
                for s in rows['Auditors'].fillna(''):
                    for name in split_people(s):
                        key = canon(name)
                        counts[key] = counts.get(key, 0) + 1
                        display[key] = name
                if not counts:
                    return f"No audits found from **{y1}** to **{y2}**."
                ranked = sorted(counts.items(), key=lambda kv: kv[1], reverse=True)[:n]
                header = f"**Top {n} auditors by number of audits from {y1} to {y2}:**"
                body = "\n".join(f"{i+1}. {display[k]} ({v})" for i, (k, v) in enumerate(ranked))
                return f"{header}\n\n{body}"

            elif tag == 'main_topn_countries_range':
                n = int(match.group(1)); y1 = int(match.group(2)); y2 = int(match.group(3))
                rows = main_df[main_df['Audit Year'].between(y1, y2)]
                grouped = rows.dropna(subset=['Country', 'Company']).groupby('Country')['Company'].nunique().sort_values(ascending=False).head(n)
                if grouped.empty:
                    return f"No audited companies found from **{y1}** to **{y2}**."
                header = f"**Top {n} countries by number of audited companies from {y1} to {y2}:**"
                body = "\n".join(f"{i+1}. {country} ({int(cnt)})" for i, (country, cnt) in enumerate(grouped.items()))
                return f"{header}\n\n{body}"

        # ============================
        # FINDING TABLE HANDLERS
        # ============================
        # Keyword-based company lists
        if qtype in ('single_year', 'year_range') and tag in ('Finding', 'Risk', 'Recommendation', 'Finding Category', 'Finding Title'):
            keyword = match.group(1).strip()
            if qtype == 'single_year':
                year = int(match.group(2))
                mask = (finding_df[tag].astype(str).str.contains(keyword, case=False, na=False)) & (finding_df['Year'] == year)
                rf = finding_df.loc[mask, ['Report File Name']].drop_duplicates()
                joined = rf.merge(main_df[['Report File Name', 'Company']], on='Report File Name', how='left')
                companies = joined['Company'].dropna().unique()
                if companies.size == 0:
                    return f"No companies found with {tag.lower()} '{keyword}' in {year}."
                header = f"Companies with {tag.lower()} '{keyword}' in {year}:"
                body = "\n".join(f"- {c}" for c in sorted(companies))
                return f"{header}\n{body}"
            else:
                y1 = int(match.group(2)); y2 = int(match.group(3))
                mask = (finding_df[tag].astype(str).str.contains(keyword, case=False, na=False)) & finding_df['Year'].between(y1, y2)
                rf_year = finding_df.loc[mask, ['Report File Name', 'Year']]
                joined = rf_year.merge(main_df[['Report File Name', 'Company']], on='Report File Name', how='left')
                if joined.empty:
                    return f"No companies found with {tag.lower()} '{keyword}' from {y1} to {y2}."
                header = f"Companies with {tag.lower()} '{keyword}' from {y1} to {y2}:"
                body = format_company_year_lines(joined, 'Company', 'Year')
                return f"{header}\n{body}"

        # Composite: enriched rows for company + keyword + field + time
        if qtype in ('company_keyword_year', 'company_keyword_year_range') and tag == 'composite':
            show_type = match.group(1).strip().lower()
            company_raw = match.group(2).strip()
            keyword = match.group(3).strip()
            field_raw = match.group(4).strip().lower()

            field_map = {
                'finding category': 'Finding Category',
                'finding title': 'Finding Title',
                'finding': 'Finding',
                'risk': 'Risk',
                'recommendation': 'Recommendation',
            }
            column_filter = field_map[field_raw]

            show_map = {
                'findings': 'Finding', 'finding': 'Finding',
                'risks': 'Risk', 'risk': 'Risk',
                'recommendations': 'Recommendation', 'recommendation': 'Recommendation',
            }
            value_col = show_map[show_type]

            company_mask = finding_df['Company_norm'].str.contains(canon(company_raw), na=False)
            keyword_mask = finding_df[column_filter].astype(str).str.contains(keyword, case=False, na=False)

            if qtype == 'company_keyword_year':
                year = int(match.group(5))
                time_mask = (finding_df['Year'] == year)
                mask = company_mask & keyword_mask & time_mask
                cols = ['Finding Category', 'Finding Title', value_col]
                rows = finding_df.loc[mask, cols].dropna(subset=[value_col]).drop_duplicates()
                if rows.empty:
                    return f"No {value_col.lower()} found for {company_raw} with '{keyword}' in the {column_filter.lower()} in {year}."
                header = f"{value_col} for {company_raw} with '{keyword}' in the {column_filter.lower()} in {year}:"
                body = format_enriched_rows(rows, value_col)
                return f"{header}\n{body}"
            else:
                y1 = int(match.group(5)); y2 = int(match.group(6))
                time_mask = finding_df['Year'].between(y1, y2)
                mask = company_mask & keyword_mask & time_mask
                cols = ['Finding Category', 'Finding Title', 'Year', value_col]
                rows = finding_df.loc[mask, cols].dropna(subset=[value_col]).drop_duplicates()
                if rows.empty:
                    return f"No {value_col.lower()} found for {company_raw} with '{keyword}' in the {column_filter.lower()} from {y1} to {y2}."
                header = f"{value_col} for {company_raw} with '{keyword}' in the {column_filter.lower()} from {y1} to {y2}:"
                def _fmt_row(r):
                    left = " | ".join([x for x in [r.get('Finding Category', ''), r.get('Finding Title', '')] if x])
                    val = str(r.get(value_col, '')).strip()
                    yr = r.get('Year', None)
                    suffix = f" ({int(yr)})" if pd.notna(yr) else ""
                    return f"- {left}: {val}{suffix}" if left else f"- {val}{suffix}"
                body = "\n".join(_fmt_row(r) for _, r in rows.iterrows())
                return f"{header}\n{body}"

        # NEW: Company-specific (no keyword) handlers
        if qtype == 'finding_company_all' and tag in ('finding_company_all_single', 'finding_company_all_range'):
            item_type = match.group(1).strip().lower()  # "finding categories|finding titles|findings|risks|recommendations"
            company_raw = match.group(2).strip()
            group_map = {
                'finding categories': ('Finding Category',),
                'finding titles': ('Finding Title',),
                'findings': ('Finding Category', 'Finding Title', 'Finding'),
                'risks': ('Finding Category', 'Finding Title', 'Risk'),
                'recommendations': ('Finding Category', 'Finding Title', 'Recommendation'),
            }
            cols = group_map[item_type]
            company_mask = finding_df['Company_norm'].str.contains(canon(company_raw), na=False)

            if tag == 'finding_company_all_single':
                year = int(match.group(3))
                time_mask = (finding_df['Year'] == year)
                period_label = f"in {year}"
                add_year = False
            else:
                y1 = int(match.group(3)); y2 = int(match.group(4))
                time_mask = finding_df['Year'].between(y1, y2)
                period_label = f"from {y1} to {y2}"
                add_year = True

            mask = company_mask & time_mask

            if item_type in ('finding categories', 'finding titles'):
                col = cols[0]
                rows = finding_df.loc[mask, [col, 'Year']].dropna(subset=[col]).drop_duplicates()
                if rows.empty:
                    return f"No {item_type} found for {company_raw} {period_label}."
                header = f"{item_type.capitalize()} for {company_raw} {period_label}:"
                if add_year:
                    tmp = rows.rename(columns={col: 'Title'})
                    body = format_title_year_lines(tmp, 'Title', 'Year')
                else:
                    values = rows[col].dropna().unique()
                    body = "\n".join(f"- {v}" for v in sorted(values))
                return f"{header}\n{body}"

            value_col = cols[-1]
            select_cols = ['Finding Category', 'Finding Title', 'Year', value_col]
            rows = (finding_df.loc[mask, select_cols]
                    .dropna(subset=[value_col])
                    .drop_duplicates()
                    .sort_values(['Year', 'Finding Category', 'Finding Title']))
            if rows.empty:
                return f"No {item_type} found for {company_raw} {period_label}."

            header = f"{item_type.capitalize()} for {company_raw} {period_label}:"
            if add_year:
                def _fmt_row2(r):
                    left = " | ".join([x for x in [r.get('Finding Category', ''), r.get('Finding Title', '')] if x])
                    val = str(r.get(value_col, '')).strip()
                    yr = r.get('Year', None)
                    suffix = f" ({int(yr)})" if pd.notna(yr) else ""
                    return f"- {left}: {val}{suffix}" if left else f"- {val}{suffix}"
                body = "\n".join(_fmt_row2(r) for _, r in rows.iterrows())
            else:
                body = format_enriched_rows(rows.drop(columns=['Year']), value_col)
            return f"{header}\n{body}"

        # Top N frequent categories/titles
        if qtype in ('topn_category_single_year', 'topn_category_year_range', 'topn_title_single_year', 'topn_title_year_range'):
            n = int(match.group(1))
            if 'single_year' in qtype:
                year = int(match.group(2))
                mask = (finding_df['Year'] == year)
                period_label = f"in **{year}**"
            else:
                y1 = int(match.group(2)); y2 = int(match.group(3))
                mask = finding_df['Year'].between(y1, y2)
                period_label = f"from **{y1}** to **{y2}**"
            group_col = 'Finding Category' if 'category' in qtype else 'Finding Title'
            counts = finding_df.loc[mask, group_col].dropna().value_counts().head(n)
            if counts.empty:
                return f"No {group_col.lower()} found {period_label}."
            header = f"**Top {n} frequent {group_col.lower()}s {period_label}:**"
            body = "\n".join([f"{i+1}. {name} ({cnt})" for i, (name, cnt) in enumerate(counts.items())])
            return f"{header}\n\n{body}"

        # Top N companies by number of keyword hits
        if qtype in ('topn_company_single_year', 'topn_company_year_range') and tag in ('Finding Category', 'Finding Title'):
            n = int(match.group(1))
            keyword = match.group(2).strip()
            if qtype == 'topn_company_single_year':
                year = int(match.group(3))
                mask = (finding_df[tag].astype(str).str.contains(keyword, case=False, na=False)) & (finding_df['Year'] == year)
                period_label = f"in **{year}**"
            else:
                y1 = int(match.group(3)); y2 = int(match.group(4))
                mask = (finding_df[tag].astype(str).str.contains(keyword, case=False, na=False)) & (finding_df['Year'].between(y1, y2))
                period_label = f"from **{y1}** to **{y2}**"
            df = finding_df.loc[mask, ['Report File Name']].merge(
                main_df[['Report File Name', 'Company']], on='Report File Name', how='left'
            )
            counts = df['Company'].dropna().value_counts().head(n)
            if counts.empty:
                return f"No companies found with **'{keyword}'** in {tag.lower()} {period_label}."
            header = f"**Top {n} companies by number of '{keyword}' in {tag.lower()} {period_label}:**"
            body = "\n".join([f"{i+1}. {name} ({cnt})" for i, (name, cnt) in enumerate(counts.items())])
            return f"{header}\n\n{body}"

        # Counting queries (Finding Category/Title)
        if qtype in ('count_single_year', 'count_year_range') and tag in ('Finding Category', 'Finding Title'):
            keyword = match.group(1).strip()
            if qtype == 'count_single_year':
                year = int(match.group(2))
                mask = (finding_df[tag].astype(str).str.contains(keyword, case=False, na=False)) & (finding_df['Year'] == year)
                period_label = f"in **{year}**"
            else:
                y1 = int(match.group(2)); y2 = int(match.group(3))
                mask = (finding_df[tag].astype(str).str.contains(keyword, case=False, na=False)) & (finding_df['Year'].between(y1, y2))
                period_label = f"from **{y1}** to **{y2}**"
            count = finding_df.loc[mask].shape[0]
            return f"**Count:** {count} records with **'{keyword}'** in {tag.lower()} {period_label}."

        # ============================
        # INTRODUCTION TABLE HANDLERS
        # ============================
        if qtype == 'intro':
            if tag == 'intro_dates':
                company = match.group(1).strip()
                year = int(match.group(2))
                mask = (intro_df['Company'].fillna('').map(canon).str.contains(canon(company), na=False)) & (intro_df['Audit Year'] == year)
                rows = intro_df.loc[mask]
                if rows.empty:
                    return f"No audit dates found for {company} in {year}."
                lines = []
                for _, r in rows.iterrows():
                    start = r.get('Audit Start Date', '')
                    end = r.get('Audit End Date', '')
                    lines.append(f"- {r.get('Company', company)}: {start} to {end}")
                return f"Audit start and end dates for {company} in {year}:\n" + "\n".join(lines)

            elif tag == 'intro_duration':
                company = match.group(1).strip()
                year = int(match.group(2))
                mask = (intro_df['Company'].fillna('').map(canon).str.contains(canon(company), na=False)) & (intro_df['Audit Year'] == year)
                rows = intro_df.loc[mask]
                if rows.empty:
                    return f"No audit duration found for {company} in {year}."
                lines = []
                for _, r in rows.iterrows():
                    duration = r.get('Audit Duration (Day)', '')
                    lines.append(f"- {r.get('Company', company)}: {duration} days")
                return f"Audit duration for {company} in {year}:\n" + "\n".join(lines)

            elif tag == 'intro_participants':
                company = match.group(1).strip()
                year = int(match.group(2))
                mask = (intro_df['Company'].fillna('').map(canon).str.contains(canon(company), na=False)) & (intro_df['Audit Year'] == year)
                rows = intro_df.loc[mask]
                if rows.empty:
                    return f"No audit participants found for {company} in {year}."
                lines = []
                for _, r in rows.iterrows():
                    participants = r.get('Participants', '')
                    lines.append(f"- {r.get('Company', company)}: {participants}")
                return f"Audit participants for {company} in {year}:\n" + "\n".join(lines)

            elif tag == 'intro_started_between':
                date1 = pd.to_datetime(match.group(1), dayfirst=True, errors='coerce')
                date2 = pd.to_datetime(match.group(2), dayfirst=True, errors='coerce')
                if pd.isna(date1) or pd.isna(date2):
                    return "Invalid date format. Please use dd.mm.yyyy."
                mask = (intro_df['Audit Start Date_dt'] >= date1) & (intro_df['Audit Start Date_dt'] <= date2)
                rows = intro_df.loc[mask]
                companies = rows['Company'].dropna().unique()
                if not companies.size:
                    return f"No companies found with audits started between {match.group(1)} and {match.group(2)}."
                return f"Companies with audits started between {match.group(1)} and {match.group(2)}:\n" + "\n".join(f"- {c}" for c in sorted(companies))

            elif tag == 'intro_ended_between':
                date1 = pd.to_datetime(match.group(1), dayfirst=True, errors='coerce')
                date2 = pd.to_datetime(match.group(2), dayfirst=True, errors='coerce')
                if pd.isna(date1) or pd.isna(date2):
                    return "Invalid date format. Please use dd.mm.yyyy."
                mask = (intro_df['Audit End Date_dt'] >= date1) & (intro_df['Audit End Date_dt'] <= date2)
                rows = intro_df.loc[mask]
                companies = rows['Company'].dropna().unique()
                if not companies.size:
                    return f"No companies found with audits ended between {match.group(1)} and {match.group(2)}."
                return f"Companies with audits ended between {match.group(1)} and {match.group(2)}:\n" + "\n".join(f"- {c}" for c in sorted(companies))

            elif tag == 'intro_duration_less_than':
                n = int(match.group(1)); y1 = int(match.group(2)); y2 = int(match.group(3))
                mask = (pd.to_numeric(intro_df['Audit Duration (Day)'], errors='coerce') < n) & (intro_df['Audit Year'].between(y1, y2))
                rows = intro_df.loc[mask, ['Company', 'Audit Year']].dropna()
                if rows.empty:
                    return f"No companies found with audit durations of less than {n} days from {y1} to {y2}."
                rows = rows.drop_duplicates().sort_values(['Audit Year', 'Company'])
                lines = [f"- {c} ({int(y)})" for c, y in rows.itertuples(index=False)]
                return f"Companies with audit durations of less than {n} days from {y1} to {y2}:\n" + "\n".join(lines)

            elif tag == 'intro_duration_more_than':
                n = int(match.group(1)); y1 = int(match.group(2)); y2 = int(match.group(3))
                mask = (pd.to_numeric(intro_df['Audit Duration (Day)'], errors='coerce') > n) & (intro_df['Audit Year'].between(y1, y2))
                rows = intro_df.loc[mask, ['Company', 'Audit Year']].dropna()
                if rows.empty:
                    return f"No companies found with audit durations of more than {n} days from {y1} to {y2}."
                rows = rows.drop_duplicates().sort_values(['Audit Year', 'Company'])
                lines = [f"- {c} ({int(y)})" for c, y in rows.itertuples(index=False)]
                return f"Companies with audit durations of more than {n} days from {y1} to {y2}:\n" + "\n".join(lines)

        # ============================
        # SUGGESTION TABLE HANDLERS
        # ============================
        if qtype == 'suggestion':
            sug = suggestion_df  # already merged with Main to have Company + Audit Year

            if tag == 'suggestion_issue_titles_company_year':
                company = match.group(1).strip()
                year = int(match.group(2))
                mask = (sug['Company'].fillna('').map(canon).str.contains(canon(company), na=False)) & (sug['Audit Year'] == year)
                rows = sug.loc[mask]
                if rows.empty:
                    return f"No issue titles found for {company} in {year}."
                titles = rows['Issue Title'].dropna().unique()
                return f"Issue titles for {company} in {year}:\n" + "\n".join(f"- {t}" for t in titles)

            elif tag == 'suggestion_issue_titles_company_range':
                company = match.group(1).strip()
                y1 = int(match.group(2)); y2 = int(match.group(3))
                mask = (sug['Company'].fillna('').map(canon).str.contains(canon(company), na=False)) & (sug['Audit Year'].between(y1, y2))
                rows = sug.loc[mask, ['Issue Title', 'Audit Year']]
                if rows.empty:
                    return f"No issue titles found for {company} from {y1} to {y2}."
                header = f"Issue titles for {company} from {y1} to {y2}:"
                body = format_title_year_lines(rows, 'Issue Title', 'Audit Year')
                return f"{header}\n{body}"

            elif tag == 'suggestion_issues_suggestions_company_year':
                company = match.group(1).strip()
                year = int(match.group(2))
                mask = (sug['Company'].fillna('').map(canon).str.contains(canon(company), na=False)) & (sug['Audit Year'] == year)
                rows = sug.loc[mask]
                if rows.empty:
                    return f"No issues or suggestions found for {company} in {year}."
                lines = []
                for _, r in rows.iterrows():
                    lines.append(
                        f"- **Issue Title:** {r.get('Issue Title','')}\n"
                        f"  **Issue:** {r.get('Issue','')}\n"
                        f"  **Suggestion:** {r.get('Suggestion','')}"
                    )
                return f"Issues, issue titles, and suggestions for {company} in {year}:\n" + "\n\n".join(lines)

            elif tag == 'suggestion_issues_suggestions_company_range':
                company = match.group(1).strip()
                y1 = int(match.group(2)); y2 = int(match.group(3))
                mask = (sug['Company'].fillna('').map(canon).str.contains(canon(company), na=False)) & (sug['Audit Year'].between(y1, y2))
                rows = sug.loc[mask, ['Issue Title', 'Issue', 'Suggestion', 'Audit Year']]
                if rows.empty:
                    return f"No issues or suggestions found for {company} from {y1} to {y2}."
                header = f"Issues, issue titles, and suggestions for {company} from {y1} to {y2}:"
                body = format_issue_suggestion_year_lines(rows, 'Issue Title', 'Issue', 'Suggestion', 'Audit Year')
                return f"{header}\n{body}"

            elif tag == 'suggestion_companies_issue_titles_keyword_range':
                keyword = match.group(1).strip()
                y1 = int(match.group(2)); y2 = int(match.group(3))
                mask = (sug['Issue Title'].astype(str).str.contains(keyword, case=False, na=False)) & (sug['Audit Year'].between(y1, y2))
                rows = (sug.loc[mask, ['Company', 'Issue Title', 'Audit Year']]
                        .dropna(subset=['Company', 'Issue Title', 'Audit Year'])
                        .drop_duplicates()
                        .sort_values(['Audit Year', 'Company', 'Issue Title']))
                if rows.empty:
                    return f"No companies found with '{keyword}' in their issue titles from {y1} to {y2}."
                header = f"Companies with '{keyword}' in their issue titles from {y1} to {y2}:"
                lines = [f"- {c}: {t} ({int(y)})" for c, t, y in rows.itertuples(index=False)]
                return f"{header}\n" + "\n".join(lines)

            elif tag == 'suggestion_companies_issues_keyword_range':
                keyword = match.group(1).strip()
                y1 = int(match.group(2)); y2 = int(match.group(3))
                mask = (sug['Issue'].astype(str).str.contains(keyword, case=False, na=False)) & (sug['Audit Year'].between(y1, y2))
                rows = (sug.loc[mask, ['Company', 'Issue', 'Suggestion', 'Audit Year']]
                        .dropna(subset=['Company', 'Issue', 'Audit Year'])
                        .drop_duplicates()
                        .sort_values(['Audit Year', 'Company']))
                if rows.empty:
                    return f"No companies found with '{keyword}' in their issues from {y1} to {y2}."
                header = f"Companies with '{keyword}' in their issues from {y1} to {y2}:"
                lines = [f"- {c} ({int(y)}):\n  Issue: {i}\n  Suggestion: {s or ''}" for c, i, s, y in rows.itertuples(index=False)]
                return f"{header}\n" + "\n\n".join(lines)

            elif tag == 'suggestion_companies_suggestions_keyword_range':
                keyword = match.group(1).strip()
                y1 = int(match.group(2)); y2 = int(match.group(3))
                mask = (sug['Suggestion'].astype(str).str.contains(keyword, case=False, na=False)) & (sug['Audit Year'].between(y1, y2))
                rows = (sug.loc[mask, ['Company', 'Issue', 'Suggestion', 'Audit Year']]
                        .dropna(subset=['Company', 'Suggestion', 'Audit Year'])
                        .drop_duplicates()
                        .sort_values(['Audit Year', 'Company']))
                if rows.empty:
                    return f"No companies found with '{keyword}' in their suggestions from {y1} to {y2}."
                header = f"Companies with '{keyword}' in their suggestions from {y1} to {y2}:"
                lines = [f"- {c} ({int(y)}):\n  Issue: {i or ''}\n  Suggestion: {s}" for c, i, s, y in rows.itertuples(index=False)]
                return f"{header}\n" + "\n\n".join(lines)

            elif tag == 'suggestion_count_companies_issue_titles_keyword_range':
                keyword = match.group(1).strip()
                y1 = int(match.group(2)); y2 = int(match.group(3))
                mask = (sug['Issue Title'].astype(str).str.contains(keyword, case=False, na=False)) & (sug['Audit Year'].between(y1, y2))
                companies = sug.loc[mask, 'Company'].dropna().unique()
                count = len(companies)
                return f"Number of companies with '{keyword}' in their issue titles from {y1} to {y2}: **{count}**"

    # Fallback guidance
    return (
        "Sorry, I couldn't understand your question.\n\n"
        "Try:\n"
        "- Which companies are located in Germany?\n"
        "- Which country is Adolf Würth GmbH & Co. KG located in?\n"
        "- Who were the auditors for Würth USA Inc. in 2021?\n"
        "- Which companies have the 'finance' in the finding category from 2020 to 2021?\n"
        "- Show all Finding categories of Würth USA Inc. in 2021\n"
        "- Show all findings for Würth USA Inc. with \"inventory\" in the finding title from 2020 to 2021\n"
        "- List the top 5 frequent finding titles in 2021."
    )

# =============================
# Streamlit UI
# =============================
st.title("Audit Chatbot")

st.write(
    "Ask a question about audit **Main**, **Introduction**, **Finding**, and **Suggestion** tables:\n"
    "- Companies, countries, KF/EVP responsibilities, auditors\n"
    "- Audit dates, durations, participants\n"
    "- Findings, risks, recommendations, categories, titles\n"
    "- Issues, issue titles, and suggestions"
)

user_input = st.text_input("Your question:")

if user_input:
    with st.spinner("Thinking..."):
        response = answer_question(user_input, finding, main, intro, suggestion)
    st.markdown(f"**Chatbot:**\n\n{response}")

st.markdown("---")
st.markdown("**Supported question examples:**")

# Main table examples
st.markdown("- Which companies are located in **Germany**?")
st.markdown("- Which country is **Adolf Würth GmbH & Co. KG** located in?")
st.markdown("- Who is the **KF** responsible for **Würth USA Inc.**?")
st.markdown("- Who is the **EVP/SVP** responsible for **Würth Russia**?")
st.markdown("- Who were the **auditors** for **Würth Canada Ltd.** in **2021**?")
st.markdown("- Which companies have **Robert Friedmann** as their **KF**?")
st.markdown("- Which companies were audited by **Marco Bambek** in **2021**?")
st.markdown("- How many companies from **China** were audited in **2021**?")
st.markdown("- List the top **5 auditors** by number of audits from **2020** to **2021**.")
st.markdown("- List the top **5 countries** by number of audited companies from **2020** to **2021**.")

# Finding table examples (keyword -> companies)
st.markdown("- Which companies have **\"inventory\"** in the **finding title** in **2021**?")
st.markdown("- Which companies have **finance** in the **finding category** from **2020** to **2021**?")

# NEW: Company-specific (no keyword)
st.markdown("- Show all **Finding categories** of **Würth USA Inc.** in **2021**.")
st.markdown("- Show all **Findings** for **RECA NORM GmbH** from **2020** to **2021**.")
st.markdown("- Show all **Risks** of **Würth Elektronik Oy** in **2021**.")
st.markdown("- Show all **Recommendations** for **Würth Elektronik Oy** from **2020** to **2021**.")

# Composite (company + keyword + field + time) with quotes/variants
st.markdown("- Show all **findings** for **Würth Elektronik Oy** with **\"inventory\"** in the **finding title** **in 2021**.")
st.markdown("- Show all **recommendations** for **Adolf Würth GmbH & Co. KG** where **IT** is in the **finding category** **from 2020 to 2021**.")

# Top-N frequency
st.markdown("- List the top **5 frequent finding categories** **in 2021**.")
st.markdown("- List the top **3 frequent finding titles** **from 2020 to 2021**.")

# Introduction table examples
st.markdown("- What were the **start and end dates** of the audit for **Würth USA Inc.** in **2021**?")
st.markdown("- Which companies had audits that **started between 01.01.2021 and 30.06.2021**?")
st.markdown("- Which companies had **audit durations** of **less than 10 days** from **2020** to **2021**?")

# Suggestion table examples
st.markdown("- What are the **issue titles** for **Würth USA Inc.** in **2021**?")
st.markdown("- What are the **issues, issue titles, and suggestions** for **Adolf Würth GmbH & Co. KG** from **2020** to **2021**?")
st.markdown("- Which companies have **\"inventory\"** in their **issue titles** from **2020** to **2021**?")
st.markdown("- Which companies have **\"fraud\"** in their **issues** from **2020** to **2021**?")
st.markdown("- How many companies have **\"authorization\"** in their **issue titles** from **2020** to **2021**?")
