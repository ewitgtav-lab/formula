import datetime as _dt
import random
from difflib import SequenceMatcher
from urllib.parse import quote_plus

import pandas as pd
import streamlit as st
from st_copy_to_clipboard import st_copy_to_clipboard


def _formulas_data() -> list[dict]:
    # Version values are intentionally simple for filtering in Community Cloud.
    # - "All": works in most Excel versions
    # - "Office 365": modern dynamic array / newer functions
    return [
        {
            "Name": "XLOOKUP",
            "Category": "Lookup",
            "Syntax": "XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])",
            "Description": "Modern lookup that replaces VLOOKUP/HLOOKUP with left/right lookup, exact match by default, and built-in not-found handling.",
            "Example_Formula": '=XLOOKUP(A2,$D$2:$D$20,$E$2:$E$20,"Not found")',
            "Version": "Office 365",
            "Tagline": "Modern lookup, safer defaults",
        },
        {
            "Name": "VLOOKUP",
            "Category": "Lookup",
            "Syntax": "VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])",
            "Description": "Finds a value in the first column of a table and returns a value in the same row from a specified column.",
            "Example_Formula": '=VLOOKUP(A2,$D$2:$F$20,3,FALSE)',
            "Version": "All",
            "Tagline": "Classic vertical lookup",
        },
        {
            "Name": "HLOOKUP",
            "Category": "Lookup",
            "Syntax": "HLOOKUP(lookup_value, table_array, row_index_num, [range_lookup])",
            "Description": "Finds a value in the first row of a table and returns a value in the same column from a specified row.",
            "Example_Formula": '=HLOOKUP(A1,$B$1:$H$5,3,FALSE)',
            "Version": "All",
            "Tagline": "Classic horizontal lookup",
        },
        {
            "Name": "INDEX",
            "Category": "Lookup",
            "Syntax": "INDEX(array, row_num, [column_num])",
            "Description": "Returns the value of a cell in a table based on row and column numbers.",
            "Example_Formula": "=INDEX($E$2:$E$20, MATCH(A2,$D$2:$D$20,0))",
            "Version": "All",
            "Tagline": "Return a value by position",
        },
        {
            "Name": "MATCH",
            "Category": "Lookup",
            "Syntax": "MATCH(lookup_value, lookup_array, [match_type])",
            "Description": "Returns the relative position of a lookup value within a range. Commonly paired with INDEX for flexible lookups.",
            "Example_Formula": "=MATCH(A2,$D$2:$D$20,0)",
            "Version": "All",
            "Tagline": "Find a position quickly",
        },
        {
            "Name": "INDEX/MATCH",
            "Category": "Lookup",
            "Syntax": "INDEX(return_range, MATCH(lookup_value, lookup_range, 0))",
            "Description": "A robust lookup pattern that can look left and is less fragile than VLOOKUP when columns move.",
            "Example_Formula": "=INDEX($E$2:$E$20, MATCH(A2,$D$2:$D$20,0))",
            "Version": "All",
            "Tagline": "Flexible lookup combo",
        },
        {
            "Name": "OFFSET",
            "Category": "Lookup",
            "Syntax": "OFFSET(reference, rows, cols, [height], [width])",
            "Description": "Returns a reference shifted by a given number of rows/columns; often used to build dynamic ranges (but can be volatile).",
            "Example_Formula": "=SUM(OFFSET(B2,0,0,5,1))",
            "Version": "All",
            "Tagline": "Build dynamic ranges (volatile)",
        },
        {
            "Name": "INDIRECT",
            "Category": "Lookup",
            "Syntax": "INDIRECT(ref_text, [a1])",
            "Description": "Converts a text string into a reference. Powerful for dynamic sheet/range references, but volatile.",
            "Example_Formula": '=SUM(INDIRECT("B2:B10"))',
            "Version": "All",
            "Tagline": "Text-to-reference (volatile)",
        },
        {
            "Name": "LET",
            "Category": "Array",
            "Syntax": "LET(name1, value1, [name2, value2], …, calculation)",
            "Description": "Assigns names to intermediate calculations to improve readability and performance—especially in complex formulas.",
            "Example_Formula": "=LET(x, A2*B2, y, C2*D2, x+y)",
            "Version": "Office 365",
            "Tagline": "Name intermediates, simplify logic",
        },
        {
            "Name": "LAMBDA",
            "Category": "Array",
            "Syntax": "LAMBDA([parameter1, parameter2, …], calculation)",
            "Description": "Creates custom reusable functions without VBA. Often paired with LET and helper functions like MAP/REDUCE.",
            "Example_Formula": "=LAMBDA(x, x^2)(A2)",
            "Version": "Office 365",
            "Tagline": "Build your own functions",
        },
        {
            "Name": "FILTER",
            "Category": "Array",
            "Syntax": "FILTER(array, include, [if_empty])",
            "Description": "Filters a range based on criteria and spills matching rows/columns.",
            "Example_Formula": '=FILTER(A2:D100, D2:D100="Active", "No matches")',
            "Version": "Office 365",
            "Tagline": "Query-like filtering in a formula",
        },
        {
            "Name": "UNIQUE",
            "Category": "Array",
            "Syntax": "UNIQUE(array, [by_col], [exactly_once])",
            "Description": "Returns a list of unique values from a range (spills).",
            "Example_Formula": "=UNIQUE(A2:A100)",
            "Version": "Office 365",
            "Tagline": "Distinct list in one step",
        },
        {
            "Name": "SORT",
            "Category": "Array",
            "Syntax": "SORT(array, [sort_index], [sort_order], [by_col])",
            "Description": "Sorts a range dynamically and spills the results.",
            "Example_Formula": "=SORT(A2:D100, 2, 1)",
            "Version": "Office 365",
            "Tagline": "Dynamic sorting",
        },
        {
            "Name": "SORTBY",
            "Category": "Array",
            "Syntax": "SORTBY(array, by_array1, [sort_order1], [by_array2], [sort_order2], …)",
            "Description": "Sorts an array by one or more other arrays (keys).",
            "Example_Formula": "=SORTBY(A2:D100, D2:D100, -1, B2:B100, 1)",
            "Version": "Office 365",
            "Tagline": "Sort by a different column",
        },
        {
            "Name": "SEQUENCE",
            "Category": "Array",
            "Syntax": "SEQUENCE(rows, [columns], [start], [step])",
            "Description": "Generates a sequence of numbers that spills into cells—great for dynamic models and dashboards.",
            "Example_Formula": "=SEQUENCE(12,1,1,1)",
            "Version": "Office 365",
            "Tagline": "Generate numbers instantly",
        },
        {
            "Name": "TEXTSPLIT",
            "Category": "Text",
            "Syntax": "TEXTSPLIT(text, col_delimiter, [row_delimiter], [ignore_empty], [match_mode], [pad_with])",
            "Description": "Splits text into columns/rows using delimiters and spills the result.",
            "Example_Formula": '=TEXTSPLIT(A2, ",")',
            "Version": "Office 365",
            "Tagline": "Split text into columns (spills)",
        },
        {
            "Name": "TEXTJOIN",
            "Category": "Text",
            "Syntax": "TEXTJOIN(delimiter, ignore_empty, text1, [text2], …)",
            "Description": "Joins text from multiple cells with a delimiter, optionally ignoring blanks.",
            "Example_Formula": '=TEXTJOIN(", ", TRUE, A2:A10)',
            "Version": "All",
            "Tagline": "Concatenate with a delimiter",
        },
        {
            "Name": "CONCAT",
            "Category": "Text",
            "Syntax": "CONCAT(text1, [text2], …)",
            "Description": "Concatenates multiple strings or ranges into one string.",
            "Example_Formula": "=CONCAT(A2, \" - \", B2)",
            "Version": "All",
            "Tagline": "Simple concatenation",
        },
        {
            "Name": "LEFT",
            "Category": "Text",
            "Syntax": "LEFT(text, [num_chars])",
            "Description": "Returns the leftmost characters from a text string.",
            "Example_Formula": "=LEFT(A2, 5)",
            "Version": "All",
            "Tagline": "Take characters from the left",
        },
        {
            "Name": "RIGHT",
            "Category": "Text",
            "Syntax": "RIGHT(text, [num_chars])",
            "Description": "Returns the rightmost characters from a text string.",
            "Example_Formula": "=RIGHT(A2, 4)",
            "Version": "All",
            "Tagline": "Take characters from the right",
        },
        {
            "Name": "MID",
            "Category": "Text",
            "Syntax": "MID(text, start_num, num_chars)",
            "Description": "Returns a specific number of characters from a text string starting at a position.",
            "Example_Formula": "=MID(A2, 3, 5)",
            "Version": "All",
            "Tagline": "Extract from the middle",
        },
        {
            "Name": "TRIM",
            "Category": "Text",
            "Syntax": "TRIM(text)",
            "Description": "Removes extra spaces from text, leaving single spaces between words.",
            "Example_Formula": "=TRIM(A2)",
            "Version": "All",
            "Tagline": "Clean up spacing",
        },
        {
            "Name": "CLEAN",
            "Category": "Text",
            "Syntax": "CLEAN(text)",
            "Description": "Removes non-printing characters from text (useful when importing data).",
            "Example_Formula": "=CLEAN(A2)",
            "Version": "All",
            "Tagline": "Remove non-printing characters",
        },
        {
            "Name": "SUBSTITUTE",
            "Category": "Text",
            "Syntax": "SUBSTITUTE(text, old_text, new_text, [instance_num])",
            "Description": "Replaces existing text with new text; optionally only the nth occurrence.",
            "Example_Formula": '=SUBSTITUTE(A2, "-", "/")',
            "Version": "All",
            "Tagline": "Targeted text replacement",
        },
        {
            "Name": "REPLACE",
            "Category": "Text",
            "Syntax": "REPLACE(old_text, start_num, num_chars, new_text)",
            "Description": "Replaces part of a text string at a given position with new text.",
            "Example_Formula": '=REPLACE(A2, 1, 3, "ID-")',
            "Version": "All",
            "Tagline": "Replace by position",
        },
        {
            "Name": "FIND",
            "Category": "Text",
            "Syntax": "FIND(find_text, within_text, [start_num])",
            "Description": "Finds the position of one text string within another (case-sensitive).",
            "Example_Formula": '=FIND("@", A2)',
            "Version": "All",
            "Tagline": "Locate text (case-sensitive)",
        },
        {
            "Name": "SEARCH",
            "Category": "Text",
            "Syntax": "SEARCH(find_text, within_text, [start_num])",
            "Description": "Finds the position of one text string within another (case-insensitive) and supports wildcards.",
            "Example_Formula": '=SEARCH("sku", A2)',
            "Version": "All",
            "Tagline": "Locate text (case-insensitive)",
        },
        {
            "Name": "UPPER",
            "Category": "Text",
            "Syntax": "UPPER(text)",
            "Description": "Converts text to uppercase.",
            "Example_Formula": "=UPPER(A2)",
            "Version": "All",
            "Tagline": "Uppercase conversion",
        },
        {
            "Name": "LOWER",
            "Category": "Text",
            "Syntax": "LOWER(text)",
            "Description": "Converts text to lowercase.",
            "Example_Formula": "=LOWER(A2)",
            "Version": "All",
            "Tagline": "Lowercase conversion",
        },
        {
            "Name": "PROPER",
            "Category": "Text",
            "Syntax": "PROPER(text)",
            "Description": "Capitalizes the first letter in each word.",
            "Example_Formula": "=PROPER(A2)",
            "Version": "All",
            "Tagline": "Title-case text",
        },
        {
            "Name": "IF",
            "Category": "Logical",
            "Syntax": "IF(logical_test, value_if_true, value_if_false)",
            "Description": "Tests a condition and returns one value if TRUE and another if FALSE.",
            "Example_Formula": '=IF(B2>=70, "Pass", "Fail")',
            "Version": "All",
            "Tagline": "Basic conditional logic",
        },
        {
            "Name": "IFS",
            "Category": "Logical",
            "Syntax": "IFS(logical_test1, value_if_true1, [logical_test2, value_if_true2], …)",
            "Description": "Evaluates multiple conditions in order and returns the result for the first TRUE condition.",
            "Example_Formula": '=IFS(B2>=90,"A", B2>=80,"B", B2>=70,"C", TRUE,"D")',
            "Version": "All",
            "Tagline": "Multi-branch conditions",
        },
        {
            "Name": "SWITCH",
            "Category": "Logical",
            "Syntax": "SWITCH(expression, value1, result1, [value2, result2], …, [default])",
            "Description": "Matches an expression against a list of values and returns the corresponding result.",
            "Example_Formula": '=SWITCH(A2,"N","New","P","In Progress","D","Done","Unknown")',
            "Version": "All",
            "Tagline": "Readable value mapping",
        },
        {
            "Name": "AND",
            "Category": "Logical",
            "Syntax": "AND(logical1, [logical2], …)",
            "Description": "Returns TRUE only if all conditions are TRUE.",
            "Example_Formula": "=AND(B2>=0, B2<=100)",
            "Version": "All",
            "Tagline": "All conditions must be true",
        },
        {
            "Name": "OR",
            "Category": "Logical",
            "Syntax": "OR(logical1, [logical2], …)",
            "Description": "Returns TRUE if any condition is TRUE.",
            "Example_Formula": '=OR(C2="Yes", D2="Yes")',
            "Version": "All",
            "Tagline": "Any condition can be true",
        },
        {
            "Name": "NOT",
            "Category": "Logical",
            "Syntax": "NOT(logical)",
            "Description": "Reverses TRUE/FALSE.",
            "Example_Formula": "=NOT(A2>0)",
            "Version": "All",
            "Tagline": "Invert a condition",
        },
        {
            "Name": "IFERROR",
            "Category": "Logical",
            "Syntax": "IFERROR(value, value_if_error)",
            "Description": "Returns a fallback value if a formula evaluates to an error; otherwise returns the original result.",
            "Example_Formula": '=IFERROR(1/0, "Check input")',
            "Version": "All",
            "Tagline": "Graceful error handling",
        },
        {
            "Name": "SUM",
            "Category": "Math",
            "Syntax": "SUM(number1, [number2], …)",
            "Description": "Adds numbers, ranges, or a combination.",
            "Example_Formula": "=SUM(B2:B10)",
            "Version": "All",
            "Tagline": "Basic aggregation",
        },
        {
            "Name": "SUMIF",
            "Category": "Math",
            "Syntax": "SUMIF(range, criteria, [sum_range])",
            "Description": "Sums values based on a single condition.",
            "Example_Formula": '=SUMIF(A2:A100, "North", B2:B100)',
            "Version": "All",
            "Tagline": "Conditional sum (one rule)",
        },
        {
            "Name": "SUMIFS",
            "Category": "Math",
            "Syntax": "SUMIFS(sum_range, criteria_range1, criteria1, [criteria_range2, criteria2], …)",
            "Description": "Sums values based on multiple conditions.",
            "Example_Formula": '=SUMIFS(C2:C100, A2:A100, "North", B2:B100, "Widgets")',
            "Version": "All",
            "Tagline": "Conditional sum (many rules)",
        },
        {
            "Name": "COUNT",
            "Category": "Math",
            "Syntax": "COUNT(value1, [value2], …)",
            "Description": "Counts cells that contain numbers.",
            "Example_Formula": "=COUNT(A2:A100)",
            "Version": "All",
            "Tagline": "Count numeric entries",
        },
        {
            "Name": "COUNTA",
            "Category": "Math",
            "Syntax": "COUNTA(value1, [value2], …)",
            "Description": "Counts non-empty cells.",
            "Example_Formula": "=COUNTA(A2:A100)",
            "Version": "All",
            "Tagline": "Count non-blank cells",
        },
        {
            "Name": "COUNTIF",
            "Category": "Math",
            "Syntax": "COUNTIF(range, criteria)",
            "Description": "Counts cells that meet a condition.",
            "Example_Formula": '=COUNTIF(B2:B100, ">=70")',
            "Version": "All",
            "Tagline": "Conditional counting",
        },
        {
            "Name": "COUNTIFS",
            "Category": "Math",
            "Syntax": "COUNTIFS(criteria_range1, criteria1, [criteria_range2, criteria2], …)",
            "Description": "Counts cells that meet multiple conditions.",
            "Example_Formula": '=COUNTIFS(A2:A100,"North", B2:B100,"Widgets")',
            "Version": "All",
            "Tagline": "Conditional counting (many rules)",
        },
        {
            "Name": "AVERAGE",
            "Category": "Math",
            "Syntax": "AVERAGE(number1, [number2], …)",
            "Description": "Returns the arithmetic mean of numbers.",
            "Example_Formula": "=AVERAGE(B2:B10)",
            "Version": "All",
            "Tagline": "Mean value",
        },
        {
            "Name": "ROUND",
            "Category": "Math",
            "Syntax": "ROUND(number, num_digits)",
            "Description": "Rounds a number to a specified number of digits.",
            "Example_Formula": "=ROUND(C2, 2)",
            "Version": "All",
            "Tagline": "Control decimals",
        },
        {
            "Name": "CEILING",
            "Category": "Math",
            "Syntax": "CEILING(number, significance)",
            "Description": "Rounds a number up to the nearest multiple of significance.",
            "Example_Formula": "=CEILING(A2, 5)",
            "Version": "All",
            "Tagline": "Round up to a multiple",
        },
        {
            "Name": "FLOOR",
            "Category": "Math",
            "Syntax": "FLOOR(number, significance)",
            "Description": "Rounds a number down to the nearest multiple of significance.",
            "Example_Formula": "=FLOOR(A2, 5)",
            "Version": "All",
            "Tagline": "Round down to a multiple",
        },
        {
            "Name": "MOD",
            "Category": "Math",
            "Syntax": "MOD(number, divisor)",
            "Description": "Returns the remainder after division.",
            "Example_Formula": "=MOD(A2, 7)",
            "Version": "All",
            "Tagline": "Remainder / cycling logic",
        },
        {
            "Name": "ABS",
            "Category": "Math",
            "Syntax": "ABS(number)",
            "Description": "Returns the absolute value of a number.",
            "Example_Formula": "=ABS(A2)",
            "Version": "All",
            "Tagline": "Remove sign",
        },
        {
            "Name": "MIN",
            "Category": "Math",
            "Syntax": "MIN(number1, [number2], …)",
            "Description": "Returns the smallest value in a set.",
            "Example_Formula": "=MIN(B2:B10)",
            "Version": "All",
            "Tagline": "Smallest value",
        },
        {
            "Name": "MAX",
            "Category": "Math",
            "Syntax": "MAX(number1, [number2], …)",
            "Description": "Returns the largest value in a set.",
            "Example_Formula": "=MAX(B2:B10)",
            "Version": "All",
            "Tagline": "Largest value",
        },
        {
            "Name": "TODAY",
            "Category": "Math",
            "Syntax": "TODAY()",
            "Description": "Returns the current date (updates when the workbook recalculates).",
            "Example_Formula": "=TODAY()",
            "Version": "All",
            "Tagline": "Current date",
        },
        {
            "Name": "NOW",
            "Category": "Math",
            "Syntax": "NOW()",
            "Description": "Returns the current date and time (updates when the workbook recalculates).",
            "Example_Formula": "=NOW()",
            "Version": "All",
            "Tagline": "Current date & time",
        },
        {
            "Name": "PMT",
            "Category": "Financial",
            "Syntax": "PMT(rate, nper, pv, [fv], [type])",
            "Description": "Returns the periodic payment for a loan or investment with constant payments and interest rate.",
            "Example_Formula": "=PMT(0.08/12, 60, -20000)",
            "Version": "All",
            "Tagline": "Loan payment calculator",
        },
        {
            "Name": "NPV",
            "Category": "Financial",
            "Syntax": "NPV(rate, value1, [value2], …)",
            "Description": "Returns the net present value of an investment based on a discount rate and a series of cash flows.",
            "Example_Formula": "=NPV(0.1, C2:C6) + C1",
            "Version": "All",
            "Tagline": "Discounted cash flow",
        },
        {
            "Name": "IRR",
            "Category": "Financial",
            "Syntax": "IRR(values, [guess])",
            "Description": "Returns the internal rate of return for a series of cash flows.",
            "Example_Formula": "=IRR(C1:C6)",
            "Version": "All",
            "Tagline": "Investment return rate",
        },
    ]


def _normalize(s: str) -> str:
    return (s or "").strip().lower()


def _fuzzy_score(query: str, text: str) -> float:
    q = _normalize(query)
    t = _normalize(text)
    if not q:
        return 1.0
    if q in t:
        return 1.0
    return SequenceMatcher(None, q, t).ratio()


def _matches_query(row: pd.Series, query: str) -> tuple[bool, float]:
    if not query.strip():
        return True, 1.0
    name = str(row.get("Name", ""))
    desc = str(row.get("Description", ""))
    syntax = str(row.get("Syntax", ""))
    tagline = str(row.get("Tagline", ""))
    combined = f"{name} {tagline} {desc} {syntax}"
    score = max(
        _fuzzy_score(query, name),
        _fuzzy_score(query, desc),
        _fuzzy_score(query, combined),
    )
    # Conservative threshold so short queries still work, but noise is reduced.
    threshold = 0.45 if len(query.strip()) >= 4 else 0.35
    return score >= threshold, score


def _init_state() -> None:
    if "favorites" not in st.session_state:
        st.session_state.favorites = set()
    if "open_formula" not in st.session_state:
        st.session_state.open_formula = None


def _apply_deep_link(formula_names: set[str]) -> None:
    qp = st.query_params
    target = qp.get("formula", None)
    if not target:
        return
    target_norm = _normalize(str(target))
    for name in formula_names:
        if _normalize(name) == target_norm:
            st.session_state.open_formula = name
            return


def _set_query_param_formula(name: str | None) -> None:
    if not name:
        st.query_params.clear()
        return
    st.query_params["formula"] = name


def _inject_css() -> None:
    st.markdown(
        """
<style>
  /* Subtle Excel-green border around expanders */
  div[data-testid="stExpander"] {
    border: 1px solid rgba(33, 115, 70, 0.35);
    border-radius: 10px;
    padding: 0.15rem 0.4rem;
    background: rgba(33, 115, 70, 0.03);
  }
  /* Slightly tighten expander header spacing */
  div[data-testid="stExpander"] summary {
    padding-top: 0.35rem !important;
    padding-bottom: 0.35rem !important;
  }
  /* Make code blocks a touch more compact */
  pre code {
    font-size: 0.92rem;
  }
</style>
        """,
        unsafe_allow_html=True,
    )


def main() -> None:
    st.set_page_config(page_title="Excel Formula Pro", layout="wide")
    _inject_css()
    _init_state()

    formulas = _formulas_data()
    df = pd.DataFrame(formulas)
    df["Name"] = df["Name"].astype(str)
    df["Category"] = df["Category"].astype(str)
    df["Version"] = df["Version"].astype(str)
    df["Tagline"] = df["Tagline"].astype(str)

    all_names = set(df["Name"].tolist())
    _apply_deep_link(all_names)

    st.title("Excel Formula Pro")
    st.caption("A master encyclopedia of high-value Excel formulas, optimized for Streamlit Community Cloud.")

    complex_pool = df[df["Name"].isin(["LET", "LAMBDA", "FILTER", "UNIQUE", "SORTBY", "SEQUENCE", "XLOOKUP"])].copy()
    if complex_pool.empty:
        complex_pool = df.copy()
    today = _dt.date.today().isoformat()
    random.seed(today)
    fod = complex_pool.sample(1, random_state=random.randint(0, 10_000)).iloc[0]

    with st.container(border=True):
        st.subheader("Formula of the Day")
        st.markdown(f"**{fod['Name']}** — {fod['Tagline']}")
        st.code(str(fod["Syntax"]), language="text")
        st.markdown(str(fod["Description"]))
        st_copy_to_clipboard(str(fod["Example_Formula"]), "Copy example")

    st.divider()

    query = st.text_input("Search formulas", placeholder="Try: lookup with fallback, dynamic arrays, multi-condition sum…")

    categories = sorted(df["Category"].unique().tolist())
    versions = sorted(df["Version"].unique().tolist())

    with st.sidebar:
        st.header("Filters")
        selected_categories = st.multiselect("Filter by Category", options=categories, default=categories)
        o365_only = st.toggle("Version Compatibility: Office 365 only", value=False)
        show_favs_only = st.checkbox("My Favorites view", value=False)
        st.divider()
        st.write(f"**Favorites:** {len(st.session_state.favorites)}")
        st.write(f"**Library size:** {len(df)} formulas")
        st.caption(f"Versions in library: {', '.join(versions)}")

    filtered = df[df["Category"].isin(selected_categories)].copy()
    if o365_only:
        filtered = filtered[filtered["Version"].str.contains("Office 365", case=False, na=False)].copy()
    if show_favs_only:
        filtered = filtered[filtered["Name"].isin(st.session_state.favorites)].copy()

    if query.strip():
        keep_rows = []
        scores = []
        for _, row in filtered.iterrows():
            ok, score = _matches_query(row, query)
            keep_rows.append(ok)
            scores.append(score)
        filtered = filtered.loc[keep_rows].copy()
        if not filtered.empty:
            filtered["__score"] = [s for (k, s) in zip(keep_rows, scores) if k]
            filtered = filtered.sort_values(["__score", "Name"], ascending=[False, True]).drop(columns=["__score"])
    else:
        filtered = filtered.sort_values("Name")

    if filtered.empty:
        st.info("No formulas match your current filters/search. Try broadening your query or selecting more categories.")
        return

    cols = st.columns([2, 1])
    with cols[0]:
        st.subheader("Formula Library")
        st.caption("Tip: share a deep link like `?formula=XLOOKUP` to open a specific formula.")
    with cols[1]:
        if st.button("Clear deep link", use_container_width=True):
            st.session_state.open_formula = None
            _set_query_param_formula(None)

    for _, row in filtered.iterrows():
        name = str(row["Name"])
        tagline = str(row["Tagline"])
        category = str(row["Category"])
        version = str(row["Version"])
        syntax = str(row["Syntax"])
        desc = str(row["Description"])
        example = str(row["Example_Formula"])

        expanded = st.session_state.open_formula == name
        label = f"{name} — {tagline}"

        with st.expander(label, expanded=expanded):
            top = st.columns([1.2, 1, 1, 1])
            with top[0]:
                is_fav = name in st.session_state.favorites
                fav_now = st.checkbox("★ Favorite", value=is_fav, key=f"fav_{name}")
                if fav_now and not is_fav:
                    st.session_state.favorites.add(name)
                if (not fav_now) and is_fav:
                    st.session_state.favorites.discard(name)
            with top[1]:
                st.markdown(f"**Category:** {category}")
            with top[2]:
                st.markdown(f"**Version:** {version}")
            with top[3]:
                if st.button("Open via link", key=f"link_{name}"):
                    st.session_state.open_formula = name
                    _set_query_param_formula(name)
                    st.rerun()

            st.code(syntax, language="text")
            st.markdown(desc)

            actions = st.columns([1, 2])
            with actions[0]:
                st_copy_to_clipboard(example, "Copy example")
            with actions[1]:
                share_url = f"?formula={quote_plus(name)}"
                st.caption(f"Deep link: `{share_url}`")


if __name__ == "__main__":
    main()

