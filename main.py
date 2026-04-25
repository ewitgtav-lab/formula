import datetime as _dt
from urllib.parse import quote_plus

import pandas as pd
import streamlit as st
from fuzzywuzzy import fuzz
from st_copy_to_clipboard import st_copy_to_clipboard


EXCEL_GREEN = "#217346"


def _normalize(s: str) -> str:
    return " ".join((s or "").strip().lower().split())


def _init_state() -> None:
    if "favorites" not in st.session_state:
        st.session_state.favorites = set()
    if "open_formula" not in st.session_state:
        st.session_state.open_formula = None


def _apply_deep_link(formula_names: list[str]) -> None:
    target = st.query_params.get("formula", None)
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
        f"""
<style>
  :root {{
    --excel-green: {EXCEL_GREEN};
  }}
  /* Make it feel like a clean mobile app */
  .block-container {{
    padding-top: 1.3rem;
    padding-bottom: 2.5rem;
    max-width: 900px;
  }}
  /* Green-accented expanders */
  div[data-testid="stExpander"] {{
    border: 1px solid rgba(33, 115, 70, 0.35);
    border-radius: 12px;
    padding: 0.15rem 0.55rem;
    background: rgba(33, 115, 70, 0.03);
  }}
  div[data-testid="stExpander"] summary {{
    padding-top: 0.45rem !important;
    padding-bottom: 0.45rem !important;
  }}
  /* Small tag pill */
  .pill {{
    display: inline-block;
    padding: 0.18rem 0.55rem;
    border-radius: 999px;
    border: 1px solid rgba(33, 115, 70, 0.35);
    color: rgba(33, 115, 70, 0.95);
    background: rgba(33, 115, 70, 0.06);
    font-size: 0.78rem;
    line-height: 1.2;
  }}
</style>
        """,
        unsafe_allow_html=True,
    )


def _formulas_data() -> list[dict]:
    # Beginner-first structure. Keep examples copy-pasteable.
    # Categories are intentionally few: Finding Stuff, Fixing Text, Quick Math, Logic/Decisions.
    return [
        {
            "Name": "XLOOKUP",
            "Human_Name": "The Finder (New & Fancy)",
            "The_Vibe": "Like asking a helpful librarian who can also walk you to the shelf.",
            "When_To_Use": "You have a list of products and want the price for “Apple” without your sheet exploding.",
            "Syntax_Simplified": "=XLOOKUP(What you want, Where to search, What to return, What to show if missing)",
            "Copy_Paste_Example": '=XLOOKUP(A2,$D$2:$D$20,$E$2:$E$20,"Not found")',
            "Category": "Finding Stuff",
            "Keywords": "lookup find match return fetch price id code search",
        },
        {
            "Name": "VLOOKUP",
            "Human_Name": "The Finder (Old School)",
            "The_Vibe": "Like asking a librarian, but only if the book is on the right side of the shelf.",
            "When_To_Use": "Your data is a simple table and you need to fetch one column from it.",
            "Syntax_Simplified": "=VLOOKUP(What you want, The table, Which column to return, Exact match?)",
            "Copy_Paste_Example": '=VLOOKUP(A2,$D$2:$F$20,3,FALSE)',
            "Category": "Finding Stuff",
            "Keywords": "lookup find match return fetch table",
        },
        {
            "Name": "INDEX",
            "Human_Name": "The Pointer",
            "The_Vibe": "You point at a row and column and Excel hands you the thing.",
            "When_To_Use": "You want a flexible lookup that doesn’t break when columns move.",
            "Syntax_Simplified": "=INDEX(The list/table, Which row, Which column)",
            "Copy_Paste_Example": "=INDEX($E$2:$E$20, 3)",
            "Category": "Finding Stuff",
            "Keywords": "position pointer row column return",
        },
        {
            "Name": "MATCH",
            "Human_Name": "The Position Finder",
            "The_Vibe": "Excel says: “It’s the 7th one down.”",
            "When_To_Use": "You want the position of a value so you can combine it with INDEX for a powerful lookup.",
            "Syntax_Simplified": "=MATCH(What you want, Where to look, Exact match or approximate)",
            "Copy_Paste_Example": "=MATCH(A2,$D$2:$D$20,0)",
            "Category": "Finding Stuff",
            "Keywords": "position index match lookup find exact",
        },
        {
            "Name": "INDEX/MATCH",
            "Human_Name": "The Power Lookup Duo",
            "The_Vibe": "Batman and Robin, but for finding stuff in spreadsheets.",
            "When_To_Use": "You need a robust lookup that can look left or right and survives column changes.",
            "Syntax_Simplified": "=INDEX(What to return, MATCH(What to find, Where to find it, Exact))",
            "Copy_Paste_Example": "=INDEX($E$2:$E$20, MATCH(A2,$D$2:$D$20,0))",
            "Category": "Finding Stuff",
            "Keywords": "lookup index match find fetch return",
        },
        {
            "Name": "FILTER",
            "Human_Name": "The Instant Filter",
            "The_Vibe": "Like a search result page that updates itself.",
            "When_To_Use": "You want to show only rows where Status = “Active” without manually filtering.",
            "Syntax_Simplified": "=FILTER(What to return, Which rows to keep, What to show if nothing matches)",
            "Copy_Paste_Example": '=FILTER(A2:D100, D2:D100="Active", "No matches")',
            "Category": "Finding Stuff",
            "Keywords": "filter rows criteria show only",
        },
        {
            "Name": "UNIQUE",
            "Human_Name": "The Duplicate Remover",
            "The_Vibe": "Like telling Excel: “Give me the guest list, not the receipts.”",
            "When_To_Use": "You need a clean list of distinct values (like all departments) from a messy column.",
            "Syntax_Simplified": "=UNIQUE(Where the list lives)",
            "Copy_Paste_Example": "=UNIQUE(A2:A100)",
            "Category": "Finding Stuff",
            "Keywords": "distinct unique remove duplicates list",
        },
        {
            "Name": "SORT",
            "Human_Name": "The Sort Button (But as a Formula)",
            "The_Vibe": "Auto-sort that doesn’t forget itself.",
            "When_To_Use": "You want a sorted view of data that updates automatically when data changes.",
            "Syntax_Simplified": "=SORT(What to sort, Which column to sort by, Up or down)",
            "Copy_Paste_Example": "=SORT(A2:D100, 2, 1)",
            "Category": "Finding Stuff",
            "Keywords": "sort order ascending descending",
        },
        {
            "Name": "SORTBY",
            "Human_Name": "Sort by Something Else",
            "The_Vibe": "Sorting one list by the vibes of another list.",
            "When_To_Use": "You want to sort a table by a specific column (or multiple columns).",
            "Syntax_Simplified": "=SORTBY(What to return, Sort key #1, Up/down, Sort key #2, Up/down)",
            "Copy_Paste_Example": "=SORTBY(A2:D100, D2:D100, -1, B2:B100, 1)",
            "Category": "Finding Stuff",
            "Keywords": "sort sortby sort by multiple",
        },
        {
            "Name": "SUMIF",
            "Human_Name": "Add Only the Matching Stuff",
            "The_Vibe": "Like totaling only the receipts that say “coffee.”",
            "When_To_Use": "You want the total sales for “North” only.",
            "Syntax_Simplified": "=SUMIF(Where to check, What to match, What to add up)",
            "Copy_Paste_Example": '=SUMIF(A2:A100, "North", B2:B100)',
            "Category": "Quick Math",
            "Keywords": "sum total add condition one criteria",
        },
        {
            "Name": "SUMIFS",
            "Human_Name": "Add Only the Matching Stuff (Hard Mode)",
            "The_Vibe": "Like a bouncer letting in only people who match multiple rules.",
            "When_To_Use": "You want total sales for North + Widgets only.",
            "Syntax_Simplified": "=SUMIFS(What to add, Check range #1, Rule #1, Check range #2, Rule #2)",
            "Copy_Paste_Example": '=SUMIFS(C2:C100, A2:A100, "North", B2:B100, "Widgets")',
            "Category": "Quick Math",
            "Keywords": "sum total add condition multiple criteria",
        },
        {
            "Name": "COUNTIF",
            "Human_Name": "Count the Matching Stuff",
            "The_Vibe": "Like counting how many tasks are marked “Done.”",
            "When_To_Use": "You need to count how many scores are 70 or higher.",
            "Syntax_Simplified": "=COUNTIF(Where to check, Rule to match)",
            "Copy_Paste_Example": '=COUNTIF(B2:B100, ">=70")',
            "Category": "Quick Math",
            "Keywords": "count how many condition",
        },
        {
            "Name": "COUNTIFS",
            "Human_Name": "Count with Multiple Rules",
            "The_Vibe": "Like counting only the red marbles that are also shiny.",
            "When_To_Use": "Count orders in North that are Widgets.",
            "Syntax_Simplified": "=COUNTIFS(Check range #1, Rule #1, Check range #2, Rule #2)",
            "Copy_Paste_Example": '=COUNTIFS(A2:A100,"North", B2:B100,"Widgets")',
            "Category": "Quick Math",
            "Keywords": "count multiple criteria",
        },
        {
            "Name": "AVERAGE",
            "Human_Name": "The Middle-ish Number",
            "The_Vibe": "Excel’s best guess at “typical.”",
            "When_To_Use": "You want the average score, average sale, average time… you get it.",
            "Syntax_Simplified": "=AVERAGE(What numbers to average)",
            "Copy_Paste_Example": "=AVERAGE(B2:B10)",
            "Category": "Quick Math",
            "Keywords": "average mean typical",
        },
        {
            "Name": "ROUND",
            "Human_Name": "The Makeup Brush",
            "The_Vibe": "Smooths out messy decimals so your report looks professional.",
            "When_To_Use": "You want prices rounded to 2 decimals (money vibes).",
            "Syntax_Simplified": "=ROUND(Number, How many decimals)",
            "Copy_Paste_Example": "=ROUND(C2, 2)",
            "Category": "Quick Math",
            "Keywords": "round decimals money",
        },
        {
            "Name": "MIN",
            "Human_Name": "The Smallest One",
            "The_Vibe": "Finds the baby number in the group.",
            "When_To_Use": "You need the lowest score, lowest price, earliest date, etc.",
            "Syntax_Simplified": "=MIN(Where the numbers live)",
            "Copy_Paste_Example": "=MIN(B2:B10)",
            "Category": "Quick Math",
            "Keywords": "min smallest lowest",
        },
        {
            "Name": "MAX",
            "Human_Name": "The Biggest One",
            "The_Vibe": "Finds the boss number in the group.",
            "When_To_Use": "You need the highest score, highest price, latest date, etc.",
            "Syntax_Simplified": "=MAX(Where the numbers live)",
            "Copy_Paste_Example": "=MAX(B2:B10)",
            "Category": "Quick Math",
            "Keywords": "max biggest highest",
        },
        {
            "Name": "IF",
            "Human_Name": "The Fork in the Road",
            "The_Vibe": "If this, then that. If not… the other thing.",
            "When_To_Use": "You want to label scores as Pass/Fail, or apply simple rules.",
            "Syntax_Simplified": '=IF(Question to test, What if YES, What if NO)',
            "Copy_Paste_Example": '=IF(B2>=70, "Pass", "Fail")',
            "Category": "Logic/Decisions",
            "Keywords": "if then else decision",
        },
        {
            "Name": "IFS",
            "Human_Name": "If… Else If… Else If… (Without Tears)",
            "The_Vibe": "A checklist of rules, first match wins.",
            "When_To_Use": "You want letter grades from a score, or buckets like High/Medium/Low.",
            "Syntax_Simplified": "=IFS(Rule 1, Result 1, Rule 2, Result 2, …, TRUE, Default)",
            "Copy_Paste_Example": '=IFS(B2>=90,"A", B2>=80,"B", B2>=70,"C", TRUE,"D")',
            "Category": "Logic/Decisions",
            "Keywords": "ifs multiple if grade buckets",
        },
        {
            "Name": "AND",
            "Human_Name": "All Conditions Must Be True",
            "The_Vibe": "Like a strict bouncer: everyone must have a ticket AND an ID.",
            "When_To_Use": "You need to check multiple conditions at once (e.g., score between 0 and 100).",
            "Syntax_Simplified": "=AND(Condition 1, Condition 2, …)",
            "Copy_Paste_Example": "=AND(B2>=0, B2<=100)",
            "Category": "Logic/Decisions",
            "Keywords": "and both conditions",
        },
        {
            "Name": "OR",
            "Human_Name": "Any Condition Can Be True",
            "The_Vibe": "Like a flexible bouncer: ticket OR VIP pass is fine.",
            "When_To_Use": "You want TRUE if any condition is met.",
            "Syntax_Simplified": "=OR(Condition 1, Condition 2, …)",
            "Copy_Paste_Example": '=OR(C2="Yes", D2="Yes")',
            "Category": "Logic/Decisions",
            "Keywords": "or either condition",
        },
        {
            "Name": "IFERROR",
            "Human_Name": "The Error Muffler",
            "The_Vibe": "Stops #N/A from screaming in your face.",
            "When_To_Use": "A lookup might fail sometimes and you want a friendly message instead.",
            "Syntax_Simplified": "=IFERROR(Your formula, What to show if it breaks)",
            "Copy_Paste_Example": '=IFERROR(XLOOKUP(A2,$D$2:$D$20,$E$2:$E$20), "Not found")',
            "Category": "Logic/Decisions",
            "Keywords": "error handle fallback",
        },
        {
            "Name": "TEXTJOIN",
            "Human_Name": "The Name Gluer",
            "The_Vibe": "Like making a sandwich, but with words and a delimiter.",
            "When_To_Use": "You want to join first + last names, or combine tags into one cell.",
            "Syntax_Simplified": "=TEXTJOIN(Separator, Ignore blanks?, Things to join)",
            "Copy_Paste_Example": '=TEXTJOIN(" ", TRUE, A2, B2)',
            "Category": "Fixing Text",
            "Keywords": "join combine concatenate names",
        },
        {
            "Name": "CONCAT",
            "Human_Name": "The Simple Gluer",
            "The_Vibe": "Tape for text.",
            "When_To_Use": "You just need to stick two or more text pieces together.",
            "Syntax_Simplified": "=CONCAT(Text 1, Text 2, …)",
            "Copy_Paste_Example": '=CONCAT(A2," ",B2)',
            "Category": "Fixing Text",
            "Keywords": "join concatenate combine",
        },
        {
            "Name": "LEFT",
            "Human_Name": "Take from the Left",
            "The_Vibe": "Cuts the first N characters like trimming a string haircut.",
            "When_To_Use": "You want the first 3 letters of a code, or the area code from a phone number.",
            "Syntax_Simplified": "=LEFT(Text, How many characters)",
            "Copy_Paste_Example": "=LEFT(A2, 3)",
            "Category": "Fixing Text",
            "Keywords": "extract left first characters",
        },
        {
            "Name": "RIGHT",
            "Human_Name": "Take from the Right",
            "The_Vibe": "Snips the end of text cleanly.",
            "When_To_Use": "You want the last 4 digits of an ID or phone number.",
            "Syntax_Simplified": "=RIGHT(Text, How many characters)",
            "Copy_Paste_Example": "=RIGHT(A2, 4)",
            "Category": "Fixing Text",
            "Keywords": "extract right last characters",
        },
        {
            "Name": "MID",
            "Human_Name": "Take from the Middle",
            "The_Vibe": "A tiny scalpel for text surgery.",
            "When_To_Use": "You need the middle part of a code like ABC-123-XYZ → 123.",
            "Syntax_Simplified": "=MID(Text, Start position, How many characters)",
            "Copy_Paste_Example": "=MID(A2, 5, 3)",
            "Category": "Fixing Text",
            "Keywords": "extract substring middle",
        },
        {
            "Name": "TRIM",
            "Human_Name": "The Space Cleaner",
            "The_Vibe": "Vacuum for extra spaces.",
            "When_To_Use": "Data from emails/exports has weird spacing and your matches keep failing.",
            "Syntax_Simplified": "=TRIM(Text)",
            "Copy_Paste_Example": "=TRIM(A2)",
            "Category": "Fixing Text",
            "Keywords": "remove spaces clean",
        },
        {
            "Name": "SUBSTITUTE",
            "Human_Name": "The Find & Replace Button (But Smarter)",
            "The_Vibe": "Swaps text like a polite shapeshifter.",
            "When_To_Use": "You want to replace dashes with slashes in dates or codes.",
            "Syntax_Simplified": "=SUBSTITUTE(Text, What to replace, Replace with what)",
            "Copy_Paste_Example": '=SUBSTITUTE(A2, "-", "/")',
            "Category": "Fixing Text",
            "Keywords": "replace substitute swap",
        },
        {
            "Name": "FIND",
            "Human_Name": "Where Is That Character?",
            "The_Vibe": "A detective that cares about capitalization.",
            "When_To_Use": "You want to find where “@” appears in an email address (case-sensitive).",
            "Syntax_Simplified": "=FIND(What to find, Where to look)",
            "Copy_Paste_Example": '=FIND("@", A2)',
            "Category": "Fixing Text",
            "Keywords": "find position character",
        },
        {
            "Name": "SEARCH",
            "Human_Name": "Find Text (Chill Mode)",
            "The_Vibe": "A detective that doesn’t care about capitalization.",
            "When_To_Use": "You want to find a word inside a cell even if the casing changes.",
            "Syntax_Simplified": "=SEARCH(What to find, Where to look)",
            "Copy_Paste_Example": '=SEARCH("sku", A2)',
            "Category": "Fixing Text",
            "Keywords": "search find position case insensitive",
        },
        {
            "Name": "UPPER",
            "Human_Name": "Shouty Text",
            "The_Vibe": "Turns everything into CAPS like it’s excited.",
            "When_To_Use": "You need consistent formatting (e.g., IDs or country codes).",
            "Syntax_Simplified": "=UPPER(Text)",
            "Copy_Paste_Example": "=UPPER(A2)",
            "Category": "Fixing Text",
            "Keywords": "uppercase caps",
        },
        {
            "Name": "LOWER",
            "Human_Name": "Quiet Text",
            "The_Vibe": "Makes text lowercase so it stops yelling.",
            "When_To_Use": "You want consistent lowercase emails/usernames.",
            "Syntax_Simplified": "=LOWER(Text)",
            "Copy_Paste_Example": "=LOWER(A2)",
            "Category": "Fixing Text",
            "Keywords": "lowercase",
        },
        {
            "Name": "PROPER",
            "Human_Name": "Make It Look Nice",
            "The_Vibe": "Fixes names like “jANE DOE” into “Jane Doe.”",
            "When_To_Use": "You’re cleaning contact lists or names for a mail merge.",
            "Syntax_Simplified": "=PROPER(Text)",
            "Copy_Paste_Example": "=PROPER(A2)",
            "Category": "Fixing Text",
            "Keywords": "title case proper names",
        },
        {
            "Name": "SUM",
            "Human_Name": "Add It Up",
            "The_Vibe": "The calculator button everyone trusts.",
            "When_To_Use": "You want a total (expenses, sales, hours…).",
            "Syntax_Simplified": "=SUM(What to add up)",
            "Copy_Paste_Example": "=SUM(B2:B10)",
            "Category": "Quick Math",
            "Keywords": "sum total add",
        },
        {
            "Name": "COUNT",
            "Human_Name": "Count Numbers Only",
            "The_Vibe": "Counts only the cells that are actually numbers (no funny business).",
            "When_To_Use": "You want to know how many numeric entries you have.",
            "Syntax_Simplified": "=COUNT(Where to count numbers)",
            "Copy_Paste_Example": "=COUNT(A2:A100)",
            "Category": "Quick Math",
            "Keywords": "count numbers how many",
        },
        {
            "Name": "COUNTA",
            "Human_Name": "Count Anything (That’s Not Blank)",
            "The_Vibe": "Counts filled cells like attendance.",
            "When_To_Use": "You want to count how many cells have something in them (text or numbers).",
            "Syntax_Simplified": "=COUNTA(Where to count filled cells)",
            "Copy_Paste_Example": "=COUNTA(A2:A100)",
            "Category": "Quick Math",
            "Keywords": "count non blank filled",
        },
        {
            "Name": "ABS",
            "Human_Name": "Make It Positive",
            "The_Vibe": "Turns negative numbers into “we’re fine” numbers.",
            "When_To_Use": "You want distance from zero regardless of sign (e.g., variance size).",
            "Syntax_Simplified": "=ABS(Number)",
            "Copy_Paste_Example": "=ABS(A2)",
            "Category": "Quick Math",
            "Keywords": "absolute positive remove negative sign",
        },
        {
            "Name": "MOD",
            "Human_Name": "The Remainder Machine",
            "The_Vibe": "Tells you what’s left after dividing—useful for patterns.",
            "When_To_Use": "You’re alternating colors/labels every N rows or checking even/odd.",
            "Syntax_Simplified": "=MOD(Number, Divide by what)",
            "Copy_Paste_Example": "=MOD(A2, 2)",
            "Category": "Quick Math",
            "Keywords": "remainder even odd cycle",
        },
        {
            "Name": "TODAY",
            "Human_Name": "What Day Is It?",
            "The_Vibe": "Excel checks the calendar so you don’t have to.",
            "When_To_Use": "You want today’s date for dashboards or aging calculations.",
            "Syntax_Simplified": "=TODAY()",
            "Copy_Paste_Example": "=TODAY()",
            "Category": "Quick Math",
            "Keywords": "date today current",
        },
        {
            "Name": "NOW",
            "Human_Name": "What Time Is It?",
            "The_Vibe": "Like a clock… inside a spreadsheet… because of course.",
            "When_To_Use": "You want a timestamp for “last updated.”",
            "Syntax_Simplified": "=NOW()",
            "Copy_Paste_Example": "=NOW()",
            "Category": "Quick Math",
            "Keywords": "date time now current timestamp",
        },
        {
            "Name": "SWITCH",
            "Human_Name": "The Label Maker",
            "The_Vibe": "Turns codes into friendly words without a giant nested IF mess.",
            "When_To_Use": "You have status codes (N/P/D) and want to show labels (New/In Progress/Done).",
            "Syntax_Simplified": "=SWITCH(Value to check, Case 1, Result 1, Case 2, Result 2, Default)",
            "Copy_Paste_Example": '=SWITCH(A2,"N","New","P","In Progress","D","Done","Unknown")',
            "Category": "Logic/Decisions",
            "Keywords": "switch map labels",
        },
        {
            "Name": "NOT",
            "Human_Name": "Flip the Answer",
            "The_Vibe": "If TRUE becomes FALSE, if FALSE becomes TRUE. Petty but useful.",
            "When_To_Use": "You want the opposite of a condition (like “not blank”).",
            "Syntax_Simplified": "=NOT(Condition)",
            "Copy_Paste_Example": '=NOT(A2="")',
            "Category": "Logic/Decisions",
            "Keywords": "not opposite invert",
        },
        # Pad with additional essential formulas using the same beginner framing (40+ total).
        {
            "Name": "REPLACE",
            "Human_Name": "Replace by Position",
            "The_Vibe": "Cuts out a slice of text and puts something else there.",
            "When_To_Use": "You need to replace the first 3 characters with “ID-”.",
            "Syntax_Simplified": "=REPLACE(Text, Start position, How many characters, New text)",
            "Copy_Paste_Example": '=REPLACE(A2, 1, 3, "ID-")',
            "Category": "Fixing Text",
            "Keywords": "replace by position",
        },
        {
            "Name": "CLEAN",
            "Human_Name": "Remove Weird Invisible Characters",
            "The_Vibe": "Exorcises spooky copy/paste ghosts.",
            "When_To_Use": "You pasted data from a PDF/website and it has invisible junk.",
            "Syntax_Simplified": "=CLEAN(Text)",
            "Copy_Paste_Example": "=CLEAN(A2)",
            "Category": "Fixing Text",
            "Keywords": "clean remove nonprinting",
        },
        {
            "Name": "TEXTSPLIT",
            "Human_Name": "Split Text into Pieces",
            "The_Vibe": "Like cutting a pizza into slices (columns).",
            "When_To_Use": "You have “Last, First” and want to split by comma.",
            "Syntax_Simplified": "=TEXTSPLIT(Text, What separates pieces)",
            "Copy_Paste_Example": '=TEXTSPLIT(A2, ",")',
            "Category": "Fixing Text",
            "Keywords": "split text delimiter comma",
        },
        {
            "Name": "SEQUENCE",
            "Human_Name": "Make a List of Numbers",
            "The_Vibe": "Excel prints 1…2…3… like it’s counting for you.",
            "When_To_Use": "You need an auto-generated index or month numbers.",
            "Syntax_Simplified": "=SEQUENCE(How many rows, How many columns, Start, Step)",
            "Copy_Paste_Example": "=SEQUENCE(12,1,1,1)",
            "Category": "Quick Math",
            "Keywords": "sequence generate numbers list",
        },
        {
            "Name": "LET",
            "Human_Name": "Name Your Steps",
            "The_Vibe": "Like giving nicknames to your math so your formula stops being a horror novel.",
            "When_To_Use": "Your formula is getting long and you want it readable and faster.",
            "Syntax_Simplified": "=LET(Name1, Value1, Name2, Value2, Final calculation)",
            "Copy_Paste_Example": "=LET(x, A2*B2, y, C2*D2, x+y)",
            "Category": "Logic/Decisions",
            "Keywords": "let variables simplify",
        },
        {
            "Name": "LAMBDA",
            "Human_Name": "Make Your Own Mini-Function",
            "The_Vibe": "Like teaching Excel a new trick so you don’t repeat yourself.",
            "When_To_Use": "You keep writing the same complex logic and want a reusable custom function.",
            "Syntax_Simplified": "=LAMBDA(Input, What to do with it)(Your input)",
            "Copy_Paste_Example": "=LAMBDA(x, x^2)(A2)",
            "Category": "Logic/Decisions",
            "Keywords": "lambda custom function reuse",
        },
        {
            "Name": "OFFSET",
            "Human_Name": "Move Over (Dynamic Range)",
            "The_Vibe": "A slippery little function that moves around a spreadsheet.",
            "When_To_Use": "You need a range that shifts/grows automatically (advanced).",
            "Syntax_Simplified": "=OFFSET(Start cell, Move rows, Move columns, Height, Width)",
            "Copy_Paste_Example": "=SUM(OFFSET(B2,0,0,5,1))",
            "Category": "Finding Stuff",
            "Keywords": "offset dynamic range shift",
        },
        {
            "Name": "INDIRECT",
            "Human_Name": "Text → Real Reference",
            "The_Vibe": "Turns a string into a real cell reference. Black magic.",
            "When_To_Use": "You want to build references from text (like sheet names) (advanced).",
            "Syntax_Simplified": '=INDIRECT("Text that looks like a cell/range")',
            "Copy_Paste_Example": '=SUM(INDIRECT("B2:B10"))',
            "Category": "Finding Stuff",
            "Keywords": "indirect reference text dynamic",
        },
        {
            "Name": "PMT",
            "Human_Name": "Monthly Payment Calculator",
            "The_Vibe": "Excel does the scary loan math so you don’t have to.",
            "When_To_Use": "You want to estimate monthly payments for a loan.",
            "Syntax_Simplified": "=PMT(Interest rate per period, Number of payments, Loan amount)",
            "Copy_Paste_Example": "=PMT(0.08/12, 60, -20000)",
            "Category": "Quick Math",
            "Keywords": "loan payment pmt finance",
        },
        {
            "Name": "NPV",
            "Human_Name": "Is This Investment Worth It?",
            "The_Vibe": "Future money, but adjusted for reality.",
            "When_To_Use": "You want the present value of future cash flows at a discount rate.",
            "Syntax_Simplified": "=NPV(Discount rate, Future cash flows) + (Cash flow today)",
            "Copy_Paste_Example": "=NPV(0.1, C2:C6) + C1",
            "Category": "Quick Math",
            "Keywords": "npv present value investment",
        },
        {
            "Name": "IRR",
            "Human_Name": "What’s the Return Rate?",
            "The_Vibe": "The “is this worth it?” percentage.",
            "When_To_Use": "You want the implied rate of return from a set of cash flows.",
            "Syntax_Simplified": "=IRR(Cash flows)",
            "Copy_Paste_Example": "=IRR(C1:C6)",
            "Category": "Quick Math",
            "Keywords": "irr return rate investment",
        },
        {
            "Name": "CEILING",
            "Human_Name": "Round Up to a Multiple",
            "The_Vibe": "Always rounds up like it’s optimistic.",
            "When_To_Use": "You want time blocks in 15-minute chunks, or prices rounded up to 5.",
            "Syntax_Simplified": "=CEILING(Number, Round up to nearest multiple of…)",
            "Copy_Paste_Example": "=CEILING(A2, 5)",
            "Category": "Quick Math",
            "Keywords": "round up ceiling multiple",
        },
        {
            "Name": "FLOOR",
            "Human_Name": "Round Down to a Multiple",
            "The_Vibe": "Always rounds down like it’s budgeting.",
            "When_To_Use": "You want to round down to the nearest multiple (like 5, 10, 0.25, etc).",
            "Syntax_Simplified": "=FLOOR(Number, Round down to nearest multiple of…)",
            "Copy_Paste_Example": "=FLOOR(A2, 5)",
            "Category": "Quick Math",
            "Keywords": "round down floor multiple",
        },
        {
            "Name": "HLOOKUP",
            "Human_Name": "The Horizontal Finder",
            "The_Vibe": "Like VLOOKUP, but sideways.",
            "When_To_Use": "Your table is arranged across the top row and you need a value from a row below.",
            "Syntax_Simplified": "=HLOOKUP(What you want, The table, Which row to return, Exact match?)",
            "Copy_Paste_Example": '=HLOOKUP(A1,$B$1:$H$5,3,FALSE)',
            "Category": "Finding Stuff",
            "Keywords": "lookup horizontal find table",
        },
    ]


def _build_search_blob(row: pd.Series) -> str:
    parts = [
        str(row.get("Name", "")),
        str(row.get("Human_Name", "")),
        str(row.get("The_Vibe", "")),
        str(row.get("When_To_Use", "")),
        str(row.get("Syntax_Simplified", "")),
        str(row.get("Category", "")),
        str(row.get("Keywords", "")),
    ]
    return _normalize(" ".join(parts))


def _score_query(query: str, blob: str) -> int:
    q = _normalize(query)
    if not q:
        return 100
    if q in blob:
        return 100
    # token_set_ratio is good for "goal" searches like "join names" / "remove spaces"
    return int(fuzz.token_set_ratio(q, blob))


def _glossary_sidebar() -> None:
    st.sidebar.subheader("Excel Basics 101")
    st.sidebar.caption("Hover the bold words for a quick explanation.")
    st.sidebar.markdown(
        """
<div style="line-height:1.65">
  <span title="A single box in Excel, like A1 or C7. It holds one value."><b>Cell</b></span>
  — one box (ex: <span class="pill">A1</span>)<br/>
  <span title="A group of cells. Written like A1:A10 (a rectangle of boxes)."><b>Range</b></span>
  — a group of cells (ex: <span class="pill">A1:A10</span>)<br/>
  <span title="An input you give a formula. Example: in SUM(A1:A10), the range A1:A10 is an argument."><b>Argument</b></span>
  — the “inputs” inside parentheses
</div>
        """,
        unsafe_allow_html=True,
    )


def main() -> None:
    st.set_page_config(page_title="Excel for Humans", layout="centered")
    _inject_css()
    _init_state()

    data = _formulas_data()
    df = pd.DataFrame(data)

    required = [
        "Name",
        "Human_Name",
        "The_Vibe",
        "When_To_Use",
        "Syntax_Simplified",
        "Copy_Paste_Example",
        "Category",
    ]
    for col in required:
        if col not in df.columns:
            df[col] = ""
        df[col] = df[col].astype(str)

    formula_names = df["Name"].tolist()
    _apply_deep_link(formula_names)

    st.title("Excel for Humans")
    st.caption("A formula encyclopedia for people who hate Excel (no judgment).")

    query = st.text_input(
        "The “I’m Lost” search bar",
        placeholder='Try goals like: "join names", "remove extra spaces", "find a price", "count done tasks"...',
    )

    with st.sidebar:
        st.header("Browse")
        categories = sorted(df["Category"].unique().tolist())
        selected_categories = st.multiselect("Category", options=categories, default=categories)
        show_favs_only = st.checkbox("Show only my favorites", value=False)
        st.divider()
        _glossary_sidebar()
        st.divider()
        st.subheader("Share")

        reddit_title = quote_plus("Excel for Humans: formula help for people who hate Excel")
        reddit_text = quote_plus(
            "I found a beginner-friendly Excel formula encyclopedia. "
            "Search by goals like 'join names' or 'remove spaces'."
        )
        st.link_button("Share on Reddit", f"https://www.reddit.com/submit?title={reddit_title}&text={reddit_text}")

        st.divider()
        st.subheader("Support")
        st.link_button("Buy me a coffee", "https://paypal.me/GewishCatedrilla")

    view = df[df["Category"].isin(selected_categories)].copy()
    if show_favs_only:
        view = view[view["Name"].isin(st.session_state.favorites)].copy()

    view["__blob"] = view.apply(_build_search_blob, axis=1)
    view["__score"] = view["__blob"].apply(lambda b: _score_query(query, b))

    # Filter mildly so beginners still see helpful suggestions.
    if _normalize(query):
        view = view[view["__score"] >= 55].copy()

    view = view.sort_values(["__score", "Human_Name"], ascending=[False, True]).drop(columns=["__blob"])

    st.write(f"**Showing:** {len(view)} formulas")
    if view.empty:
        st.info("No matches. Try simpler words like “join”, “find”, “remove spaces”, or “count”.")
        return

    # Friendly daily highlight (stable per day)
    seed = int(_dt.date.today().strftime("%Y%m%d"))
    pick = view.sample(1, random_state=seed).iloc[0]
    with st.container(border=True):
        st.subheader("Today’s friendly pick")
        st.markdown(f"**{pick['Human_Name']}**  <span class='pill'>{pick['Name']}</span>", unsafe_allow_html=True)
        st.info(pick["The_Vibe"])
        st.success(pick["When_To_Use"])
        st.code(pick["Syntax_Simplified"], language="text")
        st_copy_to_clipboard(
            pick["Copy_Paste_Example"],
            "Copy Formula",
            key=f"copy_pick_{_dt.date.today().isoformat()}_{pick['Name']}",
        )

    st.divider()

    for _, row in view.iterrows():
        name = row["Name"]
        human = row["Human_Name"]
        expanded = st.session_state.open_formula == name

        with st.expander(f"{human}", expanded=expanded):
            st.markdown(f"**{human}**  <span class='pill'>{name}</span>", unsafe_allow_html=True)

            is_fav = name in st.session_state.favorites
            fav_now = st.checkbox("★ Star this one", value=is_fav, key=f"fav_{name}")
            if fav_now and not is_fav:
                st.session_state.favorites.add(name)
                st.balloons()
            if (not fav_now) and is_fav:
                st.session_state.favorites.discard(name)

            st.info(row["The_Vibe"])
            st.success(row["When_To_Use"])
            st.code(row["Syntax_Simplified"], language="text")

            st_copy_to_clipboard(row["Copy_Paste_Example"], "Copy Formula", key=f"copy_{name}")

            share_qp = f"?formula={quote_plus(name)}"
            st.caption(f"Deep link: `{share_qp}`")
            if st.button("Open with link", key=f"open_{name}"):
                st.session_state.open_formula = name
                _set_query_param_formula(name)
                st.rerun()


if __name__ == "__main__":
    main()

