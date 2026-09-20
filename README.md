# Sanctions & PEP Screening Workflow (Excel + OpenSanctions-style data)

A no-code sanctions and PEP (Politically Exposed Person) screening
workflow, built entirely in Excel formulas — no Python, no SQL, no
VBA, no add-ins. It screens a synthetic customer book against a
reference list structured the way real sanctions/PEP data feeds are
organized, most notably **[OpenSanctions](https://www.opensanctions.org/)**
and OFAC's SDN list: an entity record with a primary name, known
aliases, entity type, country, date of birth, and the list/program
that designated it.

## ⚠️ About the data — read this first

The `Sanctions_PEP_List` file is a **small, entirely fictional
sample** built only to mirror the *schema* of a real sanctions/PEP
feed. None of the names, entities, or program designations refer to
any real sanctioned individual, real politically exposed person, or
real designated entity. It exists to demonstrate the screening
*logic*, not to serve as reference data.

For actual production or research use, screen against the real,
current data instead:
- **OpenSanctions** consolidated dataset — [opensanctions.org](https://www.opensanctions.org/) (free for personal/non-commercial use; commercial licensing applies for production use — check their current terms)
- **OFAC SDN List** — [sanctionslist.ofac.treas.gov](https://sanctionslist.ofac.treas.gov/) (US government data, public domain, free)

The column structure in this workbook (`primary_name`, `aliases`,
`entity_type`, `country`, `date_of_birth`, `list_program`) was
deliberately chosen to match an OpenSanctions-style export, so
swapping the demo data for a real downloaded extract is mostly a
copy-paste exercise.

## What's in this repo

```
Sanctions_PEP_Screening_Workflow.xlsx   ← the main workbook (open this)
README.md
data/
  sanctions_pep_list.csv                ← reference list, standalone CSV
  customers_to_screen.csv               ← customer book, standalone CSV
  screening_results_output.csv          ← computed results, exported as CSV
screenshots/
  sanctions_list.png
  screening_results.png
```

The CSVs are provided so the underlying data is viewable without
opening Excel (e.g. directly on GitHub), and so the same data could
be dropped into another tool (Google Sheets, a database, a script)
without re-typing anything.

## Workbook structure

| Tab | Purpose |
|---|---|
| `README` | Methodology and how to read the workbook (same content as this file, for anyone who only opens the workbook) |
| `Sanctions_PEP_List` | The reference list being screened against — 25 synthetic entries (mix of sanctions and PEP records) |
| `Customers_to_Screen` | 30 synthetic customers being checked |
| `Screening_Results` | The matching engine and alert queue — every formula, one row per customer |

Open `Screening_Results` and click any cell to see the live formula
that produced it — nothing is hardcoded or pasted as a value.

## Match types this workflow detects

1. **Exact name match** — the customer's stated name matches a listed
   entity's primary name exactly (case- and whitespace-insensitive).
2. **Alias match** — the customer's stated name matches one of the
   known aliases on file for a listed entity.
3. **Partial/token match** — the customer's name shares its first
   name and the remainder of its name with a list entry without
   being an exact match (catches things like a dropped middle name).
4. **Secondary identifier corroboration** — for any name-based hit,
   the workflow separately checks whether the customer's date of
   birth or country *also* matches that same list entry.

That last point matters more than it might look. Two customers in
the sample data — `C013 Dmitri Volkov` and `C014 Grace Ito` — share
an exact name with a listed entity but have a *different* date of
birth and country on file. The workbook still escalates them (a
name match is a name match — it should never be auto-cleared on
weak secondary data), but the `dob_or_country_corroboration_count`
column shows `0` for both, versus `2` for a genuine likely match
like `C015 Farid Al-Nassar`. That distinction — between "this needs
a human to check" and "this is probably a coincidence, but still
needs a human to check" — is the actual job of a screening analyst,
and it's why the workbook surfaces the signal instead of collapsing
everything into a single yes/no flag.

## Key formulas used

All standard Excel — no add-ins, no VBA:

- **`SUMPRODUCT` with `UPPER`/`TRIM`** — case- and whitespace-insensitive
  exact matching across the entire reference list in a single formula.
- **`SEARCH` (wrapped in `ISNUMBER`)** — substring matching for aliases
  and partial/token matches.
- **`FIND` + `LEFT`/`MID`** — splitting a customer's name into a first
  token and a remainder token for partial matching, with no add-in.
- **Array-entered `TEXTJOIN` + `IF`** — pulling back the actual matched
  name(s) and program(s) into readable columns, not just a count.
- **Conditional formatting** — red for confirmed name/alias hits,
  amber for partial matches needing manual review.

## Results

Running the workflow against the 30 synthetic customers correctly
identifies all 11 planted match scenarios and returns `CLEAR` for
all 19 clean customers — zero false positives, zero missed hits.

![Screening Results](screenshots/screening_results.png)

## A documented limitation

This workflow uses substring and first/remainder-token matching, not
true fuzzy (edit-distance) matching like the Jaro-Winkler or
Levenshtein algorithms used in commercial screening platforms
(World-Check, LexisNexis Bridger, Actimize). Pure Excel formulas
can't express edit-distance scoring without VBA or Power Query.

This is a deliberate, disclosed trade-off, not an oversight: the
token method will catch some transliteration and formatting variants
(a dropped hyphen, a reordered middle name) but not distant
misspellings or phonetic variants. A production deployment would
pair this kind of first-pass logic with a true fuzzy-matching engine
— the value of this project is in showing the underlying screening
*logic* (what to check, how to disambiguate, how to score) rather
than claiming formula-only Excel can fully replace one.

## Limitations (stated honestly)

- Synthetic data only, built to mirror a real feed's schema — see
  the data note above.
- 25 reference entries and 30 customers, sized for a readable
  demonstration — a real screening run processes the full consolidated
  list (tens of thousands of entries) against an entire customer book.
- No handling of non-Latin script names or transliteration variants
  beyond what's explicitly listed as an alias.
- Thresholds and match logic are illustrative; a live program would
  calibrate and validate this kind of rule against its actual false
  positive/negative rates.

## About

Built by Biswajit Das, CAMS-certified compliance analyst, as a
portfolio piece demonstrating sanctions/PEP screening logic without
a coding dependency.
