# NNM Group data cleaning

Two scripts I wrote at N.N.M. Group (logistics) for files the office types by hand.

## clean_cmr.py: every CMR in one table

The CMR forms are Excel sheets, one form per sheet, typed in an old Georgian font (AcadNusx).
They're fine to print, but you can't search or filter them.

The script reads every form and puts each CMR on one row of a single table:

- Converts AcadNusx text to real Georgian ("TariRi" → "თარიღი"), but leaves words like MAXTHERM LTD and container numbers alone
- Reads dates in every format people used: real Excel dates, "20,03,26" typed as text, 6.08 typed as a number
- Replaces impossible dates (like year 2054) with the loading date
- Splits the container cell into size, container number and seal number
- Fixes weights typed as text ("26996, 95" → 26996.95)
- Removes drivers' personal ID numbers and keeps only the name
- Splits truck and trailer plates
- Flags possible duplicate CMRs (same date, container, truck and weight) instead of deleting them

Output: `cmr_clean.xlsx` with a **CMRs** sheet and a **Data Quality** sheet that lists every fix and how many rows it touched.

## borjomi_xelfasi.py: one client's payments for a year

The accounting ledger has 235 sheets going back to 2014. One entry can take several rows: the client name is on one row,
"xelfasi" (salary/fee) and the amount are on the next one or two, and the date sits in another column.

The script:

- Goes through every sheet and links each amount back to its client and date
- Keeps only BORJOMI payments from 2025
- Fixes date typos: a date with the right month but the wrong year inside a 2025 sheet (like 16,02,24 in February 2025) counts as 2025
- Totals the payments by month, and shows months with no payments as 0

Output: `borjomi_xelfasi_2025.xlsx` with Summary, By Month, All Entries and Notes (every typo fix is written down).

## Run

    pip install pandas numpy xlrd openpyxl

Put the scripts in the same folder as the Excel files and run them. They find the files by name, so you don't need to type any paths.

The data files aren't in this repo because they contain company, client and driver information.
