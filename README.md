I’ll target a single sheet from englishfullfilefromchrome.xlsx and export it.

I found these sheet names (note leading spaces):
" Fr-01", " Fr-02", " Fr-03.1", " Fr-03.2", " Data for Scope I II & III", " 3-Quantitative Inspection (24)", " 3-Quantitative Inspection (25)", " Fr-04.1", " Fr-04.2", " Fr-05", " EF TGO AR4", " EF TGO AR5", " Update log"
I generated the first sheet as an example:

Output: english_one_sheet/Fr-01.json
Run it yourself for a specific sheet

Single JSON containing only one sheet:
python3 extract_excel.py englishfullfilefromchrome.xlsx sheet_fr01.json --sheet " Fr-01"
Or write that single sheet into a directory (one file):
python3 extract_excel.py englishfullfilefromchrome.xlsx english_one_sheet --split-sheets --sheet " Fr-01"
Notes

Sheet names include a leading space; copy the name exactly or rename the sheet in Excel first if you prefer no leading space.
To export all sheets into separate files:
python3 extract_excel.py englishfullfilefromchrome.xlsx english_sheets_out --split-sheets
Tell me which sheet you want (exact name), and I’ll export that one too.


 python3 extract_excel.py thai_empty_sheet.xlsx thai_empty_sheets_out --split-sheets