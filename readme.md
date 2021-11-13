# SC25-doc-fucker

Have u ever been driven mad by complicated designs of \*.docx? Do u think VBA is not designed for human? 

Hey, star this repo, which provides lot high-level utils based on [*python-docx*](https://github.com/python-openxml/python-docx) to handle[fʌk] docx files. And it's freaking faster than VBA!

file | feature
| - | - |
table_cell_replace.py | table text replacement
replace_by_csv.py | replace text in paragraphs by a given .csv file
sign_for_known.py | sign all items in the table for known patterns
stat.ipynb | gen stat img from \*.xlsx
change_sec_num.ipynb | change section number (1.1 -> 2.1)
detect_replaceable_by_xlsx.ipynb | detect for replace_by_xlsx.ipynb  
replace_by_xlsx.ipynb | replace the text matched only once by a given .xlsx file

check details in header comments of \*.py files.

warning: those codes are not well-tested, so check check and check.

## dependency

`pip install python-docx csv`