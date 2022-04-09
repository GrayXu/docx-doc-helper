# docx-doc-helper

VBA and Word are not appropriate for huge documents. This repo provides some high-level utils based on [*python-docx*](https://github.com/python-openxml/python-docx) for docx files.

warning: Those codes are built for maintaining GB/T documents specifically. And they are not well-tested, so check before usage.

## detail

file | feature
| - | - |
table_cell_replace.py | table text replacement
replace_by_csv.py | replace text in paragraphs by a given .csv file
sign_for_known.py | sign all items in the table for known patterns
stat.ipynb | gen stat img from \*.xlsx
change_sec_num.ipynb | change section number (1.1 -> 2.1)
detect_replaceable_by_xlsx.ipynb | detect for replace_by_xlsx.ipynb  
replace_by_xlsx.ipynb | replace the text matched only once by a given .xlsx file
sec_title_style.ipynb | auto-set section title's style based on its level
utils.py | python-docx utils
table_para_style.ipynb | set specific paragraphs to same style


## dependency

`pip3 install python-docx csv xlrd=1.2.0`