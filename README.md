xlsx2vtb
========
generate csv/vtb files from Excel(xlsx) sheets.

usage: `python xlsx2vtb <xlsx file>`

Specification and Restriction:
* csv/vtb file name is defined from sheet name
* it is assumed that data in each sheets start from A1 cell
* header row is always needed

Future Work:
* allow custom file name
  * auto-deploy to VTB path
* evaluate expression
