import xlrd
import re
from docx import Document
import locale
from itertools import takewhile
import time

locale.setlocale(locale.LC_TIME, '')  # Heure en Français

dateActuelle = time.strftime('%d %B %Y')

numberLines = 0;


def docx_replace_regex(doc_obj, regex, replace):
    for p in doc_obj.paragraphs:
        if regex.search(p.text):
            inline = p.runs
            # Loop added to work with runs (strings with same style)
            for i in range(len(inline)):
                if regex.search(inline[i].text):
                    text = regex.sub(replace, inline[i].text)
                    inline[i].text = text

    for table in doc_obj.tables:
        for row in table.rows:
            for cell in row.cells:
                docx_replace_regex(cell, regex, replace)


document = Document(r"C:\Users\ppintus\Documents\csd\courrier.docx")
wb = xlrd.open_workbook(r"C:\Users\ppintus\Documents\csd\\testexcel.xls");

sh = wb.sheet_by_name('Feuil1')

nombreLigne = (len(sh.col_values(0)))
nombreColonne = sh.row_len(0);

print(nombreLigne)
print(nombreColonne)

cell_type = sh.cell_type(5, 3)
if (cell_type == xlrd.XL_CELL_EMPTY):
    print("Cette cellule est blanche")
else:
    print(cell_type)

regex = re.compile(r"champ1");
replace = str(int(round((sh.row_values(1)[0]))))
docx_replace_regex(document, regex, replace);

regex = re.compile(r"champ2")
replace = dateActuelle;
docx_replace_regex(document, regex, replace)

regex = re.compile(r"champ3");
replace = str(sh.row_values(1)[3]);
docx_replace_regex(document, regex, replace);

regex = re.compile(r"champ4")
replace = (sh.row_values(1)[4]);
replace = str(int(round(replace)))  # Passe de float à int

docx_replace_regex(document, regex, replace)

regex = re.compile(r"champ5")
replace = str(sh.row_values(1)[5]);
docx_replace_regex(document, regex, replace)

regex = re.compile(r"champ6")
replace = str(int(round((sh.row_values(1)[6]))));
docx_replace_regex(document, regex, replace)

for i in range(1, nombreLigne):

# Save
document.save(r"C:\Users\ppintus\Documents\csd\courrier2.docx")

# document.save(r"C:\Users\ppintus\Documents\csd\courrier.docx")


# regex1 = re.compile(r"Caverne")
# replace1 = r"Monsieur le président"
# docx_replace_regex(document, regex1 , replace1)


# docx_replace_regex((document,regex,replace1));
# document.save(r"C:\Users\ppintus\Documents\csd\courrier1.docx")


