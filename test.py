import xlrd
from docx import Document



from docx import Document
from docx.shared import Inches
document = Document("D:\Documents\CSD\  \testWord")


wb = xlrd.open_workbook("D:\Documents\CSD\\test.xls");
print(wb.sheet_names());
sh= wb.sheet_by_name('Feuille1')
for rownum in range(sh.nrows):
    print(sh.row_values(rownum));

colonne1 = (sh.row_values(1));


print(colonne1);
print(colonne1[0]);

