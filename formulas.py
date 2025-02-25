from openpyxl import Workbook

#Criando uma nova planilha direto pelo Python
wb = Workbook()

sheet = wb.active

sheet["A1"].value = 100
sheet["A2"].value = 200

formula = "=SUM(A1:A2)"
sheet["A3"].value = formula

from openpyxl.formula.translate import Translator
sheet["B1"].value = 300
sheet["B2"].value = 250

#Copia a formula da A3 para a B3 igual quando arrastamos no Excel
sheet["B3"].value = Translator(formula, origin="A3").translate_formula("B3")

wb.save("formula.xlsx")

#Conhecendo as funções do excel disponível na biblioteca
from openpyxl.utils import FORMULAE
print(FORMULAE)

