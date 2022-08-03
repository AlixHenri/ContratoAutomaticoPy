#pip install python.docx
#pip install pandas
#pip install numpy
#pip install openpyxl

from docx import Document
from datetime import datetime
import pandas as pd

tabela = pd.read_excel("Dados.xlsx")

for linha in tabela.index:
    documento = Document("Contrato.docx")

    nome = tabela.loc[linha, "Nome"]
    item1 = tabela.loc[linha, "Item1"]
    item2 = tabela.loc[linha, "Item2"]
    item3 = tabela.loc[linha, "Item3"]

    #dicionario
    referencias = {
        "XXXX": nome,
        "YYYY": item1,
        "ZZZZ": item2,
        "WWWW": item3,
        "DD": str(datetime.now().day),
        "MM": str(datetime.now().month),
        "AAAA": str(datetime.now().year)
    }

    for paragrafo in documento.paragraphs:
            for codigo in referencias:
                valor = referencias[codigo]
                paragrafo.text = paragrafo.text.replace(codigo, valor)

    documento.save(f"Contrato - {nome}.docx")