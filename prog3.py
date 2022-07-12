# importer les packages
import openpyxl
import pandas as pd

# Collecter le contenu d'une cellule à partie de plusieurs fichiers
fichiers = ['D:\DevLabs\Data Science\Projet Data Science\Excel_Python\january.xlsx',
            'D:\DevLabs\Data Science\Projet Data Science\Excel_Python\/february.xlsx',
            'D:\DevLabs\Data Science\Projet Data Science\Excel_Python\march.xlsx']

for fichier in fichiers:
    wb = openpyxl.load_workbook(fichier)
    feuille = wb['Sheet1']
    feuille['E9'] = 'Total : '
    feuille['F9'] = '=SUM(F5:F8)'
    feuille['F9'].style = 'Currency'
    wb.save(fichier)

