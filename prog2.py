# importer les packages
import openpyxl
import pandas as pd

# Collecter le contenu d'une cellule à partie de plusieurs fichiers
fichiers = ['D:\DevLabs\Data Science\Projet Data Science\Excel_Python\january.xlsx',
            'D:\DevLabs\Data Science\Projet Data Science\Excel_Python\/february.xlsx',
            'D:\DevLabs\Data Science\Projet Data Science\Excel_Python\march.xlsx']
valeurs = []
valeurs_dict = {}

for fichier in fichiers:
    wb = openpyxl.load_workbook(fichier)
    feuille = wb['Sheet1']
    valeur = feuille['f5'].value
    valeurs.append(valeur)
valeurs_dict['valeurs'] = valeurs
pd.DataFrame(valeurs_dict).to_excel('valeurs_toronto.xlsx')

# Afficher la liste des valeurs collectées
print('La liste obtenue : ',valeurs)
print('Le dictionnaire obtenu : ',valeurs_dict)
