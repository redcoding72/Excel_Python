# Importer les packages
import pandas as pd

# Lister les fichiers à combiner
fichiers = ['D:\DevLabs\Data Science\Projet Data Science\Excel_Python\january.xlsx',
            'D:\DevLabs\Data Science\Projet Data Science\Excel_Python\/february.xlsx',
            'D:\DevLabs\Data Science\Projet Data Science\Excel_Python\march.xlsx']
# Créer le fichier combiné de type dataframe
fichier_combine = pd.DataFrame ()

# Combiner les trois fichiers
for fichier in fichiers:
    df = pd.read_excel (fichier, skiprows=3)
    print (df)
    fichier_combine = fichier_combine.append (df, ignore_index=True)

# Céer le fichier excel correspondant
fichier_combine.to_excel ('fich_comb.xlsx')
