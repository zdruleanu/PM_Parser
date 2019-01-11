import pandas as pd
from os import listdir


def find_csv_filenames(path_to_dir, suffix=".xls"):
    filenames = listdir(path_to_dir)
    return [path_to_dir+"/"+filename for filename in filenames if filename.endswith(suffix)]


filenames = find_csv_filenames("./data")
for name in filenames:
    print(name)

excelfile = pd.read_excel(filenames[0])

print(excelfile.head())