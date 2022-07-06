import pandas as pd

read_file = pd.read_csv ('village.csv',dtype=str)
read_file.to_excel ('village1.xlsx', index = None, header=True)