import pandas as pd

df = pd.read_excel('latest.xlsx')

print("Shipped Date column data type:", df['Shipped Date'].dtype)
print("Sample shipped dates:")
print(df['Shipped Date'].head(10))

print("Min shipped date:", df['Shipped Date'].min())
print("Max shipped date:", df['Shipped Date'].max())
