import pandas as pd

around = pd.read_excel("around.xlsx")
cafe = pd.read_excel("cafe.xlsx")

combined = pd.concat([around, cafe])

combined['순번'] = range(1, len(combined) + 1)
combined.to_excel("combined.xlsx", index=False)

