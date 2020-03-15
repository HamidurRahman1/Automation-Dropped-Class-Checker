
import pandas as pd
import os
import re

# ex = pd.read_excel("Math drop list fall 2019.xlsx")
# ex.to_csv("Math drop list fall 2019.csv", index=None, header=True)

repeat_grades = {}

print("mat 96".endswith("96"))
print(next(re.finditer(r'\d+$', 'mat 201')).group(0))
print(next(re.finditer(r'\d+$', 'mat 201')).group(0))
