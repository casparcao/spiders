import re
import pandas as pd
from pandas import DataFrame
tpl = "a{page}"

print(tpl.format(page=str(1)))

a = """
acaba
"""

m = re.findall("a.*?a(.*?)a", a, re.S)
print(m[0])

d = "."

e = d.index('.') if d else ''

df = DataFrame([
    {"title": "title0",
        "link": "link0",
        "no": 12313123,
        "district": "历下",
        "area": "泉城路",
        "price": 133.01,
        "amount": 10},
    {"title": "title1",
        "link": "link1",
        "no": 12313125,
        "district": "历下",
        "area": "泉城路",
        "price": 133.11,
        "amount": 11}
])
print(df)
with pd.ExcelWriter("./test.xlsx") as writer:
    df.to_excel(writer, sheet_name="Sheet1")
