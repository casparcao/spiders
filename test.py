import re
tpl = "a{page}"

print(tpl.format(page=str(1)))

a = """
acaba
"""

m = re.findall("a.*?a(.*?)a", a, re.S)
print(m[0])

d = "."

e = d.index('.') if d else ''

