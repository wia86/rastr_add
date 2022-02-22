import os, re , random
key = 'na=11' # 'na=11(node)'
match = re.search(re.compile(r"\((.+?)\)"), key)
if match:
    rastr_table = match[1]
    key = key.split('(', 1)[0]


name_file = 'cor(ny=15525, pn=10 qn=pn*0.4)(89)[years = 2026, season=лет, max_min=min, add_name=0°C]'
pattern = re.compile(r"\[(.+?)\]")
pattern = re.compile(r"^.+?")
match = re.search(pattern, name_file)
if match:
    print(name_file)
    print(match)

name_file = '*[2026][зим][мин][0°C]'
pattern_name = re.compile("\[.*]\[.*]\[.*]\[.*]")
match = re.search(pattern_name, name_file)
if match:
    print(name_file)
    print(match[0])
    print([match])

