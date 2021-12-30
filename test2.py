import os, re , random
import numpy as np
kkey = "df=34"
key_pop = kkey[:kkey.find('=')]
we = 0
if we:
    print('yes')

name_file = '*[2026][зим][][0°C]'
pattern_name = re.compile("(\[.*\])(\[.*\])(\[.*\])(\[.*\])")
pattern_name = re.compile("\[(.*)\]\[(.*)\]\[(.*)\]\[(.*)\]")
match = re.search(pattern_name, name_file)


years_list =  id_str.replace(" ", "").split(',')
years_list_new = nm.array([],int)
for it in years_list:
    if "-" in it:
        i_years = it.split('-')
        years_list_new = nm.hstack([years_list_new, nm.array(nm.arange(int(i_years[0]), int(i_years[1]) + 1), int)])
    else:
        years_list_new = nm.hstack([years_list_new ,int(it)])
print(nm.sort(years_list_new))

for us in nm.sort(years_list_new):
    print(us)

# return years_list_new




# get all
Uslovie_file = {"years": "2021-2027",
                 "season": "",
                 "max_min": "",
                 "add_name": ""
                 }
t= Uslovie_file.values()
tе="".join(t)

name_list = ["-", "-", "-"]
pattern_name = re.compile("(-?\d+((,|\.)\d*)?)\s?°C")
match = re.search(pattern_name, "2020 +14,45 °Cзимний (sedvw,sev,) ")
# print(match)
name_list = [match[1], match[2], match[3]]
name_list[2] = "-"
i = name_list[0]+name_list[1]+name_list[2]
е=1
# # https://docs.python.org/3/library/configparser.html
# import configparser
# config = configparser.ConfigParser()
# config['DEFAULT'] = {'ServerAliveInterval': '45',
#                      'Compression': 'yes',
#                      'CompressionLevel': '9'}
# config['bitbucket.org'] = {}
# config['bitbucket.org']['User'] = 'hg'
# config['topsecret.server.com'] = {}
# topsecret = config['topsecret.server.com']
# topsecret['Port'] = '50022'     # mutates the parser
# topsecret['ForwardX11'] = 'no'  # same here
# config['DEFAULT']['ForwardX11'] = 'yes'
# with open('example.ini', 'w') as configfile:
#   config.write(configfile)

