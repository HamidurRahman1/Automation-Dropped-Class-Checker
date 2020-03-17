
import pandas as pd
import os
import re

# ex = pd.read_excel("Math drop list fall 2019.xlsx")
# ex.to_csv("Math drop list fall 2019.csv", index=None, header=True)

# repeat_grades = {}
#
# print("mat 96".endswith("96"))
# print(next(re.finditer(r'\d+$', 'mat 201')).group(0))
# print(next(re.finditer(r'\d+$', 'mat 201')).group(0))
import platform
import os
def file_checker():
    download = os.path.join(os.path.expanduser('~'), 'downloads')
    os_name=platform.system().lower()
    path = None
    if os_name=="darwin" or os_name=="linux":
        path=download+"/test/"
    elif os_name=="windows":
        path = download+"\\test"
    if not os.path.isdir(path):
        raise FileNotFoundError("No directory")
    files=list()
    for file in os.listdir(path):
        if file.endswith(".txt"):
            files.append(file)
    if len(files)!=1:
        raise FileExistsError("expected 1 found" + str(len(files)))
    return path+str(files[0])

try:
    print(file_checker())
except Exception as e:
    print(e.args)