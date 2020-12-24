import platform
import re
import os

root_path = os.getcwd()
isMac = re.match("[Ww]indows", platform.system()) == None
config_path = root_path + "/config.txt"
exe_path = ""
para = []
with open(config_path, 'r', encoding="UTF-8") as f:
        content = f.readlines()
        for line in content:
            if(line.startswith('#') or len(line) <= 0):
                continue
            match = re.match(r"(\w+)\s*=\s*(.+)",line)
            key = match.group(1)
            key = "--"+key
            value = match.group(2)
            if("path" in key):
                value = root_path + value
            if("exe_path" in key):
                exe_path = value
                continue
            para.append(key)
            para.append(value)

cmd = "{} {} ".format(exe_path, " ".join(para))
if isMac:
    cmd = "mono " + cmd
res = os.popen(cmd)
output_str = res.read()
print(output_str)

try:
    input("Press Enter to exit...")
except:
    pass