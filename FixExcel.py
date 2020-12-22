import xlwings as xw
import re
import sys
import codecs
import atexit
import os.path


# print(sys.stdout.encoding)
if sys.stdout.encoding != 'utf8':
    sys.stdout = codecs.getwriter('utf8')(sys.stdout.buffer, 'strict')
if sys.stderr.encoding != 'utf8':
    sys.stderr = codecs.getwriter('utf8')(sys.stderr.buffer, 'strict')

if len(sys.argv) < 2:
    print("Usage: FixExcel input.file")
    sys.exit(1)

file_path = os.path.realpath(sys.argv[1])
dir_path = os.path.dirname(file_path)
# print(dirpath)
output_excel = os.path.join(dir_path, "output.xlsx")
output_txt = os.path.join(dir_path, "output.txt")

app = xw.App(visible=False, add_book=False)
nwb = app.books.add()
nwb.save(output_excel)
nsht = nwb.sheets[0]
fin = codecs.open(file_path, "r", "utf-8")
f = codecs.open(output_txt, "w", "utf-8")

header = "账务日期" + '\x1d'
header += "币种" + '\x1d'
header += "机构号" + '\x1d'
header += "科目代码" + '\x1d'
header += "科目名称" + '\x1d'
header += "借方发生额" + '\x1d'
header += "贷方发生额" + '\x1d'
header += "借方余额" + '\x1d'
header += "贷方余额"


# f.write(header + "\n")
nsht.range('A1').value = header.split("\x1d")

def _atexit():
    nwb.save()
    nwb.close()
    app.quit()

atexit.register(_atexit)

lines = fin.readlines()
row_num = len(lines)

found = False
for begin in range(0, row_num):
    prev = lines[begin]
    if prev and re.match('[0-9]@OD@', str(prev)):
        found = True
        break

if not found:
    print("Unable to find the correct format file.")
    sys.exit(1)


def procRange():
    for s in prev.split("|"):
        if re.match('S-ORG-ID:', s):
            x, orgId = s.split(":")
        elif re.match('DAT:', s):
            x, finDate = s.split(":")
        elif re.match('CCY-ID:', s):
            x, currCode = s.split(":")

    if (not orgId) or (not finDate) or (not currCode):
        print("Unable to fetch the value, skipped")
        return

    for i in range(begin + 1, end):
        content = lines[i].split()
        if len(content) < 6 or not re.match("[0-9]", str(content[0])):
            continue

        s = finDate + '\x1d' + currCode + '\x1d' + orgId  + '\x1d'
        s += str(int(content[0])) + '\x1d' # 科目代码

        # 科目名称可能含有空格
        for j in range(1, len(content)):
            if (not content[j].isdigit()) and \
               (content[j].replace(",", "").replace(".", "", 1).isdigit()):
                break
        else:
            print("Unable to fetch the account name, skipped")
            return

        temp = ""
        for k in range(1, j):
            temp += content[k] + " "
        s += temp.strip() + '\x1d'

        idx = j
        for j in range(idx, idx + 3):
            s += str(content[j]) + '\x1d'
        s += str(content[idx + 3])

        f.write(s + "\n")
        scope = nsht.range('A1').expand()
        nsht.range(scope.shape[0] + 1, 1).value = s.split("\x1d")


for end in range(begin + 1, row_num):
    value = lines[end]
    if value and re.match('[0-9]@OD@', str(value)):
        procRange()
        begin = end
        prev = value

procRange()

nsht.range('A1').expand('table').columns.autofit()

print("Generate the excel file: " + output_excel)
print("Generate the txt file: " + output_txt)

