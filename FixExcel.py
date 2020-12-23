import re
import sys
import codecs
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
output_txt = os.path.join(dir_path, "output.txt")

fin = codecs.open(file_path, "r", "utf-8")
f = codecs.open(output_txt, "w", "utf-8")

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

        # 账务日期 + 币种 + 机构号
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

        # 借方发生额 + 贷方发生额 + 借方余额 + 贷方余额
        if len(content) - j < 4:
            print("Format error, skipped")
            return

        idx = j
        for j in range(idx, idx + 3):
            s += str(content[j]) + '\x1d'
        s += str(content[idx + 3])

        f.write(s + "\n")

for end in range(begin + 1, row_num):
    value = lines[end]
    if value and re.match('[0-9]@OD@', str(value)):
        procRange()
        begin = end
        prev = value

procRange()

print("Generate the txt file: " + output_txt)

