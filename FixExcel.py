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
    print("Usage: FixExcel file.xlsx")
    sys.exit(1)

file_path = os.path.realpath(sys.argv[1])
dir_path = os.path.dirname(file_path)
# print(dirpath)
output_excel = os.path.join(dir_path, "output.xlsx")
output_txt = os.path.join(dir_path, "output.txt")

CURRENCY = ["CNY", "USD", "HKD", "EUR"]
app = xw.App(visible=False, add_book=False)
wb = app.books.open(file_path, read_only=True)
nwb = app.books.add()
nwb.save(output_excel)
nsht = nwb.sheets[0]
f = codecs.open(output_txt, "w", "utf-8")

header = "账务日期, 币种, 科目代码, 科目名称, 借方发生额, 贷方发生额, 借方余额, 贷方余额"
f.write(header + "\n")
nsht.range('A1').value = header.split(",")

def _atexit():
    wb.close()
    nwb.save()
    nwb.close()
    app.quit()

atexit.register(_atexit)

def getRowNum(sht):
    gap_cnt = 0
    old_row_num = 0
    row_num = sht['A1'].current_region.last_cell.row
    while row_num - old_row_num > 2 or gap_cnt < 5:
        if (row_num - old_row_num <= 2):
            gap_cnt += 1
        else:
            gap_cnt = 0
        old_row_num = row_num
        row_num = sht['A' + str(row_num + 2)].current_region.last_cell.row
    return row_num

found = False
for sht in wb.sheets:
    row_num = getRowNum(sht)
    # print(row_num)
    for begin in range(1, row_num):
        prev = sht.range('A' + str(begin)).value
        if prev and re.match('[0-9]@OD@', str(prev)):
            # print(prev)
            found = True
            break
    if (found):
        break

if not found:
    print("cannot find the sheet")
    sys.exit(1)


def procRange():
    for s in prev.split("|"):
        if re.match('S-ORG-ID:', s):
            x, orgId = s.split(":")
        elif re.match('DAT:', s):
            x, finDate = s.split(":")
        elif re.match('CCY-ID:', s):
            x, currCode = s.split(":")
    if re.match("1620", orgId) and currCode in CURRENCY:
        # rng = sht.range('A' + str(i + 1))
        for i in range(begin + 1, end):
            # print(i)
            try:
                content = sht.range('A' + str(i)).expand('right').options(ndim=1).value
            except:
                # ignore: dates or times are negative or too large (out of present range)
                print("row " + repr(i) + " format error, skipped.")
                continue
            if not content:
                continue
            # print(content)

            for idx, val in enumerate(content):
                # print(val)
                if val:
                    break

            if len(content) - idx >= 6 and re.match("[0-9]", str(val)):
                s = finDate + ", " + currCode + ", " + str(int(content[idx])) + ", "
                for j in range(idx + 1, idx + 5):
                    s += str(content[j]) + ", "
                s += str(content[idx + 5])
                f.write(s + "\n")
                scope = nsht.range('A1').expand()
                nsht.range(scope.shape[0] + 1, 1).value = s.split(",")


for end in range(begin + 1, row_num):
    value = sht.range('A' + str(end)).value
    if value and re.match('[0-9]@OD@', str(value)):
        procRange()
        begin = end
        prev = value

procRange()

nsht.range('A1').expand('table').columns.autofit()

print("Generate the excel file: " + output_excel)
print("Generate the txt file: " + output_txt)

