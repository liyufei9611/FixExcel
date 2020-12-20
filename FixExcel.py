import xlwings as xw
import re
import sys
import codecs


# print(sys.stdout.encoding)
if sys.stdout.encoding != 'utf8':
    sys.stdout = codecs.getwriter('utf8')(sys.stdout.buffer, 'strict')
if sys.stderr.encoding != 'utf8':
    sys.stderr = codecs.getwriter('utf8')(sys.stderr.buffer, 'strict')

CURRENCY = ["CNY", "USD", "HKD", "EUR"]
ROWNUM = 3000
app = xw.App(visible=False, add_book=False)
wb = app.books.open(r'C:\\GongCun\\Temp\\sample.xlsx')
nwb = app.books.add()
nwb.save(f'C:\\GongCun\\Temp\\output.xlsx')
nsht = nwb.sheets['Sheet1']
f = codecs.open("C:\\GongCun\\Temp\\output.txt", "w", "utf-8")

header = "账务日期, 币种, 科目代码, 科目名称, 借方发生额, 贷方发生额, 借方余额, 贷方余额"
f.write(header + "\n")

nsht.range('A1').value = header.split(",")
# nwb.save()
# nwb.close()
# app.quit()
# sys.exit(2)

found = False
for sht in wb.sheets:
    for begin in range(1, ROWNUM):
        prev = sht.range('A' + str(begin)).value
        if prev and re.match('[0-9]@OD@', prev):
            # print(prev)
            found = True
            break
    if (found):
        break

if not found:
    print("cannot find the sheet")
    sys.exit(1)


def procRows():
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
            try:
                content = sht.range('A' + str(i)).expand('right').options(ndim=1).value
            except:
                # ignore: dates or times are negative or too large (out of present range)
                continue
            # content = sht.range('A' + str(i)).expand('right').options(ndim=1).api.Value2
            if not content:
                continue
            # print(content)

            for idx, val in enumerate(content):
                # print(val)
                if val:
                    break

            if len(content) >= 6 and re.match("[0-9]", str(val)):
                s = finDate + ", " + currCode + ", " + str(int(content[idx])) + ", "
                for j in range(idx + 1, idx + 5):
                    s += str(content[j]) + ", "
                s += str(content[idx + 5])
                f.write(s + "\n")
                scope = nsht.range('A1').expand()
                nsht.range(scope.shape[0] + 1, 1).value = s.split(",")


for end in range(begin + 1, ROWNUM):
    value = sht.range('A' + str(end)).value
    if value and re.match('[0-9]@OD@', value):
        procRows()
        begin = end
        prev = value

procRows()
wb.close()
nwb.save()
nwb.close()
app.quit()
