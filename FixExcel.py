import xlwings as xw
import re
import sys
import codecs
import atexit
import os.path
import gzip


# print(sys.stdout.encoding)
if sys.stdout.encoding != 'utf8':
    sys.stdout = codecs.getwriter('utf8')(sys.stdout.buffer, 'strict')
if sys.stderr.encoding != 'utf8':
    sys.stderr = codecs.getwriter('utf8')(sys.stderr.buffer, 'strict')


def procDir(fpath):
    print("processing directory: " + fpath)
    return


def procExcel(fpath):
    print("processing excel: " + fpath)
    return


def procFile(fpath):
    # process the .txt file or .gz file
    print("processing file: " + fpath)
    file_path = os.path.realpath(fpath)
    dir_path = os.path.dirname(file_path)
    file_name = os.path.basename(fpath)
    # file_name = re.sub('\..*$', '', file_name)
    if file_name.lower().endswith(('.gz', '.gzip', '.txt')):
        arr = file_name.split('.')
        if len(arr) > 1:
            file_name = ""
            for i in range(0, len(arr) - 2):
                file_name += arr[i] + "."
            file_name += arr[-2]

    output_name = file_name + "_fix.txt"
    output_txt = os.path.join(dir_path, output_name)

    if fpath.lower().endswith(('.gz', '.gzip')):
        fin = gzip.open(fpath, 'rt', encoding='utf8')
    else:
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
        return


    def procRange():
        for s in prev.split("|"):
            if re.match('S-ORG-ID:', s):
                x, org_id = s.split(":")
            elif re.match('DAT:', s):
                x, fin_date = s.split(":")
            elif re.match('CCY-ID:', s):
                x, curr_code = s.split(":")

        if (not org_id) or (not fin_date) or (not curr_code):
            print("Unable to fetch the value, skipped")
            return

        for i in range(begin + 1, end):
            content = lines[i].split()
            if len(content) < 6 or not re.match("[0-9]", str(content[0])):
                continue

            # "账务日期" + "币种" + "机构号"
            s = fin_date + '\x1d' + curr_code + '\x1d' + org_id  + '\x1d'
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

            # "借方发生额" + "贷方发生额" + "借方余额" + "贷方余额"
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


################################
if len(sys.argv) < 2:
    print("Usage: FixExcel <file|directory>")
    sys.exit(1)

fpath = sys.argv[1]
if os.path.isfile(fpath):
    if fpath.lower().endswith(('.xlsx', '.xls')):
        procExcel(fpath)
    else:
        procFile(fpath)
elif os.path.isdir(fpath):
    procDir(fpath)
else:
    print("Unknown file type: " + fpath)
    sys.exit(1)

