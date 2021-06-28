import xlwings as xw
import re
import sys
import codecs
import atexit
import os.path
import gzip
import datetime as dt
import errno


if sys.stdout.encoding != 'utf8':
    sys.stdout = codecs.getwriter('utf8')(sys.stdout.buffer, 'strict')
if sys.stderr.encoding != 'utf8':
    sys.stderr = codecs.getwriter('utf8')(sys.stderr.buffer, 'strict')


if len(sys.argv) < 2:
    print("Usage: FixExcel < file | directory > [ output-directory ]")
    sys.exit(1)

if len(sys.argv) < 3:
    output_dir = os.path.dirname(sys.argv[0])
else:
    output_dir = sys.argv[2]
    try:
        os.mkdir(output_dir)
    except OSError as e:
        if e.errno != errno.EEXIST:
            raise

def procExcel(fpath):
    # process the .xlsx file or .xls file
    print("processing excel: " + fpath)
    file_path = os.path.realpath(fpath)
    file_name = os.path.basename(fpath)
    # file_name = re.sub('\..*$', '', file_name)
    if file_name.lower().endswith(('.xlsx', '.xls')):
        arr = file_name.split('.')
        if len(arr) > 1:
            file_name = ""
            for i in range(0, len(arr) - 2):
                file_name += arr[i] + "."
            file_name += arr[-2]

    output_name = file_name + "_CONVERT.TXT"
    output_txt = os.path.join(output_dir, output_name)
    f = codecs.open(output_txt, "w", "utf-8")

    app = xw.App(visible=False, add_book=False)
    wb = app.books.open(file_path, read_only=True)
    sht = wb.sheets[0]

    def _atexit():
        wb.close()
        app.quit()

    atexit.register(_atexit)

    content = sht.range('A1').expand('table').options(dates=dt.date).value
    for row_value in content:
        s = ""
        for val in row_value:
            s += str(val) + '\x1d'
        f.write(s.rstrip('\x1d') + "\n")

    print("Generate the txt file: " + output_txt)


def procFile(fpath):
    # process the .txt file or .gz file
    print("processing file: " + fpath)
    file_path = os.path.realpath(fpath)
    file_name = os.path.basename(fpath)
    # file_name = re.sub('\..*$', '', file_name)
    if file_name.lower().endswith(('.gz', '.gzip', '.txt')):
        arr = file_name.split('.')
        if len(arr) > 1:
            file_name = ""
            for i in range(0, len(arr) - 2):
                file_name += arr[i] + "."
            file_name += arr[-2]

    output_name = file_name + "_convert.txt"
    output_txt = os.path.join(output_dir, output_name)

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
            elif re.match('LANG:', s):
                x, lang = s.split(":")

        if (not org_id) or (not fin_date) or (not curr_code):
            print("Unable to fetch the value, skipped")
            return

        for i in range(begin + 1, end):
            content = lines[i].split()
            if len(content) < 6 or not re.match("[0-9]", str(content[0])):
                continue

            # "账务日期" + "币种" + "机构号"
            s = fin_date + '\x1d' + curr_code + '\x1d' + org_id  + '\x1d' + lang  + '\x1d'
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


def procDir(fpath):
    print("processing directory: " + fpath)
    file_path = os.path.realpath(fpath)
    file_list = os.listdir(file_path)
    for _f in file_list:
        if _f.startswith('~$'):
            continue
        # 注意对于目录会有递归调用
        file_paths = os.path.join(file_path, _f)
        callFunc(file_paths)


def callFunc(fpath):
    if os.path.isfile(fpath):
        if fpath.lower().endswith(('.xlsx', '.xls')):
            procExcel(fpath)
        else:
            procFile(fpath)
    elif os.path.isdir(fpath):
        procDir(fpath)
    else:
        print("Unknown file type: " + fpath)


################################
callFunc(sys.argv[1])
