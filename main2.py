import os
import xlwings as xw
import datetime

wb1 = xw.Book(r"C:\Users\tyler.anderson\Documents\Finance\GL Codes List for Schedules.xlsm", update_links=False)
ws = wb1.sheets['Summary']
data = ws.range("B5:B46").value
xw.apps.active.quit()


def updateGL():
    path = r"P:\PACS\Finance\Month End Close"
    path = r"C:\Users\tyler.anderson"
    main_folders = os.listdir(path)
    for folder in main_folders:
        if "Old" not in folder and "cloud" not in folder and "All -" not in folder:
            print(folder)
            newpath = path + "\\" + folder
            for dirpath, dirnames, filenames in os.walk(newpath):
                for filename in [f for f in filenames if f.endswith("GL Schedules.xlsm")]:
                    file = os.path.join(dirpath, filename)
                    print(file)
                    modified_time = os.path.getmtime(file)
                    date_value = datetime.datetime.fromtimestamp(modified_time)
                    print(f"{date_value:%Y-%m-%d %H:%M:%S}")
                    month = date_value.month
                    year = date_value.year
                    if year > 2020 and month > 3:
                        wb = xw.Book(file, update_links=False)
                        sum_sht = wb.sheets('Summary')
                        for i, code in enumerate(data):
                            sum_sht.range("B" + str(i+5)).value = code
                        addin_file = xw.Book(r'C:\Users\tyler.anderson\AppData\Roaming\Microsoft\AddIns\1005-Duplicate Sheet.xlam', update_links=False)
                        macro = addin_file.macro('DupSheet.updateSummary')
                        macro()
                        wb.save()
                        wb.close()
                    break
    xw.apps.active.quit()


if __name__ == '__main__':
    updateGL()
