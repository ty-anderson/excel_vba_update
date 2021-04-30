import os
import xlwings as xw
import datetime
import time
import pandas as pd


def updateGL():
    wb1 = xw.Book(r"C:\Users\tyler.anderson\Documents\Finance\GL Codes List for Schedules.xlsm", update_links=False)
    ws = wb1.sheets['Summary']
    data = ws.range("B5:B46").value
    df = pd.DataFrame(data)
    xw.apps.active.quit()
    path = r"P:\PACS\Finance\Month End Close"
    path = r"C:\Users\tyler.anderson"
    main_folders = os.listdir(path)
    for folder in main_folders:
        if "Old" not in folder and "cloud" not in folder and "All -" not in folder:
            print(folder)
            newpath = path + "\\" + folder
            for dirpath, dirnames, filenames in os.walk(newpath):
                for filename in [f for f in filenames if (("GL Schedules" in f) and f.endswith(".xlsm"))]:
                    file = os.path.join(dirpath, filename)
                    modified_time = os.path.getmtime(file)
                    date_value = datetime.datetime.fromtimestamp(modified_time)
                    print(file)
                    print(f"{date_value:%Y-%m-%d %H:%M:%S}")
                    month = date_value.month
                    year = date_value.year
                    if year > 2020 and month > 3:
                        wb = xw.Book(file, update_links=False)
                        app = xw.apps.active
                        win_wb = wb.api
                        for obj in win_wb.VBProject.VBComponents:
                            print(obj.Name)
                        sum_sht = wb.sheets('Summary')
                        sum_sht.range("B5").value = df.values
                        addin_file = xw.Book(r'C:\Users\tyler.anderson\AppData\Roaming\Microsoft\AddIns\1005-Duplicate Sheet.xlam', update_links=False)
                        macro = addin_file.macro('DupSheet.updateSummary')
                        macro()
                        while True:
                            try:
                                wb.save()
                                wb.close()
                                break
                            except:
                                time.sleep(5)
                    break
    xw.apps.active.quit()


if __name__ == '__main__':
    updateGL()
