import os
import xlwings as xw
import datetime
import time
import shutil


def updateGL():
    archive = r'P:\PACS\Finance\Automation\Archive\Payroll Recon01'
    path = r"P:\PACS\Finance\Month End Close"
    main_folders = os.listdir(path)
    for folder in main_folders:
        if "Old" not in folder and "cloud" not in folder and "All -" not in folder:
            print(folder)
            newpath = path + "\\" + folder
            for dirpath, dirnames, filenames in os.walk(newpath):
                for filename in [f for f in filenames if (("Payroll" in f or "Wage" in f) and f.endswith(".xlsm"))]:
                    file = os.path.join(dirpath, filename)
                    modified_time = os.path.getmtime(file)
                    date_value = datetime.datetime.fromtimestamp(modified_time)
                    month = date_value.month
                    year = date_value.year
                    if year > 2020 and month > 6:
                        # SAVE COPY OF OLD VERSION
                        new_file = os.path.join(archive, filename)
                        shutil.copy(file, new_file)
                        try:  # OPEN THE FILE AND MAKE THE CHANGES
                            wb = xw.Book(file, update_links=False)
                            win_wb = wb.api
                            for obj in win_wb.VBProject.VBComponents:
                                if obj.Name == 'Module1':
                                    obj.CodeModule.DeleteLines(1, 130)
                                    code = r'''
Sub Reconcile_Payroll()

On Error Resume Next

Dim cashreq_ws As Worksheet
Dim bank_ws As Worksheet
Dim wb As Workbook

'set worksheets
Set wb = ActiveWorkbook
Set cashreq_ws = ActiveSheet
Set bank_ws = wb.Sheets("Bank Detail")

'Excel fast working state
Application.Calculation = xlManual
Application.ScreenUpdating = False
Application.DisplayAlerts = False

bank_ws.Cells.Interior.ColorIndex = 0

Call get_bank_transactions

'start reconciliation process
x = 1
Do While KS < 50

    bank_value = bank_ws.Cells(x, 3).Value                                  'get amount from bank transaction
    If bank_ws.Cells(x, 1).Value = "Credits" Then                           'if a credit then make negative
        bank_value = bank_value * -1                                        'if a credit then make negative
    End If

    If bank_value <> 0 Then

                Set Rng = cashreq_ws.Range("D3:L50")                        'set the range to look through cash requirements report to match values
                For Each cell In Rng                                        'loop through all values of cash requirements report
                        cashreq_value = cell.Value                          'get value of cash requirement cell
                        If cashreq_value = bank_value Then                  'match amounts
                            If cell.Interior.ColorIndex <> 4 Then           'check if already reconciled
                                cell.Interior.ColorIndex = 4                'mark as reconciled
                                bank_ws.Cells(x, 3).Interior.ColorIndex = 4 'mark as reconciled
                                Exit For                                    'exit loop on successful recon
                            End If
                        End If
                Next cell
    Else
        'kill switch increase as value is not found
        KS = KS + 1
    End If


x = x + 1
Loop


'set excel back to normal
Application.ScreenUpdating = True
Application.Calculation = xlAutomatic
Application.DisplayAlerts = True
Application.StatusBar = False


MsgBox "Process was successful"
End Sub


Sub get_bank_transactions()

Dim cashreq_ws As Worksheet
Dim bank_ws As Worksheet
Dim wb As Workbook
Dim bank_report As Workbook

'set worksheets
Set wb = ActiveWorkbook
Set cashreq_ws = ActiveSheet
Set bank_ws = wb.Sheets("Bank Detail")
'Set Summary = wb.Sheets("Summary")

'Excel fast working state
Application.Calculation = xlManual
Application.ScreenUpdating = False
Application.DisplayAlerts = False

bank_ws.Cells.ClearContents

'Month and year setup for file selection
For i = 1 To 50
    If cashreq_ws.Cells(1, i).Value = "Date" Then
        Exit For
    End If
Next i
mn = Month(cashreq_ws.Cells(2, i).Value)
yr = CStr(Year(cashreq_ws.Cells(2, i).Value))
If mn = 0 Then
    mn = CStr(12)
End If
If Len(mn) = 1 Then
    mn = "0" & CStr(mn)
End If

sharedrive = "P:\PACS\Finance\Month End Close\All - Payroll Draws\" & yr & "\"
File = mn & " " & yr & " - #1027 CB&T Payroll #7775.xlsx"

Set bank_report = Application.Workbooks.Open(sharedrive & File, UpdateLinks:=0, ReadOnly:=True)
Set bank_report_ws = bank_report.Sheets("Sheet1")

UserForm1.Show

'fac_input = InputBox("Facility?")
fac_input = UserForm1.cboLocation.Value

KS = 0
b = 1
y = 1

Do While KS < 100


    If bank_report_ws.Cells(y, 14).Value = fac_input Then
        bank_report_ws.Range("A" & y & ":N" & y).Copy
        bank_ws.Range("A" & b & ":N" & b).PasteSpecial xlPasteValues
        b = b + 1
    ElseIf bank_report_ws.Cells(y, 14).Value = "" Then
        KS = KS + 1
    End If

y = y + 1
Loop

bank_report.Close

'set excel back to normal
Application.ScreenUpdating = True
Application.Calculation = xlAutomatic
Application.DisplayAlerts = True
Application.StatusBar = False

End Sub


'''
                                    obj.CodeModule.AddFromString(code)
                                    # win_wb.Application.Run('Module1.Refresh')
                                    xl = xw.apps.active.api
                                    break
                            while True:
                                try:
                                    wb.save()
                                    wb.close()
                                    xl.Quit()
                                    break
                                except:
                                    print("error exiting excel. waiting 5 seconds")
                                    time.sleep(5)
                        except Exception as e:
                            with open("GL schedules.txt", "a") as f:
                                f.write(f"Cound not run for {str(file)} due to exception {e}")
                                f.close()
                            print("Could not run for " + str(file))
                    # break


if __name__ == '__main__':
    updateGL()
