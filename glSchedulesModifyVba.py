import os
import xlwings as xw
import datetime
import time
import shutil


def updateGL():
    archive = r'P:\PACS\Finance\Automation\Archive\GL Schedules 2'
    path = r"P:\PACS\Finance\Month End Close"
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
                    month = date_value.month
                    year = date_value.year
                    if year > 2020 and month > 4:
                        # SAVE COPY OF OLD VERSION
                        new_file = os.path.join(archive, filename)
                        shutil.copy(file, new_file)
                        try:  # OPEN THE FILE AND MAKE THE CHANGES
                            wb = xw.Book(file, update_links=False)
                            win_wb = wb.api
                            for obj in win_wb.VBProject.VBComponents:
                                if obj.Name == 'Module1':
                                    obj.CodeModule.DeleteLines(1, 168)
                                    code = '''Sub Add_Acct()
    
    On Error GoTo handler
    
    'EXCEL FAST WORKING STATE
    Application.Calculation = xlManual
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Dim myValue As Variant
    Dim sht As Worksheet
    
    'INPUT SHEET NAME
    myValue = InputBox("Enter the account number you want to create a schedule for (example: 1023.5):")
    
    'IF USER CANCLES OR DOES NOT ENTER AN ACCOUNT CODE
    If (StrPtr(myValue) = 0) Or (myValue = "") Then
        'SET EXCEL BACK TO NORMAL AND EXIT
        Application.ScreenUpdating = True
        Application.Calculation = xlAutomatic
        Application.DisplayAlerts = True
        Exit Sub
    ElseIf IsNumeric(myValue) = False Then
        'SET EXCEL BACK TO NORMAL AND EXIT
        Application.ScreenUpdating = True
        Application.Calculation = xlAutomatic
        Application.DisplayAlerts = True
        MsgBox "Account number not valid. Please enter a valid account code"
        Exit Sub
    End If
    
    'CHECK IF SHEET ALREADY EXISTS
    For Each sht In Worksheets
        If sht.Name = myValue Then
            'SET EXCEL BACK TO NORMAL AND EXIT
            Application.ScreenUpdating = True
            Application.Calculation = xlAutomatic
            Application.DisplayAlerts = True
            MsgBox "This tab already exists.  Please use a different name"
            Exit Sub
        End If
    Next
    
    'ADD THE TAB
    Sheets("Template").Copy after:=Sheets(1)
    Sheets("Template (2)").Name = myValue
    Sheets(myValue).Visible = True
    Sheets(myValue).Cells(2, 1).Value = myValue
    Sheets(myValue).Cells(2, 1).NumberFormat = "0.000"
    
    'ALPHABATIZE SHEETS
    For i = 1 To Application.Sheets.Count
        For j = 1 To Application.Sheets.Count - 1
            If UCase$(Application.Sheets(j).Name) > UCase$(Application.Sheets(j + 1).Name) Then
                Sheets(j).Move after:=Sheets(j + 1)
            End If
        Next
    Next
    
    Sheets("Summary").Move Before:=Sheets(1)
    
    'UPDATE SUMMARY PAGE
    x = 5
    Set ws = ThisWorkbook.Sheets("Summary")
    
    For Each sht In Worksheets
        If IsNumeric(sht.Name) = True Then
            ws.Cells(x, 2).Value = sht.Name
            x = x + 1
        End If
    Next
        
        'EXTEND FORMULAS TO THE BOTTOM
        ws.Range("A5:A" & (x - 1)).FillDown
        ws.Range("C5:E" & (x - 1)).FillDown
        
    'SET EXCEL BACK TO NORMAL AND END
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    Application.DisplayAlerts = True
    
    ThisWorkbook.RefreshAll
    
    Exit Sub
handler:
        'SET EXCEL BACK TO NORMAL
        Application.ScreenUpdating = True
        Application.Calculation = xlAutomatic
        Application.DisplayAlerts = True
        
        MsgBox "There was an issue adding the tab"
    
    
    End Sub
    
    Sub Refresh()
    
        On Error GoTo handle
    
        'Excel fast working state
        Application.Calculation = xlManual
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
    
'        Dim wb As Workbook
'        Dim ws As Worksheet
'        Dim sumws As Worksheet
'        Dim listwb As Workbook
'        Dim coaws As Worksheet
'        Dim buildingsdw As Worksheet
'
'        Set wb = ThisWorkbook
'        Set ws = wb.Sheets("Lists")
'        Set sumws = wb.Sheets("Summary")
'
'        ' OPEN LIST SHEET
'        Set listwb = Application.Workbooks.Open("P:\PACS\Finance\Automation\Data Files\Lists.xlsx", ReadOnly:=True)
'        listwb.EnableConnections
'        Set coaws = listwb.Sheets("COA")
'        Set buildingsdw = listwb.Sheets("BuildingsDW")
'
'        ' GET LAST ROW ADDRESS
'        coa_LRow = coaws.Cells(Rows.Count, 1).End(xlUp).Row
'        buildings_LRow = buildingsdw.Cells(Rows.Count, 1).End(xlUp).Row
'
'        ' COPY OVER FROM ONE SHEET TO THE OTHER
'        coaws.Range("A1:B" & coa_LRow).Copy ws.Cells(1, 1)
'        buildingsdw.Range("A1:A" & buildings_LRow).Copy ws.Cells(1, 4)
'
'        building_range = ws.Range("D1:D" & buildings_LRow).Address
'
'        ' UPDATED DATA VALIDATION
'        With sumws.Range("A1").Validation
'            .Delete
'            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
'            xlBetween, Formula1:="=Lists!" & building_range
'            .IgnoreBlank = True
'            .InCellDropdown = True
'            .InputTitle = ""
'            .ErrorTitle = ""
'            .InputMessage = ""
'            .ErrorMessage = ""
'            .ShowInput = True
'            .ShowError = True
'        End With
'
'        listwb.Close
    
        ThisWorkbook.RefreshAll
        
        'set excel back to normal
        Application.ScreenUpdating = True
        Application.Calculation = xlAutomatic
        Application.DisplayAlerts = True
        Application.StatusBar = False
        Exit Sub
        
handle:
        
        'set excel back to normal
        Application.ScreenUpdating = True
        Application.Calculation = xlAutomatic
        Application.DisplayAlerts = True
        Application.StatusBar = False
    
    
    End Sub'''
                                    obj.CodeModule.AddFromString(code)
                                    win_wb.Application.Run('Module1.Refresh')
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
