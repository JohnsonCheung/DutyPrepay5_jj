Attribute VB_Name = "xSav"
Option Compare Text
Option Explicit
Option Base 0
Const cMod$ = cLib & ".xSav"
Function Sav_Wb_AsXla(pWb As Workbook) As Boolean
On Error GoTo R
pWb.Application.DisplayAlerts = False
pWb.SaveAs pWb.FullName, Excel.XlFileFormat.xlAddIn
pWb.Application.DisplayAlerts = True
Exit Function
R: Sav_Wb_AsXla = True
End Function
Sub Sav_AllWb()
Dim iWb As Workbook, iWin As Window
For Each iWb In Excel.Application.Workbooks
    If Not iWb.Saved Then
        iWb.Windows(1).WindowState = xlMaximized
        iWb.Save
    End If
    jj.Set_Wb_Min iWb
Next
End Sub
Function Sav_Rec() As Boolean
Const cSub$ = "Sav_Rec"
On Error GoTo R
DoCmd.DoMenuItem acFormBar, acRecordsMenu, acSaveRecord, , acMenuVer70
Exit Function
R: ss.R
E: Sav_Rec = True: ss.B cSub, cMod
End Function
Function Sav_Wb(pWb As Workbook) As Boolean
Const cSub$ = "Sav_Wb"
On Error GoTo R
pWb.Save
Exit Function
R: ss.R
E: Sav_Wb = True: ss.B cSub, cMod, "Wb", jj.ToStr_Wb(pWb)
End Function
Function Sav_Wrd(pWrd As Word.Document) As Boolean
Const cSub$ = "Sav_Wrd"
On Error GoTo R
pWrd.Save
Exit Function
R: ss.R
E: Sav_Wrd = True: ss.B cSub, cMod, "Wrd", ToStr_Wrd(pWrd)
End Function


