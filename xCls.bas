Attribute VB_Name = "xCls"
Option Compare Text
Option Explicit
Option Base 0
Const cMod$ = cLib & ".xCls"
Function Cls_AllTbl() As Boolean
Dim iAcsObj As AccessObject
For Each iAcsObj In Application.CurrentData.AllTables
    With iAcsObj
        If .IsLoaded Then DoCmd.Close acTable, .Name
    End With
Next
End Function
Function Cls_Wbs(pXls As Excel.Application) As Boolean
Const cSub$ = "Cls_Wbs"
On Error GoTo R
Dim iWb As Workbook
For Each iWb In pXls.Workbooks
    If jj.Cls_Wb(iWb, True) Then ss.A 1: GoTo E
Next
Exit Function
R: ss.R
E: Cls_Wbs = True: ss.B cSub, cMod
End Function
Function Cls_CurDb(pAcs As Access.Application) As Boolean
On Error Resume Next
pAcs.CloseCurrentDatabase
End Function
Function Cls_Db(oDb As DAO.Database) As Boolean
On Error Resume Next
oDb.Close
Set oDb = Nothing
End Function
Function Cls_Frm(pNmFrm$) As Boolean
Const cSub$ = "Cls_Frm"
On Error GoTo R
DoCmd.Close acForm, pNmFrm, acSaveYes
Exit Function
R: ss.R
E: Cls_Frm = True: ss.B cSub, cMod, "pNmFrm", pNmFrm
End Function
Function Cls_Frm_Tst() As Boolean
Const cSub$ = "Cls_Frm_Tst"
jj.Shw_Dbg cSub, cMod, "Result", jj.Cls_Frm("1Rec")
End Function
Function Cls_Rs(pRs1 As DAO.Recordset _
    , Optional pRs2 As DAO.Recordset = Nothing _
    , Optional pRs3 As DAO.Recordset = Nothing _
    , Optional pRs4 As DAO.Recordset = Nothing _
    , Optional pRs5 As DAO.Recordset = Nothing _
    ) As Boolean
On Error Resume Next
pRs1.Close: If IsNothing(pRs2) Then Exit Function
pRs2.Close: If IsNothing(pRs3) Then Exit Function
pRs3.Close: If IsNothing(pRs4) Then Exit Function
pRs4.Close: If IsNothing(pRs5) Then Exit Function
pRs5.Close
End Function
Function Cls_Wb(pWb As Workbook, Optional pSav As Boolean = False, Optional pSilent As Boolean = False) As Boolean
Const cSub$ = "Cls_Wb"
On Error GoTo R
Dim mXls As Excel.Application: Set mXls = pWb.Application
mXls.DisplayAlerts = False
pWb.Close pSav
mXls.DisplayAlerts = True
Exit Function
R: ss.R
E: Cls_Wb = True: If Not pSilent Then ss.B cSub, cMod, "Wb,Sav", jj.ToStr_Wb(pWb), pSav
End Function
Function Cls_Wrd(pWrd As Word.Document, Optional pSav As Boolean = False, Optional pSilent As Boolean = False) As Boolean
Const cSub$ = "Cls_Wd"
On Error GoTo R
Dim mWrd As Word.Application: Set mWrd = Word.Application
mWrd.DisplayAlerts = False
pWrd.Save
mWrd.DisplayAlerts = True
Exit Function
R: ss.R
E: Cls_Wrd = True: If Not pSilent Then ss.B cSub, cMod, "pWrd,pSav", ToStr_Wrd(pWrd), pSav
End Function
Function Cls_Ppt(pPpt As PowerPoint.Presentation, Optional pSav As Boolean = False) As Boolean
Const cSub$ = "Cls_Ppt"
On Error GoTo E
Dim mPpt As PowerPoint.Application: Set mPpt = pPpt.Application
mPpt.DisplayAlerts = False
pPpt.Save
pPpt.Close
mPpt.DisplayAlerts = True
Exit Function
E: Cls_Ppt = True
End Function

