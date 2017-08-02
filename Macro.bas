Attribute VB_Name = "Macro"
Option Explicit
Option Base 0
Const cMod$ = cLib & ".Macro"
Sub Opn_AllWb()
jj.Opn_Wb_ByDirLsFn "P:\AppDef_Meta\", "MetaDb,MetaLgc,MetaPgm,MetaTx"
End Sub
Sub NoWrap()
Const cSub$ = "NoWrap"
Selection.WrapText = False
End Sub
Sub SetNm()
'Aim: Set the Ws Name with the column name.  Ws Name is begin with x
Dim mWs As Worksheet: Set mWs = Excel.Application.ActiveSheet
Dim mRge As Range: Set mRge = mWs.Range("A" & g.cRnoDta)
Dim mAyCnoDta() As Byte, mAnFld$(): If jj.Fnd_AyCnoDta(mAnFld, mAyCnoDta, mRge) Then GoTo E
Dim iNm As Excel.Name
Dim J%
For J = mWs.Names.Count To 1 Step -1
    If Left(mWs.Names(J).Name, 1) = "x" Then mWs.Names(J).Delete
Next
For J = 0 To UBound(mAyCnoDta)
    Dim iCno As Byte: iCno = mAyCnoDta(J)
    mWs.Names.Add "x" & mWs.Cells(jj.g.cRnoDta - 1, iCno).Value, mWs.Columns(iCno)
Next
Exit Sub
E: MsgBox "Error", , "SetNm"
End Sub
Sub CpyFormula()
'Aim:Copy formula from row Dta to row NmFld's comment
Dim mWs As Worksheet: Set mWs = Excel.Application.ActiveSheet
Dim iCno As Byte
For iCno = 1 To 254
    Dim mRgeNmFld As Range: Set mRgeNmFld = mWs.Cells(jj.g.cRnoDta - 1, iCno)
    If IsEmpty(mRgeNmFld.Value) Then Exit Sub
    Dim mRgeDta As Range: Set mRgeDta = mWs.Cells(jj.g.cRnoDta, iCno)
    jj.Dlt_Cmt mRgeNmFld
    If Left(mRgeDta.Formula, 1) = "=" Then mRgeNmFld.AddComment mRgeDta.Formula
Next
End Sub
Sub MinAllWin()
Dim iWin As Window
For Each iWin In Excel.Application.Windows
    If iWin.WindowState <> xlMinimized Then iWin.WindowState = xlMinimized
Next
End Sub
Sub HArge()
Excel.Application.Windows.Arrange xlArrangeStyleHorizontal
End Sub
Sub VArge()
Excel.Application.Windows.Arrange xlArrangeStyleVertical
End Sub
Sub InsRowAbove()
Dim mObj As Object: Set mObj = Excel.Application.Selection
If TypeName(mObj) <> "Range" Then MsgBox "A row must be selected": Exit Sub
Dim mRge As Range: Set mRge = Excel.Application.Selection
If mRge.Rows.Count <> 1 Then MsgBox "Only a row can be selected": Exit Sub
Dim mWs As Worksheet: Set mWs = mRge.Parent
jj.Shw_AllCols mWs
mRge.EntireRow.Insert xlDown
Dim mRowFm As Range: Set mRowFm = mWs.Rows(mRge.Row)
Dim mRowTo As Range: Set mRowTo = mWs.Rows(mRge.Row - 1)
mRowFm.Copy
mRowTo.PasteSpecial xlPasteAll
mRowTo.Range("B1").Value = Null
Excel.Application.CutCopyMode = False
End Sub


