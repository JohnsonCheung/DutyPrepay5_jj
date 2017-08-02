Attribute VB_Name = "xClr"
#Const Tst = True
Option Compare Text
Option Explicit
Option Base 0
Const cMod$ = cLib & ".xClr"
Function Clr_Qt(pWs As Worksheet) As Boolean
While pWs.QueryTables.Count > 0
    pWs.QueryTables(1).Delete
Wend
End Function
Function Clr_AyV(oAyV()) As Boolean
Dim mAyV(): oAyV = mAyV
End Function
Sub Clr_AyLng(oAyLng&())
Dim mAyLng&(): oAyLng = mAyLng
End Sub
Function Clr_ImpWs(pRge As Range, Optional pRithImp% = -4) As Boolean
'Aim: Clean up {mWs} for input.
'       Format of mWs is: (See jj.Exp_Nmtq2Xls_wFmt)
'       it must be in following format
'       - #1: A1      : A1 must be in format of import:<mNmWsTar>
'       - #2: Rno4    : It the rno of field name.  All vbstring fields will be used as field name.
'                     - The cmt of the field name is used as Rno5's formula and will copy to all records
'       - #3: Rno5    : Is the data rno
'       - #4: SubTot  : Col B is be subtotal to count the # of records
Const cSub$ = "Clr_ImpWs"
On Error GoTo R
Dim mAmFld() As tMap, mAyCno() As Byte: If jj.Fnd_AyCnoImpFld(mAyCno, mAmFld, pRge, pRithImp) Then ss.A 1: GoTo E

Dim mWs As Worksheet: Set mWs = pRge.Worksheet
mWs.Activate
mWs.Application.ActiveWindow.FreezePanes = False
Shw_Sts "Clear ImpWs [" & mWs.Name & "] ..."
If jj.Shw_AllDta(mWs) Then ss.A 2: GoTo E
If jj.Clr_OutLine(mWs) Then ss.A 3: GoTo E

'Delete all the rows
Dim mRge As Range: Set mRge = pRge(1, 2)
Dim mRnoLas&
If IsEmpty(mRge) Then
    mRnoLas = pRge.Row - 1
Else
    mRnoLas = mRge.End(xlDown).Row
End If
If mRnoLas = 65536 Then mWs.Rows(mRnoLas + 1 & ":65536").EntireRow.Delete

'Delete the not using columns by referring using columns mAyCno
Do
        
    'mAyCnoToDlt
    Dim mNFld As Byte: mNFld = jj.Siz_Ay(mAyCno)
    Dim mCnoLas As Byte: mCnoLas = mAyCno(mNFld - 1)
    mWs.Cells.Copy
    mWs.Cells.PasteSpecial xlPasteValues
    'Delete all columns after mNCol
    mWs.Range(mWs.Cells(1, mCnoLas + 1), mWs.Cells(1, 256)).EntireColumn.Delete
    
    If mCnoLas > mNFld Then
        ReDim mAyCnoToDlt(mCnoLas - mNFld - 1) As Byte
        Dim mN%, iCno As Byte
        For iCno = 1 To mCnoLas
            Dim iIdx%: If jj.Fnd_IdxByt(iIdx, mAyCno, iCno) Then ss.A 1: GoTo E
            If iIdx < 0 Then mAyCnoToDlt(mN) = iCno: mN = mN + 1
        Next
    
        'Delete all column stored in mAyCnoToDlt()
        Dim J%
        For J = mN - 1 To 0 Step -1
            mWs.Columns(mAyCnoToDlt(J)).Delete
        Next
    End If
    
Loop Until True

'Delete Row1-3
mWs.Rows("1:3").Delete

mWs.Range("A1").Select
mWs.Range("A1").Activate
Exit Function
R: ss.R
E: Clr_ImpWs = True: ss.B cSub, cMod, "mWs", ToStr_Ws(mWs)
X: Clr_Sts
End Function
#If Tst Then
Function Clr_ImpWs_Tst() As Boolean
Dim mWb As Workbook: If jj.Opn_Wb_R(mWb, "p:\appdef_Meta\MetaDb.xls", , True) Then Stop: GoTo E
Dim mWs As Worksheet: Set mWs = mWb.Sheets("TblF")
If Clr_ImpWs(mWs.Range("A5")) Then Stop: GoTo E
mWs.Application.Visible = True
Stop
GoTo X
E:
X: jj.Cls_Wb mWb, False, True
End Function
#End If
'Sub Clr_Doc(oDoc As MSXML2.DOMDocument60)
'If xIs.IsNothing(oDoc) Then Set oDoc = New MSXML2.DOMDocument60: Exit Sub
'oDoc.Load ""
'End Sub
'#If Tst Then
'Function Clr_Doc_Tst() As Boolean
'Dim mDoc As New MSXML2.DOMDocument60
'mDoc.loadXML "<A></A>"
'Debug.Print mDoc.ChildNodes.Length
'jj.Clr_Doc mDoc
'Debug.Print mDoc.ChildNodes.Length
'End Function
'#End If
Sub Clr_DPgm(oDPgm As d_Pgm)
If TypeName(oDPgm) = "Nothing" Then Set oDPgm = New d_Pgm
With oDPgm
    .x_IsPrivate = False
    .x_TypFct = 0
    .x_NmPrc = ""
    .x_NmTypRet = ""
    .x_Aim = ""
    .x_PrcBody = ""
End With
End Sub
Sub Clr_AyDArg(oAyDArg() As d_Arg)
Dim mAyDArg() As d_Arg: oAyDArg = mAyDArg
End Sub
Sub Clr_Am(oAm() As tMap)
Dim mAm() As tMap: oAm = mAm
End Sub
Function Clr_OutLine(pWs As Worksheet) As Boolean
Const cSub$ = "Clr_OutLine"
'Aim: Clear all columns' outline
Dim J As Byte
For J = 1 To 9
    On Error GoTo X
    pWs.Cells.EntireColumn.Ungroup
Next
X:
End Function
Function Clr_Sess(pTrc&, pMaxSess%, pNmLgs$) As Boolean
Const cSub$ = "Clr_Sess"
'Aim: Clear the sess directory if exist, otherwise, create.
'     Sess dir is @ {DirTmp}Tp_{App}{NmLgs}{NnnnNnnn}
Dim mDir$: mDir = jj.Sdir_TmpLgc$ & pNmLgs & "\"
Dim mAyDir$(): If jj.Fnd_AyDir(mAyDir, mDir) Then ss.A 1: GoTo E
Dim N%: N = jj.Siz_Ay(mAyDir)
If pMaxSess >= N Then GoTo X

'Set Top pMaxSess elements in mAyDir() to empty
Dim J%
For J = 1 To pMaxSess
    Dim mIdx%: jj.Fnd_MaxEle mIdx, mAyDir
    mAyDir(mIdx) = ""
Next
For J = 0 To N - 1
    If mAyDir(J) <> "" Then
        Dim mA$: mA = mDir & mAyDir(J) & "\"
        jj.Dlt_Dir mA
        On Error Resume Next
        RmDir mA
        On Error GoTo 0
    End If
Next
GoTo X
R: ss.R
E: Clr_Sess = True: ss.B cSub, cMod, "pTrc&, pMaxSess%, pNmLgs$", pTrc&, pMaxSess%, pNmLgs$
X: jj.Sdir_TmpSess pNmLgs, pTrc
End Function
#If Tst Then
Function Clr_Sess_Tst() As Boolean
If jj.Clr_Sess(5, 3, "MPS") Then Stop
End Function
#End If
Sub Clr_Ays(oAys$())
Dim mA$(): oAys = mA
End Sub
#If Tst Then
Function Clr_AyLng_Tst() As Boolean
Dim mAyLng&()
ReDim mAyLng(10)
Dim J%
For J = 0 To 10
    mAyLng(J) = J * 10
Next
Stop
jj.Clr_AyLng mAyLng
Stop
End Function
#End If
Function Clr_AyByt(oAyByt() As Byte) As Boolean
Dim mAyByt() As Byte: oAyByt = mAyByt
End Function
Function Clr_Sq(oSq As tSq) As Boolean
With oSq
    .c1 = 0
    .c2 = 0
    .r1 = 0
    .r2 = 0
End With
End Function
Function Clr_Sts() As Boolean
If jj.IsAcs Then Application.SysCmd acSysCmdClearStatus
End Function
Function Clr_Tp() As Boolean
Const cSub$ = "Clr_Tp"
Dim mDirTp$: mDirTp = jj.Sdir_Tp
'-- Delete all tmp file records & insert one dummy record for those tmp*Output*
Dim iTbl As DAO.TableDef: For Each iTbl In CurrentDb.TableDefs
    If Left(iTbl.Name, 3) = "tmp" Then
        Debug.Print ": Deleting tmp table -->" & iTbl.Name
        If jj.Run_Sql("Delete * from [" & iTbl.Name & "]") Then ss.A 1: GoTo E
        If InStr(iTbl.Name, "Output") > 0 Then
            With CurrentDb.TableDefs(iTbl.Name).OpenRecordset
                .AddNew
                .Update
            End With
        End If
    End If
Next
'Loop Template file
Dim AyFn$(): If jj.Fnd_AyFn(AyFn, mDirTp, "*.xls", False) Then ss.A 1: GoTo E
If jj.Siz_Ay(AyFn) = 0 Then ss.A 1, "No template files found", eRunTimErr, "DirTp", mDirTp: GoTo E
Dim iFil
For Each iFil In AyFn
    Debug.Print "******************************"
    Debug.Print "******************************"
    Debug.Print "Open Xls: " & iFil
    Dim mWb As Workbook: If jj.Opn_Wb_RW(mWb, mDirTp & iFil) Then Stop
    If jj.Rfh_Wb(mWb) Then ss.A 1: GoTo E
    jj.Cls_Wb mWb, True
Next
MsgBox "All Templates Cleared"
Exit Function
R: ss.R
E: Clr_Tp = True: ss.B cSub, cMod, "Tp dir:jj.Sdir_Tp", mDirTp
End Function


