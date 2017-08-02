Attribute VB_Name = "xCpy"
#Const Tst = True
Option Compare Text
Option Explicit
Option Base 0
Const cMod$ = cLib & ".xCpy"
Function Cpy_WsVal(oWs As Worksheet, pFmWs As Worksheet, pToNmWs$) As Boolean
Const cSub$ = "Cpy_WsVal"
If jj.Add_Ws(oWs, pFmWs.Parent, pToNmWs) Then ss.A 1: GoTo E
pFmWs.Cells.Copy
oWs.Select
oWs.Application.Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
oWs.Application.Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Exit Function
R: ss.R
E: Cpy_WsVal = True: ss.B cSub, cMod, "pFmWs,pToNmWs", jj.ToStr_Ws(pFmWs), pToNmWs
End Function
Function Cpy_Row(pWs As Worksheet, pRnoFm&, pRnoTo&) As Boolean
'Aim: Copy {pRnoFm} to {pRnoTo} in pWs
Const cSub$ = "Cpy_Row"
On Error GoTo R
Dim mRowFm As Range: Set mRowFm = pWs.Range(pRnoFm & ":" & pRnoFm)
Dim mRowTo As Range: Set mRowTo = pWs.Range(pRnoTo & ":" & pRnoTo)
mRowFm.Copy
mRowTo.PasteSpecial xlPasteAllExceptBorders
pWs.Application.CutCopyMode = False
Exit Function
R: ss.R
E: Cpy_Row = True: ss.B cSub, cMod, "pWs,pRnoFm,pRnoTo", jj.ToStr_Ws(pWs), pRnoFm, pRnoTo
End Function
Function Cpy_FmRs(oNRec&, pRge As Range, pSqlMemo$) As Boolean
'Aim: simulate pRge.CopyFromRecordSet to handle memo field having string of len >255.
Const cSub$ = "Cpy_FmRs"
oNRec = 0
Dim mCase As Byte: mCase = 2
Select Case mCase
Case 1
    Dim mRs As DAO.Recordset
    If jj.Opn_Rs(mRs, pSqlMemo) Then ss.A 1: GoTo E
    With mRs
        Dim J As Byte
        Dim NFld%: NFld = mRs.Fields.Count
        While Not .EOF
            oNRec = oNRec + 1
            For J = 0 To NFld - 1
                pRge.Cells(oNRec, J + 1).Value = .Fields(J).Value
            Next
            .MoveNext
        Wend
        .Close
    End With
Case 2
    Dim mFxTmp$: mFxTmp = jj.Sdir_Tmp & Format(Now, "YYYYMMDD HHMMSS") & ".xls"
    Dim mCnnStr$: mCnnStr = jj.Q_S(jj.CnnStr_Xls(mFxTmp), "[*].Tmp")
    'jj.Fmt_Str
    Dim mSql$: mSql = jj.Fmt_Str("Select * into {0} from ({1})", mCnnStr, pSqlMemo)
    If jj.Run_Sql(mSql) Then ss.A 2: GoTo E
    Dim mWb As Workbook: If jj.Opn_Wb_R(mWb, mFxTmp) Then ss.A 3: GoTo E
    With mWb.Names("Tmp").RefersToRange
        .Rows(1).Delete
        .Copy
        oNRec = .Rows.Count
    End With
    pRge.PasteSpecial xlPasteValues
    jj.Cls_Wb mWb
    jj.Dlt_Fil mFxTmp
End Select
Exit Function
R: ss.R
E: Cpy_FmRs = True: ss.B cSub, cMod, "pRge,pSqlMemo", jj.ToStr_Rge(pRge), pSqlMemo
X: jj.Cls_Rs mRs
End Function
#If Tst Then
Function Cpy_FmRs_Tst() As Boolean
Const cSub$ = "Cpy_FmRs_Tst"
If jj.Crt_Tbl_FmLoFld("#Tmp", "aa text 10, bb memo") Then Stop: GoTo E
With CurrentDb.TableDefs("#Tmp").OpenRecordset
    Dim J%
    For J = 0 To 10
        .AddNew
        !aa = J
        !BB = String(245, Chr(Asc("0") + J))
        .Update
    Next
    For J = 0 To 10
        .AddNew
        !aa = J
        !BB = String(500, Chr(Asc("0") + J))
        .Update
    Next
    .Close
End With
Dim mWb As Workbook: If jj.Crt_Wb(mWb, "c:\aa.xls", True, "aa") Then Stop: GoTo E
Dim mRge As Range: Set mRge = mWb.Sheets(1).Range("e5")
Dim cSel$: cSel = "Select * from [#Tmp]"
Dim mNRec&: If jj.Cpy_FmRs(mNRec, mRge, cSel) Then Stop: GoTo E
mWb.Application.Visible = True
Exit Function
E:
End Function
#End If
Function Cpy_Ws(pWbTo As Workbook, pWbFm As Workbook, pNmWsFm$) As Boolean
'Aim: Copy {pWbFm}!{pNmWsFm$} to {pWbTo}.  If ws exist in {pWbTo}, ws will be replaced and position retended, else copy to end
Const cSub$ = "Cpy_Ws"
Dim mWsFm As Worksheet: If jj.Fnd_Ws(mWsFm, pWbFm, pNmWsFm) Then ss.A 1: GoTo E
Dim mWsTo As Worksheet
If jj.Fnd_Ws(mWsTo, pWbTo, pNmWsFm) Then
    If jj.Add_Ws(mWsTo, pWbTo, pNmWsFm) Then ss.A 2: GoTo E
    mWsFm.Copy , mWsTo
Else
    mWsTo.Cells.Clear
End If
mWsFm.Cells.Copy
mWsTo.Range("A1").PasteSpecial xlPasteAll
mWsFm.Application.CutCopyMode = False
Exit Function
R: ss.R
E: Cpy_Ws = True: ss.B cSub, cMod, "pWbTo,pWbFm,pNmWsFm", jj.ToStr_Wb(pWbTo), jj.ToStr_Wb(pWbFm), pNmWsFm
End Function
#If Tst Then
Function Cpy_Ws_Tst() As Boolean
Dim mFxFm$: mFxFm = "c:\tmp\Fm.xls"
Dim mFxTo$: mFxTo = "c:\tmp\To.xls"
Dim mWbFm As Workbook: If jj.Crt_Wb(mWbFm, mFxFm, True) Then Stop
Dim mWbTo As Workbook: If jj.Crt_Wb(mWbTo, mFxTo, True) Then Stop
Dim mWsFm As Worksheet: Set mWsFm = mWbFm.Sheets.Add
If jj.Set_Ws_ByLpAp(mWsFm, 1, 1, False, "Msg", "First Time From Wb") Then Stop
mWbFm.Application.Visible = True
MsgBox "Before Copy, Check To ws"
Stop
If jj.Cpy_Ws(mWbTo, mWbFm, mWsFm.Name) Then Stop
MsgBox "First Time Copy.  Check To ws"
Stop
If jj.Set_Ws_ByLpAp(mWsFm, 1, 1, False, "Msg", "Second Time From Wb") Then Stop
If jj.Cpy_Ws(mWbTo, mWbFm, mWsFm.Name) Then Stop
MsgBox "Second Time Copy.  Check To ws"
End Function
#End If
Function Cpy_Formula_ByCmt(pRge As Range, pNRec&) As Boolean
'Aim: Assume the row above pRge contains formula of the row pRge in Cmt.
'     After setting the formula of each cell of row pRge, copy them downward until pRnoEnd
Const cSub$ = "Cpy_Formula_ByCmt"
On Error GoTo R
Dim mCnoLas As Byte: If jj.Fnd_CnoLas(mCnoLas, pRge) Then ss.A 1: GoTo E
On Error GoTo R
Dim iCno As Byte
Dim mWs As Worksheet: Set mWs = pRge.Parent
For iCno = pRge.Column To mCnoLas
    Dim mRgeFormula As Range: Set mRgeFormula = mWs.Cells(pRge.Row - 1, iCno)
    If jj.IsCmt(mRgeFormula) Then
        Dim mFormula$: mFormula = mRgeFormula.Comment.Text
        If jj.Set_Formula(mRgeFormula(2, 1), pNRec, mFormula) Then ss.A 3, "Error in set formula for column[" & iCno & "]": GoTo E
    End If
Next
pRge.Application.CutCopyMode = False
Exit Function
R: ss.R
E: Cpy_Formula_ByCmt = True: ss.B cSub, cMod, "pRge,pNRec", jj.ToStr_Rge(pRge)
End Function
Function Cpy_Rge2Ws_ByXlsNm(pWbSrc As Workbook, pXlsNmSrc$, pWsTar As Worksheet) As Boolean
'Aim: Copy the range as pointed by {pXlsNmSrc} in {pWbSrc} to {pWsTar}
Const cSub$ = "Cpy_Rge2Ws_ByXlsNm"
On Error GoTo R
Dim mXlsNm As Excel.Name: Set mXlsNm = pWbSrc.Names(pXlsNmSrc)
Dim mRge As Range: Set mRge = mXlsNm.RefersToRange
mRge.Copy pWsTar.Range("A1")
Exit Function
R: ss.R
E: Cpy_Rge2Ws_ByXlsNm = True: ss.B cSub, cMod, "pWbSrc,pXlsNmSrc,pWbTar", jj.ToStr_Wb(pWbSrc), pXlsNmSrc, jj.ToStr_Ws(pWsTar)
End Function
#If Tst Then
Function Cpy_Rge2Ws_ByXlsNm_Tst() As Boolean
Const cFxTar$ = "C:\aa.xls"
Const cFx$ = "p:\Workingdir\Meta Db.xls"
Dim mWbTar As Workbook: If jj.Crt_Wb(mWbTar, cFxTar, True) Then Stop
Dim mWsTar As Worksheet: Set mWsTar = mWbTar.Sheets(1)
Dim mWbSrc As Workbook: If jj.Opn_Wb_R(mWbSrc, cFx) Then Stop
If jj.Cpy_Rge2Ws_ByXlsNm(mWbSrc, "DefTbl", mWsTar) Then Stop
mWbTar.Application.Visible = True
End Function
#End If
Function Cpy_Rge2Xls_ByXlsNm(pWbSrc As Workbook, pXlsNmSrc$, pFxTar$, Optional pNmWsTar$ = "", Optional pOvrWrt As Boolean = False) As Boolean
Const cSub$ = "Cpy_Rge2Xls_ByXlsNm"
'Aim: Copy the range as defined in {pXlsNmSrc} in {pWbSrc} to of {pNmWsTar} in {pFxTar}.  If {pNmWsTar} is '', use {pXlsNmSrc}
On Error GoTo R
Dim mWbTar As Workbook: If jj.Crt_Wb(mWbTar, pFxTar, pOvrWrt) Then ss.A 1: GoTo E
If jj.Dlt_AllWs_Except1(mWbTar) Then ss.A 2: GoTo E

If pNmWsTar = "" Then pNmWsTar = pXlsNmSrc
Dim mWsTar As Worksheet: Set mWsTar = mWbTar.Sheets(1)
mWsTar.Name = pNmWsTar
If jj.Cpy_Rge2Ws_ByXlsNm(pWbSrc, pXlsNmSrc, mWsTar) Then ss.A 3: GoTo E
If jj.Cls_Wb(mWbTar, True) Then ss.A 1: GoTo E
Exit Function
R: ss.R
E: Cpy_Rge2Xls_ByXlsNm = True: ss.B cSub, cMod, "pWbSrc,pXlsNmSrc,pFxTar", jj.ToStr_Wb(pWbSrc), pXlsNmSrc, pFxTar
End Function
#If Tst Then
Function Cpy_Rge2Xls_ByXlsNm_Tst() As Boolean
Const cFxTar$ = "C:\aa.xls"
Const cFx$ = "P:\WorkingDir\META Db.xls"
Dim mWb As Workbook: If jj.Opn_Wb(mWb, cFx, True) Then Stop: GoTo E
If jj.Cpy_Rge2Xls_ByXlsNm(mWb, "DefTbl", cFxTar, "Tbl", True) Then Stop: GoTo E
If jj.Cls_Wb(mWb) Then Stop: GoTo E
If jj.Opn_Wb(mWb, cFxTar, , , True) Then Stop: GoTo E
Exit Function
E: Cpy_Rge2Xls_ByXlsNm_Tst = True
End Function
#End If
Function Cpy_AndOpn(oWb As Workbook, pFxFm$, pFxTo$, Optional pOvrWrt As Boolean = False) As Boolean
Const cSub$ = "Cpy_AndOpn"
If VBA.Dir(pFxFm) = "" Then ss.A 1, "From file not exist": GoTo E

'If <pOvrWrt>, delete <pFxTo> if exist, else prompt to overwrite if exist.
If VBA.Dir(pFxTo) <> "" Then
    If pOvrWrt Then
        If jj.Dlt_Fil(pFxTo) Then
            Dim mMsg$: mMsg = "Target Xls file [" & pFxTo & "] cannot be overwritten (or killed)||" & _
                "Check:|" & _
                "1. Check if the Target Xls is openned.  If is openned, close it and re-run||" & _
                "2. Otherwise, do following:|" & _
                "   1 Close all Xls files|" & _
                "   2 Press [Ctrl]+[Alt]+[Delete], Click [Task Manager] button|" & _
                "   3 A window [Windows Task Manager] is displayed.  Click [Processes] page ta|" & _
                "   4 Click the column [Image Name] to sort [Image Name]|" & _
                "   5 If there is any [Excel.exe] in the column [Image Name], highlight it [Excel.Exe] and Click [End Process]|" & vbLf & _
                "   6 Repeat [5] until no more [Excel.exe] in the column [Image Name]|" & _
                "   7 Re-run the program"
            ss.A 2, mMsg: GoTo E
        End If
    Else
        ss.A 3, "To file exist": GoTo E
    End If
End If
'Copy <pFxFm> to <pFxTo> and open <pFxTo> in mWb
If jj.Cpy_Fil(pFxFm, pFxTo) Then ss.A 4: GoTo E
gXls.AutomationSecurity = msoAutomationSecurityForceDisable
Set oWb = gXls.Workbooks.Open(pFxTo, UpdateLinks:=False, IgnoreReadOnlyRecommended:=True)
Exit Function
R: ss.R
E: Cpy_AndOpn = True: ss.B cSub, cMod, "pFxFm,pFxTo,pOvrWrt", pFxFm, pFxTo, pOvrWrt
End Function
Function Cpy_Ay(oAy$(), pAy$()) As Boolean
Dim N%: N = UBound(pAy) + 1
ReDim oAy(N - 1)
Dim J%: For J = 0 To N% - 1
    oAy(J) = pAy(J)
Next
End Function
Function Cpy_Am(oAm_To() As tMap, pAm_Fm() As tMap, Optional pSwap As Boolean = False) As Boolean
Dim N%: N% = jj.Siz_Am(pAm_Fm): If N% = 0 Then jj.Clr_Am oAm_To: GoTo E
ReDim oAm_To(0 To N% - 1)
Dim J%
If pSwap Then
    For J% = 0 To N% - 1
        With pAm_Fm(J%)
            oAm_To(J%).F1 = .F2
            oAm_To(J%).F2 = .F1
        End With
    Next
    Exit Function
End If
For J% = 0 To N% - 1
    With pAm_Fm(J%)
        oAm_To(J%).F1 = .F1
        oAm_To(J%).F2 = .F2
    End With
Next
Exit Function
E: Cpy_Am = True
End Function
Function Cpy_Am_F1F2Swap(oAm() As tMap) As Boolean
Dim J%: For J% = 0 To jj.Siz_Am(oAm) - 1
    With oAm(J%)
        Dim A$: A = .F2: .F2 = .F1: .F1 = A
    End With
Next
End Function
Function Cpy_Am_F1ToF2(oAm() As tMap) As Boolean
Dim J%: For J% = 0 To jj.Siz_Am(oAm) - 1
    With oAm(J%)
        .F2 = .F1
    End With
Next
End Function
Function Cpy_Am_F2ToF1(oAm() As tMap) As Boolean
Dim J%: For J% = 0 To jj.Siz_Am(oAm) - 1
    With oAm(J%)
        .F1 = .F2
    End With
Next
End Function
Function Cpy_Am_Tst() As Boolean
Dim mAm_Fm(0 To 4) As tMap
Dim mAm_To() As tMap
Dim J%: For J% = 0 To 4
    With mAm_Fm(J%)
        .F1 = J%
        .F2 = J% * 10
    End With
Next
Debug.Print jj.Cpy_Am(mAm_To, mAm_Fm)
Debug.Print jj.ToStr_Am(mAm_To, , , , vbLf)
jj.Shw_DbgWin
End Function
Function Cpy_AmF1_ToAy(oAy$(), pAm() As tMap) As Boolean
Dim N%: N = jj.Siz_Am(pAm)
ReDim oAy(0 To N - 1): If N = 0 Then Exit Function
Dim J%: For J = 0 To N - 1
    oAy(J) = pAm(J).F1
Next
End Function
Function Cpy_AmF2_ToAy(oAy$(), pAm() As tMap) As Boolean
Dim N%: N = jj.Siz_Am(pAm)
ReDim oAy(0 To N - 1): If N = 0 Then Exit Function
Dim J%: For J = 0 To N - 1
    oAy(J) = pAm(J).F2
Next
End Function
Function Cpy_Fil(pFfnFm$, pFfnTo$, Optional pOvrWrt As Boolean = False) As Boolean
Const cSub$ = "Cpy_Fil"
If jj.Ovr_Wrt(pFfnTo, pOvrWrt) Then ss.A 1: GoTo E
On Error GoTo R
jj.g.gFso.CopyFile pFfnFm, pFfnTo
'VBA.FileCopy pFfnFm, pFfnTo
Exit Function
R: ss.R
E: Cpy_Fil = True: ss.B cSub, cMod, "From File,To File", pFfnFm, pFfnTo
End Function
Function Cpy_Fil_Up1Dir(pDir$, Optional pFspc$ = "*.*", Optional pToPfx$ = "") As Boolean
'Aim: Copy all files of {pFspc} in {pDir} up 1 directory with a prefix {pToPfx}.  Target file will be overwritten.
Const cSub$ = "Cpy_Fil_Up1Dir"
Dim mChk As Boolean
mChk = False
''==Start
If Not jj.IsDir(pDir) Then ss.A 1: GoTo E

'Copy Files in {pDir} up 1 directory
On Error GoTo R
Dim iFn$: iFn = VBA.Dir(pDir & pFspc)
Dim mA$
While iFn <> ""
    If jj.Cpy_Fil(pDir & iFn, pDir & "..\" & pToPfx & iFn) Then mA = jj.Add_Str(mA, iFn)
    iFn = VBA.Dir
Wend
If Len(mA) <> 0 Then ss.A 1, "Some files cannot be copied", eRunTimErr, "The Files", mA: GoTo E
If mChk Then
    MsgBox jj.Fmt_Str("Check if all the of spec [{0}] the dir [{1}] is copied up 1 dir with pfx[{2}]", pFspc, pDir, pToPfx), vbInformation, "jj.CopyFilUp1Dir"
    jj.Opn_Dir pDir
    Stop
End If
Exit Function
R: ss.R
E: Cpy_Fil_Up1Dir = True: ss.B cSub, cMod, "pDir,pFSpc"
End Function
Function Cpy_Obj_ByLn(pLnObj_Tar$, pTypObj As AcObjectType, Optional pFb_Src$ = "", Optional pLnObj_Src$ = "") As Boolean
Const cSub$ = "Cpy_Obj"
Dim mAccess As Access.Application
If pFb_Src <> "" Then
    If Not jj.IsFfn(pFb_Src) Then ss.A 1: GoTo E
    Set mAccess = g.gAcs
    If jj.Opn_CurDb(mAccess, pFb_Src) Then jj.Cls_CurDb mAccess: ss.A 1: GoTo E
End If

On Error GoTo R
Dim mAnObj_Tar$(): mAnObj_Tar = Split(pLnObj_Tar, cComma)
Dim mAnObj_Src$(): mAnObj_Src = Split(Fct.NonBlank(pLnObj_Src, pLnObj_Tar), cComma)
Dim N%: N = jj.Siz_Ay(mAnObj_Src)
If jj.Siz_Ay(mAnObj_Tar) <> N Then ss.A 1, "# of object names in Src & Tar are diff", , "Src,Tar", N, jj.Siz_Ay(mAnObj_Tar): GoTo E
Dim J%
If pFb_Src = "" Then
    For J = 0 To N - 1
        DoCmd.CopyObject , mAnObj_Tar(J), pTypObj, mAnObj_Src(J)
    Next
Else
    For J = 0 To N - 1
        mAccess.DoCmd.CopyObject CurrentDb.Name, mAnObj_Tar(J), pTypObj, mAnObj_Src(J)
    Next
    jj.Cls_CurDb mAccess
End If
Select Case pTypObj
Case Access.AcObjectType.acQuery: CurrentDb.QueryDefs.Refresh
Case Access.AcObjectType.acTable: CurrentDb.TableDefs.Refresh
End Select
Exit Function
R: ss.R
E: Cpy_Obj_ByLn = True: ss.B cSub, cMod, "J (Idx of LnObj with err),pLnObj_Tar,pTypObj,pFb_Src,pLnObj_Src", J, pLnObj_Tar, pTypObj, pFb_Src, pLnObj_Src
X: If pFb_Src <> "" Then jj.Cls_CurDb mAccess
End Function
Function Cpy_Obj_ByLn_Tst() As Boolean
Const cSub$ = "Cpy_Obj_ByPfx_Tst"
Dim mLnObj_Src$, mFb_Src$, mTypObj As Access.AcObjectType
Dim mResult As Boolean, mCase As Byte: mCase = 1
Select Case mCase
Case 1
    mLnObj_Src = "qryOdbcMPS_01_0_Prm,qryOdbcMPS_01_1_Fm_qEnv_qBrand"
    mFb_Src = "P:\MPSDetail\MPSDetail\MPS.Mdb"
    mTypObj = acQuery
End Select
mResult = jj.Cpy_Obj_ByLn(mLnObj_Src, mTypObj, mFb_Src)
jj.Shw_Dbg cSub, cMod, , "Result,mLnObj_Src,mTypObj,mFb_Src", mResult, mLnObj_Src, jj.ToStr_TypObj(mTypObj), mFb_Src
End Function
Function Cpy_Obj_ByPfx(pPfx_Tar$, pTypObj As AcObjectType, Optional pFb_Src$ = "", Optional pPfx_Src$ = "") As Boolean
Const cSub$ = "Cpy_Obj_ByPfx"

If pFb_Src = "" And pPfx_Src = "" Then ss.A 1, "Cannot both pFb_Src & pPfx_Src be blank", , "pPfx_Tar,pTypObj", pPfx_Tar, jj.ToStr_TypObj(pTypObj): GoTo E

Dim mAyTar$(): If jj.Fnd_AnObj_ByPfx_InMdb(mAyTar$, pFb_Src, pPfx_Tar, pTypObj) Then ss.A 2: GoTo E
Dim mAySrc$(): If jj.Repl_Pfx_InAy(mAySrc, pPfx_Src, mAyTar, pPfx_Tar) Then ss.A 2: GoTo E
Dim N%: N = jj.Siz_Ay(mAySrc)
If jj.Siz_Ay(mAyTar) <> N Then ss.A 1, "# of object names in Src & Tar are diff", , "Src,Tar", N, jj.Siz_Ay(mAyTar): GoTo E

Dim mAccess As Access.Application
If pFb_Src <> "" Then
    If Not jj.IsFfn(pFb_Src) Then ss.A 1: GoTo E
    Set mAccess = New Access.Application
    If jj.Opn_CurDb(mAccess, pFb_Src) Then jj.Cls_CurDb mAccess:  ss.A 1: GoTo E
End If
On Error GoTo R

Dim J%
If pFb_Src = "" Then
    For J = 0 To N - 1
        DoCmd.CopyObject , mAyTar(J), pTypObj, mAySrc(J)
    Next
Else
    For J = 0 To N - 1
        mAccess.DoCmd.CopyObject CurrentDb.Name, mAyTar(J), pTypObj, mAySrc(J)
    Next
    jj.Cls_CurDb mAccess
    mAccess.Quit
    Set mAccess = Nothing
End If
Select Case pTypObj
Case Access.AcObjectType.acQuery: CurrentDb.QueryDefs.Refresh
Case Access.AcObjectType.acTable: CurrentDb.TableDefs.Refresh
End Select
Exit Function
R: ss.R
E: Cpy_Obj_ByPfx = True: ss.B cSub, cMod, "J (Idx of Pfx with err),pPfx_Tar,pTypObj,pFb_Src,pPfx_Src", J, pPfx_Tar, pTypObj, pFb_Src, pPfx_Src
X: If pFb_Src <> "" Then jj.Cls_CurDb mAccess: mAccess.Quit: Set mAccess = Nothing
End Function
Function Cpy_Obj_ByPfx_Tst() As Boolean
Const cSub$ = "Cpy_Obj_ByPfx_Tst"
Dim mPfx_Src$, mFb_Src$, mTypObj As Access.AcObjectType
Dim mResult As Boolean, mCase As Byte: mCase = 1
Select Case mCase
Case 1
    mPfx_Src = "qryOdbcFc_0"
    mFb_Src = "P:\MPSDetail\MPSDetail\WorkingDir\PgmObj\RfhFc.Mdb"
    mTypObj = acQuery
End Select
mResult = jj.Cpy_Obj_ByPfx(mPfx_Src, mTypObj, mFb_Src)
jj.Shw_Dbg cSub, cMod, , "Result,mPfx_Src,mTypObj,mFb_Src", mResult, mPfx_Src, jj.ToStr_TypObj(mTypObj), mFb_Src
End Function
Function Cpy_Qry(pFmPfx$, pToPfx$) As Boolean
Const cSub$ = "Cpy_Qry"
If jj.Dlt_Qry_ByPfx(pToPfx) Then ss.A 1: GoTo E
Dim iQry As QueryDef
Dim mLstPart$
For Each iQry In CurrentDb.QueryDefs
    If Left(iQry.Name, Len(pFmPfx)) = pFmPfx Then
        mLstPart = mID$(iQry.Name, Len(pFmPfx) + 1)
        Debug.Print "Fm:" & iQry.Name & "  To:" & pToPfx & mLstPart
        Call DoCmd.CopyObject(, pToPfx & mLstPart, AcObjectType.acQuery, pFmPfx & mLstPart)
    End If
Next
Exit Function
R: ss.R
E: Cpy_Qry = True: ss.B cSub, cMod, "pFmPfx,pToPfx", pFmPfx, pToPfx
End Function
Function Cpy_Rs_ToFrm(pRs As DAO.Recordset, pFrm As Access.Form, pLnFld$) As Boolean
'Aim: Copy the fields value from {pRs} to the controls in {pFrm}.  Only those fields in {pLnFld} will be copied.
'     {pLnFld} is in fmt of aaa=xxx,bbb,ccc  aaa,bbb,ccc will be field name in {pFrm} & xxx,bbb,ccc will be field in {pRs}
Const cSub$ = "Cpy_Rs_ToFrm"
On Error GoTo R
Dim mAn_Frm$(), mAn_Rs$(): If jj.Brk_Lm_To2Ay(mAn_Frm, mAn_Rs, pLnFld) Then ss.A 1: GoTo E
Dim mIsEq As Boolean, mEr$, mV_Rs, mV_FrmNew
Dim J%: For J = 0 To jj.Siz_Ay(mAn_Frm) - 1
    With pFrm.Controls(mAn_Frm(J))
        mV_Rs = pRs.Fields(mAn_Rs(J)).Value
        mV_FrmNew = .Value
        If jj.IfEq(mIsEq, mV_Rs, mV_FrmNew) Then ss.A 1: GoTo E
        If Not mIsEq Then .Value = mV_Rs
    End With
Next
'Sav.Rec
Exit Function
R: ss.R
E: Cpy_Rs_ToFrm = True: ss.B cSub, cMod, "pRs,pFrm,pLnFld", jj.ToStr_Rs(pRs), jj.ToStr_Frm(pFrm), pLnFld
End Function
Function Cpy_XlsRowDown(pWs As Worksheet, pColRgeList$, pRow&, pNRow&, Optional pCopyOnly_Val_Fmt As Boolean = True) As Boolean
Dim mAyColRge$(), J As Byte, mFmAdr$, mToAdr$
mAyColRge = Split(pColRgeList, cComma)
With pWs
    For J = LBound(mAyColRge) To UBound(mAyColRge)
        mFmAdr = jj.Cv_RnoColRge2Adr(mAyColRge(J), pRow)
        mToAdr = jj.Cv_RnoColRge2Adr(mAyColRge(J), True)
        mToAdr = mToAdr & pRow + 1 & ":" & mToAdr & pRow + pNRow - 1
        .Range(mFmAdr).Copy
        If pCopyOnly_Val_Fmt Then
            .Range(mToAdr).PasteSpecial xlPasteFormulas
            .Range(mToAdr).PasteSpecial xlPasteFormats
        Else
            .Range(mToAdr).PasteSpecial xlPasteAll
        End If
    Next
End With
End Function
