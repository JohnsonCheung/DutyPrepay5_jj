Attribute VB_Name = "xIs"
#Const Tst = True
Option Compare Text
Option Explicit
Option Base 0
Const cMod$ = cLib & ".xIs"
Function IsEnd(pS$, pEnd$) As Boolean
IsEnd = (Right(pS, Len(pEnd)) = pEnd)
End Function
Function IfEqRs(oIsEq As Boolean, pRs1 As DAO.Recordset, pRs2 As DAO.Recordset) As Boolean
'Aim: Compare 2 rs if they are they are the same
Const cSub$ = "IfEqRs"
On Error GoTo R
oIsEq = False
If pRs1.Fields.Count <> pRs2.Fields.Count Then ss.A 1, "Diff in Fld Cnt": GoTo E
Dim J%
For J = 0 To pRs1.Fields.Count - 1
    If pRs1.Fields(J).Name <> pRs2.Fields(J).Name Then ss.A 2, "Diff field name in J-Th field", , "J", J: GoTo E
    If pRs1.Fields(J).Value <> pRs2.Fields(J).Value Then Exit Function
Nxt:
Next
oIsEq = True
Exit Function
R: ss.R
E: IfEqRs = True: ss.B cSub, cMod, "pRs1,pRs2", jj.ToStr_Rs_NmFld(pRs1, True), jj.ToStr_Rs_NmFld(pRs2, True)
End Function
#If Tst Then
Function IfEqRs_Tst() As Boolean
Const cSub$ = "IfEqRs_Tst"
Dim mRs1 As DAO.Recordset, mRs2 As DAO.Recordset
Set mRs1 = CurrentDb.OpenRecordset("Select * from mstBrand")
Set mRs2 = CurrentDb.OpenRecordset("Select * from mstBrand")
Dim mIsEq As Boolean: If IfEqRs(mIsEq, mRs1, mRs2) Then Stop
jj.Shw_Dbg cSub, cMod, "mIsEq", mIsEq
End Function
#End If
Function IsStrExistInQry(pStr$) As Boolean
Const cSub$ = "IsStrExistInQry"
On Error GoTo R
Dim iQry As QueryDef
For Each iQry In CurrentDb.QueryDefs
    If Left(iQry.Name, 1) <> "~" Then
        If InStr(iQry.Sql, pStr) > 0 Then IsStrExistInQry = True: Exit Function
    End If
Next
Exit Function
R: ss.R
E: ss.B cSub, cMod, "pStr"
End Function
Function IsFfn(pFfn$, Optional pSilient As Boolean) As Boolean
Const cSub$ = "IsFfn"
If VBA.Dir(pFfn) = "" Then ss.A 1, "File does not exists": GoTo E
IsFfn = True
Exit Function
E: If Not pSilient Then ss.B cSub, cMod, "pFfn", pFfn
End Function
Function IsDir(pDir$, Optional pSilient As Boolean = False) As Boolean
Const cSub$ = "IsDir"
If Right(pDir, 1) <> "\" Then ss.A 1, "Given pDir must end with \": GoTo E
If VBA.Dir(pDir, vbDirectory) = "" Then ss.A 2, "Given pDir not exist": GoTo E
IsDir = True
Exit Function
E: If Not pSilient Then ss.B cSub, cMod, "pDir", pDir
End Function
Function IfMem255(oIsMem255 As Boolean, pNmtq$, Optional pFbSrc$ = "") As Boolean
'Aim: Is pNmt has a memo field of len greater then 255
Const cSub$ = "IfMem255"
oIsMem255 = False
Dim mNmtq$: mNmtq = jj.Q_SqBkt(pNmtq)
Dim mInFb$: mInFb = jj.Cv_Fb2InFb(pFbSrc)
Dim mSql$: mSql = jj.Fmt_Str("select * from {0}{1}", mNmtq, mInFb)
Dim mRs As DAO.Recordset: If jj.Opn_Rs(mRs, mSql) Then ss.A 1: GoTo E
Dim J%, mAnFld$()
For J = 0 To mRs.Fields.Count - 1
    If mRs.Fields(J).Type = dbMemo Then If jj.Add_AyEle(mAnFld, mRs.Fields(J).Name) Then ss.A 2: GoTo E
Next
Dim N%: N = jj.Siz_Ay(mAnFld): If N = 0 Then Exit Function
Dim mCndn$: mCndn = jj.Fmt_Str_Repeat_Ay("Len({N})>0", mAnFld, , " or ")
mSql = jj.Fmt_Str("Select count(*) from {0}{1} where {2}", mNmtq, mInFb, mCndn)
Dim mRecCnt&: If jj.Fnd_ValFmSql(mRecCnt, mSql) Then ss.A 3: GoTo E
If mRecCnt >= 1 Then oIsMem255 = True
GoTo X
E: IfMem255 = True: ss.B cSub, cMod, "pNmtq,pFbSrc", pNmtq, pFbSrc
X: jj.Cls_Rs mRs
End Function
#If Tst Then
Function IsMem255_Tst() As Boolean
Dim mIsMem255 As Boolean, mA$
mA = "@OldQry": If jj.IfMem255(mIsMem255, mA) Then Stop Else Debug.Print mA$; "-->"; mIsMem255
mA = "@OldQsT":   If jj.IfMem255(mIsMem255, mA) Then Stop Else Debug.Print mA$; "-->"; mIsMem255
End Function
#End If
Function IsLik_BySet(pNm$, pSetNm$) As Boolean
End Function
Function IsIdx(pNmt$, pNmIdx$, Optional pDb As DAO.Database, Optional pSilient As Boolean = False) As Boolean
Const cSub$ = "IsIdx"
On Error GoTo R
Dim mDb As DAO.Database: Set mDb = jj.Cv_Db(pDb)
Dim mA$: mA = mDb.TableDefs(pNmt).Indexes(pNmIdx).Name
IsIdx = True
Exit Function
R: ss.R
E: If Not pSilient Then ss.B cSub, cMod, "pNmt,pNmIdx,pDb", pNmt, pNmIdx, jj.ToStr_Db(pDb)
End Function
Function IsLikAyLik(pS$, pAyLik$()) As Boolean
'Aim: If pS is like any one of the element of pAyLik$()
Const cSub$ = "IsLikAyLik"
Dim J%
For J = 0 To jj.Siz_Ay(pAyLik) - 1
    If pS Like pAyLik(J) Then IsLikAyLik = True: Exit Function
Next
End Function
#If Tst Then
Function IsLikAyLik_Tst() As Boolean
Dim mA$(1): mA(0) = "bb*"
mA(1) = "cc*": Debug.Print jj.IsLikAyLik("aa", mA)
mA(1) = "a*":  Debug.Print jj.IsLikAyLik("aa", mA)
End Function
#End If
Function IsSubSet_Ln(pSet_Ln$, pSubSet_Ln$) As Boolean
'Aim: Test if pSubSet_Ln is a subset of pSet_Ln
Const cSub$ = "IsSubSet_Ln"
Dim mAn$(): mAn = Split(pSet_Ln, cComma)
Dim N%: N = jj.Siz_Ay(mAn): If N = 0 Then IsSubSet_Ln = False: Exit Function

Dim mAn_SubSet$(): mAn_SubSet = Split(pSubSet_Ln, cComma)
Dim NSubSet%: NSubSet = jj.Siz_Ay(mAn_SubSet): If NSubSet = 0 Then IsSubSet_Ln = True: Exit Function

Dim iSubSet%, J%
For iSubSet = 0 To NSubSet - 1
    Dim mFound As Boolean: mFound = False
    For J = 0 To N - 1
        If Trim(mAn_SubSet(iSubSet)) = Trim(mAn(J)) Then mFound = True: Exit For
    Next
    If Not mFound Then Exit Function
Next
IsSubSet_Ln = True
End Function
#If Tst Then
Function IsSubSet_Ln_Tst() As Boolean
Debug.Print jj.IsSubSet_Ln("aa,bb , cc", "aa,cc")
Debug.Print jj.IsSubSet_Ln("aa,bb , cc", "aa")
End Function
#End If
Function IfFfnSamDte(oIsFfnSamDte As Boolean, pFfn$) As Boolean
Const cSub$ = "IfFfnSamDte"
If VBA.Dir(pFfn) = "" Then ss.A 1, "File not exist": GoTo E
oIsFfnSamDte = (Date = CDate(Format(VBA.FileDateTime(pFfn), "yyyy/mm/dd")))
Exit Function
E: IfFfnSamDte = True: ss.B cSub, cMod, "pFfn", pFfn
End Function
Function IsAcs() As Boolean
Static xIsSet As Boolean, xIsAcs As Boolean
If xIsSet Then IsAcs = xIsAcs: Exit Function
xIsAcs = Application.Name = "Microsoft Access"
xIsSet = True
IsAcs = xIsAcs
End Function
Function IfEq_Nm2V(oIsEq As Boolean, pNm2V As tNm2V) As Boolean
Const cSub$ = "IfEq_Nm2V"
On Error GoTo R
With pNm2V
    IfEq_Nm2V = jj.IfEq(oIsEq, .NewV, .OldV)
End With
Exit Function
R: ss.R
E: IfEq_Nm2V = True: ss.B cSub, cMod, "pNm2V", jj.ToStr_Nm2V(pNm2V)
End Function
Function IfEq(oIsEq As Boolean, pV1, pV2) As Boolean
Const cSub$ = "IfEq"
On Error GoTo R
If VarType(pV1) = vbString Then
    If VarType(pV2) = vbString Then oIsEq = (RTrim(pV1) = RTrim(pV2)): Exit Function
End If
If VarType(pV1) = vbNull Then
    If VarType(pV2) = vbNull Then oIsEq = True: Exit Function
    If VarType(pV2) = vbString Then oIsEq = (Trim(pV2) = ""): Exit Function
    oIsEq = False
    Exit Function
End If
If VarType(pV2) = vbNull Then
    If VarType(pV1) = vbString Then oIsEq = (Trim(pV1) = ""): Exit Function
    oIsEq = False
    Exit Function
End If
oIsEq = (pV1 = pV2)
Exit Function
R: ss.R
E: IfEq = True: ss.B cSub, cMod, "pV1,pV2", pV1, pV2
End Function
Function IsEq_Tst() As Boolean
Const cSub$ = "IsIsEq_Tst"
Dim mV1, mV2, mIsEq As Boolean
Dim mRslt As Boolean, mCase As Byte: mCase = 1
jj.Shw_Dbg cSub, cMod
For mCase = 1 To 8
    Select Case mCase
    Case 1: mV1 = Null:    mV2 = Null
    Case 2: mV1 = Null:    mV2 = "1"
    Case 3: mV1 = "1":     mV2 = Null
    Case 4: mV1 = Null:    mV2 = 1
    Case 5: mV1 = 1:       mV2 = Null
    Case 6: mV1 = "1 ":    mV2 = "1"
    Case 7: mV1 = "1":     mV2 = "1 "
    Case 8: mV1 = 1:       mV2 = "1"
    End Select
    mRslt = jj.IfEq(mIsEq, mV1, mV2)
    Debug.Print jj.ToStr_LpAp(vbTab & vbTab, "mRslt,mV1,mV2,mIsEq", mRslt, mV1 & "(" & TypeName(mV1) & ")", mV2 & "(" & TypeName(mV2) & ")", mIsEq)
Next
End Function
Function IsFrmOpn(pFrmNam$) As Boolean
On Error Resume Next
Dim mNam$: mNam = Forms(pFrmNam).Name
IsFrmOpn = mNam <> ""
End Function
Function IsMacro(pS$) As Boolean
Dim mP1%: mP1 = InStr(pS, "{")
Dim mP2%: mP2 = InStr(pS, "}")
IsMacro = (mP2 > mP1 And mP1 > 0)
End Function
Function IsNmtBad(pNmt$, Optional pDb As DAO.Database = Nothing) As Boolean
'Aim: Return error if Nmt is Bad, means not Exist
Const cSub$ = "IsIsNmtBad"
If jj.IsTbl(pNmt, pDb) Then ss.A 1, "Given table exist": GoTo E
IsNmtBad = True
Exit Function
E: ss.B 1, cSub, cMod, "pNmt,pDb", pNmt, jj.ToStr_Db(pDb)
End Function
Function IsNmtOK(pNmt$, Optional pDb As DAO.Database = Nothing) As Boolean
'Aim: Return error if Nmt is OK, means Exist
Const cSub$ = "IsIsNmtOK"
If Not jj.IsTbl(pNmt) Then ss.A 1, "Given table not exist": GoTo E
IsNmtOK = True
E:  ss.B 1, cSub, cMod, "pNmt,pDb", pNmt, jj.ToStr_Db(pDb)
End Function
Function IsNoRecInFrm(pFrm As Form) As Boolean
Const cSub$ = "IsIsNoRecInFrm"
On Error GoTo R
Dim mRs As DAO.Recordset: Set mRs = pFrm.Recordset
IsNoRecInFrm = mRs.RecordCount = 0
Exit Function
R: ss.R
E: ss.B cSub, cMod, "pFrm", jj.ToStr_Frm(pFrm)
End Function
Function IsNothing(p) As Boolean
IsNothing = (TypeName(p) = "Nothing")
End Function
Function IsPfx(pS$, pPfx$) As Boolean
If pPfx = "" Then IsPfx = True: Exit Function
If Right(pPfx, 1) = "*" Then
    Dim L As Byte: L = Len(pPfx) - 1
    IsPfx = (Left(pS, L) = Left(pPfx, L))
Else
    IsPfx = (pPfx = pS)
End If
End Function
Function IsQry(pNmq$, Optional pDb As DAO.Database = Nothing) As Boolean
Dim mDb As DAO.Database: Set mDb = jj.Cv_Db(pDb)
On Error GoTo R
Dim mA$: mA = mDb.QueryDefs(pNmq).Name
If Left(mA, 4) = "~sq_" Then Exit Function
On Error GoTo 0
IsQry = True
Exit Function
R: ss.R
End Function
Function IsRel(pNmRel$, Optional pDb As Database = Nothing) As Boolean
On Error GoTo R
Dim mA$: mA = Cv_Db(pDb).Relations(pNmRel).Name
IsRel = True
Exit Function
R:
End Function
Function IsSamAy(pAy1$(), pAy2$()) As Boolean
Dim N%: N = jj.Siz_Ay(pAy1)
If jj.Siz_Ay(pAy2) = N Then Exit Function
Dim J%: For J = 0 To N - 1
    If Fct.InStr_Ay(pAy1(J), pAy2) = -1 Then Exit Function
Next
IsSamAy = True
End Function
Function IsSamKey(pRs As DAO.Recordset, pAyKvLas()) As Boolean
'Aim: Return True if first N fields of {pRs} is same as pAyKvLas
Dim N%: N = jj.Siz_Ay(pAyKvLas)
Dim J As Byte: For J = 0 To N - 1
    If pRs.Fields(J).Value <> pAyKvLas(J) Then Exit Function
Next
IsSamKey = True
End Function
Function IsSamKey_ByAnFldKey(pRs As DAO.Recordset, pAnFldKey$(), pAyKvLas()) As Boolean
'Aim: Return if the list of fields of name in {pKeyFlds} of current record of {pRs} has the same values as in {pLastKey}
Dim N%: N = jj.Siz_Ay(pAnFldKey)
Dim J As Byte: For J = 0 To N - 1
    If pRs.Fields(pAnFldKey(J)).Value <> pAyKvLas(J) Then Exit Function
Next
IsSamKey_ByAnFldKey = True
End Function
Function IsSingleWsXls(pFx$, Optional oWb As Workbook, Optional oWs As Worksheet) As Boolean
Const cSub$ = "IsSingleWsXls"
On Error GoTo R
If jj.Opn_Wb_R(oWb, pFx) Then ss.A 1: GoTo E
Set oWb = g.gXls.Workbooks.Open(pFx)
If oWb.Worksheets.Count <> 1 Then ss.A 2, "Given pFx has more than 1 worksheet", , "N Ws", oWb.Worksheets.Count: GoTo E
Set oWs = oWb.Worksheets(1)
If oWs.Name <> Fct.Nam_FilNam(oWb.Name, False) Then ss.A 3, "The Ws of given pFx should have same name as the file", , "Ws Name", oWs.Name: GoTo E
IsSingleWsXls = True
Exit Function
R: ss.R
E: ss.B cSub, cMod, "pFx", pFx
End Function
Function IsCmt(pRge As Range) As Boolean
Const cSub$ = "IsIsCmt"
On Error GoTo R
IsCmt = TypeName(pRge.Comment) <> "Nothing"
Exit Function
R: ss.R
E: ss.B cSub, cMod, "pRge", jj.ToStr_Rge(pRge)
End Function
Function IsTbl(pNmt, Optional pDb As DAO.Database = Nothing) As Boolean
Const cSub$ = "IsTbl"
On Error GoTo R
Dim mA$: mA = jj.Cv_Db(pDb).TableDefs(jj.Rmv_SqBkt(CStr(pNmt))).Name
IsTbl = True
Exit Function
R:
End Function
Function IsLnt(pLnt$, Optional pDb As DAO.Database = Nothing) As Boolean
Const cSub$ = "IsLnt"
On Error GoTo R
Dim mAnt$(): If jj.Brk_Ln2Ay(mAnt, pLnt) Then ss.A 1: GoTo E
Dim J%
For J = 0 To jj.Siz_Ay(mAnt) - 1
    If Not jj.IsTbl(mAnt(J), pDb) Then IsLnt = False: Exit Function
Next
IsLnt = True
Exit Function
R: ss.R
E: ss.B cSub, cMod, "pLnt,pDb", pLnt, jj.ToStr_Db(pDb)
End Function
Function IsNmtq(pNmtq$, Optional pDb As DAO.Database = Nothing) As Boolean
IsNmtq = True
If jj.IsTbl(pNmtq, pDb) Then Exit Function
If jj.IsQry(pNmtq, pDb) Then Exit Function
IsNmtq = False
End Function
Function IfWs_InFx(oIsWs_InFx As Boolean, pFx$, pNmWs$) As Boolean
Const cSub$ = "IfWs_InFx"
Dim mWb As Workbook: If jj.Opn_Wb_R(mWb, pFx) Then ss.A 1: GoTo E
oIsWs_InFx = jj.IsWs(mWb, pNmWs)
GoTo X
E: IfWs_InFx = True: ss.B cSub, cMod, "pFx,pNmWs", pFx, pNmWs
X: jj.Cls_Wb mWb, , True
End Function
Function IsWs(pWb As Workbook, pNmWs$) As Boolean
On Error GoTo R
Dim mWs As Worksheet: Set mWs = pWb.Sheets(pNmWs)
If mWs.Name = pNmWs Then IsWs = True: Exit Function
R:
End Function
#If Tst Then
Function IsWs_Tst() As Boolean
Const cFx$ = "c:\tmp\aa.xls"
Const cNmWs$ = "xxxxx"
Dim mWb As Workbook: If jj.Crt_Wb(mWb, cFx, True, cNmWs) Then Stop
MsgBox jj.IsWs(mWb, cNmWs), , "Is.Ws? (Should be true)"
If jj.Cls_Wb(mWb, True) Then Stop
'INTO [Excel 8.0;HDR=YES;IMEX=2;DATABASE={0}].{1} FROM {2}"
End Function
#End If
Function IsXls() As Boolean
Static xIsSet As Boolean, xIsXls As Boolean
If xIsSet Then IsXls = xIsXls: Exit Function
xIsXls = Application.Name = "Microsoft Excel"
xIsSet = True
IsXls = xIsXls
End Function
Function IsWbNm(pWb As Excel.Workbook, pNm$, Optional oNm As Excel.Name) As Boolean
On Error GoTo R
Set oNm = pWb.Names(pNm)
IsWbNm = True
Exit Function
R:
End Function
Function IsWsNm(pWs As Excel.Workbook, pNm$, Optional oNm As Excel.Name) As Boolean
On Error GoTo R
Set oNm = pWs.Names(pNm)
IsWsNm = True
Exit Function
R:
End Function
