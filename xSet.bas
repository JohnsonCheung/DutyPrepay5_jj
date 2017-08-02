Attribute VB_Name = "xSet"
#Const Tst = True
Option Compare Text
Option Explicit
Option Base 0
Const cMod$ = cLib & ".xSet"
Dim x_AySilent() As Boolean
Dim x_AyNoLog() As Boolean
Function Set_Import_AtA1(pFx$) As Boolean
'Aim: each ws in {pFx} set a1 as "Import:{NmWs}"
Const cSub$ = "Set_Import_AtA1"
On Error GoTo R
Dim mWb As Workbook: If jj.Opn_Wb_RW(mWb, pFx) Then ss.A 1: GoTo E
Dim iWs As Worksheet
For Each iWs In mWb.Sheets
    iWs.Range("A1").Value = "Import:" & iWs.Name
Next
If jj.Cls_Wb(mWb, True) Then ss.A 2: GoTo E
Exit Function
R: ss.R
E: Set_Import_AtA1 = True: ss.B cSub, cMod, "pFx", pFx
End Function
#If Tst Then
Function Set_Import_AtA1_Tst() As Boolean
If jj.Set_Import_AtA1("P:\AppDef_Meta\MetaDb.xls") Then Stop
End Function
#End If
Function Set_Lm_FmSql(oLm$, pSql$ _
    , Optional pNmFld0$ = "" _
    , Optional pNmFld1$ = "" _
    , Optional pBrkChr$ = "=" _
    , Optional pSepChr$ = vbCrLf) As Boolean
'Aim: Build {oLm} from 2 fields ({pNmFld1} & {pNmFld2}) of {pRs}
Const cSub$ = "Set_Lm_FmSql"
Dim mRs As DAO.Recordset: If jj.Opn_Rs(mRs, pSql) Then ss.A 1: GoTo E
If jj.Set_Lm_ByRs(oLm, mRs, pNmFld0, pNmFld1, pBrkChr, pSepChr) Then ss.A 2: GoTo E
GoTo X
R: ss.R
E: Set_Lm_FmSql = True: ss.B cSub, cMod, "pSql,pNmFld0,pNmFld1,pBrkChr,pSepChr", pSql, pNmFld0, pNmFld1, pBrkChr, pSepChr
X: jj.Cls_Rs mRs
End Function
Function Set_Lv_ByRs(oLv$, pRs As DAO.Recordset, pLnFld$, _
    Optional pBrk$ = "=", Optional pSep$ = cComma, Optional pIsNoNm As Boolean = False) As Boolean
'Aim: Build {oLv} by {pLnFld} in {pRs}
Const cSub$ = "Set_Lv_ByRs"
Dim mAnFld_Lcl$(), mAnFld_Host$(): If jj.Brk_Lm_To2Ay(mAnFld_Lcl, mAnFld_Host, pLnFld) Then ss.A 1: GoTo E
Dim N%: N = jj.Siz_Ay(mAnFld_Lcl)
On Error GoTo R
oLv = ""
With pRs
    Dim J%, mA$
    If pIsNoNm Then
        For J = 0 To N - 1
            oLv = jj.Add_Str(oLv, jj.Q_V(.Fields(mAnFld_Lcl(J)).Value), pSep$)
        Next
    Else
        For J = 0 To N - 1
            If jj.Join_NmV(mA, mAnFld_Host(J), .Fields(mAnFld_Lcl(J)).Value, pBrk) Then ss.A 1: GoTo E
            oLv = jj.Add_Str(oLv, mA, pSep$)
        Next
    End If
End With
Exit Function
R: ss.R
E: Set_Lv_ByRs = True: ss.B cSub, cMod, "pFrm,pLnFld", jj.ToStr_Rs(pRs), pLnFld
End Function
#If Tst Then
Function Set_Lv_ByRs_Tst() As Boolean
End Function
#End If
Function Set_Lm_ByTbl(oLm$, pNmt$ _
    , Optional pBrkChr$ = "=" _
    , Optional pSepChr$ = cSemi) As Boolean
'Aim: Build {oLm} from all fields in one record of {pRs}
Const cSub$ = "Set_Lm_ByRsRec"
Dim J%
oLm = ""
With CurrentDb.TableDefs(pNmt).OpenRecordset
    For J = 0 To .Fields.Count - 1
        oLm = jj.Add_Str(oLm, .Fields(J).Name & pBrkChr & .Fields(J).Value, pSepChr)
    Next
    .Close
End With
Exit Function
R: ss.R
E: ss.B cSub, cMod, "pNmt,pNmt->Fields,pBrkChr,pSepChr", pNmt, jj.ToStr_Nmt(pNmt), pBrkChr, pSepChr
End Function
Function Set_Lm_ByRsRec_Tst()
If jj.Crt_Tbl_FmLoFld("#Tmp", "Itm Text 10,N Text 50,X Text 50") Then Stop
If jj.Run_Sql("Insert into [#Tmp] values ('Tbl','1,2,3',',x,xx,xxx')") Then Stop
Dim mLm$: If jj.Set_Lm_ByTbl(mLm, "#Tmp") Then Stop
Debug.Print mLm
End Function
Function Set_Lm_ByRs(oLm$, pRs As DAO.Recordset _
    , Optional pNmFld0$ = "" _
    , Optional pNmFld1$ = "" _
    , Optional pBrkChr$ = "=" _
    , Optional pSepChr$ = vbCrLf) As Boolean
'Aim: Build {oLm} from all records in {pRs} which have 2 fields {pNmFld1} & {pNmFld2}
Const cSub$ = "Set_Lm_ByRs"
On Error GoTo R
Dim mNmFld0$: mNmFld0 = NonBlank(pNmFld0, pRs.Fields(0).Name)
Dim mNmFld1$: mNmFld1 = NonBlank(pNmFld1, pRs.Fields(1).Name)
oLm = ""
With pRs
    While Not .EOF
        oLm = jj.Add_Str(oLm, .Fields(mNmFld0).Value & pBrkChr & .Fields(mNmFld1).Value, pSepChr)
        .MoveNext
    Wend
End With
Exit Function
R: ss.R
E: Set_Lm_ByRs = True: ss.B cSub, cMod, "pRs,pNmFld0,pNmFld1,pBrkChr,pSepChr", jj.ToStr_Flds(pRs.Fields), pNmFld0, pNmFld1, pBrkChr, pSepChr
End Function
#If Tst Then
Function Set_Am_ByLm_Tst() As Boolean
Const cSub$ = "Set_Am_ByLm_Tst"
Dim mLm$, mSepChr$, mBrkChr$, mCase As Byte
Dim mAm() As Typ.tMap
jj.Shw_Dbg cSub, cMod
For mCase = 1 To 2
    Select Case mCase
    Case 1
        mLm = "aaa=xxx,bbb=yyy,1111"
        mBrkChr = "="
        mSepChr = cComma
    Case 2
        mLm = "aaa as ccc, dd, ee as xx"
        mBrkChr = " as "
        mSepChr = cComma
    End Select
    mAm = jj.Get_Am_ByLm(mLm, mBrkChr, mSepChr)
    Debug.Print "Input----"
    Debug.Print "mLm="; mLm
    Debug.Print "Output----"
    Debug.Print "mAm="; jj.ToStr_Am(mAm)
Next
End Function
#End If
Sub Set_WbAll_Min()
Dim iWb As Workbook, iWin As Window
For Each iWb In Excel.Application.Workbooks
    Set_Wb_Min iWb
Next
End Sub
Sub Set_Wb_Min(pWb As Workbook)
Dim iWin As Window
For Each iWin In pWb.Windows
    If iWin.WindowState <> xlMinimized Then iWin.WindowState = xlMinimized
Next
End Sub
Sub Set_TBar_Toggle()
Dim mWs As Worksheet: Set mWs = Excel.Application.ActiveSheet
If jj.IsNothing(mWs) Then Exit Sub
Dim iOLEObj As Excel.OLEObject
For Each iOLEObj In mWs.OLEObjects
    If TypeName(iOLEObj.Object) = "ToolBar" Then iOLEObj.Visible = Not iOLEObj.Visible
Next
End Sub

Function Set_Ws_ByAyV(pWs As Worksheet, pRno&, pCno As Byte, pIsDown As Boolean, pAyV()) As Boolean
Set_Ws_ByAyV = Set_Ws_ByVayv(pWs, pRno, pCno, pIsDown, CVar(pAyV))
End Function
Function Set_Ws_ByVayv(pWs As Worksheet, pRno&, pCno As Byte, pIsDown As Boolean, pVayv) As Boolean
'Aim: Set {pRno} in {pWs} by {pAp}
Const cSub$ = "Set_Ws_ByVayv"
On Error GoTo R
Dim mAyV(): mAyV = pVayv
Dim J%, N%: N = jj.Siz_Ay(mAyV)
With pWs
    If pIsDown Then
        For J = 0 To N - 1
            .Cells(pRno + J, pCno).Value = mAyV(J)
        Next
        Exit Function
    End If
    For J = 0 To N - 1
        .Cells(pRno, pCno + J).Value = mAyV(J)
    Next
End With
Exit Function
R: ss.R
E: Set_Ws_ByVayv = True: ss.B cSub, cMod, "pWs,pRno,pCno,pIsDown,Vayv,", jj.ToStr_Ws(pWs), pRno, pCno, pIsDown, jj.ToStr_Vayv(pVayv)
End Function
Function Set_Ws_ByAyPrm(pWs As Worksheet, pRno&, pCno As Byte, pIsDown As Boolean, ParamArray pAp()) As Boolean
'Aim: Set {pRno} in {pWs} by {pAp}
Set_Ws_ByAyPrm = Set_Ws_ByVayv(pWs, pRno, pCno, pIsDown, CVar(pAp))
End Function
Public Sub Set_Ws_CmbBox(pWs As Excel.Worksheet, pPfx As String, pCtlCnt As Byte, pPrp As String, pVal)
'Aim: Assume there are pCtlCnt comboxbox control object in the pWs with name XXX01, ... XXXnn, where XXX is pPfx, nn is pCtlCnt
'     It is required to set the property pPrp for each of the control by the value pVal
Dim J As Byte
For J = 1 To 20
    Select Case pPrp
    Case "ListFillRange":    pWs.OLEObjects(pPfx & Format(J, "00")).ListFillRange = pVal
    Case "PrintObject":      pWs.OLEObjects(pPfx & Format(J, "00")).PrintObject = pVal
    Case "Height":           pWs.OLEObjects(pPfx & Format(J, "00")).Height = pVal
    Case "Height":           pWs.OLEObjects(pPfx & Format(J, "00")).Height = pVal
    Case "ListRows":         pWs.OLEObjects(pPfx & Format(J, "00")).ListRows = pVal
    Case Else
    Stop
    End Select
Next
End Sub
Function Set_Pth(pNmt$, pPth$, pSeg$, pNmt_Par_Chd$, Optional pSno$ = "", Optional pRoot$ = "", Optional pSepChr$ = ".") As Boolean
'Aim: Update {pNmt}->{pPth}[,pSno][,pRoot] using {pNmt}->{pSeg} as segment in the path.
'     The par-chd relation is defined in {pNmtParChd} of fmt {mNmtParChd}.{mPar}.{mChd}
'     Assume {mNmtParChd} & {pNmt} has common {mPar}
'     Assume Struct: {pNmt}:mPar,pPth[,pSno][,Root]
'     Assume Struct: {mNmtParChd}:mPar,mChd
Const cSub$ = "Set_Pth"
Dim mNmt0ParChd$, mPar$, mChd$
Do
    If jj.Brk_Str_To3Seg(mNmt0ParChd, mPar, mChd, pNmt_Par_Chd, ".") Then ss.A 1: GoTo E
    If jj.Chk_Struct_Tbl_SubSet(mNmt0ParChd, mPar & "," & mChd) Then ss.A 2: GoTo E
Loop Until True
Dim mNmt1$
Do
    If jj.Chk_Struct_Tbl_SubSet(pNmt, mPar & "," & pPth & "," & pSeg & jj.Cv_Str(pSno, ",") & jj.Cv_Str(pRoot, ",")) Then ss.A 3: GoTo E
    mNmt1 = jj.Q_SqBkt(pNmt)
Loop Until True

Dim mAyRoot&()
If jj.Fnd_AyRoot(mAyRoot, mNmt0ParChd, mPar, mChd) Then ss.A 4: GoTo E
Dim J%, mSno%
For J = 0 To jj.Siz_Ay(mAyRoot) - 1
    mSno = 1
    If Set_Pth_Upd1Nod(mNmt1, mAyRoot(J), pPth, pSeg, mNmt0ParChd, mPar, mChd, pSepChr, pSno, mSno, pRoot, mAyRoot(J)) Then ss.A 5: GoTo E
Next

'Update those nood is not chd but has no chd.
Dim mSql$: mSql = jj.Fmt_Str_ByLpAp_ExclSqBkt("Select t.{mPar}" & _
    " From ({mNmt1} t" & _
    " left join [{mNmt0ParChd}] p on t.{mPar}=p.{mPar})" & _
    " left join [{mNmt0ParChd}] c on t.{mPar}=c.{mChd}" & _
    " Where IsNull(p.{mPar})" & _
    " and   IsNull(c.{mChd})", "mPar,mChd,mNmt1,mNmt0ParChd", mPar, mChd, mNmt1, mNmt0ParChd)
If jj.Fnd_AyVFmSql(mAyRoot, mSql) Then ss.A 4: GoTo E

Dim mFmtStr$: mFmtStr = "Update {mNmt1}" & _
    " Set {pPth}={pSeg} & '{pSepChr}'" & _
    jj.Cv_Str_ByQ(pRoot, ",*={mPar}") & _
    jj.Cv_Str_ByQ(pSno, ",*=1") & _
    " Where {mPar} in (" & ToStr_AyLng(mAyRoot) & ")"
mSql = jj.Fmt_Str_ByLpAp_ExclSqBkt(mFmtStr, "mNmt1,mPar,pPth,pSeg,pSepChr,pRoot,pSno", mNmt1, mPar, pPth, pSeg, pSepChr, pRoot, pSno)
If jj.Run_Sql(mSql) Then ss.A 6: GoTo E
GoTo X
R: ss.R
E: Set_Pth = True: ss.B cSub, cMod, "pNmt,pPth,pSeg,pNmt_Par_Chd,pSepChr", pNmt, pPth, pSeg, pNmt_Par_Chd, pSepChr
X:
End Function
#If Tst Then
Function Set_Pth_Tst() As Boolean
'If jj.Crt_Tbl_FmLnkLnt("p:\workingdir\MetaAll.mdb", "$Dir,$DirC") Then Stop: GoTo E
If jj.Set_Pth("$Dir", "Pth", "NmDirSeg", "$DirC.Dir.DirChd", "SnoDir", "DirRoot", "\") Then Stop
Exit Function
E: Set_Pth_Tst = True
End Function
#End If
Private Function Set_Pth_Upd1Nod(pNmt1$, pNod&, pPth$, pSeg$, pNmtParChd$, pPar$, pChd$, pSepChr$, pSno$, ByRef oSno%, pRoot$, pRootId&, Optional pPthPar$ = "") As Boolean
'Aim: Update pNmt->Pth of record pNod by pSeg & pPthPar.  Recursively update all the child.
'     Note: pNmt1$: 1 means with []; pNmt0$: 0 means without [], pNmt$: means not sure with or without []
Const cSub$ = "Set_Pth_Upd1Nod"
On Error GoTo R
Dim mSql$
mSql = jj.Bld_SqlSel( _
    pPth & "," & pSeg & jj.Cv_Str(pSno, ",") & jj.Cv_Str(pRoot, ",") _
    , pNmt1 _
    , pPar & "=" & pNod)
Dim mRs As DAO.Recordset: If jj.Opn_Rs(mRs, mSql) Then ss.A 1: GoTo E

Dim mPthCur$
With mRs
    .Edit
    If pRoot <> "" Then .Fields(pRoot).Value = pRootId
    mPthCur = pPthPar & .Fields(pSeg).Value & pSepChr
    .Fields(pPth).Value = mPthCur
    If pSno <> "" Then .Fields(pSno).Value = oSno: oSno = oSno + 1
    .Update
    .Close
End With

mSql = jj.Bld_SqlSel( _
    pChd _
    , pNmtParChd _
    , pPar & "=" & pNod)
Dim mAyChd&(): If jj.Fnd_AyVFmSql(mAyChd, mSql) Then ss.A 2: GoTo E
Dim J%
For J = 0 To jj.Siz_Ay(mAyChd) - 1
    If Set_Pth_Upd1Nod(pNmt1, mAyChd(J), pPth, pSeg, pNmtParChd, pPar, pChd, pSepChr, pSno, oSno, pRoot, pRootId, mPthCur) Then ss.A 3: GoTo E
Next
GoTo X
R: ss.R
E: Set_Pth_Upd1Nod = True: ss.B cSub, cMod, "pNmt1,pNod,pPth,pSeg,pNmtParChd,pPar,pChd,pSepChr,oSno,pSno,pRoot,pPthPar", pNmt1, pNod, pPth, pSeg, pNmtParChd, pPar, pChd, pSepChr, oSno, pSno, pRoot, pPthPar
X:
    jj.Cls_Rs mRs
End Function
Function Set_Fld_ToAuto(pNmt$, pNmFld$) As Boolean
Const cSub$ = "Set_Fld_ToAuto"
On Error GoTo R
Dim mFldAtr&: mFldAtr = CurrentDb.TableDefs(pNmt).Fields(pNmFld).Attributes
CurrentDb.TableDefs(pNmt).Fields(pNmFld).Attributes = mFldAtr Or DAO.FieldAttributeEnum.dbAutoIncrField
Exit Function
R: ss.R
E: Set_Fld_ToAuto = True: ss.B cSub, cMod, "pNmt,pNmFld", pNmt, pNmFld
End Function
#If Tst Then
Function Set_Fld_ToAuto_Tst() As Boolean
Dim mNmt$: mNmt = "#Tmp"
If jj.Crt_Tbl_FmLoFld(mNmt, "aa long, bb text 10") Then Stop: GoTo E
If jj.Set_Fld_ToAuto(mNmt, "aa") Then Stop: GoTo E
Exit Function
E: Set_Fld_ToAuto_Tst = True
End Function
#End If
Function Set_NoLog(Optional pNoLog As Boolean = True) As Boolean
Dim N%: N = jj.Siz_Ay(x_AyNoLog)
ReDim Preserve x_AyNoLog(N + 1)
x_AyNoLog(N) = jj.g.gNoLog
jj.g.gNoLog = pNoLog
End Function
Function Set_NoLog_Rst() As Boolean
Dim N%: N = jj.Siz_Ay(x_AyNoLog)
If N = 0 Then jj.g.gNoLog = False: Exit Function
jj.g.gNoLog = x_AyNoLog(N - 1)
ReDim Preserve x_AyNoLog(N - 1)
End Function
Function Set_Silent(Optional pSilent As Boolean = True) As Boolean
Dim N%: N = jj.Siz_Ay(x_AySilent)
ReDim Preserve x_AySilent(N + 1)
x_AySilent(N) = jj.g.gSilent
jj.g.gSilent = pSilent
End Function
Function Set_Silent_Rst() As Boolean
Dim N%: N = jj.Siz_Ay(x_AySilent)
If N = 0 Then jj.g.gSilent = False: Exit Function
jj.g.gSilent = x_AySilent(N - 1)
ReDim Preserve x_AySilent(N - 1)
End Function
Function Set_WsTit_ByRs(pWs As Worksheet, pRs As DAO.Recordset, Optional pRno& = 1) As Boolean
Const cSub$ = "Set_WsTit_ByRs"
On Error GoTo R
Dim J%
With pWs
    For J = 0 To pRs.Fields.Count - 1
        pWs.Cells(pRno, J + 1).Value = pRs.Fields(J).Name
        If pRs.Fields(J).Type = dbDate Then .Columns(J + 1).EntireColumn.NumberFormat = "yyyy/mm/dd h:mm AM/PM;@"
    Next
    .Rows(pRno).AutoFilter
    .Cells.EntireColumn.AutoFit
    .Activate
    .Range("A" & pRno + 1).Activate
    .Range("A" & pRno + 1).Select
    .Application.ActiveWindow.FreezePanes = True
End With
Exit Function
R: ss.R

E:
: ss.B cSub, cMod
    Set_WsTit_ByRs = True
End Function
Function Set_Ver() As Boolean
'Aim: Set tblVer as now.  Create if needed
Const cSub$ = "Set_Ver"
If jj.Crt_TblVer Then ss.A 1: GoTo E
If jj.Run_Sql("Delete * from tblVer") Then ss.A 2: GoTo E
If jj.Run_Sql("insert into tblVer (Ver) values (Now())") Then ss.A 2: GoTo E
MsgBox "tblVer is set"
Exit Function
R: ss.R

E:
: ss.B cSub, cMod, ""
    Set_Ver = True
End Function
Function Set_Ws_ByLv(pWs As Worksheet, pRno&, pCno As Byte, pIsDown As Boolean, pLv$) As Boolean
'Aim: Set 1 row in {pWs} by {pLv}.
Const cSub$ = "Set_Ws_ByLv"
On Error GoTo R
Dim J%, N%, mAyV(): mAyV = jj.Brk_Lv2AyV(pLv, cComma): N = jj.Siz_Ay(mAyV)
Set_Ws_ByLv = jj.Set_Ws_ByVayv(pWs, pRno, pCno, pIsDown, CVar(mAyV))
Exit Function
R: ss.R
E: Set_Ws_ByLv = True: ss.B cSub, cMod, "pWs,pRno,pLv", jj.ToStr_Ws(pWs), pRno, pLv
End Function
Function Set_Ws_ByLpAp(pWs As Worksheet, pRno&, pCno As Byte, pIsDown As Boolean, pLp$, ParamArray pAp()) As Boolean
'Aim: Set first 2 rows in {pWs} {pLnFld} & {pAp}.
Const cSub$ = "Set_Ws_ByLpAp"
On Error GoTo R
Dim J%, N%, mAyV(): mAyV = jj.Brk_Lv2AyV(pLp, cComma)
If Set_Ws_ByAyV(pWs, pRno, pCno, pIsDown, mAyV) Then ss.A 1: GoTo E
If pIsDown Then
    If Set_Ws_ByVayv(pWs, pRno, pCno + 1, pIsDown, CVar(pAp)) Then ss.A 2: GoTo E
Else
    If Set_Ws_ByVayv(pWs, pRno + 1, pCno, pIsDown, CVar(pAp)) Then ss.A 3: GoTo E
End If
Exit Function
R: ss.R
E: Set_Ws_ByLpAp = True: ss.B cSub, cMod, "pWs,pLp,pAp", jj.ToStr_Ws(pWs), pLp, jj.ToStr_Vayv(CVar(pAp))
End Function
#If Tst Then
Function Set_Ws_ByLpAp_Tst() As Boolean
Dim mWb As Workbook, mWs As Worksheet
If jj.Crt_Wb(mWb, "c:\tmp\hh.xls", True) Then Stop
Set mWs = mWb.Sheets(1)
Dim mCase As Byte: mCase = 2
Select Case mCase
Case 1
    If jj.Set_Ws_ByLpAp(mWs, 1, 1, False, "abc,def,xyz", 123, #1/1/2007#, "sdfdf") Then Stop
Case 2
    If jj.Set_Ws_ByLpAp(mWs, 1, 1, True, "abc,def,xyz", 123, #1/1/2007#, "sdfdf") Then Stop
End Select
mWb.Application.Visible = True
Stop
X: jj.Cls_Wb mWb, , True
End Function
#End If
Function Set_Rs_ByLpVv(oRs As DAO.Recordset, pLnFld$, pVayv) As Boolean
'Aim: Set {oRs} by {pLnFld} & {pAyV}.  Assume oRs is already .AddNew or .Edit
Const cSub$ = "Set_Rs_ByLpVv"
On Error GoTo R
Dim J%, mAnFld$(): mAnFld = Split(pLnFld, cComma)
Dim mNmFld$, mAyV()
mAyV = pVayv
With oRs
    For J = 0 To UBound(mAnFld$)
        mNmFld = Trim(mAnFld(J))
        .Fields(mNmFld).Value = mAyV(J)
    Next
End With
Exit Function
R: ss.R

E:
: ss.B cSub, cMod, "oRs,,pLnFld,,J,NmFld,Val", jj.ToStr_Rs(oRs), , pLnFld, , J, mNmFld, mAyV(J)
    Set_Rs_ByLpVv = True
End Function
#If Tst Then
Function Set_Rs_ByLpAp_Tst() As Boolean
If jj.Dlt_Tbl("xx") Then Stop
If jj.Run_Sql("Create table xx (aa Long, bb Integer, cc Date)") Then Stop
Dim mRs As DAO.Recordset
Set mRs = CurrentDb.TableDefs("xx").OpenRecordset
mRs.AddNew
If jj.Set_Rs_ByLpAp(mRs, "aa,bb,cc", "13", 12, "2007/12/31") Then Stop ' Should have NO error
mRs.Update

mRs.AddNew
If jj.Set_Rs_ByLpAp(mRs, "aa,bb,cc", 13, 12, #1/1/2007#) Then Stop ' Should have NO error
mRs.Update

mRs.AddNew
If jj.Set_Rs_ByLpAp(mRs, "aa,bb,cc", "13a", 12, #1/1/2007#) Then Stop ' Should have error
mRs.Update
mRs.Close
DoCmd.OpenTable ("xx")
End Function
#End If
Function Set_Rs_ByLpAp(oRs As DAO.Recordset, pLnFld$, ParamArray pAp()) As Boolean
Set_Rs_ByLpAp = jj.Set_Rs_ByLpVv(oRs, pLnFld, CVar(pAp))
End Function
Function Set_Prm(pFb$, pTrc&, pNmLgc$, Optional pLm$) As Boolean
'Aim: set {pTrc&, pNmLgc$, pLm$} to table tblPrm
Const cSub$ = "Set_Prm"
Dim mDb As DAO.Database: If jj.Cv_Db_FmFb(mDb, pFb) Then ss.A 1: GoTo E
If jj.Crt_TblPrm(mDb) Then ss.A 3: GoTo E
On Error GoTo R
mDb.Execute "Delete * from tblPrm"
mDb.Execute jj.Fmt_Str("Insert into tblPrm (Trc, NmLgc, Lm) values ({0},'{1}','{2}')", pTrc, pNmLgc, pLm)
If pFb <> "" Then mDb.Close
Exit Function
R: ss.R
E: Set_Prm = True: ss.B cSub, cMod, "pTrc,pNmLgc,pLp,pLv", pTrc, pNmLgc, pLm
End Function
#If Tst Then
Function Set_Prm_Tst() As Boolean
If jj.Set_Prm(1, "abc", "sdf", "dfsdf") Then Stop
DoCmd.OpenTable "tblPrm"
End Function
#End If
Function Set_ActA1(pWb As Workbook) As Boolean
Dim iWs As Worksheet
For Each iWs In pWb.Sheets
    If iWs.Visible Then
        iWs.Activate
        iWs.Range("A1").Activate
    End If
Next
pWb.Sheets(1).Activate
End Function
Function Set_AutoRef() As Boolean
Dim mDirCur$: mDirCur = CurrentDb.Name
Dim mP%: mP = InStrRev(mDirCur, "\")
mDirCur = Left(mDirCur, mP)
Dim mDirObj$: mDirObj = mDirCur & "Working\PgmObj\"
Dim mFfnModU$: mFfnModU = mDirObj & "jj.mda"
If VBA.Dir(mFfnModU$) Then MsgBox (mFfnModU$ & "  not found."): Application.Quit
Dim iRef As Reference: For Each iRef In Application.References
    If iRef.Name = cLib Then Application.References.Remove iRef
Next
Application.References.AddFromFile mFfnModU
End Function
Function Set_Ays(oAy$(), ParamArray pAp()) As Boolean
Dim N%: N = UBound(pAp) + 1
ReDim oAy(N - 1)
Dim J%: For J = 0 To UBound(pAp)
    If IsMissing(pAp(J)) Then
        oAy(J) = ""
    Else
        oAy(J) = pAp(J)
    End If
Next
End Function
Function Set_Ays_Tst() As Boolean
Dim mAys$()
If jj.Set_Ays(mAys, "sdf", "dfsf", , "df") Then Stop
Debug.Print Join(mAys, cComma)
End Function
Function Set_AyAtEnd(oAy$(), pAySrc$(), ParamArray pAp()) As Boolean
'Aim: Return {oAy} by adding {pAy} to end of {oAySrc}
Const cSub$ = "Set_AyAtEnd"
Dim mAyNew: mAyNew = pAp
Dim NNew%: NNew = jj.Siz_Vayv(mAyNew): If NNew = 0 Then Exit Function
Dim NSrc%: NSrc = jj.Siz_Ay(pAySrc)
ReDim mAyTar$(NSrc + NNew - 1)
Dim J%
For J = 0 To NSrc - 1
    mAyTar(J) = pAySrc(J)
Next
For J = 0 To NNew - 1
    mAyTar(J + NSrc) = mAyNew(J)
Next
oAy = mAyTar
Exit Function
E:
: ss.B cSub, cMod, "pAySrc,pAyNew", jj.ToStr_Ays(pAySrc), jj.ToStr_Vayv(CVar(pAp))
    Set_AyAtEnd = True
End Function
#If Tst Then
Function Set_AyAtEnd_Tst() As Boolean
Dim mA$(3), J%, mB$()
For J = 0 To 3
    mA(J) = J * 10
Next
Call jj.Set_AyAtEnd(mB, mA, "sdf", "sdfsf")
Debug.Print Join(mB, cComma)
End Function
#End If
Function Set_Am_F1(oAm() As tMap, pAy$()) As Boolean
Const cSub$ = "Set_Am_F1"
Dim N%: N% = jj.Siz_Am(oAm): If N% <> jj.Siz_Ay(pAy) Then ss.A 1, "Size of oAm() & pAy() are diff", ePrmErr, "oAm Siz,pAy Siz", N, jj.Siz_Ay(pAy): GoTo E
Dim J%: For J% = 0 To N% - 1
    oAm(J%).F1 = pAy(J%)
Next
Exit Function
E:
: ss.B cSub, cMod, "pAy", jj.ToStr_Ays(pAy)
    Set_Am_F1 = True
End Function
Function Set_Am_F2(oAm() As tMap, pAy$()) As Boolean
Const cSub$ = "Set_Am_F2"
Dim N%: N% = jj.Siz_Am(oAm): If N% <> jj.Siz_Ay(pAy) Then ss.A 1, "Size of oAm() & pAy() are diff", ePrmErr, "oAm Siz,pAy Siz", N, jj.Siz_Ay(pAy): GoTo E
Dim J%: For J% = 0 To N% - 1
    oAm(J%).F2 = pAy(J%)
Next
Exit Function
E:
: ss.B cSub, cMod, "pAy", jj.ToStr_Ays(pAy)
    Set_Am_F2 = True
End Function
Function Set_ChdLnk(pFrm As Access.Form, pLnChd$, pMst$, pChd$) As Boolean
Const cSub$ = "Set_ChdLnk"
On Error GoTo R
Dim mAnChd$(): mAnChd = Split(pLnChd, cComma)
Dim J%: For J = 0 To jj.Siz_Ay(mAnChd) - 1
    Dim mSubFrm As SubForm: If jj.Fnd_Ctl(mSubFrm, pFrm, mAnChd(J)) Then GoTo Nxt
    With mSubFrm
        .LinkMasterFields = pMst
        .LinkChildFields = pChd
    End With
Nxt:
Next
Exit Function
R: ss.R

E:
: ss.B cSub, cMod, "pFrm,pLnChd,pChd,pMst", jj.ToStr_Frm(pFrm), , pLnChd, pChd, pMst
    Set_ChdLnk = True
End Function
Function Set_EnableEdt(pFrm As Access.Form, pEnable As Boolean) As Boolean
Const cSub$ = "Set_EnableEdt"
Dim iCtl As Access.Control: For Each iCtl In pFrm.Controls
    If iCtl.Tag = "Edt" Then
        Dim mNmTyp$: mNmTyp = TypeName(iCtl)
        Select Case mNmTyp
        Case "TextBox":  If jj.Set_EnableTBox(iCtl, pEnable) Then ss.A 1: GoTo E
        Case "Check":    If jj.Set_EnableChkB(iCtl, pEnable) Then ss.A 2: GoTo E
        Case "ComboBox": If jj.Set_EnableCBox(iCtl, pEnable) Then ss.A 3: GoTo E
        End Select
    End If
Next
E: Set_EnableEdt = True: ss.B cSub, cMod, "pFrm,pEnable", jj.ToStr_Frm(pFrm), pEnable
End Function
Function Set_EnableChkB(pChkB As Access.CheckBox, pEnable As Boolean) As Boolean
Const cSub$ = "Set_EnableChkB"
pChkB.Enabled = pEnable
pChkB.BorderColor = IIf(pEnable, 65280, 13209)
On Error Resume Next
End Function
Function Set_EnableCBox(pCBox As Access.ComboBox, pEnable As Boolean) As Boolean
Const cSub$ = "Set_EnableCBox"
pCBox.Enabled = pEnable
pCBox.ForeColor = IIf(pEnable, 0, 255)
On Error Resume Next
End Function
Function Set_EnableTBox(pTBox As Access.TextBox, pEnable As Boolean) As Boolean
Const cSub$ = "Set_EnableTBox"
pTBox.Enabled = pEnable
pTBox.ForeColor = IIf(pEnable, 0, 255)
On Error Resume Next
End Function
Function Set_CmdBtnSte(pNmCmdBar$, pLnBtn$, pBtnSte As MsoButtonState) As Boolean
Const cSub$ = "Set_CmdBtn"
On Error GoTo R
Dim mCmdBar As CommandBar: Set mCmdBar = Application.CommandBars(pNmCmdBar)
Dim mAnBtn$(): mAnBtn = Split(pLnBtn, cComma)
Dim mEnabled As Boolean: mEnabled = (pBtnSte <> msoButtonDown)
With mCmdBar
    Dim J%: For J = 0 To jj.Siz_Ay(mAnBtn) - 1
        With .Controls(mAnBtn(J))
            .State = pBtnSte
            .Enabled = mEnabled
        End With
    Next
End With
Exit Function
R: ss.R
E:
: ss.B cSub, cMod, "pNmCmdBar,pLnBtn,pBtnSte", pNmCmdBar, pLnBtn, pBtnSte
    Set_CmdBtnSte = True
End Function
Function Set_Colr_Chk(pChk As Access.CheckBox, pEnable As Boolean) As Boolean
On Error Resume Next
With pChk
    If pEnable Then
        .BorderColor = 65280
    Else
        .BorderColor = 13209
    End If
End With
End Function
Function Set_Colr_Lbl(pLbl As Label, pEnable As Boolean) As Boolean
On Error Resume Next
With pLbl
    If pEnable Then
        .BackColor = 65280
        .ForeColor = 0
    Else
        .BackColor = 13209
        .ForeColor = 16777215
    End If
End With
End Function
Function Set_CtlLayout(pCtl As Access.Control, Optional pLeft! = -1, Optional pTop! = -1, Optional pWdt! = -1, Optional pHgt! = -1) As Boolean
Const cSub$ = "Set_CtlLayout"
On Error GoTo R
With pCtl
    If pTop >= 0 Then .Top = pTop
    If pLeft >= 0 Then .Left = pLeft
    If pHgt >= 0 Then .Height = pHgt
    If pWdt >= 0 Then .Width = pWdt
End With
Exit Function
R: ss.R
E:
: ss.B cSub, cMod, "pCtl,pLeft,pTop,pWdt,pHgt", jj.ToStr_Ctl(pCtl), pLeft, pTop, pWdt, pHgt
    Set_CtlLayout = True
End Function
Function Set_CtlPrp(pCtl As Access.Control, pNmPrp$, pV) As Boolean
Const cSub$ = "Set_CtlPrp"
On Error GoTo R
pCtl.Properties(pNmPrp).Value = pV
Exit Function
R: ss.R
E:
: ss.B cSub, cMod, "pCtl,pNmPrp,pV", jj.ToStr_Ctl(pCtl), pNmPrp, pV
    Set_CtlPrp = True
End Function
Function Set_CtlPrp_InFrm(pFrm As Access.Form, pTagSubStr$, pNmPrp$, pV) As Boolean
Const cSub$ = "Set_CtlPrp_InFrm"
Dim iCtl As Access.Control: For Each iCtl In pFrm.Controls
    If InStr(iCtl.Tag, pTagSubStr) > 0 Then jj.Set_CtlPrp iCtl, pNmPrp, pV
Next
End Function
Function Set_CtlVisible(pFrm As Form, pVisibleTag$, Optional pInVisibleTag$) As Boolean
Dim iCtl As Control
On Error Resume Next
For Each iCtl In pFrm.Controls
    If InStr(iCtl.Tag, pVisibleTag$) Then iCtl.Visible = True
    If InStr(iCtl.Tag, pInVisibleTag$) Then iCtl.Visible = False
Next
End Function
Function Set_Cummulation(pRs As DAO.Recordset, pLoKey$, pValFld$, pSetFld$) As Boolean
'Aim: Set Cummulation of <pValFld> into <pSetFLd> with grouping as defined in list of key fields <pKeyFlds>
'Output: the field pRs->pSetFld will be Updated
'Input : pRs, pKeyFlds, pValFld, pSetFld
''pRs     : Assume it has been sorted in proper order
''pKeyFlds: a list of key fields used as grouping the records in pRs (same records with pKeyFlds value considered as a group)
''pValFld : pValFld is the value field name used to do the cummulation to set the pSetFld.  If ="", use 1 as value.
''pSetFld : the field required to set
'Logic : For each group of records in pRs, the pSetFld will be set to cummulate the field pValFld
'Example: in ATP.mdb: ATP_35_FullSetNew_3Upd_Qty_As_Cummulate_RunCode()
''- Input table is : tmpATP_FullSetNew
''                   FGDmdId / FG / CmpSupTypSeq / CmpSupTyp / DelveryDate / Cmp / Qty / RunningQty
''- pRs      = currentdtable("tblATP_FullSetNew").openrecordset
''             pRs.index = "PrimaryKey"
''             pRs.PrimaryKey is : FGDmdId / FG / Cmp / CmpSupTypSeq / CmpSupTyp / DeliveryDate
''- pKeyFlds = FGDmdId / FG / Cmp
''- pValFld  = "Qty"
''- pSetFld  = RunningQty
Dim mAnFldKey$(): mAnFldKey = Split(pLoKey, cComma)
Dim NKey%: NKey = jj.Siz_Ay(mAnFldKey)
ReDim mAyLasKeyVal(NKey - 1)
Dim J As Byte: For J = 0 To NKey - 1
    mAyLasKeyVal(J) = "xxxx"
Next
Dim mQ_Run As Double
With pRs
    While Not .EOF
        If jj.IsSamKey_ByAnFldKey(pRs, mAnFldKey, mAyLasKeyVal) Then
            If pValFld = "" Then
                mQ_Run = mQ_Run + 1
            Else
                mQ_Run = mQ_Run + Nz(pRs.Fields(pValFld).Value, 0)
            End If
        Else
            If pValFld = "" Then
                mQ_Run = 1
            Else
                mQ_Run = Nz(pRs.Fields(pValFld).Value, 0)
            End If
            For J = 0 To NKey - 1
                mAyLasKeyVal(J) = pRs.Fields(mAnFldKey(J)).Value
            Next
        End If
        .Edit
        .Fields(pSetFld).Value = mQ_Run
        .Update
        .MoveNext
    Wend
End With
End Function
Function Set_Dbl0Dft(pNmt$) As Boolean
'Aim: set all double fields in {pNmt} to have 0 as default value
Dim J%
For J = 0 To CurrentDb.TableDefs(pNmt).Fields.Count - 1
    If CurrentDb.TableDefs(pNmt).Fields(J).Type = dbDouble Then jj.Set_FldDftV pNmt, CurrentDb.TableDefs(pNmt).Fields(J).Name, 0
Next
End Function
Function Set_DocPrp(pWb As Workbook, pDocPrp As tDocPrp) As Boolean
With pDocPrp
    pWb.BuiltinDocumentProperties("Title").Value = .NmRpt
    pWb.BuiltinDocumentProperties("Subject").Value = .NmRptSht & "-" & .NmSess
    pWb.BuiltinDocumentProperties("Author").Value = "Johnson Cheung"
    pWb.BuiltinDocumentProperties("Comments").Value = _
        "Generated @ " & Format(Now(), "yyyy/mm/dd hh:nn:ss") & vbLf & _
        "Generated by " & CurrentDb.Name & vbLf & _
        "Data name: " & .NmData & vbLf & _
        "ExtraPrm : " & .ExtraPrm
    pWb.BuiltinDocumentProperties("Keywords").Value = .NmRptSht & cComma & .NmSess
End With
End Function
#If PDF Then
Function Set_FfnPDF(pFfnPDF$) As Boolean
Dim mDir$: mDir = Fct.Nam_DirNam(pFfnPDF)
Dim mFn$:  mFn = Fct.Nam_FilNam(pFfnPDF)
With gPDF
    .cOption("UseAutosave") = 1
    .cOption("UseAutosaveDirectory") = 1
    .cOption("AutosaveDirectory") = mDir
    .cOption("AutosaveFilename") = mFn
    .cOption("AutosaveFormat") = 0                            ' 0 = PDF
    .cStart
End With
End Function
#End If
Function Set_FilRO(pFfn$) As Boolean
Const cSub$ = "Set_FilRO"
On Error GoTo R
FileSystem.SetAttr pFfn, vbReadOnly
Exit Function
R: ss.R
E:
: ss.B cSub, cMod, "pFfn", pFfn
    Set_FilRO = True
End Function
Function Set_FilRW(pFfn$) As Boolean
Const cSub$ = "Set_FilRW"
On Error GoTo R
FileSystem.SetAttr pFfn, vbNormal
Exit Function
R: ss.R
E:
: ss.B cSub, cMod, "pFfn", pFfn
    Set_FilRW = True
End Function
Function Set_FldDftV(pNmt$, pFldNm$, pDftV) As Boolean
Const cSub$ = "Set_FldDftV"
On Error GoTo R
CurrentDb.TableDefs(pNmt).Fields(pFldNm).DefaultValue = pDftV
Exit Function
R: ss.R
E: Set_FldDftV = True: ss.B cSub, cMod, "pNmt,pFldNm,pDftV", pNmt, pFldNm, pDftV
End Function
Function Set_Formula(pRge As Range, pNRow&, pFormula$) As Boolean
'Aim: Copy formula at {pRge} download {pNRow} (including the row of {pRge}
Const cSub$ = "Set_Formula"
If pNRow <= 0 Then Exit Function
On Error GoTo R
With pRge(1, 1)
    .Formula = pFormula
    .Copy
End With
Dim mWs As Worksheet: Set mWs = pRge.Parent
mWs.Range(pRge(2, 1), pRge(pNRow, 1)).PasteSpecial xlPasteFormulas
Exit Function
R: ss.R
E: Set_Formula = True: ss.B cSub, cMod, "pRge,NRow,pFormula", jj.ToStr_Rge(pRge), pNRow, pFormula
End Function
Function Set_Formula_SumNxtN(pWs As Worksheet, pCno As Byte, pRnoBeg&, pNRow&, pNCol As Byte) As Boolean
Const cSub$ = "Set_Formula_SumNxtN"
Dim mNxt1$: mNxt1 = jj.Cv_Cno2Col(pCno + 1) & pRnoBeg
Dim mNxtN$: mNxtN = jj.Cv_Cno2Col(pCno + pNCol) & pRnoBeg
Dim mCol$: mCol = jj.Cv_Cno2Col(pCno)
With pWs.Range(mCol & pRnoBeg)
    .Formula = jj.Fmt_Str("=Sum({0}:{1})", mNxt1, mNxtN)
    .Copy
End With
pWs.Range(mCol & pRnoBeg & ":" & mCol & pRnoBeg + pNRow - 1).PasteSpecial xlPasteFormulas
Exit Function
R: ss.R
E: Set_Formula_SumNxtN = True: ss.B cSub, cMod, "pWs,pCno,pRnoBeg,pNRow,pNCol", jj.ToStr_Ws(pWs), pCno, pRnoBeg, pNRow, pNCol
End Function
Function Set_Freeze(pWs As Worksheet, pAdr$) As Boolean
With pWs.Range(pAdr)
    .Activate
    .Select
End With
ActiveWindow.FreezePanes = True
End Function
Function Set_HypLnk(pRge As Excel.Range) As Boolean
'Aim: Set any cells within the {pRge} to hyper link to A1 of worksheet if they have the same value
Const cSub$ = "Set_HypLnk"
Dim mWs As Worksheet: Set mWs = pRge.Worksheet
Dim mWb As Workbook: Set mWb = mWs.Parent
Dim mAnWs$(): If jj.Fnd_AnWs_ByWb(mAnWs, mWb) Then GoTo E
Dim N%: N = jj.Siz_Ay(mAnWs)
Dim iCell As Range, V, J%
For Each iCell In pRge
    V = iCell.Value
    If VarType(V) = vbString Then
        V = Left(V, 31)
        For J = 0 To N - 1
            If V = mAnWs(J) Then Call mWs.Hyperlinks.Add(iCell, "", cQSng & mWb.Sheets(mAnWs(J)).Name & "'!A1")
        Next
    End If
Next
Exit Function
E:
: ss.B cSub, cMod, "pAy", jj.ToStr_Rge(pRge)
    Set_HypLnk = True
End Function
Function Set_HypLnk_Tst() As Boolean
Const cFfn$ = "c:\temp\a.xls"
Dim mWb As Workbook: Set mWb = g.gXls.Workbooks.Open(cFfn)
Dim mWs As Worksheet: Set mWs = mWb.Sheets("Index")
mWb.Application.Visible = True
If jj.Set_HypLnk(mWs.Range("A1:E200")) Then Stop
End Function
Function Set_Lck(pFrm As Access.Form, pLck As Boolean, Optional pAlwAdd As Boolean = False, Optional pAlwDlt As Boolean = False) As Boolean
'Aim: Set all controls in {pFrm} as lock
Const cSub$ = "Set_Lck"
Dim iCtl As Access.Control: For Each iCtl In pFrm.Controls
    If Not Visible Then GoTo Nxt
    Dim mLck As Boolean: If iCtl.Tag = "Edt" Then mLck = pLck Else mLck = True
    Select Case TypeName(iCtl)
    Case "Label"
        If IsEnd(iCtl.Name, "_Lbl") Then GoTo Nxt
        If jj.Set_LckLbl(iCtl, mLck) Then ss.A 1: GoTo E
    Case "TextBox":  If jj.Set_LckTBox(iCtl, mLck) Then ss.A 2: GoTo E
    Case "Check":    If jj.Set_LckChkB(iCtl, mLck) Then ss.A 3: GoTo E
    Case "ComboBox": If jj.Set_LckCBox(iCtl, mLck) Then ss.A 4: GoTo E
    End Select
Nxt:
Next
With pFrm
    If pLck Then
        .AllowEdits = False
        .AllowAdditions = False
        .AllowDeletions = False
    Else
        .AllowEdits = True
        .AllowAdditions = pAlwAdd
        .AllowDeletions = pAlwDlt
    End If
End With
pFrm.Repaint
Exit Function
R: ss.R
E: Set_Lck = True: ss.B cSub, cMod, "pFrm,pLck,pAlwAdd,pAlwDlt", jj.ToStr_Frm(pFrm), pLck, pAlwAdd, pAlwDlt
End Function
#If Tst Then
Function Set_Lck_Tst() As Boolean
Const cNmFrm$ = "frmIIC_Tst"
Dim mFrm As Access.Form: If jj.Opn_Frm(cNmFrm, , , mFrm) Then Stop: GoTo E
If jj.Set_Lck(mFrm, False) Then Stop: GoTo E
Stop
If jj.Set_Lck(mFrm, True) Then Stop: GoTo E
Stop
If jj.Set_Lck(mFrm, False) Then Stop: GoTo E
Stop
GoTo X
E: Set_Lck_Tst = True
X: jj.Cls_Frm cNmFrm
End Function
#End If
Function Set_LckCBox(pCBox As Access.ComboBox, pLck As Boolean) As Boolean
Const cSub$ = "Set_LckCBox"
pCBox.Locked = pLck
pCBox.ForeColor = 0
pCBox.TabStop = Not pLck
'pCBox.ForeColor = IIf(pLck, 255, 0)
End Function
Function Set_LckChkB(pChkB As Access.CheckBox, pLck As Boolean) As Boolean
Const cSub$ = "Set_LckChkB"
pChkB.Locked = pLck
pChkB.BorderColor = IIf(pLck, 13209, 65280)
pChkB.TabStop = Not pLck
End Function
Function Set_LckLbl(pLbl As Access.Label, pLck As Boolean) As Boolean
On Error Resume Next
pLbl.ForeColor = IIf(pLck, 16777215, 0)
pLbl.BackColor = IIf(pLck, 13209, 65280)
End Function
#If Tst Then
Function Set_LckLbl_Tst() As Boolean
Const cNmFrm$ = "frmIIC_Tst"
Dim mFrm As Access.Form: If jj.Opn_Frm(cNmFrm, , , mFrm) Then Stop: GoTo E
Dim mLbl As Access.Label: Set mLbl = mFrm.Controls("ICGL_Label")
If jj.Set_LckLbl(mLbl, False) Then Stop: GoTo E
Exit Function
E: Set_LckLbl_Tst = True
End Function
#End If
Function Set_LckTBox(pTBox As Access.TextBox, pLck As Boolean) As Boolean
Const cSub$ = "Set_LckTBox"
pTBox.Locked = pLck
pTBox.ForeColor = 0
pTBox.TabStop = Not pLck
Dim mLbl As Access.Label: If jj.Fnd_Lbl(mLbl, pTBox) Then Exit Function
Set_LckLbl mLbl, pLck
End Function
#If Tst Then
Function Set_LckTBox_Tst() As Boolean
Const cNmFrm$ = "frmIIC_Tst"
Dim mFrm As Access.Form: If jj.Opn_Frm(cNmFrm, , , mFrm) Then Stop: GoTo E
Dim mTBox As Access.TextBox: Set mTBox = mFrm.Controls("ICGL")
If jj.Set_LckTBox(mTBox, False) Then Stop: GoTo E
Exit Function
E: Set_LckTBox_Tst = True
End Function
#End If
Function Set_LstCtlLayout(pFrm As Access.Form, pLnCtl$, Optional pLeft! = -1, Optional pTop! = -1, Optional pWdt! = -1, Optional pHgt! = -1) As Boolean
Const cSub$ = "Set_LstCtlLayout"
Dim mAnCtl$(): mAnCtl = Split(pLnCtl, cComma)
Dim J%: For J = 0 To jj.Siz_Ay(mAnCtl) - 1
    Dim iCtl As Access.Control: If jj.Fnd_Ctl(iCtl, pFrm, mAnCtl(J)) Then GoTo Nxt
    jj.Set_CtlLayout iCtl, pLeft, pTop, pWdt, pHgt
Nxt:
Next
End Function
Function Set_Nm_InWb(pWb As Workbook, pNm$, pReferTo$) As Boolean
Const cSub$ = "Set_Set_Nm_InWb"
On Error GoTo R
Dim mNm As Name: If jj.IsWbNm(pWb, pNm, mNm) Then mNm.RefersTo = pReferTo: Exit Function
pWb.Names.Add pNm, pReferTo$
GoTo X
R: ss.R
E: Set_Nm_InWb = True: ss.B cSub, cMod, "pWb,pNm,pReferTo", ToStr_Wb(pWb), pNm, pReferTo
X:
End Function
Function Set_Nm_InWs(pWs As Worksheet, pNm$, pReferTo$) As Boolean
Const cSub$ = "Set_Set_Nm_InWs"
On Error GoTo R
Dim mNm As Name: If jj.IsWsNm(pWs, pNm, mNm) Then mNm.RefersTo = pReferTo: Exit Function
pWs.Names.Add pNm, pReferTo$
GoTo X
R: ss.R
E: Set_Nm_InWs = True: ss.B cSub, cMod, "pWs,pNm,pReferTo", ToStr_Ws(pWs), pNm, pReferTo
X:
End Function
Function Set_PdfPrt(pSetPdfPrt As Boolean) As Boolean
Const cSub$ = "Set_PdfPrt"
On Error GoTo R
Static xSavPrt$
With gWrd
    If pSetPdfPrt Then
        If Left(.ActivePrinter, 10) = "PDFCreator" Then Exit Function
        xSavPrt = .ActivePrinter
        .ActivePrinter = "PDFCreator"
        Exit Function
    End If
    If xSavPrt <> "" Then If .ActivePrinter <> xSavPrt Then .ActivePrinter = xSavPrt
    xSavPrt = ""
    Exit Function
End With
Exit Function
R: ss.R

E:
: ss.B cSub, cMod, "pSetPdfPrt", pSetPdfPrt
    Set_PdfPrt = True
End Function
#If Tst Then
Function Set_PdfPrt_Tst() As Boolean
Const cSub$ = "Set_PdfPrt_Tst"
jj.Shw_Dbg cSub, cMod
Dim J%: For J = 0 To 10
    Debug.Print J
    jj.Set_PdfPrt True
    jj.Set_PdfPrt False
Next
End Function
#End If
Function Set_Pf_OfWb(pWb As Workbook, pLnPf$, Optional pOrientation As XlOrientation = xlHidden) As Boolean
'Aim: Hide the pivot fields {pLnPf} inside the {pWb}
'Param: pLnPf is a list Pivot Field Name separated by comma
Dim J As Byte, AnPf$(): AnPf = Split(pLnPf, cComma)
Dim iWs As Worksheet
For Each iWs In pWb.Worksheets
    Dim iPt As PivotTable
    For Each iPt In iWs.PivotTables
        For J = LBound(AnPf) To UBound(AnPf)
            On Error Resume Next
            iPt.PivotFields(AnPf(J)).Orientation = pOrientation
            On Error GoTo 0
        Next
    Next
Next
End Function
Function Set_Am_ByF1F2(oAm() As tMap, pF1$, pF2$, Optional pAlwAdd As Boolean = False) As Boolean
'Aim: set one of element of oAm.F2 by lookup by pF1
Const cSub$ = "Set_Am_ByF1F2"
Dim J%, N%: N = jj.Siz_Am(oAm)
For J = 0 To N - 1
    If oAm(J).F1 = pF1 Then oAm(J).F2 = pF2: Exit Function
Next
If Not pAlwAdd Then ss.A 1, "pF1 not in oAm": GoTo E
ReDim Preserve oAm(N)
oAm(N).F1 = pF1
oAm(N).F2 = pF2
Exit Function
E: Set_Am_ByF1F2 = True: ss.B cSub, cMod, "oAm,pF1,pF2", jj.ToStr_Am(oAm), pF1, pF2
End Function
Function Set_Am_ByF1F2_Tst() As Boolean
Dim mLm$: mLm = jj.Fmt_Str("Date={0},XX=,NmDte=aa,NmSess=123,NmRptSht=xxx,NmRpt=xxx,MGIWeekNum={1}", Format(Date, "YYYY_MM_DD"), "Wk" & Fct.MGIWeekNum(Date))
Dim mAm() As tMap: mAm = jj.Get_Am_ByLm(mLm)
Debug.Print "before ================================================-"
Debug.Print jj.ToStr_Am(mAm, , , , vbLf)
If jj.Set_Am_ByF1F2(mAm, "XX", "1111") Then Stop: GoTo E
Debug.Print "after jj.Set_Am_ByF1F2(mAm, ""XX"", ""1111"")----"
Debug.Print jj.ToStr_Am(mAm, , , , vbLf)
Shw_DbgWin
Stop
E: Set_Am_ByF1F2_Tst = True
End Function
Function Set_Prp(pNm$, pTypObj As AcObjectType, pNmPrp$, pVal$) As Boolean
Const cSub$ = "Set_Prp"
Dim mPrp As DAO.Property:
Select Case pTypObj
Case Access.AcObjectType.acQuery
    On Error GoTo Er1
    CurrentDb.QueryDefs(pNm).Properties(pNmPrp).Value = pVal
Case acForm, acModule, acReport, acTable
    Dim mNmTypObj$:  mNmTypObj = jj.ToStr_TypObj(pTypObj)
    On Error GoTo Er2
    CurrentDb.Containers(mNmTypObj).Documents(pNm).Properties(pNmPrp).Value = pVal
Case Else
    ss.A 1, "Given TypObj is not supported", , "Supported Types", "acQuery, acForm, acModule, acReport, acTable": GoTo E
End Select
Exit Function
Er1:
    On Error GoTo 0
    On Error GoTo E
    Set mPrp = CurrentDb.QueryDefs(pNm).CreateProperty(pNmPrp, DAO.DataTypeEnum.dbText, pVal)
    CurrentDb.QueryDefs(pNm).Properties.Append mPrp
    Exit Function
Er2:
    On Error GoTo 0
    On Error GoTo E
    Set mPrp = CurrentDb.Containers(mNmTypObj).Documents(pNm).CreateProperty(pNmPrp, DAO.DataTypeEnum.dbText, pVal)
    CurrentDb.Containers(mNmTypObj).Documents(pNm).Properties.Append mPrp
    Exit Function
E: Set_Prp = True
: ss.B cSub, cMod, "pNm,pTypObj,pNmPrp,pVal", pNm, jj.ToStr_TypObj(pTypObj), pNmPrp, pVal
End Function
Function Set_Prp_Tst() As Boolean
Debug.Print jj.Set_Prp("1Rec", acForm, "Description", "yy")
End Function
Function Set_QryPrp(pQry As QueryDef, pNmPrp$, pVal$) As Boolean
If pNmPrp = "Description" And pVal = "" Then
    On Error Resume Next
    pQry.Properties.Delete pNmPrp: Exit Function
    Exit Function
End If
On Error GoTo R

pQry.Properties(pNmPrp).Value = pVal
Exit Function
R: ss.R
    Dim mPrp As DAO.Property: Set mPrp = pQry.CreateProperty(pNmPrp, DAO.DataTypeEnum.dbText, pVal)
    pQry.Properties.Append mPrp
End Function
Function Set_QryPrp_Bool(pQry As QueryDef, pNmPrp$, pVal As Boolean) As Boolean
On Error GoTo R
pQry.Properties(pNmPrp).Value = pVal
Exit Function
R: ss.R
    Dim mPrp As DAO.Property: Set mPrp = pQry.CreateProperty(pNmPrp, DAO.DataTypeEnum.dbBoolean, pVal)
    pQry.Properties.Append mPrp
End Function
Function Set_Sno(pNmt$, Optional pNmFldSno$ = "Sno", Optional pOrdBy$ = "") As Boolean
Const cSub$ = "Set_Sno"
'-- Fill in <<pNmSeqFld>> starting from 1 by using PrimaryKey as the key
On Error GoTo R
Dim mNmt$: mNmt = jj.Q_S(pNmt, "[]")
Dim mRs As DAO.Recordset: If jj.Opn_Rs(mRs, jj.Fmt_Str("Select {0} from {1}{2}", pNmFldSno, mNmt, jj.Cv_OrdBy(pOrdBy))) Then ss.A 1: GoTo E
Set_Sno = jj.Set_Sno_ByRs(mRs, pNmFldSno$)
Exit Function
R: ss.R
E: Set_Sno = True
: ss.B cSub, cMod, "pNmt,pNmFldSno$", pNmt, pNmFldSno
End Function
#If Tst Then
Function Set_Sno_Tst() As Boolean
If jj.Crt_Tbl_FmLoFld("#Tmp", "aa Text 10, Sno Long") Then Stop: GoTo E
Dim J%
For J = 0 To 10
    If jj.Run_Sql("Insert into [#Tmp] (aa) values ('{0}')") Then Stop: GoTo E
Next
If jj.Set_Sno("#Tmp") Then Stop: GoTo E
Exit Function
E:
    Set_Sno_Tst = True
End Function
#End If
Function Set_Sno_ByRs(pRs As DAO.Recordset, pNmFldSno$) As Boolean
Const cSub$ = "Set_Sno_ByRs"
On Error GoTo R
With pRs
    Dim mSno&
    While Not .EOF
        .Edit
        mSno = mSno + 1
        .Fields(pNmFldSno$).Value = mSno
        .Update
        .MoveNext
    Wend
    .Close
End With
Exit Function
R: ss.R
E: Set_Sno_ByRs = True
: ss.B cSub, cMod, "pNmFldSno$", pNmFldSno$
End Function
Function Set_AyKv_ByRs(oAyKv(), pRs As DAO.Recordset) As Boolean
'Aim: Set first N fields value of {pRs} to {oAyKv}
Const cSub$ = "Set_AyKv_ByRs"
On Error GoTo R
Dim J%, mN%: mN = jj.Siz_Ay(oAyKv)
For J = 0 To mN - 1
    oAyKv(J) = pRs.Fields(J).Value
Next
Exit Function
R: ss.R
E: Set_AyKv_ByRs = True: ss.B cSub, cMod, "Siz(oAyKv),pRs", mN, jj.ToStr_Rs_NmFld(pRs)
End Function
Function Set_Sno_wGp(pNmt$, pLnFldGp$, Optional pOrdBy$ = "", Optional pNmFldSeq$ = "Sno") As Boolean
Const cSub$ = "Set_Sno_wGp"
On Error GoTo R
Dim mNmt$: mNmt = jj.Q_SqBkt(pNmt)
Dim mSql$: mSql = jj.Fmt_Str("Select {2},{0} from {1} Order by {2}{3}", pNmFldSeq, mNmt, pLnFldGp, jj.Cv_Str(pOrdBy, ","))
Dim mRs As DAO.Recordset: If jj.Opn_Rs(mRs, mSql) Then ss.A 1: GoTo E
Dim mAnFldGp$(): mAnFldGp = Split(pLnFldGp, ",")
Dim NGp%: NGp = jj.Siz_Ay(mAnFldGp)
ReDim mAyKvLas(NGp - 1)
Dim mSno%: mSno = 0
With mRs
    While Not .EOF
        If Not jj.IsSamKey(mRs, mAyKvLas) Then
            If jj.Set_AyKv_ByRs(mAyKvLas, mRs) Then ss.A 2: GoTo E
            mSno = 0
        End If
        mSno = mSno + 10
        .Edit
        .Fields(pNmFldSeq).Value = mSno
        .Update
        .MoveNext
    Wend
End With
GoTo X
R: ss.R
E: Set_Sno_wGp = True: ss.B cSub, cMod, "pNmt,pLnFldGp,pOrdBy,pNmFldSeq", pNmt, pLnFldGp, pOrdBy, pNmFldSeq
X:
    jj.Cls_Rs mRs
End Function
Function Set_SetCmdBarEnable(pNmCmdBar$, pEnable As Boolean) As Boolean
Dim iCtl As CommandBarControl
On Error Resume Next
For Each iCtl In Application.CommandBars(pNmCmdBar).Controls
    iCtl.Enabled = pEnable
Next
On Error GoTo 0
End Function
Function Set_SetColWidth(pWs As Worksheet, pColWidthLst$) As Boolean
Dim Ay$(), J%, iColNo As Byte, mRange As Range
Ay = Split(pColWidthLst, cComma)
With pWs
    For J = LBound(Ay) To UBound(Ay)
        iColNo = iColNo + 1
        Set mRange = .Cells(1, iColNo)
        mRange.EntireColumn.ColumnWidth = Ay(J)
    Next
End With
End Function
Function Set_SubTot(pWs As Worksheet, pCno As Byte, pRnoBeg&, pNRow&, Optional pFctNo As Byte = 9) As Boolean
Const cSub$ = "Set_SubTot"
On Error GoTo R
Dim mCol$: mCol = jj.Cv_Cno2Col(pCno)
pWs.Range(mCol & (pRnoBeg + pNRow)).Formula = jj.Fmt_Str("=SUBTOTAL({0},{1})", pFctNo, mCol & pRnoBeg & ":" & mCol & pRnoBeg + pNRow - 1)
Exit Function
R: ss.R
E: Set_SubTot = True: ss.B cSub, cMod, "pWs,pCno,pRnoBeg,pNRow,pFctNo", jj.ToStr_Ws(pWs), pCno, pRnoBeg, pNRow, pFctNo
End Function
Function Set_TblSeqInDesc(pQryPfx$) As Boolean
'Aim   : Set first 2 char of desc of the table tmpQQQ_TTT to NN
'Assume: - The make table query qryQQQ_NN_1_Fm_XXXX will generate tmp table tmpQQQ_TTT.
'        - The select query is  qryQQQ_NN_0_TTT.
'
'Logic : For each "Select query" in format of qryQQQ_NN_0_TTT
'          If tmpQQQ_TTT exist, Find nn & set the desc
'        Next
Const cSub$ = "Set_TblSeqInDesc"
Dim L As Byte: L = Len(pQryPfx): If L = 0 Then ss.A 1, "pQryPfx cannot zero length": GoTo E
Dim iQry As QueryDef, iNmq$
For Each iQry In CurrentDb.QueryDefs
    If Left(iQry.Name, L) <> pQryPfx Then GoTo Nxt
    If iQry.Type <> DAO.QueryDefTypeEnum.dbQSelect Then GoTo Nxt
    iNmq = iQry.Name
    'Get iNN as II_0_PPPP
    'Get iStep as 0
    Dim iNN$:      iNN = mID$(iNmq, L + 2)
    Dim iStep$:    iStep = mID$(iNN, 4, 1)
    If iStep <> "0" Then GoTo Nxt
    'Get iII as II
    'Get iPPPP as PPPP
    Dim iII$:      iII = Left(iNN, 2)
    Dim iPPPP$:    iPPPP = mID$(iNN, 6)
    '-- Get iTmpNmt as tmpMMMM_PPPP
    Dim iTmpNmt$:  iTmpNmt = "tmp" & mID$(pQryPfx, 4) & "_" & iPPPP
    If jj.IsTbl(iTmpNmt) Then
        Call jj.Set_TblSeqInDesc_SetDesc(iTmpNmt, iII)
        Debug.Print iNmq; " "; iTmpNmt; " is set to -----> " & iII
    Else
        Debug.Print iNmq; " "; iTmpNmt; " does not exist"
    End If
Nxt:
Next
Exit Function
E: Set_TblSeqInDesc = True: ss.B cSub, cMod, "pQryPfx", pQryPfx
End Function
Function Set_TblSeqInDesc_SetDesc(pNmt$, pNN$) As Boolean
Dim mDesc$: mDesc = jj.Fnd_Prp(pNmt, acTable, "Description")
jj.Set_Prp pNmt, acTable, "Description", pNN & mID$(mDesc, 3)
End Function
Function Set_TblZero2Null(pNmt, pLnFld_SubStr) As Boolean
Dim mAnFld_SubStr$(): mAnFld_SubStr = Split(pLnFld_SubStr, cComma)
With CurrentDb.TableDefs(pNmt).OpenRecordset
    While Not .EOF
        Dim iSubStr As Byte: For iSubStr = LBound(mAnFld_SubStr) To UBound(mAnFld_SubStr)
            Dim iFld As DAO.Field: For Each iFld In .Fields
                .Edit
                If InStr(iFld.Name, mAnFld_SubStr(iSubStr)) > 0 Then
                    If iFld.Value = 0 Then iFld.Value = Null
                End If
                .Update
            Next
        Next
        .MoveNext
    Wend
    .Close
End With
End Function
Function Set_Lv2ColAtEnd(oRgeLv As Range, pLv$, pWs As Worksheet _
    , Optional pRow1Val$ = "WsOfFollowRge" _
    ) As Boolean
'Aim:

Const cSub$ = "Set_Lv2ColAtEnd"
On Error GoTo R
'Do Set Row1 & pLv in an empty column & Set oRgeLv
Do
    'Do Find mCno_Empty
    Dim mCno_Empty As Byte
    Do
        mCno_Empty = jj.Fnd_Cno_EmptyCell_InRow(pWs, , 255, 1)
        If mCno_Empty = 0 Then ss.A 2: GoTo E
    Loop Until True
    With pWs
    'Set Row1
        pWs.Cells(1, mCno_Empty).Value = pRow1Val
        'Set Lv to empty column
        Dim mAy$(): mAy = Split(pLv, cComma)
        Dim J%, N%: N = jj.Siz_Ay(mAy)
        For J = 0 To N - 1
            .Cells(2 + J, mCno_Empty).Value = mAy(J)
        Next
        'Set oRgeLv
        Set oRgeLv = .Range(.Cells(2, mCno_Empty), .Cells(J + N - 1, mCno_Empty))
    End With
Loop Until True
Exit Function
R: ss.R
E: Set_Lv2ColAtEnd = True: ss.B cSub, cMod, "pLv,pWs,pRow1Val", pLv, jj.ToStr_Ws(pWs), pRow1Val
End Function
Function Set_RgeVdt_ByLv(pRge As Range, pLv$ _
    , Optional pInputTit$ = "Enter value or leave blank" _
    , Optional pInputMsg$ = "Enter one of the value in the list or leave it blank." _
    , Optional pErrTit$ = "Not in the List" _
    , Optional pErrMsg = "Please enter a value in list or leave it blank" _
    ) As Boolean
'Aim: Set the validation of {pRge} to select a list of value {pLv}.
'     'The list of value' will be the stored in the avaliable column of ws [SelectionList]
'          Ws [SelectionList] Row1=Ws Name, Row2=Rge that will use to the list to select value, Row3 and onward will be the selection value
Const cSub$ = "Set_RgeVdt_ByLv"

' Do Build mRgeLv: 'The list of value'
Dim mRgeLv As Range: If jj.Set_Lv2ColAtEnd(mRgeLv, pLv, pRge.Worksheet, pRge.Address) Then ss.A 1: GoTo E

' Do Set Vdt of pRge
Do
    On Error GoTo R
    With pRge.Validation
        .Delete
        Dim mFormula$: mFormula = "=" & mRgeLv.Address
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=mFormula
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = pInputTit
        On Error Resume Next
        .InputMessage = pInputMsg
        On Error GoTo R
        .ErrorTitle = pErrTit
        .ErrorMessage = pErrMsg
        .ShowInput = True
        .ShowError = True
    End With
Loop Until True
Exit Function
R: ss.R
E: Set_RgeVdt_ByLv = True: ss.B cSub, cMod, "pRge,pLv", jj.ToStr_Rge(pRge), pInputTit, pInputMsg, pErrTit, pErrMsg, pLv, pInputTit, pInputMsg, pErrTit, pErrMsg
End Function
#If Tst Then
Function Set_RgeVdt_ByLv_Tst() As Boolean
Const cSub$ = "Set_RgeVdt_ByLv_Tst"
Dim mRge As Range, mLv$
Dim mRslt As Boolean, mCase As Byte: mCase = 1

Dim mWb As Workbook: If jj.Crt_Wb(mWb, "c:\aa.xls", True) Then ss.A 1: GoTo E
mWb.Application.Visible = True
Select Case mCase
Case 1
    Set mRge = mWb.Sheets(1).Range("A1:D5")
    mLv = "aa,bb,cc,11,22,33"
End Select
mRslt = jj.Set_RgeVdt_ByLv(mRge, mLv)
jj.Shw_Dbg cSub, cMod, "mRslt, mRge, mLv", mRslt, jj.ToStr_Rge(mRge), mLv
Exit Function
R: ss.R
E: Set_RgeVdt_ByLv_Tst = True: ss.B cSub, cMod
End Function
#End If
Function Set_Zoom(pWs As Worksheet, pZoom As Byte) As Boolean
Dim Wb As Workbook: Set Wb = pWs.Parent
pWs.Activate
Dim W As Window: For Each W In Wb.Windows
    W.Zoom = pZoom
Next
End Function

