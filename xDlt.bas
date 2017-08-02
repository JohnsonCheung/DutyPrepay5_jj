Attribute VB_Name = "xDlt"
#Const Tst = True
Option Compare Text
Option Explicit
Option Base 0
Const cMod$ = cLib & ".xDlt"
Function Dlt_Qry_XXNN(Optional pDb As DAO.Database = Nothing) As Boolean
'Aim: delete all queies like *_XX??
Const cSub$ = "Dlt_Qry_XXNN"
On Error GoTo R
Dim mAnq$(): If jj.Fnd_Anq_ByLik(mAnq, "*_XX??", pDb) Then ss.A 1: GoTo E
If jj.Dlt_Qry_ByAnq(mAnq, pDb) Then ss.A 2: GoTo E
Exit Function
R: ss.R
E: Dlt_Qry_XXNN = True: ss.B cSub, cMod, "pDb", jj.ToStr_Db(pDb)
End Function
Function Dlt_Qry_YYNN(Optional pDb As DAO.Database = Nothing) As Boolean
'Aim: delete all queies like *_YY??
Const cSub$ = "Dlt_Qry_YYNN"
On Error GoTo R
Dim mAnq$(): If jj.Fnd_Anq_ByLik(mAnq, "*_YY??", pDb) Then ss.A 1: GoTo E
If jj.Dlt_Qry_ByAnq(mAnq, pDb) Then ss.A 2: GoTo E
Exit Function
R: ss.R
E: Dlt_Qry_YYNN = True: ss.B cSub, cMod, "pDb", jj.ToStr_Db(pDb)
End Function
Function Dlt_MnuInXls() As Boolean
'Aim:delete all mnu to each work book
Const cSub$ = "Dlt_MnuInXls"
On Error GoTo R
Dim iWb As Workbook
For Each iWb In Excel.Application.Workbooks
    If Dlt_MnuInWb(iWb) Then ss.A 1: GoTo E
Next
Exit Function
R: ss.R
E: Dlt_MnuInXls = True: ss.B cSub, cMod
End Function
Function Dlt_TBar(pWs As Worksheet, pNmTBar$) As Boolean
Dim iOLEObj As Excel.OLEObject
For Each iOLEObj In pWs.OLEObjects
    If iOLEObj.Name = pNmTBar Then iOLEObj.Delete: Exit Function
Next
End Function
Function Dlt_MnuInWb(pWb As Workbook) As Boolean
'Aim:delete all mnu to each work book
Const cSub$ = "Dlt_MnuInWb"
On Error GoTo R
If Dlt_MnuInPrj(pWb.VBProject, "Mnu" & pWb.CodeName) Then ss.A 1: GoTo E
If Dlt_Prc("J" & mID(pWb.CodeName, 3) & ".g", "Shw_MnuWb") Then ss.A 2: GoTo E
If Dlt_Prc("J" & mID(pWb.CodeName, 3) & ".g", "Shw_MnuWs") Then ss.A 2: GoTo E

Dim iWs As Worksheet
For Each iWs In pWb.Sheets
    If Dlt_MnuInPrj(pWb.VBProject, "Mnu" & iWs.CodeName) Then ss.A 1: GoTo E
Next
Exit Function
R: ss.R
E: Dlt_MnuInWb = True: ss.B cSub, cMod, "pWb", jj.ToStr_Wb(pWb)
End Function
Function Dlt_MnuInPrj(pPrj As VBProject, pNmMnu$) As Boolean
'Aim: add one userform {pNmMnu} to {pPrj} If menu exist, skip adding
Const cSub$ = "Dlt_MnuInPrj"
On Error GoTo R
Dim mVbCmp As VBComponent
If jj.Fnd_VbCmp(mVbCmp, pPrj, pNmMnu) Then Exit Function
pPrj.VBComponents.Remove mVbCmp
Exit Function
R: ss.R
E: Dlt_MnuInPrj = True: ss.B cSub, cMod, "pPrj,pNmMnu", jj.ToStr_Prj(pPrj), pNmMnu
End Function
Function Dlt_Prc(pMod$, pNmPrc$) As Boolean
'Aim: Delete {pNmPrc} in {pNmPrc_Nmm}.
Const cSub$ = "Dlt_Prc"
On Error GoTo R
Dim mMd As CodeModule: If jj.Fnd_Md_ByNm(mMd, pMod) Then ss.A 1: GoTo E
Dim mLin&: If jj.Dlt_Prc_ByMd(mLin, mMd, pNmPrc) Then ss.A 2: GoTo E
Exit Function
R: ss.R
E: Dlt_Prc = True: ss.B cSub, cMod, "pMod,pNmPrc", pMod, pNmPrc
End Function
Function Dlt_Prc_ByMd(oLin&, pMd As CodeModule, pNmPrc$) As Boolean
'Aim: dlt prc if exist and return {oLin} that it starts.  If it not exist return oLin as the last line.
Const cSub$ = "Dlt_Prc_ByMd"
On Error GoTo R
Dim mRgeRno As tRgeRno
If jj.Fnd_PrcRgeRno_ByMd(mRgeRno, pMd, pNmPrc) Then ss.A 1: GoTo E
With mRgeRno
    If .Fm > 0 And .To > 0 And .To >= .Fm Then
        pMd.DeleteLines .Fm, .To - .Fm + 1
        oLin = .Fm
        Exit Function
    End If
End With
oLin = pMd.CountOfLines + 1
Exit Function
R: ss.R
E: Dlt_Prc_ByMd = True: ss.A cSub, cMod, "pMd,pNmPrc", jj.ToStr_Md(pMd), pNmPrc
End Function
Function Dlt_Prc_ByMd_Tst() As Boolean
Dim mMd As CodeModule: If jj.Fnd_Md_ByNm(mMd, "jj.xDlt") Then Stop: GoTo E
Dim mLin&: If Dlt_Prc_ByMd(mLin, mMd, "aaa") Then Stop: GoTo E
Debug.Print mLin
jj.Shw_DbgWin
Exit Function
E: Dlt_Prc_ByMd_Tst = True
End Function
Function Dlt_Cmt(pRge As Range) As Boolean
Dim mCmt As Comment: Set mCmt = pRge.Comment
If TypeName(mCmt) = "Nothing" Then Exit Function
mCmt.Delete
End Function
Function Dlt_RowNotInAy(pRge As Range, pAy$()) As Boolean
'Aim: for all data downward from {pRge} delete any row having value not in {pAy}
Const cSub$ = "Dlt_RowNotInAy"
On Error GoTo R
Dim mRnoLas&: If jj.Fnd_RnoLas(mRnoLas, pRge) Then ss.A 1: GoTo E
Dim iRCnt&
For iRCnt = mRnoLas - pRge.Row + 1 To 1 Step -1
    Dim J%: If jj.Fnd_Idx(J, pAy, pRge(iRCnt, 1).Value) Then Stop: GoTo E
    If J = -1 Then pRge.Rows(iRCnt).EntireRow.Delete
Next
Exit Function
R: ss.R
E: Dlt_RowNotInAy = True: ss.B cSub, cMod, "pRge,pAy", jj.ToStr_Rge(pRge), jj.ToStr_Ays(pAy)
End Function
#If Tst Then
Function Dlt_RowNotInAy_Tst() As Boolean
If jj.Cpy_Fil("P:\AppDef_Meta\MetaLgc.xls", "c:\tmp\bb.xls", True) Then Stop: GoTo E
Dim mWb As Workbook: If jj.Opn_Wb_RW(mWb, "c:\tmp\bb.xls", , True) Then Stop: GoTo E
Dim mAn$(): jj.Set_Ays mAn, "qryPrepImpTblTy", "qryPrepImpTblTyp", "qryExpTblF", "qryExpTbl"
If Dlt_RowNotInAy(mWb.Sheets("OldQsT").Range("B5"), mAn) Then Stop: GoTo E
Stop
GoTo X
E: Dlt_RowNotInAy_Tst = True
X: jj.Cls_Wb mWb, False, True
End Function
#End If
Function Dlt_DupRow(pRge As Range) As Boolean
'Aim: for all data downward from {pRge} delete any duplicate row
Const cSub$ = "Dlt_DupRow"
Dim mRnoLas&: If jj.Fnd_RnoLas(mRnoLas, pRge) Then Stop: GoTo E
Stop
'Dim J%
'For J = pRnoFm To mRnoLas - 2
'    Dim I%
'    I = J + 1
'    Dim V: V = pWs.Cells(J, pCol).Value
'    While I < mRnoLas - 1
'        If V = pWs.Cells(I, pCol).Value Then
'            pWs.Rows(I).Delete
'            mRnoLas = mRnoLas - 1
'        End If
'        I = I + 1
'    Wend
'Next
Exit Function
E: Dlt_DupRow = True: ss.B cSub, cMod, "pRge", jj.ToStr_Ws(pRge)
End Function
Function Dlt_Fil_BySfx(pDir$, pSfx$) As Boolean
Const cSub$ = "Dlt_Fil_BySfx"
Dim mAyFn$()
If jj.Fnd_AyFn_ByLik(mAyFn, pDir, "*" & pSfx) Then ss.A 1: GoTo E
If jj.Dlt_Fil_ByAy(pDir, mAyFn) Then ss.A 2: GoTo E
Exit Function
E: Dlt_Fil_BySfx = True: ss.B cSub, cMod, "pDir,pSfx", pDir, pSfx
End Function
Function Dlt_Fil_ByPfx(pDir$, pPfx$) As Boolean
Const cSub$ = "Dlt_Fil_ByPfx"
Dim mAyFn$()
If jj.Fnd_AyFn_ByLik(mAyFn, pDir, pPfx & "*") Then ss.A 1: GoTo E
If jj.Dlt_Fil_ByAy(pDir, mAyFn) Then ss.A 2: GoTo E
Exit Function
E: Dlt_Fil_ByPfx = True: ss.B cSub, cMod, "pDir,pPfx", pDir, pPfx
End Function
Function Dlt_Fil_ByAy(pDir$, pAyFn$()) As Boolean
Const cSub$ = "Dlt_Fil_ByAy"
Dim J%
For J = 0 To jj.Siz_Ay(pAyFn) - 1
    If jj.Dlt_Fil(pDir & pAyFn(J)) Then ss.A 1: GoTo E
Next
Exit Function
E: Dlt_Fil_ByAy = True: ss.B cSub, cMod, "pDir,pAyFn", pDir, jj.ToStr_Ays(pAyFn)
End Function
Function Dlt_Rel(pNmRel$, Optional pDb As DAO.Database = Nothing) As Boolean
Const cSub$ = "Dlt_Rel"
On Error GoTo R
jj.Cv_Db(pDb).Relations.Delete pNmRel
R: ss.R
E: Dlt_Rel = True: ss.B cSub, cMod, "pRel,pDb", jj.ToStr_Rel(pNmRel), jj.ToStr_Db(pDb)
End Function
Function Dlt_RelAll(Optional pDb As DAO.Database = Nothing) As Boolean
Dim mDb As DAO.Database: Set mDb = jj.Cv_Db(pDb)
With mDb.Relations
    While .Count >= 1
        .Delete mDb.Relations(0).Name
    Wend
End With
End Function
#If Tst Then
Function Dlt_RelAll_Tst() As Boolean
Const cFbMeta$ = "C:\Tmp\WorkingDir\Meta_Data.Mdb"
Dim mDb As DAO.Database
If jj.Opn_Db(mDb, cFbMeta, False) Then Stop
If Dlt_RelAll(mDb) Then Stop
mDb.Close
If jj.Opn_CurDb(jj.g.gAcs, cFbMeta) Then Stop
jj.g.gAcs.Visible = True
End Function
#End If
Function Dlt_Idx(pNmt$, pNmIdx$, Optional pDb As DAO.Database = Nothing) As Boolean
Const cSub$ = "Dlt_Idx"
If Not jj.IsIdx(pNmt, pNmIdx, pDb, True) Then Exit Function
Dim mSql$: mSql = jj.Fmt_Str("Drop Index {0} on [{1}]", pNmIdx, jj.Rmv_SqBkt(pNmt))
If jj.Run_Sql_ByDbExec(mSql, pDb) Then ss.A 1: GoTo E
Exit Function
R: ss.R
E: Dlt_Idx = True: ss.B cSub, cMod, "pNmt,pNmIdx,pDb", pNmt, pNmIdx, jj.ToStr_Db(pDb)
End Function
Function Dlt_Ws_Excpt(pWb As Workbook, pWsNmExcpt$) As Boolean
'Aim: delete all ws except {pWsExcpt}
Const cSub$ = "Dlt_Ws_Excpt"
On Error GoTo R
pWb.Application.DisplayAlerts = False
While pWb.Sheets.Count >= 2
    If pWb.Sheets(1).Name = pWsNmExcpt Then
        pWb.Sheets(2).Delete
    Else
        pWb.Sheets(1).Delete
    End If
Wend
pWb.Application.DisplayAlerts = True
Exit Function
R: ss.R
E: Dlt_Ws_Excpt = True: ss.B cSub, cMod, "pWb,pWsNmExcpt", jj.ToStr_Wb(pWb), pWsNmExcpt
End Function
Function Dlt_Ws_Excpt_Tst() As Boolean
Dim mWb As Workbook: If jj.Crt_Wb(mWb, "c:\aa.xls", True) Then Stop
mWb.Sheets.Add
mWb.Sheets.Add
mWb.Sheets.Add
mWb.Application.Visible = True
Stop
If jj.Dlt_Ws_Excpt(mWb, "ToBeDelete") Then Stop
Stop
mWb.Close True
End Function
Function Dlt_TxtSpec(pNmSpec$, Optional pDb As DAO.Database = Nothing) As Boolean
'Aim: Delete all records in MSysIMEXSpecs & MSysIMEXColumns for SpecName={pNmSpec}
'     MSysIMEXSpecs  : DateDelim,DateFourDigitYear,DateLeadingZeros,DateOrder,DecimalPoint,FieldSeparator,FileType,SpecID,SpecName,SpecType,StartRow,TextDelim,TimeDelim
'     MSysIMEXColumns: Attributes,DataType,FieldName,IndexType,SkipColumn,SpecID,Start,Width
Const cSub$ = "Dlt_TxtSpec"
Dim mDb As DAO.Database: Set mDb = jj.Cv_Db(pDb)
If pNmSpec = "*" Then
    Dim mAnTxtSpec$(): If jj.Fnd_AnTxtSpec(mAnTxtSpec, pDb) Then ss.A 1: GoTo E
    If jj.Siz_Ay(mAnTxtSpec) = 0 Then MsgBox "No Txt Spec is found", , "Delete Txt Spec for importing": Exit Function
    If MsgBox("Are your sure to delete all following Txt Spec?" & vbLf & vbLf & Join(mAnTxtSpec, vbLf), vbYesNo) = vbNo Then Exit Function
    If jj.Run_Sql_ByDbExec("Delete * from MSysIMEXSpecs", mDb) Then ss.A 2: GoTo E
    If jj.Run_Sql_ByDbExec("Delete * from MSysIMEXColumns", mDb) Then ss.A 2: GoTo E
    Exit Function
End If
Dim mTxtSpecId&: If jj.Fnd_TxtSpecId(mTxtSpecId, pNmSpec, mDb) Then Exit Function
mDb.Execute "Delete * from MSysIMEXSpecs where SpecId=" & mTxtSpecId
mDb.Execute "Delete * from MSysIMEXColumns where SpecId=" & mTxtSpecId
Exit Function
R: ss.R
E: Dlt_TxtSpec = True: ss.B cSub, cMod, "pNmSpec,pDb", pNmSpec, jj.ToStr_Db(pDb)
End Function
Function Dlt_TxtSpec_Tst() As Boolean
If jj.Dlt_TxtSpec("*") Then Stop
End Function
Function Dlt_AllWs_Except1(pWb As Workbook, Optional pNmWs$ = "") As Boolean
'Aim: Delete all the worksheet except the first ws or the given {pNmWs}
Const cSub$ = "Dlt_AllWs_Except1"
Dim N%: N = pWb.Sheets.Count
While N > 1
    If pNmWs$ = "" Then
        pWb.Sheets(pWb.Sheets.Count).Delete
    Else
        If pWb.Sheets(1).Name = pNmWs Then
            pWb.Sheets(2).Delete
        Else
            pWb.Sheets(1).Delete
        End If
    End If
    N = pWb.Sheets.Count
Wend
Exit Function
R: ss.R
E: Dlt_AllWs_Except1 = True: ss.B cSub, cMod, "pWb", jj.ToStr_Wb(pWb)
End Function
#If Tst Then
Function Dlt_AllWs_Except1_Tst() As Boolean
Dim mWb As Workbook
If jj.Crt_Wb(mWb, "c:\tmp\aa.xls") Then Stop
If jj.Dlt_AllWs_Except1(mWb, "Sheet2") Then Stop
mWb.Application.Visible = True
End Function
#End If
Function Dlt_Dir(pDir$, Optional pFfnSpec$ = "*.*") As Boolean
Const cSub$ = "Dlt_Dir"
'Aim: Delete files as {pFfnSpec} in {pDir}.  Return false and show message if some file cannot be deleted.
'==
If Not jj.IsDir(pDir, True) Then ss.A 1: GoTo E
Dim mAyFn$():  If jj.Fnd_AyFn(mAyFn, pDir, pFfnSpec) Then ss.A 1: GoTo E
If jj.Siz_Ay(mAyFn) = 0 Then Exit Function
Dim iFn: For Each iFn In mAyFn
    If jj.Dlt_Fil(pDir & iFn) Then ss.A 2, "iFn in pDir cannot be deleted", eRunTimErr, "iFn", iFn: GoTo E
Next
Exit Function
R: ss.R
E: Dlt_Dir = True: ss.B cSub, cMod, "pDir,pFfnSpec", pDir, pFfnSpec
End Function
Function Dlt_Fil(pFfn$, Optional pIgnoreRO As Boolean = False) As Boolean
Const cSub$ = "Dlt_Fil"
If VBA.Dir(pFfn) = "" Then Exit Function
If pIgnoreRO Then If jj.Set_FilRW(pFfn) Then ss.A 1: GoTo E
On Error GoTo R
Kill pFfn
On Error GoTo 0
If VBA.Dir(pFfn) <> "" Then ss.A 2, "Fil exist but cannot delete": GoTo E
Exit Function
R: ss.R
E: Dlt_Fil = True: ss.B cSub, cMod, "pFfn", pFfn
End Function
Function Dlt_Host_ByFrm(pNmtHost$, pDsn$, pFrm As Access.Form, pLmPk$, pLnFld$, Optional pLmPk_FriendlyName$ = "", Optional pMsg$ = "") As Boolean
'Aim: This Function Dlt_is called Form's OnDelete.  Each record going to be deleted will have values in the Controls in {pFrm}
'     Verify the old value of in the list of {pLnFld} is same as the host table {pNmtHost} through {pDsn}
'     If some field is not same, prompt user that the local will be sync from host, then return error.
'     Prompt for delete
'     Then, delete the host record
Const cSub$ = "Dlt_Host_ByFrm"
Const cNmqOdbc_DltHost$ = "qryUpdHostByFrm_DltRec"

'Return if gIsLclMd
If jj.SysCfg_IsLclMd Then Exit Function

jj.Shw_Sts jj.Fmt_Str("Deleting  record in [{0}] through [{1}].  PK Fields [{2}] ....", pNmtHost, pDsn, pLmPk)

'ChkHost
Dim mHostSts As eHostSts
If jj.Chk_Host_ByFrm(mHostSts, pNmtHost, pDsn, pFrm, pLmPk, pLnFld) Then
    Select Case mHostSts
    Case e0Rec, eHostCpyToFrm:                  Exit Function
    Case e1Rec, e2Rec, eUnExpectedErr: GoTo E
    Case Else:                                  ss.A 1, "Logic Error in jj.Chk_Host_ByFrm: return invalid value in mHostSts[" & mHostSts & "]", eCritical: GoTo E
    End Select
End If

'Ask
Dim mPKCndn$: If jj.Bld_LExpr_InFrm(mPKCndn, pFrm, pLmPk) Then ss.A 1: GoTo E
If Not Fct.Start("Record:||" & mPKCndn, "Delete " & pMsg & "?") Then GoTo E

'Dlt Host
Dim mSql$: mSql = jj.ToSql_Dlt(pNmtHost, mPKCndn$)
If jj.Crt_Qry_ByDSN(cNmqOdbc_DltHost, mSql, pDsn, False) Then ss.A 2: GoTo E
If jj.Run_Qry_ByOpnQry(cNmqOdbc_DltHost) Then ss.A 3, "Error in deleting host": GoTo E
ss.xx 4, cSub, cMod, eUsrInfo, "Both Local and Host record are DELETED", "mSql", mSql
GoTo X
R: ss.R
E: Dlt_Host_ByFrm = True: ss.B cSub, cMod, "pNmtHost,pDsn,pFrm,pLmPk,pLnFld,pLmPk_FriendlyName,pMsg", pNmtHost, pDsn, jj.ToStr_Frm(pFrm), pLmPk, pLnFld, pLmPk_FriendlyName, pMsg
X: jj.Clr_Sts
End Function
Function Dlt_Qry(pNmq$, Optional pDb As DAO.Database = Nothing) As Boolean
Const cSub$ = "Dlt_Qry"
On Error GoTo R
Dim mDb As DAO.Database: Set mDb = jj.Cv_Db(pDb)
mDb.QueryDefs.Delete pNmq
Exit Function
R: ss.R
E: Dlt_Qry = True: ss.B cSub, cMod, "pNmq,pDb", pNmq, jj.ToStr_Db(pDb)
End Function
Function Dlt_Qry_ByAnq(pAnq$(), Optional pDb As DAO.Database = Nothing) As Boolean
Const cSub$ = "Dlt_Qry_ByAnq"
Dim A$, J%
For J = 0 To jj.Siz_Ay(pAnq) - 1
    If jj.Dlt_Qry(pAnq(J), pDb) Then A = jj.Add_Str(A, pAnq(J))
Next
If Len(A) <> 0 Then ss.A 1, "Some query cannot be deleted", eRunTimErr, "The queries cannot be deleted", A: GoTo E
Exit Function
R: ss.R
E: Dlt_Qry_ByAnq = True: ss.B cSub, cMod, "pAnq,pDb", Join(pAnq, ","), jj.ToStr_Db(pDb)
End Function
Function Dlt_Qry_ByPfx(pPfx$, Optional pDb As DAO.Database = Nothing) As Boolean
Const cSub$ = "Dlt_Qry_ByPfx"
Dim mAnq$(): If jj.Fnd_Anq_ByPfx(mAnq, pPfx, pDb) Then ss.A 1: GoTo E
Dlt_Qry_ByPfx = Dlt_Qry_ByAnq(mAnq, pDb)
Exit Function
R: ss.R
E: Dlt_Qry_ByPfx = True: ss.B cSub, cMod, "pPfx,pDb", pPfx, jj.ToStr_Db(pDb)
End Function
Function Dlt_Tbl(pNmt$, Optional pDb As DAO.Database = Nothing) As Boolean
Const cSub$ = "Dlt_Tbl"
Dim mDb As DAO.Database: Set mDb = jj.Cv_Db(pDb)
If Not jj.IsTbl(pNmt, mDb) Then Exit Function
On Error GoTo R
If Left(pNmt, 1) = "[" And Right(pNmt, 1) = "]" Then
    mDb.TableDefs.Delete mID(pNmt, 2, Len(pNmt) - 2)
Else
    mDb.TableDefs.Delete pNmt
End If
Exit Function
R: ss.R
E: Dlt_Tbl = True: ss.B cSub, cMod, "pNmt,pDb", pNmt, jj.ToStr_Db(pDb)
End Function
Function Dlt_Tbl_ByLnk() As Boolean
Const cSub$ = "Dlt_Tbl_ByLnk"
'Aim: Delete all linked table in currentdb
jj.Shw_Sts "Deleting all Link Tables  ..."
Dim mAnt_Lnk$(): If jj.Fnd_Ant_ByLnk(mAnt_Lnk$) Then ss.A 1: GoTo E
Dim J%
For J = 0 To jj.Siz_Ay(mAnt_Lnk) - 1
    If jj.Dlt_Tbl(mAnt_Lnk(J)) Then ss.A 2: GoTo E
Next
GoTo X
R: ss.R
E: Dlt_Tbl_ByLnk = True: ss.B cSub, cMod
X:
    jj.Clr_Sts
End Function
Function Dlt_Tbl_ByPfx(pPfx$, Optional pDb As DAO.Database = Nothing) As Boolean
Const cSub$ = "Dlt_Tbl_ByPfx"
Dim mDb As DAO.Database: Set mDb = jj.Cv_Db(pDb)
Dim L%: L = Len(pPfx)
Dim mColl As New VBA.Collection
Dim iTbl As TableDef: For Each iTbl In mDb.TableDefs
    If Left(iTbl.Name, L) = pPfx Then mColl.Add iTbl.Name
Next
Dim mA$
Dim mNmt: For Each mNmt In mColl
    If jj.Dlt_Tbl(CStr(mNmt), mDb) Then mA = jj.Add_Str(mA, CStr(mNmt))
Next
mDb.TableDefs.Refresh
If Len(mA) <> 0 Then ss.A 1, "These tables cannot be deleted: " & mA: GoTo E
Exit Function
E: Dlt_Tbl_ByPfx = True: ss.B cSub, cMod, "pPfx,pDb", pPfx, jj.ToStr_Db(pDb)
End Function
Function Dlt_Ws(pWs As Worksheet) As Boolean
Const cSub$ = "Dlt_Ws_InWb"
On Error GoTo R
Dim mXls As Excel.Application: Set mXls = pWs.Application
mXls.DisplayAlerts = False
pWs.Delete
mXls.DisplayAlerts = True
Exit Function
R: ss.R
E: Dlt_Ws = True: ss.B cSub, cMod, "Ws", jj.ToStr_Ws(pWs)
End Function
Function Dlt_Ws_InWb(pWb As Workbook, pNmWs$) As Boolean
Const cSub$ = "Dlt_Ws_InWb"
On Error GoTo R
If jj.Dlt_Ws(pWb.Worksheets(pNmWs)) Then ss.A 1: GoTo E
Exit Function
R: ss.R
E: Dlt_Ws_InWb = True: ss.B cSub, cMod, "pWb,pNmWs", jj.ToStr_Wb(pWb), pNmWs
End Function
