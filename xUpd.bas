Attribute VB_Name = "xUpd"
Option Compare Text
Option Explicit
Option Base 0
Const cMod$ = cLib & ".xUpd"
Function Upd_Fld_ByNm(pItm$, pNmFld$, pKey$, pV) As Boolean
'Aim: Update the {pNmFld} of table [$<pItm>] by {pV}.  Assume there is unique {Nm<pItm>} index
Const cSub$ = "Upd_Fld_ByNm"
On Error GoTo R
Dim mSql$: mSql = jj.ToSql_Sel("$" & pItm, pNmFld, "Nm" & pItm & "='" & pKey & "'")
Dim mRs As DAO.Recordset: If jj.Opn_Rs(mRs, mSql) Then ss.A 1: GoTo E
With mRs
    If .AbsolutePosition = -1 Then ss.A 2, "No record in [$<pItm>] with given value {pV} in field [Nm<pItm>]": GoTo E
    .Edit
    .Fields(pNmFld).Value = pV
    .Update
End With
GoTo X
R: ss.R
E: Upd_Fld_ByNm = True: ss.B cSub, cMod, "pItm,pNmFld,pV", pItm, pNmFld, pV
X: jj.Cls_Rs mRs
End Function
Function Upd_Host_ByFrm(pNmtHost$, pDsn$, pFrm As Form, pLmPk$, pLmFld$, Optional pLoConst$ = "") As Boolean
'Aim: This function is called Form's Before Update.  The Controls in {pFrm} contains old and new value.
'     Verify the old value of in the list of {pLmFld} is same as the host table {pNmtHost} through {pDsn}
'           If some field is not same, prompt user that the local will be sync from host, then Update the local record & exit
'     Then, Update the host record
Const cSub$ = "Upd_Host_ByFrm"
Const cNmqOdbc_UpdHost$ = "qryUpdHostByFrm_UpdRec"
Const cNmqOdbc_InsHost$ = "qryUpdHostByFrm_InsRec"
'Return if gIsLclMd
If jj.SysCfg_IsLclMd Then Exit Function
jj.Shw_Sts jj.Fmt_Str("Updating [{0}] through [{1}].  Fields [{2}] ....", pNmtHost, pDsn, pLmFld)

'ChkHost
Dim mHostSts As eHostSts
Dim mSql$
If jj.Chk_Host_ByFrm(mHostSts, pNmtHost, pDsn, pFrm, pLmPk, pLmFld) Then
    Select Case mHostSts
    Case e0Rec
        If Not pFrm.AllowAdditions Then ss.A 1, "No Host Rec & Not allowed insert": GoTo E ' Return error, but only display message.  This will let the caller no Update to local record
        ''Insert Rec to Host
        If jj.BldSql_Ins_ByFrm(mSql, pNmtHost, pFrm, pLmPk, pLmFld, pLoConst) Then ss.A 4: GoTo E
        If jj.Crt_Qry_ByDSN(cNmqOdbc_InsHost, mSql, pDsn, False) Then ss.A 4: GoTo E
        If jj.Run_Qry_ByOpnQry(cNmqOdbc_InsHost) Then ss.A 5, "Error in inserting to Host", eUsrInfo: GoTo E
        ss.A 6, "New record is added to Host", eUsrInfo, "mSql", mSql
        Exit Function
    Case e1Rec, e2Rec, eUnExpectedErr: GoTo E
    Case eHostCpyToFrm: Exit Function
    Case Else
        ss.A 6, "Logic Err in jj.Chk_Host_ByFrm.  mHostSts=[" & mHostSts & "]", eCritical: GoTo E
    End Select
End If

'UpdHost
Dim mLmFld$: mLmFld = pLmFld: If pLoConst <> "" Then mLmFld = mLmFld & cComma & pLoConst
Dim mLoChgd$: If jj.BldSql_Upd_InFrm(mSql, pNmtHost, pFrm, pLmPk, mLmFld, mLoChgd) Then ss.A 1: GoTo E
If mSql <> "" Then
    If jj.Crt_Qry_ByDSN(cNmqOdbc_UpdHost, mSql, pDsn, False) Then ss.A 8: GoTo E
    If jj.Run_Qry_ByOpnQry(cNmqOdbc_UpdHost) Then ss.A 9, "Error in updating to Host": GoTo E
    ss.xx 10, cSub, cMod, eUsrInfo, "Both Local and Host record are UpdateD", "mSql,Changed", mSql, mLoChgd
End If
GoTo X
R: ss.R
E: Upd_Host_ByFrm = True: ss.B cSub, cMod, "pNmtHost, pLmPk", pNmtHost, pLmPk
X: jj.Clr_Sts
End Function
Function Upd_Host_ByFrm_Tst() As Boolean
Const cSub$ = "Upd_Host_ByFrm_Tst"
Const cNmFrm$ = "frmIIC_Tst"
Dim mFrm As Access.Form: If jj.Opn_Frm(cNmFrm, , , mFrm) Then Stop: GoTo E
Dim mNmtHost$:
Dim mDsn$:      mDsn = "FETEST_QGPL"
Dim mLstPK$:    mLstPK = "ICLAS"
Dim mLnFld$: mLnFld = ""
Dim mLoConst$: mLoConst = ""
If jj.Upd_Host_ByFrm(mNmtHost, mDsn, mFrm, mLstPK$, mLnFld$, mLoConst$) Then Stop: GoTo E
Exit Function
E: Upd_Host_ByFrm_Tst = True
End Function
Function Upd_Tbl_ToTbl(pNmtTar$, pNmtSrc$, pNKFld As Byte) As Boolean
'Aim: Upd {pNmtSrc} to {pNmtTar} for records exist in both tables
'     assuming first {pNKFld} are common primary in both tables
Const cSub$ = "Upd_Tbl_ToTbl"
Dim mSqlAdd$, mSqlUpd$, mSqlDlt$
If jj.BldSql_AddUpdDlt(mSqlAdd, mSqlUpd, mSqlDlt, pNmtTar, pNmtSrc, pNKFld) Then ss.A 1: GoTo E
If jj.Run_Sql(mSqlUpd) Then ss.A 1: GoTo E
Exit Function
R: ss.R
E: Upd_Tbl_ToTbl = True: ss.B cSub, cMod, "pNmtTar,pNmtSrc,pNKFld", pNmtTar, pNmtSrc, pNKFld
End Function
