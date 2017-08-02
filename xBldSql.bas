Attribute VB_Name = "xBldSql"
#Const Tst = True
Option Compare Text
Option Explicit
Option Base 0
Const cMod$ = cLib & ".xBldSql"
Function BldSql_Qt(oSql$, pRge As Range, pNmtq$, pFbSrc$) As Boolean
'Build oSql for a Qt with pRge as the first data row & field field
Const cSub$ = "BldSql_Qt"
Dim mNmtq0$: mNmtq0 = jj.Rmv_SqBkt(pNmtq)
Dim mAnFld_RnoNmFld$()
Dim mCnoLas As Byte
If jj.Fnd_CnoLas(mCnoLas, pRge(0, 1)) Then ss.A 1: GoTo E
If mCnoLas = 0 Then ss.A 2, "The pRge.Row has no data": GoTo E
Do
    Dim iCno As Byte, mN%: mN = 0
    For iCno = pRge.Column To mCnoLas
        Dim mV: mV = pRge(0, iCno).Value
        If VarType(mV) = vbString Then
            ReDim Preserve mAnFld_RnoNmFld(mN)
            mAnFld_RnoNmFld(mN) = mV
            mN = mN + 1
        End If
    Next
    If mN = 0 Then ss.A 6, "No valid fields in Rno NmFld": GoTo E
Loop Until True

Dim mAnFld_Common$()
Do
    Dim mAnFld_Nmtq$(), mAyFld_Nmtq_IsBool() As Boolean
    Do
        Dim mDb As DAO.Database: If jj.Cv_Db_FmFb(mDb, pFbSrc) Then ss.A 7: GoTo E
        Dim mLnFld_Nmtq$: If jj.Fnd_LnFld_ByNmtq(mLnFld_Nmtq, pNmtq, mDb, True) Then jj.Cls_Db mDb: ss.A 8: GoTo E
        jj.Cls_Db mDb
        
        Dim mAnFld$(): mAnFld = Split(mLnFld_Nmtq, cComma)
        Dim N%: N = jj.Siz_Ay(mAnFld)
        ReDim mAnFld_Nmtq(N - 1), mAyFld_Nmtq_IsBool(N - 1)
        Dim J%
        For J = 0 To N - 1
            Dim mB$: If jj.Brk_Str_Both(mAnFld_Nmtq(J), mB, mAnFld(J), ":") Then ss.A 9: GoTo E
            mAyFld_Nmtq_IsBool(J) = (mB = "YesNo")
        Next
    Loop Until True
    If jj.Ay_Intersect(mAnFld_Common, mAnFld_Nmtq, mAnFld_RnoNmFld) Then ss.A 9: GoTo E
    If jj.Siz_Ay(mAnFld_Common) = 0 Then ss.A 10, "No common fields in mAnFld_Nmtq & mAnFld_RnoNmFld", "mAnFld_Nmtq,mAnFld_RnoNmFld", Join(mAnFld_Nmtq, cComma), Join(mAnFld_RnoNmFld, cComma): GoTo E
Loop Until True
    
Dim mLnFld$: mLnFld = ""
N = jj.Siz_Ay(mAnFld)
Dim mNNullExpr%: mNNullExpr = 0
For iCno = 1 To mCnoLas
    mV = pRge(0, iCno).Value: Dim mIsNull As Boolean
    If VarType(mV) = vbString Then
        Dim iIdx%: If jj.Fnd_Idx(iIdx, mAnFld_Common, CStr(mV)) Then ss.A 11: GoTo E
        mIsNull = (iIdx = -1)
    Else
        mIsNull = True
    End If
    
    If mIsNull Then
        mLnFld = jj.Add_Str(mLnFld, "'' as NullExpr" & mNNullExpr): mNNullExpr = mNNullExpr + 1
    Else
        'Handle Boolean Type
        If jj.Fnd_Idx(iIdx, mAnFld_Nmtq, CStr(mV)) Then ss.A 11: GoTo E
        If mAyFld_Nmtq_IsBool(iIdx) Then
            mB = jj.Fmt_Str("IIF(x.[{0}],""x"","""") As [{0}]", mV)
            mLnFld = jj.Add_Str(mLnFld, mB)
        Else
            mLnFld = jj.Add_Str(mLnFld, jj.Q_S(mV, "[]"))
        End If
    End If
Next
oSql = jj.Fmt_Str("Select {0} from [{1}] as x{2}", mLnFld, mNmtq0, jj.Cv_Fb2InFb(pFbSrc))
GoTo X
R: ss.R
E: BldSql_Qt = True: ss.B cSub, cMod, "pRge,pNmtq$,pFbSrc$", jj.ToStr_Rge(pRge), pFbSrc
X:
End Function
#If Tst Then
Function BldSql_Qt_Tst() As Boolean
Const cSub$ = "BldSql_Qt_Tst"
Dim mWb As Workbook: If jj.Opn_Wb_R(mWb, "p:\AppDef_Meta\MetaDb.xls", , True) Then Stop: GoTo E
Dim mWs As Worksheet: Set mWs = mWb.Sheets("TblUF")
Dim mRge As Range: Set mRge = mWs.Range("A5")
Dim mSql$: If jj.BldSql_Qt(mSql, mRge, "@TblUF", "p:\workingdir\pgmobj\JMtcDb.mdb") Then Stop: GoTo E
jj.Shw_Dbg cSub, cMod, "mSql", mSql
Exit Function
E:
X: jj.Cls_Wb mWb, , True
End Function
#End If
Function BldSql_AddUpdDlt(oSqlAdd$, oSqlUpd$, oSqlDlt$, pNmtTar$, pNmtSrc$, pNKFld As Byte, Optional pNKFldRmv = 0) As Boolean
'Aim: Add/Upd {pNmtTar} by {pNmtSrc}.  Both has same {pNKFld} of PK.  All fields in {pNmtSrc} should all be found in {TarNmt}
'     If pNKFldRmv>0 then some record in {pNmtTar} will be remove if they does not exist in pNmtSrc having first {pNKFldRmv} as the matching keys
'     Example, Tar & Src: a,b,c, x,y,z
'              pNKFld   : 3
'              pNKFld   : 2
'              Tar: 1,1,3, ..... Src: 1,1,4, ...
'                   1,1,4, .....      1,1,5 ...
'                   1,1,5, .....      1,1,6, ...
'                 : 1,2,3, .....
'                   1,2,4, .....
'                   1,2,5, .....
'    After
'              Tar: 1,1,4
'                   1,1,5
'                   1,1,6
Const cSub$ = "BldSql_AddUpdDlt"
Dim mNmtTar$: mNmtTar = jj.Rmv_SqBkt(pNmtTar)
Dim mNmtSrc$: mNmtSrc = jj.Rmv_SqBkt(pNmtSrc)

Dim mJoin$, mSet$
Dim mAnKey$()
Do
    Dim mLnFld$: mLnFld = jj.ToStr_Nmt(pNmtSrc)
    Dim mAnFld$(): mAnFld = Split(mLnFld, cComma)
    Dim mAnSet$(): If jj.Ay_Cut(mAnKey, mAnSet, mAnFld, CInt(pNKFld)) Then ss.A 1: GoTo E
    mJoin = jj.Fmt_Str_Repeat_Ay("t.{N}=s.{N}", mAnKey, , " and ")
    mSet = jj.Fmt_Str_Repeat_Ay("t.{N}=s.{N}", mAnSet)
Loop Until True

Dim mSql$
oSqlUpd = jj.Fmt_Str("Update [{0}] t inner join [{1}] s on {2} set {3} ", mNmtTar, mNmtSrc, mJoin, mSet)
oSqlAdd = jj.Fmt_Str("Insert into [{0}] Select s.* from [{1}] s left join [{0}] t on {2} where IsNull(t.{3})", mNmtTar, mNmtSrc, mJoin, mAnKey(0))
oSqlDlt = ""
'Find & Run {mSql} for Delete
If pNKFldRmv > 0 Then
    Dim mExprPK$: mExprPK$ = Join(mAnKey, " & ")
    
    ReDim Preserve mAnKey(pNKFldRmv - 1)
    Dim mExpr1stNFld$: mExpr1stNFld = Join(mAnKey, " & ")
    
    'DELETE *
    'From [$MdbS]
    'WHERE Mdb In (Select Mdb from [#MdbS])
    ' AND Mdb & Schm Not In (Select Mdb & Schm from [#MdbS])
    oSqlDlt = jj.Fmt_Str("delete *" & _
        " from [{0}]" & _
        " where {2} in (Select {2} from [{1}])" & _
        " and {3} not in (Select {3} from [{1}])" _
        , mNmtTar, mNmtSrc, mExpr1stNFld$, mExprPK$)
End If
Exit Function
R: ss.R
E: BldSql_AddUpdDlt = True: ss.B cSub, cMod, "pNmtTar,pNmtSrc,pNKFld,pNKFldRmv", pNmtTar, pNmtSrc, pNKFld, pNKFldRmv
End Function
Function BldSql_Ins_ByFrm(oInsSql$, pNmt$, pFrm As Access.Form, pLn1$, Optional pLn2$ = "", Optional pLn3$ = "") As Boolean
Const cSub$ = "BldSql_Ins_ByFrm"
'Aim: Bld {oInsSql} by the new value of controls in {pFrm} of name {pLn1..3}.
'     {pLn1..3} are in fmt aaa=xxx,bbb,ccc.  aaa,bbb,ccc will be used to lookup the form's control & xxx,bbb,ccc will used in the {oInsSql}
Dim mLn$: mLn = pLn1
If pLn2 <> "" Then mLn = mLn & cComma & pLn2
If pLn3 <> "" Then mLn = mLn & cComma & pLn3
Dim mAy1$(), mAy2$(): If jj.Brk_Lm_To2Ay(mAy1, mAy2, mLn$) Then ss.A 1: GoTo E

Dim mInsLst$: mInsLst = Join(mAy2, cComma)
Dim mLstQVal$: If jj.Fnd_LoQVal_ByFrm(mLstQVal, pFrm, mAy1) Then ss.A 2: GoTo E
oInsSql = jj.Fmt_Str("Insert into {0} ({1}) values ({2})", pNmt, mInsLst, mLstQVal)
Exit Function
R: ss.R
E: BldSql_Ins_ByFrm = True: ss.B cSub, cMod, "pNmt,pFrm,pLn1,pLn2,pLn3", pNmt, jj.ToStr_Frm(pFrm), pLn1, pLn2, pLn3
End Function
#If Tst Then
Function BldSql_Ins_ByFrm_Tst() As Boolean
Const cNmFrm$ = "frmIIC_Tst"
Dim mFrm As Access.Form: If jj.Opn_Frm(cNmFrm, , , mFrm) Then Stop: GoTo E
Dim mInsSql$, mIsChg As Boolean, mUpdSql$, mLoChgd$
If jj.BldSql_Ins_ByFrm(mInsSql, "IIC", mFrm, "ItemClass=ICLAS,Des=ICDES") Then Stop
If jj.BldSql_Upd_InFrm(mUpdSql, "IIC", mFrm, "ItemClass=ICLAS", "Des=ICDES", mLoChgd) Then Stop
Debug.Print mInsSql
Debug.Print mUpdSql
Exit Function
E: BldSql_Ins_ByFrm_Tst = True
End Function
#End If
Function BldSql_Upd_InFrm(oUpdSql$, pNmt$, pFrm As Access.Form, pLmPk$, pLmFld$, Optional oLoChgd$) As Boolean
'Aim: Bld {oUpdSql} by the new value of controls in {pFrm} of name {pLn1..3}.
Const cSub = "Upd_ByFrm"
oUpdSql = ""
Dim mLoAsg$: If jj.Fnd_LoAsg_InFrm(mLoAsg, pFrm, pLmFld, oLoChgd) Then ss.A 1: GoTo E
If mLoAsg = "" Then Exit Function
Dim mPKCndn$: If jj.Bld_LExpr_InFrm(mPKCndn, pFrm, pLmPk) Then ss.A 1: GoTo E
oUpdSql = jj.Fmt_Str("Update {0} Set {1} where {2}", pNmt, mLoAsg, mPKCndn)
Exit Function
R: ss.R
E: BldSql_Upd_InFrm = True: ss.B cSub, cMod, "pNmt, pFrm , pLmPk, pLmFld", pNmt, jj.ToStr_Frm(pFrm), pLmPk, pLmFld
End Function

