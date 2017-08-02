Attribute VB_Name = "xCmb"
#Const Tst = True
Option Compare Text
Option Explicit
Option Base 0
Const cMod$ = cLib & ".xCmb"
Function Cmb_Rec_ByCmd(pCmbRecCmd$, Optional pSepChr$ = cComma) As Boolean
Const cSub$ = "Cmb_Rec_ByCmd"
Dim mCmb$, mJoin$, mTo$, mInto$, mKeep$, mOrd$, mStp%
If jj.Brk_Cmb_RecCmd(mCmb, mJoin, mInto, mTo, mKeep, mOrd, mStp, pCmbRecCmd) Then ss.A 1: GoTo E
If jj.Cmb_Rec(mCmb, mJoin, mInto, mTo, mKeep, mOrd, mStp, pSepChr) Then ss.A 2: GoTo E
Exit Function
E: Cmb_Rec_ByCmd = True: ss.B cSub, cMod, "pCmbRecCmd", pCmbRecCmd
End Function
#If Tst Then
Function Cmb_Rec_ByCmd_Tst() As Boolean
Const cCmb$ = "#JnRec"
Const cJoin$ = "XXX"
Const cKeep$ = "abc,def"
'Create Fm
If jj.Crt_Tbl_FmLoFld(cCmb, "abc Text 10,def Long,Sno Int, XXX Text 20") Then Stop
Dim J%, I%, mV_Abc$, mV_Def&, mV_XXX$, mV_Sno%, mSql$
For J = 0 To 19
    mV_Abc = "abc-" & Format(J, "00")
    mV_Def = J * 2
    For I = 0 To 10 + J
        mV_Sno = I
        mV_XXX = J & "-" & I
        mSql = jj.Fmt_Str("Insert into [{0}] (abc,def,Sno,xxx) values('{1}',{2},{3},'{4}')", cCmb, mV_Abc, mV_Def, mV_Sno, mV_XXX)
        If jj.Run_Sql(mSql) Then Stop
    Next
Next
Dim mCmbRecCmd$: mCmbRecCmd = jj.Fmt_Str("Cmb {0} Join {1} Keep {2}", cCmb, cJoin, cKeep)
If jj.Cmb_Rec_ByCmd(mCmbRecCmd) Then Stop: GoTo E
Call DoCmd.OpenTable(cCmb & "_Jn")
Exit Function
E:
    Cmb_Rec_ByCmd_Tst = True
End Function
#End If
Function Cmb_Rec(ByVal pCmb$, ByVal pJoin$ _
        , Optional ByVal pInto$ = "" _
        , Optional ByVal pTo$ = "" _
        , Optional ByVal pKeep$ = "" _
        , Optional ByVal pOrd$ = "" _
        , Optional ByVal pStp% = 10 _
        , Optional ByVal pSepChr$ = cComma _
        ) As Boolean
'Aim: Cmb Join [To] [Into] [Keep] [SetOrd] [Stp]
'           {pCmb}     N record
'           {pJoin}
'           {pInto}    into 1 records   use <pBrk>_Ln
'           {pTo}      to field         use Ln<pSplit>,   if <pSplit> begins with Nm
'                                       use <pJoin>_Jn    else
'           {pKeep}                     use <pBrk>
'           {pOrd}                      use Sno
'           {pStp}
'Eg. pCmd = "Tbl", pJoin="NmFld"
'    Assume: Tbl     has fields: Tbl,LnFld
'    Then  : Tbl_Jn  has fields: Tbl,Sno,NmFld
'Eg. pBrk = "Tbl", pSplit="LoFld"
'    Assume: Tbl     has fields: Tbl,LoFld
'    Then  : Tbl_Ln  has fields: Tbl,Sno,LoFld_Brk
Const cSub$ = "Cmb_Rec"
On Error GoTo R
'Set pCmb (with []), pInto (with []), pTo, pKeep, pOrd
If Left(pCmb, 1) = "(" And Right(pCmb, 1) = ")" Then
    If pInto = "" Then ss.A 1, "pCmb is (..), but pInto is not given": GoTo E
    If pKeep = "" Then ss.A 2, "pCmb is (..), but pKeep is not given": GoTo E
Else
    pCmb = jj.Rmv_SqBkt(pCmb)
    If pInto = "" Then pInto = pCmb & "_Jn"
    pCmb = jj.Q_S(pCmb, "[]")
End If
pInto = jj.Q_S(pInto, "[]")
'
If pKeep = "" Then pKeep = pCmb
'
If pTo = "" Then
    If Left(pJoin, 2) = "Nm" Then
        pTo = "Ln" & mID(pJoin, 3)
    Else
        pTo = pJoin & "_Ln"
    End If
End If
'
If pOrd = "" Then pOrd = "Sno"
'
If pStp = 0 Then pStp = 10

'Create pInto
Dim mSql$
mSql = jj.Fmt_Str("SELECT {0} INTO {1} FROM {2} where False", pKeep, pInto, pCmb):   If jj.Run_Sql(mSql) Then ss.A 1: GoTo E
mSql = jj.Fmt_Str("Alter Table {0} Add [{1}] Memo", pInto, pTo):                     If jj.Run_Sql(mSql) Then ss.A 2: GoTo E

'Opn mRsTar
Dim mRsTar As DAO.Recordset, mRsSrc As DAO.Recordset, mAnFldKeep$(), NKeep%
Do
    Dim mOrdBy$: If pOrd <> "" Then mOrdBy = cComma & pOrd
    mSql = jj.Fmt_Str("Select {0},{1} from {2} Order By {0}{3}", pKeep, pJoin, pCmb, mOrdBy$)
    If jj.Opn_Rs(mRsSrc, mSql) Then ss.A 1: GoTo E
    If mRsSrc.EOF Then GoTo X
    
    Set mRsTar = CurrentDb.OpenRecordset("Select * from " & pInto)
    
    'Find NKeep, mAnFldKeep$()
    mAnFldKeep = Split(pKeep, cComma):
    NKeep = jj.Siz_Ay(mAnFldKeep)
    Dim J%
    For J = 0 To NKeep - 1
        mAnFldKeep(J) = Trim(mAnFldKeep(J))
    Next
Loop Until True

'Opn mRsSrc
With mRsSrc
    ReDim mAyV(NKeep - 1)
    For J = 0 To NKeep - 1
        mAyV(J) = .Fields(mAnFldKeep(J)).Value
    Next
    Dim mLoXXX$: mLoXXX = Nz(.Fields(pJoin).Value, "")
    .MoveNext
    While Not .EOF
        Dim mIsSam As Boolean: mIsSam = False
        For J = 0 To NKeep - 1
            If mAyV(J) <> .Fields(mAnFldKeep(J)).Value Then GoTo Fnd
        Next
        mIsSam = True
Fnd:
        If mIsSam Then
            mLoXXX = jj.Add_Str(mLoXXX, CStr(.Fields(pJoin).Value), pSepChr)
        Else
            mRsTar.AddNew
            For J = 0 To NKeep - 1
                mRsTar.Fields(mAnFldKeep(J)).Value = mAyV(J)
            Next
            mRsTar.Fields(pTo).Value = mLoXXX
            mRsTar.Update
            mLoXXX = .Fields(pJoin).Value
            For J = 0 To NKeep - 1
                 mAyV(J) = .Fields(mAnFldKeep(J)).Value
            Next
        End If
        .MoveNext
    Wend
    If mLoXXX <> "" Then
        mRsTar.AddNew
        For J = 0 To NKeep - 1
            mRsTar.Fields(mAnFldKeep(J)).Value = mAyV(J)
        Next
        mRsTar.Fields(pTo).Value = mLoXXX
        mRsTar.Update
    End If
End With
GoTo X
R: ss.R
E: Cmb_Rec = True: ss.B cSub, cMod, "mSql,pCmb,pJoin,pInto,pTo,pKeep,pOrd,pStp", mSql, pCmb, pJoin, pInto, pTo, pKeep, pOrd, pStp
X:
    jj.Cls_Rs mRsTar
    jj.Cls_Rs mRsSrc
End Function
#If Tst Then
Function Cmb_Rec_Tst() As Boolean
Const cCmb$ = "#JnRec"
Const cJoin$ = "XXX"
Const cKeep$ = "abc,def"
'Create Fm
If jj.Crt_Tbl_FmLoFld(cCmb, "abc Text 10,def Long,Sno Int, XXX Text 20") Then Stop
Dim J%, I%, mV_Abc$, mV_Def&, mV_XXX$, mV_Sno%, mSql$
For J = 0 To 19
    mV_Abc = "abc-" & Format(J, "00")
    mV_Def = J * 2
    For I = 0 To 10 + J
        mV_Sno = I
        mV_XXX = J & "-" & I
        mSql = jj.Fmt_Str("Insert into [{0}] (abc,def,Sno,xxx) values('{1}',{2},{3},'{4}')", cCmb, mV_Abc, mV_Def, mV_Sno, mV_XXX)
        If jj.Run_Sql(mSql) Then Stop
    Next
Next
If jj.Cmb_Rec(cCmb, cJoin, , , cKeep) Then Stop: GoTo E
Call DoCmd.OpenTable(cCmb & "_Jn")
Exit Function
E: Cmb_Rec_Tst = True
End Function
#End If
