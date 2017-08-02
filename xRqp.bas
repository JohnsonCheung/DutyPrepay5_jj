Attribute VB_Name = "xRqp"
#Const Tst = True
Option Base 0
Option Compare Text
Option Explicit
Const cMod$ = cLib$ & ".xRqp"
Function Rqp_Dta2Mdb(pNmLgc$, Optional p As d_RqpDta = Nothing) As Boolean
'Aim: #1. Get mFb_tmpRqp in 2 modes: Remote/Lcl
''        mFb_tmpRqp                    = jj.Sdir_Rqp & "tmpRqp_{NmLgc}.Mdb
'     #2. Link all tables in [mFb_tmpRqp] with pPfx_LnkTbl
Const cSub$ = "Rqp_Dta2Mdb"
If jj.IsNothing(p) Then Set p = New d_RqpDta
Dim mFb_tmpRqp$: mFb_tmpRqp = jj.Sdir_TmpRqp & "tmpRqp_" & pNmLgc & ".Mdb"
If jj.Run_Lgc(pNmLgc, , p.Lm) Then ss.A 1: GoTo E
If jj.Exp_SetNmtq2Mdb("*", mFb_tmpRqp) Then ss.A 2: GoTo E
'#2. Link all tables in [mFb_tmpRqp] with pPfx_LnkTbl
Dim mDb As DAO.Database:    If jj.Cv_Db_FmFb(mDb, mFb_tmpRqp) Then ss.A 3: GoTo E
Dim mAnt$():                If jj.Fnd_Ant_ByLik(mAnt, "*", mDb) Then ss.A 4: GoTo E
Dim mLnt$:                  mLnt = Join(mAnt, cComma)
Dim mLntNew$:               mLntNew = jj.Q_Ln(mLntNew, mLnt, p.Pfx_LnkTbl & "*" & p.Sfx_LnkTbl)
jj.Cls_Db mDb
If jj.Crt_Tbl_FmLnkLnt(mFb_tmpRqp, mLnt, mLntNew) Then ss.A 5: GoTo E
Exit Function
R: ss.R
E: Rqp_Dta2Mdb = True: ss.B cSub, cMod, "pNmLgc", pNmLgc
End Function
#If Tst Then
Function Rqp_Dta2Mdb_Tst() As Boolean
Dim mRqpDta As New d_RqpDta
If Rqp_Dta2Mdb("ExpDir", mRqpDta) Then Stop
End Function
#End If
Function zGenDoc() As Boolean
If jj.Gen_Doc("Rqp") Then Stop
End Function
