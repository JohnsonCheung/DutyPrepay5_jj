Attribute VB_Name = "xAcpt"
#Const Tst = True
Option Compare Text
Option Explicit
Const cMod$ = cLib & ".Acpt"
Function Acpt_PkVal(pFrm As Access.Form, pLmPk$) As Boolean
'Aim: Due to PK is locked in {pFrm}, this Function Acpt_is to accept PK Val to {pFrm}'s control before to add a new record
Const cSub$ = "PkVal"
Dim mAn_Frm$(), mAn_Host$(): If jj.Brk_Lm_To2Ay(mAn_Frm, mAn_Host, pLmPk) Then ss.A 1: GoTo E
Dim J%, N%: N = jj.Siz_Ay(mAn_Frm)
For J = 0 To N - 1
    Dim mA$: mA = InputBox("Key Value " & J + 1 & " ... " & mAn_Frm(J))
    If mA = "" Then GoTo E
    pFrm.Controls(mAn_Frm(J)).Value = mA
Next
Exit Function
R: ss.R

E:
: ss.B cSub, cMod, "pFrm,pLmPk", jj.ToStr_Frm(pFrm), pLmPk
    Acpt_PkVal = True
End Function
Function Acpt_Dte(oDte As Date _
    , Optional pDteDef As Date = 0 _
    , Optional pDteMin As Date = #1/1/1980# _
    , Optional pDteMax As Date = #12/31/2100# _
    , Optional pAlwTim As Boolean = False _
    , Optional pAlwNull As Boolean = False _
    , Optional oIsNull As Boolean = False _
    ) As Boolean
Const cSub$ = "Dte"
Dim mDteDef As Date: mDteDef = IIf(pDteDef = 0, Date, pDteDef)
If jj.Opn_Frm("frmSelDte", jj.Join_Lv(cComma, mDteDef, pDteMin, pDteMax, pAlwTim, pAlwNull), True) Then ss.A 1: GoTo E
If g.gIsCnl Then
    oDte = 0
    oIsNull = True
    GoTo E
End If

If IsNull(g.gDteSel) Then
    oIsNull = True
    oDte = 0
Else
    oIsNull = False
    oDte = g.gDteSel
End If
Exit Function
R: ss.R

E:
: ss.B cSub, cMod, "pDteDef,pDteMin,pDteMax,pAlwTim,pAlwNull", pDteDef, pDteMin, pDteMax, pAlwTim, pAlwNull
    Acpt_Dte = True
End Function
#If Tst Then
Function Acpt_Dte_Tst() As Boolean
Const cSub$ = "Dte_Tst"
Dim mDte As Date, mIsNull As Boolean
Dim mDteDef As Date: mDteDef = CDate(InputBox("Default Date", , Date))
If jj.Acpt_Dte(mDte, mDteDef, , , True, True, mIsNull) Then MsgBox "Select date is cancelled": Exit Function
MsgBox jj.ToStr_LpAp(vbLf, "Selected Date, IsNull", mDte, mIsNull)
End Function
#End If
