Attribute VB_Name = "xQ"
#Const Tst = True
Option Compare Text
Option Explicit
Const cMod$ = cLib & ".xQ"
Function Q_SqBkt$(pS$)
If Left(pS, 1) = "(" And Right(pS, 1) = ")" Then Q_SqBkt = pS: Exit Function
Q_SqBkt = jj.Q_S(pS, "[]")
End Function
#If Tst Then
Function Q_SqBkt_Tst() As Boolean
Debug.Print jj.Q_SqBkt("[dsfdf]")
Debug.Print jj.Q_SqBkt("dsfdf")
End Function
#End If
Function Q_Ln(oLn_wQuote$, pLn$, Optional pQ$ = cQSng) As Boolean
'Aim: Quote each name in {pLn} by {pQ} into {oLn_wQuote}
Const cSub$ = "Q_Ln"
If pLn$ = "" Then oLn_wQuote = "": Exit Function
Dim An$(): An = Split(pLn, cComma)
Dim A$: A = jj.Q_S(jj.Rmv_Q(Trim(An(0)), pQ), pQ)
Dim J%: For J = 1 To jj.Siz_Ay(An) - 1
    A = A & cComma & jj.Q_S(Trim(An(J)), pQ)
Next
oLn_wQuote = A$
Exit Function
E: Q_Ln = True
End Function
#If Tst Then
Function Q_Ln_Tst() As Boolean
Const cSub$ = "Q_Ln_Tst"
Dim mLn$, mLn_wQuote$, mQ$
mLn = "abc,def": mQ = "[*//"
If jj.Q_Ln(mLn_wQuote, mLn, mQ) Then Stop
jj.Shw_Dbg cSub, cMod, "mLn,mLn_wQuote,mQ", mLn, mLn_wQuote, mQ
End Function
#End If
Function Q_MrkUp$(pS$, pTag$)
Q_MrkUp = "<" & pTag & ">" & pS & "</" & pTag & ">"
End Function
Function Q_S$(ByVal pS$, Optional pQ$ = cQSng)
If pS = "" Then Exit Function
Dim Q1$, Q2$: jj.Brk_Q pQ, Q1, Q2
If Left(pS, Len(Q1)) = Q1 Then
    If Right(pS, Len(Q2)) = Q2 Then Q_S = pS: Exit Function
End If
Q_S = Q1 & pS & Q2
End Function
#If Tst Then
#End If
Function Q_S_Tst() As Boolean
Dim mV: mV = Now
Debug.Print Q_S(mV, "#")
End Function
Function Q_V$(pV, Optional pByDblQ As Boolean = False, Optional pNullVal$ = "Null")
Const cSub$ = "Q_V"
Dim mTypDta As VbVarType: mTypDta = VarType(pV)
If mTypDta = vbEmpty Then Exit Function
If mTypDta = vbNull Then Q_V = pNullVal: Exit Function
If (mTypDta And vbArray) <> 0 Then Q_V = "jj.Q.V(pV) Err: pV is Array": Exit Function
On Error GoTo R
Select Case jj.Cv_V2Sim(pV)
Case eTypSim_Bool, eTypSim_Num: Q_V = pV
Case eTypSim_Dte: Q_V = "#" & pV & "#"
Case eTypSim_Str: If pByDblQ Then Q_V = cQDbl & Replace(pV, """", """""") & cQDbl Else Q_V = cQSng & Replace(pV, "'", "''") & cQSng
Case Else: Q_V = pV: Exit Function
End Select
Exit Function
R: ss.R
E: ss.B cSub, cMod, "pV,TypDta(pV)", pV, jj.ToStr_TypDta(VarType(pV))
End Function
Function Q_V_Tst() As Boolean
Dim mV
mV = Now: Debug.Print jj.Q_V(mV)
mV = 11: Debug.Print jj.Q_V(mV)
mV = "12""34": Debug.Print jj.Q_V(mV, True)
mV = "12'34": Debug.Print jj.Q_V(mV)
mV = 12323&: Debug.Print jj.Q_V(mV)
mV = 12323@: Debug.Print jj.Q_V(mV)
mV = 12323!: Debug.Print jj.Q_V(mV)
mV = 12323#: Debug.Print jj.Q_V(mV)
mV = CByte(1): Debug.Print jj.Q_V(mV)
mV = Null: Debug.Print jj.Q_V(mV)
Set mV = New d_Arg: Debug.Print jj.Q_V(mV)
Shw_DbgWin
End Function
