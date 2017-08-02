Attribute VB_Name = "xSubstract"
Option Compare Text
Option Explicit
Option Base 0
Const cMod$ = cLib & ".xSubstract"
Function Substract_Lst(oLo$, pLst1$, pLst2$) As Boolean
Const cSub$ = "Substract_Lst"
On Error GoTo R
Dim mAy1$(): mAy1 = Split(pLst1, cComma)
Dim mAy2$(): mAy2 = Split(pLst2, cComma)
If jj.Siz_Ay(mAy2) = 0 Then oLo = pLst1: Exit Function
oLo = ""
Dim N%: N = jj.Siz_Ay(mAy1)
Dim J%: For J = 0 To N - 1
    If Not Fct.InStr_Ay(mAy1(J), mAy2) Then oLo = jj.Add_Str(oLo, mAy1(J), cComma)
Next
R: ss.R
E: Substract_Lst = True: ss.B cSub, cMod, "pLst1,pLst2", pLst1, pLst2
End Function

