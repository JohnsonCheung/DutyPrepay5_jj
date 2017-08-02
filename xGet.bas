Attribute VB_Name = "xGet"
#Const Tst = True
Option Compare Text
Option Explicit
Option Base 0
Const cMod$ = cLib & ".xGet"
Function Get_Am_ByLm(pLm$, Optional pBrkChr$ = "=", Optional pSepChr$ = cSemi) As tMap()
Dim mAmStr$(): mAmStr = Split(pLm, pSepChr$)
Dim NMap%: NMap = jj.Siz_Ay(mAmStr)
If NMap = 0 Then Exit Function
ReDim mAm(0 To NMap - 1) As tMap
Dim I%: For I = 0 To NMap - 1
    If jj.Brk_Str2Map(mAm(I), mAmStr(I), pBrkChr$) Then Exit Function
Next
Get_Am_ByLm = mAm
End Function
Function Get_Am_ByLpVv(pLp$, pVayv) As tMap()
Const cSub$ = "Get_Am_ByLpAp"
On Error GoTo R
Dim mAn$(): mAn = Split(pLp, cComma)
Dim mAyV(): mAyV = pVayv
Dim N1%: N1 = jj.Siz_Ay(mAn)
Dim N2%: N2 = jj.Siz_Ay(mAyV)
If N1 <> N2 Then ss.A 1, "Cnt in pAp & pLp are diff", , "Cnt in pAp,Cnt in pLp", N2, N1: GoTo E
ReDim mAm(N1 - 1) As tMap
If jj.Set_Am_F1(mAm, mAn) Then ss.A 2: GoTo E
Dim J%, mA$
For J = 0 To N1 - 1
    If Not IsMissing(mAyV(J)) Then
        If (VarType(mAyV(J)) And vbArray) Then
            mAm(J).F2 = Join(mAyV(J), ",")
        Else
            mAm(J).F2 = mAyV(J)
        End If
    End If
Next
Get_Am_ByLpVv = mAm
Exit Function
R: ss.R
E: ss.B cSub, cMod, "pLp,pVayv", pLp, jj.ToStr_Vayv(pVayv)
End Function
Function Get_Am_ByLpAp(pLp$, ParamArray pAp()) As tMap()
Get_Am_ByLpAp = Get_Am_ByLpVv(pLp, CVar(pAp))
End Function

