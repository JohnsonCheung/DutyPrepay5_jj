Attribute VB_Name = "xAy"
#Const Tst = True
Option Compare Text
Option Explicit
Const cMod$ = cLib & ".xAy"
Function Ay_Rmv1stEle(oA$()) As Boolean
Dim N%: N = jj.Siz_Ay(oA)
If N = 0 Then Exit Function
If N = 1 Then jj.Clr_Ays oA: Exit Function
Dim J%
For J = 1 To N - 1
    oA(J - 1) = oA(J)
Next
ReDim Preserve oA(N - 2)
End Function
Function Ay_Cut(oAy1$(), oAy2$(), pAy$(), pN%) As Boolean
'Aim: Split pAy$() into 2, first {pN} in {oAy1} and rest in {oAy2}
Const cSub$ = "Ay_Cut"
Dim N%: N = jj.Siz_Ay(pAy)
If pN > N Then ss.A 1, "pN must be <= siz of pAy", , "siz of pAy, pAy", N: GoTo E
If pN = 0 Then jj.Clr_Ays oAy1: oAy2 = pAy: Exit Function
If pN = N Then jj.Clr_Ays oAy2: oAy1 = pAy: Exit Function
ReDim oAy1(pN - 1)
ReDim oAy2(N - pN - 1)
Dim J%
For J = 0 To pN - 1
    oAy1(J) = pAy(J)
Next
For J = 0 To N - pN - 1
    oAy2(J) = pAy(J + pN)
Next
Exit Function
E: Ay_Cut = True: ss.B cSub, cMod, "pAy,pN", jj.ToStr_Ays(pAy), pN
End Function
#If Tst Then
Function Ay_Cut_Tst() As Boolean
Dim mA$(5), J%
For J = 0 To 5: mA(J) = J: Next
For J = 0 To 6
    Dim mA1$(), mA2$(): If jj.Ay_Cut(mA1, mA2, mA, J) Then Stop
    Debug.Print "Splitting " & J & "...."
    Debug.Print Join(mA1, cComma) & "<---"
    Debug.Print Join(mA2, cComma) & "<---"
Next
End Function
#End If
Function Ay_Intersect(oAyInersect$(), pAy1$(), pAy2$()) As Boolean
Const cSub$ = "Ay_Intersect"
Dim mA$()
Dim mN1%: mN1 = jj.Siz_Ay(pAy1): If mN1 = 0 Then oAyInersect = mA: Exit Function
Dim mN2%: mN2 = jj.Siz_Ay(pAy2): If mN2 = 0 Then oAyInersect = mA: Exit Function
Dim mN%: mN = 0
Dim J%
For J = 0 To mN1 - 1
    Dim I%: If jj.Fnd_Idx(I, pAy2, pAy1(J)) Then ss.A 1: GoTo E
    If I >= 0 Then
        ReDim Preserve oAyInersect(mN)
        oAyInersect(mN) = pAy1(J)
        mN = mN + 1
    End If
Next
Exit Function
R: ss.R
E: Ay_Intersect = True: ss.B cSub, cMod, "pAy1,pAy2", jj.ToStr_Ays(pAy1), jj.ToStr_Ays(pAy2)
End Function
Function Ay_Subtract(oAySubtract$(), pAy1$(), pAy2$()) As Boolean
'Aim: Find oAySubtract = oAy1 - oAy2 (ie, Those elements in exist in oAy1 and not in oAy2
Const cSub$ = "Ay_Subtract"
Dim mA$()
Dim mN1%: mN1 = jj.Siz_Ay(pAy1): If mN1 = 0 Then oAySubtract = mA: Exit Function
Dim mN2%: mN2 = jj.Siz_Ay(pAy2): If mN2 = 0 Then oAySubtract = pAy1: Exit Function
Dim mN%: mN = 0
Dim J%
For J = 0 To mN1 - 1
    Dim I%: If jj.Fnd_Idx(I, pAy2, pAy1(J)) Then ss.A 1: GoTo E
    If I = -1 Then
        ReDim Preserve oAySubtract(mN)
        oAySubtract(mN) = pAy1(J)
        mN = mN + 1
    End If
Next
Exit Function
R: ss.R
E: Ay_Subtract = True: ss.B cSub, cMod, "pAy1,pAy2", jj.ToStr_Ays(pAy1), jj.ToStr_Ays(pAy2)
End Function
#If Tst Then
Function Ay_Subtract_Tst() As Boolean
Dim mAy1$(10), J%
For J = 0 To 10
    mAy1(J) = J
Next
Dim mAy2$()
Dim mAy$()
If jj.Ay_Subtract(mAy, mAy1, mAy2) Then Stop
Debug.Print "mAy1=" & Join(mAy1, cComma)
Debug.Print "mAy=" & Join(mAy, cComma)
mAy1(0) = "xx"
Debug.Print "mAy1(0)=" & mAy1(0)
Debug.Print "mAy(0)=" & mAy(0)
End Function
#End If
