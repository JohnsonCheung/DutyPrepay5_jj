Attribute VB_Name = "xSiz"
Option Compare Text
Option Explicit
Option Base 0
Const cMod$ = cLib & ".xSiz"
Function Siz_AyRgeRno%(pAyRgeRno() As tRgeRno)
On Error GoTo R
Siz_AyRgeRno = UBound(pAyRgeRno) - LBound(pAyRgeRno) + 1
Exit Function
R: Siz_AyRgeRno = 0
End Function
Function Siz_AyDArg%(pAyDArg() As d_Arg)
On Error GoTo R
Siz_AyDArg = UBound(pAyDArg) - LBound(pAyDArg) + 1
Exit Function
R: Siz_AyDArg = 0
End Function
Function Siz_Ay%(pVAy)
On Error GoTo R
Siz_Ay = UBound(pVAy) - LBound(pVAy) + 1
Exit Function
R: Siz_Ay = 0
End Function
Function Siz_Ay_Tst() As Boolean
'Dim mAm(10) As tMap: Debug.Print Ay(CVar(mAm))
Dim mAyInt%(10): Debug.Print jj.Siz_Ay(mAyInt)
Dim mAyLng&(10): Debug.Print jj.Siz_Ay(mAyLng)
Dim mAyBool(9) As Boolean: Debug.Print jj.Siz_Ay(mAyBool)
Dim mAyByt(8) As Byte: Debug.Print jj.Siz_Ay(mAyByt)
Dim mAys$(7): Debug.Print jj.Siz_Ay(mAys)
Dim mAyV(6): Debug.Print jj.Siz_Ay(mAyV)
Dim mAm(5) As tMap: Debug.Print jj.Siz_Am(mAm)
End Function
Function Siz_AyDQry%(pAyDQry() As jj.d_Qry)
On Error GoTo R
Siz_AyDQry = UBound(pAyDQry) - LBound(pAyDQry) + 1
Exit Function
R: Siz_AyDQry = 0
End Function
Function Siz_AyOldQsT%(pAyOldQsT() As jj.d_QsT)
On Error GoTo R
Siz_AyOldQsT = UBound(pAyOldQsT) - LBound(pAyOldQsT) + 1
Exit Function
R: Siz_AyOldQsT = 0
End Function
Function Siz_An2V%(pAn2V() As tNm2V)
On Error GoTo R
Siz_An2V = UBound(pAn2V) - LBound(pAn2V) + 1
Exit Function
R: Siz_An2V = 0
End Function
Function Siz_Am%(pAm() As tMap)
On Error GoTo R
Siz_Am = UBound(pAm) - LBound(pAm) + 1
Exit Function
R: Siz_Am = 0
End Function
Function Siz_AyRgeCno%(pAyRgeCno() As tRgeCno)
On Error GoTo R
Siz_AyRgeCno = UBound(pAyRgeCno) - LBound(pAyRgeCno) + 1
Exit Function
R: Siz_AyRgeCno = 0
End Function
Function Siz_Coll%(pColl As VBA.Collection)
On Error GoTo R
Siz_Coll = pColl.Count
Exit Function
R: Siz_Coll = 0
End Function
Function Siz_Vayv%(pVayv)
Const cSub$ = "Vayv"
On Error GoTo R
If (VarType(pVayv) And vbArray) = 0 Then ss.A 1, "VarTyp of pVayv must be vbArray of something", , "VarTyp of pVayv", VarType(pVayv): GoTo E
Dim mAyV(): mAyV = pVayv
Dim N%: N = UBound(mAyV) - LBound(mAyV) + 1
Dim J%
For J = N - 1 To 0 Step -1
    If Not IsMissing(mAyV(J)) Then
        If Not IsNull(mAyV(J)) Then
            If VarType(mAyV(J)) <> vbString Then Siz_Vayv = J + 1: Exit Function
            If mAyV(J) <> "" Then Siz_Vayv = J + 1: Exit Function
        End If
    End If
Next
Siz_Vayv = 0
Exit Function
R: ss.R
E: ss.B cSub, cMod
   Siz_Vayv = 0
End Function
