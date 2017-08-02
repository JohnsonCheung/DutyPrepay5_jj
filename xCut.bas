Attribute VB_Name = "xCut"
#Const Tst = True
Option Compare Text
Option Explicit
Option Base 0
Const cMod$ = cLib & ".xCut"
Function Cut_Prm$(pS$)
'Aim: Cut the {oPrm} out from pS
Dim mA$: mA = Replace(pS$, "()", "  ")
Dim mP1%: mP1 = InStr(mA, "(")
Dim mP2%: mP2 = InStr(mA, ")")
Cut_Prm = ""
If mP1 = 0 Then Exit Function
If mP2 = 0 Then Exit Function
If mP1 > mP2 Then Exit Function
Cut_Prm = mID(pS, mP1 + 1, mP2 - mP1 - 1)
End Function
#If Tst Then
Function Cut_Prm_Tst() As Boolean
Debug.Print jj.Cut_Prm("lksdjf()lksdjf,(lskdjf, dg 1)klsdj")
End Function
#End If
Function Cut_NonRmk$(pS$)
'Aim: return all the begining remark lines of pS
Dim mA$(): mA = Split(pS, vbCrLf)
Dim mB$, J%
For J = 0 To UBound(mA)
    If Left(mA(J), 1) <> cQSng Then Exit For
    mB = mB & mA(J) & vbCrLf
Next
Cut_NonRmk = mB
End Function
#If Tst Then
Function Cut_NonRmk_Tst() As Boolean
Debug.Print jj.Cut_NonRmk("")
Debug.Print jj.Cut_NonRmk("'lksdjfld" & vbCrLf & "lksdjfdl")
End Function
#End If
Function Cut_Lm(oLm$, pLm$, pLnSubSet$) As Boolean
Const cSub$ = "Cut_Lm"
'Aim: Cut {pLm} into {oLm} by using {pLnSubSet} as sub-set of {pLm}
Dim mAm() As tMap: mAm = jj.Get_Am_ByLm(pLm)
Dim oAm() As tMap, N%
Dim mAnSubSet$(): mAnSubSet = Split(pLnSubSet, ",")
Dim mNSubSet%: mNSubSet = jj.Siz_Ay(mAnSubSet)
Dim J%, I%
For J = 0 To jj.Siz_Am(mAm) - 1
    For I% = 0 To mNSubSet - 1
        If mAm(J).F1 = mAnSubSet(I) Then
            ReDim Preserve oAm(N)
            oAm(N) = mAm(J)
            N = N + 1
        End If
    Next
Next
oLm = ToStr_Am(oAm)
Exit Function
R: ss.R
E: Cut_Lm = True: ss.B cSub, cMod, "pLm,pLnSubSet", pLm, pLnSubSet
End Function
Function Cut_Lv(oLv$, pLp$, pLv$, pLp_SubSet$) As Boolean
Const cSub$ = "Cut_Lv"
'Aim: Cut {pLv} into {oLv} by using {pLp_SubSet} as sub-set of {pLp}
Dim mAn$(), mAyV$(), mAn_SubSet$()
mAn = Split(pLp, cComma)
mAyV = Split(pLv, vbCrLf)
Dim N%: N = jj.Siz_Ay(mAn): If jj.Siz_Ay(mAyV) <> N Then ss.A 1: GoTo E

mAn_SubSet = Split(pLp_SubSet, cComma)
Dim N_SubSet%: N_SubSet = jj.Siz_Ay(mAn_SubSet)
ReDim mAyV_SubSet$(N_SubSet - 1)
Dim I%, I_SubSet%
For I_SubSet = 0 To N_SubSet - 1
    For I = 0 To N - 1
        If mAn(I) = mAn_SubSet(I_SubSet) Then mAyV_SubSet(I_SubSet) = mAyV(I): Exit For
    Next
Next
oLv = Join(mAyV_SubSet, vbCrLf)
Exit Function
R: ss.R
E: Cut_Lv = True: ss.B cSub, cMod, "pLp,pLv,pLp_SubSet", pLp, pLv, pLp_SubSet
End Function
#If Tst Then
Function Cut_Lv_Tst() As Boolean
Const cSub$ = "Lv_Tst"
Dim mLv_SubSet$, mLn_SubSet$, mLn$, mLv$
mLn_SubSet = "aaa,ccc,eeed"
mLn = "aaa,bbb,ccc,ddd,eee"
mLv = "aValue,bValue,cValue,dValue,eValue"
If jj.Cut_Lv(mLv_SubSet, mLn, mLv, mLn_SubSet) Then Stop
jj.Shw_Dbg cSub, cMod
Debug.Print "Result---"
Debug.Print "mLv_SubSet=" & mLv_SubSet
Debug.Print "Prm----"
Debug.Print jj.ToStr_LpAp(vbLf, "mLn,mLv,mLn_SubSet", mLn, mLv, mLn_SubSet)
End Function
#End If
Function Cut_Aft$(pItm$, pSubStr$)
'Aim: Cut {pItm} to get the part after {pSubStr}.  If no pSubStr find, keep the pItm
Dim p%: p = InStr(pItm, pSubStr)
If p <= 0 Then Cut_Aft = pItm: Exit Function
Cut_Aft = mID(pItm, p + Len(pSubStr))
End Function
Function Cut_Ext$(pFil$)
Dim mP1%: mP1 = InStrRev(pFil, ".")
If mP1 = 0 Then Cut_Ext = pFil: Exit Function
Dim mP2%: mP2 = InStrRev(pFil, "\")
If mP1 > mP2 Then Cut_Ext = Left(pFil, mP1 - 1): Exit Function
Cut_Ext = pFil
End Function
Function Cut_FirstLin$(pStr$)
Dim p%: p = InStr(pStr, vbCrLf)
If p = 0 Then Cut_FirstLin = pStr: Exit Function
Cut_FirstLin = mID$(pStr, p + 2)
End Function
Function Cut_LastLin$(pStr$)
Dim mP%: mP = InStrRev(pStr, vbCrLf)
If mP = 0 Then Cut_LastLin = pStr: Exit Function
Cut_LastLin = Left(pStr, mP - 1)
End Function
