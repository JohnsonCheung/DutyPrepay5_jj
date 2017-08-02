Attribute VB_Name = "xRmv"
#Const Tst = True
Option Compare Text
Option Explicit
Option Base 0
Const cMod$ = cLib & ".xRmv"
Function Rmv_Q$(pS$, Optional pQ$ = "'")
Dim mQ1$, mQ2$: If jj.Brk_Q(pQ, mQ1, mQ2) Then Rmv_Q = pS: Exit Function
Dim mA$: mA = pS
If Left(pS, Len(mQ1)) = mQ1 Then mA = mID(mA, Len(mQ1) + 1)
If Right(pS, Len(mQ2)) = mQ2 Then mA = Left(mA, Len(mA) - Len(mQ2))
Rmv_Q = mA
End Function
#If Tst Then
Function Rmv_Q_Tst() As Boolean
Debug.Print Rmv_Q("aaaa", "[]")
jj.Shw_DbgWin
End Function
#End If
Function Rmv_SqBkt$(pS$)
If Left(pS, 1) = "[" And Right(pS, 1) = "]" Then Rmv_SqBkt = mID(pS, 2, Len(pS) - 2): Exit Function
Rmv_SqBkt = pS
End Function
Function Rmv_FirstLin$(pLines$)
Dim p%: p = InStr(pLines, vbCrLf)
If p = 0 Then Rmv_FirstLin = "": Exit Function
Rmv_FirstLin = mID(pLines, p + 2)
End Function
Function Rmv_FirstLin_Tst() As Boolean
Debug.Print jj.Rmv_FirstLin("kdsfjsdjlf")
Debug.Print "...."
Debug.Print jj.Rmv_FirstLin(jj.Fmt_Str("Line1....|Line....2|Line....3"))
End Function
Function Rmv_FirstChr$(pLines$)
Dim Ays$(): Ays = Split(pLines, vbCrLf)
Dim mS$
Dim J%: For J = 0 To jj.Siz_Ay(Ays) - 1
    mS = jj.Add_Str(mS, mID$(Ays(J), 2), vbCrLf)
Next
Rmv_FirstChr = mS
End Function
Function Rmv_Cummulation(pRs As DAO.Recordset, pLnFldKey$, pNmFldCum$, pNmFldSet$) As Boolean
Const cSub$ = "Rmv_Cummulation"
'   Output: the field pRs->pNmFldSet will be Updated
'   Input : pRs         assume it has been sorted in proper order
'           pLnFldKey      is the list of key fields name used as grouping the records in pRs (records with pLnFldKey value considered as a group)
'           pNmFldCum   is the value fields used to do the cummulation to set the pNmFldSet
'           pNmFldSet   is the field required to set
'   Logic:
'           For each group of records in pRs, the pNmFldSet will be set by removing cummulation in the field pValFld.
'           (Note: Assuming pValFld is already in cummulation)
'
If Trim(pLnFldKey) = "" Then ss.A 1, "pLnFldKey is empty string": GoTo E
Dim mAnFldKey$(): mAnFldKey = Split(pLnFldKey, cComma)
Dim NKey%: NKey = jj.Siz_Ay(mAnFldKey)
ReDim mAyKvLas(NKey - 1)
Dim mLasRunningQty As Double
With pRs
    While Not .EOF
        If Not jj.IsSamKey_ByAnFldKey(pRs, mAnFldKey, mAyKvLas) Then
            mLasRunningQty = 0
            Dim J%
            For J = 0 To NKey - 1
                mAyKvLas(J) = pRs.Fields(mAnFldKey(J)).Value
            Next
        End If
        .Edit
        .Fields(pNmFldSet).Value = .Fields(pNmFldCum).Value - mLasRunningQty
        mLasRunningQty = .Fields(pNmFldCum).Value
        .Update
        .MoveNext
    Wend
End With
Exit Function
E: Rmv_Cummulation = True: ss.B cSub, cMod, "pRs,pLnFldKey,pNmFldCum,pNmFldSet", jj.ToStr_Rs(pRs), pLnFldKey, pNmFldCum, pNmFldSet
End Function
Function Rmv_DoubleBlackSlash$(pStr$)
Dim mA$, mP%
mA = pStr
mP = InStr(pStr, "\\")
While mP > 0
    mA = Replace(mA, "\\", "\")
    mP = InStr(mA, "\\")
Wend
Rmv_DoubleBlackSlash = mA
End Function
Function Rmv_Itm_InLst$(pLst$, pLoItmToRmv$)
'Aim: Remove {pLoItmToRmv} in {pLst}
Dim mAy$(): mAy = Split(pLst, cComma)
Dim mAyToRmv$(): mAyToRmv = Split(pLoItmToRmv, cComma)
Dim mAyTar$(), K%: K = 0
Dim mRmv As Boolean
Dim J%: For J = 0 To jj.Siz_Ay(mAy) - 1
    mRmv = False
    Dim I%: For I = 0 To jj.Siz_Ay(mAyToRmv) - 1
        If mAy(J) = mAyToRmv(I) Then mRmv = True: Exit For
    Next
    If Not mRmv Then ReDim Preserve mAyTar(K): mAyTar(K) = mAy(J): K = K + 1
Next
Rmv_Itm_InLst = Join(mAyTar, cComma)
End Function
#If Tst Then
Function Rmv_Itm_InLst_Tst() As Boolean
Const cSub$ = "Itm_InLst_Tst"
Dim mLst$, mLoItmToRmv$
Dim mRslt$, mCase As Byte
For mCase = 1 To 5
    Select Case mCase
    Case 1
        mLst = ""
        mLoItmToRmv = ""
    Case 2
        mLst = "aaa,xxx,yyy"
        mLoItmToRmv = "aaa,bbb,cc"
    Case 3
        mLst = "aaa,xxx,yyy"
        mLoItmToRmv = "111,222"
    Case 4
        mLst = "aaa,xxx,yyy"
        mLoItmToRmv = ""
    Case 5
        mLst = "aaa,xxx,yyy"
        mLoItmToRmv = "aaa,xxx,yyy"
    End Select
    mRslt = jj.Rmv_Itm_InLst(mLst, mLoItmToRmv)
    jj.Shw_Dbg cSub, cMod, "mCase,mRslt,mLst, mLoItmToRmv", mCase, mRslt, mLst, mLoItmToRmv
Next
End Function
#End If
