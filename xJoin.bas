Attribute VB_Name = "xJoin"
#Const Tst = True
Option Compare Text
Option Explicit
Option Base 0
Const cMod$ = cLib & ".xJoin"
Function Join_AyV$(pAyV(), Optional pQ$ = "", Optional pSepChr$ = cComma)
Dim mS$, J%
If pQ = "" Then
    For J = 0 To jj.Siz_Ay(pAyV) - 1
        mS = jj.Add_Str(mS, pAyV(J), pSepChr)
    Next
    Join_AyV = mS
    Exit Function
End If
For J = 0 To jj.Siz_Ay(pAyV) - 1
    mS = jj.Add_Str(mS, jj.Q_S(pAyV(J), pQ), pSepChr)
Next
Join_AyV = mS
End Function
Function Join_Prm$(ParamArray pAp())
'Aim: Join {pAp} into a string of Prm which has format:
'     - One line one parameter
'     - In each parameters, the CrLf, Lf or Cr will be replaced by |
'The Reverse Function Join_is jj.Brk_Prm
Dim J%, mA$, N%
mA = Replace(Replace(Replace(pAp(0), vbCrLf, "|"), vbCr, "|"), vbLf, "|")
For J = 1 To UBound(pAp)
    mA = mA & vbCrLf & Replace(Replace(Replace(pAp(J), vbCrLf, "|"), vbCr, "|"), vbLf, "|")
Next
Join_Prm = mA
End Function
#If Tst Then
Function Join_Prm_Tst() As Boolean
Debug.Print jj.Join_Prm("kldfdklf", "lkdjfldkf", "dskfjsdklf")
End Function
#End If
Function Join_Lin$(pLoLin$, Optional pMaxLin% = 15, Optional pSepChr$ = "          ")
'Aim: Join {pLoLin} into at most {pMaxLin} lines in column format
'     Eg. There are 25 lines: Line NN ---.  To join these line by using pMaxLin=10
'         Gives
'                Line 01 ---         Line 11 ---         Line 21 ---
'                Line 02 ---         Line 12 ---         Line 22 ---
'                Line 03 ---         Line 13 ---         Line 23 ---
'                Line 04 ---         Line 14 ---         Line 24 ---
'                Line 05 ---         Line 15 ---         Line 25 ---
'                Line 06 ---         Line 16 ---
'                Line 07 ---         Line 17 ---
'                Line 08 ---         Line 18 ---
'                Line 09 ---         Line 19 ---
'                Line 10 ---         Line 20 ---
Dim mAyLin$(): mAyLin = Split(pLoLin, vbCrLf)
Dim N%: N = jj.Siz_Ay(mAyLin)
Dim NRslt%: NRslt = Fct.MinInt(pMaxLin, N + 1)
ReDim mAyRslt$(NRslt - 1)
Dim J%, iRslt%: For J = 0 To N - 1
    iRslt = J Mod pMaxLin
    mAyRslt(iRslt) = jj.Add_Str(mAyRslt(iRslt), mAyLin(J), pSepChr)
Next
Join_Lin = Join(mAyRslt, vbCrLf)
End Function
#If Tst Then
Function Join_Lin_Tst() As Boolean
Dim mAy$(30)
Dim J%: For J = 0 To 30
    mAy(J) = "Line " & J
Next
Dim mLines$: mLines = Join(mAy, vbCrLf)
Debug.Print jj.Join_Lin(mLines, 10)
End Function
#End If
Function Join_Lv$(pSepChr$, ParamArray pAp())
Dim mA$, mFirst As Boolean
Dim J%: For J = LBound(pAp) To UBound(pAp)
    If IsMissing(pAp(J)) Then GoTo Nxt
    Dim mV$
    If VarType(pAp(J)) = vbDate Then
        mV = Format(pAp(J), "yyyy/mm/dd")
    Else
        mV = CStr(pAp(J))
    End If
    If mA = "" Then
        mA = mV
    Else
        If mV <> "" Then mA = mA & pSepChr & mV
    End If
Nxt:
Next
Join_Lv = mA
End Function
Function Join_Lv_Tst() As Boolean
Debug.Print jj.Join_Lv(cComma, "slkdf", , 12323, #1/1/2007#)
Debug.Print jj.Join_Lv(cComma, "slkdf", "", 12323, #1/1/2007#)
End Function
Function Join_NmV(oS$, pNm$, pV, Optional pBrk$ = "=", Optional pSwap As Boolean = False) As Boolean
'Aim: Join {pNm} and {pV} by {pBrk} into {oS} with optional to swap name and value.  {pV} will always quote by its type. ['] for str, [#] for date
Const cSub$ = "NmV"
On Error GoTo R
If pSwap Then
    oS = jj.Q_V(pV) & pBrk & pNm
Else
    oS = pNm & pBrk & jj.Q_V(pV)
End If
GoTo X
R: ss.R
E: Join_NmV = True: ss.B cSub, cMod, "pNm,pV", pNm, pV
X:
End Function
Function Join_Xls_ByDir(pDirFm$, pFnXlsTo$) As Boolean
'Aim: Join all Xls files in {pDirFm} into one Xls {pFnXlsTo} and delete those original Xls files
'Assume: Each Xls in {pDirFm} has only 1 ws and have ws name and the file same being the same.
Const cSub$ = "Xls_ByDir"
'==Start
If Right(pFnXlsTo, 4) = ".xls" Then ss.A 1, "Given file must be .xls": GoTo E ' After JoinXls all Xls (include pFnXlsTo if it has .xls) will be deleted
If Not jj.IsDir(pDirFm) Then ss.A 2: GoTo E
Dim FfnTo$: FfnTo = pDirFm & pFnXlsTo
If jj.Dlt_Fil(FfnTo) Then ss.A 3: GoTo E

'Create {mWbTo} by copy first Xls in {pDirFm} as {pFnXlsTo}
Dim AyFnXls$(): If jj.Fnd_AyFn(AyFnXls, pDirFm, "*.xls", False) Then ss.A 4: GoTo E
If jj.Siz_Ay(AyFnXls) <= 0 Then ss.A 5, "No file in given dir to join": GoTo E
FileSystem.FileCopy pDirFm & AyFnXls(LBound(AyFnXls)) & ".xls", FfnTo
Dim mWbTo As Workbook: Set mWbTo = gXls.Workbooks.Open(FfnTo)

gXls.DisplayAlerts = False
'Loop each Xls file started from 2nd Xls in {pDirFm}
Dim iFnXls$, J As Byte
For J = 1 To UBound(AyFnXls)
    iFnXls = AyFnXls(J)
    Dim mWbFm As Workbook: If jj.IsSingleWsXls(pDirFm & iFnXls, mWbFm) Then ss.A 6: GoTo E
    
    ''Copy the {mFmWs} to {mWbTo}, then close mWbFm
    Dim mWs As Worksheet: If jj.Crt_Ws_FmWs(mWs, mWbFm.Worksheets(1), , mWbTo) Then ss.A 7: GoTo E
    mWbFm.Close
Next
jj.Dlt_Dir pDirFm, "*.xls"
'Save mWbTo
gXls.DisplayAlerts = False
mWbTo.SaveAs pDirFm & pFnXlsTo
gXls.DisplayAlerts = True
GoTo X
R: ss.R
E: Join_Xls_ByDir = True: ss.B cSub, cMod, "pDirFm$, pFnXlsTo$", pDirFm$, pFnXlsTo$
X:
End Function
#If Tst Then
Function Join_Xls_ByDir_Tst() As Boolean
jj.Join_Xls_ByDir "C:\temp\LT\HB\", "a.xls"
MsgBox "Done"
End Function
#End If
