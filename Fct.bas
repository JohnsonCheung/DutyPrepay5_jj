Attribute VB_Name = "Fct"
#Const Tst = True
Option Compare Text
Option Explicit
Option Base 0
Const cMod$ = cLib & ".Fct"
Function CnnStr_Xls$(pFx$)
'INTO [Excel 8.0;HDR=YES;IMEX=2;DATABASE={0}].{1} FROM {2}"
CnnStr_Xls = "Excel 8.0;HDR=YES;IMEX=2;DATABASE=" & pFx & ";"
End Function
Function CnnStr_Csv$(pFfnCsv)
'Text;DSN=Delta_Tbl_08052203_20080522_033948 Link Specification;FMT=Delimited;HDR=NO;IMEX=2;CharacterSet=936;DATABASE=C:\Tmp;TABLE=Delta_Tbl_08052203_20080522_033948#csv
End Function
Function CnnStr_Mdb$(pFb$)
'    "Provider=Microsoft.JET.OLEDB.4.0;"
CnnStr_Mdb = jj.Fmt_Str( _
    "OLEDB;" & _
    "Provider=Microsoft.ACE.OLEDB.12.0;" & _
    "User ID=Admin;" & _
    "Data Source={0};" & _
    "Mode=Share Deny None;" & _
    "Jet OLEDB:Engine Type=5;" & _
    "Jet OLEDB:Database Locking Mode=1;" & _
    "Jet OLEDB:Global Partial Bulk Ops=2;" & _
    "Jet OLEDB:Global Bulk Transactions=1;" & _
    "Jet OLEDB:Create System Database=False;" & _
    "Jet OLEDB:Encrypt Database=False;" & _
    "Jet OLEDB:Don't Copy Locale on Compact=False;" & _
    "Jet OLEDB:Compact Without Replica Repair=False;" & _
    "Jet OLEDB:SFP=False", pFb)
End Function
Function Sess_ReqNo&()
Static mReqNo&
mReqNo& = mReqNo + 1
Sess_ReqNo = mReqNo
End Function
Function RecCnt%(pNmt$, Optional pDb As DAO.Database = Nothing)
On Error GoTo R
Dim mRecCnt%, mRs As DAO.Recordset
Set mRs = jj.Cv_Db(pDb).OpenRecordset("Select Count(*) from [" & jj.Rmv_SqBkt(pNmt) & "]")
With mRs
    RecCnt = .Fields(0).Value
End With
GoTo X
R: RecCnt = -1
X: jj.Cls_Rs mRs
End Function
Function TimStmp$()
TimStmp = Format(Now, "yyyymmdd_hhmmss")
End Function
Function RTrimSemiQ$(pS$)
Dim mA$: mA = Trim(pS)
While Right(mA, 1) = ";"
    mA = Left(mA, Len(mA) - 1)
Wend
While Right(mA, 1) = ")"
    mA = Left(mA, Len(mA) - 1)
Wend
While Left(mA, 1) = "("
    mA = mID(mA, 2)
Wend
RTrimSemiQ = mA
End Function
Function InStr_SqlKw%(oTypKw As Byte, oKwLen As Byte, pS$)
Dim p%, X%
p = InStrRev(pS, "FROM "):                                          oTypKw = 1: oKwLen = 5
X = InStrRev(pS, "INNER JOIN "): If X > 0 Then If X > p Then p = X: oTypKw = 1: oKwLen = 11
X = InStrRev(pS, "LEFT JOIN "): If X > 0 Then If X > p Then p = X:  oTypKw = 1: oKwLen = 10
X = InStrRev(pS, "RIGHT JOIN "): If X > 0 Then If X > p Then p = X: oTypKw = 1: oKwLen = 11

X = InStrRev(pS, "INTO "): If X > 0 Then If X > p Then p = X:       oTypKw = 2: oKwLen = 5
X = InStrRev(pS, "DELETE "): If X > 0 Then If X > p Then p = X:     oTypKw = 2: oKwLen = 7

X = InStrRev(pS, "SELECT "): If X > 0 Then If X > p Then p = X:     oTypKw = 9: oKwLen = 7
X = InStrRev(pS, "AND "): If X > 0 Then If X > p Then p = X:        oTypKw = 9: oKwLen = 4
X = InStrRev(pS, "SET "): If X > 0 Then If X > p Then p = X:        oTypKw = 9: oKwLen = 4
X = InStrRev(pS, "ON "): If X > 0 Then If X > p Then p = X:         oTypKw = 9: oKwLen = 3
X = InStrRev(pS, "AS "): If X > 0 Then If X > p Then p = X:         oTypKw = 9: oKwLen = 3
X = InStrRev(pS, "WHERE "): If X > 0 Then If X > p Then p = X:      oTypKw = 9: oKwLen = 9
X = InStrRev(pS, "ORDER BY "): If X > 0 Then If X > p Then p = X:   oTypKw = 9: oKwLen = 9
X = InStrRev(pS, "GROUP BY "): If X > 0 Then If X > p Then p = X:   oTypKw = 9: oKwLen = 9
InStr_SqlKw = p
End Function
Function CurMdbDir$()
Static xCurMdbDir$: If xCurMdbDir = "" Then xCurMdbDir = Fct.Nam_DirNam(CurrentDb.Name)
CurMdbDir = xCurMdbDir
End Function
Function CurMdbNam$()
Static xCurMdbNam$: If xCurMdbNam = "" Then xCurMdbNam = Fct.Nam_FilNam(CurrentDb.Name, False)
CurMdbNam = xCurMdbNam
End Function
Function Done() As Boolean
If g.gIsBch Then Exit Function
MsgBox "Done"
End Function
Function FilExt$(pFfn$)
Dim mP%: mP = InStrRev(pFfn, ".")
If mP = 0 Then FilExt = "": Exit Function
FilExt = mID$(pFfn, mP)
End Function
Function Fy$(pYYMM%)
'500 => FY06
If pYYMM = 9999 Then
    Fy = "FY07"
    Exit Function
End If
If pYYMM Mod 100 = 0 Then
    Fy = "FY" & VBA.Format((pYYMM \ 100) + 1, "00")
    Exit Function
End If
If pYYMM Mod 100 = 1 Then
    Fy = "FY" & VBA.Format(pYYMM \ 100, "00")
    Exit Function
End If
Fy = "FY" + VBA.Format((pYYMM \ 100) + 1, "00")
End Function
Function FY_Cur$()
Dim mYYMM%
mYYMM = VBA.Format(Date, "yymm")
FY_Cur = Fy(mYYMM)
End Function
Function FY_Prev$(pNPrev_Year As Byte)
FY_Prev = Fy((Year(Date) - pNPrev_Year - 2000) * 100 + Month(Date))
End Function
'-- pFY is in FYnn Format
'-- Return YYYYMMDD
Function FY_StartDate(pFy$) As Long
FY_StartDate = "20" & VBA.Format(CInt(Right(pFy, 2)) - 1, "00") & "0201"
End Function
Function InStr_Ay%(pStr$, pInAy$())
InStr_Ay = -1
If jj.Siz_Ay(pInAy) = 0 Then Exit Function
Dim Idx%: For Idx = LBound(pInAy) To UBound(pInAy)
    If pInAy(Idx) = pStr Then InStr_Ay = Idx: Exit Function
Next
End Function
Function MaxInt%(pI1%, pI2%)
MaxInt = IIf(pI1 > pI2, pI1, pI2)
End Function
Function MGIWeekNum(pDte As Date) As Byte
If Year(pDte) = 2005 Then
    MGIWeekNum = VBA.Format(pDte, "ww", , vbFirstFullWeek)
    Exit Function
End If
MGIWeekNum = VBA.Format(pDte, "ww")
End Function
Function MinByt(pA As Byte, pB As Byte) As Byte
If pA > pB Then MinByt = pB: Exit Function
MinByt = pA
End Function
Function MinInt%(pI1%, pI2%)
MinInt = IIf(pI1 < pI2, pI1, pI2)
End Function

Function Nam_DirNam$(pFfn$)
Dim mPos%: mPos = InStrRev(pFfn, "\")
Nam_DirNam = Left(pFfn, mPos)  ' include "\" at end
End Function
Function Nam_FilNam$(pFfn$, Optional pWithExt As Boolean = True)
Dim mPos%: mPos = InStrRev(pFfn, "\")
If pWithExt Then Nam_FilNam = mID$(pFfn, mPos + 1): Exit Function
Nam_FilNam = jj.Cut_Ext(mID$(pFfn, mPos + 1))
End Function
Function NmPC$()
'Aim: Get current computer name
NmPC = "FromPC"
End Function
Function NxtNCol$(pCol$, pNCol As Byte)
Const cSub$ = "NxtNCol"
On Error GoTo R
If pNCol = 0 Then NxtNCol = pCol: Exit Function
pCol = UCase(pCol)
If Len(pCol) = 1 Then
    If pCol = "Z" Then NxtNCol = "AA": Exit Function
    NxtNCol = Chr(Asc(pCol) + 1)
    Exit Function
End If
If Len(pCol) <> 2 Then ss.A 1, "Given pCol must be 1 or 2 char": GoTo E
If Right(pCol, 1) = "Z" Then NxtNCol = Chr(Asc(Left(pCol, 1)) + 1) & "A": Exit Function
NxtNCol = Left(pCol, 1) & Chr(Asc(Right(pCol, 1)) + 1)
Exit Function
R: ss.R
E: ss.B cSub, cMod, "pCol,pNCol", pCol, pNCol
End Function
Function Quit() As Boolean
If MsgBox("Quit?", vbYesNo + vbDefaultButton2) = vbYes Then Application.Quit
End Function
Function ChooseBool(pIf1$, pIf2$, pIs As Boolean) As Boolean
Select Case pIf1
Case "Y", "N": ChooseBool = pIf1 = "Y": Exit Function
End Select
Select Case pIf2
Case "Y", "N": ChooseBool = pIf2 = "Y": Exit Function
End Select
ChooseBool = pIs
End Function
Function NonBlank(ParamArray pPrmAy())
Dim J%
For J = 0 To UBound(pPrmAy)
    If Not IsMissing(pPrmAy(J)) Then If Nz(pPrmAy(J), "") <> "" Then NonBlank = pPrmAy(J): Exit Function
Next
End Function
#If Tst Then
Function NonBlank_Tst() As Boolean
Debug.Print TypeName(NonBlank("", 1))
End Function
#End If
Function Start(Optional pMsg$ = "Start?", Optional pTitle$ = "Start?") As Boolean
If g.gIsBch Then Start = True: Exit Function
Start = MsgBox(Replace(pMsg, "|", vbCrLf), vbQuestion + vbYesNo + vbDefaultButton1, pTitle) = vbYes
End Function
Function UnderlineStr$(pS$, Optional pUnderLineChar$ = "-")
UnderlineStr = String(Len(pS), pUnderLineChar)
End Function
Function Version$()
Version = "Verision 2007-03-14@0111"
End Function
Function WaitFor(pFfn$, Optional pMsg$ = "") As Boolean
'Aim: for a file is created.  Return true if "wait for" success.  If cancel waiting by user return false.
Const cSub$ = "WaitFor"
If jj.Opn_Frm("frmWaitFor", 1000 & cComma & pFfn & cComma & pMsg, True) Then ss.A 1, "User has cancel waiting": GoTo E
If VBA.Dir(pFfn) = "" Then ss.A 2, "Opn_Frm('frmWaitFor') returns no error, but pFfn not found.  Strange", eImpossibleReachHere: GoTo E
jj.Dlt_Fil pFfn
Exit Function
E: WaitFor = True: ss.B cSub, cMod, "pFfn,pMsg", pFfn, pMsg
End Function
